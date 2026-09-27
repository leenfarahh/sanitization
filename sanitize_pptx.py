
#python sanitize_pptx.py deck.pptx --preview (yellow highlights + CSV report)
#python sanitize_pptx.py deck.pptx (redact with "XXXX")
#python sanitize_pptx.py deck.pptx --keywords names.txt --label "[REDACTED]"

import argparse
import base64
import csv
import io
import posixpath
import re
import zipfile
from pathlib import Path
from xml.dom import minidom

#validators
def luhn_ok(digits: str) -> bool:
    total, parity = 0, len(digits) % 2
    for i, ch in enumerate(digits):
        d = int(ch)
        if i % 2 == parity:
            d *= 2
            if d > 9:
                d -= 9
        total += d
    return total % 10 == 0


IBAN_LEN = {"AE": 23, "SA": 24, "JO": 30, "QA": 29}

def iban_ok(value: str) -> bool:
    s = re.sub(r"\s", "", value).upper()
    if IBAN_LEN.get(s[:2]) != len(s):
        return False
    rearranged = s[4:] + s[:4]
    return int("".join(str(int(c, 36)) for c in rearranged)) % 97 == 1


def digits_only(v: str) -> str:
    return re.sub(r"\D", "", v)

#rules (name, regex, validator or None, context keywords or None)
#context rules only when a keyword appears within CONTEXT_WINDOW chars before the match.

SEP = r"[\s\-]?"
CONTEXT_WINDOW = 40

RULES = [
    ("EMAIL", r"[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+", None, None),

    ("IBAN_AE", rf"\bAE\d{{2}}(?:{SEP}\d){{19}}(?![0-9A-Z])", iban_ok, None),
    ("IBAN_SA", rf"\bSA\d{{2}}(?:{SEP}[0-9A-Z]){{20}}(?![0-9A-Z])", iban_ok, None),
    ("IBAN_JO", rf"\bJO\d{{2}}(?:{SEP}[A-Z]){{4}}(?:{SEP}[0-9A-Z]){{22}}(?![0-9A-Z])", iban_ok, None),
    ("IBAN_QA", rf"\bQA\d{{2}}(?:{SEP}[A-Z]){{4}}(?:{SEP}[0-9A-Z]){{21}}(?![0-9A-Z])", iban_ok, None),

    ("CARD", r"\b(?:\d[\s\-]?){12,18}\d\b",
        lambda v: 13 <= len(digits_only(v)) <= 19 and luhn_ok(digits_only(v)), None),

    ("EMIRATES_ID", rf"\b784{SEP}(?:19|20)\d{{2}}{SEP}\d{{7}}{SEP}\d\b", None, None),
    ("KSA_ID_IQAMA", r"\b[12]\d{9}\b", lambda v: luhn_ok(v), None),
    ("QATAR_QID", r"\b[23]\d{10}\b", None,
        ["qid", "qatar id", "id no", "id number", "البطاقة الشخصية", "رقم الهوية"]),
    ("JORDAN_NATIONAL_NO", r"\b\d{10}\b", None,
        ["national no", "national number", "national id", "الرقم الوطني"]),
    ("PASSPORT", r"\b[A-Z]{1,2}\d{6,8}\b", None,
        ["passport", "جواز", "رقم الجواز"]),

    #phones: international and local formats
    ("PHONE_UAE", rf"(?:(?:\+|00)971{SEP}|\b0)(?:5[024568]|[2-4679]){SEP}\d{{3}}{SEP}\d{{4}}\b", None, None),
    ("PHONE_KSA", rf"(?:(?:\+|00)966{SEP}|\b0)(?:5\d|1[1-7]){SEP}\d{{3}}{SEP}\d{{4}}\b", None, None),
    ("PHONE_JO",  rf"(?:(?:\+|00)962{SEP}|\b0)(?:7[789]|[2356]){SEP}\d{{3}}{SEP}\d{{4}}\b", None, None),
    ("PHONE_QA",  rf"(?:\+|00)974{SEP}[34567]\d{{3}}{SEP}\d{{4}}\b", None, None),
]

COMPILED = [(n, re.compile(p), v, c) for n, p, v, c in RULES]

#arabic-indic and persian digits -> ASCII (1:1 char mapping keeps indices aligned)
DIGIT_MAP = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

#detection
def find_matches(text: str, keywords=None):
    #return list of (start, end, rule_name), first/longest wins when there's overlap
    norm = text.translate(DIGIT_MAP)
    lower = norm.lower()
    hits = []

    for name, rx, validator, context in COMPILED:
        for m in rx.finditer(norm):
            if validator and not validator(m.group()):
                continue
            if context:
                window = lower[max(0, m.start() - CONTEXT_WINDOW):m.start()]
                if not any(k in window for k in context):
                    continue
            hits.append((m.start(), m.end(), name))

    for kw in keywords or []:
        for m in re.finditer(re.escape(kw), norm, re.IGNORECASE):
            hits.append((m.start(), m.end(), "KEYWORD"))

    hits.sort(key=lambda h: (h[0], -(h[1] - h[0])))
    result, last_end = [], -1
    for h in hits:
        if h[0] >= last_end:
            result.append(h)
            last_end = h[1]
    return result


def mask(value: str) -> str:
    v = value.strip()
    return "*" * max(0, len(v) - 4) + v[-4:]

#xml helpers (stdlib minidom)

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

RPR_AFTER_HIGHLIGHT = {"uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
                       "hlinkClick", "hlinkMouseOver", "rtl", "extLst"}
CLEAR_META = {"creator", "lastModifiedBy", "Company", "Manager"}
EMBEDDED_OOXML = (".xlsx", ".xlsm", ".docx", ".pptx")

BLANK_JPEG = base64.b64decode(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD0iiiigD//2Q==")
BLANK_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGM4ceIEAAS0AlkWLoFAAAAAAElFTkSuQmCC")


def is_el(node, ns=None, local=None):
    return (node.nodeType == node.ELEMENT_NODE
            and (ns is None or node.namespaceURI == ns)
            and (local is None or node.localName == local))


def iter_elements(node):
    for child in node.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            yield child
            yield from iter_elements(child)


def children(node, ns=None, local=None) -> list[minidom.Element]:
    return [c for c in node.childNodes if is_el(c, ns, local)]


def first_child(node, ns, local):
    found = children(node, ns, local)
    return found[0] if found else None


def get_text(el):
    return "".join(c.data for c in el.childNodes if c.nodeType in (c.TEXT_NODE, c.CDATA_SECTION_NODE))


def set_text(el, value):
    for c in list(el.childNodes):
        el.removeChild(c)
    el.appendChild(el.ownerDocument.createTextNode(value))


def new_el(like, local):
    #create an element in the same namespace/prefix as an existing one.
    qname = f"{like.prefix}:{local}" if like.prefix else local
    return like.ownerDocument.createElementNS(like.namespaceURI, qname)


def parse_root(data):
    root = minidom.parseString(data).documentElement
    assert root is not None #parseString always yields a root element
    return root


def serialize(root, original):
    m = re.match(rb"\s*(<\?xml[^>]*\?>)", original)
    decl = m.group(1) if m else b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    return decl + b"\r\n" + root.toxml().encode("utf-8")


class Job:
    def __init__(self, keywords, label, preview):
        self.keywords, self.label, self.preview = keywords, label, preview
        self.rows = []

    def log(self, loc, kind, value=""):
        self.rows.append((loc, kind, mask(value) if value else ""))

    def redact_string(self, s, loc):
        #plain-string redaction (attributes, chart caches, cells).
        if not s:
            return s
        matches = find_matches(s, self.keywords)
        for start, end, rule in matches:
            self.log(loc, rule, s[start:end])
        if self.preview:
            return s
        for start, end, _ in reversed(matches):
            s = s[:start] + self.label + s[end:]
        return s

#paragraphs (a:p)
def segments(p):
    segs, pos = [], 0
    for el in children(p, NS_A):
        if el.localName in ("r", "fld"):
            t = first_child(el, NS_A, "t")
            txt = get_text(t) if t else ""
            segs.append((el, pos, pos + len(txt), txt))
            pos += len(txt)
        elif el.localName == "br":
            segs.append((el, pos, pos + 1, "\n"))
            pos += 1
    return segs


def split_run(run, k):
    new = run.cloneNode(True)
    t1, t2 = first_child(run, NS_A, "t"), first_child(new, NS_A, "t")
    full = get_text(t1)
    set_text(t1, full[:k])
    set_text(t2, full[k:])
    run.parentNode.insertBefore(new, run.nextSibling)


def add_highlight(run):
    rpr = first_child(run, NS_A, "rPr")
    if rpr is None:
        rpr = new_el(run, "rPr")
        run.insertBefore(rpr, run.firstChild)
    for old in children(rpr, NS_A, "highlight"):
        rpr.removeChild(old)
    hl = new_el(run, "highlight")
    clr = new_el(run, "srgbClr")
    clr.setAttribute("val", "FFFF00")
    hl.appendChild(clr)
    anchor = next((c for c in children(rpr) if c.localName in RPR_AFTER_HIGHLIGHT), None)
    rpr.insertBefore(hl, anchor)


def process_paragraph(p, job, loc):
    text = "".join(s[3] for s in segments(p))
    matches = find_matches(text, job.keywords)
    for start, end, rule in reversed(matches): #right-to-left keeps offsets valid
        job.log(loc, rule, text[start:end])
        for boundary in (end, start): #align run edges with the match
            for el, s, e, _ in segments(p):
                if s < boundary < e and el.localName == "r":
                    split_run(el, boundary - s)
                    break
        inside = [el for el, s, e, _ in segments(p) if s < end and e > start]
        runs = [el for el in inside if el.localName in ("r", "fld")]
        if not runs:
            continue
        if job.preview:
            for r in runs:
                if r.localName == "r":
                    add_highlight(r)
        else:
            set_text(first_child(runs[0], NS_A, "t"), job.label)
            for el in inside:
                if el is not runs[0]:
                    p.removeChild(el)
    return bool(matches)

#xml parts
def has_ancestor(el, ns, local):
    node = el.parentNode
    while node is not None and node.nodeType == node.ELEMENT_NODE:
        if is_el(node, ns, local):
            return True
        node = node.parentNode
    return False


def process_xml(name, data, job, loc):
    root = parse_root(data)
    before = len(job.rows)
    changed = False

    if name.endswith(".rels"):
        for rel in children(root):
            t = rel.getAttribute("Target")
            if rel.getAttribute("TargetMode") == "External" and t.lower().startswith(("mailto:", "tel:")):
                rel.setAttribute("Target", job.redact_string(t, loc + " (hyperlink)"))
    else:
        for p in root.getElementsByTagNameNS(NS_A, "p"):
            changed |= process_paragraph(p, job, loc)

        for el in list(iter_elements(root)):
            ln, ns = el.localName, el.namespaceURI
            if ln == "cNvPr": #alt text
                for attr in ("descr", "title"):
                    if el.getAttribute(attr):
                        el.setAttribute(attr, job.redact_string(el.getAttribute(attr), loc + " (alt text)"))
            elif ns == NS_C and ln == "v" and has_ancestor(el, NS_C, "strCache"):
                set_text(el, job.redact_string(get_text(el), loc + " (chart labels)"))
            elif ns == NS_P and ln == "text": #legacy comments
                set_text(el, job.redact_string(get_text(el), loc + " (comment)"))
            elif ns == NS_S and ln == "t": #excel strings
                set_text(el, job.redact_string(get_text(el), loc))
            elif (ns == NS_S and ln == "c" and el.getAttribute("t") in ("", "n")
                  and first_child(el, NS_S, "f") is None): #numeric Excel cell
                v = first_child(el, NS_S, "v")
                if v is not None and get_text(v):
                    new = job.redact_string(get_text(v), loc)
                    if new != get_text(v):
                        el.removeChild(v)
                        el.setAttribute("t", "inlineStr")
                        is_, t = new_el(el, "is"), new_el(el, "t")
                        set_text(t, new)
                        is_.appendChild(t)
                        el.appendChild(is_)
            elif ln in ("cmAuthor", "author") and el.getAttribute("name"): #comment authors
                if not job.preview:
                    job.log(loc, "COMMENT_AUTHOR_CLEARED")
                    el.setAttribute("name", "Author")
                    if el.hasAttribute("initials"):
                        el.setAttribute("initials", "A")
                    if el.hasAttribute("userId"):
                        el.setAttribute("userId", "")

        if name.startswith("docProps/"): #file properties
            for el in list(iter_elements(root)):
                if not children(el) and get_text(el).strip():
                    if el.localName in CLEAR_META:
                        job.log("File properties", f"METADATA_{el.localName.upper()}", get_text(el))
                        if not job.preview:
                            set_text(el, "")
                    else:
                        set_text(el, job.redact_string(get_text(el), "File properties"))

        if re.match(r"ppt/slides/slide\d+\.xml$", name):
            if root.getAttribute("show") == "0":
                job.log(loc, "HIDDEN_SLIDE")
            if root.getElementsByTagNameNS(NS_A, "blip"):
                job.log(loc, "IMAGE_REVIEW_MANUALLY")

    if changed or len(job.rows) > before:
        return serialize(root, data)
    return data

#package
def rels_of(z, names, part):
    d, f = posixpath.split(part)
    rp = posixpath.join(d, "_rels", f + ".rels")
    if rp not in names:
        return []
    out = []
    for rel in children(parse_root(z.read(rp))):
        if rel.getAttribute("TargetMode") == "External":
            continue
        t = rel.getAttribute("Target")
        full = t[1:] if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
        out.append((rel.getAttribute("Id"), rel.getAttribute("Type").rsplit("/", 1)[-1], full))
    return out


def build_locations(z):
    #map every part to a readable location, following slide order in presentation.xml.
    names = set(z.namelist())
    loc = {}
    pres = "ppt/presentation.xml"
    if pres in names:
        rid = {i: full for i, _, full in rels_of(z, names, pres)}
        root = parse_root(z.read(pres))
        skip = {"slideLayout", "slideMaster", "notesMaster", "theme", "slide"}
        for n, sid in enumerate(root.getElementsByTagNameNS(NS_P, "sldId"), 1):
            slide = rid.get(sid.getAttributeNS(NS_R, "id"))
            if not slide:
                continue
            loc[slide] = f"Slide {n}"
            stack = [slide]
            while stack:
                part = stack.pop()
                for _, typ, full in rels_of(z, names, part):
                    if full in loc or typ in skip:
                        continue
                    suffix = {"notesSlide": " notes", "chart": " chart"}.get(typ, "")
                    loc[full] = loc[part] if loc[part].endswith(suffix) else loc[part] + suffix
                    stack.append(full)
    for n in names:
        if n not in loc:
            if "slideLayout" in n:
                loc[n] = "Slide layouts"
            elif "slideMaster" in n:
                loc[n] = "Slide master"
            elif n.startswith("docProps/"):
                loc[n] = "File properties"
            elif "commentAuthors" in n or n.endswith("authors.xml"):
                loc[n] = "Comment authors"
            else:
                loc[n] = posixpath.basename(n)
    return loc


def process_package(data, job, parent_loc=None):
    zin = zipfile.ZipFile(io.BytesIO(data))
    locs = build_locations(zin)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.is_dir(): #folder entries carry no data
                zout.writestr(info, b"")
                continue
            name, blob = info.filename, zin.read(info.filename)
            owner = re.sub(r"_rels/(.+)\.rels$", r"\1", name) # a .rels belongs to its part
            loc = parent_loc or locs.get(owner, name)
            if name.endswith((".xml", ".rels")) and name != "[Content_Types].xml":
                blob = process_xml(name, blob, job, loc)
            elif name.lower().endswith(EMBEDDED_OOXML):
                blob = process_package(blob, job, loc + " (embedded data)")
            elif "embeddings/" in name:
                job.log(loc, "EMBEDDED_OBJECT_NOT_SCANNED")
            elif name.startswith("docProps/thumbnail") and not job.preview:
                blob = BLANK_PNG if name.lower().endswith(".png") else BLANK_JPEG
                job.log("File properties", "THUMBNAIL_BLANKED")
            zout.writestr(info, blob)
    return out.getvalue()


def natural_key(row):
    m = re.match(r"Slide (\d+)", row[0])
    return (0, int(m.group(1))) if m else (1, 0)


def main():
    ap = argparse.ArgumentParser(description="Rule-based PII redaction for .pptx (no dependencies)")
    ap.add_argument("pptx")
    ap.add_argument("--preview", action="store_true", help="highlight only, no removal")
    ap.add_argument("--keywords", help="text file, one term per line")
    ap.add_argument("--label", default="XXXX")
    args = ap.parse_args()

    keywords = []
    if args.keywords:
        keywords = [l.strip() for l in Path(args.keywords).read_text(encoding="utf-8").splitlines() if l.strip()]

    src = Path(args.pptx)
    job = Job(keywords, args.label, args.preview)
    result = process_package(src.read_bytes(), job)

    suffix = "_preview" if args.preview else "_redacted"
    out = src.with_name(src.stem + suffix + src.suffix)
    out.write_bytes(result)

    rep = src.with_name(src.stem + suffix + "_report.csv")
    with open(rep, "w", newline="", encoding="utf-8-sig") as f: # utf-8-sig opens cleanly in Excel
        w = csv.writer(f)
        w.writerow(["location", "type", "masked_value"])
        w.writerows(sorted(job.rows, key=natural_key))

    found = sum(1 for r in job.rows if r[2])
    print(f"{found} items found, {len(job.rows) - found} review notes -> {out}\nReport -> {rep}")


if __name__ == "__main__":
    main()