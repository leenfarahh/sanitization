
#python sanitize_pptx.py deck.pptx --preview (yellow highlights + CSV report)
#python sanitize_pptx.py deck.pptx (redact with IDs + restore key .xlsx that stays with the client)
#python sanitize_pptx.py deck.pptx --keywords names.txt
#python sanitize_pptx.py returned_deck.pptx --restore deck_restore_key.xlsx (put the original values back)

import argparse
import base64
import csv
import io
import json
import posixpath
import re
import secrets
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path
from xml.dom import minidom
from xml.sax.saxutils import escape

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

#ids that replace matches in the deck, numbered per name (PhoneUAE1, PhoneUAE2, Email1...)
ID_PREFIX = {
    "EMAIL": "Email",
    "IBAN_AE": "IBAN", "IBAN_SA": "IBAN", "IBAN_JO": "IBAN", "IBAN_QA": "IBAN",
    "CARD": "CardNum",
    "EMIRATES_ID": "EmiratesID",
    "KSA_ID_IQAMA": "SaudiID",
    "QATAR_QID": "QID",
    "JORDAN_NATIONAL_NO": "NationalNo",
    "PASSPORT": "PassportNum",
    "PHONE_UAE": "PhoneUAE", "PHONE_KSA": "PhoneKSA", "PHONE_JO": "PhoneJO", "PHONE_QA": "PhoneQA",
}

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
NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

RPR_AFTER_HIGHLIGHT = {"uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
                       "hlinkClick", "hlinkMouseOver", "rtl", "extLst"}
CLEAR_META = {"creator", "lastModifiedBy", "Company", "Manager"}
AUTHOR_PLACEHOLDER = {"name": "Author", "initials": "A", "userId": ""}
EMBEDDED_OOXML = (".xlsx", ".xlsm", ".docx", ".pptx")

#the redacted deck carries a random id in its custom properties, the key carries the same one
CUSTOM_PROPS = "docProps/custom.xml"
DECK_ID_PROP = "SanitizationID"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
LAYOUT_GROWTH = 5 #restored text this many characters longer than its id is flagged for a layout check

NOTE_TEXT = {
    "IMAGE_REVIEW_MANUALLY": "Has images: text inside images is not redacted, check them by hand",
    "HIDDEN_SLIDE": "Hidden slide (scanned and redacted like the others)",
    "EMBEDDED_OBJECT_NOT_SCANNED": "Embedded object that could not be scanned: check it by hand",
    "THUMBNAIL_BLANKED": "File preview image blanked (not restored, it would show the pre-edit first slide)",
    "COMMENT_AUTHOR_CLEARED": "Comment author cleared (original kept in the Restore key sheet)",
}

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


def new_el(like, local, ns=None, prefix=None):
    #create an element in the same document, by default in the same namespace/prefix as an existing one.
    if ns is None:
        ns, prefix = like.namespaceURI, like.prefix
    qname = f"{prefix}:{local}" if prefix else local
    return like.ownerDocument.createElementNS(ns, qname)


def xmlns_prefix(root, ns):
    #prefix the root element declares for a namespace, if any
    return next((a.localName for a in root.attributes.values() if a.prefix == "xmlns" and a.value == ns), None)


def parse_root(data):
    root = minidom.parseString(data).documentElement
    assert root is not None #parseString always yields a root element
    return root


def serialize(root, original):
    m = re.match(rb"\s*(<\?xml[^>]*\?>)", original)
    decl = m.group(1) if m else b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    return decl + b"\r\n" + root.toxml().encode("utf-8")


def part_path(pkg, name):
    #parts inside embedded files are addressed as outer!inner
    return "!".join([*pkg, name])


class Job:
    def __init__(self, keywords, labels, preview):
        self.keywords, self.labels, self.preview = keywords, labels, preview
        self.rows = []
        self.ids, self.entries, self.fields = {}, {}, []
        self.counts, self.placed = Counter(), Counter()
        self.authors = 0
        self.pkg = [] #embedded file being processed, if any
        self.deck_id = None if preview else secrets.token_hex(6)

    def log(self, loc, kind, value=""):
        self.rows.append((loc, kind, mask(value) if value else ""))

    def token(self, rule, value, loc, numeric=False):
        #id that replaces value in the deck: same value, same id
        prefix = ID_PREFIX.get(rule) or self.labels.get(value.translate(DIGIT_MAP).lower(), "Keyword")
        tid = self.ids.get((prefix, value))
        if tid is None:
            self.counts[prefix] += 1
            tid = self.ids[prefix, value] = f"{prefix}{self.counts[prefix]}"
            self.entries[tid] = {"id": tid, "type": rule, "value": value, "locations": [], "numeric": False}
        entry = self.entries[tid]
        entry["numeric"] |= numeric
        if loc not in entry["locations"]:
            entry["locations"].append(loc)
        self.placed[tid] += 1
        return tid

    def keep(self, label, type_, value, loc, **target):
        #originals with no text to hold an id (file properties, comment authors)
        target["part"] = part_path(self.pkg, target["part"])
        self.fields.append({"id": label, "type": type_, "value": value, "locations": [loc], "target": target})

    def redact_string(self, s, loc, numeric=False):
        #plain-string redaction (attributes, chart caches, cells).
        if not s:
            return s
        matches = find_matches(s, self.keywords)
        for start, end, rule in matches:
            self.log(loc, rule, s[start:end])
        if self.preview:
            return s
        tokens = [self.token(rule, s[start:end], loc, numeric) for start, end, rule in matches]
        for (start, end, _), tok in reversed(list(zip(matches, tokens))):
            s = s[:start] + tok + s[end:]
        return s

    def xml(self, name, blob, loc):
        return process_xml(name, blob, self, loc)

    def other(self, name, blob, loc):
        if "embeddings/" in name:
            self.log(loc, "EMBEDDED_OBJECT_NOT_SCANNED")
        elif name.startswith("docProps/thumbnail") and not self.preview:
            self.log("File properties", "THUMBNAIL_BLANKED")
            return BLANK_PNG if name.lower().endswith(".png") else BLANK_JPEG
        return blob

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


def isolate(p, start, end):
    #split runs so [start, end) covers whole elements; return (elements inside, text runs inside)
    for boundary in (end, start):
        for el, s, e, _ in segments(p):
            if s < boundary < e and el.localName == "r":
                split_run(el, boundary - s)
                break
    inside = [el for el, s, e, _ in segments(p) if s < end and e > start]
    return inside, [el for el in inside if el.localName in ("r", "fld")]


def replace_span(p, inside, runs, value):
    #value takes the first run's formatting, the rest of the span is removed
    set_text(first_child(runs[0], NS_A, "t"), value)
    for el in inside:
        if el is not runs[0]:
            p.removeChild(el)


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
    for start, end, rule in matches:
        job.log(loc, rule, text[start:end])
    #IDs are assigned left to right, then spliced right to left so offsets stay valid
    tokens = [None if job.preview else job.token(rule, text[start:end], loc) for start, end, rule in matches]
    for (start, end, _), tok in reversed(list(zip(matches, tokens))):
        inside, runs = isolate(p, start, end)
        if not runs:
            continue
        if job.preview:
            for r in runs:
                if r.localName == "r":
                    add_highlight(r)
        else:
            replace_span(p, inside, runs, tok)
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
                    new = job.redact_string(get_text(v), loc, numeric=True)
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
                    job.authors += 1
                    for attr, placeholder in AUTHOR_PLACEHOLDER.items():
                        if el.getAttribute(attr):
                            job.keep(f"(comment author {job.authors})", f"COMMENT_AUTHOR_{attr.upper()}",
                                     el.getAttribute(attr), loc,
                                     kind="author", part=name, el=ln, id=el.getAttribute("id"), attr=attr)
                            el.setAttribute(attr, placeholder)

        if name.startswith("docProps/"): #file properties
            for el in list(iter_elements(root)):
                if not children(el) and get_text(el).strip():
                    if el.localName in CLEAR_META:
                        kind = f"METADATA_{el.localName.upper()}"
                        job.log("File properties", kind, get_text(el))
                        if not job.preview:
                            job.keep("(file property)", kind, get_text(el), loc,
                                     kind="property", part=name, el=el.localName, ns=el.namespaceURI)
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

#deck id (custom file property)
def deck_id_props(root):
    return [p for p in children(root, NS_CUSTOM, "property") if p.getAttribute("name") == DECK_ID_PROP]


def read_deck_id(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    if CUSTOM_PROPS in z.namelist():
        for p in deck_id_props(parse_root(z.read(CUSTOM_PROPS))):
            return "".join(get_text(v) for v in children(p))
    return None


def custom_props_xml(deck_id):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<Properties xmlns="{NS_CUSTOM}" xmlns:vt="{NS_VT}">'
            f'<property fmtid="{FMTID}" pid="2" name="{DECK_ID_PROP}"><vt:lpwstr>{deck_id}</vt:lpwstr></property>'
            '</Properties>').encode("utf-8")


def stamp_deck_id(name, blob, deck_id, names):
    #add the deck ID to custom.xml, or register a new custom.xml when the deck has none
    if name == CUSTOM_PROPS:
        root = parse_root(blob)
        for old in deck_id_props(root):
            root.removeChild(old)
        pids = [int(p.getAttribute("pid")) for p in children(root, NS_CUSTOM, "property")
                if p.getAttribute("pid").isdigit()]
        if xmlns_prefix(root, NS_VT) != "vt":
            root.setAttribute("xmlns:vt", NS_VT)
        prop = new_el(root, "property")
        for k, v in (("fmtid", FMTID), ("pid", str(max(pids, default=1) + 1)), ("name", DECK_ID_PROP)):
            prop.setAttribute(k, v)
        val = new_el(root, "lpwstr", NS_VT, "vt")
        set_text(val, deck_id)
        prop.appendChild(val)
        root.appendChild(prop)
        return serialize(root, blob)
    if CUSTOM_PROPS in names:
        return blob
    if name == "[Content_Types].xml":
        root = parse_root(blob)
        el = new_el(root, "Override")
        el.setAttribute("PartName", "/" + CUSTOM_PROPS)
        el.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")
        root.appendChild(el)
        return serialize(root, blob)
    if name == "_rels/.rels":
        root = parse_root(blob)
        taken = {r.getAttribute("Id") for r in children(root)}
        n = len(taken) + 1
        while f"rId{n}" in taken:
            n += 1
        el = new_el(root, "Relationship")
        el.setAttribute("Id", f"rId{n}")
        el.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
        el.setAttribute("Target", CUSTOM_PROPS)
        root.appendChild(el)
        return serialize(root, blob)
    return blob

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


def slide_key(loc):
    m = re.match(r"Slide (\d+)", loc)
    return (0, int(m.group(1))) if m else (1, 0)


def row_order(row):
    return slide_key(row[0])


def process_package(data, job, parent_loc=None):
    #job is a Job (redact/preview) or a Restore; both expose xml(), other(), pkg and deck_id
    zin = zipfile.ZipFile(io.BytesIO(data))
    names = set(zin.namelist())
    locs = build_locations(zin)
    stamp = parent_loc is None and job.deck_id

    def loc_of(name):
        owner = re.sub(r"_rels/(.+)\.rels$", r"\1", name) # a .rels belongs to its part
        return parent_loc or locs.get(owner, name)

    #process in slide order so IDs number from slide 1, write back in the original order
    done = {}
    files = [i.filename for i in zin.infolist() if not i.is_dir()]
    for name in sorted(files, key=lambda n: slide_key(loc_of(n))):
        blob, loc = zin.read(name), loc_of(name)
        if name.endswith((".xml", ".rels")) and name != "[Content_Types].xml":
            blob = job.xml(name, blob, loc)
        elif name.lower().endswith(EMBEDDED_OOXML):
            job.pkg.append(name)
            blob = process_package(blob, job, loc + " (embedded data)")
            job.pkg.pop()
        else:
            blob = job.other(name, blob, loc)
        done[name] = stamp_deck_id(name, blob, job.deck_id, names) if stamp else blob

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, b"" if info.is_dir() else done[info.filename]) #folder entries carry no data
        if stamp and CUSTOM_PROPS not in names:
            zout.writestr(CUSTOM_PROPS, custom_props_xml(job.deck_id))
    return out.getvalue()

#restore
class Restore:
    def __init__(self, entries, fields=()):
        self.entries = {e["id"].lower(): e for e in entries}
        ids = sorted(self.entries, key=len, reverse=True) #longest first, so PhoneUAE12 wins over PhoneUAE1
        self.rx = re.compile("(?:" + "|".join(map(re.escape, ids)) + r")(?![0-9])", re.I) if ids else None
        self.fields = [f for f in fields if f["restore"]]
        self.done = set() #indices of fields put back
        self.found, self.restored = Counter(), Counter()
        self.layout = set()
        self.pkg, self.deck_id, self.dirty = [], None, False

    def hits(self, text):
        #count every ID in text, return the ones marked for restore
        if not self.rx:
            return []
        found = [(m.start(), m.end(), self.entries[m.group().lower()])
                 for m in self.rx.finditer(text.translate(DIGIT_MAP))]
        for *_, e in found:
            self.found[e["id"]] += 1
        return [h for h in found if h[2]["restore"]]

    def swap(self, s):
        for start, end, e in reversed(self.hits(s)):
            s = s[:start] + e["value"] + s[end:]
            self.restored[e["id"]] += 1
            self.dirty = True
        return s

    def xml(self, name, blob, loc):
        return restore_xml(name, blob, self, loc)

    def other(self, name, blob, loc):
        return blob


def restore_paragraph(p, rj, loc):
    text = "".join(s[3] for s in segments(p))
    for start, end, e in reversed(rj.hits(text)):
        inside, runs = isolate(p, start, end)
        if not runs:
            continue
        replace_span(p, inside, runs, e["value"])
        rj.restored[e["id"]] += 1
        rj.dirty = True
        if len(e["value"]) - len(e["id"]) >= LAYOUT_GROWTH:
            rj.layout.add(loc)


def restore_number(c, rj):
    #an Excel number that became text on redaction goes back to a number
    text = "".join(get_text(t) for t in c.getElementsByTagNameNS(NS_S, "t"))
    m = rj.rx.fullmatch(text.translate(DIGIT_MAP)) if rj.rx else None
    e = m and rj.entries[m.group().lower()]
    if not (e and e["numeric"] and e["restore"]):
        return
    for is_ in children(c, NS_S, "is"):
        c.removeChild(is_)
    c.removeAttribute("t")
    v = new_el(c, "v")
    set_text(v, e["value"])
    c.insertBefore(v, first_child(c, NS_S, "extLst"))
    rj.found[e["id"]] += 1
    rj.restored[e["id"]] += 1
    rj.dirty = True


def add_child(root, ns, local):
    prefix = xmlns_prefix(root, ns)
    el = new_el(root, local, ns, prefix)
    if prefix is None and root.namespaceURI != ns:
        el.setAttribute("xmlns", ns)
    root.appendChild(el)
    return el


def restore_field(f, root):
    t = f["target"]
    if t["kind"] == "property":
        el = next((e for e in iter_elements(root) if e.localName == t["el"] and e.namespaceURI == t["ns"]), None)
        set_text(el if el is not None else add_child(root, t["ns"], t["el"]), f["value"])
        return True
    for el in iter_elements(root):
        #only authors still carrying the placeholder, never one added after redaction
        if el.localName == t["el"] and el.getAttribute("id") == t["id"]:
            if el.getAttribute(t["attr"]) != AUTHOR_PLACEHOLDER[t["attr"]]:
                return False
            el.setAttribute(t["attr"], f["value"])
            return True
    return False


def restore_xml(name, data, rj, loc):
    root = parse_root(data)
    rj.dirty = False

    if name.endswith(".rels"):
        for rel in children(root):
            if rel.getAttribute("TargetMode") == "External":
                rel.setAttribute("Target", rj.swap(rel.getAttribute("Target")))
    else:
        #paragraphs first, so IDs split across differently formatted runs still match
        for p in root.getElementsByTagNameNS(NS_A, "p"):
            restore_paragraph(p, rj, loc)
        for el in list(iter_elements(root)):
            if is_el(el, NS_S, "c") and el.getAttribute("t") == "inlineStr":
                restore_number(el, rj)
        for el in list(iter_elements(root)):
            if el.localName == "cNvPr":
                for attr in ("descr", "title"):
                    if el.getAttribute(attr):
                        el.setAttribute(attr, rj.swap(el.getAttribute(attr)))
            if not is_el(el, NS_A, "t"): #a:t was handled with its paragraph
                for c in el.childNodes:
                    if isinstance(c, minidom.Text): #cdata sections are Text too
                        c.data = rj.swap(c.data)

        path = part_path(rj.pkg, name)
        for i, f in enumerate(rj.fields):
            if i not in rj.done and f["target"]["part"] == path and restore_field(f, root):
                rj.done.add(i)
                rj.dirty = True
        if path == CUSTOM_PROPS:
            for old in deck_id_props(root):
                root.removeChild(old)
                rj.dirty = True

    return serialize(root, data) if rj.dirty else data

#restore key (.xlsx, stdlib only)
XLSX_STYLES = (
    f'<styleSheet xmlns="{NS_S}">'
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="49" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>')


def col_name(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(ref):
    n = 0
    for ch in ref.rstrip("0123456789"):
        n = n * 26 + ord(ch) - 64
    return n - 1


def sheet_xml(widths, rows, yn):
    #every cell and column is text so excel keeps leading zeros and long numbers
    cols = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1" style="1"/>' for i, w in enumerate(widths, 1))
    body = "".join(
        f'<row r="{r}">' + "".join(
            f'<c r="{col_name(c)}{r}" t="inlineStr" s="{2 if r == 1 else 1}"><is><t xml:space="preserve">{escape(v)}</t></is></c>'
            for c, v in enumerate(row) if v) + "</row>"
        for r, row in enumerate(rows, 1))
    choice = "" if yn is None else (
        f'<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" '
        f'sqref="{col_name(yn)}2:{col_name(yn)}{max(len(rows), 2)}"><formula1>"Y,N"</formula1></dataValidation></dataValidations>')
    return (f'<worksheet xmlns="{NS_S}"><sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            f'<cols>{cols}</cols><sheetData>{body}</sheetData>{choice}</worksheet>')


def write_xlsx(path, sheets):
    #sheets: [(name, column widths, rows, Y/N column index or None)], first row is the header
    decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    ct = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    n = len(sheets)
    parts = {
        "[Content_Types].xml":
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{ct}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{ct}.styles+xml"/>'
            + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="{ct}.worksheet+xml"/>'
                      for i in range(1, n + 1)) + "</Types>",
        "_rels/.rels":
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml":
            f'<workbook xmlns="{NS_S}" xmlns:r="{NS_R}"><bookViews><workbookView/></bookViews><sheets>'
            + "".join(f'<sheet name="{escape(s[0], {chr(34): "&quot;"})}" sheetId="{i}" r:id="rId{i}"/>'
                      for i, s in enumerate(sheets, 1)) + "</sheets></workbook>",
        "xl/_rels/workbook.xml.rels":
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{i}" Type="{rel}/worksheet" Target="worksheets/sheet{i}.xml"/>'
                      for i in range(1, n + 1))
            + f'<Relationship Id="rId{n + 1}" Type="{rel}/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": XLSX_STYLES,
    }
    for i, (_, widths, rows, yn) in enumerate(sheets, 1):
        parts[f"xl/worksheets/sheet{i}.xml"] = sheet_xml(widths, rows, yn)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, xml in parts.items():
            z.writestr(name, (decl + xml).encode("utf-8"))


def read_xlsx(path):
    #{sheet name: rows as lists of strings}; also reads the file after Excel has re-saved it
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    shared = []
    if "xl/sharedStrings.xml" in names:
        for si in children(parse_root(z.read("xl/sharedStrings.xml")), NS_S, "si"):
            #plain text sits in si/t, rich text in si/r/t (si/rPh/t is phonetic, skipped)
            shared.append("".join(get_text(t) for part in [si, *children(si, NS_S, "r")]
                                  for t in children(part, NS_S, "t")))
    targets = {rid: full for rid, _, full in rels_of(z, names, "xl/workbook.xml")}
    sheets = {}
    for sh in parse_root(z.read("xl/workbook.xml")).getElementsByTagNameNS(NS_S, "sheet"):
        rows = []
        for row in parse_root(z.read(targets[sh.getAttributeNS(NS_R, "id")])).getElementsByTagNameNS(NS_S, "row"):
            cells, col = {}, -1
            for c in children(row, NS_S, "c"):
                col = col_index(c.getAttribute("r")) if c.getAttribute("r") else col + 1
                v = first_child(c, NS_S, "v")
                if c.getAttribute("t") == "inlineStr":
                    cells[col] = "".join(get_text(t) for t in c.getElementsByTagNameNS(NS_S, "t"))
                elif v is not None:
                    cells[col] = shared[int(get_text(v))] if c.getAttribute("t") == "s" else get_text(v)
            rows.append([cells.get(i, "") for i in range(max(cells, default=-1) + 1)])
        sheets[sh.getAttribute("name")] = rows
    return sheets


KEY_HEADER = ["ID", "Type", "Original value", "Locations", "Restore (Y/N)", "Restore details (do not edit)"]


def id_order(entry):
    m = re.match(r"(.*?)(\d+)$", entry["id"])
    return (m.group(1), int(m.group(2))) if m else (entry["id"], 0)


def write_key(path, job, src):
    rows = [KEY_HEADER]
    for e in sorted(job.entries.values(), key=id_order):
        rows.append([e["id"], e["type"], e["value"], "; ".join(sorted(e["locations"], key=slide_key)), "Y",
                     json.dumps({"kind": "number"}) if e["numeric"] else ""])
    for f in job.fields:
        rows.append([f["id"], f["type"], f["value"], "; ".join(f["locations"]), "Y", json.dumps(f["target"])])
    notes = [["Location", "Note"]] + [
        [loc, NOTE_TEXT.get(kind, kind)]
        for loc, kind in dict.fromkeys((r[0], r[1]) for r in sorted(job.rows, key=row_order) if not r[2])]
    about = [["Item", "Value"],
             ["Source deck", src.name],
             ["Redacted on", datetime.now().strftime("%d/%m/%Y %H:%M")],
             ["Deck ID", job.deck_id],
             ["Keep this file", "It holds the original values. Keep it on this machine and never send it with the deck."],
             ["To restore", f"python sanitize_pptx.py <returned deck>.pptx --restore {path.name}"],
             ["Restore only some items", "Set Restore (Y/N) to N for anything that should stay redacted."]]
    write_xlsx(path, [("Restore key", [16, 24, 36, 44, 14, 40], rows, 4),
                      ("Review notes", [30, 80], notes, None),
                      ("About", [24, 90], about, None)])


def read_key(path):
    sheets = read_xlsx(path)
    head, *rows = sheets["Restore key"]
    col = {h: i for i, h in enumerate(head)}
    entries, fields = [], []
    for r in rows:
        if not any(r):
            continue
        get = lambda h: r[col[h]] if col[h] < len(r) else ""
        details = get(KEY_HEADER[5]).strip()
        target = json.loads(details) if details else {}
        item = {"id": get("ID").strip(), "type": get("Type"), "value": get("Original value"),
                "restore": not get("Restore (Y/N)").strip().upper().startswith("N")}
        if target.get("kind") in ("property", "author"):
            fields.append({**item, "target": target})
        else:
            entries.append({**item, "numeric": target.get("kind") == "number"})
    about = {r[0]: r[1] for r in sheets.get("About", [])[1:] if len(r) > 1}
    return entries, fields, about.get("Deck ID")

#command line
def load_keywords(path):
    #"term" or "term | Label"; the label names the term's id
    terms, labels = [], {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        term, _, label = line.partition("|")
        term, label = term.strip(), re.sub(r"\s+", "", label)
        if not term:
            continue
        if label[-1:].isdigit():
            raise SystemExit(f"Label '{label}' for '{term}' must not end with a digit (IDs add their own number)")
        terms.append(term)
        if label:
            labels[term.translate(DIGIT_MAP).lower()] = label
    return terms, labels


def redact(args):
    keywords, labels = load_keywords(args.keywords) if args.keywords else ([], {})
    src = Path(args.pptx)
    job = Job(keywords, labels, args.preview)
    result = process_package(src.read_bytes(), job)

    suffix = "_preview" if args.preview else "_redacted"
    out = src.with_name(src.stem + suffix + src.suffix)
    out.write_bytes(result)
    found = sum(1 for r in job.rows if r[2])

    if args.preview:
        rep = src.with_name(src.stem + suffix + "_report.csv")
        with open(rep, "w", newline="", encoding="utf-8-sig") as f: #utf-8-sig opens cleanly in excel
            w = csv.writer(f)
            w.writerow(["location", "type", "masked_value"])
            w.writerows(sorted(job.rows, key=row_order))
        print(f"{found} items found, {len(job.rows) - found} review notes -> {out}\nReport -> {rep}")
        return

    key_dir = Path(args.key_dir) if args.key_dir else src.parent
    key_dir.mkdir(parents=True, exist_ok=True)
    key = key_dir / (src.stem + "_restore_key.xlsx")
    write_key(key, job, src)
    print(f"{found} items redacted as {len(job.entries)} IDs, {len(job.rows) - found} review notes -> {out}")
    print(f"Restore key -> {key}\nKeep the key on this machine. Send only the redacted deck.")

    #self-check: every id must be found again exactly as often as it was placed
    check = Restore([{**e, "restore": True} for e in job.entries.values()])
    process_package(result, check)
    off = [(i, n, check.found[i]) for i, n in job.placed.items() if check.found[i] != n]
    if off:
        print("WARNING: these ids will not restore cleanly:")
        for i, placed, seen in off:
            why = "the deck already contains this text" if seen > placed else "it touches letters or digits next to it"
            print(f"  {i}: placed {placed}, found {seen} ({why})")


def restore(args):
    src = Path(args.pptx)
    data = src.read_bytes()
    entries, fields, key_id = read_key(Path(args.restore))
    deck_id = read_deck_id(data)
    if deck_id and key_id and deck_id != key_id and not args.force:
        raise SystemExit("This key was made for a different deck (deck ids differ). "
                         "Use the matching key, or add --force to restore anyway.")
    if not deck_id:
        print("Note: this deck has no deck id (it may have been rebuilt in a new file), "
              "so the key could not be matched to it. Make sure it is the right key.")

    rj = Restore(entries, fields)
    result = process_package(data, rj)
    out = src.with_name(src.stem + "_restored" + src.suffix)
    out.write_bytes(result)

    wanted = [e["id"] for e in entries if e["restore"]]
    missing = [i for i in wanted if not rj.restored[i]]
    kept = [e["id"] for e in entries if not e["restore"]]
    print(f"Restored {len(wanted) - len(missing)} of {len(wanted)} IDs "
          f"({sum(rj.restored.values())} places) -> {out}")
    if wanted and len(missing) == len(wanted):
        print("WARNING: none of the key's IDs were found. Is this the right deck and key?")
    elif missing:
        print("Not found, restore by hand (see Locations in the key): " + ", ".join(missing))
    if kept:
        print("Kept redacted (Restore = N): " + ", ".join(kept))
    if rj.fields:
        left = [f for i, f in enumerate(rj.fields) if i not in rj.done]
        print(f"File properties and comment authors restored: {len(rj.fields) - len(left)} of {len(rj.fields)}")
        for f in left:
            print(f"  not restored: {f['id']} {f['type']}")
    if rj.layout:
        print("Check layout, restored text is longer than its ID: " + ", ".join(sorted(rj.layout, key=slide_key)))


def main():
    ap = argparse.ArgumentParser(description="Rule-based PII redaction for .pptx (no dependencies)")
    ap.add_argument("pptx")
    ap.add_argument("--preview", action="store_true", help="highlight only, no removal")
    ap.add_argument("--keywords", help="text file, one term per line; 'term | Label' names its ID")
    ap.add_argument("--key-dir", help="folder for the restore key (default: next to the deck)")
    ap.add_argument("--restore", metavar="KEY", help="put the original values back using this restore key (.xlsx)")
    ap.add_argument("--force", action="store_true", help="restore even if the key was made for a different deck")
    args = ap.parse_args()
    if args.restore:
        restore(args)
    else:
        redact(args)


if __name__ == "__main__":
    main()
