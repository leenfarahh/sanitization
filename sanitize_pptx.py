
#python sanitize_pptx.py deck.pptx --preview (yellow highlights + CSV report)
#python sanitize_pptx.py deck.pptx (each confidential value becomes [X] + restore key .xlsx that stays with the client)
#python sanitize_pptx.py deck.pptx --keywords keywords.txt (only the terms in keywords.txt become [X], emails/phones/IDs/IBANs/cards are left as they are)
#python sanitize_pptx.py deck.pptx --all (all text except slide titles becomes [X], one per paragraph, line and formatting change; chart labels [X]; restorable with the key)
#python sanitize_pptx.py returned_deck.pptx --restore deck_restore_key.xlsx (put the original values back)

import argparse
import base64
import csv
import io
import json
import posixpath
import random
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
def find_matches(text: str, keywords=None, rules=COMPILED):
    #return list of (start, end, rule_name), first/longest wins when there's overlap
    norm = text.translate(DIGIT_MAP)
    lower = norm.lower()
    hits = []

    for name, rx, validator, context in rules:
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
    #preview report: [X] in place of all but the last 4 characters
    v = value.strip()
    return (X_MASK if len(v) > 4 else "") + v[-4:]

#masking: everything becomes [X]
def mask_char(ch):
    return "⁣" if ch in (MARK_EDGE, MARK_0, MARK_1) else ch #never let the original look like a marker


def bracket_pairs(s):
    #(original, shown) pieces of a text, or of one line or formatting group of a paragraph: [X] for everything
    #between the spaces around it (kept, so groups don't run together); a piece with no letters or digits stays
    if not any(ch.isalnum() for ch in s):
        return [(ch, mask_char(ch)) for ch in s]
    lead, trail = len(s) - len(s.lstrip()), len(s.rstrip())
    return [(p, p) for p in (s[:lead],) if p] + [(s[lead:trail], X_MASK)] + [(p, p) for p in (s[trail:],) if p]


def bracket_text(s):
    return "".join(shown for _, shown in bracket_pairs(s))


def marker(n):
    bits = f"{n:016b}{n % 13:04b}"
    return MARK_EDGE + "".join(MARK_1 if b == "1" else MARK_0 for b in bits) + MARK_EDGE


def read_marker(code):
    #restore id from a marker, 0 (never used) when the check bits don't match (marker damaged by an edit)
    bits = "".join("1" if c == MARK_1 else "0" for c in code)
    n = int(bits[:16], 2)
    return n if int(bits[16:], 2) == n % 13 else 0

#xml helpers (stdlib minidom)

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

RPR_AFTER_HIGHLIGHT = {"uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
                       "hlinkClick", "hlinkMouseOver", "rtl", "extLst"}
X_MASK = "[X]" #the one masking string, in every mode
CLEAR_META = {"creator", "lastModifiedBy", "Company", "Manager"} #become [X]
AUTHOR_PLACEHOLDER = {"name": X_MASK, "initials": X_MASK, "userId": ""}
EMBEDDED_OOXML = (".xlsx", ".xlsm", ".docx", ".pptx")

#the redacted deck carries a random id in its custom properties, the key carries the same one
CUSTOM_PROPS = "docProps/custom.xml"
DECK_ID_PROP = "SanitizationID"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
LAYOUT_GROWTH = 5 #restored text this many characters longer than what was on the slide is flagged for a layout check

#detected values become a single [X] whatever their length (+971 50 123 4567 -> [X]), so a slide full of values stays readable.
#--all mode: each paragraph becomes [X], one more for each line break and formatting change (a bold lead-in), so every
#piece goes back with its own formatting; chart labels [X] too. the key records where each piece starts and ends.
#each masked value or text ends with an invisible marker (zero-width characters) holding its restore id: 16-bit id + 4-bit check
TITLE_PH = {"title", "ctrTitle"} #slide titles stay readable
CORE_TEXT = {"title", "subject", "keywords", "description", "category", "contentStatus"}
MARK_EDGE, MARK_0, MARK_1 = "\u2060", "\u200b", "\u200c" #word joiner, zero-width space, zero-width non-joiner
MARK_RX = re.compile(f"{MARK_EDGE}([{MARK_0}{MARK_1}]{{20}}){MARK_EDGE}")
#run attributes PowerPoint splits runs over (spelling, language, editing) that don't show, so they don't start a new [X]
RPR_UNSEEN = ("lang", "altLang", "dirty", "err", "noProof", "smtClean", "smtId", "bmk")
#web links can't hold an invisible marker so they point to a placeholder on the reserved .invalid domain
LINK_HOST = "https://masked.invalid/"
LINK_RX = re.compile(r"https?://masked\.invalid/(Link\d+)/?", re.I)
KIND_ORDER = {"x": 0, "id": 0, "mask": 1, "placeholder": 2, "link": 3}

#backup when a marker is lost: the key records where each [X] sat (slide, item number top to bottom),
#and restore suggests unmarked [X]s there, for a person to confirm
SLIDE_RX = re.compile(r"ppt/slides/slide\d+\.xml$")
SHAPES = {"sp", "graphicFrame", "cxnSp", "pic"}
PH_FAMILY = {"ctrTitle": "title", "subTitle": "body", "obj": "body"}
ROW_BAND = 91440 #0.1 inch: shapes this close vertically count as one row, read left to right
XRUN_RX = re.compile(r"\[X\]")

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
    def __init__(self, keywords, labels, preview, full=False):
        self.keywords, self.labels, self.preview, self.full = keywords, labels, preview, full
        self.rules = [] if keywords else COMPILED #a keyword list replaces the built-in rules: only its terms are redacted
        self.rows = []
        self.ids, self.entries, self.fields = {}, {}, []
        self.marks, self.numbers, self.links = {}, {}, {} 
        self.counts, self.placed = Counter(), Counter()
        self.authors, self.last_mark = 0, 0
        self.pkg = [] #embedded file being processed, if any
        self.deck_id = None if preview else secrets.token_hex(6)
        self.rng = random.SystemRandom()
        self.spots, self.ph_pos = None, {} #X values placed on the current slide, where its placeholders sit

    def log(self, loc, kind, value=""):
        self.rows.append((loc, kind, mask(value) if value else ""))

    def place(self, tid, loc):
        entry = self.entries[tid]
        if loc not in entry["locations"]:
            entry["locations"].append(loc)
        self.placed[tid] += 1
        return tid

    def new_mark(self):
        #marker numbers are shared by X values and --all texts
        self.last_mark += 1
        if self.last_mark >= 1 << 16:
            raise SystemExit("Too many different values and texts to mask in one deck (65,535 max)")
        return self.last_mark

    def xrun(self, rule, value, loc, numeric=False):
        #detected value -> one [X] + invisible marker, whatever its length; same value, same marker.
        #the key names it by type (PhoneUAE1, Email2...), the deck shows only [X]
        prefix = ID_PREFIX.get(rule) or self.labels.get(value.translate(DIGIT_MAP).lower(), "Keyword")
        tid = self.ids.get((prefix, value))
        if tid is None:
            self.counts[prefix] += 1
            tid = self.ids[prefix, value] = f"{prefix}{self.counts[prefix]}"
            self.entries[tid] = {"id": tid, "type": rule, "value": value, "locations": [], "kind": "x",
                                 "mark": self.new_mark(), "numeric": False, "slots": [], "style": "bracket"}
        entry = self.entries[tid]
        entry["numeric"] |= numeric
        self.place(tid, loc)
        return tid, X_MASK + marker(entry["mark"])

    def mark(self, text, loc, label=False, parts=()):
        #--all: invisible marker for a masked text, same text in the same pieces same marker.
        #parts: length of each piece shown as its own [X] (lines, formatting changes), none when the text is one piece
        parts = tuple(parts) if len(parts) > 1 else ()
        tid = self.marks.get((label, text, parts))
        if tid is None:
            n = self.new_mark()
            tid = self.marks[label, text, parts] = f"Text{n}"
            self.entries[tid] = {"id": tid, "type": "CHART_LABEL" if label else "TEXT", "value": text,
                                 "locations": [], "kind": "mask", "mark": n, "style": "bracket", "parts": list(parts)}
        self.place(tid, loc)
        return marker(self.entries[tid]["mark"])

    def mask_string(self, s, loc, label=False):
        if not any(ch.isalnum() for ch in s):
            return s
        return bracket_text(s) + self.mark(s, loc, label)

    def string(self, s, loc, label=False):
        #text outside paragraphs: fully masked in --all mode ([X], chart labels too), pattern redaction otherwise
        return self.mask_string(s, loc, label) if self.full else self.redact_string(s, loc)

    def number(self, value, loc):
        #--all: chart and excel numbers become random placeholders, same number same placeholder.
        ph = self.numbers.get(value)
        if ph is None:
            lo, hi = (101, 999) if len(self.numbers) < 400 else (1001, 9999)
            ph = str(self.rng.randint(lo, hi))
            while ph.endswith("0") or ph in self.entries:
                ph = str(self.rng.randint(lo, hi))
            self.numbers[value] = ph
            self.entries[ph] = {"id": ph, "type": "CHART_NUMBER", "value": value, "locations": [], "kind": "placeholder"}
        return self.place(ph, loc)

    def link(self, target, loc):
        #web links (--all) and email/phone links point to a placeholder address instead, same address same placeholder
        tid = self.links.get(target)
        if tid is None:
            tid = self.links[target] = f"Link{len(self.links) + 1}"
            self.entries[tid] = {"id": tid, "type": "LINK", "value": target, "locations": [], "kind": "link"}
        return LINK_HOST + self.place(tid, loc)

    def keep(self, label, type_, value, loc, **target):
        #originals with no text to hold an id (file properties, comment authors)
        target["part"] = part_path(self.pkg, target["part"])
        self.fields.append({"id": label, "type": type_, "value": value, "locations": [loc], "target": target})

    def redact_string(self, s, loc, numeric=False):
        #plain-string redaction (attributes, chart caches, cells).
        if not s:
            return s
        matches = find_matches(s, self.keywords, self.rules)
        for start, end, rule in matches:
            self.log(loc, rule, s[start:end])
        if self.preview:
            return s
        tokens = [self.xrun(rule, s[start:end], loc, numeric)[1] for start, end, rule in matches]
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
    matches = find_matches(text, job.keywords, job.rules)
    for start, end, rule in matches:
        job.log(loc, rule, text[start:end])
    #markers are assigned left to right, then spliced right to left so offsets stay valid
    tokens = [None if job.preview else job.xrun(rule, text[start:end], loc) for start, end, rule in matches]
    for (start, end, _), tok in reversed(list(zip(matches, tokens))):
        inside, runs = isolate(p, start, end)
        if not runs:
            continue
        if tok is None:
            for r in runs:
                if r.localName == "r":
                    add_highlight(r)
        else:
            replace_span(p, inside, runs, tok[1])
            if job.spots is not None:
                job.spots.append((p, start, tok[0]))
    return bool(matches)

#xml parts
def has_ancestor(el, ns, local):
    node = el.parentNode
    while node is not None and node.nodeType == node.ELEMENT_NODE:
        if is_el(node, ns, local):
            return True
        node = node.parentNode
    return False


def chart_text(el):
    #chart strings: cached labels and series names, literal strings, plain series names
    return (has_ancestor(el, NS_C, "strCache") or has_ancestor(el, NS_C, "strLit")
            or is_el(el.parentNode, NS_C, "tx"))


def chart_number(el):
    return has_ancestor(el, NS_C, "numCache") or has_ancestor(el, NS_C, "numLit")


def title_paragraphs(root):
    #paragraphs inside slide title placeholders, which stay readable in --all mode
    keep = set()
    for sp in root.getElementsByTagNameNS(NS_P, "sp"):
        ph = sp.getElementsByTagNameNS(NS_P, "ph")
        if ph and ph[0].getAttribute("type") in TITLE_PH:
            keep.update(sp.getElementsByTagNameNS(NS_A, "p"))
    return keep

#position on the slide (top to bottom, then left to right), for the backup match
def xfrm_of(shape):
    #a:xfrm of a shape (p:spPr), a group (p:grpSpPr) or a table/chart frame (p:xfrm)
    for c in children(shape, NS_P):
        if c.localName in ("spPr", "grpSpPr"):
            return first_child(c, NS_A, "xfrm")
        if c.localName == "xfrm":
            return c
    return None


def xy(xfrm, local):
    #(x, y) of a:off / a:chOff, (cx, cy) of a:ext / a:chExt
    el = first_child(xfrm, NS_A, local) if xfrm is not None else None
    if el is None:
        return None
    a, b = ("cx", "cy") if local.endswith(("ext", "Ext")) else ("x", "y")
    return int(el.getAttribute(a) or 0), int(el.getAttribute(b) or 0)


def placeholder_positions(z, names, slide):
    #where the placeholders of a slide's layout (and master) sit, for placeholders that inherit their position
    parts = [full for _, typ, full in rels_of(z, names, slide) if typ == "slideLayout"][:1]
    if parts:
        parts += [full for _, typ, full in rels_of(z, names, parts[0]) if typ == "slideMaster"][:1]
    pos = {}
    for part in reversed(parts): #master first, the layout overrides it
        for sp in parse_root(z.read(part)).getElementsByTagNameNS(NS_P, "sp"):
            ph, off = sp.getElementsByTagNameNS(NS_P, "ph"), xy(xfrm_of(sp), "off")
            if ph and off:
                if ph[0].getAttribute("idx") and part == parts[0]:
                    pos["idx", ph[0].getAttribute("idx")] = off
                typ = ph[0].getAttribute("type") or "obj"
                pos["type", PH_FAMILY.get(typ, typ)] = off
    return pos


def reading_order(p, ph_pos):
    #(row, x) of the shape holding paragraph p, mapped out of any groups
    shape = p.parentNode
    while is_el(shape) and not (shape.namespaceURI == NS_P and shape.localName in SHAPES):
        shape = shape.parentNode
    if not is_el(shape):
        return (0, 0)
    off = xy(xfrm_of(shape), "off")
    ph = shape.getElementsByTagNameNS(NS_P, "ph")
    if off is None and ph:
        idx, typ = ph[0].getAttribute("idx"), ph[0].getAttribute("type") or "obj"
        off = (idx and ph_pos.get(("idx", idx))) or ph_pos.get(("type", PH_FAMILY.get(typ, typ)))
    x, y = off or (0, 0)
    grp = shape.parentNode
    while is_el(grp, NS_P, "grpSp"):
        g = xfrm_of(grp)
        (gx, gy), (ox, oy) = xy(g, "off") or (0, 0), xy(g, "chOff") or (0, 0)
        (gw, gh), (cw, ch) = xy(g, "ext") or (1, 1), xy(g, "chExt") or (1, 1)
        x, y = gx + (x - ox) * gw // (cw or 1), gy + (y - oy) * gh // (ch or 1)
        grp = grp.parentNode
    return (round(y / ROW_BAND), x)


def slide_no(loc):
    m = re.fullmatch(r"Slide (\d+)", loc or "")
    return int(m.group(1)) if m else None


def shown_format(run):
    #a run's formatting as it shows on the slide, to tell where one [X] ends and the next begins
    rpr = first_child(run, NS_A, "rPr")
    if rpr is None:
        return ""
    return " ".join([f'{k}="{v}"' for k, v in sorted(rpr.attributes.items()) if k not in RPR_UNSEEN]
                    + [c.toxml() for c in children(rpr)])


def mask_paragraph(p, job, loc):
    #--all: a paragraph becomes [X], one per line and per formatting change, each in the first run of its piece.
    #an invisible marker right after the last [X] carries the restore id
    pieces, look = [], None
    for el in children(p, NS_A):
        t = first_child(el, NS_A, "t") if el.localName == "r" else None
        if el.localName == "br":
            look = None #a line break starts a new [X]
        elif t is not None:
            if shown_format(el) != look:
                pieces.append([])
                look = shown_format(el)
            pieces[-1].append(t)
    text = "".join(get_text(t) for piece in pieces for t in piece)
    if not any(ch.isalnum() for ch in text):
        return False
    parts = []
    for piece in pieces:
        s = "".join(get_text(t) for t in piece)
        for t in piece[1:]:
            set_text(t, "")
        set_text(piece[0], bracket_text(s))
        if s:
            parts.append(len(s))
    last = pieces[-1][0]
    set_text(last, get_text(last) + job.mark(text, loc, parts=parts))
    return True


def process_xml(name, data, job, loc):
    root = parse_root(data)
    before = len(job.rows) + sum(job.placed.values())
    changed = False

    if name.endswith(".rels"):
        for rel in children(root):
            t = rel.getAttribute("Target")
            if rel.getAttribute("TargetMode") != "External" or not t:
                continue
            here = loc + " (hyperlink)"
            if t.lower().startswith(("mailto:", "tel:")): #a link address can't hold a marker, so it gets a placeholder
                matches = find_matches(t, job.keywords, job.rules)
                for start, end, rule in matches:
                    job.log(here, rule, t[start:end])
                if job.full or matches and not job.preview:
                    rel.setAttribute("Target", job.link(t, here))
            elif job.full and rel.getAttribute("Type").endswith("/hyperlink"): #web and file links
                rel.setAttribute("Target", job.link(t, here))
    else:
        titles = title_paragraphs(root) if job.full else set()
        slide = bool(SLIDE_RX.match(name)) and not job.pkg and slide_no(loc) is not None
        paras = root.getElementsByTagNameNS(NS_A, "p")
        job.spots = [] if slide else None
        for p in paras:
            if job.full and p not in titles:
                changed |= mask_paragraph(p, job, loc)
            else:
                changed |= process_paragraph(p, job, loc)
        if job.spots:
            #number the X values on this slide top to bottom, for the backup match
            index = {p: i for i, p in enumerate(paras)}
            ranked = sorted(job.spots, key=lambda s: (reading_order(s[0], job.ph_pos), index[s[0]], s[1]))
            for k, (_, _, tid) in enumerate(ranked, 1):
                job.entries[tid]["slots"].append([slide_no(loc), k])
        job.spots = None

        for el in list(iter_elements(root)):
            ln, ns = el.localName, el.namespaceURI
            if ln == "cNvPr": #alt text
                for attr in ("descr", "title"):
                    if el.getAttribute(attr):
                        el.setAttribute(attr, job.string(el.getAttribute(attr), loc + " (alt text)"))
            elif ns == NS_C and ln == "v" and chart_text(el):
                set_text(el, job.string(get_text(el), loc + " (chart labels)", label=True))
            elif ns == NS_C and ln == "v" and job.full and get_text(el).strip() and chart_number(el):
                set_text(el, job.number(get_text(el), loc + " (chart data)"))
            elif ns == NS_P and ln == "text": #legacy comments
                set_text(el, job.string(get_text(el), loc + " (comment)"))
            elif ns == NS_S and ln == "t": #excel strings, masked like chart labels so chart data and chart agree
                set_text(el, job.string(get_text(el), loc, label=True))
            elif ns == NS_S and ln == "c" and el.getAttribute("t") in ("", "n") and job.full: #every Excel number
                v = first_child(el, NS_S, "v")
                if v is not None and get_text(v):
                    set_text(v, job.number(get_text(v), loc))
            elif ns == NS_P14 and ln == "section" and job.full and el.getAttribute("name"): #slide section names
                el.setAttribute("name", job.mask_string(el.getAttribute("name"), "Sections"))
            elif ln in ("hlinkClick", "hlinkMouseOver") and job.full and el.getAttribute("tooltip"): #link screentips
                el.setAttribute("tooltip", job.mask_string(el.getAttribute("tooltip"), loc + " (link screentip)"))
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
                            set_text(el, X_MASK)
                    elif job.full and (name.endswith("core.xml") and el.localName in CORE_TEXT
                                       or name.endswith("custom.xml") and is_el(el, NS_VT) and el.localName in ("lpwstr", "bstr")):
                        set_text(el, job.mask_string(get_text(el), "File properties"))
                    else:
                        set_text(el, job.redact_string(get_text(el), "File properties"))

        if SLIDE_RX.match(name):
            if root.getAttribute("show") == "0":
                job.log(loc, "HIDDEN_SLIDE")
            if root.getElementsByTagNameNS(NS_A, "blip"):
                job.log(loc, "IMAGE_REVIEW_MANUALLY")

    if changed or len(job.rows) + sum(job.placed.values()) > before:
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
            f'<property fmtid="{FMTID}" pid="2" name="{DECK_ID_PROP}"><vt:lpwstr>{escape(deck_id)}</vt:lpwstr></property>'
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
    #the deck carries its id, an embedded file its name at redaction (PowerPoint renames them on save)
    stamp = job.deck_id if parent_loc is None else job.deck_id and job.pkg[-1]

    def loc_of(name):
        owner = re.sub(r"_rels/(.+)\.rels$", r"\1", name) # a .rels belongs to its part
        return parent_loc or locs.get(owner, name)

    #process in slide order so IDs number from slide 1, write back in the original order
    done = {}
    files = [i.filename for i in zin.infolist() if not i.is_dir()]
    for name in sorted(files, key=lambda n: slide_key(loc_of(n))):
        blob, loc = zin.read(name), loc_of(name)
        if name.endswith((".xml", ".rels")) and name != "[Content_Types].xml":
            slide = parent_loc is None and SLIDE_RX.match(name)
            job.ph_pos = placeholder_positions(zin, names, name) if slide else {}
            blob = job.xml(name, blob, loc)
        elif name.lower().endswith(EMBEDDED_OOXML):
            job.pkg.append(read_deck_id(blob) or name)
            blob = process_package(blob, job, loc + " (embedded data)")
            job.pkg.pop()
        else:
            blob = job.other(name, blob, loc)
        done[name] = stamp_deck_id(name, blob, stamp, names) if stamp else blob

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, b"" if info.is_dir() else done[info.filename]) #folder entries carry no data
        if stamp and CUSTOM_PROPS not in names:
            zout.writestr(CUSTOM_PROPS, custom_props_xml(stamp))
    return out.getvalue()

#restore
class Restore:
    def __init__(self, entries, fields=()):
        kind = lambda k: [e for e in entries if e.get("kind", "id") == k]
        self.entries = {e["id"].lower(): e for e in kind("id")} #keys from before markers: IDs like PhoneUAE1 in the deck
        self.marks = {e["mark"]: e for e in kind("x") + kind("mask")}
        self.nums = {e["id"]: e for e in kind("placeholder")}
        self.links = {e["id"].lower(): e for e in kind("link")}
        ids = sorted(self.entries, key=len, reverse=True) #longest first, so PhoneUAE12 wins over PhoneUAE1
        self.rx = re.compile("(?:" + "|".join(map(re.escape, ids)) + r")(?![0-9])", re.I) if ids else None
        self.fields = [f for f in fields if f["restore"]]
        self.done = set() #indices of fields put back
        self.found, self.restored = Counter(), Counter()
        self.layout, self.edited, self.older = set(), set(), set()
        self.xruns = [] #[X]s seen on the slides, for the backup match
        self.pkg, self.deck_id, self.dirty, self.ph_pos = [], None, False, {}

    def plan(self, text, loc):
        #per-character replacements for text holding markers, keeping the designer's formatting
        new, start = list(text), 0
        for m in MARK_RX.finditer(text):
            seg, start = start, m.end()
            e = self.marks.get(read_marker(m.group(1)))
            if e is None:
                continue
            self.found[e["id"]] += 1
            if not e["restore"]:
                continue
            new[m.start():m.end()] = [""] * (m.end() - m.start())
            masked, before = masked_form(e), text[seg:m.start()]
            if before.lower().endswith(masked.lower()): #anything the designer typed before it is kept
                #each original piece goes where its [X] sits, so it takes that piece's formatting
                pos = m.start() - len(masked)
                for orig, shown in mask_pairs(e):
                    new[pos:pos + len(shown)] = [orig] + [""] * (len(shown) - 1)
                    pos += len(shown)
                #[X] is much shorter than what it hides; chart labels and masters/layouts (placeholder prompts) unchecked
                if (e["type"] != "CHART_LABEL" and slide_key(loc)[0] == 0
                        and len(e["value"]) - len(masked) >= LAYOUT_GROWTH):
                    self.layout.add(loc)
            else:
                #edited, or masked by an older version of this tool: a detected value loses what is left of its [X]
                #(or X's) just before its marker, an --all text everything since the last marker
                rest = re.sub(r"\[?[Xx]*\]?\Z", "", before, count=1)
                begin = m.start() - (len(before) - len(rest)) if e["kind"] == "x" else seg
                new[begin:m.start()] = [""] * (m.start() - begin)
                new[m.start()] = e["value"]
                (self.edited if e.get("style") == "bracket" else self.older).add(loc)
            self.restored[e["id"]] += 1
            self.dirty = True
        return new

    def unmask(self, s, loc):
        return "".join(self.plan(s, loc)) if MARK_EDGE in s else s

    def target(self, t):
        #link targets: a --all placeholder goes back whole, IDs inside other targets (mailto:Email1) are swapped
        m = LINK_RX.fullmatch(t.strip())
        e = self.links.get(m.group(1).lower()) if m else None
        if e is None:
            return self.swap(t)
        self.found[e["id"]] += 1
        if not e["restore"]:
            return t
        self.restored[e["id"]] += 1
        self.dirty = True
        return e["value"]

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


def mask_pairs(e):
    #(original, shown) pieces of a marked entry as it looked in the redacted deck, before its marker:
    #a detected value is one [X], an --all text one [X] per piece (line, formatting change)
    if e["kind"] == "x":
        return [(e["value"], X_MASK)]
    pairs, pos = [], 0
    for n in e.get("parts") or [len(e["value"])]:
        pairs += bracket_pairs(e["value"][pos:pos + n])
        pos += n
    return pairs


def masked_form(e):
    return "".join(shown for _, shown in mask_pairs(e))


def collect_xruns(root, rj, loc):
    #every [X] on a slide, marked or not, with its reading position; [X]s inside --all text are skipped
    for i, p in enumerate(root.getElementsByTagNameNS(NS_A, "p")):
        text = "".join(s[3] for s in segments(p))
        if X_MASK not in text:
            continue
        owners = [rj.marks.get(read_marker(code)) for code in MARK_RX.findall(text)]
        in_mask = any(e is not None and e["kind"] == "mask" for e in owners)
        order = reading_order(p, rj.ph_pos)
        for m in XRUN_RX.finditer(text):
            mk = MARK_RX.match(text, m.end())
            e = rj.marks.get(read_marker(mk.group(1))) if mk else None
            if e is not None and e["kind"] == "mask" or e is None and in_mask:
                continue
            rj.xruns.append({"loc": loc, "pos": (order, i, m.start()), "entry": e})


def suggestions(rj, e):
    #unmarked [X]s for a lost value: same slide and item first, then same slide, then elsewhere,
    #nearest to where it was first (slides added or removed shift the numbers)
    slots = {tuple(s) for s in e.get("slots", [])}
    slides = {s[0] for s in slots}
    ranked = []
    for r in rj.xruns:
        if r["entry"] is None:
            n, k = slide_no(r["loc"]), r["item"]
            near = min(((abs(n - s), abs(k - i)) for s, i in slots), default=(0, 0))
            ranked.append((0 if (n, k) in slots else 1 if n in slides else 2, near, n, k))
    return [(rank, n, k) for rank, _, n, k in sorted(ranked)[:3]]


def restore_masked_paragraph(p, rj, loc):
    ts = [first_child(r, NS_A, "t") for r in children(p, NS_A, "r")]
    ts = [t for t in ts if t is not None]
    text = "".join(get_text(t) for t in ts)
    if MARK_EDGE not in text:
        return
    new, pos = rj.plan(text, loc), 0
    for t in ts: #runs may have been split or restyled by the designer, each keeps its own formatting
        n = len(get_text(t))
        set_text(t, "".join(new[pos:pos + n]))
        pos += n


def restore_placeholder(v, rj):
    #--all: chart and Excel placeholders go back to the original numbers
    e = rj.nums.get(get_text(v).strip()) if v is not None else None
    if e is None:
        return
    rj.found[e["id"]] += 1
    if e["restore"]:
        set_text(v, e["value"])
        rj.restored[e["id"]] += 1
        rj.dirty = True


def restore_number(c, rj):
    #an excel number that became text on redaction goes back to a number, when its [X] (X's, older keys) is untouched
    text = "".join(get_text(t) for t in c.getElementsByTagNameNS(NS_S, "t"))
    m = MARK_RX.search(text)
    if m:
        e = rj.marks.get(read_marker(m.group(1)))
        if not (e and m.end() == len(text) and re.fullmatch(r"\[?X*\]?", text[:m.start()].upper())):
            return
    else: #keys from before markers
        m = rj.rx.fullmatch(text.translate(DIGIT_MAP)) if rj.rx else None
        e = m and rj.entries[m.group().lower()]
    if not (e and e.get("numeric") and e["restore"]):
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
                rel.setAttribute("Target", rj.target(rel.getAttribute("Target")))
    else:
        if SLIDE_RX.match(name) and not rj.pkg and slide_no(loc) is not None:
            collect_xruns(root, rj, loc) #before anything is restored
        #paragraphs first, so values split across differently formatted runs still match
        for p in root.getElementsByTagNameNS(NS_A, "p"):
            restore_paragraph(p, rj, loc)
            restore_masked_paragraph(p, rj, loc)
        for el in list(iter_elements(root)):
            if is_el(el, NS_S, "c") and el.getAttribute("t") == "inlineStr":
                restore_number(el, rj)
            elif is_el(el, NS_S, "c") and el.getAttribute("t") in ("", "n"):
                restore_placeholder(first_child(el, NS_S, "v"), rj)
            elif is_el(el, NS_C, "v") and chart_number(el):
                restore_placeholder(el, rj)
        for el in list(iter_elements(root)):
            if el.localName == "cNvPr":
                for attr in ("descr", "title"):
                    if el.getAttribute(attr):
                        el.setAttribute(attr, rj.unmask(rj.swap(el.getAttribute(attr)), loc))
            elif is_el(el, NS_P14, "section") and el.getAttribute("name"):
                el.setAttribute("name", rj.unmask(el.getAttribute("name"), loc))
            elif el.localName in ("hlinkClick", "hlinkMouseOver") and el.getAttribute("tooltip"):
                el.setAttribute("tooltip", rj.unmask(el.getAttribute("tooltip"), loc))
            if not is_el(el, NS_A, "t"): #a:t was handled with its paragraph
                for c in el.childNodes:
                    if isinstance(c, minidom.Text): #cdata sections are Text too
                        c.data = rj.unmask(rj.swap(c.data), loc)

        path = part_path(rj.pkg, name)
        for i, f in enumerate(rj.fields):
            if i not in rj.done and f["target"]["part"] == path and restore_field(f, root):
                rj.done.add(i)
                rj.dirty = True
        if name == CUSTOM_PROPS: #the deck's and embedded files' stamps
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


def entry_details(e, count):
    #marker number, places redacted, number flag and slide positions (slide, item) as JSON, for restore
    d = {"kind": e["kind"]}
    if "mark" in e:
        d["mark"], d["count"] = e["mark"], count
    if e.get("style"):
        d["style"] = e["style"]
    if e.get("parts"):
        d["parts"] = e["parts"]
    if e.get("numeric"):
        d["numeric"] = True
    if e.get("slots"):
        d["slots"] = e["slots"]
    return json.dumps(d)


def where(e):
    #Locations column: "Slide 2 (item 3); Slide 5 notes", items numbered top to bottom on each slide
    items = {}
    for n, k in e.get("slots", []):
        items.setdefault(n, []).append(str(k))
    out = []
    for loc in sorted(e["locations"], key=slide_key):
        ks = items.get(slide_no(loc))
        out.append(f"{loc} (item{'s' if len(ks) > 1 else ''} {', '.join(ks)})" if ks else loc)
    return "; ".join(out)


def write_key(path, job, src):
    rows = [KEY_HEADER]
    for e in sorted(job.entries.values(), key=lambda e: (KIND_ORDER[e["kind"]], id_order(e))):
        rows.append([e["id"], e["type"], e["value"], where(e), "Y", entry_details(e, job.placed[e["id"]])])
    for f in job.fields:
        rows.append([f["id"], f["type"], f["value"], "; ".join(f["locations"]), "Y", json.dumps(f["target"])])
    notes = [["Location", "Note"]] + [
        [loc, NOTE_TEXT.get(kind, kind)]
        for loc, kind in dict.fromkeys((r[0], r[1]) for r in sorted(job.rows, key=row_order) if not r[2])]
    about = [["Item", "Value"],
             ["Source deck", src.name],
             ["Redacted on", datetime.now().strftime("%d/%m/%Y %H:%M")],
             ["Deck ID", job.deck_id],
             ["Mode", "All text except slide titles replaced with [X] (--all)" if job.full
                      else "Only the terms in the keyword list replaced with [X], one per value (--keywords)"
                      if not job.rules else "Confidential details replaced with [X], one per value"],
             ["Items", "Each [X] on a slide is numbered top to bottom (item 1 is highest). If a value's invisible marker "
                       "is lost, restore suggests unmarked [X]s in the same position, to check by hand."],
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
        kind = target.get("kind")
        if kind in ("property", "author"):
            fields.append({**item, "target": target})
        else:
            e = {**item, "kind": kind if kind in ("x", "mask", "placeholder", "link") else "id",
                 "numeric": kind == "number" or bool(target.get("numeric")), "slots": target.get("slots", []),
                 "count": target.get("count", 0), "style": target.get("style", ""), #no style: masked by an older version
                 "parts": target.get("parts", [])}
            if e["kind"] in ("x", "mask"): #--all keys from before the mark field: Text<n>
                e["mark"] = int(target.get("mark") or e["id"].removeprefix("Text"))
            entries.append(e)
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
    if args.keywords and not keywords: #an empty list would redact nothing at all
        raise SystemExit(f"{args.keywords} has no terms. Add one term per line, the deck was not redacted.")
    src = Path(args.pptx)
    job = Job(keywords, labels, args.preview, args.all)
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
    kinds = Counter(e["kind"] for e in job.entries.values())
    if job.full:
        print(f"{kinds['mask']} texts masked, {kinds['x']} values in titles masked with [X], "
              f"{kinds['placeholder']} chart numbers replaced, {kinds['link']} links masked, "
              f"{len(job.rows) - found} review notes -> {out}")
    else:
        print(f"{found} items masked with [X] ({kinds['x']} different values, {kinds['link']} email/phone links), "
              f"{len(job.rows) - found} review notes -> {out}")
    print(f"Restore key -> {key}\nKeep the key on this machine. Send only the redacted deck.")

    #self-check: every marker and placeholder must be found again exactly as often as it was placed
    check = Restore([{**e, "restore": True} for e in job.entries.values()])
    process_package(result, check)
    off = [(i, n, check.found[i]) for i, n in job.placed.items() if check.found[i] != n]
    if off:
        print("WARNING: these items will not restore cleanly:")
        for i, placed, seen in off:
            why = "the deck already had redaction markers, was it redacted before?" if seen > placed else "could not be read back"
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
    short = lambda items: ", ".join(items[:20]) + (f" and {len(items) - 20} more" if len(items) > 20 else "")
    print(f"Restored {len(wanted) - len(missing)} of {len(wanted)} items "
          f"({sum(rj.restored.values())} places) -> {out}")
    if wanted and len(missing) == len(wanted):
        print("WARNING: none of the key's items were found. Is this the right deck and key?")
    elif missing:
        print("Not found, restore by hand (see Locations in the key): " + short(missing))
    partial = [e for e in entries if e["restore"] and 0 < rj.restored[e["id"]] < e.get("count", 0)]
    if partial:
        print("Restored in fewer places than redacted (a marker was lost, or a copy was removed): "
              + short([f"{e['id']} ({rj.restored[e['id']]} of {e['count']})" for e in partial]))
    lost = [e for e in entries if e["kind"] == "x" and (e["id"] in missing or e in partial)]
    if lost and len(missing) < len(wanted):
        #backup: number every [X] per slide top to bottom, then match lost values by position
        by_slide = {}
        for r in rj.xruns:
            by_slide.setdefault(r["loc"], []).append(r)
        for runs in by_slide.values():
            for k, r in enumerate(sorted(runs, key=lambda r: r["pos"]), 1):
                r["item"] = k
        lines = []
        for e in lost:
            hits = suggestions(rj, e)
            if hits:
                was = ", ".join(f"Slide {n} item {k}" for n, k in e["slots"]) or where(e)
                tag = {0: " (same slide and position)", 1: " (same slide)", 2: ""}
                lines.append(f"  {e['id']} (was {was}): "
                             + ", ".join(f"Slide {n} item {k}{tag[rank]}" for rank, n, k in hits))
        if lines:
            print("Possible matches by position, unmarked [X]s (check before restoring by hand):")
            print("\n".join(lines[:20]) + (f"\n  and {len(lines) - 20} more" if len(lines) > 20 else ""))
    if kept:
        print("Kept redacted (Restore = N): " + short(kept))
    if rj.older:
        print("Masked by an older version of this tool, the originals were put back in full (check formatting): "
              + ", ".join(sorted(rj.older, key=slide_key)))
    if rj.edited:
        print("Masked text was edited here, the original was put back in full (check wording and formatting): "
              + ", ".join(sorted(rj.edited, key=slide_key)))
    if rj.fields:
        left = [f for i, f in enumerate(rj.fields) if i not in rj.done]
        print(f"File properties and comment authors restored: {len(rj.fields) - len(left)} of {len(rj.fields)}")
        for (label, type_, inner), n in Counter((f["id"], f["type"], "!" in f["target"]["part"]) for f in left).items():
            print(f"  not restored: {label} {type_}" + (f" in {n} embedded files" if inner else f" x{n}" if n > 1 else ""))
        if any("!" in f["target"]["part"] for f in left):
            print("  (embedded files such as chart data were renamed since redaction, "
                  "decks redacted with this version keep track of them)")
    if rj.layout:
        print("Check layout, restored text is longer than what was on the slide: "
              + ", ".join(sorted(rj.layout, key=slide_key)))


def main():
    ap = argparse.ArgumentParser(description="Rule-based PII redaction for .pptx (no dependencies)")
    ap.add_argument("pptx")
    ap.add_argument("--preview", action="store_true", help="highlight only, no removal")
    ap.add_argument("--all", action="store_true",
                    help="mask all text except slide titles (each paragraph, line and formatting change -> [X], chart labels -> [X]), "
                         "chart numbers become placeholders; restorable with the key")
    ap.add_argument("--keywords", help="text file, one term per line: only these terms are redacted, the built-in "
                                       "email/phone/ID/IBAN/card rules are off; 'term | Label' names it in the key")
    ap.add_argument("--key-dir", help="folder for the restore key (default: next to the deck)")
    ap.add_argument("--restore", metavar="KEY", help="put the original values back using this restore key (.xlsx)")
    ap.add_argument("--force", action="store_true", help="restore even if the key was made for a different deck")
    args = ap.parse_args()
    if args.all and args.preview:
        ap.error("--all cannot be combined with --preview")
    if args.restore:
        restore(args)
    else:
        redact(args)


if __name__ == "__main__":
    main()
