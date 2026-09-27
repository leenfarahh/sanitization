
#python sanitize_pdf.py input.pdf --preview   --> highlight matches + CSV report, nothing removed
#python sanitize_pdf.py input.pdf             --> redact with "XXXX" overlay
#python sanitize_pdf.py input.pdf --keywords names.txt --label "[REDACTED]"

import argparse
import csv
import re
import sys
from pathlib import Path

import pymupdf

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

#rules(name, regex, validator or None, context keywords or None)
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

    # Phones: international and local formats
    ("PHONE_UAE", rf"(?:(?:\+|00)971{SEP}|\b0)(?:5[024568]|[2-4679]){SEP}\d{{3}}{SEP}\d{{4}}\b", None, None),
    ("PHONE_KSA", rf"(?:(?:\+|00)966{SEP}|\b0)(?:5\d|1[1-7]){SEP}\d{{3}}{SEP}\d{{4}}\b", None, None),
    ("PHONE_JO",  rf"(?:(?:\+|00)962{SEP}|\b0)(?:7[789]|[2356]){SEP}\d{{3}}{SEP}\d{{4}}\b", None, None),
    ("PHONE_QA",  rf"(?:\+|00)974{SEP}[34567]\d{{3}}{SEP}\d{{4}}\b", None, None),
]

COMPILED = [(n, re.compile(p), v, c) for n, p, v, c in RULES]

DIGIT_MAP = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

#detection
def find_matches(text: str, keywords=None):
    #return list of (start, end, rule_name). overlaps resolved: first/longest wins
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

#pdf processing
def merge_by_line(rects):
    #merge word boxes on the same line into one box, so each match gets one label
    merged = []
    for r in rects:
        if merged and abs(merged[-1].y0 - r.y0) < 2 and abs(merged[-1].y1 - r.y1) < 2:
            merged[-1] |= r
        else:
            merged.append(pymupdf.Rect(r))
    return merged


def process_page(page, keywords):
    #join words into one string, map char spans back to word boxes.
    words = page.get_text("words")  # (x0, y0, x1, y1, text, block, line, word)
    spans, parts, pos = [], [], 0
    for w in words:
        spans.append((pos, pos + len(w[4]), pymupdf.Rect(w[:4])))
        parts.append(w[4])
        pos += len(w[4]) + 1
    text = " ".join(parts)

    found = []
    for start, end, rule in find_matches(text, keywords):
        rects = [r for s, e, r in spans if s < end and e > start]
        found.append((rule, text[start:end], merge_by_line(rects)))
    return found


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--preview", action="store_true", help="highlight only, no removal")
    ap.add_argument("--keywords", help="text file, one term per line (client names, codes)")
    ap.add_argument("--label", default="XXXX")
    args = ap.parse_args()

    #fall back to keywords.txt in the working folder when --keywords is not given
    keywords = []
    kw_file = args.keywords or ("keywords.txt" if Path("keywords.txt").exists() else None)
    if kw_file:
        keywords = [l.strip() for l in Path(kw_file).read_text(encoding="utf-8").splitlines() if l.strip()]

    src = Path(args.pdf)
    doc = pymupdf.open(src)
    report = []

    for n, page in enumerate(doc.pages(), 1):
        for rule, value, rects in process_page(page, keywords):
            report.append((n, rule, mask(value)))
            for r in rects:
                if args.preview:
                    page.add_highlight_annot(r)
                else:
                    page.add_redact_annot(r, text=args.label, fill=(0, 0, 0),
                                          text_color=(1, 1, 1), fontsize=0)
        if not args.preview:
            page.apply_redactions()

    suffix = "_preview" if args.preview else "_redacted"
    out = src.with_name(src.stem + suffix + ".pdf")
    if not args.preview:
        doc.set_metadata({})
        doc.del_xml_metadata()
    doc.save(out, garbage=4, deflate=True)

    rep = src.with_name(src.stem + suffix + "_report.csv")
    with open(rep, "w", newline="", encoding="utf-8-sig") as f: # utf-8-sig opens cleanly in Excel
        w = csv.writer(f)
        w.writerow(["page", "type", "masked_value"])
        w.writerows(report)

    print(f"{len(report)} matches -> {out}\nReport -> {rep}")
    if doc.page_count and not any(p.get_text("words") for p in doc):
        print("WARNING: no text layer found. This looks scanned; run OCR first.", file=sys.stderr)


if __name__ == "__main__":
    main()