#!/usr/bin/env python3
"""ocr_pdf_text.py — the text a PDF already carries, page by page, for the OCR tool.

Usage: python3 ocr_pdf_text.py in.pdf [max_pages]
stdout: one JSON line
  {"pageCount": n, "pages": [{"page": 1, "usable": true, "text": "…"}, …]}
  or {"error": "encrypted"} / {"error": "unreadable"} (exit 0 either way).
max_pages (default 500) bounds how many pages are looked at; pageCount is always the
whole document.

WHY (10 Oct 2026): every PDF went through OCR, a digital one too. Measured on a 12-page
PDF with a text layer: OCR took 13.9 s for the first 10 pages (the other two were silently
dropped), while its own text came out of PyMuPDF in 11 ms for all 12 pages, exactly.
The OCR service now reads this first and OCRs only the pages without usable text.

A page's text is "usable" when it has some letters (MIN_CHARS non-blank characters) and
does not look like a font without a Unicode map: those extract as private-use code points,
U+FFFD or Latin-1 mojibake ("ÇáÚÑÈíÉ" for Arabic) — such a page is OCR'd instead. A page
that is mostly a picture (a scan) with only a line of text on it (a page number, a stamp,
a caption) is OCR'd as well, since its text layer is not the page's text.
"""
import json
import sys
import unicodedata

import fitz  # PyMuPDF

MIN_CHARS = 20          # fewer non-blank characters: nothing worth keeping
MAX_BAD_SHARE = 0.05    # private use / U+FFFD / control characters
MAX_MOJIBAKE_SHARE = 0.3  # letters in U+00C0..U+00FF among all letters
MAX_STRAY_IN_ARABIC = 0.02  # Latin Extended letters in a page that is mostly Arabic
SCAN_COVERAGE = 0.5     # image area / page area from which a page counts as a scan …
SCAN_MIN_CHARS = 200    # … unless its text layer is a whole page of text (an OCR layer)


def _image_coverage(page):
    """Share of the page covered by images (0..1), as in pdf_to_images.py."""
    try:
        area = abs(page.rect)
        if not area:
            return 0.0
        covered = 0.0
        for info in page.get_image_info():
            r = fitz.Rect(info["bbox"]) & page.rect
            if not r.is_empty:
                covered += abs(r)
        return min(1.0, covered / area)
    except Exception:
        return 0.0


def _letters(text):
    """Arabic presentation forms as the letters they are.

    Some PDF writers map each shaped glyph to its presentation form (U+FB50..U+FDFF,
    U+FE70..U+FEFF): it looks right but cannot be searched or matched with the same word
    typed. NFKC maps exactly those back ("\ufed3\ufe8e" -> "\u0641\u0627"); the rest of
    the text is left alone.
    """
    return "".join(unicodedata.normalize("NFKC", c) if "\ufb50" <= c <= "\ufdff" or "\ufe70" <= c <= "\ufefe" else c
                   for c in text)


def _usable(text, coverage):
    chars = [c for c in text if not c.isspace()]
    n = len(chars)
    if n < MIN_CHARS:
        return False
    bad = sum(1 for c in chars if c == "\ufffd" or "\ue000" <= c <= "\uf8ff" or ord(c) < 32)
    if bad > MAX_BAD_SHARE * n:
        return False
    letters = [c for c in chars if c.isalpha()]
    if letters:
        latin1 = sum(1 for c in letters if "\u00c0" <= c <= "\u00ff")
        if latin1 > MAX_MOJIBAKE_SHARE * len(letters):
            return False
        # Arabic whose ligatures map to stray Latin Extended letters ("\u01c4\u01c1" for
        # lam-lam, seen in a PDF MuPDF itself wrote): OCR reads such a page better.
        arabic = sum(1 for c in letters if "\u0600" <= c <= "\u06ff")
        stray = sum(1 for c in letters if "\u0100" <= c <= "\u024f")
        if arabic > len(letters) // 2 and stray > MAX_STRAY_IN_ARABIC * len(letters):
            return False
    if coverage >= SCAN_COVERAGE and n < SCAN_MIN_CHARS:
        return False
    return True


def main(path, max_pages):
    try:
        doc = fitz.open(path)
    except Exception:
        return {"error": "unreadable"}
    if not doc.is_pdf:
        return {"error": "unreadable"}
    if doc.needs_pass and not doc.authenticate(""):
        return {"error": "encrypted"}
    pages = []
    for i in range(min(doc.page_count, max_pages)):
        try:
            page = doc[i]
            text = _letters(page.get_text("text"))
            usable = _usable(text, _image_coverage(page))
        except Exception:
            text, usable = "", False
        pages.append({"page": i + 1, "usable": usable, "text": text.strip() if usable else ""})
    return {"pageCount": doc.page_count, "pages": pages}


if __name__ == "__main__":
    try:
        limit = int(sys.argv[2]) if len(sys.argv) > 2 else 500
        out = main(sys.argv[1], max(1, limit))
    except Exception:
        out = {"error": "unreadable"}
    sys.stdout.write(json.dumps(out, ensure_ascii=False) + "\n")
