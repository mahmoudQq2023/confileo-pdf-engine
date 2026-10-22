"""Shared text extraction in READING order, built on the glyph-position rebuild
that pdf_to_word.py uses for Arabic PDF->Word (validated 7 Sept 2026 on
WeasyPrint, LibreOffice and Chrome producers).

Why this exists: `page.get_text()` returns Arabic words in VISUAL order for PDFs
made by Cairo/WeasyPrint (our own CV builder among them), so the CV parser read
"أحمد سمير عبد الله" as "الله عبد سمير أحمد" and the audio tool read documents
backwards. Both now go through `page_flow()` here.

API
  page_flow(page)      -> (page_rtl, [(rows, cx0, cx1), ...])   one entry per column,
                          rows sorted top->bottom, each row = pdf_to_word row dict
                          (tokens/size/bold/color/x0/x1/y0/y1/rtl/block)
  row_text(row)        -> the row's text in logical order
  paragraphs(rows, cx0, cx1) -> [[row, row, ...], ...] visual paragraphs
  document_lines(doc)  -> per page: list of (column_index, row) in reading order
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import fitz  # PyMuPDF

from pdf_to_word import (  # noqa: E402
    _line_chars, _merge_rows, _finish_row, _page_columns, _is_arabic_char,
    _normalize_presentation_forms, _tidy, _BULLETS,
)


def page_flow(page):
    """Rows per column in reading order. Mirrors convert_with_fitz() minus images."""
    pw = page.rect.width or 1.0
    flags = fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE
    raw = page.get_text("rawdict", flags=flags)
    frags, blocks = [], []
    for b in raw.get("blocks", []):
        if b.get("type") != 0:
            continue
        x0, y0, x1, y1 = b["bbox"]
        got = False
        for ln in b.get("lines", []):
            chars = _line_chars(ln)
            if any(c[0] not in "  " for c in chars):
                frags.append({"chars": chars, "y0": ln["bbox"][1], "y1": ln["bbox"][3],
                              "x0": ln["bbox"][0], "x1": ln["bbox"][2], "block": len(blocks)})
                got = True
        if got:
            blocks.append({"x0": x0, "x1": x1, "y0": y0, "y1": y1})
    if not frags:
        return False, []

    all_text = "".join(c[0] for f in frags for c in f["chars"])
    ar_n = sum(1 for c in all_text if _is_arabic_char(c))
    lat_n = sum(1 for c in all_text if c.isascii() and c.isalpha())
    page_rtl = ar_n > lat_n
    split = _page_columns(blocks, pw)

    def col(x0, x1):
        if split is None:
            return 0
        right = (x0 + x1) / 2.0 >= split
        return (0 if right else 1) if page_rtl else (1 if right else 0)

    ncols = 1 if split is None else 2
    flow = []
    for ci in range(ncols):
        cf = [f for f in frags if col(f["x0"], f["x1"]) == ci]
        rows = [r for r in (_finish_row(r, page_rtl) for r in _merge_rows(cf)) if r]
        rows.sort(key=lambda r: r["y0"])
        if split is None:
            cx0, cx1 = 0.0, pw
        elif (ci == 0) == page_rtl:
            cx0, cx1 = split, pw
        else:
            cx0, cx1 = 0.0, split
        flow.append((rows, cx0, cx1))
    return page_rtl, flow


def row_text(row):
    return _tidy(" ".join(_normalize_presentation_forms(t["text"]) for t in row["tokens"] if t["text"]))


def row_text_smart(row, chips=True):
    """Like row_text, with two layout-aware touches:
    * letter-spaced headings ("C L I N I C A L  E X P E R I E N C E", tracked small
      caps) are rebuilt into words from the glyph gaps: a gap wider than ~0.55 em
      between two single-letter tokens is a word break, a narrower one is tracking
      (left-to-right rows only; Arabic is never letter-spaced);
    * with chips=True, a gap wider than 1.5 em between two words becomes " · ",
      so skill chips laid out side by side ("X-ray   CT   MRI") and a title with
      its dates on the far edge ("Site Engineer        2017 – 2020") keep their
      boundary instead of fusing into one phrase."""
    toks = [t for t in row["tokens"] if t["text"]]
    size = max(row.get("size") or 0, 1.0)
    if len(toks) < 2:
        return row_text(row)
    singles = sum(1 for t in toks if len(t["text"].strip()) == 1)
    if row.get("rtl") and len(toks) >= 4 and singles >= 0.6 * len(toks):
        # Letter-spaced Arabic (an older template tracked its headings): the
        # tokens are already in logical order, so group them by the gap between
        # neighbours in that order — wide gap = word break, narrow = tracking.
        gaps = [max(a["x0"] - b["x1"], b["x0"] - a["x1"]) for a, b in zip(toks, toks[1:])]
        small = sorted(g for g in gaps if g >= 0)
        base = small[len(small) // 2] if small else 0.0
        thresh = max(0.55 * size, base * 2.2)
        words, cur = [], [toks[0]["text"].strip()]
        for g, t in zip(gaps, toks[1:]):
            if g > thresh:
                words.append("".join(cur)); cur = []
            cur.append(t["text"].strip())
        words.append("".join(cur))
        return _tidy(_normalize_presentation_forms(" ".join(w for w in words if w)))
    if row.get("rtl") or len(toks) < 4 or singles < 0.6 * len(toks):
        if not chips:
            return row_text(row)
        # Two type sizes on one row = two fields (a bold title and its small grey
        # dates). Emit the larger face first, then the smaller, each in its own
        # logical order — when the fields overlap horizontally, ordering by x
        # alone interleaves "…تحت التدريب" with "سبتمبر ٢٠١٣".
        sizes = sorted({round(t["size"] * 2) / 2 for t in toks})
        if len(sizes) == 2 and sizes[1] - sizes[0] >= 1.0:
            big = [t for t in toks if round(t["size"] * 2) / 2 == sizes[1]]
            small = [t for t in toks if round(t["size"] * 2) / 2 == sizes[0]]
            if len(big) >= 1 and len(small) >= 1 and any(ch.isdigit() for t in small for ch in t["text"]):
                a = _tidy(" ".join(_normalize_presentation_forms(t["text"]) for t in big))
                b = _tidy(" ".join(_normalize_presentation_forms(t["text"]) for t in small))
                return f"{a} · {b}"
        parts = [_normalize_presentation_forms(toks[0]["text"])]
        for a, b in zip(toks, toks[1:]):
            gap = max(b["x0"] - a["x1"], a["x0"] - b["x1"])
            wide = gap > 1.5 * size and a["text"].strip() != "·" and b["text"].strip() != "·"
            parts.append(" · " if wide else " ")
            parts.append(_normalize_presentation_forms(b["text"]))
        return _tidy("".join(parts))
    toks = sorted(toks, key=lambda t: t["x0"])
    gaps = [b["x0"] - a["x1"] for a, b in zip(toks, toks[1:])]
    if not gaps:
        return row_text(row)
    small = sorted(g for g in gaps if g >= 0)
    base = small[len(small) // 2] if small else 0.0          # typical letter gap
    thresh = max(0.55 * size, base * 2.2)
    words, cur = [], [toks[0]["text"].strip()]
    for g, t in zip(gaps, toks[1:]):
        if g > thresh:
            words.append("".join(cur)); cur = []
        cur.append(t["text"].strip())
    words.append("".join(cur))
    return _tidy(_normalize_presentation_forms(" ".join(w for w in words if w)))


def paragraphs(rows, cx0, cx1):
    """Group rows into visual paragraphs with the same rules as the Word rebuild:
    a bullet, a size/weight change, a vertical gap, or a block change that is not
    a wrapped line starts a new paragraph."""
    widest = max((r["x1"] - r["x0"] for r in rows), default=1.0)
    paras, prev = [], None
    for ln in rows:
        first = ln["tokens"][0]["text"][:1] if ln["tokens"] else ""
        style_break = (
            first in _BULLETS
            or prev is None
            or abs(ln["size"] - prev["size"]) > 0.9
            or ln["bold"] != prev["bold"]
            or (ln["y0"] - prev["y1"]) > max(2.0, 0.45 * max(ln["size"], prev["size"]))
        )
        block_break = False
        if not style_break and prev is not None and ln["block"] != prev["block"]:
            prev_text = row_text(prev).rstrip()
            wrapped = (prev["x1"] - prev["x0"]) >= 0.8 * widest and prev_text[-1:] not in ".:!?؟؛"
            block_break = not wrapped
        if style_break or block_break:
            paras.append([ln])
        else:
            paras[-1].append(ln)
        prev = ln
    return paras


def document_lines(doc):
    """Reading-order rows for a whole document: [[(col, row), ...] per page]."""
    out = []
    for page in doc:
        _, flow = page_flow(page)
        page_rows = []
        for ci, (rows, _, _) in enumerate(flow):
            page_rows.extend((ci, r) for r in rows)
        out.append(page_rows)
    return out
