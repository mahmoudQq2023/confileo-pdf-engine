#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF -> rows for the PDF->Excel tool.

Rewritten 7 Sept 2026 after a live audit found the old pipeline sending every
"auto" request (the UI default) through the Arabic OCR engine even when the PDF
had a perfect text layer — and then "un-reversing" every number it saw, so a
payroll sheet came back with 18,500 as 581 and 7,724.50 as 5.4277. Ruled tables
lost every table but the first; borderless tables and Latin scans came back as an
empty workbook that still said "success".

What this does now:
  1. A page WITH a text layer never goes near OCR.
     * Ruled tables come from PyMuPDF's line detector (all of them, not [0]); the
       text of each cell is rebuilt from glyph positions with the same machinery
       pdf_to_word.py uses, so Arabic cells read in logical order on Cairo,
       LibreOffice and Chrome PDFs alike (find_tables().extract() hands back
       visual-order Arabic: "يفاصلا" for "الصافي").
     * Everything outside a ruled table becomes text rows; a row is split into
       cells wherever a gap wider than ~1.4 em separates two words, and runs of
       such rows are aligned into columns from the gutters they share — that is
       how borderless tables (bank statements, price lists) become a real grid
       instead of an empty sheet. Wrapped cell text on the next line is folded
       back into its cell.
  2. A page WITHOUT a text layer is OCR'd (PaddleOCR microservice) with the engine
     that matches its script: "ar" or "en" as asked, and for "auto" the Arabic
     pass is checked and redone in English when the page turns out to be Latin
     (the Arabic model reads "Description" as "Descriptio N" and "0%" as "0/0").
     OCR lines then go through the same row/column alignment.
  3. Numbers are never reordered. Arabic-Indic digits are mapped to Western so
     Excel can treat them as numbers; the digit order is kept as read.

Output JSON (consumed by PdfToExcelService.cs):
  {success, isRTL, engine, pages:[{pageNumber, blocks:[{kind:"table"|"text",
   rows:[[cell,...]], headerRow}], rows, headerRow}], totalRows, totalTables}

Usage: python pdf_table_extractor.py <input.pdf> <output.json> [arabic|english|auto]
"""
import json
import os
import re
import sys
import traceback

# PyMuPDF prints a "Consider using the pymupdf_layout package" hint to stdout on
# its first table search; the reply on stdout is JSON.
os.environ.setdefault("PYMUPDF_SUGGEST_LAYOUT_ANALYZER", "0")

try:
    import fitz  # PyMuPDF
except ImportError:
    print(json.dumps({"success": False, "message": "PyMuPDF (fitz) is not installed."}))
    sys.exit(1)

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from pdf_to_word import (  # noqa: E402
    _line_chars, _merge_rows, _finish_row, _normalize_presentation_forms,
    _is_arabic_char, _tidy, _garbled_text, _find_tables,
)

try:
    import ocr_client
except Exception:
    ocr_client = None

MIN_TEXT_CHARS = 20          # fewer real glyphs than this = a scanned page
GAP_EM = 1.4                 # a gap wider than this (in em) splits a row into cells
MIN_GUTTER = 3.0             # pt; narrower gaps between columns are ignored

_AR2WE = str.maketrans("٠١٢٣٤٥٦٧٨٩٫٬", "0123456789.,")
_NUMERIC_RE = re.compile(r"^[\s$€£¥%+\-−()\d.,/:٪]+$")


def _clean(s):
    return _tidy(re.sub(r"\s+", " ", (s or "").replace("\n", " ")).strip(_EDGE_MARKS))


# Invisible direction marks and zero-width spaces at either end of a cell: Tesseract
# ends Arabic-context words with U+200F, and 33 of the 935 cells of an OCR'd page came
# out «dribbble.com\u200f», so =VLOOKUP("dribbble.com", …) and A1="dribbble.com"
# missed them (10 Oct 2026). Inside the text they stay: there they order mixed text.
_EDGE_MARKS = " \t\u200b\u200e\u200f\u061c\u202a\u202b\u202c\u202d\u202e\u2060\u2066\u2067\u2068\u2069\ufeff"


_AR_DIGIT_RUN = re.compile(r"[٠-٩٫٬]+")


def _norm_digits(cell):
    """OCR digits, measured against the live engine (7 Sept 2026): Western
    digits come back in reading order ("1494", "25,750.50") and must not be
    touched — the old blanket "undo RTL order" turned 18,500 into 581. The
    Arabic model does return ARABIC-INDIC runs reversed ("٣٨" for "٨٣", "٢١" for
    "١٢"), so only those runs are flipped, then mapped to Western so Excel can
    treat them as numbers."""
    if not cell or not any("٠" <= c <= "٩" for c in cell):
        return cell
    return _AR_DIGIT_RUN.sub(lambda m: m.group(0)[::-1], cell).translate(_AR2WE)


def _is_numeric(text):
    t = (text or "").strip()
    return bool(t) and any(ch.isdigit() for ch in t) and bool(_NUMERIC_RE.match(t))


def _has_number(text):
    return any(_is_numeric(w) for w in (text or "").split())


def _script_counts(text):
    ar = sum(1 for c in text if _is_arabic_char(c))
    lat = sum(1 for c in text if c.isascii() and c.isalpha())
    return ar, lat


# ── text-layer pages ───────────────────────────────────────────────────────

def page_frags(page):
    """rawdict lines -> fragments (chars with x per glyph, y per line)."""
    flags = fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE
    raw = page.get_text("rawdict", flags=flags)
    frags, nblocks = [], 0
    for b in raw.get("blocks", []):
        if b.get("type") != 0:
            continue
        got = False
        for ln in b.get("lines", []):
            chars = _line_chars(ln)
            if any(c[0] not in "  " for c in chars):
                frags.append({"chars": chars, "y0": ln["bbox"][1], "y1": ln["bbox"][3],
                              "x0": ln["bbox"][0], "x1": ln["bbox"][2], "block": nblocks})
                got = True
        if got:
            nblocks += 1
    return frags


def _rows_from_frags(frags, page_rtl):
    copies = [dict(f, chars=[list(c) for c in f["chars"]]) for f in frags]
    rows = [r for r in (_finish_row(r, page_rtl) for r in _merge_rows(copies)) if r]
    rows.sort(key=lambda r: r["y0"])
    return rows


def _row_text(row):
    return _tidy(" ".join(_normalize_presentation_forms(t["text"]) for t in row["tokens"] if t["text"]))


def _cell_text(frags, page_rtl):
    if not frags:
        return ""
    return _clean(" ".join(_row_text(r) for r in _rows_from_frags(frags, page_rtl)))


def _split_frag_by_cells(frag, cell_of):
    """Split one line's glyphs among table cells. cell_of(xc, yc) -> key or None."""
    yc = (frag["y0"] + frag["y1"]) / 2.0
    groups = {}
    for ch in frag["chars"]:
        key = cell_of((ch[1] + ch[2]) / 2.0, yc)
        groups.setdefault(key, []).append(ch)
    out = {}
    for key, chars in groups.items():
        wide = [c for c in chars if c[2] - c[1] > 0.01] or chars
        out[key] = {"chars": chars, "y0": frag["y0"], "y1": frag["y1"],
                    "x0": min(c[1] for c in wide), "x1": max(c[2] for c in wide), "block": frag["block"]}
    return out


def ruled_tables(page, frags, page_rtl):
    """Every ruled table on the page, with logical-order cell text. Returns
    (tables, leftover_frags): the glyphs not inside any table."""
    try:
        found = _find_tables(page)        # restores MuPDF's glyph boxes when it raises
    except Exception:
        return [], frags
    tables = []
    # PyMuPDF 1.26.7 returns None where 1.26.5 raises (a Type3 web font in a
    # Chrome PDF): the whole sheet failed with "'NoneType' object has no
    # attribute 'tables'" (10 Oct 2026)
    for t in getattr(found, "tables", None) or []:
        if t.row_count < 1 or t.col_count < 1 or (t.row_count < 2 and t.col_count < 2):
            continue
        cells = []                                   # (row, col, rect)
        for ri, row in enumerate(t.rows):
            for ci, cb in enumerate(row.cells):
                if cb is None:
                    continue
                cells.append((ri, ci, fitz.Rect(cb)))
        if not cells:
            continue
        tables.append({"table": t, "cells": cells, "bbox": fitz.Rect(t.bbox)})
    if not tables:
        return [], frags

    def cell_of(xc, yc):
        for ti, tb in enumerate(tables):
            if not tb["bbox"].contains(fitz.Point(xc, yc)):
                continue
            for ri, ci, r in tb["cells"]:
                if r.x0 - 0.5 <= xc <= r.x1 + 0.5 and r.y0 - 0.5 <= yc <= r.y1 + 0.5:
                    return (ti, ri, ci)
            return (ti, -1, -1)                     # inside the table, in no cell (ruling gap)
        return None

    per_cell, leftover = {}, []
    fills = None
    for f in frags:
        parts = _split_frag_by_cells(f, cell_of)
        for key, part in parts.items():
            if key is None:
                leftover.append(part)
            elif key[1] >= 0:
                per_cell.setdefault(key, []).append(part)
            # glyphs that fall between cells of a table are dropped: they are
            # ruling noise, never a value the sheet is missing.

    out = []
    for ti, tb in enumerate(tables):
        t = tb["table"]
        grid = [["" for _ in range(t.col_count)] for _ in range(t.row_count)]
        for (tti, ri, ci), fr in per_cell.items():
            if tti == ti:
                grid[ri][ci] = _cell_text(fr, page_rtl)
        keep = [ri for ri in range(t.row_count) if any(c for c in grid[ri])]
        rows = [grid[ri] for ri in keep]
        if not rows:
            continue
        if page_rtl:
            for r in rows:
                r.reverse()
        # Geometry for PDF->Word (a zebra-striped table is detected from its
        # shaded rows only; the white rows between them are put back by x/y).
        # extract_tables() keeps only kind/rows/headerRow, so Excel is unchanged.
        grid_xs, layout = _merged_layout(t, tb["cells"], keep, page_rtl)
        if layout:
            if fills is None:
                fills = _FillIndex(page)
            rect_of = {(ri, ci): r for ri, ci, r in tb["cells"]}
            for entries in layout:
                for e in entries:
                    key = e.pop("cell", None)
                    if key is not None:
                        fill = fills.at(rect_of[key])
                        look = _cell_look(per_cell.get((ti,) + key, ()), rect_of[key])
                        if fill:
                            e["fill"] = fill
                        color = look.pop("color", None)
                        if color and (fill or not _light(color)):
                            e["color"] = color
                        e.update(look)
        col_x = []
        if grid_xs:
            # the rulings themselves: a cell merged across columns no longer
            # widens the column it starts in
            col_x = [[grid_xs[c], grid_xs[c + 1]] for c in range(t.col_count)]
        else:
            for ci in range(t.col_count):
                rs = [r for (_, c, r) in tb["cells"] if c == ci]
                if rs:
                    col_x.append([min(r.x0 for r in rs), max(r.x1 for r in rs)])
        row_y = []
        for ri in keep:
            try:
                row_y.append(float(t.rows[ri].bbox[1]))
            except Exception:  # noqa: BLE001
                row_y.append(float(tb["bbox"].y0))
        sizes = {}
        for (tti, _, _), fr in per_cell.items():
            if tti == ti:
                for part in fr:
                    for ch in part["chars"]:
                        if ch[0].strip() and ch[4]:
                            k = round(float(ch[4]) * 2) / 2
                            sizes[k] = sizes.get(k, 0) + 1
        out.append({"kind": "table", "y0": tb["bbox"].y0, "rows": rows, "headerRow": 0,
                    "y1": tb["bbox"].y1, "x0": tb["bbox"].x0, "x1": tb["bbox"].x1,
                    "row_y": row_y, "col_x": col_x if len(col_x) == t.col_count else None,
                    "layout": layout, "size": max(sizes, key=sizes.get) if sizes else None})
    return out, leftover


class _FillIndex:
    """The page's filled shapes, to find the colour a table cell is painted:
    the last shape (painter's order) that covers most of the cell. A white
    shape painted over a coloured one clears it."""

    def __init__(self, page, step=12.0):
        self.step = step
        self.bins = {}
        try:
            drawings = page.get_drawings()
        except Exception:  # noqa: BLE001 - no colours then
            drawings = []
        for n, d in enumerate(drawings[:5000]):
            f, r = d.get("fill"), d.get("rect")
            if not f or r is None or r.is_empty or (d.get("fill_opacity") or 1.0) < 0.5:
                continue
            rgb = tuple(int(round(max(0.0, min(1.0, v)) * 255)) for v in f[:3])
            color = None if min(rgb) >= 247 else "%02X%02X%02X" % rgb
            item = (n, fitz.Rect(r), color)
            for b in range(int(r.y0 // step), int(r.y1 // step) + 1):
                self.bins.setdefault(b, []).append(item)

    def at(self, rect):
        area = rect.width * rect.height
        if area <= 0:
            return None
        best = None
        for n, fr, color in self.bins.get(int(((rect.y0 + rect.y1) / 2.0) // self.step), ()):
            inter = fitz.Rect(fr) & rect
            if not inter.is_empty and inter.width * inter.height >= 0.85 * area \
                    and (best is None or n > best[0]):
                best = (n, color)
        return best[1] if best else None


def _cell_look(parts, rect):
    """How a cell's text is printed: {"color": "RRGGBB" (not black), "bold": True,
    "align": "left" | "center" | "right" (on the page), "valign": "center"}."""
    counts, bold, n = {}, 0, 0
    tx0 = ty0 = float("inf")
    tx1 = ty1 = float("-inf")
    for part in parts:
        for ch in part["chars"]:
            sp = ch[3] if isinstance(ch[3], dict) else None
            if sp is None or not ch[0].strip():
                continue
            n += 1
            c = int(sp.get("color", 0)) & 0xFFFFFF
            counts[c] = counts.get(c, 0) + 1
            name = (sp.get("font") or "").lower()
            if (sp.get("flags", 0) & 16) or "bold" in name or "black" in name or "heavy" in name:
                bold += 1
            if ch[2] - ch[1] > 0.01:
                tx0, tx1 = min(tx0, ch[1]), max(tx1, ch[2])
        ty0, ty1 = min(ty0, part["y0"]), max(ty1, part["y1"])
    look = {}
    if not n:
        return look
    c = max(counts, key=counts.get)
    if c:
        look["color"] = "%06X" % c
    if bold * 2 > n:
        look["bold"] = True
    if tx1 > tx0:
        left, right = tx0 - rect.x0, rect.x1 - tx1
        if left > 3 and right > 3 and abs(left - right) <= max(3.0, 0.12 * rect.width):
            look["align"] = "center"
        else:
            look["align"] = "right" if right < left else "left"
    if ty1 > ty0 and rect.height > (ty1 - ty0) + 4:
        mid = (ty0 + ty1) / 2.0 - (rect.y0 + rect.y1) / 2.0
        if abs(mid) <= max(2.0, 0.12 * rect.height):
            look["valign"] = "center"
    return look


def _light(color):
    return min(int(color[i:i + 2], 16) for i in (0, 2, 4)) >= 200


def _merged_layout(t, cells, keep, rtl):
    """Merged cells of a ruled table, for PDF->Word: (column edges, layout).

    layout has one list per kept row: its cells in reading order as
    {"span": grid columns covered, "vmerge": None | "restart" | "continue",
    "col": index of the cell's text in that row}. PyMuPDF reports a merged cell
    once, with its full rectangle, and None for the slots it covers; those used
    to become empty bordered cells beside a squeezed one ("الإجمالي الكلي" in one
    narrow column, three blanks beside it). (None, None) when the rulings do not
    form a plain grid: the table then stays one cell per slot, as before."""
    ncol, nrow = t.col_count, t.row_count

    def edges(vals):
        out = []
        for v in sorted(vals):
            if not out or v - out[-1] > 1.0:
                out.append(v)
        return out

    xs = edges([r.x0 for _, _, r in cells] + [r.x1 for _, _, r in cells])
    ys = edges([r.y0 for _, _, r in cells] + [r.y1 for _, _, r in cells])
    if len(xs) != ncol + 1 or len(ys) != nrow + 1:
        return None, None
    owner = {}
    for ri, ci, r in cells:
        cs = [c for c in range(ncol) if r.x0 - 1 <= (xs[c] + xs[c + 1]) / 2.0 <= r.x1 + 1]
        rs = [q for q in range(nrow) if r.y0 - 1 <= (ys[q] + ys[q + 1]) / 2.0 <= r.y1 + 1]
        if not cs or not rs or cs[0] != ci or rs[0] != ri:
            return None, None
        for q in rs:
            for c in cs:
                if (q, c) in owner:
                    return None, None
                owner[(q, c)] = (ri, ci)
    layout = []
    for k, q in enumerate(keep):
        entries, c = [], 0
        while c < ncol:
            own = owner.get((q, c))
            span = 1
            while own is not None and c + span < ncol and owner.get((q, c + span)) == own:
                span += 1
            vm, col = None, c
            if own is not None and own[0] == q:
                col = own[1]
                if k + 1 < len(keep) and owner.get((keep[k + 1], c)) == own:
                    vm = "restart"
            elif own is not None and k > 0 and owner.get((keep[k - 1], c)) == own:
                vm = "continue"
            # "cell": the PyMuPDF cell drawn here, for its colours (dropped once read)
            entries.append({"span": span, "vmerge": vm, "col": col, "cell": own})
            c += span
        if rtl:
            # the row's text was reversed into reading order; so are its cells
            entries.reverse()
            for e in entries:
                e["col"] = ncol - 1 - e["col"]
        layout.append(entries)
    return xs, layout


# ── rows -> cells -> columns (shared by text rows and OCR lines) ───────────

def _gap_threshold(gaps, size):
    """Where does a word gap end and a cell gap begin on this row? The gaps of
    a table row fall in two clusters (0.3 em inside a cell, 1 em and more
    between cells); justified prose has one smooth cluster. Look for the jump."""
    positive = sorted(g for g in gaps if g > 0)
    if not positive:
        return GAP_EM * size
    for lo, hi in zip(positive, positive[1:]):
        if hi >= 0.8 * size and hi >= 1.8 * lo:
            return hi - 0.01
    if positive[0] >= 0.75 * size:               # every gap is a cell gap
        return positive[0] - 0.01
    return GAP_EM * size


def _row_segments(row):
    """Split a rebuilt row into cells at wide gaps. Segments come back in
    reading order (right-to-left for an RTL row) with their visual x-range."""
    toks = [t for t in row["tokens"] if t["text"].strip()]
    if not toks:
        return []
    size = max(row.get("size") or 0.0, 1.0)
    order = {id(t): i for i, t in enumerate(toks)}
    byx = sorted(toks, key=lambda t: t["x0"])
    gaps = [b["x0"] - a["x1"] for a, b in zip(byx, byx[1:])]
    threshold = _gap_threshold(gaps, size)
    groups = [[byx[0]]]
    for g, b in zip(gaps, byx[1:]):
        if g >= threshold:
            groups.append([b])
        else:
            groups[-1].append(b)
    segs = []
    for g in groups:
        g.sort(key=lambda t: order[id(t)])
        toks = [{"x0": t["x0"], "x1": t["x1"], "text": _normalize_presentation_forms(t["text"]).strip()} for t in g]
        segs.append({"x0": min(t["x0"] for t in g), "x1": max(t["x1"] for t in g),
                     "text": _clean(" ".join(t["text"] for t in toks)), "toks": toks})
    if row.get("rtl"):
        segs.reverse()
    return [s for s in segs if s["text"]]


def _column_bounds(seg_rows):
    """Column x-intervals for a run of rows. An x is a gutter when most of the
    rows that reach across it leave it empty — a vote, so one row whose cells
    ran together (a header that wrapped, a merged "Subtotal" cell) cannot erase
    a column boundary the other rows agree on."""
    multi = [r for r in seg_rows if len(r["segs"]) >= 2]
    if len(multi) < 2:
        return None
    xmin = min(s["x0"] for r in multi for s in r["segs"])
    xmax = max(s["x1"] for r in multi for s in r["segs"])
    n = int(xmax - xmin) + 2
    span = [0] * n
    free = [0] * n
    for r in multi:
        segs = sorted(r["segs"], key=lambda s: s["x0"])
        a = max(0, int(segs[0]["x0"] - xmin))
        b = min(n - 1, int(segs[-1]["x1"] - xmin) + 1)
        for x in range(a, b + 1):
            span[x] += 1
        for s1, s2 in zip(segs, segs[1:]):
            for x in range(max(0, int(s1["x1"] - xmin) + 1), min(n, int(s2["x0"] - xmin))):
                free[x] += 1
    gutter = [span[x] > 0 and free[x] > 0.5 * span[x] for x in range(n)]
    cols, x = [], 0
    while x < n:
        if not gutter[x] and span[x] > 0:
            s = x
            while x < n and not gutter[x] and span[x] > 0:
                x += 1
            cols.append([s + xmin, x - 1 + xmin])
        else:
            x += 1
    merged = []
    for c in cols:
        if merged and c[0] - merged[-1][1] < MIN_GUTTER:
            merged[-1][1] = c[1]
        else:
            merged.append(c)
    return merged if len(merged) >= 2 else None


def _overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def _place(segs, cols):
    """Distribute a row over the columns. Each WORD goes to the column it overlaps
    most, so a segment whose cells ran together on this row (a header that had
    tighter gaps than its data) still lands in the right cells; words that share
    a column keep their reading order."""
    cells = [[] for _ in cols]
    for s in segs:
        last = None
        for t in s.get("toks") or [s]:
            best, bi = -1.0, None
            for i, (c0, c1) in enumerate(cols):
                ov = _overlap(t["x0"], t["x1"], c0, c1)
                if ov > best:
                    best, bi = ov, i
            if bi is None or best <= 0:
                # the word sits in a gutter: it belongs with the word before it
                # in the same visual chunk (a long description running past its
                # column), or, first in its chunk, with the nearest column
                if last is not None:
                    bi = last
                else:
                    cx = (t["x0"] + t["x1"]) / 2.0
                    bi = min(range(len(cols)), key=lambda i: abs((cols[i][0] + cols[i][1]) / 2.0 - cx))
            if t["text"]:
                cells[bi].append(t["text"])
            last = bi
    return [_clean(" ".join(c)) for c in cells]


def _fits_one_column(seg, cols):
    hits = [i for i, (c0, c1) in enumerate(cols) if _overlap(seg["x0"], seg["x1"], c0, c1) > 0.2 * (seg["x1"] - seg["x0"] + 0.1)]
    return hits[0] if len(hits) == 1 else None


def build_blocks(seg_rows, rtl):
    """seg_rows: [{y0, y1, segs:[{x0,x1,text}]}] in reading order -> blocks.
    Consecutive rows that share column gutters become one table; the rest are
    text rows (already split into cells at wide gaps)."""
    blocks = []
    i, n = 0, len(seg_rows)
    while i < n:
        r = seg_rows[i]
        if len(r["segs"]) < 2:
            blocks.append({"kind": "text", "y0": r["y0"], "rows": [[s["text"] for s in r["segs"]]], "headerRow": -1})
            i += 1
            continue
        # grow a run of table-ish rows
        j = i + 1
        last_multi = r                                  # the last real data row of the run
        while j < n:
            nxt, prev = seg_rows[j], seg_rows[j - 1]
            gap = nxt["y0"] - prev["y1"]
            h = max(nxt["y1"] - nxt["y0"], 1.0)
            ns, nm = len(nxt["segs"]), len(last_multi["segs"])
            similar = abs(ns - nm) <= 3 or (ns >= 4 and nm >= 4)
            if ns >= 2 and gap < 2.5 * h and similar:
                last_multi = nxt
                j += 1
            elif 1 <= ns < nm and gap < 0.9 * h:
                j += 1                                  # a wrapped cell line, tight under its row
            else:
                break
        run = seg_rows[i:j]
        cols = _column_bounds(run)
        if cols is None:
            for rr in run:
                blocks.append({"kind": "text", "y0": rr["y0"], "rows": [[s["text"] for s in s_order(rr)]], "headerRow": -1})
            i = j
            continue
        if rtl:
            cols = cols[::-1]
        rows, prev_row, prev_y1 = [], None, None
        for rr in run:
            single = _fits_one_column(rr["segs"][0], cols) if len(rr["segs"]) == 1 else None
            if single is not None and prev_row is not None and prev_y1 is not None \
                    and (rr["y0"] - prev_y1) < 0.9 * max(rr["y1"] - rr["y0"], 1.0) \
                    and prev_row[single] and not _is_numeric(prev_row[single]) and not _has_number(rr["segs"][0]["text"]):
                prev_row[single] = prev_row[single] + " " + rr["segs"][0]["text"]
                prev_y1 = rr["y1"]
                continue
            cells = _place(rr["segs"], cols)
            # a continuation line: fewer cells than columns, every one of them
            # text under text, and sitting tight beneath the previous row
            filled = sum(1 for c in cells if c)
            if prev_row is not None and prev_y1 is not None and 0 < filled <= max(1, len(cols) // 2) \
                    and (rr["y0"] - prev_y1) < 0.6 * max(rr["y1"] - rr["y0"], 1.0) \
                    and all((not c) or (prev_row[k] and not _is_numeric(prev_row[k]) and not _has_number(c)) for k, c in enumerate(cells)):
                for k, c in enumerate(cells):
                    if c:
                        prev_row[k] = prev_row[k] + " " + c
                prev_y1 = rr["y1"]
                continue
            rows.append(cells)
            prev_row, prev_y1 = cells, rr["y1"]
        # y1: where the run ends, for PDF to Word's OCR pages (Excel keeps kind/rows/headerRow)
        blocks.append({"kind": "table", "y0": run[0]["y0"], "y1": max(rr["y1"] for rr in run),
                       "rows": rows, "headerRow": 0})
        i = j
    return blocks


def s_order(rr):
    return rr["segs"]


def text_layer_page(page, page_rtl_hint=None):
    frags = page_frags(page)
    if not frags:
        return None, ""
    all_text = "".join(c[0] for f in frags for c in f["chars"])
    if sum(1 for c in all_text if not c.isspace()) < MIN_TEXT_CHARS:
        return None, all_text
    if _garbled_text(page.get_text("text")):
        # a text layer that is mojibake (a broken ToUnicode map, an old Arabic
        # DTP encoding): its "text" is not the page's, so the page is read
        # like a scan (see pdf_to_word._garbled_text)
        return None, ""
    ar, lat = _script_counts(all_text)
    page_rtl = ar > lat if page_rtl_hint is None else page_rtl_hint
    tables, leftover = ruled_tables(page, frags, page_rtl)
    seg_rows = []
    for row in _rows_from_frags(leftover, page_rtl):
        segs = _row_segments(row)
        if segs:
            seg_rows.append({"y0": row["y0"], "y1": row["y1"], "segs": segs})
    blocks = tables + build_blocks(seg_rows, page_rtl)
    blocks.sort(key=lambda b: b["y0"])
    return blocks, all_text


# ── scanned pages ──────────────────────────────────────────────────────────

def ocr_seg_rows(op, rtl):
    """OCR lines -> seg rows (y-clustered, reading order within a row)."""
    lines = [l for l in op.get("lines", []) if l.get("box") and _clean(l.get("text"))]
    if not lines:
        return []
    sx = (op.get("pdf_width") or 0) and 1.0
    lines.sort(key=lambda l: (l["box"][1] + l["box"][3]) / 2.0)
    heights = sorted(l["box"][3] - l["box"][1] for l in lines)
    tol = max(6.0, heights[len(heights) // 2] * 0.55)
    rows, cur = [], [lines[0]]
    cy = (lines[0]["box"][1] + lines[0]["box"][3]) / 2.0
    for l in lines[1:]:
        c = (l["box"][1] + l["box"][3]) / 2.0
        if abs(c - cy) <= tol:
            cur.append(l)
            cy = (cy + c) / 2.0
        else:
            rows.append(cur)
            cur, cy = [l], c
    rows.append(cur)
    out = []
    for r in rows:
        r.sort(key=lambda l: (l["box"][0] + l["box"][2]) / 2.0, reverse=rtl)
        segs = [{"x0": l["box"][0], "x1": l["box"][2], "text": _clean(l["text"]),
                 "toks": [{"x0": l["box"][0], "x1": l["box"][2], "text": _clean(l["text"])}]} for l in r]
        out.append({"y0": min(l["box"][1] for l in r), "y1": max(l["box"][3] for l in r), "segs": segs})
    return out


def osd_script(pdf_path, pno):
    """'ar' or 'en' from Tesseract's script detection on one page (0.6 s), or None.

    "auto" used to OCR every scanned page with the Arabic engine and, when the text came
    back mostly Latin, OCR every page AGAIN in English: an English invoice took 21.8 s
    and a one-page French table 48 s (22-24 Sept 2026). Asking OSD first means one pass
    in the right language."""
    import subprocess
    import tempfile
    png = None
    try:
        doc = fitz.open(pdf_path)
        pix = doc[pno - 1].get_pixmap(dpi=150)
        fd, png = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        pix.save(png)
        doc.close()
        out = subprocess.run(["tesseract", png, "stdout", "--psm", "0", "-l", "osd"],
                             capture_output=True, text=True, timeout=20)
        text = (out.stdout or "") + (out.stderr or "")
        m = re.search(r"Script:\s*(\w+)", text)
        c = re.search(r"Script confidence:\s*([\d.]+)", text)
        if not m or not c or float(c.group(1)) < 1.0:
            return None
        return "ar" if m.group(1).lower() == "arabic" else "en"
    except Exception:
        return None
    finally:
        if png:
            try:
                os.remove(png)
            except OSError:
                pass


def ocr_pages(pdf_path, page_numbers, lang):
    """OCR only the given pages (1-based) through the microservice."""
    if not page_numbers:
        return {}
    src = fitz.open(pdf_path)
    sub = fitz.open()
    for pno in page_numbers:
        sub.insert_pdf(src, from_page=pno - 1, to_page=pno - 1)
    tmp = pdf_path + f".ocr-{lang}.pdf"
    sub.save(tmp)
    sub.close(); src.close()
    try:
        resp = ocr_client.ocr_pdf(tmp, lang=lang)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    pages = resp.get("pages", []) if isinstance(resp, dict) else []
    return {page_numbers[i]: p for i, p in enumerate(pages) if i < len(page_numbers)}


def tesseract_pages(pdf_path, page_numbers, lang):
    """{page: {"lines": [{"text", "box"}]}} for the given pages (1-based) from
    Tesseract's word boxes, or {} when Tesseract is missing.

    Scanned pages used to need the PaddleOCR service; with it down (or absent)
    a scanned table came back as "No readable text or tables were found" (2 Oct
    2026). Tesseract is what PDF to Word reads scans with. Its words are grouped
    into the pieces of a line that stand apart — a wide gap is a cell border —
    so the rows and columns are rebuilt from positions as for the service's
    lines; the order Tesseract reads its blocks in does not matter here. Boxes
    are in PDF points."""
    import shutil
    import subprocess
    from concurrent.futures import ThreadPoolExecutor
    tess = shutil.which("tesseract")
    if not tess or not page_numbers:
        return {}
    env = dict(os.environ, OMP_THREAD_LIMIT="1")
    dpi = 300

    def read(img):
        try:
            # textord_noise_rejrows=0: Tesseract's noise-row test dropped whole lines of text,
            # a title included (see OcrService, 19 Oct 2026)
            r = subprocess.run([tess, "stdin", "stdout", "-l", lang, "--dpi", str(dpi), "-c", "textord_noise_rejrows=0",
                                "tsv"], input=img,
                               stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, env=env, timeout=120)
        except (subprocess.TimeoutExpired, OSError):
            return None
        return r.stdout.decode("utf-8", "replace") if r.returncode == 0 else None

    src = fitz.open(pdf_path)
    try:
        jobs = {}
        workers = max(1, min(4, (os.cpu_count() or 2) - 1))
        with ThreadPoolExecutor(max_workers=workers) as pool:
            for pno in page_numbers:
                pix = src[pno - 1].get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False)
                jobs[pno] = pool.submit(read, pix.tobytes("pgm"))
                del pix
            out = {}
            for pno, job in jobs.items():
                tsv = job.result()
                if tsv is not None:
                    out[pno] = {"lines": _tsv_pieces(tsv, 72.0 / dpi)}
            return out
    finally:
        src.close()


def _tsv_pieces(tsv, scale):
    """Tesseract TSV -> [{"text", "box"}]: each line cut where two words stand
    more than a line height apart."""
    lines = {}
    for row in tsv.splitlines()[1:]:
        f = row.split("\t")
        if len(f) < 12 or f[0] != "5":
            continue
        text = f[11].strip()
        try:
            conf = float(f[10])
            x, y, w, h = (int(v) for v in f[6:10])
        except ValueError:
            continue
        if not text or conf < 0:
            continue
        lines.setdefault((f[2], f[3], f[4]), []).append((x, y, x + w, y + h, text))
    pieces = []
    for words in lines.values():
        words.sort()
        hs = sorted(w[3] - w[1] for w in words)
        gap_max = 1.0 * hs[len(hs) // 2]
        group = []
        for w in words + [None]:
            # a ruling read as "|" or "[" is a cell border: it ends the piece
            rule = w is not None and _RULE_RE.match(w[4])
            if w is not None and not rule and (not group or w[0] - group[-1][2] <= gap_max):
                group.append(w)
                continue
            if not group:
                continue
            ar = sum(1 for g in group for c in g[4] if _is_arabic_char(c))
            lat = sum(1 for g in group for c in g[4] if c.isascii() and c.isalpha())
            ordered = sorted(group, key=lambda g: -g[0]) if ar > lat else group
            pieces.append({"text": " ".join(g[4] for g in ordered),
                           "box": [min(g[0] for g in group) * scale, min(g[1] for g in group) * scale,
                                   max(g[2] for g in group) * scale, max(g[3] for g in group) * scale]})
            group = [w] if w is not None and not rule else []
    return pieces


_RULE_RE = re.compile(r"^[|\[\]{}¦]+$")


# ── driver ─────────────────────────────────────────────────────────────────

def _table_merges(rows, layout):
    """Merged cells of a ruled table as [first row, first col, last row, last
    col] (0-based, in the row's reading order), with each merged cell's text
    moved to the first cell of its range, where Excel shows it. A header such
    as «2025» over Budget / Actual came back as «2025» beside an empty cell
    (2 Oct 2026); the layout PDF to Word builds its merged cells from was
    already there and was dropped on the way to Excel."""
    if not layout or len(layout) != len(rows):
        return []
    width = max(len(r) for r in rows)
    merges, open_v = [], {}
    for k, entries in enumerate(layout):
        if sum(e["span"] for e in entries) != width:
            return []
        pos = 0
        for e in entries:
            span = e["span"]
            if e["vmerge"] == "continue" and pos in open_v and open_v[pos][1] == span:
                open_v[pos][2] = k
            else:
                if pos in open_v:
                    merges.append(_closed(pos, open_v.pop(pos)))
                col = e["col"]
                if col != pos and 0 <= col < len(rows[k]) and not rows[k][pos]:
                    rows[k][pos], rows[k][col] = rows[k][col], ""
                if e["vmerge"] == "restart":
                    open_v[pos] = [k, span, k]
                elif span > 1:
                    merges.append([k, pos, k, pos + span - 1])
            pos += span
    merges += [_closed(pos, v) for pos, v in open_v.items()]
    return [m for m in merges if m[2] > m[0] or m[3] > m[1]]


def _header_at(rows):
    """Index of a table's header row when one or two rows of something else
    stand above it: the first row, among the first three, that fills most of
    the columns with labels (no numbers) and has numbers under it. 0 when the
    first row is the header or nothing says otherwise. The first row used to be
    styled as the header whatever it was."""
    if len(rows) < 3:
        return 0
    width = len(rows[0])
    for i in range(min(3, len(rows) - 1)):
        cells = [c for c in rows[i] if c]
        labels = len(cells) >= max(2, 0.6 * width) and not any(_is_numeric(c) for c in cells)
        if labels and any(_is_numeric(c) for c in rows[i + 1] if c):
            return i
    return 0


def _closed(pos, v):
    start, span, end = v
    return [start, pos, end, pos + span - 1]


def _running_lines(per_page, heights, page_count):
    """ids of the one-line text blocks a multi-page PDF repeats at the top or
    the bottom of most pages — a running head, "Page 3 of 12". In a sheet they
    are noise between the parts of one table, and they kept a table that runs
    over pages from being read as one (2 Oct 2026)."""
    if page_count < 2:
        return set()
    seen = {}
    for pno, blocks in per_page.items():
        h = heights.get(pno) or 0.0
        if h <= 0:
            continue
        for b in blocks:
            if b["kind"] != "text" or len(b["rows"]) != 1 or b.get("y0") is None:
                continue
            if 0.08 * h < b["y0"] < 0.90 * h:
                continue
            sig = re.sub(r"[0-9\u0660-\u0669]+", "#", " ".join(b["rows"][0]).strip())
            if sig:
                seen.setdefault(sig, []).append((pno, b))
    need = max(2, (page_count + 1) // 2)
    out = set()
    for occ in seen.values():
        if len({pno for pno, _ in occ}) >= need:
            out.update(id(b) for _, b in occ)
    return out


def extract_tables(input_pdf, output_json, language="auto"):
    doc = fitz.open(input_pdf)
    if doc.needs_pass:
        doc.close()
        # PdfToExcelService reads "code"; no page can be read without the password
        print(json.dumps({"success": False, "code": "passwordProtected",
                          "message": "This PDF is protected by an open password."}))
        sys.exit(3)
    per_page, texts, scanned, heights = {}, [], [], {}
    for pno, page in enumerate(doc, start=1):
        blocks, text = text_layer_page(page)
        if blocks is None:
            scanned.append(pno)
        else:
            heights[pno] = page.rect.height          # text-layer pages: y in points
            per_page[pno] = blocks
            texts.append(text)
    page_count = len(doc)
    doc.close()

    engine = "text-layer" if per_page else "none"
    if scanned and ocr_client is not None and ocr_client.is_available():
        lang = {"arabic": "ar", "ar": "ar", "english": "en", "en": "en"}.get(language, None)
        if lang is None:
            lang = osd_script(input_pdf, scanned[0])   # None = undecided: the old two-pass check below
        try:
            first = lang or "ar"
            got = ocr_pages(input_pdf, scanned, first)
            if lang is None and got:
                joined = " ".join(l.get("text", "") for p in got.values() for l in p.get("lines", []))
                ar, lat = _script_counts(joined)
                if lat > 2 * ar and lat >= 10:
                    got = ocr_pages(input_pdf, scanned, "en")
                    first = "en"
            for pno in scanned:
                op = got.get(pno)
                if not op:
                    per_page[pno] = []
                    continue
                joined = " ".join(l.get("text", "") for l in op.get("lines", []))
                texts.append(joined)
                ar, lat = _script_counts(joined)
                rtl = ar > lat
                blocks = build_blocks(ocr_seg_rows(op, rtl), rtl)
                for b in blocks:
                    b["rows"] = [[_norm_digits(c) for c in r] for r in b["rows"]]
                per_page[pno] = blocks
            engine = "ocr-" + first if not texts[: len(texts) - len(scanned)] else "mixed"
        except Exception as e:                        # engine down mid-call: keep what we have
            sys.stderr.write(f"OCR failed: {e}\n")
    # the pages the service did not read (it is down, absent or failed): Tesseract
    left = [pno for pno in scanned if not per_page.get(pno)]
    if left:
        tlang = {"arabic": "ara+eng", "ar": "ara+eng", "english": "eng", "en": "eng"}.get(language, "eng+ara")
        try:
            got = tesseract_pages(input_pdf, left, tlang)
        except Exception as e:  # noqa: BLE001 - those pages stay empty
            sys.stderr.write(f"Tesseract failed: {e}\n")
            got = {}
        for pno in left:
            op = got.get(pno)
            if not op:
                continue
            joined = " ".join(l.get("text", "") for l in op.get("lines", []))
            texts.append(joined)
            ar, lat = _script_counts(joined)
            rtl = ar > lat
            blocks = build_blocks(ocr_seg_rows(op, rtl), rtl)
            for b in blocks:
                # Tesseract hands digits over in reading order: mapped, never flipped
                b["rows"] = [[c.translate(_AR2WE) for c in r] for r in b["rows"]]
            per_page[pno] = blocks
        if got:
            engine = "ocr-tesseract" if len(left) == page_count else "mixed"
    for pno in scanned:
        per_page.setdefault(pno, [])

    all_text = " ".join(texts)
    ar, lat = _script_counts(all_text)
    is_rtl = ar > lat

    running = _running_lines(per_page, heights, page_count)

    pages_out, total_rows, total_tables = [], 0, 0
    for pno in range(1, page_count + 1):
        blocks = per_page.get(pno, [])
        clean_blocks = []
        for b in blocks:
            if id(b) in running:
                continue
            rows = [[_clean(c) for c in r] for r in b["rows"]]
            merges = _table_merges(rows, b.get("layout")) if b["kind"] == "table" else []
            keep = [i for i, r in enumerate(rows) if any(r)]
            if not keep:
                continue
            if len(keep) < len(rows):
                at = {old: new for new, old in enumerate(keep)}
                merges = [[at[r0], c0, at[r1], c1] for r0, c0, r1, c1 in merges if r0 in at and r1 in at]
                rows = [rows[i] for i in keep]
            width = max(len(r) for r in rows)
            for r in rows:
                r += [""] * (width - len(r))
            lead = _header_at(rows) if b["kind"] == "table" and not merges else 0
            if lead:
                # rows above the header ("Account: 0012-3456-789 …" over a statement)
                # are text before the table; the header row is the table's first
                clean_blocks.append({"kind": "text", "rows": [[c for c in r if c] for r in rows[:lead]],
                                     "headerRow": -1})
                total_rows += lead
                rows = rows[lead:]
            out = {"kind": b["kind"], "rows": rows, "headerRow": 0 if b["kind"] == "table" else -1}
            if merges:
                out["merges"] = merges
            clean_blocks.append(out)
            total_rows += len(rows)
            if b["kind"] == "table":
                total_tables += 1
        flat = [r for b in clean_blocks for r in b["rows"]]
        pages_out.append({"pageNumber": pno, "blocks": clean_blocks, "rows": flat,
                          "headerRow": 0 if clean_blocks and clean_blocks[0]["kind"] == "table" else -1})

    result = {"success": True, "isRTL": is_rtl, "engine": engine, "pages": pages_out,
              "totalRows": total_rows, "totalTables": total_tables}
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)
    print(json.dumps({"success": True, "isRTL": is_rtl, "engine": engine, "totalTables": total_tables,
                      "message": f"Extracted {total_rows} rows / {total_tables} tables from {page_count} pages ({engine})"},
                     ensure_ascii=False))


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"success": False, "message": "Usage: pdf_table_extractor.py <input.pdf> <output.json> [arabic|english|auto]"}))
        sys.exit(1)
    input_pdf, output_json = sys.argv[1], sys.argv[2]
    language = sys.argv[3] if len(sys.argv) > 3 else "auto"
    try:
        extract_tables(input_pdf, output_json, language)
    except Exception as e:
        print(json.dumps({"success": False, "message": f"Extraction failed: {e}\n{traceback.format_exc()}"}))
        sys.exit(1)


if __name__ == "__main__":
    main()
