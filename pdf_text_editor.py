#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Text Editor using PyMuPDF (fitz) — extract editable text lines with
positions, replace only what the user changed.

Rewritten 25 Sept 2026 after the second audit (the editor was the weakest PDF
tool, 9/10):

  * Every edit re-typeset the WHOLE line from the extracted text, and that text
    had lost the letters Chrome and LibreOffice draw inside /ActualText: editing
    a name in a Chrome lease rewrote the line as «الطر ال ا: محمد رباه م عل».
    Extraction is fixed in pdf_to_word.py (shared with PDF→Word); an edit now
    diffs the old and new line word by word and removes and redraws ONLY the
    words that changed. The other words keep their original glyphs, font and
    position (the page's content stream is re-emitted by MuPDF's redaction
    filter, so untouched text is equal operator for operator, not byte for byte).

  * Replacement text is set by pdf_text_layer.py: HarfBuzz shaping, the face that
    matches the original's script, weight, slant and class (Liberation / DejaVu
    for Latin, IBM Plex Sans Arabic / Noto Naskh / Amiri / Noto Kufi for Arabic)
    and a text layer that extracts in logical order. Before, Arabic went through
    insert_htmlbox, whose text layer read «ﻪğﺤ4Ěﺻ3» for «صفحة 3 من 20», and Latin
    through base-14 fonts: "İstanbul ğüş Łódź €" became "?stanbul ?ü?? ?ód? ?".
    When every new character already exists in the ORIGINAL embedded font (its
    ToUnicode map says which), a Latin replacement is written with that very font.

  * A longer replacement first takes the room the line leaves free, then the
    rest of the line is re-set with it, then the type shrinks (to 60%), then it
    wraps into free space below, and only then shrinks further: never past the
    page edge (the old box grew 50% and ran to x=599 on a 595 pt page).

  * Lines are rows: MuPDF splits a Chrome line wherever a glyph cluster lost its
    width, and Latin spans used to be one element per word; fragments on one
    baseline of a block are one editable line now. Rotated pages (/Rotate 90,
    180, 270) are handled in DISPLAYED coordinates, which is what the page
    preview shows.

Usage:
    python3 pdf_text_editor.py extract <input.pdf> <output.json>
    python3 pdf_text_editor.py replace <input.pdf> <output.pdf> <edits.json>
"""

import json
import os
import re
import sys
import time

import fitz  # PyMuPDF

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import pdf_geometry as geom  # noqa: E402
from pdf_to_word import (  # noqa: E402
    _line_chars, _merge_rows, _finish_row, _normalize_presentation_forms, _is_arabic_char, _tidy,
)

try:
    import pdf_text_layer as ptl  # noqa: E402
except Exception as _e:  # noqa: BLE001 — the legacy writer below still works
    ptl = None
    sys.stderr.write(f"pdf_text_layer unavailable ({_e}); using the legacy writer\n")

_AR_RE = re.compile(u"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]")
_FLAGS = fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE


def _has_arabic(t):
    return bool(_AR_RE.search(t or ""))


# ── the page model shared by extract and replace ───────────────────────────

def _displayed_raw(page):
    """rawdict in DISPLAYED coordinates: on a /Rotate 90 page MuPDF reports a
    line that reads left to right on screen as running up the unrotated page,
    and every x-based rule of the row rebuild would read it as one column."""
    raw = page.get_text("rawdict", flags=_FLAGS)
    if not page.rotation:
        return raw
    m = page.rotation_matrix
    o = fitz.Point(0, 0) * m

    def rect(b):
        return tuple(fitz.Rect(b) * m)

    def pt(p):
        return tuple(fitz.Point(p) * m)
    for b in raw.get("blocks", []):
        b["bbox"] = rect(b["bbox"])
        for ln in b.get("lines", []):
            ln["bbox"] = rect(ln["bbox"])
            d = fitz.Point(ln["dir"]) * m - o
            ln["dir"] = (d.x, d.y)
            for sp in ln.get("spans", []):
                sp["bbox"] = rect(sp["bbox"])
                if "origin" in sp:
                    sp["origin"] = pt(sp["origin"])
                for ch in sp.get("chars", []):
                    ch["bbox"] = rect(ch["bbox"])
                    ch["origin"] = pt(ch["origin"])
    return raw


def page_rows(page):
    """-> (page_rtl, [row]) — a row is pdf_to_word's row dict plus id/bbox/base."""
    raw = _displayed_raw(page)
    blocks = [b for b in raw.get("blocks", []) if b.get("type") == 0]
    all_chars = "".join(ch.get("c", "") for b in blocks for ln in b["lines"] for sp in ln["spans"] for ch in sp["chars"])
    ar_n = sum(1 for c in all_chars if _is_arabic_char(c))
    lat_n = sum(1 for c in all_chars if c.isascii() and c.isalpha())
    page_rtl = ar_n > lat_n
    merged = []
    for bi, b in enumerate(blocks):
        frags = []
        for ln in b["lines"]:
            d = ln.get("dir") or (1, 0)
            if not (d[0] > 0.99 and abs(d[1]) < 0.01):
                continue                         # vertical / tilted text is not editable here
            chars = _line_chars(ln)
            if not any(c[0] not in "  " for c in chars):
                continue
            frags.append({"chars": chars, "y0": ln["bbox"][1], "y1": ln["bbox"][3],
                          "x0": ln["bbox"][0], "x1": ln["bbox"][2], "block": bi})
        for ri, r in enumerate(_merge_rows(frags)):
            xs = [c[1] for c in r["chars"]] + [c[2] for c in r["chars"]]
            r.update(x0=min(xs), x1=max(xs), ri=ri)
            merged.append(r)
    # Words set later on the same line (an earlier edit, a form filler) sit in a
    # block of their own; a word-sized gap on the same baseline joins them again.
    # Table cells are further apart and stay separate.
    merged.sort(key=lambda r: (r["y0"], r["x0"]))
    joined = []
    for r in merged:
        size = max((c[4] for c in r["chars"]), default=10.0) or 10.0
        mate = None
        for q in reversed(joined):
            if q["y1"] < r["y0"] - 2 * size:
                break                          # rows are sorted by top: the rest are higher up
            qb, rb = q.get("base"), r.get("base")
            same_base = (qb is not None and rb is not None and abs(qb[0] - rb[0]) < 0.3 * size) or \
                        (qb is None or rb is None) and abs((q["y0"] + q["y1"]) - (r["y0"] + r["y1"])) < 0.3 * size
            gap = max(r["x0"] - q["x1"], q["x0"] - r["x1"])
            if same_base and gap < 0.8 * size and q["block"] != r["block"]:
                mate = q
                break
        if mate is None:
            joined.append(r)
        else:
            mate["chars"].extend(r["chars"])
            mate.update(x0=min(mate["x0"], r["x0"]), x1=max(mate["x1"], r["x1"]),
                        y0=min(mate["y0"], r["y0"]), y1=max(mate["y1"], r["y1"]))
    rows = []
    for r in joined:
        whole = _finish_row(r, page_rtl)
        if not whole:
            continue
        # a gap of more than 3 em on one baseline separates two things (page
        # corners, a label and its value across a form): two editable lines
        groups, prev = [[]], None
        for t in whole["tokens"]:
            if prev is not None:
                gap = (prev["x0"] - t["x1"]) if whole["rtl"] else (t["x0"] - prev["x1"])
                if gap > 3.0 * max(t["size"], prev["size"], 1.0):
                    groups.append([])
            groups[-1].append(t)
            prev = t
        for gi, toks in enumerate(groups):
            row = dict(whole, tokens=toks, x0=min(t["x0"] for t in toks), x1=max(t["x1"] for t in toks))
            row["id"] = f"p{page.number + 1}_b{r['block']}_r{r['ri']}" + (f"_g{gi}" if len(groups) > 1 else "")
            row["text"] = _row_text(row)
            if not row["text"].strip():
                continue
            spans = [c[3] for t in toks for c in t["chars"] if isinstance(c[3], dict)]
            base = [sp["origin"][1] for sp in spans if sp.get("origin")]
            row["base"] = max(set(base), key=base.count) if base else row["y1"] - 0.2 * (row["y1"] - row["y0"])
            rows.append(row)
    return page_rtl, rows


def _row_text(row):
    return _tidy(" ".join(_normalize_presentation_forms(t["text"]) for t in row["tokens"] if t["text"]))


def _norm(s):
    return _normalize_presentation_forms(s).replace(" ", "")


def _row_words(row):
    """The words the user sees (row text split on spaces), each with the tokens
    it is made of. None when the two do not line up."""
    toks = [t for t in row["tokens"] if t["text"]]
    words = row["text"].split()
    out, k = [], 0
    for w in words:
        target, acc, idx = _norm(w), "", []
        while k < len(toks) and len(acc) < len(target):
            acc += _norm(toks[k]["text"])
            idx.append(k)
            k += 1
        if acc != target:
            return None
        out.append((w, [toks[i] for i in idx]))
    return out if k == len(toks) else None


def _dominant_span(tokens):
    weights = {}
    for t in tokens:
        for c in t["chars"]:
            sp = c[3]
            if isinstance(sp, dict) and c[0] not in "  ":
                k = id(sp)
                weights[k] = (weights.get(k, (0, sp))[0] + 1, sp)
    return max(weights.values(), key=lambda v: v[0])[1] if weights else {}


def _style(span):
    font = span.get("font") or ""
    f = font.lower()
    flags = int(span.get("flags", 0))
    return {
        "font": font,
        "size": float(span.get("size") or 12.0),
        "color": int(span.get("color", 0)),
        "bold": bool(flags & 16) or any(w in f for w in ("bold", "black", "heavy", "semibold", "demibold")),
        "italic": bool(flags & 2) or any(w in f for w in ("italic", "oblique")),
    }


# ── extraction ─────────────────────────────────────────────────────────────

def extract_text_blocks(input_pdf, output_json):
    t0 = time.time()
    doc = fitz.open(input_pdf)
    pages_data = []
    for page in doc:
        _rtl, rows = page_rows(page)
        page_data = {"pageNumber": page.number + 1, "width": page.rect.width, "height": page.rect.height,
                     "rotation": page.rotation, "textBlocks": []}
        for row in rows:
            x0 = min(t["x0"] for t in row["tokens"])
            x1 = max(t["x1"] for t in row["tokens"])
            st = _style(_dominant_span(row["tokens"]))
            page_data["textBlocks"].append({
                "id": row["id"], "text": row["text"],
                "x": round(x0, 2), "y": round(row["y0"], 2),
                "width": round(x1 - x0, 2), "height": round(row["y1"] - row["y0"], 2),
                "fontSize": round(st["size"], 2), "fontFamily": st["font"], "color": f"#{st['color']:06x}",
                "pageNumber": page.number + 1, "isRTL": bool(row["rtl"]),
                "bold": st["bold"], "italic": st["italic"],
            })
        pages_data.append(page_data)
    doc.close()
    result = {"success": True, "pages": pages_data, "totalBlocks": sum(len(p["textBlocks"]) for p in pages_data)}
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False)
    print(json.dumps({"success": True, "message": f"Extracted {result['totalBlocks']} text lines",
                      "seconds": round(time.time() - t0, 2)}))


# ── replacement ────────────────────────────────────────────────────────────

def _hex_color(c):
    try:
        c = (c or "#000000").lstrip("#")
        return tuple(int(c[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except Exception:  # noqa: BLE001
        return (0, 0, 0)


def _int_color(c):
    return ((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255


def _to_unrotated(page, rect):
    return fitz.Rect(rect) * ~page.rotation_matrix if page.rotation else fitz.Rect(rect)


def _find_row(rows, edit):
    """The row an edit refers to: by id when its text still matches, else the
    row with the same text nearest to the edit's box."""
    orig = (edit.get("originalText") or edit.get("text") or "").strip()
    by_id = next((r for r in rows if r["id"] == edit.get("id")), None)
    if by_id is not None and (not orig or _norm(by_id["text"]) == _norm(orig)):
        return by_id
    same = [r for r in rows if orig and _norm(r["text"]) == _norm(orig)]
    if same:
        x, y = float(edit.get("x", 0)), float(edit.get("y", 0))
        return min(same, key=lambda r: abs(r["y0"] - y) + abs(min(t["x0"] for t in r["tokens"]) - x))
    return None


def _changed_range(old_words, new_words):
    """Common prefix / suffix at the word level -> (a, old_end, new_end)."""
    n, m = len(old_words), len(new_words)
    a = 0
    while a < n and a < m and _norm(old_words[a]) == _norm(new_words[a]):
        a += 1
    b = 0
    while b < n - a and b < m - a and _norm(old_words[n - 1 - b]) == _norm(new_words[m - 1 - b]):
        b += 1
    return a, n - b, m - b


def _free_limits(page, rows, row, region, rtl):
    """How far a replacement may run past its region in the reading direction:
    up to the next text on the same band, the page's text column, and never
    closer than 6 pt to the page edge."""
    pw = page.rect.width
    band0, band1 = row["y0"] + 0.2 * (row["y1"] - row["y0"]), row["y1"] - 0.2 * (row["y1"] - row["y0"])
    others = [r for r in rows if r is not row and r["y1"] > band0 and r["y0"] < band1]
    xs0 = [min(t["x0"] for t in r["tokens"]) for r in rows]
    xs1 = [max(t["x1"] for t in r["tokens"]) for r in rows]
    if not rtl:
        limit = min(pw - 6.0, max(max(xs1, default=region.x1), region.x1))
        for r in others:
            rx0 = min(t["x0"] for t in r["tokens"])
            if rx0 >= region.x1 - 0.5:
                limit = min(limit, rx0 - 4.0)
        return region.x0, max(limit, region.x1)
    limit = max(6.0, min(min(xs0, default=region.x0), region.x0))
    for r in others:
        rx1 = max(t["x1"] for t in r["tokens"])
        if rx1 <= region.x0 + 0.5:
            limit = max(limit, rx1 + 4.0)
    return min(limit, region.x0), region.x1


def _room_below(page, rows, row, x0, x1, line_h):
    """Extra lines that fit under the row without touching other text."""
    ph = page.rect.height
    y = row["y1"]
    n = 0
    while n < 6:
        top, bot = y + n * line_h, y + (n + 1) * line_h
        if bot > ph - 6:
            break
        hit = any(r is not row and r["y0"] < bot and r["y1"] > top
                  and min(t["x0"] for t in r["tokens"]) < x1 and max(t["x1"] for t in r["tokens"]) > x0
                  for r in rows)
        if hit:
            break
        n += 1
    return n


# ---- reuse of the original embedded font (Latin, no shaping needed) ----

_HEX = re.compile(r"<([0-9A-Fa-f]+)>")


def _parse_tounicode(data):
    """bfchar / bfrange of a ToUnicode CMap -> {code(int): str}."""
    text = data.decode("latin-1", "ignore")
    out = {}
    for blk in re.findall(r"beginbfchar(.*?)endbfchar", text, re.S):
        for src, dst in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>", blk):
            try:
                out[int(src, 16)] = bytes.fromhex(dst).decode("utf-16-be", "ignore") if dst else ""
            except ValueError:
                pass
    for blk in re.findall(r"beginbfrange(.*?)endbfrange", text, re.S):
        for m in re.finditer(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(\[[^\]]*\]|<[0-9A-Fa-f]+>)", blk):
            lo, hi = int(m.group(1), 16), int(m.group(2), 16)
            if hi - lo > 4096:
                continue
            dst = m.group(3)
            if dst.startswith("["):
                for k, h in enumerate(_HEX.findall(dst)):
                    out[lo + k] = bytes.fromhex(h).decode("utf-16-be", "ignore")
            else:
                base = bytes.fromhex(dst[1:-1])
                for k in range(hi - lo + 1):
                    b = bytearray(base)
                    b[-1] = (b[-1] + k) & 255
                    out[lo + k] = bytes(b).decode("utf-16-be", "ignore")
    return out


def _parse_widths_W(s):
    """/W array of a CIDFont -> {cid: width}."""
    toks = re.findall(r"\[|\]|-?\d+(?:\.\d+)?", s or "")
    out, i = {}, 1 if toks[:1] == ["["] else 0
    while i < len(toks):
        if toks[i] == "]":
            break
        c = int(float(toks[i]))
        if i + 1 < len(toks) and toks[i + 1] == "[":
            j = i + 2
            while j < len(toks) and toks[j] != "]":
                out[c] = float(toks[j])
                c += 1
                j += 1
            i = j + 1
        elif i + 2 < len(toks):
            c2, w = int(float(toks[i + 1])), float(toks[i + 2])
            for k in range(c, c2 + 1):
                out[k] = w
            i += 3
        else:
            break
    return out


_FONT_CODES = {}


def _font_codes(doc, page, span_font):
    """The page's own font resources named like the span, readable for reuse:
    [(resource name, bytes per code, {char: code}, {code: width/1000})]."""
    key = (id(doc), page.number, span_font)
    if key in _FONT_CODES:
        return _FONT_CODES[key]
    found = []
    for (xref, _ext, ftype, basefont, resname, _enc, referencer) in page.get_fonts(full=True):
        if basefont.split("+")[-1] != span_font or ftype == "Type3" or referencer:
            continue
        tu = doc.xref_get_key(xref, "ToUnicode")
        if tu[0] != "xref":
            continue
        try:
            code_to = _parse_tounicode(doc.xref_stream(int(tu[1].split()[0])))
            inv = {}
            for code, s in code_to.items():
                if len(s) == 1:
                    inv.setdefault(s, code)
            if ftype == "Type0":
                desc = doc.xref_get_key(xref, "DescendantFonts")
                dxref = int(re.findall(r"(\d+) 0 R", desc[1])[0])
                dw = doc.xref_get_key(dxref, "DW")
                default = float(dw[1]) if dw[0] in ("int", "float") else 1000.0
                w = doc.xref_get_key(dxref, "W")
                if w[0] == "xref":
                    w = ("array", doc.xref_object(int(w[1].split()[0])))
                table = _parse_widths_W(w[1]) if w[0] == "array" else {}
                widths = {code: table.get(code, default) / 1000.0 for code in inv.values()}
                found.append((resname, 2, inv, widths))
            else:
                fc = doc.xref_get_key(xref, "FirstChar")
                wl = doc.xref_get_key(xref, "Widths")
                if wl[0] == "xref":
                    wl = ("array", doc.xref_object(int(wl[1].split()[0])))
                if fc[0] != "int" or wl[0] != "array":
                    continue
                first = int(fc[1])
                nums = [float(x) for x in re.findall(r"-?\d+(?:\.\d+)?", wl[1])]
                inv = {c: code for c, code in inv.items() if code <= 255}
                widths = {code: (nums[code - first] if 0 <= code - first < len(nums) else 0.0) / 1000.0
                          for code in inv.values()}
                found.append((resname, 1, inv, widths))
        except Exception:  # noqa: BLE001 — an unreadable resource is simply not reused
            continue
    _FONT_CODES[key] = found
    return found


def _reuse_font(doc, page, span_font, text):
    """The original font resource, when it can draw every character of `text`
    (its ToUnicode gives each a code: the glyph is in the embedded subset) and
    the text needs no shaping. A space the subset lacks is left as a gap."""
    if ptl is not None and ptl.needs_shaping(text):
        return None
    want = set(text) - {" "}
    for cand in _font_codes(doc, page, span_font):
        if want <= set(cand[2]):
            return cand
    return None


def _space_w(reuse):
    inv, widths = reuse[2], reuse[3]
    return widths.get(inv[" "], 0.28) if " " in inv else 0.28


def _draw_reused(page, reuse, text, x, base, size, color):
    resname, nbytes, inv, widths = reuse
    m = geom.displayed_to_pdf(page)  # exact on rotated pages whose box origin is not (0,0)
    o = fitz.Point(0, 0) * m
    ex = fitz.Point(1, 0) * m - o
    ey = fitz.Point(0, -1) * m - o
    r, g, b = color
    ops = ["q", "BT", "0 Tc 0 Tw 100 Tz 0 Ts 0 Tr", "%.4g %.4g %.4g rg" % (r, g, b), "/%s %.3f Tf" % (resname, size)]
    pen = x
    for k, word in enumerate(text.split(" ")):
        if k:
            pen += _space_w(reuse) * size
        if not word:
            continue
        p = fitz.Point(pen, base) * m
        ops.append("%.5g %.5g %.5g %.5g %.3f %.3f Tm <%s> Tj" % (
            ex.x, ex.y, ey.x, ey.y, p.x, p.y, "".join("%0*x" % (2 * nbytes, inv[c]) for c in word)))
        pen += sum(widths.get(inv[c], 0.5) for c in word) * size
    ops += ["ET", "Q"]
    ptl.append_content(page, ("\n".join(ops) + "\n").encode())


def _reused_width(reuse, text, size):
    return (sum(reuse[3].get(reuse[2][c], 0.5) for c in text if c != " ") + text.count(" ") * _space_w(reuse)) * size


# ---- the legacy writer (kept as the fallback) ----

_FONT_DIRS = [os.path.join(HERE, "..", "Fonts"), os.path.join(HERE, "Fonts"), "/usr/share/fonts/google-noto",
              "/usr/share/fonts/confileo", "/usr/share/fonts/truetype/noto", "/usr/share/fonts/amiri-fonts",
              os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")]


def _find_font(*names):
    for d in _FONT_DIRS:
        for n in names:
            p = os.path.join(d, n)
            if os.path.exists(p):
                return os.path.abspath(p)
    return None


def _legacy_write(page, rect, text, size, color, family, rtl):
    """The pre-rewrite path: MuPDF's HTML box for Arabic, a base-14 face otherwise.

    Both calls map their rectangle to the file themselves, and on a rotated page whose box
    does not start at (0,0) both landed away from the word they replace (10 Oct 2026). At
    /Rotate 0 they are exact, so they run on the page turned upright (pdf_geometry.unrotated),
    in unrotated coordinates, with the text turned by the page's rotation instead."""
    bold = any(k in (family or "").lower() for k in ("bold", "black", "heavy"))
    rotation = page.rotation
    rect = _to_unrotated(page, rect)
    with geom.unrotated(page):
        if _has_arabic(text):
            fontfile = _find_font(*((["IBMPlexSansArabic-Bold.ttf"] if bold else []) + ["IBMPlexSansArabic-Regular.ttf",
                                                                                       "NotoNaskhArabic-Regular.ttf"]))
            css, fam, arch = "", "sans-serif", None
            if fontfile:
                fam = "confileoAr"
                css = f"@font-face {{ font-family: {fam}; src: url({os.path.basename(fontfile)}); }}"
                arch = fitz.Archive(os.path.dirname(fontfile))
            r, g, b = (int(c * 255) for c in color)
            esc = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            html = (f'<div dir="{"rtl" if rtl else "ltr"}" style="font-family:{fam};font-size:{size:.2f}pt;'
                    f'color:rgb({r},{g},{b});line-height:1.15;margin:0;{"font-weight:bold;" if bold else ""}">{esc}</div>')
            try:
                page.insert_htmlbox(rect, html, css=css, archive=arch, scale_low=0.3, rotate=rotation)
                return
            except Exception as e:  # noqa: BLE001
                sys.stderr.write(f"htmlbox failed: {e}\n")
        s = size
        while s >= 3:
            if page.insert_textbox(rect, text, fontsize=s, fontname="hebo" if bold else "helv", color=color,
                                   align=2 if rtl else 0, rotate=rotation) >= 0:
                break
            s -= 0.5


# ---- one edit ----

class _Plan:
    """Where and how one edit's new text goes (displayed coordinates)."""

    def __init__(self):
        self.remove = []          # displayed rects to redact
        self.lines = []           # (text, x_left, baseline, size, reuse or stack, style, rtl, width)
        self.legacy = None        # (rect, text, size, color, family, rtl)


def _plan_edit(doc, page, rows, edit):
    plan = _Plan()
    new_text = " ".join((edit.get("newText") or "").split())[:4000]
    row = _find_row(rows, edit)
    if row is None:
        # the edit's own box (a PDF that changed since extraction, or the old client)
        x, y = float(edit.get("x", 0)), float(edit.get("y", 0))
        w, h = float(edit.get("width", 100)), float(edit.get("height", 20))
        rect = fitz.Rect(x - 0.5, y + 0.12 * h, x + w + 0.5, y + h - 0.12 * h)
        plan.remove.append(rect)
        if new_text:
            plan.legacy = (fitz.Rect(x, y, x + w, y + h) & page.rect, new_text, float(edit.get("fontSize") or 12),
                           _hex_color(edit.get("color")), edit.get("fontFamily") or "", bool(edit.get("isRTL")))
        return plan
    rtl = bool(row["rtl"])
    words = _row_words(row)
    new_words = new_text.split()
    if words is None:
        a, old_end, new_end = 0, None, len(new_words)
        changed_tokens = [t for t in row["tokens"] if t["text"]]
        old_words = None
    else:
        old_words = [w for w, _t in words]
        a, old_end, new_end = _changed_range(old_words, new_words)
        if a == old_end and a == new_end:
            return plan                                   # nothing changed
        if a == old_end:                                  # pure insertion: take a neighbour along
            if old_end < len(old_words):
                old_end += 1
                new_end += 1
            else:
                a -= 1
        changed_tokens = [t for _w, ts in words[a:old_end] for t in ts]
    # visual contiguity: an unchanged word between changed ones (mixed-direction
    # lines) means the whole line is re-set
    if words is not None:
        cx0 = min(t["x0"] for t in changed_tokens)
        cx1 = max(t["x1"] for t in changed_tokens)
        between = [t for _w, ts in (words[:a] + words[old_end:]) for t in ts
                   if cx0 + 0.5 < (t["x0"] + t["x1"]) / 2.0 < cx1 - 0.5]
        if between:
            a, old_end, new_end = 0, len(old_words), len(new_words)
            changed_tokens = [t for _w, ts in words for t in ts]
    span = _dominant_span(changed_tokens)
    st = _style(span)
    size = st["size"]
    color = _int_color(st["color"])
    base = row["base"]

    def region_of(tokens):
        return fitz.Rect(min(t["x0"] for t in tokens), row["y0"], max(t["x1"] for t in tokens), row["y1"])

    def text_for(k0, k1):
        return " ".join(new_words[k0:k1]) if old_words is not None else new_text

    region = region_of(changed_tokens)
    seg = text_for(a, new_end)
    stack = ptl.font_stack(st["font"], st["bold"], st["italic"], seg or new_text) if ptl is not None else []
    usable = ptl is not None and stack and ptl.available(new_text)

    def measure(text, sz):
        reuse = _reuse_font(doc, page, st["font"], text) if usable else None
        if reuse:
            return reuse, _reused_width(reuse, text, sz)
        line = ptl.layout(text, stack, sz, base_rtl=rtl)
        return None, line.width

    if not usable:
        plan.remove.append(fitz.Rect(region.x0, row["y0"] + 0.12 * region.height, region.x1, row["y1"] - 0.12 * region.height))
        if seg:
            plan.legacy = (fitz.Rect(region.x0, row["y0"], region.x1, row["y1"]), seg, size, color, st["font"], rtl)
        return plan

    tail_tokens = []
    if seg:
        reuse, wn = measure(seg, size)
        lo, hi = _free_limits(page, rows, row, region, rtl)
        if words is not None and old_end < len(old_words):
            # room up to the next unchanged word only
            nxt = [t for _w, ts in words[old_end:] for t in ts]
            if rtl:
                lo = max(lo, max(t["x1"] for t in nxt) + 0.3 * size)
            else:
                hi = min(hi, min(t["x0"] for t in nxt) - 0.3 * size)
        if wn > (hi - lo) + 0.5 and words is not None and old_end < len(old_words):
            # take the rest of the line along and re-set it after the new words
            tail_tokens = [t for _w, ts in words[old_end:] for t in ts]
            changed_tokens = changed_tokens + tail_tokens
            region = region_of(changed_tokens)
            seg = text_for(a, len(new_words))
            reuse, wn = measure(seg, size)
            lo, hi = _free_limits(page, rows, row, region, rtl)
        avail = max(hi - lo, 1.0)
        sz = size
        lines_out = [(seg, reuse, wn)]
        if wn > (avail + 0.5) / 0.85 and old_words is not None:
            flow = _reflow(page, rows, row, new_words, stack, st, rtl)
            if flow is not None:
                plan.lines.extend(flow[0])
                plan.remove.extend(flow[1])
                return plan
        if wn > avail + 0.5:
            sz = size * avail / wn
            if sz < 0.6 * size:
                line_h = (row["y1"] - row["y0"]) or size * 1.2
                extra = _room_below(page, rows, row, lo, hi, line_h)
                wrapped = None
                if extra:
                    for trial in (size, 0.9 * size, 0.8 * size, 0.7 * size, 0.6 * size):
                        wl = ptl.wrap(seg, stack, trial, avail, base_rtl=rtl)
                        if len(wl) <= extra + 1 and max(l.width for l in wl) <= avail + 0.5:
                            wrapped = (trial, wl)
                            break
                if wrapped:
                    sz, wl = wrapped
                    lines_out = [(l.text, None, l.width) for l in wl]
                else:
                    sz = max(size * avail / wn, 0.3 * size)
                    lines_out = [(seg, reuse, wn * sz / size)]
            else:
                lines_out = [(seg, reuse, wn * sz / size)]
        line_h = (row["y1"] - row["y0"]) * sz / size if len(lines_out) > 1 else 0
        for k, (text, reuse_k, width) in enumerate(lines_out):
            if len(lines_out) > 1 or sz != size:
                reuse_k, width = measure(text, sz)
            x_left = (hi - width) if rtl else lo
            if rtl and len(lines_out) == 1 and not tail_tokens:
                x_left = region.x1 - width            # keep the words' right edge
            x_left = min(max(x_left, 6.0), page.rect.width - 6.0 - width)
            plan.lines.append((text, x_left, base + k * line_h * 1.15, sz, reuse_k, stack, st, rtl, width))
    for t in changed_tokens:
        plan.remove.append(fitz.Rect(t["x0"], row["y0"] + 0.12 * (row["y1"] - row["y0"]),
                                     t["x1"], row["y1"] - 0.12 * (row["y1"] - row["y0"])))
    return plan


def _reflow(page, rows, row, new_words, stack, st, rtl):
    """A replacement too long for its line wraps within its paragraph: the
    edited line and the lines after it in the same block are set again at the
    block's width and line pitch (then 5% smaller at a time, down to 70%),
    using the room under the block when there is some. Returns (lines, rects)
    or None when the paragraph has no room for it."""
    para = [row]
    for r in sorted((r for r in rows if r is not row and r.get("block") == row.get("block")
                     and r["y0"] > row["y0"] + 0.5 * (row["y1"] - row["y0"])), key=lambda r: r["y0"]):
        last = para[-1]
        if r["y0"] - last["y1"] > 1.2 * (last["y1"] - last["y0"]):
            break
        if min(r["x1"], row["x1"]) - max(r["x0"], row["x0"]) < 0.3 * (row["x1"] - row["x0"]):
            break
        if (last["x1"] - last["x0"]) < 0.85 * (row["x1"] - row["x0"]) and last is not row:
            break                          # the line before was the paragraph's last
        para.append(r)
    x0 = min(r["x0"] for r in para)
    x1 = max(r["x1"] for r in para)
    width = x1 - x0
    size = st["size"]
    bases = [r["base"] for r in para]
    pitch = (bases[-1] - bases[0]) / (len(bases) - 1) if len(bases) > 1 else 1.2 * size
    pitch = max(pitch, 1.05 * size)
    extra = _room_below(page, rows, para[-1], x0, x1, pitch)
    text = " ".join(new_words + [w for r in para[1:] for w in r["text"].split()])
    for f in (1.0, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7):
        sz = size * f
        wl = ptl.wrap(text, stack, sz, width, base_rtl=rtl)
        if len(wl) <= len(para) + extra and max(l.width for l in wl) <= width + 0.5:
            lines = []
            for k, l in enumerate(wl):
                xl = (x1 - l.width) if rtl else x0
                lines.append((l.text, xl, bases[0] + k * pitch * f, sz, None, stack, st, rtl, l.width))
            rects = [fitz.Rect(t["x0"], r["y0"] + 0.12 * (r["y1"] - r["y0"]), t["x1"], r["y1"] - 0.12 * (r["y1"] - r["y0"]))
                     for r in para for t in r["tokens"]]
            return lines, rects
    return None


def replace_text(input_pdf, output_pdf, edits_json):
    t0 = time.time()
    with open(edits_json, "r", encoding="utf-8") as f:
        edits = json.load(f)
    doc = fitz.open(input_pdf)
    by_page = {}
    for edit in edits:
        by_page.setdefault(int(edit.get("pageNumber", 1)) - 1, []).append(edit)
    writer = ptl.TextLayerWriter(doc) if ptl is not None else None
    applied = 0
    report = []
    for pno, page_edits in sorted(by_page.items()):
        if pno < 0 or pno >= len(doc):
            continue
        page = doc[pno]
        _rtl, rows = page_rows(page)
        plans = []
        for edit in page_edits:
            try:
                plans.append((edit, _plan_edit(doc, page, rows, edit)))
            except Exception as e:  # noqa: BLE001 — this edit falls back to the legacy writer
                sys.stderr.write(f"edit {edit.get('id')} planning failed: {e}\n")
                x, y = float(edit.get("x", 0)), float(edit.get("y", 0))
                w, h = float(edit.get("width", 100)), float(edit.get("height", 20))
                p = _Plan()
                p.remove.append(fitz.Rect(x - 0.5, y + 0.12 * h, x + w + 0.5, y + h - 0.12 * h))
                p.legacy = (fitz.Rect(x, y, x + w, y + h), " ".join((edit.get("newText") or "").split()),
                            float(edit.get("fontSize") or 12), _hex_color(edit.get("color")),
                            edit.get("fontFamily") or "", bool(edit.get("isRTL")))
                plans.append((edit, p))
        # 1) remove only the changed words. No fill: rules, shading and pictures stay.
        for _edit, p in plans:
            for r in p.remove:
                page.add_redact_annot(_to_unrotated(page, r), fill=False)
        if any(p.remove for _e, p in plans):
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
        # 2) the new words
        for edit, p in plans:
            try:
                for text, x_left, base, sz, reuse, stack, st, rtl, width in p.lines:
                    color = _int_color(st["color"])
                    if reuse:
                        _draw_reused(page, reuse, text, x_left, base, sz, color)
                        how = "original font"
                    else:
                        line = ptl.layout(text, stack, sz, base_rtl=rtl)
                        face = line.glyphs[0][0] if line.glyphs else None
                        fake_bold = st["bold"] and face is not None and not face.bold
                        fake_italic = st["italic"] and face is not None and not face.italic
                        writer.draw(page, line, x_left, base, color, fake_bold, fake_italic)
                        how = os.path.basename(face.path) if face else "?"
                    report.append({"id": edit.get("id"), "text": text, "size": round(sz, 2), "font": how})
                if p.legacy:
                    _legacy_write(page, *p.legacy)
                    report.append({"id": edit.get("id"), "text": p.legacy[1], "font": "legacy"})
                applied += 1
            except Exception as e:  # noqa: BLE001
                sys.stderr.write(f"edit {edit.get('id')} failed: {e}\n")
        if writer is not None:
            writer.flush()
    # garbage=3 also merges duplicate objects but is quadratic in the object count: 2.7 s of
    # a 2.8 s edit on a 9.1k-object file was the save (10 Oct 2026). Same rule as
    # pdf_protect.py and pdf_redact.py — garbage=2 (still drops everything unreferenced) past
    # 4,000 objects — with object streams, which keep that output at 0.43 MB instead of 1.0 MB.
    doc.save(output_pdf, garbage=3 if doc.xref_length() <= 4000 else 2, deflate=True, use_objstms=1)
    doc.close()
    print(json.dumps({"success": True, "message": f"Applied {applied} edits", "outputFile": output_pdf,
                      "edits": report, "seconds": round(time.time() - t0, 2)}, ensure_ascii=False))


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"success": False, "message": "Usage: pdf_text_editor.py <command> [args]"}))
        sys.exit(1)
    command = sys.argv[1]
    if command == "extract" and len(sys.argv) >= 4:
        extract_text_blocks(sys.argv[2], sys.argv[3])
    elif command == "replace" and len(sys.argv) >= 5:
        replace_text(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        print(json.dumps({"success": False, "message": "Invalid command or arguments"}))
        sys.exit(1)


if __name__ == "__main__":
    main()
