#!/usr/bin/env python3
"""
pdf_crop.py — cut margins off PDF pages the way the reader sees them.

Usage: python3 pdf_crop.py in.pdf out.pdf opts.json
opts.json: {"top": 10, "right": 10, "bottom": 10, "left": 10,
            "unit": "mm" | "cm" | "in" | "pt" | "percent",
            "pages": [1, 3] | null (= all)}
Prints one JSON line ({"pages": n, "cropped": k}); exit 0. Exit 3 = encrypted; exit 2 =
failure; exit 4 = the margins leave nothing of a page (message on stderr); exit 5 = the page
range names a page the file does not have ("no such page: 50 (4 pages)" on stderr).

WHY (measured 2026-09-14 on the live tool): the PdfSharp path took the margins in the
page's *unrotated* coordinate system, so on a page carrying /Rotate 90 "top 30 mm" cut the
LEFT edge of what the user sees, and it only wrote a /CropBox while the /MediaBox kept the
old size — so a viewer or printer that honours the MediaBox showed the page uncropped.

Here the margins are applied to page.rect (what is displayed, rotation included), mapped
back through the page's derotation matrix into the unrotated box, offset by the existing
CropBox, and written as BOTH the MediaBox and the CropBox, with Trim/Bleed/Art boxes
removed so every viewer agrees on the page. Percent margins are a percent of the displayed
width (left/right) and height (top/bottom) of that page, so an A5 page in an A4 file is
cropped by its own size.
"""
import json
import re
import sys

import fitz  # PyMuPDF

PT = {"mm": 72.0 / 25.4, "cm": 72.0 / 2.54, "in": 72.0, "pt": 1.0, "px": 0.75}


def to_pt(value, unit, reference):
    try:
        v = float(value or 0)
    except (TypeError, ValueError):
        v = 0.0
    if v <= 0:
        return 0.0
    u = (unit or "pt").strip().lower()
    if u in ("percent", "%"):
        return reference * (v / 100.0)
    return v * PT.get(u, 1.0)


_DIGITS = str.maketrans("\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669"
                        "\u06f0\u06f1\u06f2\u06f3\u06f4\u06f5\u06f6\u06f7\u06f8\u06f9",
                        "01234567890123456789")


def parse_range(spec, n):
    """"1-3, 5" -> ([1, 2, 3, 5], []); "all"/empty -> (None, []) (every page).

    Reads the range the way people type it (5 Oct 2026): "1 - 3" used to split on the
    spaces and crop pages 1 and 3 only; an en dash from a phone keyboard ("2–4"), Arabic
    digits and comma ("٢-٤، ٧") and an open end ("5-", "-3") matched nothing and fell
    through to every page.

    The second list holds the parts that name no page of this file: "50" or "7-9" on a
    4-page PDF, "abc". Until 10 Oct 2026 a range that matched nothing still fell through to
    every page — "50" on a 4-page file cropped all 4 — and a stray "50" next to valid pages
    was dropped without a word. A range that runs past the end ("3-10") keeps its pages."""
    if not spec or str(spec).strip().lower() == "all":
        return None, []
    s = str(spec).translate(_DIGITS)
    s = re.sub("[\u2010-\u2015\u2212]", "-", s).replace("\u060c", ",").replace("\u061b", ",")
    s = re.sub(r"\s*-\s*", "-", s)
    pages, missing = set(), []
    for part in re.split(r"[,\s;]+", s.strip()):
        if not part:
            continue
        m = re.fullmatch(r"(\d*)-(\d*)", part)
        if m and (m.group(1) or m.group(2)):
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else n
            hit = range(max(1, a), min(n, b) + 1)
            if len(hit):
                pages.update(hit)
            else:
                missing.append(part)
        elif part.isdecimal() and 1 <= int(part) <= n:
            pages.add(int(part))
        else:
            missing.append(part)
    if not pages and not missing:
        return None, []  # only separators: as good as empty
    return sorted(pages), missing


def main(src, dst, opts_path):
    with open(opts_path, "r", encoding="utf-8") as f:
        o = json.load(f)
    doc = fitz.open(src)
    if doc.needs_pass:
        print("encrypted", file=sys.stderr)
        return 3
    n = len(doc)
    if o.get("pages"):
        wanted = [int(p) for p in o["pages"]]
        missing = [str(p) for p in wanted if not 1 <= p <= n]
    else:
        wanted, missing = parse_range(o.get("pageRange"), n)
    if missing:
        print("no such page: %s (%d pages)" % (", ".join(missing), n), file=sys.stderr)
        return 5
    targets = set(range(n)) if not wanted else {p - 1 for p in wanted}
    unit = o.get("unit", "mm")
    cropped = 0
    for i in range(n):
        if i not in targets:
            continue
        page = doc[i]
        vis = page.rect  # displayed rectangle: rotation and existing CropBox applied
        left = to_pt(o.get("left"), unit, vis.width)
        right = to_pt(o.get("right"), unit, vis.width)
        top = to_pt(o.get("top"), unit, vis.height)
        bottom = to_pt(o.get("bottom"), unit, vis.height)
        new_vis = fitz.Rect(vis.x0 + left, vis.y0 + top, vis.x1 - right, vis.y1 - bottom)
        if new_vis.width < 10 or new_vis.height < 10:
            print(f"page {i + 1}: margins leave nothing visible", file=sys.stderr)
            return 4
        # displayed -> unrotated, top-left origin, relative to the current CropBox
        unrot = (new_vis * page.derotation_matrix).normalize()
        cb = page.cropbox  # y-down, MediaBox-relative
        box = fitz.Rect(unrot.x0 + cb.x0, unrot.y0 + cb.y0, unrot.x1 + cb.x0, unrot.y1 + cb.y0)
        # MediaBox is written in PDF coordinates (y up), so flip against the old MediaBox.
        # x needs no offset: PyMuPDF's cropbox x is already absolute, so adding mb.x0 counted a
        # non-zero origin twice — re-cropping a file cropped on the left, or any PDF whose page
        # box does not start at x=0, shifted the content and left a blank strip (25 Sept 2026).
        mb = page.mediabox
        pdf_box = fitz.Rect(box.x0, mb.y1 - box.y1, box.x1, mb.y1 - box.y0)
        pdf_box.normalize()
        page.set_mediabox(pdf_box)  # resets CropBox to the full new MediaBox
        for key in ("CropBox", "TrimBox", "BleedBox", "ArtBox"):
            try:
                doc.xref_set_key(page.xref, key, "null")
            except Exception:
                pass
        cropped += 1
    if cropped == 0:
        print("no pages matched", file=sys.stderr)
        return 2
    doc.save(dst, garbage=1, deflate=True)
    print(json.dumps({"pages": n, "cropped": cropped}))
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("usage: pdf_crop.py in.pdf out.pdf opts.json", file=sys.stderr)
        sys.exit(2)
    try:
        sys.exit(main(sys.argv[1], sys.argv[2], sys.argv[3]))
    except Exception as e:
        print(f"pdf_crop: {e}", file=sys.stderr)
        sys.exit(2)
