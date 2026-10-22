#!/usr/bin/env python3
"""
pdf_page_numbers.py — stamp page numbers where the READER sees them, in a real Arabic font.

Usage: python3 pdf_page_numbers.py in.pdf out.pdf options.json
options: {"position": "bottom-center", "format": "Page {n} of {total}", "fontSize": 11,
          "color": "#333333", "startNumber": 1, "skipFirstPage": false, "margin": 20,
          "bold": false, "fontFile": "/usr/share/fonts/cairo/Cairo.ttf"}
Prints one JSON line ({"pages": n, "numbered": n}); exit 0. Exit 3 = encrypted; 2 = failure
(caller falls back to its PdfSharp path).

WHY (measured 2026-09-08 on the live tool): the PdfSharp engine drew in the page's UNROTATED
coordinate space, so on a page carrying /Rotate 90 (every landscape scan from a phone) the
"bottom-centre" number appeared at the left-middle edge, and on /Rotate 180 at the top; and an
Arabic template ("صفحة {n} من {total}") came out as empty boxes because the reshaped
presentation forms were not in the font. MuPDF's Story shapes Arabic with HarfBuzz and takes a
`rotate` argument, so the footer is placed in the displayed frame and reads correctly.

SPEED (25 Sept 2026). A 563-page LibreOffice PDF took 87 s — past Cloudflare's 100 s once the
upload is counted — and 1,200 pages took 154 s. Three costs, measured one by one:
  * insert_htmlbox() per page (28 s): every call builds a Story, searches for a fitting scale by
    laying it out several times, writes a one-page temp PDF with the WHOLE font in it and grafts
    that PDF onto the page, so the font was copied once per page;
  * save(garbage=4) (59 s): deduplicating those 563 font copies object by object;
  * subset_fonts() on the result (3 s).
Now every label is laid out once into ONE stamp document (the font is embedded there once), and
each stamp is drawn onto its page through one shared graft map. Placing the stamp goes through
the low-level call, because the public show_pdf_page() lists every XObject, font and image of the
page to invent an unused name — on a LibreOffice PDF all pages share one /Resources, which grows
by one entry per stamp, so that listing is quadratic (95% of the stamping time in cProfile).
Measured on the 563-page file: 87 s → 1.1 s (Latin and Arabic alike), 3.1 MB → 1.4 MB.

The font now comes from the app's own Fonts/ folder (IBM Plex Sans Arabic, which has Latin,
Arabic and a real bold face) instead of whatever the host happens to have: "bold" used to render
pixel-identical to regular because only one face was ever loaded, and its presentation-form
cmap gives the stamped Arabic a readable text layer («ﺻﻔﺤﺔ 3 ﻣﻦ 20», which NFKC-normalises to
«صفحة 3 من 20») where the host fonts produced «ﻪğﺤ4Ěﺻ3». The fontFile the service passes stays
the fallback.
"""
import io
import json
import os
import re
import sys
import uuid

import fitz  # PyMuPDF >= 1.24 (Story, DocumentWriter)

import pdf_geometry as geom  # same folder

AR = re.compile(r"[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]")
HERE = os.path.dirname(os.path.abspath(__file__))


def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _bundled_font(bold):
    """IBM Plex Sans Arabic from the app's Fonts/ folder (next to Scripts/ in dev and in the
    published build), or None."""
    name = "IBMPlexSansArabic-Bold.ttf" if bold else "IBMPlexSansArabic-Regular.ttf"
    for d in (os.path.join(HERE, "..", "Fonts"), os.path.join(os.getcwd(), "Fonts")):
        p = os.path.normpath(os.path.join(d, name))
        if os.path.isfile(p):
            return p
    return None


def _matrix(src_rect, tar_rect, rotate):
    """Same transform PyMuPDF's show_pdf_page computes: centre the source on the target,
    rotate, and scale uniformly to fit."""
    smp = (src_rect.tl + src_rect.br) / 2.0
    tmp = (tar_rect.tl + tar_rect.br) / 2.0
    m = fitz.Matrix(1, 0, 0, 1, -smp.x, -smp.y) * fitz.Matrix(rotate)
    sr1 = src_rect * m
    f = min(tar_rect.width / sr1.width, tar_rect.height / sr1.height)
    m *= fitz.Matrix(f, f)
    m *= fitz.Matrix(1, 0, 0, 1, tmp.x, tmp.y)
    return m


def _stamp(page, rect, stamps, sno, rotate, graftmap, name):
    """Draw stamps[sno] into rect (the page's unrotated space) on page, rotated by `rotate`
    degrees. The rectangle goes to the file through pdf_geometry: PyMuPDF's own mapping drops
    the box origin of a rotated page, which put the number off the page on a /Rotate 270 page
    whose CropBox does not start at (0,0) (10 Oct 2026)."""
    src_page = stamps[sno]
    to_pdf = geom.to_pdf(page)
    if graftmap is not None and hasattr(page, "_show_pdf_page"):
        try:
            tar = rect * to_pdf
            src = src_page.rect * ~src_page.transformation_matrix
            page.wrap_contents()  # a balanced graphics state, as show_pdf_page ensures
            page._show_pdf_page(src_page, overlay=True, matrix=tuple(_matrix(src, tar, rotate)),
                                xref=0, oc=0, clip=src, graftmap=graftmap, _imgname=name)
            return
        except TypeError:
            pass  # a PyMuPDF whose private signature changed: the public call below still works
    page.show_pdf_page(rect * to_pdf * page.transformation_matrix, stamps, sno, rotate=rotate)


def main(src, dst, opts_path):
    with open(opts_path, "r", encoding="utf-8") as f:
        o = json.load(f)
    doc = fitz.open(src)
    if doc.needs_pass:
        print("encrypted", file=sys.stderr)
        return 3

    template = o.get("format") or "{n}"
    position = (o.get("position") or "bottom-center").lower()
    vert, _, horiz = position.partition("-")
    horiz = horiz or "center"
    size = float(o.get("fontSize") or 11)
    color = o.get("color") or "#333333"
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", color or ""):
        color = "#333333"
    # 0 is a real start number (a cover page counted as page 0); `or 1` turned it into 1.
    try:
        start = int(o.get("startNumber")) if o.get("startNumber") is not None else 1
    except (TypeError, ValueError):
        start = 1
    skip_first = bool(o.get("skipFirstPage"))
    # 0 is a real choice (numbers on the edge); "or 20" turned it into 20 pt.
    margin = float(o["margin"]) if o.get("margin") is not None else 20.0
    bold = bool(o.get("bold"))

    font_file = _bundled_font(bold) or o.get("fontFile") or "/usr/share/fonts/cairo/Cairo.ttf"
    archive = None
    css = "body { margin: 0; }"
    measure = None
    if os.path.isfile(font_file):
        archive = fitz.Archive(os.path.dirname(font_file))
        css += " @font-face { font-family: PgNum; src: url(%s); }" % os.path.basename(font_file)
        try:
            measure = fitz.Font(fontfile=font_file)
        except Exception:  # noqa: BLE001
            measure = None
    family = "PgNum, sans-serif" if archive else "sans-serif"
    # A real bold face carries its own weight; a regular one is asked for 700 so MuPDF at least
    # tries (it has nothing else when the bundled fonts are missing).
    weight = "700" if bold and "Bold" not in os.path.basename(font_file) else "400"

    skip = 1 if skip_first and len(doc) > 1 else 0
    # {total} is the number the LAST numbered page carries, so "{n} of {total}" ends at
    # "7 of 7" whatever the start number. It was the count of numbered pages, which with a
    # custom start read "Page 5 of 3" … "Page 7 of 3" (start 5, cover skipped, 4 pages;
    # 10 Oct 2026). With the default start of 1 the two are the same number.
    total = start + (len(doc) - skip) - 1
    from datetime import date
    today = date.today().isoformat()
    align = {"left": "left", "right": "right"}.get(horiz, "center")

    # 1) Lay every label out once, into one stamp document (one font copy for all of them).
    buf = io.BytesIO()
    writer = fitz.DocumentWriter(buf)
    plan = []  # (page index, rect in page coordinates, rotation)
    number = start
    for i, page in enumerate(doc):
        if i < skip:
            continue
        text = template.replace("{n}", str(number)).replace("{total}", str(total)).replace("{date}", today)
        number += 1
        rtl = bool(AR.search(text))
        # page.rect is ALREADY the displayed frame (PyMuPDF applies /Rotate to it; the
        # unrotated box is page.mediabox), so no swapping — that mistake put the footer off-page.
        rot = page.rotation % 360
        W, H = page.rect.width, page.rect.height
        box_w = max(W - 2 * margin, 10.0)
        # Shrink a label that would not fit on one line (a long template on a narrow page)
        # instead of letting it wrap out of its box; insert_htmlbox did this with scale_low=0.5.
        fs = size
        if measure is not None:
            try:
                need = measure.text_length(text, fontsize=size) * 1.08
                if need > box_w:
                    fs = max(size * 0.5, size * box_w / need)
            except Exception:  # noqa: BLE001
                pass
        box_h = fs * 1.9
        y0 = margin if vert == "top" else H - margin - box_h
        rect_disp = fitz.Rect(margin, y0, margin + box_w, y0 + box_h)
        rect_page = rect_disp * page.derotation_matrix  # back to the page's own coordinates
        rect_page.normalize()
        a = align
        # with dir=rtl, CSS left/right are logical in Story: map physical → logical
        if rtl and a in ("left", "right"):
            a = "right" if a == "left" else "left"
        html = (f'<div dir="{"rtl" if rtl else "ltr"}" style="font-family:{family}; font-size:{fs}pt; '
                f'color:{color}; font-weight:{weight}; text-align:{a}; line-height:1.4; margin:0; padding:0; '
                f'white-space:nowrap">{esc(text)}</div>')
        where = fitz.Rect(0, 0, box_w, box_h)
        story = fitz.Story(html=html, user_css=css, archive=archive)
        dev = writer.begin_page(where)
        story.place(where)
        story.draw(dev)
        writer.end_page()
        plan.append((i, rect_page, rot))
    writer.close()
    stamps = fitz.open("pdf", buf.getvalue())
    # One font, a few glyphs: subsetting keeps the result small (237 KB face → 16 KB). Only the
    # STAMP document is subset. Until 10 Oct 2026 the whole output was (doc.subset_fonts()),
    # which also cut the document's own fonts down to the glyphs drawn on its pages: a form
    # whose field font is an embedded Arial lost every other glyph (1,045,720 → 51,291 bytes),
    # so a viewer drawing newly typed values with it showed nothing for "Z", "k", "H"…
    try:
        stamps.subset_fonts()
    except Exception:  # noqa: BLE001 — size, not correctness
        pass

    # 2) Draw each label on its page through ONE graft map, so the font is copied once.
    try:
        graftmap = fitz.Graftmap(doc)
    except Exception:  # noqa: BLE001
        graftmap = None
    tag = uuid.uuid4().hex[:8]
    for sno, (i, rect_page, rot) in enumerate(plan):
        _stamp(doc[i], rect_page, stamps, sno, rot, graftmap, f"CfPn{tag}{sno}")

    # garbage=1 drops what is unreferenced; the deduplication of garbage=4 is no longer needed
    # (the font exists once) and was the single biggest cost on files whose pages share one
    # /Resources. Object streams keep the cross-reference data small.
    try:
        doc.save(dst, garbage=1, deflate=True, use_objstms=1)
    except TypeError:
        doc.save(dst, garbage=1, deflate=True)
    res = {"pages": len(doc), "numbered": len(plan)}
    doc.close()
    print(json.dumps(res))
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("usage: pdf_page_numbers.py in.pdf out.pdf options.json", file=sys.stderr)
        sys.exit(2)
    try:
        sys.exit(main(sys.argv[1], sys.argv[2], sys.argv[3]))
    except Exception as e:  # noqa: BLE001
        print(f"pdf_page_numbers failed: {e}", file=sys.stderr)
        sys.exit(2)
