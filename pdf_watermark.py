#!/usr/bin/env python3
"""
pdf_watermark.py — a text watermark that fits the page, follows its rotation and shapes Arabic.

Usage: python3 pdf_watermark.py in.pdf out.pdf opts.json
opts.json: {"text": "CONFIDENTIAL", "fontSize": 72, "rotation": 45, "opacity": 0.3,
            "color": "#ff0000", "fontDir": "/path/to/Fonts",
            "fontLatin": "arialbd.ttf", "fontArabic": "IBMPlexSansArabic-Bold.ttf"}
Prints one JSON line ({"pages": n, "fontSize": used}); exit 0. Exit 3 = encrypted; exit 2 =
failure (the caller falls back to its PdfSharp path).

WHY (measured 2026-09-14 on the live tool): the PdfSharp path drew the string once at the
requested size, so a sentence-length watermark ran off both edges of the page ("RIBUTE -
Mahmou" was all that stayed on an A4 sheet), on a page carrying /Rotate 90 it was laid down
in the unrotated frame and came out sideways and clipped, and Arabic went through a
presentation-forms reshaper into Arial, which joins letters but has no real kerning or
ligatures. Here the text is typeset by MuPDF's Story engine (proper shaping and bidi, the
site's own Plex Arabic face for Arabic, Arial Bold for Latin) into a small stamp page, the
largest size that fits inside the page — on up to three lines for long text — is searched
from the requested size downwards, and the stamp is placed on the displayed page rectangle
with show_pdf_page, which handles rotation for us. Positive angles run bottom-left to
top-right, the convention every other watermark tool uses.
"""
import json
import math
import os
import re
import sys
import uuid

import fitz  # PyMuPDF

import pdf_geometry as geom  # same folder

ARABIC = re.compile(r"[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]")


def esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


def pick_font(o, text):
    font_dir = o.get("fontDir") or ""
    name = o.get("fontArabic") if ARABIC.search(text) else o.get("fontLatin")
    if not name:
        name = "IBMPlexSansArabic-Bold.ttf" if ARABIC.search(text) else "arialbd.ttf"
    path = os.path.join(font_dir, name)
    if not os.path.exists(path):
        # any Arabic-capable face on the box beats a missing one
        for cand in ("/usr/share/fonts/confileo/IBMPlexSansArabic-Bold.ttf",
                     "/usr/share/fonts/google-noto/NotoSansArabic-Bold.ttf",
                     "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Bold.ttf"):
            if os.path.exists(cand):
                return os.path.dirname(cand), os.path.basename(cand)
        return None, None
    return font_dir, name


def fits(text, size, box_w, box_h, css_font, archive, color, opacity):
    """Lay the text out at `size` in a box_w x box_h stamp; return the stamp doc or None."""
    tmp = fitz.open()
    page = tmp.new_page(width=box_w, height=box_h)
    css = (css_font + " p{font-size:%.2fpx;color:%s;margin:0;padding:0;text-align:center;line-height:1.25;}"
           % (size, color))
    html = '<p dir="auto">%s</p>' % esc(text).replace("\n", "<br>")
    spare, scale = page.insert_htmlbox(fitz.Rect(2, 2, box_w - 2, box_h - 2), html, css=css,
                                       archive=archive, opacity=opacity, scale_low=1)
    if spare < 0 or scale < 1:
        tmp.close()
        return None
    # trim the stamp to the used height so it centres on the page properly
    used_h = min(box_h, box_h - spare + 4)
    return tmp, used_h


def _matrix(src_rect, tar_rect, rotate):
    """The transform PyMuPDF's show_pdf_page computes: centre the source on the target,
    rotate, and scale uniformly to fit."""
    smp = (src_rect.tl + src_rect.br) / 2.0
    tmp = (tar_rect.tl + tar_rect.br) / 2.0
    m = fitz.Matrix(1, 0, 0, 1, -smp.x, -smp.y) * fitz.Matrix(rotate)
    sr1 = src_rect * m
    f = min(tar_rect.width / sr1.width, tar_rect.height / sr1.height)
    m *= fitz.Matrix(f, f)
    m *= fitz.Matrix(1, 0, 0, 1, tmp.x, tmp.y)
    return m


def _stamp(page, target, stamps, sno, clip, rotate, graftmap, name):
    """show_pdf_page(target, stamps, sno, rotate=rotate, clip=clip) without its name search.

    The public call lists every XObject, font and image of the page to invent an unused
    resource name. When all pages share one /Resources (every LibreOffice PDF), that dict gains
    one entry per stamped page and the listing turns quadratic: 563 pages took 20.8 s and
    1,200 pages 57 s (25 Sept 2026). A unique name per run makes the search unnecessary.

    `target` is in the page's unrotated space and reaches the file through pdf_geometry:
    PyMuPDF's own mapping drops the box origin of a rotated page, which pushed the watermark
    off-centre and past the edge on rotated pages whose CropBox does not start at (0,0)
    (10 Oct 2026)."""
    to_pdf = geom.to_pdf(page)
    if graftmap is not None and hasattr(page, "_show_pdf_page"):
        try:
            src_page = stamps[sno]
            src = (src_page.rect & clip) * ~src_page.transformation_matrix
            tar = target * to_pdf
            page.wrap_contents()  # a balanced graphics state, as show_pdf_page ensures
            page._show_pdf_page(src_page, overlay=True, matrix=tuple(_matrix(src, tar, rotate)),
                                xref=0, oc=0, clip=src, graftmap=graftmap, _imgname=name)
            return
        except TypeError:
            pass  # a PyMuPDF whose private signature changed: the public call still works
    page.show_pdf_page(target * to_pdf * page.transformation_matrix, stamps, sno, rotate=rotate,
                       overlay=True, clip=clip)


def main(src, dst, opts_path):
    with open(opts_path, "r", encoding="utf-8") as f:
        o = json.load(f)
    text = str(o.get("text") or "").strip()
    if not text:
        print("empty text", file=sys.stderr)
        return 2
    size = float(o.get("fontSize") or 72)
    size = max(8.0, min(size, 400.0))
    angle = float(o.get("rotation") or 0) % 360
    opacity = float(o.get("opacity") if o.get("opacity") is not None else 0.3)
    opacity = max(0.02, min(opacity, 1.0))
    color = str(o.get("color") or "#808080")
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
        color = "#808080"

    doc = fitz.open(src)
    if doc.needs_pass:
        print("encrypted", file=sys.stderr)
        return 3

    font_dir, font_name = pick_font(o, text)
    archive = fitz.Archive(font_dir) if font_dir else None
    css_font = ("@font-face{font-family:wm;src:url(%s);} p{font-family:wm,sans-serif;font-weight:700;}"
                % font_name) if font_name else "p{font-family:sans-serif;font-weight:700;}"

    th = math.radians(angle)
    c, s = abs(math.cos(th)), abs(math.sin(th))
    # 1) One stamp per distinct page size, all of them pages of ONE stamp document, so that
    # subsetting it leaves one font file that the graft map copies into the output once.
    stamps = fitz.open()
    cache = {}  # (w, h) rounded -> (stamp page number, used size, box_w, used_h)
    plan = []
    used_size = size
    for page in doc:
        W, H = page.rect.width, page.rect.height
        key = (round(W), round(H))
        if key not in cache:
            # longest segment through the centre at this angle, with a 6% margin each side
            L = min(W / c if c > 1e-6 else 1e9, H / s if s > 1e-6 else 1e9) * 0.88
            L = min(L, math.hypot(W, H))
            stamp = None
            trial = size
            while trial >= 8:
                box_h = trial * 1.3 * 3 + 8  # room for up to three wrapped lines
                res = fits(text, trial, L, box_h, css_font, archive, color, opacity)
                if res is not None:
                    stamp, used_h = res
                    # a stamp taller than the page's short side would spill: shrink further
                    if used_h <= min(W, H) * 0.9:
                        break
                    stamp.close(); stamp = None
                trial *= 0.9
            if stamp is None:
                res = fits(text, 8, L, 8 * 1.3 * 3 + 8, css_font, archive, color, opacity)
                if res is None:
                    print("text cannot be placed", file=sys.stderr)
                    return 2
                stamp, used_h = res; trial = 8
            stamps.insert_pdf(stamp)
            stamp.close()
            cache[key] = (len(stamps) - 1, trial, L, used_h)
        used_size = min(used_size, cache[key][1])
        plan.append((page.number, key))
    # The Story embeds the whole face (Arial Bold is 990 KB, one copy per distinct page size),
    # which made a 112 KB file 2.3 MB; subsetting keeps only the watermark's glyphs (back to
    # ~130 KB, 14 Sept 2026). Only the STAMPS are subset: until 10 Oct 2026 the whole output
    # was (doc.subset_fonts()), which also cut the document's own fonts down to the glyphs on
    # its pages — a form's embedded field font went from 1,045,720 to 51,291 bytes and could
    # no longer draw newly typed values.
    try:
        stamps.subset_fonts()
    except Exception as e:  # noqa: BLE001 — a face it cannot subset: size, not correctness
        print(f"subset_fonts: {e}", file=sys.stderr)

    # 2) Place them, all through one graft map.
    try:
        graftmap = fitz.Graftmap(doc)
    except Exception:  # noqa: BLE001
        graftmap = None
    tag = uuid.uuid4().hex[:8]
    for pno, key in plan:
        page = doc[pno]
        sno, _used, L, used_h = cache[key]
        W, H = page.rect.width, page.rect.height
        cx, cy = page.rect.x0 + W / 2, page.rect.y0 + H / 2
        # show_pdf_page fits the ROTATED stamp into the target while keeping proportions, so
        # the target has to be the bounding box of the rotated band, not the band itself —
        # otherwise a 45-degree stamp is shrunk until its tilted bounds fit a flat strip.
        bw = L * c + used_h * s
        bh = L * s + used_h * c
        target = fitz.Rect(cx - bw / 2, cy - bh / 2, cx + bw / 2, cy + bh / 2)
        # the stamp page is L x box_h; show only the used band so the rotation pivots on the text
        clip = fitz.Rect(0, 0, L, used_h)
        # page.rect is the DISPLAYED frame, but show_pdf_page places in the page's own
        # (unrotated) coordinates and the viewer then turns the page by /Rotate. Until 25 Sept
        # 2026 the target was never mapped back, so with the default 45° a /Rotate 90 page got
        # the text at -45° and off-centre (0.65, 0.70), and /Rotate 180 and 270 upside down.
        # Map the box through the derotation matrix and add the page's rotation to the angle.
        rot = page.rotation % 360
        if rot:
            target = (target * page.derotation_matrix).normalize()
        _stamp(page, target, stamps, sno, clip, (angle + rot) % 360, graftmap, f"CfWm{tag}{pno}")
    # garbage=1 drops what is unreferenced; garbage=3's object-by-object comparison was the
    # other big cost on files whose pages share resources (563 LibreOffice pages: 20.8 s → see
    # the commit), and there is nothing left to deduplicate once the font is grafted once.
    try:
        doc.save(dst, garbage=1, deflate=True, use_objstms=1)
    except TypeError:
        doc.save(dst, garbage=1, deflate=True)
    stamps.close()
    print(json.dumps({"pages": len(doc), "fontSize": round(used_size, 1)}))
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("usage: pdf_watermark.py in.pdf out.pdf opts.json", file=sys.stderr)
        sys.exit(2)
    try:
        sys.exit(main(sys.argv[1], sys.argv[2], sys.argv[3]))
    except Exception as e:
        print(f"pdf_watermark: {e}", file=sys.stderr)
        sys.exit(2)
