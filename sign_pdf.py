#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sign_pdf.py — stamp signatures and company stamps onto a PDF.

    python sign_pdf.py <input.pdf> <output.pdf> <placements.json> <images_dir>

placements.json:
    {
      "items": [
        {"image": "0.png", "page": 0, "x": 0.61, "y": 0.78, "w": 0.22, "h": 0.07,
         "opacity": 1.0, "rotate": 0}
      ]
    }

x/y/w/h are FRACTIONS of the page box, origin top-left. Fractions rather than
points on purpose: the browser places the overlay against a pdf.js viewport
rendered at some arbitrary zoom, and any attempt to hand over points has to
agree with that zoom, with the device pixel ratio and with /Rotate. Fractions
survive all three, because both sides divide by the same page box.

The artwork arrives as a transparent PNG that the BROWSER drew. That is the
whole trick behind Arabic support here: no font is embedded, no reshaping and
no bidi runs on the server, so a company stamp reading "شركة كونفيليو للتجارة"
lands as the browser shaped it. The server only ever moves pixels.
"""
import json
import os
import sys
import uuid

import fitz  # PyMuPDF

import pdf_geometry as geom  # same folder


def _page_box(page):
    """The box the browser measured against: page.rect honours /Rotate, exactly
    like pdf.js getViewport(), so a fraction means the same thing on both sides."""
    return page.rect


def _place(page, pdf_rect, xref, rotate, name):
    """Show image `xref` (already in the file) in pdf_rect, PDF coordinates, under the page
    resource name `name`. This is page.insert_image(xref=…) without its search for an unused
    name: that search lists every image, form and font of the page, and on a file whose pages
    share one /Resources (every LibreOffice or Chrome PDF) that dictionary gains an entry per
    placement, so the listing grew with every page — 97 % of the run time in cProfile when a
    stamp went on all 400 pages of such a file: 11.9 s (10 Oct 2026). One name per artwork,
    unique to this run, needs no search (those 400 pages: 0.5 s); the same image under the
    same name on every page also keeps a shared /Resources at one entry per artwork."""
    page.wrap_contents()  # a balanced graphics state, as insert_image ensures
    page._insert_image(clip=pdf_rect, overlay=1, rotate=rotate, keep_proportion=0, xref=xref,
                       _imgname=name, digests=None)


def sign(input_pdf, output_pdf, placements_json, images_dir):
    with open(placements_json, "r", encoding="utf-8") as fh:
        spec = json.load(fh)

    items = spec.get("items") or []
    if not items:
        raise ValueError("no items to place")

    doc = fitz.open(input_pdf)

    if doc.needs_pass:
        doc.close()
        raise ValueError("encrypted")

    # Each PNG is read once and embedded once: "put the stamp on every page" is the common
    # case, and every later placement shows the same image object.
    blobs = {}
    xrefs = {}   # artwork name -> image xref, after its first placement
    names = {}   # artwork name -> resource name used for the later placements
    tag = uuid.uuid4().hex[:8]

    def blob(name):
        if name not in blobs:
            path = os.path.join(images_dir, os.path.basename(name))
            with open(path, "rb") as fh:
                blobs[name] = fh.read()
        return blobs[name]

    placed = 0
    for item in items:
        pno = int(item.get("page", 0))
        if pno < 0 or pno >= doc.page_count:
            continue

        page = doc[pno]
        box = _page_box(page)

        x = box.x0 + float(item.get("x", 0)) * box.width
        y = box.y0 + float(item.get("y", 0)) * box.height
        w = float(item.get("w", 0.2)) * box.width
        h = float(item.get("h", 0.08)) * box.height
        if w <= 0 or h <= 0:
            continue

        rect = fitz.Rect(x, y, x + w, y + h)

        # insert_image works in the page's UNROTATED space, while page.rect (and
        # the browser's viewport, and therefore the fractions above) describe the
        # rotated space the reader sees. On a /Rotate 90 page the two disagree and
        # the signature lands off the page entirely — measured, not assumed.
        # derotation_matrix maps visual -> stored; the image then has to be turned
        # by the same angle so it still reads upright once the viewer re-applies
        # the page rotation. From there pdf_geometry takes it to the file: PyMuPDF's
        # own mapping drops the box origin of a rotated page, and a signature placed at
        # 60 %/75 % of a cropped landscape page landed at 43 %/62 % (10 Oct 2026).
        rotation = page.rotation % 360
        if rotation:
            rect = (rect * page.derotation_matrix).normalize()
        to_pdf = geom.to_pdf(page)
        rotate = (int(item.get("rotate", 0)) + rotation) % 360
        name = item["image"]

        if name in xrefs and hasattr(page, "_insert_image"):
            try:
                _place(page, rect * to_pdf, xrefs[name], rotate, names[name])
                placed += 1
                continue
            except (TypeError, AttributeError):
                pass  # a PyMuPDF whose private signature changed: the public call still works

        xref = page.insert_image(
            rect * to_pdf * page.transformation_matrix,
            stream=None if name in xrefs else blob(name),
            xref=xrefs.get(name, 0),
            overlay=True,
            # keep_proportion=False: the user already fixed the aspect ratio by
            # dragging the corner handle, and re-fitting here would shift the
            # artwork out from under the box they positioned.
            keep_proportion=False,
            rotate=rotate,
        )
        if name not in xrefs:
            xrefs[name] = xref
            names[name] = "CfSig%s%d" % (tag, len(names))
        placed += 1

    if placed == 0:
        doc.close()
        raise ValueError("nothing was placed")

    # A signature that can be dragged off the page is not a signature. Flatten
    # so the artwork is page content, not an annotation a reader can move.
    # Each PNG is embedded once (above). garbage=3 also merges duplicate objects but is
    # quadratic in the object count: stamping 3 pages of a 9.1k-object file took 2.78 s, of
    # which the save was 2.6 s (10 Oct 2026). Same rule as pdf_protect.py and pdf_redact.py:
    # garbage=2 (still drops everything unreferenced) past 4,000 objects, with object
    # streams, which made that output 0.43 MB instead of 1.03 MB, in 0.21 s.
    doc.save(output_pdf, garbage=3 if doc.xref_length() <= 4000 else 2, deflate=True, use_objstms=1)
    doc.close()

    print(json.dumps({"success": True, "placed": placed}, ensure_ascii=False))


def main():
    if len(sys.argv) < 5:
        print(json.dumps({
            "success": False,
            "message": "Usage: sign_pdf.py <in.pdf> <out.pdf> <placements.json> <images_dir>",
        }))
        sys.exit(1)

    try:
        sign(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
    except Exception as exc:  # surfaced verbatim to the controller's log
        print(json.dumps({"success": False, "message": str(exc)}, ensure_ascii=False),
              file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
