#!/usr/bin/env python3
"""
pdf_geometry.py — where a rectangle the reader sees lands in the PDF, on every kind of page.

Shared by the scripts that draw onto existing pages: page numbers, watermark, sign & stamp,
the redaction fill and the PDF editor's replacement text.

WHY (10 Oct 2026). PyMuPDF's Page.transformation_matrix is exact on an unrotated page — it
carries the CropBox/MediaBox origin, (1,0,0,-1,-50,700) for a CropBox [50 100 450 700] — but
on every rotated page it returns (1,0,0,-1,0,H) and drops that origin. Every insert and draw
call written in Python (insert_image, show_pdf_page, insert_htmlbox, insert_text, Shape)
maps its rectangle to the file through the inverse of that matrix, so on a rotated page whose
box does not start at (0,0) everything landed off by the box origin. Measured on PyMuPDF
1.26.7 (and unchanged in 1.28): page numbers invisible on a /Rotate 270 page, a signature
placed at 60 %/75 % of the page drawn at 43 %/62 %, the redaction's black bar drawn away
from the text it removed. Such pages are common: every rotated page cropped by Confileo's
own Crop tool has one (MediaBox (85,43,538,814) with /Rotate 90), and so do scanner and
imposition output. Text extraction, get_drawings() and redaction annotations were never
affected: PyMuPDF runs those with the page set to /Rotate 0, where MuPDF's own transform
is used.

So this module does the same: it reads the transform at /Rotate 0, where it is MuPDF's own
(CropBox ∩ MediaBox, UserUnit and all), and puts the rotation back.

Spaces
  displayed   page.rect: /Rotate applied, origin top-left; what pdf.js shows and what the
              browser's fractions and percentages are measured against
  unrotated   PyMuPDF's page space at /Rotate 0 (get_text, get_drawings, search_for,
              add_redact_annot); displayed * page.derotation_matrix
  pdf         the file's own user space (y up): content streams, /Rect, show_pdf_page's
              private target

API
  to_pdf(page)            Matrix  unrotated -> pdf
  from_pdf(page)          Matrix  pdf -> unrotated
  displayed_to_pdf(page)  Matrix  displayed -> pdf
  for_api(page)           Matrix  unrotated -> the rectangle the calls that map through
                                  ~page.transformation_matrix take (insert_image,
                                  show_pdf_page, insert_htmlbox, Shape.draw_*); identity
                                  unless the page is rotated and its box origin is not (0,0)
  unrotated(page)         context the page at /Rotate 0 for the duration: insert_text and
                                  insert_textbox do their own arithmetic with the box origin,
                                  so they are called inside it with unrotated coordinates and
                                  rotate=<the page's rotation>
"""
import contextlib

import fitz  # PyMuPDF


@contextlib.contextmanager
def unrotated(page):
    """The page at /Rotate 0 while the block runs, its rotation back afterwards."""
    rotation = page.rotation
    if rotation % 360 == 0:
        yield page
        return
    page.set_rotation(0)
    try:
        yield page
    finally:
        page.set_rotation(rotation)


def to_pdf(page):
    """Matrix from the page's unrotated space to PDF user space, as MuPDF maps it."""
    # transformation_matrix is MuPDF's own transform only at /Rotate 0: read it there.
    # PyMuPDF's get_text(), get_drawings() and add_redact_annot() flip the page the same way.
    with unrotated(page):
        return ~page.transformation_matrix


def from_pdf(page):
    """Matrix from PDF user space to the page's unrotated space."""
    return ~to_pdf(page)


def displayed_to_pdf(page):
    """Matrix from the displayed page (what the reader sees, y down) to PDF user space."""
    return page.derotation_matrix * to_pdf(page)


def for_api(page):
    """Matrix from the unrotated page space to the rectangle PyMuPDF's Python-side insert and
    draw calls expect: they map it back with ~page.transformation_matrix, which is exact at
    /Rotate 0 and off by the box origin on a rotated page."""
    return to_pdf(page) * page.transformation_matrix
