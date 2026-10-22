#!/usr/bin/env python3
"""
pdf_flatten.py — flatten forms and annotations WITHOUT touching the text layer.

Usage: python3 pdf_flatten.py in.pdf out.pdf
Prints one JSON line on success ({"widgets": n, "annots": n, "links": n}); exit 0.
Exit 3 = encrypted (needs a password), exit 2 = failure — the caller falls back to
its Ghostscript path.

WHY (measured 2026-09-08 on the live tool): the Ghostscript pdfwrite route re-encodes
every font and rebuilds ToUnicode from the subset cmap. On Arabic PDFs that turns a
searchable text layer into presentation-form glyph codes in visual order (the same
defect that hit Compress: text similarity 0.088 on an Arabic CV). Here MuPDF's bake()
draws each form field and annotation appearance into the page content stream and
leaves every other object exactly as it was: text searchable, Arabic logical, links,
bookmarks and metadata kept, file size unchanged. Fields that were filled by a tool
that wrote no appearance stream (NeedAppearances) get one generated first, so their
values are visible in the flattened file instead of vanishing.

Those generated appearances come from pdf_form.py's layout (25 Sept 2026), not from
MuPDF's widget.update(): update() clipped "محمد إبراهيم علي" to the bottom of its
glyphs, drew NOTHING for an auto-size (0 Tf) field, printed a combo's export value
("ALX") instead of its label and dropped list selections. The shared code draws the
same boxes Fill PDF Form draws (alignment, fit, wrapping, comb cells, list lines)
with MuPDF's HTML engine, which shapes Arabic.
"""
import json
import sys

import fitz  # PyMuPDF >= 1.24 (Document.bake)

ANNOT_HIDDEN = 1 << 1  # /F bit 2


def _int_key(doc, xref, key):
    try:
        kind, val = doc.xref_get_key(xref, key)
        return int(val) if kind == "int" else 0
    except Exception:  # noqa: BLE001
        return 0


def _form_appearances(doc):
    """Text/choice fields holding a value but no (or a stale) appearance: build it
    the way Fill PDF Form does. Returns the widget xrefs handled; any failure
    returns an empty set and the per-widget update() below runs as before."""
    try:
        import pdf_form  # same folder
        cat = doc.pdf_catalog()
        need_app = doc.xref_get_key(cat, "AcroForm/NeedAppearances")[1] == "true"
        _, ws = pdf_form.read_form(doc, want_labels=False)
        targets = [d for d in ws if d["type"] in ("text", "combobox", "listbox") and not d["hidden"]
                   and (not d["hasAP"] or need_app) and d.get("value")]
        if not targets:
            return set()
        boxes = [pdf_form.make_box(d) for d in targets]
        pw = max(36.0, max(b["w"] for b in boxes))
        ph = max(36.0, max(b["h"] for b in boxes))
        pdf_form.apply_appearances(doc, targets, boxes, pdf_form._mupdf_overlay(boxes, pw, ph))
        return {d["xref"] for d in targets}
    except Exception:  # noqa: BLE001
        return set()


def flatten(src, dst):
    doc = fitz.open(src)
    if doc.needs_pass:
        doc.close()
        return None
    built = _form_appearances(doc)
    widgets = annots = links = 0
    for page in doc:
        links += len(page.get_links())
        # Widgets: hidden ones must not be printed; value-only fields need an appearance.
        for w in list(page.widgets()):
            widgets += 1
            if _int_key(doc, w.xref, "F") & ANNOT_HIDDEN:
                try:
                    page.delete_widget(w)
                except Exception:  # noqa: BLE001
                    pass
                continue
            kind, _ = doc.xref_get_key(w.xref, "AP")
            if kind in ("null", "none") and w.xref not in built:
                try:
                    w.update()  # builds /AP from /V so the value is drawn
                except Exception:  # noqa: BLE001
                    pass
        # Annotations: drop popups (the yellow note windows) and hidden ones.
        for a in list(page.annots()):
            annots += 1
            try:
                if a.type[0] == fitz.PDF_ANNOT_POPUP or (a.flags & ANNOT_HIDDEN):
                    page.delete_annot(a)
            except Exception:  # noqa: BLE001
                pass
    doc.bake(annots=True, widgets=True)
    # No fields are left, so the interactive form dictionary must go too — some viewers
    # otherwise still show a "this document contains form fields" bar.
    try:
        cat = doc.pdf_catalog()
        if doc.xref_get_key(cat, "AcroForm")[0] != "null":
            doc.xref_set_key(cat, "AcroForm", "null")
    except Exception:  # noqa: BLE001
        pass
    # garbage=3 also merges duplicate objects but is quadratic in the object count: 1.97 s on
    # a 9.1k-object file (10 Oct 2026). Same rule as pdf_protect.py and pdf_redact.py —
    # garbage=2 (still drops everything unreferenced) past 4,000 objects — with object streams.
    doc.save(dst, garbage=3 if doc.xref_length() <= 4000 else 2, deflate=True, clean=False, use_objstms=1)
    doc.close()
    return {"widgets": widgets, "annots": annots, "links": links}


def main(argv):
    if len(argv) < 2:
        print("usage: pdf_flatten.py in.pdf out.pdf", file=sys.stderr)
        return 2
    res = flatten(argv[0], argv[1])
    if res is None:
        print("encrypted", file=sys.stderr)
        return 3
    print(json.dumps(res))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except Exception as e:  # noqa: BLE001
        print(f"pdf_flatten failed: {e}", file=sys.stderr)
        sys.exit(2)
