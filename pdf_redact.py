#!/usr/bin/env python3
"""
pdf_redact.py — TRUE redaction with PyMuPDF.

Usage: python3 pdf_redact.py in.pdf out.pdf areas.json [--color #000000]
       python3 pdf_redact.py --preview in.pdf out.json [--dpi 100]   (old clients only)

areas.json: [{"pageNumber": 1, "x": 10.5, "y": 20.0, "width": 30.0, "height": 4.2}, ...]
            x/y/width/height are PERCENT of the page as displayed (rotation applied).
Prints one JSON line on success; exit 0.
Exit 3 = encrypted (needs a password), 4 = unreadable, 5 = no usable box: a fallback
engine would not do better, the caller reports it. Exit 2 = anything else (the caller
falls back to its raster engine, which is safe but image-only).

WHY (measured 2026-09-08 on the live tool): the previous engine rasterised the WHOLE
document through ImageMagick at 200 dpi and painted a rectangle on the bitmap. The
sensitive pixels were gone, but so was everything else: the text layer of every page
(no search, no copy, no accessibility), vector crispness, links, bookmarks, and the file
grew several times. This engine removes exactly what is under the box and leaves the rest
of the page as it was: text stays searchable, Arabic stays Arabic, the file stays small.

WHAT "REMOVED" MEANS (leak test 2026-09-25): the first version of this engine still
shipped, inside the "redacted" file, a vector signature under a box, the value of a form
field under a box ("SECRET-FIELD-123" in pdftotext), a sticky note half under a box, the
Keywords and the XMP metadata — and, on every Chrome/Skia PDF, the removed words
themselves: Skia wraps each glyph cluster in /ActualText, MuPDF drops the glyphs but keeps
the marked-content span, so an Arabic line redacted from a Chrome export read back letter
by letter. Now:
  * under a box: glyphs (visible and invisible OCR text) with their /ActualText, image
    pixels, and vector paths — the ones the box covers and the curved or small ones that
    are mostly under it, so a signature poking out of the box does not survive whole
    under the black fill (recoverable by deleting one rectangle);
  * annotations touching a box are deleted with their popups and replies (notes, free
    text, highlights, ink, attachments, links, JavaScript links); form fields touching a
    box are deleted with their value, appearance and entry in the form;
  * document-wide hidden data that can repeat what was redacted: Info + XMP metadata
    (every /Metadata stream), attached files (/EmbeddedFiles, /AF), JavaScript, XFA data,
    application private data (/PieceInfo), page thumbnails, and the alt text of tagged
    elements on redacted pages;
  * the file is rewritten from scratch with garbage collection, so earlier revisions of
    an incrementally-saved file and every replaced object are gone from the bytes.
Nothing else changes: text, links, notes, fields, bookmarks and the invisible OCR layer
outside the boxes stay (scrub()'s defaults would reset every field in the file, drop every
link and strip a scan's whole text layer). Each redacted page is checked afterwards; a page
where text or a covered path still sits under a box is replaced by an image of itself —
the one remaining way to keep the promise — rather than shipped as is.
"""
import collections
import json
import math
import re
import secrets
import sys
import zlib

import fitz  # PyMuPDF >= 1.24 (PDF_REDACT_LINE_ART_*), production pin 1.26

import pdf_geometry as geom  # same folder

EXIT_ENCRYPTED, EXIT_UNREADABLE, EXIT_NO_AREAS = 3, 4, 5

# MuPDF decides "the box covers this path" with the path's stroke bounds, and for mitred
# joins (the PDF default) it widens them by line width x miter limit (10 by default): a
# 2 pt signature needed a box 20 pt larger than itself on every side (measured 2026-09-25,
# MuPDF 1.26.12), so in practice PDF_REDACT_LINE_ART_REMOVE_IF_COVERED removed nothing a
# person would draw a box around. While the boxes are applied the page therefore runs with
# a miter limit of 1.0001 (bounds ~ the real ink); the marker is put back to the default
# 10 afterwards, so the page draws exactly as before.
MITER_MARK = b"1.0001"
_MITER_RE = re.compile(rb"(?<![\d.])1\.0001\s+M(?![A-Za-z])")


class Refused(Exception):
    def __init__(self, code, reason):
        super().__init__(reason)
        self.code = code


def hex_to_rgb(s):
    s = (s or "#000000").strip().lstrip("#")
    if len(s) == 3:
        s = "".join(c * 2 for c in s)
    try:
        return tuple(int(s[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except Exception:  # noqa: BLE001
        return (0.0, 0.0, 0.0)


def _num(v):
    f = float(v)
    if not math.isfinite(f):
        raise ValueError("not finite")
    return f


def page_boxes(doc, areas):
    """Percent boxes on the DISPLAYED page -> Rects in unrotated page space, per page."""
    out = {}
    for a in areas:
        try:
            p = int(a.get("pageNumber", 1)) - 1
            if p < 0 or p >= len(doc):
                continue
            page = doc[p]
            r = page.rect  # displayed size, origin 0,0
            x0 = r.x0 + _num(a["x"]) / 100.0 * r.width
            y0 = r.y0 + _num(a["y"]) / 100.0 * r.height
            x1 = x0 + _num(a["width"]) / 100.0 * r.width
            y1 = y0 + _num(a["height"]) / 100.0 * r.height
            box = fitz.Rect(x0, y0, x1, y1).normalize() & r
            if box.is_empty or box.width < 0.5 or box.height < 0.5:
                continue
            if page.rotation:
                box = (box * page.derotation_matrix).normalize()
            out.setdefault(p, []).append(box)
        except Exception:  # noqa: BLE001
            continue
    return out


# ───────────────────────── annotations, links, form fields ─────────────────────────

_REF = re.compile(r"(\d+)\s+(\d+)\s+R")


def _ref_list(doc, xref, key):
    """Indirect references of an array value (None when the key is not an array)."""
    kind, val = doc.xref_get_key(xref, key)
    if kind != "array":
        return None
    return [int(m.group(1)) for m in _REF.finditer(val)]


def _set_ref_list(doc, xref, key, refs):
    doc.xref_set_key(xref, key, "[" + " ".join(f"{r} 0 R" for r in refs) + "]")


def _drop_from_annots(doc, page, xref):
    """Last resort: unlink an annotation from the page's /Annots array by hand."""
    refs = _ref_list(doc, page.xref, "Annots")
    if refs is not None and xref in refs:
        _set_ref_list(doc, page.xref, "Annots", [r for r in refs if r != xref])


def _raw_rect(doc, page, xref, from_pdf):
    """/Rect of an annotation in unrotated page space (the space of the boxes). PyMuPDF
    reports links in rotated space and widgets unrotated, so the file's own /Rect is the
    one reference that means the same thing for every annotation type. `from_pdf` is
    pdf_geometry's mapping: page.transformation_matrix drops the box origin of a rotated
    page, so a field or link under a box on a rotated, cropped page was not seen, and the
    page was rasterised by the check below instead (measured 10 Oct 2026)."""
    try:
        kind, val = doc.xref_get_key(xref, "Rect")
        if kind != "array":
            return None
        nums = [float(v) for v in val.strip("[] ").split()]
        if len(nums) != 4:
            return None
        return (fitz.Rect(nums) * from_pdf).normalize()
    except Exception:  # noqa: BLE001
        return None


def _touches(rect, boxes):
    """Any overlap at all — a sticky note half under a box goes."""
    if rect is None:
        return False
    grown = fitz.Rect(rect.x0 - .01, rect.y0 - .01, rect.x1 + .01, rect.y1 + .01)
    return any(grown.intersects(b) for b in boxes)


def remove_touching_annots(doc, page, boxes, stats):
    """Delete every annotation, link and form widget that touches a box. Their text (note
    contents, field values, URIs, scripts, attached files) is not drawn where the box is,
    so redacting the page content alone leaves it in the file."""
    doomed = []
    from_pdf = geom.from_pdf(page)
    for xref, atype, _ in page.annot_xrefs():
        if atype in (fitz.PDF_ANNOT_POPUP, fitz.PDF_ANNOT_REDACT):
            continue  # popups go with their parent; pending redactions are applied below
        hit = _touches(_raw_rect(doc, page, xref, from_pdf), boxes)
        if not hit and atype not in (fitz.PDF_ANNOT_LINK, fitz.PDF_ANNOT_WIDGET):
            try:  # NoRotate icons (sticky notes) are drawn away from /Rect on rotated pages
                a = page.load_annot(xref)
                hit = a is not None and _touches(a.rect, boxes)
            except Exception:  # noqa: BLE001
                hit = False
        if hit:
            doomed.append((xref, atype))
    removed_widgets = set()
    if not doomed:
        return removed_widgets
    links = {l.get("xref"): l for l in page.get_links() if l.get("xref")}
    for xref, atype in doomed:
        try:
            if atype == fitz.PDF_ANNOT_LINK:
                link = links.get(xref)
                if link is not None:
                    page.delete_link(link)
                else:
                    _drop_from_annots(doc, page, xref)
                stats["links"] += 1
            elif atype == fitz.PDF_ANNOT_WIDGET:
                removed_widgets.add(xref)
                kind, val = doc.xref_get_key(xref, "Parent")
                if kind == "xref":
                    removed_widgets.add(-int(val.split()[0]))  # negative: a parent field
                w = next((w for w in page.widgets() if w.xref == xref), None)
                if w is not None:
                    page.delete_widget(w)
                else:
                    _drop_from_annots(doc, page, xref)
                stats["fields"] += 1
            else:
                a = page.load_annot(xref)
                if a is not None:
                    page.delete_annot(a)  # also deletes /IRT replies and the popup
                else:
                    _drop_from_annots(doc, page, xref)
                stats["annotations"] += 1
        except Exception:  # noqa: BLE001
            _drop_from_annots(doc, page, xref)
    return removed_widgets


def prune_form(doc, removed):
    """After widgets were deleted, drop form fields that have no widget left (with their
    /V value) from /AcroForm /Fields, the parents' /Kids and the calculation order /CO —
    any of those references alone would keep the value in the file. A button group that
    lost a widget loses its value too: /V names the selected option."""
    cat = doc.pdf_catalog()
    kind, val = doc.xref_get_key(cat, "AcroForm")
    if kind == "null":
        return
    if kind == "xref":
        holder, prefix = int(val.split()[0]), ""
    else:
        holder, prefix = cat, "AcroForm/"
    live = set()
    for page in doc:
        for x, t, _ in page.annot_xrefs():
            if t == fitz.PDF_ANNOT_WIDGET:
                live.add(x)
    seen = {}

    def alive(x, depth=0):
        if x in seen:
            return seen[x]
        seen[x] = x in live
        kids = _ref_list(doc, x, "Kids")
        if kids and depth < 50:
            keep = [k for k in kids if alive(k, depth + 1)]
            if len(keep) != len(kids):
                _set_ref_list(doc, x, "Kids", keep)
            # a radio group whose SELECTED button went: /V would still name it
            if -x in removed and doc.xref_get_key(x, "FT")[1] == "/Btn":
                value = doc.xref_get_key(x, "V")[1]
                shown = {doc.xref_get_key(k, "AS")[1] for k in keep}
                if value not in ("null", "/Off") and value not in shown:
                    doc.xref_set_key(x, "V", "/Off")
            seen[x] = bool(keep) or x in live
        return seen[x]

    fields = _ref_list(doc, holder, prefix + "Fields") or []
    keep = [f for f in fields if alive(f)]
    if len(keep) != len(fields):
        _set_ref_list(doc, holder, prefix + "Fields", keep)
    co = _ref_list(doc, holder, prefix + "CO")
    if co is not None:
        keep_co = [c for c in co if seen.get(c, c in live)]
        if len(keep_co) != len(co):
            _set_ref_list(doc, holder, prefix + "CO", keep_co)


# ───────────────────────────────── vector paths ─────────────────────────────────

def _ink_rect(d):
    r = fitz.Rect(d.get("rect"))
    w = (d.get("width") or 0) / 2 if d.get("type") in ("s", "fs") else 0
    return fitz.Rect(r.x0 - w, r.y0 - w, r.x1 + w, r.y1 + w)


def _area(r):
    """Rect.get_area() is missing from PyMuPDF 1.26.5, which requirements.txt allows: there
    every redaction stopped here and the caller fell back to rasterising the pages
    (10 Oct 2026). width x height, 0 for an empty rectangle, on every version."""
    return 0.0 if r.is_empty else r.width * r.height


def path_targets(page, boxes):
    """Paths to remove for the boxes: every path whose ink lies inside a box, curved paths
    (signatures, handwriting, logos) with at least half their area under a box, and small
    paths (≤ 30 pt: text converted to outlines, check marks) touching a box. Straight
    rules merely crossing a box (table borders) are kept, so the layout stays intact."""
    out = []
    try:
        drawings = page.get_drawings()
    except Exception:  # noqa: BLE001
        return out
    for d in drawings:
        r = _ink_rect(d)
        near = [b for b in boxes if r.intersects(b) or b.contains(r)]
        if not near:
            continue
        if r.is_empty or r.width < 0.01 or r.height < 0.01:
            if any(b.contains(r.tl) and b.contains(r.br) for b in near):
                out.append(d)
            continue
        if any(b.contains(r) for b in near):
            out.append(d)
            continue
        small = max(r.width, r.height) <= 30
        curved = any(it[0] in ("c", "qu") for it in d.get("items", ()))
        area = max(_area(r), 1e-6)
        share = max(_area(fitz.Rect(r) & b) / area for b in near)
        if (small and share > 0) or (curved and share >= 0.5):
            out.append(d)
    return out


def _signature(d):
    r = fitz.Rect(d.get("rect"))
    return (round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1),
            tuple(it[0] for it in d.get("items", ())))


def remove_paths(page, targets, factor):
    """Line-art-only redaction (no fill, no text, no image change) of each target's ink
    rectangle, grown by line width x factor so MuPDF's stroke bounds fit inside."""
    for d in targets:
        r = _ink_rect(d)
        g = (d.get("width") or 1) * factor + 0.5
        page.add_redact_annot(fitz.Rect(r.x0 - g, r.y0 - g, r.x1 + g, r.y1 + g), fill=None, cross_out=False)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                          graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                          text=fitz.PDF_REDACT_TEXT_NONE)


def _surviving(page, targets):
    want = collections.Counter(_signature(d) for d in targets)
    left = 0
    for d in page.get_drawings():
        s = _signature(d)
        if want.get(s):
            want[s] -= 1
            left += 1
    return left


def _wrap_miter(doc, page):
    pre = doc.get_new_xref()
    doc.update_object(pre, "<<>>")
    doc.update_stream(pre, b"q " + MITER_MARK + b" M\n", new=True)
    post = doc.get_new_xref()
    doc.update_object(post, "<<>>")
    doc.update_stream(post, b"\nQ\n", new=True)
    conts = page.get_contents()
    doc.xref_set_key(page.xref, "Contents", "[" + " ".join(f"{x} 0 R" for x in [pre] + conts + [post]) + "]")


def _page_streams(doc, page):
    """Content streams of the page and of the form XObjects it draws."""
    xs = list(page.get_contents())
    try:
        xs += [x[0] for x in page.get_xobjects()]
    except Exception:  # noqa: BLE001
        pass
    return list(dict.fromkeys(xs))


def _restore_miter(doc, page):
    for x in _page_streams(doc, page):
        try:
            data = doc.xref_stream(x)
        except Exception:  # noqa: BLE001
            continue
        if data and MITER_MARK in data:
            new = _MITER_RE.sub(b"10 M", data)
            if new != data:
                doc.update_stream(x, new)


# ─────────────────── marked content: /ActualText, /Alt, /E in the page ───────────────────

_WS = frozenset(b" \t\r\n\f\x00")
_DELIM = frozenset(b"()<>[]{}/%")
_ESC = {0x6E: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12}
_HIDDEN_TEXT_KEYS = (b"ActualText", b"Alt", b"E")


def _lex(data):
    """Tokens of a content stream as (kind, start, end, value); kinds: num name str op
    [ ] << >> img (a whole inline image, BI … EI)."""
    i, n = 0, len(data)
    while i < n:
        c = data[i]
        if c in _WS:
            i += 1
        elif c == 0x25:  # % comment
            while i < n and data[i] not in (10, 13):
                i += 1
        elif c == 0x28:  # (literal string)
            j, depth, buf = i + 1, 1, bytearray()
            while j < n:
                ch = data[j]
                if ch == 0x5C:
                    j += 1
                    if j >= n:
                        break
                    e = data[j]
                    if e in _ESC:
                        buf.append(_ESC[e])
                        j += 1
                    elif 0x30 <= e <= 0x37:
                        k, v = j, 0
                        while k < n and k < j + 3 and 0x30 <= data[k] <= 0x37:
                            v = v * 8 + data[k] - 0x30
                            k += 1
                        buf.append(v & 0xFF)
                        j = k
                    elif e in (10, 13):
                        j += 2 if (e == 13 and j + 1 < n and data[j + 1] == 10) else 1
                    else:
                        buf.append(e)
                        j += 1
                    continue
                if ch == 0x28:
                    depth += 1
                elif ch == 0x29:
                    depth -= 1
                    if depth == 0:
                        j += 1
                        break
                buf.append(ch)
                j += 1
            yield ("str", i, j, bytes(buf))
            i = j
        elif c == 0x3C:
            if i + 1 < n and data[i + 1] == 0x3C:
                yield ("<<", i, i + 2, None)
                i += 2
                continue
            j = data.find(b">", i + 1)
            j = n if j == -1 else j
            hx = bytes(ch for ch in data[i + 1:j] if ch not in _WS)
            if len(hx) % 2:
                hx += b"0"
            try:
                val = bytes.fromhex(hx.decode("ascii"))
            except ValueError:
                val = b""
            yield ("str", i, min(j + 1, n), val)
            i = j + 1
        elif c == 0x3E:
            if i + 1 < n and data[i + 1] == 0x3E:
                yield (">>", i, i + 2, None)
                i += 2
            else:
                i += 1
        elif c == 0x5B:
            yield ("[", i, i + 1, None)
            i += 1
        elif c == 0x5D:
            yield ("]", i, i + 1, None)
            i += 1
        elif c in (0x7B, 0x7D):
            i += 1
        else:
            j = i + 1
            while j < n and data[j] not in _WS and data[j] not in _DELIM:
                j += 1
            tok = data[i:j]
            if c == 0x2F:
                yield ("name", i, j, tok[1:])
            elif tok[:1] in b"+-.0123456789":
                yield ("num", i, j, tok)
            elif tok == b"BI":
                k = data.find(b"ID", j)
                end = n
                if k != -1:
                    p = k + 3
                    while True:
                        q = data.find(b"EI", p)
                        if q == -1:
                            break
                        if data[q - 1] in _WS and (q + 2 >= n or data[q + 2] in _WS):
                            end = q + 2
                            break
                        p = q + 1
                yield ("img", i, end, data[i:end])
                j = end
            else:
                yield ("op", i, j, tok)
            i = j


def _operations(data):
    """(operator token, operands) pairs; arrays and dicts are grouped as
    (kind, start, end, items)."""
    stack, starts = [[]], []
    for tok in _lex(data):
        kind = tok[0]
        if kind in ("[", "<<"):
            stack.append([])
            starts.append((kind, tok[1]))
        elif kind in ("]", ">>"):
            if len(stack) == 1:
                continue
            items = stack.pop()
            k, s = starts.pop()
            stack[-1].append(("array" if k == "[" else "dict", s, tok[2], items))
        elif kind == "img" and len(stack) == 1:
            yield tok, []
            stack[0] = []
        elif kind == "op" and len(stack) == 1 and tok[3] not in (b"true", b"false", b"null"):
            yield tok, stack[0]
            stack[0] = []
        else:
            stack[-1].append(tok)


def _dict_entries(prop):
    items = prop[3]
    out = {}
    for i in range(0, len(items) - 1, 2):
        k, v = items[i], items[i + 1]
        if k[0] == "name":
            out[k[3]] = (k[1], v[2], v[3] if v[0] == "str" else None)
    return out


def _sections(data, xobject_id):
    """Marked-content sections carrying /ActualText, /Alt or /E, each with what it draws:
    the bytes of its strings, the identity of its images, its inline images."""
    out, stack = [], []
    for tok, ops in _operations(data):
        op = tok[3] if tok[0] == "op" else b"BI"
        if op in (b"BDC", b"BMC"):
            sec = None
            if op == b"BDC" and ops:
                prop = ops[-1]
                if prop[0] == "dict":
                    entries = _dict_entries(prop)
                    hits = {k: v for k, v in entries.items() if k in _HIDDEN_TEXT_KEYS}
                    if hits:
                        key = tuple(sorted((k, v[2] or b"") for k, v in hits.items()))
                        sec = {"key": key, "hits": hits, "payload": []}
                elif prop[0] == "name":
                    sec = {"key": (b"/" + prop[3],), "named": prop[3], "payload": []}
            stack.append(sec)
            if sec is not None:
                out.append(sec)
            continue
        if op == b"EMC":
            if stack:
                stack.pop()
            continue
        piece = None
        if op in (b"Tj", b"'", b'"'):
            s = [o for o in ops if o[0] == "str"]
            piece = b"T" + (s[-1][3] if s else b"")
        elif op == b"TJ":
            arr = [o for o in ops if o[0] == "array"]
            piece = b"T" + b"".join(x[3] for x in (arr[-1][3] if arr else []) if x[0] == "str")
        elif op == b"Do":
            nm = [o for o in ops if o[0] == "name"]
            piece = b"D" + xobject_id(nm[-1][3] if nm else b"")
        elif op == b"BI":
            piece = b"I%d" % len(tok[3] or b"")
        if piece is not None and len(piece) > 1:
            for sec in stack:
                if sec is not None:
                    sec["payload"].append(piece)
    return out


def _xobject_ids(doc, page):
    ids = {}
    try:
        for img in page.get_images(full=True):
            if img[9] == 0:
                ids[img[7].encode("latin-1", "replace")] = b"img%d" % img[0]
        for xo in page.get_xobjects():
            if xo[2] == 0:
                ids.setdefault(xo[1].encode("latin-1", "replace"), b"form")
    except Exception:  # noqa: BLE001
        pass
    return lambda name: ids.get(name, b"?")


def page_marked_content(doc, page):
    try:
        data = page.read_contents()
    except Exception:  # noqa: BLE001
        return collections.Counter()
    xid = _xobject_ids(doc, page)
    return collections.Counter((s["key"], tuple(s["payload"])) for s in _sections(data, xid))


def scrub_marked_content(doc, page, before, stats):
    """Drop /ActualText, /Alt and /E from marked-content spans whose content the
    redaction emptied or changed. Spans that draw exactly what they drew before keep
    them — Chrome's Arabic depends on them for copy and search."""
    xid = _xobject_ids(doc, page)
    remaining = collections.Counter(before)
    contents = set(page.get_contents())
    for x in _page_streams(doc, page):
        try:
            data = doc.xref_stream(x)
        except Exception:  # noqa: BLE001
            continue
        if not data or b"BDC" not in data:
            continue
        cuts = []
        is_page = x in contents
        for sec in _sections(data, xid if is_page else (lambda n: b"?")):
            sig = (sec["key"], tuple(sec["payload"]))
            if sec["payload"] and is_page and remaining.get(sig):
                remaining[sig] -= 1
                continue
            if sec["payload"] and not is_page:
                continue  # form XObjects: only emptied spans can be told apart safely
            if "hits" in sec:
                cuts.extend((v[0], v[1]) for v in sec["hits"].values())
            elif "named" in sec:
                for k in ("ActualText", "Alt", "E"):
                    path = "Resources/Properties/" + sec["named"].decode("latin-1") + "/" + k
                    try:
                        if doc.xref_get_key(page.xref, path)[0] != "null":
                            doc.xref_set_key(page.xref, path, "null")
                            stats["hiddenText"] += 1
                    except Exception:  # noqa: BLE001
                        pass
        if cuts:
            out = bytearray(data)
            for s, e in sorted(set(cuts), reverse=True):
                del out[s:e]
            doc.update_stream(x, bytes(out))
            stats["hiddenText"] += len(set(cuts))


# ───────────────────────────── checks and last resorts ─────────────────────────────

def leftover_text(page, boxes):
    """Characters whose centre still lies inside a box after redaction. Plain float
    comparisons: Rect/Point objects per character made this check 80% of the run time
    (6.3 of 7.9 s on 80 pages)."""
    try:
        raw = page.get_text("rawdict", flags=0)
    except Exception:  # noqa: BLE001
        return 0
    inner = [(b.x0 + .25, b.y0 + .25, b.x1 - .25, b.y1 - .25) for b in boxes]
    left = 0
    for block in raw.get("blocks", ()):
        for line in block.get("lines", ()):
            lx0, ly0, lx1, ly1 = line["bbox"]
            near = [b for b in inner if lx0 <= b[2] and lx1 >= b[0] and ly0 <= b[3] and ly1 >= b[1]]
            if not near:
                continue
            for span in line.get("spans", ()):
                for ch in span.get("chars", ()):
                    if not ch.get("c", "").strip():
                        continue
                    x0, y0, x1, y1 = ch["bbox"]
                    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                    for b in near:
                        if b[0] <= cx <= b[2] and b[1] <= cy <= b[3]:
                            left += 1
                            break
    return left


def rasterize_page(doc, pno):
    """Replace a page by a picture of itself (boxes already painted)."""
    page = doc[pno]
    pix = page.get_pixmap(dpi=150, annots=True)
    w, h = page.rect.width, page.rect.height
    doc.xref_set_key(page.xref, "Annots", "null")
    for x in page.get_contents():
        doc.update_stream(x, b" ")
    doc.xref_set_key(page.xref, "Resources", "<<>>")
    page.set_rotation(0)
    page.set_mediabox(fitz.Rect(0, 0, w, h))
    page.insert_image(page.rect, stream=pix.tobytes("jpeg", jpg_quality=88))


def _plain_colorspace(doc, xref, n):
    """True when a JPEG of n components can keep the image's colour space as is."""
    kind, val = doc.xref_get_key(xref, "ColorSpace")
    if kind == "name":
        return val in ("/DeviceGray", "/DeviceRGB") and (n == 1) == (val == "/DeviceGray")
    if kind in ("xref", "array"):
        src = doc.xref_object(int(val.split()[0])) if kind == "xref" else val
        return ("/ICCBased" in src or "/CalRGB" in src or "/CalGray" in src) and not any(
            t in src for t in ("/Indexed", "/Separation", "/DeviceN", "/Lab"))
    return False


def recompress_images(doc, page, before):
    """MuPDF writes an image whose pixels it blanked uncompressed, and the save deflates
    it: a noisy 300 dpi A4 scan went from a 1.7 MB JPEG to a 14.8 MB file (2026-09-25).
    Images that were JPEG before are written as JPEG again, in their own colour space;
    lossless ones (PNG, line art) stay lossless."""
    jpeg_sizes = {(img[2], img[3]) for img in before.values() if img[8] == "DCTDecode"}
    for img in page.get_images(full=True):
        xref = img[0]
        if xref in before or img[4] != 8 or (img[2], img[3]) not in jpeg_sizes:
            continue
        try:
            if doc.xref_get_key(xref, "Filter")[1] not in ("null", "/FlateDecode"):
                continue
            if doc.xref_get_key(xref, "SMask")[0] != "null" or doc.xref_get_key(xref, "Mask")[0] != "null":
                continue
            if doc.xref_get_key(xref, "Decode")[0] != "null":
                continue
            raw = doc.xref_stream_raw(xref)
            if len(raw) < 64 * 1024:
                continue
            pix = fitz.Pixmap(doc, xref)
            if pix.alpha or pix.n not in (1, 3) or not _plain_colorspace(doc, xref, pix.n):
                continue
            jpg = pix.tobytes("jpeg", jpg_quality=90)
            # what the save would write (it deflates the raw pixels): flat synthetic pages
            # compress better losslessly, and then stay lossless
            flate = len(raw) if doc.xref_get_key(xref, "Filter")[1] == "/FlateDecode" else len(zlib.compress(raw, 6))
            if len(jpg) * 1.25 >= flate:
                continue
            doc.update_stream(xref, jpg, compress=False)
            doc.xref_set_key(xref, "Filter", "/DCTDecode")
            doc.xref_set_key(xref, "DecodeParms", "null")
        except Exception:  # noqa: BLE001
            continue


# ──────────────────────────────── document-wide ────────────────────────────────

def _struct_page(doc, x, depth=0):
    kind, val = doc.xref_get_key(x, "Pg")
    if kind == "xref":
        return int(val.split()[0])
    if depth > 40:
        return None
    kind, val = doc.xref_get_key(x, "P")
    if kind == "xref":
        return _struct_page(doc, int(val.split()[0]), depth + 1)
    return None


def scrub_hidden_data(doc, redacted_page_xrefs, stats):
    """Document-wide data that can repeat what was redacted and that nobody sees on the
    page. Visible content outside the boxes is NOT touched."""
    meta = doc.metadata or {}
    if any(meta.get(k) for k in ("title", "author", "subject", "keywords", "creator", "producer", "creationDate", "modDate")):
        stats["metadata"] = True
    doc.set_metadata({})
    if doc.get_xml_metadata():
        stats["metadata"] = True
    doc.del_xml_metadata()
    for name in doc.embfile_names():
        doc.embfile_del(name)
        stats["attachments"] += 1
    cat = doc.pdf_catalog()
    for key in ("Names/JavaScript", "Names/EmbeddedFiles", "AcroForm/XFA", "AA", "PieceInfo", "AF", "Metadata"):
        try:
            if doc.xref_get_key(cat, key)[0] != "null":
                doc.xref_set_key(cat, key, "null")
                if key in ("Names/JavaScript", "AA"):
                    stats["scripts"] += 1
        except Exception:  # noqa: BLE001
            pass
    try:
        if doc.xref_get_key(cat, "OpenAction/S")[1] == "/JavaScript":
            doc.xref_set_key(cat, "OpenAction", "null")
            stats["scripts"] += 1
    except Exception:  # noqa: BLE001
        pass
    for page in doc:
        for key in ("Thumb", "PieceInfo", "AA", "Metadata", "AF"):
            if doc.xref_get_key(page.xref, key)[0] != "null":
                doc.xref_set_key(page.xref, key, "null")
    for x in range(1, doc.xref_length()):
        try:
            src = doc.xref_object(x, compressed=True)
        except Exception:  # noqa: BLE001
            continue
        if not any(t in src for t in ("/JavaScript", "/JS", "/AA", "/PieceInfo", "/Metadata", "/AF", "/XFA", "/Alt", "/ActualText")):
            continue
        try:
            if doc.xref_get_key(x, "S")[1] == "/JavaScript":
                doc.update_object(x, "null")  # a whole action object (names tree, /A, /Next …)
                stats["scripts"] += 1
                continue
            if doc.xref_get_key(x, "Type")[1] == "/Metadata":
                doc.update_object(x, "<<>>")  # XMP of a page, image or font
                doc.update_stream(x, b" ", new=True)
                continue
            for key in ("A", "OpenAction", "A/Next", "OpenAction/Next"):
                if doc.xref_get_key(x, key + "/S")[1] == "/JavaScript":
                    doc.xref_set_key(x, key, "null")
                    stats["scripts"] += 1
            for key in ("AA", "JavaScript", "JS", "PieceInfo", "Metadata", "AF", "XFA"):
                if doc.xref_get_key(x, key)[0] != "null":
                    doc.xref_set_key(x, key, "null")
                    if key in ("AA", "JavaScript", "JS"):
                        stats["scripts"] += 1
            # Tagged PDFs: the alt text of a figure on a redacted page describes what the
            # box hid ("photo of the ID card"), /ActualText can repeat the removed words.
            # Only elements of redacted pages lose them; other pages keep their tags.
            if ("/Alt" in src or "/ActualText" in src) and doc.xref_get_key(x, "Subtype")[0] == "null" \
                    and doc.xref_get_key(x, "S")[0] == "name":
                pg = _struct_page(doc, x)
                if pg is None or pg in redacted_page_xrefs:
                    for key in ("Alt", "ActualText", "E"):
                        if doc.xref_get_key(x, key)[0] != "null":
                            doc.xref_set_key(x, key, "null")
        except Exception:  # noqa: BLE001
            continue


# ─────────────────────────────────── main ───────────────────────────────────

def redact_page(doc, pno, boxes, fill, stats):
    page = doc[pno]
    removed_widgets = remove_touching_annots(doc, page, boxes, stats)
    images_before = {img[0]: img for img in page.get_images(full=True)}
    marked_before = page_marked_content(doc, page)
    _wrap_miter(doc, page)
    # pass 1 — the boxes: glyphs (visible and invisible), image pixels, covered paths.
    # No fill yet: the paths pass below must not see (and remove) the black boxes.
    for box in boxes:
        page.add_redact_annot(box, fill=None, cross_out=False)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS,
                          graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                          text=fitz.PDF_REDACT_TEXT_REMOVE)
    # pass 2 — line art only: paths that still belong to a box (the ones it only mostly
    # covers). Looked for after the text is gone, because Type3 glyphs (Chrome, TeX) are
    # paths to get_drawings() but text to the redaction. A file that sets its own miter
    # limit defeats the tight bounds; those paths get a second try with MuPDF's wide ones.
    targets = path_targets(page, boxes)
    ok = True
    if targets:
        remove_paths(page, targets, 1.0)
        if _surviving(page, targets):
            remove_paths(page, targets, 10.0)
        ok = _surviving(page, targets) == 0
        stats["paths"] += len(targets)
    _restore_miter(doc, page)
    scrub_marked_content(doc, page, marked_before, stats)
    recompress_images(doc, page, images_before)
    if fill is not None:
        # The boxes are in the unrotated page space the redaction annotations use; the
        # drawing call maps through page.transformation_matrix, which drops the box origin of
        # a rotated page — the text went, but the black bar was drawn 57-151 pt away, over
        # unrelated content, on rotated pages whose CropBox does not start at (0,0) (10 Oct
        # 2026). pdf_geometry gives the rectangle that call expects.
        api = geom.for_api(page)
        shape = page.new_shape()
        for box in boxes:
            shape.draw_rect(box * api)
            shape.finish(fill=fill, color=fill, width=0)
        shape.commit()
    if not ok or leftover_text(page, boxes):
        rasterize_page(doc, pno)
        stats["rasterized"].append(pno + 1)
    return removed_widgets


def redact(src, dst, areas, color):
    try:
        doc = fitz.open(src)
    except Exception as e:  # noqa: BLE001
        raise Refused(EXIT_UNREADABLE, f"cannot open: {e}")
    if doc.needs_pass:
        raise Refused(EXIT_ENCRYPTED, "encrypted")
    if not doc.is_pdf or len(doc) == 0:
        raise Refused(EXIT_UNREADABLE, "not a PDF")
    boxes_by_page = page_boxes(doc, areas)
    if not boxes_by_page:
        raise Refused(EXIT_NO_AREAS, "no usable box")
    meta_enc, perms = (doc.metadata or {}).get("encryption"), doc.permissions
    fill = hex_to_rgb(color)
    stats = {"annotations": 0, "fields": 0, "links": 0, "attachments": 0, "scripts": 0,
             "hiddenText": 0, "metadata": False, "paths": 0, "rasterized": []}
    n = sum(len(b) for b in boxes_by_page.values())
    removed_widgets = set()
    redacted_xrefs = set()
    for pno, boxes in sorted(boxes_by_page.items()):
        redacted_xrefs.add(doc[pno].xref)
        removed_widgets |= redact_page(doc, pno, boxes, fill, stats)
    if removed_widgets:
        try:
            prune_form(doc, removed_widgets)
        except Exception:  # noqa: BLE001
            pass
    scrub_hidden_data(doc, redacted_xrefs, stats)
    pages = len(doc)
    # A file that opened without a password but carried permissions (no editing, no
    # copying …) keeps them: MuPDF would otherwise write it unencrypted. Its owner password
    # is unknown, so a random one replaces it — the restrictions stay what they were.
    enc = {}
    if (meta_enc or "").strip():
        enc = dict(encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="",
                   owner_pw=secrets.token_hex(20), permissions=perms)
    # Full rewrite: unreferenced objects (old revisions, replaced streams, deleted notes)
    # are dropped. garbage=3 also merges duplicates but is quadratic in the object count:
    # 7.6 s on a 563-page LibreOffice file (13.7k objects) against 0.08 s for garbage=2.
    garbage = 3 if doc.xref_length() <= 4000 else 2
    doc.save(dst, garbage=garbage, deflate=True, use_objstms=1, **enc)
    doc.close()
    return n, pages, stats


def preview(src, dst, dpi):
    import base64
    doc = fitz.open(src)
    if doc.needs_pass:
        raise Refused(EXIT_ENCRYPTED, "encrypted")
    pages = []
    dpi = max(40, min(int(dpi or 100), 200))
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csRGB, alpha=False)
        b = pix.tobytes("jpeg", jpg_quality=70)
        pages.append({"pageNumber": i + 1, "width": pix.width, "height": pix.height,
                      "imageBase64": "data:image/jpeg;base64," + base64.b64encode(b).decode("ascii")})
    with open(dst, "w", encoding="utf-8") as f:
        json.dump({"pageCount": len(doc), "pages": pages}, f)
    doc.close()
    return len(pages)


def main(argv):
    if len(argv) >= 3 and argv[0] == "--preview":
        dpi = 100
        if "--dpi" in argv:
            dpi = int(argv[argv.index("--dpi") + 1])
        n = preview(argv[1], argv[2], dpi)
        print(json.dumps({"pages": n}))
        return 0
    if len(argv) < 3:
        print("usage", file=sys.stderr)
        return 2
    src, dst, areas_path = argv[0], argv[1], argv[2]
    color = "#000000"
    if "--color" in argv:
        color = argv[argv.index("--color") + 1]
    with open(areas_path, "r", encoding="utf-8") as f:
        areas = json.load(f)
    if not isinstance(areas, list):
        areas = areas.get("areas", [])
    n, pages, stats = redact(src, dst, areas, color)
    print(json.dumps({"redacted": n, "pages": pages, "removed": stats}))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except Refused as r:
        print(json.dumps({"error": str(r)}))
        print(f"pdf_redact refused: {r}", file=sys.stderr)
        sys.exit(r.code)
    except Exception as e:  # noqa: BLE001
        print(f"pdf_redact failed: {e}", file=sys.stderr)
        sys.exit(2)
