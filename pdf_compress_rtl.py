"""Text-preserving compression (PyMuPDF) for PDFs Ghostscript would damage.

Usage: python3 pdf_compress_rtl.py in.pdf out.pdf --level {0,1,2} [--target BYTES] [--force] [--budget SECONDS]
Exit codes: 0 = wrote out.pdf (JSON line on stdout), 10 = nothing here needs this
engine (the caller uses Ghostscript), 11 = the file needs an open password,
12 = damaged (the pages could not be read: repaired and blank), 2 = failure.

Why: Ghostscript's pdfwrite re-encodes every font and rewrites ToUnicode from the
subset's cmap. For Arabic that turns a searchable, copyable text layer into
presentation-form glyph codes in visual order (measured 7 Sept 2026: text similarity
0.088 before/after on a CV); for Devanagari it splits vowel signs from their letters
(0.849 on a Chrome PDF, 28 Sept 2026); and it drops the fields of a fillable form.
PyMuPDF's rewrite_images() recompresses only the image streams and leaves every text
object, form field, link and bookmark byte-for-byte intact.

The API (CompressPdfService) decides with PdfPig whether a file needs this engine and
then passes --force; without --force the script decides itself (a file PdfPig cannot
read), and exits 10 when there is no such text and no form.
"""
import argparse
import io
import json
import math
import os
import re
import shutil
import sys
import time

import fitz  # PyMuPDF >= 1.24.4 (rewrite_images); 1.26 on the server
from fitz import mupdf

# level -> (image dpi, JPEG quality). Same numbers as the Ghostscript levels in
# CompressPdfService (QFactor 0.4 / 0.76 / 1.2 = JPEG quality 80 / 62 / 42), so the two
# engines give the same kind of file for the same card.
LEVELS = {
    0: (150, 80),  # Low: print-ready
    1: (100, 62),  # Medium
    2: (80, 42),   # High: smallest that still reads (OCR word recall 0.96 on a 300-dpi scan)
}

# Target-size ladder, best first — the Ghostscript ladder's rungs in JPEG quality. Sizes
# fall down the ladder, so the first rung that fits is found by bisection. The last rung
# is the readability floor: below it scanned text stops being readable (OCR word recall
# 0.89 at 75 dpi, 0.71 at 72 dpi with more compression, 0.29 at 50 dpi), so a target that
# needs less is reported as not met instead of being forced.
TARGET_LADDER = [(220, 80), (150, 80), (130, 75), (115, 67), (100, 62), (90, 55), (80, 42), (75, 38)]

SCRIPTS = (
    (0x0590, 0x08FF),  # Hebrew, Arabic, Syriac, Thaana, N'Ko ... Arabic Extended
    (0xFB1D, 0xFDFF),  # Hebrew and Arabic presentation forms A
    (0xFE70, 0xFEFF),  # Arabic presentation forms B
    (0x0900, 0x0DFF),  # Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil, Telugu, Kannada, Malayalam, Sinhala
    (0x0E00, 0x0FFF),  # Thai, Lao, Tibetan
    (0x1000, 0x109F),  # Myanmar
    (0x1780, 0x17FF),  # Khmer
)


def needs_this_engine(doc, max_pages=6):
    """A text layer Ghostscript would garble, or a fillable form it would flatten."""
    try:
        if doc.is_form_pdf:
            return True
    except Exception:  # noqa: BLE001
        pass
    for i, page in enumerate(doc):
        if i >= max_pages:
            break
        t = page.get_text()
        if any(lo <= ord(c) <= hi for c in t for (lo, hi) in SCRIPTS):
            return True
    return False


def looks_damaged(doc):
    """MuPDF had to rebuild the file and the first pages draw nothing at all."""
    if not doc.is_repaired:
        return False
    for i, page in enumerate(doc):
        if i >= 3:
            break
        if page.get_text().strip() or page.get_images() or page.get_drawings():
            return False
    return True


def rewrite_options(dpi, quality, lossy_all=False):
    """JPEG for photos and scans, lossless (Flate) for screenshots and line art, fax for
    1-bit scans. Bicubic resampling reaches the exact dpi (MuPDF's average method only
    halves: 300 -> 150 -> 75, so 72 and 100 both came out at 105 dpi).

    lossy_all: screenshots and charts become JPEG too. Only the target search uses it, and
    only once the whole ladder has failed with them kept lossless (see main)."""
    o = mupdf.PdfImageRewriterOptions()
    thr = int(dpi * 1.5)  # Ghostscript's downsample threshold: leave images within 1.5x alone
    q = str(quality)
    for kind in ("color_lossy", "gray_lossy"):
        setattr(o, f"{kind}_image_recompress_method", mupdf.FZ_RECOMPRESS_JPEG)
        setattr(o, f"{kind}_image_subsample_method", mupdf.FZ_SUBSAMPLE_BICUBIC)
        setattr(o, f"{kind}_image_subsample_threshold", thr)
        setattr(o, f"{kind}_image_subsample_to", dpi)
        setattr(o, f"{kind}_image_recompress_quality", q)
    for kind in ("color_lossless", "gray_lossless"):
        # Screenshots, logos, charts: resampled like everything else but kept lossless.
        # Turning them into JPEG put ringing around every letter of a screenshot.
        setattr(o, f"{kind}_image_recompress_method", mupdf.FZ_RECOMPRESS_JPEG if lossy_all else mupdf.FZ_RECOMPRESS_LOSSLESS)
        setattr(o, f"{kind}_image_subsample_method", mupdf.FZ_SUBSAMPLE_BICUBIC)
        setattr(o, f"{kind}_image_subsample_threshold", thr)
        setattr(o, f"{kind}_image_subsample_to", dpi)
        setattr(o, f"{kind}_image_recompress_quality", q)
    # 1-bit scans are already small (fax/JBIG2) and are only ever re-packed as fax here
    # (MuPDF 1.26 does not resample them; the old average-to-72/100-dpi setting was a
    # no-op). The API sends a mostly 1-bit scan to Ghostscript, which takes it to 150 dpi.
    o.bitonal_image_recompress_method = mupdf.FZ_RECOMPRESS_FAX
    o.bitonal_image_subsample_method = mupdf.FZ_SUBSAMPLE_AVERAGE
    o.bitonal_image_subsample_threshold = 450
    o.bitonal_image_subsample_to = 300
    o.bitonal_image_recompress_quality = q
    return o


def unshare_resources(doc):
    """Give every page its own /Resources dictionary before the images are rewritten.
    When several pages point at one shared dictionary, MuPDF's image rewriter left the
    text of those pages unreadable to copy and search (glyph codes instead of letters;
    measured 28 Sept 2026 on a report whose pages 3-4 share their resources). A direct
    copy per page costs a few bytes and keeps every font and image object shared."""
    users = {}
    for p in doc:
        r = doc.xref_get_key(p.xref, "Resources")
        if r[0] == "xref":
            users.setdefault(int(r[1].split()[0]), []).append(p.xref)
    for x, pages in users.items():
        if len(pages) > 1:
            src = doc.xref_object(x, compressed=True)
            for px in pages:
                doc.xref_set_key(px, "Resources", src)


def same_text(src_path, out_path, limit=40):
    """The engine's promise: every checked page reads exactly as before."""
    a, b = fitz.open(src_path), fitz.open(out_path)
    try:
        if a.page_count != b.page_count:
            return False
        n = a.page_count
        step = max(1, n // limit)
        return all(a[i].get_text() == b[i].get_text() for i in range(0, n, step))
    finally:
        a.close()
        b.close()


def drop_inherited_resources(doc):
    """rewrite_images() gives every page its own /Resources that names the new images,
    but a /Resources inherited from the page tree (LibreOffice writes one on /Pages)
    still names the old ones, so they stayed in the file: an Arabic LibreOffice PDF with
    two photos came out 828 KB from 691 KB and the original was returned (0 %). Once
    every page carries its own, the inherited dictionaries are dead weight."""
    if not all(doc.xref_get_key(p.xref, "Resources")[0] != "null" for p in doc):
        return
    seen = set()
    for p in doc:
        par = doc.xref_get_key(p.xref, "Parent")
        while par[0] == "xref":
            x = int(par[1].split()[0])
            if x in seen:
                break
            seen.add(x)
            if doc.xref_get_key(x, "Resources")[0] != "null":
                doc.xref_set_key(x, "Resources", "null")
            par = doc.xref_get_key(x, "Parent")


def colour_components(doc, xref):
    """Components of an image's /ColorSpace: 1 grey, 3 RGB, 4 CMYK, None when unsure."""
    kind, val = doc.xref_get_key(xref, "ColorSpace")
    simple = {"/DeviceGray": 1, "/CalGray": 1, "/DeviceRGB": 3, "/CalRGB": 3, "/DeviceCMYK": 4}
    if kind == "name":
        return simple.get(val)
    if kind == "xref":
        val = doc.xref_object(int(val.split()[0]), compressed=True)
    elif kind != "array":
        return None
    m = re.search(r"/ICCBased\s*(\d+)\s+0\s+R", val)
    if m:
        n = doc.xref_get_key(int(m.group(1)), "N")
        return int(n[1]) if n[0] == "int" else None
    for name, comps in simple.items():
        if val.lstrip("[ ").startswith(name):
            return comps
    return None


def fix_and_reencode_jpegs(doc, before, quality):
    """Two things about the JPEGs MuPDF's rewriter writes (every new image object):

    - A grey image (DeviceGray, or ICCBased with one component) comes back as a
      three-component JPEG under its one-component colour space, and every viewer then
      draws it as interleaved stripes: a grey 300-dpi scan fell to SSIM 0.47 and OCR 0 —
      the text-preserving engine did this to every grey image of an Arabic PDF before
      28 Sept 2026. Such an image is converted back to grey and re-encoded, always.
    - MuPDF keeps colour at full resolution (4:4:4); Ghostscript halves it (4:2:0), which
      the eye does not see. At the same quality MuPDF's files were 14 % larger for the same
      SSIM (24 MB of phone photos at Medium: 680 KB vs Ghostscript's 596 KB, SSIM 0.9436 /
      0.9435). Colour JPEGs are re-encoded once at the same quality with 4:2:0 and kept
      when smaller: 574 KB, SSIM 0.9434.
    - Each re-encode is written baseline and progressive (same pixels, same quality; only
      the entropy coding differs) and the smaller is kept. Measured 2 Oct 2026: scanned
      pages came out 6-9 % smaller progressive (grey scan at Medium 837 -> 759 KB, colour
      scan 580 -> 535 KB, Arabic scan 283 -> 260 KB); photos 0-2 %; small images are
      sometimes smaller baseline, hence both.
    CMYK, Decode arrays and anything Pillow cannot read stay as MuPDF wrote them."""
    try:
        from PIL import Image
    except Exception:  # noqa: BLE001 - Pillow missing: MuPDF's own JPEGs stay
        return
    done = {}  # (raw bytes, components) -> new bytes or None: one image on every page is one job
    for x in range(1, doc.xref_length()):
        if x in before:
            continue
        try:
            if doc.xref_get_key(x, "Subtype")[1] != "/Image" or doc.xref_get_key(x, "Filter")[1] != "/DCTDecode":
                continue
            if doc.xref_get_key(x, "Decode")[0] != "null":
                continue
            raw = doc.xref_stream_raw(x)
            want = colour_components(doc, x)
            if (raw, want) in done:
                data = done[(raw, want)]
                if data is not None:
                    doc.update_stream(x, data, compress=False)
                    doc.xref_set_key(x, "Filter", "/DCTDecode")
                    doc.xref_set_key(x, "DecodeParms", "null")
                continue
            done[(raw, want)] = None
            im = Image.open(io.BytesIO(raw))
            must = False
            if want == 1 and im.mode != "L":
                im, must = im.convert("L"), True
            elif want == 3 and im.mode == "L":
                im, must = im.convert("RGB"), True
            if im.mode not in ("L", "RGB"):
                continue
            data = None
            for progressive in (False, True):
                buf = io.BytesIO()
                im.save(buf, "JPEG", quality=quality, subsampling=2 if im.mode == "RGB" else 0,
                        optimize=True, progressive=progressive)
                if data is None or len(buf.getvalue()) < len(data):
                    data = buf.getvalue()
            if must or len(data) < len(raw):
                done[(raw, want)] = data
                doc.update_stream(x, data, compress=False)
                doc.xref_set_key(x, "Filter", "/DCTDecode")
                doc.xref_set_key(x, "DecodeParms", "null")
        except Exception:  # noqa: BLE001 - this image stays as MuPDF wrote it
            continue


# libjpeg's luminance table at quality 50, in natural order: a JPEG's own table, compared
# with this one scaled, tells the quality it was saved at.
_STD_LUMA = (16, 11, 10, 16, 24, 40, 51, 61, 12, 12, 14, 19, 26, 58, 60, 55,
             14, 13, 16, 24, 40, 57, 69, 56, 14, 17, 22, 29, 51, 87, 80, 62,
             18, 22, 37, 56, 68, 109, 103, 77, 24, 35, 55, 64, 81, 104, 113, 92,
             49, 64, 78, 87, 103, 121, 120, 101, 72, 92, 95, 98, 112, 100, 103, 99)


def jpeg_quality(im):
    """The quality a JPEG was saved at, read back from its luminance table: exact for
    libjpeg and Pillow files (checked at 30-95), close for cameras and editors."""
    try:
        qt = list(im.quantization[0])
    except Exception:  # noqa: BLE001
        return None
    if len(qt) != 64:
        return None
    best, best_q = None, None
    for q in range(1, 101):
        s = 5000 / q if q < 50 else 200 - 2 * q
        err = sum((a - min(255, max(1, int((v * s + 50) // 100)))) ** 2 for a, v in zip(qt, _STD_LUMA))
        if best is None or err < best:
            best, best_q = err, q
    return best_q


def looks_like_photo(im):
    """Continuous tone: thousands of colours and few runs of identical pixels. Screenshots,
    charts, logos and line art are mostly flat (screenshots: over half the neighbouring
    pixels the same) and stay lossless, since JPEG rings round every letter of them."""
    import numpy as np
    grey = im.mode == "L"
    t = im.copy() if grey else im.convert("RGB")
    t.thumbnail((480, 480))
    a = np.asarray(t).astype(np.int32)
    if a.shape[0] < 16 or a.shape[1] < 16:
        return False
    if grey:
        same = (np.diff(a, axis=1) == 0).mean()
        return same < 0.30 and len(np.unique(a)) > 120
    same = (np.abs(np.diff(a, axis=1)).sum(axis=2) == 0).mean()
    colours = len(np.unique((a[..., 0] << 16) | (a[..., 1] << 8) | a[..., 2]))
    return same < 0.30 and colours > 3000


def _jpeg_bytes(im, quality):
    """The smaller of baseline and progressive, 4:2:0 for colour (as fix_and_reencode_jpegs)."""
    data = None
    for progressive in (False, True):
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=quality, subsampling=2 if im.mode == "RGB" else 0,
                optimize=True, progressive=progressive)
        if data is None or len(buf.getvalue()) < len(data):
            data = buf.getvalue()
    return data


def recompress_photos(doc, dpi, quality, everything=False, deadline=None):
    """Two kinds of photo MuPDF's rewriter hands back as they came (19 Oct 2026):

    - A photo stored lossless (Flate: what PowerPoint, Canva and most exporters write for a
      PNG or a pasted picture). The rewriter keeps lossless images lossless, which is right
      for a screenshot and useless for a photo: an 8-slide deck of 1600x900 photos came back
      at 23.8 MB, 0 % at every level. A photo (see looks_like_photo) becomes a JPEG at the
      level's quality, resampled to its dpi when it is drawn above 1.5x that, as the rest.
    - A JPEG the rewriter does not resample (drawn within 1.5x of the level's dpi) is
      left at its own quality: a q92 photo stayed q92 at High (q42). One saved well above
      the level's quality is re-encoded at it, kept when that saves 15 % or more.

    everything: the rewriter is not run after this (it garbled the file's text, see
    main), so the JPEGs it would have resampled are resampled and re-encoded here too.

    Images that are soft masks, stencils, colour-keyed, Indexed, CMYK, 16-bit or carry a
    Decode array stay as they are. Returns the number of images changed."""
    try:
        from PIL import Image
        import numpy  # noqa: F401 - looks_like_photo
    except Exception:  # noqa: BLE001 - no Pillow/numpy: the rewriter alone, as before
        return 0
    thr = dpi * 1.5
    ppi, masks = {}, set()
    for page in doc:
        for item in page.get_images(full=True):
            x, smask = item[0], item[1]
            if smask:
                masks.add(smask)
            ppi.setdefault(x, 0.0)                     # 0: where it is drawn is not known
            try:
                box = page.get_image_bbox(item)
            except Exception:  # noqa: BLE001
                continue
            if box.is_empty or box.is_infinite or box.width <= 0 or box.height <= 0:
                continue
            here = min(item[2] / (box.width / 72.0), item[3] / (box.height / 72.0))
            ppi[x] = min(ppi[x], here) if ppi[x] else here   # the largest it is drawn decides
    changed = 0
    for x, at in sorted(ppi.items()):
        if deadline is not None and time.monotonic() > deadline:
            break
        try:
            if x in masks or doc.xref_get_key(x, "Subtype")[1] != "/Image":
                continue
            if doc.xref_get_key(x, "ImageMask")[1] == "true" or doc.xref_get_key(x, "Mask")[0] != "null":
                continue
            if doc.xref_get_key(x, "Decode")[0] != "null" or doc.xref_get_key(x, "BitsPerComponent")[1] != "8":
                continue
            comps = colour_components(doc, x)
            if comps not in (1, 3):
                continue
            filt = doc.xref_get_key(x, "Filter")
            raw = doc.xref_stream_raw(x)
            if filt[1] == "/DCTDecode":
                if at > thr and not everything:
                    continue                           # the rewriter resamples and re-encodes it
                im = Image.open(io.BytesIO(raw))
                if im.mode not in ("L", "RGB") or (im.mode == "L") != (comps == 1):
                    continue
                if at > thr:
                    s = dpi / at
                    im = im.resize((max(1, round(im.width * s)), max(1, round(im.height * s))), Image.LANCZOS)
                else:
                    q = jpeg_quality(im)
                    if q is None or q <= quality + 10:
                        continue
                data = _jpeg_bytes(im, quality)
                if len(data) > 0.85 * len(raw):
                    continue
                doc.update_stream(x, data, compress=False)
                doc.xref_set_key(x, "Width", str(im.width))
                doc.xref_set_key(x, "Height", str(im.height))
            elif filt[0] in ("name", "null") and filt[1] in ("/FlateDecode", "/LZWDecode", "/RunLengthDecode", "null"):
                pix = fitz.Pixmap(doc, x)
                if pix.alpha or pix.n != comps:
                    continue
                im = Image.frombytes("L" if comps == 1 else "RGB", (pix.width, pix.height), pix.samples)
                if not looks_like_photo(im):
                    continue
                if at > thr:
                    s = dpi / at
                    im = im.resize((max(1, round(im.width * s)), max(1, round(im.height * s))), Image.LANCZOS)
                data = _jpeg_bytes(im, quality)
                if len(data) > 0.9 * len(raw):
                    continue
                doc.update_stream(x, data, compress=False)
                doc.xref_set_key(x, "Width", str(im.width))
                doc.xref_set_key(x, "Height", str(im.height))
            else:
                continue
            doc.xref_set_key(x, "Filter", "/DCTDecode")
            doc.xref_set_key(x, "DecodeParms", "null")
            changed += 1
        except Exception:  # noqa: BLE001 - this image stays as it is
            continue
    return changed


# 1-bit pages for the last resort of a target search (see find_text_scans).
BITONAL_DPI = 150


def looks_like_text_scan(im):
    """A scanned page of text: paper and ink, nothing in between. Judged at about 120 dpi
    against the page's own local paper colour (the brightest pixels around each point), so
    a yellowed page or a shadow at one edge still counts while photos, grey drawings and
    colour pictures do not. Measured 2 Oct 2026 (grey, colour and Arabic 300-dpi scans
    against phone photos, screenshots and charts):
    - colour: at most 1.5 % of the pixels clearly coloured (a stamp, a signature) — photos
      28-83 %;
    - paper: at most 5 % of the page where the paper itself is darker than mid-grey —
      photos 51-96 %;
    - at least 80 % of the pixels within 25 levels of the paper (scans 87-93 %) and at most
      10 % that are neither paper nor ink (anti-aliased edges: scans 4-7 %)."""
    from PIL import Image, ImageFilter
    import numpy as np
    if im.mode == "RGB":
        c = im.copy()
        c.thumbnail((400, 400))   # coloured area is a share of the page: a small copy will do
        a = np.asarray(c).astype(np.int16)
        if ((a.max(axis=2) - a.min(axis=2)) > 48).mean() > 0.015:
            return False
    t = im.convert("L")
    t.thumbnail((1400, 1400))
    g = np.asarray(t).astype(np.int16)
    # The paper colour: the brightest pixel of each 4x4 block, smoothed, scaled back up
    # (a 9x9 MaxFilter did the same in 1.1 s a page; this takes 0.02 s).
    h4, w4 = g.shape[0] // 4 * 4, g.shape[1] // 4 * 4
    if h4 < 8 or w4 < 8:
        return False
    blocks = g[:h4, :w4].reshape(h4 // 4, 4, w4 // 4, 4).max(axis=(1, 3)).astype(np.uint8)
    paper = Image.fromarray(blocks).filter(ImageFilter.BoxBlur(3)).resize((g.shape[1], g.shape[0]), Image.BILINEAR)
    paper = np.asarray(paper).astype(np.int16)
    if (paper < 150).mean() > 0.05:
        return False
    off = paper - g
    near = (off <= 25).mean()
    mid = ((off > 25) & (g > 110)).mean()
    ink = (g <= 110).mean()
    return near >= 0.80 and mid <= 0.10 and 0.003 <= ink <= 0.30


def binarize(gray):
    """Ink black, paper white: below the local paper colour by a margin (an uneven or
    yellowed page keeps a clean background) and darker than Otsu's split plus a little
    (no speckles from paper grain)."""
    from PIL import Image, ImageFilter
    import numpy as np
    a = np.asarray(gray).astype(np.int16)
    hist = np.bincount(a.ravel(), minlength=256).astype(np.float64)
    p = hist / hist.sum()
    w = np.cumsum(p)
    mu = np.cumsum(p * np.arange(256))
    between = (mu[-1] * w - mu) ** 2 / (w * (1 - w) + 1e-12)
    otsu = int(np.argmax(between))
    r = max(8, gray.width // 50)
    local = np.asarray(gray.filter(ImageFilter.BoxBlur(r))).astype(np.int16)
    ink = (a < local - 18) & (a < otsu + 25)
    return Image.fromarray(np.where(ink, 0, 255).astype(np.uint8)).convert("1")


def find_text_scans(src, dpi=BITONAL_DPI):
    """Scanned text pages (grey or colour images that are only paper and ink) as 1-bit
    images at 150 dpi: {xref: (flate data, width, height)}. Photos, drawings and
    everything else are left out. Worked out once per request; every probe that uses it
    only swaps the streams in (bitonal_scans).

    Only the target search reaches for this, after even its floor (75 dpi, JPEG 38) has
    missed. Measured 2 Oct 2026 on 300-dpi scans, floor -> 1-bit (tesseract word recall
    on page 1 in brackets): grey, 6 pages 364 -> 277 KB (0.98 -> 1.00); colour with a
    stamp, 4 pages 255 -> 186 KB (0.92 -> 0.91, the stamp turns black), so a 200 KB
    target is met; Arabic, 3 pages 127 -> 98 KB (0.68 -> 0.93), so 100 KB is met. Written
    as 1-bit Flate: CCITT G4 from Pillow was only 7 % smaller, needs libtiff and a single
    strip, and MuPDF stopped decoding Pillow's multi-strip G4 halfway down the page."""
    try:
        import zlib
        from PIL import Image
        import numpy  # noqa: F401 - looks_like_text_scan and binarize need it
    except Exception:  # noqa: BLE001 - no Pillow/numpy: no 1-bit rung
        return {}
    found, seen = {}, set()
    doc = fitz.open(src)
    try:
        for page in doc:
            for item in page.get_images(full=True):
                x, smask, width, height, bpc = item[0], item[1], item[2], item[3], item[4]
                if x in seen:
                    continue
                seen.add(x)
                try:
                    if smask or bpc != 8 or colour_components(doc, x) not in (1, 3):
                        continue
                    if doc.xref_get_key(x, "Decode")[0] != "null" or doc.xref_get_key(x, "ImageMask")[1] == "true":
                        continue
                    # Where the page draws it (get_image_rects took 0.3 s a page; this 0.01 s)
                    box = page.get_image_bbox(item)
                    if box.is_empty or box.is_infinite:
                        continue
                    w_pt, h_pt = box.width, box.height
                    if w_pt < 72 or h_pt < 72:   # a logo or an icon, not a page
                        continue
                    if min(width / (w_pt / 72.0), height / (h_pt / 72.0)) < dpi * 0.9:
                        continue                 # already too coarse to make good 1-bit
                    tw, th = int(round(w_pt / 72.0 * dpi)), int(round(h_pt / 72.0 * dpi))
                    im = None
                    if doc.xref_get_key(x, "Filter")[1] == "/DCTDecode":
                        # A JPEG decodes straight at half or quarter size (libjpeg's scaled
                        # IDCT): a 300-dpi page in 0.05 s instead of 0.3 s at full size.
                        im = Image.open(io.BytesIO(doc.xref_stream_raw(x)))
                        if im.mode not in ("L", "RGB"):
                            im = None
                        else:
                            im.draft(im.mode, (int(tw * 0.95), int(th * 0.95)))
                            im = im.convert(im.mode)
                    if im is None:
                        pix = fitz.Pixmap(doc, x)
                        if pix.alpha or pix.n not in (1, 3):
                            continue
                        im = Image.frombytes("L" if pix.n == 1 else "RGB", (pix.w, pix.h), pix.samples)
                    if not looks_like_text_scan(im):
                        continue
                    gray = im.convert("L")
                    if tw < gray.width:
                        gray = gray.resize((tw, th), Image.LANCZOS)
                    bw = binarize(gray)
                    found[x] = (zlib.compress(bw.tobytes(), 9), bw.width, bw.height)
                except Exception:  # noqa: BLE001 - this image stays as it is
                    continue
    finally:
        doc.close()
    return found


def bitonal_scans(doc, scans):
    """Swaps in the 1-bit versions find_text_scans made (same file, so same xrefs)."""
    for x, (data, w, h) in scans.items():
        doc.update_stream(x, data, compress=False)
        for k, v in (("Filter", "/FlateDecode"), ("DecodeParms", "null"), ("Width", str(w)), ("Height", str(h)),
                     ("BitsPerComponent", "1"), ("ColorSpace", "/DeviceGray"), ("Interpolate", "null")):
            doc.xref_set_key(x, k, v)


def compress_once(src, out, dpi, quality, lossy_all=False, scans=None, rewriter=True):
    """rewriter=False: the images are done by recompress_photos alone (see main)."""
    doc = fitz.open(src)
    try:
        before = set(range(1, doc.xref_length()))
        unshare_resources(doc)
        if scans:
            bitonal_scans(doc, scans)
        recompress_photos(doc, dpi, quality, everything=not rewriter)
        if rewriter:
            doc.rewrite_images(options=rewrite_options(dpi, quality, lossy_all))
        drop_inherited_resources(doc)
        fix_and_reencode_jpegs(doc, before, quality)
        try:
            doc.subset_fonts()
        except Exception:  # noqa: BLE001 - fonts stay as they are
            pass
        # garbage=4 also merges identical objects. Every page has its own /Resources (see
        # unshare_resources), so the rewriter makes one copy of an image per page that uses
        # it: a letterhead on all 11 pages of a LibreOffice Arabic letter came out 11 times
        # (82 of 139 KB at High), and Low and Medium returned the file unchanged (2 Oct 2026).
        # The copies are byte-identical and merge back into one.
        doc.save(out, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, use_objstms=1)
    finally:
        doc.close()
    return os.path.getsize(out)


def lossless_once(src, out):
    """The file with every image left alone: fonts subset, identical objects merged,
    streams deflated, objects packed into object streams. The fallback for a file whose
    image pass changed its text or gained nothing (see main)."""
    doc = fitz.open(src)
    try:
        try:
            doc.subset_fonts()
        except Exception:  # noqa: BLE001 - fonts stay as they are
            pass
        doc.save(out, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, use_objstms=1)
    finally:
        doc.close()
    return os.path.getsize(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("out")
    ap.add_argument("--level", type=int, default=1)
    ap.add_argument("--target", type=int, default=0)
    ap.add_argument("--force", action="store_true", help="compress even without such text or form")
    ap.add_argument("--budget", type=float, default=0, help="seconds the API lets this process run (0 = no limit)")
    a = ap.parse_args()
    started = time.monotonic()

    try:
        doc = fitz.open(a.src)
        if doc.needs_pass:
            print(json.dumps({"ok": False, "reason": "password"})); sys.exit(11)
        if looks_damaged(doc):
            print(json.dumps({"ok": False, "reason": "damaged"})); sys.exit(12)
        needed = a.force or needs_this_engine(doc)
        doc.close()
    except SystemExit:
        raise
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": f"open: {e}"})); sys.exit(2)
    if not needed:
        print(json.dumps({"ok": False, "reason": "not-needed"})); sys.exit(10)

    in_size = os.path.getsize(a.src)
    probed = {}         # (dpi, quality, lossy_all, bitonal) -> (path, size)
    spent = []          # seconds per probe, to tell whether one more still fits the budget
    scans = {}          # find_text_scans, worked out when the second last resort is reached

    def probe(setting):
        if setting not in probed:
            path = f"{a.out}.{len(probed)}.tmp"
            t = time.monotonic()
            dpi_, q_, lossy_all_, bitonal_ = setting
            probed[setting] = (path, compress_once(a.src, path, dpi_, q_, lossy_all_, scans if bitonal_ else None))
            spent.append(time.monotonic() - t)
        return probed[setting][1]

    def changes_anything(fallback):
        """Whether a last resort can make this file any different from what was tried:
        lossless colour or grey images for the first, text scans for the second."""
        if not fallback[1]:
            d = fitz.open(a.src)
            try:
                for x in range(1, d.xref_length()):
                    if (d.xref_get_key(x, "Subtype")[1] == "/Image" and d.xref_get_key(x, "BitsPerComponent")[1] == "8"
                            and d.xref_get_key(x, "Filter")[1] not in ("/DCTDecode", "/JPXDecode")):
                        return True
                return False
            finally:
                d.close()
        scans.update(find_text_scans(a.src))
        return bool(scans)

    def time_for_one_more():
        """The optional probes (the lossy retry, the refinement) only run when one more probe
        of the kind already made still ends well inside the time the API gave this process;
        past it the API would kill the process and lose the file already found."""
        if a.budget <= 0 or not spent:
            return True
        return (time.monotonic() - started) + 1.3 * max(spent) < a.budget - 3

    def search(mode):
        """Index of the first (best-looking) rung that fits with this mode (lossy_all,
        bitonal), or None: rung sizes fall down the ladder, so it is a bisection."""
        lo, hi, best = 0, len(TARGET_LADDER) - 1, None
        while lo <= hi:
            mid = (lo + hi) // 2
            if probe(TARGET_LADDER[mid] + mode) <= a.target:
                best, hi = mid, mid - 1
            else:
                lo = mid + 1
        return best

    extra = {}
    try:
        if a.target > 0 and in_size <= a.target:
            # Already under the size asked for: the Low level at most, never less.
            dpi, quality = LEVELS[0]
            out_size = compress_once(a.src, a.out, dpi, quality)
            met = True
        elif a.target > 0:
            floor = len(TARGET_LADDER) - 1
            mode = (False, False)
            best = search(mode)
            # Even the floor misses: two last resorts, each tried at the floor first and
            # searched up the ladder only when it fits there. The levels never use them.
            # 1. Screenshots and charts as JPEG too. Measured 2 Oct 2026: a 1.2 MB review
            #    with a full-page screenshot asked to fit 100 KB had a 196 KB floor, 158 KB
            #    of it the screenshot kept lossless; as JPEG (what Ghostscript's own filter
            #    choice does with it) the floor is 84 KB and the search fits at 93 KB.
            # 2. Scanned text pages as 1-bit at 150 dpi (find_text_scans): about a quarter
            #    smaller than the floor's 75-dpi JPEG, and easier to read.
            for fallback in ((True, False), (True, True)):
                if best is not None or not time_for_one_more():
                    break
                if not changes_anything(fallback):
                    continue
                if probe(TARGET_LADDER[floor] + fallback) <= a.target:
                    mode = fallback
                    best = search(mode)
            pick = TARGET_LADDER[best if best is not None else floor] + mode
            if best is None:
                # Nothing fits: the smallest file made, reported as not met.
                pick = min(probed, key=lambda k: probed[k][1])
            elif best > 0:
                # The rung above missed and this one fits, often with a third of the
                # budget unused (12 MB of phone photos asked for 500 KB: 361 KB at 115 dpi,
                # while 130 dpi made 524 KB; now 479 KB at 127 dpi). One probe in between,
                # interpolated on the log of the size towards 95 % of the target; a second,
                # halfway, if it missed.
                miss = TARGET_LADDER[best - 1] + mode
                s_fit, s_miss = probed[pick][1], probed[miss][1]
                if s_miss > s_fit > 0:
                    t = (math.log(0.95 * a.target) - math.log(s_fit)) / (math.log(s_miss) - math.log(s_fit))
                    for step in (t, t / 2):
                        if step < 0.2 or not time_for_one_more():
                            break
                        cand = (int(round(pick[0] + step * (miss[0] - pick[0]))),
                                int(round(pick[1] + step * (miss[1] - pick[1])))) + mode
                        if cand[:2] == pick[:2] or cand[:2] == miss[:2]:
                            break
                        if probe(cand) <= a.target:
                            pick = cand
                            break
            shutil.move(probed[pick][0], a.out)
            for key, (path, _) in probed.items():
                if key != pick and os.path.exists(path):
                    os.remove(path)
            (dpi, quality), out_size = pick[:2], os.path.getsize(a.out)
            met = out_size <= a.target
            extra = {"lossy_all": pick[2], "bitonal": pick[3]}
        else:
            dpi, quality = LEVELS.get(a.level, LEVELS[1])
            try:
                out_size = compress_once(a.src, a.out, dpi, quality)
            except Exception as e:  # noqa: BLE001 - the lossless pass below, not Ghostscript
                print(f"image pass failed: {e}", file=sys.stderr)
                out_size = in_size
            met = None
        unchanged = False
        if out_size < in_size and not same_text(a.src, a.out):
            # Never hand back a smaller file whose text reads differently: the whole
            # point of this engine is that copy and search keep working.
            out_size = in_size
            # MuPDF's rewriter garbles some files' text (see the lossless pass below), and
            # only once it resamples a picture drawn inside a form: a Chrome deck's first
            # page read «\x00V» and «ي ح ي» at Medium and High, so those levels came back
            # 1 % smaller while Low made it 55 % (19 Oct 2026). The same dpi and quality
            # without the rewriter: the photos and JPEGs done by recompress_photos alone.
            if dpi and time_for_one_more():
                print("text changed after compression; trying without the rewriter", file=sys.stderr)
                path = a.out + ".own.tmp"
                try:
                    lossy = bool(extra.get("lossy_all"))
                    size = compress_once(a.src, path, dpi, quality, lossy, scans if extra.get("bitonal") else None,
                                         rewriter=False)
                    if size <= in_size * 0.99 and same_text(a.src, path):
                        shutil.move(path, a.out)
                        out_size = size
                        extra = {**extra, "rewriter": False}
                except Exception as e:  # noqa: BLE001 - the lossless pass then
                    print(f"own image pass failed: {e}", file=sys.stderr)
                finally:
                    if os.path.exists(path):
                        os.remove(path)
                if a.target > 0:
                    met = out_size <= a.target
            if out_size >= in_size:
                print("text changed after compression; trying the lossless pass", file=sys.stderr)
        if out_size > in_size * 0.99 and time_for_one_more():
            # The image pass changed the text or gained nothing: the file without touching an
            # image. PyMuPDF 1.26.5's rewrite_images() loses the form XObjects that draw a
            # Type3 font's glyphs («cannot find XObject resource»), and every page of a
            # Chrome-on-Windows PDF with 100+ such fonts read differently, so the original
            # came back at every level (0 %); the lossless pass made the same 1,169 KB file
            # 497 KB (-57 %) with the text unchanged, in 1.2 s (10 Oct 2026).
            path = a.out + ".lossless.tmp"
            try:
                size = lossless_once(a.src, path)
                if size <= in_size * 0.99 and same_text(a.src, path):
                    shutil.move(path, a.out)
                    out_size, dpi, quality = size, None, None
                    extra = {"lossless": True}
            except Exception as e:  # noqa: BLE001 - the original then, as before
                print(f"lossless pass failed: {e}", file=sys.stderr)
            finally:
                if os.path.exists(path):
                    os.remove(path)
            if a.target > 0:
                met = out_size <= a.target
        if out_size > in_size * 0.99:  # under 1 % gained (Ghostscript's rule too): hand back the original untouched
            shutil.copyfile(a.src, a.out)
            out_size, unchanged = in_size, True
            if a.target > 0:
                met = out_size <= a.target
        print(json.dumps({"ok": True, "in": in_size, "out": out_size, "dpi": dpi, "quality": quality,
                          "target_met": met, "unchanged": unchanged,
                          "reduction": round((1 - out_size / in_size) * 100, 1), **extra}))
    except Exception as e:  # noqa: BLE001
        for path, _ in probed.values():
            if os.path.exists(path) and path != a.out:
                os.remove(path)
        print(json.dumps({"ok": False, "error": str(e)})); sys.exit(2)


if __name__ == "__main__":
    main()
