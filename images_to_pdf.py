#!/usr/bin/env python3
"""
images_to_pdf.py — build a PDF from images the way a person expects a "PDF of my photos"
to look: real paper pages (A4/Letter, or the image's own size), the picture fitted and
centred, phone photos upright, transparency on white, and no second lossy generation
when the source is already a JPEG.

Usage: python3 images_to_pdf.py out.pdf manifest.json
manifest: {"quality": "high|medium|low", "pageSize": "a4|letter|fit",
           "margin": "none|small|normal", "orientation": "auto|portrait|landscape",
           "files": ["/tmp/x/1.jpg", ...]}
Prints one JSON line on success ({"pages": n, "skipped": [...]}), exit 0.
Exit 11 = an input this engine cannot decode (HEIC/HEIF without pillow-heif, corrupt
file): the caller falls back to its ImageMagick engine for the whole batch. Exit 2 = failure.

WHY (measured 2026-09-08 on the live tool): the previous engine made every page exactly
the image's pixel size divided by a DPI that CHANGED WITH THE QUALITY CARD — a 300x200
screenshot became a 1.0x0.67-inch page, a phone photo a 10x13-inch page at "high" and a
12.5x16.7-inch page at "low", and one PDF mixed five page sizes. It also re-encoded every
JPEG at q96 (a 9.6 MB photo came back as 11.3 MB). Here the quality card only decides the
image's resolution/compression; the page size is a page size.

2026-10-02, measured on 50 generated 12 MP phone photos (4032x3024, EXIF orientations 1/3/6/8)
and a set of edge-case files:
- An upside-down photo (EXIF orientation 3) and a square sideways one came out upside down /
  sideways: the "untouched JPEG" test compared the decoded size with the file's size, and a
  180-degree turn (or a quarter turn of a square) does not change the size. The orientation tag
  is now read from the header, and the original bytes are embedded and TURNED ON THE PAGE
  (insert_image rotate), so sideways photos stay lossless too instead of being re-encoded —
  which is what made 156 MB of photos come back as a 175 MB PDF at "high".
- A 16-bit greyscale PNG (scanners, medical/astronomy exports) came out white: Pillow's
  I;16 -> L conversion clips at 255 instead of scaling. It is scaled now.
- Phone JPEGs that carry a second picture (MPO: Samsung/iPhone depth or preview frames) were
  re-encoded because Pillow calls them "MPO", not "JPEG". The first frame's bytes are embedded.
- Embedded JPEGs no longer carry EXIF/XMP (camera serial, GPS position of the photo) into a
  PDF people send around; ICC (APP2) and Adobe (APP14) segments are kept, the pixels are untouched.
- Images are prepared on a small thread pool (Pillow releases the GIL while decoding/encoding),
  JPEG downscales decode at reduced DCT scale (draft mode), and lossless images are handed to
  MuPDF without Pillow's slow "optimize" pass, since MuPDF re-compresses them anyway.
"""
import io
import json
import os
import struct
import sys
from concurrent.futures import ThreadPoolExecutor

from PIL import Image, ImageOps
import fitz  # PyMuPDF

try:  # optional: HEIC/HEIF straight from iPhones. Without it the caller falls back (exit 11).
    from pillow_heif import register_heif_opener  # type: ignore
    register_heif_opener()
except Exception:  # noqa: BLE001
    pass

Image.MAX_IMAGE_PIXELS = 300_000_000

PAGE_SIZES = {"a4": (595.276, 841.89), "letter": (612.0, 792.0)}
MARGINS = {"none": 0.0, "small": 18.0, "normal": 36.0}  # points
# quality -> (max long side in px or None, jpeg quality, dpi used for "fit" pages without metadata)
PROFILES = {
    "high": (None, 92, 150),
    "medium": (2600, 82, 150),
    "low": (1600, 65, 96),
}
LOSSLESS_SOURCES = {"PNG", "GIF", "BMP", "TIFF", "PCX", "TGA"}
JPEG_FORMATS = {"JPEG", "MPO"}
# EXIF orientation -> anti-clockwise turn for insert_image (PyMuPDF's sign). Mirrored
# orientations (2, 4, 5, 7) are decoded and transposed by Pillow instead.
TURNS = {1: 0, 3: 180, 6: -90, 8: 90}
# At medium/low an in-cap JPEG is re-encoded only when that saves at least this much;
# otherwise the original (better) bytes are smaller or close enough.
REENCODE_GAIN = 0.8
WORKERS = max(1, min(4, os.cpu_count() or 1))


def jpeg_clean_bytes(data):
    """The first JPEG image of the file (MPO extra frames and their APP2 index dropped)
    without the EXIF/XMP (APP1), Photoshop/IPTC (APP13) and comment segments. Entropy-coded data and the
    tables are copied byte for byte, so the pixels are exactly the original's.
    Returns None when the file is not a well-formed JPEG; the caller then re-encodes."""
    if len(data) < 4 or data[0:2] != b"\xff\xd8":
        return None
    out = bytearray(b"\xff\xd8")
    i = 2
    n = len(data)
    while i + 4 <= n:
        if data[i] != 0xFF:
            return None
        marker = data[i + 1]
        if marker == 0xFF:  # fill byte
            i += 1
            continue
        seg_len = struct.unpack(">H", data[i + 2:i + 4])[0]
        if seg_len < 2 or i + 2 + seg_len > n:
            return None
        if marker == 0xDA:  # SOS: the rest up to the first EOI is the image
            end = data.find(b"\xff\xd9", i + 2 + seg_len)
            # inside entropy-coded data 0xFF is always followed by 0x00 or RSTn, so the
            # first FFD9 after the first scan header is this image's end (EXIF thumbnails,
            # which also contain FFD9, live in APP1 and were skipped by length above)
            if end < 0:
                return None
            out += data[i:end + 2]
            return bytes(out)
        # APP2 "MPF" is the multi-picture index of an MPO; its frames are being dropped
        drop = marker == 0xE1 or marker == 0xED or marker == 0xFE or \
            (marker == 0xE2 and data[i + 4:i + 8] == b"MPF\x00")
        if not drop:
            out += data[i:i + 2 + seg_len]
        i += 2 + seg_len
    return None


def _stated_dpi(im):
    try:
        d = im.info.get("dpi")
        # 72 is JFIF's "unspecified" default (Pillow writes it on every JPEG it saves), so
        # a phone photo would become a 41x55-inch page in "fit" mode — treat it as unknown.
        if d and d[0] and 30 <= float(d[0]) <= 1200 and round(float(d[0])) != 72:
            return float(d[0])
    except Exception:  # noqa: BLE001
        pass
    return None


def _orientation(im):
    try:
        return int(im.getexif().get(0x0112) or 1)
    except Exception:  # noqa: BLE001
        return 1


def _to_page_mode(im):
    """Any decoded image -> RGB or L, transparency composited on white paper."""
    if im.mode in ("RGBA", "LA", "PA") or (im.mode == "P" and "transparency" in im.info):
        rgba = im.convert("RGBA")
        bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        return Image.alpha_composite(bg, rgba).convert("RGB")
    if im.mode == "CMYK":
        icc = im.info.get("icc_profile")
        if icc:
            try:
                from PIL import ImageCms
                src = ImageCms.ImageCmsProfile(io.BytesIO(icc))
                dst = ImageCms.createProfile("sRGB")
                return ImageCms.profileToProfile(im, src, dst, outputMode="RGB")
            except Exception:  # noqa: BLE001
                pass
        return im.convert("RGB")
    if im.mode in ("I;16", "I;16B", "I;16L", "I;16N"):
        # Pillow's I;16 -> L conversion CLIPS at 255 (a 16-bit scan came out white); scale it.
        return im.convert("I").point(lambda v: v * (1.0 / 256.0)).convert("L")
    if im.mode in ("I", "F"):
        lo, hi = im.getextrema()
        if hi is None or lo is None or hi == lo:
            return im.convert("L")
        scale = 255.0 / (hi - lo)
        return im.point(lambda v: (v - lo) * scale).convert("L")
    if im.mode == "1":
        return im.convert("L")
    if im.mode not in ("RGB", "L"):
        return im.convert("RGB")
    return im


def is_graphic(im):
    """Few colours = screenshot / drawing / scan of a form: keep it lossless."""
    try:
        probe = im
        if max(im.size) > 1200:
            probe = im.copy()
            probe.thumbnail((1200, 1200))
    except Exception:  # noqa: BLE001
        probe = im
    try:
        return probe.getcolors(4096) is not None
    except Exception:  # noqa: BLE001
        return False


def _encode_jpeg(im, jpg_q, icc):
    buf = io.BytesIO()
    kw = {}
    if icc:
        kw["icc_profile"] = icc
    im.save(buf, "JPEG", quality=jpg_q, optimize=True, progressive=False,
            subsampling=0 if jpg_q >= 90 else 2, **kw)
    return buf.getvalue()


def _encode_png(im, icc):
    # compress_level 1: MuPDF decodes the PNG and Flate-compresses the pixels itself, so
    # Pillow's optimize pass (seconds on a 4K screenshot) bought nothing in the output.
    buf = io.BytesIO()
    kw = {}
    if icc:
        kw["icc_profile"] = icc
    im.save(buf, "PNG", compress_level=1, **kw)
    return buf.getvalue()


def prepare(path, quality, frame=0):
    """Decode/encode one image (one frame of it). Returns dict(data, w, h, rotate, dpi, kind)
    where w/h are the size as the reader sees it (after orientation, before downscale), or raises."""
    max_side, jpg_q, _ = PROFILES[quality]
    im = Image.open(path)
    try:
        if frame:
            im.seek(frame)
        fmt = (im.format or "").upper()
        dpi = _stated_dpi(im)
        icc = im.info.get("icc_profile")
        orient = _orientation(im)
        w0, h0 = im.size
        quarter = orient in (5, 6, 7, 8)
        disp_w, disp_h = (h0, w0) if quarter else (w0, h0)
        within_cap = not max_side or max(w0, h0) <= max_side

        # 1) Untouched JPEG: the original entropy-coded bytes, turned on the page.
        original = None
        if fmt in JPEG_FORMATS and within_cap and orient in TURNS and im.mode in ("RGB", "L"):
            with open(path, "rb") as f:
                original = jpeg_clean_bytes(f.read())
            if original is not None and quality == "high":
                return dict(data=original, w=disp_w, h=disp_h, rotate=TURNS[orient], dpi=dpi, kind="jpeg-passthrough")

        # 2) Decode. A JPEG that is going to be downscaled decodes at reduced DCT scale.
        if fmt in JPEG_FORMATS and max_side and not within_cap:
            scale = max_side / float(max(w0, h0))
            try:
                im.draft(im.mode if im.mode in ("RGB", "L") else "RGB",
                         (max(1, int(w0 * scale)), max(1, int(h0 * scale))))
            except Exception:  # noqa: BLE001
                pass
        if getattr(im, "is_animated", False):
            im = im.copy()  # the CURRENT frame: frame 0, or the TIFF page _jobs asked for
        try:
            im = ImageOps.exif_transpose(im) or im
        except Exception:  # noqa: BLE001
            pass
        im = _to_page_mode(im)
        if max_side and max(im.size) > max_side:
            s = max_side / float(max(im.size))
            im = im.resize((max(1, round(im.size[0] * s)), max(1, round(im.size[1] * s))), Image.LANCZOS)

        if fmt == "CMYK" or (icc and im.mode == "RGB" and fmt in JPEG_FORMATS and b"CMYK" in icc[:40]):
            icc = None  # converted to sRGB above; the CMYK profile no longer describes the pixels

        lossless = (fmt in LOSSLESS_SOURCES and (quality == "high" or is_graphic(im))) or \
                   (fmt == "WEBP" and is_graphic(im))
        if lossless:
            return dict(data=_encode_png(im, icc), w=disp_w, h=disp_h, rotate=0, dpi=dpi, kind="png")
        data = _encode_jpeg(im, jpg_q, icc if im.mode == "RGB" else None)
        if original is not None and len(data) > REENCODE_GAIN * len(original):
            return dict(data=original, w=disp_w, h=disp_h, rotate=TURNS[orient], dpi=dpi, kind="jpeg-passthrough")
        return dict(data=data, w=disp_w, h=disp_h, rotate=0, dpi=dpi, kind="jpeg")
    finally:
        try:
            im.close()
        except Exception:  # noqa: BLE001
            pass


def _jobs(files):
    """One (path, frame) per PDF page. Every page of a multi-page TIFF (what a scanner
    saves) becomes a page; until 5 Oct 2026 only the first one did, so a 3-page scan came
    out as a 1-page PDF with no warning. Animated GIF/WebP stay one page (first frame)."""
    out = []
    for p in files:
        frames = 1
        try:
            with Image.open(p) as im:
                if (im.format or "").upper() == "TIFF":
                    frames = max(1, int(getattr(im, "n_frames", 1) or 1))
        except Exception:  # noqa: BLE001
            pass  # prepare() reports what cannot be opened
        out.extend((p, i) for i in range(frames))
    return out


def _prepare_all(files, quality):
    """Yields ((path, frame), prepared-or-exception) in order, at most 2*WORKERS images in flight."""
    def job(p):
        try:
            return prepare(p[0], quality, p[1])
        except Exception as e:  # noqa: BLE001
            return e
    if WORKERS == 1 or len(files) == 1:
        for p in files:
            yield p, job(p)
        return
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        pending = []
        it = iter(files)
        for p in it:
            pending.append((p, ex.submit(job, p)))
            if len(pending) >= 2 * WORKERS:
                break
        while pending:
            p, fut = pending.pop(0)
            res = fut.result()
            nxt = next(it, None)
            if nxt is not None:
                pending.append((nxt, ex.submit(job, nxt)))
            yield p, res


def main(out_path, manifest_path):
    with open(manifest_path, "r", encoding="utf-8") as f:
        m = json.load(f)
    quality = (m.get("quality") or "high").lower()
    if quality not in PROFILES:
        quality = "high"
    page_size = (m.get("pageSize") or "a4").lower()
    margin = MARGINS.get((m.get("margin") or "none").lower(), 0.0)
    orientation = (m.get("orientation") or "auto").lower()
    files = m.get("files") or []
    if not files:
        print("no files", file=sys.stderr)
        return 2

    doc = fitz.open()
    skipped = []
    kinds = {}
    for (path, _frame), prep in _prepare_all(_jobs(files), quality):
        if isinstance(prep, Exception):
            print(f"cannot decode {os.path.basename(path)}: {prep}", file=sys.stderr)
            doc.close()
            return 11
        w_px, h_px = prep["w"], prep["h"]
        if page_size in PAGE_SIZES:
            pw, ph = PAGE_SIZES[page_size]
            land = (w_px > h_px) if orientation == "auto" else (orientation == "landscape")
            if land:
                pw, ph = ph, pw
        else:  # "fit": the page IS the image, at its own DPI (or the profile's default)
            use_dpi = prep["dpi"] or PROFILES[quality][2]
            pw, ph = w_px * 72.0 / use_dpi + 2 * margin, h_px * 72.0 / use_dpi + 2 * margin
            pw = max(36.0, min(pw, 14400.0)); ph = max(36.0, min(ph, 14400.0))
        page = doc.new_page(width=pw, height=ph)
        box = fitz.Rect(margin, margin, pw - margin, ph - margin)
        # Fitted and centred here, not by keep_proportion (19 Oct 2026): PyMuPDF 1.26 reads a
        # JPEG whose JFIF header gives an aspect ratio instead of a DPI (what Pillow writes,
        # and many phone and web pictures carry) as having no proportion to keep, and a
        # square photo was stretched over the whole A4 page.
        s = min(box.width / max(1, w_px), box.height / max(1, h_px))
        fw, fh = w_px * s, h_px * s
        x0, y0 = box.x0 + (box.width - fw) / 2.0, box.y0 + (box.height - fh) / 2.0
        page.insert_image(fitz.Rect(x0, y0, x0 + fw, y0 + fh), stream=prep["data"],
                          keep_proportion=False, rotate=prep["rotate"])
        kinds[prep["kind"]] = kinds.get(prep["kind"], 0) + 1
    if len(doc) == 0:
        doc.close()
        return 11
    first = os.path.splitext(os.path.basename(files[0]))[0] if files else "images"
    doc.set_metadata({"title": first if len(files) == 1 else "Images", "producer": "Confileo", "creator": "Confileo Images to PDF"})
    doc.save(out_path, garbage=3, deflate=True)
    n = len(doc)
    doc.close()
    print(json.dumps({"pages": n, "skipped": skipped, "kinds": kinds}))
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("usage: images_to_pdf.py out.pdf manifest.json", file=sys.stderr)
        sys.exit(2)
    try:
        sys.exit(main(sys.argv[1], sys.argv[2]))
    except Exception as e:  # noqa: BLE001
        print(f"images_to_pdf failed: {e}", file=sys.stderr)
        sys.exit(2)
