#!/usr/bin/env python3
"""
pdf_to_images.py — render PDF pages to JPG / PNG / WebP with MuPDF, several pages at once.

Usage: python3 pdf_to_images.py in.pdf out_dir options.json
options: {"dpi": 150, "format": "jpg", "quality": 90, "pages": [1, 2, 5] | null (= all),
          "stem": "Report", "grayscale": false, "rotation": 0, "scale": null,
          "width": null, "height": null, "keepAspect": true, "background": "#ffffff",
          "workers": 4}
stdout: one JSON line {"total": n, "pages": [{"page": 1, "file": "Report-01.jpg", "w": .., "h": .., "bytes": ..}]}
exit 3 = the PDF needs a password; exit 4 = nothing to render (bad page list);
exit 5 = the file is damaged: MuPDF had to rebuild it and at least half of the pages came out
         empty (stdout still carries {"damaged": {"blank": n, "pages": m}});
exit 2 = any other failure (the caller falls back to the ImageMagick path).

WHY (measured 24 Sept 2026 on the live box): the ImageMagick path asked Ghostscript for EVERY
page of the PDF at the requested DPI, held all of them in memory, and only then wrote the
first image, on one core. The page-count call the page makes when a file is picked rendered
every page a second time at 72 DPI. MuPDF opens the document once per worker, renders one
page at a time and hands each pixmap straight to the encoder, so memory stays at one page
per worker and the pages of a long PDF are spread across cores.

Output choices, and why:
- The page is painted on white (alpha off): a PDF paints glyphs, not paper, and a
  transparent page turns into a black sheet in any dark viewer (the 4 Sept PNG bug).
- JPEG at quality >= 85 keeps full-resolution colour (4:4:4). With the default 4:2:0,
  coloured headings, links and stamps get soft fringes at 150 DPI; black text is unaffected
  either way. Progressive + optimised Huffman tables: smaller files at the same quality.
- Every file carries its DPI (JFIF density / PNG pHYs), so it prints at the page's real size.
- A pixel cap keeps a poster-sized page at 600 DPI from asking for gigabytes.
- Pages that are mostly picture (a scan, a full-page photo: images cover >= 25% of the page)
  are written with cheaper encoder settings, because the expensive ones buy almost nothing
  there (2 Oct 2026, one core): PNG at zlib level 3 instead of 6 is 3.0x faster on a 300 DPI
  scan for +2% bytes (2.0-2.6x faster, +6% on half-photo pages), and the whole PNG export of a
  50-page scan at 300 DPI took 92-106 s on a 4-core test box, past Cloudflare's 100 s: an error page.
  Text pages keep level 6, where it costs little and saves up to 30%. A JPEG of a page that is
  one big picture (>= 90%) is not progressive: 0.26 s -> 0.12 s per 300 DPI page, and the file
  is the same size (-2% at 300 DPI, +0.6% at 150); on text pages progressive saves 5-9% and stays.
  Measured on whole files (CPU seconds, 4 workers): 50-page scan at 300 DPI, PNG 203 -> 72 s
  (wall 92 -> 44 s on a busy 4-core box), JPG 17.7 -> 11.6 s; 60-page mixed report at 300 DPI
  PNG 41 -> 25 s. Text-only pages are encoded exactly as before.
- A damaged PDF (truncated download, broken xref) is rebuilt by MuPDF, and a page whose
  content was lost then renders as white paper. Until 2 Oct 2026 a 60-page PDF cut off after
  20 KB "converted" into 59 empty JPGs and one real one, sent as a success. Now when MuPDF had
  to rebuild the file and at least half of the pages are empty, the script reports the file
  as damaged instead (exit 5); a rebuilt file with a few empty pages converts as before.
"""
import json
import os
import re
import sys

import fitz  # PyMuPDF

MAX_PIXELS = 80_000_000  # ~ an A0 page at 200 DPI; beyond this the zoom is reduced
PICTURE_PAGE = 0.25      # image area / page area from which a page counts as a picture
SCAN_PAGE = 0.90         # ... and from which it is one big picture (a scan)


def _image_coverage(page):
    """Share of the page covered by images (0..1), from their placement boxes; ~1 ms a page."""
    try:
        area = abs(page.rect)
        if not area:
            return 0.0
        covered = 0.0
        for info in page.get_image_info():
            r = fitz.Rect(info["bbox"]) & page.rect
            if not r.is_empty:
                covered += abs(r)
        return min(1.0, covered / area)
    except Exception:
        return 0.0


def _parse_color(value):
    if not value:
        return (255, 255, 255)
    v = value.strip().lstrip("#")
    if re.fullmatch(r"[0-9a-fA-F]{6}", v):
        return tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))
    return (255, 255, 255)


def _render_one(doc, index, opts, out_dir, name_for, check_blank=False):
    from PIL import Image, ImageOps

    page = doc[index]
    coverage = _image_coverage(page)
    dpi = opts["dpi"]
    zoom = dpi / 72.0
    rect = page.rect  # the displayed frame: /Rotate and the CropBox are already applied
    px = rect.width * zoom * rect.height * zoom
    if px > MAX_PIXELS:
        zoom *= (MAX_PIXELS / px) ** 0.5
    bg = _parse_color(opts.get("background"))
    white = bg == (255, 255, 255)
    gray = bool(opts.get("grayscale"))

    if white:
        cs = fitz.csGRAY if gray else fitz.csRGB
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=cs, alpha=False, annots=True)
        mode = "L" if gray else "RGB"
        img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    else:
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csRGB, alpha=True, annots=True)
        rgba = Image.frombytes("RGBA", (pix.width, pix.height), pix.samples)
        img = Image.new("RGB", rgba.size, bg)
        img.paste(rgba, mask=rgba.getchannel("A"))
        if gray:
            img = img.convert("L")
    pix = None

    rot = int(opts.get("rotation") or 0) % 360
    if rot:
        img = img.rotate(-rot, expand=True, fillcolor=(255 if gray else bg))

    scale = opts.get("scale")
    tw, th = opts.get("width"), opts.get("height")
    if scale and int(scale) != 100:
        f = max(10, min(400, int(scale))) / 100.0
        img = img.resize((max(1, round(img.width * f)), max(1, round(img.height * f))), Image.LANCZOS)
    elif tw or th:
        if opts.get("keepAspect", True):
            if tw and th:
                img = ImageOps.contain(img, (int(tw), int(th)), Image.LANCZOS)
            elif tw:
                img = img.resize((int(tw), max(1, round(img.height * int(tw) / img.width))), Image.LANCZOS)
            else:
                img = img.resize((max(1, round(img.width * int(th) / img.height)), int(th)), Image.LANCZOS)
        else:
            img = img.resize((int(tw or img.width), int(th or img.height)), Image.LANCZOS)

    # A page with no colour at all (black text on white, a grey scan) is stored as greyscale:
    # the same pixels in a third of the bytes, and a JPEG then has no chroma to smear.
    if img.mode == "RGB":
        from PIL import ImageChops
        r, g, b = img.split()
        if ImageChops.difference(r, g).getbbox() is None and ImageChops.difference(g, b).getbbox() is None:
            img = r

    # Only asked for when MuPDF had to rebuild the file: a page that lost its content is paper.
    blank = False
    if check_blank:
        ext = img.getextrema()
        blank = all(lo >= 250 for lo, _ in ext) if isinstance(ext[0], tuple) else ext[0] >= 250

    fmt = opts["format"]
    q = int(opts.get("quality") or 90)
    q = max(1, min(100, q))
    path = os.path.join(out_dir, name_for(index + 1))
    if fmt == "png":
        img.save(path, "PNG", compress_level=3 if coverage >= PICTURE_PAGE else 6, dpi=(dpi, dpi))
    elif fmt == "webp":
        img.save(path, "WEBP", quality=q, method=4)
    else:
        img.save(path, "JPEG", quality=q, subsampling=0 if q >= 85 else 2,
                 progressive=coverage < SCAN_PAGE, optimize=True, dpi=(dpi, dpi))
    out = {"page": index + 1, "file": os.path.basename(path), "w": img.width, "h": img.height,
           "bytes": os.path.getsize(path)}
    if blank:
        out["blank"] = True
    return out


def _worker(args):
    pdf_path, indices, opts, out_dir, pad = args
    doc = fitz.open(pdf_path)
    if doc.needs_pass:
        doc.authenticate("")
    stem = opts.get("stem") or "page"
    ext = {"png": ".png", "webp": ".webp"}.get(opts["format"], ".jpg")
    single = opts.get("_single", False)
    check_blank = bool(doc.is_repaired)

    def name_for(n):
        return f"{stem}{ext}" if single else f"{stem}-{n:0{pad}d}{ext}"

    return [_render_one(doc, i, opts, out_dir, name_for, check_blank) for i in indices]


def main(pdf_path, out_dir, options_path):
    with open(options_path, "r", encoding="utf-8") as f:
        opts = json.load(f)
    opts["format"] = {"jpeg": "jpg", "jpg": "jpg", "png": "png", "webp": "webp"}.get(
        str(opts.get("format") or "jpg").lower(), "jpg")
    opts["dpi"] = max(36, min(600, int(opts.get("dpi") or 150)))

    doc = fitz.open(pdf_path)
    # An owner password only restricts editing; the pages still open with the empty user
    # password, exactly as every viewer opens them. Only a real user password stops here.
    if doc.needs_pass and not doc.authenticate(""):
        print("encrypted", file=sys.stderr)
        return 3
    total = doc.page_count
    wanted = opts.get("pages")
    if not wanted and opts.get("pageNumber"):
        wanted = [int(opts["pageNumber"])]
    elif not wanted and (opts.get("startPage") or opts.get("endPage")):
        first = max(1, int(opts.get("startPage") or 1))
        last = min(total, int(opts.get("endPage") or total))
        wanted = list(range(first, last + 1))
    indices = sorted({p - 1 for p in wanted if 1 <= p <= total}) if wanted else list(range(total))
    doc.close()
    if not indices:
        print("no pages", file=sys.stderr)
        return 4

    opts["_single"] = len(indices) == 1
    pad = max(2, len(str(total)))
    os.makedirs(out_dir, exist_ok=True)
    workers = max(1, min(int(opts.get("workers") or 4), os.cpu_count() or 1, len(indices)))
    # Interleaved slices keep the workers busy evenly when page cost varies (scans vs text).
    jobs = [(pdf_path, indices[w::workers], opts, out_dir, pad) for w in range(workers)]
    if workers == 1:
        parts = [_worker(jobs[0])]
    else:
        import multiprocessing as mp
        ctx = mp.get_context("fork") if hasattr(os, "fork") else mp.get_context("spawn")
        with ctx.Pool(workers) as pool:
            parts = pool.map(_worker, jobs)
    pages = sorted((p for part in parts for p in part), key=lambda p: p["page"])
    blank = sum(1 for p in pages if p.get("blank"))
    if blank and blank * 2 >= len(pages):
        for p in pages:
            try:
                os.remove(os.path.join(out_dir, p["file"]))
            except OSError:
                pass
        print(json.dumps({"total": total, "damaged": {"blank": blank, "pages": len(pages)}}), flush=True)
        print(f"damaged: {blank} of {len(pages)} pages empty after repair", file=sys.stderr)
        return 5
    print(json.dumps({"total": total, "pages": pages}), flush=True)
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print(__doc__, file=sys.stderr)
        sys.exit(2)
    try:
        sys.exit(main(sys.argv[1], sys.argv[2], sys.argv[3]))
    except Exception as ex:  # the caller falls back to ImageMagick
        print(f"{type(ex).__name__}: {ex}", file=sys.stderr)
        sys.exit(2)
