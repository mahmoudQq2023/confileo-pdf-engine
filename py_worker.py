#!/usr/bin/env python3
"""A warm python3 for the API's model-heavy scripts (23 Sept 2026).

Remove Background and the passport-photo measurement each started a fresh
python3 per request, imported rembg (and mediapipe), and loaded the 178 MB
isnet model before doing any work. Measured on the server: the API took 4.4 s
for a 300x120 PNG and 6.4 s for a 1200x1600 photo; the same cut-out in a python
that is already running takes 0.8 s and 2.0 s. PythonWorkerPool.cs keeps a few
of these processes alive and sends them jobs; each job runs one of the
ALLOWED scripts exactly as `python3 <script> <args…>` would, in-process.

Protocol, one JSON object per line:
  worker → API   {"ready": true}                       once, after warm-up
  API → worker   {"id": "…", "script": "remove_bg.py", "args": ["in", "out"]}
  worker → API   {"id": "…", "exit": 0, "stdout": "…", "stderr": "…", "ms": 812}

The script's own stdout is captured into "stdout"; everything else a library
prints, from Python or from C, goes to stderr, so it can never corrupt a reply.
What the script itself writes to sys.stderr is also returned in "stderr" (its
last 4,000 characters, 10 Oct 2026): Crop reads the page number of a refusal
from it and Sign & Stamp the word "encrypted", so a caller moving to the worker
keeps the same contract as `python3 <script>`.
"""
import contextlib
import io
import json
import os
import runpy
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
# cv_builder_export.py (24 Sept 2026): CV export/import paid ~1 s of imports (WeasyPrint alone
# 0.6 s, then PyMuPDF) on every request; the script only uses absolute paths, so it runs
# in-process like the others. (xhtml2pdf and reportlab left it on 28 Sept 2026.)
#
# The MuPDF file-in/file-out scripts (10 Oct 2026): each request paid 0.12-0.17 s to start
# python3 and import PyMuPDF for a job that itself takes 0.1-0.6 s (crop, page numbers,
# watermark, flatten, protect/unlock, sign, redact, organize). They take absolute paths,
# change no process-wide state at import or run time, and run here with the same exit
# codes, stdout and stderr as from the command line (checked one by one). pdf_split.py
# (forks) and pdf_form.py (a stdin pipe) stay out. Their C# callers still start their own
# process; a caller that wants the warm path asks PythonWorkerPool first.
PDF_SCRIPTS = {"pdf_crop.py", "pdf_page_numbers.py", "pdf_watermark.py", "pdf_flatten.py",
               "pdf_protect.py", "sign_pdf.py", "pdf_redact.py", "pdf_organize.py"}
# 19 Oct 2026: compress to a size (Pillow + MozJPEG, ~0.2-2 s), the Object Remover (LaMa: its
# ~6 s model load is paid once per worker through vision_cache) and PDF to PowerPoint (MuPDF).
# The Image Upscaler streams its progress, so it runs in its own process (ScriptRunner).
PDF_SCRIPTS = PDF_SCRIPTS | {"pdf_to_pptx.py"}
ALLOWED = {"remove_bg.py", "face_geometry.py", "cv_builder_export.py",
           "image_target_size.py", "image_inpaint.py"} | PDF_SCRIPTS

# Replies go to a private copy of the original stdout; fd 1 itself is pointed at
# stderr, so a C extension writing to fd 1 lands in the API's debug log instead.
_reply = os.fdopen(os.dup(1), "w", buffering=1, encoding="utf-8")
os.dup2(2, 1)
sys.stdout = sys.stderr


def _share_rembg_sessions():
    """Make rembg.new_session return one cached session per model name.

    Both scripts call new_session() on every run; a session IS the loaded model,
    so reusing it is the whole point of this process. Sessions are safe to reuse
    across calls (that is how rembg's own server uses them).
    """
    try:
        import rembg
    except Exception:
        return
    original = rembg.new_session
    cache = {}

    def new_session(model_name="u2net", *args, **kwargs):
        if args or kwargs:
            return original(model_name, *args, **kwargs)
        if model_name not in cache:
            cache[model_name] = original(model_name)
        return cache[model_name]

    rembg.new_session = new_session


def _warm():
    _share_rembg_sessions()
    for mod in ("numpy", "PIL.Image", "mediapipe", "weasyprint", "fitz", "docx", "pptx",
                "mozjpeg_lossless_optimization", "onnxruntime"):
        try:
            __import__(mod)
        except Exception:
            pass
    try:
        import rembg
        rembg.new_session("isnet-general-use")
    except Exception:
        traceback.print_exc()


class _Tee(io.StringIO):
    """Keeps what the script writes to sys.stderr for the reply and still logs it."""

    def __init__(self, log):
        super().__init__()
        self._log = log

    def write(self, text):
        try:
            self._log.write(text)
        except Exception:
            pass
        return super().write(text)


def _run(script, args):
    path = os.path.join(HERE, script)
    captured = io.StringIO()
    errors = _Tee(sys.__stderr__)
    code = 0
    saved_argv = sys.argv
    sys.argv = [path] + [str(a) for a in args]
    try:
        with contextlib.redirect_stdout(captured), contextlib.redirect_stderr(errors):
            try:
                runpy.run_path(path, run_name="__main__")
            except SystemExit as e:
                code = e.code if isinstance(e.code, int) else (0 if e.code is None else 1)
            except BaseException:
                traceback.print_exc()
                code = 1
    finally:
        sys.argv = saved_argv
    if script in PDF_SCRIPTS:
        _release_pdf_memory()
    return code, captured.getvalue(), errors.getvalue()[-4000:]


def _release_pdf_memory():
    """MuPDF keeps decoded fonts and images in a store of up to 256 MB for the next page; in a
    process that lives for 300 jobs that is memory held for nothing between requests."""
    try:
        fitz = sys.modules.get("fitz") or sys.modules.get("pymupdf")
        if fitz is not None:
            fitz.TOOLS.store_shrink(100)
    except Exception:
        pass


def main():
    _warm()
    _reply.write(json.dumps({"ready": True}) + "\n")
    while True:
        line = sys.stdin.readline()
        if not line:
            return  # the API closed our stdin: shut down
        line = line.strip()
        if not line:
            continue
        rid = None
        try:
            req = json.loads(line)
            rid = req.get("id")
            script = req.get("script")
            if script not in ALLOWED:
                raise ValueError(f"script not allowed: {script!r}")
            t0 = time.time()
            code, out, err = _run(script, req.get("args") or [])
            reply = {"id": rid, "exit": code, "stdout": out, "stderr": err, "ms": int((time.time() - t0) * 1000)}
        except Exception as e:
            reply = {"id": rid, "exit": 1, "stdout": "", "error": str(e)}
        _reply.write(json.dumps(reply) + "\n")


if __name__ == "__main__":
    main()
