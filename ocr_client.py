# -*- coding: utf-8 -*-
"""Thin client for the central Confileo OCR microservice (PaddleOCR, :8090).

Every tool that needs OCR imports this instead of calling Tesseract. Pure stdlib
(urllib) so it works under the .NET service's system python with no extra deps.
"""
import os
import json
import urllib.request

OCR_URL = os.environ.get("OCR_SERVICE_URL", "http://127.0.0.1:8090").rstrip("/")


def ocr_pdf(path, lang="ar", dpi=None, max_pages=0, timeout=900):
    """OCR a whole PDF via the engine. Returns {success, page_count, pages:[...]}.
    Each page: {page, width, height, dpi, lines:[{text, box:[x0,y0,x1,y1], conf}]}."""
    if dpi is None:
        dpi = int(os.environ.get("OCR_DPI", "200"))
    body = json.dumps({"path": path, "lang": lang, "dpi": dpi, "max_pages": max_pages}).encode("utf-8")
    req = urllib.request.Request(OCR_URL + "/ocr/pdf", data=body,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def ocr_image(path, lang="ar", timeout=300):
    """OCR a single image file. Returns {success, lines:[...]}."""
    body = json.dumps({"path": path, "lang": lang}).encode("utf-8")
    req = urllib.request.Request(OCR_URL + "/ocr/image", data=body,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def is_available(timeout=4):
    try:
        with urllib.request.urlopen(OCR_URL + "/health", timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8")).get("status") == "ok"
    except Exception:
        return False
