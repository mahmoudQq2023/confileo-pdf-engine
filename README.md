# Confileo PDF engine

The Python scripts behind the PDF tools of [confileo.com](https://confileo.com).

They use [PyMuPDF](https://pymupdf.readthedocs.io/), which is licensed under the GNU Affero General Public
License v3. Because these scripts serve a public online service together with PyMuPDF, they are published here under
the same licence, as the AGPL requires. This repository follows the version that runs on confileo.com.

The rest of confileo.com (website, API, accounts, other tools) is not part of this repository.

## Tools and scripts

| Tool on confileo.com | Script |
|---|---|
| PDF to Word | `pdf_to_word.py`; Arabic documents: `pdf_to_word_arabic.py`, `pdf_to_word_arabic_v2.py` |
| PDF to Excel | `pdf_table_extractor.py` |
| PDF to PowerPoint | `pdf_to_pptx.py` |
| PDF to JPG / images, Word to image, handwriting pages | `pdf_to_images.py` |
| PDF to audio (text extraction) | `pdf_audio_text.py` |
| Compress PDF | `pdf_compress_rtl.py` |
| Merge PDF | `pdf_merge.py` |
| Split PDF | `pdf_split.py` |
| Organize PDF (rotate, delete, extract, reorder) | `pdf_organize.py` |
| Crop PDF | `pdf_crop.py` |
| Page numbers | `pdf_page_numbers.py` |
| Watermark | `pdf_watermark.py` |
| Sign PDF | `sign_pdf.py` |
| Protect and unlock PDF | `pdf_protect.py` |
| Redact PDF | `pdf_redact.py` |
| PDF editor | `pdf_text_editor.py` |
| PDF form filler | `pdf_form.py` |
| Flatten PDF | `pdf_flatten.py` |
| Images to PDF | `images_to_pdf.py` |
| OCR (searchable PDF text) | `ocr_pdf_text.py` |
| CV builder (PDF import and export) | `cv_builder_export.py` |
| Shared helpers | `pdf_geometry.py`, `pdf_text_rows.py`, `pdf_text_layer.py`, `ocr_client.py` |
| Warm Python process the API keeps running | `py_worker.py` |

## Running

- Python 3.9 or later: `pip install -r requirements.txt`
- System tools some scripts call: Tesseract (OCR) and Ghostscript (compression).
- `ocr_client.py` talks to an OCR service on `127.0.0.1:8090`, which is not part of this repository.
- Each script is run as `python3 <script>.py <arguments>`; its arguments are described at the top of the file.

## Licence

Copyright (C) 2026 Confileo (https://confileo.com)

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General
Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any
later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU Affero General Public License in
[LICENSE](LICENSE) for more details.

The libraries in `requirements.txt` are installed from PyPI under their own licences (PyMuPDF: AGPL-3.0;
pdf2docx, python-docx, fonttools: MIT; WeasyPrint: BSD; Pillow: MIT-CMU; pytesseract, uharfbuzz: Apache-2.0).

## بالعربي

دي سكريبتات البايثون اللي بتشغّل أدوات الـ PDF في موقع كونفيليو. السكريبتات دي بتستخدم مكتبة PyMuPDF، ورخصتها AGPL،
فبننشرها هنا بنفس الرخصة زي ما الرخصة بتطلب. باقي الموقع مش جزء من الريبو ده.
