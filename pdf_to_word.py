#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF → Word (DOCX) — world-class converter for Confileo.

Stack:
  • pdf2docx  — primary engine: real layout reconstruction (paragraphs, tables,
                images, columns) producing genuinely editable Word. It extracts
                text in LOGICAL Unicode order via PyMuPDF.
  • Arabic    — post-process the generated .docx to set RTL direction on Arabic
                paragraphs/runs (w:bidi / w:rtl + right alignment). We do NOT
                reshape or bidi-reorder the text: Word shapes logical Unicode
                itself, and reshaping would produce broken/disconnected glyphs.
  • OCR       — scanned / image-only PDFs (no text layer) fall back to Tesseract
                (ara+eng) so they still produce text instead of empty pages.

CLI (matches the .NET PdfToWordService contract):
    python pdf_to_word.py <input.pdf> <output.docx> [--json] [--no-images]

On --json, prints a single JSON line: {"success", "engine", "rtl", "message", "code"}.
Exit code 0 on success, 1 on failure (so .NET can fall back to LibreOffice), 3 when
the PDF needs an open password (no engine can read it).
"""
import sys
import os
import io
import re
import json
import argparse
import functools
import time
import unicodedata

# Control chars that are invalid in OOXML (python-docx rejects them). Keep
# tab/newline/CR; strip the rest. PDF text extraction often carries stray ones.
_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _xml_clean(s: str) -> str:
    return _CTRL_RE.sub("", s or "")

# Arabic Unicode blocks (base letters + supplements; NOT presentation forms).
_AR_RANGES = (
    (0x0600, 0x06FF),  # Arabic
    (0x0750, 0x077F),  # Arabic Supplement
    (0x08A0, 0x08FF),  # Arabic Extended-A
    (0xFB50, 0xFDFF),  # Arabic Presentation Forms-A
    (0xFE70, 0xFEFF),  # Arabic Presentation Forms-B
)


def _is_arabic_char(ch: str) -> bool:
    o = ord(ch)
    if o < 0x0600:          # Latin, digits, punctuation: most calls end here
        return False
    return any(lo <= o <= hi for lo, hi in _AR_RANGES)


def is_arabic_text(text: str) -> bool:
    """True if the text has a meaningful share of Arabic letters."""
    if not text:
        return False
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return False
    ar = sum(1 for c in letters if _is_arabic_char(c))
    return ar / len(letters) >= 0.30


def _has_presentation_forms(text: str) -> bool:
    """Arabic Presentation Forms (A/B) — a sign the PDF stored *shaped, visually
    ordered* glyphs instead of logical Unicode."""
    return any(0xFB50 <= ord(c) <= 0xFDFF or 0xFE70 <= ord(c) <= 0xFEFF for c in text)


# ── Presentation-form normalisation (NEVER reorder) ─────────────────────────
# THE central Arabic correctness problem. A PDF can store Arabic two ways:
#   • base Unicode letters (modern, well-built PDFs), or
#   • pre-shaped Arabic Presentation Forms (U+FB50–FDFF / FE70–FEFF), emitted by
#     older/CAD/print-to-PDF exporters.
# In BOTH cases PyMuPDF/pdf2docx return the text in LOGICAL reading order — this
# was verified end-to-end (Scripts/_rtltest) against a LibreOffice-saved base-
# letter PDF AND a reportlab presentation-form PDF: both extracted logically.
# Therefore we must NEVER reverse: reversing logical text mirrors it into
# gibberish (the bug this replaced). For presentation-form text we only map the
# shaped glyphs back to base letters via NFKC so Word can re-shape them with a
# connected Arabic face; base-letter text is left verbatim.

def _normalize_presentation_forms(text: str) -> str:
    """Map Arabic Presentation Forms and LAM-ALEF ligatures back to base Unicode
    letters via NFKC, WITHOUT reordering, so Word shapes them correctly."""
    text = unicodedata.normalize("NFKC", text)
    if "اأ" in text or "اإ" in text or "اآ" in text:
        text = _ALEF_LAM_HAMZA_RE.sub(r"\1ال\2", text)
    return text


# Microsoft Office's "Save as PDF" (and Chrome under PyMuPDF 1.26 for some
# faces) names the lam-alef-hamza ligature alef first: «الأساسي» extracts as
# «األساسي», «والإحصاء» as «واإلحصاء» — 60 words of the 15-page Access 2010
# exam table the backend keeps as test.pdf, none of them right (10 Oct 2026).
# An alef followed by a hamza on alef never occurs in Arabic spelling, so at a
# word's start, or after the prefixes و ف ب ك that can stand before «ال», the
# letters go back in order. (Office's yeh glyphs: see _office_yeh.)
_ALEF_LAM_HAMZA_RE = re.compile("(?<![\u0621-\u064A])([\u0648\u0641]?[\u0628\u0643]?)"
                                "ا([أإآ])ل")


# ── DOCX helpers (python-docx low-level OOXML) ──────────────────────────────
# The children of w:pPr, w:rPr, w:tblPr and w:sectPr form an xsd:sequence, and
# Word checks the order: w:bidi appended after w:jc, w:rFonts after w:sz,
# w:bidiVisual after w:tblLook and w:bidi after w:docGrid (all four were written
# that way until 28 Sept 2026) make a schema-invalid file that LibreOffice
# forgives and Word may refuse or offer to "repair". Everything this script
# writes by hand goes through _put(), which inserts at the schema position.
_PPR_ORDER = ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
              "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
              "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
              "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
              "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
              "sectPr", "pPrChange")
_RPR_ORDER = ("rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
              "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
              "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
              "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")
_TBLPR_ORDER = ("tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
                "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
                "tblLook", "tblCaption", "tblDescription")
_TCPR_ORDER = ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
               "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge",
               "tcPrChange")
_SECTPR_ORDER = ("headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
                 "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
                 "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange")

# The complex-script face for Arabic runs when the PDF does not name a usable one.
_AR_FALLBACK_FONT = "Arial"


def _put(parent, tag, order, **attrs):
    """The child `tag` ("w:bidi") of `parent`, created at its schema position if
    missing; the given attributes (without prefix) are set on it."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        rank = order.index(tag.split(":", 1)[1])
        for child in parent:
            name = child.tag.rsplit("}", 1)[-1] if isinstance(child.tag, str) else ""
            if name in order and order.index(name) > rank:
                child.addprevious(el)
                break
        else:
            parent.append(el)
    for k, v in attrs.items():
        el.set(qn("w:" + k), v)
    return el


def _cs_mirror(rPr):
    """Arabic is sized, emboldened and italicised from the COMPLEX-SCRIPT properties:
    a run with w:sz 28 and no w:szCs printed its Arabic at the template's 11 pt, and
    w:b without w:bCs left every Arabic heading regular, in Word and LibreOffice
    alike (measured 28 Sept 2026 on Chrome and LibreOffice PDFs). The Latin value
    is copied into the complex-script slot wherever that slot is still empty."""
    from docx.oxml.ns import qn
    sz = rPr.find(qn("w:sz"))
    if sz is not None and rPr.find(qn("w:szCs")) is None:
        _put(rPr, "w:szCs", _RPR_ORDER, val=sz.get(qn("w:val")))
    for latin, cs in (("w:b", "w:bCs"), ("w:i", "w:iCs")):
        el = rPr.find(qn(latin))
        if el is not None and rPr.find(qn(cs)) is None:
            val = el.get(qn("w:val"))
            if val is None:
                _put(rPr, cs, _RPR_ORDER)
            else:
                _put(rPr, cs, _RPR_ORDER, val=val)


def _style_run(run, size=None, bold=False, italic=False, color=None, font=None, cs_font=None, rtl=False):
    """Formats a run in both script slots at once (w:sz + w:szCs, w:b + w:bCs,
    w:i + w:iCs, rFonts ascii/hAnsi + cs), every element at its schema position."""
    from docx.oxml.ns import qn
    rPr = run._r.get_or_add_rPr()
    if font or cs_font:
        f = _put(rPr, "w:rFonts", _RPR_ORDER)
        if font:
            f.set(qn("w:ascii"), font)
            f.set(qn("w:hAnsi"), font)
        f.set(qn("w:cs"), cs_font or font)
    if bold:
        _put(rPr, "w:b", _RPR_ORDER)
        _put(rPr, "w:bCs", _RPR_ORDER)
    if italic:
        _put(rPr, "w:i", _RPR_ORDER)
        _put(rPr, "w:iCs", _RPR_ORDER)
    if color:
        _put(rPr, "w:color", _RPR_ORDER, val=color)
    if size:
        v = str(int(round(size * 2)))
        _put(rPr, "w:sz", _RPR_ORDER, val=v)
        _put(rPr, "w:szCs", _RPR_ORDER, val=v)
    if rtl:
        _put(rPr, "w:rtl", _RPR_ORDER)
    return run


def _set_paragraph_rtl(paragraph):
    from docx.oxml.ns import qn
    pPr = paragraph._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is not None:
        return      # laid out right-to-left already, its alignment chosen with it
    _put(pPr, "w:bidi", _PPR_ORDER)
    jc = pPr.find(qn("w:jc"))
    if jc is None or jc.get(qn("w:val")) not in ("center", "both", "distribute"):
        # Inside a bidi paragraph Word and LibreOffice read w:jc left/right as
        # START/END: "right" (written here until 28 Sept 2026) put every Arabic
        # paragraph on the LEFT margin. "left" is the start, the right margin.
        _put(pPr, "w:jc", _PPR_ORDER, val="left")


def _set_run_rtl(run):
    from docx.oxml.ns import qn
    rPr = run._r.get_or_add_rPr()
    # Bind Arabic to the complex-script font slot so Word shapes it with an
    # Arabic-capable face instead of falling back to tofu boxes.
    rFonts = _put(rPr, "w:rFonts", _RPR_ORDER)
    if not rFonts.get(qn("w:cs")):
        rFonts.set(qn("w:cs"), _AR_FALLBACK_FONT)
    _put(rPr, "w:rtl", _RPR_ORDER)
    _cs_mirror(rPr)


def _set_section_rtl(doc):
    for section in doc.sections:
        _put(section._sectPr, "w:bidi", _SECTPR_ORDER)


def _rebuild_paragraph_basenorm(paragraph):
    """Rebuild a presentation-form Arabic paragraph as a single run whose text is
    NFKC-normalised to base letters (no reordering), preserving the first run's
    basic font so Word can shape it correctly."""
    new_text = _normalize_presentation_forms(paragraph.text)
    runs = paragraph.runs
    bold = italic = size = name = None
    if runs:
        f = runs[0].font
        bold, italic, size, name = f.bold, f.italic, f.size, f.name
    for r in list(runs):
        r._element.getparent().remove(r._element)
    run = paragraph.add_run(new_text)
    run.font.bold, run.font.italic = bold, italic
    if size:
        run.font.size = size
    if name:
        run.font.name = name
    return run


def _process_paragraphs(paragraphs):
    found = False
    for p in paragraphs:
        if not p.text or not is_arabic_text(p.text):
            continue
        found = True
        # NEVER reorder: PyMuPDF/pdf2docx hand back Arabic in logical order, so
        # reversing would mirror it into gibberish. If the text carries pre-shaped
        # presentation forms, map them to base letters (no reordering) so Word can
        # re-shape; otherwise keep it verbatim. Either way, set RTL.
        if _has_presentation_forms(p.text):
            run = _rebuild_paragraph_basenorm(p)
            _set_run_rtl(run)
        else:
            for run in p.runs:
                _set_run_rtl(run)
        _set_paragraph_rtl(p)
    return found


def postprocess_rtl(docx_path: str, title=None) -> bool:
    """Set RTL on every Arabic paragraph/run (body + tables). Returns True if
    any Arabic was found. Text is never reordered — Word does the shaping; only
    pre-shaped presentation forms are normalised to base letters. Every engine's
    output passes here, so the document's properties are set here too."""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import _Cell
    doc = Document(docx_path)
    any_ar = _process_paragraphs(doc.paragraphs)
    # Every cell once, from the XML (row.cells rebuilds the table grid per row,
    # quadratic on a long table). A merged cell used to be visited once per grid
    # slot it spans; its paragraphs are the same objects, so once is enough.
    for table in doc.tables:
        for tc in table._tbl.iter(qn("w:tc")):
            if _process_paragraphs(_Cell(tc, table).paragraphs):
                any_ar = True
    # running heads and page footers the rebuild moved into the sections
    for section in doc.sections:
        for part in (section.header, section.footer):
            if not part.is_linked_to_previous and _process_paragraphs(part.paragraphs):
                any_ar = True
    if any_ar:
        _set_section_rtl(doc)
    _unnest_hyperlinks(doc.element)
    # pdf2docx writes w:widowControl after w:autoSpaceDE and similar; every
    # engine's output leaves here with its property children in schema order
    _normalize_order(doc.element)
    for section in doc.sections:
        for part in (section.header, section.footer):
            if not part.is_linked_to_previous:
                _normalize_order(part._element)
    _set_properties(doc, title)
    doc.save(docx_path)
    return any_ar


def _set_properties(doc, title):
    """File > Info in Word: every document said author "python-docx", created and
    modified 23 Dec 2013 and "generated by python-docx" (python-docx's template,
    10 Oct 2026). Now: the title the caller gives (the uploaded file's name), the
    site as author, and the time of the conversion."""
    import datetime
    cp = doc.core_properties
    now = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0, tzinfo=None)
    cp.title = _xml_clean(title or "")[:255]
    cp.author = cp.last_modified_by = "Confileo"
    cp.comments = ""
    cp.created = cp.modified = now
    cp.revision = 1


def _unnest_hyperlinks(root):
    """pdf2docx writes a web link as <w:r><w:rPr/><w:hyperlink><w:r>…</w:r>
    </w:hyperlink></w:r>: a hyperlink inside a run, which the schema forbids.
    LibreOffice drops the link and its words («…and a .» for «…and a web link.»)
    and Word reports unreadable content (2 Oct 2026, every Latin PDF with a
    table and a link goes through pdf2docx). The hyperlink moves up to the
    paragraph in the run's place and its runs take the outer run's look."""
    import copy
    from docx.oxml.ns import qn
    for h in list(root.iter(qn("w:hyperlink"))):
        r = h.getparent()
        if r is None or r.tag != qn("w:r"):
            continue
        outer = r.find(qn("w:rPr"))
        for inner in h.findall(qn("w:r")):
            if outer is None:
                break
            rpr = inner.find(qn("w:rPr"))
            if rpr is None:
                rpr = copy.deepcopy(outer)
                inner.insert(0, rpr)
                continue
            have = {c.tag for c in rpr}
            for c in outer:
                if c.tag not in have:
                    rpr.append(copy.deepcopy(c))
        # other content of the outer run (rare) stays in a run before the link
        rest = [c for c in r if c is not h and c.tag != qn("w:rPr")]
        if rest:
            r.addnext(h)              # lxml moves the element
        else:
            r.getparent().replace(r, h)


def _normalize_order(root):
    """Sorts the children of every w:pPr / w:rPr / w:tblPr / w:sectPr under `root`
    into schema order (stable; foreign or unknown children keep their order at
    the end, tracked-change markers first in a paragraph mark's w:rPr)."""
    from docx.oxml.ns import qn
    w_ns = qn("w:p").split("}")[0][1:]
    for tag, order in (("w:pPr", _PPR_ORDER), ("w:rPr", _RPR_ORDER), ("w:tblPr", _TBLPR_ORDER),
                       ("w:sectPr", _SECTPR_ORDER)):
        rank = {n: i for i, n in enumerate(order)}
        if tag == "w:rPr":
            rank.update({"ins": -4, "del": -3, "moveFrom": -2, "moveTo": -1})
        last = len(order) + 1
        for el in root.iter(qn(tag)):
            kids = list(el)
            keys = []
            for c in kids:
                if not isinstance(c.tag, str) or not c.tag.startswith("{" + w_ns + "}"):
                    keys.append(last)
                else:
                    keys.append(rank.get(c.tag.split("}", 1)[1], last))
            if keys == sorted(keys):
                continue
            for c in kids:
                el.remove(c)
            for _, _, c in sorted(zip(keys, range(len(kids)), kids), key=lambda t: (t[0], t[1])):
                el.append(c)


# ── Engines ─────────────────────────────────────────────────────────────────
def has_text_layer(pdf_path: str) -> bool:
    """True if the PDF has an extractable text layer (vs scanned images)."""
    import fitz
    doc = fitz.open(pdf_path)
    try:
        for page in doc:
            if page.get_text("text").strip():
                return True
        return False
    finally:
        doc.close()


# A text layer that is not the text printed on the page. Measured 2 Oct 2026:
# a WeasyPrint PDF set in Noto Sans Arabic carries a broken ToUnicode map and
# extracts as «ةه2ن هن\ الوشي4ةUة اإلh» for «تهدف هذه الوثيقة إلى» — 10% of the
# Arabic words in the .docx were real words, against 100% for the same page in
# Amiri. Arabic letters glued to Latin letters, backslashes or control codes
# are that signature (38-48% of the Arabic letters on such a page, 0 on the
# twelve sound Arabic pages of the test set, tables with digits included).
# Old Arabic DTP fonts with their own encoding come out as Latin-1 accented
# letters («ÇáÔÑßÉ»); a font with no Unicode map at all as U+FFFD or
# private-use code points. Such a page is read like a scan.
_GLUED_AR_RE = re.compile(r"[\u0600-\u06FF][A-Za-z\x00-\x08\x0b-\x1f\\]|[A-Za-z\x00-\x08\x0b-\x1f\\][\u0600-\u06FF]")


def _garbled_text(text):
    """True when a page's extracted text is mojibake rather than its words."""
    ar = sum(1 for c in text if "\u0600" <= c <= "\u06FF")
    if ar >= 30 and len(_GLUED_AR_RE.findall(text)) >= 0.08 * ar:
        return True
    letters = [c for c in text if c.isalpha()]
    if len(letters) >= 50 and sum(1 for c in letters if "\u00c0" <= c <= "\u00ff") >= 0.4 * len(letters):
        return True
    shown = [c for c in text if not c.isspace()]
    unmapped = sum(1 for c in shown if c == "\ufffd" or "\ue000" <= c <= "\uf8ff")
    return len(shown) >= 50 and unmapped >= 0.3 * len(shown)


def _text_pages(pdf_path):
    """(pages with a text layer, the ones among them whose text is garbled)."""
    import fitz
    doc = fitz.open(pdf_path)
    try:
        has, bad = set(), set()
        for i, page in enumerate(doc):
            t = page.get_text("text")
            if t.strip():
                has.add(i)
                if _garbled_text(t):
                    bad.add(i)
        return has, bad
    finally:
        doc.close()


def _ensure_pdf2docx_compat():
    """pdf2docx 0.5.8 calls fitz.Rect.get_area(), which newer PyMuPDF (1.26.5+)
    removed. Without this shim pdf2docx crashes with
    "'Rect' object has no attribute 'get_area'", the .NET service falls back to
    LibreOffice, and Arabic comes out visually REVERSED.

    We deliberately keep a *recent* PyMuPDF: it extracts Arabic in correct LOGICAL
    order, whereas PyMuPDF < 1.24 returns raw VISUAL order (reversed). So the right
    combination is: new PyMuPDF (logical text) + this shim (no crash)."""
    import fitz
    _area = lambda self, unit="px": self.width * self.height
    for cls in ("Rect", "IRect"):
        rc = getattr(fitz, cls, None)
        if rc is not None and not hasattr(rc, "get_area"):
            rc.get_area = _area


def convert_with_pdf2docx(pdf_path: str, docx_path: str):
    _ensure_pdf2docx_compat()
    from pdf2docx import Converter
    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path, start=0, end=None)
    finally:
        cv.close()


# ── Position-aware rebuild (replaces the "words"-order rebuild) ─────────────
# The text order a PDF *stores* is not the order a reader reads. Measured on
# 2026-09-07 against three producers of the same Arabic CV:
#   • WeasyPrint/Cairo (our own CV builder): words come back in VISUAL order
#     ("الله عبد سمير أحمد" for "أحمد سمير عبد الله") and a glyph cluster such
#     as "شر" is handed over as a zero-width "ر" *before* the "ش" it follows —
#     so "مشروعات" read "مرشوعات" and "بخبرة" read "بخربة".
#   • LibreOffice / Word: words already in LOGICAL order, and the only zero-width
#     characters are the second half of a lam-alef ligature, already in place.
#   • PdfPig (the engine this replaced for Arabic): logical order but wrong
#     letters on subsetted fonts ("معاومات" for "معلومات", "ميندس" for "مهندس").
# Reversing or not reversing the string therefore fixes one producer and breaks
# the other. What never lies is the glyph's POSITION on the page: a word that is
# further right is read first in an RTL line. So this rebuild orders words by
# their x coordinate, keeps each word's own letters in the order PyMuPDF gives
# (logical, once a stray cluster tail is moved back behind its base), and
# un-reverses embedded Latin/digit runs so "Microsoft Office 2019" and
# "2019 – 2022" keep their own direction inside an Arabic line.

_ZW_EPS = 0.5           # pt; a zero-width char sits on a neighbour's edge (italics drift a little)
_BULLETS = "•·▪■●○◦‣-–—*✓✔➢►▶"
# Word and LibreOffice PDFs carry list bullets as Symbol/Wingdings private-use
# code points (U+F0B7 is Symbol's bullet); in any other font they print as a box.
_PUA_BULLETS = {"\uf0b7": "•", "\uf0a7": "▪", "\uf0fc": "✓", "\uf0d8": "➢", "\uf076": "❖",
                "\uf0a8": "◦", "\uf06e": "■", "\uf0e0": "➢", "\uf0d7": "•", "\uf02d": "–"}
_LIST_MARKS = set("•·▪■●○◦‣✓✔➢►▶❖") | set(_PUA_BULLETS)
_ALEFS = "اأإآ"
_OPENERS = "([{«"
_CLOSERS = ")]}»"
_BRACKET_PAIR = {"(": ")", ")": "(", "[": "]", "]": "[", "{": "}", "}": "{", "«": "»", "»": "«"}
_END_PUNCT = ".,،:;!?؟"
_BIDI_CONTROLS = "‎‏‪‫‬‭‮⁦⁧⁨⁩"
_SPACE_BEFORE_RE = re.compile(r"\s+([،,\.:;!?؟\)\]\}»%])")
_SPACE_AFTER_RE = re.compile(r"([\(\[\{«])\s+")
_SLASH_RE = re.compile(r"(\d)\s*/\s*(\d)")


def _width(c):
    return c[2] - c[1]


def _carrier_mark(chars, i):
    """Index of the combining mark WITH width whose glyph draws the zero-width
    letter at i, or -1. Walks back over zero-width marks stacked on that glyph."""
    z = chars[i]
    k = i - 1
    while k >= 0 and unicodedata.combining(chars[k][0]) and _width(chars[k]) < 0.01:
        k -= 1
    if k < 0 or not unicodedata.combining(chars[k][0]):
        return -1
    m = chars[k]
    if _width(m) >= 0.01 and m[1] - _ZW_EPS <= z[1] <= m[2] + _ZW_EPS:
        return k
    return -1


def _glyph_before(chars, i, spill=False):
    """(index, left edge) of the glyph with a width that the zero-width char at
    i hangs on: back over zero-width letters and marks, and over a space MuPDF
    made up (synthetic) inside that glyph's own box; (-1, None) when a real
    space, or a made-up one outside the glyph, comes first. With spill=True a
    made-up space may stick out of the glyph's left side by up to half an em,
    and the left edge is then that space's: Chrome's DejaVu Sans parks the lam
    of «لاً» at its far end («جزيلاً» came out «جزيًا», 10 Oct 2026)."""
    k = i - 1
    synth = []
    while k >= 0:
        c = chars[k]
        if c[0] in "  ":
            if not _is_synthetic(c):
                return -1, None
            synth.append(c)
        elif _width(c) >= 0.01:
            break
        k -= 1
    if k < 0:
        return -1, None
    g = chars[k]
    reach = g[1] - (0.6 * max(g[4] or 1.0, 1.0) if spill else 0.0) - _ZW_EPS
    if any(s[1] < reach or s[2] > g[2] + _ZW_EPS for s in synth):
        return -1, None
    return k, min([g[1]] + [s[1] for s in synth])


def _once(marks):
    """Zero-width marks of one cluster, each kept once and without the spaces
    MuPDF made up inside the glyph (Chrome's «أولاً» at a line end came with
    two fathatan)."""
    out, seen = [], set()
    for c in marks:
        if _is_synthetic(c):
            continue
        if unicodedata.combining(c[0]) and _width(c) < 0.01:
            if c[0] in seen:
                continue
            seen.add(c[0])
        out.append(c)
    return out


def _lam_host(chars, i, j):
    """For an alef with no width at i that nothing in its stretch of the stream
    claims (chars[i:j] are it and its marks): the index of the lam whose left
    edge it sits on, on the same line, or -1. Chrome writes the line's last
    glyph cluster first in the line's text; for «أولاً.» at the end of a line
    that was ". ا ً ً" before the first word, and the alef and its tanween were
    dropped as a cluster without a home («أول.»)."""
    z = chars[i]
    hosts = []
    for k, c in enumerate(chars):
        if i <= k < j or c[0] != "ل" or _width(c) < 0.01 or abs(c[1] - z[1]) >= _ZW_EPS:
            continue
        if not _same_line(c, z):
            continue
        n = k + 1
        while n < len(chars) and _width(chars[n]) < 0.01 and chars[n][0] not in "  " \
                and chars[n][0] not in _ALEFS:
            n += 1
        if n < len(chars) and chars[n][0] in _ALEFS and _width(chars[n]) < 0.01:
            continue                    # this lam carries its own alef already
        hosts.append(k)
    return hosts[0] if len(hosts) == 1 else -1      # two candidates: no guess


def _fix_zero_width_clusters(chars):
    """chars: list of [c, x0, x1, span, size]. A zero-width char is the extra
    half of a glyph cluster: one glyph drawn for two or more characters, with the
    width on one of them and the rest parked at zero width on an edge of it.

    Measured 22 Sept 2026 on the same Arabic and English text printed by Chrome,
    Microsoft Word and LibreOffice (_wf3/audit11/rows_regress.py):
      • All three emit the extra letter BEFORE the glyph that draws it, parked on
        that glyph's right edge ("ني" arrives ي(0) ن, "لأ" arrives أ(0) ل), so it
        moves behind the following glyph. Cairo/WeasyPrint does the same.
      • When the glyph is a letter carrying a shadda or tanween, Chrome and
        LibreOffice give the WIDTH to the mark and park the letter at zero width
        inside the mark's box: "دّ" arrives ّ د(0), "اً" arrives ً ا(0). Read with
        the rule above, the letter was moved into the NEXT letter ("جرّب" read
        "جبر", "بيدّيك" read "بييدك") or across the space ("مجاناً تماماً" read
        "مجانً اتمامً ا") — about one Arabic word in nine on a Chrome PDF.
      • Latin ligatures (fi, fl, ff, ft, tt, ti) put the second letter at zero
        width AFTER the first, already in order. They were deleted as orphans:
        "file" came out "fle" and "automatically" "automatcally" for every
        LibreOffice and Chrome PDF — in Word output, narration and CV import.
    """
    i = 0
    while i < len(chars):
        z = chars[i]
        w = z[2] - z[1]
        size = z[4] if len(z) > 4 and z[4] else 1.0
        zero = w < 0.01
        if zero and z[0] not in "  " and not unicodedata.combining(z[0]):
            if not _is_arabic_char(z[0]):
                # an LTR ligature's tail is already where it belongs
                i += 1
                continue
            prev_c = chars[i - 1] if i > 0 else None
            if (prev_c is not None and prev_c[0] == z[0] and 0.0 < _width(prev_c) < 0.1 * size
                    and abs(prev_c[1] - z[1]) < 1.0):
                # Chrome can draw one letter as two pieces, a sliver and a
                # zero-width copy, both mapped to the same letter: "بخبرة" came
                # out "بخبخرة" in a CV import. The copy is not a second letter.
                del chars[i]
                continue
            k = _carrier_mark(chars, i)
            if k >= 0:
                # the letter first, the marks it carries after it
                chars[k:i + 1] = [z] + chars[k:i]
                i += 1
                continue
            g, left = _glyph_before(chars, i, spill=True)
            a = chars[g] if g >= 0 else None
            if z[0] == "ل" and a is not None and a[0] in _ALEFS \
                    and not any(_is_letter_zw(c) for c in chars[g + 1:i]):
                # "لاً" whose ToUnicode came back reversed: ً(0) ا ل(0), the
                # lam parked on the alef's LEFT edge (LibreOffice and Chrome both
                # do this for a lam-alef that carries a tanween). Lam first, then
                # the alef, then the marks that sat on the glyph's right edge.
                s = g
                while (s - 1 >= 0 and unicodedata.combining(chars[s - 1][0])
                       and _width(chars[s - 1]) < 0.01 and abs(chars[s - 1][1] - a[2]) < _ZW_EPS):
                    s -= 1
                n = i + 1
                while n < len(chars) and _width(chars[n]) < 0.01 and chars[n][0] not in "  ":
                    n += 1
                word_ends = n >= len(chars) or chars[n][0] in "  " or not _is_arabic_char(chars[n][0])
                # Only with the reversed cluster's signature — a mark on the
                # glyph's right edge, or the word ending here. Otherwise the lam
                # is the tail of the next glyph and the rule below handles it.
                on_left = abs(left - z[1]) < _ZW_EPS and (s < g or word_ends)
                # Chrome also gives the ligature's whole width to the tanween
                # and parks the alef on its right edge and the lam INSIDE it:
                # «جميلًا» arrives ا(0) ً ل(0), and once the alef took the mark's
                # box the lam inside it was dropped as a stray tail — «جميًا»,
                # «مثًا», «جزيًا» in every Chrome PDF set in IBM Plex, Noto
                # Kufi or Noto Sans Arabic Bold (58 words in 8 test PDFs, 10 Oct
                # 2026, PyMuPDF 1.26.5). A lam well inside the alef's own box is
                # never the next glyph's tail: the two are one lam-alef.
                inside = (a[1] + _ZW_EPS < z[1] < a[2] - _ZW_EPS
                          and not any(_is_letter_zw(c) for c in chars[i + 1:n]))
                if on_left or inside:
                    cluster = [z, a] + _once(chars[s:g] + chars[g + 1:i])
                    chars[s:i + 1] = cluster
                    i = s + len(cluster)
                    continue
        if i == len(chars) - 1:
            break
        # Cairo's Regular faces give the cluster tail a SMALL width instead of
        # zero (ر = 0.37 em where a real ر is 0.69 em) and park it INSIDE the
        # box of the base glyph that follows it in the stream: "شركة" arrived as
        # ر ش ك ة and "للمقاولات" as ل ل م ق ا و ا ل ت (7 Sept 2026, IBM Plex
        # Sans Arabic Regular 8.8 pt). Same cluster, same cure: move it behind
        # its base. Only Arabic letters, only when the next glyph contains it.
        # Width window matters: Cairo tails are 0.35-0.4 em; Amiri's genuinely
        # narrow medial forms (م 0.15 em, خ 0.05 em) are in the right order
        # already and must not be touched.
        narrow = (not zero and 0.2 * size <= w < 0.45 * size and _is_arabic_char(z[0])
                  and not unicodedata.combining(z[0]))
        if (zero or narrow) and z[0] not in "  " and not unicodedata.combining(z[0]):
            prev_c = chars[i - 1] if i > 0 else None
            prev = prev_c[0] if prev_c is not None else ""
            j = i + 1
            if zero:
                while j < len(chars) and (chars[j][2] - chars[j][1]) < 0.01:
                    j += 1
                nxt = chars[j] if j < len(chars) and chars[j][0] not in "  " else None
                placed = nxt is not None and abs(nxt[2] - z[1]) < _ZW_EPS
                on_prev = (prev_c is not None and prev not in "  " and _width(prev_c) >= 0.01
                           and abs(prev_c[1] - z[1]) < _ZW_EPS)
                if not on_prev and not placed:
                    # The editor writes a ligature's further letters as zero-width
                    # glyphs just inside its left edge, after it («الله» in Naskh
                    # is one glyph and two of them): in order too. So does Chrome
                    # for the alef of «لًا», with a space MuPDF made up inside the
                    # lam's box between them: «جميلًا» came out «جميلً» in Noto
                    # Naskh (10 Oct 2026); that space no longer hides the lam.
                    k, _ = _glyph_before(chars, i)
                    own = chars[k] if k >= 0 else None
                    on_prev = (own is not None and own[1] - _ZW_EPS <= z[1] <= own[1] + 0.3 * _width(own))
                if z[0] in _ALEFS and prev == "ل":
                    # A lam-alef emitted lam first (base-first) is already right —
                    # unless the alef is the tail of the NEXT lam: in "للأصل" it
                    # touches both lams, and the ligature is the wider glyph
                    # (the initial lam is ~0.2 em, the lam-alef ~0.55 em).
                    if not (placed and nxt[0] == "ل" and (not on_prev or _width(nxt) > _width(prev_c))):
                        i += 1
                        continue
                elif not placed and on_prev:
                    # a tail emitted after its glyph, on its left edge: in order
                    i += 1
                    continue
            else:
                if z[0] in _ALEFS and prev == "ل":
                    i += 1
                    continue
                # the tail lies INSIDE the base's box (it shares one of its edges);
                # an adjacent glyph that merely touches the box is left alone.
                eps = 0.05 * size
                placed = (j < len(chars) and (chars[j][2] - chars[j][1]) >= 0.45 * size
                          and chars[j][1] - eps <= z[1] and z[2] <= chars[j][2] + eps)
            if placed:
                base = chars[j]
                run = chars[i:j]
                if narrow:
                    # park the tail on the base's left edge with no width, like a
                    # real zero-width tail, so the tokenizer's gap test measures
                    # from the base and does not split "شركة" into "شر كة".
                    z[1] = z[2] = base[1]
                # A run of SEVERAL zero-width letters is one ligature glyph whose
                # ToUnicode string was written in visual order: LibreOffice hands
                # "الله" back as ا + [ه ل] + ل, and keeping the run as-is spelt it
                # "الهل" in every PDF→Word, CV import and Excel export (7 Sept
                # 2026). Reverse the letters of the run; a combining mark stays
                # glued behind its own letter. A one-letter run (Cairo's tails)
                # is unchanged.
                groups = []
                for c in run:
                    if groups and unicodedata.combining(c[0]):
                        groups[-1].append(c)
                    else:
                        groups.append([c])
                run = [c for g in reversed(groups) for c in g]
                chars[i:j + 1] = [base] + run
                i = j + 1
                continue
            if narrow:
                i += 1
                continue
            if z[0] in _ALEFS and all(unicodedata.combining(c[0]) for c in chars[i + 1:j]):
                h = _lam_host(chars, i, j)
                if h >= 0:
                    # the lam-alef's alef written ahead of its line (_lam_host):
                    # back behind its lam, the lam's own marks staying with it
                    group = _once(chars[i:j])
                    del chars[i:j]
                    if h > i:
                        h -= j - i
                    q = h + 1
                    while (q < len(chars) and _width(chars[q]) < 0.01 and chars[q][0] not in "  "
                           and chars[h][1] - _ZW_EPS <= chars[q][1] <= chars[h][2] + _ZW_EPS):
                        q += 1
                    chars[q:q] = group
                    if q <= i:
                        i += len(group)
                    continue
            # An orphan: MuPDF parks a cluster it could not place (typically the
            # "اً" ending a word at the start of a Cairo text object) at the end
            # of the previous object, which can be another column entirely.
            # There is no way back to its word, and leaving it would print the
            # mark inside a date, so it is dropped.
            del chars[i:j]
            continue
        i += 1


def _token_kind(text):
    if any(_is_arabic_char(c) for c in text):
        return "R"
    if any(c.isalpha() for c in text):
        return "L"
    if any(c.isdigit() for c in text):
        return "D"
    return "N"


def _order_tokens(tokens, rtl):
    """tokens: list of dicts with text/x0/x1. Returns them in reading order."""
    if not rtl:
        return _order_ltr_tokens(tokens)
    toks = sorted(tokens, key=lambda t: -t["x1"])
    kinds = [_token_kind(t["text"]) for t in toks]
    # A Latin/digit run keeps its own left-to-right order. A neutral token joins
    # the run only between two WORDS: between two numbers ("2019 – 2022") the
    # Unicode bidi algorithm resolves it to the paragraph direction (rule N1),
    # so the numbers already sit in reading order and must not be flipped.
    out = []
    i = 0
    while i < len(toks):
        if kinds[i] not in "LD":
            out.append(toks[i])
            i += 1
            continue
        j = i + 1
        while j < len(toks):
            if kinds[j] in "LD":
                j += 1
            elif kinds[j] == "N" and kinds[j - 1] == "L" and j + 1 < len(toks) and kinds[j + 1] == "L":
                j += 1
            else:
                break
        out.extend(reversed(toks[i:j]))
        i = j
    return out


def _order_ltr_tokens(tokens):
    """A left-to-right row in reading order. Words follow x, except inside an
    Arabic phrase: the bidi algorithm lays an RTL run out right to left, so its
    words stand in x order LAST word first. "Seller: … Co. — شركة المستقبل
    للحلول الرقمية" came back "… Co. — الرقمية للحلول المستقبل شركة" in every
    bilingual invoice (28 Sept 2026). A run is Arabic words plus what the
    algorithm attaches to them: neutrals and numbers between two Arabic words
    (rule N1; a number after Arabic letters is an Arabic number, rule W2) and
    the numbers right after the last one ("الرياض 12345")."""
    toks = sorted(tokens, key=lambda t: t["x0"])
    kinds = [_token_kind(t["text"]) for t in toks]
    n = len(toks)
    out = []
    i = 0
    while i < n:
        if kinds[i] != "R":
            out.append(toks[i])
            i += 1
            continue
        j = i + 1
        while j < n:
            if kinds[j] == "R":
                j += 1
                continue
            k = j
            while k < n and kinds[k] in "ND":
                k += 1
            if k < n and kinds[k] == "R":
                j = k + 1
                continue
            last_digit = None
            for m in range(j, k):
                if kinds[m] == "D":
                    last_digit = m
            if last_digit is not None:
                j = last_digit + 1
            break
        run = list(reversed(toks[i:j]))
        if any(ch in _BRACKET_PAIR for t in run for ch in t["text"]):
            _normalize_brackets(run)
        out.extend(run)
        i = j
    return out


def _fix_digit_runs(tok):
    """PyMuPDF treats Arabic-Indic digits as RTL letters and hands "٢٠٠٩" back as
    "٩٠٠٢". Digits read left-to-right in every script, so every run of digits
    (with the separators that live between digits) is put in ascending x."""
    chars = tok["chars"]
    if not any(_is_arabic_char(c[0]) for c in chars):
        return
    i = 0
    out = []
    while i < len(chars):
        if unicodedata.category(chars[i][0]) == "Nd":
            j = i
            while j < len(chars) and (unicodedata.category(chars[j][0]) == "Nd"
                                      or (chars[j][0] in ".,:/٫٬" and j + 1 < len(chars)
                                          and unicodedata.category(chars[j + 1][0]) == "Nd")):
                j += 1
            run = sorted(chars[i:j], key=lambda c: (c[1] + c[2]) / 2.0)
            out.extend(run)
            i = j
        else:
            out.append(chars[i])
            i += 1
    tok["chars"] = out


def _settle_token_punctuation(tok):
    """Inside an RTL word the letters are already logical, but a comma or bracket
    glued to the word can sit on either side in the stream. Put punctuation that
    is drawn LEFT of every letter at the end, and RIGHT of every letter first."""
    chars = tok["chars"]
    if not any(_is_arabic_char(c[0]) for c in chars):
        return
    letters = [c for c in chars if c[0].isalnum() or unicodedata.combining(c[0])]
    if not letters or len(letters) == len(chars):
        return
    lx0 = min(c[1] for c in letters)
    lx1 = max(c[2] for c in letters)
    lead, body, trail = [], [], []
    for c in chars:
        if c in letters:
            body.append(c)
        elif (c[1] + c[2]) / 2.0 > lx1 - 0.01:
            lead.append(c)
        elif (c[1] + c[2]) / 2.0 < lx0 + 0.01:
            trail.append(c)
        else:
            body.append(c)
    tok["chars"] = lead + body + trail


def _is_letter_zw(c):
    """A zero-width Arabic LETTER (not a mark, not a space)."""
    return (_width(c) < 0.01 and c[0] not in "  " and _is_arabic_char(c[0])
            and not unicodedata.combining(c[0]))


def _is_synthetic(c):
    return c[0] == " " and isinstance(c[3], dict) and bool(c[3].get("_synthetic"))


def _is_zw_mark(c):
    return unicodedata.combining(c[0]) and _width(c) < 0.01


def _is_body(c):
    return _width(c) >= 0.01 and _is_arabic_char(c[0]) and not unicodedata.combining(c[0])


def _drop_marked_letter_copies(chars):
    """LibreOffice writes a letter that carries a shadda or tanween as the mark
    at zero width on the body's RIGHT edge, the body, then the mark again and a
    zero-width copy of the letter on its LEFT edge: "تقدّم" came out "تقّدّمد",
    "المنفّذة" "المنّفّذفة", "المعدّات" "المعّدّادت" in every LibreOffice 24.2
    PDF, whatever the font (28 Sept 2026). The copy of the letter and the second
    mark go, and a zero-width mark parked on the right edge of the body that
    follows it — its visual position, ahead of its letter — moves behind that
    body, where the letter it sits on is read.

    Only in LibreOffice's text. WeasyPrint, MuPDF and Chrome write a mark AFTER
    its own letter, and a mark on the edge between two letters is then the
    previous letter's: moving it put the fatha of «الْعَالَمِينَ» on the meem
    («الْعَالمَِينَ», 28 Sept 2026). A lone mark is moved only when these chars
    show LibreOffice's cluster somewhere — a zero-width copy of a letter or a
    mark written twice around its letter — or right after that letter."""
    eps = 0.05
    seen_lo = False
    # the zero-width copy of a body, with the body's marks between them
    i = 0
    while i < len(chars):
        z = chars[i]
        if _is_letter_zw(z):
            k = i - 1
            while k >= 0 and _is_zw_mark(chars[k]):
                k -= 1
            if 0 <= k < i - 1 and chars[k][0] == z[0] and _width(chars[k]) >= 0.01 \
                    and abs(chars[k][1] - z[1]) < 0.12:
                del chars[i]
                seen_lo = True
                continue
        i += 1

    def twin_of(i):
        """(body, the same mark again right behind it) for a mark at chars[i]
        standing on the right edge of the body after it, else (body, None);
        (None, None) when chars[i] is not such a mark."""
        m = chars[i]
        b = chars[i + 1] if i + 1 < len(chars) else None
        if b is None or not (_is_zw_mark(m) and _is_body(b) and abs(m[1] - b[2]) < eps):
            return None, None
        j = i + 2
        while j < len(chars) and _is_zw_mark(chars[j]) and b[1] - 2 * eps <= chars[j][1] <= b[2] + eps:
            if chars[j][0] == m[0]:
                return b, chars[j]
            j += 1
        return b, None

    seen_lo = seen_lo or any(twin_of(i)[1] is not None for i in range(len(chars) - 1))
    # a mark before its body, and the same mark again after it: once, after it
    settled = set()
    i = 0
    while i < len(chars) - 1:
        m = chars[i]
        b, twin = twin_of(i) if id(m) not in settled else (None, None)
        if b is not None:
            if twin is not None:
                settled.add(id(twin))      # the copy behind the body is the one read
                del chars[i]
                continue
            # alone: it belongs to the body it sits on — in LibreOffice's text
            if seen_lo or _lo_cluster_follows(chars, i + 2, m, b):
                chars[i:i + 2] = [b, m]
                settled.add(id(m))
                i += 2
                continue
        i += 1


def _lo_cluster_follows(chars, j, mark, body):
    """True when, before the next letter with a width, chars[j:] holds the same
    mark again or a zero-width copy of `body` (LibreOffice's mark / body / mark /
    copy cluster, with a synthetic space or other marks in between)."""
    while j < len(chars):
        c = chars[j]
        if _is_zw_mark(c):
            if c[0] == mark[0]:
                return True
        elif _is_letter_zw(c):
            if c[0] == body[0]:
                return True
        elif not (_is_synthetic(c) or c[0] == "\u200d"):
            return False
        j += 1
    return False


def _repair_actualtext_clusters(chars):
    """Give back the letters MuPDF strips from glyph clusters in /ActualText.

    Chrome (Skia) and LibreOffice write a cluster whose glyphs do not map one
    to one onto its letters as /Span <</ActualText (…)>>, glyphs in visual order,
    so a mark comes BEFORE the body it sits on. MuPDF hands the text out to the
    glyphs in stream order: the letter lands on the mark (zero width) and the
    body is dropped, or gets the mark ("اً" arrives ً with the alef's box and ا
    at zero width somewhere near it). Worst hit is a font that draws dots as
    separate glyphs — Noto Sans Arabic 2.x from Google Fonts, the face the site
    itself ships: every dotted letter loses its body, MuPDF fills the hole with
    a SYNTHETIC space, or starts a new line when it is wider than 0.8 em, and
    files the letter wherever. Measured 25 Sept 2026 on a Chrome PDF of a lease
    in that face: 47 of 132 letters lost, «الطرف الثاني: محمد إبراهيم علي» read
    «الطر ال ا: محمد رباه م عل» — in the editor (which rewrote the whole line
    from that text on every edit), PDF→Word, audio, CV import and Excel.

      1. LibreOffice maps some bodies to the letter itself (ق, ة): MuPDF then
         gives the letter twice, the body with its width and a zero-width copy
         within 0.03 pt of the body's LEFT edge. The copy goes. (A ligature
         tail sits on its glyph's RIGHT edge; the editor's own ligature letters
         sit 0.25 pt inside the left edge, so "الله" is safe either way.)
      2. A mark WITH width is a body that got the mark's code point: the
         zero-width letter nearest to it (within 0.3 em) takes that box and
         goes before it. The old ±0.5 pt test missed Chrome's Amiri and
         LibreOffice's Arial ("مجاناً" read "مجانً").
      3. A zero-width letter standing in a hole of at least 0.15 em between the
         glyphs and real spaces around it (synthetic spaces not counted) lost
         its body: the letters of one hole share it, split halfway between
         them, and the synthetic spaces in it go. Once a row shows such holes,
         a zero-width letter at the row's right end is its first letter, whose
         body lay beyond the last glyph MuPDF kept.
      4. Each letter so rebuilt moves right behind the glyph touching its right
         edge (its predecessor in reading order), past that glyph's own zero-
         width cluster, or before the glyph touching its left edge when a space
         precedes it.
    A ligature tail parked on its glyph's edge has no hole and no mark with
    width near it, so it is left to _fix_zero_width_clusters as before.
    """
    _drop_marked_letter_copies(chars)
    if not any(_is_letter_zw(c) for c in chars):
        return
    eps = _ZW_EPS
    # 1) LibreOffice's zero-width copy of a body that already carries the letter
    i = 0
    while i < len(chars):
        z = chars[i]
        if _is_letter_zw(z) and any(
                0 <= j < len(chars) and chars[j][0] == z[0] and _width(chars[j]) >= 0.01
                and abs(chars[j][1] - z[1]) < 0.06 for j in (i - 1, i + 1)):
            nxt = chars[i + 1] if i + 1 < len(chars) else None
            if not (z[0] in _ALEFS and nxt is not None and nxt[0] == "ل" and abs(nxt[2] - z[1]) < eps):
                del chars[i]           # (an alef on a lam's right edge is «الا»'s lam-alef tail)
                continue
        i += 1
    grown = []
    # 2) a mark with width is the body of the zero-width letter next to it
    marks = [m for m in chars if unicodedata.combining(m[0]) and _width(m) >= 0.01]
    if marks:
        solid_letters = [c for c in chars if _width(c) >= 0.01 and c[0] not in "  "
                         and not unicodedata.combining(c[0])]
        pairs = []
        for m in marks:
            size = max(m[4], 1.0)
            for z in chars:
                if not _is_letter_zw(z):
                    continue
                dist = max(m[1] - z[1], z[1] - m[2], 0.0)
                if dist > 0.3 * size:
                    continue
                lo, hi = sorted((z[1], (m[1] + m[2]) / 2.0))
                if any(w[1] >= lo + eps and w[2] <= hi - eps for w in solid_letters):
                    continue                       # a letter lies between them
                pairs.append((dist, id(m), id(z), m, z))
        pairs.sort(key=lambda p: p[0])
        used = set()
        for dist, mid, zid, m, z in pairs:
            if mid in used or zid in used:
                continue
            used.update((mid, zid))
            z[1], z[2] = m[1], m[2]
            m[2] = m[1]
            chars.remove(z)
            chars.insert(next(k for k, c in enumerate(chars) if c is m), z)
    # 3) letters that lost their body: share the hole they stand in
    solid = [c for c in chars if _width(c) >= 0.01 and not _is_synthetic(c)]
    synth = [c for c in chars if _is_synthetic(c) and _width(c) >= 0.01]
    holes = {}
    for i, z in enumerate(chars):
        if not _is_letter_zw(z) or _carrier_mark(chars, i) >= 0:
            continue
        if any(w[1] + eps < z[1] < w[2] - eps for w in solid):
            continue                                   # inside a glyph: not ours
        nxt = next((c for c in chars[i + 1:i + 6] if _width(c) >= 0.01), None)
        if nxt is not None and nxt[0] not in "  " and abs(nxt[2] - z[1]) < 0.005:
            # parked EXACTLY on the right edge of the glyph that follows it: a
            # ligature tail (Office gives «محمد» as م + three such tails and a
            # table gap may follow). A mark's origin never lands there to the
            # hundredth of a point (min 0.02 pt over 139 Chrome letters).
            continue
        left = max((w[2] for w in solid if w[2] <= z[1] + eps), default=None)
        right = min((w[1] for w in solid if w[1] >= z[1] - eps), default=None)
        holes.setdefault((left, right), []).append(z)
    rebuilt = False
    edge = []
    for (left, right), zs in holes.items():
        zs.sort(key=lambda c: c[1])
        inner = [s for s in synth if (left is None or s[1] >= left - eps) and (right is None or s[2] <= right + eps)]
        lo = left if left is not None else min([zs[0][1]] + [s[1] for s in inner])
        hi = right if right is not None else max([zs[-1][1]] + [s[2] for s in inner])
        if hi - lo < 0.15 * max(zs[0][4], 1.0):
            if right is None and left is not None:
                edge.append((lo, zs))
            continue
        xs = [c[1] for c in zs]
        cuts = [lo] + [(xs[k] + xs[k + 1]) / 2.0 for k in range(len(xs) - 1)] + [hi]
        for k, z in enumerate(zs):
            z[1], z[2] = cuts[k], cuts[k + 1]
            grown.append(z)
        inner_ids = {id(s) for s in inner}
        chars[:] = [c for c in chars if id(c) not in inner_ids]
        rebuilt = True
    if rebuilt:
        for lo, zs in edge:
            for z in zs:
                if z[0] in _ALEFS:
                    continue                           # a lam-alef tail, not a first letter
                z[1] = lo
                grown.append(z)
    # 4) each rebuilt letter goes where reading order puts it
    grown_ids = {id(z) for z in grown}
    for z in sorted(grown, key=lambda c: -c[2]):
        # the zero-width marks drawn over the rebuilt letter travel with it,
        # unless they sit on the letter they already follow; a mark MuPDF gave
        # twice (Chrome's «عالٍ» came with two kasratan) is kept once
        cluster = [z]
        for k, c in enumerate(chars):
            if c is z or _width(c) >= 0.01 or not unicodedata.combining(c[0]):
                continue
            if not (z[1] - eps <= c[1] <= z[2] + eps):
                continue
            prev = chars[k - 1] if k > 0 else None
            if prev is not None and prev is not z and _width(prev) >= 0.01 and prev[0] not in "  " \
                    and not unicodedata.combining(prev[0]) and prev[1] - eps <= c[1] <= prev[2] + eps:
                continue
            cluster.append(c)
        pos = next(k for k, c in enumerate(chars) if c is z)
        cluster_ids = {id(c) for c in cluster}
        chars[:] = [c for c in chars if id(c) not in cluster_ids]
        seen = set()
        item = [c for c in cluster if c is z or not (c[0] in seen or seen.add(c[0]))]
        pos = min(pos, len(chars))
        pred = [k for k, c in enumerate(chars) if c[0] not in "  " and _width(c) >= 0.01
                and abs(c[1] - z[2]) < eps]
        if len(pred) == 1:
            p = chars[pred[0]]
            q = pred[0] + 1
            while q < len(chars) and _width(chars[q]) < 0.01 and chars[q][0] not in "  " \
                    and id(chars[q]) not in grown_ids and p[1] - eps <= chars[q][1] <= p[2] + eps:
                q += 1      # the zero-width cluster of the predecessor stays with it
        else:
            succ = [k for k, c in enumerate(chars) if c[0] not in "  " and _width(c) >= 0.01
                    and id(c) not in grown_ids and abs(c[2] - z[1]) < eps]
            q = succ[0] if len(succ) == 1 else pos
        chars[q:q] = item


_KASHIDA_STANDINS = set("+&=*#^~|<>_\\@$`") | set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")


def _mark_ascii_kashida(chars):
    """LibreOffice 24.x leaves the kashida it stretches a justified Arabic word
    with out of the ToUnicode map, and MuPDF then names it after whatever byte
    the subset font gave the glyph: "+", "&", "=", "u" (28 Sept 2026: "تقديري
    ةuuu", "لض&بط", "تُ عتم===د" in PDF→Word, PDF→Excel and the editor alike;
    7.1 gave U+FFFD, handled in _line_chars). A run of one such character,
    narrower than a third of an em, drawn in the font of an Arabic letter it
    touches in the stream, becomes the same zero-width joiner: it holds the word
    together and is not text. Punctuation that can stand between Arabic letters
    ("ص.ب", "و/أو") is never taken."""
    i = 0
    while i < len(chars):
        c = chars[i]
        if c[0] not in _KASHIDA_STANDINS:
            i += 1
            continue
        j = i
        while j < len(chars) and chars[j][0] == c[0]:
            j += 1
        prev = chars[i - 1] if i > 0 else None
        nxt = chars[j] if j < len(chars) else None

        def arabic_letter(x):
            return x is not None and _is_arabic_char(x[0]) and not unicodedata.combining(x[0])

        def same_font(x):
            return isinstance(x[3], dict) and isinstance(c[3], dict) and x[3].get("font") == c[3].get("font")

        touching = [x for x in (prev, nxt) if arabic_letter(x) and same_font(x)]
        other_side_ok = all(x is None or x[0] in "  " or arabic_letter(x) or x[0] == "\u200d" for x in (prev, nxt))
        narrow = all(_width(chars[k]) < 0.35 * max(chars[k][4] or 1.0, 1.0) for k in range(i, j))
        # LibreOffice writes a justified word in pieces and the kashida at the END of
        # the next piece, so its stream neighbours are unrelated; what gives it away
        # is the gap it fills exactly, between two letters of its own font
        x0, x1 = min(chars[k][1] for k in range(i, j)), max(chars[k][2] for k in range(i, j))
        bridges = any(arabic_letter(a) and same_font(a) and abs(a[1] - x1) < 0.6 for a in chars) and \
            any(arabic_letter(b) and same_font(b) and abs(b[2] - x0) < 0.6 for b in chars)
        if narrow and j - i <= 8 and ((touching and other_side_ok) or bridges):
            for k in range(i, j):
                chars[k][0] = "\u200d"
                chars[k][3] = dict(chars[k][3], _kashida=True) if isinstance(chars[k][3], dict) else {"_kashida": True}
        i = j


def _line_chars(line):
    """rawdict line → list of [c, x0, x1, span, size] with control chars gone.
    A space MuPDF invented (no glyph in the PDF) carries a span copy marked
    _synthetic, which _repair_actualtext_clusters needs."""
    chars = []
    d = line.get("dir") or (1.0, 0.0)
    horizontal = d[0] > 0.99 and abs(d[1]) < 0.01
    flat = [(ch, sp, float(sp.get("size", 0.0))) for sp in line.get("spans", []) for ch in sp.get("chars", [])]
    for k, (ch, sp, size) in enumerate(flat):
        c = ch.get("c", "")
        if not c:
            continue
        x0, x1 = ch["bbox"][0], ch["bbox"][2]
        if c in "\r\n\t":
            c = " "
        elif c == "\ufffd" and horizontal and 0.01 < x1 - x0 <= max(size, 1.0) and chars \
                and _is_arabic_char(chars[-1][0]) and not unicodedata.combining(chars[-1][0]) \
                and k + 1 < len(flat) and _is_arabic_char(flat[k + 1][0].get("c") or " "):
            # An unmapped glyph with width between two Arabic letters is a
            # KASHIDA: LibreOffice justifies Arabic by stretching words with
            # tatweel glyphs it leaves out of the ToUnicode map. Dropping it
            # left a hole the tokenizer read as a word break — a justified
            # paragraph came out «الإج راءات المتبع ة» (26 of 51 words right,
            # 25 Sept 2026). It stays as a joiner that holds the word together
            # and is left out of the text (see _tokenize); a ZWJ, not a tatweel,
            # so it never counts as an Arabic letter when a page's direction is
            # decided.
            c = "\u200d"
            sp = dict(sp, _kashida=True)
        elif ord(c) < 32 or c in "\ufeff\ufffd" or c in _BIDI_CONTROLS:
            continue
        elif c == " " and horizontal and ch.get("synthetic"):
            sp = dict(sp, _synthetic=True)
        chars.append([c, x0, x1, sp, size])
    _mark_ascii_kashida(chars)
    # A glyph that claims to be an ASCII letter but is narrower than a third of
    # an em, drawn from an Arabic face between Arabic letters, is a mapping
    # artefact (LibreOffice emits a 0.2 em "Q" where it broke a line inside a
    # table cell). A real Latin capital is never that narrow.
    out = []
    for k, ch in enumerate(chars):
        c, x0, x1, sp, size = ch
        if c.isascii() and c.isalpha() and size > 0 and (x1 - x0) < 0.3 * size:
            fname = (sp.get("font") or "").lower()
            if any(w in fname for w in ("arab", "naskh", "amiri", "kufi", "cairo", "tajawal", "almarai")):
                near = [x for x in chars[max(0, k - 2):k] + chars[k + 1:k + 3] if x[0] not in "  "]
                if near and all(_is_arabic_char(x[0]) for x in near):
                    continue
        out.append(ch)
    return out


def _same_line(a, b):
    """True when two chars come from spans on the same printed line. A row can
    merge several lines (a table header whose cells wrap), and x alone would
    then match a space on one line against a glyph on another."""
    ya = (a[3] or {}).get("bbox") if isinstance(a[3], dict) else None
    yb = (b[3] or {}).get("bbox") if isinstance(b[3], dict) else None
    if not ya or not yb:
        return False
    overlap = min(ya[3], yb[3]) - max(ya[1], yb[1])
    return overlap > 0.5 * max(0.1, min(ya[3] - ya[1], yb[3] - yb[1]))


def _is_indic(ch):
    """Devanagari to Sinhala (U+0900-U+0DFF): scripts whose vowel signs and viramas
    PyMuPDF reports as zero-width glyphs."""
    return "\u0900" <= ch <= "\u0dff"


def _office_font(span):
    """True for a font named the way Microsoft Office's PDF export names it:
    the subset tag ABCDEE+, or no subset tag with Windows' "Arial,Bold" style
    (Chrome, LibreOffice and WeasyPrint tag every subset AAAAAA+, BAAAAA+ …)."""
    name = (span.get("font") or "") if isinstance(span, dict) else ""
    return bool(name) and (name.startswith("ABCDEE+") or "+" not in name)


def _office_yeh(chars):
    """Microsoft Office maps Arial's medial yeh glyph to U+064C (dammatan) and
    its final yeh to U+064B (fathatan): the backend's own test.pdf (Access 2010,
    Arial Bold) read «التعلٌم», «رئٌس», «المدٌرٌات» and «العظم ً» for
    «التعليم», «رئيس», «المديريات», «العظمي» — 399 glyphs on 15 pages, 5% of
    its Arabic words (10 Oct 2026). Those code points then carry a letter's
    full advance; a real tanween has no width of its own, and when Chrome gives
    it its cluster's width the letter it carries stands at zero width beside it
    (which _repair_actualtext_clusters, run after this, takes back). So a dammatan
    or fathatan at least a quarter of an em wide, with no zero-width letter
    within 0.3 em, in a font named the Office way, is the yeh."""
    for c in chars:
        if c[0] not in "\u064B\u064C":
            continue
        size = max(c[4] or 1.0, 1.0)
        if _width(c) < 0.25 * size or not _office_font(c[3]):
            continue
        reach = 0.3 * size
        if any(_is_letter_zw(o) and c[1] - reach <= o[1] <= c[2] + reach for o in chars):
            continue
        c[0] = "\u064A"


def _tokenize(chars):
    """Split a row's chars into words: on spaces, and on a visible gap between
    neighbouring glyphs (LibreOffice writes no space character between an Arabic
    word and a digit in another font — the gap is the only evidence)."""
    _office_yeh(chars)
    _repair_actualtext_clusters(chars)
    _fix_zero_width_clusters(chars)
    tokens = []
    cur = None
    prev = None
    prev_w = None                 # the last glyph that has a width
    gap_at = None                 # the space that ended the last word
    for k, ch in enumerate(chars):
        c, x0, x1, sp, size = ch
        if c in "  ":
            # Word, LibreOffice and Chrome all write a narrow "space" INSIDE the
            # glyph of a letter carrying a shadda ("المحوّل" arrives المحوّ ␠ ل,
            # "خصّص" as خصّ ␠ ص), which split one word in two. A real space
            # starts where the previous glyph ends; this one lies within it.
            # (Rotated text gives every glyph the same x box, so the space must
            # also be narrower than the glyph it sits in.)
            if (cur is not None and prev_w is not None and _same_line(ch, prev_w)
                    and x0 >= prev_w[1] - _ZW_EPS and x1 <= prev_w[2] + _ZW_EPS
                    and (x1 - x0) < _width(prev_w) - _ZW_EPS):
                continue
            # A space MuPDF made up inside the NEXT Arabic glyph is no gap
            # either: LibreOffice's «شكرًا» arrives ر ً ␠ ا with that space
            # inside the alef's box, and came out «شكرً ا» (10 Oct 2026).
            if cur is not None and _is_synthetic(ch) and cur["chars"] and _is_arabic_char(cur["chars"][-1][0]):
                nxt = next((n for n in chars[k + 1:k + 6] if _width(n) >= 0.01), None)
                if (nxt is not None and nxt[0] not in "  " and _is_arabic_char(nxt[0]) and _same_line(ch, nxt)
                        and x0 >= nxt[1] - _ZW_EPS and x1 <= nxt[2] + _ZW_EPS):
                    continue
            cur = None
            prev = None
            prev_w = None
            gap_at = ch
            continue
        # The gap is measured from the last glyph with a width: a zero-width
        # tail (the ى of "إلى") used to skip the test, and the Latin word printed
        # right after it joined the Arabic one ("إلىPDF", "فيPDF").
        # Not between two Indic characters: Devanagari vowel signs (ा ी ि) arrive
        # with zero width, so the next consonant looked 2.9 pt away from the last
        # glyph with a width and every word broke at its matras ("लगातार" read
        # "लगा ता र": 73 of 88 words of a LibreOffice Hindi page, 28 Sept 2026).
        if cur is not None and prev_w is not None and (x1 - x0) > 0.01 \
                and not (_is_indic(c) and prev is not None and _is_indic(prev[0])):
            gap = max(prev_w[1] - x1, x0 - prev_w[2])
            if gap > 0.2 * max(size, prev_w[4], 1.0):
                cur = None
        # A size change inside a run of Arabic letters is a field boundary (a 10 pt
        # title running into its 8 pt date on the same row), never one word.
        if cur is not None and prev is not None and abs(size - prev[4]) >= 1.0 \
                and (_is_arabic_char(c) or _is_arabic_char(prev[0])):
            cur = None
        if cur is None and unicodedata.combining(c) and tokens:
            if (x1 - x0) < 0.01 and gap_at is not None and gap_at[1] - _ZW_EPS <= x0 <= gap_at[2] + _ZW_EPS:
                # a stray copy of the last word's tanween, drawn in the space after
                # it (Chrome, DejaVu Sans): it stays with that word, and the next
                # letter starts a word of its own — «أولاً مثلاً» came out as one
                # word «أولاًًمثلاً» (10 Oct 2026)
                tokens[-1]["chars"].append(ch)
                continue
            cur = tokens[-1]          # a mark never starts a word
        if cur is None:
            cur = {"chars": [], "x0": x0, "x1": x1, "span": sp, "size": size}
            tokens.append(cur)
        cur["chars"].append(ch)
        cur["x0"] = min(cur["x0"], x0)
        cur["x1"] = max(cur["x1"], x1)
        prev = ch
        if (x1 - x0) > 0.01:
            prev_w = ch
    for t in tokens:
        _token_text(t)
    tokens = [t for t in tokens if t["text"]]
    return tokens


def _token_text(t):
    _settle_token_punctuation(t)
    _fix_digit_runs(t)
    # a justification kashida (see _line_chars) holds the word together
    # but is not part of the text
    text = "".join(c[0] for c in t["chars"] if not (isinstance(c[3], dict) and c[3].get("_kashida")))
    # The same harakah twice in a row is never text: Chrome hands a line-end
    # tanween over twice, once with its cluster and once on its own (10 Oct 2026).
    t["text"] = _DOUBLE_MARK_RE.sub(r"\1", text) if "\u064B" <= max(text or " ") else text


_DOUBLE_MARK_RE = re.compile(r"([\u064B-\u0652])\1+")


def _join_touching_words(tokens):
    """RTL tokens in reading order. Two Arabic words that touch — no gap, no
    space between them — are one word whose pieces MuPDF filed in two places:
    Chrome writes a line's last glyph cluster at the START of the line's text,
    so «مثلاً أولاً.» at the end of an Amiri line came out «مثلاً أول اً.» and
    was read aloud as two words (10 Oct 2026). Words of one size only (a size
    change is a field boundary, see _tokenize)."""
    out = []
    for t in tokens:
        p = out[-1] if out else None
        if (p is not None and _token_kind(p["text"]) == "R" and _token_kind(t["text"]) == "R"
                and abs(p["x0"] - t["x1"]) < 0.25 and abs(p["size"] - t["size"]) < 1.0
                and p.get("chars") and t.get("chars") and _same_line(p["chars"][0], t["chars"][0])):
            joined = dict(p, chars=p["chars"] + t["chars"], x0=min(p["x0"], t["x0"]), x1=max(p["x1"], t["x1"]))
            _token_text(joined)
            out[-1] = joined
            continue
        out.append(t)
    return out


def _join_kashida_splits(tokens, chars):
    """A justified Arabic word that LibreOffice stretched with a kashida is written
    in two pieces with the kashida between them and filed at the end of the NEXT
    piece, so the pieces came out as two words: "وخط ة", "التط وير", "الثقيل ة"
    (28 Sept 2026, every justified paragraph of a LibreOffice 24.2 PDF). The
    kashida's box still sits exactly in the gap: the piece to its right and the
    piece to its left are one word, read right piece first."""
    kash = sorted((c[1], c[2]) for c in chars
                  if isinstance(c[3], dict) and c[3].get("_kashida") and c[2] - c[1] > 0.3)
    if not kash:
        return tokens
    spans = []
    for k0, k1 in kash:                       # consecutive kashida glyphs are one stretch
        if spans and k0 - spans[-1][1] < 0.6:
            spans[-1][1] = max(spans[-1][1], k1)
        else:
            spans.append([k0, k1])
    toks = list(tokens)
    for k0, k1 in spans:
        right = [t for t in toks if abs(t["x0"] - k1) < 1.0 and _token_kind(t["text"]) == "R"]
        left = [t for t in toks if abs(t["x1"] - k0) < 1.0 and _token_kind(t["text"]) == "R"]
        if len(right) != 1 or len(left) != 1 or right[0] is left[0]:
            continue
        a, b = right[0], left[0]
        a["text"] = a["text"] + b["text"]
        a["chars"] = a["chars"] + b["chars"]
        a["x0"] = min(a["x0"], b["x0"])
        toks = [t for t in toks if t is not b]
    return toks


def _normalize_brackets(tokens):
    """Cairo maps a mirrored bracket glyph to the mirrored character, so an RTL
    "(النص)" extracts as ")النص(". In reading order the first bracket of a
    balanced pair must open, so re-derive each bracket from its nesting."""
    depth = 0
    for t in tokens:
        if not any(ch in _BRACKET_PAIR for ch in t["text"]):
            continue
        out = []
        for ch in t["text"]:
            if ch in _OPENERS or ch in _CLOSERS:
                if depth <= 0:
                    ch = ch if ch in _OPENERS else _BRACKET_PAIR[ch]
                    depth += 1
                else:
                    ch = ch if ch in _CLOSERS else _BRACKET_PAIR[ch]
                    depth -= 1
            out.append(ch)
        t["text"] = "".join(out)


def _finish_row(row, page_rtl):
    """Order a merged row's words and pick its face. Returns None if empty."""
    tokens = _tokenize(row["chars"])
    if not tokens:
        return None
    text = "".join(t["text"] for t in tokens)
    ar = sum(1 for c in text if _is_arabic_char(c))
    lat = sum(1 for c in text if c.isascii() and c.isalpha())
    # A row that carries Arabic reads right-to-left on an Arabic page even when
    # an e-mail address outnumbers the Arabic letters; on a Latin page it only
    # does so when Arabic dominates.
    rtl = ar > 0 and (page_rtl or ar >= lat)
    if rtl:
        tokens = _join_kashida_splits(tokens, row["chars"])
    tokens = _order_tokens(tokens, rtl)
    if rtl:
        tokens = _join_touching_words(tokens)
        _normalize_brackets(tokens)
        for t in tokens:
            # A full stop drawn LEFT of a number or Latin word that ends an RTL
            # line closes the sentence: «رقم قومي 29801011234567.» came out
            # «قومي.29801011234567» for every producer (25 Sept 2026).
            ch = t.get("chars") or []
            if ch and ch[0][0] in _END_PUNCT and not any(_is_arabic_char(c[0]) for c in ch):
                k = 0
                while k < len(ch) and ch[k][0] in _END_PUNCT:
                    k += 1
                if k < len(ch) and ch[k][0].isalnum() and ch[0][1] <= min(c[1] for c in ch) + 0.01:
                    t["text"] = t["text"][k:] + t["text"][:k]
    elif page_rtl and tokens:
        # An English sentence in an RTL document ends with its full stop on the
        # LEFT: Chrome's «.This line is English: … month» (25 Sept 2026).
        t = tokens[0]
        ch = t.get("chars") or []
        k = 0
        while k < len(ch) and ch[k][0] in _END_PUNCT:
            k += 1
        if 0 < k < len(ch) and ch[k][0].isalnum():
            tokens[-1]["text"] = tokens[-1]["text"] + t["text"][:k]
            t["text"] = t["text"][k:]
    weights = {}
    for t in tokens:
        key = id(t["span"])
        weights[key] = (weights.get(key, (0, None))[0] + len(t["text"]), t["span"])
    sp = max(weights.values(), key=lambda v: v[0])[1] or {}
    fname = (sp.get("font") or "").lower()
    bold = bool(sp.get("flags", 0) & 16) or "bold" in fname or "black" in fname or "heavy" in fname
    return {"tokens": tokens, "rtl": rtl, "size": max(t["size"] for t in tokens),
            "bold": bold, "color": int(sp.get("color", 0)), "y0": row["y0"], "y1": row["y1"],
            "x0": min(t["x0"] for t in tokens), "x1": max(t["x1"] for t in tokens),
            "block": row["block"]}


def _frag_baseline(f):
    """(baseline y, font size) of a fragment's largest span, or None."""
    best = None
    for c in f["chars"]:
        sp = c[3]
        if isinstance(sp, dict) and sp.get("origin") is not None and (best is None or c[4] > best[1]):
            best = (float(sp["origin"][1]), float(c[4] or 0.0))
    return best


def _merge_rows(frags):
    """Fragments (chars + y-range) that share a baseline become one row, whatever
    block PyMuPDF filed them under — a job title and the date printed beside it
    are one line to the reader.

    The boxes alone are not enough: Chrome grows the box of a line that carries
    a shadda or tanween by half a line, so two consecutive lines of one Arabic
    paragraph "overlapped" and were merged into one row, their words interleaved
    ("عالج المؤسسات. ملفاتك في ثوانٍ ملفاتك وليس مشفرة دقائق", 22 Sept 2026).
    Two fragments whose baselines are most of a line apart are never one row;
    a superscript or a smaller date beside a title still is."""
    frags.sort(key=lambda f: (f["y0"] + f["y1"]) / 2.0)
    rows = []
    for f in frags:
        cy = (f["y0"] + f["y1"]) / 2.0
        h = max(f["y1"] - f["y0"], 1.0)
        fb = _frag_baseline(f)
        same_row = False
        if rows:
            r = rows[-1]
            overlap = min(f["y1"], r["y1"]) - max(f["y0"], r["y0"])
            same_row = overlap > 0.5 * min(h, max(r["y1"] - r["y0"], 1.0))
            rb = r.get("base")
            if same_row and fb is not None and rb is not None \
                    and abs(fb[0] - rb[0]) > 0.6 * max(fb[1], rb[1], 1.0):
                same_row = False
        if same_row:
            r = rows[-1]
            r["chars"].extend(f["chars"])
            r["y0"] = min(r["y0"], f["y0"])
            r["y1"] = max(r["y1"], f["y1"])
            if fb is not None and (r.get("base") is None or fb[1] > r["base"][1]):
                r["base"] = fb
            if (f["x1"] - f["x0"]) > r["bw"]:
                r["bw"] = f["x1"] - f["x0"]
                r["block"] = f["block"]
        else:
            rows.append({"chars": list(f["chars"]), "y0": f["y0"], "y1": f["y1"], "cy": cy,
                         "block": f["block"], "bw": f["x1"] - f["x0"], "base": fb})
    return rows


def _page_columns(blocks, pw):
    """Return a split x (or None) if the text blocks sit in two clear columns."""
    if len(blocks) < 4:
        return None
    best = None
    for frac in range(25, 76):
        x = pw * frac / 100.0
        if any(b["x0"] < x - 2 and b["x1"] > x + 2 for b in blocks):
            continue
        left = sum(1 for b in blocks if b["x1"] <= x)
        right = sum(1 for b in blocks if b["x0"] >= x)
        if left >= 2 and right >= 2:
            score = min(left, right)
            if best is None or score > best[0]:
                best = (score, x)
    return best[1] if best else None


def _side_by_side(ls, rs):
    """True when the lines of the left and right sides of a two-column part
    are two columns of reading text rather than the cells of a borderless
    table: at least three lines a side, and when the lines pair up row by row
    on the same baselines, both sides must be filled to their width (justified
    or ragged prose; a description and an amount column are not)."""
    if len(ls) < 3 or len(rs) < 3:
        return False
    few, many = (ls, rs) if len(ls) <= len(rs) else (rs, ls)
    paired = sum(1 for a in few if any(abs((a[2] + a[3]) - (b[2] + b[3])) < 0.3 * (a[3] - a[2]) for b in many))
    if paired < 0.7 * len(few):
        return True

    def full(lines):
        x0 = min(l[0] for l in lines)
        x1 = max(l[1] for l in lines)
        widths = sorted(l[1] - l[0] for l in lines)
        return widths[len(widths) // 2] >= 0.75 * max(x1 - x0, 1.0)
    return full(ls) and full(rs)


def _page_bands(blocks, spanning, pw):
    """Two columns under (or between) full-width parts: a paper's title over its
    two columns, a CV's name over its sidebar and main column, a letterhead with
    the sender on the left and the date on the right above the letter.

    _page_columns() gives up as soon as one block crosses the split, so all of
    these were read as one column and the two sides were woven together line by
    line: «Python Lead Data Engineer, Fintech Ltd», «Spark The quarterly review
    covers…», «Acme Trading Ltd Ref: AC/2026/114 12 High Street Date: 2 October
    2026» (2 Oct 2026, a WeasyPrint paper, CV and letter; the paper's text kept
    18% of its word order). It also wants two blocks a side, so a page holding a
    heading and a paragraph beside the column carried over from the page before
    was woven together too. Here a block that crosses the split is allowed when
    no block of either column stands beside it: the page is then a stack of
    full-width parts and two-column parts, read top to bottom, each two-column
    part left column first (right first on an Arabic page).

    `blocks` are text blocks {x0, x1, y0, y1, lines: [(x0, x1, y0, y1)]},
    `spanning` pictures and ruled tables (they may cross the split but never
    start a column). A two-column part is read column by column only when
    _side_by_side() says its sides are reading text; otherwise (a borderless
    table, a title with its date on the right, a "Name … Date …" line) its
    rows keep their words side by side as before.
    Returns (split x, [(y0, y1) of each full-width part], [True when part k
    is read column by column], ids of the one-sided blocks read full width)
    or None."""
    if len(blocks) < 3:
        return None
    best = None
    for frac in range(20, 81):
        x = pw * frac / 100.0
        lone = set()
        for _ in range(4):
            got = _bands_at(x, blocks, spanning, lone)
            if got is None or got[0] == lone:
                break
            lone = got[0]
        if got is None:
            continue
        found, cuts, modes, score, gap, mid = got
        if not score:
            continue
        key = (score, gap, -abs(x - mid))
        if best is None or key > best[0]:
            best = (key, x, cuts, modes, found)
    return best[1:] if best else None


def _bands_at(x, blocks, spanning, lone):
    """_page_bands() for one split x. `lone` holds the ids of blocks that sit on
    one side only, above or below the rows where both sides have text and set
    apart from them by more than a paragraph gap ("Dear Ms Jones," under a
    letterhead whose date stands on the right): they are read full width.
    Returns (lone ids found now, cuts, modes, score, gutter width, gutter middle)
    or None when x cannot split the page."""
    def crosses(b):
        return id(b) in lone or (b["x0"] < x - 2 and b["x1"] > x + 2)
    crossing = [b for b in blocks + spanning if crosses(b)]
    left = [b for b in blocks if not crosses(b) and (b["x0"] + b["x1"]) / 2.0 < x]
    right = [b for b in blocks if not crosses(b) and (b["x0"] + b["x1"]) / 2.0 >= x]
    if not left or not right:
        return None
    # nothing that crosses the split may stand beside a column block
    if any(min(c["y1"], b["y1"]) - max(c["y0"], b["y0"]) > 2.0 for c in crossing for b in left + right):
        return None
    # the full-width parts: crossing blocks with no column block between them
    cuts = []
    for c in sorted(crossing, key=lambda c: c["y0"]):
        if cuts and not any(cuts[-1][1] - 2.0 <= (b["y0"] + b["y1"]) / 2.0 <= c["y0"] + 2.0
                            for b in left + right):
            cuts[-1][1] = max(cuts[-1][1], c["y1"])
        else:
            cuts.append([c["y0"], c["y1"]])
    modes, score, found = [], 0, set(lone)
    for k in range(len(cuts) + 1):
        top = cuts[k - 1][1] if k > 0 else float("-inf")
        bottom = cuts[k][0] if k < len(cuts) else float("inf")
        lb = [b for b in left if top - 2.0 <= (b["y0"] + b["y1"]) / 2.0 <= bottom + 2.0]
        rb = [b for b in right if top - 2.0 <= (b["y0"] + b["y1"]) / 2.0 <= bottom + 2.0]
        ls = [ln for b in lb for ln in b["lines"]]
        rs = [ln for b in rb for ln in b["lines"]]
        cols = bool(ls and rs) and _side_by_side(ls, rs)
        modes.append(cols)
        if not cols:
            continue
        score += min(len(ls), len(rs))
        heights = sorted(ln[3] - ln[2] for ln in ls + rs)
        para_gap = 1.6 * heights[len(heights) // 2]
        both0 = max(min(b["y0"] for b in lb), min(b["y0"] for b in rb))
        both1 = min(max(b["y1"] for b in lb), max(b["y1"] for b in rb))
        for side in (lb, rb):
            col = sorted(side, key=lambda b: b["y0"])
            # downwards from the last block beside the other side
            edge = None
            for b in col:
                if b["y0"] < both1 - 2.0:
                    edge = b["y1"] if edge is None else max(edge, b["y1"])
                elif edge is not None and b["y0"] - edge > para_gap:
                    found.update(id(c) for c in col if c["y0"] >= b["y0"])
                    break
                else:
                    edge = b["y1"]
            # upwards from the first block beside the other side
            edge = None
            for b in reversed(col):
                if b["y1"] > both0 + 2.0:
                    edge = b["y0"] if edge is None else min(edge, b["y0"])
                elif edge is not None and edge - b["y1"] > para_gap:
                    found.update(id(c) for c in col if c["y1"] <= b["y1"])
                    break
                else:
                    edge = b["y0"]
    gap = min(b["x0"] for b in right) - max(b["x1"] for b in left)
    return found, [tuple(c) for c in cuts], modes, score, gap, max(b["x1"] for b in left) + gap / 2.0


def _tidy(text):
    text = _SPACE_BEFORE_RE.sub(r"\1", text)
    text = _SPACE_AFTER_RE.sub(r"\1", text)
    text = _SLASH_RE.sub(r"\1/\2", text)
    return text.strip()


def _valid_layout(layout, rows, ncols):
    """The merged-cell layout if it still fits the rows (every row covers the
    grid exactly, a continued cell sits under a cell of the same span), else None."""
    if not layout or len(layout) != len(rows):
        return None
    above = {}
    for entries in layout:
        pos, here = 0, {}
        for e in entries:
            if e["span"] < 1 or not 0 <= e["col"] < ncols:
                return None
            if e["vmerge"] == "continue" and above.get(pos) != e["span"]:
                return None
            if e["vmerge"]:
                here[pos] = e["span"]
            pos += e["span"]
        if pos != ncols:
            return None
        above = here
    return layout


def _add_word_table(doc, rows, rtl, layout=None, col_x=None, size=None, row_y=None, y1=None):
    """A ruled PDF table as a real Word table: header row bold, Arabic cells
    right-aligned, and the column order mirrored for an RTL table (the rows
    arrive in reading order, first cell = rightmost). Cells the PDF merges are
    merged (w:gridSpan / w:vMerge), the columns keep the PDF's widths, the rows
    at least the PDF's heights and the text its size; until 28 Sept 2026 every
    column was equally wide, a merged cell was one squeezed cell beside empty
    ones, and the cells took the body text's size."""
    from docx.oxml.ns import qn
    from docx.shared import Twips
    ncols = max(len(r) for r in rows)
    layout = _valid_layout(layout, rows, ncols)
    tbl = doc.add_table(rows=len(rows), cols=ncols)
    try:
        tbl.style = "Table Grid"
    except Exception:  # noqa: BLE001 - a template without the style still gets a table
        pass
    if rtl:
        _put(tbl._tbl.tblPr, "w:bidiVisual", _TBLPR_ORDER)
    widths = None
    if col_x and len(col_x) == ncols:
        widths = [max(144, int(round((c1 - c0) * 20))) for c0, c1 in col_x]   # twips, 0.1" at least
        if rtl:
            widths.reverse()            # the grid runs in reading order too
        for gc, w in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), widths):
            gc.set(qn("w:w"), str(w))
        _put(tbl._tbl.tblPr, "w:tblW", _TBLPR_ORDER, w=str(sum(widths)), type="dxa")
        _put(tbl._tbl.tblPr, "w:tblLayout", _TBLPR_ORDER, type="fixed")
    heights = None
    if row_y and y1 and len(row_y) == len(rows) and all(b > a for a, b in zip(row_y, row_y[1:])) \
            and y1 > row_y[-1]:
        heights = [b - a for a, b in zip(row_y, list(row_y[1:]) + [y1])]
    from docx.table import _Cell, _Row
    from docx.enum.table import WD_ROW_HEIGHT_RULE
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    # Cells straight from the XML: python-docx's table.cell() rebuilds the whole
    # grid on every call, and a 15-page table report spent 188 of 201 s there.
    for ri, (r, tr) in enumerate(zip(rows, tbl._tbl.tr_lst)):
        if heights and 6 <= heights[ri] <= 400:
            row = _Row(tr, tbl)
            row.height = Twips(int(round(heights[ri] * 20)))
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        tcs = list(tr.tc_lst)
        entries = layout[ri] if layout else [{"span": 1, "vmerge": None, "col": ci} for ci in range(ncols)]
        pos = 0
        for tc, e in zip(tcs, entries):
            span = e["span"]
            if span > 1:
                tc.grid_span = span
            if e["vmerge"]:
                tc.vMerge = e["vmerge"]
            if widths:
                tc.width = Twips(sum(widths[pos:pos + span]))
            if e.get("fill"):
                # the cell's own colour (a shaded header, zebra rows)
                _put(tc.get_or_add_tcPr(), "w:shd", _TCPR_ORDER, val="clear", color="auto", fill=e["fill"])
            if e.get("valign") == "center":
                _put(tc.get_or_add_tcPr(), "w:vAlign", _TCPR_ORDER, val="center")
            pos += span
            ci = e["col"]
            text = _xml_clean(r[ci]) if ci < len(r) and e["vmerge"] != "continue" else ""
            cell = _Cell(tc, tbl)
            cell.text = text
            para = cell.paragraphs[0]
            for run in para.runs:
                # with the cells' own look read from the PDF, a header row is bold
                # only when it is printed bold (a plain header came back bold)
                _style_run(run, size=size, bold=bool(e.get("bold")) if layout else (ri == 0 or e.get("bold")),
                           color=e.get("color"))
            arabic_cell = is_arabic_text(text)
            if arabic_cell:
                _set_paragraph_rtl(para)
            side = e.get("align")
            if text and side:
                # where the PDF put the text in its cell; a bidi paragraph reads
                # left/right as start/end, so the page's right is its "left"
                if side == "center":
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif arabic_cell:
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT if side == "right" else WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT if side == "right" else WD_ALIGN_PARAGRAPH.LEFT
        for tc in tcs[len(entries):]:
            tr.remove(tc)
    doc.add_paragraph()


def _same_columns(a, b, tol=4.0):
    ca, cb = a.get("col_x"), b.get("col_x")
    if not ca or not cb or len(ca) != len(cb):
        return False
    return all(abs(p[0] - q[0]) <= tol and abs(p[1] - q[1]) <= tol for p, q in zip(ca, cb))


def _merge_stacked_tables(tables):
    """A zebra-striped table comes back from find_tables() as one small table
    per shaded band. Bands with the same columns, stacked with at most a couple
    of rows between them, are one table."""
    out = []
    for t in sorted(tables, key=lambda t: t["y0"]):
        prev = out[-1] if out else None
        if prev is not None and prev.get("row_y") and t.get("row_y") and _same_columns(prev, t):
            pitch = (prev["y1"] - prev["y0"]) / max(1, len(prev["rows"]))
            if t["y0"] - prev["y1"] < 2.5 * max(pitch, 1.0):
                prev["rows"] = prev["rows"] + t["rows"]
                prev["row_y"] = prev["row_y"] + t["row_y"]
                prev["layout"] = (prev["layout"] + t["layout"]
                                  if prev.get("layout") and t.get("layout") else None)
                prev["y1"] = max(prev["y1"], t["y1"])
                prev["x0"] = min(prev["x0"], t["x0"])
                prev["x1"] = max(prev["x1"], t["x1"])
                continue
        out.append(t)
    return out


def _absorb_rows_into_tables(rows, tables, rtl):
    """With the bands merged, the white rows between them are still loose text
    rows (no ruling of their own). A text row inside a table's vertical span
    that fills at least two of its columns is put back in as a table row, and
    so is a white row right under the last band that fills most of them: an
    invoice's eighth item came back as a line of text under a seven-row table
    (2 Oct 2026). An absorbed row takes the cell alignment of the row above.
    Returns the rows that stay text."""
    keep = []
    for r in sorted(rows, key=lambda r: r["y0"]):
        absorbed = False
        cy = (r["y0"] + r["y1"]) / 2.0
        for t in tables:
            cols = t.get("col_x")
            if not cols or not t.get("row_y") or r["x1"] < t["x0"] - 2 or r["x0"] > t["x1"] + 2:
                continue
            pitch = (t["y1"] - t["y0"]) / max(1, len(t["rows"]))
            below = not (t["y0"] <= cy <= t["y1"]) and 0 <= r["y0"] - t["y1"] < 0.6 * pitch
            if not (t["y0"] <= cy <= t["y1"]) and not below:
                continue
            cells = [[] for _ in cols]
            placed = []
            for tok in r["tokens"]:
                best, bi = 0.0, None
                for i, (c0, c1) in enumerate(cols):
                    ov = min(tok["x1"], c1) - max(tok["x0"], c0)
                    if ov > best:
                        best, bi = ov, i
                if bi is not None:
                    cells[bi].append(_normalize_presentation_forms(tok["text"]))
                    placed.append((tok["x0"], tok["x1"], bi))
            texts = [_tidy(" ".join(c)) for c in cells]
            filled = sum(1 for x in texts if x)
            if filled < 2 or (below and filled < max(2, 0.6 * len(cols))):
                continue
            if below:
                # a row of cells, not a line of prose that happens to cross the
                # columns: its words change column only across a wide gap
                placed.sort()
                if any(b[2] != a[2] and b[0] - a[1] < 0.8 * max(r["size"], 1.0) for a, b in zip(placed, placed[1:])):
                    continue
            if rtl:
                texts.reverse()
            at = sum(1 for y in t["row_y"] if y <= r["y0"])
            t["rows"].insert(at, texts)
            t["row_y"].insert(at, r["y0"])
            if below:
                t["y1"] = max(t["y1"], r["y1"])
            if t.get("layout"):
                entries = [{"span": 1, "vmerge": None, "col": i} for i in range(len(texts))]
                above = t["layout"][at - 1] if at > 0 else None
                if above and len(above) == len(entries) and all(e["span"] == 1 for e in above):
                    for e, a in zip(entries, above):
                        for k in ("align", "valign"):
                            if a.get(k):
                                e[k] = a[k]
                t["layout"].insert(at, entries)
            absorbed = True
            break
        if not absorbed:
            keep.append(r)
    return keep


# ── Page layout → Word ──────────────────────────────────────────────────────
# PDF font names → the family Word knows. Subset prefixes and style suffixes go
# ("BAAAAA+ArialMT" → "Arial", "NotoNaskhArabic-Bold" → "Noto Naskh Arabic"),
# and the metric twins LibreOffice and Chrome embed are named after the Office
# face they stand in for, so the text wraps in Word as it did in the PDF.
_FAMILY_ALIASES = {
    "arial": "Arial", "liberationsans": "Arial", "arimo": "Arial", "helvetica": "Arial",
    "timesnewroman": "Times New Roman", "liberationserif": "Times New Roman", "tinos": "Times New Roman",
    "times": "Times New Roman", "timesroman": "Times New Roman",
    "calibri": "Calibri", "carlito": "Calibri", "cambria": "Cambria", "caladea": "Cambria",
    "couriernew": "Courier New", "liberationmono": "Courier New", "cousine": "Courier New", "courier": "Courier New",
    "notonaskharabic": "Noto Naskh Arabic", "notosansarabic": "Noto Sans Arabic", "notokufiarabic": "Noto Kufi Arabic",
    "ibmplexsansarabic": "IBM Plex Sans Arabic", "dejavusans": "DejaVu Sans", "dejavuserif": "DejaVu Serif",
    "traditionalarabic": "Traditional Arabic", "simplifiedarabic": "Simplified Arabic",
    "sakkalmajalla": "Sakkal Majalla", "arabictypesetting": "Arabic Typesetting", "segoeui": "Segoe UI",
    "scheherazadenew": "Scheherazade New", "opensymbol": None, "symbol": None, "wingdings": None,
}
_STYLE_SUFFIXES = ("bolditalic", "boldoblique", "semibold", "extrabold", "bold", "italic", "oblique",
                   "regular", "medium", "light", "black", "book")


@functools.lru_cache(maxsize=512)
def _font_family(name):
    """"BAAAAA+NotoNaskhArabic-Bold" → "Noto Naskh Arabic", "ArialMT" → "Arial",
    "TimesNewRomanPSMT" → "Times New Roman", "Cairo-SemiBold" → "Cairo"; None for
    symbol faces and nameless fonts."""
    if not name:
        return None
    n = name.split("+", 1)[-1].split(",", 1)[0]
    # WeasyPrint and Cairo write the family's words with hyphens
    # ("Liberation-Serif-Bold", "Noto-Sans-Arabic", "IBM-Plex-Sans-Arabic"):
    # only the first word was kept, and "Liberation" or "Noto" is no font, so
    # Word and LibreOffice drew every run in their fallback face (2 Oct 2026).
    # The style words at the end go; the family's own words stay.
    parts = [q for q in n.split("-") if q]
    while len(parts) > 1 and _is_style_word(parts[-1]):
        parts.pop()
    if not parts:
        return None
    n = re.sub(r"(PSMT|PS|MT)$", "", parts[-1])
    for suffix in _STYLE_SUFFIXES:
        if n.lower().endswith(suffix) and len(n) > len(suffix) + 2:
            n = n[: -len(suffix)]
            break
    parts[-1] = n
    key = re.sub(r"[^a-z]", "", "".join(parts).lower())
    if key in _FAMILY_ALIASES:
        return _FAMILY_ALIASES[key]
    if len(parts) > 1:
        return " ".join(parts).strip() or None
    spaced = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", n).strip()
    return spaced or None


def _is_style_word(part):
    """"Bold", "BoldItalic", "SemiBoldOblique", "BoldMT", "Regular": a weight or
    slant, not a word of the family's name."""
    w = re.sub(r"(PSMT|PS|MT)$", "", part).lower()
    if not w:
        return True
    while w:
        for suffix in _STYLE_SUFFIXES + ("semi", "extra", "ultra", "demi", "condensed", "roman"):
            if w.endswith(suffix):
                w = w[: -len(suffix)]
                break
        else:
            return False
    return True


def _row_fonts(lines):
    """(Latin family, Arabic family) that drew most of these rows' letters."""
    from collections import Counter
    lat, ar = Counter(), Counter()
    for ln in lines:
        for t in ln["tokens"]:
            for c in t.get("chars", ()):
                sp = c[3] if isinstance(c[3], dict) else None
                fam = _font_family(sp.get("font")) if sp else None
                if not fam:
                    continue
                if _is_arabic_char(c[0]):
                    ar[fam] += 1
                elif c[0].isascii() and c[0].isalpha():
                    lat[fam] += 1
    return (lat.most_common(1)[0][0] if lat else None, ar.most_common(1)[0][0] if ar else None)


# The pictures Word is given (19 Oct 2026). python-docx reads a JPEG only when it starts with
# a JFIF or Exif header, and a picture add_picture refused was left out without a word: a
# 17-page report Chrome had printed came back with one of its seven photos, because Skia
# writes its JPEGs without that header (Ghostscript too, so a PDF that had been through any
# compressor lost every photo). A picture with a soft mask (a logo on a transparent ground)
# went in without its mask, on black, and a picture the page shows only part of (Chrome's
# object-fit, a crop made in Word or PowerPoint) went in whole, taller than the page had it.
_JFIF_APP0 = b"\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00"
_TRACE_NUM_RE = re.compile(r'(\w+)="([^"]*)"')


def _rect_area(r):
    """A rectangle's area; Rect.get_area() is gone from PyMuPDF 1.26."""
    return max(0.0, r.x1 - r.x0) * max(0.0, r.y1 - r.y0)


def _trace_attrs(tag):
    return dict(_TRACE_NUM_RE.findall(tag))


def _trace_matrix(attrs):
    import fitz
    try:
        a, b, c, d, e, f = (float(v) for v in attrs.get("transform", "").split())
        return fitz.Matrix(a, b, c, d, e, f)
    except ValueError:
        return fitz.Matrix(1, 0, 0, 1, 0, 0)


def _visible_parts(page):
    """[(transform, visible Rect, opacity)] for every picture MuPDF draws on the page, in
    drawing order: the picture's rectangle cut down to the clips it is drawn under, and
    the opacity it is drawn with (its own, times its groups'). MuPDF's trace device writes
    every clip, pop, group and picture of the page as XML lines; the stacks are kept here.
    A clip by glyphs or by a mask only keeps the box it is inside of."""
    import fitz
    try:
        from pymupdf import mupdf
        buf = mupdf.FzBuffer(8192)
        out = mupdf.FzOutput(buf)
        dev = mupdf.fz_new_trace_device(out)
        try:
            mupdf.fz_run_page(page.this, dev, mupdf.FzMatrix(), mupdf.FzCookie())
        finally:
            mupdf.fz_close_device(dev)
            out.fz_close_output()
        xml = bytes(mupdf.fz_buffer_extract(buf)).decode("utf-8", "replace")
    except Exception:  # noqa: BLE001 - no trace device in this PyMuPDF: whole pictures, as before
        return None
    big = fitz.Rect(-1e7, -1e7, 1e7, 1e7)
    stack, parts, opacity = [big], [], [1.0]
    path, path_m, stroke, hidden = None, None, False, 0
    for line in xml.splitlines():
        s = line.strip()
        if not s.startswith("<"):
            continue
        if s.startswith(("<clip_path", "<clip_stroke_path")):
            path, path_m, stroke = [], _trace_matrix(_trace_attrs(s)), s.startswith("<clip_stroke")
            if s.endswith("/>"):
                stack.append(stack[-1])
                path = None
        elif path is not None and s.startswith(("<moveto", "<lineto", "<curveto")):
            a = _trace_attrs(s)
            for xk, yk in (("x", "y"), ("x1", "y1"), ("x2", "y2"), ("x3", "y3")):
                if xk in a and yk in a:
                    path.append(fitz.Point(float(a[xk]), float(a[yk])) * path_m)
        elif s.startswith(("</clip_path", "</clip_stroke_path")):
            box = stack[-1]
            if path and not stroke:
                box = fitz.Rect(stack[-1]) & fitz.Rect(min(p.x for p in path), min(p.y for p in path),
                                                       max(p.x for p in path), max(p.y for p in path))
            stack.append(box)
            path = None
        elif s.startswith(("<clip_text", "<clip_stroke_text", "<clip_image_mask")):
            stack.append(stack[-1])
        elif s.startswith("<mask"):
            # a soft mask: its content draws the mask, not the page, and it clips what follows
            stack.append(stack[-1])
            hidden += 1
        elif s.startswith("</mask"):
            hidden = max(0, hidden - 1)
        elif s.startswith("<group"):
            try:
                opacity.append(opacity[-1] * float(_trace_attrs(s).get("alpha", "1")))
            except ValueError:
                opacity.append(opacity[-1])
        elif s.startswith("</group"):
            if len(opacity) > 1:
                opacity.pop()
        elif s.startswith("<tile"):
            hidden += 1
        elif s.startswith("</tile"):
            hidden = max(0, hidden - 1)
        elif s.startswith("<pop_clip"):
            if len(stack) > 1:
                stack.pop()
        elif s.startswith("<fill_image ") and not hidden:
            a = _trace_attrs(s)
            m = _trace_matrix(a)
            whole = fitz.Rect(0, 0, 1, 1).transform(m)
            try:
                alpha = opacity[-1] * float(a.get("alpha", "1"))
            except ValueError:
                alpha = opacity[-1]
            parts.append(((m.a, m.b, m.c, m.d, m.e, m.f), fitz.Rect(whole) & stack[-1], alpha))
    return parts


def _with_jfif(data):
    """The same JPEG with a JFIF header, which python-docx needs: no pixel is touched."""
    if data[:2] != b"\xff\xd8" or data[2:4] in (b"\xff\xe0", b"\xff\xe1"):
        return data
    return data[:2] + _JFIF_APP0 + data[2:]


def _word_picture(im, crop=None):
    """(bytes, ext) Word can show for an image block, or None. A JPEG in RGB or grey keeps
    its own bytes; a soft mask, CMYK, JPEG 2000 or a crop goes through a pixmap."""
    import fitz
    data, ext = im.get("image"), (im.get("ext") or "").lower()
    mask, ncomp = im.get("mask"), im.get("colorspace") or 3
    if not data:
        return None
    if not mask and crop is None:
        if ext in ("jpeg", "jpg") and ncomp in (1, 3):
            return _with_jfif(data), "jpeg"
        if ext in ("png", "gif", "bmp"):
            return data, ext
    try:
        from PIL import Image
        if not mask and ext in ("jpeg", "jpg") and ncomp in (1, 3):
            # a cut JPEG is saved with its own quantisation tables: no softer, no larger
            from PIL import JpegImagePlugin
            src = Image.open(io.BytesIO(data))
            out = io.BytesIO()
            src.crop(crop).save(out, "JPEG", qtables=src.quantization,
                                subsampling=JpegImagePlugin.get_sampling(src))
            return out.getvalue(), "jpeg"
        pix = fitz.Pixmap(data)
        if pix.colorspace is None or pix.colorspace.n not in (1, 3) or pix.alpha:
            pix = fitz.Pixmap(fitz.csRGB, pix) if pix.colorspace and pix.colorspace.n not in (1, 3) else pix
        mode = {1: "L", 3: "RGB"}.get(pix.n - pix.alpha, "RGB")
        if pix.alpha:
            pic = Image.frombytes(mode + "A" if mode == "RGB" else "LA", (pix.width, pix.height), pix.samples)
        else:
            pic = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
        if mask:
            mpix = fitz.Pixmap(mask)
            if mpix.n - mpix.alpha != 1:
                mpix = fitz.Pixmap(fitz.csGRAY, mpix)
            alpha = Image.frombytes("L", (mpix.width, mpix.height), mpix.samples[: mpix.width * mpix.height]
                                    if not mpix.alpha else bytes(mpix.samples)[::2])
            if alpha.size != pic.size:
                alpha = alpha.resize(pic.size)
            pic = pic.convert("RGBA")
            pic.putalpha(alpha)
        if crop is not None:
            pic = pic.crop(crop)
        out = io.BytesIO()
        if pic.mode in ("RGBA", "LA") or ext not in ("jpeg", "jpg", "jpx"):
            pic.save(out, "PNG", optimize=False)
            return out.getvalue(), "png"
        pic.save(out, "JPEG", quality=90)
        return out.getvalue(), "jpeg"
    except Exception:  # noqa: BLE001 - a picture no decoder here reads (JBIG2 and the like)
        return None


def _word_pictures(page, images):
    """The page's pictures cut to what the page shows of them, in bytes Word can show. A
    picture the page hides completely is left out, and so is one drawn faint (under 60 %
    opacity): Chrome draws a card's shadow as a grey picture at 35 % in a clip with a hole,
    and Word showed it as a black box. One Word cannot take is drawn from the page
    instead of disappearing."""
    import fitz
    parts = {}
    for tf, rect, alpha in _visible_parts(page) or ():
        parts.setdefault(tuple(round(v, 2) for v in tf), []).append((rect, alpha))
    kept = []
    for im in images:
        b = im["block"]
        crop = None
        whole = fitz.Rect(im["x0"], im["y0"], im["x1"], im["y1"])
        seen = parts.get(tuple(round(v, 2) for v in b.get("transform") or ()), [])
        vis, alpha = seen.pop(0) if seen else (whole, 1.0)
        if vis.is_empty or vis.width < 12 or vis.height < 12 or alpha < 0.6:
            continue
        if _rect_area(vis) < 0.97 * _rect_area(whole) and b.get("transform") and b.get("width") and b.get("height"):
            inv = ~fitz.Matrix(*b["transform"])
            corners = [p * inv for p in (vis.tl, vis.tr, vis.bl, vis.br)]
            us = [min(1.0, max(0.0, p.x)) for p in corners]
            vs = [min(1.0, max(0.0, p.y)) for p in corners]
            box = (int(min(us) * b["width"]), int(min(vs) * b["height"]),
                   int(round(max(us) * b["width"])), int(round(max(vs) * b["height"])))
            if box[2] - box[0] >= 8 and box[3] - box[1] >= 8:
                crop = box
        pic = _word_picture(b, crop)
        if pic is None and crop is not None:
            pic = _word_picture(b)
            vis = whole if pic else vis
        if pic is None:
            try:
                pix = page.get_pixmap(clip=vis, dpi=150, alpha=False)
                pic = (pix.tobytes("png"), "png")
            except Exception:  # noqa: BLE001
                continue
        kept.append({"x0": vis.x0, "x1": vis.x1, "y0": vis.y0, "y1": vis.y1, "image": pic[0], "ext": pic[1]})
    return kept


def _page_drawings(page):
    """page.get_drawings(), read once per page: the vector pictures and the list marks
    both look at every path, and a page of a few hundred paths took a second to read."""
    cached = getattr(page, "_confileo_drawings", None)
    if cached is None:
        try:
            cached = page.get_drawings()
        except Exception:  # noqa: BLE001 - no paths to look at then
            cached = []
        try:
            page._confileo_drawings = cached
        except Exception:  # noqa: BLE001 - a page object that takes no attributes
            pass
    return cached


def _vector_pictures(page, frags, images):
    """Charts, diagrams and logos the PDF draws as vector paths, as pictures in the flow
    (19 Oct 2026): only pictures stored as images reached Word, and a report's bar chart
    was simply gone. Paths are grouped by nearness, groups standing side by side (a bar
    chart's bars) join, and a group becomes a picture of the page there only when it is a
    real shape with little text: not a frame, a ruled table, a form or a box of text.
    Text drawn on a filled shape it nearly fills (a chip, a banner, a callout), heading-
    size text and more than a few words keep the group out, so text stays editable;
    labels beside a chart's marks go with its picture.
    Returns (pictures, frags left, images left)."""
    import fitz
    drawings = _page_drawings(page)
    if not drawings or len(drawings) > 3000:
        return [], frags, images
    pw, ph = page.rect.width or 1.0, page.rect.height or 1.0
    shapes = []
    for dr in drawings:
        r = dr.get("rect")
        if r is None or (r.width <= 0 and r.height <= 0):
            continue
        if r.width * r.height > 0.6 * pw * ph:      # a page background or a frame round the page
            continue
        shapes.append((fitz.Rect(r), dr))
    if not shapes:
        return [], frags, images

    def union_groups(rects, near):
        """Indexes of rects (x0, y0, x1, y1) grouped by `near(a, b)`, over a coarse grid."""
        parent = list(range(len(rects)))

        def find(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]
                i = parent[i]
            return i
        grid, cell = {}, 64.0
        for i, r in enumerate(rects):
            for gx in range(int((r[0] - 48) // cell), int((r[2] + 48) // cell) + 1):
                for gy in range(int((r[1] - 48) // cell), int((r[3] + 48) // cell) + 1):
                    for j in grid.get((gx, gy), ()):
                        if find(i) != find(j) and near(r, rects[j]):
                            parent[find(i)] = find(j)
                    grid.setdefault((gx, gy), []).append(i)
        out = {}
        for i in range(len(rects)):
            out.setdefault(find(i), []).append(i)
        return list(out.values())

    def touching(a, b, gap=12.0):
        return a[0] - gap <= b[2] and b[0] <= a[2] + gap and a[1] - gap <= b[3] and b[1] <= a[3] + gap

    def side_by_side(a, b):
        # one row of marks (bars, swatches, icons): mostly the same height band, close apart
        overlap = min(a[3], b[3]) - max(a[1], b[1])
        return overlap >= 0.5 * min(a[3] - a[1], b[3] - b[1]) and max(a[0], b[0]) - min(a[2], b[2]) <= 40

    def bounds(rs):
        return (min(r[0] for r in rs), min(r[1] for r in rs), max(r[2] for r in rs), max(r[3] for r in rs))

    plain = [(r.x0, r.y0, r.x1, r.y1) for r, _ in shapes]
    groups = union_groups(plain, touching)
    boxes = [bounds([plain[i] for i in g]) for g in groups]
    merged = []
    for gg in union_groups(boxes, side_by_side):
        merged.append([shapes[i] for k in gg for i in groups[k]])

    def straight(dr):
        for it in dr.get("items", ()):
            if it[0] in ("re", "qu"):
                continue
            if it[0] == "l" and (abs(it[1].x - it[2].x) < 0.5 or abs(it[1].y - it[2].y) < 0.5):
                continue
            return False
        return True

    def line_box(f):
        return fitz.Rect(f["x0"], f["y0"], f["x1"], f["y1"])

    sizes = sorted(c[4] for f in frags for c in f["chars"] if c[4] and c[0].strip())
    body = sizes[len(sizes) // 2] if sizes else 11.0
    # the text column: a certificate's corner ornaments and frame stand outside it or touch
    # the page's edge, and stacked into the flow as pictures they read as clutter
    col = (min(f["x0"] for f in frags), max(f["x1"] for f in frags)) if frags else None
    edge_x, edge_y = 0.02 * pw, 0.02 * ph
    pictures = []
    for members in merged:
        box = fitz.Rect(members[0][0])
        for r, _ in members[1:]:
            box |= r
        if box.width < 36 or box.height < 24:
            continue
        if box.x0 <= edge_x or box.y0 <= edge_y or box.x1 >= pw - edge_x or box.y1 >= ph - edge_y:
            continue                                  # a frame, a bleed, a page-edge decoration
        if col and (box.x1 <= col[0] or box.x0 >= col[1]):
            continue                                  # in the margin beside the text
        if not any(min(r.width, r.height) >= 6 for r, _ in members):
            continue                                  # rules, underlines, a grid of hairlines
        if all(straight(dr) for _, dr in members) and all(dr.get("fill") is None for _, dr in members):
            continue                                  # frames and empty boxes
        inside = [f for f in frags if box.contains(line_box(f).tl + (line_box(f).br - line_box(f).tl) * 0.5)]
        if inside:
            if all(straight(dr) for _, dr in members):
                continue                              # a table, a form, a shaded or boxed paragraph
            if sum(len("".join(c[0] for c in f["chars"]).split()) for f in inside) > 25:
                continue
            if any(c[4] >= max(14.0, 1.3 * body) for f in inside for c in f["chars"] if c[0].strip()):
                continue                              # a heading on a decoration stays a heading
            boxed = False
            for r, dr in members:
                if dr.get("fill") is None or _rect_area(r) <= 0:
                    continue
                on = sum(_rect_area(line_box(f)) for f in inside
                         if r.contains(line_box(f).tl + (line_box(f).br - line_box(f).tl) * 0.5))
                if on >= _rect_area(r) / 8.0:
                    boxed = True                      # a chip, a badge, a banner: text on its shape
                    break
            if boxed:
                continue
        pictures.append((box, inside))
    if not pictures:
        return [], frags, images
    out, gone = [], set()
    for box, inside in pictures:
        # a photo the drawing sits on goes into the same picture, which is then a JPEG
        photo = False
        for im in images:
            r = fitz.Rect(im["x0"], im["y0"], im["x1"], im["y1"])
            if _rect_area(r & box) > 0.3 * min(_rect_area(r), _rect_area(box)):
                box |= r
                gone.add(id(im))
                photo = photo or im.get("ext") == "jpeg"
        clip = fitz.Rect(box.x0 - 1.5, box.y0 - 1.5, box.x1 + 1.5, box.y1 + 1.5) & page.rect
        try:
            pix = page.get_pixmap(clip=clip, dpi=200, alpha=False)
            data, ext = (pix.tobytes("jpg", jpg_quality=88), "jpeg") if photo else (pix.tobytes("png"), "png")
        except Exception:  # noqa: BLE001
            continue
        gone.update(id(f) for f in inside)
        out.append({"x0": clip.x0, "x1": clip.x1, "y0": clip.y0, "y1": clip.y1, "image": data, "ext": ext})
    return out, [f for f in frags if id(f) not in gone], [im for im in images if id(im) not in gone]


def _layout_page(page, include_images, ruled_tables):
    """One page as columns of rows / images / tables in reading order, plus what
    the Word section needs: the page size and the box the content occupies."""
    import fitz
    pw = page.rect.width or 1.0
    info = {"width": page.rect.width, "height": page.rect.height, "flow": [], "body": 11.0,
            "rtl": False, "box": None}
    flags = fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE
    if include_images:
        flags |= fitz.TEXT_PRESERVE_IMAGES
    raw = page.get_text("rawdict", flags=flags)

    frags, images, blocks = [], [], []
    pictured = 0.0
    for b in raw.get("blocks", []):
        x0, y0, x1, y1 = b["bbox"]
        if b.get("type") == 1:
            pictured += max(0.0, x1 - x0) * max(0.0, y1 - y0)
            if include_images and (x1 - x0) >= 12 and (y1 - y0) >= 12 and b.get("image"):
                images.append({"x0": x0, "x1": x1, "y0": y0, "y1": y1, "block": b})
            continue
        got = []
        for ln in b.get("lines", []):
            chars = _line_chars(ln)
            if any(c[0] not in "  " for c in chars):
                frags.append({"chars": chars, "y0": ln["bbox"][1], "y1": ln["bbox"][3],
                              "x0": ln["bbox"][0], "x1": ln["bbox"][2], "block": len(blocks)})
                got.append((ln["bbox"][0], ln["bbox"][2], ln["bbox"][1], ln["bbox"][3]))
        if got:
            blocks.append({"x0": x0, "x1": x1, "y0": y0, "y1": y1, "lines": got})
    block_of = list(blocks)
    if images:
        images = _word_pictures(page, images)
    if include_images:
        drawn, frags, images = _vector_pictures(page, frags, images)
        images += drawn
    # a scanned page inside a PDF that has text elsewhere: no text, and pictures
    # over half of it (the whole-scan case never gets here: main() OCRs it)
    info["scan"] = not frags and pictured >= 0.5 * pw * (page.rect.height or 1.0)
    if not frags and not images:
        return info

    boxes = [(f["x0"], f["y0"], f["x1"], f["y1"]) for f in frags] + \
            [(im["x0"], im["y0"], im["x1"], im["y1"]) for im in images]
    info["box"] = (min(b[0] for b in boxes), min(b[1] for b in boxes),
                   max(b[2] for b in boxes), max(b[3] for b in boxes))

    all_text = "".join(c[0] for f in frags for c in f["chars"])
    ar_n = sum(1 for c in all_text if _is_arabic_char(c))
    lat_n = sum(1 for c in all_text if c.isascii() and c.isalpha())
    page_rtl = ar_n > lat_n
    info["rtl"] = page_rtl
    page_tables = []
    if ruled_tables is not None and frags:
        try:
            page_tables, frags = ruled_tables(page, frags, page_rtl)
        except Exception:  # noqa: BLE001 - a table we cannot read stays text
            page_tables = []
        if page_tables:
            page_tables = _merge_stacked_tables(page_tables)
            # the glyphs left outside the tables decide the columns
            keep = {f["block"] for f in frags}
            blocks = [b for i, b in enumerate(blocks) if i in keep]
    split = _page_columns(blocks, pw)
    cuts, lone = None, ()
    if split is None:
        spanning = [{"x0": im["x0"], "x1": im["x1"], "y0": im["y0"], "y1": im["y1"]} for im in images] + \
                   [{"x0": t["x0"], "x1": t["x1"], "y0": t["y0"], "y1": t["y1"]} for t in page_tables]
        banded = _page_bands(blocks, spanning, pw)
        if banded is not None:
            split, cuts, modes, lone = banded

    def side(x0, x1):
        right = (x0 + x1) / 2.0 >= split
        return (0 if right else 1) if page_rtl else (1 if right else 0)

    if split is None:
        slots = [(0.0, pw)]

        def slot_of(x0, x1, y0, y1, full=False):
            return 0
    elif cuts is None:
        # two columns from top to bottom
        slots = [(split, pw), (0.0, split)] if page_rtl else [(0.0, split), (split, pw)]

        def slot_of(x0, x1, y0, y1, full=False):
            return side(x0, x1)
    else:
        # top to bottom: part k (its two columns, or one flow when its sides
        # are not reading text), then full-width part k
        slots, first = [], []
        for k in range(len(cuts) + 1):
            first.append(len(slots))
            if modes[k]:
                slots += [(split, pw), (0.0, split)] if page_rtl else [(0.0, split), (split, pw)]
            else:
                slots.append((0.0, pw))
            if k < len(cuts):
                slots.append((0.0, pw))

        def slot_of(x0, x1, y0, y1, full=False):
            cy = (y0 + y1) / 2.0
            if full or (x0 < split - 2 and x1 > split + 2):
                k = min(range(len(cuts)), key=lambda i: 0.0 if cuts[i][0] - 2 <= cy <= cuts[i][1] + 2
                        else min(abs(cy - cuts[i][0]), abs(cy - cuts[i][1])))
                return first[k] + (2 if modes[k] else 1)
            k = sum(1 for c in cuts if c[1] <= cy)
            return first[k] + (side(x0, x1) if modes[k] else 0)

    # every line goes where its block goes: a crossing paragraph whose last
    # line is short stays in the full-width part
    def frag_slot(f):
        b = block_of[f["block"]] if 0 <= f["block"] < len(block_of) else f
        return slot_of(b["x0"], b["x1"], b["y0"], b["y1"], id(b) in lone)

    # rows per column, then paragraphs; images and tables join the flow by position
    for si, (cx0, cx1) in enumerate(slots):
        cf = [f for f in frags if frag_slot(f) == si]
        rows = [r for r in (_finish_row(r, page_rtl) for r in _merge_rows(cf)) if r]
        if cuts is None:
            slot_tables = page_tables if si == 0 else []
        else:
            slot_tables = [t for t in page_tables if slot_of(t["x0"], t["x1"], t["y0"], t["y1"]) == si]
        if slot_tables:
            rows = _absorb_rows_into_tables(rows, slot_tables, page_rtl)
        col_items = [{"kind": "row", "y0": r["y0"], "row": r} for r in rows]
        col_items += [{"kind": "image", "y0": im["y0"], "image": im}
                      for im in images if slot_of(im["x0"], im["x1"], im["y0"], im["y1"]) == si]
        col_items += [{"kind": "table", "y0": t["y0"], "table": t} for t in slot_tables]
        if not col_items and len(slots) > 1:
            continue
        col_items.sort(key=lambda it: it["y0"])
        info["flow"].append((col_items, cx0, cx1))
    _mark_vector_bullets(page, info["flow"])
    _mark_links(page, info["flow"])

    sizes = {}
    for col_items, _, _ in info["flow"]:
        for it in col_items:
            if it["kind"] == "row":
                k = round(it["row"]["size"] * 2) / 2
                sizes[k] = sizes.get(k, 0) + sum(len(t["text"]) for t in it["row"]["tokens"])
    info["body"] = max(sizes, key=sizes.get) if sizes else 11.0
    return info


def _mark_vector_bullets(page, flow):
    """Chrome draws a list's disc or square as a small filled path, not a glyph,
    so the bullet vanished and two items could run into one paragraph ("Private
    health insurance … Hybrid working …"). A filled mark of 1.2-9 pt standing on
    a row's line, just before its start (left of a Latin row, right of an Arabic
    one), makes that row a list item."""
    drawings = _page_drawings(page)
    if len(drawings) > 3000:
        return
    marks = []
    for dr in drawings:
        r = dr.get("rect")
        if r is None or dr.get("fill") is None:
            continue
        w, h = r.width, r.height
        if 1.2 <= w <= 9 and 1.2 <= h <= 9 and 0.6 <= w / max(h, 0.01) <= 1.6:
            marks.append((r.x0, r.y0, r.x1, r.y1))
    if not marks:
        return
    for col_items, _, _ in flow:
        for it in col_items:
            if it["kind"] != "row":
                continue
            row = it["row"]
            reach = 2.5 * max(row["size"], 6.0)
            for x0, y0, x1, y1 in marks:
                if not row["y0"] - 1 <= (y0 + y1) / 2.0 <= row["y1"] + 1:
                    continue
                if row["rtl"]:
                    hit = row["x1"] - 1 <= x0 <= row["x1"] + reach
                else:
                    hit = row["x0"] - reach <= x1 <= row["x0"] + 1
                if hit:
                    row["bullet"] = True
                    break


_LINK_SCHEMES = ("http://", "https://", "mailto:", "tel:")


def _mark_links(page, flow):
    """Words under a PDF link to a web page or an address carry its target."""
    try:
        links = [(l["from"], l["uri"]) for l in page.get_links()
                 if l.get("from") is not None and (l.get("uri") or "").lower().startswith(_LINK_SCHEMES)]
    except Exception:  # noqa: BLE001 - no links then
        return
    if not links:
        return
    for col_items, _, _ in flow:
        for it in col_items:
            if it["kind"] != "row":
                continue
            row = it["row"]
            cy = (row["y0"] + row["y1"]) / 2.0
            for t in row["tokens"]:
                cx = (t["x0"] + t["x1"]) / 2.0
                for rect, uri in links:
                    if rect.x0 - 1 <= cx <= rect.x1 + 1 and rect.y0 - 2 <= cy <= rect.y1 + 2:
                        t["link"] = uri
                        break


def _token_look(t):
    """(bold, italic, colour) most of a word's letters are printed in, or None."""
    n = b = i = 0
    colors = {}
    for c in t.get("chars") or ():
        sp = c[3] if isinstance(c[3], dict) else None
        if sp is None or not c[0].strip():
            continue
        n += 1
        name = (sp.get("font") or "").lower()
        flags = int(sp.get("flags", 0))
        if flags & 16 or "bold" in name or "black" in name or "heavy" in name:
            b += 1
        if flags & 2 or "italic" in name or "oblique" in name:
            i += 1
        col = int(sp.get("color", 0)) & 0xFFFFFF
        colors[col] = colors.get(col, 0) + 1
    if not n:
        return None
    return (b * 2 > n, i * 2 > n, max(colors, key=colors.get))


def _add_hyperlink(p, uri, text, style):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), p.part.relate_to(uri, RT.HYPERLINK, is_external=True))
    run = _style_run(p.add_run(text), **dict(style, color=style.get("color") or "0563C1"))
    run.font.underline = True
    p._p.remove(run._r)
    link.append(run._r)
    p._p.append(link)


def _emit_runs(p, lines, text, heading, base, **style):
    """The paragraph's text as runs: a word printed bold, italic or in its own
    colour keeps that look, and words under a PDF web link become a Word
    hyperlink. Until 28 Sept 2026 a paragraph was one run in one look — the
    bold "€78,500" and the italic job title of a letter came back plain, and
    its handbook link as plain text."""
    n = len(text)
    looks, links = [None] * n, [None] * n
    cur = 0
    for ln in lines:
        for t in ln["tokens"]:
            w = "".join(_PUA_BULLETS.get(ch, ch) for ch in _xml_clean(_normalize_presentation_forms(t["text"])))
            if not w:
                continue
            k = text.find(w, cur)          # _tidy only takes spaces out: every word is there, in order
            if k < 0:
                continue
            look, link = _token_look(t), t.get("link")
            for j in range(k, k + len(w)):
                looks[j], links[j] = look, link
            cur = k + len(w)
    # a space between two words takes their look (and link) when they share it,
    # else the look of the word before it (one run per word, not per space)
    prev = [None] * n
    last = None
    for j in range(n):
        if looks[j] is not None or links[j] is not None:
            last = j
        prev[j] = last
    nxt, last = [None] * n, None
    for j in range(n - 1, -1, -1):
        if looks[j] is not None or links[j] is not None:
            last = j
        nxt[j] = last
    for j in range(n):
        if looks[j] is None and links[j] is None and prev[j] is not None and nxt[j] is not None:
            a, b = prev[j], nxt[j]
            looks[j] = looks[a]
            if links[a] == links[b]:
                links[j] = links[a]
    j = 0
    while j < n:
        k = j
        while k < n and looks[k] == looks[j] and links[k] == links[j]:
            k += 1
        bold, italic, color = looks[j] or base
        run_style = dict(style, bold=bold or heading, italic=italic,
                         color=("%06X" % color) if color else None)
        if links[j]:
            _add_hyperlink(p, links[j], text[j:k], run_style)
        else:
            _style_run(p.add_run(text[j:k]), **run_style)
        j = k


def _strip_list_mark(lines):
    """True, and the bullet glyph removed from the first row, when the paragraph
    is a list item (a glyph bullet as its own first word, or a vector one)."""
    first = lines[0]
    toks = first["tokens"]
    if toks and toks[0]["text"] and all(ch in _LIST_MARKS for ch in toks[0]["text"]):
        first["tokens"] = toks[1:]
        return True
    if toks and toks[0]["text"][:1] in _PUA_BULLETS:
        toks[0] = dict(toks[0], text=toks[0]["text"][1:].lstrip())
        return True
    return bool(first.get("bullet"))


_ANY_DIGIT_RE = re.compile(r"[0-9\u0660-\u0669\u06f0-\u06f9]+")
_ASCII_NUM_RE = re.compile(r"([0-9]+)")
# a page label in the site's eleven languages; a counter that does not start at 1
# on the first page is a page number only beside one (or alone: "- 3 -")
_PAGE_WORD_RE = re.compile(r"(?i)\b(page|pages|pg|p|seite|página|pagina|strona|str|sayfa|halaman|hal)\b"
                           r"|صفحة|صفحه|पृष्ठ")


def _row_text(row):
    return _tidy(" ".join(w for w in (_xml_clean(_normalize_presentation_forms(t["text"]))
                                      for t in row["tokens"]) if w))


def _number_roles(occurrences):
    """For one running row seen on several pages as [(page_no, text)]: per number
    in the text, "page" (it counts the pages), "pages" (the page count, beside a
    page number) or None (the same on every page); plus the offset of the page
    counter. None when a number changes in any other way (a date, a chapter)."""
    nums = [[int(n) for n in _ASCII_NUM_RE.findall(t)] for _, t in occurrences]
    if len({len(n) for n in nums}) != 1:
        return None
    if any(re.search(r"[\u0660-\u0669\u06f0-\u06f9]", t) for _, t in occurrences) \
            and len({t for _, t in occurrences}) > 1:
        return None      # Arabic-Indic page numbers stay text: a PAGE field prints 1, 2, 3
    roles, offset = [], None
    for j in range(len(nums[0])):
        vals = [n[j] for n in nums]
        if len(set(vals)) == 1:
            roles.append(None)
            continue
        offs = {v - (pno + 1) for v, (pno, _) in zip(vals, occurrences)}
        if len(offs) != 1 or (offset is not None and offs != {offset}):
            return None
        offset = offs.pop()
        roles.append("page")
    if offset:
        # "Printed 2026-09-20" on page 1, "…-21" on page 2 counts along too
        t = occurrences[0][1]
        if not _PAGE_WORD_RE.search(t) and re.sub(r"[\W\d_]", "", t):
            return None
    return roles, (offset or 0)


def _running_items(pages):
    """Rows and pictures the PDF repeats at the same place at the top or bottom of
    most pages — a running head, a letterhead logo, "Page 3 of 12" — taken out of
    the body. Until 28 Sept 2026 they came back as body text once per PDF page, in
    the middle of the text flow («شركة النيل للمقاولات — تقرير داخلي سرّي» and
    «صفحة 2 من 5» between the two halves of a sentence); they now fill the Word
    header and footer, the page number and page count as PAGE / NUMPAGES fields.
    Returns None or {"header": [...], "footer": [...], "first": bool} where "first"
    says the first page carries none of them (a title page)."""
    import hashlib
    live = [(i, pg) for i, pg in enumerate(pages) if pg["flow"]]
    if len(live) < 2:
        return None
    groups = {}
    for pno, pg in live:
        h = pg["height"] or 1.0
        for col_items, _, _ in pg["flow"]:
            for it in col_items:
                if it["kind"] == "row":
                    r = it["row"]
                    y0, y1 = r["y0"], r["y1"]
                elif it["kind"] == "image":
                    r = it["image"]
                    y0, y1 = r["y0"], r["y1"]
                else:
                    continue
                # the page's top and bottom tenth, where Word puts its header and
                # footer; the body of a page starts below 0.1 of its height
                if y1 <= 0.10 * h:
                    band = "header"
                elif y0 >= 0.90 * h:
                    band = "footer"
                else:
                    continue
                if it["kind"] == "row":
                    text = _row_text(r)
                    if not text:
                        continue
                    sig = ("row", _ANY_DIGIT_RE.sub("#", text))
                else:
                    sig = ("image", hashlib.md5(r["image"]).hexdigest(), round(r["x1"] - r["x0"]),
                           round(r["y1"] - r["y0"]))
                    text = None
                groups.setdefault((band,) + sig, []).append(
                    {"page": pno, "item": it, "pg": pg, "text": text, "y0": y0, "y1": y1,
                     "x0": r["x0"], "x1": r["x1"]})

    need = max(2, (len(live) + 1) // 2)
    first_live = live[0][0]
    found = {"header": [], "footer": [], "first": False}
    drop = set()
    for key, occ in groups.items():
        if len({o["page"] for o in occ}) < need:
            continue
        # the same height on every page (±3 pt around the most common one)
        ys = sorted(round(o["y0"]) for o in occ)
        mid = ys[len(ys) // 2]
        seen, keep = set(), []
        for o in sorted(occ, key=lambda o: o["page"]):
            if abs(o["y0"] - mid) <= 3 and o["page"] not in seen:
                seen.add(o["page"])
                keep.append(o)
        if len(keep) < need:
            continue
        spec = {"kind": key[1], "y0": min(o["y0"] for o in keep), "y1": max(o["y1"] for o in keep),
                "pages": {o["page"] for o in keep}, "sample": keep[0]}
        if key[1] == "row":
            roles = _number_roles([(o["page"], o["text"]) for o in keep])
            if roles is None:
                continue
            spec["roles"], spec["offset"] = roles
            if "page" in spec["roles"]:
                # a constant number beside the page number that equals the page
                # count ("of 5") becomes NUMPAGES
                total = len(pages) + spec["offset"]
                nums = [int(n) for n in _ASCII_NUM_RE.findall(keep[0]["text"])]
                spec["roles"] = [("pages" if r is None and n == total else r)
                                 for r, n in zip(spec["roles"], nums)]
        found[key[0]].append(spec)
        drop.update(id(o["item"]) for o in keep)
    if not drop:
        return None
    offsets = {s["offset"] for b in ("header", "footer") for s in found[b] if "page" in (s.get("roles") or ())}
    if len(offsets) > 1:
        # a header and a footer that count differently: keep the numbers as text
        for b in ("header", "footer"):
            for s in list(found[b]):
                if "page" in (s.get("roles") or ()):
                    found[b].remove(s)
                    for o in groups.get((b, "row", _ANY_DIGIT_RE.sub("#", s["sample"]["text"])), []):
                        drop.discard(id(o["item"]))
    found["offset"] = offsets.pop() if len(offsets) == 1 else 0
    for b in ("header", "footer"):
        found[b].sort(key=lambda s: s["y0"])
    found["first"] = all(first_live not in s["pages"] for b in ("header", "footer") for s in found[b])
    if not (found["header"] or found["footer"]):
        return None

    # out of the body, and the body's own box decides the margins
    for pno, pg in live:
        box = None
        for col_items, _, _ in pg["flow"]:
            col_items[:] = [it for it in col_items if id(it) not in drop]
            for it in col_items:
                o = it["row"] if it["kind"] == "row" else it.get("image") or it.get("table")
                if o is None:
                    continue
                b = (o["x0"], o["y0"], o["x1"], o["y1"])
                box = b if box is None else (min(box[0], b[0]), min(box[1], b[1]),
                                             max(box[2], b[2]), max(box[3], b[3]))
        if box is not None and pg["box"] is not None:
            # side margins stay the page's own; only the top and bottom move
            box = (pg["box"][0], box[1], pg["box"][2], box[3])
        pg["box"] = box
    return found


def _add_field(p, instr, shown, style):
    """A Word field (PAGE, NUMPAGES) showing `shown` until the reader updates it."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    def part(el):
        r = _style_run(p.add_run(), **style)
        r._r.append(el)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    part(begin)
    code = OxmlElement("w:instrText")
    code.set(qn("xml:space"), "preserve")
    code.text = " %s " % instr
    part(code)
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    part(sep)
    _style_run(p.add_run(shown), **style)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    part(end)


def _fill_running(part, specs, page_w, latin, arabic):
    """Writes running rows / pictures into a header or footer part."""
    from docx.shared import Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    part.is_linked_to_previous = False
    first = True
    for spec in specs:
        p = part.paragraphs[0] if first and part.paragraphs else part.add_paragraph()
        first = False
        o = spec["sample"]
        cx = (o["x0"] + o["x1"]) / 2.0
        centred = abs(cx - page_w / 2.0) < 0.06 * page_w
        right = cx > page_w / 2.0
        if spec["kind"] == "image":
            im = o["item"]["image"]
            try:
                p.add_run().add_picture(io.BytesIO(im["image"]), width=Inches(max(0.3, (im["x1"] - im["x0"]) / 72.0)))
            except Exception:  # noqa: BLE001 - an image Word cannot take is left out
                continue
            p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if centred else
                           WD_ALIGN_PARAGRAPH.RIGHT if right else WD_ALIGN_PARAGRAPH.LEFT)
            continue
        row = o["item"]["row"]
        rtl = row["rtl"]
        lat, ar = _row_fonts([row])
        style = {"size": round(row["size"] * 2) / 2 if row["size"] > 0 else None, "bold": row["bold"],
                 "color": ("%06X" % (row["color"] & 0xFFFFFF)) if row["color"] else None,
                 "font": lat or latin, "cs_font": ar or arabic, "rtl": rtl}
        pieces = _ASCII_NUM_RE.split(o["text"])
        roles = iter(spec.get("roles") or ())
        for k, piece in enumerate(pieces):
            if k % 2 == 0:
                if piece:
                    _style_run(p.add_run(piece), **style)
                continue
            role = next(roles, None)
            if role == "page":
                _add_field(p, "PAGE", piece, style)
            elif role == "pages":
                _add_field(p, "NUMPAGES", piece, style)
            else:
                _style_run(p.add_run(piece), **style)
        if rtl:
            _set_paragraph_rtl(p)
        if centred:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif rtl:
            # in a bidi paragraph "left" is the start (the right margin), "right" the end
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if right else WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if right else WD_ALIGN_PARAGRAPH.LEFT


def _write_running(section, running, page_w, page_h, latin, arabic):
    """Header / footer of the document's first section (later sections link to it)."""
    from docx.shared import Pt
    if running["header"]:
        _fill_running(section.header, running["header"], page_w, latin, arabic)
        section.header_distance = Pt(max(12.0, min(72.0, running["header"][0]["y0"] - 1)))
    if running["footer"]:
        _fill_running(section.footer, running["footer"], page_w, latin, arabic)
        section.footer_distance = Pt(max(12.0, min(72.0, page_h - running["footer"][-1]["y1"] - 1)))
    if running["first"]:
        # the PDF's first page has none of them: a title page
        section.different_first_page_header_footer = True
    if running.get("offset"):
        _put(section._sectPr, "w:pgNumType", _SECTPR_ORDER, start=str(1 + running["offset"]))


def _section_setup(section, width, height, box):
    """Page size, orientation and margins of the PDF page. Until 28 Sept 2026 every
    result was US Letter with 1.25" side margins (python-docx's template), so an A4
    page's lines wrapped differently and printed on the wrong paper."""
    from docx.shared import Pt
    from docx.enum.section import WD_ORIENT
    section.orientation = WD_ORIENT.LANDSCAPE if width > height else WD_ORIENT.PORTRAIT
    section.page_width, section.page_height = Pt(width), Pt(height)
    if box:
        left, top, right, bottom = box[0], box[1], width - box[2], height - box[3]
    else:
        left = right = top = bottom = 72.0

    def clamp(v, lo, hi):
        return max(lo, min(hi, v))
    # a little narrower than the PDF's own margins: a line that exactly filled the
    # PDF measure must not wrap in Word because of a hair of metric difference
    section.left_margin = Pt(clamp(left - 4, 18, 90))
    section.right_margin = Pt(clamp(right - 4, 18, 90))
    section.top_margin = Pt(clamp(top - 2, 18, 72))
    section.bottom_margin = Pt(clamp(bottom - 2, 18, 72))
    section.header_distance = Pt(clamp(top / 2, 12, 36))
    section.footer_distance = Pt(clamp(bottom / 2, 12, 36))


def _doc_defaults(doc, latin, cs, size):
    """Document-wide fonts and size (w:docDefaults). python-docx's template names
    theme fonts — Cambria for body text — so every run the rebuild left unstyled
    came out in Cambria; the PDF's own dominant faces are used instead."""
    from docx.oxml.ns import qn
    rPr = doc.styles.element.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr"))
    if rPr is None:
        return
    f = _put(rPr, "w:rFonts", _RPR_ORDER)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme"):
        if f.get(qn(a)) is not None:
            del f.attrib[qn(a)]
    f.set(qn("w:ascii"), latin)
    f.set(qn("w:hAnsi"), latin)
    f.set(qn("w:cs"), cs)
    if size:
        v = str(int(round(size * 2)))
        _put(rPr, "w:sz", _RPR_ORDER, val=v)
        _put(rPr, "w:szCs", _RPR_ORDER, val=v)


def _list_style(doc):
    """The template's "List Bullet" paragraph style (a real Word list), or None."""
    try:
        return doc.styles["List Bullet"]
    except KeyError:
        return None


def _paragraphs_of(col_items, span=0.0):
    """Rows → paragraphs (lists of rows); images and tables pass through as dicts.
    `span` is the width of the column the rows stand in."""
    widest = max((it["row"]["x1"] - it["row"]["x0"] for it in col_items if it["kind"] == "row"), default=1.0)
    # a short line is short against the column, even when every line in it is short
    short = 0.6 * max(widest, 0.5 * span)
    paras = []
    prev = None
    for it in col_items:
        if it["kind"] in ("image", "table"):
            paras.append(it)
            prev = None
            continue
        ln = it["row"]
        first = ln["tokens"][0]["text"][:1] if ln["tokens"] else ""
        style_break = (
            first in _BULLETS
            or first in _LIST_MARKS
            or ln.get("bullet")
            or prev is None
            or abs(ln["size"] - prev["size"]) > 0.9
            or ln["bold"] != prev["bold"]
            or (ln["y0"] - prev["y1"]) > max(2.0, 0.45 * max(ln["size"], prev["size"]))
        )
        # A new PyMuPDF block is a new paragraph, except when the previous
        # line ran the full width and stopped mid-sentence: that is a wrapped
        # line the producer happened to file under another block.
        block_break = False
        if not style_break and prev is not None and ln["block"] != prev["block"]:
            prev_text = " ".join(t["text"] for t in prev["tokens"]).rstrip()
            wrapped = (prev["x1"] - prev["x0"]) >= 0.8 * widest and prev_text[-1:] not in ".:!?؟؛"
            block_break = not wrapped
        elif not style_break and prev is not None and (prev["x1"] - prev["x0"]) < short \
                and (ln["x1"] - ln["x0"]) < short:
            # Two short lines of one block are lines the author broke, not a
            # paragraph wrapping: an address, a list of skills, "Name: …" over
            # "Date: …". They were run together into one line («12 High Street
            # London EC1A 1BB United Kingdom», 2 Oct 2026). A short line
            # followed by a long one is still the end of a wrapped heading.
            block_break = True
        if style_break or block_break:
            paras.append([ln])
        else:
            paras[-1].append(ln)
        prev = ln
    return paras


def _flow_paragraphs(flow):
    """[(col_items, cx0, cx1, paragraphs)] for a page's flow. A paragraph that
    runs from the foot of one column to the head of the next is one paragraph:
    the column's last line reaches its right edge with no full stop, and the
    next column goes on in lower case, or (Arabic, which has no case) in the
    same size and weight. It came back as two paragraphs split mid-sentence
    («…The quarterly review» / «covers revenue, costs…», 2 Oct 2026)."""
    out = []
    prev = None                                    # the last paragraph of the slot before
    for col_items, cx0, cx1 in flow:
        paras = _paragraphs_of(col_items, cx1 - cx0)
        if prev is not None and paras and isinstance(paras[0], list) and paras[0] and paras[0][0]["tokens"]:
            last, widest = prev
            nxt = paras[0][0]
            tail = " ".join(t["text"] for t in last[-1]["tokens"]).rstrip()
            head = nxt["tokens"][0]["text"][:1]
            full = (last[-1]["x1"] - last[-1]["x0"]) >= 0.85 * widest
            if tail and tail[-1:] not in ".:!?؟؛" and full and len(last) >= 2 \
                    and abs(nxt["size"] - last[-1]["size"]) < 0.6 and nxt["bold"] == last[-1]["bold"] \
                    and (head.islower() or (nxt["rtl"] and last[-1]["rtl"])) \
                    and not nxt.get("bullet") and head not in _LIST_MARKS and head not in _PUA_BULLETS:
                last.extend(paras.pop(0))
        out.append((col_items, cx0, cx1, paras))
        widest = max((it["row"]["x1"] - it["row"]["x0"] for it in col_items if it["kind"] == "row"), default=1.0)
        prev = (paras[-1], widest) if paras and isinstance(paras[-1], list) else None
    return out


def _para_size(lines):
    """The type size of a paragraph's words, list marks left out."""
    sizes = [c[4] for ln in lines for t in ln["tokens"] for c in (t.get("chars") or ())
             if c[4] and c[0].strip() and c[0] not in _LIST_MARKS]
    return max(sizes) if sizes else max(ln["size"] for ln in lines)


def _outline_worthy(lines, text):
    """A heading short enough for the navigation pane and a table of contents, and
    made of words: a large-type paragraph of a cover page or a pull quote, big step
    numbers («01 02 03 04»), figures («30+ 0 6 100%») and letter-spaced labels
    («C H A P T E R 1», a brochure's kickers) keep their look but get no level."""
    words = text.split()
    spaced = len(words) >= 4 and sum(1 for w in words if len(w) == 1) >= 0.6 * len(words)
    return len(lines) <= 3 and len(text) <= 200 and sum(1 for ch in text if ch.isalpha()) >= 3 and not spaced


def _heading_levels(live, flows):
    """{size: 1 | 2 | 3}: the document's heading sizes, largest first, as Word's
    Heading 1-3. Headings used to be Normal paragraphs in bigger bold type, so the
    rebuilt document had no outline: Word's navigation pane and an automatic table
    of contents found nothing (0 heading paragraphs in every output, 10 Oct 2026)."""
    sizes = set()
    for pg in live:
        if pg.get("ocr"):
            continue
        for _, _, _, paras in flows[id(pg)]:
            for lines in paras:
                if not isinstance(lines, list):
                    continue
                size = _para_size(lines)
                text = " ".join(t["text"] for ln in lines for t in ln["tokens"])
                if size >= pg["body"] * 1.25 and _outline_worthy(lines, text):
                    sizes.add(round(size * 2) / 2)
    return {s: min(3, k + 1) for k, s in enumerate(sorted(sizes, reverse=True))}


def _plain_heading_styles(doc):
    """Heading 1-3 keep what makes them headings (outline level, kept with the next
    paragraph) and lose the template's look (Cambria, blue, its sizes): each run
    carries the PDF's own font, size, weight and colour, and a black heading must not
    turn blue where the run names no colour."""
    from docx.oxml.ns import qn
    for name in ("Heading 1", "Heading 2", "Heading 3"):
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        rpr = st.element.find(qn("w:rPr"))
        if rpr is None:
            continue
        for tag in ("w:rFonts", "w:color", "w:sz", "w:szCs", "w:i", "w:iCs"):
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)


def _emit_document(doc, pages, running=None):
    """Writes the laid-out pages: one Word section per page size, the PDF's
    margins, running heads and footers, and every run formatted in both script
    slots."""
    from collections import Counter
    from docx.shared import Pt, Inches
    from docx.enum.section import WD_SECTION
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    live = [pg for pg in pages if pg["flow"] or pg.get("ocr")]
    lat_all, ar_all, body_all = Counter(), Counter(), Counter()
    for pg in live:
        if not pg.get("ocr"):
            body_all[pg["body"]] += 1
        for col_items, _, _ in pg["flow"]:
            lat, ar = _row_fonts([it["row"] for it in col_items if it["kind"] == "row"])
            if lat:
                lat_all[lat] += 1
            if ar:
                ar_all[ar] += 1
    latin = lat_all.most_common(1)[0][0] if lat_all else "Calibri"
    arabic = ar_all.most_common(1)[0][0] if ar_all else _AR_FALLBACK_FONT
    _doc_defaults(doc, latin, arabic, body_all.most_common(1)[0][0] if body_all else None)
    flows = {id(pg): _flow_paragraphs(pg["flow"]) for pg in live if not pg.get("ocr")}
    levels = _heading_levels(live, flows)
    if levels:
        _plain_heading_styles(doc)

    section = None
    size = None
    breaks = []
    for pg in live:
        this = (round(pg["width"], 1), round(pg["height"], 1))
        if section is None or this != size:
            first_section = section is None
            if first_section:
                section = doc.sections[0]
            else:
                section = doc.add_section(WD_SECTION.NEW_PAGE)
            same = [q["box"] for q in live if q["box"] and (round(q["width"], 1), round(q["height"], 1)) == this]
            box = (min(b[0] for b in same), min(b[1] for b in same), max(b[2] for b in same),
                   max(b[3] for b in same)) if same else None
            _section_setup(section, pg["width"], pg["height"], box)
            if running and first_section:
                _write_running(section, running, pg["width"], pg["height"], latin, arabic)
            size = this
        else:
            # the body ends with its sectPr, so this page's first block will sit here
            breaks.append(len(doc.element.body) - 1)
        text_width = (section.page_width - section.left_margin - section.right_margin) / 914400.0

        if pg.get("ocr"):
            _add_ocr_paragraphs(doc, pg["ocr"])
            continue
        for col_items, cx0, cx1, paras in flows[id(pg)]:
            body = pg["body"]
            for lines in paras:
                if isinstance(lines, dict) and lines.get("kind") == "table":
                    t = lines["table"]
                    _add_word_table(doc, t["rows"], pg["rtl"], t.get("layout"), t.get("col_x"),
                                    t.get("size"), t.get("row_y"), t.get("y1"))
                    continue
                if isinstance(lines, dict):
                    try:
                        im = lines["image"]
                        p = doc.add_paragraph()
                        w_in = min(text_width, max(0.3, (im["x1"] - im["x0"]) / 72.0))
                        p.add_run().add_picture(io.BytesIO(im["image"]), width=Inches(w_in))
                    except Exception:
                        pass
                    continue
                listed = _strip_list_mark(lines)
                words = []
                for ln in lines:
                    words.extend(_xml_clean(_normalize_presentation_forms(t["text"])) for t in ln["tokens"])
                text = _tidy(" ".join(w for w in words if w))
                # any bullet left inside the text prints as a box outside Symbol/Wingdings
                text = "".join(_PUA_BULLETS.get(ch, ch) for ch in text)
                if not text:
                    continue
                # the size of the paragraph's own words: the row's size counted the
                # bullet _strip_list_mark took off, and 10 pt list items under a
                # 12 pt bullet came out 12 pt (10 Oct 2026)
                size_pt = _para_size(lines)
                heading = size_pt >= body * 1.25
                level = levels.get(round(size_pt * 2) / 2) if heading and _outline_worthy(lines, text) else None
                p = doc.add_paragraph(style=_list_style(doc) if listed else ("Heading %d" % level if level else None))
                rtl = lines[0]["rtl"]
                bold = all(ln["bold"] for ln in lines)
                color = lines[0]["color"]
                bx0 = min(ln["x0"] for ln in lines)
                bx1 = max(ln["x1"] for ln in lines)
                lm = bx0 - cx0
                rm = cx1 - bx1
                width = cx1 - cx0
                centred = len(lines) == 1 and lm > 0.12 * width and rm > 0.12 * width and abs(lm - rm) < 0.08 * width
                if rtl:
                    _set_paragraph_rtl(p)
                if centred:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif not rtl:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                lat, ar = _row_fonts(lines)
                _emit_runs(p, lines, text, heading, (bold, False, color & 0xFFFFFF),
                           size=round(size_pt * 2) / 2 if size_pt > 0 else None,
                           font=lat or latin, cs_font=ar or arabic, rtl=rtl)
                if heading:
                    p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(4 if heading else 3)
    _page_breaks_before(doc, breaks)


def _page_breaks_before(doc, starts):
    """Each later page of a section begins with "page break before" on its first block
    (19 Oct 2026). The break was a paragraph of its own at the foot of the page before,
    and on a page the PDF fills down to its bottom margin that paragraph spilled onto a
    page of its own: a 17-page report came back as 25 pages, one blank after every page
    with a picture. A page that starts with a table gets a small empty paragraph to carry
    the break, since Word keeps no such property on a table."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
    body = doc.element.body

    def blank(p):
        return p is not None and p.tag == qn("w:p") and next(
            p.iter(qn("w:t"), qn("w:drawing"), qn("w:sectPr"), qn("w:br")), None) is None

    for i in reversed(starts):      # from the end: an inserted paragraph moves no earlier index
        kids = list(body)
        el = kids[i] if i < len(kids) else None
        # the empty paragraph _add_word_table leaves after a table that ends the page
        # spilled onto a page of its own like the break did; the next page's first block
        # keeps that table apart from a table that follows
        spacer = el.getprevious() if el is not None else None
        if blank(spacer) and spacer.getprevious() is not None and spacer.getprevious().tag == qn("w:tbl"):
            body.remove(spacer)
        starts_section = el is not None and el.tag == qn("w:p") and el.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None
        if el is not None and el.tag == qn("w:p") and not starts_section:
            Paragraph(el, doc._body).paragraph_format.page_break_before = True
            continue
        carrier = OxmlElement("w:p")
        if el is not None:
            el.addprevious(carrier)
        else:
            body.append(carrier)
        p = Paragraph(carrier, doc._body)
        p.paragraph_format.page_break_before = True
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(1)


def _ocr_scanned_pages(pdf_path, pages):
    """A PDF with a text layer can still carry scanned pages (a signed page, an
    attachment scanned in). Those came back as a picture only, their text not
    editable, because OCR ran for whole scans alone; each such page is now read
    like a scan, and its text replaces the picture when OCR finds text on it (a
    photo or a drawing keeps its picture)."""
    todo = [i for i, pg in enumerate(pages) if pg.get("scan")]
    if not todo:
        return
    texts = _ocr_pages(pdf_path, todo, tables=True) or {}
    for i in todo:
        items = _ocr_items(texts.get(i))
        if sum(ch.isalpha() for ch in _items_text(items)) >= 40:
            pages[i]["ocr"] = items
            pages[i]["flow"] = []
            pages[i]["box"] = None         # the scan's edges are not the text's margins


def convert_with_fitz(pdf_path: str, docx_path: str, include_images: bool = True, tables: bool = False,
                      garbled=()):
    """Text-layer PDFs: rebuild a clean, fully editable Word document from the
    glyph POSITIONS (see the note above), one paragraph per visual paragraph,
    with size / bold / colour / alignment kept and images placed in flow.

    tables=True also rebuilds every ruled table as a Word table, its cell text
    read with the same glyph-position machinery (pdf_table_extractor). Arabic
    PDFs with a table used to go through pdf2docx, which reads PyMuPDF's raw
    order: a Chrome invoice came back "الرضيبة", "رشكة", "إىل", "مالحظات"
    for الضريبة, شركة, إلى, ملاحظات, with cells duplicated — and a 15-page
    table report took 419 s there against ~2 s here (22 Sept 2026)."""
    import fitz
    from docx import Document

    ruled_tables = None
    if tables:
        from pdf_table_extractor import ruled_tables  # noqa: E402 - imports this module

    src = fitz.open(pdf_path)
    try:
        pages = [_layout_page(page, include_images, ruled_tables) for page in src]
    finally:
        src.close()
    for i in garbled:
        if i < len(pages):
            # read like a scanned page; keeps its own text if OCR finds none
            pages[i]["scan"] = True
    _ocr_scanned_pages(pdf_path, pages)
    running = _running_items(pages)
    doc = Document()
    _emit_document(doc, pages, running)
    doc.save(docx_path)


def _find_tables(page):
    """page.find_tables(), leaving MuPDF's text settings as they were.

    find_tables() switches MuPDF to "small glyph heights" for its own character
    list and switches it back only when it returns. On some Chrome pages set in
    a web font Chrome embedded as Type3 (the site's Cairo face) it raises «no
    font file for digest» halfway, and every page read after that came back
    with other glyph boxes, rows merged and split differently: a two-page
    Chrome PDF, a Cairo page then an IBM Plex page, read 147 of its 212 words
    right in PDF to Word and PDF to Excel against 188 with the setting put back
    (10 Oct 2026, PyMuPDF 1.26.5; «يف» for «في», «األول» for «الأول»).

    On such a page it is asked once more (_find_tables_type3)."""
    import fitz
    try:
        old = bool(fitz.TOOLS.set_small_glyph_heights())
    except Exception:  # noqa: BLE001 - a PyMuPDF without the switch: as before
        return _find_tables_type3(page)
    try:
        return _find_tables_type3(page)
    finally:
        fitz.TOOLS.set_small_glyph_heights(old)


def _find_tables_type3(page):
    """page.find_tables(), asked again without TEXT_COLLECT_STYLES when a Type3 face
    makes it raise «no font file for digest»: the style collection is the only part
    that wants a font file, and a ruled table is found from its rulings. Until
    20 Oct 2026 such a page had no ruled tables at all, and PDF to Excel cut its
    columns from the gaps in the text instead: a Chrome-printed Arabic payroll of
    three ruled columns came out as five, the header and the values under
    different letters (PyMuPDF 1.26.5; asked again it finds the 3 x 4 grid)."""
    try:
        return page.find_tables()
    except RuntimeError as e:
        if "no font file" not in str(e):
            raise
    import fitz
    try:
        from pymupdf import table as tables_mod
    except ImportError:  # an older PyMuPDF, where the module is fitz.table
        from fitz import table as tables_mod
    flags = tables_mod.FLAGS
    tables_mod.FLAGS = flags & ~fitz.TEXT_COLLECT_STYLES
    try:
        return page.find_tables()
    finally:
        tables_mod.FLAGS = flags


def has_tables(pdf_path: str, max_pages: int = 6) -> bool:
    """True if the PDF looks table-heavy (checked on the first few pages).
    Table layouts reconstruct better through pdf2docx than flat text extraction."""
    import fitz
    _ensure_pdf2docx_compat()
    doc = fitz.open(pdf_path)
    try:
        for i, page in enumerate(doc):
            if i >= max_pages:
                break
            try:
                if _find_tables(page).tables:
                    return True
            except Exception:
                continue
        return False
    finally:
        doc.close()


def _docx_has_text(path: str) -> bool:
    try:
        from docx import Document
        d = Document(path)
        return (any(p.text.strip() for p in d.paragraphs)
                or any(c.text.strip() for t in d.tables for r in t.rows for c in r.cells))
    except Exception:
        return False


def _doc_is_rtl(pdf_path: str, max_pages: int = 6) -> bool:
    """True when Arabic letters are a real share of the first pages' text."""
    import fitz
    doc = fitz.open(pdf_path)
    try:
        ar = lat = 0
        for i, page in enumerate(doc):
            if i >= max_pages:
                break
            t = page.get_text()
            ar += sum(1 for c in t if _is_arabic_char(c))
            lat += sum(1 for c in t if c.isascii() and c.isalpha())
        return ar > 0 and ar >= 0.3 * (ar + lat)
    finally:
        doc.close()


def _pdf2docx_child(pdf_path, docx_path):
    import logging
    logging.disable(logging.WARNING)       # pdf2docx logs every page at INFO/WARNING
    convert_with_pdf2docx(pdf_path, docx_path)


PDF2DOCX_BUDGET_S = 45


def convert_with_pdf2docx_within(pdf_path: str, docx_path: str, budget_s: float = PDF2DOCX_BUDGET_S) -> bool:
    """pdf2docx in a child process with a time budget. A dense 15-page table
    report took 419 s in it (22 Sept 2026) while the visitor's request dies at
    Cloudflare after 100 s; past the budget the caller uses the fast engine."""
    import multiprocessing as mp
    methods = mp.get_all_start_methods()
    ctx = mp.get_context("fork" if "fork" in methods else "spawn")
    proc = ctx.Process(target=_pdf2docx_child, args=(pdf_path, docx_path))
    proc.start()
    proc.join(budget_s)
    if proc.is_alive():
        proc.terminate()
        proc.join(5)
        try:
            os.remove(docx_path)
        except OSError:
            pass
        return False
    return proc.exitcode == 0 and os.path.exists(docx_path) and os.path.getsize(docx_path) > 0


# .NET stops the whole request at 85 s (WorkBudget). Run from the command line,
# OCR reads pages until the process has run OCR_BUDGET_S seconds; a page it has
# not read by then comes back as its picture, so a 36-page scan returns every
# page (35 read, 65 s) instead of a timeout with nothing (85.6 s, 28 Sept 2026).
OCR_BUDGET_S = 65.0
_ocr_deadline = None           # set by main(); none when used as a library
_deadline = None               # the API's --budget, from the process start; none without one
# The end of the work after the last page OCR reads: the unread pages as pictures and
# the RTL pass over the document. Measured 10 Oct 2026 on a 36-page A4 scan (4 busy
# cores): with --budget 50 the reply came after 45.0 s (21 pages read, 15 pictures), with
# --budget 35 after 29.8 s. Without one the script ran 61.1 s, so a request that had
# waited 25 s of its 85 in the queue was killed at its deadline with nothing to show.
FINISH_S = 8.0


def _time_left():
    return float("inf") if _deadline is None else _deadline - time.monotonic()


def _pdf2docx_budget():
    """pdf2docx's time: its own cap, and never past the API's deadline."""
    return max(1.0, min(PDF2DOCX_BUDGET_S, _time_left() - FINISH_S))


def _pdf_title(pdf_path):
    """The PDF's own title, when the caller gives none."""
    try:
        import fitz
        doc = fitz.open(pdf_path)
        try:
            return ((doc.metadata or {}).get("title") or "").strip()
        finally:
            doc.close()
    except Exception:  # noqa: BLE001 - no title then
        return ""


def _ocr_pages(pdf_path, pages=None, dpi=300, deadline=None, tables=False):
    """{page index: text} from Tesseract (ara+eng) for `pages` (all by default),
    or None when Tesseract is missing. A page absent from the result was not
    read (a failure, or out of time). With tables=True each page is (text, its
    word boxes as Tesseract's TSV, the dpi it was read at), both written by one
    recognition pass (_ocr_items builds the page's tables from the boxes).

    Each page is rendered grey on this thread and piped to `tesseract stdin
    stdout` as PGM while the pages before it are being read. Until 28 Sept 2026
    every page was first encoded to PNG (MuPDF), decoded (PIL) and encoded to a
    temp PNG again (pytesseract) before the first page was read: 40.8 s for a
    12-page A4 scan that now takes 16 s on the same 4 cores. PyMuPDF is not
    thread-safe, so rendering stays here; the threads only wait for their
    tesseract process, which is held to one core (OMP_THREAD_LIMIT=1).

    Tesseract, not the Paddle service: measured 8 Sept 2026, PaddleOCR's Arabic
    model drops the spaces between words ("شركةالنيلللمقاولات") and scored
    0.31/0.41 word recall on an Arabic payroll and CV scan against Tesseract's
    0.38/0.59 (English 0.71 vs 0.72)."""
    import shutil
    import subprocess
    import fitz
    from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
    tess = shutil.which("tesseract")
    if not tess:
        return None
    env = dict(os.environ, OMP_THREAD_LIMIT="1")
    if deadline is None:
        deadline = _ocr_deadline if _ocr_deadline is not None else float("inf")

    def script_order(img, res):
        """"ara+eng" for an Arabic page, "eng+ara" otherwise, from Tesseract's
        script detection on the first page (0.8 s). With Arabic first, Tesseract
        reads a page's text blocks right to left, so an English table scan came
        back column by column from its last column («Amount 17,915.73 15,240.94
        … Unit price 853.13 …»: 27% of the words in order, 99.9% with English
        first; an Arabic scan reads the same either way, 2 Oct 2026)."""
        try:
            r = subprocess.run([tess, "stdin", "stdout", "--psm", "0", "-l", "osd", "--dpi", str(res)],
                               input=img, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env, timeout=20)
            m = re.search(r"Script:\s*(\w+)", r.stdout.decode("utf-8", "replace"))
        except (subprocess.TimeoutExpired, OSError):
            m = None
        if m is None:
            return "ara+eng"          # no script detection here: as before
        return "ara+eng" if m.group(1).lower() == "arabic" else "eng+ara"

    order = None

    def read(img, res):
        lang = order.result() if order is not None else "ara+eng"
        left = deadline - time.monotonic()
        if left <= 1.0:
            return None
        timeout = None if left == float("inf") else left
        if tables:
            import tempfile
            tmp = tempfile.mkdtemp(prefix="ocr-")
            try:
                base = os.path.join(tmp, "page")
                # textord_noise_rejrows=0: Tesseract's noise-row test dropped whole lines of
                # text, a title included (see OcrService, 19 Oct 2026)
                r = subprocess.run([tess, "stdin", base, "-l", lang, "--dpi", str(res),
                                    "-c", "textord_noise_rejrows=0", "txt", "tsv"],
                                   input=img, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                   env=env, timeout=timeout)
                if r.returncode != 0:
                    return None
                with open(base + ".txt", encoding="utf-8", errors="replace") as f:
                    text = f.read()
                with open(base + ".tsv", encoding="utf-8", errors="replace") as f:
                    return text, f.read(), res
            except (subprocess.TimeoutExpired, OSError):
                return None
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
        try:
            r = subprocess.run([tess, "stdin", "stdout", "-l", lang, "--dpi", str(res),
                                "-c", "textord_noise_rejrows=0"],
                               input=img, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                               env=env, timeout=timeout)
        except (subprocess.TimeoutExpired, OSError):
            return None
        return r.stdout.decode("utf-8", "replace") if r.returncode == 0 else None

    workers = max(1, min(4, (os.cpu_count() or 2) - 1))
    out, futs = {}, {}
    src = fitz.open(pdf_path)
    try:
        todo = range(len(src)) if pages is None else pages
        with ThreadPoolExecutor(max_workers=workers) as pool:
            if len(todo):
                first = src[todo[0]]
                z = 150 / 72.0
                longest = max(first.rect.width, first.rect.height) or 1.0
                if longest * z > 2500:
                    z = 2500 / longest
                pix = first.get_pixmap(matrix=fitz.Matrix(z, z), colorspace=fitz.csGRAY, alpha=False)
                order = pool.submit(script_order, pix.tobytes("pgm"), int(round(z * 72)))
                del pix
            for pno in todo:
                if time.monotonic() > deadline - 2.0:
                    break
                # two rendered pages per reader at most: a 300 dpi A4 page is 8.7 MB
                waiting = [f for f in futs.values() if not f.done()]
                while len(waiting) >= 2 * workers:
                    wait(waiting, return_when=FIRST_COMPLETED)
                    waiting = [f for f in futs.values() if not f.done()]
                page = src[pno]
                zoom = dpi / 72.0
                longest = max(page.rect.width, page.rect.height) or 1.0
                if longest * zoom > 5000:          # a poster-sized page: cap the raster
                    zoom = 5000 / longest
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csGRAY, alpha=False)
                futs[pno] = pool.submit(read, pix.tobytes("pgm"), int(round(zoom * 72)))
                del pix
            for pno, f in futs.items():
                try:
                    text = f.result()
                except Exception:  # noqa: BLE001 - that page stays unread
                    text = None
                if text is not None:
                    out[pno] = text
    finally:
        src.close()
    return out


def _ocr_paragraphs(text):
    """Tesseract's text output → paragraphs (its blocks are separated by a blank line).

    A list item starts a paragraph of its own: Tesseract puts a list in one
    block, and the items ran together («• New warehouse opened in March •
    Customer churn fell to 3.1% • …»). With Arabic loaded it reads an English
    bullet as the Arabic-Indic zero «٠» (2 Oct 2026); on a Latin line that is
    the bullet it stands for."""
    paras, cur = [], []
    for line in text.splitlines() + [""]:
        line = _xml_clean(line.strip())
        if line[:2] == "\u0660 " and not is_arabic_text(line):
            line = "\u2022" + line[1:]
        if line and cur and line[:1] in _LIST_MARKS and line[1:2] == " ":
            paras.append(" ".join(cur))
            cur = []
        if line:
            cur.append(line)
        elif cur:
            paras.append(" ".join(cur))
            cur = []
    return paras


def _add_ocr_paragraphs(doc, paras):
    """A read page: its paragraphs, and its tables (_ocr_items) as Word tables."""
    for text in paras:
        if isinstance(text, dict):
            _add_word_table(doc, text["rows"], text["rtl"])
            continue
        p = doc.add_paragraph(text)
        if is_arabic_text(text):
            _set_paragraph_rtl(p)


def _items_text(items):
    return "".join(c for it in items for c in (" ".join(" ".join(r) for r in it["rows"])
                                                if isinstance(it, dict) else it))


def _ocr_items(got):
    """A page _ocr_pages read → [paragraph text | {"rows", "rtl"} table] in reading
    order.

    Tesseract's text came back one line per paragraph, so a page that is a table
    lost it: the 40-row link table on page 4 of a Chrome PDF whose text layer is
    garbled (read like a scan) became ~80 one-line paragraphs, its row numbers
    121…140 a column of their own, while PDF to Excel rebuilt the same page as a
    49-row table from the same OCR (10 Oct 2026). The word boxes now go through
    PDF to Excel's row and column builder (pdf_table_extractor: _tsv_pieces,
    ocr_seg_rows, build_blocks) too. Its tables are kept only where the page is
    mostly table: at least 4 rows of short cells, most of them with two or more
    filled, together at least 60 % of the page's rows. Two columns of prose came
    out of the builder as a "table" of 42-character cells, and a bulleted list as
    one beside a small table; such pages stay paragraphs, exactly as before."""
    if got is None:
        return []
    if isinstance(got, str):
        return _ocr_paragraphs(got)
    text, tsv, res = got
    try:
        return _ocr_page_with_tables(text, tsv, res)
    except Exception:  # noqa: BLE001 - the page's paragraphs, as before
        return _ocr_paragraphs(text)


def _ocr_page_with_tables(text, tsv, res):
    import statistics
    from pdf_table_extractor import _tsv_pieces, ocr_seg_rows, build_blocks, _script_counts
    pieces = _tsv_pieces(tsv, 72.0 / res)
    ar, lat = _script_counts(" ".join(p["text"] for p in pieces))
    rtl = ar > lat
    seg_rows = ocr_seg_rows({"lines": pieces}, rtl)
    tables = []
    for b in build_blocks(seg_rows, rtl):
        if b["kind"] != "table" or len(b["rows"]) < 4 or max(len(r) for r in b["rows"]) < 2:
            continue
        cells = [c for r in b["rows"] for c in r if c]
        multi = sum(1 for r in b["rows"] if sum(1 for c in r if c) >= 2)
        if cells and statistics.median(len(c) for c in cells) <= 25 and multi >= 0.7 * len(b["rows"]):
            tables.append(b)
    if not tables or sum(len(b["rows"]) for b in tables) < 0.6 * len(seg_rows):
        return _ocr_paragraphs(text)
    # Tesseract's lines (its text, with the direction marks it writes around Latin words
    # in an Arabic line) beside their boxes from the TSV: the lines inside a table go,
    # the others keep their paragraphs, between the tables by height
    boxes = {}
    for row in tsv.splitlines()[1:]:
        f = row.split("\t")
        if len(f) < 12 or f[0] != "5" or not f[11].strip():
            continue
        try:
            y, h = int(f[7]), int(f[9])
        except ValueError:
            continue
        key = (int(f[2]), int(f[3]), int(f[4]))
        y0, y1 = boxes.get(key, (y, y + h))
        boxes[key] = (min(y0, y), max(y1, y + h))
    lines = [ln for ln in text.splitlines()]
    shown = [ln for ln in lines if ln.strip()]
    if len(shown) != len(boxes):
        return _ocr_paragraphs(text)                  # the two do not line up: as before
    scale = 72.0 / res
    mids = iter([(a + b) / 2.0 * scale for _, (a, b) in sorted(boxes.items())])
    tables.sort(key=lambda b: b["y0"])
    parts = [[] for _ in range(len(tables) + 1)]
    for ln in lines:
        if not ln.strip():
            for part in parts:
                part.append("")
            continue
        y = next(mids)
        if any(b["y0"] - 1 <= y <= b["y1"] + 1 for b in tables):
            continue
        parts[sum(1 for b in tables if b["y0"] < y)].append(ln)
    items = []
    for k, part in enumerate(parts):
        items.extend(_ocr_paragraphs("\n".join(part)))
        if k < len(tables):
            items.append({"rows": tables[k]["rows"], "rtl": rtl})
    return items


def _add_page_picture(doc, page, section):
    """The page itself, as a picture that fits the text area (a page OCR did not read)."""
    import fitz
    from docx.shared import Inches
    pix = page.get_pixmap(matrix=fitz.Matrix(120 / 72.0, 120 / 72.0), alpha=False)
    try:
        data = pix.tobytes("jpg", jpg_quality=80)
    except Exception:  # noqa: BLE001 - an older PyMuPDF: PNG
        data = pix.tobytes("png")
    tw = (section.page_width - section.left_margin - section.right_margin) / 914400.0
    th = (section.page_height - section.top_margin - section.bottom_margin) / 914400.0 - 0.2
    ratio = (page.rect.width or 1.0) / (page.rect.height or 1.0)
    doc.add_paragraph().add_run().add_picture(io.BytesIO(data), width=Inches(max(0.5, min(tw, th * ratio))))


def convert_with_ocr(pdf_path: str, docx_path: str, dpi: int = 300):
    """Scanned PDFs: OCR into paragraphs in reading order — Tesseract first (see
    _ocr_pages), the PaddleOCR service when Tesseract is missing. Every page
    keeps the PDF's paper size (it was US Letter until 28 Sept 2026)."""
    import fitz
    from docx import Document
    from docx.enum.section import WD_SECTION
    doc = Document()
    texts = _ocr_pages(pdf_path, dpi=dpi, tables=True)
    if texts is not None:
        src = fitz.open(pdf_path)
        try:
            section, size = None, None
            for pno, page in enumerate(src):
                this = (round(page.rect.width, 1), round(page.rect.height, 1))
                if section is None or this != size:
                    section = doc.sections[0] if section is None else doc.add_section(WD_SECTION.NEW_PAGE)
                    _section_setup(section, page.rect.width, page.rect.height, None)
                    size = this
                else:
                    doc.add_page_break()
                if pno in texts:
                    _add_ocr_paragraphs(doc, _ocr_items(texts[pno]))
                else:
                    _add_page_picture(doc, page, section)
        finally:
            src.close()
        doc.save(docx_path)
        return

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import ocr_client
    pages = ocr_client.ocr_pdf(pdf_path, lang="ar").get("pages", [])
    for i, pg in enumerate(pages):
        lines = [l for l in pg.get("lines", []) if (l.get("text") or "").strip()]
        lines.sort(key=lambda l: (l["box"][1] if l.get("box") else 0))
        for l in lines:
            doc.add_paragraph(_xml_clean((l["text"] or "").strip()))
        if i < len(pages) - 1:
            doc.add_page_break()
    doc.save(docx_path)


# Exit codes the .NET service reads (Service/PdfToWordService.cs).
EXIT_OK = 0
EXIT_FAILED = 1            # .NET may try another engine
EXIT_PASSWORD = 3          # the PDF needs an open password: no engine can read it


def _needs_password(pdf_path):
    """True when the PDF cannot be opened without a user password. An owner
    password alone (printing/copying restrictions) opens and converts as before."""
    try:
        import fitz
        doc = fitz.open(pdf_path)
        try:
            return bool(doc.needs_pass)
        finally:
            doc.close()
    except Exception:  # noqa: BLE001 - a damaged file is the engines' business
        return False


def main():
    global _ocr_deadline, _deadline
    started = time.monotonic()
    ap = argparse.ArgumentParser(description="PDF to Word (pdf2docx + Arabic RTL + OCR fallback)")
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--no-images", action="store_true")  # accepted for CLI compatibility
    ap.add_argument("--title", default="", help="the document's title (the uploaded file's name)")
    ap.add_argument("--budget", type=float, default=0,
                    help="seconds the API lets this process run (0 = the script's own limits)")
    args = ap.parse_args()
    if args.budget > 0:
        # what is left of the request when the API starts this process (it waited in the
        # queue first), less the time the end of the work needs: the unread pages as
        # pictures and the RTL pass over the document
        _deadline = started + args.budget
        _ocr_deadline = min(started + OCR_BUDGET_S, _deadline - FINISH_S)
    else:
        _ocr_deadline = started + OCR_BUDGET_S

    # The reply is ONE JSON line on stdout. PyMuPDF prints a "Consider using the
    # pymupdf_layout package…" hint to stdout on its first table search, and the
    # .NET side failed to parse every reply (28 Sept 2026); anything a library
    # prints now goes to stderr and the hint is switched off.
    os.environ.setdefault("PYMUPDF_SUGGEST_LAYOUT_ANALYZER", "0")
    reply = sys.stdout
    sys.stdout = sys.stderr

    result = {"success": False, "engine": None, "rtl": False, "message": "", "code": None}
    code = EXIT_FAILED
    try:
        if not os.path.exists(args.input):
            raise FileNotFoundError(f"Input not found: {args.input}")

        if _needs_password(args.input):
            # Every engine failed on these one after another ("document closed or
            # encrypted"), and the visitor was told the file was corrupt.
            result["code"] = "passwordProtected"
            result["message"] = "The PDF is protected by an open password"
            code = EXIT_PASSWORD
        else:
            # Hybrid router: pick the engine that gives the best result per document.
            #   • no text layer (scanned)      -> OCR (Tesseract, pages in parallel)
            #   • tables                       -> fitz + Word tables (pdf2docx only when that fails)
            #   • text layer, flowing text     -> fitz (clean, fully-editable, no jamming)
            # Latin PDFs with a table went to pdf2docx first until 2 Oct 2026. Measured
            # on six table PDFs against fitz + tables: a 2-page, 90-row ledger took 30 s
            # there (1.6 s here) and came back with fixed row heights over too-narrow
            # columns, so "description" was cut out of every row; an invoice came back
            # with page 1 empty and the whole invoice on page 2; a merged-header table
            # was stretched to the page width. pdf2docx did better on two of the six
            # (a Word table style's borders and spacing), fitz never lost text, and a
            # long table no longer spends 45 s in pdf2docx before it is given up.
            #   • text layer that is mojibake   -> OCR for those pages (all of them: the OCR engine)
            text_pages, garbled = _text_pages(args.input)
            ocr_whole = not text_pages
            if text_pages and garbled >= text_pages:
                try:
                    convert_with_ocr(args.input, args.output)
                    ocr_whole = True
                except Exception as e:  # noqa: BLE001 - no OCR engine: its own text, as before
                    sys.stderr.write(f"OCR of a garbled text layer failed: {e}\n")
                    garbled = set()
            if ocr_whole:
                if not text_pages:
                    convert_with_ocr(args.input, args.output)
                result["engine"] = "ocr"
            elif has_tables(args.input):
                try:
                    convert_with_fitz(args.input, args.output, include_images=not args.no_images, tables=True,
                                      garbled=garbled)
                    result["engine"] = "fitz+tables"
                except Exception as e:  # noqa: BLE001 - pdf2docx gets its turn (Latin only)
                    sys.stderr.write(f"fitz+tables failed: {type(e).__name__}: {e}\n")
                if (result["engine"] is None or not _docx_has_text(args.output)) and not garbled \
                        and not _doc_is_rtl(args.input) and _time_left() > 5 \
                        and convert_with_pdf2docx_within(args.input, args.output, _pdf2docx_budget()):
                    result["engine"] = "pdf2docx"
            else:
                convert_with_fitz(args.input, args.output, include_images=not args.no_images, garbled=garbled)
                result["engine"] = "fitz"
                if not _docx_has_text(args.output) and _time_left() > 5 \
                        and convert_with_pdf2docx_within(args.input, args.output, _pdf2docx_budget()):   # safety net
                    result["engine"] = "pdf2docx"

            if not (os.path.exists(args.output) and os.path.getsize(args.output) > 0):
                raise RuntimeError("No output file was produced")

            result["rtl"] = postprocess_rtl(args.output, args.title or _pdf_title(args.input))
            result["success"] = True
            result["message"] = "Converted successfully"
            code = EXIT_OK
    except Exception as e:  # noqa: BLE001 — report to .NET so it can fall back
        result["message"] = f"{type(e).__name__}: {e}"
        try:
            sys.stderr.write(result["message"] + "\n")
        except Exception:
            pass
    finally:
        sys.stdout = reply

    if args.json:
        reply.write(json.dumps(result, ensure_ascii=False) + "\n")
        reply.flush()
    sys.exit(code)


if __name__ == "__main__":
    main()
