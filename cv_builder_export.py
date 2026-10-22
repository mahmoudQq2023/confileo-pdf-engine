#!/usr/bin/env python3
"""
CV Builder — PDF export (WeasyPrint), Word export (python-docx) and CV import (PyMuPDF).

Produces ATS-friendly, text-based PDFs that match the Angular preview, in
fonts bundled in Fonts/cv/, with right-to-left layout for Arabic.

Usage:
    python cv_builder_export.py export <input.json> <output.pdf>
    python cv_builder_export.py docx   <input.json> <output.docx>
    python cv_builder_export.py parse  <input.pdf>  <output.json>
"""

import json
import os
import re
import sys
from io import BytesIO

try:
    import fitz  # PyMuPDF
    _HAS_FITZ = True
except ImportError:
    _HAS_FITZ = False


_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


# ══════════════════════════════════════════════════════════════════
# Utility Helpers
# ══════════════════════════════════════════════════════════════════

def _has_arabic(text: str) -> bool:
    """Check if a string contains Arabic characters."""
    if not text:
        return False
    return bool(re.search(r'[\u0600-\u06FF]', text))


def _esc(text: str) -> str:
    """Escape HTML special characters."""
    if not text:
        return ""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


# ── Bidi isolation (28 Sept 2026) ─────────────────────────────────────
# Inside a right-to-left paragraph the Unicode bidi algorithm lays out every
# number group and every Latin run with the paragraph's direction deciding the
# neutral characters between and around them. Measured on an Arabic CV
# (WeasyPrint 70, rawdict glyph order): the header phone "+20 100 123 4567"
# printed as "4567 123 100 20+", "C#" in a bullet as "#C", and a comma list of
# skills put its commas at the wrong ends. U+2068 FIRST STRONG ISOLATE /
# U+2069 POP DIRECTIONAL ISOLATE wrap a run so it is laid out on its own
# (left-to-right for Latin, and for a run with no strong character at all such
# as a phone number) and counts as ONE neutral for the text around it. Pango
# and Chrome consume the two marks; they never draw a glyph.
FSI, PDI = "⁨", "⁩"
# Arabic LETTERS and Arabic-Indic digits. Arabic punctuation (، ؛ ؟) is left
# out on purpose: "Node.js،" is a Latin word followed by an Arabic comma.
_AR_LETTER = re.compile(r"[ؠ-ي٠-٩ٮ-ۓەۮ-ۿ"
                        r"ݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]")
_LATIN_OR_DIGIT = re.compile(r"[A-Za-z0-9À-ɏ]")
_ISO_LEAD = "([{«\"'"
_ISO_TRAIL = ".,;:!?)]}»\"'،؛؟"


def _isolate(text: str) -> str:
    """The whole string as one isolated run (a phone number, an e-mail, a skill)."""
    return f"{FSI}{text}{PDI}" if text and text.strip() else text


def _isolate_ltr_runs(text: str) -> str:
    """Isolate every run of Latin words / numbers inside right-to-left text.

    Consecutive non-Arabic words form one run ("SQL Server 2016", "+20 100 123
    4567"); brackets and sentence punctuation around a run stay OUTSIDE it, so
    they follow the Arabic sentence ("(Full-Stack .NET)." keeps its parentheses
    mirrored and its full stop at the end of the Arabic line). A string with no
    Arabic letter at all is isolated whole. Used for right-to-left documents
    only; a left-to-right CV is left exactly as it was.
    """
    if not text or not _LATIN_OR_DIGIT.search(text):
        return text
    if not _AR_LETTER.search(text):
        return _isolate(text)
    out, run = [], []

    def flush():
        if not run:
            return
        raw = "".join(run)
        core = raw.rstrip()
        tail_ws = raw[len(core):]
        lead = ""
        while core and core[0] in _ISO_LEAD:
            lead, core = lead + core[0], core[1:]
        trail = ""
        while core and core[-1] in _ISO_TRAIL:
            trail, core = core[-1] + trail, core[:-1]
        out.append(lead + (_isolate(core) if core else "") + trail + tail_ws)
        run.clear()

    for tok in re.split(r"(\s+)", text):
        if not tok:
            continue
        if tok.isspace():
            (run if run else out).append(tok)
        elif not _AR_LETTER.search(tok) and _LATIN_OR_DIGIT.search(tok):
            run.append(tok)
        else:
            flush()
            out.append(tok)
    flush()
    return "".join(out)


# ══════════════════════════════════════════════════════════════════
# HTML Builders — Match Angular Preview Exactly
# ══════════════════════════════════════════════════════════════════

def _build_contact_html(personal: dict, isolate: bool = False) -> str:
    """Build the contact info line with dot separators.

    Label-free links to stay clean and ATS-friendly, and to exactly match
    the live preview (cv-builder.service.ts buildContactHtml).

    `isolate` wraps every value in FSI…PDI (see _isolate_ltr_runs): each item is
    then laid out on its own, so a phone number keeps its digit groups in order
    inside an Arabic header. The href never carries the marks.
    """
    iso = _isolate if isolate else (lambda s: s)
    sep = ' <span style="display:inline-block;margin:0 0.5em;color:#999999;font-weight:600;font-size:1.1em;line-height:1;vertical-align:middle">&middot;</span> '
    parts = []
    if personal.get("email"):
        parts.append(f'<span>{_esc(iso(personal["email"]))}</span>')
    if personal.get("phone"):
        parts.append(f'<span>{_esc(iso(personal["phone"]))}</span>')
    if personal.get("location"):
        parts.append(f'<span>{_esc(iso(personal["location"]))}</span>')
    for key in ("linkedin", "github", "portfolio", "website"):
        if personal.get(key):
            url = personal[key]
            href = url if url.startswith("http") else f"https://{url}"
            parts.append(f'<a href="{_esc(href)}" style="color:#000;text-decoration:none">{_esc(iso(url))}</a>')
    return sep.join(parts)


def _fmt_dates(start: str, end: str, present: str = "Present") -> str:
    """Join a start/end date into a single range; blanks handled gracefully.

    `present` is the only word this function writes, so it is also the only
    English word that used to leak onto a CV rendered in another language.
    """
    start = (start or "").strip()
    end = (end or "").strip()
    if start and end:
        return f"{start} &ndash; {end}"
    if start:
        return f"{start} &ndash; {present}"
    return end


# Bullet/dash glyphs that leak in from imported CVs (Word/PDF) and render as
# ugly squares or double bullets. We strip leading ones and drop glyph-only lines.
_BULLET_GLYPHS = "•▪■●○◆◇‣⁃·∙°⁌⁍-–—*»«"


_INVISIBLE = {0x00ad, 0x200b, 0x200c, 0x200d, 0x2060, 0xfeff}


def _clean_line(text: str) -> str:
    """Strip leftover bullet/dash glyphs from a line; return '' if only glyphs.
    Also drops zero-width/BOM/soft-hyphen/control chars so blank-looking lines
    (which otherwise render as a lone empty bullet) collapse to empty."""
    t = "".join(ch for ch in (text or "")
                if ord(ch) not in _INVISIBLE and (ord(ch) >= 32 or ch == "\t"))
    t = t.strip()
    while t and t[0] in _BULLET_GLYPHS:
        t = t[1:].lstrip()
    return t if t.strip(_BULLET_GLYPHS + " .\t ") else ""


def _languages_inline(languages: list, isolate: bool = False) -> str:
    """Inline languages list. Skips empty names and meaningless levels ('-').

    `isolate`: right-to-left documents isolate the name and the level, so
    "English (C1)" or "الألمانية (A2)" keeps its brackets around the level."""
    iso = _isolate if isolate else (lambda s: s)
    parts = []
    for lang in languages:
        raw_name = str(lang.get("name", "")).strip()
        if not raw_name:
            continue
        name = _esc(iso(raw_name))
        level = str(lang.get("level", "")).strip()
        if level and level not in ("-", "–", "—"):
            parts.append(f'<b style="font-weight:700;color:#000">{name}</b> <span style="color:#666666">({_esc(iso(level))})</span>')
        else:
            parts.append(f'<b style="font-weight:700;color:#000">{name}</b>')
    sep = ' <span style="color:#bbbbbb;margin:0 0.4em">&middot;</span> '
    return sep.join(parts)


# ══════════════════════════════════════════════════════════════════
# Free-tier attribution footer
# ══════════════════════════════════════════════════════════════════
# A CV is the most widely forwarded document this site produces — one file
# reaches dozens of employers and recruiters, all of them people who work with
# documents daily. The invoice and certificate generators already carry this
# footer; the CV, which travels furthest, was going out with no trace of where
# it was made.
#
# `watermark` is decided by the API from the user's subscription and is NOT read
# from the client payload — a browser that could set it would be a free upgrade
# button. Paid plans render nothing at all.

_WM_TEXT = {
    "en": 'Created free with <a href="https://confileo.com/tools/cv-builder/">Confileo</a> · confileo.com',
    "ar": 'أُنشئت مجاناً عبر <a href="https://confileo.com/ar/sinaat-cv/">كونفيليو</a> · confileo.com',
    "es": 'Creado gratis con <a href="https://confileo.com/es/crear-cv/">Confileo</a> · confileo.com',
    "fr": 'Créé gratuitement avec <a href="https://confileo.com/fr/creer-cv/">Confileo</a> · confileo.com',
    "hi": 'Confileo से मुफ़्त बनाया गया · <a href="https://confileo.com/hi/cv-banaye/">confileo.com</a>',
    "de": 'Kostenlos erstellt mit <a href="https://confileo.com/de/tools/cv-builder/">Confileo</a> · confileo.com',
    "pt": 'Criado grátis com <a href="https://confileo.com/pt/criar-curriculo/">Confileo</a> · confileo.com',
    "id": 'Dibuat gratis dengan <a href="https://confileo.com/id/buat-cv/">Confileo</a> · confileo.com',
    "it": 'Creato gratis con <a href="https://confileo.com/it/crea-curriculum/">Confileo</a> · confileo.com',
    "pl": 'Utworzono za darmo w <a href="https://confileo.com/pl/kreator-cv/">Confileo</a> · confileo.com',
    "tr": '<a href="https://confileo.com/tr/cv-olustur/">Confileo</a> ile ücretsiz oluşturuldu · confileo.com',
}


def _watermark_html(data: dict) -> str:
    """One line at the very bottom, or nothing for paid users.

    Kept visually quiet on purpose: a CV is the user's professional document, and
    a loud banner would push them to rebuild it somewhere else — which costs the
    reach this footer exists to earn.
    """
    if not data.get("watermark"):
        return ""

    lang = (data.get("docLang") or data.get("doc_lang") or "").lower()[:2]
    if lang not in _WM_TEXT:
        # 'auto' or unset: fall back to the script the CV itself is written in.
        name = str((data.get("personal") or {}).get("full_name", ""))
        lang = "ar" if _has_arabic(name + " " + str(data.get("summary", ""))) else "en"

    return (
        '<div style="margin-top:18px;padding-top:7px;border-top:1px solid #e5e7eb;'
        'text-align:center;font-size:7.5pt;color:#9ca3af;">'
        + _WM_TEXT[lang]
        + "</div>"
    )


def _normalize_cv_data(data: dict) -> dict:
    """Normalize input data to a consistent internal format.

    Supports two input formats:
    1. Sections-based (from backend/ project):
       { "title": "...", "sections": [ { "section_type": "...", "content": ... }, ... ] }
    2. Flat (from confileo3 frontend):
       { "title": "...", "personal": {...}, "summary": "...", "experience": [...], ... }
    """
    if "sections" in data:
        # Sections-based format — parse into flat
        sections_map = {}
        for section in data["sections"]:
            content = section.get("content")
            if isinstance(content, str):
                try:
                    content = json.loads(content)
                except (json.JSONDecodeError, TypeError):
                    pass
            sections_map[section.get("section_type", "")] = content

        font_settings = sections_map.get("font_settings", {})
        if isinstance(font_settings, str):
            font_settings = {}

        return {
            "title": data.get("title", "CV"),
            "personal": sections_map.get("personal_info", {}),
            "summary": sections_map.get("summary", ""),
            "experience": sections_map.get("experience", []),
            "education": sections_map.get("education", []),
            "skills": sections_map.get("skills", []),
            "soft_skills": sections_map.get("soft_skills", []),
            "languages": sections_map.get("languages", []),
            "font_family": font_settings.get("font_family", data.get("fontFamily", "")),  # "" = template default
            "font_size": int(font_settings.get("font_size", data.get("fontSize", 13))),
            "template_id": data.get("templateId", data.get("template_id", "modern_professional")).lower(),
        }
    else:
        # Flat format — normalize keys (camelCase -> snake_case where needed)
        personal = data.get("personal", {})
        # Handle camelCase keys from C# serialization
        if "fullName" in personal:
            personal = {
                "full_name": personal.get("fullName", ""),
                "job_title": personal.get("jobTitle", ""),
                "email": personal.get("email", ""),
                "phone": personal.get("phone", ""),
                "location": personal.get("location", ""),
                "linkedin": personal.get("linkedin", ""),
                "portfolio": personal.get("portfolio", ""),
                "website": personal.get("website", ""),
                "github": personal.get("github", ""),
            }

        # Normalize experience entries
        experience = data.get("experience", [])
        for exp in experience:
            if "startDate" in exp:
                exp["start_date"] = exp.pop("startDate", "")
            if "endDate" in exp:
                exp["end_date"] = exp.pop("endDate", "")

        # Normalize education entries
        education = data.get("education", [])
        for edu in education:
            if "startDate" in edu:
                edu["start_date"] = edu.pop("startDate", "")
            if "endDate" in edu:
                edu["end_date"] = edu.pop("endDate", "")

        return {
            "title": data.get("title", "CV"),
            "personal": personal,
            "summary": data.get("summary", ""),
            "experience": experience,
            "education": education,
            "skills": data.get("skills", []),
            "soft_skills": data.get("softSkills", data.get("soft_skills", [])),
            "languages": data.get("languages", []),
            "font_family": data.get("fontFamily", data.get("font_family", "")),  # "" = template default
            "font_size": int(data.get("fontSize", data.get("font_size", 13))),
            "template_id": data.get("templateId", data.get("template_id", "modern_professional")).lower(),
        }


# ══════════════════════════════════════════════════════════════════
# PREMIUM template engine (WeasyPrint) — FlowCV-class output
# ══════════════════════════════════════════════════════════════════
# WeasyPrint supports real CSS (web fonts, backgrounds, borders, proper
# typography) and shapes Arabic natively via Pango.
# Per-template config for the premium engine. Each is a distinct visual
# identity (accent + header alignment + chip palette). All single-column for
# ATS-safe, reliable page flow; richer layouts (sidebar) come as a later batch.
# Each template carries a full layout + style spec. `layout` selects the
# renderer: "standard" (sans single-column), "serif" (serif single-column),
# "banner" (full-width coloured header band) or "sidebar" (two-column with a
# full-height coloured side panel). These mirror the 7 reference designs.
# Tracking (letter-spacing) that a résumé parser still reads as words. Until
# 10 Oct 2026 the section headings, the role line and the minimal template's
# name were tracked 1.5-5 px (up to 0.36em), and the PDF's text came back
# letter by letter — "E X P E R I E N C E", "S A R A H M I T C H E L L" — in
# pdfminer.six (the basis of many ATS parsers), PDFium (Chrome's viewer, copy
# and Ctrl+F) and MuPDF: 231 / 158 / 159 such runs over the seven templates in
# English, German and Turkish. A sweep over the CV's own faces (DejaVu Sans and
# Liberation Serif, weights 600-800, 8.4-36.8 px, i.e. every size the Size menu
# allows) read every word correctly up to 0.08em in all three and lost words
# from 0.1em, so tracking is capped at 0.06em, in em so it scales with the
# Size menu. The names already tracked 0.5px / 1px (< 0.05em) are unchanged.
# Mirrored in cv-builder.service.ts (CV_ATS_LS).
_ATS_LS = "0.06em"

_TEMPLATE_CFG = {
    "modern_professional": {
        "layout": "standard", "font": "sans", "accent": "#0f766e",
        "header_align": "left", "name_transform": "uppercase", "name_ls": "0.5px",
        "header_rule": "2.5px solid #0f766e", "head_rule": "1.3px solid #0f766e",
        "head_align": "start", "head_ls": _ATS_LS, "skills_style": "chips",
        "chip_bg": "#eef7f5", "chip_border": "#cfe8e3", "chip_text": "#155e54", "chip_radius": 12,
    },
    "modern": {
        "layout": "standard", "font": "sans", "accent": "#0f766e",
        "header_align": "left", "name_transform": "uppercase", "name_ls": "0.5px",
        "header_rule": "2.5px solid #0f766e", "head_rule": "1.3px solid #0f766e",
        "head_align": "start", "head_ls": _ATS_LS, "skills_style": "chips",
        "chip_bg": "#eef7f5", "chip_border": "#cfe8e3", "chip_text": "#155e54", "chip_radius": 12,
    },
    "classic": {
        "layout": "serif", "font": "serif", "accent": "#1f3a5f",
        "header_align": "center", "name_transform": "uppercase", "name_ls": "1px",
        "header_rule": "4px double #1f3a5f", "double_top": True,
        "head_rule": "1.5px solid #1f3a5f", "head_align": "center", "head_ls": _ATS_LS,
        "skills_style": "comma",
        "chip_bg": "#eef1f6", "chip_border": "#c9d6e5", "chip_text": "#1f3a5f", "chip_radius": 4,
    },
    "elegant": {
        "layout": "serif", "font": "serif", "accent": "#7c2d12",
        "header_align": "left", "name_transform": "uppercase", "name_ls": "1px",
        "header_rule": "1.5px solid #7c2d12", "head_rule": "1.5px solid #7c2d12",
        "head_align": "start", "head_underline": "short", "head_ls": _ATS_LS,
        "skills_style": "comma", "page_bg": "#faf6ef",
        "chip_bg": "#f7efe9", "chip_border": "#e6d2c5", "chip_text": "#7c2d12", "chip_radius": 4,
    },
    "minimal": {
        "layout": "serif", "font": "serif", "accent": "#1a1a1a",
        "header_align": "center", "name_transform": "uppercase", "name_ls": _ATS_LS,
        "header_rule": "1px solid #cccccc", "head_rule": "1px solid #cccccc",
        "head_align": "start", "head_ls": _ATS_LS, "skills_style": "comma",
        "chip_bg": "#f3f4f6", "chip_border": "#d1d5db", "chip_text": "#374151", "chip_radius": 3,
    },
    "banner": {
        "layout": "banner", "font": "sans", "accent": "#27317a",
        "banner_bg": "#27317a", "banner_text": "#ffffff", "head_ls": _ATS_LS,
        "skills_style": "comma",
        "chip_bg": "#eaecf6", "chip_border": "#c7cde8", "chip_text": "#27317a", "chip_radius": 4,
    },
    "onyx": {
        "layout": "sidebar", "font": "sans", "accent": "#0f766e",
        "sidebar_bg": "#1f2937", "sidebar_text": "#ffffff", "name_in_sidebar": True,
        "side_head_color": "#ffffff", "side_rule": "rgba(255,255,255,0.28)",
        "side_width": "34%", "skills_style": "list",
        "chip_bg": "#eef7f5", "chip_border": "#cfe8e3", "chip_text": "#155e54", "chip_radius": 4,
    },
    "sidebar": {
        "layout": "sidebar", "font": "sans", "accent": "#2563eb",
        "sidebar_bg": "#eef2f7", "sidebar_text": "#33415a", "name_in_sidebar": False,
        "side_head_color": "#2563eb", "side_rule": "#cbd5e1", "main_name_color": "#1c2330",
        "side_width": "33%", "skills_style": "list",
        "chip_bg": "#e8effc", "chip_border": "#c5d8f7", "chip_text": "#2563eb", "chip_radius": 4,
    },
}
_TEMPLATE_DEFAULT = _TEMPLATE_CFG["modern_professional"]

# Ids of the two designs the xhtml2pdf engine drew before the gallery existed.
# Drafts saved then still carry them; each is drawn by its gallery twin (the
# preview maps them the same way in cv-builder.service.ts).
_LEGACY_TEMPLATE_IDS = {"minimalist": "minimal", "executive": "classic"}


def _template_id(raw) -> str:
    """A template id _TEMPLATE_CFG knows: legacy ids mapped, unknown ids -> the default."""
    t = str(raw or "").strip().lower()
    t = _LEGACY_TEMPLATE_IDS.get(t, t)
    return t if t in _TEMPLATE_CFG else "modern_professional"

# Résumé section labels for all six supported languages. The document language
# is chosen explicitly by the user (doc_lang); 'auto' falls back to Arabic when
# Arabic content is detected, otherwise English.
# The last five keys are not section headings: they are the words the renderer
# writes when a field is blank ("Role", "Degree", "Your Full Name"), the
# open-ended date ("Present") and the GPA prefix. Each was a hard-coded English
# string, so an Arabic CV with a current job printed "2020 - Present".
# Mirrors CV_LABELS in cv-builder.service.ts; the two must stay in step.
_PREMIUM_LABELS = {
    "en": {"Profile": "Profile", "Experience": "Experience", "Education": "Education", "Skills": "Skills", "Soft Skills": "Soft Skills", "Languages": "Languages", "Contact": "Contact", "Present": "Present", "Role": "Role", "Degree": "Degree", "Your Full Name": "Your Full Name", "GPA": "GPA"},
    "ar": {"Profile": "نبذة", "Experience": "الخبرات", "Education": "التعليم", "Skills": "المهارات", "Soft Skills": "المهارات الشخصية", "Languages": "اللغات", "Contact": "معلومات التواصل", "Present": "حتى الآن", "Role": "المسمى الوظيفي", "Degree": "الدرجة العلمية", "Your Full Name": "اسمك الكامل", "GPA": "المعدل"},
    "es": {"Profile": "Perfil", "Experience": "Experiencia", "Education": "Educación", "Skills": "Habilidades", "Soft Skills": "Habilidades blandas", "Languages": "Idiomas", "Contact": "Contacto", "Present": "Actualidad", "Role": "Puesto", "Degree": "Título", "Your Full Name": "Tu nombre completo", "GPA": "Nota media"},
    "fr": {"Profile": "Profil", "Experience": "Expérience", "Education": "Formation", "Skills": "Compétences", "Soft Skills": "Compétences personnelles", "Languages": "Langues", "Contact": "Contact", "Present": "À ce jour", "Role": "Poste", "Degree": "Diplôme", "Your Full Name": "Votre nom complet", "GPA": "Moyenne"},
    "hi": {"Profile": "प्रोफ़ाइल", "Experience": "अनुभव", "Education": "शिक्षा", "Skills": "कौशल", "Soft Skills": "सॉफ्ट स्किल्स", "Languages": "भाषाएँ", "Contact": "संपर्क", "Present": "वर्तमान", "Role": "पद", "Degree": "डिग्री", "Your Full Name": "आपका पूरा नाम", "GPA": "जीपीए"},
    "de": {"Profile": "Profil", "Experience": "Berufserfahrung", "Education": "Ausbildung", "Skills": "Kenntnisse", "Soft Skills": "Soft Skills", "Languages": "Sprachen", "Contact": "Kontakt", "Present": "heute", "Role": "Position", "Degree": "Abschluss", "Your Full Name": "Ihr vollständiger Name", "GPA": "Note"},
    "pt": {"Profile": "Perfil", "Experience": "Experiência profissional", "Education": "Formação acadêmica", "Skills": "Habilidades", "Soft Skills": "Competências comportamentais", "Languages": "Idiomas", "Contact": "Contato", "Present": "atual", "Role": "Cargo", "Degree": "Curso", "Your Full Name": "Seu nome completo", "GPA": "Média"},
    "id": {"Profile": "Profil", "Experience": "Pengalaman kerja", "Education": "Pendidikan", "Skills": "Keahlian", "Soft Skills": "Soft skill", "Languages": "Bahasa", "Contact": "Kontak", "Present": "sekarang", "Role": "Posisi", "Degree": "Gelar", "Your Full Name": "Nama lengkap Anda", "GPA": "IPK"},
    "it": {"Profile": "Profilo", "Experience": "Esperienza professionale", "Education": "Istruzione e formazione", "Skills": "Competenze", "Soft Skills": "Competenze trasversali", "Languages": "Lingue", "Contact": "Contatti", "Present": "in corso", "Role": "Ruolo", "Degree": "Titolo di studio", "Your Full Name": "Nome e cognome", "GPA": "Voto"},
    "pl": {"Profile": "Profil zawodowy", "Experience": "Doświadczenie zawodowe", "Education": "Wykształcenie", "Skills": "Umiejętności", "Soft Skills": "Umiejętności miękkie", "Languages": "Języki obce", "Contact": "Kontakt", "Present": "obecnie", "Role": "Stanowisko", "Degree": "Kierunek", "Your Full Name": "Imię i nazwisko", "GPA": "Średnia"},
    "tr": {"Profile": "Profil", "Experience": "İş deneyimi", "Education": "Eğitim", "Skills": "Yetenekler", "Soft Skills": "Kişisel beceriler", "Languages": "Yabancı diller", "Contact": "İletişim", "Present": "devam ediyor", "Role": "Pozisyon", "Degree": "Bölüm", "Your Full Name": "Adınız soyadınız", "GPA": "Not ortalaması"},
}


# ── Bundled fonts ────────────────────────────────────────────────
# The CV names only fonts it ships in Fonts/cv/ (licences next to them), so a
# PDF no longer depends on what the server happens to have installed. It used
# to ask for 'Cairo', 'Noto Naskh Arabic', 'Amiri' and got whatever fontconfig
# had: with the fonts setup_arabic_fonts.sh installs, an Arabic CV came out in
# Noto Naskh with every Latin word in Amiri Bold, and the serif templates came
# out bold from top to bottom (only Amiri-Bold was installed).
#
# The Arabic faces were picked by what an ATS gets back from the PDF, measured
# with pdftotext and PyMuPDF on a two-page Arabic CV (298 Arabic words, among
# them «عبد الله», «سمير», «نصر», «محمد عبد الرحمن»): Noto Kufi Arabic and Noto
# Naskh Arabic give back 100% of the words once the ligature named in
# `features` is off; Cairo gives «سمري» for «سمير» and «نرص» for «نصر» whatever
# is switched off. The only visible cost is the one-glyph ligature for «الله»,
# which is then spelled out; lam-alef is unaffected.
#
# The families carry their own names so a system font of the same name can
# never stand in for them. The Angular preview declares the same families with
# WOFF2 files cut from these TTFs (frontend public/assets/fonts/cv/cv-fonts.css), and the stacks below
# are byte-identical to cvFontStack in cv-builder.service.ts.
_CV_FONT_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, '..', 'Fonts', 'cv'))
_CV_FACES = {
    # family: ([(file, weight, style)], @font-face font-feature-settings or None)
    "Confileo CV Sans": ([("DejaVuSans.ttf", 400, "normal"), ("DejaVuSans-Bold.ttf", 700, "normal"),
                          ("DejaVuSans-Oblique.ttf", 400, "italic"), ("DejaVuSans-BoldOblique.ttf", 700, "italic")], None),
    "Confileo CV Serif": ([("LiberationSerif-Regular.ttf", 400, "normal"), ("LiberationSerif-Bold.ttf", 700, "normal"),
                           ("LiberationSerif-Italic.ttf", 400, "italic"), ("LiberationSerif-BoldItalic.ttf", 700, "italic")], None),
    "Confileo CV Kufi": ([("NotoKufiArabic-Regular.ttf", 400, "normal"), ("NotoKufiArabic-Bold.ttf", 700, "normal")], '"rlig" 0'),
    "Confileo CV Naskh": ([("NotoNaskhArabic-Regular.ttf", 400, "normal"), ("NotoNaskhArabic-Bold.ttf", 700, "normal")], '"liga" 0'),
    "Confileo CV Deva Sans": ([("NotoSansDevanagari-Regular.ttf", 400, "normal"), ("NotoSansDevanagari-Bold.ttf", 700, "normal")], None),
    "Confileo CV Deva Serif": ([("NotoSerifDevanagari-Regular.ttf", 400, "normal"), ("NotoSerifDevanagari-Bold.ttf", 700, "normal")], None),
}

# The Font menu: "" keeps the template's own face; "sans" and "serif" switch
# the design, each in the faces of the CV's script. Anything else, such as a
# value saved by the old 18-font menu (none of those fonts was installed on the
# server, so every one of them already printed as the template default), is
# the template default. Amiri is not offered: in the same measurement it gave
# back 94% of the words to pdftotext («الإسكندرية», «الإلكتروني», «بخبرة»
# came out misspelled) with any combination of its ligatures switched off.


def _font_stack(picked: str, is_rtl: bool, lang: str, kind: str) -> str:
    """Font stack for a rendered CV: the script's face first, then the face its
    other script falls back to (Latin in an Arabic CV, Arabic in a Latin one).

    Byte-identical to CvBuilderService.cvFontStack in the Angular preview.
    """
    pick = str(picked or "").strip()
    if pick in ("sans", "serif"):
        kind = pick
    sans = kind == "sans"
    if is_rtl:
        fams = ["Confileo CV Kufi", "Confileo CV Sans"] if sans else ["Confileo CV Naskh", "Confileo CV Serif"]
    elif lang == "hi":
        fams = ["Confileo CV Deva Sans", "Confileo CV Sans"] if sans else ["Confileo CV Deva Serif", "Confileo CV Serif"]
    else:
        fams = ["Confileo CV Sans", "Confileo CV Kufi"] if sans else ["Confileo CV Serif", "Confileo CV Naskh"]
    # DejaVu Sans closes every stack: it has the symbols people type into a CV
    # (✓ → ★ ☎ ✉ •) that the serif faces lack.
    if "Confileo CV Sans" not in fams:
        fams.append("Confileo CV Sans")
    return ", ".join(f"'{f}'" for f in fams) + (", sans-serif" if sans else ", serif")


def _font_faces_css(stack: str) -> str:
    """@font-face rules (file:// URLs into Fonts/cv/) for the families a stack names."""
    from pathlib import Path
    css = ""
    for family, (faces, features) in _CV_FACES.items():
        if f"'{family}'" not in stack:
            continue
        for file, weight, style in faces:
            url = Path(os.path.join(_CV_FONT_DIR, file)).as_uri()
            css += ("@font-face{font-family:'" + family + "';src:url('" + url + "');"
                    "font-weight:" + str(weight) + ";font-style:" + style + ";"
                    + (("font-feature-settings:" + features + ";") if features else "") + "}")
    return css


_EMOJI = re.compile("[\U0001F000-\U0001FAFF\uFE0F]")


def _cv_language(cv: dict, data: dict):
    """(language the CV is set in, right-to-left?). The user's pick wins; 'auto'
    is Arabic when the name, title, summary, skills or job titles hold Arabic."""
    personal = cv["personal"] if isinstance(cv["personal"], dict) else {}
    skills = cv["skills"] if isinstance(cv["skills"], list) else []
    all_text = " ".join([personal.get("full_name", ""), personal.get("job_title", ""), str(cv["summary"] or ""), str(skills)])
    for exp in cv["experience"] if isinstance(cv["experience"], list) else []:
        all_text += " " + exp.get("title", "") + " " + exp.get("company", "")
    picked = str(data.get("doc_lang", data.get("docLang", "auto")) or "auto")
    eff_lang = picked if (picked != "auto" and picked in _PREMIUM_LABELS) else ("ar" if _has_arabic(all_text) else "en")
    return eff_lang, eff_lang == "ar"


def _build_premium_html(data: dict) -> str:
    cv = _normalize_cv_data(data)
    personal = cv["personal"] if isinstance(cv["personal"], dict) else {}
    summary = cv["summary"] or ""
    experience = cv["experience"] if isinstance(cv["experience"], list) else []
    education = cv["education"] if isinstance(cv["education"], list) else []
    skills = cv["skills"] if isinstance(cv["skills"], list) else []
    soft_skills = cv["soft_skills"] if isinstance(cv["soft_skills"], list) else []
    languages = cv["languages"] if isinstance(cv["languages"], list) else []
    body_size = cv["font_size"]
    cfg = _TEMPLATE_CFG[_template_id(cv.get("template_id"))]
    layout = cfg.get("layout", "standard")
    accent = cfg["accent"]
    font_kind = cfg.get("font", "sans")

    eff_lang, is_rtl = _cv_language(cv, data)
    direction = "rtl" if is_rtl else "ltr"

    # Arabic is a JOINED script. `letter-spacing` inserts gaps between letters
    # that are supposed to share a connecting stroke, so a heading like
    # "الخبرات" comes out visibly torn apart — to an Arabic reader it looks
    # like corrupted text, not like styling. The Latin templates use tracking
    # on names and section headings deliberately, so this is not a value we can
    # simply drop from the themes: it has to be neutralised per-render, and
    # only when the document is actually being set in Arabic.
    def ls(value: str) -> str:
        return "normal" if is_rtl else value
    text_align = "right" if is_rtl else "left"
    hdr_align = "center" if cfg.get("header_align") == "center" else text_align
    head_align_css = "center" if cfg.get("head_align") == "center" else text_align
    pad_side = "right" if is_rtl else "left"
    end_side = "left" if is_rtl else "right"

    body_font = _font_stack(cv.get("font_family", ""), is_rtl, eff_lang, font_kind)

    def E(t):
        # WeasyPrint/Pango shapes Arabic itself; a right-to-left CV isolates the
        # Latin runs and numbers inside its text (see _isolate_ltr_runs).
        return _esc(_isolate_ltr_runs(t) if is_rtl and t else t)

    def T(t):
        """Escape a SHORT standalone token (a skill chip, a contact line, a
        date, a language level) and isolate it from the paragraph direction.

        Inside an RTL document the bidi algorithm resolves punctuation at the
        two ends of a run using the surrounding direction, so a Latin skill
        written ".NET / C#" is laid out as "#NET / C." — the leading dot and
        the trailing hash swap ends. It is not a font or shaping problem; the
        characters are simply reordered, and it lands on exactly the tokens a
        CV is full of: ".NET", "C++", "Node.js", "(AWS)", "+20 100…".

        U+2068 FIRST STRONG ISOLATE starts a run whose direction is taken from
        its own first strong character, and U+2069 ends it — so a Latin token
        is laid out LTR and an Arabic one RTL, each correct, without the
        surrounding paragraph leaking in. These are Cf format characters:
        Pango consumes them, and they can never render as a visible box the
        way a missing glyph would.
        """
        s = _esc(t)
        return f"⁨{s}⁩" if s.strip() else s

    def lbl(en):
        return _PREMIUM_LABELS.get(eff_lang, _PREMIUM_LABELS["en"]).get(en, en)

    full_name = E(personal.get("full_name", "")) or E(lbl("Your Full Name"))
    job_title = E(personal.get("job_title", ""))
    contact_inline = _build_contact_html(personal, isolate=True)

    # WeasyPrint implements text-transform:uppercase with str.upper(), which knows no
    # language: a Turkish CV came out with «İŞ DENEYIMI» / «EĞITIM» (dotless I) instead
    # of «İŞ DENEYİMİ» / «EĞİTİM». For Turkish, the text the stylesheet sets in capitals
    # (section headings, the name, the role line) is capitalised here with the Turkish
    # rule — the CSS leaves text that is already in capitals as it is. Applied to the raw
    # text BEFORE escaping, so entities such as &amp; are never uppercased.
    def tu(raw):
        return raw.replace("i", "İ").replace("ı", "I").upper() if eff_lang == "tr" else raw
    name_caps = E(tu(personal.get("full_name", ""))) or E(tu(lbl("Your Full Name")))
    title_caps = E(tu(personal.get("job_title", "")))

    def contact_items():
        items = []
        if personal.get("phone"):
            items.append(T(personal["phone"]))
        if personal.get("email"):
            items.append(T(personal["email"]))
        if personal.get("location"):
            items.append(T(personal["location"]))
        for key in ("linkedin", "github", "portfolio", "website"):
            if personal.get(key):
                items.append(T(personal[key]))
        return items

    def section(title, inner):
        return f'<div class="sec"><div class="sec-h">{tu(lbl(title))}</div>{inner}</div>'

    def entry(title, meta, dates, bullets):
        bs = [c for c in (_clean_line(str(b)) for b in (bullets or []) if b is not None) if c]
        bl = ('<ul>' + ''.join(f'<li>{E(b)}</li>' for b in bs) + '</ul>') if bs else ''
        m = f'<div class="meta">{meta}</div>' if meta else ''
        d = f'<td class="dt">{dates}</td>' if dates else ''
        return f'<div class="entry"><table><tr><td class="ti">{title}</td>{d}</tr></table>{m}{bl}</div>'

    def chips(items):
        return ''.join(f'<span class="chip">{T(s)}</span>' for s in items if str(s).strip())

    def skill_list(items):
        return ''.join(f'<div class="sk-item">{T(s)}</div>' for s in items if str(s).strip())

    def skill_comma(items):
        # One isolated run per skill, joined by the Arabic comma in a right-to-left
        # CV: ", " between two Latin skills used to glue the whole list into one
        # left-to-right run with a stray comma at its far end.
        sep = '، ' if is_rtl else ', '
        return '<p class="sum">' + sep.join((T(s) if is_rtl else E(s)) for s in items if str(s).strip()) + '</p>'

    skills_style = cfg.get("skills_style", "chips")

    def render_skills(items):
        if skills_style == "chips":
            return chips(items)
        if skills_style == "list":
            return skill_list(items)
        return skill_comma(items)

    def lang_inline_html():
        lt = _languages_inline(languages, isolate=is_rtl)
        return f'<p class="sum">{lt}</p>' if lt else ''

    def lang_list_html():
        out = ''
        for lang in languages:
            name = E(str(lang.get("name", ""))).strip()
            if not name:
                continue
            level = str(lang.get("level", "")).strip()
            if level and level not in ("-", "–", "—"):
                out += f'<div class="sk-item">{name} <span class="lvl">({T(level)})</span></div>'
            else:
                out += f'<div class="sk-item">{name}</div>'
        return out

    def experience_rows():
        return ''.join(entry(
            E(exp.get("title", "")) or E(lbl("Role")),
            " &middot; ".join(x for x in [E(_clean_line(exp.get("company", ""))), E(_clean_line(exp.get("location", "")))] if x),
            _fmt_dates(E(exp.get("start_date", "") or exp.get("startDate", "")), E(exp.get("end_date", "") or exp.get("endDate", "")), lbl("Present")),
            exp.get("description", [])) for exp in experience)

    def education_rows():
        def edu_row(edu):
            parts = [E(_clean_line(edu.get("institution", ""))), E(_clean_line(edu.get("location", "")))]
            if edu.get("gpa"):
                parts.append(E(lbl("GPA")) + ": " + E(str(edu["gpa"])))
            return entry(E(edu.get("degree", "")) or E(lbl("Degree")), " &middot; ".join(x for x in parts if x),
                         _fmt_dates(E(edu.get("start_date", "") or edu.get("startDate", "")), E(edu.get("end_date", "") or edu.get("endDate", "")), lbl("Present")), [])
        return ''.join(edu_row(e) for e in education)

    # ── Main column (profile / experience / education) ──────────────
    main_inner = ''
    if summary and str(summary).strip():
        main_inner += section("Profile", f'<p class="sum">{E(summary)}</p>')
    if experience:
        main_inner += section("Experience", experience_rows())
    if education:
        main_inner += section("Education", education_rows())

    # ── Aside content (skills / soft skills / languages) ────────────
    aside_inner = ''
    if skills:
        aside_inner += section("Skills", render_skills(skills))
    if soft_skills:
        aside_inner += section("Soft Skills", render_skills(soft_skills))
    if languages:
        lang_html = lang_list_html() if layout == "sidebar" else lang_inline_html()
        if lang_html:
            aside_inner += section("Languages", lang_html)

    # ── Shared item-level CSS ───────────────────────────────────────
    chip_css = (".chip{display:inline-block;border:1px solid " + cfg.get("chip_border", "#ccc")
                + ";background:" + cfg.get("chip_bg", "#f3f4f6") + ";color:" + cfg.get("chip_text", "#333")
                + ";border-radius:" + str(cfg.get("chip_radius", 4)) + "px;padding:3px 11px;margin:0 6px 7px 0;font-size:0.86em;}")
    common_css = (
        "*{box-sizing:border-box;}"
        "a{color:inherit;text-decoration:none;}"
        "table{border-collapse:collapse;width:100%;}"
        "td{padding:0;vertical-align:top;}"
        ".sec{margin-bottom:14px;}"
        ".sum{font-size:0.96em;line-height:1.6;color:#333;margin:0;text-align:justify;}"
        # Entries flow/split across pages (do NOT keep-together) so a long entry
        # fills the page instead of jumping wholesale to the next page and
        # leaving a big empty gap. Headings stay glued to the content that
        # follows them; paragraphs/bullets avoid 1-line orphans/widows.
        ".entry{margin-bottom:10px;}"
        ".sec-h{page-break-after:avoid;break-after:avoid;}"
        ".entry table{page-break-after:avoid;break-after:avoid;}"
        # The company line stays with the first bullet as well: WeasyPrint already moved a
        # job whose bullets did not fit to the next page rather than leave its title and
        # company line alone at the foot of a page; saying so lets the preview do the same.
        ".meta{page-break-after:avoid;break-after:avoid;}"
        "p,li{orphans:2;widows:2;}"
        ".entry .ti{font-weight:700;color:#1c2330;font-size:1.03em;text-align:" + text_align + ";}"
        ".dt{font-size:0.84em;color:#777;white-space:nowrap;width:130px;text-align:" + end_side + ";}"
        # The company line is italic only in scripts that have italics. Arabic
        # and Devanagari faces have none: the browser slanted them by force in
        # the preview, and WeasyPrint named a fake "…-Oblique" face in the PDF
        # but printed it upright, so the two never matched.
        ".meta{font-size:0.9em;color:#666;font-style:" + ("normal" if is_rtl or eff_lang == "hi" else "italic") + ";margin-top:1px;}"
        "ul{margin:4px 0 0;padding-" + pad_side + ":16px;}"
        "li{font-size:0.95em;line-height:1.45;color:#333;margin-bottom:2px;}"
        ".sk-item{font-size:0.92em;margin-bottom:5px;line-height:1.4;}"
        + chip_css
    )

    # ── Layout assembly ─────────────────────────────────────────────
    if layout == "sidebar":
        sidebar_bg = cfg.get("sidebar_bg", "#1f2937")
        sidebar_text = cfg.get("sidebar_text", "#ffffff")
        side_width = cfg.get("side_width", "33%")
        side_head_color = cfg.get("side_head_color", sidebar_text)
        side_rule = cfg.get("side_rule", "rgba(255,255,255,0.28)")
        main_name_color = cfg.get("main_name_color", "#1c2330")
        # The panel's side follows is_rtl, NOT is_ar (the table takes the body's
        # direction, the colour strip below says it outright): with the CV
        # language picked explicitly the two can disagree — an Arabic-labelled
        # CV whose content is all Latin, or an English CV with an Arabic name —
        # and the panel then landed on the opposite side from the text
        # direction, which the preview never did.
        name_in_side = cfg.get("name_in_sidebar", False)

        side = ''
        if name_in_side:
            side += f'<div class="sb-name">{name_caps}</div>'
            if job_title:
                side += f'<div class="sb-role">{job_title}</div>'
            side += '<div class="sec"><div class="sb-contact">' + ''.join(f'<div class="ci">{c}</div>' for c in contact_items()) + '</div></div>'
        else:
            side += section("Contact", '<div class="sb-contact">' + ''.join(f'<div class="ci">{c}</div>' for c in contact_items()) + '</div>')
        side += aside_inner

        main = ''
        if not name_in_side:
            main += f'<div class="main-name">{name_caps}</div>'
            if job_title:
                main += f'<div class="main-role">{title_caps}</div>'
        main += main_inner

        # Two cells of one table row: WeasyPrint 70 splits a row across pages
        # with both columns continuing side by side. The panel used to be an
        # absolutely positioned box, which lives on page 1 only: whatever did
        # not fit there was dropped without a trace (a CV with 45 skills, 15
        # soft skills and 8 languages printed 39 skills and nothing else).
        # A float does not split across pages either, and flex put the main
        # column below the panel. The panel's colour is the page background,
        # so it runs edge to edge on every page, through the page margins.
        # Pages after the first get a 12 mm top margin (the second page began
        # 0.5 mm from the paper's edge) and every page a 12 mm bottom margin.
        strip = ("linear-gradient(to " + ("left" if is_rtl else "right") + ", " + sidebar_bg + " " + side_width
                 + ", #ffffff " + side_width + ")")
        css = common_css + (
            "@page{size:A4;margin:12mm 0;background:" + strip + ";}"
            "@page:first{margin-top:0;}"
            "body{font-family:" + body_font + ";color:#1f2430;font-size:" + str(body_size) + "px;line-height:1.5;margin:0;direction:" + direction + ";text-align:" + text_align + ";}"
            ".cols{table-layout:fixed;}"
            ".sidebar{width:" + side_width + ";padding:30px 22px;color:" + sidebar_text + ";}"
            ".main{padding:30px 28px;}"
            ".sb-name{font-size:1.7em;font-weight:800;color:" + sidebar_text + ";line-height:1.12;margin:0 0 4px;text-transform:uppercase;letter-spacing:" + ls("0.5px") + ";}"
            ".sb-role{font-size:0.9em;color:" + sidebar_text + ";opacity:0.85;margin-bottom:14px;}"
            ".main-name{font-size:2.05em;font-weight:800;color:" + main_name_color + ";line-height:1.1;margin:0;text-transform:uppercase;letter-spacing:" + ls("0.5px") + ";}"
            ".main-role{font-size:0.98em;color:#666;margin:3px 0 14px;text-transform:uppercase;letter-spacing:" + ls(_ATS_LS) + ";}"
            ".sidebar .sec-h{font-size:0.84em;font-weight:700;text-transform:uppercase;letter-spacing:" + ls(_ATS_LS) + ";color:" + side_head_color + ";border-bottom:1px solid " + side_rule + ";padding-bottom:3px;margin-bottom:8px;}"
            ".sidebar .sk-item,.sidebar .ci,.sidebar .sum{color:" + sidebar_text + ";opacity:0.93;}"
            ".sidebar .ci{margin-bottom:5px;font-size:0.85em;word-break:break-word;}"
            ".sidebar .lvl{opacity:0.7;}"
            ".main .sec-h{font-size:0.85em;font-weight:700;text-transform:uppercase;letter-spacing:" + ls(_ATS_LS) + ";color:" + accent + ";border-bottom:1.5px solid " + accent + ";padding-bottom:3px;margin-bottom:9px;}"
        )
        body_html = ('<table class="cols"><tr><td class="sidebar">' + side + '</td>'
                     '<td class="main">' + main + '</td></tr></table>')

    elif layout == "banner":
        banner_bg = cfg.get("banner_bg", accent)
        banner_text = cfg.get("banner_text", "#ffffff")
        head_ls = cfg.get("head_ls", _ATS_LS)
        # The band sits on the paper's top edge on page 1 only. Every page keeps
        # 12 mm at the bottom and the next pages 12 mm at the top: with margin 0
        # page 2 began 0.5 mm from the edge (-0.6 mm on an Arabic CV, cut off)
        # and page 1 ended 0.5 mm from the bottom.
        css = common_css + (
            "@page{size:A4;margin:12mm 0;}"
            "@page:first{margin-top:0;}"
            "body{font-family:" + body_font + ";color:#1f2430;font-size:" + str(body_size) + "px;line-height:1.5;margin:0;direction:" + direction + ";text-align:" + text_align + ";}"
            ".banner{background:" + banner_bg + ";color:" + banner_text + ";padding:26px 40px;}"
            ".banner .name{font-size:2.3em;font-weight:800;color:" + banner_text + ";letter-spacing:" + ls("0.5px") + ";margin:0;line-height:1.06;text-transform:uppercase;}"
            ".banner .role{font-size:1em;font-weight:600;text-transform:uppercase;letter-spacing:" + ls(_ATS_LS) + ";color:" + banner_text + ";opacity:0.92;margin-top:5px;}"
            ".banner .contact{font-size:0.82em;color:" + banner_text + ";opacity:0.9;margin-top:9px;line-height:1.7;}"
            ".bd{padding:20px 40px 30px;}"
            ".sec-h{font-size:0.85em;font-weight:700;text-transform:uppercase;letter-spacing:" + ls(head_ls) + ";color:" + accent + ";border-bottom:1.5px solid " + accent + ";padding-bottom:3px;margin-bottom:9px;}"
        )
        header = (f'<div class="banner"><h1 class="name">{name_caps}</h1>'
                  + (f'<div class="role">{title_caps}</div>' if job_title else '')
                  + f'<div class="contact">{contact_inline}</div></div>')
        body_html = header + '<div class="bd">' + main_inner + aside_inner + '</div>'

    else:  # standard / serif single column
        page_bg = cfg.get("page_bg", "#ffffff")
        header_rule = cfg.get("header_rule", "2.5px solid " + accent)
        head_rule = cfg.get("head_rule", "1.3px solid " + accent)
        name_transform = cfg.get("name_transform", "none")
        name_ls = cfg.get("name_ls", "0.5px")
        head_ls = cfg.get("head_ls", _ATS_LS)
        # The short underline (elegant) is drawn by a heading as wide as its text. It was an
        # inline-block, and break-after:avoid does not apply to an inline box: until 10 Oct
        # 2026 such a CV could end page 1 on "SOFT SKILLS" or "EDUCATION" with the section
        # on page 2. A table is as wide as its text too, and is a block that stays with what follows.
        sec_h_display = "table" if cfg.get("head_underline") == "short" else "block"
        if cfg.get("double_top"):
            hdr_rule_css = "border-top:" + header_rule + ";border-bottom:" + header_rule + ";padding:7px 0;"
        else:
            hdr_rule_css = "border-bottom:" + header_rule + ";padding:0 0 10px;"
        page_bg_css = ("background:" + page_bg + ";") if page_bg != "#ffffff" else ""
        css = common_css + (
            "@page{size:A4;margin:" + cfg.get("page_margin", "1.25cm 1.45cm") + ";" + page_bg_css + "}"
            "body{font-family:" + body_font + ";" + page_bg_css + "color:#1f2430;font-size:" + str(body_size) + "px;line-height:1.5;margin:0;direction:" + direction + ";text-align:" + text_align + ";}"
            ".hdr{" + hdr_rule_css + "margin-bottom:15px;text-align:" + hdr_align + ";}"
            ".name{font-size:2.05em;font-weight:800;color:" + accent + ";letter-spacing:" + ls(name_ls) + ";margin:0;line-height:1.1;text-transform:" + name_transform + ";}"
            ".role{font-size:0.94em;font-weight:600;text-transform:uppercase;letter-spacing:" + ls(_ATS_LS) + ";color:#777;margin-top:4px;}"
            ".contact{font-size:0.8em;color:#666;margin-top:7px;line-height:1.7;}"
            ".sec-h{display:" + sec_h_display + ";font-size:0.85em;font-weight:700;text-transform:uppercase;letter-spacing:" + ls(head_ls) + ";color:" + accent + ";border-bottom:" + head_rule + ";padding-bottom:3px;margin-bottom:9px;text-align:" + head_align_css + ";}"
        )
        header = (f'<div class="hdr"><h1 class="name">{name_caps if name_transform == "uppercase" else full_name}</h1>'
                  + (f'<div class="role">{title_caps}</div>' if job_title else '')
                  + f'<div class="contact">{contact_inline}</div></div>')
        body_html = header + main_inner + aside_inner

    # Emoji and pictographs (U+1F000-1FAFF, and the emoji-style selector U+FE0F)
    # have no face among the CV fonts and printed as empty boxes; they are left
    # out, as in the preview. Symbols the fonts do have (✓ → ★ ☎ ✉) stay.
    body_html = _EMOJI.sub("", body_html)
    return ('<!DOCTYPE html><html dir="' + direction + '"><head><meta charset="UTF-8"/><style>'
            + _font_faces_css(body_font) + css + '</style></head><body>' + body_html + _watermark_html(data)
            + '</body></html>')


# ══════════════════════════════════════════════════════════════════
# PDF Generation
# ══════════════════════════════════════════════════════════════════

def export_cv_to_pdf(data: dict) -> bytes:
    """Export CV data to PDF, always with the WeasyPrint template engine.

    There used to be a second engine, xhtml2pdf, for template ids outside
    _TEMPLATE_CFG ("minimalist", "executive", saved by drafts older than the
    gallery) and for any exception inside WeasyPrint. It printed Arabic as
    solid black boxes: xhtml2pdf 0.2.21 refuses to read a font outside its
    allowed folders and falls back to Helvetica. Old ids are now mapped to
    their gallery twins (_template_id), and a WeasyPrint failure is an error the
    API reports, which the page shows with a retry, not a PDF of boxes.
    """
    from weasyprint import HTML
    return HTML(string=_build_premium_html(data)).write_pdf()


# ══════════════════════════════════════════════════════════════════
# Word (.docx) — the editable copy
# ══════════════════════════════════════════════════════════════════
# The Word button used to download the preview's HTML under a .doc name: `file`
# called it "HTML document", python-docx could not open it, and LibreOffice laid
# the two-column templates out as three pages with the name at body size. This
# is a real Word document: one column (what applicant tracking systems read
# best, and what someone edits for the next application), Arial for Latin and
# Arabic (installed with Windows and macOS; Hindi names Nirmala UI), headings
# in the template's colour, real bullet paragraphs, and right-to-left
# paragraphs, runs and section for an Arabic CV.
#
# Direction inside a line: Word decides the order of a Latin run inside Arabic
# text by the Unicode bidi algorithm, and LibreOffice ignores the run's own
# direction flag, so each Latin run (the ones _isolate_ltr_runs finds) is also
# fenced with U+200E LEFT-TO-RIGHT MARK on both sides. That keeps "+20 100 123
# 4567", "C#" and ".NET 8" in order in both programs; LRM is an invisible
# format character every editor knows (the PDF uses U+2068/U+2069, which older
# Word versions may not).

_W_PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
                "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
                "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
                "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
                "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
                "sectPr", "pPrChange"]
_W_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
                "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
                "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
                "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]


def _w_set(parent, tag: str, order: list, **attrs):
    """Put <w:tag attrs/> into a pPr/rPr at its schema position (Word refuses a
    document whose property elements are out of order), replacing one already there."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    old = parent.find(qn("w:" + tag))
    if old is not None:
        parent.remove(old)
    el = OxmlElement("w:" + tag)
    for k, v in attrs.items():
        el.set(qn("w:" + k), str(v))
    pos = order.index(tag)
    for i, child in enumerate(list(parent)):
        name = child.tag.split("}")[-1]
        if name in order and order.index(name) > pos:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def export_cv_to_docx(data: dict) -> bytes:
    from docx import Document
    from docx.shared import Pt, Mm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    cv = _normalize_cv_data(data)
    personal = cv["personal"] if isinstance(cv["personal"], dict) else {}
    summary = str(cv["summary"] or "")
    experience = cv["experience"] if isinstance(cv["experience"], list) else []
    education = cv["education"] if isinstance(cv["education"], list) else []
    skills = cv["skills"] if isinstance(cv["skills"], list) else []
    soft_skills = cv["soft_skills"] if isinstance(cv["soft_skills"], list) else []
    languages = cv["languages"] if isinstance(cv["languages"], list) else []
    cfg = _TEMPLATE_CFG.get(_template_id(cv.get("template_id")), _TEMPLATE_DEFAULT)
    accent = RGBColor.from_string(cfg["accent"].lstrip("#").upper())
    grey = RGBColor(0x66, 0x66, 0x66)
    eff_lang, is_rtl = _cv_language(cv, data)
    labels = _PREMIUM_LABELS.get(eff_lang, _PREMIUM_LABELS["en"])

    def lbl(en):
        return labels.get(en, en)

    def tu(raw):  # Turkish capitals, as in the PDF
        return raw.replace("i", "İ").replace("ı", "I").upper() if eff_lang == "tr" else raw.upper()

    latin_font = "Arial"
    cs_font = "Nirmala UI" if eff_lang == "hi" else "Arial"
    base_pt = max(9.0, min(12.0, float(cv.get("font_size") or 13) * 0.8))  # the PDF's 13 px ≈ 10.5 pt

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(18)
    sec.top_margin = sec.bottom_margin = Mm(16)
    if is_rtl:
        sectPr = sec._sectPr
        grid = sectPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docGrid")
        from docx.oxml import OxmlElement
        bidi = OxmlElement("w:bidi")
        if grid is not None:
            grid.addprevious(bidi)
        else:
            sectPr.append(bidi)

    normal = doc.styles["Normal"]
    normal.font.name = latin_font
    normal.font.size = Pt(base_pt)
    rpr = normal.element.get_or_add_rPr()
    _w_set(rpr, "rFonts", _W_RPR_ORDER, ascii=latin_font, hAnsi=latin_font, cs=cs_font, eastAsia=latin_font)
    _w_set(rpr, "szCs", _W_RPR_ORDER, val=int(base_pt * 2))
    _w_set(rpr, "lang", _W_RPR_ORDER, val="en-US" if eff_lang in ("ar", "hi") else eff_lang, bidi="ar-SA" if is_rtl else "en-US")
    ppr = normal.element.get_or_add_pPr()
    if is_rtl:
        _w_set(ppr, "bidi", _W_PPR_ORDER)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.15

    def runs(p, text, bold=False, size=None, color=None, italic=False, caps=False):
        """Add text to paragraph p as direction runs (see the note above)."""
        text = str(text or "")
        if not text:
            return
        pieces = []  # (text, is_latin_run)
        if is_rtl:
            iso = _isolate_ltr_runs(text)
            for chunk in re.split("(⁨[^⁩]*⁩)", iso):
                if not chunk:
                    continue
                if chunk.startswith("⁨"):
                    pieces.append(("‎" + chunk[1:-1] + "‎", True))
                else:
                    pieces.append((chunk, False))
        else:
            pieces.append((text, True))
        for piece, latin in pieces:
            r = p.add_run(piece)
            r.bold = bold or None
            r.italic = italic or None
            if size:
                r.font.size = Pt(size)
            if color is not None:
                r.font.color.rgb = color
            rp = r._r.get_or_add_rPr()
            if bold:
                _w_set(rp, "bCs", _W_RPR_ORDER)
            if italic:
                _w_set(rp, "iCs", _W_RPR_ORDER)
            if size:
                _w_set(rp, "szCs", _W_RPR_ORDER, val=int(size * 2))
            if is_rtl and not latin:
                _w_set(rp, "rtl", _W_RPR_ORDER)

    def para(style=None, space_before=0, space_after=2, center=False, keep=False):
        p = doc.add_paragraph(style=style)
        pf = p.paragraph_format
        pf.space_before, pf.space_after = Pt(space_before), Pt(space_after)
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if keep:
            pf.keep_with_next = True
        if is_rtl:
            _w_set(p._p.get_or_add_pPr(), "bidi", _W_PPR_ORDER)
        return p

    def heading(en):
        p = para(space_before=10, space_after=4, keep=True)
        runs(p, lbl(en) if is_rtl else tu(lbl(en)), bold=True, size=base_pt + 0.5, color=accent)
        if not is_rtl:
            for r in p.runs:
                _w_set(r._r.get_or_add_rPr(), "spacing", _W_RPR_ORDER, val=20)  # 1 pt tracking, as the PDF's headings
        pbdr = _w_set(p._p.get_or_add_pPr(), "pBdr", _W_PPR_ORDER)
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        bottom = OxmlElement("w:bottom")
        for k, v in (("val", "single"), ("sz", "8"), ("space", "1"), ("color", cfg["accent"].lstrip("#").upper())):
            bottom.set(qn("w:" + k), v)
        pbdr.append(bottom)
        return p

    def dates(start, end):
        start, end = str(start or "").strip(), str(end or "").strip()
        if start and end:
            return f"{start} – {end}"
        if start:
            return f"{start} – {lbl('Present')}"
        return end

    # Between two items: in a right-to-left CV the dot is fenced with U+200F
    # RIGHT-TO-LEFT MARK, so two Latin items (an e-mail and a URL) stay two
    # items read right to left instead of merging into one left-to-right run.
    sep = "\u200f  ·  \u200f" if is_rtl else "  ·  "

    def meta_line(parts):
        parts = [x for x in parts if x]
        if not parts:
            return
        p = para(space_after=2)
        for i, part in enumerate(parts):
            if i:
                runs(p, sep, color=grey, size=base_pt - 0.5)
            runs(p, part, color=grey, size=base_pt - 0.5)

    # ── Header: name, role, contact line ──
    name = str(personal.get("full_name", "")).strip() or lbl("Your Full Name")
    p = para(space_after=1)
    runs(p, name if is_rtl else tu(name) if cfg.get("name_transform") == "uppercase" else name,
         bold=True, size=base_pt * 2 + 1, color=accent)
    role = str(personal.get("job_title", "")).strip()
    if role:
        p = para(space_after=3)
        runs(p, role, bold=True, size=base_pt + 0.5, color=grey)
    contact = [str(personal.get(k, "")).strip() for k in ("phone", "email", "location", "linkedin", "github", "portfolio", "website")]
    contact = [c for c in contact if c]
    if contact:
        p = para(space_after=6)
        for i, c in enumerate(contact):
            if i:
                runs(p, sep, color=grey, size=base_pt - 1)
            runs(p, c, color=grey, size=base_pt - 1)

    if summary.strip():
        heading("Profile")
        runs(para(space_after=2), _clean_line(summary) if "\n" not in summary else summary.strip())

    if experience:
        heading("Experience")
        for exp in experience:
            p = para(space_before=4, space_after=1, keep=True)
            runs(p, _clean_line(exp.get("title", "")) or lbl("Role"), bold=True, size=base_pt + 0.5)
            meta_line([_clean_line(exp.get("company", "")), _clean_line(exp.get("location", "")),
                       dates(exp.get("start_date") or exp.get("startDate"), exp.get("end_date") or exp.get("endDate"))])
            for b in exp.get("description") or []:
                b = _clean_line(str(b or ""))
                if b:
                    runs(para(style="List Bullet", space_after=1), b)

    if education:
        heading("Education")
        for edu in education:
            p = para(space_before=4, space_after=1, keep=True)
            runs(p, _clean_line(edu.get("degree", "")) or lbl("Degree"), bold=True, size=base_pt + 0.5)
            gpa = str(edu.get("gpa", "") or "").strip()
            meta_line([_clean_line(edu.get("institution", "")), _clean_line(edu.get("location", "")),
                       (lbl("GPA") + ": " + gpa) if gpa else "",
                       dates(edu.get("start_date") or edu.get("startDate"), edu.get("end_date") or edu.get("endDate"))])

    comma = "، " if is_rtl else ", "
    for title, items in (("Skills", skills), ("Soft Skills", soft_skills)):
        items = [str(x).strip() for x in items if str(x).strip()]
        if items:
            heading(title)
            p = para(space_after=2)
            for i, it in enumerate(items):
                if i:
                    runs(p, comma)
                runs(p, it)

    langs = []
    for lang in languages:
        n = str(lang.get("name", "")).strip()
        if not n:
            continue
        lv = str(lang.get("level", "")).strip()
        langs.append((n, lv if lv not in ("", "-", "–", "—") else ""))
    if langs:
        heading("Languages")
        p = para(space_after=2)
        for i, (n, lv) in enumerate(langs):
            if i:
                runs(p, sep, color=grey)
            runs(p, n, bold=True)
            if lv:
                runs(p, " (" + lv + ")", color=grey)

    props = doc.core_properties
    props.title = str(cv.get("title") or name)
    props.author = str(personal.get("full_name", "")).strip()
    props.language = "ar-SA" if is_rtl else eff_lang
    props.comments = ""
    out = BytesIO()
    doc.save(out)
    return out.getvalue()


# ══════════════════════════════════════════════════════════════════
# CV Parser — Extract Structured Data from PDF
# ══════════════════════════════════════════════════════════════════

# Section header patterns (case-insensitive) — English + Arabic
SECTION_PATTERNS = {
    "personal_info": r"(?:personal\s*info|contact\s*info|contact\s*details|personal\s*details|contact|personal|المعلومات\s*الشخصية|معلومات\s*الاتصال|معلومات\s*التواصل|بيانات\s*التواصل|بيانات\s*الاتصال|بيانات\s*شخصية|معلومات\s*شخصية|البيانات\s*الشخصية|التواصل)",
    "summary": r"(?:summary|profile|objective|about\s*me|professional\s*summary|career\s*objective|professional\s*profile|الملخص|نبذة\s*عني|الملف\s*الشخصي|نبذة\s*مهنية|الهدف\s*الوظيفي|نبذة\s*عامة|ملخص\s*مهني|نبذة|ملخص|الهدف)",
    "experience": r"(?:experience|work\s*experience|employment|employment\s*history|professional\s*experience|work\s*history|الخبرات|الخبرة|خبرات\s*العمل|الخبرة\s*العملية|الخبرة\s*المهنية|الخبرات\s*العملية|الخبرات\s*المهنية|خبرة\s*عملية|الخبرات\s*الوظيفية|التاريخ\s*الوظيفي|الوظائف\s*السابقة)",
    "education": r"(?:education|academic|qualifications|educational\s*background|academic\s*background|التعليم|المؤهلات|المؤهلات\s*العلمية|الخلفية\s*الأكاديمية|المؤهلات\s*الدراسية|التعليم\s*والمؤهلات|المؤهل\s*العلمي|الدراسة|التعليم\s*والتدريب)",
    "skills": r"(?:skills|technical\s*skills|core\s*competencies|competencies|technologies|tech\s*stack|المهارات|المهارات\s*التقنية|المهارات\s*الفنية|الكفاءات|مهارات\s*تقنية|المهارات\s*المهنية|مهارات)",
    "soft_skills": r"(?:soft\s*skills|interpersonal\s*skills|personal\s*skills|المهارات\s*الشخصية|المهارات\s*الناعمة|مهارات\s*شخصية)",
    "languages": r"(?:languages|language\s*proficiency|language\s*skills|اللغات|المهارات\s*اللغوية|اللغات\s*المتقنة)",
    "projects": r"(?:projects|personal\s*projects|key\s*projects|المشاريع|مشاريع\s*شخصية|مشاريع\s*رئيسية)",
    "certifications": r"(?:certifications?|certificates?|licenses?|professional\s*development|الشهادات|الدورات|الدورات\s*التدريبية|شهادات\s*مهنية|التطوير\s*المهني|الشهادات\s*والدورات)",
}

# The headings this tool itself prints, in all eleven CV languages (_PREMIUM_LABELS),
# as the PDF spells them: as written, in capitals (most templates set headings in
# capitals), and in Turkish capitals (İ, I). Only English and Arabic headings were
# known, so a French CV made here imported with no job, degree, skill or language.
# They are matched as a whole line or before a colon ("Kenntnisse: …"): words such
# as "Bahasa" or "Profil" also begin ordinary lines ("Bahasa Indonesia (native)").
_LABEL_SECTIONS = {"Profile": "summary", "Experience": "experience", "Education": "education",
                   "Skills": "skills", "Soft Skills": "soft_skills", "Languages": "languages",
                   "Contact": "personal_info"}


def _label_patterns() -> dict:
    alts = {sec: set() for sec in _LABEL_SECTIONS.values()}
    for lang, labels in _PREMIUM_LABELS.items():
        for key, sec in _LABEL_SECTIONS.items():
            label = labels[key]
            forms = {label, label.upper()}
            if lang == "tr":
                forms.add(label.replace("i", "İ").replace("ı", "I").upper())
            for form in forms:
                alts[sec].add(r"\s*".join(re.escape(w) for w in form.split()))
    return {sec: "(?:" + "|".join(sorted(a, key=len, reverse=True)) + ")" for sec, a in alts.items()}


_LABEL_PATTERNS = _label_patterns()

_DEVA = re.compile(r"[\u0900-\u097F]")


def _deva_skeleton(text: str) -> str:
    """Devanagari with its vowel signs, virama, nukta and nasal marks removed."""
    return re.sub(r"[\u0900-\u0903\u093A-\u094F\u0951-\u0957\u0962\u0963\s]", "", text)


_DEVA_HEADINGS = {_deva_skeleton(_PREMIUM_LABELS["hi"][k]): sec for k, sec in _LABEL_SECTIONS.items()}
_SPACELESS_PATTERNS = {k: v.replace("\\s*", "") for k, v in SECTION_PATTERNS.items()}
for _sec, _pat in _LABEL_PATTERNS.items():
    _SPACELESS_PATTERNS[_sec] = "(?:" + _SPACELESS_PATTERNS[_sec] + "|" + _pat.replace("\\s*", "") + ")"

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
# Up to seven digit groups: "+33 6 12 34 56 78" has six and was cut to "+33 6 12 34 56".
PHONE_RE = re.compile(r"(?:\+|00)?\d{1,4}(?:[\s.-]?\(?\d{1,5}\)?){1,6}")
URL_RE = re.compile(r"https?://[^\s,;]+|www\.[^\s,;]+")
LINKEDIN_RE = re.compile(r"(?:https?://)?(?:www\.)?linkedin\.com/in/[^\s,;]+", re.IGNORECASE)
GITHUB_RE = re.compile(r"(?:https?://)?(?:www\.)?github\.com/[^\s,;]+", re.IGNORECASE)

_DATE_SEP = r"(?:[-–—]+|\bto\b|إلى)"
_AR_MONTHS = r"(?:يناير|فبراير|مارس|أبريل|مايو|يونيو|يوليو|أغسطس|سبتمبر|أكتوبر|نوفمبر|ديسمبر)"
# Whole month names, then a space, then the year: "Mar 2021" and "March 2021"
# match, "Marketing 2019" (an MBA line) does not.
_EN_MONTHS = (r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?"
              r"|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)")
# Month names of the other CV languages (full and the usual abbreviations), not
# preceded by a letter so "mar" never starts inside a word.
_OTHER_MONTHS = "(?<![^\\W\\d_])(?:" + "|".join(sorted({
    # es
    "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "setiembre",
    "octubre", "noviembre", "diciembre", "ene", "abr", "ago", "dic",
    # fr
    "janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre",
    "novembre", "décembre", "janv", "févr", "fév", "avr", "juil", "déc",
    # de
    "januar", "februar", "märz", "juni", "juli", "oktober", "dezember", "mär", "okt", "dez",
    # pt
    "janeiro", "fevereiro", "março", "maio", "junho", "julho", "setembro", "outubro", "novembro",
    "dezembro", "fev", "set", "out",
    # it
    "gennaio", "febbraio", "aprile", "maggio", "giugno", "luglio", "settembre", "ottobre", "dicembre",
    "gen", "mag", "giu", "lug", "ott",
    # pl (nominative and genitive)
    "styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec", "lipiec", "sierpień", "wrzesień",
    "październik", "listopad", "grudzień", "stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca",
    "lipca", "sierpnia", "września", "października", "listopada", "grudnia",
    "sty", "lut", "kwi", "cze", "lip", "sie", "wrz", "paź", "lis", "gru",
    # tr
    "ocak", "şubat", "mart", "nisan", "mayıs", "haziran", "temmuz", "ağustos", "eylül", "ekim", "kasım",
    "aralık", "oca", "şub", "nis", "haz", "tem", "ağu", "eyl", "eki", "kas", "ara",
    # id
    "januari", "februari", "maret", "mei", "agustus", "desember", "agu", "agus", "des",
    # hi
    "जनवरी", "फ़रवरी", "फरवरी", "मार्च", "अप्रैल", "मई", "जून", "जुलाई", "अगस्त", "सितंबर", "सितम्बर",
    "अक्टूबर", "नवंबर", "नवम्बर", "दिसंबर", "दिसम्बर",
}, key=len, reverse=True)) + ")"
# The open end of a range as the CV languages write it (the renderer's own
# "Present" label in each language first).
_PRESENT_WORDS = "|".join(sorted({re.escape(v["Present"]) for v in _PREMIUM_LABELS.values()} | {
    "present", "current", "now", "presente", "actual", "actualmente", "présent", "aujourd'hui", "heute",
    "atualmente", "oggi", "halen", "günümüz", "saat ini"}, key=len, reverse=True))
DATE_RE = re.compile(
    rf"(?:(?:{_EN_MONTHS}|{_AR_MONTHS}|{_OTHER_MONTHS})\.?\s+\d{{4}}|"
    rf"\d{{1,2}}/\d{{4}}|\d{{4}})\s*{_DATE_SEP}\s*"
    rf"(?:(?:{_EN_MONTHS}|{_AR_MONTHS}|{_OTHER_MONTHS})\.?\s+\d{{4}}|"
    rf"\d{{1,2}}/\d{{4}}|\d{{4}}|{_PRESENT_WORDS}|حتى\s*الآن|الحالي|مستمر|حالي[اً]?)",
    re.IGNORECASE,
)

_KNOWN_HEADERS = {
    "WORKEXPERIENCE": "WORK EXPERIENCE",
    "EMPLOYMENTHISTORY": "EMPLOYMENT HISTORY",
    "PROFESSIONALEXPERIENCE": "PROFESSIONAL EXPERIENCE",
    "WORKHISTORY": "WORK HISTORY",
    "SOFTSKILLS": "SOFT SKILLS",
    "TECHNICALSKILLS": "TECHNICAL SKILLS",
    "CORECOMPETENCIES": "CORE COMPETENCIES",
    "PERSONALINFO": "PERSONAL INFO",
    "CONTACTINFO": "CONTACT INFO",
    "CONTACTDETAILS": "CONTACT DETAILS",
    "PERSONALDETAILS": "PERSONAL DETAILS",
    "LANGUAGEPROFICIENCY": "LANGUAGE PROFICIENCY",
    "LANGUAGESKILLS": "LANGUAGE SKILLS",
    "PROFESSIONALSUMMARY": "PROFESSIONAL SUMMARY",
    "CAREEROBJECTIVE": "CAREER OBJECTIVE",
    "PROFESSIONALPROFILE": "PROFESSIONAL PROFILE",
    "EDUCATIONALBACKGROUND": "EDUCATIONAL BACKGROUND",
    "ACADEMICBACKGROUND": "ACADEMIC BACKGROUND",
    "PERSONALPROJECTS": "PERSONAL PROJECTS",
    "KEYPROJECTS": "KEY PROJECTS",
    "PROFESSIONALDEVELOPMENT": "PROFESSIONAL DEVELOPMENT",
    "ABOUTME": "ABOUT ME",
    "INTERPERSONALSKILLS": "INTERPERSONAL SKILLS",
    "PERSONALSKILLS": "PERSONAL SKILLS",
    "المعلوماتالشخصية": "المعلومات الشخصية",
    "بياناتشخصية": "بيانات شخصية",
    "معلوماتالاتصال": "معلومات الاتصال",
    "الخبرةالعملية": "الخبرة العملية",
    "الخبرةالمهنية": "الخبرة المهنية",
    "خبراتالعمل": "خبرات العمل",
    "الخبراتالعملية": "الخبرات العملية",
    "الخبراتالمهنية": "الخبرات المهنية",
    "المؤهلاتالعلمية": "المؤهلات العلمية",
    "المؤهلاتالدراسية": "المؤهلات الدراسية",
    "التعليموالمؤهلات": "التعليم والمؤهلات",
    "المهاراتالتقنية": "المهارات التقنية",
    "المهاراتالفنية": "المهارات الفنية",
    "المهاراتالشخصية": "المهارات الشخصية",
    "المهاراتالناعمة": "المهارات الناعمة",
    "المهاراتاللغوية": "المهارات اللغوية",
    "اللغاتالمتقنة": "اللغات المتقنة",
    "الدوراتالتدريبية": "الدورات التدريبية",
    "التطويرالمهني": "التطوير المهني",
    "نبذةعني": "نبذة عني",
    "نبذةمهنية": "نبذة مهنية",
    "الملفالشخصي": "الملف الشخصي",
    "ملخصمهني": "ملخص مهني",
    "الهدفالوظيفي": "الهدف الوظيفي",
    "مشاريعشخصية": "مشاريع شخصية",
    "شهاداتمهنية": "شهادات مهنية",
}


def _normalize_spaced_headers(text: str) -> str:
    """Collapse spaced-out headers like 'WO R K  E X P E R I E N C E' → 'WORK EXPERIENCE'."""
    result_lines = []
    for line in text.split("\n"):
        stripped = line.strip()
        if stripped and len(stripped) < 80:
            upper_chars = sum(1 for c in stripped if c.isupper())
            space_chars = sum(1 for c in stripped if c == " ")
            total = len(stripped)
            if total > 5 and space_chars / total > 0.35 and upper_chars / total > 0.35:
                # "C L I N I C A L  E X P E R I E N C E" -> "CLINICAL EXPERIENCE":
                # words are separated by a wider gap than the letters are.
                tokens = stripped.split(" ")
                singles = sum(1 for t in tokens if len(t) <= 2 and t)
                if singles >= 0.6 * max(1, len([t for t in tokens if t])):
                    words = [re.sub(r"\s+", "", w) for w in re.split(r"\s{2,}", stripped) if w.strip()]
                    generic = " ".join(words).upper()
                    if len(words) == 1:
                        key = words[0].upper()
                        if key in _KNOWN_HEADERS:
                            generic = _KNOWN_HEADERS[key]
                        else:
                            m = re.match(r"^([A-Z]{2,})?(EXPERIENCE|EMPLOYMENT|SKILLS|EDUCATION|LANGUAGES|SUMMARY"
                                         r"|PROFILE|OBJECTIVE|PROJECTS|CERTIFICATIONS?|QUALIFICATIONS|COMPETENCIES"
                                         r"|CONTACT|PERSONAL|REFERENCES|INTERESTS|HOBBIES|AWARDS|PUBLICATIONS|TRAINING"
                                         r"|ACHIEVEMENTS)$", key)
                            if m:
                                generic = " ".join(p for p in m.groups() if p)
                    if generic:
                        result_lines.append(generic)
                        continue
                collapsed = re.sub(r"\s+", "", stripped).upper()
                if collapsed in _KNOWN_HEADERS:
                    result_lines.append(_KNOWN_HEADERS[collapsed])
                    continue
                matched = False
                for section_pattern in SECTION_PATTERNS.values():
                    if re.match(rf"^(?:{section_pattern})$", collapsed, re.IGNORECASE):
                        result_lines.append(collapsed)
                        matched = True
                        break
                if matched:
                    continue
        result_lines.append(line)
    return "\n".join(result_lines)


def _detect_column_split(blocks: list, page_width: float):
    """Detect a two-column / sidebar layout and return the gutter x, or None.

    Reading a two-column CV line-by-line interleaves the sidebar with the main
    body and destroys the structure (the classic cause of scrambled bullets and
    languages on upload). When a clean vertical gutter exists we read the left
    column fully, then the right column."""
    if len(blocks) < 4:
        return None
    # Try a centred split first, then progressively narrower sidebar splits.
    for frac in (0.5, 0.40, 0.36, 0.33, 0.30):
        split = page_width * frac
        left = [b for b in blocks if b["x1"] <= split + page_width * 0.02]
        right = [b for b in blocks if b["x0"] >= split - page_width * 0.01]
        crossing = [b for b in blocks
                    if b["x0"] < split - page_width * 0.04 and b["x1"] > split + page_width * 0.04]
        # Need real content on both sides and almost nothing straddling the
        # gutter (a single full-width header band is tolerated).
        need_right = 2 if frac == 0.5 else 3
        if len(left) >= 2 and len(right) >= need_right and len(crossing) <= 1:
            return split
    return None


_LAST_HINTS = {}


def extract_text_from_pdf(file_bytes: bytes) -> str:
    """Extract text from a PDF in READING order.

    Uses the glyph-position rebuild shared with PDF->Word (Scripts/pdf_text_rows.py):
    Arabic words come back in logical order even for Cairo/WeasyPrint PDFs (our
    own CV builder's output), two-column and sidebar layouts are read column by
    column, and one row of the page is one line of text. Also records font-size
    hints in _LAST_HINTS: the largest row on page 1 is almost always the name and
    the row under it the job title, which beats guessing from line position."""
    global _LAST_HINTS
    _LAST_HINTS = {}
    if not _HAS_FITZ:
        raise RuntimeError("PyMuPDF (fitz) is required for PDF parsing")
    try:
        from pdf_text_rows import page_flow, row_text_smart
    except Exception as exc:  # pragma: no cover - keeps parsing alive if the module is missing
        sys.stderr.write(f"pdf_text_rows unavailable ({exc}); falling back to plain extraction\n")
        return _extract_text_plain(file_bytes)

    doc = fitz.open(stream=file_bytes, filetype="pdf")
    out_lines = []
    best = None  # (size, text, column, index)
    first_page_rows = []
    for pno, page in enumerate(doc):
        _, flow = page_flow(page)
        for ci, (rows, _cx0, _cx1) in enumerate(flow):
            if ci:
                # a sidebar's last section must not swallow the main column
                out_lines.append("---COLUMN---")
            for r in rows:
                t = row_text_smart(r)
                if not t:
                    continue
                out_lines.append(t)
                if pno == 0:
                    first_page_rows.append((ci, r, t))
        out_lines.append("")
    doc.close()

    # name = largest row on page 1 that looks like a name; title = next row in that column
    for idx, (ci, r, t) in enumerate(first_page_rows):
        words = t.split()
        if not (1 <= len(words) <= 6) or EMAIL_RE.search(t) or URL_RE.search(t) or re.search(r"\d{3,}", t):
            continue
        if _is_section_header(t):
            continue
        if best is None or r["size"] > best[0] + 0.5:
            best = (r["size"], t, ci, idx)
    if best is not None:
        name = best[1]
        for ci, r, t in first_page_rows[best[3] + 1:]:
            if ci != best[2]:
                continue
            if EMAIL_RE.search(t) or URL_RE.search(t) or PHONE_RE.match(t) or _is_section_header(t):
                continue
            if abs(r["size"] - best[0]) <= 0.5 and len(t.split()) <= 3 and not re.search(r"\d", t):
                name = f"{name} {t}"  # a name wrapped onto a second line
                continue
            if r["size"] < best[0] and 2 <= len(t) <= 90:
                _LAST_HINTS["job_title"] = t
            break
        _LAST_HINTS["name"] = name

    full_text = "\n".join(out_lines)
    full_text = _normalize_spaced_headers(full_text)
    return full_text.strip()


def _extract_text_plain(file_bytes: bytes) -> str:
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = "\n".join(page.get_text() for page in doc)
    doc.close()
    return _normalize_spaced_headers(text).strip()


def _is_section_header(line: str):
    """Check if a line starts with a section header. Returns the section type or None."""
    stripped = line.strip()
    if not stripped or len(stripped) > 120:
        return None
    if stripped == "---COLUMN---":
        return "break"  # ends whatever section was open, starts none
    # A whole-line heading first, for every section: "Soft Skills" / "المهارات
    # الشخصية" / "Compétences personnelles" must not be taken for "Skills" /
    # "المهارات" / "Compétences" followed by text, which is what made an Arabic
    # CV's soft skills disappear on import.
    for section_type, pattern in SECTION_PATTERNS.items():
        if re.match(rf"^\s*{pattern}\s*:?\s*$", stripped, re.IGNORECASE):
            return section_type
    for section_type, pattern in _LABEL_PATTERNS.items():
        if re.match(rf"^\s*{pattern}\s*:?\s*$", stripped, re.IGNORECASE) or \
           re.match(rf"^\s*{pattern}\s*:\s*\S", stripped, re.IGNORECASE):
            return section_type
    # A heading with its content on the same line ("Skills: Python, SQL",
    # "PROFILE Engineer with ten years of…"): a colon or a full stop, or at least
    # four more words. "التواصل الفعال" (effective communication, a soft skill)
    # is not the "التواصل" (contact) heading, and it used to end the soft-skills
    # section after its first item.
    for section_type, pattern in SECTION_PATTERNS.items():
        if re.match(rf"^\s*{pattern}\s*(?:[:.]\s*\S.*|\s+\S+(?:\s+\S+){{3,}})$", stripped, re.IGNORECASE):
            return section_type
    # Devanagari: the PDF text layer puts vowel signs in the wrong place
    # ("कौशल" comes back as "कौशिल") and can drop the first consonant of a
    # conjunct ("सॉफ्ट स्किल्स" as "सॉफ्ट कल्स"), so compare consonant skeletons,
    # allowing one missing consonant in a long heading.
    if _DEVA.search(stripped) and len(stripped) <= 40:
        skel = _deva_skeleton(stripped)
        for label, section_type in _DEVA_HEADINGS.items():
            if skel == label or (len(label) >= 5 and len(skel) == len(label) - 1
                                 and any(label[:i] + label[i + 1:] == skel for i in range(len(label)))):
                return section_type
    # A tracked heading can split inside a word ("ال خبرات"): compare with all
    # whitespace removed on both sides.
    if len(stripped) <= 40:
        collapsed = re.sub(r"\s+", "", stripped).rstrip(":")
        for section_type, pattern in _SPACELESS_PATTERNS.items():
            if re.fullmatch(pattern, collapsed, re.IGNORECASE):
                return section_type
    # "CLINICAL EXPERIENCE", "KEY SKILLS", "RELEVANT WORK EXPERIENCE": a short
    # heading whose LAST words are a known section name.
    if len(stripped.split()) <= 4 and not stripped.endswith((".", "،", ",")):
        for section_type, pattern in SECTION_PATTERNS.items():
            if re.match(rf"^\s*(?:[^\W\d_]+\s+){{1,2}}{pattern}\s*:?\s*$", stripped, re.IGNORECASE):
                return section_type
    return None


def _extract_section_text(text: str, section_name: str) -> str:
    """Extract raw text for a given section."""
    lines = text.split("\n")
    section_start = -1
    section_end = len(lines)

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        sec_type = _is_section_header(stripped)
        if sec_type:
            if sec_type == section_name and section_start == -1:
                section_start = i + 1
            elif section_start != -1:
                section_end = i
                return "\n".join(lines[section_start:section_end]).strip()

    if section_start != -1:
        return "\n".join(lines[section_start:section_end]).strip()
    return ""


def _extract_personal_info(text: str) -> dict:
    """Extract personal info from the top of the CV."""
    lines = text.split("\n")

    name = ""
    for line in lines[:5]:
        stripped = line.strip()
        if stripped and len(stripped) > 1 and not EMAIL_RE.search(stripped) and not PHONE_RE.match(stripped):
            name = stripped
            break

    job_title = ""
    found_name = False
    for line in lines[:8]:
        stripped = line.strip()
        if not stripped:
            continue
        if not found_name:
            if stripped == name:
                found_name = True
            continue
        if not EMAIL_RE.search(stripped) and not PHONE_RE.match(stripped) and not URL_RE.search(stripped):
            is_header = _is_section_header(stripped) is not None
            if not is_header and len(stripped) > 3:
                job_title = stripped
                break

    email = ""
    email_match = EMAIL_RE.search(text[:1500])
    if email_match:
        email = email_match.group()

    phone = ""
    phone_match = PHONE_RE.search(text[:1500])
    if phone_match:
        candidate = phone_match.group().strip()
        digits = re.sub(r"\D", "", candidate)
        if len(digits) >= 7:
            phone = candidate

    linkedin = ""
    linkedin_match = LINKEDIN_RE.search(text)
    if linkedin_match:
        linkedin = linkedin_match.group()

    github = ""
    github_match = GITHUB_RE.search(text)
    if github_match:
        github = github_match.group()

    location = ""
    for line in lines[:15]:
        stripped = line.strip()
        loc_match = re.match(
            r"(?:Location|Address|City|الموقع|العنوان|المدينة|محل\s*الإقامة)\s*:?\s*"
            r"(.+?)(?=\s+(?:Phone|Tel|Mobile|E-?mail|LinkedIn|GitHub|Web(?:site)?|الهاتف|الجوال|البريد|الإيميل)\s*:|\s*[|•·]|$)",
            stripped, re.IGNORECASE)
        if loc_match:
            location = loc_match.group(1).strip()
            break

    portfolio = ""
    website = ""
    for url_match in URL_RE.finditer(text[:2000]):
        url = url_match.group()
        if "linkedin.com" in url.lower() or "github.com" in url.lower():
            continue
        if "portfolio" in url.lower() or "netlify" in url.lower() or "vercel" in url.lower():
            portfolio = url
        elif not website:
            website = url

    return {
        "full_name": name,
        "job_title": job_title,
        "email": email,
        "phone": phone,
        "location": location,
        "linkedin": linkedin,
        "portfolio": portfolio,
        "website": website,
        "github": github,
    }


_BULLET_MARKERS = ("•", "●", "◦", "▪", "∙", "*", "►", "·", "‣", "▸", "»")


def _split_dates(date_str: str):
    """Split a matched date range into (start, end)."""
    parts = re.split(r"\s*(?:[-–—]+|\bto\b|إلى)\s*", date_str.strip(), maxsplit=1)
    start = parts[0].strip() if parts else ""
    end = parts[1].strip() if len(parts) > 1 else ""
    return start, end


_CONT_WORDS = {
    "and", "or", "with", "to", "the", "of", "in", "for", "a", "an", "&", "but",
    "on", "at", "as", "by", "from", "that", "which", "using", "into", "across",
}


def _looks_like_continuation(prev: str, text: str) -> bool:
    """Heuristic: is `text` a wrapped continuation of the previous bullet rather
    than a new bullet? CV exporters wrap long sentences across lines; without
    this each wrapped line became its own bullet."""
    if not prev:
        return False
    prev_stripped = prev.rstrip()
    prev_end = prev_stripped[-1:] if prev_stripped else ""
    first = text.lstrip()[:1]

    # Previous line ends on a conjunction/preposition → definitely continues.
    last_word = prev_stripped.rstrip(",").split()[-1].lower() if prev_stripped.split() else ""
    if last_word in _CONT_WORDS:
        return True

    # Previous line ended mid-sentence (no terminal punctuation, or a comma).
    if prev_end and prev_end not in ".!?:;":
        if first and (first.islower() or first.isdigit() or prev_end == ","):
            return True
        if len(text.split()) <= 4 and first.islower():
            return True
    return False


def _add_bullet(entry: dict, text: str, is_marked: bool):
    """Append a description line, merging wrapped continuations into the last bullet."""
    text = text.strip()
    if not text:
        return
    desc = entry["description"]
    if not is_marked and desc and _looks_like_continuation(desc[-1], text):
        desc[-1] = (desc[-1].rstrip() + " " + text).strip()
    else:
        desc.append(text)


def _parse_experience_entries(section_text: str) -> list:
    """Parse experience section text into structured entries.

    Handles entries whose title / dates / company sit on separate lines, and
    merges wrapped sentence fragments back into single bullets."""
    if not section_text.strip():
        return []

    entries = []
    lines = section_text.split("\n")
    current_entry = None

    def new_entry(title="", company="", location="", start="", end=""):
        return {"title": title, "company": company, "location": location,
                "start_date": start, "end_date": end, "description": []}

    # Pre-pass: a short, unmarked line immediately followed by a dates-only line
    # is the title of a NEW entry ("Radiology Trainee" / "Apr 2024 – Jun 2024").
    # Without this every entry after the first was swallowed as description.
    clean = [ln.strip() for ln in lines]

    def _dates_only(s):
        m = DATE_RE.search(s)
        if not m:
            return False
        return len(s.replace(m.group(), "").strip(" |·–—-,")) <= 3

    title_lines = set()
    for i, s in enumerate(clean):
        if not s or s[0] in _BULLET_MARKERS or DATE_RE.search(s):
            continue
        nxt = clean[i + 1] if i + 1 < len(clean) else ""
        if nxt and _dates_only(nxt) and len(s.split()) <= 10 and not s.endswith((".", ";", "؛")):
            title_lines.add(i)

    for idx, raw in enumerate(lines):
        stripped = raw.strip()
        if not stripped:
            continue

        if idx in title_lines:
            if current_entry:
                entries.append(current_entry)
            current_entry = new_entry(title=stripped)
            continue

        is_marked = stripped[0] in _BULLET_MARKERS or bool(re.match(r"^[-–—]\s+\S", stripped))
        bullet_text = re.sub(r"^[•●◦▪∙*►·‣▸»\-–—]+\s*", "", stripped) if is_marked else stripped

        date_match = DATE_RE.search(stripped)
        date_str = date_match.group() if date_match else ""
        # A line is "date-ish" when removing the date leaves little/no other text.
        residual = stripped.replace(date_str, "").strip(" |·–—-,") if date_str else stripped
        date_only = bool(date_str) and len(residual) <= 3

        has_started_bullets = bool(current_entry and current_entry["description"])

        # ── Marked bullet → always a description line ──
        if is_marked:
            if current_entry is None:
                current_entry = new_entry(title="Experience")
            _add_bullet(current_entry, bullet_text, True)
            continue

        # ── A pure date line ── belongs to the current header, else starts one.
        if date_only:
            start, end = _split_dates(date_str)
            if current_entry and not current_entry["start_date"] and not has_started_bullets:
                current_entry["start_date"] = start
                current_entry["end_date"] = end
            else:
                if current_entry:
                    entries.append(current_entry)
                current_entry = new_entry(start=start, end=end)
            continue

        # ── A header line that also carries a date (Title | Company | Dates) ──
        # Also when bullets have started: "Site Engineer · Jan 2017 – Mar 2020"
        # after the previous entry's bullets is the NEXT entry, not a bullet.
        residual_words = [w for w in re.split(r"\s+", re.sub(r"[·|,]", " ", residual)) if re.search(r"\w", w)]
        if date_str and (not has_started_bullets or (not is_marked and 1 <= len(residual_words) <= 8)):
            if current_entry and not current_entry["title"]:
                pass  # fall through to fill below
            elif current_entry and not current_entry["start_date"] and current_entry["title"]:
                start, end = _split_dates(date_str)
                current_entry["start_date"], current_entry["end_date"] = start, end
                rest = re.split(r"\s*[|·–—]\s*", residual)
                if rest and not current_entry["company"]:
                    current_entry["company"] = rest[0].strip()
                continue
            else:
                if current_entry:
                    entries.append(current_entry)
                start, end = _split_dates(date_str)
                parts = re.split(r"\s*[|·]\s*", residual)
                current_entry = new_entry(
                    title=parts[0].strip() if parts else "",
                    company=parts[1].strip() if len(parts) > 1 else "",
                    location=parts[2].strip() if len(parts) > 2 else "",
                    start=start, end=end)
                continue

        # ── Plain text line ──
        # A sentence-like line (has metrics, ends with ., or is long) is a
        # description, never a company/location label.
        # Separators (· | ,) are not words: "Fayoum Cancer Institute · Fayoum, Egypt ·
        # Internship" is a company line, not a sentence.
        words_only = [w for w in re.split(r"\s+", re.sub(r"[·|,]", " ", stripped)) if re.search(r"\w", w)]
        is_sentence = (len(words_only) >= 9
                       or stripped.rstrip().endswith((".", "!", ";", "؛"))
                       or (bool(re.search(r"\d{2,}", stripped)) and len(words_only) >= 6))
        is_location = (not is_sentence and "," in stripped
                       and len(stripped.split()) <= 4 and len(stripped) < 40)

        if current_entry is None:
            current_entry = new_entry(title=stripped)
        elif not current_entry["title"]:
            current_entry["title"] = stripped
        elif not has_started_bullets and not current_entry["company"] and not is_sentence and len(stripped) < 80:
            parts = [p.strip() for p in re.split(r"\s*[|·]\s*", stripped) if p.strip()]
            current_entry["company"] = parts[0] if parts else stripped
            if len(parts) > 1 and not current_entry["location"]:
                current_entry["location"] = parts[1]
        elif not has_started_bullets and not current_entry["location"] and is_location:
            current_entry["location"] = stripped
        else:
            _add_bullet(current_entry, stripped, False)

    if current_entry:
        entries.append(current_entry)

    return entries


def _parse_education_entries(section_text: str) -> list:
    """Parse education section text into structured entries."""
    if not section_text.strip():
        return []

    entries = []
    lines = section_text.split("\n")
    current_entry = None

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        date_match = DATE_RE.search(stripped)

        if date_match:
            date_str = date_match.group()
            remaining = stripped.replace(date_str, "").strip(" |·–—-,")
            date_parts = re.split(r"\s*(?:[-–—]+|\bto\b)\s*", date_str, maxsplit=1)
            start_date = date_parts[0].strip() if date_parts else ""
            end_date = date_parts[1].strip() if len(date_parts) > 1 else ""

            if current_entry and not current_entry["start_date"]:
                current_entry["start_date"] = start_date
                current_entry["end_date"] = end_date
                if remaining:
                    current_entry["institution"] = remaining
            else:
                if current_entry:
                    entries.append(current_entry)
                current_entry = {
                    "degree": remaining or "", "institution": "", "location": "",
                    "start_date": start_date, "end_date": end_date, "gpa": "",
                }
        elif current_entry is None:
            current_entry = {
                "degree": stripped, "institution": "", "location": "",
                "start_date": "", "end_date": "", "gpa": "",
            }
        else:
            gpa_match = re.search(
                r"(?:GPA|CGPA|Grade|المعدل|التقدير)\s*:?\s*"
                r"([\d.]+(?:\s*(?:/|out of|من)\s*[\d.]+)?(?:\s*\([^)]+\))?|(?:Excellent|Very\s*Good|Good|Fair|Pass|Distinction|ممتاز|جيد\s*جداً?|جيد|مقبول|امتياز))",
                stripped, re.IGNORECASE,
            )
            if gpa_match:
                current_entry["gpa"] = gpa_match.group(1).strip()
            elif not current_entry["institution"]:
                current_entry["institution"] = stripped
            elif not current_entry["location"]:
                current_entry["location"] = stripped

    if current_entry:
        entries.append(current_entry)

    return entries


def _parse_skills(section_text: str) -> list:
    """Parse skills section into a list of skill strings."""
    if not section_text.strip():
        return []

    skills = []
    for line in section_text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        stripped = re.sub(r"^[•●◦▪\-–—∙*►]\s*", "", stripped)
        # "IMAGING MODALITIES" / "TECHNICAL": a sub-heading inside the skills
        # column, not a skill.
        # An acronym is a skill, not a sub-heading: "AWS", "SQL", "HTML", "CI/CD".
        if (re.fullmatch(r"[A-Z][A-Z &/-]{2,}", stripped) and len(stripped.split()) <= 3
                and len(stripped.replace(" ", "")) > 5 and "/" not in stripped):
            continue
        parts = re.split(r"[,;|·•]+", stripped)
        for part in parts:
            clean = part.strip()
            if clean and len(clean) > 1 and len(clean) < 60:
                skills.append(clean)

    return skills


_LEVEL_WORDS = {
    "native", "fluent", "proficient", "proficiency", "intermediate", "beginner",
    "basic", "advanced", "conversational", "professional", "elementary", "limited",
    "working", "bilingual", "mother tongue", "mothertongue",
    "a1", "a2", "b1", "b2", "c1", "c2",
    "أصلية", "بطلاقة",
    "متقدم", "متوسط",
    "مبتدئ", "جيد", "ممتاز",
}


def _parse_languages(section_text: str) -> list:
    """Parse the languages section into clean {name, level} entries.

    Splits on commas/newlines, pairs "Name: Level" / "Name (Level)", attaches a
    stray level word to the preceding language, and never emits a level word
    (Native, C2, ...) as its own language."""
    if not section_text.strip():
        return []

    cleaned = re.sub(r"^[•●◦▪∙*►·‣▸»\-–—]+\s*",
                     "", section_text, flags=re.MULTILINE)
    units = [u.strip().strip("•-–—:·").strip() for u in re.split(r"[,\n;·|]+", cleaned)]

    languages = []
    for unit in units:
        if not unit:
            continue
        name, level = unit, ""
        sep = re.match(r"^(.+?)\s*[:\-–—·|]\s*(.+)$", unit)
        paren = re.match(r"^(.+?)\s*\(([^)]+)\)\s*$", unit)
        if sep:
            name, level = sep.group(1).strip(), sep.group(2).strip()
        elif paren:
            name, level = paren.group(1).strip(), paren.group(2).strip()
        else:
            name = unit.strip()

        low = name.lower().strip()
        if low in _LEVEL_WORDS:
            if languages and not languages[-1]["level"]:
                languages[-1]["level"] = name
            continue
        if not name or len(name) > 30:
            continue
        languages.append({"name": name, "level": level})

    return languages


def _parse_summary(text: str) -> str:
    """Extract summary/profile section text."""
    section_text = _extract_section_text(text, "summary")
    if section_text:
        lines = []
        for line in section_text.split("\n"):
            stripped = line.strip()
            if stripped:
                stripped = re.sub(r"^[•●◦▪\-–—∙*►]\s*", "", stripped)
                lines.append(stripped)
        return " ".join(lines)
    return ""


def parse_cv_heuristic(text: str) -> dict:
    """Parse CV text using heuristic/regex patterns into structured data."""
    personal = _extract_personal_info(text)
    summary = _parse_summary(text)

    exp_text = _extract_section_text(text, "experience")
    experience = _parse_experience_entries(exp_text)

    edu_text = _extract_section_text(text, "education")
    education = _parse_education_entries(edu_text)

    skills_text = _extract_section_text(text, "skills")
    skills = _parse_skills(skills_text)

    soft_skills_text = _extract_section_text(text, "soft_skills")
    soft_skills = _parse_skills(soft_skills_text)

    lang_text = _extract_section_text(text, "languages")
    languages = _parse_languages(lang_text)

    return {
        "personal_info": personal,
        "summary": summary,
        "experience": experience,
        "education": education,
        "skills": skills,
        "soft_skills": soft_skills,
        "languages": languages,
    }


def parse_uploaded_cv(file_bytes: bytes) -> dict:
    """Parse a PDF CV and return structured data."""
    text = extract_text_from_pdf(file_bytes)

    if not text.strip():
        return {
            "personal_info": {"full_name": "", "job_title": "", "email": "", "phone": "",
                              "location": "", "linkedin": "", "portfolio": "", "website": "", "github": ""},
            "summary": "",
            "experience": [],
            "education": [],
            "skills": [],
            "soft_skills": [],
            "languages": [],
        }

    result = parse_cv_heuristic(text)
    hints = dict(_LAST_HINTS)
    personal = result.get("personal_info") or {}
    name = (personal.get("full_name") or "").strip()
    # The line-position guess fails on sidebar CVs (first line = a contact header)
    # and on Arabic ones; the largest-font row on page 1 is the name.
    if hints.get("name") and (not name or _is_section_header(name) or len(name.split()) > 6
                              or name != hints["name"]):
        personal["full_name"] = hints["name"]
    # The row right under the name, in the name's own column, is the job title;
    # the line-order guess picks a sidebar's address or the name's second line.
    if hints.get("job_title"):
        personal["job_title"] = hints["job_title"]
    result["personal_info"] = personal
    return result


# ══════════════════════════════════════════════════════════════════
# CLI Entry Point
# ══════════════════════════════════════════════════════════════════

def main():
    if len(sys.argv) < 2:
        print(json.dumps({"success": False, "message": "Usage: cv_builder_export.py <export|docx|parse> [args]"}))
        sys.exit(1)

    command = sys.argv[1]

    if command == "export" and len(sys.argv) >= 4:
        input_json = sys.argv[2]
        output_pdf = sys.argv[3]

        try:
            with open(input_json, "r", encoding="utf-8") as f:
                data = json.load(f)

            pdf_bytes = export_cv_to_pdf(data)

            with open(output_pdf, "wb") as f:
                f.write(pdf_bytes)

            print(json.dumps({
                "success": True,
                "message": f"PDF exported ({len(pdf_bytes)} bytes)",
                "outputFile": output_pdf,
                "fileSize": len(pdf_bytes),
            }))
        except Exception as e:
            print(json.dumps({"success": False, "message": str(e)}))
            sys.exit(1)

    elif command == "docx" and len(sys.argv) >= 4:
        try:
            with open(sys.argv[2], "r", encoding="utf-8") as f:
                data = json.load(f)
            docx_bytes = export_cv_to_docx(data)
            with open(sys.argv[3], "wb") as f:
                f.write(docx_bytes)
            print(json.dumps({"success": True, "message": f"DOCX exported ({len(docx_bytes)} bytes)",
                              "outputFile": sys.argv[3], "fileSize": len(docx_bytes)}))
        except Exception as e:
            print(json.dumps({"success": False, "message": str(e)}))
            sys.exit(1)

    elif command == "parse" and len(sys.argv) >= 4:
        input_pdf = sys.argv[2]
        output_json = sys.argv[3]

        try:
            with open(input_pdf, "rb") as f:
                file_bytes = f.read()

            parsed = parse_uploaded_cv(file_bytes)

            with open(output_json, "w", encoding="utf-8") as f:
                json.dump(parsed, f, ensure_ascii=False, indent=2)

            print(json.dumps({
                "success": True,
                "message": "CV parsed successfully",
            }))
        except Exception as e:
            print(json.dumps({"success": False, "message": str(e)}))
            sys.exit(1)

    else:
        print(json.dumps({"success": False, "message": "Invalid command. Use: export <input.json> <output.pdf> | docx <input.json> <output.docx> | parse <input.pdf> <output.json>"}))
        sys.exit(1)


if __name__ == "__main__":
    main()
