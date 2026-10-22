#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shaped text with a correct, searchable text layer on a PyMuPDF page.

Why this exists (25 Sept 2026): MuPDF's insert_htmlbox shapes Arabic nicely but
writes a ToUnicode map that extracts as garbage — the PDF editor's replacement
«صفحة 3 من 20» came back as «ﻪğﺤ4Ěﺻ3» from its own extraction, pdftotext and
Chrome, so an edited word could not be found or edited again. Watermark and
page-number stamps use the same call and have the same text layer.

What it does: HarfBuzz shapes each bidi run (Arabic contextual forms, lam-alef,
marks, kerning); the glyphs are written in VISUAL order, as Chrome, Word and
LibreOffice do, from a subset of the face that is embedded with its own
ToUnicode, and every drawn glyph maps to exactly ONE character:
  • a glyph used for two different characters gets a clone (fonts that share a
    body between letters stay extractable);
  • the further letters of a ligature (lam-alef, «في» in IBM Plex) ride on
    invisible zero-width glyphs placed just inside the ligature's end edge and
    written before it, which is where a reader of visual-order text expects
    them. Chosen by measurement on 12 Arabic lines, words recovered by MuPDF /
    pdftotext / pdf.js / Confileo's own extraction: ligature text as a
    multi-letter ToUnicode entry 70/72/78/96 of 96 (Chrome's own output of the
    same lines: 70/-/8/96), ActualText 70/64/67/70, this 88/80-90/86-94/96.
  • optional ligatures (liga, clig, dlig) are off; required ones (rlig, lam-alef)
    stay. A Naskh «الله» ligature cost pdftotext and pdf.js one word each.

Without uharfbuzz (install.sh adds it) only text that needs no shaping can be
laid out (Latin, Greek, Cyrillic, digits...): available(text) tells the caller,
which keeps its old path for the rest.

API
  find_face(family_hint, bold, italic, script)      -> Face or None
  font_stack(family_hint, bold, italic, text)       -> [Face, ...] (primary first)
  layout(text, stack, size, base_rtl=None)          -> Line (visual-order glyphs, width)
  wrap(text, stack, size, width, base_rtl=None)     -> [Line, ...]
  TextLayerWriter(doc).draw(page, line, x, baseline, color, ...) ; .flush()
"""

import copy
import io
import logging
import os
import re
import sys
import unicodedata

import fitz

import pdf_geometry as geom  # same folder

logging.getLogger("fontTools").setLevel(logging.ERROR)

try:
    import uharfbuzz as _hb
except Exception:  # noqa: BLE001 — optional: shaping is then unavailable
    _hb = None

HERE = os.path.dirname(os.path.abspath(__file__))

_FONT_ROOTS = [
    os.path.join(HERE, "..", "Fonts"),
    os.path.join(HERE, "Fonts"),
    "/usr/share/fonts",
    "/usr/local/share/fonts",
    os.path.expanduser("~/.fonts"),
    os.path.expanduser("~/.local/share/fonts"),
    os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts"),
]

# Optional ligatures off, required ones (rlig: lam-alef) on — see the module doc.
_FEATURES = {"kern": True, "liga": False, "clig": False, "dlig": False}

_CARRIER_INSET = 0.6      # pt inside the ligature's end edge (clear of the 0.5 pt edge rules in pdf_to_word)
_CARRIER_STEP = 0.05


def _is_rtl_char(c):
    return unicodedata.bidirectional(c) in ("R", "AL")


def needs_shaping(text):
    """True when the text has a script whose glyphs depend on their neighbours."""
    for c in text or "":
        o = ord(c)
        if (0x0590 <= o <= 0x08FF or 0xFB1D <= o <= 0xFDFF or 0xFE70 <= o <= 0xFEFF
                or 0x0900 <= o <= 0x0DFF or 0x0E00 <= o <= 0x0EFF or 0x1000 <= o <= 0x109F
                or unicodedata.combining(c)):
            return True
    return False


def available(text=""):
    """Can this module set `text`? Shaping-free text always; the rest needs HarfBuzz."""
    return _hb is not None or not needs_shaping(text)


# ── fonts ──────────────────────────────────────────────────────────────────

_INDEX = None


def _font_index():
    global _INDEX
    if _INDEX is None:
        idx = {}
        for root in _FONT_ROOTS:
            if not os.path.isdir(root):
                continue
            for d, _dirs, files in os.walk(root):
                for f in files:
                    if f.lower().endswith((".ttf", ".otf")):
                        idx.setdefault(f.lower(), os.path.join(d, f))
        _INDEX = idx
    return _INDEX


class Face:
    """One font file: HarfBuzz for shaping, fontTools for coverage and metrics."""
    _cache = {}

    def __init__(self, path):
        from fontTools.ttLib import TTFont
        self.path = path
        with open(path, "rb") as f:
            self.data = f.read()
        tt = TTFont(io.BytesIO(self.data), lazy=True, fontNumber=0)
        self.upem = tt["head"].unitsPerEm
        self.cmap = tt.getBestCmap() or {}
        self.hmtx = tt["hmtx"].metrics
        self.order = tt.getGlyphOrder()
        os2 = tt["OS/2"] if "OS/2" in tt else None
        hhea = tt["hhea"]
        self.ascent = (getattr(os2, "sTypoAscender", 0) or hhea.ascent) / self.upem
        self.descent = abs(getattr(os2, "sTypoDescender", 0) or hhea.descent) / self.upem
        self.cff = "CFF " in tt or "CFF2" in tt
        name = os.path.basename(path).lower()
        self.bold = any(w in name for w in ("bold", "black", "heavy"))
        self.italic = any(w in name for w in ("italic", "oblique"))
        self.gid_of = {n: i for i, n in enumerate(self.order)}
        self._hbfont = None
        self._decomposes = None

    @classmethod
    def get(cls, path):
        if path not in cls._cache:
            cls._cache[path] = cls(path)
        return cls._cache[path]

    def covers(self, c):
        return ord(c) in self.cmap

    def hbfont(self):
        if self._hbfont is None:
            face = _hb.Face(self.data)
            self._hbfont = _hb.Font(face)
        return self._hbfont

    def decomposes(self):
        """Does the face draw Arabic dots as separate glyphs (Noto Sans Arabic
        2.013 from Google Fonts does)? Such bodies are shared between letters;
        the clone rule still keeps them extractable, but a face that keeps
        letters whole reads better in every viewer, so it is preferred."""
        if self._decomposes is None:
            self._decomposes = False
            if _hb is not None and all(ord(c) in self.cmap for c in "بتث"):
                g = _shape_run(self, "بتث", True)
                self._decomposes = len(g) > 3
        return self._decomposes


def _find_file(*names):
    idx = _font_index()
    for n in names:
        p = idx.get(n.lower())
        if p:
            return p
    return None


def _family_class(hint):
    h = (hint or "").lower()
    if any(w in h for w in ("mono", "courier", "consolas", "menlo", "code")):
        return "mono"
    if any(w in h for w in ("kufi",)):
        return "kufi"
    if any(w in h for w in ("naskh", "amiri", "traditional", "simplified", "scheherazade", "lateef",
                            "kfgq", "times", "georgia", "garamond", "serif", "cambria", "book", "palatino",
                            "minion", "roman", "charter", "century")) and "sans" not in h:
        return "serif"
    return "sans"


def _candidates(cls, bold, italic, script):
    b, i = bold, italic
    if script == "arabic":
        if cls in ("serif",):
            return (["NotoNaskhArabic-Bold.ttf", "Amiri-Bold.ttf"] if b else []) + \
                   ["NotoNaskhArabic-Regular.ttf", "Amiri-Regular.ttf"] + \
                   (["IBMPlexSansArabic-Bold.ttf"] if b else []) + ["IBMPlexSansArabic-Regular.ttf"]
        if cls == "kufi":
            return (["NotoKufiArabic-Bold.ttf"] if b else []) + ["NotoKufiArabic-Regular.ttf"] + \
                   (["IBMPlexSansArabic-Bold.ttf"] if b else []) + ["IBMPlexSansArabic-Regular.ttf"]
        return (["IBMPlexSansArabic-Bold.ttf", "NotoSansArabic-Bold.ttf"] if b else []) + \
               ["IBMPlexSansArabic-Regular.ttf", "NotoSansArabic-Regular.ttf", "NotoNaskhArabic-Regular.ttf",
                "Amiri-Regular.ttf"]
    style = {(False, False): ("Regular", "", "Regular"), (True, False): ("Bold", "-Bold", "Bold"),
             (False, True): ("Italic", "-Oblique", "Italic"), (True, True): ("BoldItalic", "-BoldOblique", "BoldItalic")}[(b, i)]
    lib, dv, noto = style
    if cls == "serif":
        dvs = {"": "DejaVuSerif.ttf", "-Bold": "DejaVuSerif-Bold.ttf", "-Oblique": "DejaVuSerif-Italic.ttf",
               "-BoldOblique": "DejaVuSerif-BoldItalic.ttf"}[dv]
        return ["LiberationSerif-%s.ttf" % lib, "NotoSerif-%s.ttf" % noto, dvs,
                "LiberationSerif-%s.ttf" % ("Bold" if b else "Regular"), "DejaVuSerif.ttf"]
    if cls == "mono":
        return ["LiberationMono-%s.ttf" % lib, "DejaVuSansMono%s.ttf" % dv, "NotoSansMono-%s.ttf" % noto,
                "LiberationMono-Regular.ttf", "DejaVuSansMono.ttf"]
    return ["LiberationSans-%s.ttf" % lib, "Arimo-%s.ttf" % noto, "DejaVuSans%s.ttf" % dv,
            "NotoSans-%s.ttf" % noto, ("arialbd.ttf" if b else "arial.ttf"),
            "LiberationSans-Regular.ttf", "DejaVuSans.ttf", "arial.ttf"]


# Faces the server has, by the family name a PDF gives them: the original's own
# family wins when it is here (a Noto Sans Arabic document keeps Noto Sans Arabic).
_FAMILIES = {
    "notosansarabic": ("arabic", "NotoSansArabic-%s.ttf"),
    "ibmplexsansarabic": ("arabic", "IBMPlexSansArabic-%s.ttf"),
    "notonaskharabic": ("arabic", "NotoNaskhArabic-%s.ttf"),
    "notokufiarabic": ("arabic", "NotoKufiArabic-%s.ttf"),
    "amiri": ("arabic", "Amiri-%s.ttf"),
    "arial": ("latin", None),
    "liberationsans": ("latin", "LiberationSans-%s.ttf"),
    "liberationserif": ("latin", "LiberationSerif-%s.ttf"),
    "liberationmono": ("latin", "LiberationMono-%s.ttf"),
    "dejavusans": ("latin", None),
    "notosans": ("latin", "NotoSans-%s.ttf"),
}


def _own_family(hint, bold, italic, script):
    h = re.sub(r"[^a-z]", "", (hint or "").lower().split("+")[-1])
    for key, (sc, pattern) in _FAMILIES.items():
        if not h.startswith(key) or sc != script:
            continue
        if key == "arial":
            return ["arialbd.ttf" if bold else "arial.ttf"]
        if key == "dejavusans":
            return ["DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"]
        style = ("Bold" if bold else "") + ("Italic" if italic else "") or "Regular"
        return [pattern % style, pattern % ("Bold" if bold else "Regular")]
    return []


def find_face(family_hint="", bold=False, italic=False, script="latin"):
    own = _own_family(family_hint, bold, italic, script)
    for n in own:
        p = _find_file(n)
        if p:
            try:
                return Face.get(p)
            except Exception:  # noqa: BLE001
                pass
    cls = _family_class(family_hint)
    names = _candidates(cls, bold, italic, script)
    first_decomposing = None
    for n in names:
        p = _find_file(n)
        if not p:
            continue
        try:
            f = Face.get(p)
        except Exception:  # noqa: BLE001 — unreadable font file
            continue
        if script == "arabic" and f.decomposes():
            first_decomposing = first_decomposing or f
            continue
        return f
    return first_decomposing


def font_stack(family_hint="", bold=False, italic=False, text=""):
    """Faces to try per character, the one matching the text's first strong script first."""
    first_rtl = next((_is_rtl_char(c) for c in text if unicodedata.bidirectional(c) in ("L", "R", "AL")), False)
    order = ["arabic", "latin"] if first_rtl else ["latin", "arabic"]
    stack = []
    for sc in order:
        f = find_face(family_hint, bold, italic, sc)
        if f is not None and f not in stack:
            stack.append(f)
    for extra in ("DejaVuSans.ttf", "NotoSans-Regular.ttf", "arial.ttf"):
        p = _find_file(extra)
        if p:
            try:
                f = Face.get(p)
            except Exception:  # noqa: BLE001
                continue
            if f not in stack:
                stack.append(f)
    return stack


# ── bidi (UAX #9 for one paragraph without explicit embeddings) ────────────

def bidi_runs(text, base_rtl=None):
    """-> (base_rtl, [(start, end, level)]) in logical order."""
    types = [unicodedata.bidirectional(c) or "ON" for c in text]
    if base_rtl is None:
        base_rtl = next((t in ("R", "AL") for t in types if t in ("L", "R", "AL")), False)
    e = "R" if base_rtl else "L"
    for i, t in enumerate(types):                        # W1
        if t == "NSM":
            types[i] = types[i - 1] if i else e
    last = e
    for i, t in enumerate(types):                        # W2
        if t in ("L", "R", "AL"):
            last = t
        elif t == "EN" and last == "AL":
            types[i] = "AN"
    types = ["R" if t == "AL" else t for t in types]     # W3
    for i in range(1, len(types) - 1):                   # W4
        if types[i] == "ES" and types[i - 1] == types[i + 1] == "EN":
            types[i] = "EN"
        elif types[i] == "CS" and types[i - 1] == types[i + 1] and types[i - 1] in ("EN", "AN"):
            types[i] = types[i - 1]
    i = 0
    while i < len(types):                                # W5
        if types[i] == "ET":
            j = i
            while j < len(types) and types[j] == "ET":
                j += 1
            if (i > 0 and types[i - 1] == "EN") or (j < len(types) and types[j] == "EN"):
                for k in range(i, j):
                    types[k] = "EN"
            i = j
        else:
            i += 1
    types = ["ON" if t in ("ES", "ET", "CS") else t for t in types]   # W6
    last = e
    for i, t in enumerate(types):                        # W7
        if t in ("L", "R"):
            last = t
        elif t == "EN" and last == "L":
            types[i] = "L"

    def strong(t):
        return "R" if t in ("R", "EN", "AN") else ("L" if t == "L" else None)
    i = 0
    while i < len(types):                                # N1, N2
        if strong(types[i]) is None:
            j = i
            while j < len(types) and strong(types[j]) is None:
                j += 1
            before = strong(types[i - 1]) if i > 0 else e
            after = strong(types[j]) if j < len(types) else e
            for k in range(i, j):
                types[k] = before if before == after else e
            i = j
        else:
            i += 1
    levels = []
    for t in types:                                      # I1, I2
        if not base_rtl:
            levels.append(1 if t == "R" else 2 if t in ("AN", "EN") else 0)
        else:
            levels.append(2 if t in ("L", "EN", "AN") else 1)
    runs, s = [], 0
    for i in range(1, len(text) + 1):
        if i == len(text) or levels[i] != levels[s]:
            runs.append((s, i, levels[s]))
            s = i
    return base_rtl, runs


def visual_runs(text, base_rtl=None):
    base_rtl, runs = bidi_runs(text, base_rtl)
    if not runs:
        return base_rtl, []
    maxl = max(r[2] for r in runs)
    lowest_odd = min((r[2] for r in runs if r[2] % 2), default=maxl + 1)
    order = list(runs)
    for lvl in range(maxl, max(lowest_odd, 1) - 1, -1):  # L2
        i = 0
        while i < len(order):
            if order[i][2] >= lvl:
                j = i
                while j < len(order) and order[j][2] >= lvl:
                    j += 1
                order[i:j] = order[i:j][::-1]
                i = j
            else:
                i += 1
    return base_rtl, order


# ── shaping and layout ─────────────────────────────────────────────────────

def _shape_run(face, text, rtl):
    """-> [[gid, advance_em, x_offset_em, y_offset_em, text_of_cluster], ...] in visual order."""
    if _hb is None:
        out = []
        for c in text:                                   # no shaping: one glyph per character
            gid = face.gid_of.get(face.cmap.get(ord(c)), 0)
            adv = face.hmtx.get(face.order[gid], (0, 0))[0] / face.upem
            out.append([gid, adv, 0.0, 0.0, c])
        return out[::-1] if rtl else out
    buf = _hb.Buffer()
    buf.add_str(text)
    buf.guess_segment_properties()
    buf.direction = "rtl" if rtl else "ltr"
    buf.cluster_level = _hb.BufferClusterLevel.MONOTONE_CHARACTERS
    _hb.shape(face.hbfont(), buf, _FEATURES)
    infos, poss = buf.glyph_infos, buf.glyph_positions
    clusters = sorted(set(i.cluster for i in infos))
    nxt = {c: (clusters[k + 1] if k + 1 < len(clusters) else len(text)) for k, c in enumerate(clusters)}
    # a cluster's text goes to its widest glyph: the body, not a dot or mark
    # that a face like Noto Sans Arabic 2.x draws as a glyph of its own
    carrier = {}
    for k, (info, pos) in enumerate(zip(infos, poss)):
        c = info.cluster
        if c not in carrier or pos.x_advance > poss[carrier[c]].x_advance:
            carrier[c] = k
    out = []
    upem = face.upem
    for k, (info, pos) in enumerate(zip(infos, poss)):
        c = info.cluster
        t = text[c:nxt[c]] if carrier[c] == k else ""
        out.append([info.codepoint, pos.x_advance / upem, pos.x_offset / upem, pos.y_offset / upem, t])
    return out


class Line:
    """Glyphs in visual order: (face, gid, x_em, y_em, advance_em, text, rtl)."""

    def __init__(self, glyphs, width_em, size, base_rtl, text):
        self.glyphs, self.width_em, self.size, self.base_rtl, self.text = glyphs, width_em, size, base_rtl, text

    @property
    def width(self):
        return self.width_em * self.size

    def scaled(self, size):
        return Line(self.glyphs, self.width_em, size, self.base_rtl, self.text)


def _split_by_face(text, stack):
    """Consecutive characters that the same face covers. Spaces and marks stay
    with their neighbours when that face has them."""
    parts = []
    for c in text:
        pick = None
        if parts and (c.isspace() or unicodedata.combining(c)) and parts[-1][0].covers(c):
            pick = parts[-1][0]
        if pick is None:
            pick = next((f for f in stack if f.covers(c)), stack[0])
        if parts and parts[-1][0] is pick:
            parts[-1][1].append(c)
        else:
            parts.append((pick, [c]))
    return [(f, "".join(cs)) for f, cs in parts]


def layout(text, stack, size, base_rtl=None):
    base_rtl, runs = visual_runs(text, base_rtl)
    glyphs, pen = [], 0.0
    for st, en, lvl in runs:
        rtl = lvl % 2 == 1
        pieces = _split_by_face(text[st:en], stack)
        if rtl:
            pieces = pieces[::-1]
        for face, piece in pieces:
            for gid, adv, xo, yo, t in _shape_run(face, piece, rtl):
                glyphs.append((face, gid, pen + xo, yo, adv, t, rtl))
                pen += adv
    return Line(glyphs, pen, size, base_rtl, text)


def wrap(text, stack, size, width, base_rtl=None):
    """Greedy word wrap into lines no wider than `width` (a single word that is
    wider stays alone on its line)."""
    words = text.split()
    lines, cur = [], []
    for w in words:
        trial = " ".join(cur + [w])
        if cur and layout(trial, stack, size, base_rtl).width > width:
            lines.append(layout(" ".join(cur), stack, size, base_rtl))
            cur = [w]
        else:
            cur.append(w)
    if cur:
        lines.append(layout(" ".join(cur), stack, size, base_rtl))
    return lines


# ── embedding and drawing ──────────────────────────────────────────────────

def _tounicode(mapping):
    lines = ["<%04x> <%s>" % (gid, "".join("%04x" % ord(c) for c in t if ord(c) < 0x10000))
             for gid, t in sorted(mapping.items()) if t and any(ord(c) < 0x10000 for c in t)]
    body = "".join("%d beginbfchar\n%s\nendbfchar\n" % (len(lines[k:k + 100]), "\n".join(lines[k:k + 100]))
                   for k in range(0, len(lines), 100))
    return ("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n"
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n"
            "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <ffff>\n"
            "endcodespacerange\n" + body + "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n").encode()


class _Embedding:
    """The glyphs one page draws from one face: each (gid, character) pair gets
    its own glyph id, cloning the outline when a gid is drawn for two
    characters, and invisible zero-width glyphs carry ligature letters."""

    def __init__(self, face):
        self.face = face
        self.uses = {}        # (gid or None, char) -> ("orig", gid) | ("new", k)
        self.extra = []       # (kind, src gid, char)

    def key(self, gid, ch):
        k = (gid, ch)
        if k not in self.uses:
            if gid is not None and not any(g == gid for g, _c in self.uses):
                self.uses[k] = ("orig", gid)
            else:
                self.extra.append(("clone" if gid is not None else "carrier", gid, ch))
                self.uses[k] = ("new", len(self.extra) - 1)
        return k

    def build(self):
        from fontTools.ttLib import TTFont
        from fontTools import subset as ftsubset
        from fontTools.ttLib.tables._g_l_y_f import Glyph
        tt = TTFont(io.BytesIO(self.face.data), fontNumber=0)
        keep = {v for kind, v in self.uses.values() if kind == "orig"} | \
               {g for kind, g, _c in self.extra if kind == "clone"} | {0}
        opts = ftsubset.Options()
        opts.retain_gids = True
        opts.notdef_outline = True
        opts.layout_features = []
        opts.name_IDs = ["*"]
        opts.name_languages = ["*"]
        opts.drop_tables += ["GSUB", "GPOS", "GDEF", "BASE", "JSTF", "kern", "DSIG", "meta", "hdmx", "VDMX", "LTSH"]
        sub = ftsubset.Subsetter(opts)
        sub.populate(gids=sorted(keep))
        sub.subset(tt)
        order = list(tt.getGlyphOrder())
        newid = {}
        if self.extra:
            if "glyf" not in tt:
                raise ValueError("CFF face: no room for clone/carrier glyphs")
            glyf, hmtx = tt["glyf"], tt["hmtx"]
            for k, (kind, src, _ch) in enumerate(self.extra):
                name = "cfl%d" % k
                if kind == "clone":
                    glyf.glyphs[name] = copy.deepcopy(glyf[order[src]])
                    hmtx.metrics[name] = hmtx.metrics[order[src]]
                else:
                    g = Glyph()
                    g.numberOfContours = 0
                    glyf.glyphs[name] = g
                    hmtx.metrics[name] = (0, 0)
                newid[k] = len(order)
                order.append(name)
            tt.setGlyphOrder(order)
            if "post" in tt:
                tt["post"].formatType = 3.0
        buf = io.BytesIO()
        tt.save(buf)
        final, tomap = {}, {}
        for (g0, ch), (kind, v) in self.uses.items():
            gid = v if kind == "orig" else newid[v]
            final[(g0, ch)] = gid
            if ch:
                tomap[gid] = ch
        return buf.getvalue(), final, tomap


def _page_to_pdf(page):
    """Matrix from the DISPLAYED page (rotation applied, y down) to PDF space. Through
    pdf_geometry: page.transformation_matrix drops the box origin of a rotated page, and on a
    rotated page whose CropBox does not start at (0,0) an edited word was removed in place
    but redrawn ~100 pt away (10 Oct 2026)."""
    return geom.displayed_to_pdf(page)


def append_content(page, data):
    """Append a content stream after the page's own, whose graphics state is sealed off."""
    doc = page.parent
    cont = doc.xref_get_key(page.xref, "Contents")
    if cont[0] != "null":
        page.wrap_contents()
        cont = doc.xref_get_key(page.xref, "Contents")
    nx = doc.get_new_xref()
    doc.update_object(nx, "<<>>")
    doc.update_stream(nx, data, new=True)
    if cont[0] == "null":
        doc.xref_set_key(page.xref, "Contents", "%d 0 R" % nx)
    elif cont[0] == "array":
        doc.xref_set_key(page.xref, "Contents", cont[1].rstrip()[:-1] + " %d 0 R]" % nx)
    else:
        doc.xref_set_key(page.xref, "Contents", "[%s %d 0 R]" % (cont[1], nx))


class TextLayerWriter:
    """Queue lines per page, then flush(): one subset per face and page, own
    ToUnicode, one appended content stream per page."""

    def __init__(self, doc):
        self.doc = doc
        self.pages = {}       # page.number -> (page, [draw])
        self.serial = 0

    def draw(self, page, line, x, baseline, color=(0, 0, 0), bold_fake=False, italic_fake=False):
        """x: LEFT edge of the line, baseline: y, both in DISPLAYED page coordinates."""
        self.pages.setdefault(page.number, (page, []))[1].append(
            (line, float(x), float(baseline), color, bold_fake, italic_fake))

    def flush(self):
        for pno, (page, draws) in sorted(self.pages.items()):
            self._flush_page(page, draws)
        self.pages = {}

    def _flush_page(self, page, draws):
        emb = {}
        plan = []
        for line, x, base, color, bold_fake, italic_fake in draws:
            size = line.size
            items = []
            for face, gid, xe, ye, adv, t, rtl in line.glyphs:
                e = emb.setdefault(face.path, _Embedding(face))
                first, rest = (t[:1], t[1:]) if t else ("", "")
                gx, gy = x + xe * size, base - ye * size
                glyph = (face.path, e.key(gid, first), gx, gy)
                carriers = []
                if rest:
                    inset = min(_CARRIER_INSET, 0.25 * adv * size)
                    m = len(rest)
                    for k, ch in enumerate(rest):
                        if rtl:
                            cx = gx + inset + (m - 1 - k) * _CARRIER_STEP
                        else:
                            cx = gx + adv * size - inset - (m - 1 - k) * _CARRIER_STEP
                        carriers.append((face.path, e.key(None, ch), cx, gy))
                if rtl:
                    items.extend(reversed(carriers))
                    items.append(glyph)
                else:
                    items.append(glyph)
                    items.extend(carriers)
            plan.append((items, size, color, bold_fake, italic_fake))
        names = {}
        finals = {}
        for path, e in emb.items():
            data, final, tomap = e.build()
            self.serial += 1
            fname = "Fcfl%d%s" % (self.serial, os.urandom(2).hex())
            xref = page.insert_font(fontname=fname, fontbuffer=data)
            tu = self.doc.xref_get_key(xref, "ToUnicode")
            if tu[0] == "xref":
                self.doc.update_stream(int(tu[1].split()[0]), _tounicode(tomap))
            else:
                nx = self.doc.get_new_xref()
                self.doc.update_object(nx, "<<>>")
                self.doc.update_stream(nx, _tounicode(tomap), new=True)
                self.doc.xref_set_key(xref, "ToUnicode", "%d 0 R" % nx)
            names[path] = fname
            finals[path] = final
        m = _page_to_pdf(page)
        # the glyph's own axes in PDF space: x along the displayed baseline, y up
        ex = fitz.Point(1, 0) * m - fitz.Point(0, 0) * m
        ey = fitz.Point(0, -1) * m - fitz.Point(0, 0) * m
        ops = ["q", "BT", "0 Tc 0 Tw 100 Tz 0 Ts"]
        for items, size, color, bold_fake, italic_fake in plan:
            r, g, b = color
            ops.append("%.4g %.4g %.4g rg %.4g %.4g %.4g RG" % (r, g, b, r, g, b))
            if bold_fake:
                ops.append("2 Tr %.3f w" % (size * 0.03))
            else:
                ops.append("0 Tr")
            skew = 0.2 if italic_fake else 0.0
            cur_font = None
            for path, key, gx, gy in items:
                if cur_font != path:
                    ops.append("/%s %.3f Tf" % (names[path], size))
                    cur_font = path
                p = fitz.Point(gx, gy) * m
                a, b_, c, d = ex.x, ex.y, ey.x + skew * ex.x, ey.y + skew * ex.y
                ops.append("%.5g %.5g %.5g %.5g %.3f %.3f Tm <%04x> Tj" % (a, b_, c, d, p.x, p.y, finals[path][key]))
        ops += ["ET", "Q"]
        append_content(page, ("\n".join(ops) + "\n").encode())


if __name__ == "__main__":
    # smoke test: python3 pdf_text_layer.py out.pdf "text" [family]
    out, text = sys.argv[1], sys.argv[2]
    fam = sys.argv[3] if len(sys.argv) > 3 else ""
    doc = fitz.open()
    pg = doc.new_page()
    w = TextLayerWriter(doc)
    st = font_stack(fam, False, False, text)
    w.draw(pg, layout(text, st, 18), 50, 100)
    w.flush()
    doc.save(out)
    print(fitz.open(out)[0].get_text())
