"""Narration text for PDF -> Audio.

Usage: python3 pdf_audio_text.py input.pdf output.txt
Prints one JSON line: {"ok": true, "pages": N, "chars": M, "dropped": {...}}

What it does that the old PdfPig extractor did not:
  * reading order from glyph positions (Arabic logical order), two-column pages even
    when a title, abstract, table or caption runs across both columns,
  * drops running headers/footers: a margin line repeated on half the pages (a book
    title) or on a nearby page (a chapter title heading that chapter's pages),
  * drops page numbers ("3", "Page 3 of 10", "- 3 -", "صفحة ٣ من ١٠"),
  * drops footnote calls ("year¹") and footnote text (small type under the body),
  * speaks a link or an e-mail address as its site ("northwind.com") instead of
    spelling it out, and never cuts it out of the middle of a sentence,
  * strips what a voice should not read: bullets, emoji and pictographs, tatweel,
    zero-width marks, table-of-contents dot leaders and their page numbers; arrows
    become a pause,
  * re-joins words hyphenated at a line end (-, U+2010, soft hyphen), merges wrapped
    lines into paragraphs, carries a paragraph over a page or column break without a
    full stop, and ends every paragraph with one so the voice pauses.
Exit code 0 on success (even with little text), 2 on failure.
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import fitz  # noqa: E402

from pdf_text_rows import page_flow, row_text_smart, paragraphs  # noqa: E402
from pdf_to_word import _BULLETS, _is_arabic_char, _line_chars, _merge_rows, _finish_row  # noqa: E402


def row_text(row):
    return row_text_smart(row, chips=False)


_URL = re.compile(r"(?i)\b(?:https?://\s?|www\.)\S+|\b[a-z0-9.-]+\.(?:com|net|org|io|co|edu|gov|me|info|eg|sa|ae|uk|de|fr|es|in)(?:/\S*)?")
_EMAIL = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
_HOST = re.compile(r"(?i)^(?:https?://\s?)?(?:www\.)?([a-z0-9.-]+?)\.?(?:[/?#:].*)?$")
# A link or address set inside an Arabic line can come out with spaces around every dot and
# hyphen ("https: //www. nile - trading. com. eg/ar/reports"); it was then read piece by piece.
_SPACED_URL = re.compile(r"(?i)\bhttps?\s*:\s*/\s*/\s*(?:www\s*\.\s*)?([a-z0-9]+(?:\s*[-.]\s*[a-z0-9]+)+)(?:\s*/\s*[\w.\-/]*)?")
_SPACED_EMAIL = re.compile(r"(?i)\b[\w.+-]+\s*@\s*([a-z0-9]+(?:\s*[-.]\s*[a-z0-9]+)+)")
_PAGE_NO = re.compile(
    r"^\s*[-–—•·|]*\s*(?:page|p\.?|pg\.?|صفحة|ص)?\s*[\d٠-٩]{1,4}\s*(?:(?:of|من|/|\\)\s*[\d٠-٩]{1,4})?\s*[-–—•·|]*\s*$",
    re.IGNORECASE,
)
_WS = re.compile(r"\s+")

# What a voice should not read (the text-to-speech page's PrepareSpokenText rules, plus arrows):
# BOM, zero-width space/joiners, word joiner, variation selectors and tatweel ("مـحـمـد").
_INVISIBLE = re.compile("[\ufeff\u200b\u200c\u200d\u2060\ufe0e\ufe0f\u0640]")
# Emoji and pictographs: the U+1F000-1FAFF planes, Misc Symbols, Dingbats and the
# symbol half of Misc Symbols and Arrows ("📌", "✅", "★", "✓").
_PICTO = re.compile("[\U0001F000-\U0001FAFF\u2600-\u2793\u2B1B-\u2BFF]")
# Arrows are a pause, not a word: "Key dates → registration opens" -> "Key dates, registration opens".
_ARROWS = re.compile("\\s*[\u2190-\u21FF\u2794-\u27BF\u27F0-\u27FF\u2900-\u297F\u2B00-\u2B1A]+\\s*")
_REPEAT_PUNCT = re.compile(r"([!?؟،,;؛])\1+")
# "Introduction ........ 1", "Chapter 3 — Pricing ______ 25", "المقدمة ....... ٣"
_LEADER = re.compile(r"\s*(?:[.·•…‥_]\s*){4,}(?:[\d٠-٩]{1,4}|[ivxlcdm]{1,6})?\s*$", re.IGNORECASE)

_HYPHENS = "-\u2010\u00ad"
_UNITS = {"m", "cm", "mm", "km", "dm", "ft", "in", "µm", "nm", "m/s"}


def _norm_key(text):
    """Header/footer identity: digits collapse so 'Page 3' == 'Page 4'."""
    t = re.sub(r"[\d٠-٩]+", "#", text.lower())
    return _WS.sub(" ", t).strip()


def _letters(text):
    return sum(1 for c in text if c.isalpha())


def _unspace_letters(text):
    """'T H E  D E F I N I T I V E  H A N D B O O K' -> 'THE DEFINITIVE HANDBOOK' so the
    voice reads words, not letters. Only when most tokens are single characters."""
    toks = [t for t in text.split(" ") if t]
    if len(toks) >= 4 and sum(1 for t in toks if len(t) == 1) >= 0.6 * len(toks):
        words = ["".join(w.split()) for w in re.split(r"\s{2,}", text) if w.strip()]
        if len(words) == 1 and len(toks) > 12:   # no wide gaps: fall back to letter groups
            return "".join(toks)
        return " ".join(words)
    return text


def _say_url(m):
    """A link is spoken as its site ("northwind-example.com"), like the text-to-speech page
    does. Deleting it left sentences such as "write to  for details"; a link wrapped over two
    lines left its second half behind ("al.pdf?lang=en&v=2")."""
    s = m.group(0)
    tail = re.search(r"[.,;:)]+$", s)
    tail = tail.group(0) if tail else ""
    h = _HOST.match(s[: len(s) - len(tail)] if tail else s)
    return (" " + h.group(1) + tail + " ") if h and h.group(1) else (" " + tail)


def _say_email(m):
    s = m.group(0)
    tail = re.search(r"[.,;:]+$", s)
    tail = tail.group(0) if tail else ""
    dom = s[: len(s) - len(tail)].partition("@")[2]
    return " " + dom + tail + " "


def _spoken(text):
    """Symbols a voice reads as noise (or by their emoji name) go; arrows become a pause."""
    text = _INVISIBLE.sub("", text)
    text = _PICTO.sub(" ", text)
    text = _ARROWS.sub(", ", text)
    text = _REPEAT_PUNCT.sub(r"\1", text)
    text = re.sub(r"(?:,\s*){2,}", ", ", text)
    return text


def _join_lines(a, b):
    """Rows of one paragraph, or the two halves of a paragraph split by a page break.
    A word hyphenated at the line end is whole again ("informa-" + "tion", WeasyPrint's
    U+2010 "précé‐" + "dent", a soft hyphen); a hyphen before a caseless script (Arabic,
    Devanagari) or inside a link is a compound hyphen and stays, without the space."""
    if not a:
        return b
    if a[-1:] in _HYPHENS and a[-2:-1].isalnum() and b[:1].isalnum():
        last = a.rsplit(" ", 1)[-1]
        if a[-1] == "\u00ad":
            return a[:-1] + b
        if "@" in last or "://" in last or last.lower().startswith("www."):
            return a[:-1] + "-" + b
        if b[:1].islower():
            return a[:-1] + b
        if not b[:1].isupper():
            return a[:-1] + "-" + b
    return a + " " + b


def _superscript_free(row):
    """Footnote calls ("year¹", "quarter2") are superscript digits in a smaller size: those
    glyphs go, so the voice does not read "year one". PyMuPDF flags superscript spans (bit 0).
    An exponent after a unit ("m²", "cm³") or after a digit ("10³") stays."""
    size = row.get("size") or 0
    if not size:
        return row
    changed = False
    toks = []
    for t in row["tokens"]:
        chars = t.get("chars") or []
        keep = []
        for k, c in enumerate(chars):
            sup = (isinstance(c[3], dict) and c[3].get("flags", 0) & 1) and c[4] < 0.8 * size \
                and (c[0].isdigit() or c[0] in "*†‡§,")
            if sup:
                before = "".join(x[0] for x in keep)
                if before[-1:].isdigit() or (c[0] in "23" and before.lower() in _UNITS):
                    sup = False
            if not sup:
                keep.append(c)
        if len(keep) != len(chars):
            changed = True
            if not keep:
                continue
            t = dict(t, chars=keep, text="".join(c[0] for c in keep))
        toks.append(t)
    return dict(row, tokens=toks) if changed else row


def _split_x(blocks, pw):
    """Column gutter x for a page whose two columns share it with full-width blocks (a title,
    an abstract, a table or a figure caption across both columns). pdf_to_word._page_columns
    gives up as soon as ONE block crosses the gutter, and then the rows of the two columns were
    merged line by line ("...supply chain rose to 32, and voluntary..."). Only multi-line blocks
    count as column text, so the cells of a table do not make a page two-column."""
    body = [b for b in blocks if b["lines"] >= 2]
    if len(body) < 2:
        return None
    best = None
    for frac in range(30, 71):
        x = pw * frac / 100.0
        left = [b for b in body if b["x1"] <= x]
        right = [b for b in body if b["x0"] >= x]
        if not left or not right:
            continue
        if not any(min(a["y1"], b["y1"]) - max(a["y0"], b["y0"]) > 0 for a in left for b in right):
            continue   # side by side, not stacked
        score = min(len(left), len(right))
        if best is None or score > best[0]:
            best = (score, x)
    return best[1] if best else None


def page_flow_bands(page):
    """pdf_text_rows.page_flow, with the page cut into bands at every block that runs across
    the column gutter: each band's columns are read in turn, then the full-width block."""
    pw = page.rect.width or 1.0
    flags = fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE
    raw = page.get_text("rawdict", flags=flags)
    frags, blocks = [], []
    for b in raw.get("blocks", []):
        if b.get("type") != 0:
            continue
        x0, y0, x1, y1 = b["bbox"]
        n = 0
        for ln in b.get("lines", []):
            chars = _line_chars(ln)
            if any(c[0] not in "  " for c in chars):
                frags.append({"chars": chars, "y0": ln["bbox"][1], "y1": ln["bbox"][3],
                              "x0": ln["bbox"][0], "x1": ln["bbox"][2], "block": len(blocks)})
                n += 1
        if n:
            blocks.append({"x0": x0, "x1": x1, "y0": y0, "y1": y1, "lines": n})
    if not frags:
        return False, []
    split = _split_x(blocks, pw)
    if split is None:
        return page_flow(page)
    all_text = "".join(c[0] for f in frags for c in f["chars"])
    page_rtl = sum(1 for c in all_text if _is_arabic_char(c)) > sum(1 for c in all_text if c.isascii() and c.isalpha())
    full = sorted((i for i, b in enumerate(blocks) if b["x0"] < split - 2 and b["x1"] > split + 2),
                  key=lambda i: blocks[i]["y0"])
    rank = {i: k for k, i in enumerate(full)}

    def region_of(y0, y1):
        cy = (y0 + y1) / 2.0
        return sum(1 for i in full if blocks[i]["y1"] <= cy)

    # Only a band that holds multi-line text on both sides of the gutter is read column by
    # column; a table or a list between two full-width blocks keeps its rows whole.
    sides = {}
    for i, b in enumerate(blocks):
        if i not in rank and b["lines"] >= 2:
            sides.setdefault(region_of(b["y0"], b["y1"]), set()).add("R" if (b["x0"] + b["x1"]) / 2.0 >= split else "L")
    two_col = {r for r, s in sides.items() if len(s) == 2}
    # Bands top to bottom: region 0, full-width block 0, region 1, ... A two-column region is
    # read column by column; everything else between them (full-width blocks, a table, a list
    # whose bullets sit left of the gutter) stays one flow, so rows keep their bullets and cells.
    band_group, gid, single = {}, 0, None
    for band in [b for r in range(len(full) + 1) for b in (("region", r), ("full", r))][:-1]:
        if band[0] == "region" and band[1] in two_col:
            band_group[band] = gid
            gid, single = gid + 1, None
        else:
            if single is None:
                single, gid = gid, gid + 1
            band_group[band] = single
    groups = {}
    for f in frags:
        if f["block"] in rank:
            key = (band_group[("full", rank[f["block"]])], -1)
        else:
            region = region_of(f["y0"], f["y1"])
            if region in two_col:
                right = (f["x0"] + f["x1"]) / 2.0 >= split
                key = (band_group[("region", region)], (0 if right else 1) if page_rtl else (1 if right else 0))
            else:
                key = (band_group[("region", region)], -1)
        groups.setdefault(key, []).append(f)
    flow = []
    for key in sorted(groups):
        rows = [r for r in (_finish_row(r, page_rtl) for r in _merge_rows(groups[key])) if r]
        rows.sort(key=lambda r: r["y0"])
        if key[1] == -1:
            cx0, cx1 = 0.0, pw
        elif (key[1] == 0) == page_rtl:
            cx0, cx1 = split, pw
        else:
            cx0, cx1 = 0.0, split
        if rows:
            flow.append((rows, cx0, cx1))
    return page_rtl, flow


def _clean_row(text):
    text = _unspace_letters(text)
    text = _EMAIL.sub(_say_email, text)
    text = _URL.sub(_say_url, text)
    text = _SPACED_EMAIL.sub(lambda m: " " + _WS.sub("", m.group(1)) + " ", text)
    text = _SPACED_URL.sub(lambda m: " " + _WS.sub("", m.group(1)) + " ", text)
    text = _spoken(text)
    text = text.lstrip(_BULLETS + " \t,")
    text = re.sub(r"^[\(\[]?[\d٠-٩]{1,2}[\)\].]\s+", "", text) if len(text) > 6 else text  # "1. " list numbers
    text = _WS.sub(" ", text).strip(" \t-–—|·•→")
    text = re.sub(r"\s+([.,;:!?؟،؛])", r"\1", text)
    return _tanween_alef(text)


# The alef of a word-final tanween ("مثلاً", "أولاً", "جداً") is the part of an Arabic
# PDF's text layer most often lost or split off: until 10 Oct 2026 Chrome PDFs were read
# "مثلً أول" ("mathan" and "or" for "for example, firstly"), "دائمً ا", "أول اً". The row
# rebuild in pdf_to_word.py now keeps those clusters whole; this is the net under it for a
# voice. A lone alef (with its tanween) standing after a word goes back onto it, and a
# fathatan that ends a word on a letter that writes its tanween with an alef — anything but
# ة, ى, a hamza or an alef — gets the alef back.
_ALEF_TANWEEN_LETTER = "[\u0628\u062A-\u063A\u0641-\u0648\u064A]"     # not ة ى ء آ أ ؤ إ ئ ا ـ
_SPLIT_TANWEEN = re.compile("(?<=" + _ALEF_TANWEEN_LETTER + "\u064B)\\s+\u0627(?![\u0621-\u0652])"
                            "|(?<=" + _ALEF_TANWEEN_LETTER + ")\\s+(?=(?:\u0627\u064B|\u064B\u0627)(?![\u0621-\u0652]))")
_BARE_TANWEEN = re.compile("(?<=" + _ALEF_TANWEEN_LETTER + ")\u064B(?![\u0621-\u0652])")


def _tanween_alef(text):
    if "\u064B" not in text:
        return text
    text = _SPLIT_TANWEEN.sub(lambda m: "\u0627" if m.group(0).strip() else "", text)
    return _BARE_TANWEEN.sub("\u064B\u0627", text)


def _end_sentence(text):
    if not text:
        return text
    if text[-1] in ".!?؟؛:…।॥":
        return text
    return text + "."


def _continues(prev, nxt):
    """Does the first paragraph of a page (or column) carry on the last one before it?"""
    if not prev["open"] or not nxt["text"]:
        return False
    first = nxt["text"][:1]
    if first.islower() or first in ",;)":
        return True
    # caseless scripts (Arabic, Devanagari): same face and size, and no end mark before the break
    return not first.isupper() and abs(prev["size"] - nxt["size"]) < 0.6 and prev["bold"] == nxt["bold"]


_LIST_START = re.compile(r"^\s*(?:[\(\[]?[\d٠-٩]{1,2}[\)\].]\s|[•●▪◦‣►▶✓✔➢*]\s?)")


def _list_items(rows):
    """"1. First ... 2. Second ..." arrive as one paragraph when the numbers are text, not
    bullets: each item becomes its own paragraph, so it ends with a pause of its own."""
    out = []
    for r in rows:
        if out and _LIST_START.match(row_text(r)):
            out.append([r])
        elif out:
            out[-1].append(r)
        else:
            out.append([r])
    return out


def _in_margin(it, items):
    """A top line with a gap under it (or a bottom line with a gap over it) of more than one
    and a half of its own line heights: the page margin, not the first or last line of text."""
    h = max(it["y1"] - it["y0"], 1.0)
    others = [o for o in items if o is not it]
    if it["top"] <= 0.12:
        below = [o["y0"] - it["y1"] for o in others if o["y0"] >= it["y1"] - 0.5]
        return not below or min(below) > 1.5 * h
    above = [it["y0"] - o["y1"] for o in others if o["y1"] <= it["y0"] + 0.5]
    return not above or min(above) > 1.5 * h


def build(pdf_path):
    doc = fitz.open(pdf_path)
    n_pages = len(doc)
    pages = []  # per page: list of dict(rows, col, colstart, y0, y1, top, bot, size, bold, chars)
    for page in doc:
        ph = page.rect.height or 842.0
        _, flow = page_flow_bands(page)
        items = []
        for ci, (rows, cx0, cx1) in enumerate(flow):
            rows = [r for r in (_superscript_free(r) for r in rows) if r["tokens"]]
            if not rows:
                continue
            for pi, para in enumerate(p for whole in paragraphs(rows, cx0, cx1) for p in _list_items(whole)):
                items.append({"rows": para, "col": ci, "colstart": pi == 0,
                              "y0": para[0]["y0"], "y1": para[-1]["y1"],
                              "top": para[0]["y0"] / ph, "bot": para[-1]["y1"] / ph,
                              "size": max(r["size"] for r in para), "bold": para[0]["bold"],
                              "chars": sum(len(t["text"]) for r in para for t in r["tokens"])})
        pages.append(items)

    # the body size: the size most characters are set in
    by_size = {}
    for items in pages:
        for it in items:
            by_size[round(it["size"])] = by_size.get(round(it["size"]), 0) + it["chars"]
    body = max(by_size, key=by_size.get) if by_size else 0

    # Running headers / footers: a short line near the top or bottom edge that repeats on half
    # the pages (a book title) - or, when it stands apart in the margin (one row, at most 80
    # characters, not larger than the body text, a clear gap from the text next to it), on a
    # page within six of this one: a chapter title that heads only that chapter's odd pages was
    # read aloud 20 times in a 40-page test book. 3+ pages only.
    repeated = set()
    if n_pages >= 3:
        seen, near = {}, {}
        for pi, items in enumerate(pages):
            for it in items:
                if not (it["top"] <= 0.12 or it["bot"] >= 0.88) or len(it["rows"]) > 2:
                    continue
                k = _norm_key(" ".join(row_text(r) for r in it["rows"]))
                if not k:
                    continue
                seen[k] = seen.get(k, 0) + 1
                if len(it["rows"]) == 1 and len(k) <= 80 and it["size"] <= body + 0.5 and _in_margin(it, items):
                    near.setdefault(k, []).append(pi)
        repeated = {k for k, c in seen.items() if c >= max(2, n_pages // 2)}
        repeated |= {k for k, pis in near.items() if any(0 < b - a <= 6 for a, b in zip(pis, pis[1:]))}

    dropped = {"header_footer": 0, "page_number": 0, "link_only": 0, "decoration": 0, "footnote": 0}
    paras = []   # dict(text, open, size, bold)
    for items in pages:
        first = True
        last_body_bot = max((it["bot"] for it in items if round(it["size"]) >= body), default=0)
        for it in items:
            raw = " ".join(row_text(r) for r in it["rows"])
            key = _norm_key(raw)
            if key in repeated and (it["top"] <= 0.12 or it["bot"] >= 0.88):
                dropped["header_footer"] += 1
                continue
            if _PAGE_NO.match(raw):
                dropped["page_number"] += 1
                continue
            # footnote text: small type below the page's last paragraph of body text
            if body and it["size"] <= 0.85 * body and it["top"] >= max(0.6, last_body_bot - 0.01):
                dropped["footnote"] += 1
                continue
            joined = ""
            for r in it["rows"]:
                joined = _join_lines(joined, _LEADER.sub("", row_text(r)))
            text = _clean_row(joined)
            if not text:
                if _URL.search(raw) or _EMAIL.search(raw):
                    dropped["link_only"] += 1
                else:
                    dropped["decoration"] += 1
                continue
            if _letters(text) < 2:
                dropped["decoration"] += 1
                continue
            cur = {"text": text, "open": text[-1] not in ".!?؟؛:…।॥\"'»”)",
                   "size": it["size"], "bold": it["bold"]}
            starts_flow = first or it["colstart"]
            first = False
            # A paragraph that runs over a page or column break is one paragraph: no full stop,
            # no pause in the middle of the sentence (9 of 10 page breaks in an 11-page report).
            if starts_flow and paras and _continues(paras[-1], cur):
                paras[-1]["text"] = _join_lines(paras[-1]["text"], cur["text"])
                paras[-1]["open"] = cur["open"]
                paras[-1]["size"], paras[-1]["bold"] = cur["size"], cur["bold"]
                continue
            paras.append(cur)

    text = "\n\n".join(_end_sentence(p["text"]) for p in paras)
    return text, n_pages, dropped


# ── The document's language, for choosing the voice ─────────────────────────────
# The page used to preselect a voice by the PAGE's language: an English PDF on the Arabic page
# was read by an Arabic voice, an Arabic PDF on the English page by an English one ("spelled
# letters", as the text-to-speech page's own warning puts it). The API now picks the voice from
# this summary. Latin-script languages are told apart by their commonest words.
_STOP = {
    "en": "the and of to in is that for with as on are this by be it from at or an was which have not",
    "fr": "le la les des et de du un une est que pour dans qui sur par pas au aux avec ce cette sont ont",
    "es": "el la los las de del y en que un una es por para con se no su al lo como más",
    "de": "der die das und ist nicht mit den dem ein eine zu von für auf sich des im auch als",
    "pt": "o a os as de do da dos das e em um uma que para com não por se no na é",
    "it": "il lo la gli le di del della e che un una per con non è nel sono alla",
    "tr": "ve bir bu da de ile için olarak olan gibi daha çok en ne değil ama",
    "id": "yang dan di ke dari ini itu dengan untuk pada adalah tidak akan juga dalam",
    "pl": "i w na z się nie do to że jest o jak po ale od za przez dla",
}
_STOP = {k: set(v.split()) for k, v in _STOP.items()}


def _script_of(ch):
    o = ord(ch)
    if _is_arabic_char(ch):
        return "arab"
    if 0x0900 <= o <= 0x097F:
        return "deva"
    if ch.isalpha() and (o < 0x0250 or 0x1E00 <= o <= 0x1EFF):
        return "latn"
    return None


_NOT_WORDS = re.compile(r"\S*[.@/:]\S*[a-z]\S*|\S*\d\S*", re.IGNORECASE)   # sites, addresses, numbers


def paragraph_script(text):
    """The script of a paragraph: the one at least 60 % of its letters and at least 20 letters
    are in, not counting site names or addresses ("nile-trading.com.eg" inside an Arabic
    sentence). None for a short or mixed line, which then stays with its neighbours' voice.
    Scripts/tts_parallel.py applies the same rule when it assigns voices."""
    counts = {}
    for ch in _NOT_WORDS.sub(" ", text):
        s = _script_of(ch)
        if s:
            counts[s] = counts.get(s, 0) + 1
    if not counts:
        return None
    best = max(counts, key=counts.get)
    return best if counts[best] >= 20 and counts[best] >= 0.6 * sum(counts.values()) else None


def language_summary(text):
    """{"lang": "ar"|"hi"|"en"|"fr"|..|None, "scripts": {..: share}, "latin_lang": .., "latin_hits": n,
    "mixed": bool}. "mixed" = paragraphs in more than one script (an Arabic document with English
    paragraphs): each script is then read by its own voice."""
    counts = {}
    for ch in text:
        s = _script_of(ch)
        if s:
            counts[s] = counts.get(s, 0) + 1
    total = sum(counts.values()) or 1
    scripts = {k: round(v / total, 3) for k, v in counts.items()}
    paras = [p for p in text.split("\n") if p.strip()]
    para_scripts = [paragraph_script(p) for p in paras]
    latin_words = []
    for p, sc in zip(paras, para_scripts):
        if sc == "latn":
            latin_words.extend(re.findall(r"[a-zà-ÿçğışöüąćęłńóśźż]+", p.lower()))
    hits = {k: sum(1 for w in latin_words if w in stop) for k, stop in _STOP.items()}
    latin_lang = max(hits, key=hits.get) if latin_words else None
    latin_hits = hits.get(latin_lang, 0) if latin_lang else 0
    if latin_hits < 5:
        latin_lang = None          # too little text to tell: the API falls back to the page's language
    primary = max(counts, key=counts.get) if counts else None
    lang = {"arab": "ar", "deva": "hi"}.get(primary, latin_lang) if primary else None
    return {"lang": lang, "scripts": scripts, "latin_lang": latin_lang, "latin_hits": latin_hits,
            "mixed": len({sc for sc in para_scripts if sc}) > 1}


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"ok": False, "error": "usage: pdf_audio_text.py in.pdf out.txt"}))
        sys.exit(2)
    try:
        text, n_pages, dropped = build(sys.argv[1])
        with open(sys.argv[2], "w", encoding="utf-8") as f:
            f.write(text)
        summary = {"ok": True, "pages": n_pages, "chars": len(text), "dropped": dropped,
                   "arabic": sum(1 for c in text if _is_arabic_char(c)) > 0}
        summary.update(language_summary(text))
        print(json.dumps(summary, ensure_ascii=False))
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": str(e)}))
        sys.exit(2)


if __name__ == "__main__":
    main()
