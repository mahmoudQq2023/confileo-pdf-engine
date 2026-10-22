#!/usr/bin/env python3
"""
pdf_form.py — read and fill AcroForm fields, and flatten the result.

Usage:
  python3 pdf_form.py fields in.pdf
      → one JSON line: {"pages":[{"index","width","height","rotation"}],
        "fields":[…], "groups":[…], "hasForm", "xfa"}
        Rectangles are in DISPLAYED page space (rotation applied, CropBox
        origin, y down) — the space pdf.js draws the page in.
  python3 pdf_form.py page in.pdf <index> out.png [scale]
      → one page as PNG. Kept for old clients; the page renders its own
        previews with pdf.js now.
  python3 pdf_form.py fill in.pdf out.pdf values.json [--flatten] [--overlay-pipe]
      → writes the values and builds every text/choice appearance; with
        --flatten the fields are then baked into the page.
        --overlay-pipe: after the values are written, prints ONE JSON line
          {"overlay": {"html": path, "width": pt, "height": pt, "boxes": n}}
        and reads ONE line from stdin: the path of that HTML printed to PDF
        (one page per box, see _overlay_html), or an empty line to have MuPDF
        draw the boxes instead. The caller keeps one Chromium for this.
Exit 0 on success, 3 = encrypted (needs a password), 4 = not a readable PDF,
2 = any other failure. Errors are one JSON line on stderr: {"error", "code"}.

What "correct output" means here, and why each piece is built the way it is
(measured 25 Sept 2026 on a 3-page test form: A4, a Letter page with /Rotate 90,
a CropBox-offset page; 11 text fields incl. a 14-cell comb, checkboxes whose
on-states are /Ja /Oui /Approved, two radio groups, a combo and a list box):

* Checkboxes and radios are NAMES, not booleans. An unticked box is /Off and a
  ticked one is whatever on-state its appearance dictionary defines (/Ja,
  /Oui, /1 …). The old reader returned the string 'Off', which is truthy, so
  every box arrived ticked on the page; and fill turned on EVERY widget of a
  radio group, leaving /V at the last one ("female" came back as /other).
  Now: checkboxes are real booleans, radio groups are one question with
  options, and fill writes the chosen widget's own on-state into the parent's
  /V and every kid's /AS (the chosen one on, the rest /Off).

* Text appearances are built here, not by MuPDF's widget.update(). For a
  value like "محمد إبراهيم Ali" MuPDF writes the runs in the wrong visual order
  and clips the second word; with an auto (0) font size it writes "0 Tf" and a
  zero-height clip, i.e. nothing at all. Each text/choice value is laid out as
  one "box" (alignment from /Q, font size from /DA with 0 = fit, multiline
  wrapping, comb cells for /MaxLen fields, list-box lines), printed by
  Chromium (the caller's shared browser) into a one-box-per-page PDF, and that
  page becomes the widget's /AP /N — background and border from /MK drawn
  underneath, the widget's /MK /R rotation in the form /Matrix. Chrome keeps
  Arabic as logical text (Amiri, same finding as Text-to-PDF): the value is
  searchable and copyable, where MuPDF's own shaping writes presentation forms
  (U+FB50..) in visual order. Without Chromium the same boxes are drawn with
  MuPDF's HTML engine (HarfBuzz shaping, correct on screen, presentation-form
  text layer) so the file is still right to look at.

* NeedAppearances is cleared once every field has a real appearance. Left on,
  viewers rebuild the appearance from the field's base-14 font and the Arabic
  is lost again (poppler, pdf.js, PDFium).

* A flattened file is the SAME appearances baked into the page (doc.bake), so
  field borders and backgrounds survive and every value is extracted once.
  The old flatten deleted the text widgets and stamped one overlay page using
  the FIRST page's size for every page, which put values on rotated or
  differently sized pages in the wrong place.
"""
import base64
import html
import io
import json
import math
import os
import re
import sys
import tempfile
import unicodedata

import fitz  # PyMuPDF >= 1.26

HERE = os.path.dirname(os.path.abspath(__file__))

TYPE_NAMES = {
    fitz.PDF_WIDGET_TYPE_TEXT: 'text',
    fitz.PDF_WIDGET_TYPE_CHECKBOX: 'checkbox',
    fitz.PDF_WIDGET_TYPE_RADIOBUTTON: 'radio',
    fitz.PDF_WIDGET_TYPE_COMBOBOX: 'combobox',
    fitz.PDF_WIDGET_TYPE_LISTBOX: 'listbox',
    fitz.PDF_WIDGET_TYPE_SIGNATURE: 'signature',
    fitz.PDF_WIDGET_TYPE_BUTTON: 'button',
}

# field flags (PDF 32000-1, 12.7.3.1 and the per-type tables)
FF_READONLY = 1
FF_REQUIRED = 1 << 1
FF_MULTILINE = 1 << 12
FF_PASSWORD = 1 << 13
FF_NOTOGGLEOFF = 1 << 14
FF_FILESELECT = 1 << 20
FF_COMB = 1 << 24
FF_EDIT = 1 << 18
FF_MULTISELECT = 1 << 21

# Fonts for the Chromium overlay, embedded as data URIs (the page is loaded
# with SetContent, so no file:// URL would resolve). Arial is metric-compatible
# with the Helvetica almost every form declares; Arabic goes to Amiri through
# unicode-range because Chrome writes Amiri's Arabic as REAL letters in the PDF
# text layer while Arial, Noto Sans Arabic and IBM Plex Sans Arabic come back
# with broken lam-alef / split words (measured 25 Sept 2026, Chrome 149:
# "لا إله إلا الله" → Amiri exact; Arial "ل إله إال هللا"; Noto "يمحمد  إ براه يم").
FONT_DIRS = [os.path.join(HERE, '..', 'Fonts'), os.path.join(HERE, '..', 'Fonts', 'handwriting')]
ARABIC_RANGES = 'U+0600-06FF, U+0750-077F, U+0870-08FF, U+FB50-FDFF, U+FE70-FEFF'
NON_ARABIC_RANGES = 'U+0000-05FF, U+0700-074F, U+0780-086F, U+0900-FB4F, U+FE00-FE6F, U+FF00-10FFFF'

# Same lookup Text-to-PDF uses for the MuPDF fallback; MuPDF also carries its
# own Noto Naskh for any glyph the face lacks.
LIST_HIGHLIGHT = '#99c1da'   # Acrobat's list-box selection colour (0.6 0.757 0.855)


def die(code, msg, key='failed'):
    print(json.dumps({'error': msg, 'code': key}), file=sys.stderr)
    sys.exit(code)


def open_doc(path):
    try:
        doc = fitz.open(path)
    except Exception as e:  # noqa: BLE001
        die(4, f'cannot open: {e}', 'invalid')
    if not doc.is_pdf:
        die(4, 'not a PDF', 'invalid')
    if doc.needs_pass:
        die(3, 'encrypted', 'encrypted')
    return doc


def r2(v):
    return round(float(v), 2)


# ── small PDF-syntax helpers ──────────────────────────────────────────
def unescape_name(tok):
    """'Choice#201' → 'Choice 1', '#D9#86#D8#B9#D9#85' → 'نعم'.

    PyMuPDF hands out on-states in PDF syntax (on_state(), button_states())
    but /V and /AS already decoded (xref_get_key), so everything is compared
    in decoded form and written back with the widget's own raw token.
    """
    if not tok or '#' not in tok:
        return tok or ''
    out, i = bytearray(), 0
    while i < len(tok):
        if tok[i] == '#' and re.fullmatch(r'[0-9A-Fa-f]{2}', tok[i + 1:i + 3] or ''):
            out.append(int(tok[i + 1:i + 3], 16))
            i += 3
        else:
            out.extend(tok[i].encode('utf-8'))
            i += 1
    try:
        return out.decode('utf-8')
    except UnicodeDecodeError:
        return out.decode('latin-1')


def escape_name(s):
    out = []
    for b in s.encode('utf-8'):
        c = chr(b)
        if 0x21 <= b <= 0x7e and c not in '()<>[]{}/%#':
            out.append(c)
        else:
            out.append('#%02X' % b)
    return ''.join(out)


def _pdf_strings(src):
    """Decode the strings of a PDF array source text: '[(a)<FEFF…>]' → ['a', …]."""
    vals, i, n = [], 0, len(src)
    while i < n:
        c = src[i]
        if c == '(':
            depth, buf, i = 1, bytearray(), i + 1
            while i < n and depth:
                ch = src[i]
                if ch == '\\' and i + 1 < n:
                    nx = src[i + 1]
                    esc = {'n': 10, 'r': 13, 't': 9, 'b': 8, 'f': 12}
                    if nx in esc:
                        buf.append(esc[nx]); i += 2
                    elif nx.isdigit():
                        m = re.match(r'[0-7]{1,3}', src[i + 1:])
                        buf.append(int(m.group(0), 8) & 0xFF); i += 1 + len(m.group(0))
                    elif nx in '\r\n':
                        i += 2
                    else:
                        buf.extend(nx.encode('latin-1', 'replace')); i += 2
                    continue
                if ch == '(':
                    depth += 1
                elif ch == ')':
                    depth -= 1
                    if not depth:
                        i += 1
                        break
                buf.extend(ch.encode('latin-1', 'replace'))
                i += 1
            vals.append(_decode_text(bytes(buf)))
        elif c == '<':
            j = src.find('>', i)
            hx = re.sub(r'\s', '', src[i + 1:j])
            if len(hx) % 2:
                hx += '0'
            vals.append(_decode_text(bytes.fromhex(hx)))
            i = j + 1
        else:
            i += 1
    return vals


def _decode_text(b):
    if b[:2] == b'\xfe\xff':
        return b[2:].decode('utf-16-be', 'replace')
    if b[:3] == b'\xef\xbb\xbf':
        return b[3:].decode('utf-8', 'replace')
    return b.decode('latin-1')


def _get(doc, xref, key):
    try:
        return doc.xref_get_key(xref, key)
    except Exception:  # noqa: BLE001
        return ('null', 'null')


def _parent(doc, xref):
    t, v = _get(doc, xref, 'Parent')
    if t == 'xref':
        try:
            return int(v.split()[0])
        except ValueError:
            return 0
    return 0


def _inherited(doc, xref, key, acro_default=None):
    """Walk the field hierarchy for an inheritable key (/Q, /DA, /V, /Opt…)."""
    seen = set()
    while xref and xref not in seen and len(seen) < 16:
        seen.add(xref)
        t, v = _get(doc, xref, key)
        if t != 'null':
            return t, v
        xref = _parent(doc, xref)
    return acro_default or ('null', 'null')


def _color(vals):
    """PDF colour array → '#rrggbb' (None for no colour / transparent)."""
    if not vals:
        return None
    try:
        v = [float(x) for x in vals]
    except (TypeError, ValueError):
        return None
    if len(v) == 1:
        r = g = b = v[0]
    elif len(v) == 3:
        r, g, b = v
    elif len(v) == 4:
        c, m, y, k = v
        r, g, b = (1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)
    else:
        return None
    return '#%02x%02x%02x' % tuple(max(0, min(255, round(x * 255))) for x in (r, g, b))


def _rgb_ops(hexcol, stroke=False):
    r, g, b = (int(hexcol[i:i + 2], 16) / 255 for i in (1, 3, 5))
    return f'{r:.3f} {g:.3f} {b:.3f} {"RG" if stroke else "rg"}'


def first_strong_rtl(s):
    """Base direction the way dir="auto" decides it: the first strong character."""
    for ch in s or '':
        bd = unicodedata.bidirectional(ch)
        if bd in ('R', 'AL'):
            return True
        if bd == 'L':
            return False
    return False


def graphemes(s):
    """Split into user-perceived characters (base + combining marks / joiners)."""
    out = []
    for ch in s:
        if out and (unicodedata.combining(ch) or ch in '‍️︎' or out[-1].endswith('‍')):
            out[-1] += ch
        else:
            out.append(ch)
    return out


# ── reading the form ──────────────────────────────────────────────────
def _nearby_label(words, r, side, stop_rects=(), reach=0.0):
    """Words printed next to a widget, in displayed space.

    side 'right'/'left': on the same line, starting close to the box and
    continuing while the gaps stay word-sized; 'above': the line just over it.
    Used only when the field has no tooltip (/TU) — a radio kid named /Choice2
    or a field named "topmostSubform[0].Page1[0].f1_01[0]" means nothing to a
    person, the words printed beside the box do.
    """
    h = max(4.0, r.height)
    if h > 28:
        # a tall box (list, multi-line) is labelled beside its FIRST line
        r = fitz.Rect(r.x0, r.y0, r.x1, r.y0 + 20)
        h = 20.0
    # text inside another widget is that widget's own appearance ("Giza"), not a label
    words = [w for w in words if not any(s.x0 <= (w[0] + w[2]) / 2 <= s.x1 and s.y0 <= (w[1] + w[3]) / 2 <= s.y1
                                         for s in stop_rects)]
    if side in ('right', 'left'):
        band = [w for w in words if not (w[3] < r.y0 + h * 0.15 or w[1] > r.y1 - h * 0.15)]
        band = [w for w in band if (w[1] + w[3]) / 2 >= r.y0 - h * 0.4 and (w[1] + w[3]) / 2 <= r.y1 + h * 0.4]
        if side == 'right':
            cand = sorted((w for w in band if w[0] >= r.x1 - 1), key=lambda w: w[0])
            limit = r.x1 + max(3 * h, 18, reach)
        else:
            cand = sorted((w for w in band if w[2] <= r.x0 + 1), key=lambda w: -w[2])
            limit = r.x0 - max(3 * h, 18, reach)
        mid = (r.y0 + r.y1) / 2
        between = [s for s in stop_rects if s.y0 - 2 <= mid <= s.y1 + 2]
        picked, edge = [], None
        for w in cand:
            wh = max(4.0, w[3] - w[1])
            if side == 'right':
                start, prev = w[0], (r.x1 if edge is None else edge)
                if (edge is None and start > limit) or (edge is not None and start - edge > 1.2 * wh):
                    break
                # another widget in the gap: this word is ITS label
                if any(s.x0 >= prev - 1 and s.x1 <= start + 1 for s in between):
                    break
                picked.append(w)
                edge = w[2]
            else:
                end, prev = w[2], (r.x0 if edge is None else edge)
                if (edge is None and end < limit) or (edge is not None and edge - end > 1.2 * wh):
                    break
                if any(s.x0 >= end - 1 and s.x1 <= prev + 1 for s in between):
                    break
                picked.append(w)
                edge = w[0]
            if sum(len(p[4]) for p in picked) > 60:
                break
        picked.sort(key=lambda w: w[0])
    else:
        cand = [w for w in words if w[3] <= r.y0 + 1 and w[3] >= r.y0 - max(1.6 * h, 14)
                and w[2] >= r.x0 - 2 and w[0] <= r.x1 + 2]
        if not cand:
            return ''
        line_y = max(w[3] for w in cand)
        row = sorted((w for w in cand if abs(w[3] - line_y) < 3), key=lambda w: w[0])
        # the run of words nearest the box's left edge (a bilingual row holds two labels)
        start = min(range(len(row)), key=lambda i: abs(row[i][0] - r.x0))
        lo = hi = start
        while lo > 0 and row[lo][0] - row[lo - 1][2] < 1.2 * max(4.0, row[lo][3] - row[lo][1]):
            lo -= 1
        while hi + 1 < len(row) and row[hi + 1][0] - row[hi][2] < 1.2 * max(4.0, row[hi][3] - row[hi][1]):
            hi += 1
        picked = row[lo:hi + 1]
    # NFKC folds Arabic presentation forms (U+FB50..) back to letters
    text = unicodedata.normalize('NFKC', ' '.join(w[4] for w in picked)).strip(' :：*')
    if first_strong_rtl(text):
        text = unicodedata.normalize('NFKC', ' '.join(w[4] for w in sorted(picked, key=lambda w: -w[0]))).strip(' :：*')
    return text[:80]


def read_form(doc, want_labels=True):
    """One pass over every widget: what the page needs and what fill needs."""
    cat = doc.pdf_catalog()
    q_default = _get(doc, cat, 'AcroForm/Q')
    pages, widgets = [], []
    for pno in range(doc.page_count):
        page = doc[pno]
        rm = page.rotation_matrix
        pr = page.rect
        pages.append({'index': pno, 'width': r2(pr.width), 'height': r2(pr.height), 'rotation': page.rotation})
        words = None
        page_widgets = []
        for w in page.widgets():
            t = TYPE_NAMES.get(w.field_type)
            if t is None or t == 'button':
                continue
            xref = w.xref
            fx = xref if _get(doc, xref, 'T')[0] != 'null' else (_parent(doc, xref) or xref)
            ff = w.field_flags or 0
            try:
                mk_r = int(float(_get(doc, xref, 'MK/R')[1])) % 360
            except ValueError:
                mk_r = 0
            rect = fitz.Rect(w.rect)
            disp = rect * rm
            d = {
                'xref': xref, 'fx': fx, 'page': pno, 'name': w.field_name or f'field_{pno}_{xref}', 'type': t,
                'rect': rect, 'disp': fitz.Rect(disp).normalize(), 'flags': ff, 'mkR': mk_r if mk_r in (90, 180, 270) else 0,
                'readOnly': bool(ff & FF_READONLY), 'required': bool(ff & FF_REQUIRED),
                'tu': (w.field_label or '').strip(), 'hasAP': _get(doc, xref, 'AP/N')[0] != 'null',
            }
            ft, fv = _get(doc, xref, 'F')
            d['hidden'] = ft == 'int' and bool(int(fv) & 2)
            if t == 'text':
                maxlen = int(w.text_maxlen or 0)
                d.update(value=w.field_value if isinstance(w.field_value, str) else '',
                         multiline=bool(ff & FF_MULTILINE), password=bool(ff & FF_PASSWORD), maxLen=maxlen,
                         comb=bool(ff & FF_COMB) and maxlen > 0 and not (ff & (FF_MULTILINE | FF_PASSWORD | FF_FILESELECT)))
            elif t in ('checkbox', 'radio'):
                try:
                    states = [s for s in ((w.button_states() or {}).get('normal') or []) if s]
                except Exception:  # noqa: BLE001
                    states = []
                raw_on = next((s for s in states if s != 'Off'), None)
                if not raw_on and _get(doc, xref, 'AP/N')[0] == 'null':
                    raw_on = None   # no appearance at all: MuPDF draws one later
                # button_states() speaks PDF syntax ('Choice#201'); a token with no
                # '#' may still hold characters a name must escape
                d['rawOn'] = (raw_on if '#' in raw_on else escape_name(raw_on)) if raw_on else 'Yes'
                d['on'] = unescape_name(raw_on) if raw_on else 'Yes'
                d['apStates'] = len(states)
                vt, vv = _inherited(doc, fx, 'V')
                at, av = _get(doc, xref, 'AS')
                if vt == 'name':
                    d['checked'] = vv[1:] == d['on']
                else:
                    d['checked'] = at == 'name' and av[1:] == d['on']
                d['fieldV'] = vv[1:] if vt == 'name' else None
                if t == 'radio' or fx != xref:
                    # a kid's position in the parent's /Kids indexes the parent's /Opt
                    ot, ov = _inherited(doc, fx, 'Opt')
                    kt, kv = _get(doc, fx, 'Kids')
                    if ot == 'array' and kt == 'array':
                        kids = [int(k) for k in re.findall(r'(\d+)\s+0\s+R', kv)]
                        opts = _pdf_strings(ov)
                        if xref in kids and kids.index(xref) < len(opts):
                            d['optLabel'] = opts[kids.index(xref)].strip()
            elif t in ('combobox', 'listbox'):
                opts, labels = [], []
                for c in (w.choice_values or []):
                    if isinstance(c, (list, tuple)):
                        opts.append(str(c[0])); labels.append(str(c[1]) if len(c) > 1 else str(c[0]))
                    else:
                        opts.append(str(c)); labels.append(str(c))
                d.update(options=opts, optionLabels=labels, multiSelect=bool(ff & FF_MULTISELECT),
                         editable=t == 'combobox' and bool(ff & FF_EDIT))
                vt, vv = _inherited(doc, fx, 'V')
                if vt == 'array':
                    sel = _pdf_strings(vv)
                elif vt == 'string':
                    sel = [w.field_value] if isinstance(w.field_value, str) and w.field_value else _pdf_strings(vv)
                else:
                    sel = []
                if not sel:
                    it, iv = _inherited(doc, fx, 'I')
                    if it == 'array':
                        sel = [opts[int(i)] for i in re.findall(r'\d+', iv) if int(i) < len(opts)]
                d['value'] = sel if d['multiSelect'] else (sel[0] if sel else '')
                try:
                    ti = int(float(_inherited(doc, fx, 'TI')[1]))
                except ValueError:
                    ti = 0
                d['topIndex'] = max(0, ti)
            if t in ('text', 'combobox', 'listbox'):
                qt, qv = _inherited(doc, xref, 'Q', q_default)
                try:
                    d['align'] = int(float(qv)) if qt in ('int', 'float') else 0
                except ValueError:
                    d['align'] = 0
                d['fontSize'] = float(w.text_fontsize or 0)
                d['font'] = w.text_font or 'Helv'
                d['color'] = _color(w.text_color) or '#000000'
                d['bg'] = _color(w.fill_color)
                d['bc'] = _color(w.border_color)
                try:
                    d['bw'] = float(w.border_width if w.border_width is not None else 1)
                except (TypeError, ValueError):
                    d['bw'] = 1.0
                d['bs'] = (w.border_style or 'Solid')
                d['dashes'] = list(w.border_dashes or [])
            page_widgets.append(d)

        if want_labels and page_widgets:
            needs = [d for d in page_widgets if not d['tu'] or d['type'] in ('checkbox', 'radio')]
            if needs:
                try:
                    # extraction boxes are unrotated like the widget rects; move
                    # both to displayed space, where the printed text is horizontal
                    words = []
                    for wd in page.get_text('words'):
                        wr = (fitz.Rect(wd[:4]) * rm).normalize()
                        words.append((wr.x0, wr.y0, wr.x1, wr.y1, wd[4]))
                except Exception:  # noqa: BLE001
                    words = []
                boxes = [d['disp'] for d in page_widgets]
                # A right-to-left page mirrors the layout: a button's label is on
                # its LEFT and a field's label on its RIGHT (LibreOffice's Arabic
                # form, 25 Sept 2026, got every radio labelled with its neighbour's).
                text = ''.join(w[4] for w in words)
                rtl_page = len(re.findall('[\u0600-\u06ff\ufb50-\ufdff\ufe70-\ufeff]', text)) > len(re.findall('[A-Za-z]', text))
                after, before = ('left', 'right') if rtl_page else ('right', 'left')
                for d in needs:
                    r = d['disp']
                    others = [b for b in boxes if b != r]
                    if d['type'] in ('checkbox', 'radio'):
                        d['near'] = _nearby_label(words, r, after, others) or _nearby_label(words, r, before, others)
                        # the question a radio group asks sits further out, in the label column
                        d['nearQuestion'] = _nearby_label(words, r, before, others, reach=180)
                        d['rtlPage'] = rtl_page
                    else:
                        d['near'] = (_nearby_label(words, r, before, others, reach=160) or _nearby_label(words, r, 'above', others)
                                     or _nearby_label(words, r, after, others))
        widgets.extend(page_widgets)
    return pages, widgets


def _xfa_state(doc):
    t, _ = _get(doc, doc.pdf_catalog(), 'AcroForm/XFA')
    return t != 'null'


def _acro_set(doc, key, value):
    """Set/remove an /AcroForm key. Through the object itself: xref_set_key
    with a 'AcroForm/Key' path leaves a 'fitz: replace me!' placeholder when
    /AcroForm is an indirect object (seen 25 Sept 2026, PyMuPDF 1.26.7)."""
    cat = doc.pdf_catalog()
    t, v = _get(doc, cat, 'AcroForm')
    if t == 'xref':
        doc.xref_set_key(int(v.split()[0]), key, value)
    elif t == 'dict':
        doc.xref_set_key(cat, f'AcroForm/{key}', value)


def cmd_fields(path):
    doc = open_doc(path)
    pages, widgets = read_form(doc)
    fields, groups, by_name = [], [], {}
    for d in widgets:
        if d['type'] == 'signature':
            continue
        by_name.setdefault(d['name'], []).append(d)
    for d in widgets:
        t = d['type']
        if t == 'signature':
            continue
        r = d['disp']
        item = {
            'id': str(d['xref']), 'page': d['page'], 'name': d['name'], 'type': t,
            'label': d['tu'] or d.get('near') or None,
            'x': r2(r.x0), 'y': r2(r.y0), 'w': r2(r.width), 'h': r2(r.height),
            'readOnly': d['readOnly'], 'required': d['required'], 'multiline': bool(d.get('multiline')),
        }
        if d['hidden']:
            item['hidden'] = True
        if t == 'text':
            item.update(value=d['value'], maxLen=d['maxLen'], comb=d['comb'], align=d['align'],
                        fontSize=r2(d['fontSize']), password=d['password'])
        elif t in ('checkbox', 'radio'):
            # A real boolean. The old reader sent the string 'Off' for an
            # unticked box, and the page's !!value turned it into "ticked".
            item.update(value=bool(d['checked']), onState=d['on'])
            if t == 'radio' or len({x['on'] for x in by_name[d['name']]}) > 1:
                item['label'] = d.get('optLabel') or d.get('near') or d['on']
        else:
            item.update(value=d['value'], options=d['options'], optionLabels=d['optionLabels'],
                        multiSelect=d['multiSelect'], editable=d['editable'], align=d['align'],
                        fontSize=r2(d['fontSize']))
        fields.append(item)

    # Radio groups (and the rare checkbox set whose widgets have DIFFERENT
    # on-states, which behaves like a radio group that can be left empty).
    for name, ws in by_name.items():
        kind = ws[0]['type']
        if kind not in ('radio', 'checkbox'):
            continue
        ons = []
        for d in ws:
            if d['on'] not in ons:
                ons.append(d['on'])
        if kind == 'checkbox' and len(ons) < 2:
            continue
        first = ws[0]
        options, seen = [], set()
        for d in ws:
            if d['on'] in seen:
                continue
            seen.add(d['on'])
            options.append({'value': d['on'], 'label': d.get('optLabel') or d.get('near') or d['on'],
                            'page': d['page']})
        selected = next((d['on'] for d in ws if d['checked']), None)
        label = first['tu']
        if not label:
            # the option furthest towards the label column (leftmost, or rightmost on an RTL page)
            outer = min(ws, key=lambda d: (d['page'], -d['disp'].x1 if d.get('rtlPage') else d['disp'].x0))
            label = outer.get('nearQuestion') or name
            if any(label == o['label'] for o in options):
                label = name
        groups.append({'name': name, 'kind': kind, 'label': label, 'value': selected,
                       'required': first['required'], 'readOnly': first['readOnly'],
                       'noToggleToOff': bool(first['flags'] & FF_NOTOGGLEOFF) and kind == 'radio',
                       'options': options})

    fillable = [f for f in fields if not f['readOnly']]
    print(json.dumps({'pages': pages, 'fields': fields, 'groups': groups, 'hasForm': bool(fillable),
                      'xfa': _xfa_state(doc)}, ensure_ascii=False))


# ── page preview ──────────────────────────────────────────────────────
def cmd_page(path, index, out, scale):
    doc = open_doc(path)
    if index < 0 or index >= len(doc):
        die(2, 'page out of range')
    page = doc[index]
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    pix.save(out)
    print(json.dumps({'width': pix.width, 'height': pix.height, 'scale': scale}))


# ── fill: values ──────────────────────────────────────────────────────
def _as_bool(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        return v.strip().lower() in ('true', '1', 'yes', 'on')
    return False


def write_values(doc, widgets, values, warnings):
    """Write every value into the fields' own dictionaries. Returns the set of
    field names whose appearance must be rebuilt."""
    by_name = {}
    for d in widgets:
        by_name.setdefault(d['name'], []).append(d)
    written, changed = 0, set()
    for name, v in values.items():
        ws = by_name.get(name)
        if not ws:
            warnings.append(f'no field named {name!r}')
            continue
        if all(d['readOnly'] for d in ws):
            continue
        t = ws[0]['type']
        if t == 'text':
            s = '' if v is None else (v if isinstance(v, str) else str(v))
            s = s.replace('\r\n', '\n').replace('\r', '\n')
            maxlen = max((d['maxLen'] for d in ws), default=0)
            if maxlen and len(s) > maxlen:
                warnings.append(f'{name}: cut to {maxlen} characters (MaxLen)')
                s = s[:maxlen]
            if not ws[0]['multiline']:
                s = s.replace('\n', ' ')
            for fx in {d['fx'] for d in ws}:
                doc.xref_set_key(fx, 'V', fitz.get_pdf_str(s))
            for d in ws:
                d['value'] = s
            changed.add(name)
            written += 1
        elif t in ('checkbox', 'radio'):
            # Which on-state is chosen: a string names it (radio option, or a
            # checkbox set); a boolean means "this box's own on-state" for a
            # checkbox, and for a radio only false (clear) has a meaning.
            ons = [d['on'] for d in ws]
            if isinstance(v, str) and v not in ('', 'Off') and v.lower() not in ('true', 'false', '1', '0', 'yes', 'no', 'on', 'off'):
                chosen = v if v in ons else None
                if chosen is None:
                    warnings.append(f'{name}: {v!r} is not one of its options')
                    continue
            elif isinstance(v, str) and v in ons:
                chosen = v
            elif t == 'checkbox':
                chosen = ons[0] if _as_bool(v) else None
            else:
                if _as_bool(v):
                    # legacy request (old page sent true for the whole group):
                    # keep whatever the form already says rather than guess
                    continue
                chosen = None
            raw = next((d['rawOn'] for d in ws if d['on'] == chosen), None)
            for fx in {d['fx'] for d in ws}:
                doc.xref_set_key(fx, 'V', f'/{raw}' if raw else '/Off')
            for d in ws:
                on = chosen is not None and d['on'] == chosen
                doc.xref_set_key(d['xref'], 'AS', f'/{d["rawOn"]}' if on else '/Off')
                d['checked'] = on
            written += 1
        elif t in ('combobox', 'listbox'):
            d0 = ws[0]
            opts = d0['options']
            if isinstance(v, (list, tuple)):
                sel = [str(x) for x in v if x is not None and str(x) != '']
            else:
                sel = [] if v is None or str(v) == '' else [str(v)]
            if not d0['multiSelect']:
                sel = sel[:1]
            # accept a display label for an export value (old page sent labels)
            sel = [opts[d0['optionLabels'].index(s)] if s not in opts and s in d0['optionLabels'] else s for s in sel]
            if not d0['editable'] and any(s not in opts for s in sel):
                bad = [s for s in sel if s not in opts]
                warnings.append(f'{name}: {bad!r} not among the options')
                sel = [s for s in sel if s in opts]
            idx = sorted(opts.index(s) for s in sel if s in opts)
            for fx in {d['fx'] for d in ws}:
                if not sel:
                    doc.xref_set_key(fx, 'V', 'null')
                elif len(sel) == 1:
                    doc.xref_set_key(fx, 'V', fitz.get_pdf_str(sel[0]))
                else:
                    doc.xref_set_key(fx, 'V', '[' + ''.join(fitz.get_pdf_str(s) for s in sel) + ']')
                doc.xref_set_key(fx, 'I', ('[' + ' '.join(map(str, idx)) + ']') if idx else 'null')
            for d in ws:
                d['value'] = sel if d['multiSelect'] else (sel[0] if sel else '')
            changed.add(name)
            written += 1
    return written, changed


# ── fill: layout of the text/choice appearances ───────────────────────
def _family(fontname):
    n = (fontname or '').lower().replace('-', '').replace(',', '')
    bold = 'bold' in n or n in ('hebo', 'tibo', 'cobo', 'hebi', 'tibi', 'cobi')
    if n.startswith(('tiro', 'tibo', 'tiit', 'tibi', 'times')) or ('serif' in n and 'sans' not in n) or 'roman' in n:
        fam = 'serif'
    elif n.startswith(('cour', 'cobo', 'coob', 'cobi')) or 'mono' in n or 'courier' in n:
        fam = 'mono'
    else:
        fam = 'sans'
    return fam, bold


def make_box(d):
    """The appearance of one text/choice widget, in its own upright space
    (width/height swapped for a widget rotated 90/270 by /MK /R)."""
    r = d['rect']
    w, h = (r.height, r.width) if d['mkR'] in (90, 270) else (r.width, r.height)
    bw = d['bw'] if d.get('bc') else 0.0
    if d.get('bs', '').lower().startswith(('bevel', 'inset')):
        bw *= 2
    pad = max(1.0, bw) + 1.0
    fam, bold = _family(d['font'])
    fs = d['fontSize']
    box = {'w': r2(w), 'h': r2(h), 'fam': fam, 'bold': bold, 'color': d['color'], 'pad': r2(pad)}
    t = d['type']
    if t == 'text':
        text = d['value'] or ''
        if d['password']:
            text = '•' * len(text)
        rtl = first_strong_rtl(text)
        q = d['align']
        box.update(kind='text', text=text, rtl=rtl,
                   align='center' if q == 1 else ('right' if q == 2 else ('right' if rtl else 'left')))
        if d['comb'] and not d['multiline']:
            n = d['maxLen']
            cells = graphemes(text)[:n]
            if rtl:
                cells = cells[::-1]
            empty = n - len(cells)
            lead = {'left': 0, 'center': empty // 2, 'right': empty}[box['align']]
            box['cells'] = [''] * lead + cells + [''] * (empty - lead)
            cw = w / max(1, n)
            box['size'] = r2(fs if fs > 0 else max(4.0, min((h - 2) * 0.72, cw * 0.95)))
            box['auto'] = fs <= 0
        elif d['multiline']:
            box['multiline'] = True
            box['size'] = r2(fs if fs > 0 else 12.0)
            box['auto'] = fs <= 0
        else:
            box['size'] = r2(fs if fs > 0 else max(4.0, min(24.0, (h - 2 * pad) * 0.75)))
            box['auto'] = fs <= 0
    elif t == 'combobox':
        val = d['value'] or ''
        text = d['optionLabels'][d['options'].index(val)] if val in d['options'] else val
        rtl = first_strong_rtl(text)
        q = d['align']
        box.update(kind='text', text=text, rtl=rtl,
                   align='center' if q == 1 else ('right' if q == 2 else ('right' if rtl else 'left')),
                   size=r2(fs if fs > 0 else max(4.0, min(24.0, (h - 2 * pad) * 0.75))), auto=fs <= 0)
    else:  # listbox
        sel = d['value'] if isinstance(d['value'], list) else ([d['value']] if d['value'] else [])
        size = fs if fs > 0 else 12.0
        lh = size * 1.15
        visible = max(1, int((h - 2 * pad + 0.01) // lh))
        idx = [i for i, o in enumerate(d['options']) if o in sel]
        top = min(d.get('topIndex', 0), max(0, len(d['options']) - 1))
        if idx and not (top <= idx[0] < top + visible):
            top = max(0, min(idx[0], len(d['options']) - visible))
        lines = []
        for i in range(top, min(len(d['options']), top + visible + 1)):
            lab = d['optionLabels'][i]
            lines.append({'text': lab, 'sel': i in idx, 'rtl': first_strong_rtl(lab)})
        box.update(kind='list', lines=lines, size=r2(size), lh=r2(lh))
    return box


FIT_JS = """
(function(){
  var PX = 96 / 72;
  function fit(){
    var els = document.querySelectorAll('[data-fit]');
    for (var i = 0; i < els.length; i++) {
      var el = els[i], fs = parseFloat(el.getAttribute('data-fs')), min = 4;
      var aw = parseFloat(el.getAttribute('data-aw')) * PX, ah = parseFloat(el.getAttribute('data-ah')) * PX;
      var ml = el.getAttribute('data-fit') === 'ml';
      for (var k = 0; k < 60 && fs > min; k++) {
        var over = ml ? el.scrollHeight > ah + 0.5 : el.getBoundingClientRect().width > aw + 0.5;
        if (!over) break;
        fs = Math.max(min, fs * 0.94);
        el.style.fontSize = fs + 'pt';
      }
    }
    return els.length;
  }
  // This runs while the page is still parsing, before any layout has asked for
  // a face, so fonts.ready alone could resolve with nothing loaded: load every
  // declared face (only the ones these boxes need are declared) first.
  var loads = [];
  if (document.fonts && document.fonts.forEach) document.fonts.forEach(function (f) { loads.push(f.load().catch(function () {})); });
  window.__cfFit = Promise.all(loads)
    .then(function () { return document.fonts && document.fonts.ready; })
    .then(fit);
})();
"""


# What each face is used for, as subsetting ranges (the @font-face unicode-range
# decides which face draws a character; these only decide what the file keeps).
SUBSET_RANGES = {
    'latin': [(0x20, 0x24F), (0x2B0, 0x36F), (0x370, 0x4FF), (0x1E00, 0x1EFF), (0x2000, 0x206F),
              (0x20A0, 0x20CF), (0x2100, 0x22FF), (0x25A0, 0x25FF), (0xFB00, 0xFB06)],
    'arabic': [(0x20, 0x40), (0x600, 0x6FF), (0x750, 0x77F), (0x870, 0x8FF), (0xFB50, 0xFDFF),
               (0xFE70, 0xFEFF), (0x200C, 0x200F), (0x2010, 0x2027), (0x25CC, 0x25CC)],
}


def _font_bytes(path, ranges_key):
    """The face as embedded in the overlay page: a cached subset when fontTools is
    there (it is, via pdf2docx), else the whole file.

    The fonts ARE the overlay page's weight: 1.9 MB of base64 took Chrome 330 ms
    to take in per fill; subset (Arial 1021 -> 211 KB, Amiri 421 -> 331 KB, every
    GSUB/GPOS feature kept, Arabic extracted identically) 155 ms (25 Sept 2026).
    Subsetting costs ~1 s per face, so it is done once and kept outside the temp
    folder the API sweeps, keyed by the font's content."""
    with open(path, 'rb') as f:
        data = f.read()
    try:
        import hashlib
        digest = hashlib.sha1(data).hexdigest()[:16]
        base = os.environ.get('XDG_CACHE_HOME') or os.path.join(os.path.expanduser('~'), '.cache')
        for root in (os.path.join(base, 'confileo', 'form-fonts'), os.path.join(tempfile.gettempdir(), 'confileo-form-fonts')):
            cached = os.path.join(root, f'{os.path.splitext(os.path.basename(path))[0]}-{ranges_key}-{digest}.ttf')
            if os.path.exists(cached):
                with open(cached, 'rb') as f:
                    return f.read()
            try:
                os.makedirs(root, exist_ok=True)
            except OSError:
                continue
            from fontTools import subset
            from fontTools.ttLib import TTFont
            opts = subset.Options()
            opts.layout_features = ['*']      # shaping (init/medi/fina, lam-alef, marks) must survive
            opts.name_IDs = ['*']
            opts.notdef_outline = True
            opts.hinting = False              # PDF text needs no hinting
            opts.drop_tables += ['DSIG']
            font = TTFont(io.BytesIO(data))
            sub = subset.Subsetter(opts)
            sub.populate(unicodes={u for a, b in SUBSET_RANGES[ranges_key] for u in range(a, b + 1)})
            sub.subset(font)
            buf = io.BytesIO()
            font.save(buf)
            out = buf.getvalue()
            tmp = f'{cached}.{os.getpid()}.tmp'
            with open(tmp, 'wb') as f:
                f.write(out)
            os.replace(tmp, cached)   # atomic: two fills racing both end with a whole file
            return out
    except Exception:  # noqa: BLE001
        pass
    return data


def _font_faces(boxes):
    """@font-face rules for the families the boxes actually use (each face is
    ~0.4–1 MB, so a Latin-only form never carries Amiri and vice versa)."""
    has_ar = any(re.search('[؀-ۿݐ-ݿࡰ-ࣿﭐ-﷿ﹰ-﻿]',
                           b.get('text', '') + ''.join(c for c in b.get('cells', [])) + ''.join(l['text'] for l in b.get('lines', [])))
                 for b in boxes)
    need_bold = any(b['bold'] for b in boxes)
    fams = {b['fam'] for b in boxes}

    def face(family, fname, weight, ranges):
        for d in FONT_DIRS:
            p = os.path.join(d, fname)
            if os.path.exists(p):
                b64 = base64.b64encode(_font_bytes(p, 'arabic' if ranges == ARABIC_RANGES else 'latin')).decode('ascii')
                return (f"@font-face{{font-family:'{family}';font-weight:{weight};font-style:normal;"
                        f"src:url(data:font/ttf;base64,{b64}) format('truetype');unicode-range:{ranges};font-display:block}}")
        return ''

    css = []
    weights = [(400, 'Amiri-Regular.ttf', 'arial.ttf')] + ([(700, 'Amiri-Bold.ttf', 'arialbd.ttf')] if need_bold else [])
    for weight, amiri, arial in weights:
        if 'sans' in fams:
            css.append(face('cfSans', arial, weight, NON_ARABIC_RANGES))
        if has_ar:
            for fam in fams:
                css.append(face({'sans': 'cfSans', 'serif': 'cfSerif', 'mono': 'cfMono'}[fam], amiri, weight, ARABIC_RANGES))
    return ''.join(css)


def _overlay_html(boxes, pw, ph):
    """One page per box, each box at the page's top-left corner. One page SIZE
    for all (the largest box) — Chrome then needs no named pages and the
    appearance just clips the box's corner out of its page."""
    stacks = {'sans': "'cfSans',Arial,'Liberation Sans',Helvetica,sans-serif",
              'serif': "'cfSerif','Liberation Serif','Times New Roman',Times,serif",
              'mono': "'cfMono','Liberation Mono','Courier New',Courier,monospace"}
    parts = []
    for b in boxes:
        font = (f"font-family:{stacks[b['fam']]};font-weight:{700 if b['bold'] else 400};color:{b['color']};")
        pad = b['pad']
        inner = ''
        if b['kind'] == 'list':
            for ln in b['lines']:
                bg = f'background:{LIST_HIGHLIGHT};' if ln['sel'] else ''
                inner += (f'<div dir="{"rtl" if ln["rtl"] else "ltr"}" style="{bg}height:{b["lh"]}pt;line-height:{b["lh"]}pt;'
                          f'padding:0 {pad}pt;white-space:pre;overflow:hidden;font-size:{b["size"]}pt;'
                          f'text-align:{"right" if ln["rtl"] else "left"}">{html.escape(ln["text"])}</div>')
            inner = f'<div style="padding-top:{pad - 1}pt">{inner}</div>'
        elif 'cells' in b:
            n = len(b['cells'])
            cw = b['w'] / max(1, n)
            for c in b['cells']:
                inner += (f'<span style="display:flex;align-items:center;justify-content:center;width:{cw:.3f}pt;height:{b["h"]}pt;'
                          f'font-size:{b["size"]}pt;white-space:pre">{html.escape(c)}</span>')
            inner = f'<div style="display:flex;flex-direction:row;direction:ltr">{inner}</div>'
        elif b.get('multiline'):
            aw, ah = b['w'] - 2 * pad, b['h'] - 2 * pad
            inner = (f'<div data-fit="ml" data-fs="{b["size"]}" data-aw="{aw:.2f}" data-ah="{ah:.2f}" dir="{"rtl" if b["rtl"] else "ltr"}" '
                     f'style="margin:{pad}pt;width:{aw:.2f}pt;font-size:{b["size"]}pt;line-height:1.15;white-space:pre-wrap;'
                     f'overflow-wrap:anywhere;text-align:{b["align"]}">{html.escape(b["text"])}</div>')
        else:
            aw = b['w'] - 2 * pad
            just = {'left': 'flex-start', 'center': 'center', 'right': 'flex-end'}[b['align']]
            inner = (f'<div style="display:flex;align-items:center;justify-content:{just};width:{aw:.2f}pt;height:{b["h"]}pt;'
                     f'padding:0 {pad}pt"><span data-fit="1" data-fs="{b["size"]}" data-aw="{aw:.2f}" data-ah="{b["h"]}" '
                     f'dir="{"rtl" if b["rtl"] else "ltr"}" style="font-size:{b["size"]}pt;line-height:1.15;white-space:pre">'
                     f'{html.escape(b["text"])}</span></div>')
        parts.append(f'<div class="pg"><div class="bx" style="width:{b["w"]}pt;height:{b["h"]}pt;{font}">{inner}</div></div>')
    return ('<!doctype html><html><head><meta charset="utf-8"><style>' + _font_faces(boxes) +
            f'@page{{size:{pw}pt {ph}pt;margin:0}}'
            'html,body{margin:0;padding:0;background:transparent}'
            f'.pg{{position:relative;width:{pw}pt;height:{ph}pt;overflow:hidden}}'
            # break BEFORE every page but the first: the trailing <script> must
            # never leave a blank last page (the page count is checked)
            '.pg+.pg{break-before:page;page-break-before:always}'
            '.bx{position:absolute;left:0;top:0;overflow:hidden;box-sizing:border-box;'
            'font-kerning:normal;font-variant-ligatures:normal}'
            '</style></head><body>' + ''.join(parts) + '<script>' + FIT_JS + '</script></body></html>')


def _mupdf_overlay(boxes, pw, ph):
    """The same boxes drawn by MuPDF's HTML engine, when Chromium is not there.

    One DocumentWriter for all boxes, one Story per piece of text: every
    page.insert_htmlbox() call embeds its own copy of the fonts, which took the
    3-page test form from 40 KB to 1.4 MB and 192 fields from 0.8 to 13.6 MB
    (25 Sept 2026); through one writer the fonts are written once. Built-in
    faces only (Nimbus Sans/Roman/Mono, MuPDF's own Noto Naskh for Arabic):
    loading the 1 MB Arial from disk cost 123 ms per box, the built-ins ~3 ms.
    """
    fams = {'sans': 'sans-serif', 'serif': 'serif', 'mono': 'monospace'}
    no_overflow = getattr(getattr(fitz, 'mupdf', None), 'FZ_PLACE_STORY_FLAG_NO_OVERFLOW', 1)

    def put(dev, rect, text, size, b, align, wrap, middle, rtl=False, bg=None):
        if rtl:
            # MuPDF lays every block out LTR unless dir="rtl" ("الجيزة مصر Egypt"
            # came back as [مصر][الجيزة][Egypt]), and under dir="rtl" it mirrors
            # text-align — so ask for the mirror image of what is wanted.
            align = {'left': 'right', 'right': 'left'}.get(align, align)
        css = (f"body{{margin:0}}*{{font-family:{fams[b['fam']]};font-size:{size}pt;color:{b['color']};line-height:1.15;"
               f"font-weight:{'bold' if b['bold'] else 'normal'};margin:0;padding:0}}"
               f"div{{text-align:{align};{'' if wrap else 'white-space:pre;'}{f'background-color:{bg};' if bg else ''}}}")
        # one block per line, each with the field's direction
        attr = ' dir="rtl"' if rtl else ''
        body = ''.join(f'<div{attr}>{html.escape(ln) or "&#160;"}</div>' for ln in text.split('\n'))
        story = fitz.Story(html=body, user_css=css)
        # the insert_htmlbox recipe: grow the rect until the text fits, then
        # scale the drawing back down into the real rect (shrink-to-fit)
        fit = story.fit_scale(fitz.Rect(0, 0, rect.width, rect.height), scale_min=1, scale_max=None, flags=no_overflow)
        scale = 1 / fit.parameter if fit.big_enough and fit.parameter else 1.0
        frame = fitz.Rect(fit.rect)
        filled = fitz.Rect(fit.filled)
        off = max((frame.y1 - filled.y1) * scale, 0) / 2 if middle else 0.0
        story.reset()
        story.place(frame)
        story.draw(dev, fitz.Matrix(scale, 0, 0, scale, rect.x0, rect.y0 + off))

    buf = io.BytesIO()
    writer = fitz.DocumentWriter(buf)
    for b in boxes:
        dev = writer.begin_page(fitz.Rect(0, 0, pw, ph))
        pad = b['pad']
        if b['kind'] == 'list':
            y = pad - 1
            for ln in b['lines']:
                put(dev, fitz.Rect(0, y, b['w'], y + b['lh']), ln['text'] or ' ', b['size'], b,
                    'right' if ln['rtl'] else 'left', False, False, ln['rtl'], LIST_HIGHLIGHT if ln['sel'] else None)
                y += b['lh']
        elif 'cells' in b:
            n = len(b['cells'])
            cw = b['w'] / max(1, n)
            for i, c in enumerate(b['cells']):
                if c.strip():
                    put(dev, fitz.Rect(i * cw, 0, (i + 1) * cw, b['h']), c, b['size'], b, 'center', False, True)
        elif b.get('multiline'):
            put(dev, fitz.Rect(pad, pad, b['w'] - pad, b['h'] - pad), b['text'], b['size'], b, b['align'], True, False, b['rtl'])
        elif b['text']:
            put(dev, fitz.Rect(pad, 0, b['w'] - pad, b['h']), b['text'], b['size'], b, b['align'], False, True, b['rtl'])
        writer.end_page()
    writer.close()
    return fitz.open('pdf', buf.getvalue())


_ROT = {0: '1 0 0 1 0 0', 90: '0 1 -1 0 0 0', 180: '-1 0 0 -1 0 0', 270: '0 -1 1 0 0 0'}


def _frame_ops(d, b):
    """Background and border from /MK, drawn under the text like Acrobat does."""
    w, h = b['w'], b['h']
    ops = []
    if d.get('bg'):
        ops.append(f'{_rgb_ops(d["bg"])} 0 0 {w:.2f} {h:.2f} re f')
    bw = d['bw'] if d.get('bc') else 0
    if bw > 0:
        style = (d.get('bs') or 'Solid').lower()
        dash = ''
        if style.startswith('dash'):
            arr = ' '.join(f'{x:g}' for x in (d.get('dashes') or [3]))
            dash = f'[{arr}] 0 d '
        col = _rgb_ops(d['bc'], stroke=True)
        if style.startswith('under'):
            ops.append(f'{col} {bw:.2f} w {dash}0 {bw / 2:.2f} m {w:.2f} {bw / 2:.2f} l S')
        else:
            ops.append(f'{col} {bw:.2f} w {dash}{bw / 2:.2f} {bw / 2:.2f} {w - bw:.2f} {h - bw:.2f} re S')
            if 'cells' in b and len(b['cells']) > 1:
                cw = w / len(b['cells'])
                for i in range(1, len(b['cells'])):
                    ops.append(f'{i * cw:.2f} 0 m {i * cw:.2f} {h:.2f} l S')
    return ('q ' + ' '.join(ops) + ' Q\n') if ops else ''


def apply_appearances(doc, targets, boxes, ovl):
    """Turn overlay page i into the /AP /N of widget i (see the module note)."""
    n0 = doc.page_count
    doc.insert_pdf(ovl, links=0, annots=0, widgets=0)
    try:
        for i, (d, b) in enumerate(zip(targets, boxes)):
            p = doc[n0 + i]
            content = b'\n'.join((doc.xref_stream(x) or b'') for x in p.get_contents())
            rt, rv = _get(doc, p.xref, 'Resources')
            res = rv if rt in ('xref', 'dict') else '<<>>'
            mb = p.mediabox
            # the box sits at the overlay page's top-left corner
            tx, ty = -mb.x0, -(mb.y1 - b['h'])
            stream = (_frame_ops(d, b).encode() + b'/Tx BMC\nq\n' +
                      f'0 0 {b["w"]:.2f} {b["h"]:.2f} re W n\n1 0 0 1 {tx:.3f} {ty:.3f} cm\n'.encode() +
                      content + b'\nQ\nEMC\n')
            ap = doc.get_new_xref()
            doc.update_object(ap, f'<</Type/XObject/Subtype/Form/BBox[0 0 {b["w"]:.2f} {b["h"]:.2f}]'
                                  f'/Matrix[{_ROT[d["mkR"]]}]/Resources {res}>>')
            doc.update_stream(ap, stream)
            doc.xref_set_key(d['xref'], 'AP', f'<</N {ap} 0 R>>')
    finally:
        doc.delete_pages(n0, doc.page_count - 1)


def _ensure_button_appearances(doc, widgets):
    """A checkbox/radio whose file carries no appearance for its state (forms
    that relied on NeedAppearances) gets MuPDF's standard check/circle."""
    by_page = {}
    for d in widgets:
        if d['type'] in ('checkbox', 'radio') and not d.get('apStates'):
            by_page.setdefault(d['page'], []).append(d)
    fixed = 0
    for pno, ds in by_page.items():
        page = doc[pno]
        for d in ds:
            try:
                w = page.load_widget(d['xref'])
                w.field_value = w.on_state() if d['checked'] else 'Off'
                w.update()
                fixed += 1
            except Exception:  # noqa: BLE001
                pass
    return fixed


def cmd_fill(path, out, values_path, flatten, pipe):
    doc = open_doc(path)
    # utf-8-sig swallows a BOM if the caller left one; a stray BOM must never
    # turn into a 500 for the user.
    with open(values_path, encoding='utf-8-sig') as f:
        values = json.load(f)
    if not isinstance(values, dict):
        die(2, 'values must be an object of name -> value')

    cat = doc.pdf_catalog()
    need_app = _get(doc, cat, 'AcroForm/NeedAppearances')[1] == 'true'
    if need_app:
        # before any page is loaded: MuPDF would otherwise rebuild every
        # appearance itself (with the Arabic problems described above)
        _acro_set(doc, 'NeedAppearances', 'null')
    xfa = _xfa_state(doc)
    if xfa:
        # A hybrid AcroForm+XFA form shows its XFA data in Acrobat, so values
        # written to the AcroForm fields would look unfilled there. Dropping
        # the XFA packet is what every AcroForm filler does.
        _acro_set(doc, 'XFA', 'null')

    _, widgets = read_form(doc, want_labels=False)
    warnings = []
    written, changed = write_values(doc, widgets, values, warnings)

    targets = [d for d in widgets if d['type'] in ('text', 'combobox', 'listbox')
               and (d['name'] in changed or not d['hasAP'] or need_app)]
    boxes = [make_box(d) for d in targets]
    engine = None
    if boxes:
        pw = r2(max(36.0, max(b['w'] for b in boxes)))
        ph = r2(max(36.0, max(b['h'] for b in boxes)))
        ovl = None
        if pipe:
            fd, html_path = tempfile.mkstemp(prefix='form_ovl_', suffix='.html')
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                f.write(_overlay_html(boxes, pw, ph))
            try:
                print(json.dumps({'overlay': {'html': html_path, 'width': pw, 'height': ph, 'boxes': len(boxes)}}), flush=True)
                reply = sys.stdin.readline().strip()
            finally:
                try:
                    os.remove(html_path)
                except OSError:
                    pass
            if reply and os.path.exists(reply):
                try:
                    cand = fitz.open(reply)
                    ok = cand.page_count == len(boxes) and all(
                        abs(p.rect.width - pw) < 2 and abs(p.rect.height - ph) < 2 for p in cand)
                    if ok:
                        ovl, engine = cand, 'chrome'
                    else:
                        warnings.append('overlay did not match the boxes; drawn with MuPDF instead')
                except Exception:  # noqa: BLE001
                    warnings.append('overlay unreadable; drawn with MuPDF instead')
        if ovl is None:
            ovl, engine = _mupdf_overlay(boxes, pw, ph), 'mupdf'
        apply_appearances(doc, targets, boxes, ovl)
    _ensure_button_appearances(doc, widgets)

    if flatten:
        # Bake every appearance into the page. Same call Flatten PDF uses:
        # text stays real text, nothing is rasterised.
        try:
            doc.bake(annots=True, widgets=True)
        except TypeError:
            doc.bake()
    doc.save(out, garbage=3, deflate=True)
    print(json.dumps({'written': written, 'drawn': len(boxes), 'flattened': bool(flatten), 'engine': engine,
                      'xfaRemoved': xfa, 'warnings': warnings[:20]}, ensure_ascii=False))


if __name__ == '__main__':
    a = sys.argv[1:]
    if not a:
        die(2, 'usage')
    cmd = a[0]
    try:
        if cmd == 'fields' and len(a) == 2:
            cmd_fields(a[1])
        elif cmd == 'page' and len(a) in (4, 5):
            cmd_page(a[1], int(a[2]), a[3], float(a[4]) if len(a) == 5 else 1.5)
        elif cmd == 'fill' and len(a) >= 4:
            extra = a[4:]
            cmd_fill(a[1], a[2], a[3], '--flatten' in extra, '--overlay-pipe' in extra)
        else:
            die(2, 'usage')
    except SystemExit:
        raise
    except Exception as e:  # noqa: BLE001
        die(2, f'{type(e).__name__}: {e}')
