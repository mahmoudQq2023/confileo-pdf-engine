#!/usr/bin/env python3
"""
pdf_split.py — write several page subsets of one PDF and keep each of them whole.

Usage: python3 pdf_split.py in.pdf manifest.json
       python3 pdf_split.py --info in.pdf      -> {"pages": n}; exit 3 encrypted, 4 damaged
manifest.json: {"outputs": [{"path": "/tmp/x_part1.pdf", "pages": [1, 2, 3]}, ...]}   # 1-based
Prints one JSON line ({"written": n}); exit 0. Exit 3 = encrypted; exit 4 = the file is
damaged beyond repair / has no pages (stderr says so); exit 2 = failure (the caller falls
back to its PdfSharp path).

WHY (measured 2026-09-14 on the live tool with a 6-page probe): PdfSharp's AddPage copies
page objects one by one into an empty document, so every split output lost the outline
(7 bookmarks -> 0), the metadata (title/author -> empty) and the page labels (i, ii, 1, 2 ->
none). Form fields, links, rotation and crop boxes survived. MuPDF's select() keeps the
outline entries that point at surviving pages and re-targets them, keeps the metadata, and
the page labels are rebuilt here from the source's label rules so "page 12" of a chapter
still reads "12" in the extracted file. Saving with garbage collection also drops the
fonts and images that only the removed pages used, which PdfSharp already did implicitly.

2026-10-02 — defects found on a 330-page WeasyPrint/Chrome-style test book and a form:
  * SIZE BLOAT. Chrome, WeasyPrint, wkhtmltopdf and many reporting tools give every page
    ONE shared /Resources dictionary that lists every font and image of the document.
    Garbage collection cannot drop an image that the shared dictionary still names, so a
    one-page extract of a 649 KB book weighed 407 KB, and "each page" turned that book into
    330 x ~407 KB = 134 MB. _prune_resources() gives each kept page a resource dictionary
    holding only the fonts and images its content stream names: the same extract is now
    12.9 KB and "each page" 27.8 MB (the pages that really show a 195 KB photo keep it).
    Every page renders pixel-identical to the source and extracts the same text.
  * BOOKMARKS LOST BY CHAPTER. select() drops an outline item whose target page is gone
    TOGETHER WITH ITS CHILDREN, so a part covering pages 2-5 of a chapter that starts on
    page 1 lost the chapter's section bookmarks too (3 kept of 6 that point inside the
    part). The outline is rebuilt when that happens: every item pointing into the part is
    kept, an ancestor that points outside it is kept as a header opening its first child.
  * FORM FIELDS STOPPED BEING FIELDS. select() removes /AcroForm: 8 fields -> 0 (the
    widgets stayed as plain annotations). _restore_acroform() puts it back.
  * DAMAGED FILES. A truncated upload that MuPDF repairs to zero pages now exits 4 with a
    reason instead of failing somewhere inside; --info lets the caller ask MuPDF for the
    page count of a file PdfSharp cannot parse.
  * SPEED. The source is read once into memory, outputs are spread over up to 4 forked
    processes when there are many, and saved with object streams + duplicate-stream
    merging (garbage=4, use_objstms) like pdf_merge.py: 330 single-page parts ~3-4 s ->
    ~1.0-2.3 s depending on how busy the cores are; a 2-part split of the same book costs
    ~80 ms more (trimming 330 pages' resources + the denser save).
"""
import json
import os
import re
import sys

import fitz  # PyMuPDF


def _rules(doc):
    try:
        rules = doc.get_page_labels() or []
    except Exception:
        rules = []
    rules = [r for r in rules if isinstance(r, dict) and "startpage" in r]
    rules.sort(key=lambda r: int(r.get("startpage", 0)))
    return rules


def _label_of(rules, i):
    """(rule, number) for original page index i, or None when the page has no label rule."""
    hit = None
    for r in rules:
        if int(r.get("startpage", 0)) <= i:
            hit = r
        else:
            break
    if hit is None:
        return None
    return hit, int(hit.get("firstpagenum", 1)) + (i - int(hit.get("startpage", 0)))


def _rebuild_labels(rules, kept):
    """Label rules for the new document whose page j came from original index kept[j]."""
    out = []
    prev = None
    for j, i in enumerate(kept):
        cur = _label_of(rules, i)
        if cur is None:
            prev = None
            continue
        rule, num = cur
        if prev is not None and prev[0] is rule and prev[1] + 1 == num:
            prev = (rule, num)
            continue
        out.append({
            "startpage": j,
            "prefix": rule.get("prefix", "") or "",
            "style": rule.get("style", "") or "",
            "firstpagenum": num,
        })
        prev = (rule, num)
    return out


def _drop_struct_tree(doc):
    """A tagged PDF (Chrome, Word) keeps every page alive through /StructTreeRoot, so the
    removed pages — and their fonts and images — would survive garbage collection and a
    one-page extract would weigh as much as the whole file (110 KB vs 46 KB measured). The
    tag tree of a partial document is broken anyway, so it goes; PdfSharp never kept it."""
    try:
        cat = doc.pdf_catalog()
        if doc.xref_get_key(cat, "StructTreeRoot")[0] != "null":
            doc.xref_set_key(cat, "StructTreeRoot", "null")
            doc.xref_set_key(cat, "MarkInfo", "null")
    except Exception as e:
        print(f"struct tree: {e}", file=sys.stderr)


# ── outline ───────────────────────────────────────────────────────────────────────────

def _subset_toc(full_toc, kept):
    """Outline for a part made of original page indices `kept` (0-based, in output order).

    full_toc is get_toc(simple=False) of the source. An item pointing at a kept page is kept
    and re-targeted (destination point preserved); an item pointing elsewhere is kept only
    when one of its descendants is, and then opens that descendant's page."""
    newpos = {}
    for j, i in enumerate(kept):
        newpos.setdefault(i, j)

    n = len(full_toc)
    keep = [False] * n
    target = [None] * n          # new 0-based page of each kept item
    # children follow their parent with a higher level; walk backwards so an ancestor
    # sees whether anything below it survived and which page its first survivor opens
    for k in range(n - 1, -1, -1):
        lvl, _title, page = full_toc[k][0], full_toc[k][1], full_toc[k][2]
        own = newpos.get(page - 1) if page and page > 0 else None
        # first kept descendant: scan the direct run of deeper items after k
        below = None
        m = k + 1
        while m < n and full_toc[m][0] > lvl:
            if keep[m]:
                below = target[m]
                break
            m += 1
        if own is not None:
            keep[k], target[k] = True, own
        elif below is not None:
            keep[k], target[k] = True, below

    out = []
    for k in range(n):
        if not keep[k]:
            continue
        lvl, title = full_toc[k][0], full_toc[k][1]
        page = target[k]
        entry = [lvl, title, page + 1]
        src_page = full_toc[k][2] - 1 if full_toc[k][2] else -1
        dest = full_toc[k][3] if len(full_toc[k]) > 3 and isinstance(full_toc[k][3], dict) else None
        if dest is not None and newpos.get(src_page) == page and dest.get("kind") == fitz.LINK_GOTO:
            d = {"kind": fitz.LINK_GOTO, "page": page, "to": dest.get("to", fitz.Point(0, 0)),
                 "zoom": dest.get("zoom", 0)}
            if "collapse" in dest:
                d["collapse"] = dest["collapse"]
            entry.append(d)
        out.append(entry)
    return _fix_levels(out)


def _fix_levels(toc):
    """set_toc() refuses a level that jumps by more than one; clamp instead of failing."""
    fixed = []
    prev = 0
    for e in toc:
        lvl = max(1, min(int(e[0]), prev + 1))
        fixed.append([lvl] + list(e[1:]))
        prev = lvl
    return fixed


def _set_toc(doc, toc):
    try:
        doc.set_toc(toc)
    except Exception:
        try:
            doc.set_toc([e[:3] for e in toc])
        except Exception as e:  # bookmarks are a nicety; never fail the split over them
            print(f"toc: {e}", file=sys.stderr)


# ── form fields ───────────────────────────────────────────────────────────────────────

def _acroform_before(doc):
    try:
        return doc.xref_get_key(doc.pdf_catalog(), "AcroForm")
    except Exception:
        return ("null", "null")


def _restore_acroform(doc, acro):
    """select() keeps the widgets on the kept pages but removes the catalog's /AcroForm, so
    a split part of a fillable form showed boxes that were no longer form fields (pdf.js and
    Acrobat treat them as plain annotations; measured 2026-10-02: 8 fields -> 0). The form
    dictionary comes back with /Fields listing only the fields that still have a widget, and
    a field's /Kids lose the widgets of removed pages (which would otherwise keep those
    pages and their images alive through garbage collection)."""
    if acro[0] not in ("xref", "dict"):
        return
    try:
        keep = set()
        roots = []
        for page in doc:
            for w in page.widgets():
                x = w.xref
                chain = [x]
                for _ in range(32):
                    par = doc.xref_get_key(x, "Parent")
                    if par[0] != "xref":
                        break
                    x = int(par[1].split()[0])
                    chain.append(x)
                keep.update(chain)
                if x not in roots:
                    roots.append(x)
        if not roots:
            return
        for x in keep:
            kids = doc.xref_get_key(x, "Kids")
            if kids[0] != "array":
                continue
            refs = re.findall(r"(\d+) 0 R", kids[1])
            kept = [r for r in refs if int(r) in keep]
            if len(kept) != len(refs):
                doc.xref_set_key(x, "Kids", "[" + " ".join(f"{r} 0 R" for r in kept) + "]")
        if acro[0] == "xref":
            ax = int(acro[1].split()[0])
        else:
            ax = doc.get_new_xref()
            doc.update_object(ax, acro[1])
        doc.xref_set_key(ax, "Fields", "[" + " ".join(f"{r} 0 R" for r in roots) + "]")
        if doc.xref_get_key(ax, "XFA")[0] != "null":
            doc.xref_set_key(ax, "XFA", "null")  # an XFA packet describes the whole form
        doc.xref_set_key(doc.pdf_catalog(), "AcroForm", f"{ax} 0 R")
    except Exception as e:  # never fail a split over the form dictionary
        print(f"acroform: {e}", file=sys.stderr)


# ── resources ─────────────────────────────────────────────────────────────────────────

_NAME_RE = re.compile(rb"/([^\s/\[\]()<>{}%]+)")
_SAFE_KEY = re.compile(r"^[A-Za-z0-9_.+\-]+$")
_HEX_ESC = re.compile(rb"#([0-9A-Fa-f]{2})")


def _names_in(data):
    names = set()
    for m in _NAME_RE.finditer(data or b""):
        raw = m.group(1)
        if b"#" in raw:
            raw = _HEX_ESC.sub(lambda h: bytes([int(h.group(1), 16)]), raw)
        names.add(raw.decode("latin-1"))
    return names


def _dict_items(doc, value):
    """(key, value-string) pairs of a dictionary given as xref_get_key() result, or None."""
    kind, text = value
    if kind == "xref":
        xref = int(text.split()[0])
    elif kind == "dict":
        xref = doc.get_new_xref()          # scratch object, unreferenced -> garbage on save
        doc.update_object(xref, text)
    else:
        return None
    items = []
    for k in doc.xref_get_keys(xref):
        vk, vt = doc.xref_get_key(xref, k)
        if vk not in ("xref", "dict", "array", "name", "int", "float", "bool", "null"):
            return None
        items.append((k, vt))
    return items


def _resources_of(doc, page_xref):
    """The page's /Resources value, following /Parent for inherited resources."""
    xref = page_xref
    for _ in range(32):
        val = doc.xref_get_key(xref, "Resources")
        if val[0] != "null":
            return val
        parent = doc.xref_get_key(xref, "Parent")
        if parent[0] != "xref":
            return val
        xref = int(parent[1].split()[0])
    return ("null", "null")


def _parse_resources(doc, res):
    """(items of the resource dict, {"Font"/"XObject": items of that sub-dictionary}) or None
    when something is unusual enough to leave the page alone."""
    items = _dict_items(doc, res)
    if items is None or not all(_SAFE_KEY.match(k) for k, _ in items):
        return None
    subs = {}
    for key, _val in items:
        if key not in ("Font", "XObject"):
            continue
        if res[0] == "xref":
            sub_val = doc.xref_get_key(int(res[1].split()[0]), key)
        else:
            sub_val = _inline_sub(doc, res[1], key)
        sub = _dict_items(doc, sub_val)
        if sub is None or not all(_SAFE_KEY.match(k) for k, _ in sub):
            return None
        subs[key] = sub
    return items, subs


def _prune_resources(doc):
    """Give every page a resource dictionary with only the fonts and XObjects its content
    names. Conservative: a page is left alone when anything is unusual (a Form XObject that
    inherits the page's resources, odd key names, unreadable dictionaries). Pages that end
    up needing the same subset share one new dictionary."""
    parsed = {}     # shared resource dict -> _parse_resources()
    form_ok = {}    # XObject reference -> _form_ok()
    made = {}       # (shared dict, kept names) -> xref of the trimmed dict
    pruned = 0
    for page in doc:
        try:
            res = _resources_of(doc, page.xref)
            if res[0] not in ("xref", "dict"):
                continue
            if res[0] == "xref":
                if res not in parsed:
                    parsed[res] = _parse_resources(doc, res)
                info = parsed[res]
            else:
                info = _parse_resources(doc, res)
            if info is None:
                continue
            items, subs = info
            used = _names_in(page.read_contents())
            kept_by_key = {}
            changed = False
            ok = True
            for key, sub in subs.items():
                kept = []
                for k, v in sub:
                    if k not in used:
                        continue
                    if key == "XObject":
                        if v not in form_ok:
                            form_ok[v] = _form_ok(doc, v)
                        if not form_ok[v]:
                            ok = False
                            break
                    kept.append((k, v))
                if not ok:
                    break
                changed = changed or len(kept) != len(sub)
                kept_by_key[key] = tuple(kept)
            if not ok or not changed:
                continue
            sig = (res if res[0] == "xref" else None, tuple(sorted(kept_by_key.items())))
            xref = made.get(sig) if sig[0] is not None else None
            if xref is None:
                parts = []
                for key, val in items:
                    if key in kept_by_key:
                        val = "<<" + "".join(f"/{k} {v}" for k, v in kept_by_key[key]) + ">>"
                    parts.append(f"/{key} {val}")
                xref = doc.get_new_xref()
                doc.update_object(xref, "<<" + "".join(parts) + ">>")
                if sig[0] is not None:
                    made[sig] = xref
            doc.xref_set_key(page.xref, "Resources", f"{xref} 0 R")
            pruned += 1
        except Exception as e:  # never fail a split over an optimisation
            print(f"prune p{page.number}: {e}", file=sys.stderr)
    return pruned


def _inline_sub(doc, dict_text, key):
    tmp = doc.get_new_xref()
    doc.update_object(tmp, dict_text)
    return doc.xref_get_key(tmp, key)


def _form_ok(doc, value):
    """A Form XObject without its own /Resources draws with the PAGE's resources, so pruning
    by the page's content alone could remove what it needs: refuse those pages."""
    if not value.endswith(" R"):
        return True
    try:
        xref = int(value.split()[0])
        if doc.xref_get_key(xref, "Subtype")[1] != "/Form":
            return True
        return doc.xref_get_key(xref, "Resources")[0] != "null"
    except Exception:
        return False


# ── main ──────────────────────────────────────────────────────────────────────────────

_CTX = {}


def _workers(jobs):
    """Each output re-opens the source (~6 ms for a 330-page book) and saves (~3 ms): a
    "one file per page" split of a long PDF is hundreds of independent jobs, so they are
    spread over up to 4 forked processes once there are enough of them to pay for it.
    Measured 2026-10-02 on 4 cores, 330 one-page parts: ~4.0 s serial -> ~1.0-1.5 s with
    the pool on a quiet machine; with the cores already busy (load 9) the gain shrinks to
    ~10-20 %, never a loss beyond noise. A failed pool falls back to serial."""
    if jobs < 24 or not hasattr(os, "fork"):
        return 1
    try:
        cpus = len(os.sched_getaffinity(0))
    except Exception:
        cpus = os.cpu_count() or 1
    return max(1, min(4, cpus, jobs // 12))


def _write_one(job):
    path, pages = job
    data, rules, full_toc = _CTX["data"], _CTX["rules"], _CTX["toc"]
    doc = fitz.open("pdf", data)
    acro = _acroform_before(doc)
    doc.select(pages)
    _restore_acroform(doc, acro)
    _drop_struct_tree(doc)
    if full_toc:
        want = _subset_toc(full_toc, pages)
        try:
            have = doc.get_toc(simple=True)
        except Exception:
            have = None
        # rewriting the outline costs ~50 ms per 180 bookmarks: only when select() lost some
        if have is None or [e[:3] for e in have] != [e[:3] for e in want]:
            _set_toc(doc, want)
    if rules:
        try:
            doc.set_page_labels(_rebuild_labels(rules, pages))
        except Exception as e:  # labels are a nicety; never fail the split over them
            print(f"labels: {e}", file=sys.stderr)
    _prune_resources(doc)
    doc.save(path, garbage=4, deflate=True, use_objstms=1)
    doc.close()
    return 1

def info(src):
    """Page count of a file PdfSharp could not open (a wrong xref offset or a missing trailer
    is common in files saved by broken editors; MuPDF repairs those on open)."""
    try:
        doc = fitz.open(src, filetype="pdf")
    except Exception as e:
        print(f"damaged: {e}", file=sys.stderr)
        return 4
    if doc.needs_pass:
        print("encrypted", file=sys.stderr)
        return 3
    if len(doc) == 0:
        print("damaged: no readable pages", file=sys.stderr)
        return 4
    print(json.dumps({"pages": len(doc)}))
    return 0


def main(src, manifest_path):
    with open(manifest_path, "r", encoding="utf-8") as f:
        outputs = json.load(f).get("outputs", [])
    with open(src, "rb") as f:
        data = f.read()
    try:
        probe = fitz.open("pdf", data)
    except Exception as e:
        print(f"damaged: {e}", file=sys.stderr)
        return 4
    if probe.needs_pass:
        print("encrypted", file=sys.stderr)
        return 3
    n = len(probe)
    if n == 0:
        print("damaged: no readable pages", file=sys.stderr)
        return 4
    rules = _rules(probe)
    try:
        full_toc = probe.get_toc(simple=False)
    except Exception:
        full_toc = []
    probe.close()

    jobs = []
    for spec in outputs:
        pages = [int(p) - 1 for p in spec.get("pages", [])]
        pages = [p for p in pages if 0 <= p < n]
        if not pages:
            print(f"output {spec.get('path')}: no valid pages", file=sys.stderr)
            return 2
        jobs.append((spec["path"], pages))

    _CTX.update(data=data, rules=rules, toc=full_toc)
    written = 0
    workers = _workers(len(jobs))
    if workers > 1:
        try:
            import multiprocessing
            with multiprocessing.get_context("fork").Pool(workers) as pool:
                written = sum(pool.imap_unordered(_write_one, jobs, chunksize=8))
        except Exception as e:  # fork unavailable / a worker died: do it in this process
            print(f"pool: {e}", file=sys.stderr)
            written = 0
    if written != len(jobs):
        written = sum(_write_one(j) for j in jobs)
    print(json.dumps({"written": written, "pages": n}))
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: pdf_split.py in.pdf manifest.json | --info in.pdf", file=sys.stderr)
        sys.exit(2)
    try:
        if sys.argv[1] == "--info":
            sys.exit(info(sys.argv[2]))
        sys.exit(main(sys.argv[1], sys.argv[2]))
    except Exception as e:
        print(f"pdf_split: {e}", file=sys.stderr)
        sys.exit(2)
