#!/usr/bin/env python3
"""
pdf_merge.py — merge PDFs and KEEP what people paid for: bookmarks, links, form fields,
annotations, page labels, metadata.

Usage: python3 pdf_merge.py out.pdf manifest.json
manifest: {"files": ["/tmp/a.pdf", "/tmp/b.pdf", ...], "names": ["Contract.pdf", "Annex.pdf"],
           "title": "optional title"}
Prints one JSON line on success ({"pages": n, "toc": n, "widgets": n, "counts": [pages per
file]}); exit 0. Exit 3 = a file is encrypted (stderr: "encrypted: <name>"); exit 4 = a file
is damaged beyond repair or has no pages (stderr: "damaged: <name>"); exit 2 = failure →
caller falls back to PdfSharp.

WHY (measured 2026-09-08 on the live tool): PdfSharp's AddPage copies page content and page-
level annotations but drops the document outline (bookmarks) of every input, the AcroForm
dictionary (so filled form fields stop being fields), and all metadata. MuPDF's insert_pdf
carries pages with their annotations, widgets and links; the outlines are re-based here by
page offset and nested under one bookmark per source file, the way Acrobat "Combine files"
does it, so a merged 300-page bundle is still navigable.

2026-10-02, measured on a test set of WeasyPrint/Chrome-style books (330 + 330 + 18 + 12
pages, 2.13 MB in total), a fillable form and damaged files:
  * INTERNAL LINKS WERE DROPPED. Chrome, Word, LibreOffice and WeasyPrint point table-of-
    contents links at NAMED destinations; insert_pdf only re-creates links with an explicit
    page target, so every "go to chapter" link vanished (66 of 66 lost). Named links are now
    re-created as explicit page links into the merged file (_named_links).
  * THE OUTPUT WAS BIGGER THAN THE INPUTS: 2.25 MB for 2.13 MB in. Two files from the same
    source repeat the same fonts and images; garbage=3 merges duplicate objects but not
    duplicate streams. garbage=4 (stream comparison) + object streams: 0.67 MB, pixel-
    identical on every sampled page, +0.15 s.
  * A DAMAGED FILE WAS MERGED AS NOTHING. A truncated upload that MuPDF repairs to zero pages
    was "merged" silently (or failed with "cannot find page tree" and fell back to PdfSharp,
    whose error did not say which file). It is now refused by name (exit 4).
  * Bookmarks with no resolvable target made set_toc() raise on the level jump of their
    children, and the fallback raised the same way → the whole merge fell back to PdfSharp
    and lost every bookmark. Levels are clamped now.
  * Page labels (i, ii, 1, 2 …) of the inputs are re-based instead of dropped.
"""
import json
import os
import sys

import fitz  # PyMuPDF


def _fix_levels(toc):
    """set_toc() refuses a level that jumps by more than one; clamp instead of failing."""
    fixed = []
    prev = 0
    for e in toc:
        lvl = max(1, min(int(e[0]), prev + 1))
        fixed.append([lvl] + list(e[1:]))
        prev = lvl
    return fixed


def _named_links(out, src, offset):
    """Re-create the source's links to named destinations as explicit page links.

    get_links() resolves a named destination to a page and a point in PDF coordinates
    (bottom-left origin); insert_link() wants the point in MuPDF page coordinates."""
    made = 0
    for pno in range(len(src)):
        try:
            links = [l for l in src[pno].get_links()
                     if l.get("kind") == fitz.LINK_NAMED and 0 <= l.get("page", -1) < len(src)]
        except Exception:
            continue
        if not links:
            continue
        page = out[offset + pno]
        for l in links:
            target = out[offset + l["page"]]
            to = l.get("to")
            if isinstance(to, fitz.Point):
                to = to * target.transformation_matrix
            else:
                to = fitz.Point(0, 0)
            try:
                page.insert_link({"kind": fitz.LINK_GOTO, "from": l["from"], "page": offset + l["page"],
                                  "to": to, "zoom": l.get("zoom", 0) or 0})
                made += 1
            except Exception as e:  # a link is never worth failing the merge for
                print(f"link p{offset + pno}: {e}", file=sys.stderr)
    return made


def _labels(src):
    try:
        rules = [r for r in (src.get_page_labels() or []) if isinstance(r, dict) and "startpage" in r]
    except Exception:
        rules = []
    return sorted(rules, key=lambda r: int(r.get("startpage", 0)))


def main(out_path, manifest_path):
    with open(manifest_path, "r", encoding="utf-8") as f:
        m = json.load(f)
    files = m.get("files") or []
    names = m.get("names") or [os.path.basename(p) for p in files]
    if not files:
        print("no files", file=sys.stderr)
        return 2

    out = fitz.open()
    toc = []
    widgets = 0
    links = 0
    counts = []
    labels = []          # (offset, page count, rules of that file)
    first_meta = None
    for path, name in zip(files, names):
        try:
            src = fitz.open(path, filetype="pdf")
        except Exception:
            print(f"damaged: {name}", file=sys.stderr)
            return 4
        if src.needs_pass:
            print(f"encrypted: {name}", file=sys.stderr)
            return 3
        if len(src) == 0:
            print(f"damaged: {name}", file=sys.stderr)
            return 4
        if first_meta is None:
            first_meta = dict(src.metadata or {})
        offset = len(out)
        out.insert_pdf(src, annots=True, widgets=True, links=True)
        links += _named_links(out, src, offset)
        if src.is_form_pdf:
            widgets += sum(1 for p in src for _ in p.widgets())
        counts.append(len(src))
        labels.append((offset, len(src), _labels(src)))
        label = os.path.splitext(name)[0] or name
        toc.append([1, label, offset + 1])
        try:
            src_toc = src.get_toc(simple=False)
        except Exception:
            src_toc = []
        for lvl, title, page, *rest in src_toc:
            if page < 1:
                continue
            entry = [lvl + 1, title, page + offset]
            if rest and isinstance(rest[0], dict):
                d = dict(rest[0])
                # keep the destination geometry when present; kind/page are re-based
                if "page" in d:
                    d["page"] = page + offset - 1
                entry.append(d)
            toc.append(entry)
        src.close()

    # A single-file "merge" (or files without any outline) still gets the file bookmarks —
    # they are what makes the bundle navigable — unless the caller asked for none.
    if m.get("bookmarks", True) and len(files) > 1:
        toc = _fix_levels(toc)
        try:
            out.set_toc(toc)
        except Exception:  # noqa: BLE001 - a malformed destination must not fail the merge
            try:
                out.set_toc([[e[0], e[1], e[2]] for e in toc])
            except Exception as e:  # noqa: BLE001
                print(f"toc: {e}", file=sys.stderr)
    elif files and len(files) == 1:
        try:
            out.set_toc(_fix_levels([[e[0] - 1] + e[1:3] for e in toc[1:]]))
        except Exception:  # noqa: BLE001
            pass

    # Page labels: only when an input has them. A file without labels then gets plain
    # numbers from its first page in the merged file, so it does not inherit the label
    # sequence of the file before it.
    if any(rules for _, _, rules in labels):
        merged = []
        for offset, count, rules in labels:
            if not rules or int(rules[0].get("startpage", 0)) > 0:
                merged.append({"startpage": offset, "prefix": "", "style": "D", "firstpagenum": offset + 1})
            for r in rules:
                sp = int(r.get("startpage", 0))
                if 0 <= sp < count:
                    merged.append({"startpage": offset + sp, "prefix": r.get("prefix", "") or "",
                                   "style": r.get("style", "") or "", "firstpagenum": int(r.get("firstpagenum", 1) or 1)})
        try:
            out.set_page_labels(merged)
        except Exception as e:  # noqa: BLE001 - labels are a nicety
            print(f"labels: {e}", file=sys.stderr)

    meta = {k: (first_meta or {}).get(k) or "" for k in ("author", "subject", "keywords")}
    meta["title"] = m.get("title") or (first_meta or {}).get("title") or "Merged document"
    meta["producer"] = "Confileo"
    meta["creator"] = "Confileo Merge PDF"
    out.set_metadata(meta)
    out.save(out_path, garbage=4, deflate=True, use_objstms=1)
    res = {"pages": len(out), "toc": len(out.get_toc()), "widgets": widgets, "links": links, "counts": counts}
    out.close()
    print(json.dumps(res))
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("usage: pdf_merge.py out.pdf manifest.json", file=sys.stderr)
        sys.exit(2)
    try:
        sys.exit(main(sys.argv[1], sys.argv[2]))
    except Exception as e:  # noqa: BLE001
        print(f"pdf_merge failed: {e}", file=sys.stderr)
        sys.exit(2)
