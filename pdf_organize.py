#!/usr/bin/env python3
"""
pdf_organize.py — reorder / delete / extract / rotate pages and keep the document whole.

Usage: python3 pdf_organize.py in.pdf out.pdf ops.json
ops.json: {"ops": [{"op": "reorder", "order": [3,1,2,...]},         # 1-based, full permutation
                   {"op": "delete",  "pages": [2,4]},
                   {"op": "extract", "pages": [5,6,7]},
                   {"op": "rotate",  "pages": [1,2], "angle": 90}]}
Prints one JSON line ({"pages": n, "toc": n}); exit 0. Exit 3 = encrypted; exit 2 = failure
(caller falls back to its PdfSharpCore path).

WHY (measured 2026-09-08 on the live tool): PdfSharpCore's AddPage rebuilds the file page by
page, so every reorder/delete/rotate dropped the outline (bookmarks: 5 → 0 on the test deck)
and the metadata. MuPDF's select() keeps the outline entries of the pages that survive and
re-points them to the new positions, keeps links/annotations/form fields, and the metadata.
"""
import json
import sys

import fitz  # PyMuPDF


def main(src, dst, ops_path):
    with open(ops_path, "r", encoding="utf-8") as f:
        ops = json.load(f).get("ops", [])
    doc = fitz.open(src)
    if doc.needs_pass:
        print("encrypted", file=sys.stderr)
        return 3
    for op in ops:
        kind = (op.get("op") or "").lower()
        n = len(doc)
        if kind == "reorder":
            order = [int(p) - 1 for p in op.get("order", [])]
            if sorted(order) != list(range(n)):
                print(f"reorder: bad permutation for {n} pages", file=sys.stderr)
                return 2
            doc.select(order)
        elif kind == "delete":
            drop = {int(p) - 1 for p in op.get("pages", [])}
            keep = [i for i in range(n) if i not in drop]
            if not keep:
                print("delete: no pages left", file=sys.stderr)
                return 2
            doc.select(keep)
        elif kind == "extract":
            keep = [int(p) - 1 for p in op.get("pages", []) if 1 <= int(p) <= n]
            if not keep:
                print("extract: no valid pages", file=sys.stderr)
                return 2
            doc.select(keep)
        elif kind == "rotate":
            angle = int(op.get("angle", 90)) % 360
            for p in op.get("pages", []) or range(1, n + 1):
                i = int(p) - 1
                if 0 <= i < n:
                    page = doc[i]
                    page.set_rotation((page.rotation + angle) % 360)
        else:
            print(f"unknown op {kind}", file=sys.stderr)
            return 2
    meta = doc.metadata or {}
    meta["producer"] = "Confileo"
    doc.set_metadata({k: v for k, v in meta.items() if v is not None})
    # garbage=3 also merges duplicate objects but is quadratic in the object count: deleting
    # one page of a 9.1k-object file took 2.76 s, nearly all of it this save (10 Oct 2026).
    # Same rule as pdf_protect.py and pdf_redact.py — garbage=2 (still drops everything
    # unreferenced, the pages taken out included) past 4,000 objects — with object streams.
    doc.save(dst, garbage=3 if doc.xref_length() <= 4000 else 2, deflate=True, use_objstms=1)
    res = {"pages": len(doc), "toc": len(doc.get_toc())}
    doc.close()
    print(json.dumps(res))
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("usage: pdf_organize.py in.pdf out.pdf ops.json", file=sys.stderr)
        sys.exit(2)
    try:
        sys.exit(main(sys.argv[1], sys.argv[2], sys.argv[3]))
    except Exception as e:  # noqa: BLE001
        print(f"pdf_organize failed: {e}", file=sys.stderr)
        sys.exit(2)
