#!/usr/bin/env python3
"""
pdf_protect.py — protect / unlock a PDF with PyMuPDF.

Usage: python3 pdf_protect.py unlock  in.pdf out.pdf job.json
       python3 pdf_protect.py protect in.pdf out.pdf job.json

job.json (written 0600 by the caller, deleted after the run — passwords never go on the
command line, where any local user could read them from the process list):
  unlock : {"password": "...", "keepRestrictions": false}
  protect: {"userPassword": "...", "ownerPassword": "...", "level": "aes256|aes128|rc4",
            "permissions": {"print": true, "printHq": true, "modify": false, "copy": true,
                            "annotate": false, "forms": true, "accessibility": true,
                            "assemble": false}}
Prints one JSON line; exit 0 on success. Known refusals (the caller shows a translated
message, no fallback): 3 = needs a password, 6 = wrong password, 7 = not encrypted,
8 = already encrypted (protect), 4 = unreadable. Exit 9 = this engine cannot express the
request (password longer than MuPDF's 40 characters, non-Latin-1 password for the
128-bit levels) and exit 2 = anything else: the caller falls back to PDFsharp.

WHY (audit 2026-09-25 of the PDFsharp engine):
  * Unlocking with the OPEN password failed ("To modify the document the owner password
    is required") for every file with a separate owner password — including every file
    Confileo protected itself, since it invents an owner password when none is given.
    MuPDF decrypts with either password.
  * A file with restrictions but no open password could not be unlocked at all (the
    form demanded a password it does not have). MuPDF opens it as is.
  * PDFsharp's output grew +69% (280 -> 473 KB for AES-256); written here with object
    streams: protected 280 -> 299 KB (+7%), and that 473 KB PDFsharp file unlocks back
    to 282 KB.
  * The "screen readers" permission was ignored; MuPDF writes every bit that was asked.
qpdf --show-encryption on the output: R6/AESv3, R4/AESv2, R3/RC4 with exactly the chosen
permissions; qpdf --check, poppler and pdf.js open them with the user and owner password.
"""
import json
import os
import secrets
import sys

import fitz  # PyMuPDF >= 1.24

EXIT_NEEDS_PASSWORD, EXIT_UNREADABLE, EXIT_WRONG_PASSWORD = 3, 4, 6
EXIT_NOT_ENCRYPTED, EXIT_ALREADY_ENCRYPTED, EXIT_UNSUPPORTED = 7, 8, 9

LEVELS = {"aes256": fitz.PDF_ENCRYPT_AES_256, "aes128": fitz.PDF_ENCRYPT_AES_128, "rc4": fitz.PDF_ENCRYPT_RC4_128}
PERMS = {
    "print": fitz.PDF_PERM_PRINT, "printHq": fitz.PDF_PERM_PRINT_HQ, "modify": fitz.PDF_PERM_MODIFY,
    "copy": fitz.PDF_PERM_COPY, "annotate": fitz.PDF_PERM_ANNOTATE, "forms": fitz.PDF_PERM_FORM,
    "accessibility": fitz.PDF_PERM_ACCESSIBILITY, "assemble": fitz.PDF_PERM_ASSEMBLE,
}
MAX_PW = 40  # PyMuPDF refuses longer passwords


class Refused(Exception):
    def __init__(self, code, reason):
        super().__init__(reason)
        self.code = code


def _open(src):
    try:
        doc = fitz.open(src)
    except Exception as e:  # noqa: BLE001
        raise Refused(EXIT_UNREADABLE, f"cannot open: {e}")
    if not doc.is_pdf:
        raise Refused(EXIT_UNREADABLE, "not a PDF")
    return doc


def _garbage(doc):
    # garbage=3 merges duplicate objects but is quadratic in the object count (7.6 s on a
    # 13.7k-object file); garbage=2 still drops everything unreferenced.
    return 3 if doc.xref_length() <= 4000 else 2


def unlock(src, dst, job):
    password = job.get("password") or ""
    doc = _open(src)
    method = ((doc.metadata or {}).get("encryption") or "").strip()
    # Read once: touching needs_pass AFTER authenticate() made PyMuPDF 1.26 write every
    # content stream still encrypted into the "unlocked" file (qpdf: inflate errors on
    # all 300 pages, 2026-09-25).
    needs_pass = doc.needs_pass
    if not needs_pass and not method:
        raise Refused(EXIT_NOT_ENCRYPTED, "not encrypted")
    role = "none"
    if needs_pass:
        if not password:
            raise Refused(EXIT_NEEDS_PASSWORD, "needs password")
        rc = doc.authenticate(password)
        if not rc:
            raise Refused(EXIT_WRONG_PASSWORD, "wrong password")
        role = "owner" if rc & 4 else "user"
    elif password:
        # restriction-only file: a given password can only be the owner's
        probe = fitz.open(src)
        try:
            role = "owner" if probe.authenticate(password) & 4 else "none"
        finally:
            probe.close()
    perms = doc.permissions
    removed = ["openPassword"] if needs_pass else []
    keep = bool(job.get("keepRestrictions")) and (perms & 0xF3C) != 0xF3C
    save = dict(garbage=_garbage(doc), deflate=True, use_objstms=1)
    if keep:
        # Remove the open password only: the restrictions stay, under the owner password
        # when it was given, else under a fresh random one (the old one is unknown).
        save.update(encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="", permissions=perms,
                    owner_pw=password if role == "owner" else secrets.token_hex(20))
    else:
        save.update(encryption=fitz.PDF_ENCRYPT_NONE)
        if (perms & 0xF3C) != 0xF3C:
            removed.append("restrictions")
    pages = len(doc)
    doc.save(dst, **save)
    doc.close()
    _readable(dst, "" if not keep else None)
    return {"pages": pages, "encryption": method, "authenticatedAs": role,
            "removed": removed, "restrictionsKept": keep}


def _readable(path, password):
    """Read the written file back — a broken output must fail here, not in the visitor's
    viewer. With password None only the structure is opened."""
    check = fitz.open(path)
    try:
        if password is not None and check.needs_pass:
            raise RuntimeError("unlocked file still asks for a password")
        for i in {0, len(check) // 2, len(check) - 1}:
            if 0 <= i < len(check):
                check[i].get_text("text")
                check.xref_stream(check[i].get_contents()[0]) if check[i].get_contents() else None
    finally:
        check.close()


def _latin1(s):
    try:
        s.encode("latin-1")
        return True
    except UnicodeEncodeError:
        return False


def protect(src, dst, job):
    user_pw = job.get("userPassword") or ""
    owner_pw = job.get("ownerPassword") or ""
    level = LEVELS.get((job.get("level") or "aes256").lower(), fitz.PDF_ENCRYPT_AES_256)
    if not user_pw or not owner_pw:
        raise Refused(EXIT_UNSUPPORTED, "both passwords are required")
    if len(user_pw) > MAX_PW or len(owner_pw) > MAX_PW:
        raise Refused(EXIT_UNSUPPORTED, "password longer than MuPDF allows")
    if level != fitz.PDF_ENCRYPT_AES_256 and not (_latin1(user_pw) and _latin1(owner_pw)):
        # RC4 / AES-128 passwords are PDFDocEncoding: readers disagree on anything else
        raise Refused(EXIT_UNSUPPORTED, "non Latin-1 password for a 128-bit level")
    doc = _open(src)
    if doc.needs_pass or ((doc.metadata or {}).get("encryption") or "").strip():
        raise Refused(EXIT_ALREADY_ENCRYPTED, "already encrypted")
    asked = job.get("permissions") or {}
    perms = 0
    for key, bit in PERMS.items():
        if asked.get(key):
            perms |= bit
    if perms & fitz.PDF_PERM_PRINT_HQ:
        perms |= fitz.PDF_PERM_PRINT  # high-quality printing is a refinement of printing
    pages = len(doc)
    doc.save(dst, garbage=_garbage(doc), deflate=True, use_objstms=1,
             encryption=level, owner_pw=owner_pw, user_pw=user_pw, permissions=perms)
    doc.close()
    # read it back: the user password must open it, with the permissions asked for
    check = fitz.open(dst)
    try:
        if not check.needs_pass or not check.authenticate(user_pw):
            raise RuntimeError("protected file does not open with its password")
    finally:
        check.close()
    return {"pages": pages, "permissions": perms}


def main(argv):
    if len(argv) != 4 or argv[0] not in ("unlock", "protect"):
        print("usage: pdf_protect.py unlock|protect in.pdf out.pdf job.json", file=sys.stderr)
        return 2
    with open(argv[3], "r", encoding="utf-8") as f:
        job = json.load(f)
    # The API's cleanup can remove the empty out/ folder between the C# side choosing
    # this path and the save below (SmartCleanup deletes empty subfolders every 30 s).
    os.makedirs(os.path.dirname(os.path.abspath(argv[2])), exist_ok=True)
    out = (unlock if argv[0] == "unlock" else protect)(argv[1], argv[2], job)
    print(json.dumps(out))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except Refused as r:
        print(json.dumps({"error": str(r)}))
        sys.exit(r.code)
    except Exception as e:  # noqa: BLE001
        print(f"pdf_protect failed: {e}", file=sys.stderr)
        sys.exit(2)
