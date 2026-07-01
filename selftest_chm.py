# -*- coding: utf-8 -*-
"""
Validate chm_reader against real .chm files.

CHM cannot be synthesized without an LZX *compressor*, so this test uses real
help files shipped with Windows. It searches a few standard locations and the
session scratchpad; if none are found it SKIPS (exit 0) so the suite stays
portable.

Correctness argument without a reference decompressor: the ResetTable stores the
exact uncompressed length, and any LZX bit-drift would either raise (bad Huffman
symbol / block type) or yield the wrong length. Decoding to *exactly* that many
bytes, with every .htm parsing as well-formed-enough HTML and ending in a real
closing tag, is strong evidence the decompressor is correct.
"""
import os
import sys
import glob
import struct
import tempfile

import chm_reader as cr
import html_translator as ht

_SEARCH = [
    r"C:\Windows\Help\mui\0409",
    r"C:\Windows\Help",
    r"C:\Windows\IME\IMEJP\help",
    os.path.join(tempfile.gettempdir(), "claude"),   # scratchpad copies
]


def _find_chms(limit=4):
    found = []
    for root in _SEARCH:
        if not os.path.isdir(root):
            continue
        for path in glob.glob(os.path.join(root, "**", "*.chm"), recursive=True) \
                + glob.glob(os.path.join(root, "**", "*.CHM"), recursive=True):
            if path not in found:
                found.append(path)
            if len(found) >= limit:
                return found
    return found


def _reset_table_uncomp_len(chm):
    rt = chm._read_raw(cr._RESET_TABLE)
    return struct.unpack_from("<Q", rt, 16)[0]


def main():
    chms = _find_chms()
    if not chms:
        print("SKIP: no .chm files found on this system")
        return 0

    fails = []
    def check(cond, msg):
        print(("  ok  " if cond else " FAIL ") + msg)
        if not cond:
            fails.append(msg)

    for path in chms:
        name = os.path.basename(path)
        print("== %s ==" % name)
        try:
            chm = cr.CHMFile(path)
        except Exception as e:
            check(False, "%s: parse failed (%s)" % (name, e))
            continue

        # 1. decompress whole LZX section to the declared length
        try:
            blob = chm._blob()
        except Exception as e:
            check(False, "%s: LZX decode raised (%s)" % (name, e))
            continue
        want = _reset_table_uncomp_len(chm)
        check(len(blob) == want,
              "%s: decoded %d bytes == ResetTable uncompressed length %d"
              % (name, len(blob), want))

        # 2. every content file reads at its declared length
        length_ok = True
        for n in chm.names():
            sec, off, ln = chm.entries[n]
            if len(chm.read(n)) != ln:
                length_ok = False
                break
        check(length_ok, "%s: all files read at their directory length" % name)

        # 3. html files look like real, complete HTML
        htmls = [n for n in chm.names() if n.lower().endswith((".htm", ".html"))]
        good = 0
        for n in htmls[:8]:
            b = chm.read(n).lstrip(b"\xef\xbb\xbf").lstrip()
            low = b[:200].lower()
            tail = b.rstrip()[-16:].lower()
            if (low.startswith(b"<!doctype") or low.startswith(b"<html")) \
                    and tail.endswith(b"</html>"):
                good += 1
        if htmls:
            check(good == min(8, len(htmls)),
                  "%s: %d/%d sampled html files are well-formed"
                  % (name, good, min(8, len(htmls))))

    # 4. end-to-end: decompile one CHM and run the HTML translator on a page
    print("== pipeline: decompile -> html_translator ==")
    chm = cr.CHMFile(chms[0])
    work = tempfile.mkdtemp(prefix="chm_out_")
    written = chm.extract_all(work)
    check(len(written) > 0, "extract_all wrote %d files to a folder" % len(written))
    htmls = [p for p in written if p.lower().endswith((".htm", ".html"))]
    if htmls:
        txt = ht.convert_html_2_text(htmls[0])
        n_lines = sum(1 for _ in open(txt, encoding="utf-8"))
        check(n_lines > 0,
              "html_translator extracted %d translatable lines from a CHM page"
              % n_lines)

    print()
    if fails:
        print("RESULT: %d FAILURE(S)" % len(fails)); return 1
    print("RESULT: ALL CHECKS PASSED  (validated on %d real CHM files)" % len(chms))
    return 0


if __name__ == "__main__":
    # keep prints ASCII-safe regardless of console codepage
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    sys.exit(main())
