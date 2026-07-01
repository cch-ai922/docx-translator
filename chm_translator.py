# -*- coding: utf-8 -*-
"""
chm_translator.py  --  CHM translation via decompile -> translate HTML -> recompile.

A pure-Python `.chm` *writer* is impractical (it needs an LZX compressor + ITSS
writer that the stdlib doesn't provide), so this module covers the two halves it
can do losslessly and hands the final recompile to Microsoft's HTML Help
Workshop (`hhc.exe`):

    step 1  chm_2_texts(chm)      decompile -> <stem>_chm/ folder of pages,
                                  one <page>.txt beside each .htm  (+ project.hhp)
    step 2  (external engine translates every <page>.txt -> <page>_tran.txt)
    step 3  texts_2_html(folder)  rebuild every <page>_tran.html in place
            then:  hhc.exe <stem>_chm/project.hhp   ->  a translated .chm

Everything except that last `hhc.exe` call is pure standard library.
"""

import os
import glob

import chm_reader as cr
import html_translator as ht


def decompile_chm(chm_path, dst_dir):
    """Extract every internal file of a .chm into dst_dir. Returns file list."""
    return cr.extract_chm(chm_path, dst_dir)


def _work_dir(chm_path):
    d = os.path.dirname(chm_path)
    stem = os.path.splitext(os.path.basename(chm_path))[0]
    return os.path.join(d, stem + "_chm")


def _html_pages(folder):
    pages = []
    for root, _dirs, files in os.walk(folder):
        for fn in files:
            if fn.lower().endswith((".htm", ".html")):
                pages.append(os.path.join(root, fn))
    return sorted(pages)


# --------------------------------------------------------------------------
# Step 1:  chm -> folder of pages + one .txt per page
# --------------------------------------------------------------------------

def chm_2_texts(chm_path, dst_dir=None):
    dst_dir = dst_dir or _work_dir(chm_path)
    decompile_chm(chm_path, dst_dir)
    pages = _html_pages(dst_dir)
    txts = [ht.convert_html_2_text(p) for p in pages]      # <page>.txt beside each page
    _write_hhp(chm_path, dst_dir, pages)
    return dst_dir, list(zip(pages, txts))


# --------------------------------------------------------------------------
# Step 3:  rebuild translated pages  (expects <page>_tran.txt next to <page>)
# --------------------------------------------------------------------------

def texts_2_html(dst_dir):
    out = []
    for page in _html_pages(dst_dir):
        if page.endswith("_tran.html"):
            continue
        stem = os.path.splitext(page)[0]
        tran_txt = stem + "_tran.txt"
        if os.path.exists(tran_txt):
            out.append(ht.convert_txt_2_html(page, tran_txt))
    return out


# --------------------------------------------------------------------------
# Emit a minimal HTML Help project so `hhc.exe project.hhp` can recompile.
# --------------------------------------------------------------------------

def _write_hhp(chm_path, dst_dir, pages):
    stem = os.path.splitext(os.path.basename(chm_path))[0]
    rel_pages = [os.path.relpath(p, dst_dir).replace("/", "\\") for p in pages]
    # a decompiled .hhc/.hhk (table of contents / index), if any survived
    hhc = next((os.path.relpath(p, dst_dir) for p in
                glob.glob(os.path.join(dst_dir, "**", "*.hhc"), recursive=True)), None)
    hhk = next((os.path.relpath(p, dst_dir) for p in
                glob.glob(os.path.join(dst_dir, "**", "*.hhk"), recursive=True)), None)
    default = rel_pages[0] if rel_pages else ""

    lines = ["[OPTIONS]",
             "Compatibility=1.1 or later",
             "Compiled file=%s_tran.chm" % stem,
             "Default topic=%s" % default,
             "Display compile progress=No",
             "Language=0x409 English (United States)"]
    if hhc:
        lines.append("Contents file=%s" % hhc.replace("/", "\\"))
    if hhk:
        lines.append("Index file=%s" % hhk.replace("/", "\\"))
    lines += ["", "[FILES]"] + rel_pages + [""]
    hhp = os.path.join(dst_dir, "project.hhp")
    with open(hhp, "w", encoding="mbcs" if os.name == "nt" else "utf-8",
              errors="replace") as f:
        f.write("\r\n".join(lines))
    return hhp


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _main(argv):
    if len(argv) == 2 and argv[0] == "extract":
        folder, pairs = chm_2_texts(argv[1])
        print("decompiled + extracted %d page(s) to %s" % (len(pairs), folder))
        print("translate each <page>.txt -> <page>_tran.txt, then run:")
        print("  python chm_translator.py rebuild \"%s\"" % folder)
        return 0
    if len(argv) == 2 and argv[0] == "rebuild":
        outs = texts_2_html(argv[1])
        print("rebuilt %d translated page(s)" % len(outs))
        print("recompile with:  hhc.exe \"%s\""
              % os.path.join(argv[1], "project.hhp"))
        return 0
    if len(argv) == 3 and argv[0] == "decompile":
        n = len(decompile_chm(argv[1], argv[2]))
        print("extracted %d files to %s" % (n, argv[2]))
        return 0
    print("usage:\n"
          "  python chm_translator.py decompile file.chm out_dir/\n"
          "  python chm_translator.py extract   file.chm   # step 1 (+ .hhp)\n"
          "  python chm_translator.py rebuild   file_chm/  # step 3")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_main(sys.argv[1:]))
