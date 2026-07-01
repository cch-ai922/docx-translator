# -*- coding: utf-8 -*-
"""
translate.py -- thin dispatcher over the docx / pptx / xlsx translators.

    python translate.py extract  source.docx
    python translate.py rebuild  source.docx  source_tran.txt

Picks the right engine from the file extension. Step 2 (the actual machine
translation of the .txt) is performed by your external engine in between.
"""
import os
import sys

import docx_translator as _docx
import pptx_translator as _pptx
import xlsx_translator as _xlsx
import html_translator as _html

_EXTRACT = {
    ".docx": _docx.convert_docx_2_text,
    ".pptx": _pptx.convert_pptx_2_text,
    ".xlsx": _xlsx.convert_xlsx_2_text,
    ".html": _html.convert_html_2_text,
    ".htm": _html.convert_html_2_text,
}
_REBUILD = {
    ".docx": _docx.convert_txt_2_docx,
    ".pptx": _pptx.convert_txt_2_pptx,
    ".xlsx": _xlsx.convert_txt_2_xlsx,
    ".html": _html.convert_txt_2_html,
    ".htm": _html.convert_txt_2_html,
}


def _ext(path):
    e = os.path.splitext(path)[1].lower()
    if e not in _EXTRACT:
        raise SystemExit("unsupported type %r (want .docx/.pptx/.xlsx/.html)" % e)
    return e


def extract(source_path):
    return _EXTRACT[_ext(source_path)](source_path)


def rebuild(source_path, trantext_path):
    return _REBUILD[_ext(source_path)](source_path, trantext_path)


def _main(argv):
    if len(argv) == 2 and argv[0] == "extract":
        print(extract(argv[1])); return 0
    if len(argv) == 3 and argv[0] == "rebuild":
        print(rebuild(argv[1], argv[2])); return 0
    print(__doc__)
    return 0 if (argv and argv[0] in ("-h", "--help")) else 2


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
