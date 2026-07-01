# -*- coding: utf-8 -*-
"""
xlsx_translator.py  --  pure-stdlib Excel (.xlsx) translator (steps 1 & 3).

    convert_xlsx_2_text(source_xlsx_path)        -> writes  <stem>.txt
    convert_txt_2_xlsx(source_xlsx_path,         -> writes  <stem>_tran.xlsx
                       source_trantext_path)

For a spreadsheet the natural translatable unit is a *shared string*
(xl/sharedStrings.xml -> <si>), not a paragraph: numbers, dates and formulas
live in the sheets and are left untouched, while almost all visible cell TEXT
is pooled here exactly once.  One <si> == one line.

A cell string may contain hard line breaks; those are reversibly encoded
(see ooxml_core.pack_newlines) so the one-unit-per-line contract holds.

Notes / limits
--------------
* Rich-text runs inside an <si> are flattened to a single run (keeping the
  first run's formatting) -- the same trade-off as the docx translator.
* Phonetic guides (<rPh>) are dropped on rebuild (their character offsets no
  longer match translated text); <phoneticPr> is kept.
* Inline strings (<c t="inlineStr"><is>...) are uncommon and not handled here.
* Excel has no real page watermark; an optional centered header text is offered
  via watermark_text but defaults to off.
"""

import copy
import xml.etree.ElementTree as ET

from ooxml_core import (
    q, local, parse_xml, serialize_xml, read_part, rewrite_zip,
    write_lines, read_lines, derive_path, pack_newlines, unpack_newlines,
    XML_SPACE,
)

SHARED_STRINGS = "xl/sharedStrings.xml"


def _shared_items(root):
    return [si for si in root if local(si) == "si"]


def _si_text(si):
    """Concatenate the translatable text of one <si> (ignoring phonetic runs)."""
    parts = []
    for child in si:
        ln = local(child)
        if ln == "t":                       # plain string
            parts.append(child.text or "")
        elif ln == "r":                     # rich-text run
            t = child.find(q("x", "t"))
            if t is not None:
                parts.append(t.text or "")
        # rPh (phonetic) / phoneticPr -> not translatable
    return "".join(parts)


# --------------------------------------------------------------------------
# Step 1:  xlsx -> txt
# --------------------------------------------------------------------------

def convert_xlsx_2_text(source_xlsx_path):
    data = read_part(source_xlsx_path, SHARED_STRINGS)
    lines = []
    if data is not None:
        for si in _shared_items(parse_xml(data)):
            lines.append(pack_newlines(_si_text(si)))
    out_txt = derive_path(source_xlsx_path, "", ".txt")
    write_lines(out_txt, lines)
    return out_txt


# --------------------------------------------------------------------------
# Step 3:  xlsx + translated txt -> translated xlsx
# --------------------------------------------------------------------------

def _rebuild_si(si, line):
    text = unpack_newlines(line)

    # keep the first run's formatting + any phoneticPr; drop everything else
    first_rpr = None
    phonetic_pr = None
    for child in si:
        ln = local(child)
        if ln == "r" and first_rpr is None:
            rpr = child.find(q("x", "rPr"))
            if rpr is not None:
                first_rpr = copy.deepcopy(rpr)
        elif ln == "phoneticPr":
            phonetic_pr = child
    for child in list(si):
        si.remove(child)

    if first_rpr is not None:
        r = ET.SubElement(si, q("x", "r"))
        r.append(first_rpr)
        t = ET.SubElement(r, q("x", "t"))
    else:
        t = ET.SubElement(si, q("x", "t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    if phonetic_pr is not None:
        si.append(phonetic_pr)


def convert_txt_2_xlsx(source_xlsx_path, source_trantext_path,
                       watermark_text=None):
    data = read_part(source_xlsx_path, SHARED_STRINGS)
    if data is None:
        # nothing to translate (workbook has no shared strings)
        out_xlsx = derive_path(source_xlsx_path, "_tran", ".xlsx")
        rewrite_zip(source_xlsx_path, out_xlsx)
        return out_xlsx

    root = parse_xml(data)
    items = _shared_items(root)
    lines = read_lines(source_trantext_path, expected_count=len(items))
    for si, line in zip(items, lines):
        _rebuild_si(si, line)

    replacements = {SHARED_STRINGS: serialize_xml(root)}
    if watermark_text:
        _add_sheet_headers(source_xlsx_path, watermark_text, replacements)

    out_xlsx = derive_path(source_xlsx_path, "_tran", ".xlsx")
    rewrite_zip(source_xlsx_path, out_xlsx, replacements)
    return out_xlsx


# --------------------------------------------------------------------------
# Optional watermark analog: a centered odd-page header on each worksheet.
# --------------------------------------------------------------------------

# worksheet children that must come AFTER <headerFooter>; insert before the
# first of these we find, else append.
_AFTER_HF = ("rowBreaks", "colBreaks", "customProperties", "cellWatches",
             "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
             "legacyDrawingHF", "picture", "oleObjects", "controls",
             "webPublishItems", "tableParts", "extLst")


def _add_sheet_headers(source_xlsx_path, text, replacements):
    for part in _worksheet_parts(source_xlsx_path):
        data = read_part(source_xlsx_path, part)
        if data is None:
            continue
        root = parse_xml(data)
        if root.find(q("x", "headerFooter")) is not None:
            hf = root.find(q("x", "headerFooter"))
        else:
            hf = ET.Element(q("x", "headerFooter"))
            insert_at = len(root)
            for i, child in enumerate(list(root)):
                if local(child) in _AFTER_HF:
                    insert_at = i
                    break
            root.insert(insert_at, hf)
        odd = hf.find(q("x", "oddHeader"))
        if odd is None:
            odd = ET.SubElement(hf, q("x", "oddHeader"))
        odd.text = "&C" + text
        replacements[part] = serialize_xml(root)


def _worksheet_parts(source_xlsx_path):
    import re
    from ooxml_core import list_parts
    return [n for n in list_parts(source_xlsx_path)
            if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print("usage:\n"
              "  python xlsx_translator.py extract  source.xlsx\n"
              "  python xlsx_translator.py rebuild  source.xlsx  source_tran.txt")
        return 0
    if argv[0] == "extract" and len(argv) == 2:
        print(convert_xlsx_2_text(argv[1]))
    elif argv[0] == "rebuild" and len(argv) == 3:
        print(convert_txt_2_xlsx(argv[1], argv[2]))
    else:
        print("bad arguments; use --help")
        return 2
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_main(sys.argv[1:]))
