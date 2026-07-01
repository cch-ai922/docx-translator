# -*- coding: utf-8 -*-
"""
pptx_translator.py  --  pure-stdlib PowerPoint (.pptx) translator (steps 1 & 3).

Same shape as the docx translator:

    convert_pptx_2_text(source_pptx_path)        -> writes  <stem>.txt
    convert_txt_2_pptx(source_pptx_path,         -> writes  <stem>_tran.pptx
                       source_trantext_path)         (+ "machine translation"
                                                       watermark on every slide)

A translatable unit is a DrawingML paragraph <a:p>.  Slides are processed in
true presentation order (per ppt/presentation.xml -> sldIdLst), and every
<a:p> on every slide contributes exactly one line, so the text file's line
count equals the total paragraph count across the deck.
"""

import re
import xml.etree.ElementTree as ET

from ooxml_core import (
    NS, q, local, parse_xml, serialize_xml, read_part, list_parts, rewrite_zip,
    write_lines, read_lines, derive_path, resolve_target,
    RunSchema, extract_unit_line, rebuild_unit,
)

# DrawingML paragraph: <a:p> -> [<a:pPr>, <a:r>(<a:rPr><a:t>), <a:br>, <a:fld>, <a:endParaRPr>]
PPTX = RunSchema(ns="a", para="p", run="r", text="t", rpr="rPr",
                 props="pPr", tail="endParaRPr", set_xml_space=False)


# --------------------------------------------------------------------------
# Slide discovery (true display order)
# --------------------------------------------------------------------------

def _slide_parts_in_order(path):
    pres = read_part(path, "ppt/presentation.xml")
    rels = read_part(path, "ppt/_rels/presentation.xml.rels")
    if pres is not None and rels is not None:
        relmap = {rel.get("Id"): rel.get("Target") for rel in parse_xml(rels)}
        lst = parse_xml(pres).find(q("p", "sldIdLst"))
        if lst is not None:
            order = []
            for sld in lst.findall(q("p", "sldId")):
                tgt = relmap.get(sld.get(q("r", "id")))
                if tgt:
                    order.append(resolve_target("ppt", tgt))
            if order:
                return order
    # fallback: numeric filename sort
    slides = [n for n in list_parts(path)
              if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    return sorted(slides, key=lambda n: int(re.search(r"(\d+)", n).group(1)))


def _slide_paragraphs(root):
    # DrawingML never nests <a:p> inside a run, so iter() is unambiguous and
    # yields paragraphs (incl. those in tables and grouped shapes) in order.
    return list(root.iter(q("a", "p")))


# --------------------------------------------------------------------------
# Step 1:  pptx -> txt
# --------------------------------------------------------------------------

def convert_pptx_2_text(source_pptx_path):
    lines = []
    for part in _slide_parts_in_order(source_pptx_path):
        data = read_part(source_pptx_path, part)
        if data is None:
            continue
        root = parse_xml(data)
        for p in _slide_paragraphs(root):
            lines.append(extract_unit_line(p, PPTX))
    out_txt = derive_path(source_pptx_path, "", ".txt")
    write_lines(out_txt, lines)
    return out_txt


# --------------------------------------------------------------------------
# Step 3:  pptx + translated txt -> translated pptx (+ watermark)
# --------------------------------------------------------------------------

def convert_txt_2_pptx(source_pptx_path, source_trantext_path,
                       watermark_text="machine translation"):
    parts = _slide_parts_in_order(source_pptx_path)

    # first pass: count paragraphs for a sane line-count check
    parsed = []
    total = 0
    for part in parts:
        data = read_part(source_pptx_path, part)
        if data is None:
            continue
        root = parse_xml(data)
        paras = _slide_paragraphs(root)
        parsed.append((part, root, paras))
        total += len(paras)

    lines = read_lines(source_trantext_path, expected_count=total)

    replacements = {}
    idx = 0
    for part, root, paras in parsed:
        for p in paras:
            if idx < len(lines):
                rebuild_unit(p, lines[idx], PPTX)
            idx += 1
        if watermark_text:
            _add_slide_watermark(root, watermark_text)
        replacements[part] = serialize_xml(root)

    out_pptx = derive_path(source_pptx_path, "_tran", ".pptx")
    rewrite_zip(source_pptx_path, out_pptx, replacements)
    return out_pptx


# --------------------------------------------------------------------------
# Watermark: a centered, semi-transparent text box added to each slide.
# (PowerPoint has no "watermark" concept; a master-level shape can be hidden
# by a layout, so adding it per-slide is the reliable "every slide" choice.)
# --------------------------------------------------------------------------

def _add_slide_watermark(slide_root, text):
    spTree = slide_root.find(q("p", "cSld") + "/" + q("p", "spTree"))
    if spTree is None:
        return
    safe = (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    sp_xml = (
        '<p:sp xmlns:p="%s" xmlns:a="%s">'
        '<p:nvSpPr><p:cNvPr id="990001" name="MT Watermark"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm rot="19800000"><a:off x="1257300" y="2971800"/>'
        '<a:ext cx="6629400" cy="1828800"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        '</p:spPr>'
        '<p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/>'
        '<a:r><a:rPr lang="en-US" sz="4000" b="1">'
        '<a:solidFill><a:srgbClr val="D9D9D9"><a:alpha val="45000"/></a:srgbClr>'
        '</a:solidFill></a:rPr><a:t>%s</a:t></a:r></a:p>'
        '</p:txBody></p:sp>' % (NS["p"], NS["a"], safe)
    )
    spTree.append(parse_xml(sp_xml))


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print("usage:\n"
              "  python pptx_translator.py extract  source.pptx\n"
              "  python pptx_translator.py rebuild  source.pptx  source_tran.txt")
        return 0
    if argv[0] == "extract" and len(argv) == 2:
        print(convert_pptx_2_text(argv[1]))
    elif argv[0] == "rebuild" and len(argv) == 3:
        print(convert_txt_2_pptx(argv[1], argv[2]))
    else:
        print("bad arguments; use --help")
        return 2
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_main(sys.argv[1:]))
