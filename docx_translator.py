# -*- coding: utf-8 -*-
"""
docx_translator.py  --  pure-stdlib Word (.docx) translator (steps 1 & 3).

Public API (exactly as specified):

    convert_docx_2_text(source_docx_path)        -> writes  <stem>.txt
    convert_txt_2_docx(source_docx_path,         -> writes  <stem>_tran.docx
                       source_trantext_path)         (with an every-page
                                                       "machine translation"
                                                       watermark)

Step 2 (the actual translation of the .txt file) is done by a 3rd-party engine
and is intentionally NOT part of this module.

Implementation notes
---------------------
* The main story lives in `word/document.xml`.
* A translatable unit is a Word paragraph <w:p>, taken in document order,
  descending through tables and block-level content controls but NOT into
  runs / drawings / text boxes (their inner text is preserved untranslated).
* Inline images, math, hyperlinks, bookmarks, breaks, tabs and fields become
  SYM_00n placeholders; only <w:t> text is extracted.  See ooxml_core.py.
"""

import os
import copy
import xml.etree.ElementTree as ET

from ooxml_core import (
    NS, q, local, parse_xml, serialize_xml, read_part, rewrite_zip,
    write_lines, read_lines, derive_path,
    RunSchema, extract_unit_line, rebuild_unit,
)

DOCUMENT_PART = "word/document.xml"

# WordprocessingML paragraph: <w:p> -> [<w:pPr>, runs, hyperlinks, math, ...]
DOCX = RunSchema(ns="w", para="p", run="r", text="t", rpr="rPr",
                 props="pPr", tail=None, set_xml_space=True)


# --------------------------------------------------------------------------
# Paragraph collection (document order, into tables / SDTs, not into runs)
# --------------------------------------------------------------------------

def iter_block_paragraphs(parent):
    """Yield every <w:p> reachable as block content, in document order."""
    for child in parent:
        ln = local(child)
        if ln == "p":
            yield child
        elif ln == "tbl":
            for tr in child:
                if local(tr) == "tr":
                    for tc in tr:
                        if local(tc) == "tc":
                            for p in iter_block_paragraphs(tc):
                                yield p
        elif ln == "sdt":
            content = child.find(q("w", "sdtContent"))
            if content is not None:
                for p in iter_block_paragraphs(content):
                    yield p


def _body(root):
    body = root.find(q("w", "body"))
    return body if body is not None else root


# --------------------------------------------------------------------------
# Step 1:  docx -> txt
# --------------------------------------------------------------------------

def convert_docx_2_text(source_docx_path):
    data = read_part(source_docx_path, DOCUMENT_PART)
    if data is None:
        raise ValueError("%s is not a valid .docx (no %s)"
                         % (source_docx_path, DOCUMENT_PART))
    root = parse_xml(data)
    lines = [extract_unit_line(p, DOCX) for p in iter_block_paragraphs(_body(root))]

    out_txt = derive_path(source_docx_path, "", ".txt")
    write_lines(out_txt, lines)
    return out_txt


# --------------------------------------------------------------------------
# Step 3:  docx + translated txt -> translated docx (+ watermark)
# --------------------------------------------------------------------------

def convert_txt_2_docx(source_docx_path, source_trantext_path,
                       watermark_text="machine translation"):
    data = read_part(source_docx_path, DOCUMENT_PART)
    if data is None:
        raise ValueError("%s is not a valid .docx" % source_docx_path)
    root = parse_xml(data)
    paragraphs = list(iter_block_paragraphs(_body(root)))

    lines = read_lines(source_trantext_path, expected_count=len(paragraphs))
    for p, line in zip(paragraphs, lines):
        rebuild_unit(p, line, DOCX)

    replacements = {DOCUMENT_PART: serialize_xml(root)}
    additions = {}
    if watermark_text:
        _add_watermark(source_docx_path, watermark_text, replacements, additions)

    out_docx = derive_path(source_docx_path, "_tran", ".docx")
    rewrite_zip(source_docx_path, out_docx, replacements, additions)
    return out_docx


# --------------------------------------------------------------------------
# Watermark ("machine translation" on every page)
# --------------------------------------------------------------------------
# A Word watermark is a VML text shape living in the section *header* part(s).
# To cover every page we:
#   1. inject the watermark shape into every header the document already uses
#      (covers pages that already show a header), and
#   2. add a new watermark-only header and reference it (as default/first/even)
#      from every <w:sectPr> that lacks a header of that type.
# Inactive references (e.g. a first-page header when titlePg is off) are simply
# ignored by Word, so adding all three types is safe.

_WM_HEADER_PART = "word/header_mt_watermark.xml"
_WM_REL_ID = "rIdMtWatermark"

# Canonical WordArt textplate shape type (t136), reused by the watermark shape.
_VML_SHAPETYPE = (
    '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136"'
    ' adj="10800" path="m@7,0l@8,0m@5,21600l@6,21600e">'
    '<v:formulas>'
    '<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/>'
    '<v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>'
    '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>'
    '<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/>'
    '<v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/>'
    '<v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>'
    '<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/>'
    '</v:formulas>'
    '<v:path textpathok="t" o:connecttype="custom"'
    ' o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"'
    ' o:connectangles="270,180,90,0"/>'
    '<v:textpath on="t" fitshape="t"/>'
    '</v:shapetype>'
)


def _watermark_paragraph_xml(text, spid):
    """A <w:p> (as a namespaced fragment string) holding the watermark shape."""
    safe = (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))
    return (
        '<w:p xmlns:w="%s" xmlns:v="%s" xmlns:o="%s" xmlns:w10="%s">'
        '<w:r><w:rPr><w:noProof/></w:rPr><w:pict>'
        '%s'
        '<v:shape id="%s" o:spid="%s" type="#_x0000_t136"'
        ' style="position:absolute;left:0;text-align:left;'
        'margin-left:0;margin-top:0;width:450pt;height:90pt;rotation:315;'
        'z-index:-251654144;mso-position-horizontal:center;'
        'mso-position-horizontal-relative:margin;mso-position-vertical:center;'
        'mso-position-vertical-relative:margin"'
        ' fillcolor="#d9d9d9" stroked="f">'
        '<v:fill opacity=".5"/>'
        '<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt"'
        ' string="%s"/>'
        '</v:shape></w:pict></w:r></w:p>'
        % (NS["w"], NS["v"], NS["o"], NS["w10"],
           _VML_SHAPETYPE, spid, spid, safe)
    )


def _new_watermark_header_xml(text):
    inner = _watermark_paragraph_xml(text, "_x0000_s2049")
    # strip the fragment's own <w:p ...> namespace decls duplication by simply
    # nesting it: ET will re-emit clean declarations on the header root.
    hdr = ('<w:hdr xmlns:w="%s" xmlns:v="%s" xmlns:o="%s" xmlns:w10="%s"'
           ' xmlns:r="%s">%s</w:hdr>'
           % (NS["w"], NS["v"], NS["o"], NS["w10"], NS["r"],
              _strip_root(inner)))
    return serialize_xml(parse_xml(hdr))


def _strip_root(p_fragment):
    """Return the inner <w:p>...</w:p> with namespace decls removed (re-declared
    by the wrapping <w:hdr>)."""
    # The fragment already declares the namespaces it needs; re-parsing and
    # re-serializing inside the header is simplest and keeps it valid.
    el = parse_xml(p_fragment)
    return ET.tostring(el, encoding="unicode")


def _existing_header_parts(source_docx_path):
    """List header part names referenced by word/_rels/document.xml.rels."""
    rels_data = read_part(source_docx_path, "word/_rels/document.xml.rels")
    if rels_data is None:
        return []
    rels = parse_xml(rels_data)
    headers = []
    for rel in rels:
        if rel.get("Type", "").endswith("/header"):
            target = rel.get("Target", "")
            if target.startswith("/"):
                name = target.lstrip("/")
            else:
                name = "word/" + target
            headers.append(name)
    return headers


def _inject_into_existing_headers(source_docx_path, text, replacements, spid_base):
    for i, part in enumerate(_existing_header_parts(source_docx_path)):
        data = replacements.get(part) or read_part(source_docx_path, part)
        if data is None:
            continue
        hdr = parse_xml(data)
        frag = _watermark_paragraph_xml(text, "_x0000_s20%02d" % (50 + i))
        hdr.insert(0, parse_xml(frag))
        replacements[part] = serialize_xml(hdr)


def _add_watermark(source_docx_path, text, replacements, additions):
    # 1. watermark the headers the document already uses
    _inject_into_existing_headers(source_docx_path, text, replacements, 0)

    # 2. ship a watermark-only header part ...
    additions[_WM_HEADER_PART] = _new_watermark_header_xml(text)

    # ... declare its content type ...
    ct_name = "[Content_Types].xml"
    ct_data = replacements.get(ct_name) or read_part(source_docx_path, ct_name)
    ct = parse_xml(ct_data)
    override = ET.SubElement(ct, q("ct", "Override"))
    override.set("PartName", "/" + _WM_HEADER_PART)
    override.set("ContentType",
                 "application/vnd.openxmlformats-officedocument."
                 "wordprocessingml.header+xml")
    replacements[ct_name] = serialize_xml(ct)

    # ... register a relationship from the main document ...
    rels_name = "word/_rels/document.xml.rels"
    rels_data = replacements.get(rels_name) or read_part(source_docx_path, rels_name)
    rels = parse_xml(rels_data)
    rel = ET.SubElement(rels, q("rel", "Relationship"))
    rel.set("Id", _WM_REL_ID)
    rel.set("Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/"
            "relationships/header")
    rel.set("Target", os.path.basename(_WM_HEADER_PART))
    replacements[rels_name] = serialize_xml(rels)

    # 3. reference the watermark header from every section that lacks a header
    doc_data = replacements[DOCUMENT_PART]
    root = parse_xml(doc_data)
    for sectPr in root.iter(q("w", "sectPr")):
        present = set()
        for hr in sectPr.findall(q("w", "headerReference")):
            present.add(hr.get(q("w", "type")))
        idx = 0
        for htype in ("default", "first", "even"):
            if htype in present:
                continue
            hr = ET.Element(q("w", "headerReference"))
            hr.set(q("w", "type"), htype)
            hr.set(q("r", "id"), _WM_REL_ID)
            sectPr.insert(idx, hr)     # header/footer refs go first in sectPr
            idx += 1
    replacements[DOCUMENT_PART] = serialize_xml(root)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print("usage:\n"
              "  python docx_translator.py extract  source.docx\n"
              "  python docx_translator.py rebuild  source.docx  source_tran.txt")
        return 0
    cmd = argv[0]
    if cmd == "extract" and len(argv) == 2:
        print(convert_docx_2_text(argv[1]))
    elif cmd == "rebuild" and len(argv) == 3:
        print(convert_txt_2_docx(argv[1], argv[2]))
    else:
        print("bad arguments; use --help")
        return 2
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_main(sys.argv[1:]))
