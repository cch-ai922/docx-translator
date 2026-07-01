# -*- coding: utf-8 -*-
"""Synthesize minimal .pptx and .xlsx, run the pipelines, validate. Pure stdlib."""
import os
import sys
import zipfile
import tempfile
import xml.etree.ElementTree as ET

import pptx_translator as pt
import xlsx_translator as xt
from ooxml_core import q

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# ----------------------------- PPTX fixtures ------------------------------
PPTX_CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
 <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"""
PPTX_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""
PPTX_PRES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="%s" xmlns:r="%s">
 <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>""" % (P, R)
PPTX_PRES_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"""
PPTX_SLIDE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="%s" xmlns:a="%s" xmlns:r="%s"><p:cSld><p:spTree>
 <p:sp><p:txBody><a:bodyPr/><a:lstStyle/>
   <a:p><a:r><a:t>Title text</a:t></a:r></a:p>
   <a:p><a:r><a:t xml:space="preserve">Second </a:t></a:r><a:br/><a:r><a:t>line</a:t></a:r></a:p>
 </p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>""" % (P, A, R)


def build_pptx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", PPTX_CT)
        z.writestr("_rels/.rels", PPTX_ROOT_RELS)
        z.writestr("ppt/presentation.xml", PPTX_PRES)
        z.writestr("ppt/_rels/presentation.xml.rels", PPTX_PRES_RELS)
        z.writestr("ppt/slides/slide1.xml", PPTX_SLIDE)


# ----------------------------- XLSX fixtures ------------------------------
XLSX_CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
 <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
 <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""
XLSX_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
XLSX_WB = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="%s" xmlns:r="%s"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""" % (X, R)
XLSX_WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
XLSX_SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="%s"><sheetData>
 <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>
 <row r="2"><c r="A2" t="s"><v>1</v></c></row>
 <row r="3"><c r="A3" t="s"><v>2</v></c></row>
</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>""" % X
XLSX_SST = ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="%s" count="3" uniqueCount="3">
 <si><t>Apple</t></si>
 <si><r><rPr><b/></rPr><t xml:space="preserve">Red </t></r><r><t>Banana</t></r></si>
 <si><t xml:space="preserve">line1\nline2</t></si>
</sst>""" % X)


def build_xlsx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", XLSX_CT)
        z.writestr("_rels/.rels", XLSX_ROOT_RELS)
        z.writestr("xl/workbook.xml", XLSX_WB)
        z.writestr("xl/_rels/workbook.xml.rels", XLSX_WB_RELS)
        z.writestr("xl/worksheets/sheet1.xml", XLSX_SHEET)
        z.writestr("xl/sharedStrings.xml", XLSX_SST)


# ------------------------------- helpers ----------------------------------
fails = []
def check(cond, msg):
    print(("  ok  " if cond else " FAIL ") + msg)
    if not cond:
        fails.append(msg)

def read_lines(p):
    ls = open(p, encoding="utf-8").read().split("\n")
    if ls and ls[-1] == "":
        ls.pop()
    return ls

def fake_translate(txt, out):
    with open(out, "w", encoding="utf-8", newline="") as f:
        for ln in read_lines(txt):
            f.write(("《%s》" % ln) if ln else "")
            f.write("\n")

def all_text(xml_bytes, ns):
    return "".join((e.text or "") for e in ET.fromstring(xml_bytes).iter("{%s}t" % ns))


# -------------------------------- PPTX test -------------------------------
def test_pptx(tmp):
    print("== PPTX ==")
    src = os.path.join(tmp, "deck.pptx"); build_pptx(src)
    txt = pt.convert_pptx_2_text(src)
    lines = read_lines(txt)
    for i, ln in enumerate(lines):
        print("   [%d] %r" % (i, ln))
    check(lines == ["Title text", "Second SYM_001line"],
          "paragraphs extracted, <a:br> -> SYM_001")
    tran = txt[:-4] + "_tran.txt"; fake_translate(txt, tran)
    out = pt.convert_txt_2_pptx(src, tran)
    slide = zipfile.ZipFile(out).read("ppt/slides/slide1.xml")
    ET.fromstring(slide)  # well-formed?
    check(True, "rebuilt slide is well-formed XML")
    txt_blob = all_text(slide, A)
    check("《Title text》" in txt_blob, "translated slide text present")
    check(sum(1 for _ in ET.fromstring(slide).iter(q("a", "br"))) == 1,
          "<a:br> preserved (1)")
    check("machine translation" in txt_blob, "watermark added to slide")


# -------------------------------- XLSX test -------------------------------
def test_xlsx(tmp):
    print("== XLSX ==")
    src = os.path.join(tmp, "book.xlsx"); build_xlsx(src)
    txt = xt.convert_xlsx_2_text(src)
    lines = read_lines(txt)
    for i, ln in enumerate(lines):
        print("   [%d] %r" % (i, ln))
    check(lines == ["Apple", "Red Banana", "line1\\nline2"],
          "shared strings extracted; rich run merged; newline packed")
    tran = txt[:-4] + "_tran.txt"; fake_translate(txt, tran)
    out = xt.convert_txt_2_xlsx(src, tran, watermark_text="machine translation")
    sst = zipfile.ZipFile(out).read("xl/sharedStrings.xml")
    ET.fromstring(sst)
    check(True, "rebuilt sharedStrings is well-formed XML")
    blob = all_text(sst, X)
    check("《Apple》" in blob and "《Red Banana》" in blob, "translated strings present")
    check("line1\nline2" in blob, "packed newline restored to real newline")
    sheet = zipfile.ZipFile(out).read("xl/worksheets/sheet1.xml")
    ET.fromstring(sheet)
    check(b"machine translation" in sheet, "watermark header added to sheet")
    # numeric cell untouched
    check(b'<v>42</v>' in sheet or b"42" in sheet, "numeric cell value untouched")


def main():
    tmp = tempfile.mkdtemp(prefix="ooxmltest_")
    test_pptx(tmp)
    test_xlsx(tmp)
    print()
    if fails:
        print("RESULT: %d FAILURE(S)" % len(fails)); return 1
    print("RESULT: ALL CHECKS PASSED"); return 0


if __name__ == "__main__":
    sys.exit(main())
