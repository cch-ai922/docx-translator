# -*- coding: utf-8 -*-
"""
selftest_docx.py -- synthesize a real .docx, run the full pipeline, validate.
Pure stdlib; no external .docx needed.  Run:  python selftest_docx.py
"""
import os
import sys
import zipfile
import tempfile
import xml.etree.ElementTree as ET

import docx_translator as dt
from ooxml_core import q

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Default Extension="png" ContentType="image/png"/>
 <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
 <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
 <Relationship Id="rIdLink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com" TargetMode="External"/>
 <Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>"""

HEADER1 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="%s"><w:p><w:r><w:t>Original header</w:t></w:r></w:p></w:hdr>""" % W

DRAWING = """
   <w:r><w:drawing>
     <wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="{WP}">
       <wp:extent cx="914400" cy="914400"/>
       <wp:docPr id="1" name="Picture 1"/>
       <a:graphic xmlns:a="{A}">
         <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
           <pic:pic xmlns:pic="{PIC}">
             <pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>
             <pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
             <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
               <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
           </pic:pic>
         </a:graphicData>
       </a:graphic>
     </wp:inline>
   </w:drawing></w:r>""".format(WP=WP, A=A, PIC=PIC)

DOCUMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:m="{M}">
 <w:body>
  <w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>
  <w:p><w:r><w:t xml:space="preserve">See picture: </w:t></w:r>{DRAW}<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>
  <w:p><w:r><w:t xml:space="preserve">Formula </w:t></w:r><m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath><w:r><w:t xml:space="preserve"> done</w:t></w:r></w:p>
  <w:p><w:r><w:t xml:space="preserve">Link: </w:t></w:r><w:hyperlink r:id="rIdLink1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>click here</w:t></w:r></w:hyperlink></w:p>
  <w:p></w:p>
  <w:p>{DRAW}</w:p>
  <w:tbl>
   <w:tr><w:tc><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:sectPr>
   <w:headerReference w:type="default" r:id="rIdHdr1"/>
   <w:pgSz w:w="12240" w:h="15840"/>
  </w:sectPr>
 </w:body>
</w:document>""".format(W=W, R=R, M=M, DRAW=DRAWING)

# 1x1 transparent PNG
PNG = (b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
       b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00"
       b"\x01\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82")


def build_docx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", DOCUMENT)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/header1.xml", HEADER1)
        z.writestr("word/media/image1.png", PNG)


def count_tags(xml_bytes, *clark_tags):
    root = ET.fromstring(xml_bytes)
    return {t: sum(1 for _ in root.iter(t)) for t in clark_tags}


def main():
    tmp = tempfile.mkdtemp(prefix="docxtest_")
    src = os.path.join(tmp, "source.docx")
    build_docx(src)

    fails = []
    def check(cond, msg):
        print(("  ok  " if cond else " FAIL ") + msg)
        if not cond:
            fails.append(msg)

    # ---- step 1: extract ------------------------------------------------
    txt = dt.convert_docx_2_text(src)
    with open(txt, encoding="utf-8") as f:
        lines = f.read().split("\n")
    if lines and lines[-1] == "":
        lines.pop()

    print("Extracted lines:")
    for i, ln in enumerate(lines):
        print("   [%d] %r" % (i, ln))

    check(len(lines) == 7, "line count == paragraph count (7)")
    check(lines[0] == "Hello world", "para0 two runs merged -> 'Hello world'")
    check("SYM_001" in lines[1] and lines[1].startswith("See picture:"),
          "para1 image -> SYM token inline with text")
    check("SYM_001" in lines[2], "para2 math -> SYM token")
    check(lines[3] == "Link: SYM_001", "para3 hyperlink preserved whole as SYM")
    check(lines[4] == "", "para4 empty -> empty line")
    check(lines[5] == "", "para5 image-only -> empty line (left untouched)")
    check(lines[6] == "Cell text", "para6 table cell text extracted")

    # ---- step 2 (fake): wrap each non-empty line, keep SYM tokens --------
    tran = txt[:-4] + "_tran.txt"
    with open(tran, "w", encoding="utf-8", newline="") as f:
        for ln in lines:
            f.write(("《%s》" % ln) if ln else "")
            f.write("\n")

    # ---- step 3: rebuild ------------------------------------------------
    out = dt.convert_txt_2_docx(src, tran)
    check(os.path.exists(out), "output .docx written")

    # ---- validate output ------------------------------------------------
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        doc = z.read("word/document.xml")
        ct = z.read("[Content_Types].xml")
        rels = z.read("word/_rels/document.xml.rels")
        hdr1 = z.read("word/header1.xml")

    # well-formed?
    try:
        ET.fromstring(doc); ET.fromstring(ct); ET.fromstring(rels)
        check(True, "all rewritten parts are well-formed XML")
    except ET.ParseError as e:
        check(False, "rewritten XML well-formed (%s)" % e)

    src_counts = count_tags(DOCUMENT.encode("utf-8"),
                            q("w", "drawing"), q("m", "oMath"), q("w", "hyperlink"))
    out_counts = count_tags(doc,
                            q("w", "drawing"), q("m", "oMath"), q("w", "hyperlink"))
    check(out_counts[q("w", "drawing")] == src_counts[q("w", "drawing")] == 2,
          "both images preserved (2 w:drawing)")
    check(out_counts[q("m", "oMath")] == 1, "math preserved (1 m:oMath)")
    check(out_counts[q("w", "hyperlink")] == 1, "hyperlink preserved (1 w:hyperlink)")

    text_blob = b"".join(
        (e.text or "").encode("utf-8")
        for e in ET.fromstring(doc).iter(q("w", "t")))
    check("《Hello world》".encode("utf-8") in text_blob,
          "translated text present in output")
    check("Hello world".encode("utf-8") not in text_blob.replace(
              "《Hello world》".encode("utf-8"), b""),
          "original untranslated text replaced")

    # watermark wiring
    check(dt._WM_HEADER_PART in names, "watermark header part added")
    check(b"machine translation" in z_read(out, dt._WM_HEADER_PART),
          "watermark text in new header")
    check(b"machine translation" in hdr1, "watermark injected into existing header1")
    check(b"/word/header_mt_watermark.xml" in ct, "watermark content-type override")
    check(b"rIdMtWatermark" in rels, "watermark relationship registered")
    sect = ET.fromstring(doc).find(".//" + q("w", "sectPr"))
    refs = [hr.get(q("w", "type")) for hr in sect.findall(q("w", "headerReference"))]
    check("default" in refs, "section still references a default header")

    print()
    if fails:
        print("RESULT: %d FAILURE(S)" % len(fails))
        return 1
    print("RESULT: ALL CHECKS PASSED   ->  %s" % out)
    return 0


def z_read(path, part):
    with zipfile.ZipFile(path) as z:
        return z.read(part)


if __name__ == "__main__":
    sys.exit(main())
