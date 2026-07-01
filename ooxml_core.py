# -*- coding: utf-8 -*-
"""
ooxml_core.py  --  shared, pure-stdlib core for the OOXML family translators
(docx / pptx / xlsx).

No third-party packages. Uses only:
    zipfile               -- an OOXML file is a ZIP of XML parts
    xml.etree.ElementTree -- the stdlib XML parser/serializer

Python 3.7+ (x86/x64).

The whole family shares ONE idea:

    * A document is a sequence of *translatable units* (a Word paragraph,
      a PowerPoint paragraph, an Excel shared-string).
    * Each unit is written to exactly one line of a UTF-8 text file, so
      `line count == unit count`.  This is the contract the external
      translator (step 2) relies on.
    * Inside a unit, anything that must NOT be translated (image, math,
      hyperlink, bookmark, break, field ...) is replaced by a literal
      placeholder token  SYM_001, SYM_002, ...  The original elements are
      never written to disk: on rebuild we re-parse the original file and
      walk the units in the same order, so the N-th placeholder of a line
      maps back to the N-th preserved element of that unit.

This module provides the pieces every format reuses; the per-format logic
lives in docx_translator.py / pptx_translator.py / xlsx_translator.py.
"""

import os
import re
import copy
import zipfile
import xml.etree.ElementTree as ET

# --------------------------------------------------------------------------
# Namespaces
# --------------------------------------------------------------------------
# Registering the canonical prefixes keeps the rewritten XML clean (Word/Excel
# read by namespace *URI*, so even an unregistered prefix would be valid -- but
# clean output keeps diffs small and avoids surprises).

NS = {
    # WordprocessingML
    "w":    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14":  "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15":  "http://schemas.microsoft.com/office/word/2012/wordml",
    # DrawingML (shared by docx/pptx/xlsx) + PresentationML / SpreadsheetML
    "a":    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p":    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "x":    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pic":  "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wp":   "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    # Math
    "m":    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    # Relationships / content types / packaging
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel":  "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct":   "http://schemas.openxmlformats.org/package/2006/content-types",
    # Markup compatibility
    "mc":   "http://schemas.openxmlformats.org/markup-compatibility/2006",
    # VML (used by the Word watermark) + office drawing
    "v":    "urn:schemas-microsoft-com:vml",
    "o":    "urn:schemas-microsoft-com:office:office",
    "w10":  "urn:schemas-microsoft-com:office:word",
}
XML_NS = "http://www.w3.org/XML/1998/namespace"     # built-in xml: prefix

for _prefix, _uri in NS.items():
    ET.register_namespace(_prefix, _uri)
# NB: do NOT register an empty default prefix for any OOXML namespace. If w:
# were the default namespace, ElementTree would serialize namespaced ATTRIBUTES
# (e.g. w:type) without their prefix -> a bare `type` attribute lands in *no*
# namespace and Word stops recognizing it. Keeping explicit prefixes is correct.


def q(prefix, local):
    """Clark-notation qualified name, e.g. q('w','p') -> '{...}p'."""
    return "{%s}%s" % (NS[prefix], local)


def local(el_or_tag):
    """
    Local name, accepting either an Element or a Clark-notation tag string.
    Comment / processing-instruction nodes (whose .tag is a callable) yield ''.
        local(child)        -> 'p'
        local('{ns}p')      -> 'p'
    """
    tag = getattr(el_or_tag, "tag", el_or_tag)
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


XML_SPACE = "{%s}space" % XML_NS


# --------------------------------------------------------------------------
# Placeholder tokens
# --------------------------------------------------------------------------
# Format is fixed at 3 zero-padded digits so the splitter can match exactly
# 3 digits and never swallow following text ("SYM_001123" -> SYM_001 + "123").
# That caps a single unit at 999 preserved elements, which is far beyond any
# real paragraph.

SYM_FMT = "SYM_%03d"
SYM_RE = re.compile(r"SYM_(\d{3})")


def sym_token(n):
    """1-based index -> 'SYM_001'."""
    return SYM_FMT % n


def split_syms(line):
    """
    Split a translated line into an ordered list of pieces:
        ('text', str)   -- literal text
        ('sym',  int)   -- 1-based placeholder index
    """
    pieces = []
    last = 0
    for m in SYM_RE.finditer(line):
        if m.start() > last:
            pieces.append(("text", line[last:m.start()]))
        pieces.append(("sym", int(m.group(1))))
        last = m.end()
    if last < len(line):
        pieces.append(("text", line[last:]))
    return pieces


# --------------------------------------------------------------------------
# XML (de)serialization
# --------------------------------------------------------------------------

def parse_xml(data):
    """bytes -> ElementTree root element."""
    return ET.fromstring(data)


def serialize_xml(root):
    """
    Element -> bytes, with the standalone XML declaration OOXML parts use.
    ET.tostring emits the needed xmlns declarations itself (for namespaces
    actually used in the tree, via the registered prefixes).
    """
    body = ET.tostring(root, encoding="unicode")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            + body).encode("utf-8")


# --------------------------------------------------------------------------
# ZIP read / rewrite
# --------------------------------------------------------------------------

def read_part(zip_path, part_name):
    """Read one part from the package as bytes (None if absent)."""
    with zipfile.ZipFile(zip_path, "r") as z:
        try:
            return z.read(part_name)
        except KeyError:
            return None


def list_parts(zip_path):
    with zipfile.ZipFile(zip_path, "r") as z:
        return z.namelist()


def rewrite_zip(src_path, dst_path, replacements=None, additions=None):
    """
    Copy `src_path` to `dst_path`, swapping in `replacements` (name -> bytes)
    and appending `additions` (name -> bytes).  Every other part is copied
    byte-for-byte with its original compression, so the package stays valid
    and minimal-diff.
    """
    replacements = replacements or {}
    additions = additions or {}
    seen = set()
    with zipfile.ZipFile(src_path, "r") as zin, \
            zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            name = info.filename
            data = replacements.get(name)
            if data is None:
                data = zin.read(name)
            zi = zipfile.ZipInfo(name, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zout.writestr(zi, data)
            seen.add(name)
        for name, data in additions.items():
            if name not in seen:
                zout.writestr(name, data)


# --------------------------------------------------------------------------
# Translation text file (the step-1 output / step-3 input)
# --------------------------------------------------------------------------
# One unit per line, '\n' terminated, UTF-8.  We tolerate a single trailing
# newline added by the translator.

def write_lines(txt_path, lines):
    with open(txt_path, "w", encoding="utf-8", newline="") as f:
        for ln in lines:
            f.write(ln)
            f.write("\n")


def read_lines(txt_path, expected_count=None):
    with open(txt_path, "r", encoding="utf-8", newline="") as f:
        content = f.read()
    lines = content.split("\n")
    if lines and lines[-1] == "":          # drop the trailing terminator
        lines.pop()
    if expected_count is not None and len(lines) != expected_count:
        # Don't crash on a translator that changed the line count; warn and
        # let the caller zip by the shorter length.
        import sys
        sys.stderr.write(
            "[ooxml] warning: translation has %d lines but document has %d "
            "units; using min().\n" % (len(lines), expected_count))
    return lines


# --------------------------------------------------------------------------
# Shared paragraph tokenizer for *run-based* formats (docx & pptx)
# --------------------------------------------------------------------------
# docx and pptx both model a paragraph as <p> -> [props, runs, inline objects].
# A run is <r> -> [props, <t> text, or a preserved object].  Only the tag names
# and namespace differ, captured by a small RunSchema.

class RunSchema(object):
    def __init__(self, ns, para, run, text, rpr, props=None, tail=None,
                 set_xml_space=True):
        self.ns = ns
        self.para = q(ns, para)          # paragraph tag, e.g. w:p / a:p
        self.run = q(ns, run)            # run tag,        e.g. w:r / a:r
        self.text = q(ns, text)          # text tag,       e.g. w:t / a:t
        self.rpr = q(ns, rpr)            # run props,      e.g. w:rPr / a:rPr
        self.props_local = props         # paragraph props local name (pPr)
        self.tail_local = tail           # trailing element kept last (a:endParaRPr)
        self.set_xml_space = set_xml_space

    def find_rpr(self, run):
        return run.find(self.rpr)


# A token is a 3-tuple: (kind, payload, ctx)
#   ('text', str,     run_elem)   -- translatable text; ctx run supplies base rPr
#   ('sym',  element, None)       -- preserved paragraph-level child (insert as-is)
#   ('sym',  element, run_elem)   -- preserved run sub-element (wrap in a run on rebuild)

def iter_run_tokens(para, S):
    """Yield tokens for a run-based paragraph element. Shared by docx & pptx."""
    rpr_l = local(S.rpr)
    text_l = local(S.text)
    props_l = S.props_local
    tail_l = S.tail_local
    run_l = local(S.run)
    for child in para:
        ln = local(child)
        if ln == props_l or ln == tail_l:
            continue                      # paragraph props / trailing props kept verbatim
        if ln == run_l:
            for sub in child:
                sln = local(sub)
                if sln == rpr_l:
                    continue
                if sln == text_l:
                    yield ("text", sub.text or "", child)
                else:
                    yield ("sym", sub, child)     # run sub-element -> wrap on rebuild
        else:
            yield ("sym", child, None)            # paragraph-level object -> as-is


def extract_unit_line(para, S):
    """
    Build the text-file line for one paragraph: text with SYM placeholders.
    Returns '' when the paragraph holds no translatable text (so rebuild leaves
    it untouched and any images/math survive automatically).
    """
    parts = []
    n = 0
    has_text = False
    for kind, payload, _ctx in iter_run_tokens(para, S):
        if kind == "text":
            has_text = True
            parts.append(payload)
        else:
            n += 1
            parts.append(sym_token(n))
    return "".join(parts) if has_text else ""


def _make_text_run(S, base_rpr, text):
    r = ET.Element(S.run)
    if base_rpr is not None:
        r.append(copy.deepcopy(base_rpr))
    t = ET.Element(S.text)
    t.text = text
    if S.set_xml_space:
        t.set(XML_SPACE, "preserve")
    r.append(t)
    return r


def rebuild_unit(para, translated_line, S):
    """
    Rewrite one paragraph in place: replace its translatable runs with the
    translated text, splicing preserved elements back where their SYM tokens
    appear.  Paragraph props (and any trailing props) are preserved.
    """
    tokens = list(iter_run_tokens(para, S))
    if not any(k == "text" for k, _p, _c in tokens):
        return                                   # nothing was translated; leave as-is

    # Ordered preserved elements + the run context that supplied each one.
    syms = [(payload, ctx) for kind, payload, ctx in tokens if kind == "sym"]

    # Base formatting = rPr of the first translatable run.
    base_rpr = None
    for kind, _payload, ctx in tokens:
        if kind == "text":
            rpr = S.find_rpr(ctx) if ctx is not None else None
            base_rpr = copy.deepcopy(rpr) if rpr is not None else None
            break

    def materialize_sym(idx):
        elem, ctx = syms[idx - 1]
        if ctx is None:
            return elem                          # already a valid paragraph child
        # run sub-element: wrap in a fresh run carrying the original run's rPr
        nr = ET.Element(S.run)
        rpr = S.find_rpr(ctx)
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        nr.append(elem)
        return nr

    new_children = []
    used = set()
    for kind, val in split_syms(translated_line):
        if kind == "text":
            if val:
                new_children.append(_make_text_run(S, base_rpr, val))
        else:  # sym
            if 1 <= val <= len(syms) and val not in used:
                new_children.append(materialize_sym(val))
                used.add(val)
    # Safety net: never drop a preserved element the translator forgot to echo.
    for i in range(1, len(syms) + 1):
        if i not in used:
            new_children.append(materialize_sym(i))

    # Reassemble: [pPr] + new content + [trailing props].
    props_el = None
    tail_el = None
    for child in list(para):
        ln = local(child)
        if ln == S.props_local and props_el is None:
            props_el = child
        elif ln == S.tail_local and S.tail_local is not None:
            tail_el = child
        para.remove(child)
    if props_el is not None:
        para.append(props_el)
    for c in new_children:
        para.append(c)
    if tail_el is not None:
        para.append(tail_el)


# --------------------------------------------------------------------------
# Path helpers
# --------------------------------------------------------------------------

def resolve_target(base_dir, target):
    """
    Resolve a relationship Target to a package part name.
        resolve_target('ppt', 'slides/slide1.xml') -> 'ppt/slides/slide1.xml'
        resolve_target('ppt', '/customXml/item1.xml') -> 'customXml/item1.xml'
    """
    if target.startswith("/"):
        return target.lstrip("/")
    # normalize any ../ segments relative to base_dir
    parts = (base_dir.split("/") if base_dir else []) + target.split("/")
    stack = []
    for seg in parts:
        if seg in ("", "."):
            continue
        if seg == "..":
            if stack:
                stack.pop()
        else:
            stack.append(seg)
    return "/".join(stack)


# --------------------------------------------------------------------------
# Inline-newline packing (for formats whose unit may contain hard line breaks,
# e.g. an Excel shared string).  A text-file line cannot hold a raw newline, so
# we reversibly encode them.  Backslash is escaped first to stay reversible.
# --------------------------------------------------------------------------

def pack_newlines(s):
    return (s.replace("\\", "\\\\")
             .replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n"))


def unpack_newlines(s):
    out = []
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if c == "\\" and i + 1 < n:
            nxt = s[i + 1]
            if nxt == "n":
                out.append("\n")
                i += 2
                continue
            if nxt == "\\":
                out.append("\\")
                i += 2
                continue
        out.append(c)
        i += 1
    return "".join(out)


def derive_path(source_path, new_suffix, new_ext=None):
    """
    /a/b/source.docx + ('', '.txt')      -> /a/b/source.txt
    /a/b/source.docx + ('_tran', None)   -> /a/b/source_tran.docx
    """
    d = os.path.dirname(source_path)
    base = os.path.basename(source_path)
    stem, ext = os.path.splitext(base)
    if new_ext is not None:
        ext = new_ext
    return os.path.join(d, stem + new_suffix + ext)
