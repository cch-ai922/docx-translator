# -*- coding: utf-8 -*-
"""
html_translator.py  --  pure-stdlib HTML translator (steps 1 & 3).

    convert_html_2_text(source_html_path)        -> writes  <stem>.txt
    convert_txt_2_html(source_html_path,         -> writes  <stem>_tran.html
                       source_trantext_path)         (+ optional watermark)

Same extract -> translate -> rebuild contract as the OOXML translators, but the
container/parser differ:

* HTML is a single text file and is frequently NOT well-formed XML (unclosed
  <br>/<img>, bare &, unquoted attributes), so we use the stdlib tolerant
  `html.parser.HTMLParser` rather than xml.etree.
* The translatable UNIT (one line) is a *block* element's text: p, div, li,
  h1-h6, td, title, option, ... Concatenated visible text inside the block.
* *Inline* elements (a, b, i, span, img, br, code, sub, sup, ...) and comments
  become SYM_00n placeholders -- exactly the role runs/drawings play in OOXML.
* <script>/<style> bodies, comments, doctype and processing instructions are
  echoed back verbatim.

Determinism: nothing is written to a side-file. Rebuild re-parses the original
HTML with the identical tokenizer, so the N-th SYM of a translated line maps
positionally back to the N-th preserved inline element of that unit.
"""

import re
import html as _html
from html.parser import HTMLParser

from ooxml_core import (
    sym_token, split_syms, write_lines, read_lines, derive_path,
    pack_newlines, unpack_newlines,
)

# Elements that delimit a translatable unit (everything else is "inline" and is
# preserved as a SYM token inside the surrounding unit).
BLOCK_ELEMENTS = frozenset("""
html head body title p div
h1 h2 h3 h4 h5 h6
ul ol li dl dt dd menu
table thead tbody tfoot tr td th caption colgroup col
blockquote pre hr address
section article aside header footer nav main figure figcaption
form fieldset legend details summary dialog
option optgroup textarea
script style noscript template
""".split())

# Elements whose body is CDATA (never translated, echoed verbatim).
CDATA_ELEMENTS = frozenset(("script", "style"))


# --------------------------------------------------------------------------
# Tokenizer
# --------------------------------------------------------------------------

def _fix_meta_charset(raw):
    """Rewrite a <meta> tag's declared charset to utf-8 (output is utf-8)."""
    return re.sub(r'(charset\s*=\s*["\']?\s*)([\w\-]+)',
                  lambda m: m.group(1) + "utf-8", raw, flags=re.I)


class _Tokenizer(HTMLParser):
    def __init__(self):
        # convert_charrefs=True -> entity/char refs in text arrive pre-decoded
        # in handle_data (but NOT inside script/style CDATA, which stays raw).
        HTMLParser.__init__(self, convert_charrefs=True)
        self.tokens = []
        self._cdata_depth = 0

    def handle_starttag(self, tag, attrs):
        raw = self.get_starttag_text()
        name = tag.lower()
        if name == "meta":
            raw = _fix_meta_charset(raw)
        self.tokens.append(("start", name, raw))
        if name in CDATA_ELEMENTS:
            self._cdata_depth += 1

    def handle_startendtag(self, tag, attrs):
        raw = self.get_starttag_text()
        name = tag.lower()
        if name == "meta":
            raw = _fix_meta_charset(raw)
        self.tokens.append(("startend", name, raw))

    def handle_endtag(self, tag):
        name = tag.lower()
        self.tokens.append(("end", name, "</%s>" % name))
        if name in CDATA_ELEMENTS and self._cdata_depth > 0:
            self._cdata_depth -= 1

    def handle_data(self, data):
        self.tokens.append(("rawdata" if self._cdata_depth > 0 else "data", data))

    def handle_comment(self, data):
        self.tokens.append(("raw", "<!--%s-->" % data))

    def handle_decl(self, decl):
        self.tokens.append(("raw", "<!%s>" % decl))

    def handle_pi(self, data):
        self.tokens.append(("raw", "<?%s>" % data))

    def unknown_decl(self, data):
        self.tokens.append(("raw", "<![%s]>" % data))


# --------------------------------------------------------------------------
# Tokens -> events
# --------------------------------------------------------------------------
# An event is one of:
#   ('raw',  string)                 -- echo verbatim
#   ('unit', text_with_syms, [raw])  -- one translatable line + preserved inline
#                                       markup in order

def _build_events(tokens):
    events = []
    buf = []                 # list of ('text', decoded) | ('sym', raw_markup)
    st = {"has_text": False}

    def flush():
        if not buf:
            return
        if st["has_text"]:
            parts, preserved, n = [], [], 0
            for kind, val in buf:
                if kind == "text":
                    parts.append(val)
                else:
                    n += 1
                    parts.append(sym_token(n))
                    preserved.append(val)
            # a block's text may contain hard newlines (source wrapping); pack
            # them so the unit stays on a single text-file line.
            events.append(("unit", pack_newlines("".join(parts)), preserved))
        else:
            # no translatable text -> echo the buffer verbatim (standalone
            # comments, whitespace, image-only blocks, ...)
            s = "".join(_html.escape(v, quote=False) if k == "text" else v
                        for k, v in buf)
            events.append(("raw", s))
        del buf[:]
        st["has_text"] = False

    for tok in tokens:
        ttype = tok[0]
        if ttype == "data":
            buf.append(("text", tok[1]))
            if tok[1].strip():
                st["has_text"] = True
        elif ttype == "rawdata":
            flush()
            events.append(("raw", tok[1]))
        elif ttype in ("start", "end", "startend"):
            name, raw = tok[1], tok[2]
            if name in BLOCK_ELEMENTS:
                flush()
                events.append(("raw", raw))
            else:
                buf.append(("sym", raw))            # inline element -> placeholder
        elif ttype == "raw":
            buf.append(("sym", tok[1]))             # comment/decl/pi -> placeholder
    flush()
    return events


# --------------------------------------------------------------------------
# Reading (encoding detection)
# --------------------------------------------------------------------------

def _read_html(path):
    with open(path, "rb") as f:
        raw = f.read()
    enc = "utf-8"
    if raw.startswith(b"\xef\xbb\xbf"):
        enc = "utf-8-sig"
    elif raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        enc = "utf-16"
    else:
        m = re.search(rb'charset\s*=\s*["\']?\s*([\w\-]+)', raw[:4096], re.I)
        if m:
            try:
                enc = m.group(1).decode("ascii").strip()
            except Exception:
                enc = "utf-8"
    try:
        return raw.decode(enc, errors="replace")
    except LookupError:
        return raw.decode("utf-8", errors="replace")


def _tokenize(path):
    parser = _Tokenizer()
    parser.feed(_read_html(path))
    parser.close()
    return parser.tokens


# --------------------------------------------------------------------------
# Step 1:  html -> txt
# --------------------------------------------------------------------------

def convert_html_2_text(source_html_path):
    events = _build_events(_tokenize(source_html_path))
    lines = [ev[1] for ev in events if ev[0] == "unit"]
    out_txt = derive_path(source_html_path, "", ".txt")
    write_lines(out_txt, lines)
    return out_txt


# --------------------------------------------------------------------------
# Step 3:  html + translated txt -> translated html
# --------------------------------------------------------------------------

def convert_txt_2_html(source_html_path, source_trantext_path,
                       watermark_text="machine translation"):
    events = _build_events(_tokenize(source_html_path))
    n_units = sum(1 for ev in events if ev[0] == "unit")
    lines = read_lines(source_trantext_path, expected_count=n_units)

    out = []
    li = 0
    for ev in events:
        if ev[0] == "raw":
            out.append(ev[1])
            continue
        preserved = ev[2]
        line = lines[li] if li < len(lines) else ev[1]
        li += 1
        used = set()
        for kind, val in split_syms(unpack_newlines(line)):
            if kind == "text":
                out.append(_html.escape(val, quote=False))
            elif 1 <= val <= len(preserved):
                out.append(preserved[val - 1])
                used.add(val)
        for i in range(1, len(preserved) + 1):      # safety net: never lose markup
            if i not in used:
                out.append(preserved[i - 1])
    html_out = "".join(out)

    if watermark_text:
        html_out = _inject_watermark(html_out, watermark_text)

    out_html = derive_path(source_html_path, "_tran", ".html")
    with open(out_html, "w", encoding="utf-8", newline="") as f:
        f.write(html_out)
    return out_html


def _inject_watermark(html_out, text):
    overlay = (
        '<div style="position:fixed;top:50%%;left:50%%;'
        'transform:translate(-50%%,-50%%) rotate(-45deg);'
        'font-size:64px;color:#d9d9d9;opacity:.5;pointer-events:none;'
        'z-index:2147483647;white-space:nowrap;font-family:sans-serif">%s</div>'
        % _html.escape(text)
    )
    idx = html_out.lower().rfind("</body>")
    if idx != -1:
        return html_out[:idx] + overlay + html_out[idx:]
    return html_out + overlay


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print("usage:\n"
              "  python html_translator.py extract  source.html\n"
              "  python html_translator.py rebuild  source.html  source_tran.txt")
        return 0
    if argv[0] == "extract" and len(argv) == 2:
        print(convert_html_2_text(argv[1]))
    elif argv[0] == "rebuild" and len(argv) == 3:
        print(convert_txt_2_html(argv[1], argv[2]))
    else:
        print("bad arguments; use --help")
        return 2
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_main(sys.argv[1:]))
