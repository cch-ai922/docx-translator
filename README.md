# OOXML Translator (pure Python, no third-party packages)

A 3-step document translation toolkit. Steps **1** and **3** are implemented
here; **step 2** (the actual translation of a plain-text file) is left to your
external engine.

```
source.docx ──①extract──▶ source.txt ──②translate──▶ source_tran.txt ──③rebuild──▶ source_tran.docx
            (this tool)                (your engine)                   (this tool)   (+ watermark)
```

* **Pure standard library** — `zipfile` + `xml.etree.ElementTree`. Nothing to
  `pip install`. CPython **3.7+**, x86/x64.
* Same engine for **Word (.docx)**, **PowerPoint (.pptx)** and **Excel (.xlsx)**.

> ### Why not `lxml`?
> `lxml` is a third-party C extension (`pip install lxml`), so it violates the
> "no third-party packages" rule. An OOXML file is just a **ZIP of XML parts**,
> so `zipfile` + the stdlib `xml.etree.ElementTree` cover everything we need —
> with zero install footprint.

---

## Files

| File | Role |
|------|------|
| `ooxml_core.py` | Shared core: ZIP I/O, XML (de)serialization, the `SYM_00n` placeholder machinery, the text-file contract, and the run-based paragraph tokenizer reused by docx/pptx. |
| `docx_translator.py` | `convert_docx_2_text`, `convert_txt_2_docx` (+ every-page watermark). |
| `pptx_translator.py` | `convert_pptx_2_text`, `convert_txt_2_pptx` (+ per-slide watermark). |
| `xlsx_translator.py` | `convert_xlsx_2_text`, `convert_txt_2_xlsx` (shared strings; optional sheet-header watermark). |
| `html_translator.py` | `convert_html_2_text`, `convert_txt_2_html` (stdlib `HTMLParser`; + optional watermark overlay). |
| `chm_reader.py` | Pure-stdlib CHM decompiler: ITSF/ITSP/PMGL parser **and a from-scratch LZX decompressor**. `CHMFile(path).extract_all(dir)`. |
| `chm_translator.py` | Orchestrates CHM: decompile → per-page HTML translation → emits an `.hhp` project for recompiling with `hhc.exe`. |
| `translate.py` | Dispatcher CLI that picks the engine by file extension (.docx/.pptx/.xlsx/.html). |
| `selftest_docx.py`, `selftest_pptx_xlsx.py`, `selftest_html.py`, `selftest_chm.py` | Self-contained tests. The first three synthesize real files; `selftest_chm.py` validates the LZX decompressor against real Windows `.chm` help files. |

---

## Usage

### As a library

```python
from docx_translator import convert_docx_2_text, convert_txt_2_docx

txt = convert_docx_2_text("source.docx")          # -> source.txt  (step 1)
#   ... step 2: translate source.txt -> source_tran.txt ...
out = convert_txt_2_docx("source.docx", "source_tran.txt")   # -> source_tran.docx (step 3)
```

`pptx_translator` and `xlsx_translator` expose the analogous
`convert_pptx_2_text` / `convert_txt_2_pptx` and
`convert_xlsx_2_text` / `convert_txt_2_xlsx`.

### From the command line

```bash
python translate.py extract source.docx                 # writes source.txt
python translate.py rebuild source.docx source_tran.txt # writes source_tran.docx
```

### Tests

```bash
python selftest_docx.py
python selftest_pptx_xlsx.py
```

---

## How the "extract only translatable text" principle is implemented

For every paragraph (the translatable **unit**) we walk its content in document
order and classify each child:

* **`<w:t>` text** → extracted as the words to translate.
* **anything else** — inline image (`w:drawing`/`w:pict`/`w:object`), math
  (`m:oMath`), hyperlink, bookmark, break, tab, field — is replaced by a literal
  placeholder token **`SYM_001`, `SYM_002`, …** (reset per line).

One unit → one line, so **line count == paragraph count** (the contract step 2
relies on). A paragraph with *no* translatable text yields an **empty line** and
is left completely untouched on rebuild, so a picture-only paragraph survives
for free.

### No side-file, no overhead

The preserved elements are **never serialized to disk.** On rebuild we re-open
the *original* document and walk the units in the identical order, so the N-th
`SYM` token of a translated line maps positionally back to the N-th preserved
element of that unit. This is the cheapest possible scheme: a single extra
integer counter during extraction, and a single regex split during rebuild.

```
extract:  "Hello "·<img>·" world"   ->  line:  "Hello SYM_001 world"
translate:                              line:  "Bonjour SYM_001 monde"
rebuild:  re-read original, splice the original <img> where SYM_001 sits
          ->  run("Bonjour ") · <img> · run(" monde")
```

A safety net re-appends any preserved element whose token the translator
dropped, so **content is never lost** even if step 2 mangles a token.

---

## Design decisions & trade-offs (worth knowing)

1. **Placeholder format is fixed at 3 digits** (`SYM_%03d`) and matched as
   exactly 3 digits, so a token can never swallow following text
   (`"SYM_001123"` → `SYM_001` + `"123"`). Cap: 999 preserved elements per
   paragraph — far beyond anything real. Your translator must pass these tokens
   through verbatim (they are short, all-caps + digits; most engines leave them
   alone — keep them on a glossary/do-not-translate list if available).

2. **Intra-paragraph formatting is flattened.** Translated text goes into one
   run carrying the first run's `<w:rPr>`. A *bold word in the middle* of a
   sentence won't keep its bold, because machine translation reorders words and
   the original character spans no longer apply. This is the standard, robust
   trade-off; paragraph-level styles, and the formatting of preserved elements
   themselves, are kept.

3. **Hyperlinks are preserved whole** (URL **and** display text), per the spec —
   the link becomes a single `SYM` token, so its visible text stays in the
   source language. If you'd rather translate link text, change the
   `hyperlink` branch in `ooxml_core.iter_run_tokens` to recurse into the
   hyperlink's runs.

4. **Tables and block content controls are translated** (paragraphs inside them
   are collected in document order). Text boxes / SmartArt / tracked-changes /
   inline content controls are **preserved untranslated** (kept intact rather
   than risk corrupting the structure).

5. **Watermark.** Word stores a watermark as a VML WordArt shape in the section
   **header**. We (a) inject the "machine translation" shape into every header
   the document already uses, and (b) reference a new watermark-only header from
   any section/header-type that lacks one — so it shows on every page. PowerPoint
   has no watermark concept, so we add a centered, semi-transparent text box to
   **every slide**. Excel has no page watermark at all; an optional centered
   **sheet header** (`watermark_text=...`, off by default) is the closest analog.

---

## HTML and CHM

The **extract → translate → rebuild** pattern and the `SYM` placeholder idea
carry over directly. The container and parser are what change.

### HTML — implemented (`html_translator.py`), 100% stdlib

A `.html` file is **not** a ZIP and is frequently **not well-formed XML**
(unclosed `<br>`, `<img>`, bare `&`, etc.), so `xml.etree.ElementTree` would
choke. The module uses the stdlib tolerant **`html.parser.HTMLParser`** and
**echoes every token verbatim, rewriting only text nodes**:

* **Translatable unit (a "line")** = a *block* element (`p`, `div`, `li`,
  `h1`–`h6`, `td`, `title`, `option`, …): its concatenated visible text.
* **`SYM` tokens** = *inline* elements (`a`, `b`, `i`, `span`, `img`, `br`,
  `sub`/`sup`, `code`, …) **and** comments. Exactly the role runs/drawings play
  in OOXML. A nice bonus over OOXML: because each `<b>`/`</b>` becomes its own
  token, a bolded *span* inside a sentence can actually survive translation.
* **Echoed verbatim**: `<script>`/`<style>` bodies, comments, doctype, PIs, and
  whitespace/structure between blocks. Attributes are not translated (extension
  point: `alt`/`title`/`placeholder`/`<meta name=description>` could each become
  their own line).
* **Handled**: HTML entities (decoded on extract via `convert_charrefs`,
  re-escaped on rebuild), input **encoding** detection (BOM / `<meta charset>`,
  default UTF-8) with output normalized to UTF-8 (and the `<meta charset>`
  rewritten to match). Optional `position:fixed` watermark overlay before
  `</body>` (repeats per printed page).

```bash
python translate.py extract source.html                 # writes source.txt
python translate.py rebuild source.html source_tran.txt # writes source_tran.html
python selftest_html.py                                 # round-trip test
```

### CHM — decompile implemented (`chm_reader.py`); recompile needs `hhc.exe`

A `.chm` is **not** a ZIP. It is a compiled Microsoft **ITSS/ITSF** container
holding HTML pages plus a TOC (`.hhc`) and index (`.hhk`), with the content
**LZX-compressed**.

* **Decompiling is done and pure-stdlib.** `chm_reader.py` parses the
  ITSF → ITSP → PMGL directory and includes a **from-scratch LZX decompressor**
  (16-bit-LE bitstream, canonical Huffman pre/main/length/aligned trees,
  position-slot offsets, repeated-offset R0/R1/R2, per-frame 16-bit realignment,
  and the ResetTable-driven per-interval reset). Validated against real Windows
  help files — it decodes each to *exactly* the ResetTable's uncompressed length
  with every page well-formed (`python selftest_chm.py`).
  *Not handled*: LZX Intel-E8 call translation (never enabled for HTML content —
  raises rather than emit corrupt bytes).
* **Recompiling** back to `.chm` is the real wall: it needs an **LZX
  *compressor*** + ITSS writer, which nobody ships in pure Python. Microsoft's
  **HTML Help Workshop** (`hhc.exe`) is the standard recompiler.

`chm_translator.py` runs the whole pipeline and even writes the `.hhp` project
for the final step:

```
.chm ─chm_reader─▶ folder of pages ─html_translator (steps 1&3)─▶ translated pages
     └─ chm_translator also emits project.hhp ─▶  hhc.exe project.hhp ─▶ translated .chm
        └─ or just ship the translated HTML folder / a zipped site
```

```bash
python chm_reader.py file.chm                 # list internal files
python chm_reader.py file.chm out_dir/         # decompile to a folder
python chm_translator.py extract file.chm      # decompile + one .txt per page (+ .hhp)
python chm_translator.py rebuild file_chm/     # rebuild translated pages, then run hhc.exe
```

So the entire chain is pure-stdlib **except** the final `.chm` recompile, which
delegates to the external HHW tool because a pure-Python LZX *compressor* is
impractical.
