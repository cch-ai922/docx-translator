# -*- coding: utf-8 -*-
"""Round-trip test for html_translator. Pure stdlib."""
import os
import sys
import tempfile

import html_translator as ht

SAMPLE = """<!DOCTYPE html>
<html>
<head>
  <meta charset="iso-8859-1">
  <title>My Page</title>
  <style>.a{color:red & blue}</style>
  <script>if (a < b && c > d) alert("<p>not html</p>");</script>
</head>
<body>
  <h1>Hello &amp; welcome</h1>
  <p>This is <b>bold</b> and a <a href="http://x.com?a=1&b=2">link</a>.<br>Second line.</p>
  <ul>
    <li>First item</li>
    <li>Second item</li>
  </ul>
  <p><img src="pic.png" alt="x"></p>
  <!-- a comment -->
  <table><tr><td>Cell A</td><td>Cell B</td></tr></table>
</body>
</html>
"""


def read_lines(p):
    ls = open(p, encoding="utf-8").read().split("\n")
    if ls and ls[-1] == "":
        ls.pop()
    return ls


def main():
    tmp = tempfile.mkdtemp(prefix="htmltest_")
    src = os.path.join(tmp, "source.html")
    # write as latin-1 to exercise encoding detection from <meta charset>
    with open(src, "w", encoding="iso-8859-1", newline="") as f:
        f.write(SAMPLE)

    fails = []
    def check(cond, msg):
        print(("  ok  " if cond else " FAIL ") + msg)
        if not cond:
            fails.append(msg)

    # ---- step 1 --------------------------------------------------------
    txt = ht.convert_html_2_text(src)
    lines = read_lines(txt)
    print("Extracted lines:")
    for i, ln in enumerate(lines):
        print("   [%d] %r" % (i, ln))

    check(lines == [
        "My Page",
        "Hello & welcome",
        "This is SYM_001boldSYM_002 and a SYM_003linkSYM_004.SYM_005Second line.",
        "First item",
        "Second item",
        "Cell A",
        "Cell B",
    ], "units + inline SYM tokens extracted correctly")
    check(all("not html" not in ln for ln in lines), "script body NOT extracted")
    check(all("color:red" not in ln for ln in lines), "style body NOT extracted")
    check("<img" not in "".join(lines), "image-only <p> produced no unit (left as-is)")

    # ---- step 2 (fake): wrap each line, keep SYM tokens ----------------
    tran = txt[:-4] + "_tran.txt"
    with open(tran, "w", encoding="utf-8", newline="") as f:
        for ln in lines:
            f.write(("[%s]" % ln) if ln else "")
            f.write("\n")

    # ---- step 3 --------------------------------------------------------
    out = ht.convert_txt_2_html(src, tran)
    result = open(out, encoding="utf-8").read()
    print("\n--- rebuilt (excerpt) ---")
    print(result[result.find("<body"):result.find("</ul>") + 5])

    check("[This is " in result, "translated paragraph text present")
    check("<b>bold</b>" in result, "inline <b>bold</b> markup preserved in place")
    check('<a href="http://x.com?a=1&b=2">link</a>' in result,
          "hyperlink markup + href preserved exactly")
    check("<br>" in result, "<br> preserved")
    check('<img src="pic.png" alt="x">' in result, "image-only block preserved")
    check('if (a < b && c > d)' in result, "script body untouched (not escaped)")
    check(".a{color:red & blue}" in result, "style body untouched")
    check("<!-- a comment -->" in result, "comment preserved")
    check('charset="utf-8"' in result, "meta charset rewritten to utf-8")
    check("machine translation" in result and "position:fixed" in result,
          "watermark overlay injected")

    # re-tokenizing the OUTPUT must keep the same inline-element counts
    def count_tag(s, tag):
        return len(__import__("re").findall(r"<%s\b" % tag, s))
    check(count_tag(result, "b") == 1 and count_tag(result, "a") == 1
          and count_tag(result, "img") == 1, "inline element counts unchanged")

    # regression: a block whose text contains a hard newline must stay one line
    ml_src = os.path.join(tmp, "ml.html")
    with open(ml_src, "w", encoding="utf-8", newline="") as f:
        f.write("<html><body><p>alpha\nbeta\ngamma</p><p>solo</p></body></html>")
    ml_txt = ht.convert_html_2_text(ml_src)
    ml_lines = read_lines(ml_txt)
    check(len(ml_lines) == 2, "multi-line block stays one text-file line (2 units)")
    with open(ml_txt[:-4] + "_tran.txt", "w", encoding="utf-8", newline="") as f:
        for ln in ml_lines:
            f.write(("X" + ln) if ln else "")
            f.write("\n")
    ml_out = ht.convert_txt_2_html(ml_src, ml_txt[:-4] + "_tran.txt", watermark_text=None)
    ml_res = open(ml_out, encoding="utf-8").read()
    check("alpha\nbeta\ngamma" in ml_res, "hard newlines restored on rebuild")

    print()
    if fails:
        print("RESULT: %d FAILURE(S)" % len(fails)); return 1
    print("RESULT: ALL CHECKS PASSED  ->  %s" % out); return 0


if __name__ == "__main__":
    sys.exit(main())
