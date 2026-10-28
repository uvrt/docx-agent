#!/usr/bin/env python3
"""Measure which paragraphs Word lists in a table of contents, and at which level, for
``TOC``'s ``\\o``, ``\\t`` and ``\\u`` alone and combined (production round 21, T1).

Made-up documents: "Contents", a ``TOC`` field whose cached result is ``?``, then a body of
paragraphs in made-up styles -- built-in headings, custom styles with and without an
outline level, paragraphs stating their own (``w:outlineLvl``), empty ones, a heading in a
table cell, a built-in heading whose definition lost (or changed) its outline level.  One
copy per instruction.  Word -- through ``tests/oracle.py``'s machine-wide lock -- opens each,
updates its fields (``update field``) and saves it; the entries it wrote (TOC level, text)
are read back into ``tests/observations/toc-word.json``.

Word splits ``\\t``'s list on its locale's list separator: this Mac's Word is Dutch (``;``),
so most cases are measured there; the ``en-US`` ones are Word started with
``-AppleLocale en_US -AppleLanguages (en-US)`` (``open --args``), where the separator is
``,`` and the built-in styles' display names English.

    python tools/toc_probe.py [case ...]

Nothing Word writes is committed: only these facts.
"""

from __future__ import annotations

import io
import json
import subprocess
import sys
import time
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "src"))

OBSERVED = ROOT / "tests" / "observations" / "toc-word.json"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
NS = {"w": W}

#: The made-up bodies: (text, style name or None for Normal, own w:outlineLvl or None).
BODIES = {
    "base": [
        ("Cover line", "heading 2", None),
        ("Chapter alpha", "Section title", None),
        ("Sub point alpha", "heading 2", None),
        ("", "Section title", None),
        ("", "heading 1", None),
        ("Own level one", None, 0),
        ("Heading one", "heading 1", None),
        ("Heading three", "heading 3", None),
        ("Leveled para", "Leveled", None),
        ("Leveled child para", "Leveled child", None),
        ("Heading one body level", "heading 1", 9),
        ("Report name", "Report title", None),
        ("Section own level three", "Section title", 2),
        ("Plain body", None, None),
        ("Chapter beta", "Section title", None),
    ],
    "outline": [
        ("Heading one own three", "heading 1", 2),
        ("Leveled own one", "Leveled", 0),
        ("Leveled own body", "Leveled", 9),
        ("Derived heading own three", "Derived heading", 2),
        ("Derived heading plain", "Derived heading", None),
        ("Heading four no level", "heading 4", None),
        ("Heading five level eight", "heading 5", None),
        ("Own level nine", None, 8),
        ("Cell heading two", "heading 2", None),      # in a one-cell table
        ("Heading two", "heading 2", None),
        ("Section own level two", "Section title", 1),
        ("Normal own body", None, 9),
    ],
}
#: Per body: style id -> the outline level its definition states (None: none stated).
STYLE_LEVELS = {"base": {"Leveled": 1}, "outline": {"Leveled": 1, "Heading4": None, "Heading5": 7}}
CELL = "Cell heading two"

_T = ' TOC \\h \\z \\t "{}" '
#: case -> (body, instruction, locale).
CASES = {
    # \t alone, its separator, levels and names (Dutch Word: ";").
    "t": ("base", _T.format("Section title;1"), "nl-NL"),
    "t-comma-nl": ("base", _T.format("Section title,1"), "nl-NL"),
    "t-pairs": ("base", _T.format("Section title;1;Report title;2;Leveled;3"), "nl-NL"),
    "t-spaces": ("base", _T.format("Section title; 1; Report title; 2"), "nl-NL"),
    "t-padded": ("base", _T.format(" Section title ; 2 "), "nl-NL"),
    "t-mixed": ("base", _T.format("Section title;1,Report title;2"), "nl-NL"),
    "t-case": ("base", _T.format("SECTION title;1"), "nl-NL"),
    "t-style-id": ("base", _T.format("Sectiontitle;1"), "nl-NL"),
    "t-no-level": ("base", _T.format("Section title"), "nl-NL"),
    "t-odd": ("base", _T.format("Section title;1;Report title"), "nl-NL"),
    "t-missing": ("base", _T.format("No such style;1;Report title;2"), "nl-NL"),
    "t-level-5": ("base", _T.format("Section title;5"), "nl-NL"),
    "t-level-0": ("base", _T.format("Section title;0;Report title;2"), "nl-NL"),
    "t-level-10": ("base", _T.format("Section title;10;Report title;2"), "nl-NL"),
    "t-level-x": ("base", _T.format("Section title;x;Report title;2"), "nl-NL"),
    "t-twice": ("base", ' TOC \\h \\z \\t "Section title;1" \\t "Report title;2" ', "nl-NL"),
    "t-duplicate": ("base", _T.format("Section title;1;Section title;3"), "nl-NL"),
    "t-leveled": ("base", _T.format("Leveled;3"), "nl-NL"),
    "t-heading-en-name-nl": ("base", _T.format("Heading 2;1"), "nl-NL"),
    "t-heading-nl-name-nl": ("outline", _T.format("Kop 2;1"), "nl-NL"),
    "t-heading-id": ("outline", _T.format("Heading2;1"), "nl-NL"),
    "t-overrides-style-level": ("outline", _T.format("Leveled;5;Heading 2;6"), "nl-NL"),
    # \o alone and with \t.
    "none": ("base", " TOC \\h \\z ", "nl-NL"),
    "o": ("base", " TOC \\o \\h \\z ", "nl-NL"),
    "o12": ("base", ' TOC \\o "1-2" \\h \\z ', "nl-NL"),
    "o13": ("base", ' TOC \\o "1-3" \\h \\z ', "nl-NL"),
    "o19": ("base", ' TOC \\o "1-9" \\h \\z ', "nl-NL"),
    "o13-outline": ("outline", ' TOC \\o "1-3" \\h \\z ', "nl-NL"),
    "o19-outline": ("outline", ' TOC \\o "1-9" \\h \\z ', "nl-NL"),
    "o12-t": ("base", ' TOC \\o "1-2" \\h \\z \\t "Section title;1" ', "nl-NL"),
    "o13-t-leveled": ("base", ' TOC \\o "1-3" \\h \\z \\t "Leveled;3" ', "nl-NL"),
    "o22-t": ("base", ' TOC \\o "2-2" \\h \\z \\t "Report title;1" ', "nl-NL"),
    "o-t": ("base", ' TOC \\o \\h \\z \\t "Report title;2" ', "nl-NL"),
    "o12-t-outline": ("outline", ' TOC \\o "1-2" \\h \\z \\t "Leveled;5" ', "nl-NL"),
    # \u alone and combined.
    "u": ("base", " TOC \\h \\z \\u ", "nl-NL"),
    "u-outline": ("outline", " TOC \\h \\z \\u ", "nl-NL"),
    "o12-u": ("base", ' TOC \\o "1-2" \\h \\z \\u ', "nl-NL"),
    "o13-u": ("base", ' TOC \\o "1-3" \\h \\z \\u ', "nl-NL"),
    "o19-u-outline": ("outline", ' TOC \\o "1-9" \\h \\z \\u ', "nl-NL"),
    "u-t": ("base", ' TOC \\h \\z \\u \\t "Report title;2" ', "nl-NL"),
    "t-u": ("base", ' TOC \\h \\z \\t "Report title;2" \\u ', "nl-NL"),
    "u-t-own": ("base", ' TOC \\h \\z \\u \\t "Section title;1" ', "nl-NL"),
    "o19-u-t": ("base", ' TOC \\o "1-9" \\h \\z \\u \\t "Report title;2" ', "nl-NL"),
    "o19-u-t-outline": ("outline", ' TOC \\o "1-9" \\h \\z \\u \\t "Section title;4" ', "nl-NL"),
    # English Word: "," separates, English display names.
    "t-en": ("base", _T.format("Section title,1"), "en-US"),
    "t-pairs-en": ("base", _T.format("Section title,1,Report title,2,Leveled,3"), "en-US"),
    "t-semicolon-en": ("base", _T.format("Section title;1;Report title;2"), "en-US"),
    "t-heading-en": ("base", _T.format("Heading 2,1"), "en-US"),
    "o13-t-heading-en": ("base", ' TOC \\o "1-3" \\h \\z \\t "Heading 2,1" ', "en-US"),
    "o12-t-en": ("base", ' TOC \\o "1-2" \\h \\z \\t "Section title,1" ', "en-US"),
    "u-en": ("base", " TOC \\h \\z \\u ", "en-US"),
}

SCRIPT = '''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file inputPath)
      repeat 60 times
        if (count of documents) > 0 then exit repeat
        delay 0.5
      end repeat
      set d to active document
      set n to count of fields of d
      repeat with i from 1 to n
        try
          update field (field i of d)
        end try
      end repeat
      save as d file name outputPath file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
end run
'''


def make(body: str, instruction: str) -> bytes:
    """The probe document of ``body`` with a TOC field of ``instruction``."""
    from docx_agent import Document

    document = Document.new(created="2026-10-28T09:00:00Z")
    for name, based in (("Section title", "Normal"), ("Report title", "Normal"), ("Leveled", "Normal"),
                        ("Leveled child", "Leveled"), ("Derived heading", "heading 1")):
        document.styles.add(name, based_on=based)
    first = document.paragraphs()[0]
    first.set_text("Contents")
    last = first.id
    for text, style, _ in BODIES[body]:
        made = document.paragraph(last).insert_after(text)
        document.paragraph(made.id).set_style(style or "Normal")
        last = made.id
    return _patch(document.to_bytes(), body, instruction)


def _patch(data: bytes, body: str, instruction: str) -> bytes:
    source = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            raw = source.read(item.filename)
            if item.filename == "word/styles.xml":
                root = etree.fromstring(raw)
                for style in root.findall("w:style", NS):
                    style_id = style.get(_W + "styleId")
                    if style_id not in STYLE_LEVELS[body]:
                        continue
                    properties = style.find("w:pPr", NS)
                    if properties is None:
                        properties = etree.Element(_W + "pPr")
                        run = style.find("w:rPr", NS)
                        (run.addprevious if run is not None else style.append)(properties)
                    for old in properties.findall("w:outlineLvl", NS):
                        properties.remove(old)
                    if STYLE_LEVELS[body][style_id] is not None:
                        etree.SubElement(properties, _W + "outlineLvl").set(_W + "val",
                                                                            str(STYLE_LEVELS[body][style_id]))
                raw = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "word/document.xml":
                root = etree.fromstring(raw)
                paragraphs = root.find("w:body", NS).findall("w:p", NS)
                for (text, _, own), paragraph in zip(BODIES[body], paragraphs[1:]):
                    if own is not None:
                        properties = paragraph.find("w:pPr", NS)
                        if properties is None:
                            properties = etree.Element(_W + "pPr")
                            paragraph.insert(0, properties)
                        etree.SubElement(properties, _W + "outlineLvl").set(_W + "val", str(own))
                        mark = properties.find("w:rPr", NS)
                        if mark is not None:
                            properties.append(mark)
                    if text == CELL:
                        _into_cell(paragraph)
                toc = etree.Element(_W + "p")
                for kind in ("begin", None, "separate", "text", "end"):
                    run = etree.SubElement(toc, _W + "r")
                    if kind is None:
                        node = etree.SubElement(run, _W + "instrText")
                        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        node.text = instruction
                    elif kind == "text":
                        etree.SubElement(run, _W + "t").text = "?"
                    else:
                        etree.SubElement(run, _W + "fldChar").set(_W + "fldCharType", kind)
                paragraphs[0].addnext(toc)
                raw = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            target.writestr(item, raw)
    return out.getvalue()


def _into_cell(paragraph) -> None:
    table = etree.Element(_W + "tbl")
    properties = etree.SubElement(table, _W + "tblPr")
    etree.SubElement(properties, _W + "tblW").attrib.update({_W + "w": "0", _W + "type": "auto"})
    etree.SubElement(etree.SubElement(table, _W + "tblGrid"), _W + "gridCol").set(_W + "w", "4000")
    cell = etree.SubElement(etree.SubElement(table, _W + "tr"), _W + "tc")
    etree.SubElement(etree.SubElement(cell, _W + "tcPr"), _W + "tcW").attrib.update(
        {_W + "w": "4000", _W + "type": "dxa"})
    paragraph.addprevious(table)
    cell.append(paragraph)


def entries(data: bytes) -> list[list]:
    """The TOC entries of a saved document: ``[level, text]`` (the page number's digits
    stripped from the end), in order."""
    archive = zipfile.ZipFile(io.BytesIO(data))
    root = etree.fromstring(archive.read("word/document.xml"))
    styles = etree.fromstring(archive.read("word/styles.xml"))
    names = {s.get(_W + "styleId"): s.find("w:name", NS).get(_W + "val") for s in styles.findall("w:style", NS)}
    out = []
    for paragraph in root.iter(_W + "p"):
        style = paragraph.find("w:pPr/w:pStyle", NS)
        name = (names.get(style.get(_W + "val")) or "") if style is not None else ""
        if name.lower().startswith("toc "):
            text = "".join(t.text or "" for t in paragraph.iter(_W + "t") if not _in_pageref(t))
            out.append([int(name[4:]), text])
    return out


def _in_pageref(node) -> bool:
    """Whether a ``w:t`` is in an entry's ``PAGEREF`` field (its page number)."""
    paragraph = node.getparent()
    while paragraph is not None and paragraph.tag != _W + "p":
        paragraph = paragraph.getparent()
    open_fields: list[str] = []
    for item in paragraph.iter(_W + "fldChar", _W + "instrText", _W + "t"):
        if item.tag == _W + "fldChar":
            kind = item.get(_W + "fldCharType")
            if kind == "begin":
                open_fields.append("")
            elif kind == "end" and open_fields:
                open_fields.pop()
        elif item.tag == _W + "instrText" and open_fields:
            open_fields[-1] += item.text or ""
        elif item is node:
            return any("PAGEREF" in field for field in open_fields)
    return False


def _run_english(word, data: bytes, case: str) -> bytes | None:
    """Word started in English (US) for one case: ``open --args -AppleLocale``."""
    import oracle

    staged = oracle.ORACLE_DIR / f"toc-{case}-in.docx"
    out = oracle.ORACLE_DIR / f"toc-{case}-out.docx"
    staged.write_bytes(data)
    out.unlink(missing_ok=True)
    subprocess.run(["open", "-g", "-a", str(oracle.WORD_APP), "--args", "-AppleLocale", "en_US",
                    "-AppleLanguages", "(en-US)"], check=True)
    time.sleep(10)
    subprocess.run(["osascript", "-e", SCRIPT, str(staged), str(out)], capture_output=True, timeout=180)
    staged.unlink(missing_ok=True)
    saved = out.read_bytes() if out.exists() else None
    out.unlink(missing_ok=True)
    word.recover()
    return saved


def main(wanted: list[str]) -> None:
    import oracle

    observed = json.loads(OBSERVED.read_text()) if OBSERVED.exists() else {}
    with oracle.session() as word:
        for case in wanted or list(CASES):
            body, instruction, locale = CASES[case]
            data = make(body, instruction)
            if locale == "en-US":
                saved = _run_english(word, data, case)
            else:
                got = word.run_script(SCRIPT, data, name=f"toc-{case}", tag="updated")
                saved = got.path.read_bytes() if got.ok else None
                if got.ok:
                    got.path.unlink()
            if saved is None:
                print(case, "Word gave nothing back")
                continue
            observed[case] = {"body": body, "instruction": instruction, "locale": locale,
                              "entries": entries(saved)}
            print(case, observed[case]["entries"])
    OBSERVED.write_text(json.dumps(observed, indent=1, ensure_ascii=False) + "\n")


if __name__ == "__main__":
    main(sys.argv[1:])
