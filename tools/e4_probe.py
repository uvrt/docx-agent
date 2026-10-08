#!/usr/bin/env python3
"""Measure what Word writes for E4's structure: sections, headers and footers, notes and
fields.

Probe documents are written here (mode 15, the handful of styles ``e1_probe.py`` writes,
no notes parts), and Word -- through ``tests/oracle.py``'s machine-wide lock -- opens each
and, by AppleScript, makes one family of edits and saves:

* ``breaks``: a section break of each kind -- next page at a paragraph's text's end, a
  continuous one at a paragraph's start, an even-page one at the end of a paragraph that
  ends in text;
* ``setup``: in a document of three sections, the second's orientation, margins, gutter,
  columns (count, spacing, separator), vertical alignment, line numbering, a different
  first page, page numbers restarted in another format; the third's odd and even pages;
* ``headers``: default, first and even headers and a footer in a document of three
  sections, the second section's header unlinked and written, the third left linked;
* ``headers2``: page-number fields (``PAGE``, ``NUMPAGES``, ``SECTIONPAGES``) in footers, a
  footer unlinked without being written (Word copies the one before), a header unlinked,
  written and linked again (its part goes);
* ``notes``: footnotes and endnotes, a section's footnote numbering restarted in lower
  Roman, the footnotes beneath the text; ``notes2``: the last section's numbering in
  upper-case letters from 3, a footnote deleted (and endnotes at each section's end, which
  Word's AppleScript refuses to set);
* ``fields``: a caption (``SEQ``), a bookmark with a ``REF`` and a ``PAGEREF`` to it, a
  ``DATE`` and a ``HYPERLINK`` field, a continuous break at a heading's start;
* ``tocfield``, ``tocstyles``: a table of contents over headings on several pages, put in
  by Insert Field (``TOC \\o "1-3" \\h \\z \\u``) and updated -- AppleScript cannot make
  one through ``table of contents`` -- on a page of another text width, and over nine
  levels (for the TOC 1-9 styles);
* ``removal``: a section break deleted (which section's properties the text keeps);
* ``tracked``: the same kinds of edit with Word tracking: a section's margins and columns,
  a section break, a header's text, a footnote, a field, a table of contents.

What it wrote is read back -- each part normalised (rsids, authors and dates replaced) --
into ``tests/observations/e4-word.json``, and the TOC 1-9 styles Word adds with a table of
contents into ``src/docx_agent/edit/word_toc_styles.py``.

    python tools/e4_probe.py

Nothing Word writes is committed: only these facts.
"""

from __future__ import annotations

import io
import json
import pprint
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))

import oracle  # noqa: E402
from e1_probe import minimal_styles  # noqa: E402
from make_fixtures import DECL, NS, REL, WML, content_types, relationships, settings  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "e4-word.json"
TOC_STYLES = ROOT / "src" / "docx_agent" / "edit" / "word_toc_styles.py"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W

FILLER = ("This paragraph is filler that takes room on the page, so that the headings fall on "
          "different pages and the table of contents has page numbers to show. ")
SECTION = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" '
           'w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
           '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>')


def para(text: str, k: int, *, section: str = "") -> str:
    para_id = f"1{k:07X}"
    properties = f"<w:pPr>{section}</w:pPr>" if section else ""
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p w14:paraId="{para_id}" w14:textId="77777777">{properties}{run}</w:p>'


def package(body: str, *, last: str = SECTION) -> bytes:
    parts = {
        "[Content_Types].xml": content_types({"word/document.xml": WML + ".document.main+xml",
                                              "word/styles.xml": WML + ".styles+xml",
                                              "word/settings.xml": WML + ".settings+xml"}),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False)]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}{last}</w:body></w:document>",
        "word/styles.xml": minimal_styles(),
        "word/settings.xml": settings(15),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 10, 4, 12, 0, 0)), data)
    return buffer.getvalue()


def plain_document(count: int = 12) -> bytes:
    return package("".join(para(f"Paragraph {k}. " + FILLER, k) for k in range(1, count + 1)))


def sectioned_document() -> bytes:
    """Three sections of four paragraphs each, every section break a next-page one."""
    body = ""
    for k in range(1, 13):
        section = SECTION.replace("<w:sectPr>", '<w:sectPr><w:type w:val="nextPage"/>') if k in (4, 8) else ""
        body += para(f"Section {(k - 1) // 4 + 1}, paragraph {k}. " + FILLER, k, section=section)
    return package(body)


def headings_document() -> bytes:
    """Headings (styled by the script) over several pages, filler between them."""
    body = para("Contents follow.", 1)
    k = 2
    for chapter in range(1, 4):
        body += para(f"Chapter {chapter}", k)
        k += 1
        for part in range(1, 3):
            body += para(f"Part {chapter}.{part}", k)
            k += 1
            body += para(f"Detail {chapter}.{part}", k)
            k += 1
            for _ in range(6):
                body += para(FILLER * 3, k)
                k += 1
    return package(body)


def _guard(line: str, k: int) -> str:
    """One step, its failure written at the document's end instead of stopping the script."""
    return (f"try\n  {line}\non error m\n  insert text (\"[step {k} failed: \" & m & \"]\") "
            "at end of text object of d\nend try")


def _script(lines: list[str]) -> str:
    lines = [_guard(line, k) for k, line in enumerate(lines)]
    body = "\n      ".join(line.replace("\n", "\n      ") for line in lines)
    return f'''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file inputPath)
      set d to active document
      {body}
      save as d file name outputPath file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
end run
'''


def _end(k: int) -> str:
    """A collapsed range at paragraph ``k``'s text's end (before its mark)."""
    return f"(create range d start ((end of content of text object of paragraph {k} of d) - 1) " \
           f"end ((end of content of text object of paragraph {k} of d) - 1))"


def _start(k: int) -> str:
    return f"(create range d start (start of content of text object of paragraph {k} of d) " \
           f"end (start of content of text object of paragraph {k} of d))"


def breaks_script() -> list[str]:
    return [
        f"insert break at {_end(9)} break type section break even page",
        f"insert break at {_start(6)} break type section break continuous",
        f"insert break at {_end(3)} break type section break next page",
    ]


def setup_script() -> list[str]:
    s2 = "page setup of section 2 of d"
    return [
        f"set orientation of {s2} to orient landscape",
        f"set top margin of {s2} to 54",
        f"set left margin of {s2} to 90",
        f"set gutter of {s2} to 18",
        f"set number of text columns ({s2}) number of columns 2",
        f"set spacing between text columns of {s2} to 24",
        f"set line between text columns of {s2} to true",
        f"set vertical alignment of {s2} to align vertical center",
        f"set ln to line numbering of {s2}",
        "set active line of ln to true",
        "set count by of ln to 5",
        f"set line numbering of {s2} to ln",
        f"set different first page header footer of {s2} to true",
        "set pno to page number options of (get footer section 2 of d index header footer primary)",
        "set restart numbering at section of pno to true",
        "set starting number of pno to 5",
        "set number style of pno to page number style lowercase roman",
        "set odd and even pages header footer of page setup of section 3 of d to true",
    ]


def headers_script() -> list[str]:
    return [
        "set different first page header footer of page setup of section 1 of d to true",
        "set odd and even pages header footer of page setup of section 1 of d to true",
        'set content of text object of (get header section 1 of d index header footer primary) to "Default header one"',
        'set content of text object of (get header section 1 of d index header footer first page) to "First header one"',
        'set content of text object of (get header section 1 of d index header footer even pages) to "Even header one"',
        'set content of text object of (get footer section 1 of d index header footer primary) to "Footer one"',
        "set h2 to get header section 2 of d index header footer primary",
        "set link to previous of h2 to false",
        'set content of text object of h2 to "Header two, unlinked"',
    ]


def notes_script() -> list[str]:
    return [
        f"make new footnote at d with properties {{text range:{_end(2)}}}",
        f"make new footnote at d with properties {{text range:{_end(6)}}}",
        f"make new footnote at d with properties {{text range:{_end(10)}}}",
        f"make new endnote at d with properties {{text range:{_end(3)}}}",
        f"make new endnote at d with properties {{text range:{_end(7)}}}",
        "set fo to footnote options of (text object of section 2 of d)",
        "set footnote number style of fo to note number style lowercase roman",
        "set footnote numbering rule of fo to restart section",
        "set footnote starting number of fo to 1",
        "set footnote location of fo to beneath text",
    ]


def heading_levels() -> dict[int, int]:
    """Paragraph number (1-based) -> heading level in :func:`headings_document`."""
    out = {}
    k = 2
    for chapter in range(1, 4):
        out[k] = 1
        k += 1
        for part in range(1, 3):
            out[k] = 2
            out[k + 1] = 3
            k += 2 + 6
    return out


def _style_headings() -> list[str]:
    return [f"set style of paragraph {k} of d to style heading{level}" for k, level in heading_levels().items()]


def _span(k: int, length: int) -> str:
    """The first ``length`` characters of paragraph ``k``."""
    return f"(create range d start (start of content of text object of paragraph {k} of d) " \
           f"end ((start of content of text object of paragraph {k} of d) + {length}))"


def _count() -> int:
    return 1 + 3 * (1 + 2 * (2 + 6))


def fields_script() -> list[str]:
    levels = heading_levels()
    chapter2 = [k for k, level in levels.items() if level == 1][1]
    last = _count()
    return _style_headings() + [
        # From the end backwards, so the paragraph numbers hold.
        f'insert date time at {_end(last)} date time format "d MMMM yyyy" insert as field true',
        f'create new field text range {_end(last - 1)} field type field empty field text "HYPERLINK \\"https://example.com/e4\\""',
        f'insert caption at {_end(last - 2)} caption label caption figure title ": a figure" caption position caption position below',
        f'make new bookmark at d with properties {{name:"E4Target", text object:{_span(chapter2 + 2, 6)}}}',
        f"insert cross reference at {_end(last - 4)} reference type reference type bookmark reference kind reference content text reference item \"E4Target\" insert as hyperlink true",
        f"insert cross reference at {_end(last - 5)} reference type reference type bookmark reference kind reference page number reference item \"E4Target\" insert as hyperlink true",
        f"insert break at {_start(chapter2)} break type section break continuous",
    ]


_TOC = 'field text "\\\\o \\"1-{levels}\\" \\\\h \\\\z \\\\u"'


def tocfield_script() -> list[str]:
    """A table of contents put in an empty paragraph (the first), as Word's Insert Field
    and update make one, on a page whose text width is not A4's."""
    return _style_headings() + [
        "set left margin of page setup of section 1 of d to 90",
        "set right margin of page setup of section 1 of d to 36",
        f"create new field text range {_start(1)} field type field toc " + _TOC.format(levels=3),
        "update field (field 1 of d)",
    ]


def nine_document() -> bytes:
    body = para("", 1)
    for level in range(1, 10):
        body += para(f"Level {level} heading", level + 1)
        body += para(FILLER, level + 20)
    return package(body)


def tocstyles_script() -> list[str]:
    return [f"set style of paragraph {2 * level} of d to style heading{level}" for level in range(1, 10)] + [
        f"create new field text range {_start(1)} field type field toc " + _TOC.format(levels=9),
        "update field (field 1 of d)",
    ]


def removal_script() -> list[str]:
    """The break between the first two sections deleted, the sections told apart by their
    margins and headers."""
    return [
        "set left margin of page setup of section 1 of d to 36",
        "set left margin of page setup of section 2 of d to 108",
        "set h1 to get header section 1 of d index header footer primary",
        'set content of text object of h1 to "Header one"',
        "set h2 to get header section 2 of d index header footer primary",
        "set link to previous of h2 to false",
        'set content of text object of h2 to "Header two"',
        "set e to end of content of text object of section 1 of d",
        "delete (create range d start (e - 1) end e)",
    ]


def tracked_script() -> list[str]:
    levels = heading_levels()
    chapter2 = [k for k, level in levels.items() if level == 1][1]
    return _style_headings() + [
        "set h1 to get header section 1 of d index header footer primary",
        'set content of text object of h1 to "Header before"',
        "set track revisions of d to true",
        'insert text " tracked" at end of text object of h1',
        "set top margin of page setup of section 1 of d to 54",
        f"insert break at {_start(chapter2)} break type section break next page",
        "set number of text columns (page setup of section 2 of d) number of columns 2",
        f"make new footnote at d with properties {{text range:{_end(chapter2 + 2)}}}",
        f"create new field text range {_end(chapter2 + 4)} field type field page",
        f"create new field text range {_start(1)} field type field toc " + _TOC.format(levels=3),
        "update field (field 1 of d)",
    ]


def tracked_removal_script() -> list[str]:
    return [
        "set left margin of page setup of section 2 of d to 108",
        "set track revisions of d to true",
        "set e to end of content of text object of section 1 of d",
        "delete (create range d start (e - 1) end e)",
    ]


def headers2_script() -> list[str]:
    """Page-number fields as Word puts them in a footer; a footer unlinked without being
    written (does Word copy the previous one?); a header unlinked, then linked again."""
    return [
        "set different first page header footer of page setup of section 1 of d to true",
        "set odd and even pages header footer of page setup of section 1 of d to true",
        'set content of text object of (get footer section 1 of d index header footer primary) to "Footer one"',
        "create new field text range (text object of (get footer section 1 of d index header footer primary)) "
        "field type field page",
        "create new field text range (text object of (get footer section 1 of d index header footer first page)) "
        "field type field num pages",
        "create new field text range (text object of (get footer section 1 of d index header footer even pages)) "
        "field type field section pages",
        'set content of text object of (get header section 1 of d index header footer primary) to "Header one"',
        "set link to previous of (get footer section 2 of d index header footer primary) to false",
        "set h3 to get header section 3 of d index header footer primary",
        "set link to previous of h3 to false",
        'set content of text object of h3 to "Header three"',
        "set link to previous of h3 to true",
    ]


def notes2_script() -> list[str]:
    """Footnote options on the last section only; endnotes at each section's end; a note
    deleted."""
    return [
        f"make new footnote at d with properties {{text range:{_end(2)}}}",
        f"make new footnote at d with properties {{text range:{_end(6)}}}",
        f"make new footnote at d with properties {{text range:{_end(10)}}}",
        f"make new endnote at d with properties {{text range:{_end(3)}}}",
        "set fo to footnote options of (text object of section 3 of d)",
        "set footnote number style of fo to note number style uppercase letter",
        "set footnote starting number of fo to 3",
        "set eo to endnote options of (text object of section 1 of d)",
        "set endnote location of eo to end_of_section",
        "set endnote number style of eo to note number style arabic",
        "delete footnote 2 of d",
    ]


RUNS = {
    "breaks": (breaks_script, plain_document),
    "removal": (removal_script, sectioned_document),
    "setup": (setup_script, sectioned_document),
    "headers": (headers_script, sectioned_document),
    "headers2": (headers2_script, sectioned_document),
    "notes": (notes_script, sectioned_document),
    "notes2": (notes2_script, sectioned_document),
    "fields": (fields_script, headings_document),
    "tocfield": (tocfield_script, headings_document),
    "tocstyles": (tocstyles_script, nine_document),
    "tracked": (tracked_script, headings_document),
    "tracked-removal": (tracked_removal_script, sectioned_document),
}

#: The parts each probe's reading keeps besides the body.
_KEEP = re.compile(r"word/(header\d+|footer\d+|footnotes|endnotes|settings|people)\.xml$")


def run(names: list[str]) -> dict[str, bytes]:
    out = {}
    with oracle.session() as session:
        for name in names:
            script, document = RUNS[name]
            outcome = session.run_script(_script(script()), document(), name=f"e4-{name}", tag="saved")
            if not outcome:
                raise SystemExit(f"Word did not save {name}: {outcome.outcome} {outcome.detail}")
            out[name] = outcome.path.read_bytes()
    return out


def observe(data: bytes) -> dict:
    files = parts(data)
    entry: dict = {"blocks": blocks(data)}
    failed = re.findall(r"\[step \d+ failed: [^\]]*\]", files["word/document.xml"].decode("utf-8"))
    if failed:
        entry["failed"] = failed
    for part in sorted(files):
        if _KEEP.match(part):
            text = normalise(files[part].decode("utf-8"))
            if part.endswith("settings.xml"):
                text = re.sub(r"<w:rsids>.*?</w:rsids>|<w15:docId [^>]*/>", "", text)
            entry[part] = text
    rels = files.get("word/_rels/document.xml.rels", b"").decode("utf-8")
    entry["relationships"] = sorted(re.findall(r'Type="[^"]*/([^/"]+)" Target="([^"]+)"', rels))
    entry["content_types"] = sorted(re.findall(r'PartName="([^"]+)"', files["[Content_Types].xml"].decode()))
    styles = etree.fromstring(files["word/styles.xml"])
    entry["styles"] = sorted(s.find(_W + "name").get(_W + "val") for s in styles.findall(_W + "style"))
    return entry


def write_toc_styles(data: bytes) -> None:
    """The TOC 1-9 styles Word added with a table of contents, as ``word_styles.py``'s
    entries are written (``e1_probe.read_styles``)."""
    from e1_probe import read_styles

    found = {name: entry for name, entry in read_styles(parts(data)["word/styles.xml"], b"").items()
             if re.fullmatch(r"toc \d", name)}
    if len(found) != 9:
        raise SystemExit(f"expected TOC 1-9, found {sorted(found)}")
    TOC_STYLES.write_text(
        '"""Word\'s TOC 1-9 styles, as Word 16.106 writes them with a table of contents.\n\n'
        "Generated by ``tools/e4_probe.py`` from the probe Word saved (the observations are in\n"
        "``tests/observations/e4-word.json``); do not edit by hand.  The same form as\n"
        "``word_styles.py``: keyed by ``w:name``, ``{id}`` and the references filled in when written.\n"
        '"""\n\nSTYLES: dict[str, dict] = ' + pprint.pformat(found, width=100) + "\n", encoding="utf-8")
    print(f"wrote {TOC_STYLES.relative_to(ROOT)}")


def main(names: list[str]) -> None:
    saved = run(names or list(RUNS))
    if "tocstyles" in saved:
        write_toc_styles(saved["tocstyles"])
    if names:
        for name, data in saved.items():
            print(name, json.dumps(observe(data), indent=1)[:4000])
        return
    OBSERVATIONS.write_text(json.dumps({
        "_about": ("What Word 16.106 for Mac wrote for tools/e4_probe.py: each probe's body blocks and its "
                   "header, footer, notes, settings and people parts, its relationships, content types "
                   "and style names. Authors and dates replaced by placeholders, rsids dropped. A step "
                   "Word refused is listed under 'failed'. Regenerate with python tools/e4_probe.py."),
        **{name: observe(data) for name, data in saved.items()}}, indent=1, ensure_ascii=False) + "\n",
        encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


# -- reading ---------------------------------------------------------------------------------

_DROP = re.compile(r' w:rsid\w*="[^"]*"')


def normalise(xml: str) -> str:
    xml = re.sub(r' xmlns:\w+="[^"]*"', "", xml)
    xml = _DROP.sub("", xml)
    xml = re.sub(r'w:author="[^"]*"', 'w:author="{author}"', xml)
    xml = re.sub(r'w:date="[^"]*"', 'w:date="{date}"', xml)
    xml = re.sub(r'(w16du):dateUtc="[^"]*"', r'\1:dateUtc="{date}"', xml)
    xml = re.sub(r'w15:author="[^"]*"', 'w15:author="{author}"', xml)
    xml = re.sub(r'w15:userId="[^"]*"', 'w15:userId="{user}"', xml)
    return xml


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def blocks(data: bytes) -> list[str]:
    root = etree.fromstring(parts(data)["word/document.xml"])
    return [normalise(etree.tostring(child, encoding="unicode")) for child in root.find(_W + "body")]


if __name__ == "__main__":
    main(sys.argv[1:])
