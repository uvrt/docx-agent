#!/usr/bin/env python3
"""Measure what Word writes for the two things ``insert_markdown`` makes that E1 and E3 did
not measure: a table in Word's own Table Grid style, and a document's first footnote.

A probe document is written here (mode 15, the handful of styles ``e1_probe.py`` writes,
no notes parts), and Word -- through ``tests/oracle.py``'s machine-wide lock -- opens it
and, by AppleScript:

* makes a table of two rows and three columns before the second paragraph and gives it the
  Table Grid style.  This Word's interface is Dutch, styles are set by their *localised*
  name, and Word's dictionary has no constant for Table Grid, so each candidate name is
  tried and the one Word took is read back from the saved file (``w:name`` is English);
* puts a footnote after the third paragraph's text.

and saves.  A second probe has Word itself, tracking, type two paragraphs at the end of a
document whose last paragraph is a heading and make them Normal -- the form
``insert_markdown`` writes at a container's end (E3's typing form) -- and saves it, and
then its own Reject All and Accept All of it.  What it wrote is read back:

* the Table Grid definition, keyed by ``w:name`` as ``word_styles.py``'s entries are
  (``e1_probe.read_styles``), into ``src/docx_agent/edit/word_table_styles.py``;
* the table's properties and grid, and how it divided the text width among its columns;
* the parts, relationships, content types and settings a first footnote brings, and the
  footnote's paragraph;
* the typed paragraphs' revision form, and what Word's own review leaves of it (its Reject
  All leaves the last paragraph's property change unrejected, the heading's style lost).

Word's ``auto format as you type`` border for a ``---`` line (the thematic break) cannot be
measured this way: typing through AppleScript does not trigger AutoFormat (tried: the
text stays ``---``).

    python tools/e2_probe.py        # writes tests/observations/e2-word.json and
                                    # src/docx_agent/edit/word_table_styles.py

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
from e1_probe import minimal_styles, read_styles  # noqa: E402
from make_fixtures import DECL, NS, REL, WML, content_types, p, relationships, settings  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "e2-word.json"
TABLE = ROOT / "src" / "docx_agent" / "edit" / "word_table_styles.py"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
#: Table Grid's name in each interface language this probe may meet.
TABLE_GRID_NAMES = ("Table Grid", "Tabelraster")


def probe_document() -> bytes:
    body = "".join(p(f"Probe paragraph {k}", f"1{k:07X}") for k in range(1, 5))
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" '
             'w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    parts = {
        "[Content_Types].xml": content_types({"word/document.xml": WML + ".document.main+xml",
                                              "word/styles.xml": WML + ".styles+xml",
                                              "word/settings.xml": WML + ".settings+xml"}),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False)]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": minimal_styles(),
        "word/settings.xml": settings(15),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 10, 3, 12, 0, 0)), data)
    return buffer.getvalue()


def script() -> str:
    tries = "\n".join(f'''        try
          set style of t to "{name}"
        end try''' for name in TABLE_GRID_NAMES)
    return f'''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file inputPath)
      set d to active document
      set e to (end of content of text object of paragraph 3 of d) - 1
      make new footnote at d with properties {{text range:(create range d start e end e)}}
      set s to start of content of text object of paragraph 2 of d
      set t to make new table at d with properties {{text object:(create range d start s end s), number of rows:2, number of columns:3}}
{tries}
      save as d file name outputPath file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
end run
'''


def typed_document() -> bytes:
    from make_fixtures import styles

    body = (p("Plain first.", "10000001") + p("A heading last.", "10000002", style="Heading1")
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" '
            'w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    parts = {
        "[Content_Types].xml": content_types({"word/document.xml": WML + ".document.main+xml",
                                              "word/styles.xml": WML + ".styles+xml",
                                              "word/settings.xml": WML + ".settings+xml"}),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False)]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": styles(),
        "word/settings.xml": settings(15),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 10, 3, 12, 0, 0)), data)
    return buffer.getvalue()


TYPED = ["set track revisions of d to true",
         "set e to end of content of text object of paragraph 2 of d",
         "select (create range d start (e - 1) end (e - 1))",
         "type paragraph selection", 'type text selection text "First new."',
         "type paragraph selection", 'type text selection text "Second new."',
         "set style of paragraph 3 of d to style normal",
         "set style of paragraph 4 of d to style normal"]


def _body(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
    xml = re.sub(r'w:author="[^"]*"', 'w:author="{author}"', _clean(xml))
    xml = re.sub(r'(w:date|w16du:dateUtc)="[^"]*"', r'\1="{date}"', xml)
    return xml[xml.index("<w:body>") + 8:xml.index("<w:sectPr")]


def _clean(xml: str) -> str:
    xml = re.sub(r' xmlns:\w+="[^"]*"', "", xml)
    xml = re.sub(r' w:rsid\w*="[^"]*"', "", xml)
    xml = re.sub(r' w14:(paraId|textId)="[^"]*"', "", xml)
    return re.sub(r'<w15:docId [^>]*/>|<w:rsids>.*?</w:rsids>', "", xml)


def main() -> None:
    oracle.ORACLE_DIR.mkdir(parents=True, exist_ok=True)
    from e3_probe import _script, review_script

    with oracle.session() as session:
        outcome = session.run_script(script(), probe_document(), name="e2-probe", tag="table-footnote",
                                     timeout=200)
        typed = session.run_script(_script(TYPED), typed_document(), name="e2-probe", tag="typed", timeout=200)
        if not typed:
            raise SystemExit(f"Word did not save the typed probe: {typed.outcome} {typed.detail}")
        reviews = {action: session.run_script(review_script(action), typed.path.read_bytes(), name="e2-probe-typed",
                                              tag=action, timeout=200) for action in ("accept", "reject")}
    typing = {"typed": _body(typed.path.read_bytes()),
              **{action: _body(result.path.read_bytes()) for action, result in reviews.items()}}
    if not outcome:
        raise SystemExit(f"Word did not save the probe: {outcome.outcome} {outcome.detail}")
    with zipfile.ZipFile(outcome.path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    outcome.path.unlink(missing_ok=True)

    measured = read_styles(files["word/styles.xml"], files.get("word/numbering.xml", b""))
    if "Table Grid" not in measured:
        raise SystemExit("Word did not apply Table Grid under any of " + ", ".join(TABLE_GRID_NAMES))
    document = etree.fromstring(files["word/document.xml"])
    body = document.find(_W + "body")
    table = body.find(_W + "tbl")
    grid = [int(col.get(_W + "w")) for col in table.iter(_W + "gridCol")]
    section = body.find(_W + "sectPr")
    size, margins = section.find(_W + "pgSz"), section.find(_W + "pgMar")
    text_width = int(size.get(_W + "w")) - int(margins.get(_W + "left")) - int(margins.get(_W + "right"))
    referencing = next(p for p in body.iter(_W + "p") if p.find(f".//{_W}footnoteReference") is not None)
    settings_root = etree.fromstring(files["word/settings.xml"])
    notes_settings = [_clean(etree.tostring(node, encoding="unicode"))
                      for node in settings_root if etree.QName(node).localname in ("footnotePr", "endnotePr")]
    rels = files["word/_rels/document.xml.rels"].decode()
    types = files["[Content_Types].xml"].decode()
    observations = {
        "word": "16.106 (macOS), Dutch interface",
        "table": {
            "properties": _clean(etree.tostring(table.find(_W + "tblPr"), encoding="unicode")),
            "grid": grid,
            "text_width": text_width,
            "cell": _clean(etree.tostring(table.find(f"{_W}tr/{_W}tc/{_W}tcPr"), encoding="unicode")),
        },
        "typing_at_the_end": typing,
        "footnote": {
            "parts": sorted(name for name in files if re.fullmatch(r"word/(foot|end)notes\.xml", name)),
            "relationships": sorted(t.rpartition("/")[2] for t in re.findall(r'Type="([^"]+)"', rels)
                                    if t.endswith(("footnotes", "endnotes"))),
            "content_types": sorted(re.findall(r'PartName="(/word/(?:foot|end)notes\.xml)"', types)),
            "settings": notes_settings,
            "footnotes.xml": _clean(files["word/footnotes.xml"].decode()).split("?>", 1)[1],
            "endnotes.xml": _clean(files["word/endnotes.xml"].decode()).split("?>", 1)[1]
            if "word/endnotes.xml" in files else None,
            "reference": _clean(etree.tostring(referencing, encoding="unicode")),
        },
    }
    OBSERVATIONS.write_text(json.dumps(observations, indent=1, ensure_ascii=False, sort_keys=True) + "\n")
    entry = {"Table Grid": measured["Table Grid"]}
    TABLE.write_text(
        '"""Word\'s Table Grid style, as Word 16.106 writes it on first use.\n\n'
        "Generated by ``tools/e2_probe.py`` from the probe Word saved (the observations are in\n"
        "``tests/observations/e2-word.json``); do not edit by hand.  The same form as\n"
        "``word_styles.py``: keyed by ``w:name``, ``{id}`` and the references filled in when written.\n"
        '"""\n\n'
        f"STYLES: dict[str, dict] = {pprint.pformat(entry, width=110)}\n")
    print(f"Table Grid measured; grid {grid} of {text_width}; footnote parts {observations['footnote']['parts']}")


if __name__ == "__main__":
    main()
