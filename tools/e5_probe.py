#!/usr/bin/env python3
"""Measure what Word writes for E5's tables, drawings and content controls.

Probe documents are written here (mode 15, ``e1_probe``'s handful of styles and Word's
Table Grid), and Word -- through ``tests/oracle.py``'s machine-wide lock -- opens each and,
by AppleScript (no clipboard), makes one family of edits and saves:

* ``tables`` (and ``tables-tracked``, the same with tracking on): on tables made as Word
  makes them (E2), one property each -- a header row, fixed layout, a width in percent,
  alignment, indent, the style options, a row's exact and at-least height, a row that may
  not split, a cell's shading, borders, margins, vertical alignment and text direction,
  column and cell widths, a table floated beside the text, a table nested in a cell;
* ``merges`` (and ``merges-tracked``): cells merged across, down and as a rectangle; a cell
  split into columns, into rows, into both; a merged cell split again; a column inserted
  beside a cell another row spans across, on either side; a column deleted that a span
  covers; a row inserted inside a vertical merge and one deleted from its middle;
* ``drawings`` (and ``drawings-tracked``): inline pictures floated (``convert to shape``),
  each wrap type, its side and distances, positions against each frame, by offset and
  alignment, the order (``z order``), a locked anchor, overlap refused; one floated and put
  inline again; a text box and a shape made by Word;
* ``controls`` (and ``controls-tracked``): content controls of each kind written here as
  Word 365 writes them (Word's dictionary has no content-control object, so Word cannot make
  one by script), filled by Word typing into them; a data-bound one, its custom XML part
  stale on purpose.

What it wrote is read back -- each part normalised (rsids, authors and dates replaced) --
into ``tests/observations/e5-word.json``.

    python tools/e5_probe.py            # every probe, the observations written
    python tools/e5_probe.py tables     # one probe, printed

Nothing Word writes is committed: only these facts.
"""

from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "src"))

import oracle  # noqa: E402
from e1_probe import minimal_styles  # noqa: E402
from make_fixtures import DECL, NS, REL, WML, content_types, png, relationships, settings  # noqa: E402

from docx_agent.edit.word_table_styles import STYLES as TABLE_STYLES  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "e5-word.json"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
SECTION = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" '
           'w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
           '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>')
#: The namespaces a drawing needs, beside ``make_fixtures.NS``.
DRAWING_NS = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
              'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
              'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')


class Ids:
    """Deterministic paraIds for a probe document."""

    def __init__(self, base: int) -> None:
        self.next = base

    def __call__(self) -> str:
        self.next += 1
        return f"{self.next:08X}"


def para(ids: Ids, text: str = "", props: str = "") -> str:
    properties = f"<w:pPr>{props}</w:pPr>" if props else ""
    pieces = [f'<w:t xml:space="preserve">{piece}</w:t>' if piece else "" for piece in text.split("\t")]
    run = f'<w:r>{"<w:tab/>".join(pieces)}</w:r>' if text else ""
    return f'<w:p w14:paraId="{ids()}" w14:textId="77777777">{properties}{run}</w:p>'


def table(ids: Ids, rows: list[list], grid: list[int], *, props: str = "", look: bool = True) -> str:
    """A table as Word makes one (E2): Table Grid, ``tblW`` auto, ``tblLook`` 04A0.  A cell is
    its text, or ``(text, tcPr extra, span)``; ``None`` is a cell a span covers."""
    look_xml = ('<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" '
                'w:noHBand="0" w:noVBand="1"/>') if look else ""
    out = [f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>{props}{look_xml}'
           "</w:tblPr><w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in grid) + "</w:tblGrid>"]
    for row in rows:
        out.append(f'<w:tr w14:paraId="{ids()}" w14:textId="77777777">')
        column = 0
        for cell in row:
            if cell is None:
                continue
            text, extra, span = (cell, "", 1) if isinstance(cell, str) else cell
            width = sum(grid[column:column + span])
            span_xml = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
            out.append(f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{span_xml}{extra}</w:tcPr>'
                       + "".join(para(ids, line) for line in text.split("\n")) + "</w:tc>")
            column += span
        out.append("</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def grid_table(ids: Ids, key: str, rows: int = 3, columns: int = 3, width: int = 2000, **options) -> str:
    return table(ids, [[f"{key} r{i}c{j}" for j in range(1, columns + 1)] for i in range(1, rows + 1)],
                 [width] * columns, **options)


def styles() -> str:
    grid = TABLE_STYLES["Table Grid"]["xml"].replace("{id}", "TableGrid").replace("{basedOn}", "TableNormal")
    grid = grid.replace(f' xmlns:w="{W}"', "")
    return minimal_styles().replace("</w:styles>", grid + "</w:styles>")


def package(body: str, *, extra_parts: dict[str, bytes | str] | None = None,
            extra_types: dict[str, str] | None = None, extra_defaults: dict[str, str] | None = None,
            document_rels: list[tuple[str, str, str, bool]] = (), namespaces: str = "") -> bytes:
    types = {"word/document.xml": WML + ".document.main+xml", "word/styles.xml": WML + ".styles+xml",
             "word/settings.xml": WML + ".settings+xml", **(extra_types or {})}
    content = content_types(types)
    for extension, kind in (extra_defaults or {}).items():
        content = content.replace("<Default ", f'<Default Extension="{extension}" ContentType="{kind}"/><Default ', 1)
    parts = {
        "[Content_Types].xml": content,
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False),
                                                       *document_rels]),
        "word/document.xml": f"{DECL}<w:document {NS} {namespaces}><w:body>{body}{SECTION}</w:body></w:document>",
        "word/styles.xml": styles(),
        "word/settings.xml": settings(15),
        **(extra_parts or {}),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 10, 4, 12, 0, 0)), data)
    return buffer.getvalue()


# -- tables ----------------------------------------------------------------------------------

#: The property cases, one table each, in order: key -> the AppleScript steps on ``t``
#: (the table) -- ``{c11}`` is its first cell, ``{c22}`` the middle one.
TABLE_CASES: dict[str, list[str]] = {
    "Head": ["set heading format of row 1 of t to true"],
    "Fixed": ["set allow auto fit of t to false"],
    "Pct": ["set preferred width type of t to preferred width percent", "set preferred width of t to 80"],
    "Center": ["set alignment of row options of t to align row center"],
    "Right": ["set alignment of row options of t to align row right"],
    "Indent": ["set row left indent of row options of t to 36"],
    "Look": ["set apply style heading rows of t to false", "set apply style first column of t to false",
             "set apply style last row of t to true", "set apply style last column of t to true"],
    "Exact": ["set height rule of row 1 of t to row height exactly", "set height of row 1 of t to 30"],
    "AtLeast": ["set height rule of row 2 of t to row height at least", "set height of row 2 of t to 40"],
    "CantSplit": ["set allow break across pages of row 1 of t to false"],
    "Shade": ["set texture of shading of {c22} to texture solid",
              "set background pattern color index of shading of {c22} to turquoise",
              "set background pattern color index of shading of {c11} to yellow"],
    "Border": ["set b to get border {c22} which border border top", "set line style of b to line style double",
               "set b to get border {c22} which border border left",
               "set line style of b to line style single", "set line width of b to line width225 point",
               "set color index of b to red"],
    "Margin": ["set top padding of {c22} to 5", "set left padding of {c22} to 12",
               "set left padding of t to 9"],
    "VAlign": ["set vertical alignment of {c22} to cell align vertical center",
               "set vertical alignment of {c11} to cell align vertical bottom"],
    "TextDir": ["set orientation of text object of {c22} to text orientation upward",
                "set orientation of text object of {c11} to text orientation downward"],
    "Widths": ["set preferred width type of {c11} to preferred width points", "set preferred width of {c11} to 72",
               "set width of column 2 of t to 150"],
    "Float": ["set wrap around text of row options of t to true",
              "set relative horizontal position of row options of t to relative horizontal position page",
              "set horizontal position of row options of t to 100",
              "set relative vertical position of row options of t to relative vertical position paragraph",
              "set vertical position of row options of t to 20",
              "set distance left of row options of t to 9", "set distance right of row options of t to 12",
              "set distance top of row options of t to 3", "set distance bottom of row options of t to 6",
              "set allow overlap of row options of t to false"],
    "Spacing": ["set spacing of t to 3"],
}


def tables_document() -> bytes:
    ids = Ids(0x30000000)
    body = para(ids, "Tables follow.")
    for key in TABLE_CASES:
        body += para(ids, f"Before {key}.") + grid_table(ids, key)
    # The table a nested one goes in.
    body += para(ids, "Before Nest.")
    body += table(ids, [["Nest r1c1\nInner one\tInner two", "Nest r1c2"], ["Nest r2c1", "Nest r2c2"]], [4500, 4500])
    body += para(ids, "The end.")
    return package(body)


def _guard(line: str, k: int) -> str:
    """One step, its failure written at the document's end instead of stopping the script."""
    return (f"try\n  {line}\non error m\n  insert text (\"[step {k} failed: \" & m & \"]\") "
            "at end of text object of d\nend try")


def _script(lines: list[str], *, track: bool = False) -> str:
    lines = (["set track revisions of d to true"] if track else []) + lines
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


def _cell(t: str, r: int, c: int) -> str:
    return f"(get cell from table ({t}) row {r} column {c})"


def tables_script() -> list[str]:
    lines = []
    for k, (key, steps) in enumerate(TABLE_CASES.items(), start=1):
        t = f"table {k} of d"
        lines.append(f"set t to {t}")
        for step in steps:
            lines.append(step.format(c11=_cell("t", 1, 1), c22=_cell("t", 2, 2)))
    nest = len(TABLE_CASES) + 1
    # A table nested in the outer table's first cell, at its second paragraph (AppleScript's
    # make new table is Word 97's form: no style, fixed layout; Word 2000's, which Insert
    # Table makes, is E2's).
    lines += [
        f"set c to {_cell(f'table {nest} of d', 1, 1)}",
        "make new table at text object of paragraph 2 of text object of c with properties "
        "{number of rows:2, number of columns:2}",
    ]
    return lines


# -- merges, splits, and columns and rows across them ----------------------------------------

#: key -> (rows, grid, steps).  Rows as :func:`table` takes them.
def _merge_cases() -> dict[str, tuple[list, list[int], list[str]]]:
    plain = lambda key: [[f"{key} r{i}c{j}" for j in range(1, 4)] for i in range(1, 4)]  # noqa: E731
    vmerge_start = '<w:vMerge w:val="restart"/>'
    vmerge_cont = "<w:vMerge/>"

    def spanned(key: str) -> list:
        return [[(f"{key} r1c1", "", 2), None, f"{key} r1c3"], *plain(key)[1:]]

    def vmerged(key: str) -> list:
        rows = plain(key)
        rows[0][0] = (f"{key} r1c1", vmerge_start, 1)
        rows[1][0] = ("", vmerge_cont, 1)
        rows[2][0] = ("", vmerge_cont, 1)
        return rows

    c = lambda t, r, k: _cell(t, r, k)  # noqa: E731
    return {
        "MergeAcross": (plain("MergeAcross"), [2000] * 3, ["merge cell {} with {}".format(c("t", 1, 1), c("t", 1, 2))]),
        "MergeDown": (plain("MergeDown"), [2000] * 3, ["merge cell {} with {}".format(c("t", 1, 1), c("t", 2, 1))]),
        "MergeRect": (plain("MergeRect"), [2000] * 3, ["merge cell {} with {}".format(c("t", 1, 1), c("t", 2, 2))]),
        "SplitCols": (plain("SplitCols"), [2000] * 3,
                      ["split cell {} number of rows 1 number of columns 2".format(c("t", 2, 2))]),
        "SplitRows": (plain("SplitRows"), [2000] * 3,
                      ["split cell {} number of rows 2 number of columns 1".format(c("t", 2, 2))]),
        "SplitBoth": (plain("SplitBoth"), [2000] * 3,
                      ["split cell {} number of rows 2 number of columns 2".format(c("t", 2, 2))]),
        "Unspan": (spanned("Unspan"), [2000] * 3,
                   ["split cell {} number of rows 1 number of columns 2".format(c("t", 1, 1))]),
        "Unmerge": (vmerged("Unmerge"), [2000] * 3,
                    ["split cell {} number of rows 3 number of columns 1".format(c("t", 1, 1))]),
        "ColRightOfSpan": (spanned("ColRightOfSpan"), [2000] * 3,
                           ["select text object of {}".format(c("t", 2, 1)),
                            "insert columns selection position insert on the right"]),
        "ColLeftOfSpan": (spanned("ColLeftOfSpan"), [2000] * 3,
                          ["select text object of {}".format(c("t", 2, 2)), "insert columns selection"]),
        "ColDelSpan": (spanned("ColDelSpan"), [2000] * 3,
                       ["select text object of {}".format(c("t", 2, 2)), "select column selection",
                        "delete column 1 of selection"]),
        "RowInMerge": (vmerged("RowInMerge"), [2000] * 3,
                       ["select text object of {}".format(c("t", 2, 2)),
                        "insert rows selection position below"]),
        "RowDelMerge": (vmerged("RowDelMerge"), [2000] * 3,
                        ["select text object of {}".format(c("t", 2, 2)), "select row selection",
                         "type backspace selection"]),
        "RowDelFirst": (vmerged("RowDelFirst"), [2000] * 3,
                        ["select text object of {}".format(c("t", 1, 2)), "select row selection",
                         "type backspace selection"]),
        "RowInSpan": (spanned("RowInSpan"), [2000] * 3,
                      ["select text object of {}".format(c("t", 1, 1)),
                       "insert rows selection position below"]),
        "ColDelPlain": (plain("ColDelPlain"), [2000] * 3, ["delete column 2 of t"]),
    }


def merges_document() -> bytes:
    ids = Ids(0x34000000)
    body = para(ids, "Merges follow.")
    for key, (rows, grid, _) in _merge_cases().items():
        body += para(ids, f"Before {key}.") + table(ids, rows, grid)
    body += para(ids, "The end.")
    return package(body)


def merges_script() -> list[str]:
    lines = []
    for k, (key, (_, _, steps)) in enumerate(_merge_cases().items(), start=1):
        lines.append(f"set t to table {k} of d")
        lines += steps
    return lines


# -- the tracked merge across: which form Word accepts and rejects -----------------------------

AGENT = 'w:author="E5 Agent" w:date="2026-10-04T12:00:00Z"'


def mergeforms_document() -> bytes:
    """One merge across (r1c1 and r1c2) written in each candidate tracked form, for Word's
    Accept All and Reject All: ``E3`` the first cell's span in a ``w:tcPrChange`` and the
    second ``w:cellDel`` (E3's form); ``Grid`` with the old grid in a ``w:tblGridChange``;
    ``GridDel`` with the deleted cell's old properties too; ``Row`` with every cell of the
    row's (``w:tcPrChange`` on each, Word's own form for a table's properties); ``Legacy``
    the old ``w:hMerge``."""
    ids = Ids(0x38000000)
    counter = [600]

    def nid() -> int:
        counter[0] += 1
        return counter[0]

    def p(text: str, mark: str = "", deleted: bool = False, inserted: bool = False) -> str:
        props = f"<w:pPr><w:rPr>{mark}</w:rPr></w:pPr>" if mark else ""
        run = ""
        if text:
            kind = "delText" if deleted else "t"
            run = f'<w:r><w:{kind} xml:space="preserve">{text}</w:{kind}></w:r>'
            if deleted:
                run = f'<w:del w:id="{nid()}" {AGENT}>{run}</w:del>'
            if inserted:
                run = f'<w:ins w:id="{nid()}" {AGENT}>{run}</w:ins>'
        return f'<w:p w14:paraId="{ids()}" w14:textId="77777777">{props}{run}</w:p>'

    def tc(content: str, extra: str = "", width: int = 2000) -> str:
        return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{extra}</w:tcPr>{content}</w:tc>'

    def change() -> str:
        return f'<w:tcPrChange w:id="{nid()}" {AGENT}><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr></w:tcPrChange>'

    def tbl(key: str, row1: list[str], grid_change: bool) -> str:
        cols = '<w:gridCol w:w="2000"/>' * 3
        if grid_change:
            cols += f'<w:tblGridChange w:id="{nid()}"><w:tblGrid>{cols}</w:tblGrid></w:tblGridChange>'
        rows = [row1] + [[tc(p(f"{key} r{i}c{j}")) for j in (1, 2, 3)] for i in (2, 3)]
        body = "".join(f'<w:tr w14:paraId="{ids()}" w14:textId="77777777">{"".join(r)}</w:tr>' for r in rows)
        return (p(f"Before {key}.") + '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
                '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" '
                f'w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>{cols}</w:tblGrid>{body}</w:tbl>')

    def lead(key: str) -> str:
        return tc(p(f"{key} r1c1", f'<w:ins w:id="{nid()}" {AGENT}/>') + p(f"{key} r1c2", inserted=True),
                  '<w:gridSpan w:val="2"/>' + change(), 4000)

    def gone(key: str, recorded: bool) -> str:
        return tc(p(f"{key} r1c2", f'<w:del w:id="{nid()}" {AGENT}/>', deleted=True),
                  f'<w:cellDel w:id="{nid()}" {AGENT}/>' + (change() if recorded else ""))

    body = p("Merge forms follow.")
    body += tbl("E3", [lead("E3"), gone("E3", False), tc(p("E3 r1c3"))], False)
    body += tbl("Grid", [lead("Grid"), gone("Grid", False), tc(p("Grid r1c3"))], True)
    body += tbl("GridDel", [lead("GridDel"), gone("GridDel", True), tc(p("GridDel r1c3"))], True)
    body += tbl("Row", [lead("Row"), gone("Row", True), tc(p("Row r1c3"), change())], True)
    body += tbl("Legacy", [tc(p("Legacy r1c1", f'<w:ins w:id="{nid()}" {AGENT}/>') + p("Legacy r1c2", inserted=True),
                              '<w:hMerge w:val="restart"/>' + change()),
                           tc(p("Legacy r1c2", f'<w:del w:id="{nid()}" {AGENT}/>', deleted=True),
                              "<w:hMerge/>" + change()), tc(p("Legacy r1c3"))], False)
    body += p("The end.")
    return package(body)


# -- drawings -------------------------------------------------------------------------------

FILLER = "Text that runs beside and around the drawing so that its wrapping shows on the page. "

#: Each case floats the picture of its paragraph (``s``) and sets what it names.  The left
#: and top of -999995 are Word's ``wdShapeCenter`` (-999996 right, -999997 bottom,
#: -999998 left, -999999 top).
DRAWING_CASES: dict[str, list[str]] = {
    "Float": [],
    "Square": ["set wrap type of wrap format of {s} to wrap square"],
    "Tight": ["set wrap type of wrap format of {s} to wrap tight"],
    "Through": ["set wrap type of wrap format of {s} to wrap through"],
    "TopBottom": ["set wrap type of wrap format of {s} to wrap top bottom"],
    "Front": ["set wrap type of wrap format of {s} to wrap front"],
    "Behind": ["set wrap type of wrap format of {s} to wrap behind"],
    "Sides": ["set wrap type of wrap format of {s} to wrap square", "set wrap side of wrap format of {s} to wrap left",
              "set distance top of wrap format of {s} to 3", "set distance bottom of wrap format of {s} to 6",
              "set distance left of wrap format of {s} to 9", "set distance right of wrap format of {s} to 12"],
    "Page": ["set relative horizontal position of {s} to relative horizontal position page",
             "set left position of {s} to 100",
             "set relative vertical position of {s} to relative vertical position page", "set top of {s} to 200"],
    "Aligned": ["set relative horizontal position of {s} to relative horizontal position margin",
                "set left position of {s} to -999995",
                "set relative vertical position of {s} to relative vertical position margin", "set top of {s} to -999997"],
    "Column": ["set relative horizontal position of {s} to relative horizontal position column",
               "set left position of {s} to 20",
               "set relative vertical position of {s} to relative vertical position paragraph", "set top of {s} to 10"],
    "Character": ["set relative horizontal position of {s} to relative horizontal position character",
                  "set left position of {s} to 5",
                  "set relative vertical position of {s} to relative vertical position line", "set top of {s} to 2"],
    "Inside": ["set relative horizontal position of {s} to relative horizontal position page",
               "set left position of {s} to -999994",
               "set relative vertical position of {s} to relative vertical position top margin",
               "set top of {s} to -999999"],
    "Front2": ["set wrap type of wrap format of {s} to wrap front", "z order {s} z order command bring shape to front"],
    "Back2": ["set wrap type of wrap format of {s} to wrap front", "z order {s} z order command send shape to back"],
    "Lock": ["set lock anchor of {s} to true"],
    "Overlap": ["set allow overlap of wrap format of {s} to false"],
    "Inline": ["convert to inline shape {s}"],
    "Size": ["set width of {s} to 144", "set height of {s} to 72"],
}


def drawings_document() -> bytes:
    from docx_agent.edit.pictures import _drawing

    ids = Ids(0x3A000000)
    body = para(ids, "Drawings follow.")
    for k, key in enumerate(DRAWING_CASES, start=1):
        drawing = etree.tostring(_drawing("rIdImg", k, f"Picture {k}", f"Case {key}", (72.0, 36.0),
                                          f"{0x3B000000 + 2 * k:08X}", f"{0x3B000001 + 2 * k:08X}"),
                                 encoding="unicode")
        drawing = re.sub(r' xmlns:\w+="[^"]*"', "", drawing)
        body += (f'<w:p w14:paraId="{ids()}" w14:textId="77777777"><w:r><w:t xml:space="preserve">Case {key}: </w:t>'
                 f"</w:r><w:r>{drawing}</w:r><w:r><w:t xml:space=\"preserve\"> {FILLER * 3}</w:t></w:r></w:p>")
    body += para(ids, "Box anchor. " + FILLER) + para(ids, "Shape anchor. " + FILLER) + para(ids, "The end.")
    return package(body, namespaces=DRAWING_NS, extra_parts={"word/media/image1.png": png(16, 8, (200, 60, 40))},
                   extra_defaults={"png": "image/png"},
                   document_rels=[("rIdImg", REL + "image", "media/image1.png", False)])


def drawings_script() -> list[str]:
    lines = []
    for k, (key, steps) in enumerate(DRAWING_CASES.items(), start=1):
        paragraph = k + 1
        lines.append(f"convert to shape (inline shape 1 of text object of paragraph {paragraph} of d)")
        # Word's references to shapes are by index, and a conversion reorders them: each
        # case finds its own by name.
        lines.append(f'repeat with i from 1 to (count of shapes of d)\n  if name of (shape i of d) is "Picture {k}" '
                     "then set s to shape i of d\nend repeat")
        lines += [step.format(s="s") for step in steps]
    box = len(DRAWING_CASES) + 2
    lines += [
        f"set r to text object of paragraph {box} of d",
        "set b to make new shape at d with properties {auto shape type:autoshape rectangle, left position:300, "
        "top:20, width:150, height:60, anchor:r}",
        'set content of text range of text frame of b to "Text box line"',
        f"select text object of paragraph {box + 2} of d",
        "create textbox selection",
        f"set r to text object of paragraph {box + 1} of d",
        "set e to make new shape at d with properties {auto shape type:autoshape oval, left position:200, "
        "top:30, width:90, height:45, anchor:r}",
    ]
    return lines


# -- content controls -----------------------------------------------------------------------

#: The custom XML part a bound control maps, and its store id.
STORE_ID = "{6C3C8BC8-F283-45AE-878A-BAB7291F4D35}"
BOUND_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<root xmlns="urn:docx-agent:e5"><name>Bound from XML</name><when>2026-10-04</when></root>')
ITEM_PROPS = ('<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n'
              f'<ds:datastoreItem ds:itemID="{STORE_ID}" '
              'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/>'
              "</ds:datastoreItem>")
CHECKBOX = ('<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
            '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>')
GOTHIC = '<w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/></w:rPr>'

#: key -> (sdtPr kind elements, the content's runs, the text Word types over, its replacement).
CONTROL_CASES: dict[str, tuple[str, str, str | None, str | None]] = {
    "Plain": ("<w:text/>", "Plain value", "Plain value", "Plain typed"),
    "Rich": ("", "Rich value", "Rich value", "Rich typed"),
    "Drop": ('<w:dropDownList><w:listItem w:displayText="Alpha" w:value="A"/>'
             '<w:listItem w:displayText="Beta" w:value="B"/></w:dropDownList>', "Alpha", "Alpha", "Beta"),
    "Combo": ('<w:comboBox><w:listItem w:displayText="Red" w:value="R"/><w:listItem w:displayText="Green" '
              'w:value="G"/></w:comboBox>', "Red", "Red", "Purple"),
    "Date": ('<w:date w:fullDate="2026-10-04T00:00:00Z"><w:dateFormat w:val="d-M-yyyy"/><w:lid w:val="en-GB"/>'
             '<w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>', "4-10-2026",
             "4-10-2026", "5-10-2026"),
    "Check": (CHECKBOX, "\u2610", None, None),
    "Bound": (f'<w:dataBinding w:prefixMappings="xmlns:ns0=\'urn:docx-agent:e5\'" '
              f'w:xpath="/ns0:root[1]/ns0:name[1]" w:storeItemID="{STORE_ID}"/><w:text/>', "Stale cache", None, None),
    "BoundTyped": (f'<w:dataBinding w:prefixMappings="xmlns:ns0=\'urn:docx-agent:e5\'" '
                   f'w:xpath="/ns0:root[1]/ns0:when[1]" w:storeItemID="{STORE_ID}"/><w:text/>', "2026-10-04",
                   "2026-10-04", "Typed into bound"),
}


def controls_document() -> bytes:
    ids = Ids(0x3C000000)
    body = para(ids, "Controls follow.")
    for k, (key, (kind, text, _, _)) in enumerate(CONTROL_CASES.items(), start=1):
        rpr = GOTHIC if key == "Check" else ""
        body += (f'<w:p w14:paraId="{ids()}" w14:textId="77777777"><w:r><w:t xml:space="preserve">{key}: </w:t></w:r>'
                 f'<w:sdt><w:sdtPr>{rpr}<w:id w:val="{1000 + k}"/><w:tag w:val="{key}"/>{kind}</w:sdtPr>'
                 f"<w:sdtContent><w:r>{rpr}<w:t>{text}</w:t></w:r></w:sdtContent></w:sdt>"
                 '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>')
    # A block-level rich text control of two paragraphs.
    body += (f'<w:sdt><w:sdtPr><w:id w:val="1100"/><w:tag w:val="Block"/></w:sdtPr><w:sdtContent>'
             f'{para(ids, "Block one.")}{para(ids, "Block two.")}</w:sdtContent></w:sdt>')
    body += para(ids, "The end.")
    return package(body,
                   extra_parts={"customXml/item1.xml": BOUND_XML, "customXml/itemProps1.xml": ITEM_PROPS,
                                "customXml/_rels/item1.xml.rels": relationships(
                                    [("rId1", REL + "customXmlProps", "itemProps1.xml", False)])},
                   extra_types={"customXml/itemProps1.xml": "application/vnd.openxmlformats-officedocument."
                                                            "customXmlProperties+xml"},
                   document_rels=[("rIdXml", REL + "customXml", "../customXml/item1.xml", False)])


def controls_script() -> list[str]:
    lines = []
    for k, (key, (_, _, old, new)) in enumerate(CONTROL_CASES.items(), start=2):
        if old is None:
            continue
        lines.append(f'execute find (find object of text object of paragraph {k} of d) find text "{old}" '
                     f'replace with "{new}" replace replace one')
    return lines


# -- reading ---------------------------------------------------------------------------------

_DROP = re.compile(r' w:rsid\w*="[^"]*"')


def normalise(xml: str) -> str:
    xml = re.sub(r' xmlns:\w+="[^"]*"', "", xml)
    # Word's binary drawing cache for older readers: not a fact about the XML.
    xml = re.sub(r' o:gfxdata="[^"]*"', ' o:gfxdata="{gfxdata}"', xml)
    xml = _DROP.sub("", xml)
    xml = re.sub(r'w:author="[^"]*"', 'w:author="{author}"', xml)
    xml = re.sub(r'w:date="[^"]*"', 'w:date="{date}"', xml)
    xml = re.sub(r'(w16du):dateUtc="[^"]*"', r'\1:dateUtc="{date}"', xml)
    xml = re.sub(r'w15:author="[^"]*"', 'w15:author="{author}"', xml)
    xml = re.sub(r'w15:userId="[^"]*"', 'w15:userId="{user}"', xml)
    xml = re.sub(r'w:initials="[^"]*"', 'w:initials="{initials}"', xml)
    return xml


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def blocks(data: bytes) -> list[str]:
    root = etree.fromstring(parts(data)["word/document.xml"])
    return [normalise(etree.tostring(child, encoding="unicode")) for child in root.find(_W + "body")]


_KEEP = re.compile(r"(word/(settings|people|comments\w*)\.xml|customXml/item\d+\.xml)$")


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
    styles_root = etree.fromstring(files["word/styles.xml"])
    entry["styles"] = sorted(s.find(_W + "name").get(_W + "val") for s in styles_root.findall(_W + "style"))
    return entry


RUNS = {
    "tables": (lambda: _script(tables_script()), tables_document),
    "tables-tracked": (lambda: _script(tables_script(), track=True), tables_document),
    "merges": (lambda: _script(merges_script()), merges_document),
    "merges-tracked": (lambda: _script(merges_script(), track=True), merges_document),
    "mergeforms-accept": (lambda: _script(["accept all revisions d"]), mergeforms_document),
    "mergeforms-reject": (lambda: _script(["reject all revisions d"]), mergeforms_document),
    "drawings": (lambda: _script(drawings_script()), drawings_document),
    "drawings-tracked": (lambda: _script(drawings_script(), track=True), drawings_document),
    "controls": (lambda: _script(controls_script()), controls_document),
    "controls-tracked": (lambda: _script(controls_script(), track=True), controls_document),
    "controls-saved": (lambda: _script([]), controls_document),
}


def run(names: list[str]) -> dict[str, bytes]:
    out = {}
    with oracle.session() as session:
        for name in names:
            script, document = RUNS[name]
            outcome = session.run_script(script(), document(), name=f"e5-{name}", tag="saved", timeout=200)
            if not outcome:
                raise SystemExit(f"Word did not save {name}: {outcome.outcome} {outcome.detail}")
            out[name] = outcome.path.read_bytes()
    return out


def main(names: list[str]) -> None:
    saved = run(names or list(RUNS))
    if names:
        for name, data in saved.items():
            print(name, json.dumps(observe(data), indent=1)[:200000])
        return
    OBSERVATIONS.write_text(json.dumps({
        "_about": ("What Word 16.106 for Mac wrote for tools/e5_probe.py: each probe's body blocks and its "
                   "settings, people and comments parts, custom XML items, relationships, content types and "
                   "style names. Authors and dates replaced by placeholders, rsids dropped. A step Word "
                   "refused is listed under 'failed'. Regenerate with python tools/e5_probe.py."),
        **{name: observe(data) for name, data in saved.items()}}, indent=1, ensure_ascii=False) + "\n",
        encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main(sys.argv[1:])
