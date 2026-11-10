"""A table fitted to the window (round 48): ``word_format_table`` ``width_of: "window"`` --
Word's AutoFit to Window, measured with Word 16.106 (``test_word_fits_tables_to_the_window_so``):

* ``w:tblW`` 5000 ``pct``; no indent (the table's own ``w:tblInd`` removed; ``0`` where its
  style indents it); the layout autofit; every cell's ``w:tcW`` in percent;
* the grid: the room less the outer halves of the outer borders in mode 15 (10,338 of a
  10,348-twip text width with borders of ``w:sz`` 4; 10,348 with none), the room and the
  first and last cells' margins below it (10,564 with margins of 108);
* the room: the section's text width; one column's in a section of columns; the cell's text
  width for a table in a cell.

docx2svg lays the result out where Word does: a 517.4 pt text width gives a 517.4 pt table
at the margin.  Word re-saves it unchanged (``test_word_keeps_our_fitted_tables``).

Tool design rules: 1 (``width_percent`` is the share of the window; ``width`` with it is
refused naming the way), 3 (a table that already fits changes nothing, said so), 6 (Word's
own form; the style's margins and borders decide the grid).  Made-up text.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from ooxml_edit.tools import Toolbox

_BORDERS = ('<w:tblBorders>' + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            for s in ("top", "left", "bottom", "right", "insideH", "insideV")) + '</w:tblBorders>')
_PAGE = ('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="779" w:bottom="1134" w:left="779" '
         'w:header="708" w:footer="708" w:gutter="0"/>')


def _styles(borders: bool) -> str:
    return ('<w:style w:type="table" w:styleId="ProbeGrid"><w:name w:val="Probe Grid"/><w:basedOn w:val="TableNormal"/>'
            f'<w:tblPr>{_BORDERS if borders else ""}</w:tblPr></w:style>'
            '<w:style w:type="table" w:styleId="ProbeWide"><w:name w:val="Probe Wide"/><w:basedOn w:val="TableNormal"/>'
            f'<w:tblPr><w:tblInd w:w="360" w:type="dxa"/>{_BORDERS}<w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
            '<w:left w:w="216" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="216" w:type="dxa"/>'
            '</w:tblCellMar></w:tblPr></w:style>')


def _body() -> str:
    """Tables in a 517.4 pt text width (A4, margins of 779 twips): one in the style's
    defaults; one whose style has cell margins of 10.8 pt and an indent; one 488.6 pt wide at
    an indent of 28.8 pt, fixed; one indented holding a table; then, in a section of two
    columns, one more."""
    n = iter(range(0x30000001, 0x30001000))

    def p(text: str, properties: str = "") -> str:
        run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
        return f'<w:p w14:paraId="{next(n):08X}" w14:textId="77777777">{properties}{run}</w:p>'

    def table(style, widths, width='<w:tblW w:w="0" w:type="auto"/>', extra="", label="T", first=None):
        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
        out = (f'<w:tbl><w:tblPr><w:tblStyle w:val="{style}"/>{width}{extra}<w:tblLook w:val="04A0"/></w:tblPr>'
               f'<w:tblGrid>{grid}</w:tblGrid>')
        for r in range(2):
            out += f'<w:tr w14:paraId="{next(n):08X}" w14:textId="77777777">'
            for c, w in enumerate(widths):
                inner = first if (first and r == 0 and c == 0) else p(f"{label} r{r}c{c}")
                out += f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{inner}</w:tc>'
            out += '</w:tr>'
        return out + '</w:tbl>'

    return "".join([
        p("Before one."),
        table("ProbeGrid", [2880, 2880], label="A"),
        p("Between one."),
        table("ProbeWide", [2400, 2400, 2400], label="B"),
        p("Indented body text.", '<w:pPr><w:ind w:left="576"/></w:pPr>'),
        table("ProbeGrid", [4886, 4886], '<w:tblW w:w="9772" w:type="dxa"/>',
              '<w:tblInd w:w="576" w:type="dxa"/><w:tblLayout w:type="fixed"/>', label="C"),
        p("Indented after.", '<w:pPr><w:ind w:left="576"/></w:pPr>'),
        table("ProbeGrid", [3000, 3000], extra='<w:tblInd w:w="300" w:type="dxa"/>', label="D",
              first=table("ProbeGrid", [1000, 1000], label="N") + p("")),
        p("End of section one."),
        f'<w:p w14:paraId="{next(n):08X}" w14:textId="77777777"><w:pPr><w:sectPr>{_PAGE}<w:cols w:space="708"/>'
        '</w:sectPr></w:pPr></w:p>',
        table("ProbeGrid", [2000, 2000], label="E"),
        p("Column text " * 30)])


def document_bytes(*, mode: int = 15, borders: bool = True) -> bytes:
    data = Document.new().to_bytes()
    source = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for item in source.infolist():
            raw = source.read(item.filename)
            if item.filename == "word/document.xml":
                xml = raw.decode()
                start = xml.index("<w:body>") + len("<w:body>")
                raw = (xml[:start] + _body() + f'<w:sectPr><w:type w:val="continuous"/>{_PAGE}'
                       '<w:cols w:num="2" w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body>'
                       '</w:document>').encode()
            elif item.filename == "word/styles.xml":
                raw = raw.decode().replace("</w:styles>", _styles(borders) + "</w:styles>").encode()
            elif item.filename == "word/settings.xml" and mode != 15:
                raw = raw.replace(b'w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" '
                                  b'w:val="15"', b'w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/'
                                  b'office/word" w:val="%d"' % mode)
            archive.writestr(item, raw)
    return out.getvalue()


def tables(data: bytes) -> list[str]:
    """Each table's properties (style options left out) and grid, in document order, and
    then its cells' widths."""
    xml = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml").decode()
    xml = re.sub(r' w(?:14)?:(?:rsid\w*|paraId|textId)="[^"]*"', "", xml)
    out = [re.sub(r"<w:tblLook[^>]*/>", "", m.group(1)) for m in re.finditer(r"<w:tblPr>(.*?)</w:tblGrid>", xml)]
    return out + [" ".join(re.findall(r"<w:tcW [^>]*/>", xml))]


def _grid(*widths: int) -> str:
    return "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)


_PCT = '<w:tblW w:w="5000" w:type="pct"/>'
_ZERO = '<w:tblInd w:w="0" w:type="dxa"/>'


def _fitted(two: int, three: tuple, column: int) -> list[str]:
    """Word's tables after AutoFit to Window on every top-level table (the nested one not)."""
    plain = f'<w:tblStyle w:val="ProbeGrid"/>{_PCT}</w:tblPr>{_grid(two, two)}'
    return [plain, f'<w:tblStyle w:val="ProbeWide"/>{_PCT}{_ZERO}</w:tblPr>{_grid(*three)}', plain, plain,
            f'<w:tblStyle w:val="ProbeGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>{_grid(1000, 1000)}',
            f'<w:tblStyle w:val="ProbeGrid"/>{_PCT}</w:tblPr>{_grid(column, column)}',
            " ".join(['<w:tcW w:w="2500" w:type="pct"/>'] * 4 + ['<w:tcW w:w="1667" w:type="pct"/>'] * 6
                     + ['<w:tcW w:w="2500" w:type="pct"/>'] * 5 + ['<w:tcW w:w="1000" w:type="dxa"/>'] * 4
                     + ['<w:tcW w:w="2500" w:type="pct"/>'] * 7)]


#: What Word 16.106 saved after AutoFit to Window (``feedback-round48/probe-C``).
WORD = {
    "mode 15": _fitted(5169, (3446, 3447, 3445), 2405),
    "no borders": _fitted(5174, (3446, 3447, 3445), 2410),
    "mode 14": _fitted(5282, (3594, 3594, 3592), 2518),
}
OPTIONS = {"mode 15": {}, "no borders": {"borders": False}, "mode 14": {"mode": 14}}
#: The nested table alone fitted (its cell: 3,000 twips less margins of 108 and 108).
WORD_NESTED = f'<w:tblStyle w:val="ProbeGrid"/>{_PCT}</w:tblPr>{_grid(1387, 1387)}'


def fitted(data: bytes, which: list[int] | None = None, **extra) -> tuple[bytes, list]:
    """Tables ``which`` (indexes among the top-level ones; ``None``: the nested one) fitted
    through the tool; the results."""
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session()
        session.open(data, "tables.docx")
        found = session.entry("d1").document
        top = [b["id"] for b in found.state()["blocks"] if b["id"].startswith("t:")]
        targets = [top[k] for k in which] if which is not None else [
            t.id for t in found.tables() if t.id not in top]
        results = [box.dispatch(session, "word_format_table", {"doc": "d1", "table": t, "scope": "table",
                                                               "width_of": "window", **extra}) for t in targets]
        return found.to_bytes(), results


@pytest.mark.parametrize("case", sorted(WORD))
def test_a_table_fitted_to_the_window_is_what_word_writes(case):
    data, results = fitted(document_bytes(**OPTIONS[case]), [0, 1, 2, 3, 4])
    assert all(r.ok for r in results), [r.error for r in results if not r.ok]
    assert tables(data) == WORD[case]


def test_a_nested_table_fits_its_cell():
    data, results = fitted(document_bytes(), None)
    assert results and all(r.ok for r in results)
    assert tables(data)[4] == WORD_NESTED


def test_docx2svg_lays_the_fitted_table_out_at_the_text_width():
    """The case of a 517.4 pt text width: a 488.6 pt table at an indent of 28.8 pt is drawn,
    once fitted, from the margin across all 517.4 pt (outer border edges)."""
    from docx2svg import parse_package
    from docx2svg.model import Table
    from docx2svg.table import grid_lines, outer_part, resolve_table

    def edges(data: bytes, k: int, room: int) -> tuple[float, float]:
        model = parse_package(data)
        table = [b for b in model.body if isinstance(b, Table)][k]
        resolved = resolve_table(model, table)
        lines = grid_lines(model, resolved, room)
        first, last = resolved.rows[0].cells[0], resolved.rows[0].cells[-1]
        return (float(lines[0] - outer_part(first.left_edge[0])) / 20,
                float(lines[-1] + outer_part(last.right_edge[0])) / 20)

    before = document_bytes()
    after, _ = fitted(before, [0, 1, 2, 3, 4])
    assert edges(before, 2, 10348) == (28.8, 517.9)
    for k in range(4):
        assert edges(after, k, 10348) == (0.0, 517.4)
    assert edges(after, 4, 4820) == (0.0, 241.0)     # one of two columns: (10,348 - 708) / 2 twips


def test_a_share_of_the_window():
    data, results = fitted(document_bytes(), [0], width_percent=50)
    assert results[0].ok
    first = tables(data)[0]
    # Each grid line at the floor of its share of the half: 5,169 twips.
    assert '<w:tblW w:w="2500" w:type="pct"/>' in first and _grid(2584, 2585) in first
    assert tables(data)[-1].startswith('<w:tcW w:w="1250" w:type="pct"/> <w:tcW w:w="1250" w:type="pct"/>')


def test_a_table_that_fits_already_is_not_changed():
    data, _ = fitted(document_bytes(), [2])
    again, results = fitted(data, [2])
    assert results[0].ok and results[0].changed == [] and results[0].summary.startswith("No change")
    assert any("already fits the window" in w for w in results[0].warnings)
    assert again == data or tables(again) == tables(data)


def test_a_width_with_the_window_is_refused_naming_the_way():
    _, results = fitted(document_bytes(), [0], width=300)
    assert not results[0].ok and results[0].error.field == "width_of"
    assert "width_percent" in results[0].error.message


def test_fitting_tracked_is_a_property_change_word_can_reject():
    document = Document.open(document_bytes())
    table = [b["id"] for b in document.state()["blocks"] if b["id"].startswith("t:")][2]
    before = tables(document.to_bytes())
    with document.tracking("Reviewer"):
        document.fit_table_to_window(table)
    xml = zipfile.ZipFile(io.BytesIO(document.to_bytes())).read("word/document.xml").decode()
    assert "<w:tblPrChange" in xml and "<w:tcPrChange" in xml and "<w:tblGridChange" in xml
    document.reject_all()
    assert tables(document.to_bytes()) == before


# -- Word itself -----------------------------------------------------------------------------


def _oracle():
    oracle = pytest.importorskip("oracle")
    if not oracle.available():
        pytest.skip("Microsoft Word or docx2svg's oracle scripts are not available")
    return oracle


@pytest.mark.oracle
@pytest.mark.parametrize("case", sorted(WORD))
def test_word_fits_tables_to_the_window_so(case):
    """Word itself: AutoFit to Window on each top-level table (the measurement of :data:`WORD`)."""
    oracle = _oracle()
    script = '''
on run argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file (item 1 of argv))
      set d to active document
      repeat with k from 1 to (count of tables of d)
        auto fit behavior (table k of d) behavior auto fit window
      end repeat
      save as d file name (item 2 of argv) file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
end run
'''
    try:
        with oracle.session() as word:
            out = word.run_script(script, document_bytes(**OPTIONS[case]), name="r48-window", tag="word")
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))
    assert out.ok, out.detail
    assert tables(out.path.read_bytes()) == WORD[case]


@pytest.mark.oracle
@pytest.mark.parametrize("case", sorted(WORD))
def test_word_keeps_our_fitted_tables(case):
    """Word opens ours without a prompt and its re-save keeps every table as we wrote it."""
    oracle = _oracle()
    data, _ = fitted(document_bytes(**OPTIONS[case]), [0, 1, 2, 3, 4])
    try:
        with oracle.session() as word:
            out = word.resave(data, name=f"r48-window-ours-{case.replace(' ', '-')}")
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))
    assert out.ok, (out.outcome, out.detail)
    assert tables(out.path.read_bytes()) == tables(data)
