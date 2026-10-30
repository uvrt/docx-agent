"""A table's rows and keep with next (round 25): check's kept-row fact by Word's rule, a
table kept together (``word_format_table`` ``keep_together``) and what check says of it,
and a table's lines measured to their last character for ``past_margins``.

Word decides whether a row keeps with the next by the first paragraph of its first cell (a
nested table first in that cell: its first paragraph), measured in round 24 (docx2svg's
``tools/make_table_keep_probe.py``); docx-agent asks docx2svg's rule
(:func:`docx2svg.table.row_keeps_with_next`).  Made-up documents, Document.new()'s package.
"""

from __future__ import annotations

import random
import re

import pytest
from ooxml_edit.tools import Toolbox
from test_tools_layout_facts import FILLER, docx, facts, ok, p, with_style

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.layout import PX_PER_PT

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
KEEP = "<w:keepNext/>"
#: A paragraph style that keeps with the next, for the first column's labels.
LABEL_STYLE = ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="RowLabel"><w:name w:val="Row label"/>'
               '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr></w:style>')
LABEL = '<w:pStyle w:val="RowLabel"/>'


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


@pytest.fixture
def session(box):
    return box.session()


def _cell(content: str, width: int) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{content}</w:tc>'


def _nested(text: str, ppr: str) -> str:
    return ('<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/>'
            f'</w:tblGrid><w:tr>{_cell(p(text, ppr=ppr), 2000)}</w:tr></w:tbl>')


def grid(rows: int, first: str = "", second: str = "", *, nested: str | None = None, label: str = "Row",
         width: int = 9000, margins: bool = True) -> str:
    """A bordered two-column table of ``rows`` rows: the first cell's paragraph with
    ``first`` as its properties, the second's with ``second``; ``nested``: the first cell
    starts with a one-cell table whose paragraph has these properties."""
    half = width // 2
    borders = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    out = (f'<w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
           f'<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="{half}"/><w:gridCol w:w="{half}"/>'
           '</w:tblGrid>')
    for r in range(rows):
        lead = (_nested(f"Inner {r}", nested) if nested is not None else "") + p(f"{label} {r}", ppr=first)
        out += f"<w:tr>{_cell(lead, half)}{_cell(p(f'value {r}', ppr=second), half)}</w:tr>"
    return out + "</w:tbl>"


# -- W1: the kept-row fact by Word's rule ------------------------------------------------------


@pytest.mark.parametrize("case, keyword, reported", [
    ("first cell only", {"first": KEEP}, True),
    ("every cell", {"first": KEEP, "second": KEEP}, True),
    ("second cell only", {"second": KEEP}, False),
    ("nested table first", {"nested": KEEP}, True),
    ("outer paragraph after a nested table", {"nested": "", "first": KEEP, "second": KEEP}, False),
])
def test_a_split_after_a_kept_row_by_its_first_cells_first_paragraph(box, session, case, keyword, reported):
    """Seventy rows that keep are taller than a page: the table breaks after a kept row.
    Word keeps a row by its first cell's first paragraph alone (0.1.23 asked every
    paragraph of the row)."""
    d = session.open(docx(p("Tables") + grid(70, **keyword) + p("End")), "kept.docx")
    table = session.entry(d).document.tables()[0].id
    entries = [e for e in facts(box, session, d)["keep_with_next_at_foot"] if "table" in e]
    if not reported:
        assert entries == [], case
        return
    assert entries and {e["table"] for e in entries} == {table}, case
    assert all(e["next_row"] == e["row"] + 1 and e["next_page"] == e["page"] + 1 for e in entries)


# -- W2: a table kept together -----------------------------------------------------------------


def _keeps(document: Document, table: str) -> list:
    """Each row's first cell's first paragraph's direct w:keepNext: True, False or None."""
    element = document._table_entry(table)[1].element
    out = []
    for row in element.findall(_W + "tr"):
        cell = row.find(_W + "tc")
        first = next(child for child in cell if child.tag in (_W + "p", _W + "tbl"))
        while first.tag == _W + "tbl":
            first = next(child for child in first.find(_W + "tr").find(_W + "tc") if child.tag in (_W + "p", _W + "tbl"))
        node = first.find(f"{_W}pPr/{_W}keepNext")
        out.append(None if node is None else node.get(_W + "val") not in ("0", "false"))
    return out


def test_keep_together_where_a_style_keeps_the_rows(box, session):
    """A label style keeps every row's first paragraph: keep_together writes only the last
    row's release; check reports it as the table's, not as direct formatting, and offers
    no fix that clears it; off removes the release and nothing else."""
    data = with_style(docx(p("Before") + grid(6, LABEL) + p("After")), LABEL_STYLE)
    d = session.open(data, "label.docx")
    document = session.entry(d).document
    table = document.tables()[0].id
    assert document.table_properties(table)["keep_together"] is False     # every row keeps: a chain
    made = ok(box, session, "word_format_table", doc=d, table=table, scope="table", keep_together=True)
    table = made.renamed.get(table, table)                       # stamped by the edit
    assert table in made.changed and table == document.tables()[0].id
    assert _keeps(document, table) == [None] * 5 + [False]
    assert document.table_properties(table)["keep_together"] is True
    last = document.table(table).cell(5, 0).paragraphs[0].id
    layout = facts(box, session, d)
    assert layout["keep_together"] == [{"table": table, "pages": [1], "direct": [last]}]
    assert last not in {e["para"] for e in layout["direct_formatting"]}
    assert all(last not in fix.get("targets", []) for e in layout["direct_formatting"]
               for fix in e.get("fix", {}).get("calls", []))
    described = ok(box, session, "describe", doc=d).data["tables"]
    assert described[0]["keep_together"] is True

    ok(box, session, "word_format_table", doc=d, table=table, scope="table", keep_together=False)
    assert _keeps(document, table) == [None] * 6
    assert document.table_properties(table)["keep_together"] is False
    assert facts(box, session, d)["keep_together"] == []
    assert "keep_together" not in ok(box, session, "describe", doc=d).data["tables"][0]


def test_keep_together_on_plain_rows_and_off_removes_only_its_own(box, session):
    """No style keeps: the first cell's first paragraph of every row but the last gets a
    direct keep.  Off removes those and leaves a direct keep elsewhere in the row and the
    paragraphs' other direct formatting."""
    rows = grid(4, '<w:jc w:val="center"/>', KEEP)
    d = session.open(docx(p("Before") + rows + p("After")), "plain.docx")
    document = session.entry(d).document
    table = document.tables()[0].id
    ok(box, session, "word_format_table", doc=d, table=table, scope="table", keep_together=True)
    table = document.tables()[0].id
    assert _keeps(document, table) == [True, True, True, None]
    layout = facts(box, session, d)
    [kept] = layout["keep_together"]
    leads = [document.table(table).cell(r, 0).paragraphs[0].id for r in range(4)]
    assert kept["direct"] == leads[:3]
    centred = {e["para"]: e for e in layout["direct_formatting"]}
    for lead in leads:              # the centring is still direct formatting; the keep is not
        assert "keep_with_next" not in centred[lead]["props"]
        assert centred[lead]["fix"]["calls"][0]["clear"] == ["alignment"]

    ok(box, session, "word_format_table", doc=d, table=table, scope="table", keep_together=False)
    assert _keeps(document, table) == [None] * 4
    element = document._table_entry(table)[1].element
    assert len(element.findall(f".//{_W}jc")) == 4               # the centring stays
    assert len(element.findall(f".//{_W}keepNext")) == 4         # the second cells' keeps stay


def test_keep_together_mends_a_released_row_and_keeps_a_nested_tables_first_paragraph(box, session):
    """A row whose style keeps but a direct release stops: the release is removed.  A first
    cell that starts with a nested table is kept by that table's first paragraph."""
    body = grid(3, LABEL).replace(f'<w:pPr>{LABEL}</w:pPr><w:r><w:t xml:space="preserve">Row 1',
                                  f'<w:pPr>{LABEL}<w:keepNext w:val="0"/></w:pPr><w:r><w:t xml:space="preserve">Row 1')
    nested = grid(3, nested="", label="Outer")
    d = session.open(with_style(docx(p("Before") + body + p("Between") + nested + p("After")), LABEL_STYLE),
                     "mixed.docx")
    document = session.entry(d).document
    first, second = [t.id for t in document.tables()][:2]
    assert _keeps(document, first) == [None, False, None]
    ok(box, session, "word_format_table", doc=d, table=first, scope="table", keep_together=True)
    first, second = [t.id for t in document.tables()][:2]
    assert _keeps(document, first) == [None, None, False]
    ok(box, session, "word_format_table", doc=d, table=second, scope="table", keep_together=True)
    first, second = [t.id for t in document.tables()][:2]
    assert _keeps(document, second) == [True, True, None]
    element = document._table_entry(second)[1].element
    inner = [node for node in element.iter(_W + "tbl") if node is not element]
    assert all(node.find(f".//{_W}keepNext") is not None for node in inner[:2])
    assert not any(node.find(f"{_W}pPr/{_W}keepNext") is not None
                   for node in element.findall(f"{_W}tr/{_W}tc/{_W}p"))   # the outer paragraphs: none
    assert document.table_properties(second)["keep_together"] is True


def test_keep_together_refuses_what_it_cannot_do(box, session):
    d = session.open(docx(p("Before") + grid(1) + p("Between") + grid(3) + p("After")), "refuse.docx")
    one, three = [t.id for t in session.entry(d).document.tables()]
    for arguments, field in (({"table": one, "scope": "table", "keep_together": True}, None),
                             ({"table": three, "scope": "table", "keep_together": False}, None),
                             ({"table": three, "scope": "row", "row": 0, "keep_together": True}, "keep_together")):
        refused = box.dispatch(session, "word_format_table", {"doc": d, **arguments})
        assert not refused.ok, arguments
        if field:
            assert refused.error.field == field


def test_a_table_kept_together_does_not_take_the_next_block_along(box, session):
    """The label rows' chain takes the first table, the heading after it and the second
    table to the next page together, though the first fits where it is; kept together, the
    first stays (its last row does not keep with the heading) and the second goes on whole."""
    filler = "".join(p(f"{FILLER} {k}") for k in range(10))
    heading = p("Next part", ppr="<w:keepNext/>")
    body = filler + grid(10, LABEL) + heading + grid(14, LABEL, label="More") + p("End")
    d = session.open(with_style(docx(body), LABEL_STYLE), "foot.docx")
    document = session.entry(d).document

    def pages() -> list[list[int]]:
        layout = document.layout()
        return [sorted({w.page for r in range(len(t.rows)) for w in layout.where(t.cell(r, 0).paragraphs[0].id)})
                for t in document.tables()]

    assert pages() == [[2], [2]]
    for k in range(2):
        ok(box, session, "word_format_table", doc=d, table=document.tables()[k].id, scope="table",
           keep_together=True)
    assert pages() == [[1], [2]]
    assert len(facts(box, session, d)["keep_together"]) == 2


def test_a_table_kept_together_taller_than_a_page_is_warned_of_and_reported_once(box, session):
    """Word keeps it together all the same -- it starts a new page and breaks where that is
    full -- so the tool does it and warns; check says ``taller_than_page`` on the table's
    ``keep_together`` and no split after a kept row of it.  A table not kept together that
    splits after kept rows is reported as before; a short table kept together is not warned of."""
    body = (p("Before") + grid(70, LABEL) + p("Between", ppr="<w:pageBreakBefore/>") + grid(70, LABEL, label="Other")
            + p("Then", ppr="<w:pageBreakBefore/>") + grid(4) + p("End"))
    d = session.open(with_style(docx(body), LABEL_STYLE), "tall.docx")
    document = session.entry(d).document
    made = ok(box, session, "word_format_table", doc=d, table=document.tables()[0].id, scope="table",
              keep_together=True)
    tall, other, short = [t.id for t in document.tables()]
    assert made.warnings == [f"{tall} is taller than a page: Word starts it on a new page and breaks it where "
                             "the page is full"]
    assert ok(box, session, "word_format_table", doc=d, table=short, scope="table", keep_together=True).warnings == []
    tall, other, short = [t.id for t in document.tables()]
    layout = facts(box, session, d)
    kept = {e["table"]: e for e in layout["keep_together"]}
    assert kept[tall]["taller_than_page"] is True and "taller_than_page" not in kept[short]
    split = {e["table"] for e in layout["keep_with_next_at_foot"] if "table" in e}
    assert split == {other}


# -- W3: past the margins, to the last character -----------------------------------------------


def test_a_column_wide_table_whose_lines_end_in_spaces_is_not_past_the_margin(box, session):
    """A table as wide as the text column with no cell margins and no borders: its lines'
    text ends inside the cell; a line's last space, which Word lets run past the edge, is
    not drawn and is no fact.  Text pushed past the edge by an indent still is."""
    import io
    import zipfile

    section = zipfile.ZipFile(io.BytesIO(Document.new().to_bytes())).read("word/document.xml").decode()
    page = int(re.search(r'<w:pgSz w:w="(\d+)"', section).group(1))
    left, right = (int(re.search(rf'w:{side}="(\d+)"', section).group(1)) for side in ("left", "right"))
    width = page - left - right
    letters = random.Random(1)
    words = " ".join("".join(letters.choice("abcdefghijklmnopqrstuvwxyz") for _ in range(letters.randint(1, 9)))
                     for _ in range(500))
    margins = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'

    def wide(text: str, ppr: str = "") -> str:
        return (f'<w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/><w:tblLayout w:type="fixed"/>{margins}'
                f'</w:tblPr><w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid><w:tr>{_cell(p(text, ppr=ppr), width)}'
                "</w:tr></w:tbl>")

    d = session.open(docx(p("Before") + wide(words) + p("Between") + wide("Pushed out", '<w:ind w:right="-400"/>'
                                                                         '<w:jc w:val="right"/>') + p("After")),
                     "column.docx")
    first, second = [t.id for t in session.entry(d).document.tables()]
    page = session.entry(d).document.layout().pages[0]
    lines = [line for line in page.lines if re.match(r"^w:body/w:tbl\[1\]", line.path)]
    # Lines whose last space runs past the cell's edge, as 0.1.23 measured them.
    assert sum(max(span.end for span in line.spans) > page.column[1] + PX_PER_PT for line in lines) >= 3
    past = {e.get("table"): e for e in facts(box, session, d)["past_margins"]}
    assert first not in past
    assert past[second]["past"]["right"] > 15


def test_the_ink_of_a_line_leaves_out_its_spaces():
    from fractions import Fraction

    from docx2svg.layout import Span

    from docx_agent.layout_facts import _ink

    def span(chars: str, start: int) -> Span:
        xs = [Fraction(start + 5 * k) for k in range(len(chars))]
        return Span("w:p", list(chars), xs, 0, Fraction(start + 5 * len(chars)), "Face", False, False, 22)

    assert _ink([span(" ab ", 10)]) == (15, 25)
    assert _ink([span("ab", 0), span("  ", 10)]) == (0, 10)
    assert _ink([span("   ", 0)]) is None
