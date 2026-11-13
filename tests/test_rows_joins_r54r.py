"""Round 54 follow-up (lane R): row positions as accepted, the final view at a table's edge as
Word's Accept All leaves it, and the text Accept All loses.  Made-up documents only.

* **Rows as accepted** (user decision 2026-11-13, as lane M did for columns):
  ``word_edit_table``'s ``row``, ``row_label``, ``index`` (insert/delete row), ``first``/
  ``last``, a column fill's ``rows``, a ``from`` block's rows and ``word_format_table``'s
  ``row`` skip a row whose tracked deletion is pending; results count ``rows`` so and list
  ``rows_deleted_pending``; a warning names the row skipped; past the end is refused with
  the accepted shape.  An insert beside the row named keeps the deleted row on its other
  side (Word's Accept All and Reject All, measured: ``feedback-round54/probe-R``).  Cell ids
  stay on the document's grid.
* **The final view at a table's edge** (lane J measured Word 16.106's Accept All; docx2svg
  0.11.13 draws it): a paragraph whose mark is deleted before a table joins the first
  cell's first paragraph; a cell whose last mark is deleted takes in the next cell of its
  row (a row's first cell's other paragraphs leave the table, which splits there in a later
  row); a row's last cell goes with its text.  ``word_read`` reads what render draws (held
  here against docx2svg's own parse), and positions follow it.
* **check's ``accept_loses_text``**: the cell Accept All removes with its text, with the
  ``word_changes`` reject that keeps it; ``word_edit_table`` says it too.

Tool-design rules: 1 (positions read as the reader sees the table), 2 (refusals give the
accepted shape and the way out), 7 (what Accept All loses is a fact), 8 (the view follows
Word's Accept All, measured).
"""

from __future__ import annotations

import datetime as dt

import pytest
from docx2svg.model import Paragraph, Table
from docx2svg.parse.document import parse_package
from lxml import etree
from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool

from docx_agent import Document
from docx_agent.layout_facts import accept_loses_text
from docx_agent.markdown import to_markdown
from docx_agent.markdown.read import ParagraphRecord, Reader, TableRecord
from docx_agent.tools import FORMAT, GROUPS, TOOLS

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
ME = "Editor A"
CLOCK = lambda: dt.datetime(2026, 11, 13, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
ROWS = [("Area", "Lead", "Phase", "Remark"), ("Harbour", "Ines", "draft", "awaits figures"),
        ("Orchard", "Tomas", "done", "signed off"), ("Quarry", "Aleid", "open", "two gaps")]
MARKDOWN = "\n".join(["| " + " | ".join(ROWS[0]) + " |", "|---|---|---|---|"]
                     + ["| " + " | ".join(r) + " |" for r in ROWS[1:]])
SETS: dict[str, tuple[tuple, list[tuple]]] = {}


@tool("rows_put", "Keep a made-up dataset.", {"id": string("Its id.")}, documents=())
def rows_put(call, id):
    columns, rows = SETS[id]
    call.define_dataset(id, Dataset(columns, rows), source=None)
    return Result(summary=f"Kept d:{id}")


@pytest.fixture(scope="module")
def box():
    with Toolbox([*TOOLS, rows_put], formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def call(box, session, name, /, **arguments):
    return box.dispatch(session, name, arguments)


def ok(box, session, name, /, **arguments):
    result = call(box, session, name, **arguments)
    assert result.ok, (name, result.to_json())
    return result


def _deleted_row(box, deleted: int = 2):
    """The report table, row ``deleted`` deleted as a tracked change (tracking left on)."""
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "report.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", doc="d1", at=f"after:{first}",
       markdown=f"Field report.\n\n{MARKDOWN}\n\nAfter the table.")
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    table = session.entry("d1").document.tables()[-1].id
    ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_row", index=deleted)
    return session, session.entry("d1").document.tables()[-1].id


def _texts(session, table):
    t = session.entry("d1").document.table(table)
    return [[t.cell(r, c).text for c in range(t.column_count)] for r in range(len(t.rows))]


def _first_cells(session, table):
    """Each document row's first cell's text, deleted text included."""
    entry = session.entry("d1").document._table_entry(table)[1]
    return ["".join(n.text or "" for n in row.find(_W + "tc").iter(_W + "t", _W + "delText"))
            for row, _ in entry.rows]


# -- rows as accepted ---------------------------------------------------------------------------


def test_row_positions_skip_a_row_whose_deletion_is_pending(box):
    session, table = _deleted_row(box)
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                cells=[{"row": 2, "col": 3, "text": "three gaps"}])
    assert result.data["rows"] == 3 and result.data["rows_deleted_pending"] == [2]
    assert _texts(session, table)[3][3] == "three gaps"            # Quarry, not the deleted Orchard
    deleted = session.entry("d1").document._table_entry(table)[1].rows[2][0]
    assert "signed off" in "".join(n.text for n in deleted.iter(_W + "delText"))   # left as it was
    assert any("document row 2 ('Orchard') has a tracked deletion" in w and "cell ids keep" in w
               for w in result.warnings)


def test_a_row_past_the_end_is_refused_with_the_accepted_shape(box):
    session, table = _deleted_row(box)
    result = call(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                  cells=[{"row": 3, "col": 0, "text": "x"}])
    error = result.to_json()["error"]
    assert error["code"] == "not_found" and "rows 0 to 2 as accepted" in error["message"]
    assert error["valid_options"] == ["0", "1", "2"]


def test_row_label_and_a_column_fill_count_the_rows_as_accepted(box):
    session, table = _deleted_row(box)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
       cells=[{"row_label": "Quarry", "col_label": "Phase", "text": "closed"}])
    assert _texts(session, table)[3][2] == "closed"
    # The deleted row's own text is not a label any more (word_changes reaches it).
    missed = call(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                  cells=[{"row_label": "Orchard", "col": 2, "text": "x"}])
    assert not missed.ok and "Orchard" not in missed.to_json()["error"]["valid_options"]
    filled = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                cells=[{"col": 1, "rows": "*", "text": "Noor"}])
    assert filled.data["filled"][0]["rows"] == "1..2" and filled.data["filled"][0]["cells"] == 2
    assert [r[1] for r in _texts(session, table)] == ["Lead", "Noor", "", "Noor"]   # the deleted row left


def test_delete_row_and_format_row_count_the_rows_as_accepted(box):
    session, table = _deleted_row(box)
    ok(box, session, "word_format_table", doc="d1", table=table, scope="row", row=2, shading="FFEEDD")
    element = session.entry("d1").document._table_entry(table)[1].rows[3][0]
    assert element.find(f".//{_W}shd") is not None
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_row", index=2)
    assert result.data["rows"] == 2 and result.data["rows_deleted_pending"] == [2, 3]


@pytest.mark.parametrize("index,side", [(1, "after"), (2, "before"), (None, None)],
                         ids=["after the row before", "before the row after", "default last"])
def test_an_insert_beside_a_pending_deleted_row_goes_beside_the_row_named(box, index, side):
    session, table = _deleted_row(box, deleted=2 if index is not None else 3)
    args = {k: v for k, v in (("index", index), ("side", side)) if v is not None}
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="insert_row", **args)
    rows = _first_cells(session, result.data["id"])
    if index is None:   # after the last row as accepted (Orchard), the deleted Quarry after it
        assert rows == ["Area", "Harbour", "Orchard", "", "Quarry"]
    elif side == "after":
        assert rows == ["Area", "Harbour", "", "Orchard", "Quarry"]
    else:
        assert rows == ["Area", "Harbour", "Orchard", "", "Quarry"]
    document = session.entry("d1").document
    accepted = Document.open(document.to_bytes())
    accepted.accept_all()
    assert len(accepted.tables()[-1].rows) == 4


def test_a_merge_over_a_pending_deleted_row_is_refused_with_the_accept(box):
    session, table = _deleted_row(box)
    ok(box, session, "word_set_tracking", doc="d1", on=False)
    result = call(box, session, "word_edit_table", doc="d1", table=table, action="merge",
                  first={"row": 1, "col": 0}, last={"row": 2, "col": 0})
    error = result.to_json()["error"]
    assert error["code"] == "refused" and "a row deleted" in error["message"]
    assert error["details"]["fix"]["calls"] == [{"action": "accept", "within": table}]


def test_a_from_block_fills_the_rows_as_accepted(box):
    session, table = _deleted_row(box)
    SETS["shelf"] = (("Area", "Lead"), [("Harbour", "Ines"), ("Quarry", "Aleid")])
    ok(box, session, "rows_put", id="shelf")
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells", **{"from": "d:shelf"})
    assert result.data["rows"] == 3 and result.data["from"]["first_row"] == 1
    assert [r[:2] for r in _texts(session, table)] == [["Area", "Lead"], ["Harbour", "Ines"], ["", ""],
                                                       ["Quarry", "Aleid"]]


def test_sync_and_describe_count_the_rows_as_accepted(box):
    session, table = _deleted_row(box)
    SETS["shelf2"] = (("Area", "Lead"), [("Harbour", "Ines"), ("Quarry", "Aleid")])
    ok(box, session, "rows_put", id="shelf2")
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="sync", **{"from": "d:shelf2"})
    assert result.data["rows"] == 3
    described = ok(box, session, "word_describe", doc="d1")
    assert [t["rows"] for t in described.data["tables"]] == [3]


# -- the final view at a table's edge -----------------------------------------------------------


def _p(text: str, deleted_mark: bool = False) -> str:
    mark = f'<w:pPr><w:rPr><w:del w:id="9{len(text)}" w:author="Reviewer B" w:date="2026-11-13T09:00:00Z"/>' \
           '</w:rPr></w:pPr>' if deleted_mark else ""
    return f"<w:p>{mark}<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>"


def _tc(content: str) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{content}</w:tc>'


def _tbl(rows: list[list[str]]) -> str:
    grid = "".join('<w:gridCol w:w="3000"/>' for _ in rows[0])
    return (f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>'
            + "".join("<w:tr>" + "".join(_tc(c) for c in row) + "</w:tr>" for row in rows) + "</w:tbl>")


def _edge(at: tuple[int, int], rows: int = 2, columns: int = 2) -> str:
    cells = [[_p(f"r{r + 1}c{c + 1}") for c in range(columns)] for r in range(rows)]
    cells[at[0]][at[1]] = _p("Kept.") + _p("Last.", deleted_mark=True)
    return _tbl(cells)


SHAPES = {
    "before a table": _p("Before. ", True) + _tbl([[_p("A1"), _p("B1")]]),
    "chain into a nested table": _p("One. ", True) + _p("Two. ", True)
    + _tbl([[_tbl([[_p("Inner"), _p("Side")]]) + _p("Under"), _p("B1")]]),
    "a first cell's end": _edge((0, 0)),
    "a later row's first cell": _edge((1, 0)),
    "a middle cell's end": _edge((0, 1), rows=1, columns=3),
    "a row's end": _edge((0, 1)),
    "the table's last cell": _edge((1, 1)),
}


def _document(body: str) -> bytes:
    document = Document.new()
    tree = document.package.tree(document.package.document_part())
    container = tree.find(_W + "body")
    section = container.find(_W + "sectPr")
    for child in list(container):
        if child is not section:
            container.remove(child)
    wrapped = etree.fromstring(f'<w:body xmlns:w="{W}">{body}{_p("End.")}</w:body>')
    for child in list(wrapped):
        section.addprevious(child)
    document.package.mark_dirty(document.package.document_part())
    return document.to_bytes()


def _ours(blocks):
    out = []
    for b in blocks:
        if isinstance(b, ParagraphRecord):
            out.append(("p", "".join(b.text.split())))
        elif isinstance(b, TableRecord):
            out.append(("t", [[_ours(c.records) for c in line] for line in b.rows]))
    return out


def _render(blocks):
    out = []
    for b in blocks:
        if isinstance(b, Paragraph):
            out.append(("p", "".join(b.text.split())))
        elif isinstance(b, Table):
            out.append(("t", [[_render(c) for c in line] for line in b.rows]))
    return out


@pytest.mark.parametrize("shape", list(SHAPES))
def test_the_final_view_reads_what_render_draws_at_a_tables_edge(shape):
    data = _document(SHAPES[shape])
    document = Document.open(data)
    ours = _ours(Reader(document).story(document.package.document_part()))
    assert ours == _render(parse_package(data).body)


def test_the_joins_read_as_accept_all_leaves_them():
    document = Document.open(_document(SHAPES["before a table"]))
    markdown = to_markdown(document, ids=False)
    assert "| Before. A1 | B1 |" in markdown and "\nBefore." not in markdown
    split = to_markdown(Document.open(_document(SHAPES["a later row's first cell"])), ids=True)
    # The paragraph leaves the table, which goes on after it as the same table.
    assert "Kept." in split.split("continued from row 1")[0] and "Last.r2c2" in split.split("continued")[1]
    gone = to_markdown(Document.open(_document(SHAPES["a row's end"])), ids=False)
    assert "Kept." not in gone and "Last." not in gone and "r1c1" in gone


def test_a_table_read_alone_takes_the_paragraph_joined_into_it(box):
    session = box.session(clock=CLOCK)
    session.open(_document(SHAPES["before a table"]), "edge.docx")
    table = session.entry("d1").document.tables()[0].id
    read = ok(box, session, "word_read", doc="d1", range=table)
    assert "Before. A1" in read.data


def test_positions_follow_the_cells_accept_all_takes_in_and_removes(box):
    session = box.session(clock=CLOCK)
    session.open(_document(_edge((0, 1), rows=1, columns=3)), "edge.docx")
    table = session.entry("d1").document.tables()[0].id
    read = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
              cells=[{"row": 0, "col": 0, "text": "r1c1"}])
    assert read.data["columns"] == 2         # the third cell is taken into the second
    past = call(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                cells=[{"row": 0, "col": 2, "text": "x"}])
    assert past.to_json()["error"]["valid_options"] == ["0", "1"]
    session = box.session(clock=CLOCK)
    session.open(_document(_edge((0, 1))), "edge.docx")
    table = session.entry("d1").document.tables()[0].id
    refused = call(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                   cells=[{"row": 0, "col": 1, "text": "lost"}])
    assert refused.to_json()["error"]["code"] == "refused" and "would be lost" in refused.to_json()["error"]["message"]


# -- check: accept_loses_text --------------------------------------------------------------------


@pytest.mark.parametrize("shape", ["a row's end", "the table's last cell"])
def test_check_names_the_text_accept_all_loses_with_the_reject_that_keeps_it(box, shape):
    session = box.session(clock=CLOCK)
    session.open(_document(SHAPES[shape]), "edge.docx")
    found = ok(box, session, "check", doc="d1", include=["layout"]).data["layout"]["accept_loses_text"]
    assert len(found) == 1
    entry = found[0]
    assert entry["text"] == "Kept. Last." and entry["author"] == "Reviewer B" and entry["change"].startswith("rev:")
    assert entry["cell"].endswith("/c0,1" if shape == "a row's end" else "/c1,1")
    assert entry["fix"] == {"tool": "word_changes", "calls": [{"action": "reject", "ids": [entry["change"]]}]}
    ok(box, session, "word_changes", doc="d1", **entry["fix"]["calls"][0])
    assert accept_loses_text(session.entry("d1").document) == []
    assert "Kept." in ok(box, session, "word_read", doc="d1").data


def test_no_loss_where_the_next_cell_takes_the_text_in():
    for shape in ("a first cell's end", "a middle cell's end", "before a table"):
        assert accept_loses_text(Document.open(_document(SHAPES[shape]))) == []


def test_an_edit_of_the_table_says_what_accept_all_loses(box):
    session = box.session(clock=CLOCK)
    session.open(_document(SHAPES["a row's end"]), "edge.docx")
    table = session.entry("d1").document.tables()[0].id
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                cells=[{"row": 1, "col": 0, "text": "r2c1 again"}])
    assert any("Word's Accept All drops this cell and its text from view" in w and "word_changes reject rev:" in w
               for w in result.warnings)
