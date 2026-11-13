"""Round 54 follow-up (lane A): docx-agent's own accept gives Word's Accept All at a table's
edge, in a clean file.  Made-up documents only.

A paragraph mark deleted at a table's edge, accepted (``word_changes`` accept: all, by id or
within a block), as Word 16.106's Accept All leaves it (measured by lanes J and R:
``feedback-round54/probe-J``, ``probe-R/word/rows``): a paragraph before a table joins the
first cell's first paragraph; a cell whose last mark is deleted takes the next cell of its
row in (a row's first cell's other paragraphs leave the table before the row, the table
split there in a later row; the row ends short where another row keeps the column, with
``w:gridAfter``/``w:wAfter`` as Word saves it).  A row's last cell: Word's screen loses it,
its saved file holds the text in a broken table (``gridSpan`` 0); ours, by user decision
(2026-11-13), a clean file: the text moves after the row in paragraphs of its own, the table
split there, the result saying so.  Reject stays the original.

After our accept ``word_read``'s final view is the one before it, and docx2svg's parse of
the accepted file is its final view of the pending one, wherever Word's result is clean.

Tool-design rules: 7 (the moved text is said in the result, and check's
``accept_loses_text`` says what our accept does), 8 (the accepted file is what Word's Accept
All shows, measured), 3 (an accept that changes nothing stays a no-op).
"""

from __future__ import annotations

import pytest
from docx2svg.parse.document import parse_package
from ooxml_edit.tools import Toolbox
from test_rows_joins_r54r import CLOCK, SHAPES, _document, _edge, _ours, _p, _render, _tbl, call, ok

from docx_agent import Document
from docx_agent.layout_facts import accept_loses_text
from docx_agent.markdown.read import Reader
from docx_agent.tools import FORMAT, GROUPS, TOOLS

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W

CLEAN = [s for s in SHAPES if s not in ("a row's end", "the table's last cell")]


@pytest.fixture(scope="module")
def box():
    with Toolbox(list(TOOLS), formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _read(document: Document) -> list:
    return _ours(Reader(document).story(document.package.document_part()))


def _accepted(data: bytes) -> Document:
    document = Document.open(data)
    document.accept_all()
    return Document.open(document.to_bytes())


def _body(document: Document):
    return document.package.tree(document.package.document_part()).find(_W + "body")


def _blocks(document: Document) -> list[str]:
    return [child.tag.rpartition("}")[2] for child in _body(document) if child.tag != _W + "sectPr"]


def _grid(table) -> list[str]:
    return [c.get(_W + "w") for c in table.findall(f"{_W}tblGrid/{_W}gridCol")]


@pytest.mark.parametrize("shape", CLEAN)
def test_after_our_accept_the_final_view_is_the_one_before_it(shape):
    data = _document(SHAPES[shape])
    accepted = _accepted(data)
    assert _read(accepted) == _read(Document.open(data))
    # ... and render's (docx2svg's parse of the accepted file: its final view before).
    assert _render(parse_package(accepted.to_bytes()).body) == _render(parse_package(data).body)
    assert accepted.revisions() == []


def test_a_paragraph_before_a_table_joins_the_first_cells_first_paragraph():
    data = _document(_p("Before. ", True) + _tbl([[_p("A1").replace("<w:p>", '<w:p><w:pPr><w:jc w:val="right"/>'
                                                                   "</w:pPr>"), _p("B1")]]))
    accepted = _accepted(data)
    assert _blocks(accepted) == ["tbl", "p"]
    first = _body(accepted).find(f"{_W}tbl/{_W}tr/{_W}tc/{_W}p")
    assert "".join(t.text for t in first.iter(_W + "t")) == "Before. A1"
    assert first.find(f"{_W}pPr/{_W}jc").get(_W + "val") == "right"     # the cell paragraph's properties


def test_a_first_cells_end_takes_the_next_cell_in_and_its_other_paragraphs_leave_the_table():
    accepted = _accepted(_document(_edge((0, 0))))
    assert _blocks(accepted) == ["p", "tbl", "p"]
    table = _body(accepted).find(_W + "tbl")
    first, second = table.findall(_W + "tr")
    assert len(first.findall(_W + "tc")) == 1 and len(second.findall(_W + "tc")) == 2
    # The row ends short (row 2 keeps the column), as Word saves it.
    assert first.find(f"{_W}trPr/{_W}gridAfter").get(_W + "val") == "1"
    assert first.find(f"{_W}trPr/{_W}wAfter").get(_W + "w") == "3000"
    assert first.find(f"{_W}tc/{_W}tcPr/{_W}tcW").get(_W + "w") == "3000"   # its own width kept
    assert _grid(table) == ["3000", "3000"]


def test_in_a_later_row_the_table_splits_each_part_with_its_own_grid():
    accepted = _accepted(_document(_edge((1, 0))))
    assert _blocks(accepted) == ["tbl", "p", "tbl", "p"]
    upper, lower = _body(accepted).findall(_W + "tbl")
    assert _grid(upper) == ["3000", "3000"] and _grid(lower) == ["3000"]
    assert lower.find(f"{_W}tr/{_W}trPr/{_W}gridAfter") is None
    assert _read(accepted)[1] == ("p", "Kept.")


def test_a_middle_cell_takes_in_the_rows_last_and_the_grid_loses_its_column():
    accepted = _accepted(_document(_edge((0, 1), rows=1, columns=3)))
    table = _body(accepted).find(_W + "tbl")
    assert _grid(table) == ["3000", "3000"] and len(table.findall(f"{_W}tr/{_W}tc")) == 2
    assert table.find(f"{_W}tr/{_W}trPr") is None


@pytest.mark.parametrize("at,blocks", [((0, 1), ["tbl", "p", "p", "tbl", "p"]), ((1, 1), ["tbl", "p", "p", "p"])],
                         ids=["a row's end", "the table's last cell"])
def test_a_rows_last_cell_goes_and_its_text_moves_after_the_row_in_a_clean_file(at, blocks):
    accepted = _accepted(_document(_edge(at)))
    assert _blocks(accepted) == blocks
    texts = [b.text for b in Reader(accepted).story(accepted.package.document_part()) if hasattr(b, "text")]
    assert "Kept." in texts and "Last." in texts
    # No broken table: every cell spans at least one column, every row reaches the grid.
    for table in _body(accepted).iter(_W + "tbl"):
        width = len(_grid(table))
        for row in table.findall(_W + "tr"):
            spans = sum(int(c.get(_W + "val")) for c in row.iter(_W + "gridSpan")) + \
                len([c for c in row.findall(_W + "tc") if c.find(f"{_W}tcPr/{_W}gridSpan") is None])
            after = row.find(f"{_W}trPr/{_W}gridAfter")
            assert spans + (int(after.get(_W + "val")) if after is not None else 0) == width
            assert all(int(c.get(_W + "val")) > 0 for c in row.iter(_W + "gridSpan"))
    # The row that lost its cell keeps the grid (Word keeps the table whole): it ends short.
    first = _body(accepted).find(_W + "tbl")
    assert _grid(first) == ["3000", "3000"]
    short = first.findall(_W + "tr")[at[0]]
    assert short.find(f"{_W}trPr/{_W}gridAfter").get(_W + "val") == "1"
    assert accept_loses_text(accepted) == []


@pytest.mark.parametrize("shape", list(SHAPES))
def test_reject_stays_the_original(shape):
    data = _document(SHAPES[shape])
    original = Document.open(data)
    rejected = Document.open(data)
    rejected.reject_all()
    assert rejected.revisions() == []
    assert _read(rejected) == _ours(Reader(original, view="original").story(original.package.document_part()))
    assert _blocks(rejected) == _blocks(original)


#: The chain with revision ids of its own (the shared shape gives both marks one w:id).
OWN_IDS = {**SHAPES, "chain into a nested table": _p("One. ", True) + _p("Two, two. ", True)
           + _tbl([[_tbl([[_p("Inner"), _p("Side")]]) + _p("Under"), _p("B1")]])}


@pytest.mark.parametrize("shape", list(SHAPES))
def test_accepting_by_id_or_within_the_table_gives_what_accept_all_gives(box, shape):
    data = _document(OWN_IDS[shape])
    expected = _read(_accepted(data))
    by_id = box.session(clock=CLOCK)
    by_id.open(data, "edge.docx")
    ids = [c["id"] for c in ok(box, by_id, "word_changes", doc="d1", action="list").data]
    ok(box, by_id, "word_changes", doc="d1", action="accept", ids=ids)
    assert _read(by_id.entry("d1").document) == expected
    within = box.session(clock=CLOCK)
    within.open(data, "edge.docx")
    table = within.entry("d1").document.tables()[0].id
    if shape.startswith(("before", "chain")):
        return  # the mark is before the table, not within it
    ok(box, within, "word_changes", doc="d1", action="accept", within=table)
    assert _read(within.entry("d1").document) == expected


def test_the_result_says_where_a_rows_last_cell_text_went_and_undo_brings_it_back(box):
    session = box.session(clock=CLOCK)
    session.open(_document(_edge((0, 1))), "edge.docx")
    before = ok(box, session, "word_read", doc="d1").data
    result = ok(box, session, "word_changes", doc="d1", action="accept", all=True)
    said = [w for w in result.warnings if "row's last cell" in w]
    assert said and "'Kept. Last.'" in said[0] and "Word's Accept All takes that text out of view" in said[0]
    assert result.to_json()["checks"]["validate"]["new"] == []
    read = ok(box, session, "word_read", doc="d1").data
    for identifier in said[0].split(" in ")[1].split(", the table")[0].split(", "):
        assert identifier.startswith("p:") and identifier in read
    ok(box, session, "word_undo", doc="d1")
    assert ok(box, session, "word_read", doc="d1").data == before


def test_check_and_the_table_warning_say_what_our_accept_does(box):
    session = box.session(clock=CLOCK)
    session.open(_document(_edge((0, 1))), "edge.docx")
    table = session.entry("d1").document.tables()[0].id
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                cells=[{"row": 1, "col": 0, "text": "r2c1 again"}])
    assert any("word_changes accept moves the text after the row" in w for w in result.warnings)
    refused = call(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                   cells=[{"row": 0, "col": 1, "text": "x"}])
    assert "word_changes accept moves the text after the row" in refused.to_json()["error"]["message"]


def test_a_cell_merged_down_by_a_pending_tracked_merge_reads_merged_in_the_final_view():
    # docx2svg's revision probe, "cellMerge alone": render draws the cell merged down (Word's
    # Accept All); word_read's final view now reads it so, as after our accept.
    merge = '<w:cellMerge w:id="7" w:author="Editor B" w:date="2026-11-13T09:00:00Z" w:vMerge="cont" ' \
            'w:vMergeOrig="rest"/>'
    top = ('<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>'
           + _p("a1") + "</w:tc><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>" + _p("b1") + "</w:tc></w:tr>")
    below = (f'<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>{merge}</w:tcPr>' + _p("a2")
             + "</w:tc><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>" + _p("b2") + "</w:tc></w:tr>")
    table = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>'
             f'<w:gridCol w:w="3000"/></w:tblGrid>{top}{below}</w:tbl>')
    data = _document(table)
    assert _read(Document.open(data)) == _read(_accepted(data))
    assert _read(Document.open(data))[0][1][1] == [[("p", "b2")]]


def test_an_accept_with_numbered_paragraphs_and_no_numbering_part_does_not_fail():
    numbered = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:rPr>' \
               '<w:del w:id="5" w:author="Editor B" w:date="2026-11-13T09:00:00Z"/></w:rPr></w:pPr>' \
               '<w:r><w:t>First</w:t></w:r></w:p>'
    document = Document.open(_document(numbered + _p("Second")))
    document.accept_all()
    assert [b.text for b in Reader(document).story(document.package.document_part())][0] == "FirstSecond"


def test_a_column_scale_counts_the_rows_as_accepted(box):
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "shelf.docx")
    ok(box, session, "word_insert_markdown", doc="d1", at="end",
       markdown="| Item | Grade |\n|---|---|\n| S-01 | 2 |\n| S-02 | 4 |\n| S-03 | 5 |\n")
    table = session.entry("d1").document.tables()[0].id
    ok(box, session, "word_set_tracking", doc="d1", on=True, author="Editor A")
    ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_row", index=2)
    scale = {"name": "Grade", "steps": [{"min": 1, "fill": "C62828", "label": "Low"},
                                        {"min": 4, "fill": "2E7D32", "label": "High"}],
             "targets": [f"{table}/col1"]}
    result = ok(box, session, "style_by_value", doc="d1", scale=scale, preview=True)
    # The pending deleted row (S-02, 4) is not among the values: rows as accepted.
    assert [(t["target"], t.get("value")) for t in result.data["targets"]] == [
        (f"{table}/c1,1", "2"), (f"{table}/c3,1", "5")]
