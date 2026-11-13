"""Round 54 (R110-R112): a table's structure changed with tracking on, read and drawn as it
will be once accepted, never merged over a pending change, and every change listed under
its author.  Made-up documents only.

* R110: a tracked column deletion (``w:cellDel`` on the column's cells, their content
  deleted) is left out of ``word_read``'s final view and of render's: the table as Word's
  Accept All leaves it (measured, round 54: the column gone, the others keeping their
  widths); the markup view keeps the cells, marked deleted.
* R111: with a tracked structure change pending in the table, a merge or split was
  refused; round 55 (R115) writes them tracked over it instead.
* R112: ``word_changes`` lists every tracked change with its author, a table's grid change
  (``w:tblGridChange``, which has no author in the schema, nor in Word's own writing)
  under the change that recorded it.
* A tracked merge writes a comment's or note's reference once (Word does not open a
  document with one written twice: measured, round 54).
"""

from __future__ import annotations

import datetime as dt
import re

import pytest
from lxml import etree
from ooxml_edit.tools import Toolbox

import docx2svg
from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
ME = "Editor A"
CLOCK = lambda: dt.datetime(2026, 11, 12, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
ROWS = [("Area", "Lead", "Phase", "Remark"), ("Harbour", "Ines", "draft", "awaits figures"),
        ("Orchard", "Tomas", "done", "signed off"), ("Quarry", "Aleid", "open", "two gaps")]
MARKDOWN = "\n".join(["| " + " | ".join(ROWS[0]) + " |", "|---|---|---|---|"]
                     + ["| " + " | ".join(r) + " |" for r in ROWS[1:]])


@pytest.fixture(scope="module")
def box():
    with Toolbox(list(TOOLS), formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def call(box, session, name, /, **arguments):
    return box.dispatch(session, name, arguments)


def ok(box, session, name, /, **arguments):
    result = call(box, session, name, **arguments)
    assert result.ok, (name, result.to_json())
    return result


def _opened(box, tracked: bool = True):
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "report.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", doc="d1", at=f"after:{first}",
       markdown=f"Field report.\n\n{MARKDOWN}\n\nAfter the table.")
    if tracked:
        ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    return session, session.entry("d1").document.tables()[-1].id


def _document(session) -> Document:
    return session.entry("d1").document


def _glyph_lines(data: bytes) -> list[tuple]:
    """Every line docx2svg lays out: its text and where each glyph starts (the drawing
    without the structural paths, which name the document's cells)."""
    layout = docx2svg.convert_docx_to_layout(data)
    out = []
    for page in layout._lines:
        for _, line in page:
            out.append(tuple((span.kind, tuple(span.chars), round(float(span.x), 3)) for span in line.spans
                             if span.kind == "text"))
    return out


def _svg(data: bytes) -> str:
    return re.sub(r' data-docx-path="[^"]*"', "", "".join(docx2svg.convert_docx_to_svg(data)))


# -- R110 ----------------------------------------------------------------------------------------


def test_a_tracked_column_deletion_reads_as_accepted(box):
    session, table = _opened(box)
    deleted = ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_column", index=2)
    # Rule 2: the result says what the reader now sees and what positions count (round 54
    # lane M: the table as accepted, test_tracked_merge_r54m.py).
    assert deleted.data["columns"] == 3 and deleted.data["columns_deleted_pending"] == [2]
    assert "| Area | Lead | Remark |" in deleted.data["markdown"]
    assert any("col positions count the table as accepted" in w and f"within {table}" in w
               for w in deleted.warnings)
    final = ok(box, session, "word_read", doc="d1").data
    assert "| Area | Lead | Remark |" in final and "| Quarry | Aleid | two gaps |" in final
    assert "Phase" not in final and "|  |" not in final
    markup = ok(box, session, "word_read", doc="d1", view="markup").data
    assert "{--Phase--}" in markup and "{--open--}" in markup
    original = ok(box, session, "word_read", doc="d1", view="original").data
    assert "| Area | Lead | Phase | Remark |" in original


def test_render_draws_a_tracked_column_deletion_as_accept_all_leaves_it(box):
    """Word's Accept All of this document: three columns of 2,268 twips, the table that much
    narrower (measured round 54); docx2svg draws the tracked document exactly so."""
    session, table = _opened(box)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_column", index=2)
    tracked = _document(session).to_bytes()
    accepted = Document.open(tracked)
    accepted.accept_all()
    grid = [c.get(_W + "w") for c in accepted.package.tree(accepted.package.document_part()).iter(_W + "gridCol")]
    assert grid == ["2268", "2268", "2268"]
    assert _svg(tracked) == _svg(accepted.to_bytes())


def test_an_inserted_cell_is_left_out_of_the_original_view(box):
    session, table = _opened(box)
    document = _document(session)
    part = document.package.document_part()
    cell = next(document.package.tree(part).iter(_W + "tc"))
    properties = cell.find(_W + "tcPr")
    properties.append(etree.Element(_W + "cellIns", {_W + "id": "900", _W + "author": ME,
                                                      _W + "date": "2026-11-12T09:00:00Z"}))
    document.package.mark_dirty(part)
    reopened = Document.open(document.to_bytes())
    assert "Area" not in reopened.to_markdown(view="original", ids=False)
    assert "Area" in reopened.to_markdown(view="final", ids=False)


# -- R111 ----------------------------------------------------------------------------------------


def test_a_merge_over_a_pending_column_deletion_is_written_tracked(box):
    """Round 55 (R115) replaces R111's refusal: the merge is written over the pending
    deletion (tests/test_tracked_reshape_r55.py holds the forms Word settles)."""
    session, table = _opened(box)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_column", index=2)
    merged = ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
                first={"row": 0, "col": 0}, last={"row": 0, "col": 1})
    assert merged.data["columns"] == 3
    changes = ok(box, session, "word_changes", doc="d1", action="list").data
    assert len(changes) == 2 and all(c["author"] == ME for c in changes)


def test_a_split_beside_a_pending_merge_is_written_tracked(box):
    session, table = _opened(box)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
       first={"row": 1, "col": 0}, last={"row": 1, "col": 1})
    ok(box, session, "word_edit_table", doc="d1", table=table, action="split", first={"row": 2, "col": 0})


def test_untracked_a_merge_is_not_refused(box):
    session, table = _opened(box, tracked=False)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
       first={"row": 0, "col": 0}, last={"row": 0, "col": 1})


# -- R112 ----------------------------------------------------------------------------------------


#: Every tracked edit of a table's structure and its formatting, through the tools.
SWEEP = [
    ("insert_row", {"action": "insert_row", "index": 1}),
    ("delete_row", {"action": "delete_row", "index": 2}),
    ("insert_column", {"action": "insert_column", "index": 1}),
    ("delete_column", {"action": "delete_column", "index": 2}),
    ("merge across", {"action": "merge", "first": {"row": 0, "col": 0}, "last": {"row": 0, "col": 1}}),
    ("merge down", {"action": "merge", "first": {"row": 1, "col": 0}, "last": {"row": 2, "col": 0}}),
    ("merge rectangle", {"action": "merge", "first": {"row": 1, "col": 1}, "last": {"row": 2, "col": 2}}),
    ("split", {"action": "split", "first": {"row": 1, "col": 1}, "columns": 3}),
    ("split down", {"action": "split", "first": {"row": 1, "col": 1}, "rows": 2}),
    ("set_cells", {"action": "set_cells", "cells": [{"row": 1, "col": 3, "text": "revised"}]}),
]
FORMATS = [
    ("table width", {"scope": "table", "width": 400}),
    ("column width", {"scope": "column", "column": 1, "width": 120}),
    ("cell shading", {"scope": "cell", "row": 1, "column": 1, "shading": "FFEE99"}),
    ("table style options", {"scope": "table", "look": ["header_row"]}),
]
#: Revision records the schema gives no author (CT_Markup: an id only); Word writes none.
NO_AUTHOR = {_W + "tblGridChange"}
RECORDS = {_W + name for name in ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange",
                                   "tblPrChange", "trPrChange", "tcPrChange", "tblGridChange", "numberingChange",
                                   "cellIns", "cellDel", "cellMerge")}


def _sweep_case(box, tool, arguments):
    session, table = _opened(box)
    ok(box, session, tool, doc="d1", table=table, **arguments)
    return session


@pytest.mark.parametrize("name, tool, arguments",
                         [(n, "word_edit_table", a) for n, a in SWEEP] + [(n, "word_format_table", a) for n, a in FORMATS])
def test_every_tracked_table_change_is_listed_under_its_author(box, name, tool, arguments):
    session = _sweep_case(box, tool, arguments)
    changes = ok(box, session, "word_changes", doc="d1", action="list").data
    assert changes, name
    assert [c["author"] for c in changes] == [ME] * len(changes), (name, changes)
    assert all(c["date"] for c in changes), name
    # Every record written carries the author and date, except the grid change, which
    # the schema gives none (as Word writes it).
    document = _document(session)
    root = document.package.tree(document.package.document_part())
    for record in root.iter(*RECORDS):
        if record.tag in NO_AUTHOR:
            assert record.get(_W + "author") is None
            continue
        assert record.get(_W + "author") == ME and record.get(_W + "date"), (name, record.tag)


def test_a_grid_change_from_another_document_without_its_partner_is_listed_as_found(box):
    """A grid change no other record of its table names an author for is listed as it is:
    author null (a document arriving that way is reported as-is)."""
    session, table = _opened(box, tracked=False)
    document = Document.open(_document(session).to_bytes())
    part = document.package.document_part()
    grid = next(document.package.tree(part).iter(_W + "tblGrid"))
    change = etree.SubElement(grid, _W + "tblGridChange", {_W + "id": "901"})
    old = etree.SubElement(change, _W + "tblGrid")
    for _ in range(4):
        etree.SubElement(old, _W + "gridCol", {_W + "w": "2000"})
    document.package.mark_dirty(part)
    session = box.session(clock=CLOCK)
    session.open(document.to_bytes(), "arrived.docx")
    changes = ok(box, session, "word_changes", doc="d1", action="list").data
    assert [(c["kind"], c["author"]) for c in changes] == [("table", None)]


# -- references in a tracked merge -----------------------------------------------------------------


def test_a_tracked_merge_writes_each_comment_reference_once(box):
    session, table = _opened(box, tracked=False)
    ok(box, session, "word_comments", doc="d1", action="add", author="Other",
       items=[{"find": "Ines", "text": "Check the lead."}])
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
       first={"row": 1, "col": 0}, last={"row": 1, "col": 1})
    root = _document(session).package.tree(_document(session).package.document_part())
    for tag in ("commentReference", "commentRangeStart", "commentRangeEnd"):
        ids = [node.get(_W + "id") for node in root.iter(_W + tag)]
        assert len(ids) == len(set(ids)), (tag, ids)
    # Accepting keeps the comment, in the merged cell.
    accepted = Document.open(_document(session).to_bytes())
    accepted.accept_all()
    assert "Check the lead." in accepted.to_markdown(view="markup")


def test_a_deleted_cell_whose_content_is_kept_widens_the_cell_before_in_the_final_view(box):
    """Word's Accept All gives such a cell's columns to the cell before it, its content
    going (measured, round 54); the final view reads it so, as docx2svg draws it."""
    session, table = _opened(box, tracked=False)
    document = Document.open(_document(session).to_bytes())
    root = document.package.tree(document.package.document_part())
    row = list(root.iter(_W + "tr"))[1]
    cell = list(row.iter(_W + "tc"))[1]
    etree.SubElement(cell.find(_W + "tcPr"), _W + "cellDel",
                     {_W + "id": "903", _W + "author": ME, _W + "date": "2026-11-12T09:00:00Z"})
    document.package.mark_dirty(document.package.document_part())
    final = Document.open(document.to_bytes()).to_markdown(view="final", ids=False)
    assert '<td colspan="2"><p>Harbour</p></td>' in final and "Ines" not in final
