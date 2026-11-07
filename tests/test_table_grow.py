"""``word_edit_table`` ``set_cells`` ``from`` with ``grow`` (round 40): a block of rows longer
than the table adds the rows it needs, each as ``insert_row`` makes one -- the last row's
properties and cells, not a header row -- inside the call's undo step; without ``grow`` the
refusal gives the call that works (``details.call``).  Made-up data only (the toy dataset of
:mod:`test_tables_from_data`)."""

from __future__ import annotations

import json

import pytest

pytest.importorskip("ooxml_edit.tools.sources", reason="needs ooxml-edit 0.22 (sources)")

from ooxml_edit.tools import Toolbox  # noqa: E402
from test_tables_from_data import GRID, _session, _table, call, ds_put, failed, ok, texts  # noqa: E402

from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from docx_agent.tools.objects import _table_entry  # noqa: E402

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def session(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
    return session


def edit(toolbox, session, table, **arguments):
    return call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells", **arguments)


def _rows(session, table):
    return [row for row, _ in _table_entry(session.entry("d1").document, table).rows]


def _fills(row) -> list[str | None]:
    return [None if shd is None else shd.get(_W + "fill")
            for shd in (tc.find(f"{_W}tcPr/{_W}shd") for tc in row.iter(_W + "tc"))]


FROM = {"from": "d:scores", "from_columns": ["check", "note"]}


def test_grow_adds_the_rows_the_data_needs_each_like_the_last(toolbox, session):
    table = _table(toolbox, session, 2, 2)
    ok(call(toolbox, session, "word_format_table", doc="d1", table=table, scope="row", row=1, shading="FFCC00",
            height=20))
    result = ok(edit(toolbox, session, table, grow=True, **FROM))
    assert texts(session, table) == [["check", "note"], ["Alpha", "fine"], ["Beta", ""],
                                     ["Gamma", "late\vvery"]]
    rows = _rows(session, table)
    assert [_fills(r) for r in rows[1:]] == [["FFCC00", "FFCC00"]] * 3
    heights = [r.find(f"{_W}trPr/{_W}trHeight").get(_W + "val") for r in rows[1:]]
    assert len(set(heights)) == 1
    assert len(result.created) == 2 and all(c.startswith("tr:") for c in result.created)
    assert "2 row(s) added after row 1" in result.summary
    assert result.data["rows"] == 4 and result.data["from"]["columns"][0]["rows"] == "1-3"
    assert result.checks["validate"]["new"] == []


def test_without_grow_the_refusal_is_the_call_that_works(toolbox, session):
    table = _table(toolbox, session, 2, 2)
    error = failed(edit(toolbox, session, table, **FROM), "refused")
    assert "grow: true" in error.message and "2 row(s)" in error.message
    assert error.valid_options[0] == "grow" and error.field == "from"
    meant = error.details["call"]
    assert meant["tool"] == "word_edit_table"
    assert meant["arguments"] == {"doc": "d1", "table": table, "action": "set_cells", "grow": True, **FROM}
    assert texts(session, table) == [["old", "old"]] * 2           # nothing changed
    ok(call(toolbox, session, meant["tool"], **meant["arguments"]))
    assert len(texts(session, table)) == 4


def test_a_header_row_last_is_copied_as_a_body_row(toolbox, session):
    # A table holding its header row only: the rows added are not header rows.
    table = _table(toolbox, session, 1, 2)
    ok(call(toolbox, session, "word_format_table", doc="d1", table=table, scope="table", header_rows=1))
    ok(edit(toolbox, session, table, grow=True, **FROM))
    rows = _rows(session, table)
    assert rows[0].find(f"{_W}trPr/{_W}tblHeader") is not None
    assert all(r.find(f"{_W}trPr/{_W}tblHeader") is None for r in rows[1:])
    assert texts(session, table)[3] == ["Gamma", "late\vvery"]


def test_grow_is_one_undo_step(toolbox, session):
    table = _table(toolbox, session, 2, 2)
    ok(edit(toolbox, session, table, grow=True, **FROM))
    ok(call(toolbox, session, "undo", doc="d1"))
    assert texts(session, table) == [["old", "old"]] * 2


def test_grow_tracks_the_rows_it_adds(toolbox, session):
    table = _table(toolbox, session, 2, 2)
    ok(call(toolbox, session, "word_set_tracking", doc="d1", on=True, author="Reviewer"))
    result = ok(edit(toolbox, session, table, grow=True, **FROM))
    assert result.checks["validate"]["new"] == []
    rows = _rows(session, table)
    assert all(r.find(f"{_W}trPr/{_W}ins") is not None for r in rows[2:])


def test_grow_from_the_row_after_the_last_appends(toolbox, session):
    table = _table(toolbox, session, 2, 2)
    error = failed(edit(toolbox, session, table, index=2, header=[], **FROM), "not_found")
    assert error.field == "index" and error.valid_options == ["grow"]
    ok(edit(toolbox, session, table, index=2, header=[], grow=True, **FROM))
    assert texts(session, table)[2:] == [["Alpha", "fine"], ["Beta", ""], ["Gamma", "late\vvery"]]
    assert texts(session, table)[:2] == [["old", "old"]] * 2


def test_grow_needs_from_and_false_asks_nothing(toolbox, session):
    table = _table(toolbox, session, 2, 2)
    error = failed(edit(toolbox, session, table, grow=True, cells=[{"row": 0, "col": 0, "text": "x"}]),
                   "invalid_arguments")
    assert error.field == "grow" and "insert_row" in error.valid_options
    ok(edit(toolbox, session, table, grow=False, cells=[{"row": 0, "col": 0, "text": "x"}]))
    # Room enough: grow adds nothing.
    big = _table(toolbox, session, 5, 2)
    result = ok(edit(toolbox, session, big, grow=True, **FROM))
    assert result.created == [] and "added" not in result.summary


def test_grow_over_a_merged_last_row_is_refused_before_anything_changes(toolbox, session):
    table = _table(toolbox, session, 2, 2)
    ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="merge",
            first={"row": 1, "col": 0}, last={"row": 1, "col": 1}))
    before = texts(session, table)
    error = failed(edit(toolbox, session, table, grow=True, **FROM), "refused")
    assert "merged" in error.message
    assert texts(session, table) == before and len(_rows(session, table)) == 2


def test_more_columns_than_the_table_are_refused_with_grow_too(toolbox, session):
    table = _table(toolbox, session, 2, 1)
    error = failed(edit(toolbox, session, table, grow=True, **FROM), "refused")
    assert error.field == "from_columns" and "the table has 1" in error.message


def test_fewer_rows_than_the_table_says_which_rows_keep_their_text(toolbox, session):
    table = _table(toolbox, session, 6, 2)
    result = ok(edit(toolbox, session, table, **FROM))
    warning = " ".join(result.warnings)
    assert "rows 4-5 below the data" in warning and "delete_row index 4, 2 time(s)" in warning
    assert texts(session, table)[4:] == [["old", "old"]] * 2


@pytest.mark.parametrize("atomic", [True, False])
def test_in_a_batch(toolbox, session, atomic):
    table = _table(toolbox, session, 2, 2)
    op = {"tool": "word_edit_table", "arguments": {"doc": "d1", "table": table, "action": "set_cells", **FROM}}
    other = {"tool": "word_insert_markdown", "arguments": {"doc": "d1", "at": "end", "markdown": "After."}}
    result = call(toolbox, session, "batch", ops=[other, op], atomic=atomic)
    if atomic:
        assert not result.ok and "grow: true" in result.error.message
        assert result.error.details["error"]["details"]["call"]["arguments"]["grow"] is True
        assert texts(session, table) == [["old", "old"]] * 2
    else:
        assert "grow: true" in json.dumps(result.to_json(), default=str)
    op["arguments"]["grow"] = True
    ok(call(toolbox, session, "batch", ops=[other, op], atomic=atomic))
    assert len(texts(session, table)) == 4
    ok(call(toolbox, session, "undo", doc="d1"))
    assert len(texts(session, table)) == 2
