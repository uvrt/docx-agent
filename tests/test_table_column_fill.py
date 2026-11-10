"""Round 48 (R89): ``word_edit_table`` ``set_cells`` fills a column -- or a range of its rows --
with one text or one value by reference, in one entry: ``{col: 2, rows: "*", text: "Yes"}``.

``rows`` is ``*`` (every row below the header rows) or a range from 0 as ``row`` counts
(``1..40``, ``5..``, ``..9``, ``7``); the header rows are those marked to repeat
(``w:tblHeader``), else the first where the table's look has its header row on.  The entry
writes what as many entries of one cell each write (the same bytes); the result says the
range and counts, not every cell.  Made-up tables and data only."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import zipfile

import pytest

pytest.importorskip("ooxml_edit.tools.sources", reason="needs ooxml-edit 0.22 (sources)")

from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent import blocks as _blocks  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 11, 10, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
HEAD = "| code | name | state |\n|---|---|---|\n"


@tool("parts_put", "Keep a made-up dataset.", {"id": string("Its id.")}, documents=())
def parts_put(call, id):
    call.define_dataset(id, Dataset(("code", "state"), [(f"P-{i}", "open" if i % 2 else "shut")
                                                         for i in range(5)]), source=None)
    return Result(summary=f"Kept d:{id}")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, parts_put], formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


def digest(data: bytes) -> str:
    archive = zipfile.ZipFile(io.BytesIO(data))
    found = hashlib.sha256()
    for info in archive.infolist():
        if info.filename != "docProps/core.xml":
            found.update(info.filename.encode() + b"\0" + archive.read(info.filename))
    return found.hexdigest()


def _table(toolbox, rows: int = 6, *, markdown: str | None = None):
    session = toolbox.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "parts.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    body = markdown or HEAD + "\n".join(f"| P-{i} | part {i} |  |" for i in range(rows))
    assert toolbox.dispatch(session, "word_insert_markdown", {"doc": "d1", "at": f"after:{first}",
                                                              "markdown": body}).ok
    assert toolbox.dispatch(session, "parts_put", {"id": "parts"}).ok
    return session, session.entry("d1").document.tables()[-1].id


def _fill(toolbox, session, table, cells, **more):
    return toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "set_cells",
                                                         "cells": cells, **more})


def _column(session, table, col):
    view = session.entry("d1").document.table(table)
    return [view.cell(r, col).text for r in range(len(view.rows))]


def test_the_cells_entry_takes_rows_and_the_strict_set_is_unchanged():
    from ooxml_edit.tools.adapters import anthropic_strict_plan

    from docx_agent.tools.format import STRICT_FIRST

    edit = next(t for t in TOOLS if t.name == "word_edit_table")
    assert "cells[].rows" in edit.optional_parameters()
    plan = anthropic_strict_plan(TOOLS, first=STRICT_FIRST)
    assert not plan["word_edit_table"]          # never strict for Claude: 19 optional of its own
    assert sum(len(t.optional_parameters()) for t in TOOLS if plan[t.name]) == 24


def test_a_column_below_its_header_takes_one_text(toolbox):
    session, table = _table(toolbox)
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "*", "text": "Yes"}])
    assert result.ok, result.to_json()
    assert _column(session, table, 2) == ["state"] + ["Yes"] * 6
    assert result.data["filled"] == [{"entry": 0, "col": 2, "rows": "1..6", "cells": 6, "changed": 6,
                                      "header_rows": 1}]
    # rule 4: the range and counts, not the table's text or each cell's paragraph
    assert "markdown" not in result.data and result.changed == [table]
    assert result.created == [] and result.removed == []


def test_the_same_bytes_as_one_entry_per_cell(toolbox):
    for value in ({"text": "Yes"}, {"text": "two\nlines"}, {"from": "d:parts/#2/state"}):
        session, table = _table(toolbox, 9)
        assert _fill(toolbox, session, table, [{"col": 1, "rows": "2..7", **value}]).ok
        at_once = digest(session.entry("d1").document.to_bytes())
        session, table = _table(toolbox, 9)
        assert _fill(toolbox, session, table, [{"row": r, "col": 1, **value} for r in range(2, 8)]).ok
        assert digest(session.entry("d1").document.to_bytes()) == at_once, value


def test_tracked_changes_too(toolbox):
    out = []
    for cells in ([{"col": 2, "rows": "*", "text": "Yes"}],
                  [{"row": r, "col": 2, "text": "Yes"} for r in range(1, 7)]):
        session, table = _table(toolbox)
        assert toolbox.dispatch(session, "word_set_tracking", {"doc": "d1", "on": True, "author": "Editor"}).ok
        assert _fill(toolbox, session, table, cells).ok
        out.append(digest(session.entry("d1").document.to_bytes()))
    assert out[0] == out[1]


def test_rows_forms(toolbox):
    session, table = _table(toolbox, 8)
    for said, rows in (("3..5", [3, 4, 5]), ("6..", [6, 7, 8]), ("..2", [1, 2]), ("4", [4]),
                       ("1-2", [1, 2]), ("7:8", [7, 8]), ("all", list(range(1, 9)))):
        assert _fill(toolbox, session, table, [{"col": 2, "rows": said, "text": said}]).ok, said
        column = _column(session, table, 2)
        assert all(column[r] == said for r in rows), (said, column)


def test_a_range_naming_the_header_writes_it_and_says_so(toolbox):
    session, table = _table(toolbox)
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "0..2", "text": "x"}])
    assert result.ok and _column(session, table, 2)[:4] == ["x", "x", "x", ""]
    assert any("header row" in w and "written as asked" in w for w in result.warnings)


def test_header_rows_marked_to_repeat_or_none(toolbox):
    session, table = _table(toolbox)
    session.entry("d1").document.set_table(table, header_rows=2)
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "*", "text": "Yes"}])
    assert result.ok and _column(session, table, 2)[:3] == ["state", "", "Yes"]
    assert result.data["filled"][0]["header_rows"] == 2
    # none marked and the look's header row off: every row is the column's
    document = session.entry("d1").document
    document.set_table(table, header_rows=0, look={"first_row": False})
    result = _fill(toolbox, session, table, [{"col": 0, "rows": "*", "text": "z"}])
    assert result.ok and _column(session, table, 0) == ["z"] * 7
    assert "header_rows" not in result.data["filled"][0]


def test_a_column_by_its_header_text_and_beside_single_cells(toolbox):
    session, table = _table(toolbox)
    result = _fill(toolbox, session, table, [{"col_label": "state", "rows": "*", "text": "Yes"},
                                             {"row": 3, "col": 2, "text": "No"},
                                             {"row_label": "P-0", "col": 1, "text": "first"}])
    assert result.ok, result.to_json()
    assert _column(session, table, 2) == ["state", "Yes", "Yes", "No", "Yes", "Yes", "Yes"]
    assert _column(session, table, 1)[1] == "first"


def test_a_merged_cell_is_written_once_and_a_missing_one_is_said(toolbox):
    session, table = _table(toolbox, 5)
    document = session.entry("d1").document
    document.merge_cells(table, (2, 2), (3, 2))
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "*", "text": "Yes"}])
    assert result.ok and result.data["filled"][0]["cells"] == 4
    assert document.table(table).cell(2, 2).text == "Yes"
    view = document.table(table)
    assert view.cell(3, 2).id == view.cell(2, 2).id


def test_a_value_by_reference_once_kept_in_every_cell(toolbox):
    session, table = _table(toolbox)
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "*", "from": "d:parts/#2/state"}])
    assert result.ok, result.to_json()
    assert _column(session, table, 2)[1:] == ["open"] * 6        # #2: the second row
    facts = result.data["facts"]
    assert facts == [{"from": "d:parts/#2/state", "text": "open", "col": 2, "rows": "1..6", "cells": 6,
                      **{k: v for k, v in facts[0].items() if k in ("sources", "stale")}}]
    kept = session.entry("d1").document.text_facts()
    view = session.entry("d1").document.table(table)
    for r in range(1, 7):
        assert view.cell(r, 2).paragraphs[0].id in kept


def test_a_reference_in_the_text_kept_in_every_cell(toolbox):
    session, table = _table(toolbox)
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "2..4", "text": "state: {{d:parts/#1/state}}"}])
    assert result.ok, result.to_json()
    assert _column(session, table, 2)[2:5] == ["state: shut"] * 3
    kept = session.entry("d1").document.text_facts()
    view = session.entry("d1").document.table(table)
    for r in (2, 3, 4):
        found = kept[view.cell(r, 2).paragraphs[0].id]
        assert [(f["text"], f["at"]) for f in found] == [("shut", 7)]
    assert len(result.data["facts"]) == 1 and result.data["facts"][0]["rows"] == "2..4"


def test_a_top_level_one_value_fills_the_column(toolbox):
    session, table = _table(toolbox)
    result = toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "set_cells",
                                                           "from": "d:parts/#3/state",
                                                           "cells": [{"col": 2, "rows": "*"}]})
    assert result.ok and _column(session, table, 2)[1:] == ["shut"] * 6


def test_placeholders_are_none(toolbox):
    # rule 1: a strict decoder's "" for rows is no rows -- one cell, as before
    session, table = _table(toolbox)
    result = _fill(toolbox, session, table, [{"row": 1, "col": 2, "rows": "", "text": "one",
                                              "row_label": "", "col_label": "", "from": ""}])
    assert result.ok and _column(session, table, 2)[1:3] == ["one", ""]
    assert "filled" not in result.data and "markdown" in result.data


def test_one_undo_step(toolbox):
    session, table = _table(toolbox)
    before = digest(session.entry("d1").document.to_bytes())
    assert _fill(toolbox, session, table, [{"col": 2, "rows": "*", "from": "d:parts/#2/state"}]).ok
    assert toolbox.dispatch(session, "undo", {"doc": "d1"}).ok
    assert digest(session.entry("d1").document.to_bytes()) == before


def test_already_holding_the_text_is_not_an_error(toolbox):
    session, table = _table(toolbox)
    assert _fill(toolbox, session, table, [{"col": 2, "rows": "*", "text": "Yes"}]).ok
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "*", "text": "Yes"}])
    assert result.ok and result.data["filled"][0]["changed"] == 0
    assert any("held that text already" in w for w in result.warnings)


@pytest.mark.parametrize("cell, said, option", [
    ({"col": 2, "rows": "1..40", "text": "x"}, "no row 40", "1..6"),
    ({"col": 7, "rows": "*", "text": "x"}, "grid columns 0 to 2", "2"),
    ({"rows": "*", "text": "x"}, "no column", "col"),
    ({"row": 1, "col": 2, "rows": "*", "text": "x"}, "gives rows and row", "row"),
    ({"col": 2, "rows": "5..2", "text": "x"}, "runs backwards", "2..5"),
    ({"col": 2, "rows": "every", "text": "x"}, "is not rows", "*"),
])
def test_a_refusal_says_the_table_and_the_rows_that_work(toolbox, cell, said, option):
    session, table = _table(toolbox)
    before = digest(session.entry("d1").document.to_bytes())
    result = _fill(toolbox, session, table, [cell])
    assert not result.ok and said in result.error.message, result.error.message
    assert option in (result.error.valid_options or []), result.error.valid_options
    assert digest(session.entry("d1").document.to_bytes()) == before


def test_rows_without_a_cell_are_said_not_failed(toolbox):
    # rule 5: the cells there are written; a row without one at that column is listed
    session, table = _table(toolbox, 4)
    document = session.entry("d1").document
    from docx_agent.edit.document import qn

    row = document._resolve(table)[1].rows[2][0]
    cells = row.findall(qn("w:tc"))
    row.remove(cells[-1])                          # row 2 ends a column early
    trpr = row.find(qn("w:trPr"))
    if trpr is None:
        trpr = row.makeelement(qn("w:trPr"), {})
        row.insert(0, trpr)
    after = trpr.makeelement(qn("w:gridAfter"), {qn("w:val"): "1"})
    trpr.append(after)
    document._invalidate()
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "*", "text": "Yes"}])
    assert result.ok, result.to_json()
    assert result.data["filled"][0]["cells"] == 3 and result.data["filled"][0]["no_cell"] == 1
    assert any("row(s) 2 have no cell at column 2" in w for w in result.warnings)


def test_a_column_of_431_rows_keeps_its_values_in_one_edit(toolbox, monkeypatch):
    session, table = _table(toolbox, 431)
    calls = []
    original = Document._follow_text_facts
    monkeypatch.setattr(Document, "_follow_text_facts", lambda self: (calls.append(1), original(self))[1])
    result = _fill(toolbox, session, table, [{"col": 2, "rows": "*", "from": "d:parts/#2/state"}])
    assert result.ok and result.data["filled"][0]["cells"] == 431
    assert len(calls) < 20, len(calls)              # 431 edits of its own each, before
    assert len(json.dumps(result.to_json(), default=str)) < 1500


def test_a_column_fill_is_drafted_as_a_cell_is():
    assert _blocks.admits("word_edit_table", {"action": "set_cells", "table": "t:1A2B3C4D",
                                              "cells": [{"col": 2, "rows": "*", "from": "d:parts/#2/state"}]})
