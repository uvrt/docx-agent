"""Round 56 (R120): ``word_edit_table`` given the document's own table as ``from``.

A sync (or set_cells, or a new table from data) whose ``from`` names a table of the document
being edited -- bare (``t:458D9849``) or with its handle (``d1/t:458D9849``) -- is refused
saying so: ``from`` takes the rows to place, a dataset (``d:...``) or a workbook table
(``t:...``); the datasets and workbook tables there are are listed, with the call shape, and
the call itself when there is one source alone.  Every address a refusal offers resolves:
``from``'s, ``key``'s and ``table``'s.

The workbook is a made-up grid (a stand-in whose ``t:<name>`` resolves to its rows and whose
``value_sources`` lists them), as in ``test_block_sources``.  Made-up documents and values only.
"""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.sources", reason="needs ooxml-edit 0.114 (own_source)")

from ooxml_edit.tools import Dataset, DocumentFormat, Result, Toolbox, ToolError, string, tool  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from docx_agent.tools.table_data import own_source  # noqa: E402

if "own_source" not in DocumentFormat.__dataclass_fields__:
    pytest.skip("needs ooxml-edit 0.114 (own_source)", allow_module_level=True)

COLUMNS = ("code", "item", "count")
ROWS = [("K-01", "Bolts", 12), ("K-02", "Nuts", 7), ("K-03", "Pins", 3)]


class Grid:
    def __init__(self) -> None:
        self.tables = {"stock": (COLUMNS, [list(row) for row in ROWS])}


def _table(grid: Grid, address: str) -> Dataset:
    name, _, key = address[len("t:"):].partition("/")
    if name not in grid.tables:
        raise ToolError("not_found", f"no table, name or region t:{name}; there are "
                        f"{['t:' + n for n in grid.tables]}", valid_options=[f"t:{n}" for n in grid.tables])
    header, rows = grid.tables[name]
    rows = [tuple(row) for row in rows]
    if key:
        rows = [row for i, row in enumerate(rows) if key in (f"#{i + 1}", row[0])]
    return Dataset(header, rows, keys=1)


GRID = DocumentFormat(kind="xlsx", open=lambda data: Grid(), detect=lambda data, name: False,
                      datasets={"t:": _table}, value_sources=lambda grid: [f"t:{n}" for n in grid.tables])


@tool("rows_put", "Keep a made-up dataset.", {"id": string("Its id.")}, documents=())
def rows_put(call, id):
    call.define_dataset(id, Dataset(COLUMNS, list(ROWS)), source=None)
    return Result(summary=f"Kept d:{id}")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, rows_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


def _report(toolbox, *, grids: int = 1, datasets: tuple = ()):
    """A made-up report d1 with one table (the grid's rows, by hand), grids d2..., datasets."""
    session = toolbox.session()
    session.open(Document.new().to_bytes(), "report.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    markdown = "| code | item | count |\n|---|---|---|\n| K-01 | Bolts | 12 |"
    assert toolbox.dispatch(session, "word_insert_markdown",
                            {"doc": "d1", "at": f"after:{first}", "markdown": markdown}).ok
    for n in range(grids):
        session.adopt(Grid(), "xlsx", f"stock{n}.xlsx")
    for name in datasets:
        assert toolbox.dispatch(session, "rows_put", {"id": name}).ok
    return session, session.entry("d1").document.tables()[-1].id


def _sync(toolbox, session, table, **more):
    return toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "sync", **more})


def _fresh(toolbox, grids, datasets):
    return _report(toolbox, grids=grids, datasets=datasets)


# -- the report's own table as from (the reported refusals) --------------------------------------


@pytest.mark.parametrize("prefix", ["", "d1/"])
def test_the_own_table_is_named_as_the_documents_and_the_rows_are_listed(toolbox, prefix):
    session, table = _report(toolbox, datasets=("counts",))
    result = _sync(toolbox, session, table, **{"from": prefix + table})
    assert not result.ok and result.error.code == "invalid_arguments", result.to_json()
    error = result.error
    assert error.message.startswith(
        f"{table} is a table of d1, the document this call edits, not rows to take: from takes a dataset "
        "(d:<id>) or rows of an open xlsx document (t:<name>); there are datasets d:counts and tables t:stock")
    assert 'e.g. from: "d:counts"' in error.message
    assert error.field == "from" and error.valid_options == ["d:counts", "t:stock"]
    assert error.details["names"] == {"doc": "d1", "address": table, "is": "a table"}
    assert "d2/" + table not in error.valid_options


@pytest.mark.parametrize("prefix", ["", "d1/"])
def test_one_source_alone_gives_the_call_that_works(toolbox, prefix):
    session, table = _report(toolbox)
    error = _sync(toolbox, session, table, **{"from": prefix + table}).error
    assert error.valid_options == ["t:stock"]
    meant = error.details["call"]
    assert meant == {"tool": "word_edit_table",
                     "arguments": {"doc": "d1", "table": table, "action": "sync", "from": "t:stock"}}
    done = toolbox.dispatch(session, meant["tool"], meant["arguments"])
    assert done.ok, done.to_json()
    view = session.entry("d1").document.table(table)
    assert [view.cell(r, 0).text for r in range(len(view.rows))] == ["code", "K-01", "K-02", "K-03"]


def test_with_two_workbooks_each_table_is_named_with_its_document(toolbox):
    session, table = _report(toolbox, grids=2, datasets=("counts",))
    error = _sync(toolbox, session, table, **{"from": table}).error
    assert error.code == "invalid_arguments"
    assert error.valid_options == ["d:counts", "d2/t:stock", "d3/t:stock"]
    assert "call" not in error.details


def test_nothing_to_take_rows_from_says_how_to_get_some(toolbox):
    session, table = _report(toolbox, grids=0)
    error = _sync(toolbox, session, table, **{"from": table}).error
    assert error.code == "invalid_arguments" and error.valid_options == []
    assert error.message.endswith("there is none: define a dataset or open a document with tables first")


@pytest.mark.parametrize("tool, arguments", [
    ("word_insert_markdown", {"at": "end"}),
    ("word_edit_table", {"action": "set_cells"}),
])
def test_other_froms_of_the_document_say_the_same(toolbox, tool, arguments):
    session, table = _report(toolbox, datasets=("counts",))
    arguments = {"doc": "d1", **arguments, "from": table}
    if tool == "word_edit_table":
        arguments["table"] = table
    error = toolbox.dispatch(session, tool, arguments).error
    assert error.code == "invalid_arguments" and error.message.startswith(f"{table} is a table of d1")
    assert error.valid_options == ["d:counts", "t:stock"]


def test_a_paragraph_is_named_as_one(toolbox):
    session, table = _report(toolbox)
    paragraph = session.entry("d1").document.paragraphs()[0].id
    error = _sync(toolbox, session, table, **{"from": paragraph}).error
    assert error.message.startswith(f"{paragraph} is a paragraph of d1, the document this call edits")


def test_a_reference_to_the_own_table_is_told_what_a_reference_names(toolbox):
    session, table = _report(toolbox)
    error = toolbox.dispatch(session, "word_edit_table", {
        "doc": "d1", "table": table, "action": "set_cells",
        "cells": [{"row": 1, "col": 0, "text": "{{" + table + "/#1/code}}"}]}).error
    assert error.code == "invalid_arguments"
    assert error.message.startswith(f"{table}/#1 is part of a table of d1, the document this call edits, not a "
                                    "value to quote")
    assert "{{<source>/#<row>/<column>}}" in error.message and "call" not in error.details


def test_own_source_names_the_blocks():
    document = Document.new()
    first = document.paragraphs()[0].id
    assert own_source(document, first) == "a paragraph"
    assert own_source(document, "t:00000000") is None and own_source(document, "d:counts") is None


# -- every suggestion resolves (tool design rule 2) ---------------------------------------------


def _source_refusal(result, field="from"):
    return not result.ok and result.error.field == field and result.error.code in (
        "not_found", "invalid_arguments", "ambiguous")


@pytest.mark.parametrize("grids", [1, 2])
@pytest.mark.parametrize("value", ["OWN", "d1/OWN", "d2/OWN", "t:nope", "d1/t:stock", "d1/t:nope",
                                   "d9/t:stock", "x:1", "PARAGRAPH"])
def test_every_from_suggestion_resolves(toolbox, grids, value):
    session, table = _report(toolbox, grids=grids, datasets=("counts",))
    paragraph = session.entry("d1").document.paragraphs()[0].id
    given = value.replace("OWN", table).replace("PARAGRAPH", paragraph)
    result = _sync(toolbox, session, table, **{"from": given})
    assert _source_refusal(result), result.to_json()
    assert result.error.valid_options, result.to_json()
    for option in result.error.valid_options:
        fresh, again = _fresh(toolbox, grids, ("counts",))
        tried = _sync(toolbox, fresh, again, **{"from": option})
        assert tried.ok, (given, option, tried.to_json())
        if "call" in result.error.details:
            meant = result.error.details["call"]
            assert toolbox.dispatch(session, meant["tool"], meant["arguments"]).ok


def test_every_key_suggestion_resolves(toolbox):
    session, table = _report(toolbox)
    error = _sync(toolbox, session, table, **{"from": "t:stock", "key": "nope"}).error
    assert error.code == "not_found" and error.field == "key"
    assert error.valid_options == ["code", "item", "count"]
    for option in error.valid_options:
        fresh, again = _report(toolbox)
        tried = _sync(toolbox, fresh, again, **{"from": "t:stock", "key": option})
        assert not _source_refusal(tried, "key"), (option, tried.to_json())


@pytest.mark.parametrize("given", ["t:stock", "d2/t:stock", "t:nope"])
def test_a_table_not_in_the_document_offers_its_tables(toolbox, given):
    session, table = _report(toolbox, datasets=("counts",))
    error = _sync(toolbox, session, given, **{"from": "d:counts"}).error
    assert error.code == "not_found" and error.field == "table" and error.valid_options == [table]
    if given == "t:nope":
        assert error.message == f"no table 't:nope' in d1; d1's tables: {table}"
    else:
        assert error.message.startswith(f"{given} is a table of d2 (xlsx), not of d1: word_edit_table changes a "
                                        f"table of d1 ({table})")
    for option in error.valid_options:
        assert _sync(toolbox, session, option, **{"from": "d:counts"}).ok


# -- the schema ----------------------------------------------------------------------------------


def test_no_parameter_was_added_and_the_strict_set_is_full():
    from ooxml_edit.tools.adapters import anthropic_strict_plan

    from docx_agent.tools.format import STRICT_FIRST

    edit = next(t for t in TOOLS if t.name == "word_edit_table")
    assert "own_source" in FORMAT.__dataclass_fields__ and FORMAT.own_source is own_source
    plan = anthropic_strict_plan(TOOLS, first=STRICT_FIRST)
    assert sum(len(t.optional_parameters()) for t in TOOLS if plan[t.name]) == 24
    assert "from" in edit.canonical["properties"]
