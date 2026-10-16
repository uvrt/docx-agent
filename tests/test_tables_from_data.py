"""Word tables from data (X1): ``word_insert_markdown`` makes a table, and
``word_edit_table``'s ``set_cells`` fills one, from a dataset (``d:<id>``) or a table of
another open document (``t:<name>``), picking its columns by name and reporting where every
placed value came from -- as pptx-agent's ``ppt_add_table`` and ``ppt_edit_table`` do.

The toys, all made up: a dataset a toy tool keeps (``d:scores``), and a toy grid format
whose documents are named tables of cells, resolving ``t:<name>`` to rows with a source per
cell (``G!B2``) and merging runs of cells in one column into ranges (``G!B2:B4``).

``d:`` is also a Word drawing's prefix (``d:3``): a ``from`` resolves datasets only and says
so when given a drawing, and no argument that takes a drawing is a source.
"""

from __future__ import annotations

import datetime as dt
import json

import pytest

pytest.importorskip("ooxml_edit.tools.sources", reason="needs ooxml-edit 0.22 (sources)")

from ooxml_common.locales import Locale  # noqa: E402
from ooxml_edit.tools import (Dataset, DocumentFormat, Result, ToolError, Toolbox,  # noqa: E402
                              string, strict_problems, tool)

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 10, 13, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731

#: A made-up checks table: a name, a rating in half points, a date as ISO text, a flag, a note.
COLUMNS = ("check", "rating", "due", "done", "note")
ROWS = [("Alpha", 3.5, "2026-01-05", True, "fine"),
        ("Beta", 4, "2026-02-10", False, None),
        ("Gamma", 1.25, "2026-03-15", True, "late\nvery")]


#: The checks' column types, as a workbook's date column gives them: ``due`` is a date.
TYPES = ("text", "number", "date", "bool", "text")


def _sources(rows) -> list:
    return [[(f"G!{chr(65 + c)}{r + 2}",) for c in range(len(COLUMNS))] for r in range(len(rows))]


class Grid:
    """Named tables, each a header and rows; cell ``B2`` is the second column's first row."""

    def __init__(self, tables: dict) -> None:
        self.tables = tables


def _grid_table(grid: Grid, address: str) -> Dataset:
    """``t:<name>``, or one row of it by key, ``t:<name>/<key>`` (as a workbook's ``t:``
    takes one, for a text fact)."""
    name, _, key = address[len("t:"):].partition("/")
    if name not in grid.tables:
        raise ToolError("not_found", f"no table {name!r}", valid_options=sorted(grid.tables))
    header, rows = grid.tables[name]
    sources = _sources(rows)
    if key:
        found = [i for i, row in enumerate(rows) if row[0] == key]
        if not found:
            raise ToolError("not_found", f"no row {key!r} in {name}")
        rows, sources = [rows[found[0]]], [sources[found[0]]]
    return Dataset(header, rows, sources=sources, keys=1)


def _merge(addresses):
    """Runs of cells in one column as ranges: G!B2, G!B3 -> G!B2:B3."""
    out: list[list] = []
    for address in addresses:
        sheet, cell = address.split("!")
        column, row = cell[0], int(cell[1:])
        last = out[-1] if out else None
        if last and last[:2] == [sheet, column] and last[3] == row - 1:
            last[3] = row
        else:
            out.append([sheet, column, row, row])
    return [f"{s}!{c}{a}" if a == b else f"{s}!{c}{a}:{c}{b}" for s, c, a, b in out]


GRID = DocumentFormat(kind="grid", open=lambda data: Grid({}), detect=lambda data, name: False,
                      datasets={"t:": _grid_table}, merge_sources=_merge)


@tool("ds_put", "Keep a made-up dataset.",
      {"id": string("Its id."), "size": string("rows x columns of filler.", optional=True),
       "typed": string("Label the columns' types, as a workbook's t: does.", optional=True)},
      documents=())
def ds_put(call, id, size=None, typed=None):
    if typed:
        call.define_dataset(id, Dataset(COLUMNS, ROWS, sources=_sources(ROWS), keys=1,
                                        types=TYPES), source=None)
    elif size:
        rows, columns = (int(n) for n in size.split("x"))
        names = [f"c{i}" for i in range(columns)]
        call.define_dataset(id, Dataset(names, [[r] * columns for r in range(rows)]), source=None)
    else:
        call.define_dataset(id, Dataset(COLUMNS, ROWS, sources=_sources(ROWS), keys=1),
                            source=None)
    return Result(summary=f"Kept d:{id}")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


#: The rectangle every test document has: a drawing, ``d:<docPr id>``.
DRAWING = "d:1"


def _new_document() -> bytes:
    document = Document.new()
    paragraph = document.paragraphs()[0]
    paragraph.set_text("Checks")
    document.insert_shape(f"{document.paragraphs()[0].id}@0", "rect", width=40, height=40)
    return document.to_bytes()


def _session(toolbox, **options):
    session = toolbox.session(clock=CLOCK, **options)
    session.open(_new_document(), "report.docx")                                 # d1
    session.adopt(Grid({"checks": (COLUMNS, ROWS)}), "grid", "book.grid")        # d2
    return session


@pytest.fixture
def session(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
    return session


def call(toolbox, session, name, **arguments):
    return toolbox.dispatch(session, name, arguments)


def ok(result):
    assert result.ok, json.dumps(result.to_json(), indent=1, default=str)
    return result


def failed(result, code):
    assert not result.ok, json.dumps(result.to_json(), indent=1, default=str)
    assert result.error.code == code, result.to_json()
    return result.error


def insert(toolbox, session, **arguments):
    return call(toolbox, session, "word_insert_markdown", **{"doc": "d1", "at": "end", **arguments})


def texts(session, table: str) -> list[list[str]]:
    view = session.entry("d1").document.table(table)
    return [[view.cell(r, c).text for c in range(view.column_count)] for r in range(len(view.rows))]


# -- the definitions -----------------------------------------------------------------------------


def test_both_tools_declare_from_and_stay_strict_clean(toolbox):
    for name in ("word_insert_markdown", "word_edit_table"):
        tool_ = toolbox.tools[name]
        assert tool_.sources == ("from",)
        properties = tool_.schema["properties"]
        assert properties["from"]["type"] == "string"
        assert properties["from_columns"]["items"] == {"type": "string"}
        assert properties["header"]["items"] == {"type": "string"}
        assert not {"from", "from_columns", "header"} & set(tool_.schema.get("required", []))
    assert ("markdown", "blob", "from") in [tuple(g) for g in
                                            toolbox.tools["word_insert_markdown"].exactly_one]
    assert strict_problems(list(TOOLS)) == []


def test_no_argument_that_takes_a_drawing_is_a_source():
    for tool_ in TOOLS:
        assert set(tool_.sources) <= {"from"}, tool_.name
        for path in tool_.sources:
            described = tool_.schema["properties"][path].get("description", "")
            assert "drawing" not in described, tool_.name
    drawing_tools = {"word_drawings", "edit_chart", "edit_smartart", "word_delete"}
    for tool_ in TOOLS:
        if tool_.name in drawing_tools:
            assert "target" not in tool_.sources, tool_.name


# -- a new table: word_insert_markdown ------------------------------------------------------------


def test_a_table_from_a_dataset_header_rows_and_provenance(toolbox, session):
    result = ok(insert(toolbox, session, **{"from": "d:scores",
                                            "from_columns": ["check", "rating", "due", "done"]}))
    data = result.data
    table = data["id"]
    assert table in data["blocks"] and table in result.created
    assert (data["rows"], data["columns"]) == (4, 4)
    assert "markdown" not in data                       # the caller has the data
    assert texts(session, table) == [["check", "rating", "due", "done"],
                                     ["Alpha", "3.5", "2026-01-05", "TRUE"],
                                     ["Beta", "4", "2026-02-10", "FALSE"],
                                     ["Gamma", "1.25", "2026-03-15", "TRUE"]]
    placed = data["from"]
    assert placed["source"] == "d:scores" and placed["doc"] is None
    assert placed["header_row"] == 0 and placed["first_row"] == 1
    assert placed["columns"][1] == {"col": 1, "column": "rating", "rows": "1-3",
                                    "sources": ["G!B2", "G!B3", "G!B4"]}
    document = session.entry("d1").document
    assert list(document.provenance()) == [table]       # kept in the file (test_provenance)
    xml = document.package.tree(document.package.document_part())
    tbl = xml.find(".//{*}tbl")
    assert tbl.find("{*}tr/{*}trPr/{*}tblHeader") is not None      # the header row repeats
    assert result.checks["validate"]["new"] == []


def test_columns_are_picked_and_ordered_by_name_and_default_to_all(toolbox, session):
    picked = ok(insert(toolbox, session, **{"from": "d:scores",
                                            "from_columns": ["Rating", "check"]})).data
    rows = texts(session, picked["id"])
    assert rows[0] == ["rating", "check"] and rows[1] == ["3.5", "Alpha"]
    every = ok(insert(toolbox, session, **{"from": "d:scores"})).data
    rows = texts(session, every["id"])
    assert every["columns"] == 5 and rows[0] == list(COLUMNS)
    assert rows[2][4] == ""                             # a blank is an empty cell


def test_a_values_lines_are_line_breaks_in_one_paragraph(toolbox, session):
    data = ok(insert(toolbox, session, **{"from": "d:scores", "from_columns": ["note"]})).data
    cell = session.entry("d1").document.table(data["id"]).cell(3, 0)
    assert len(cell.paragraphs) == 1 and cell.text == "late\vvery"


def test_a_header_given_or_none(toolbox, session):
    given = ok(insert(toolbox, session, **{"from": "d:scores", "from_columns": ["check", "rating"],
                                           "header": ["Check", "Rating (1-5)"]})).data
    assert texts(session, given["id"])[0] == ["Check", "Rating (1-5)"] and given["rows"] == 4
    none = ok(insert(toolbox, session, **{"from": "d:scores", "from_columns": ["check"],
                                          "header": []})).data
    assert none["rows"] == 3 and texts(session, none["id"])[0] == ["Alpha"]
    assert none["from"]["header_row"] is None and none["from"]["first_row"] == 0
    document = session.entry("d1").document
    tbl = document.package.tree(document.package.document_part()).findall(".//{*}tbl")[-1]
    assert tbl.find(".//{*}tblHeader") is None
    look = tbl.find("{*}tblPr/{*}tblLook")
    assert look.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstRow") == "0"
    error = failed(insert(toolbox, session, **{"from": "d:scores", "from_columns": ["check"],
                                               "header": ["a", "b"]}), "invalid_arguments")
    assert error.field == "header"


def test_text_goes_in_literally_not_as_markdown(toolbox, session):
    @tool("ds_marks", "A dataset of Markdown-looking text.", {}, documents=())
    def ds_marks(call):
        call.define_dataset("marks", Dataset(("a",), [("**not bold** | [x](y) <b>",)]),
                            source=None)
        return Result(summary="ok")

    with Toolbox([*TOOLS, ds_marks], formats=[FORMAT], groups=GROUPS, workers=0) as box:
        own = box.session(clock=CLOCK)
        own.open(_new_document(), "report.docx")
        ok(box.dispatch(own, "ds_marks", {}))
        data = ok(box.dispatch(own, "word_insert_markdown", {"doc": "d1", "at": "end",
                                                              "from": "d:marks"})).data
        view = own.entry("d1").document.table(data["id"])
        assert view.cell(1, 0).text == "**not bold** | [x](y) <b>"


def test_numbers_are_written_in_the_sessions_locale(toolbox):
    dutch = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                   language="nl-NL", list_separator=";")
    session = _session(toolbox, locale=dutch)
    data = ok(insert(toolbox, session, **{"from": "t:checks",
                                          "from_columns": ["rating", "due"]})).data
    rows = texts(session, data["id"])
    assert [row[0] for row in rows[1:]] == ["3,5", "4", "1,25"]
    assert rows[1][1] == "2026-01-05"                   # a date as the source gives it


def test_a_table_from_another_documents_table_merges_its_sources(toolbox, session):
    data = ok(insert(toolbox, session, **{"from": "t:checks",
                                          "from_columns": ["check", "note"]})).data
    placed = data["from"]
    assert placed["source"] == "d2/t:checks" and placed["doc"] == "d2"
    assert placed["columns"][0]["sources"] == ["G!A2:A4"]
    assert placed["columns"][1] == {"col": 1, "column": "note", "rows": "1-3",
                                    "sources": ["G!E2:E4"]}


def test_a_row_range_of_a_source_is_a_table_whose_header_repeats(toolbox, session):
    data = ok(insert(toolbox, session, **{"from": "d:scores/#2:",
                                          "from_columns": ["check", "rating"]})).data
    assert texts(session, data["id"]) == [["check", "rating"], ["Beta", "4"],
                                          ["Gamma", "1.25"]]
    placed = data["from"]
    assert placed["source"] == "d:scores/#2:"
    assert placed["columns"][1] == {"col": 1, "column": "rating", "rows": "1-2",
                                    "sources": ["G!B3", "G!B4"]}
    document = session.entry("d1").document
    tbl = document.package.tree(document.package.document_part()).findall(".//{*}tbl")[-1]
    assert tbl.find("{*}tr/{*}trPr/{*}tblHeader") is not None      # Word repeats it per page
    failed(insert(toolbox, session, **{"from": "d:scores/#2:#9"}), "invalid_arguments")


def test_from_with_markdown_or_without_a_source_is_refused(toolbox, session):
    error = failed(insert(toolbox, session, **{"from": "d:scores", "markdown": "x"}),
                   "invalid_arguments")
    assert set(error.valid_options) == {"markdown", "blob", "from"}
    error = failed(insert(toolbox, session, markdown="x", from_columns=["check"]),
                   "invalid_arguments")
    assert error.field == "from_columns"
    error = failed(insert(toolbox, session, markdown="x", header=[]), "invalid_arguments")
    assert error.field == "header"
    error = failed(insert(toolbox, session, **{"from": "d:scores", "from_columns": ["nope"]}),
                   "not_found")
    assert "rating" in error.valid_options
    failed(insert(toolbox, session, **{"from": "d:scores", "from_columns": ["check", "CHECK"]}),
           "invalid_arguments")
    failed(insert(toolbox, session, **{"from": "d:missing"}), "not_found")


def test_a_source_too_big_is_refused_never_cut(toolbox, session):
    ok(call(toolbox, session, "ds_put", id="wide", size="2x64"))
    error = failed(insert(toolbox, session, **{"from": "d:wide"}), "limit")
    assert error.field == "from_columns" and "64 columns" in error.message and "63" in error.message
    ok(call(toolbox, session, "ds_put", id="cells", size="100x60"))
    error = failed(insert(toolbox, session, **{"from": "d:cells"}), "limit")
    assert "5,000" in error.message and "nothing is cut" in error.message
    tables = len(session.entry("d1").document.tables())
    ok(call(toolbox, session, "ds_put", id="fits", size="99x50"))
    ok(insert(toolbox, session, **{"from": "d:fits"}))                 # 100 x 50: 5,000 cells
    assert len(session.entry("d1").document.tables()) == tables + 1


def test_a_value_a_cell_cannot_take_is_refused_naming_it():
    @tool("ds_odd", "A dataset with a value that is not finite.", {}, documents=())
    def ds_odd(call):
        call.define_dataset("odd", Dataset(("a",), [(1,), (float("nan"),)],
                                           sources=[[("G!A2",)], [("G!A3",)]]), source=None)
        return Result(summary="ok")

    with Toolbox([*TOOLS, ds_odd], formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session(clock=CLOCK)
        session.open(_new_document(), "report.docx")
        ok(box.dispatch(session, "ds_odd", {}))
        error = failed(box.dispatch(session, "word_insert_markdown",
                                    {"doc": "d1", "at": "end", "from": "d:odd"}),
                       "invalid_arguments")
        assert "G!A3" in error.message and "not a finite number" in error.message
        assert session.entry("d1").document.tables() == []


def test_a_new_table_from_a_source_is_one_undo_step_and_tracks(toolbox, session):
    ok(insert(toolbox, session, **{"from": "d:scores"}))
    ok(call(toolbox, session, "undo", doc="d1"))
    assert session.entry("d1").document.tables() == []
    ok(call(toolbox, session, "word_set_tracking", doc="d1", on=True, author="Reviewer"))
    result = ok(insert(toolbox, session, **{"from": "d:scores", "from_columns": ["check"]}))
    assert result.checks["validate"]["new"] == []
    assert session.entry("d1").document.revisions()


# -- an existing table: word_edit_table set_cells ----------------------------------------------


def _table(toolbox, session, rows, columns) -> str:
    row = "| " + " | ".join(["old"] * columns) + " |"
    markdown = "\n".join([row, "|" + "---|" * columns] + [row] * (rows - 1))
    return ok(insert(toolbox, session, markdown=markdown)).data["blocks"][0]


def edit(toolbox, session, table, **arguments):
    return call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                **arguments)


def test_set_cells_from_a_source_rewrites_the_block(toolbox, session):
    table = _table(toolbox, session, 4, 2)
    result = ok(edit(toolbox, session, table, **{"from": "d:scores",
                                                 "from_columns": ["check", "note"]}))
    assert texts(session, table) == [["check", "note"], ["Alpha", "fine"], ["Beta", ""],
                                     ["Gamma", "late\vvery"]]
    assert result.data["from"]["first_row"] == 1 and not result.warnings
    assert f"{table}/c2,1" in result.changed
    assert result.data["id"] == table and "markdown" not in result.data


def test_set_cells_from_row_index_under_a_kept_header(toolbox, session):
    table = _table(toolbox, session, 5, 3)
    result = ok(edit(toolbox, session, table, index=1, header=[],
                     **{"from": "d:scores", "from_columns": ["check", "rating"]}))
    rows = texts(session, table)
    assert rows[0] == ["old"] * 3 and rows[1] == ["Alpha", "3.5", "old"]
    assert rows[4] == ["old"] * 3
    placed = result.data["from"]
    assert placed["header_row"] is None and placed["first_row"] == 1
    assert result.warnings and "keep their text" in result.warnings[0]


def test_set_cells_from_a_source_too_big_for_the_table_is_refused(toolbox, session):
    table = _table(toolbox, session, 3, 2)
    error = failed(edit(toolbox, session, table, **{"from": "d:scores",
                                                    "from_columns": ["check"]}), "refused")
    assert "4 rows" in error.message and "insert_row" in error.message
    # A row range of the source places fewer: the refusal offers the one that fits.
    assert error.valid_options == ["d:scores/#1:#2"]
    error = failed(edit(toolbox, session, table, header=[], **{"from": "d:scores"}), "refused")
    assert error.field == "from_columns"
    error = failed(edit(toolbox, session, table, index=3, **{"from": "d:scores"}), "not_found")
    assert error.field == "index"
    error = failed(call(toolbox, session, "word_edit_table", doc="d1", table=table,
                        action="delete_row", index=0, **{"from": "d:scores"}),
                   "invalid_arguments")
    assert error.field == "from"
    error = failed(edit(toolbox, session, table, cells=[], **{"from": "d:scores"}),
                   "invalid_arguments")
    assert error.field == "cells"
    error = failed(call(toolbox, session, "word_edit_table", doc="d1", table=table,
                        action="insert_row", header=[]), "invalid_arguments")
    assert error.field == "header"
    assert texts(session, table)[0] == ["old", "old"]


def test_set_cells_from_a_source_over_a_merged_cell_is_refused(toolbox, session):
    table = _table(toolbox, session, 4, 2)
    ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="merge",
            first={"row": 2, "col": 0}, last={"row": 2, "col": 1}))
    error = failed(edit(toolbox, session, table, **{"from": "d:scores",
                                                    "from_columns": ["check", "note"]}),
                   "refused")
    assert "(2, 1)" in error.message and "split it first" in error.message
    assert texts(session, table)[0] == ["old", "old"]


def test_set_cells_from_a_source_is_one_undo_step(toolbox, session):
    table = _table(toolbox, session, 4, 1)
    ok(edit(toolbox, session, table, **{"from": "t:checks", "from_columns": ["check"]}))
    ok(call(toolbox, session, "undo", doc="d1"))
    assert texts(session, table) == [["old"]] * 4


# -- d: is a drawing too ---------------------------------------------------------------------------


def test_a_drawings_id_as_from_is_named_a_drawing_not_a_dataset(toolbox, session):
    drawing = DRAWING
    assert session.entry("d1").document.drawing(drawing) is not None
    error = failed(insert(toolbox, session, **{"from": drawing}), "not_found")
    assert error.field == "from"
    assert f"{drawing} is a drawing in this document, not a dataset" in error.message
    assert error.valid_options == ["d:scores"]          # the datasets there are
    table = _table(toolbox, session, 2, 2)
    error = failed(edit(toolbox, session, table, **{"from": drawing}), "not_found")
    assert "is a drawing in this document" in error.message
    error = failed(call(toolbox, session, "edit_chart", doc="d1", target=drawing,
                        action="set_values", **{"from": drawing}), "not_found")
    assert "is a drawing in this document" in error.message
    error = failed(insert(toolbox, session, **{"from": "d:999"}), "not_found")
    assert "d:999 is a drawing's id (none here), not a dataset" in error.message


def test_a_dataset_named_like_nothing_drawn_gets_no_drawing_hint(toolbox, session):
    error = failed(insert(toolbox, session, **{"from": "d:missing"}), "not_found")
    assert "drawing" not in error.message


def test_a_drawings_id_still_names_the_drawing_where_drawings_are_taken(toolbox, session):
    drawing = DRAWING
    ok(call(toolbox, session, "word_drawings", doc="d1", action="resize", target=drawing,
            width=30))


def test_a_drawings_id_in_a_batch_is_named_a_drawing(toolbox, session):
    drawing = DRAWING
    result = call(toolbox, session, "batch", ops=[
        {"tool": "word_insert_markdown",
         "arguments": {"doc": "d1", "at": "end", "from": drawing}}])
    assert not result.ok
    assert "is a drawing in this document" in json.dumps(result.to_json())


def test_column_names_given_as_columns_are_read_as_from_columns(toolbox, session):
    made = ok(insert(toolbox, session, **{"from": "d:scores", "columns": ["rating", "check"]}))
    assert texts(session, made.data["id"])[0] == ["rating", "check"]
    assert made.warnings[0] == "columns read as from_columns"
    table = _table(toolbox, session, 4, 2)
    result = ok(edit(toolbox, session, table, **{"from": "d:scores",
                                                 "columns": ["check", "note"]}))
    assert texts(session, table)[0] == ["check", "note"]
    assert result.warnings[0] == "columns read as from_columns"
    # split's count of columns stays a count.
    error = failed(edit(toolbox, session, table, columns=[2]), "invalid_arguments")
    assert error.field == "columns"


# -- dates (T18b) --------------------------------------------------------------------------------

DUTCH = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")
#: The checks' due dates as each locale writes them.
DUE = {"en-US": ["1/5/2026", "2/10/2026", "3/15/2026"],
       "nl-BE": ["5/1/2026", "10/2/2026", "15/3/2026"]}


def _dated(toolbox, locale):
    session = _session(toolbox, locale=locale)
    ok(call(toolbox, session, "ds_put", id="dated", typed="yes"))
    return session


@pytest.mark.parametrize("locale, due", [(None, DUE["en-US"]), (DUTCH, DUE["nl-BE"])])
def test_a_date_columns_iso_text_is_written_in_the_sessions_locale(toolbox, locale, due):
    session = _dated(toolbox, locale)
    result = ok(insert(toolbox, session, **{"from": "d:dated",
                                            "from_columns": ["check", "due"]}))
    table = result.data["id"]
    assert [row[1] for row in texts(session, table)[1:]] == due
    # The record names the cells, never their values.
    kept = session.entry("d1").document.provenance()[table]
    assert kept["columns"][1]["sources"] == ["G!C2", "G!C3", "G!C4"]
    assert not any(text in json.dumps(kept) for text in due + ["2026-01-05"])
    # set_cells writes the same.
    other = _table(toolbox, session, 4, 2)
    ok(edit(toolbox, session, other, **{"from": "d:dated", "from_columns": ["check", "due"]}))
    assert [row[1] for row in texts(session, other)[1:]] == due


def test_an_unlabelled_iso_column_is_written_as_it_is(toolbox):
    session = _dated(toolbox, DUTCH)
    table = ok(insert(toolbox, session, **{"from": "t:checks",
                                           "from_columns": ["due"]})).data["id"]
    assert texts(session, table)[1] == ["2026-01-05"]           # no type: text


@pytest.mark.parametrize("locale, key", [(None, "en-US"), (DUTCH, "nl-BE")])
def test_a_chart_from_a_date_column_and_set_values_matched_either_way(toolbox, locale, key):
    session = _dated(toolbox, locale)
    anchor = session.entry("d1").document.paragraphs()[0].id
    made = ok(call(toolbox, session, "edit_chart", doc="d1", target=anchor, action="add",
                   chart_type="column",
                   **{"from": "d:dated", "columns": ["due", "rating"]})).created[-1]
    read = ok(call(toolbox, session, "edit_chart", doc="d1", target=made, action="read"))
    assert read.data["categories"] == DUE[key]
    # An older document's chart, its categories written ISO and in another order.
    older = ok(call(toolbox, session, "edit_chart", doc="d1", target=anchor, action="add",
                    chart_type="column", categories=["2026-03-15", "2026-01-05", "2026-02-10"],
                    data=[{"name": "rating", "values": [0, 0, 0]}])).created[-1]
    for target, order in ((made, [3.5, 4, 1.25]), (older, [1.25, 3.5, 4])):
        ok(call(toolbox, session, "edit_chart", doc="d1", target=target, action="set_values",
                series="rating", **{"from": "d:dated", "columns": ["due", "rating"]}))
        read = ok(call(toolbox, session, "edit_chart", doc="d1", target=target,
                       action="read"))
        assert read.data["series"][0]["values"] == order


# -- date_format (round 6, R9) ------------------------------------------------------------------


def _locale(language):
    return Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
                  language=language, list_separator=";")


#: The checks' due dates (5 January, 10 February, 15 March 2026) in ``dddd d mmmm``, the
#: names measured in Excel for Mac, 2026-10-16.
@pytest.mark.parametrize("language, due", [
    ("en-GB", ["Monday 5 January", "Tuesday 10 February", "Sunday 15 March"]),
    ("nl-NL", ["maandag 5 januari", "dinsdag 10 februari", "zondag 15 maart"]),
    ("es-ES", ["lunes 5 enero", "martes 10 febrero", "domingo 15 marzo"]),
    ("it-IT", ["lunedì 5 gennaio", "martedì 10 febbraio", "domenica 15 marzo"]),
])
def test_a_tables_dates_in_a_date_format(toolbox, language, due):
    session = _dated(toolbox, _locale(language))
    result = ok(insert(toolbox, session, date_format="dddd d mmmm",
                       **{"from": "d:dated", "from_columns": ["check", "due", "rating"]}))
    rows = texts(session, result.data["id"])[1:]
    assert [row[1] for row in rows] == due
    assert rows[0][2] == "3,5"                                   # a number as before
    other = _table(toolbox, session, 4, 2)
    ok(edit(toolbox, session, other, date_format="dd.mm.yy",
            **{"from": "d:dated", "from_columns": ["check", "due"]}))
    assert [row[1] for row in texts(session, other)[1:]] == ["05.01.26", "10.02.26",
                                                             "15.03.26"]


def test_a_date_format_is_checked_for_a_table(toolbox):
    session = _dated(toolbox, _locale("sv-SE"))
    error = failed(insert(toolbox, session, date_format="d mmm yyyy",
                          **{"from": "d:dated", "from_columns": ["due"]}), "invalid_arguments")
    assert "no month or day names are measured for sv-SE" in error.message
    assert error.field == "date_format"
    ok(insert(toolbox, session, date_format="yyyy-mm-dd",
              **{"from": "d:dated", "from_columns": ["due"]}))
    error = failed(insert(toolbox, session, date_format="dd/mm/yyyy",
                          **{"from": "d:dated", "from_columns": ["check"]}), "invalid_arguments")
    assert "date_format is for dates" in error.message
    error = failed(insert(toolbox, session, markdown="Text", date_format="dd/mm/yyyy"),
                   "invalid_arguments")
    assert "date_format is for from" in error.message


def test_a_charts_dates_are_kept_in_the_locales_code(toolbox):
    from lxml import etree

    session = _dated(toolbox, _locale("fr-FR"))
    anchor = session.entry("d1").document.paragraphs()[0].id
    made = ok(call(toolbox, session, "edit_chart", doc="d1", target=anchor, action="add",
                   chart_type="column", date_format="d mmmm yyyy",
                   **{"from": "d:dated", "columns": ["due", "rating"]})).created[-1]
    read = ok(call(toolbox, session, "edit_chart", doc="d1", target=made, action="read"))
    # Read back as ISO dates (ooxml-edit 0.39.0), with the code that shows them.
    assert read.data["categories"] == ["2026-01-05", "2026-02-10", "2026-03-15"]
    assert read.data["category_format"] == r"[$-40C]d\ mmmm\ yyyy"
    # A category is named by its date, as the read lists it.
    ok(call(toolbox, session, "edit_chart", doc="d1", target=made, action="set_value",
            series="rating", category="2026-02-10", value=2))
    read = ok(call(toolbox, session, "edit_chart", doc="d1", target=made, action="read"))
    assert read.data["series"][0]["values"] == [3.5, 2, 1.25]
    error = failed(call(toolbox, session, "edit_chart", doc="d1", target=made, action="set_value",
                        series="rating", category="46063", value=2), "not_found")
    assert "named as YYYY-MM-DD" in error.message and "2026-02-10" in error.valid_options
    document = session.entry("d1").document
    chart = document.chart(made)
    root = etree.fromstring(document.package.read(chart.part))
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    assert root.xpath("string(//c:cat//c:formatCode)", namespaces=ns) == r"[$-40C]d\ mmmm\ yyyy"
    assert root.xpath("string(//c:catAx/c:auto/@val)", namespaces=ns) == "0"
    # The Markdown outline's chart line lists them as dates too.
    assert "categories: 2026-01-05, 2026-02-10, 2026-03-15" in document.to_markdown()
