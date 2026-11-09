"""Round 45 (R67): a table filled from a dataset in one pass, and ``word_fields`` update's
reading done once -- the same bytes and results as before, and a ``timing`` block saying where
the time went.

``word_edit_table`` ``set_cells`` ``from`` wrote its block as one edit per row added and one or
more per cell, each throwing every index away, so the next read the whole document again: a
431-row block took minutes.  ``Document._rows_after_last`` and ``Document._fill_cells`` make
the same edits in one pass; the one-by-one path is still taken where something needs stamping,
and is what the one pass is compared with here (``_stamping_due`` forced true), on many table
shapes.  ``word_fields`` update read the part from its root for every field's result, every
line of the layout for every heading of a table of contents, the numbering part for every
paragraph's list, and parsed the whole package for the toc styles' looks on every pass; each
is compared with the reading it replaces.  Made-up documents and data only."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import sys
import zipfile
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.sources", reason="needs ooxml-edit 0.22 (sources)")

from docx2svg.paths import resolve_path  # noqa: E402
from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.edit import fields as _fields  # noqa: E402
from docx_agent.edit import ids as _ids  # noqa: E402
from docx_agent.edit.effective import PathFinder  # noqa: E402
from docx_agent.edit.numbering import membership, reading  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

sys.path.insert(0, str(Path(__file__).parent))
import toc_number_cases  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 11, 8, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
FIXTURES = Path(__file__).parent / "fixtures"

#: A made-up stock list: a code, a name (some empty), an amount (some missing), a date, a note
#: (some missing, some of two lines, some long, some with French punctuation).
COLUMNS = ("code", "name", "amount", "due", "note")
TYPES = ("text", "text", "number", "date", "text")


def _rows(count: int, long: bool = False) -> list[tuple]:
    out = []
    for i in range(count):
        note = None if i % 4 == 1 else ("line one\nline two" if i % 4 == 2 else f"note {i} : fin !")
        if long and i % 3 == 0:
            note = ("lantern meadow river stone harbour quiet orchard copper " * 12).strip()
        out.append((f"K-{i:03d}", f"Item {i}" if i % 5 else "", 1.5 * i if i % 6 else None,
                    f"2026-0{1 + i % 9}-1{i % 10}", note))
    return out


@tool("stock_put", "Keep a made-up dataset.", {"id": string("Its id."), "rows": string("How many."),
                                                 "long": string("Long notes.", optional=True),
                                                 "typed": string("Typed columns.", optional=True)},
      documents=())
def stock_put(call, id, rows, long=None, typed=None):
    call.define_dataset(id, Dataset(COLUMNS, _rows(int(rows), bool(long)), types=TYPES if typed else None),
                        source=None)
    return Result(summary=f"Kept d:{id}")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, stock_put], formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


def digest(data: bytes) -> str:
    """The package's parts, but its core properties (a new document's creation time)."""
    archive = zipfile.ZipFile(io.BytesIO(data))
    found = hashlib.sha256()
    for info in archive.infolist():
        if info.filename != "docProps/core.xml":
            found.update(info.filename.encode() + b"\0" + archive.read(info.filename))
    return found.hexdigest()


def plain(result) -> dict:
    """A result's JSON without what differs from run to run: its duration and timing."""
    out = result.to_json()
    out.pop("duration", None)
    if isinstance(out.get("data"), dict):
        out["data"] = {k: v for k, v in out["data"].items() if k != "timing"}
    return out


HEAD = "| code | name | amount | due | note |\n|---|---|---|---|---|\n"
ROW = "| a | b | c | d | e |"

#: name -> (setup, fills, dataset): the table shapes compared.
SHAPES = {
    "grow": ({}, [{"grow": True}], {}),
    "header-repeated": ({"header_rows": True}, [{"grow": True}], {}),
    "header-row-only": ({"header_rows": True, "delete_row": True}, [{"grow": True}], {}),
    "number-formats": ({}, [{"grow": True, "date_format": "dd/MM/yyyy"}], {"typed": "1"}),
    "some-columns": ({"markdown": "| x | y |\n|---|---|\n| 1 | 2 |"},
                     [{"grow": True, "from_columns": ["note", "code"]}], {}),
    "below-the-header": ({}, [{"grow": True, "header": [], "index": 1}], {}),
    "rows-left-below": ({"markdown": HEAD + "\n".join([ROW] * 20)}, [{}], {"rows": "12"}),
    "long-text": ({}, [{"grow": True}], {"long": "1"}),
    "cells-of-paragraphs": ({"paragraphs": True}, [{"grow": True}], {}),
    "filled-again": ({}, [{"grow": True}, {}], {}),
    "tracked": ({"tracked": True}, [{"grow": True}], {}),
    "tracked-header": ({"tracked": True, "header_rows": True}, [{"grow": True}], {}),
    "merged-header": ({"merge_header": True}, [{"grow": True, "header": [], "index": 1}], {}),
    "merged-down-above": ({"markdown": HEAD + ROW + "\n" + ROW, "merge_down": True},
                          [{"grow": True, "header": [], "index": 2}], {}),
    "shaded-rows": ({"shade": True}, [{"grow": True}], {}),
    "french": ({"language": "fr-FR"}, [{"grow": True}], {}),
}


def _run(toolbox, setup: dict, fills: list[dict], dataset: dict) -> tuple:
    session = toolbox.session(clock=CLOCK)
    document = Document.new()
    document.paragraphs()[0].set_text("Stock report")
    session.open(document.to_bytes(), "report.docx")
    results = []

    def call(name, **arguments):
        result = toolbox.dispatch(session, name, arguments)
        assert result.ok, json.dumps(result.to_json(), indent=1, default=str)
        results.append(plain(result))
        return result

    if setup.get("language"):
        call("word_styles", doc="d1", action="set_language", language=setup["language"])
    first = session.entry("d1").document.paragraphs()[0].id
    call("word_insert_markdown", doc="d1", at=f"after:{first}", markdown=setup.get("markdown", HEAD + ROW))
    table = session.entry("d1").document.tables()[-1].id
    if setup.get("header_rows"):
        call("word_format_table", doc="d1", table=table, scope="table", header_rows=1)
    if setup.get("delete_row"):
        call("word_edit_table", doc="d1", table=table, action="delete_row", index=1)
    if setup.get("shade"):
        call("word_format_table", doc="d1", table=table, scope="row", row=1, shading="FFCC00", height=20)
    if setup.get("merge_header"):
        call("word_edit_table", doc="d1", table=table, action="merge", first={"row": 0, "col": 0},
             last={"row": 0, "col": 1})
    if setup.get("merge_down"):
        call("word_edit_table", doc="d1", table=table, action="merge", first={"row": 0, "col": 4},
             last={"row": 1, "col": 4})
    if setup.get("paragraphs"):
        call("word_edit_table", doc="d1", table=table, action="set_cells",
             cells=[{"row": 1, "col": k, "text": "one\ntwo\nthree"} for k in range(5)])
    if setup.get("tracked"):
        call("word_set_tracking", doc="d1", on=True, author="Reviewer")
    call("stock_put", id="stock", rows=dataset.get("rows", "23"), long=dataset.get("long"),
         typed=dataset.get("typed"))
    timings = []
    for fill in fills:
        table = session.entry("d1").document.tables()[-1].id
        timings.append(call("word_edit_table", doc="d1", table=table, action="set_cells",
                            **{"from": "d:stock", **fill}).data["timing"])
    filled = digest(session.entry("d1").document.to_bytes())
    call("undo", doc="d1")
    return filled, digest(session.entry("d1").document.to_bytes()), results, timings


@pytest.mark.parametrize("name", sorted(SHAPES))
def test_one_pass_writes_what_the_edits_one_by_one_wrote(toolbox, monkeypatch, name):
    """The same bytes, the same results and the same undo, cell for cell and id for id."""
    setup, fills, dataset = SHAPES[name]
    at_once = _run(toolbox, setup, fills, dataset)
    monkeypatch.setattr(Document, "_stamping_due", lambda self: True)
    one_by_one = _run(toolbox, setup, fills, dataset)
    assert at_once[:3] == one_by_one[:3]


def test_the_result_says_where_the_time_went(toolbox):
    _, _, results, timings = _run(toolbox, {}, [{"grow": True}], {})
    timing = timings[0]
    assert list(timing)[:6] == ["resolve", "check", "rows", "cells", "record", "result"]
    assert {"spacing", "validate"} <= set(timing)
    assert all(isinstance(v, float) and v >= 0 for v in timing.values())
    assert len(json.dumps(timing)) < 300                       # a small part of the result's budget


def test_a_block_of_many_rows_reads_the_document_a_few_times_not_per_cell(toolbox, monkeypatch):
    """The work of a fill does not grow with rows x document: the story's index is built a
    few times in all, not once or more per cell (2 x 23 x 5 before)."""
    built = []
    original = _ids.PartIndex.__init__

    def counted(self, *args, **kwargs):
        built.append(1)
        original(self, *args, **kwargs)

    session = toolbox.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "report.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    assert toolbox.dispatch(session, "word_insert_markdown", {"doc": "d1", "at": f"after:{first}",
                                                              "markdown": HEAD + ROW}).ok
    assert toolbox.dispatch(session, "stock_put", {"id": "stock", "rows": "120"}).ok
    table = session.entry("d1").document.tables()[-1].id
    monkeypatch.setattr(_ids.PartIndex, "__init__", counted)
    result = toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "set_cells",
                                                           "from": "d:stock", "grow": True})
    assert result.ok and result.data["rows"] == 121
    assert len(built) < 40, len(built)


# -- word_fields update: each reading against the one it replaces ------------------------------


def _old_between(start, end):
    first, last = _fields._run_of(start), _fields._run_of(end)
    return _fields._between_from_root(start, first, last)


def _field_documents():
    yield Document.open((FIXTURES / "generated" / "markdown" / "constructs.docx").read_bytes())
    yield Document.open(toc_number_cases.CASES["n-tab"])


def test_a_fields_result_is_read_from_its_start_onward():
    seen = 0
    for document in _field_documents():
        document.update_fields()
        for found in document._all_fields():
            if found.simple or found.separate is None or found.end is None:
                continue
            assert _fields._between(found.separate, found.end) == _old_between(found.separate, found.end)
            # End before start, or the same run: what the walk from the root says.
            assert _fields._between(found.end, found.separate) == _old_between(found.end, found.separate)
            assert _fields._between(found.begin, found.begin) == _old_between(found.begin, found.begin)
            seen += 1
    assert seen >= 3


def _old_where(layout, identifier):
    from fractions import Fraction  # noqa: F401 -- the lines' measures

    from docx_agent.layout import PX_PER_PT, Placement, _story_kind

    wanted = layout._wanted(identifier)
    out = []
    for number, lines in enumerate(layout._lines):
        columns: dict = {}
        for key, line in lines:
            if key is not None and layout._canon(key) in wanted:
                columns.setdefault(line.column or 0, []).append(line)
        for column, mine in columns.items():
            top = min(line.top for line in mine)
            bottom = max(line.top + line.pitch for line in mine)
            story = mine[0].story or _story_kind(mine[0].path)
            out.append(Placement(number + 1, round(float(top / PX_PER_PT), 2),
                                 round(float(bottom / PX_PER_PT), 2), len(mine), story, column))
    return out or layout._unknown(identifier)


def test_where_a_block_is_drawn_is_looked_up_not_searched():
    document = Document.open((FIXTURES / "wordto" / "sample-with-table.docx").read_bytes())
    layout = document.layout()
    ids = [p.id for p in document.paragraphs()] + [t.id for t in document.tables()] + ["p:7FFFFFFF"]
    for identifier in ids:
        assert repr(layout.where(identifier)) == repr(_old_where(layout, identifier)), identifier
    # A rename seen through the aliases: the index is made again.
    paragraph = document.paragraphs()[1]
    before = layout.where(paragraph.id)
    document._aliases["p:00ABCDEF"] = paragraph.id
    assert repr(layout.where("p:00ABCDEF")) == repr(_old_where(layout, "p:00ABCDEF")) == repr(before)


def test_paths_resolve_as_docx2svg_resolves_them():
    count = 0
    for name in ("sample-5pages.docx", "sample-with-table.docx", "sample-with-images.docx"):
        document = Document.open((FIXTURES / "wordto" / name).read_bytes())
        layout = document.layout()
        finder = PathFinder()
        root = document.package.tree(document.package.document_part())
        for page in layout.pages:
            for line in page.text_lines():
                assert finder.resolve(root, line.path) is resolve_path(root, line.path), line.path
                count += 1
        for path in ("w:body/w:p[1]", "w:body/w:p[999]", "w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]",
                     "w:body/w:p[1]/w:r[1]", "w:body/w:sectPr[1]"):
            assert finder.resolve(root, path) is resolve_path(root, path), path
    assert count > 20


def test_lists_read_once_say_what_they_said():
    document = Document.open((FIXTURES / "generated" / "lists-and-styles.docx").read_bytes())
    elements = [entry.element for entry in document._index(document.package.document_part()).paragraphs]
    plain_read = [membership(document, element) for element in elements]
    with reading(document):
        held = [membership(document, element) for element in elements]
    assert held == plain_read and any(found is not None for found in held)
    assert document.__dict__["_numbering_held"] is None


@pytest.mark.parametrize("name", ["n-tab", "n-lvl", "t-end", "s-arial12"])
def test_toc_looks_kept_by_their_parts_write_the_same_entries(name):
    cases = {**toc_number_cases.CASES, **toc_number_cases.CASES2}
    _fields._LOOKS.clear()
    fresh = Document.open(cases[name])
    fresh.update_fields()
    assert _fields._LOOKS                                      # kept, by what they are made of
    again = Document.open(cases[name])
    again.update_fields()                                      # found kept: no parse for the looks
    assert digest(fresh.to_bytes()) == digest(again.to_bytes())
    # Another body, the same styles: the same key; other styles: another.
    other = Document.open(cases["n-11pt"])
    assert _fields._looks_key(other) != _fields._looks_key(fresh)


def test_the_update_says_where_its_time_went(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open(toc_number_cases.CASES["n-tab"], "toc.docx")
    result = toolbox.dispatch(session, "word_fields", {"doc": "d1", "action": "update"})
    assert result.ok, result.to_json()
    timing = result.data["timing"]
    assert list(timing)[:2] == ["ahead", "update"] and "validate" in timing
    assert len(json.dumps(timing)) < 200


# -- set_cells by position: one pass too (a thousand cells took 80 s one by one) ----------------


def _cells_run(toolbox, tracked: bool) -> tuple:
    session = toolbox.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "report.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    results = []

    def call(name, **arguments):
        result = toolbox.dispatch(session, name, arguments)
        assert result.ok, json.dumps(result.to_json(), indent=1, default=str)
        results.append(plain(result))

    call("word_insert_markdown", doc="d1", at=f"after:{first}", markdown=HEAD + "\n".join([ROW] * 6))
    table = session.entry("d1").document.tables()[-1].id
    call("word_edit_table", doc="d1", table=table, action="merge", first={"row": 1, "col": 0},
         last={"row": 2, "col": 1})
    if tracked:
        call("word_set_tracking", doc="d1", on=True, author="Reviewer")
    cells = [{"row": r, "col": c, "text": f"r{r} c{c} : ok" + ("\nsecond line" if (r + c) % 3 == 0 else "")}
             for r in range(7) for c in range(5)]
    cells += [{"row": 2, "col": 1, "text": "the merged cell again"}, {"row": 3, "col": 3, "text": "d"}]
    call("word_edit_table", doc="d1", table=table, action="set_cells", cells=cells)
    call("word_edit_table", doc="d1", table=table, action="set_cells",
         cells=[{"row": 4, "col": 0, "text": "one"}, {"row": 4, "col_label": "r0 c4 : ok", "text": "by label"}])
    filled = digest(session.entry("d1").document.to_bytes())
    call("undo", doc="d1")
    return filled, digest(session.entry("d1").document.to_bytes()), results


@pytest.mark.parametrize("tracked", [False, True])
def test_cells_by_position_write_what_they_wrote_one_by_one(toolbox, monkeypatch, tracked):
    at_once = _cells_run(toolbox, tracked)
    monkeypatch.setattr(Document, "_stamping_due", lambda self: True)
    assert at_once == _cells_run(toolbox, tracked)


def test_a_cell_the_table_lacks_is_refused_as_before(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "report.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    assert toolbox.dispatch(session, "word_insert_markdown", {"doc": "d1", "at": f"after:{first}",
                                                              "markdown": HEAD + ROW}).ok
    table = session.entry("d1").document.tables()[-1].id
    result = toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "set_cells",
                                                           "cells": [{"row": 0, "col": 0, "text": "x"},
                                                                     {"row": 9, "col": 0, "text": "y"}]})
    assert not result.ok and "no cell at (9, 0)" in result.error.message
    assert session.entry("d1").document.table(table).cell(0, 0).text == "code"
