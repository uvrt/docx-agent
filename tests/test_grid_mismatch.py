"""Round 53 (R107): a table whose row reaches past its grid.

A row with more cells (or columns spanned) than ``w:tblGrid`` -- Word's own form of a merge or
a column deletion tracked (the cell merged away kept, ``w:cellDel``), a row from a table of
another grid -- stopped docx2svg's layout with an ``IndexError``, which the tool layer reported
as ``not_found``: render, check and the live view of the document all failed, again and again.
docx2svg (0.11.11) lays such a table out as Word does; here:

* ``check`` says it (``layout``'s ``table_grid_mismatch``: the table, its rows, how Word draws
  it, the call that settles a tracked one), from the document alone, so even where the layout
  failed (rule 7);
* a layout that fails is ``layout_failed``, naming the table where docx2svg names it -- never
  ``not_found`` -- and a version that failed fails again at once, logged once (rule 2);
* the edits that left a table off its grid keep it in step: a row inserted beside a column
  whose tracked deletion or insertion is pending goes with that change, and a review leaves
  every table it touched fitting its grid.

Made-up "Shelf" reports only.
"""

from __future__ import annotations

import datetime as dt
import io
import logging
import zipfile

import docx2svg
import pytest
from lxml import etree
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent import layout as layout_module
from docx_agent.layout import LayoutFailed
from docx_agent.layout_facts import table_grid_mismatch
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.validate import grid_problems

CLOCK = lambda: dt.datetime(2026, 11, 12, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SHELF = ("| item | room | notes |\n|---|---|---|\n| Lamp | hall | Outcome: partly done |\n"
         "| Kettle | kitchen | Outcome: done |\n| Stool | attic | Outcome: open |")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


def _shelf() -> Document:
    document = Document.new()
    document.insert_markdown("Shelf report.\n\n" + SHELF + "\n\nThe end.", at="end")
    return document


def _with_extra_cells(document: Document, row: int = 1, extra: int = 1) -> bytes:
    """``document``'s bytes with ``extra`` cells more in its first table's row ``row``: a row
    off its grid, as a row copied from a table of another grid leaves it."""
    data = document.to_bytes()
    archive = zipfile.ZipFile(io.BytesIO(data))
    root = etree.fromstring(archive.read("word/document.xml"))
    table = root.find(f".//{_W}tbl")
    target = table.findall(f"{_W}tr")[row]
    for _ in range(extra):
        cell = etree.fromstring(f'<w:tc xmlns:w="{_W[1:-1]}"><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>'
                                "<w:p><w:r><w:t>Spare</w:t></w:r></w:p></w:tc>")
        target.append(cell)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as written:
        for info in archive.infolist():
            body = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True) \
                if info.filename == "word/document.xml" else archive.read(info.filename)
            written.writestr(info, body)
    return out.getvalue()


def _open(toolbox, data: bytes):
    session = toolbox.session(clock=CLOCK)
    session.open(data, "shelf.docx")
    return session, session.entry("d1").document


def _fits(document: Document) -> bool:
    return all(not grid_problems(table._entry()[1].element) for table in document.tables())


def _reviewed(document: Document, how: str) -> Document:
    copy = Document.open(document.to_bytes())
    getattr(copy, how)()
    return copy


# -- check says it (rule 7) ------------------------------------------------------------------------


def test_check_names_the_table_and_row_off_the_grid_and_render_draws_it(toolbox):
    session, document = _open(toolbox, _with_extra_cells(_shelf(), extra=2))
    table = document.tables()[0].id
    result = toolbox.dispatch(session, "check", {"doc": "d1", "include": ["validate", "reflow", "layout"]})
    assert result.ok, result.to_json()
    assert "error" not in result.data["reflow"]
    entry = result.data["layout"]["table_grid_mismatch"]
    assert entry == [{"table": table, "grid": 3, "rows": [{"row": 1, "reaches": 5}], "drawn": "widened",
                      "pages": [1]}]
    assert any("grid-row-mismatch" in problem for problem in result.data["validate"]["problems"])
    assert any("table-grid-mismatch" in w for w in result.data["reflow"]["layout_warnings"])
    rendered = toolbox.dispatch(session, "render", {"doc": "d1", "pages": [1]})
    assert rendered.ok, rendered.to_json()


def test_a_tracked_merge_is_off_its_grid_until_accepted_and_check_gives_the_call(toolbox):
    session, document = _open(toolbox, _shelf().to_bytes())
    table = document.tables()[0].id
    assert toolbox.dispatch(session, "word_set_tracking", {"doc": "d1", "on": True, "author": "Shelf Editor"}).ok
    merged = toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "merge",
                                                           "first": {"row": 1, "col": 0}, "last": {"row": 1, "col": 2}})
    assert merged.ok, merged.to_json()
    # Word's own form: the two cells merged away kept, w:cellDel -- five columns' worth over three.
    result = toolbox.dispatch(session, "check", {"doc": "d1", "include": ["layout"]})
    assert result.ok and "error" not in result.data["layout"], result.to_json()
    entry, = result.data["layout"]["table_grid_mismatch"]
    assert entry["rows"] == [{"row": 1, "reaches": 5}] and entry["tracked"] is True
    fix = entry["fix"]
    assert fix == {"tool": "word_changes", "calls": [{"action": "accept", "within": table}]}
    accepted = toolbox.dispatch(session, fix["tool"], {"doc": "d1", **fix["calls"][0]})
    assert accepted.ok, accepted.to_json()
    again = toolbox.dispatch(session, "check", {"doc": "d1", "include": ["layout"]})
    assert again.data["layout"]["table_grid_mismatch"] == [] and _fits(document)


def test_rows_off_the_grid_are_listed_within_the_cap():
    document = Document.open(_with_extra_cells(_shelf(), row=1))
    found = table_grid_mismatch(document)
    assert len(found) == 1 and found[0]["rows"] == [{"row": 1, "reaches": 4}] and "rows_left_out" not in found[0]


# -- a layout that fails is layout_failed, once (rule 2) -------------------------------------------------


def _failing(monkeypatch, error: BaseException) -> list:
    calls = []

    def lay_out_docx(*args, **kwargs):
        calls.append(1)
        raise error

    monkeypatch.setattr(layout_module.docx2svg, "lay_out_docx", lay_out_docx)
    return calls


def test_a_failed_layout_is_layout_failed_naming_the_table_never_not_found(toolbox, monkeypatch):
    session, document = _open(toolbox, _with_extra_cells(_shelf()))
    table = document.tables()[0].id
    _failing(monkeypatch, docx2svg.LayoutError("w:body/w:tbl[1]", "IndexError: list index out of range"))
    rendered = toolbox.dispatch(session, "render", {"doc": "d1", "pages": [1]})
    assert not rendered.ok and rendered.error.code == "layout_failed"
    assert rendered.error.valid_options == [table] and table in rendered.error.message
    # check still says what it can: the validation, and the layout facts read from the document.
    checked = toolbox.dispatch(session, "check", {"doc": "d1", "include": ["validate", "reflow", "layout"]})
    assert checked.ok, checked.to_json()
    assert checked.data["reflow"]["error"].startswith("layout_failed: ") and table in checked.data["reflow"]["error"]
    assert checked.data["layout"]["table_grid_mismatch"][0]["table"] == table


def test_any_failure_in_the_layout_is_layout_failed(toolbox, monkeypatch):
    session, _ = _open(toolbox, _shelf().to_bytes())
    _failing(monkeypatch, IndexError("list index out of range"))
    rendered = toolbox.dispatch(session, "render", {"doc": "d1", "pages": [1]})
    assert not rendered.ok and rendered.error.code == "layout_failed"
    assert "IndexError" in rendered.error.message and rendered.error.valid_options == []


def test_a_version_that_failed_fails_again_at_once_logged_once(toolbox, monkeypatch, caplog):
    from docx_agent.liveview import view

    session, document = _open(toolbox, _shelf().to_bytes())
    calls = _failing(monkeypatch, docx2svg.LayoutError("w:body/w:tbl[1]", "IndexError: boom"))
    with caplog.at_level(logging.WARNING, logger="docx_agent.layout"):
        for _ in range(5):
            with pytest.raises(LayoutFailed) as raised:
                view(session, "d1")
    assert raised.value.at == document.tables()[0].id
    assert len(calls) == 1
    assert len([r for r in caplog.records if r.name == "docx_agent.layout"]) == 1
    # A new version is laid out again.
    document.insert_markdown("Another line.", at="end")
    with pytest.raises(LayoutFailed):
        view(session, "d1")
    assert len(calls) == 2


def test_layout_failed_survives_a_worker():
    import pickle

    failure = LayoutFailed("IndexError: boom", at="t:1A2B3C4D", path="w:body/w:tbl[2]")
    again = pickle.loads(pickle.dumps(failure))
    assert isinstance(again, LayoutFailed) and again.code == "layout_failed"
    assert (again.at, again.path, again.reason) == ("t:1A2B3C4D", "w:body/w:tbl[2]", "IndexError: boom")


# -- the edits keep the grid in step -----------------------------------------------------------------


AUTHOR = {"author": "Shelf Editor", "date": "2026-11-12T09:00:00Z"}


def _tracked(document: Document, *steps) -> None:
    with document.tracking(**AUTHOR):
        for step in steps:
            step(document, document.tables()[0].id)


@pytest.mark.parametrize("name, steps", [
    # a row inserted beside a column whose tracked deletion is pending (accepted: a row a
    # column wider than the grid)
    ("delete column, insert row", [lambda d, t: d.delete_column(t, 1), lambda d, t: d.insert_row(t, 1)]),
    ("delete column, insert row above", [lambda d, t: d.delete_column(t, 2),
                                         lambda d, t: d.insert_row(t, 0, below=False)]),
    # a tracked split records no grid (Word's form); a later tracked change recorded the split
    # grid, which rejecting set back under the unsplit cells
    ("split, then cell width", [lambda d, t: d.split_cell(t, 1, 1, columns=2),
                                lambda d, t: d.set_cell(t, 1, 1, width=90)]),
    ("split, then merge", [lambda d, t: d.split_cell(t, 1, 0, columns=2),
                           lambda d, t: d.merge_cells(t, (1, 0), (2, 1))]),
    ("merge across, then insert row", [lambda d, t: d.merge_cells(t, (1, 0), (1, 2)),
                                       lambda d, t: d.insert_row(t, 1)]),
])
def test_tracked_table_edits_accepted_or_rejected_fit_the_grid(name, steps):
    document = _shelf()
    _tracked(document, *steps)
    for how in ("accept_all", "reject_all"):
        assert _fits(_reviewed(document, how)), (name, how)


def test_a_row_inserted_untracked_beside_a_tracked_column_goes_with_it():
    document = _shelf()
    with document.tracking(**AUTHOR):
        document.insert_column(document.tables()[0].id, 1)
    document.insert_row(document.tables()[0].id, 1)          # not tracked
    rejected = _reviewed(document, "reject_all")
    assert _fits(rejected)
    assert [len(row) for row in rejected.table(rejected.tables()[0].id).grid()] == [3] * 5
    assert _fits(_reviewed(document, "accept_all"))


def test_rows_a_sync_adds_beside_a_tracked_column_deletion_go_with_it(toolbox):
    from ooxml_edit.tools import Dataset, Result, string, tool

    @tool("shelf_rows", "Keep made-up rows.", {"id": string("Its id.")}, documents=())
    def shelf_rows(call, id):
        call.define_dataset(id, Dataset(("item", "room", "notes"),
                                        [("Lamp", "hall", "Outcome: partly done"), ("Rug", "den", "Outcome: open")]),
                            source=None)
        return Result(summary="kept")

    with Toolbox([*TOOLS, shelf_rows], formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session(clock=CLOCK)
        session.open(_shelf().to_bytes(), "shelf.docx")
        document = session.entry("d1").document
        assert box.dispatch(session, "shelf_rows", {"id": "next"}).ok
        assert box.dispatch(session, "word_set_tracking", {"doc": "d1", "on": True, "author": "Shelf Editor"}).ok
        table = document.tables()[0].id
        assert box.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "delete_column",
                                                         "index": 1}).ok
        synced = box.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "sync",
                                                           "from": "d:next", "key": "item"})
        assert synced.ok, synced.to_json()
        assert synced.data["added"]["count"] == 1
        for how in ("accept_all", "reject_all"):
            assert _fits(_reviewed(document, how)), how


def test_a_review_leaves_alone_a_table_it_did_not_touch():
    document = Document.open(_with_extra_cells(_shelf()))      # off its grid, before any change
    document.insert_markdown("Shelf notes.", at="end")
    with document.tracking(**AUTHOR):
        document.insert_markdown("A tracked line.", at="end")
    accepted = _reviewed(document, "accept_all")
    assert table_grid_mismatch(accepted)[0]["rows"] == [{"row": 1, "reaches": 4}]
