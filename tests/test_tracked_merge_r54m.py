"""Round 54 follow-up (lane M): a tracked merge in the form Word reviews exactly, and column
positions that count a table as accepted.  Made-up documents only.

* **The merge form** (measured, Word 16.106 for Mac, ``feedback-round54/probe-M``): each
  row's first cell keeps its properties; the cells merged away are ``w:cellDel`` with their
  content and paragraph marks kept; merged down, ``w:cellMerge`` on the rows' first cells.
  Word's Accept All gives a deleted cell that keeps content or a mark to the cell before it
  (its columns and width, its content going) -- an empty cell too -- and Reject All keeps it:
  docx-agent's own accept and reject give the same table, and the pending table stays on its
  grid (``table_grid_mismatch`` empty).  The earlier form (the first cell's span applied, the
  others' content deleted, the whole table's old properties recorded) Word accepted 54 px
  off for a body row merged across, 6 px for a rectangle.
* **Positions as accepted**: ``word_edit_table``'s (and ``word_format_table``'s) ``col``,
  ``col_label``, ``index``, ``first``/``last`` and a ``from`` block's columns skip a column
  whose tracked deletion is pending, as ``word_read``'s final view leaves it out; a position
  a pending merge takes in is the merged cell's.  The result says when it skipped one.

Tool-design rules: 1 (a position read as the reader sees the table), 2 (refusals give the
shape as accepted and the accept call), 3 (an unchanged merge is a no-op), 8 (the form and
the view follow Word's Accept All, measured).
"""

from __future__ import annotations

import datetime as dt

import pytest
from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool

from docx_agent import Document
from docx_agent.layout_facts import table_grid_mismatch
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.validate import grid_problems

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
ME = "Editor A"
CLOCK = lambda: dt.datetime(2026, 11, 12, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
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


def _opened(box, tracked: bool = True, empty: tuple[int, int] | None = None):
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "report.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", doc="d1", at=f"after:{first}",
       markdown=f"Field report.\n\n{MARKDOWN}\n\nAfter the table.")
    table = session.entry("d1").document.tables()[-1].id
    if empty is not None:
        ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
           cells=[{"row": empty[0], "col": empty[1], "text": ""}])
    if tracked:
        ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    return session, table


def _document(session) -> Document:
    return session.entry("d1").document


def _shape(document: Document) -> list[list[tuple]]:
    """Each row's cells: (grid column, span, vertical merge, text)."""
    table = document.tables()[-1]
    _part, entry = table._entry()
    out = []
    for row, _ in entry.rows:
        column, line = 0, []
        for cell in row.iter(_W + "tc"):
            properties = cell.find(_W + "tcPr")
            span = properties.find(_W + "gridSpan") if properties is not None else None
            merge = properties.find(_W + "vMerge") if properties is not None else None
            width = int(span.get(_W + "val")) if span is not None else 1
            text = " / ".join("".join(t.text or "" for t in p.iter(_W + "t")) for p in cell.iter(_W + "p"))
            line.append((column, width, None if merge is None else merge.get(_W + "val", "continue"), text))
            column += width
        out.append(line)
    return out


MERGES = {
    "across": ({"row": 1, "col": 0}, {"row": 1, "col": 1}, None),
    "header": ({"row": 0, "col": 1}, {"row": 0, "col": 3}, None),
    "down": ({"row": 1, "col": 0}, {"row": 2, "col": 0}, None),
    "rectangle": ({"row": 1, "col": 1}, {"row": 2, "col": 2}, None),
    "empty cell": ({"row": 1, "col": 0}, {"row": 1, "col": 1}, (1, 1)),
    "last two": ({"row": 3, "col": 2}, {"row": 3, "col": 3}, None),
}


# -- the merge form ------------------------------------------------------------------------------


@pytest.mark.parametrize("case", list(MERGES))
def test_a_tracked_merge_keeps_the_table_on_its_grid_and_reviews_as_the_untracked_merge(box, case):
    first, last, empty = MERGES[case]
    session, table = _opened(box, empty=empty)
    original = _shape(_document(session))
    ok(box, session, "word_edit_table", doc="d1", table=table, action="merge", first=first, last=last)
    document = _document(session)
    data = document.to_bytes()
    # Pending: the rows reach exactly across the grid (no grid change recorded, no mismatch).
    assert table_grid_mismatch(document) == []
    assert all(grid_problems(t._entry()[1].element) == [] for t in document.tables())
    assert b"tblGridChange" not in data and b"tcPrChange" not in data
    # Accepted here = merged untracked; rejected = the original.
    plain_session, plain_table = _opened(box, tracked=False, empty=empty)
    ok(box, plain_session, "word_edit_table", doc="d1", table=plain_table, action="merge", first=first, last=last)
    accepted = Document.open(data)
    accepted.accept_all()
    assert _shape(accepted) == _shape(_document(plain_session))
    rejected = Document.open(data)
    rejected.reject_all()
    assert _shape(rejected) == original
    for reviewed in (accepted, rejected):
        assert all(grid_problems(t._entry()[1].element) == [] for t in reviewed.tables())


def test_the_cell_merged_away_keeps_its_text_and_marks(box):
    """Word's form: the cell merged away is w:cellDel with its content kept (what makes
    Accept All widen the cell before it); the first cell keeps its span of one."""
    session, table = _opened(box)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
       first={"row": 1, "col": 0}, last={"row": 1, "col": 1})
    root = _document(session).package.tree(_document(session).package.document_part())
    row = list(root.iter(_W + "tr"))[1]
    lead, away = list(row.iter(_W + "tc"))[:2]
    assert lead.find(f"{_W}tcPr/{_W}gridSpan") is None
    assert away.find(f"{_W}tcPr/{_W}cellDel").get(_W + "author") == ME
    assert [t.text for t in away.iter(_W + "t")] == ["Ines"] and not list(away.iter(_W + "delText"))
    assert away.find(f"{_W}p/{_W}pPr/{_W}rPr/{_W}del") is None
    # The text is inserted at the end of the first cell.
    assert [t.text for t in lead.iter(_W + "t")] == ["Harbour", "Ines"]
    assert lead.find(f"{_W}p/{_W}ins") is not None or any(lead.iter(_W + "ins"))


def test_the_final_view_reads_the_pending_merge_as_accepted(box):
    session, table = _opened(box, empty=(1, 1))
    merged = ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
                first={"row": 1, "col": 0}, last={"row": 1, "col": 1})
    # An empty cell merged away is given to the cell before it too (measured: Word widens it).
    assert '<td colspan="2"><p>Harbour</p></td>' in merged.data["markdown"]
    assert merged.data["columns"] == 4 and "columns_deleted_pending" not in merged.data


def test_a_comment_in_the_cell_merged_away_is_written_once_and_reviews_as_word_does(box):
    """Measured: Word's Accept All keeps the comment (in the merged cell); Reject All drops
    it with the inserted copy that holds its reference -- as docx-agent's review does."""
    session, table = _opened(box, tracked=False)
    ok(box, session, "word_comments", doc="d1", action="add", author="Other",
       items=[{"find": "Ines", "within": table, "text": "Check the name."}])
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
       first={"row": 1, "col": 0}, last={"row": 1, "col": 1})
    data = _document(session).to_bytes()
    root = _document(session).package.tree(_document(session).package.document_part())
    for tag in ("commentReference", "commentRangeStart", "commentRangeEnd"):
        ids = [node.get(_W + "id") for node in root.iter(_W + tag)]
        assert len(ids) == len(set(ids)), (tag, ids)
    accepted = Document.open(data)
    accepted.accept_all()
    assert [c.text for c in accepted.comments()] == ["Check the name."]
    rejected = Document.open(data)
    rejected.reject_all()
    assert [c.text for c in rejected.comments()] == []


def test_a_merge_already_made_is_a_no_op(box):
    session, table = _opened(box)
    before = _document(session).to_bytes()
    same = ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
              first={"row": 1, "col": 1}, last={"row": 1, "col": 1})
    assert same.to_json()["changed"] == [] or _document(session).to_bytes() == before


# -- positions count the table as accepted --------------------------------------------------------


def _deleted(box, index: int = 1, tracked: bool = True):
    session, table = _opened(box)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_column", index=index)
    if not tracked:
        ok(box, session, "word_set_tracking", doc="d1", on=False)
    return session, table


def _final_rows(box, session) -> list[str]:
    final = ok(box, session, "word_read", doc="d1").data
    return [line for line in final.splitlines() if line.startswith("| ") and "---" not in line]


def test_the_deletion_says_positions_now_count_the_table_as_accepted(box):
    session, table = _opened(box)
    deleted = ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_column", index=1)
    assert deleted.data["columns"] == 3 and deleted.data["columns_deleted_pending"] == [1]
    assert any("col positions count the table as accepted" in w and f"within {table}" in w
               for w in deleted.warnings)


def test_set_cells_col_skips_a_column_whose_deletion_is_pending(box):
    session, table = _deleted(box)
    written = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                 cells=[{"row": 1, "col": 1, "text": "ready"}, {"row": 2, "col": 0, "text": "Orchard"}])
    assert _final_rows(box, session)[1] == "| Harbour | ready | awaits figures |"
    assert any("grid column 1 ('Lead')" in w and "col 1 is grid column 2" in w for w in written.warnings)
    # The deleted column is untouched: its deletion still pending, Reject All gives it back.
    rejected = Document.open(_document(session).to_bytes())
    rejected.reject_all()
    assert "| Harbour | Ines | draft | awaits figures |" in rejected.to_markdown(ids=False)


def test_a_position_before_the_deleted_column_says_nothing(box):
    session, table = _deleted(box, index=2)
    written = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                 cells=[{"row": 1, "col": 1, "text": "Ines B"}])
    assert not any("col positions" in w for w in written.warnings)
    assert _final_rows(box, session)[1] == "| Harbour | Ines B | awaits figures |"


def test_col_past_the_accepted_columns_is_refused_with_the_shape(box):
    session, table = _deleted(box)
    refused = call(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                   cells=[{"row": 1, "col": 3, "text": "x"}])
    error = refused.to_json()["error"]
    assert error["code"] == "not_found" and "columns 0 to 2 as accepted" in error["message"]
    assert "grid column(s) 1 deleted by a pending tracked change" in error["message"]
    assert error["valid_options"] == ["0", "1", "2"]


def test_col_label_finds_the_columns_as_accepted(box):
    session, table = _deleted(box)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
       cells=[{"row": 2, "col_label": "Remark", "text": "filed"}])
    assert _final_rows(box, session)[2] == "| Orchard | done | filed |"
    missing = call(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                   cells=[{"row": 2, "col_label": "Lead", "text": "x"}])
    assert not missing.ok and "Lead" not in missing.to_json()["error"]["valid_options"]


def test_a_column_fill_counts_the_columns_as_accepted(box):
    session, table = _deleted(box)
    filled = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                cells=[{"col": 2, "rows": "*", "text": "checked"}])
    assert filled.data["filled"][0]["col"] == 2
    assert [line.split(" | ")[-1] for line in _final_rows(box, session)[1:]] == ["checked |"] * 3


def test_a_block_from_a_source_fills_the_columns_as_accepted(box):
    SETS["three"] = (("Area", "Phase", "Remark"), [("North", "late", "one gap"), ("South", "done", "none")])
    session, table = _deleted(box)
    ok(box, session, "rows_put", id="three")
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                index=1, **{"from": "d:three"})
    # the source's header row at row 1, its rows below it
    assert _final_rows(box, session)[2:4] == ["| North | late | one gap |", "| South | done | none |"]
    assert any("as accepted" in w for w in result.warnings)


def test_sync_counts_the_columns_as_accepted(box):
    SETS["sync"] = (("Area", "Phase", "Remark"), [("Harbour", "late", "awaits figures"),
                                                  ("Orchard", "done", "signed off"),
                                                  ("Quarry", "open", "two gaps")])
    session, table = _deleted(box)
    ok(box, session, "rows_put", id="sync")
    synced = ok(box, session, "word_edit_table", doc="d1", table=table, action="sync", key="Area",
                **{"from": "d:sync"})
    assert synced.data["columns"] == 3 and synced.data["changed"]["count"] == 1
    assert _final_rows(box, session)[1] == "| Harbour | late | awaits figures |"


def test_delete_column_index_counts_the_columns_as_accepted(box):
    session, table = _deleted(box)
    deleted = ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_column", index=1)
    assert deleted.data["columns"] == 2 and deleted.data["columns_deleted_pending"] == [1, 2]
    assert _final_rows(box, session)[0] == "| Area | Remark |"


@pytest.mark.parametrize("index, side, header", [
    (0, "after", "| Area | New | Phase | Remark |"),
    (1, "before", "| Area | New | Phase | Remark |"),
    (None, None, "| Area | Phase | Remark | New |"),
])
def test_a_column_inserted_beside_a_pending_deletion_lands_where_accept_all_has_it(box, index, side, header):
    """Defined (and measured: Word's Accept All of these documents is what docx-agent's
    accept gives): the new column goes beside the column named, the pending deleted one on
    its other side."""
    session, table = _deleted(box)
    args = {k: v for k, v in (("index", index), ("side", side)) if v is not None}
    inserted = ok(box, session, "word_edit_table", doc="d1", table=table, action="insert_column", **args)
    new = 3 if index is None else 1
    ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
       cells=[{"row": 0, "col": new, "text": "New"}])
    assert inserted.data["columns"] == 4
    assert _final_rows(box, session)[0] == header
    accepted = Document.open(_document(session).to_bytes())
    accepted.accept_all()
    assert header in accepted.to_markdown(ids=False)
    rejected = Document.open(_document(session).to_bytes())
    rejected.reject_all()
    assert "| Area | Lead | Phase | Remark |" in rejected.to_markdown(ids=False)


def test_word_format_table_column_counts_the_columns_as_accepted(box):
    session, table = _deleted(box)
    formatted = ok(box, session, "word_format_table", doc="d1", table=table, scope="column", column=1,
                   shading="FFEE99")
    assert any("col 1 is grid column 2" in w for w in formatted.warnings)
    root = _document(session).package.tree(_document(session).package.document_part())
    shaded = [k for k, cell in enumerate(list(root.iter(_W + "tr"))[1].iter(_W + "tc"))
              if cell.find(f"{_W}tcPr/{_W}shd") is not None]
    assert shaded == [2]


def test_untracked_a_merge_over_a_pending_deletion_is_refused_with_the_way_out(box):
    session, table = _deleted(box, tracked=False)
    before = _document(session).to_bytes()
    refused = call(box, session, "word_edit_table", doc="d1", table=table, action="merge",
                   first={"row": 1, "col": 0}, last={"row": 1, "col": 1})
    error = refused.to_json()["error"]
    assert error["code"] == "refused"
    assert error["details"]["fix"] == {"tool": "word_changes", "calls": [{"action": "accept", "within": table}]}
    assert _document(session).to_bytes() == before
    # Beside it, untracked, a merge works -- on the columns as accepted.
    merged = ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
                first={"row": 1, "col": 1}, last={"row": 1, "col": 2})
    assert '<td colspan="2"><p>draft</p><p>awaits figures</p></td>' in merged.data["markdown"]


def test_a_position_a_pending_merge_takes_in_is_the_merged_cell(box):
    session, table = _opened(box)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="merge",
       first={"row": 1, "col": 0}, last={"row": 1, "col": 1})
    written = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                 cells=[{"row": 1, "col": 1, "text": "Harbour and Ines"}])
    assert any("(1, 1) is in the cell a pending tracked merge makes of (1, 0)" in w for w in written.warnings)
    assert '<td colspan="2"><p>Harbour and Ines</p></td>' in written.data["markdown"]


def test_without_a_pending_change_positions_are_the_grid_columns(box):
    """Nothing changes for a table without a pending structure change (rule 8: as before)."""
    session, table = _opened(box)
    written = ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                 cells=[{"row": 1, "col": 1, "text": "Ines B"}])
    assert written.data["columns"] == 4 and not written.warnings


# -- Word itself (opt-in: pytest -m oracle) --------------------------------------------------------


@pytest.mark.oracle
@pytest.mark.parametrize("case", ["across", "rectangle", "empty cell", "header"])
@pytest.mark.parametrize("action", ["accept", "reject"])
def test_words_review_of_a_tracked_merge_is_ours(box, case, action):
    """Word's own Accept All / Reject All of the merge form give the table docx-agent's
    review gives: the same cells, spans, merges and text (widths Word re-saves its own way)."""
    import sys
    from pathlib import Path

    import oracle

    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
    from e3_probe import review_script

    first, last, empty = MERGES[case]
    session, table = _opened(box, empty=empty)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="merge", first=first, last=last)
    data = _document(session).to_bytes()
    try:
        with oracle.session() as word:
            outcome = word.run_script(review_script(action), data, name=f"r54m-{case}", tag=action, timeout=200)
            assert outcome, f"Word did not {action}: {outcome.outcome} {outcome.detail}"
            theirs = Document.open(outcome.path.read_bytes())
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))
    ours = Document.open(data)
    ours.accept_all() if action == "accept" else ours.reject_all()
    assert _shape(theirs) == _shape(ours)
