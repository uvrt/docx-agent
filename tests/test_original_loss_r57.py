"""Round 57 (R122, R123): text the file held when opened, taken out of the final view; and no
comment describing a tracked structure change.

R122: agents rebuilt a column (a new one inserted, the old one deleted) where ``set_cells``
would have done; tracked, the old column's values left the final view and nothing said so.
Now ``delete_row``, ``delete_column``, ``sync`` with ``remove`` and ``word_delete`` of a
paragraph, table or span say in the summary and in ``data.original_removed`` when what they
removed held text the document had when opened -- counted, sampled with places, with the
``undo`` call and, tracked, the ``word_changes`` reject that keeps it.  Removing only this
session's own rows, columns and text says nothing.  ``check``'s ``accept_loses_original``
lists what a pending deletion of this session takes out of the final view, with the reject.

R123: a tracked column deletion, merge or split writes no describing comment; a document that
has one keeps it, and deleting the row or column holding it is refused naming it and the
calls that clear the way.  Made-up documents only.

Tool-design rules: 2 (the refusal names the comment and the calls; data carries undo and
reject calls that work), 4 (sample capped, the rest counted), 7 (check says what a person
would miss in the final view).
"""

from __future__ import annotations

import datetime as dt
import io
import re
import zipfile

import pytest
from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool

from docx_agent import Document
from docx_agent.revisions.review import DESCRIBES
from docx_agent.tools import FORMAT, GROUPS, TOOLS

CLOCK = lambda: dt.datetime(2026, 11, 13, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
ME = "Editor A"
TABLE = ("| Site | Crew | Shift | Note |\n|---|---|---|---|\n| Harbour | Ines | Early | Pump check |\n"
         "| Orchard | Tomas | Late |  |\n| Quarry | Lena | Night | Fence repair |")
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


def ok(box, session, name, /, **arguments):
    result = box.dispatch(session, name, {"doc": "d1", **arguments})
    assert result.ok, (name, result.to_json())
    return result


@pytest.fixture(scope="module")
def report(box) -> bytes:
    """A made-up report: a heading, a line, a 4-column table (one empty cell), a line."""
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "crews.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", at=f"after:{first}",
       markdown=f"# Crews\n\nIntro line.\n\n{TABLE}\n\nClosing line.")
    return session.entry("d1").document.to_bytes()


def _without_para_ids(data: bytes) -> bytes:
    source, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w") as target:
        for item in source.infolist():
            content = source.read(item)
            if item.filename == "word/document.xml":
                content = re.sub(rb' w14:(paraId|textId)="[0-9A-Fa-f]+"', b"", content)
            target.writestr(item, content)
    return out.getvalue()


def _open(box, data: bytes, tracked: bool):
    session = box.session(clock=CLOCK)
    session.open(data, "crews.docx")
    if tracked:
        ok(box, session, "word_set_tracking", on=True, author=ME)
    return session, session.entry("d1").document.tables()[0].id


def _texts(session) -> list[str]:
    return [" ".join(c.text for c in row) for row in session.entry("d1").document.tables()[0].grid()] \
        if session.entry("d1").document.tables() else []


# -- R122: the edits say so ------------------------------------------------------------------


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
@pytest.mark.parametrize("ids", [True, False], ids=["paraids", "no_paraids"])
def test_a_column_deleted_names_the_files_text_it_removed(box, report, tracked, ids):
    session, table = _open(box, report if ids else _without_para_ids(report), tracked)
    result = ok(box, session, "word_edit_table", table=table, action="delete_column", index=3)
    found = result.data["original_removed"]
    assert found["cells"] == 3 and found["paragraphs"] == 0      # the empty cell is not counted
    assert [s["text"] for s in found["sample"]] == ["Note", "Pump check", "Fence repair"]
    assert all(s["at"].endswith(",3") for s in found["sample"])
    assert found["undo"] == {"tool": "undo", "arguments": {"doc": "d1"}}
    assert "3 cells holding text the file had when opened left the final view" in result.summary
    assert "'Pump check'" in result.summary and "undo takes the call back" in result.summary
    if not tracked:
        assert "reject" not in found
        return
    reject = found["reject"]
    assert reject["tool"] == "word_changes" and reject["arguments"]["action"] == "reject"
    assert f"word_changes reject ids {reject['arguments']['ids']} keeps them" in result.summary
    # The call given works: the column is back.
    ok(box, session, "word_changes", **{k: v for k, v in reject["arguments"].items() if k != "doc"})
    assert "Pump check" in " ".join(_texts(session))


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_a_row_deleted_names_its_cells_and_undo_takes_it_back(box, report, tracked):
    session, table = _open(box, report, tracked)
    before = _texts(session)
    result = ok(box, session, "word_edit_table", table=table, action="delete_row", index=1)
    found = result.data["original_removed"]
    assert found["cells"] == 4 and [s["at"] for s in found["sample"]] == [f"{table}/c1,{c}" for c in range(4)]
    box.dispatch(session, found["undo"]["tool"], found["undo"]["arguments"])
    assert _texts(session) == before


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_removing_only_this_sessions_own_says_nothing(box, report, tracked):
    session, table = _open(box, report, tracked)
    ok(box, session, "word_edit_table", table=table, action="insert_column", index=3)
    ok(box, session, "word_edit_table", table=table, action="set_cells",
       cells=[{"row": r, "col": 4, "text": f"New {r}"} for r in range(4)])
    gone = ok(box, session, "word_edit_table", table=table, action="delete_column", index=4)
    assert "original_removed" not in gone.data and "file had" not in gone.summary
    session, table = _open(box, report, tracked)
    ok(box, session, "word_edit_table", table=table, action="insert_row", index=3)
    ok(box, session, "word_edit_table", table=table, action="set_cells", cells=[{"row": 4, "col": 0, "text": "Dock"}])
    gone = ok(box, session, "word_edit_table", table=table, action="delete_row", index=4)
    assert "original_removed" not in gone.data
    # Text written this session into the file's empty cell is this session's too.
    ok(box, session, "word_edit_table", table=table, action="set_cells", cells=[{"row": 2, "col": 3, "text": "Fresh"}])
    gone = ok(box, session, "word_edit_table", table=table, action="delete_column", index=3)
    assert [s["text"] for s in gone.data["original_removed"]["sample"]] == ["Note", "Pump check", "Fence repair"]


@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_word_delete_of_a_table_a_paragraph_and_a_span(box, report, tracked):
    session, table = _open(box, report, tracked)
    result = ok(box, session, "word_delete", target=table)
    found = result.data["original_removed"]
    # Rule 4: 15 cells with text, 5 sampled, the rest counted.
    assert found["cells"] == 15 and len(found["sample"]) == 5 and found["sample_left_out"] == 10
    assert "12 more" in result.summary
    session, table = _open(box, report, tracked)
    intro = session.entry("d1").document.paragraphs()[2]
    assert intro.text == "Intro line."
    intro = intro.id
    result = ok(box, session, "word_delete", target=intro)
    assert result.data["original_removed"]["sample"] == [{"at": intro, "text": "Intro line."}]
    session, table = _open(box, report, tracked)
    heading = session.entry("d1").document.paragraphs()[1].id
    result = ok(box, session, "word_delete", target=f"{heading}..{table}", force=True)
    assert (result.data["original_removed"]["paragraphs"], result.data["original_removed"]["cells"]) == (2, 15)
    # Text found and deleted is a text edit, not a block: nothing more.
    session, _ = _open(box, report, tracked)
    assert "original_removed" not in (ok(box, session, "word_delete", find="Closing line.").data or {})


def test_a_paragraph_made_this_session_deleted_says_nothing(box, report):
    session, _ = _open(box, report, True)
    made = ok(box, session, "word_insert_markdown", at="end", markdown="A line of mine.")
    result = ok(box, session, "word_delete", target=made.created[0])
    assert "original_removed" not in (result.data or {})


def test_sync_remove_names_the_rows_text(box, report):
    session, table = _open(box, report, True)
    SETS["crews"] = (("Site", "Crew"), [("Harbour", "Ines"), ("Quarry", "Lena")])
    assert box.dispatch(session, "rows_put", {"id": "crews"}).ok
    result = ok(box, session, "word_edit_table", table=table, action="sync", remove=True, **{"from": "d:crews"})
    found = result.data["original_removed"]
    assert found["cells"] == 3 and [s["text"] for s in found["sample"]] == ["Orchard", "Tomas", "Late"]
    assert found["reject"]["arguments"]["ids"]


# -- R122: check's fact ----------------------------------------------------------------------


def _facts(box, session, **arguments):
    return ok(box, session, "check", include=["layout"], **arguments).data["layout"]


def test_check_lists_the_files_text_a_pending_deletion_removes(box, report):
    session, table = _open(box, report, True)
    # The rebuild: a new column filled, the old one deleted.
    ok(box, session, "word_edit_table", table=table, action="insert_column", index=3)
    ok(box, session, "word_edit_table", table=table, action="set_cells",
       cells=[{"row": r, "col": 4, "text": f"New {r}"} for r in range(4)])
    ok(box, session, "word_edit_table", table=table, action="delete_column", index=3)
    lost = _facts(box, session)["accept_loses_original"]
    assert len(lost) == 1
    entry = lost[0]
    assert entry["table"] == table and entry["cells"] == 3 and entry["author"] == ME
    assert [s["text"] for s in entry["sample"]] == ["Note", "Pump check", "Fence repair"]
    assert entry["fix"] == {"tool": "word_changes", "calls": [{"action": "reject", "ids": [entry["change"]]}]}
    # Scoped to another block: not about it.
    intro = session.entry("d1").document.paragraphs()[2].id
    assert ok(box, session, "check", include=["layout"], range=intro).data["layout"]["accept_loses_original"] == []
    # The fix keeps the text, and the fact is gone.
    ok(box, session, "word_changes", **entry["fix"]["calls"][0])
    assert "Pump check" in " ".join(_texts(session))
    assert _facts(box, session)["accept_loses_original"] == []


def test_check_says_nothing_of_untracked_deletions_or_an_earlier_sessions(box, report):
    session, table = _open(box, report, False)
    ok(box, session, "word_edit_table", table=table, action="delete_column", index=3)
    assert _facts(box, session)["accept_loses_original"] == []
    session, table = _open(box, report, True)
    ok(box, session, "word_edit_table", table=table, action="delete_row", index=1)
    data = session.entry("d1").document.to_bytes()
    later = box.session(clock=lambda: dt.datetime(2026, 11, 14, 9, 0, tzinfo=dt.timezone.utc))
    later.open(data, "crews.docx")
    assert _facts(box, later)["accept_loses_original"] == []   # pending when opened: not this session's
    intro = later.entry("d1").document.paragraphs()[2].id
    ok(box, later, "word_set_tracking", on=True, author="Editor B")
    ok(box, later, "word_delete", target=intro)
    lost = _facts(box, later)["accept_loses_original"]
    assert [(e["author"], e["paras"], e["paragraphs"]) for e in lost] == [("Editor B", [intro], 1)]


# -- R123: no describing comment ---------------------------------------------------------------


def test_tracked_structure_changes_write_no_comment(box, report):
    session, table = _open(box, report, True)
    ok(box, session, "word_edit_table", table=table, action="delete_column", index=3)
    ok(box, session, "word_edit_table", table=table, action="merge", first={"row": 1, "col": 0},
       last={"row": 1, "col": 1})
    ok(box, session, "word_edit_table", table=table, action="split", first={"row": 3, "col": 2}, columns=2)
    assert session.entry("d1").document.comments() == []
    listed = ok(box, session, "word_changes", action="list").data
    assert {"column deletion"} <= {c.get("structure") for c in listed}
    # The column next to it goes too now: nothing anchored in the header row stops it.
    ok(box, session, "word_edit_table", table=table, action="delete_row", index=0)


def _with_old_comment(report: bytes) -> bytes:
    """A document as an earlier version left it: a tracked column deletion and the comment
    that version wrote on the header cell before it."""
    document = Document.open(report)
    table = document.tables()[0]
    with document.tracking(author=ME):
        document.delete_column(table.id, 3)
    cell = document.tables()[0].grid()[0][2]
    document.add_comment(f"{cell.paragraphs[0].id}@0:0", f"{DESCRIBES} column 4 deleted; Word does not track "
                                                         "deleting a column, so it is written as deleted cells.")
    return document.to_bytes()


@pytest.mark.parametrize("action,index", [("delete_row", 0), ("delete_column", 2)])
def test_an_old_describing_comment_is_kept_and_its_refusal_names_the_way_out(box, report, action, index):
    session, table = _open(box, _with_old_comment(report), True)
    comment = session.entry("d1").document.comments()[0].id
    refused = box.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": action, "index": index})
    assert not refused.ok
    message = refused.to_json()["error"]["message"]
    assert f"holds comment {comment} (docx-agent's note on a tracked structure change" in message
    assert f"word_comments {{action: 'delete', items: [{{comment: '{comment}'}}]}}" in message
    assert f"word_changes accept or reject within {table}" in message
    ok(box, session, "word_comments", action="delete", items=[{"comment": comment}])
    ok(box, session, "word_edit_table", table=table, action=action, index=index)


def test_an_old_describing_comment_goes_when_its_table_is_settled(box, report):
    session, table = _open(box, _with_old_comment(report), True)
    assert len(session.entry("d1").document.comments()) == 1
    ok(box, session, "word_changes", action="accept", within=table)
    assert session.entry("d1").document.comments() == []
