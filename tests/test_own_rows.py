"""Round 51 (R103): a tracked deletion of a row or table that is the author's own pending
insertion removes it outright, as Word does (``tools/rows_probe.py``,
``tests/observations/rows-word.json``):

* a row wholly one's own insertion (its mark, its cells' marks, every run) goes, no record
  left -- deleted by ``delete_row``, a table's deletion, a ``sync`` removal;
* another author's inserted row is marked deleted beside their insertion, each cell's mark
  too;
* one's own row holding another author's insertion goes, and what was theirs stays, deleted,
  at the start of the next paragraph; one's own row holding what was there before keeps its
  insertion, its content and marks deleted;
* a table goes with its last row.

The same edits Word made in the probe are made here on the same made-up document, and the
two compared shape for shape (records by author, text); Word's Accept All and Reject All of
its result are compared with ours."""

from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path

import pytest
from lxml import etree
from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import rows_probe  # noqa: E402

OBSERVED = json.loads((ROOT / "tests" / "observations" / "rows-word.json").read_text(encoding="utf-8"))
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
ME, OTHER = "Writer", rows_probe.OTHER
CLOCK = lambda: dt.datetime(2026, 11, 11, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
_NS = (f'xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
       'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"')


# -- shapes --------------------------------------------------------------------------------------


def _who(element) -> str:
    author = element.get(_W + "author")
    return "me" if author in ("{author}", ME) else "other"


def _records(properties) -> tuple:
    if properties is None:
        return ()
    return tuple(f"{c.tag[len(_W):]}:{_who(c)}" for c in properties if c.tag in (_W + "ins", _W + "del"))


def _paragraph(p) -> tuple:
    runs = []
    for run in p.iter(_W + "r"):
        chain, node = [], run.getparent()
        while node is not p:
            if node.tag in (_W + "ins", _W + "del"):
                chain.append(f"{node.tag[len(_W):]}:{_who(node)}")
            node = node.getparent()
        text = "".join(x.text or "" for x in run if x.tag in (_W + "t", _W + "delText"))
        if text:
            runs.append((">".join(reversed(chain)), text))
    return _records(p.find(f"{_W}pPr/{_W}rPr")), tuple(runs)


def _block(element) -> tuple:
    if element.tag == _W + "tbl":
        return ("table", tuple((_records(row.find(_W + "trPr")), tuple(_paragraph(p) for p in row.iter(_W + "p")))
                               for row in element.iter(_W + "tr")))
    return ("p", _paragraph(element))


def _observed(name: str) -> list[tuple]:
    out = []
    for block in OBSERVED[name]:
        element = etree.fromstring(re.sub(r"^<(w:\w+)", r"<\1 " + _NS, block, count=1))
        if element.tag != _W + "sectPr":
            out.append(_block(element))
    return out


def _ours(document: Document) -> list[tuple]:
    body = document.package.tree(document.package.document_part()).find(_W + "body")
    return [_block(child) for child in body if child.tag != _W + "sectPr"]


# -- Word's edits, made here ---------------------------------------------------------------------


def _probe_document() -> Document:
    return Document.open(rows_probe.document(ME))


def _made_as_word_made_them(document: Document) -> None:
    """The probe's edits (``rows_probe.script``), tracked by the probe's Word user."""
    tables = {key: t.id for key, t in zip([k for k in rows_probe.TABLES if k != "TableLive"], document.tables())}
    with document.tracking(author=ME, date="2026-11-11T09:00:00Z"):
        here = next(p for p in document.paragraphs() if p.text == "Here")
        made = document.insert_markdown("| x | y |\n|---|---|\n| TableLive typed | |", at=f"before:{here.id}")
        document.delete_block(next(i for i in made.created if i.startswith("t:")))
        document.delete_block(tables["TableMixed"])
        document.delete_block(tables["TableMine"])
        typed = document.table(tables["Typed"])
        second = typed.cell(1, 1).paragraphs[0]
        second.set_text("Typed r2c2 TYPEDTWO")
        second.set_text("")
        first = typed.cell(1, 0).paragraphs[0]
        first.set_text("Typed r2c1 TYPEDONE")
        first.set_text("Typed r2c1")
        document.delete_row(tables["Two"], 2)
        document.delete_row(tables["Two"], 1)
        for key in ("MinePlain", "OtherMineText"):
            document.delete_row(tables[key], 2)
        document.delete_row(tables["MineOtherMid"], 1)
        for key in ("MineOtherText", "Other", "Mine"):
            document.delete_row(tables[key], 2)
        clear = document.table(tables["LiveClear"])
        document.insert_row(tables["LiveClear"], 1)
        cell = clear.cell(2, 0).paragraphs[0]
        cell.set_text("CLEARED")
        cell.set_text("")
        document.insert_row(tables["LiveEmpty"], 1)
        document.delete_row(tables["LiveEmpty"], 2)
        document.insert_row(tables["Live"], 1)
        document.table(tables["Live"]).cell(2, 0).paragraphs[0].set_text("LIVENEW")
        document.delete_row(tables["Live"], 2)


def test_word_removes_ones_own_pending_row_and_table_outright():
    """The measured facts this follows (Word 16.106 for Mac)."""
    tracked = _observed("tracked")
    tables = [b for b in tracked if b[0] == "table"]
    live, empty, clear, mine, other, mine_other, mine_mid, other_mine, plain, two, typed, mixed = tables
    for table in (live, empty, mine):             # inserted and deleted: no record left
        assert all(row[0] == () for row in table[1])
    assert len(live[1]) == len(empty[1]) == 3 and len(mine[1]) == 2
    assert other[1][2][0] == ("ins:other", "del:me")           # another's row: marked deleted
    assert other[1][2][1][0] == (("ins:other", "del:me"), (("ins:other>del:me", "Other r3c1"),))
    assert other_mine[1][2][1][0] == (("ins:other", "del:me"), ())  # one's own text in it: gone
    assert plain[1][2][0] == ("ins:me",)                       # what was there before: deleted
    assert plain[1][2][1][0] == (("del:me",), (("del:me", "MinePlain r3c1"),))
    assert len(mine_other[1]) == 2 and len(mine_mid[1]) == 2   # one's own row holding another's
    after = tracked[tracked.index(mine_other) + 1]
    assert after == ("p", ((), (("ins:other>del:me", "MineOtherText r3c1"), ("", "After MineOtherText."))))
    assert mine_mid[1][1][1][0][1][0] == ("ins:other>del:me", "MineOtherMid r2c1")
    assert [row[0] for row in two[1]] == [(), ("del:me",)]     # own row gone, the other marked
    assert [row[0] for row in mixed[1]] == [("ins:other", "del:me")]
    assert not any(b[0] == "table" and "TableMine" in str(b) for b in tracked)
    assert typed[1][1][1] == (((), (("", "Typed r2c1"),)), ((), (("del:me", "Typed r2c2"),)))


def test_the_probes_edits_made_here_give_words_result():
    document = _probe_document()
    _made_as_word_made_them(document)
    assert _ours(document) == _observed("tracked")
    assert document.validate() == []


@pytest.mark.parametrize("review", ["accept", "reject"])
def test_accepting_or_rejecting_gives_words_review(review):
    document = _probe_document()
    _made_as_word_made_them(document)
    (document.accept_all if review == "accept" else document.reject_all)()
    assert _ours(document) == _observed(review)


# -- through the tools ---------------------------------------------------------------------------


SETS: dict[str, list[tuple]] = {}


@tool("rows_keep", "Keep a made-up dataset.", {"id": string("Its id.")}, documents=())
def rows_keep(call, id):
    call.define_dataset(id, Dataset(("code", "title"), SETS[id]), source=None)
    return Result(summary=f"Kept d:{id}")


@pytest.fixture(scope="module")
def box():
    with Toolbox([*TOOLS, rows_keep], formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def ok(box, session, name, /, **arguments):
    result = box.dispatch(session, name, arguments)
    assert result.ok, (name, result.to_json())
    return result


def _opened(box, rows: list[tuple]):
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "plan.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", doc="d1", at=f"after:{first}",
       markdown="| code | title |\n|---|---|\n" + "\n".join(f"| {a} | {b} |" for a, b in rows))
    ok(box, session, "word_insert_markdown", doc="d1", at="end", markdown="Middle.\n\nTail.")
    return session, session.entry("d1").document.tables()[-1].id


def _rows(session, table) -> tuple[int, int]:
    """The table's rows in the document, and in the final view (less those deleted)."""
    _part, entry = session.entry("d1").document._resolve(table)
    rows = [row for row, _ in entry.rows]
    return len(rows), sum(1 for row in rows if row.find(f"{_W}trPr/{_W}del") is None)


def _xml(session) -> bytes:
    document = session.entry("d1").document
    return etree.tostring(document.package.tree(document.package.document_part()))


def test_a_row_inserted_then_deleted_leaves_no_change(box):
    """The reported case: tracking on, a 3-row table, a row inserted at 2, row 3 deleted."""
    session, table = _opened(box, [("A1", "one"), ("A2", "two")])
    before = _xml(session)
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="insert_row", index=2)
    deleted = ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_row", index=3)
    table = session.entry("d1").document.tables()[-1].id
    assert _rows(session, table) == (3, 3)
    assert ok(box, session, "word_changes", doc="d1", action="list").data == []
    assert _xml(session) == before                    # as it was: no trace of either edit
    assert any(i.startswith("tr:") for i in deleted.removed)


def test_a_row_with_text_typed_into_it_goes_too(box):
    session, table = _opened(box, [("A1", "one"), ("A2", "two")])
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="insert_row", index=1)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
       cells=[{"row": 2, "col": 0, "text": "A9"}, {"row": 2, "col": 1, "text": "nine"}])
    ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_row", index=2)
    assert _rows(session, table) == (3, 3)
    assert session.entry("d1").document.revisions() == []


def test_sync_removing_rows_it_added_earlier_leaves_final_rows(box):
    """Rows a sync added, removed by a later sync in the same pending session: the document's
    rows are the final view's; a row there before is marked deleted (it stays until
    accepted)."""
    base = [(f"K{i:03d}", f"Item {i}") for i in range(1, 21)]
    session, table = _opened(box, base)
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    grown = base + [(f"K{i:03d}", f"Item {i}") for i in range(21, 61)]
    SETS["grown"] = grown
    ok(box, session, "rows_keep", id="grown")
    ok(box, session, "word_edit_table", doc="d1", table=table, action="sync", **{"from": "d:grown", "remove": True})
    table = session.entry("d1").document.tables()[-1].id
    assert _rows(session, table) == (61, 61)
    kept = [r for r in grown if not 30 <= int(r[0][1:]) < 45]
    SETS["kept"] = kept
    ok(box, session, "rows_keep", id="kept")
    result = ok(box, session, "word_edit_table", doc="d1", table=table, action="sync",
                **{"from": "d:kept", "remove": True})
    table = session.entry("d1").document.tables()[-1].id
    assert _rows(session, table) == (len(kept) + 1, len(kept) + 1)
    assert result.data["rows"] == len(kept) + 1
    assert len([i for i in result.removed if i.startswith("tr:")]) == 15
    # One that was there before: marked, as before.
    SETS["less"] = [r for r in kept if r[0] != "K003"]
    ok(box, session, "rows_keep", id="less")
    ok(box, session, "word_edit_table", doc="d1", table=table, action="sync", **{"from": "d:less", "remove": True})
    table = session.entry("d1").document.tables()[-1].id
    assert _rows(session, table) == (len(kept) + 1, len(kept))


def test_another_authors_pending_row_is_marked_deleted_beside_their_insertion(box):
    session, table = _opened(box, [("A1", "one"), ("A2", "two")])
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=OTHER)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="insert_row", index=2)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="set_cells",
       cells=[{"row": 3, "col": 0, "text": "A3"}])
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    ok(box, session, "word_edit_table", doc="d1", table=table, action="delete_row", index=3)
    assert _rows(session, table) == (4, 3)
    _part, entry = session.entry("d1").document._resolve(table)
    row = entry.rows[3][0]
    assert _records(row.find(_W + "trPr")) == ("ins:other", "del:me")
    assert [_paragraph(p) for p in row.iter(_W + "p")] == [
        (("ins:other", "del:me"), (("ins:other>del:me", "A3"),)), (("ins:other", "del:me"), ())]


def test_a_whole_table_one_inserted_goes_with_its_rows(box):
    session, _table = _opened(box, [("A1", "one")])
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    last = session.entry("d1").document.paragraphs()[-1].id
    before = _xml(session)
    made = ok(box, session, "word_insert_markdown", doc="d1", at=f"before:{last}",
              markdown="| new | table |\n|---|---|\n| B1 | b |")
    new = next(i for i in made.created if i.startswith("t:"))
    deleted = ok(box, session, "word_delete", doc="d1", target=new)
    assert new in deleted.removed and len(session.entry("d1").document.tables()) == 1
    assert _xml(session) == before and ok(box, session, "word_changes", doc="d1", action="list").data == []


def test_a_table_with_another_authors_row_keeps_that_row_marked(box):
    """Word's TableMixed: one's own rows go, another author's is marked deleted, the table
    stays with it."""
    session, _table = _opened(box, [("A1", "one")])
    last = session.entry("d1").document.paragraphs()[-1].id
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    made = ok(box, session, "word_insert_markdown", doc="d1", at=f"before:{last}",
              markdown="| new | table |\n|---|---|\n| B1 | b |\n| B2 | c |")
    new = next(i for i in made.created if i.startswith("t:"))
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=OTHER)
    ok(box, session, "word_edit_table", doc="d1", table=new, action="insert_row", index=1)
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    deleted = ok(box, session, "word_delete", doc="d1", target=new)
    assert _rows(session, new) == (1, 0)
    _part, entry = session.entry("d1").document._resolve(new)
    assert _records(entry.rows[0][0].find(_W + "trPr")) == ("ins:other", "del:me")
    assert len([i for i in deleted.removed if i.startswith("tr:")]) == 3


def test_the_first_tracked_call_after_untracked_drafts_drafts_at_once():
    """Switching tracking makes no new version: the version drafts are made from carries
    it, so the next call does not find "tracking was switched" (ooxml-edit 0.109.0)."""
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, block_edits="always") as toolbox:
        session = toolbox.session(clock=CLOCK)
        session.open(Document.new().to_bytes(), "plan.docx")
        first = session.entry("d1").document.paragraphs()[0].id
        made = ok(toolbox, session, "word_insert_markdown", doc="d1", at=f"after:{first}", markdown="Some text.")
        assert made.block
        for on in (True, False, True):
            ok(toolbox, session, "word_set_tracking", doc="d1", on=on, **({"author": ME} if on else {}))
            target = session.entry("d1").document.paragraphs()[-1].id
            result = ok(toolbox, session, "word_set_text", doc="d1", target=target, text=f"Changed {on}.")
            assert result.block and result.block["tries"] == 1 and result.block["conflicts"] == [], result.block


def test_tracked_syncs_and_a_table_deletion_still_run_on_a_draft():
    """Drafted as before (round 50, R97): a sync confined to its table, adding rows and then
    removing them, and a deletion of a known table prepare on a draft; one's own rows go."""
    with Toolbox([*TOOLS, rows_keep], formats=[FORMAT], groups=GROUPS, workers=0,
                 block_edits="always") as toolbox:
        base = [("A1", "one"), ("A2", "two")]
        session, table = _opened(toolbox, base)
        ok(toolbox, session, "word_set_tracking", doc="d1", on=True, author=ME)
        SETS["more"] = base + [("A3", "three"), ("A4", "four")]
        ok(toolbox, session, "rows_keep", id="more")
        added = ok(toolbox, session, "word_edit_table", doc="d1", table=table, action="sync",
                   **{"from": "d:more", "remove": True})
        SETS["back"] = base
        ok(toolbox, session, "rows_keep", id="back")
        removed = ok(toolbox, session, "word_edit_table", doc="d1", table=table, action="sync",
                     **{"from": "d:back", "remove": True})
        last = session.entry("d1").document.paragraphs()[-1].id
        made = ok(toolbox, session, "word_insert_markdown", doc="d1", at=f"before:{last}",
                  markdown="| new | table |\n|---|---|\n| B1 | b |")
        deleted = ok(toolbox, session, "word_delete", doc="d1",
                     target=next(i for i in made.created if i.startswith("t:")))
        # Round 51: the first tracked call after drafts made untracked drafts on its first try
        # too (before: "tracking was switched" three times, then the lock).
        for result in (added, removed, made, deleted):
            assert result.block and result.block["tries"] == 1 and result.block["conflicts"] == [], result.block
        assert _rows(session, table) == (3, 3) and session.entry("d1").document.revisions() == []


def test_a_control_deleted_with_its_table_deletes_the_table_tracked():
    """Before round 51 a table inside a block control deleted with it, tracked, was left as
    it was: its rows are now deleted as Word deletes rows -- one's own pending row outright."""
    from docx_agent.oxml.xml import make

    document = Document.new()
    document.insert_markdown("Before.\n\n| a | b |\n|---|---|\n| 1 | 2 |\n\nAfter.", at="end")
    element = document._resolve(document.tables()[-1].id)[1].element
    sdt = make("w:sdt")
    properties = make("w:sdtPr")
    properties.append(make("w:id"))
    properties[0].set(_W + "val", "4711")
    sdt.append(properties)
    content = make("w:sdtContent")
    element.addprevious(sdt)
    sdt.append(content)
    content.append(element)
    document = Document.open(document.to_bytes())
    table = document.tables()[-1].id
    with document.tracking(author=ME):
        document.insert_row(table, 1)
        document.delete_block("cc:4711")
    rows = list(document.package.tree(document.package.document_part()).iter(_W + "tr"))
    assert len(rows) == 2 and all(_records(row.find(_W + "trPr")) == ("del:me",) for row in rows)
    assert all(_paragraph(p)[0] == ("del:me",) for row in rows for p in row.iter(_W + "p"))
    assert document.validate() == []


# -- the paragraph a table came with (round 51) ------------------------------------------------

SEPARATOR = json.loads((ROOT / "tests" / "observations" / "separator-word.json").read_text(encoding="utf-8"))


def _separator_observed(name: str) -> list[tuple]:
    out = []
    for block in SEPARATOR[name]:
        element = etree.fromstring(re.sub(r"^<(w:\w+)", r"<\1 " + _NS, block, count=1))
        if element.tag != _W + "sectPr":
            out.append(_block(element))
    return out


def test_word_puts_a_new_table_beside_a_table_and_deleting_it_leaves_no_trace():
    """The measured facts (``tools/separator_probe.py``): Word inserts a table next to a table
    with no paragraph between, and deleting it leaves the document as it was; a table
    inserted at the end of a paragraph's text splits it, and the split stays."""
    made, deleted = _separator_observed("made"), _separator_observed("deleted")
    kinds = [b[0] for b in made]
    assert kinds[1:3] == ["table", "table"] and kinds[-3:] == ["table", "table", "p"]
    split = next(k for k, b in enumerate(deleted) if b[0] == "p" and b[1][1] == (("", "Last paragraph."),))
    assert deleted[split][1][0] == ("ins:me",) and deleted[split + 1] == ("p", ((), ()))
    assert [b[0] for b in deleted] == ["p", "table", "p", "table", "p", "p", "p", "table", "p"]


@pytest.mark.parametrize("where", ["between", "end"])
def test_a_table_one_inserted_beside_a_table_goes_with_the_paragraph_it_came_with(where):
    """docx-agent keeps two tables apart with an empty paragraph (inserted, tracked); the
    table deleted, that paragraph goes too -- the document as it was, as Word leaves it."""
    document = Document.open(rows_probe._package(
        rows_probe._paragraph("Start.") + rows_probe._table([rows_probe._row("One", 1)])
        + rows_probe._paragraph() + rows_probe._table([rows_probe._row("Two", 1)]) + rows_probe._paragraph()))
    def body():
        found = document.package.tree(document.package.document_part()).find(_W + "body")
        return [re.sub(r' xmlns:\w+="[^"]*"', "", etree.tostring(child).decode()) for child in found]

    before = body()
    tables = document.tables()
    element = document._resolve(tables[0 if where == "between" else 1].id)[1].element.getnext()
    place = document._index(document.package.document_part()).entry_for(element).id
    with document.tracking(author=ME):
        made = document.insert_markdown("| new | |\n|---|---|\n| | |", at=f"before:{place}")
        new = next(i for i in made.created if i.startswith("t:"))
        assert sum(1 for b in _ours(document) if b == ("p", (("ins:me",), ()))) == 1   # the separator
        result = document.delete_block(new)
    assert body() == before      # (the root keeps the w16du declaration the tracking added)
    assert len(result.removed) == len(set(result.removed)) and new in result.removed


def test_a_paragraph_still_keeping_two_tables_apart_stays():
    document = Document.open(rows_probe._package(
        rows_probe._paragraph("Start.") + rows_probe._table([rows_probe._row("One", 1)]) + rows_probe._paragraph()))
    last = document.paragraphs()[-1].id
    with document.tracking(author=ME):
        first = next(i for i in document.insert_markdown("| a | |\n|---|---|\n| | |", at=f"before:{last}").created
                     if i.startswith("t:"))
        document.insert_markdown("| b | |\n|---|---|\n| | |", at=f"before:{last}")
        document.delete_block(first)
    blocks = [b[0] for b in _ours(document)]
    assert "table" in blocks and all(not (a == b == "table") for a, b in zip(blocks, blocks[1:]))
    assert document.validate() == []
