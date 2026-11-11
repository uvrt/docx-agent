"""Round 49 (R90): ``word_edit_table`` ``sync`` keeps a table in step with a source's rows by a
key column -- new keys added in the body rows' formatting, changed cells rewritten (equal ones
untouched), keys gone listed or (``remove``) deleted, header and total rows kept, tracked when
tracking is on, values keeping their source, duplicate keys refused.  Made-up tables and data
only."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import zipfile

import pytest

pytest.importorskip("ooxml_edit.tools.sources", reason="needs ooxml-edit 0.22 (sources)")

from lxml import etree  # noqa: E402
from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent import blocks as _blocks  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 11, 10, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
COLUMNS = ("code", "title", "issue", "checked")
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def made(n: int, *, drop: tuple = (), add: tuple = (), change: tuple = ()) -> list[tuple]:
    """Made-up rows ``Q7-001`` ... with a title, an issue and a date checked."""
    out = [(f"Q7-{i:03d}", f"Leaflet {i}", "B" if i in change else "A", f"2026-0{1 + i % 9}-1{i % 10}")
           for i in range(1, n + 1) if i not in drop]
    return out + [(f"Q7-{i:03d}", f"Leaflet {i}", "A", "2026-11-01") for i in add]


SETS: dict[str, tuple[tuple, list[tuple]]] = {}


@tool("rows_put", "Keep a made-up dataset.", {"id": string("Its id.")}, documents=())
def rows_put(call, id):
    columns, rows = SETS[id]
    call.define_dataset(id, Dataset(columns, rows), source=None)
    return Result(summary=f"Kept d:{id}")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, rows_put], formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


def digest(data: bytes) -> str:
    archive = zipfile.ZipFile(io.BytesIO(data))
    found = hashlib.sha256()
    for info in archive.infolist():
        if info.filename != "docProps/core.xml":
            found.update(info.filename.encode() + b"\0" + archive.read(info.filename))
    return found.hexdigest()


def _markdown(rows: list[tuple], head: tuple = COLUMNS) -> str:
    return ("| " + " | ".join(head) + " |\n|" + "---|" * len(head) + "\n"
            + "\n".join("| " + " | ".join(row) + " |" for row in rows))


def _table(toolbox, rows: list[tuple], data: dict[str, list[tuple]], *, head: tuple = COLUMNS,
           columns: tuple = COLUMNS):
    session = toolbox.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "leaflets.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    assert toolbox.dispatch(session, "word_insert_markdown", {"doc": "d1", "at": f"after:{first}",
                                                              "markdown": _markdown(rows, head)}).ok
    for name, values in data.items():
        SETS[name] = (columns, values)
        assert toolbox.dispatch(session, "rows_put", {"id": name}).ok
    return session, session.entry("d1").document.tables()[-1].id


def _sync(toolbox, session, table, **more):
    return toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "sync", **more})


def _texts(session, table):
    view = session.entry("d1").document.table(table)
    return [tuple(view.cell(r, c).text for c in range(view.column_count)) for r in range(len(view.rows))]


def _row_xml(session, table) -> list[bytes]:
    document = session.entry("d1").document
    _part, entry = document._resolve(table)
    return [etree.tostring(row) for row, _ in entry.rows]


# -- the schema ----------------------------------------------------------------------------------


def test_sync_is_an_action_and_the_strict_set_is_unchanged():
    from ooxml_edit.tools.adapters import anthropic_strict_plan

    from docx_agent.tools.format import STRICT_FIRST

    edit = next(t for t in TOOLS if t.name == "word_edit_table")
    assert "sync" in edit.canonical["properties"]["action"]["enum"]
    assert {"key", "remove"} <= set(edit.optional_parameters())
    plan = anthropic_strict_plan(TOOLS, first=STRICT_FIRST)
    assert not plan["word_edit_table"]
    assert sum(len(t.optional_parameters()) for t in TOOLS if plan[t.name]) == 24


# -- the rules -----------------------------------------------------------------------------------


def test_one_added_one_gone_some_changed(toolbox):
    session, table = _table(toolbox, made(12), {"next": made(12, drop=(5,), add=(13,), change=(2, 9))})
    before = _row_xml(session, table)
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert result.ok, result.to_json()
    data = result.data
    assert data["added"] == {"count": 1, "keys": ["Q7-013"], "placed": "key order"}
    assert data["changed"] == {"count": 2, "cells": 2, "keys": {"Q7-002": ["issue"], "Q7-009": ["issue"]}}
    assert data["missing"] == {"count": 1, "keys": ["Q7-005"]} and "removed" not in data
    assert data["unchanged"] == 9 and data["key"] == {"col": 0, "column": "code"}
    texts = _texts(session, table)
    assert texts[1:] == sorted(made(12, add=(13,), change=(2, 9)))    # Q7-005 left (reported)
    # rule 3: rows with nothing changed are byte for byte as they were
    after = _row_xml(session, table)
    for r in (1, 3, 4, 5, 6, 7, 8, 10, 11, 12):
        assert after[r] == before[r], r
    assert result.created and result.created[0].startswith("tr:") and result.changed == [table]
    assert "missing" in result.summary and "remove: true" in result.summary


def test_gone_keys_removed_on_request(toolbox):
    session, table = _table(toolbox, made(8), {"next": made(8, drop=(3, 6))})
    result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True})
    assert result.ok, result.to_json()
    assert result.data["removed"] == {"count": 2, "keys": ["Q7-003", "Q7-006"]}
    assert _texts(session, table)[1:] == made(8, drop=(3, 6))
    assert len(result.removed) == 2 and all(i.startswith("tr:") for i in result.removed)


def test_nothing_to_change_is_no_change(toolbox):
    session, table = _table(toolbox, made(6), {"same": made(6), "fewer": made(6, drop=(2,))})
    version = session.entry("d1").version
    for name in ("same", "fewer"):
        result = _sync(toolbox, session, table, **{"from": f"d:{name}"})
        assert result.ok and result.changed == [] and session.entry("d1").version == version
        assert "nothing changed" in result.warnings[0]
    assert result.data["missing"] == {"count": 1, "keys": ["Q7-002"]}


def test_placed_in_the_source_order_where_the_table_is_not_in_key_order(toolbox):
    shuffled = [made(6)[k] for k in (3, 0, 5, 1, 4, 2)]
    data = shuffled[:2] + [("Q7-100", "Leaflet 100", "A", "2026-11-01")] + shuffled[2:]
    data = [("Q7-050", "Leaflet 50", "A", "2026-11-01")] + data
    session, table = _table(toolbox, shuffled, {"next": data})
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert result.ok, result.to_json()
    assert result.data["added"]["placed"] == "source order"
    assert [t[0] for t in _texts(session, table)[1:]] == [row[0] for row in data]


def test_placed_in_key_order_where_the_table_is(toolbox):
    data = [("Q7-010", "x", "A", "d")] + made(6) + [("Q7-000", "y", "A", "d"), ("Q7-004a", "z", "A", "d")]
    session, table = _table(toolbox, made(6), {"next": data})
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert result.ok and result.data["added"]["placed"] == "key order"
    assert [t[0] for t in _texts(session, table)[1:]] == ["Q7-000", "Q7-001", "Q7-002", "Q7-003", "Q7-004",
                                                          "Q7-004a", "Q7-005", "Q7-006", "Q7-010"]


def test_a_new_row_is_formatted_as_a_body_row(toolbox):
    session, table = _table(toolbox, made(3), {"next": made(4)})
    # the row it is placed after: a cell shaded
    assert toolbox.dispatch(session, "word_format_table", {"doc": "d1", "table": table, "scope": "cell", "row": 3,
                                                           "column": 1, "shading": "DDEEFF"}).ok
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert result.ok, result.to_json()
    xml = _row_xml(session, table)
    assert b'w:fill="DDEEFF"' in xml[4] and b"w:tblHeader" not in xml[4]
    assert _texts(session, table)[4] == made(4)[3]


def test_header_rows_and_the_total_row_are_kept(toolbox):
    rows = made(4) + [("Total", "", "", "")]
    session, table = _table(toolbox, rows, {"next": made(5, drop=(1,))})
    assert toolbox.dispatch(session, "word_format_table", {"doc": "d1", "table": table, "scope": "table",
                                                           "look": ["header_row", "total_row", "banded_rows"]}).ok
    result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True})
    assert result.ok, result.to_json()
    texts = _texts(session, table)
    assert texts[0] == COLUMNS and texts[-1] == ("Total", "", "", "")
    assert texts[1:-1] == made(5, drop=(1,))
    assert "Total" not in str(result.data["removed"])


def test_a_row_with_a_formula_field_is_a_total_row(toolbox):
    # a "Sum" row computing over the table (=SUM(ABOVE)), no look's total row: never matched,
    # added to or removed -- also with remove
    session, table = _table(toolbox, made(4) + [("Sum", "", "", "")], {})
    document = Document.open(session.entry("d1").document.to_bytes())
    cell = document.table(table).cell(5, 2).paragraphs[0].id
    document.insert_field(f"{cell}@0", "=SUM(ABOVE)", result="0")
    session = toolbox.session(clock=CLOCK)
    session.open(document.to_bytes(), "leaflets.docx")
    SETS["next"] = (COLUMNS, made(6, drop=(1,), change=(2,)))
    assert toolbox.dispatch(session, "rows_put", {"id": "next"}).ok
    result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True})
    assert result.ok, result.to_json()
    assert result.data["removed"] == {"count": 1, "keys": ["Q7-001"]} and result.data["total_rows"] == 1
    assert result.data["added"]["keys"] == ["Q7-005", "Q7-006"]
    texts = _texts(session, table)
    assert texts[1:-1] == made(6, drop=(1,), change=(2,)) and texts[-1][0] == "Sum"
    assert b"SUM(ABOVE)" in _row_xml(session, table)[-1]


def test_keys_that_would_match_loosely_are_said(toolbox):
    # round 52: the note names every loose match (case, spaces, separators, leading zeros) and the way out
    session, table = _table(toolbox, made(3), {"next": [("q7-001",) + made(3)[0][1:],
                                                        ("Q7.2",) + made(3)[1][1:], made(3)[2]]})
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert result.ok, result.to_json()
    assert result.data["added"]["keys"] == ["q7-001", "Q7.2"]
    assert result.data["missing"]["keys"] == ["Q7-001", "Q7-002"] and "loose_matches" not in result.data
    assert any("match a source key loosely" in w and "'Q7-001' / 'q7-001', 'Q7-002' / 'Q7.2'" in w
               and "loose: 'table'" in w for w in result.warnings), result.warnings


def test_rows_like_a_row_merged_across_columns_taken_are_refused_first(toolbox):
    # Q7-003 (the merged row) gone from the source, so only the new row, made like it, is refused
    session, table = _table(toolbox, made(3), {"next": made(4, drop=(3,))})
    assert toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "merge",
                                                         "first": {"row": 3, "col": 1},
                                                         "last": {"row": 3, "col": 2}}).ok
    version = session.entry("d1").version
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert not result.ok and result.error.field == "from_columns", result.to_json()
    assert result.error.message.startswith("new rows are made like row 3")
    assert "merged across columns 1, 2" in result.error.message and "title, issue" in result.error.message
    assert result.error.details["from_columns"] == ["code", "title", "", "checked"]
    assert session.entry("d1").version == version                       # nothing changed
    # the way out works
    result = _sync(toolbox, session, table, **{"from": "d:next", "from_columns": ["code", "title", "", "checked"]})
    assert result.ok and result.data["added"]["keys"] == ["Q7-004"], result.to_json()
    # a matched row whose merged cell would take two values: refused the same way
    session, table = _table(toolbox, made(3), {"next": made(3, change=(3,))})
    assert toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "merge",
                                                         "first": {"row": 3, "col": 1},
                                                         "last": {"row": 3, "col": 2}}).ok
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert not result.ok and result.error.message.startswith("row 3 of")
    assert result.error.details["from_columns"] == ["code", "title", "", "checked"]


def test_rows_with_no_key_are_left_and_counted(toolbox):
    session, table = _table(toolbox, made(3) + [("", "a note row", "", "")], {"next": made(3)})
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert result.ok and result.changed == [] and result.data["no_key_rows"] == 1


def test_a_source_row_with_a_blank_key_is_said(toolbox):
    session, table = _table(toolbox, made(2), {"next": made(2) + [("", "no code", "A", "d")]})
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert result.ok and result.changed == []
    assert any("blank key not placed: #3" in w for w in result.warnings)


def test_tracked_when_tracking_is_on(toolbox):
    session, table = _table(toolbox, made(6), {"next": made(6, drop=(4,), add=(7,), change=(1,))})
    assert toolbox.dispatch(session, "word_set_tracking", {"doc": "d1", "on": True, "author": "Editor"}).ok
    result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True})
    assert result.ok, result.to_json()
    xml = _row_xml(session, table)
    assert len(xml) == 8                                        # the deleted row stays, tracked
    assert b"<w:del " in xml[4] and b"<w:ins " in xml[7]
    assert b"<w:delText" in xml[1] and b"<w:ins " in xml[1]       # the issue rewritten, tracked
    # the row deleted as a tracked change is no data row: the next sync has nothing to do
    again = _sync(toolbox, session, table, **{"from": "d:next", "remove": True})
    assert again.ok and again.changed == [], again.to_json()


def test_duplicate_keys_refused_naming_them(toolbox):
    session, table = _table(toolbox, made(3) + [made(3)[1]], {"next": made(3)})
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert not result.ok and "'Q7-002' in rows 2, 4" in result.error.message
    session, table = _table(toolbox, made(3), {"twice": made(3) + [made(3)[0]]})
    result = _sync(toolbox, session, table, **{"from": "d:twice"})
    assert not result.ok and "'Q7-001' in rows #1, #4" in result.error.message
    assert result.error.valid_options == ["Q7-001"]


def test_columns_by_label_or_given(toolbox):
    # labels matched the case aside; a table column with no field left as it is
    head = ("Code", "Title", "Notes", "ISSUE")
    rows = [(c, t, "keep me", i) for c, t, i, _ in made(3)]
    session, table = _table(toolbox, rows, {"next": made(3, change=(2,))}, head=head)
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert result.ok and result.data["mapped"] == {"0": "code", "1": "title", "3": "issue"}
    assert _texts(session, table)[2] == ("Q7-002", "Leaflet 2", "keep me", "B")
    # labels that name no field: refused with the call that works
    head = ("Ref", "Name", "Rev", "When")
    session, table = _table(toolbox, made(3), {"next": made(3)}, head=head)
    result = _sync(toolbox, session, table, **{"from": "d:next"})
    assert not result.ok and result.error.field == "from_columns"
    assert result.error.details["from_columns"] == list(COLUMNS)
    result = _sync(toolbox, session, table, **{"from": "d:next", "from_columns": ["code", "", "issue", "checked"],
                                               "key": "Ref"})
    assert result.ok and result.changed == [], result.to_json()


def test_the_key_column_by_label_index_or_field(toolbox):
    data = [(c, t, i, d) for c, t, i, d in made(4)]
    session, table = _table(toolbox, [(t, c, i, d) for c, t, i, d in data], {"next": data},
                            head=("title", "code", "issue", "checked"))
    for key in ("code", "1", "CODE"):
        result = _sync(toolbox, session, table, **{"from": "d:next", "key": key})
        assert result.ok and result.data["key"]["col"] == 1, (key, result.to_json())
    result = _sync(toolbox, session, table, **{"from": "d:next", "key": "number"})
    assert not result.ok and result.error.field == "key" and "0: title" in result.error.valid_options


def test_a_sync_repeated_with_the_same_key_column_runs_again(toolbox):
    """Round 51: key is the column, not a retry key -- the same sync key on another table, or
    on the same table with new rows, syncs; it is never "already done"."""
    session, table = _table(toolbox, made(3), {"first": made(3, add=(4,)), "then": made(3, add=(4, 5))})
    first = _sync(toolbox, session, table, **{"from": "d:first", "key": "code"})
    then = _sync(toolbox, session, table, **{"from": "d:then", "key": "code"})
    assert first.ok and then.ok and "already done" not in then.summary, then.to_json()
    assert [row[0] for row in _texts(session, table)][-2:] == ["Q7-004", "Q7-005"]
    assert then.version == first.version + 1


def test_values_keep_their_source(toolbox):
    # every data row a source row after: the table's record, its rows in the table's order
    session, table = _table(toolbox, made(5), {"next": made(6, drop=(2,), change=(3,))})
    assert _sync(toolbox, session, table, **{"from": "d:next", "remove": True}).ok
    record = session.entry("d1").document.provenance()[table]
    assert record["source"] == "d:next" and record["first_row"] == 1
    assert [c["col"] for c in record["columns"]] == [0, 1, 2, 3] and record["columns"][0]["rows"] == "1-5"
    # a key left in the table: each cell written keeps its value's source
    session, table = _table(toolbox, made(5), {"next": made(6, drop=(2,), change=(3,))})
    assert _sync(toolbox, session, table, **{"from": "d:next"}).ok
    document = session.entry("d1").document
    assert table not in document.provenance()
    kept = document.text_facts()
    view = document.table(table)
    issue = kept[view.cell(3, 2).paragraphs[0].id]
    assert issue[0]["source"] == "d:next/#2/issue" and issue[0]["text"] == "B"
    assert view.cell(1, 2).paragraphs[0].id not in kept                  # not written: not kept
    assert all(view.cell(6, c).paragraphs[0].id in kept for c in range(4))   # the row added


def test_a_record_the_sync_makes_untrue_is_dropped(toolbox):
    session, table = _table(toolbox, made(4), {"first": made(4), "next": made(5, change=(1,))})
    assert toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "set_cells",
                                                         "from": "d:first"}).ok
    assert table in session.entry("d1").document.provenance()
    session.entry("d1").document.table(table)
    SETS["other"] = (COLUMNS, made(5, drop=(4,)))
    assert toolbox.dispatch(session, "rows_put", {"id": "other"}).ok
    result = _sync(toolbox, session, table, **{"from": "d:other"})         # Q7-004 left: partial
    assert result.ok and any("dropped" in w for w in result.warnings)
    assert table not in session.entry("d1").document.provenance()


def test_lists_capped_with_counts(toolbox):
    session, table = _table(toolbox, made(2), {"many": made(30)})
    result = _sync(toolbox, session, table, **{"from": "d:many"})
    assert result.ok
    added = result.data["added"]
    assert added["count"] == 28 and len(added["keys"]) == 20 and added["more"] == 8
    assert _texts(session, table)[1:] == made(30)


def test_one_pass_and_one_by_one_make_the_same_bytes(toolbox, monkeypatch):
    out = []
    for slow in (False, True):
        with monkeypatch.context() as patch:
            if slow:
                patch.setattr(Document, "_stamping_due", lambda self: True)
            session, table = _table(toolbox, made(9), {"next": made(12, drop=(1, 5), add=(20,), change=(4,))})
            data = [("Q7-000", "zero", "A", "d")] + made(12, drop=(1, 5), change=(4,))
            SETS["next"] = (COLUMNS, data)
            assert toolbox.dispatch(session, "rows_put", {"id": "next"}).ok
            result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True})
            assert result.ok, result.to_json()
        out.append((digest(session.entry("d1").document.to_bytes()), _texts(session, table)))
    assert out[0] == out[1]


def test_a_headerless_table_renamed_by_a_row_above_its_first(toolbox):
    session, table = _table(toolbox, made(3), {"next": [("Q7-000", "zero", "A", "d")] + made(3)})
    deleted = toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "delete_row",
                                                            "index": 0})      # no header row now
    table = deleted.renamed.get(table) or session.entry("d1").document.tables()[-1].id
    # the look's header row off too (else Word draws the first row as a header, and so it is one)
    assert toolbox.dispatch(session, "word_format_table", {"doc": "d1", "table": table, "scope": "table",
                                                           "look": ["banded_rows"]}).ok
    result = _sync(toolbox, session, table, **{"from": "d:next", "from_columns": list(COLUMNS)})
    assert result.ok, result.to_json()
    now = result.data["id"]
    assert now != table and result.renamed == {table: now}
    assert _texts(session, now)[0][0] == "Q7-000"


def test_arguments_read_as_meant(toolbox):
    session, table = _table(toolbox, made(3), {"next": made(4)})
    result = _sync(toolbox, session, table, **{"from": "d:next", "grow": True, "index": 3})
    assert result.ok and any("grow" in w for w in result.warnings) and any("index 3" in w for w in result.warnings)
    result = _sync(toolbox, session, table, cells=[{"row": 1, "col": 1, "text": "x"}], **{"from": "d:next"})
    assert not result.ok and result.error.field == "cells"
    result = _sync(toolbox, session, table)
    assert not result.ok and result.error.field == "from" and "sync" in result.error.message
    result = toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "insert_row",
                                                           "remove": True})
    assert not result.ok and result.error.field == "remove"


def test_dates_in_a_pattern(toolbox):
    data = [(c, t, i, dt.date(2026, 3, 2)) for c, t, i, _ in made(2)]
    session, table = _table(toolbox, [(c, t, i, "2 March 2026") for c, t, i, _ in made(2)], {"next": data})
    result = _sync(toolbox, session, table, **{"from": "d:next", "date_format": "d mmmm yyyy"})
    assert result.ok and result.changed == [], result.to_json()


# -- drafts --------------------------------------------------------------------------------------


def test_a_sync_from_a_dataset_runs_on_a_draft():
    assert _blocks.admits("word_edit_table", {"action": "sync", "table": "t:1A2B3C4D", "from": "d:next"})
    assert _blocks.alone("word_edit_table", {"action": "sync", "table": "t:1A2B3C4D", "from": "d:next"})
    # Round 50 (R100): from another document's table too (test_block_sources.py).
    assert _blocks.admits("word_edit_table", {"action": "sync", "table": "t:1A2B3C4D", "from": "d2/t:x"})
    assert not _blocks.admits("word_edit_table", {"action": "delete_row", "table": "t:1A2B3C4D", "index": 1})


def test_concurrent_syncs_equal_their_commit_order_run_alone():
    import block_stress

    found = block_stress.run(3, threads=4, calls=8, chapters=3, sync=True)
    assert found["same"], found
    assert found["drafted"] > 0


def test_a_sync_on_a_draft_names_only_its_own_table(monkeypatch):
    """Round 51: finding the synced table again after the sync reads the tables before it on
    the way; on a draft those are not the call's -- a sync of an earlier table meanwhile is
    no conflict, and the later sync applies at once (before, "a block near t:... changed",
    prepared again, up to running under the lock)."""
    import threading

    from test_block_edits import held

    with Toolbox([*TOOLS, rows_put], formats=[FORMAT], groups=GROUPS, workers=0, block_edits="always") as box:
        session = box.session(clock=CLOCK)
        session.open(Document.new().to_bytes(), "leaflets.docx")
        first = session.entry("d1").document.paragraphs()[0].id
        two = _markdown(made(2)) + "\n\nBetween the two tables.\n\nAnd more words.\n\n" + _markdown(made(2))
        assert box.dispatch(session, "word_insert_markdown", {"doc": "d1", "at": f"after:{first}", "markdown": two}).ok
        earlier, later = [t.id for t in session.entry("d1").document.tables()][-2:]
        SETS["more"] = (COLUMNS, made(2, add=(3,)))
        SETS["first"] = (COLUMNS, made(2, change=(1,)))
        assert box.dispatch(session, "rows_put", {"id": "more"}).ok
        assert box.dispatch(session, "rows_put", {"id": "first"}).ok
        for table in (earlier, later):          # their records made first (the part both add)
            assert _sync(box, session, table, **{"from": "d:first"}).ok
        with held(monkeypatch) as hold:
            out: dict = {}
            thread = threading.Thread(target=lambda: out.setdefault("r", _sync(box, session, later, **{"from": "d:more"})))
            thread.start()
            assert hold.ready.wait(60)
            assert _sync(box, session, earlier, **{"from": "d:more"}).ok
            hold.release()
            thread.join()
        result = out["r"]
        assert result.ok and result.block["tries"] == 1 and not result.block["conflicts"], result.block
        assert [row[0] for row in _texts(session, later)][-1] == "Q7-003"
        assert [row[0] for row in _texts(session, earlier)][-1] == "Q7-003"
