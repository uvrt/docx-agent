"""Round 52 (R106): ``sync`` with keys spelled differently -- ``loose`` matches the keys left
after the exact ones by ``ooxml_edit.tools.loose_key`` (case, white space, separators between
runs, numbers as numbers), keeping the table's spelling or taking the source's; ``match``
pairs keys no rule should guess, first.  Ambiguity refused before any change; tracked when
tracking is on; drafted.  Made-up tables and keys only."""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.keys", reason="needs ooxml-edit 0.110 (loose_key)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent import blocks as _blocks  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from test_table_sync import (CLOCK, COLUMNS, SETS, _markdown, _row_xml, _sync, _table, _texts,  # noqa: E402
                             rows_put)

#: The table's spelling of five keys, and the source's.
SPELT = {"K7-12.3": "k7-12.03", "K7-12.4": "K7_12_4", "DOC-4.2": "DOC-004.2", "DOC-4.3": "doc 4 3",
         "K8-1": "K8.01"}


def rows(keys, issue="A") -> list[tuple]:
    return [(k, f"Leaflet {n}", issue, f"2026-0{1 + n % 9}-1{n % 10}") for n, k in enumerate(keys, 1)]


TABLE_KEYS = ["K7-1", *SPELT, "K9-9"]
DATA_KEYS = ["K7-1", *SPELT.values(), "K9-9"]


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, rows_put], formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


def test_the_parameters_are_optional_and_the_strict_set_unchanged():
    from ooxml_edit.tools.adapters import anthropic_strict_plan

    from docx_agent.tools.format import STRICT_FIRST

    edit = next(t for t in TOOLS if t.name == "word_edit_table")
    assert {"loose", "match"} <= set(edit.optional_parameters())
    assert edit.canonical["properties"]["loose"]["enum"] == ["table", "data"]
    assert "orMap" not in str(edit.canonical)
    plan = anthropic_strict_plan(TOOLS, first=STRICT_FIRST)
    assert not plan["word_edit_table"]
    assert sum(len(t.optional_parameters()) for t in TOOLS if plan[t.name]) == 24


def test_exact_keys_only_by_default(toolbox):
    session, table = _table(toolbox, rows(TABLE_KEYS), {"next": rows(DATA_KEYS)})
    result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True})
    assert result.ok, result.to_json()
    assert result.data["added"]["count"] == 5 and result.data["removed"]["count"] == 5
    assert any("5 table key(s) match a source key loosely" in w for w in result.warnings)


def test_loose_keeps_the_tables_spelling(toolbox):
    session, table = _table(toolbox, rows(TABLE_KEYS), {"next": rows(DATA_KEYS, issue="B")})
    result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True, "loose": "table"})
    assert result.ok, result.to_json()
    data = result.data
    assert data["added"]["count"] == 0 and data["removed"]["count"] == 0
    assert data["loose_matches"] == {"count": 5, "pairs": SPELT}
    assert result.summary.endswith("; 5 key(s) matched loosely")
    assert data["changed"]["count"] == 7 and set(data["changed"]["keys"]) == set(TABLE_KEYS)
    assert [row[0] for row in _texts(session, table)[1:]] == TABLE_KEYS
    assert all(row[2] == "B" for row in _texts(session, table)[1:])


def test_loose_data_takes_the_sources_spelling_tracked(toolbox):
    session, table = _table(toolbox, rows(TABLE_KEYS), {"next": rows(DATA_KEYS)})
    assert toolbox.dispatch(session, "word_set_tracking", {"doc": "d1", "on": True, "author": "Editor"}).ok
    result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True, "loose": "data"})
    assert result.ok, result.to_json()
    assert result.data["loose_matches"]["count"] == 5
    assert result.data["changed"] == {"count": 5, "cells": 5, "keys": {d: ["code"] for d in SPELT.values()}}
    xml = _row_xml(session, table)
    assert len(xml) == 1 + len(TABLE_KEYS)                    # no row added or deleted
    assert b"<w:delText" in xml[2] and b"<w:ins " in xml[2]   # the key rewritten, tracked
    assert b"<w:ins " not in xml[1] and b"<w:ins " not in xml[-1]
    # the next sync finds every key as it is
    again = _sync(toolbox, session, table, **{"from": "d:next", "remove": True, "loose": "data"})
    assert again.ok and again.changed == [] and again.data["loose_matches"]["count"] == 0, again.to_json()


def test_nothing_to_change_is_no_change_and_says_the_matches(toolbox):
    session, table = _table(toolbox, rows(TABLE_KEYS), {"next": rows(DATA_KEYS)})
    result = _sync(toolbox, session, table, **{"from": "d:next", "loose": "table"})
    assert result.ok and result.changed == [], result.to_json()
    assert result.data["loose_matches"]["count"] == 5 and result.data["missing"]["count"] == 0
    assert any("nothing changed" in w for w in result.warnings)


def test_exact_wins_over_loose(toolbox):
    # K7-1 in the table twice spelled; the source's K7-1 is K7-1's, so K7.1 is left missing
    session, table = _table(toolbox, rows(["K7-1", "K7.01", "K9-9"]), {"next": rows(["K7-1", "K9-9"])})
    result = _sync(toolbox, session, table, **{"from": "d:next", "loose": "table"})
    assert result.ok, result.to_json()
    assert result.data["missing"]["keys"] == ["K7.01"] and result.data["loose_matches"]["count"] == 0


def test_a_key_matching_two_is_refused_before_any_change(toolbox):
    session, table = _table(toolbox, rows(["K7-1", "K7.01", "K9-9"]), {"next": rows(["k7_1", "K9-9", "K10"])})
    version = session.entry("d1").version
    result = _sync(toolbox, session, table, **{"from": "d:next", "loose": "table"})
    assert not result.ok and result.error.code == "ambiguous" and result.error.field == "loose", result.to_json()
    assert "'K7-1', 'K7.01' (table) / 'k7_1' (source)" in result.error.message
    assert "match" in result.error.message and result.error.valid_options == ["match"]
    assert session.entry("d1").version == version
    # the source's side too
    session, table = _table(toolbox, rows(["K7-1", "K9-9"]), {"next": rows(["K7.1", "k7_01", "K9-9"])})
    result = _sync(toolbox, session, table, **{"from": "d:next", "loose": "data"})
    assert not result.ok and "'K7-1' (table) / 'K7.1', 'k7_01' (source)" in result.error.message
    # the way out: match pairs one, the other added
    result = _sync(toolbox, session, table, **{"from": "d:next", "loose": "data", "match": {"K7-1": "k7_01"}})
    assert result.ok, result.to_json()
    assert result.data["paired"] == 1 and result.data["added"]["keys"] == ["K7.1"]
    assert any("read as 1 pair(s)" in w for w in result.warnings)
    assert [row[0] for row in _texts(session, table)[1:]][0] == "k7_01"      # data's spelling


def test_match_pairs_keys_no_rule_would(toolbox):
    session, table = _table(toolbox, rows(["K7-1", "OLD-3", "K9-9"]), {"next": rows(["K7-1", "NEW-3", "K9-9"])})
    result = _sync(toolbox, session, table, **{"from": "d:next", "remove": True,
                                               "match": [{"table": "OLD-3", "data": "NEW-3"}]})
    assert result.ok, result.to_json()
    assert result.data["paired"] == 1 and "loose_matches" not in result.data and result.changed == []
    assert result.data["added"]["count"] == 0 and result.data["missing"]["count"] == 0
    assert [row[0] for row in _texts(session, table)[1:]] == ["K7-1", "OLD-3", "K9-9"]


def test_match_written_the_other_way_round_is_read_as_meant(toolbox):
    session, table = _table(toolbox, rows(["K7-1", "OLD-3"]), {"next": rows(["K7-1", "NEW-3"])})
    result = _sync(toolbox, session, table, **{"from": "d:next", "match": [{"table": "NEW-3", "data": "OLD-3"}]})
    assert result.ok and result.data["paired"] == 1, result.to_json()
    assert any("read the other way round" in w for w in result.warnings)


def test_match_naming_no_key_is_refused_with_the_keys(toolbox):
    session, table = _table(toolbox, rows(["K7-1", "OLD-3"]), {"next": rows(["K7-1", "NEW-3"])})
    result = _sync(toolbox, session, table, **{"from": "d:next", "match": {"OLD-4": "NEW-3"}})
    assert not result.ok and result.error.code == "not_found" and result.error.field == "match[0]"
    assert result.error.valid_options == ["K7-1", "OLD-3"]
    result = _sync(toolbox, session, table, **{"from": "d:next", "match": {"OLD-3": "new_03"}})
    assert not result.ok and "loosely equal: 'NEW-3'" in result.error.message
    assert result.error.valid_options[0] == "NEW-3"
    result = _sync(toolbox, session, table, **{"from": "d:next", "match": [{"table": "OLD-3", "data": "NEW-3"},
                                                                           {"table": "K7-1", "data": "NEW-3"}]})
    assert not result.ok and "paired twice" in result.error.message


def test_loose_and_match_are_for_sync(toolbox):
    session, table = _table(toolbox, rows(["K7-1"]), {})
    result = toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "insert_row",
                                                           "loose": "table"})
    assert not result.ok and result.error.field == "loose" and result.error.valid_options == ["sync"]
    # a strict decoder's placeholder is none
    result = toolbox.dispatch(session, "word_edit_table", {"doc": "d1", "table": table, "action": "insert_row",
                                                           "match": []})
    assert result.ok, result.to_json()


def test_loose_matches_capped_with_a_count(toolbox):
    table_keys = [f"K7-{n}" for n in range(1, 31)]
    data_keys = [f"k7.{n:02d}" for n in range(1, 31)]
    session, table = _table(toolbox, rows(table_keys), {"next": rows(data_keys)})
    result = _sync(toolbox, session, table, **{"from": "d:next", "loose": "table"})
    assert result.ok, result.to_json()
    found = result.data["loose_matches"]
    assert found["count"] == 30 and len(found["pairs"]) == 20 and found["more"] == 10


def test_values_keep_their_source_on_loosely_matched_rows(toolbox):
    # every data row a source row after, the loose ones too: the table's record, rows in its order
    session, table = _table(toolbox, rows(TABLE_KEYS), {"next": rows(list(reversed(DATA_KEYS)), issue="B")})
    result = _sync(toolbox, session, table, **{"from": "d:next", "loose": "table"})
    assert result.ok, result.to_json()
    record = session.entry("d1").document.provenance()[result.data["id"]]
    assert record["source"] == "d:next" and record["columns"][0]["rows"] == "7,6,5,4,3,2,1"


def test_a_loose_sync_runs_on_a_draft():
    args = {"action": "sync", "table": "t:1A2B3C4D", "from": "d:next", "loose": "data",
            "match": [{"table": "K7-1", "data": "k7.01"}]}
    assert _blocks.admits("word_edit_table", args) and _blocks.alone("word_edit_table", args)
    with Toolbox([*TOOLS, rows_put], formats=[FORMAT], groups=GROUPS, workers=0, block_edits="always") as box:
        session = box.session(clock=CLOCK)
        session.open(Document.new().to_bytes(), "leaflets.docx")
        first = session.entry("d1").document.paragraphs()[0].id
        assert box.dispatch(session, "word_insert_markdown", {"doc": "d1", "at": f"after:{first}",
                                                              "markdown": _markdown(rows(TABLE_KEYS))}).ok
        SETS["drafted"] = (COLUMNS, rows(DATA_KEYS, issue="B"))
        assert box.dispatch(session, "rows_put", {"id": "drafted"}).ok
        table = session.entry("d1").document.tables()[-1].id
        result = _sync(box, session, table, **{"from": "d:drafted", "loose": "data"})
        assert result.ok and result.block is not None and result.block["tries"] == 1, result.to_json()
        assert [row[0] for row in _texts(session, table)[1:]] == DATA_KEYS
