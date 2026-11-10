"""Edits on a draft that read another open document's table (round 50): ``from: "t:..."`` and
``{{t:...}}``, as ``d:`` datasets do; braces shaped like a reference that names no source; a
batch's ``replace_text`` without ``expect``.

The other document is a made-up grid (a workbook's stand-in: ``t:<table>`` resolves to its rows),
changed by ``grid_set``.  On a draft the grid is read at its committed version; the change is
applied only while the grid is at that version still, else prepared again -- so the calls, run
from threads, answer and write as the same calls one at a time in the order they were applied
(:func:`block_stress._order`).  Made-up documents and values only.
"""

from __future__ import annotations

import hashlib
import threading

import pytest

from ooxml_edit.tools import Dataset, DocumentFormat, Result, Toolbox, ToolError, integer, string, tool
from ooxml_edit.tools.limits import Limits

from docx_agent import Document, blocks
from docx_agent.tools import FORMAT, GROUPS, TOOLS

import block_stress

pytestmark = pytest.mark.any_faces


class Grid:
    def __init__(self) -> None:
        self.tables = {"stock": (("item", "count", "note"),
                                 [["bolts", 12, "small"], ["nuts", 7, "round"], ["pins", 3, "thin"]])}


def _table(grid: Grid, address: str) -> Dataset:
    name, _, key = address[len("t:"):].partition("/")
    if name not in grid.tables:
        raise ToolError("not_found", f"no table {name!r}", valid_options=sorted(grid.tables))
    header, rows = grid.tables[name]
    rows = [tuple(row) for row in rows]
    if key:
        rows = [row for i, row in enumerate(rows) if key in (f"#{i + 1}", row[0])]
    return Dataset(header, rows, keys=1)


GRID = DocumentFormat(kind="xlsx", open=lambda data: Grid(), detect=lambda data, name: False,
                      datasets={"t:": _table})


@tool("grid_set", "Test only: set a count of the grid.",
      {"doc": string("Document id."), "row": integer("From 1."), "count": integer("The count.")},
      kind="xlsx", mutates=True)
def grid_set(call, doc, row, count):
    call.document.tables["stock"][1][row - 1][1] = count
    return Result(summary=f"row {row}: {count}")


def _box(mode: str) -> Toolbox:
    return Toolbox([*TOOLS, block_stress.ds_rows, grid_set], formats=[FORMAT, GRID], groups=GROUPS, workers=0,
                   block_edits=mode, limits=Limits(render_timeout=600, layout_timeout=600))


def _session(box):
    session = box.session(clock=block_stress.CLOCK)
    doc = session.open(block_stress.made(3), "report.docx")             # d1
    session.adopt(Grid(), "xlsx", "stock.xlsx")                          # d2
    box.dispatch(session, "ds_rows", {"id": "rows", "rows": "6"})
    return session, doc, block_stress.Layout(session.entry(doc).document)


# -- what runs on a draft ------------------------------------------------------------------------


@pytest.mark.parametrize("tool, arguments", [
    ("word_set_text", {"target": "p:1A", "text": "{{t:stock/#1/count}}"}),
    ("word_set_text", {"target": "p:1A", "text": "{{d2/t:stock/#1/count}} and {{d:rows/#1/name}}"}),
    ("word_insert_markdown", {"at": "end", "from": "t:stock"}),
    ("word_insert_markdown", {"at": "end", "from": "d2/t:stock"}),
    ("word_insert_markdown", {"at": "end", "markdown": "{{t:stock/#1/note}}"}),
    ("word_edit_table", {"table": "t:1A", "action": "set_cells", "from": "d2/t:stock"}),
    ("word_edit_table", {"table": "t:1A", "action": "set_cells", "cells": [{"row": 1, "col": 0, "from": "t:stock/#1"}]}),
    ("replace_text", {"find": "a", "replace": "{{t:stock/#1/count}}", "expect": "one", "range": "p:1A"}),
])
def test_another_documents_table_is_read_on_a_draft(tool, arguments):
    assert blocks.admits(tool, arguments) is True


@pytest.mark.parametrize("tool, arguments, alone", [
    ("word_insert_markdown", {"at": "end", "from": "d2/t:stock"}, True),
    ("word_edit_table", {"table": "t:1A", "action": "set_cells", "from": "t:stock"}, True),
    ("word_set_text", {"target": "p:1A", "text": "{{t:stock/#1/count}}"}, False),
])
def test_a_fill_from_another_documents_table_is_worth_a_draft_alone(tool, arguments, alone):
    assert blocks.alone(tool, arguments) is alone


def _calls(doc, layout):
    body = layout.bodies[1]
    return [
        ("word_insert_markdown", {"doc": doc, "at": f"after:{layout.tables[2]}", "from": "t:stock"}),
        ("word_set_text", {"doc": doc, "target": body[0], "text": "Bolts in stock: {{t:stock/bolts/count}}."}),
        ("word_edit_table", {"doc": doc, "table": layout.fill, "action": "set_cells", "grow": True,
                             "from": "t:stock"}),
        ("grid_set", {"doc": "d2", "row": 1, "count": 40}),
        ("word_insert_markdown", {"doc": doc, "at": f"before:{layout.ends[1]}",
                                  "markdown": "Now {{t:stock/#1/count}} bolts, {{d:rows/#2/name}}."}),
        ("batch", {"ops": [
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": body[1], "text": "{{t:stock/#3/note}}"}},
            {"tool": "replace_text", "arguments": {"doc": doc, "find": "thin", "replace": "{{t:stock/#2/note}}",
                                                   "expect": "one", "range": body[1]}}]}),
    ]


def test_a_call_reading_a_table_on_a_draft_answers_and_writes_as_on_the_document():
    out = {}
    for mode in ("off", "always"):
        with _box(mode) as box:
            session, doc, layout = _session(box)
            results = []
            for name, arguments in _calls(doc, layout):
                found = box.dispatch(session, name, arguments)
                assert found.ok, found.to_json()
                results.append((block_stress.normal(found), found.block))
            entry = session.entry(doc)
            out[mode] = ([r for r, _ in results], [b for _, b in results],
                         hashlib.sha256(entry.document.to_bytes()).hexdigest(), entry.document.id_state())
    assert out["off"][0] == out["always"][0]
    assert out["off"][2:] == out["always"][2:]
    blocks_seen = out["always"][1]
    assert blocks_seen[3] is None                   # the grid's own change: no draft
    assert all(block is not None and block.get("lock") is None for k, block in enumerate(blocks_seen) if k != 3), \
        blocks_seen


def test_a_table_changed_while_the_draft_is_prepared_is_read_again(monkeypatch):
    """The grid changes after a draft read it, before its change is applied: prepared again,
    the paragraph gets the value now, and its source's version is the grid's now."""
    from test_block_edits import held

    with _box("always") as box, held(monkeypatch) as hold:
        session, doc, layout = _session(box)
        target = layout.bodies[0][0]
        out: dict = {}
        thread = threading.Thread(target=lambda: out.setdefault("r", box.dispatch(
            session, "word_set_text", {"doc": doc, "target": target, "text": "Bolts: {{t:stock/#1/count}}."})))
        thread.start()
        assert hold.ready.wait(60)
        assert box.dispatch(session, "grid_set", {"doc": "d2", "row": 1, "count": 55}).ok
        hold.release()
        thread.join()
        result = out["r"]
        assert result.ok and result.block["tries"] == 2 and result.block["conflicts"] == ["d2 changed"], result.block
        assert session.entry(doc).document.paragraph(target).text == "Bolts: 55."
        assert result.data["facts"][0]["text"] == "55"


def test_calls_reading_a_table_from_threads_replay_serially():
    """A short replay (the long stress runs are block_stress.py's): three threads edit their
    own chapters by reference to the grid, a fourth changes the grid; the results and the
    document are those of the same calls one at a time, in the order they were applied."""
    data = block_stress.made(3)
    layout = block_stress.Layout(Document.open(data))
    plans = []
    for k in range(3):
        body = layout.bodies[k]
        plans.append([
            ("word_set_text", {"doc": "d1", "target": body[0], "text": f"Chapter {k}: {{{{t:stock/#{k + 1}/count}}}}."}),
            ("word_insert_markdown", {"doc": "d1", "at": f"before:{layout.ends[k]}",
                                      "markdown": f"Note {k}: {{{{t:stock/#{k + 1}/note}}}}, {{{{t:stock/#1/count}}}}."}),
            ("word_set_text", {"doc": "d1", "target": body[1], "text": f"Again {{{{t:stock/#{3 - k}/count}}}}."}),
        ])
    plans.append([("grid_set", {"doc": "d2", "row": 1 + c % 3, "count": 100 + c}) for c in range(4)])
    runs = []
    for _ in range(2):
        with _box("always") as box, block_stress.Recorder() as recorder:
            session, doc, _layout = _session(box)
            made: list[dict] = []
            lock = threading.Lock()

            def worker(k: int) -> None:
                for c, (name, arguments) in enumerate(plans[k]):
                    number = k * 100 + c
                    recorder.local.call = number
                    try:
                        result = box.dispatch(session, name, arguments)
                    finally:
                        recorder.local.call = None
                    with lock:
                        made.append({"call": number, "thread": k, "name": name, "arguments": arguments,
                                     "result": block_stress.normal(result), "version": result.version,
                                     "start": 0, "end": len(recorder.commits), "block": result.block})

            threads = [threading.Thread(target=worker, args=(k,)) for k in range(len(plans))]
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join()
            final = session.entry(doc).document.to_bytes()
            order = block_stress._order(made, list(recorder.commits))
        assert len(order) == sum(len(p) for p in plans)
        with _box("off") as box:
            session, doc, _layout = _session(box)
            again = [block_stress.normal(box.dispatch(session, item["name"], item["arguments"])) for item in order]
            assert [item["result"] for item in order] == again
            assert session.entry(doc).document.to_bytes() == final
        runs.append(sum(1 for item in made if item["block"] and not item["block"].get("lock")))
    assert max(runs) > 0                                            # drafted


# -- braces shaped like a reference, and a document's own braces ---------------------------------


def test_a_reference_without_its_prefix_is_refused_naming_the_one_meant():
    with _box("off") as box:
        session, doc, layout = _session(box)
        result = box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": "end",
                                                                "markdown": "Name: {{rows/#1/name}}"})
        assert not result.ok and result.error.code == "invalid_arguments"
        assert "did you mean {{d:rows/#1/name}}?" in result.error.message
        assert "\\{{rows/#1/name}}" in result.error.message and result.error.field == "markdown"
        handle = box.dispatch(session, "word_set_text", {"doc": doc, "target": layout.bodies[0][0],
                                                         "text": "{{d2/#1/count}}"})
        assert not handle.ok and "{{d2/t:<table>/#1/count}}" in handle.error.message


def test_a_documents_own_braces_and_escaped_ones_stay_text():
    with _box("off") as box:
        session, doc, layout = _session(box)
        target = layout.bodies[0][0]
        written = box.dispatch(session, "word_set_text", {"doc": doc, "target": target,
                                                          "text": "Dear {{name}}, see {{TODO: a/b}}."})
        assert written.ok and written.warnings == []
        paragraph = session.entry(doc).document.paragraph(target)
        assert paragraph.text == "Dear {{name}}, see {{TODO: a/b}}."
        replaced = box.dispatch(session, "replace_text", {"doc": doc, "find": "{{name}}", "replace": "Ann",
                                                          "expect": "one"})
        assert replaced.ok and session.entry(doc).document.paragraph(target).text == "Dear Ann, see {{TODO: a/b}}."
        escaped = box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": "end",
                                                                 "markdown": "Literal \\{{rows/#1/name}} here."})
        assert escaped.ok
        assert session.entry(doc).document.paragraphs()[-1].text == "Literal {{rows/#1/name}} here."


def test_replace_text_in_a_batch_without_expect_replaces_one_as_alone():
    with _box("off") as box:
        session, doc, layout = _session(box)
        target = layout.bodies[0][0]
        box.dispatch(session, "word_set_text", {"doc": doc, "target": target, "text": "One fox, two foxes."})
        batched = box.dispatch(session, "batch", {"ops": [
            {"tool": "replace_text", "arguments": {"doc": doc, "find": "One", "replace": "A", "range": target}}]})
        assert batched.ok and "ops[0]: expect not given: read as 'one'" in batched.warnings
        assert session.entry(doc).document.paragraph(target).text == "A fox, two foxes."
        several = box.dispatch(session, "batch", {"ops": [
            {"tool": "replace_text", "arguments": {"doc": doc, "find": "fox", "replace": "cat", "range": target}}]})
        assert "2" in str(several.to_json()) and session.entry(doc).document.paragraph(target).text == "A fox, two foxes."


def test_an_entry_written_anew_comes_after_the_others_as_on_the_document():
    """The provenance item merged with another call's change: a paragraph's entry whose
    values the call wrote anew (none outlived its new text) goes after the entries others
    added meanwhile, where the call run on the document puts it; one it only changed stays."""
    def item(*entries):
        rows = "".join(f'<prov:source kind="text" key="{key}">{{"v":1,"paragraph":"{text}","facts":[{{"source":'
                       f'"d:x/#1/a","text":"{value}","p":0,"at":{len(text) - len(value)}}}]}}</prov:source>'
                       for key, text, value in entries)
        return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<prov:sources '
                f'xmlns:prov="urn:uvrt:docx-agent:provenance:v1">{rows}</prov:sources>').encode()

    before = item(("p:1", "Old 1", "1"), ("p:2", "Keep 2", "2"))
    after = item(("p:2", "Kept 2", "2"), ("p:1", "New 9", "9"))            # p:1 anew; p:2 edited before its value
    live = item(("p:1", "Old 1", "1"), ("p:2", "Keep 2", "2"), ("p:3", "Theirs 3", "3"))
    merged = blocks._merge_sources(live, before, after)[1].decode()
    keys = [part.split('"')[0] for part in merged.split('key="')[1:]]
    assert keys == ["p:2", "p:3", "p:1"]
