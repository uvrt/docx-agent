"""Edits of different blocks of one Word document at the same time (round 45, lane B):
``docx_agent.blocks``, ooxml-edit's ``DocumentFormat.blocks`` for Word.

A call that runs on a draft is held between its preparation and its commit (:func:`held`)
while another call commits, so each interleaving is made, not hoped for: edits of other blocks
apply beside it, an edit of the same blocks (or the blocks around them) makes it prepare again,
and everything a patch cannot carry -- or that changed since its version -- keeps the
document's lock.  Each rule of the shared-state inventory (``docx_agent/blocks.py``) has a test
here.  Made-up documents only."""

from __future__ import annotations

import contextlib
import hashlib
import threading

import pytest

from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool
from ooxml_edit.tools.limits import Limits

from docx_agent import Document, blocks
from docx_agent.tools import FORMAT, GROUPS, TOOLS

import block_stress

pytestmark = pytest.mark.any_faces


def _box(mode: str = "always", tries: int = 3, redo: float = 2.0) -> Toolbox:
    return Toolbox([*TOOLS, block_stress.ds_rows], formats=[FORMAT], groups=GROUPS, workers=0, block_edits=mode,
                   limits=Limits(render_timeout=600, layout_timeout=600, block_tries=tries, block_redo_cpu=redo))


@contextlib.contextmanager
def opened(mode: str = "always", tries: int = 3, redo: float = 2.0, first_fill: bool = False):
    """A session on the made-up report (``block_stress.made``), its layout of ids."""
    data = block_stress.made(3, first_fill=first_fill)
    with _box(mode, tries, redo) as box:
        session = box.session(clock=block_stress.CLOCK)
        doc = session.open(data, "report.docx")
        box.dispatch(session, "ds_rows", {"id": "rows", "rows": "6"})
        yield box, session, doc, block_stress.Layout(session.entry(doc).document)


class held:
    """Hold the next call prepared on a draft after its preparation, before its commit, until
    released: ``with held(monkeypatch) as hold: ... hold.ready.wait(); ...; hold.release()``."""

    def __init__(self, monkeypatch) -> None:
        self.ready = threading.Event()
        self.go = threading.Event()
        self.taken = False
        original = blocks.Blocks.prepared
        outer = self

        def prepared(self, draft, result, arguments):
            patch = original(self, draft, result, arguments)
            if not outer.taken:
                outer.taken = True
                outer.ready.set()
                assert outer.go.wait(60)
            return patch

        monkeypatch.setattr(blocks.Blocks, "prepared", prepared)

    def __enter__(self) -> "held":
        return self

    def __exit__(self, *exc) -> None:
        self.go.set()

    def release(self) -> None:
        self.go.set()


def _shape(block):
    """``Result.block`` less ``cpu`` (the drafts' own preparation time)."""
    return {key: value for key, value in (block or {}).items() if key != "cpu"}


def _in_thread(fn):
    out: dict = {}
    thread = threading.Thread(target=lambda: out.setdefault("result", fn()))
    thread.start()
    return thread, out


def _text(session, doc, identifier):
    return session.entry(doc).document.paragraph(identifier).text


# -- what runs on a draft ------------------------------------------------------------------------


@pytest.mark.parametrize("tool, arguments, admitted", [
    ("word_set_text", {"target": "p:1A", "text": "x"}, True),
    ("word_set_text", {"items": [{"target": "$intro", "text": "x {{d:rows/#1/name}}"}]}, True),
    ("word_set_text", {"target": "p:1A..p:2B", "text": "x"}, False),
    ("word_set_text", {"target": "p:1A", "text": "{{t:Costs/#1/name}}"}, False),
    ("word_insert_markdown", {"at": "before:p:1A", "markdown": "# T"}, True),
    ("word_insert_markdown", {"at": "end", "markdown": "text"}, True),
    ("word_insert_markdown", {"at": "replace:p:1A", "markdown": "text"}, False),
    ("word_insert_markdown", {"at": "end", "markdown": "![a](b1)"}, False),
    ("word_insert_markdown", {"at": "end", "markdown": "a[^1]\n\n[^1]: note"}, False),
    ("word_insert_markdown", {"at": "end", "from": "d:rows"}, True),
    ("word_insert_markdown", {"at": "end", "from": "t:Costs"}, False),
    ("word_edit_table", {"table": "t:1A", "action": "set_cells", "from": "d:rows"}, True),
    ("word_edit_table", {"table": "t:1A", "action": "set_cells", "cells": [{"row": 0, "col": 0, "from": ""}]}, True),
    ("word_edit_table", {"table": "t:1A", "action": "insert_row"}, False),
    # Round 48 (R83): edits of known blocks.
    ("word_format", {"targets": ["p:1A"], "bold": True}, True),
    ("word_format", {"targets": ["p:1A@2:7", "p:1A..t:2B"], "italic": True}, True),
    ("word_format", {"find": [{"text": "x", "within": "p:1A..p:2B"}], "bold": True}, True),
    ("word_format", {"find": [{"text": "x"}], "bold": True}, False),               # anywhere: the whole document
    ("word_format", {"targets": ["s:2"], "bold": True}, False),                     # a section
    ("word_format", {"targets": ["h:p:1A"], "bold": True}, False),                  # a heading's range
    ("word_format", {"targets": ["$intro"], "bold": True}, True),                   # a ref: resolved on the draft
    ("word_format", {"targets": ["Revenue"], "bold": True}, False),                 # a bookmark
    ("replace_text", {"find": "a", "replace": "b", "expect": "all", "range": "p:1A..p:2B"}, True),
    ("replace_text", {"find": "a", "replace": "b", "expect": "all"}, False),         # the whole document
    ("replace_text", {"find": "a", "replace": "b", "expect": "all", "range": "p:1A", "stories": "all"}, False),
    ("word_format_table", {"table": "t:1A", "scope": "table", "width_of": "body"}, True),
    ("word_delete", {"target": "p:1A..p:2B"}, True),
    ("word_delete", {"find": "x", "within": "p:1A"}, True),
    ("word_delete", {"find": "x"}, False),
    ("word_links", {"action": "add_hyperlink", "target": "p:1A@0:4", "url": "https://example.com"}, True),
    ("word_links", {"action": "add_bookmark", "target": "p:1A", "bookmark": "here"}, True),
    ("word_links", {"action": "add_hyperlink", "find": "x", "url": "https://example.com"}, False),
    ("word_links", {"action": "remove_bookmark", "bookmark": "here"}, False),
    ("word_fields", {"action": "update"}, False),
])
def test_what_a_draft_takes(tool, arguments, admitted):
    assert blocks.admits(tool, arguments) is admitted


def test_one_caller_alone_runs_as_before():
    with opened("auto") as (box, session, doc, layout):
        result = box.dispatch(session, "word_set_text", {"doc": doc, "target": layout.bodies[0][0], "text": "Alone."})
        assert result.ok and result.block is None


@pytest.mark.parametrize("tool, arguments, alone", [
    ("word_edit_table", {"table": "t:1A", "action": "set_cells", "from": "d:rows"}, True),
    ("word_edit_table", {"table": "t:1A", "action": "set_cells", "cells": [{"row": 1, "col": 0, "from": "d:rows"}]}, True),
    ("word_edit_table", {"table": "t:1A", "action": "set_cells", "cells": [{"row": 1, "col": 0, "text": "x"}]}, False),
    ("word_insert_markdown", {"at": "end", "from": "d:rows"}, True),
    ("word_insert_markdown", {"at": "end", "markdown": "text {{d:rows/#1/name}}"}, False),
    ("word_set_text", {"target": "p:1A", "text": "x"}, False),
])
def test_a_fill_from_a_dataset_is_worth_a_draft_alone(tool, arguments, alone):
    assert blocks.alone(tool, arguments) is alone


def test_one_caller_alone_fills_from_a_dataset_on_a_draft():
    """``auto``: a fill from a dataset, which may run long, drafts with no other edit in
    flight, so an edit that comes while it runs need not wait for it; its result and bytes
    are those of the same fill on the document."""
    outcomes = []
    for mode in ("auto", "off"):
        with opened(mode) as (box, session, doc, layout):
            result = box.dispatch(session, "word_edit_table", {"doc": doc, "table": layout.fill, "action": "set_cells",
                                                               "grow": True, "from": "d:rows"})
            assert result.ok, result.to_json()
            assert (result.block is not None) is (mode == "auto")
            outcomes.append((block_stress.normal(result), session.entry(doc).document.to_bytes()))
    assert outcomes[0] == outcomes[1]


def test_a_slow_fill_runs_once_when_calls_that_take_the_lock_come_meanwhile(monkeypatch):
    """The team's slow fill: a fill from a dataset (on a draft, alone) still preparing when
    calls that take the document's lock come -- a style changed (a part the fill's patch
    must find unchanged) and the fields updated.  They wait for the fill to apply, so it
    runs once, and they change the version it made."""
    import time

    from docx_agent.tools import objects

    original = objects._set_from
    runs = []

    def slow(*args, **kwargs):
        runs.append(1)
        time.sleep(2)
        return original(*args, **kwargs)

    monkeypatch.setattr(objects, "_set_from", slow)
    with opened("auto", redo=0.3) as (box, session, doc, layout):
        out = {}
        thread = threading.Thread(target=lambda: out.setdefault("fill", box.dispatch(session, "word_edit_table", {
            "doc": doc, "table": layout.fill, "action": "set_cells", "grow": True, "from": "d:rows"})))
        thread.start()
        time.sleep(0.8)
        style = box.dispatch(session, "word_styles", {"doc": doc, "action": "modify", "name": "Normal", "size": 12})
        fields = box.dispatch(session, "word_fields", {"doc": doc, "action": "update"})
        thread.join()
        fill = out["fill"]
        assert fill.ok and _shape(fill.block) == {"tries": 1, "conflicts": []}, fill.block
        assert len(runs) == 1
        assert style.ok and style.block is not None and style.block["waited"] > 0.5
        assert style.version == fill.version + 1
        assert fields.ok and fields.version >= style.version


def test_a_call_on_a_draft_answers_and_writes_as_the_call_on_the_document():
    """Every admitted kind, run alone on a draft and on the document: the same results (what
    they say about time aside), the same touched, the same bytes and id state."""
    def calls(doc, layout):
        body = layout.bodies[1]
        return [
            ("word_insert_markdown", {"doc": doc, "at": f"before:{layout.ends[1]}",
                                      "markdown": "### A note\n\nText.\n\n- one\n- two\n\n1. first\n2. second"}),
            ("word_set_text", {"doc": doc, "target": body[0], "text": "New text **strong** here."}),
            ("word_set_text", {"doc": doc, "target": body[1], "text": "Two\nlines"}),
            ("word_set_text", {"doc": doc, "target": body[2], "text": "From the rows: {{d:rows/#3/name}}."}),
            ("word_edit_table", {"doc": doc, "table": layout.fill, "action": "set_cells", "grow": True,
                                 "from": "d:rows"}),
            ("word_edit_table", {"doc": doc, "table": layout.tables[1], "action": "set_cells",
                                 "cells": [{"row": 1, "col": 1, "text": "9", "from": ""}]}),
            ("word_insert_markdown", {"doc": doc, "at": f"after:{layout.tables[2]}",
                                      "markdown": "| A | B |\n|---|---|\n| 1 | 2 |"}),
            ("batch", {"ops": [
                {"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": "end", "markdown": "The end."}},
                {"tool": "word_set_text", "arguments": {"doc": doc, "target": body[3], "text": "Batched."}},
                {"tool": "word_read", "arguments": {"doc": doc, "range": f"{layout.heads[1]}..{layout.ends[1]}"}}]}),
            ("word_set_text", {"doc": doc, "target": body[0], "text": "New text **strong** here."}),
            ("word_set_text", {"doc": doc, "target": "p:7FFFFFF0", "text": "gone"}),
        ]

    out = {}
    for mode in ("off", "always"):
        with opened(mode) as (box, session, doc, layout):
            results = []
            for name, arguments in calls(doc, layout):
                found = box.dispatch(session, name, arguments)
                results.append((block_stress.normal(found), found.block))
            entry = session.entry(doc)
            out[mode] = ([r for r, _ in results], [b for _, b in results],
                         hashlib.sha256(entry.document.to_bytes()).hexdigest(), entry.document.id_state())
    assert out["off"][0] == out["always"][0]
    assert out["off"][2:] == out["always"][2:]
    drafted = out["always"][1]
    assert all(b is not None for b in drafted)
    assert [b.get("lock") for b in drafted] == [None] * len(drafted)


# -- other blocks apply beside it; the same blocks make it prepare again -------------------------------


def test_an_edit_of_another_chapter_applies_beside_a_call_on_a_draft(monkeypatch):
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        mine, theirs = layout.bodies[0][0], layout.ends[2]
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_set_text",
                                                       {"doc": doc, "target": mine, "text": "Mine, held."}))
        assert hold.ready.wait(60)
        fast = box.dispatch(session, "word_insert_markdown",
                            {"doc": doc, "at": f"before:{theirs}", "markdown": "Theirs, first."})
        hold.release()
        thread.join()
        result = slow["result"]
        assert fast.ok and result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
        assert result.version == fast.version + 1
        assert _text(session, doc, mine) == "Mine, held."
        entry = session.entry(doc)
        assert entry.document.paragraph(fast.created[0]).text == "Theirs, first."
        assert not entry.document.validate()


def test_ids_drawn_meanwhile_are_drawn_again_and_the_result_says_the_documents(monkeypatch):
    """Both calls make paragraphs: the one held drew the same values the other drew first, so
    its ids are drawn again when it applies -- never an id the document has -- and its result
    names the paragraphs the document has."""
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_insert_markdown", {
            "doc": doc, "at": f"before:{layout.ends[0]}", "markdown": "Held one.\n\nHeld two."}))
        assert hold.ready.wait(60)
        fast = box.dispatch(session, "word_insert_markdown",
                            {"doc": doc, "at": f"before:{layout.ends[2]}", "markdown": "First one.\n\nFirst two."})
        hold.release()
        thread.join()
        result = slow["result"]
        assert result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
        document = session.entry(doc).document
        assert [document.paragraph(i).text for i in result.created] == ["Held one.", "Held two."]
        assert [document.paragraph(i).text for i in fast.created] == ["First one.", "First two."]
        assert not set(result.created) & set(fast.created)
        assert result.data["blocks"] == result.created
        ids = [p.id for p in document.paragraphs()]
        assert len(ids) == len(set(ids)) and not document.validate()


def test_a_first_fill_while_others_draw_ids_is_prepared_once_its_added_part_renamed(monkeypatch):
    """A document's first fill from a dataset adds the provenance item (a part); another
    builder draws ids while the fill is held on its draft.  The fill's ids are drawn again
    when it applies -- in the body and in the part it adds -- and it is prepared once: every
    row's id is the document's, and every place the provenance names is one the document has."""
    with opened(first_fill=True) as (box, session, doc, layout), held(monkeypatch) as hold:
        document = session.entry(doc).document
        assert not any(name.startswith("/customXml/") or name.startswith("customXml/")
                       for name in document.package.part_names)
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_edit_table", {
            "doc": doc, "table": layout.fill, "action": "set_cells", "grow": True, "from": "d:rows"}))
        assert hold.ready.wait(60)
        fast = box.dispatch(session, "word_set_text",       # draws textIds, as the fill does
                            {"doc": doc, "target": layout.bodies[2][0], "text": "Changed meanwhile."})
        hold.release()
        thread.join()
        result = slow["result"]
        assert fast.ok and result.ok, result.to_json()
        assert _shape(result.block) == {"tries": 1, "conflicts": []}
        document = session.entry(doc).document
        ids = [p.id for p in document.paragraphs()]
        assert len(ids) == len(set(ids)) and not document.validate()
        provenance = session.provenance(doc)
        assert provenance
        table = document.table(layout.fill)
        for address in provenance:
            assert address.split("/")[0] == layout.fill, address
        assert table.cell(1, 0).text and table.cell(6, 0).text
        assert any(name.endswith("customXml/item1.xml") for name in document.package.part_names)
        concurrent = document.to_bytes()
    # The same two calls one at a time, in the order they committed: the same bytes (the
    # item's properties name it by a hash of its first key, renamed too).
    with opened(first_fill=True) as (box, session, doc, layout):
        assert box.dispatch(session, "word_set_text", {"doc": doc, "target": layout.bodies[2][0],
                                                       "text": "Changed meanwhile."}).ok
        assert box.dispatch(session, "word_edit_table", {"doc": doc, "table": layout.fill, "action": "set_cells",
                                                         "grow": True, "from": "d:rows"}).ok
        assert session.entry(doc).document.to_bytes() == concurrent


def test_an_edit_of_the_same_paragraph_makes_the_call_prepare_again(monkeypatch):
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        target = layout.bodies[1][1]
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_set_text",
                                                       {"doc": doc, "target": target, "text": "Held."}))
        assert hold.ready.wait(60)
        assert box.dispatch(session, "word_set_text", {"doc": doc, "target": target, "text": "First."}).ok
        hold.release()
        thread.join()
        result = slow["result"]
        assert result.ok and result.block["tries"] == 2
        assert result.block["conflicts"][0].startswith("a block near")
        assert _text(session, doc, target) == "Held."


def test_an_insert_right_beside_a_paragraph_does_not_make_its_edit_prepare_again(monkeypatch):
    """Round 47 (R78): a paragraph's text set depends on that paragraph alone -- blocks put
    right after it meanwhile leave the change as it was."""
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        target = layout.bodies[1][1]
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_set_text",
                                                       {"doc": doc, "target": target, "text": "Held."}))
        assert hold.ready.wait(60)
        beside = box.dispatch(session, "word_insert_markdown",
                              {"doc": doc, "at": f"after:{target}", "markdown": "Right after."})
        assert beside.ok
        hold.release()
        thread.join()
        result = slow["result"]
        assert result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
        assert _text(session, doc, target) == "Held."
        ids = [p.id for p in session.entry(doc).document.paragraphs()]
        assert ids[ids.index(target) + 1] == beside.created[0]


def _neighbours(first, then, *, mode="always"):
    """``first`` held on its draft while ``then`` runs (each ``(name, arguments)`` made of
    the doc and layout); both results and the body's paragraph texts after."""
    import pytest as _pytest

    monkeypatch = _pytest.MonkeyPatch()
    try:
        with opened(mode) as (box, session, doc, layout), held(monkeypatch) as hold:
            name, arguments = first(doc, layout)
            thread, slow = _in_thread(lambda: box.dispatch(session, name, arguments))
            assert hold.ready.wait(60)
            other = box.dispatch(session, *then(doc, layout))
            hold.release()
            thread.join()
            document = session.entry(doc).document
            return slow["result"], other, document, layout
    finally:
        monkeypatch.undo()


@pytest.mark.parametrize("first_at, then_at", [
    ("before:{end}", "before:{end}"),        # the same place, both before the same block
    ("after:{last}", "before:{end}"),        # the same gap, from either side
    ("before:{end}", "after:{last}"),
])
def test_inserts_into_one_gap_apply_in_the_order_they_committed(first_at, then_at):
    """An insert before a block depends on that block, one after a block on that one: two
    put into the same gap meanwhile both apply, once, each where it would have gone had it
    run on the document as it was then (the later before:X right before X, the later after:P
    right after P)."""
    def at(text, layout):
        return text.format(end=layout.ends[1], last=layout.tables[1])     # the chapter's table, then its end

    held_call, other, document, layout = _neighbours(
        lambda doc, layout: ("word_insert_markdown", {"doc": doc, "at": at(first_at, layout), "markdown": "Held."}),
        lambda doc, layout: ("word_insert_markdown", {"doc": doc, "at": at(then_at, layout), "markdown": "Other."}))
    assert other.ok and held_call.ok
    assert _shape(held_call.block) == {"tries": 1, "conflicts": []}
    texts = [p.text for p in document.paragraphs()]
    end = texts.index(document.paragraph(layout.ends[1]).text)
    got = texts[end - 2:end]
    # "Other." committed first; the held insert then runs, in the serial order, on it.
    expected = {("before:{end}", "before:{end}"): ["Other.", "Held."],
                ("after:{last}", "before:{end}"): ["Held.", "Other."],
                ("before:{end}", "after:{last}"): ["Other.", "Held."]}[(first_at, then_at)]
    assert got == expected


def test_a_table_put_beside_a_table_put_there_meanwhile_prepares_again():
    """A table reads its neighbours (beside a table it takes a separator): a table put in the
    same gap meanwhile makes the held one prepare again, so the separator is there."""
    table = "| A | B |\n|---|---|\n| 1 | 2 |"
    held_call, other, document, layout = _neighbours(
        lambda doc, layout: ("word_insert_markdown", {"doc": doc, "at": f"before:{layout.ends[1]}", "markdown": table}),
        lambda doc, layout: ("word_insert_markdown", {"doc": doc, "at": f"before:{layout.ends[1]}", "markdown": table}))
    assert other.ok and held_call.ok
    assert held_call.block["tries"] == 2 and held_call.block["conflicts"]
    assert not document.validate()


def test_an_insert_whose_anchor_changed_meanwhile_prepares_again():
    held_call, other, document, layout = _neighbours(
        lambda doc, layout: ("word_insert_markdown", {"doc": doc, "at": f"after:{layout.bodies[1][0]}",
                                                      "markdown": "Held."}),
        lambda doc, layout: ("word_set_text", {"doc": doc, "target": layout.bodies[1][0], "text": "Changed."}))
    assert other.ok and held_call.ok and held_call.block["tries"] == 2


def test_past_its_tries_a_call_takes_the_documents_lock(monkeypatch):
    with opened(tries=1) as (box, session, doc, layout), held(monkeypatch) as hold:
        target = layout.bodies[1][1]
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_set_text",
                                                       {"doc": doc, "target": target, "text": "Held."}))
        assert hold.ready.wait(60)
        assert box.dispatch(session, "word_set_text", {"doc": doc, "target": target, "text": "First."}).ok
        hold.release()
        thread.join()
        result = slow["result"]
        assert result.ok and result.block["lock"] == "conflicts"
        assert _text(session, doc, target) == "Held."


# -- the shared-state inventory, rule by rule ------------------------------------------------------


def _held_then(monkeypatch, first, then, *, mode="always"):
    """``first`` (name, arguments) held on its draft while ``then`` runs; both results."""
    with opened(mode) as (box, session, doc, layout), held(monkeypatch) as hold:
        name, arguments = first(doc, layout)
        thread, slow = _in_thread(lambda: box.dispatch(session, name, arguments))
        assert hold.ready.wait(60)
        other_name, other_arguments = then(doc, layout)
        other = box.dispatch(session, other_name, other_arguments)
        hold.release()
        thread.join()
        return slow["result"], other, session, doc


def test_a_change_to_another_part_since_makes_it_prepare_again(monkeypatch):
    """styles.xml (a style changed by a whole-document call): every part but the body must be
    as it was."""
    result, other, _s, _d = _held_then(
        monkeypatch,
        lambda doc, layout: ("word_set_text", {"doc": doc, "target": layout.bodies[0][0], "text": "Held."}),
        lambda doc, layout: ("word_styles", {"doc": doc, "action": "modify", "name": "Normal", "size": 12}))
    assert other.ok, other.to_json()
    assert result.ok and result.block["tries"] == 2 and "styles.xml changed" in result.block["conflicts"][0]


def test_a_comment_added_since_makes_it_prepare_again(monkeypatch):
    result, other, _s, _d = _held_then(
        monkeypatch,
        lambda doc, layout: ("word_set_text", {"doc": doc, "target": layout.bodies[0][0], "text": "Held."}),
        lambda doc, layout: ("word_comments", {"doc": doc, "action": "add", "items": [
            {"target": layout.bodies[2][0], "text": "A remark."}]}))
    assert other.ok, other.to_json()
    assert result.ok and result.block["tries"] == 2
    assert "comments" in result.block["conflicts"][0] or "a block near" in result.block["conflicts"][0]


def test_a_value_written_by_reference_meanwhile_is_merged_not_a_conflict(monkeypatch):
    """Round 47 (R78): the provenance item keeps an entry per table and paragraph; another
    call's entry added meanwhile is merged with the call's own -- each where the calls, run
    in the order they committed, put it -- not a conflict."""
    result, other, session, doc = _held_then(
        monkeypatch,
        lambda doc, layout: ("word_set_text", {"doc": doc, "target": layout.bodies[0][0],
                                               "text": "Held: {{d:rows/#2/name}}."}),
        lambda doc, layout: ("word_set_text", {"doc": doc, "target": layout.bodies[2][0],
                                               "text": "From the rows: {{d:rows/#4/name}}."}))
    assert other.ok and result.ok
    assert _shape(result.block) == {"tries": 1, "conflicts": []}
    facts = session.entry(doc).document.text_facts()
    assert list(facts)[-2:] == [other.changed[0], result.changed[0]]


def test_sources_merge_entry_by_entry():
    from docx_agent.blocks import _merge_sources
    from docx_agent.edit.provenance import NS

    def item(*entries):
        inner = "".join(f'<prov:source kind="{k}" key="{key}">{text}</prov:source>' for k, key, text in entries)
        return f'<prov:sources xmlns:prov="{NS}">{inner}</prov:sources>'.encode()

    before = item(("table", "t:1", "a"), ("text", "p:2", "b"))
    mine = item(("table", "t:1", "A"), ("text", "p:3", "c"))                       # t:1 changed, p:2 dropped, p:3 new
    theirs = item(("table", "t:1", "a"), ("text", "p:2", "b"), ("text", "p:9", "z"))   # p:9 new
    _live, merged = _merge_sources(theirs, before, mine)
    from ooxml_edit.xml import parse_xml, serialize
    assert merged == serialize(parse_xml(item(("table", "t:1", "A"), ("text", "p:9", "z"), ("text", "p:3", "c"))))
    with pytest.raises(blocks.BlockConflict):         # both changed t:1
        _merge_sources(item(("table", "t:1", "x"), ("text", "p:2", "b")), before, mine)
    with pytest.raises(blocks.BlockConflict):         # both added p:3
        _merge_sources(item(("table", "t:1", "a"), ("text", "p:2", "b"), ("text", "p:3", "q")), before, mine)


def test_a_section_change_since_makes_it_prepare_again(monkeypatch):
    """The body's last section (its page size, margins: a table's width) must be as it was."""
    result, other, _s, _d = _held_then(
        monkeypatch,
        lambda doc, layout: ("word_insert_markdown", {"doc": doc, "at": f"before:{layout.ends[0]}",
                                                      "markdown": "| A | B |\n|---|---|\n| 1 | 2 |"}),
        lambda doc, layout: ("word_sections", {"doc": doc, "action": "set", "section": "s:body",
                                               "orientation": "landscape"}))
    assert other.ok, other.to_json()
    assert result.ok and result.block["tries"] == 2


def test_tracked_changes_keep_the_lock():
    with opened() as (box, session, doc, layout):
        assert box.dispatch(session, "word_set_tracking", {"doc": doc, "on": True, "author": "Writer"}).ok
        result = box.dispatch(session, "word_set_text", {"doc": doc, "target": layout.bodies[0][0], "text": "Tracked."})
        assert result.ok and "tracked" in result.block["lock"]
        assert session.entry(doc).document.revisions()


def test_a_document_not_yet_stamped_takes_the_lock_for_its_first_edit():
    document = Document.new()
    document.paragraphs()[0].set_text("One.")
    data = document.to_bytes()
    with _box("always") as box:
        session = box.session(clock=block_stress.CLOCK)
        doc = session.open(Document.open(data).to_bytes(), "plain.docx")
        first = session.entry(doc).document.paragraphs()[0].id
        result = box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"after:{first}", "markdown": "Two."})
        assert result.ok
        assert result.block.get("lock") in (None, "ids stamped")


def test_facts_of_the_whole_document_are_recorded_and_checked_again(monkeypatch):
    """A Markdown table in a document with no default table style takes the style its tables
    use most -- a fact of the whole document, recorded on the draft and found the same when
    the patch applies, else a conflict (as is any such fact: the list definition the
    document's lists use most, whether the body is one empty paragraph)."""
    from ooxml_edit.tools.blocks import BlockConflict

    seen: dict = {}
    original = blocks.Blocks.prepared

    def spy(self, draft, result, arguments):
        seen["patch"] = original(self, draft, result, arguments)
        return seen["patch"]

    monkeypatch.setattr(blocks.Blocks, "prepared", spy)
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_insert_markdown", {
            "doc": doc, "at": f"before:{layout.ends[0]}", "markdown": "| A | B |\n|---|---|\n| 1 | 2 |"}))
        assert hold.ready.wait(60)
        patch = seen["patch"]
        assert patch.facts == {("table_style_in_use", ()): "TableGrid"}
        entry = session.entry(doc)
        with entry.lock:
            blocks.BLOCKS.check(entry.document, patch, frozenset())          # as it is: applies
            patch.facts[("table_style_in_use", ())] = "TableNormal"
            with pytest.raises(BlockConflict, match="table style in use changed"):
                blocks.BLOCKS.check(entry.document, patch, frozenset())
            patch.facts[("table_style_in_use", ())] = "TableGrid"
        hold.release()
        thread.join()
        assert slow["result"].ok and _shape(slow["result"].block) == {"tries": 1, "conflicts": []}


def test_a_dataset_redefined_meanwhile_makes_it_prepare_again(monkeypatch):
    result, other, session, doc = _held_then(
        monkeypatch,
        lambda doc, layout: ("word_edit_table", {"doc": doc, "table": layout.fill, "action": "set_cells",
                                                 "grow": True, "from": "d:rows"}),
        lambda doc, layout: ("ds_rows", {"id": "rows", "rows": "4"}))
    assert other.ok and result.ok
    assert result.block["tries"] == 2 and result.block["conflicts"] == ["dataset d:rows changed"]
    table = session.entry(doc).document.table(result.data["id"] if "id" in (result.data or {}) else result.changed[0])
    assert len(table.rows) == 5                     # the header and the dataset's 4 rows now


def test_a_ref_used_and_redefined_meanwhile_makes_it_prepare_again(monkeypatch):
    with opened() as (box, session, doc, layout):
        assert box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{layout.ends[0]}",
                                                              "markdown": "Named.", "ref": "named"}).ok
        hold = held(monkeypatch)
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_set_text",
                                                       {"doc": doc, "target": "$named", "text": "Through the ref."}))
        assert hold.ready.wait(60)
        assert box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{layout.ends[2]}",
                                                              "markdown": "Other.", "ref": "other"}).ok
        hold.release()
        thread.join()
        result = slow["result"]
        assert result.ok and result.block["tries"] == 2 and result.block["conflicts"] == ["the document's refs changed"]


def test_a_retry_key_used_meanwhile_answers_as_done(monkeypatch):
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        arguments = {"doc": doc, "at": f"before:{layout.ends[0]}", "markdown": "Once.", "key": "k1"}
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_insert_markdown", dict(arguments)))
        assert hold.ready.wait(60)
        first = box.dispatch(session, "word_insert_markdown", dict(arguments))
        hold.release()
        thread.join()
        result = slow["result"]
        assert first.ok and result.ok and result.summary.startswith("already done")
        texts = [p.text for p in session.entry(doc).document.paragraphs()]
        assert texts.count("Once.") == 1


def test_each_commit_is_one_version_and_one_step_of_its_callers(monkeypatch):
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        a, b = session.view(context={"agent": "a"}), session.view(context={"agent": "b"})
        thread, slow = _in_thread(lambda: box.dispatch(a, "word_set_text", {
            "doc": doc, "target": layout.bodies[0][0], "text": "A's."}))
        assert hold.ready.wait(60)
        fast = box.dispatch(b, "word_set_text", {"doc": doc, "target": layout.bodies[2][0], "text": "B's."})
        hold.release()
        thread.join()
        result = slow["result"]
        entry = session.entry(doc)
        assert (fast.version, result.version) == (entry.version - 1, entry.version)
        assert entry.owners[fast.version].caller == {"agent": "b"}
        assert entry.owners[result.version].caller == {"agent": "a"}
        # B's own undo is refused: A's change came after it, and the document undoes whole.
        refused = box.dispatch(b, "undo", {"doc": doc})
        assert not refused.ok and refused.error.code == "entangled", refused.to_json()
        assert "word_set_text by agent=a" in refused.error.message
        assert _text(session, doc, layout.bodies[2][0]) == "B's."
        undone = box.dispatch(a, "undo", {"doc": doc})            # A's is the latest: undone
        assert undone.ok and _text(session, doc, layout.bodies[0][0]) != "A's."
        assert _text(session, doc, layout.bodies[2][0]) == "B's."


def test_checks_and_reads_are_of_the_version_the_change_made(monkeypatch):
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        thread, slow = _in_thread(lambda: box.dispatch(session, "batch", {"ops": [
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": layout.bodies[0][0], "text": "Held."}},
            {"tool": "word_read", "arguments": {"doc": doc, "range": f"{layout.heads[2]}..{layout.ends[2]}"}}]}))
        assert hold.ready.wait(60)
        fast = box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{layout.ends[2]}",
                                                              "markdown": "Seen by the held batch's read."})
        hold.release()
        thread.join()
        result = slow["result"]
        assert result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
        read = result.data["ops"][1]
        assert "Seen by the held batch's read." in str(read["data"]) and "version" not in read
        assert result.checks["validate"]["new"] == []
        assert fast.version < result.version


def test_formats_without_the_hook_are_unchanged():
    import dataclasses

    plain = dataclasses.replace(FORMAT, blocks=None)
    with Toolbox(TOOLS, formats=[plain], groups=GROUPS, workers=0, block_edits="always") as box:
        session = box.session(clock=block_stress.CLOCK)
        doc = session.open(block_stress.made(3), "report.docx")
        layout = block_stress.Layout(session.entry(doc).document)
        result = box.dispatch(session, "word_set_text", {"doc": doc, "target": layout.bodies[0][0], "text": "x"})
        assert result.ok and result.block is None


def test_a_cells_empty_from_is_no_source():
    """Lane D's find: a strict decoder's ``from: ""`` beside a cell's text is no source."""
    with opened("off") as (box, session, doc, layout):
        result = box.dispatch(session, "word_edit_table", {"doc": doc, "table": layout.tables[0], "action": "set_cells",
                                                           "cells": [{"row": 1, "col": 1, "text": "7", "from": ""}],
                                                           "from": ""})
        assert result.ok, result.to_json()
        assert session.entry(doc).document.table(layout.tables[0]).cell(1, 1).text == "7"


def test_a_body_read_from_an_earlier_state_reads_as_the_whole_body():
    """``read_body`` reads only the blocks between the first and last byte that differ from
    the body it read last; what it finds is what a read of the whole part finds."""
    import random

    with opened("off") as (_box, session, doc, _layout):
        document = session.entry(doc).document
        data = document.package.read(document.package.document_part())
    base = blocks._read_body(data)
    assert base is not None
    rng = random.Random(45)
    reread = 0
    for _ in range(300):
        k = rng.randrange(base.count)
        last = min(base.count, k + rng.randrange(1, 4))
        begin, end = base.spans[k][0], base.spans[last - 1][1]
        choice = rng.random()
        if choice < 0.3:
            new = data[:begin] + data[end:]
        elif choice < 0.6:
            new = data[:begin] + data[begin:end].replace(b"e", b"E", 3) + data[end:]
        elif choice < 0.7:
            j = rng.randrange(base.count)
            copy = data[base.spans[j][0]:base.spans[j][1]]
            new = data[:begin] + copy.replace(b'paraId="', b'paraId="7', 1) + data[begin:]
        elif choice < 0.85:
            j = rng.randrange(base.count)      # an exact copy: its ids twice, not stamped
            new = data[:begin] + data[base.spans[j][0]:base.spans[j][1]] + data[begin:]
        else:
            cut = rng.randrange(len(data))
            new = data[:cut] + b"<w:p>" + data[cut:]
        whole, found = blocks._read_body(new), blocks._reread_body(new, base)
        if found is None:
            continue
        reread += 1
        assert whole is not None
        assert (found.start, found.end, found.spans, found.ids, found.stamps) == \
            (whole.start, whole.end, whole.spans, whole.ids, whole.stamps)
    assert reread > 100


# -- "changed while this call read it": only what the read covered (round 47) ------------------


def _changed_by(edit):
    """The ooxml-edit Change one edit of the made-up report makes, and its layout of ids."""
    document = Document.open(block_stress.made(3))
    layout = block_stress.Layout(document)
    with document.history.watch() as change:
        edit(document, layout)
    return change, document, layout


def _scope_of(document, *ids):
    from docx_agent.scope import ReadScope

    return ReadScope(document.package.document_part(), frozenset(ids))


def test_a_change_touches_a_read_of_blocks_it_changed_or_put_between():
    change, document, layout = _changed_by(lambda d, l: d.paragraph(l.bodies[1][0]).set_text("New."))
    assert blocks.touches(change, _scope_of(document, layout.bodies[1][0]))
    assert not blocks.touches(change, _scope_of(document, layout.bodies[0][0], layout.ends[0]))
    change, document, layout = _changed_by(
        lambda d, l: d.insert_markdown("Between.", at=f"after:{l.bodies[1][0]}"))
    ids = [p.id for p in document.paragraphs()]
    span = ids[ids.index(layout.heads[1]):ids.index(layout.ends[1]) + 1]
    # Put inside the span read; right after its last block (not inside); far away.
    assert blocks.touches(change, _scope_of(document, *span))
    assert not blocks.touches(change, _scope_of(document, layout.bodies[1][0]))
    assert not blocks.touches(change, _scope_of(document, layout.heads[2], layout.ends[2]))


def test_a_change_before_a_range_touches_a_read_of_its_pages_one_after_it_does_not():
    change, document, layout = _changed_by(lambda d, l: d.paragraph(l.bodies[0][0]).set_text("New."))
    ids = [p.id for p in document.paragraphs()]
    later = ids[ids.index(layout.heads[1]):ids.index(layout.ends[1]) + 1]
    from docx_agent.scope import ReadScope

    part = document.package.document_part()
    assert blocks.touches(change, ReadScope(part, frozenset(later), layout=True))
    assert not blocks.touches(change, ReadScope(part, frozenset(later), layout=False))
    change, document, layout = _changed_by(lambda d, l: d.paragraph(l.bodies[2][0]).set_text("New."))
    assert not blocks.touches(change, ReadScope(part, frozenset(later), layout=True))


def test_a_change_to_a_style_touches_any_read():
    change, document, layout = _changed_by(lambda d, l: d.styles.modify("Normal", size=12))
    assert blocks.touches(change, _scope_of(document, layout.bodies[0][0]))


def test_a_ranged_check_is_warned_only_when_its_blocks_changed_meanwhile(monkeypatch):
    """A check of one chapter, laid out with the lock let go: an edit of another chapter
    meanwhile is not news to it; one of its own chapter is ("changed while this call read
    it"); a check of the whole document hears of either."""
    import contextlib
    import time

    from ooxml_edit.tools import dispatch

    original = dispatch.Call.unlocked

    @contextlib.contextmanager
    def slow(self):
        with original(self) as released:
            if released:
                time.sleep(1.0)
            yield released

    monkeypatch.setattr(dispatch.Call, "unlocked", slow)

    def run(arguments, target):
        with opened("auto") as (box, session, doc, layout):
            out = {}
            reader = threading.Thread(target=lambda: out.setdefault("check", box.dispatch(
                session, "check", {"doc": doc, **arguments(layout)})))
            reader.start()
            time.sleep(0.4)
            assert box.dispatch(session, "word_set_text", {"doc": doc, "target": target(layout),
                                                           "text": "Meanwhile."}).ok
            reader.join()
            result = out["check"]
            assert result.ok, result.to_json()
            return any("changed while this call read it" in w for w in result.warnings)

    chapter = lambda layout: {"range": [f"{layout.heads[1]}..{layout.ends[1]}"]}      # noqa: E731
    assert not run(chapter, lambda layout: layout.bodies[2][0])      # a later chapter
    assert run(chapter, lambda layout: layout.bodies[1][0])          # its own
    assert run(chapter, lambda layout: layout.bodies[0][0])          # an earlier one: its pages may move
    assert run(lambda layout: {}, lambda layout: layout.bodies[2][0])


@tool("ds_notes", "Test only: notes with Markdown bodies.", {"id": string("Its id.")}, documents=())
def _ds_notes(call, id):
    call.define_dataset(id, Dataset(("title", "body"), [("One", "Some **text** here.\n\n- a\n- b"),
                                                        ("Two", "More text.")]), source=None)
    return Result(summary="kept")


@pytest.mark.parametrize("mode", ["auto", "always"])
def test_one_caller_filling_an_all_but_empty_document_prepares_each_call_once(mode):
    """Round 47: a call that writes text by reference looks up the paragraphs it made, and
    0.1.74 then looked for them in the document -- where its patch has not put them yet:
    "is gone" three times, and the call ran under the lock.  Alone on a document of one
    paragraph, inserts and fills from a dataset are each prepared once."""
    from ooxml_edit.tools.limits import Limits as _Limits

    data = Document.new()
    data.paragraphs()[0].set_text("Title")
    with Toolbox([*TOOLS, _ds_notes, block_stress.ds_rows], formats=[FORMAT], groups=GROUPS, workers=0,
                 block_edits=mode, limits=_Limits(render_timeout=600, layout_timeout=600)) as box:
        session = box.session(clock=block_stress.CLOCK)
        doc = session.open(data.to_bytes(), "report.docx")
        box.dispatch(session, "ds_notes", {"id": "notes"})
        box.dispatch(session, "ds_rows", {"id": "rows", "rows": "4"})
        calls = [{"at": "end", "from": "d:notes/#1/body"}, {"at": "end", "from": "d:notes/#2/body"},
                 {"at": "end", "from": "d:rows"}, {"at": "end", "markdown": "Plain {{d:notes/#1/title}}."}]
        for arguments in calls:
            result = box.dispatch(session, "word_insert_markdown", {"doc": doc, **arguments})
            assert result.ok, result.to_json()
            if mode == "always" or result.block is not None:
                assert _shape(result.block) == {"tries": 1, "conflicts": []}, (arguments, result.block)


def test_a_first_value_by_reference_in_a_new_paragraph_renamed_names_its_item_again(monkeypatch):
    """The provenance item's properties name it by a hash of the key it was made for (the
    new paragraph's id): drawn again when the call applies, the hash follows the new id --
    the bytes those of the two calls run one at a time in the order they committed."""
    held_args = {"at": None, "markdown": "Value {{d:rows/#1/name}}."}
    with opened(first_fill=True) as (box, session, doc, layout), held(monkeypatch) as hold:
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_insert_markdown", {
            "doc": doc, **held_args, "at": f"before:{layout.ends[1]}"}))
        assert hold.ready.wait(60)
        other = box.dispatch(session, "word_insert_markdown",
                             {"doc": doc, "at": f"before:{layout.ends[2]}", "markdown": "Other one.\n\nOther two."})
        hold.release()
        thread.join()
        result = slow["result"]
        assert other.ok and result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
        assert set(result.created).isdisjoint(other.created)
        concurrent = session.entry(doc).document.to_bytes()
    with opened(first_fill=True) as (box, session, doc, layout):
        assert box.dispatch(session, "word_insert_markdown",
                            {"doc": doc, "at": f"before:{layout.ends[2]}", "markdown": "Other one.\n\nOther two."}).ok
        assert box.dispatch(session, "word_insert_markdown", {"doc": doc, **held_args,
                                                              "at": f"before:{layout.ends[1]}"}).ok
        assert session.entry(doc).document.to_bytes() == concurrent


def test_a_call_that_draws_ids_then_fails_spends_them_as_under_the_lock():
    """Round 47 (lane F's find): an atomic batch whose blocks are made, then an op fails --
    under the lock its ids are spent, so the next call draws others; on a draft it would not
    be.  It runs under the lock instead: the next call's ids and the id state the same."""
    found = []
    for mode in ("off", "always"):
        with opened(mode) as (box, session, doc, layout):
            failed = box.dispatch(session, "batch", {"atomic": True, "ops": [
                {"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": f"before:{layout.ends[0]}",
                                                               "markdown": "A.\n\nB."}},
                {"tool": "word_set_text", "arguments": {"doc": doc, "target": "p:7FFFFFF0", "text": "x"}}]})
            assert not failed.ok
            if mode == "always":
                assert failed.block["lock"] == "drew ids, then failed"
            after = box.dispatch(session, "word_insert_markdown",
                                 {"doc": doc, "at": f"before:{layout.ends[1]}", "markdown": "C."})
            found.append((after.created, session.entry(doc).document.id_state()))
    assert found[0] == found[1]


# -- round 48 (R83): edits of known blocks ---------------------------------------------------------


def _round48_calls(doc, layout):
    """Each newly admitted kind, and a batch mixing them as callers send them."""
    body, other, span = layout.bodies[1], layout.bodies[2], f"{layout.heads[1]}..{layout.ends[1]}"
    word = Document.open(block_stress.made(3)).paragraph(body[0]).text.split()[1].strip(".,").lower()
    return [
        ("replace_text", {"doc": doc, "find": word, "replace": word.upper(), "expect": "all", "range": span}),
        ("word_format", {"doc": doc, "targets": [body[0]], "bold": True}),
        ("word_format", {"doc": doc, "targets": [f"{body[1]}@0:5", span], "italic": True}),
        ("word_format", {"doc": doc, "find": [{"text": word.upper(), "within": span, "expect": "all"}],
                         "underline": True}),
        ("word_format_table", {"doc": doc, "table": layout.tables[1], "scope": "table", "width_of": "body"}),
        ("word_format_table", {"doc": doc, "table": layout.tables[1], "scope": "row", "row": 0, "shading": "DDEEFF"}),
        ("word_delete", {"doc": doc, "target": f"{body[2]}@0:3"}),
        ("word_links", {"doc": doc, "action": "add_hyperlink", "target": f"{body[3]}@0:4",
                        "url": "https://example.com/a"}),
        ("word_links", {"doc": doc, "action": "add_hyperlink", "target": f"{body[3]}@6:9", "bookmark": "first_section"}),
        ("word_links", {"doc": doc, "action": "add_bookmark", "target": f"{body[0]}@1:6", "bookmark": "a_mark"}),
        ("batch", {"ops": [
            {"tool": "replace_text", "arguments": {"doc": doc, "find": word.upper(), "replace": word, "expect": "all",
                                                   "range": span}},
            {"tool": "word_format", "arguments": {"doc": doc, "targets": [body[1]], "bold": True}},
            {"tool": "word_format_table", "arguments": {"doc": doc, "table": layout.tables[1], "scope": "table",
                                                        "borders": "single"}},
            {"tool": "word_delete", "arguments": {"doc": doc, "target": other[0]}},
            {"tool": "word_links", "arguments": {"doc": doc, "action": "add_hyperlink", "target": f"{body[0]}@0:3",
                                                 "url": "https://example.com/b"}}]}),
        ("word_delete", {"doc": doc, "target": other[2]}),
    ]


def test_edits_of_known_blocks_answer_and_write_on_a_draft_as_on_the_document():
    """Round 48 (R83): replace_text in a span, word_format, word_format_table, word_delete and
    word_links (a link, a bookmark) -- alone and mixed in one batch -- each run on a draft:
    the same results, touched, bytes and id state as on the document, and none of them on
    the lock."""
    out = {}
    for mode in ("off", "always"):
        with opened(mode) as (box, session, doc, layout):
            results = []
            for name, arguments in _round48_calls(doc, layout):
                found = box.dispatch(session, name, arguments)
                assert found.ok, (name, found.to_json())
                results.append((block_stress.normal(found), found.block))
            entry = session.entry(doc)
            out[mode] = ([r for r, _ in results], [b for _, b in results],
                         hashlib.sha256(entry.document.to_bytes()).hexdigest(), entry.document.id_state())
    assert out["off"][0] == out["always"][0]
    assert out["off"][2:] == out["always"][2:]
    assert [(b or {}).get("lock") for b in out["always"][1]] == [None] * len(out["always"][1]), out["always"][1]


@pytest.mark.parametrize("name, arguments, why", [
    ("replace_text", {"find": "the", "replace": "a", "expect": "all"}, None),           # not admitted
    ("word_format", {"targets": ["{span}"], "paragraph_style": "Caption"}, "styles.xml"),  # a style made
    ("word_format", {"targets": ["h:{head}"], "bold": True}, None),                      # not admitted
])
def test_what_reads_or_changes_the_whole_document_keeps_the_lock(name, arguments, why):
    with opened("always") as (box, session, doc, layout):
        def fill(value):
            if isinstance(value, str):
                return value.format(span=f"{layout.heads[1]}..{layout.ends[1]}", head=layout.heads[1])
            return [fill(v) for v in value] if isinstance(value, list) else value

        filled = {key: fill(value) for key, value in arguments.items()}
        found = box.dispatch(session, name, {"doc": doc, **filled})
        if why is None:
            assert found.block is None
        else:
            assert why in found.block["lock"], found.block


def _held_and_serial(monkeypatch, first, then):
    """``first`` held on its draft while ``then`` runs; then the two run one at a time in the
    order they committed (``then`` first) on a fresh session: both results and whether the
    bytes and id state are the same."""
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        name, arguments = first(doc, layout)
        thread, slow = _in_thread(lambda: box.dispatch(session, name, arguments))
        assert hold.ready.wait(60)
        other = box.dispatch(session, *then(doc, layout))
        hold.release()
        thread.join()
        result = slow["result"]
        entry = session.entry(doc)
        concurrent = (entry.document.to_bytes(), entry.document.id_state(), block_stress.normal(result))
    with opened("off") as (box, session, doc, layout):
        assert box.dispatch(session, *then(doc, layout)).ok
        alone = box.dispatch(session, *first(doc, layout))
        entry = session.entry(doc)
        serial = (entry.document.to_bytes(), entry.document.id_state(), block_stress.normal(alone))
    return result, other, concurrent == serial


def test_links_added_meanwhile_merge_their_relationships(monkeypatch):
    """Two links to different addresses in different chapters: the held one's relationship
    is the next the document has free when it applies (its r:id renamed), the bytes those of
    the two calls one at a time."""
    result, other, same = _held_and_serial(
        monkeypatch,
        lambda doc, layout: ("word_links", {"doc": doc, "action": "add_hyperlink",
                                            "target": f"{layout.bodies[0][0]}@0:4", "url": "https://example.com/x"}),
        lambda doc, layout: ("word_links", {"doc": doc, "action": "add_hyperlink",
                                            "target": f"{layout.bodies[2][0]}@0:4", "url": "https://example.com/y"}))
    assert other.ok and result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
    assert same


def test_a_link_to_an_address_linked_meanwhile_shares_its_relationship(monkeypatch):
    result, other, same = _held_and_serial(
        monkeypatch,
        lambda doc, layout: ("word_links", {"doc": doc, "action": "add_hyperlink",
                                            "target": f"{layout.bodies[0][0]}@0:4", "url": "https://example.com/x"}),
        lambda doc, layout: ("word_links", {"doc": doc, "action": "add_hyperlink",
                                            "target": f"{layout.bodies[2][0]}@0:4", "url": "https://example.com/x"}))
    assert other.ok and result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
    assert same


def test_bookmarks_added_meanwhile_draw_their_ids_again(monkeypatch):
    result, other, same = _held_and_serial(
        monkeypatch,
        lambda doc, layout: ("word_links", {"doc": doc, "action": "add_bookmark",
                                            "target": f"{layout.bodies[0][0]}@0:4", "bookmark": "held_mark"}),
        lambda doc, layout: ("word_links", {"doc": doc, "action": "add_bookmark",
                                            "target": f"{layout.bodies[2][0]}@0:4", "bookmark": "other_mark"}))
    assert other.ok and result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
    assert same


def test_a_bookmark_of_the_same_name_added_meanwhile_prepares_again(monkeypatch):
    result, other, same = _held_and_serial(
        monkeypatch,
        lambda doc, layout: ("word_links", {"doc": doc, "action": "add_bookmark",
                                            "target": f"{layout.bodies[0][0]}@0:4", "bookmark": "Mark"}),
        lambda doc, layout: ("word_links", {"doc": doc, "action": "add_bookmark",
                                            "target": f"{layout.bodies[2][0]}@0:4", "bookmark": "mark"}))
    assert other.ok and result.block["tries"] == 2 and "bookmark" in result.block["conflicts"][0]
    assert not result.ok and same            # as one at a time: the name is taken


def test_text_replaced_in_a_span_a_block_was_put_into_meanwhile_prepares_again(monkeypatch):
    """A span's blocks are what a replace in it reads: one put between them meanwhile is one
    the call would have changed too."""
    result, other, same = _held_and_serial(
        monkeypatch,
        lambda doc, layout: ("replace_text", {"doc": doc, "find": "Section", "replace": "Part", "expect": "all",
                                              "range": f"{layout.heads[1]}..{layout.ends[1]}"}),
        lambda doc, layout: ("word_insert_markdown", {"doc": doc, "at": f"after:{layout.bodies[1][0]}",
                                                      "markdown": "Section inserted meanwhile."}))
    assert other.ok and result.ok and result.block["tries"] == 2
    assert same


def test_edits_of_other_chapters_apply_beside_a_held_batch_of_block_edits(monkeypatch):
    result, other, same = _held_and_serial(
        monkeypatch,
        lambda doc, layout: ("batch", {"ops": [
            {"tool": "word_format", "arguments": {"doc": doc, "targets": [layout.bodies[0][0]], "bold": True}},
            {"tool": "word_format_table", "arguments": {"doc": doc, "table": layout.tables[0], "scope": "table",
                                                        "shading": "DDEEFF"}},
            {"tool": "word_delete", "arguments": {"doc": doc, "target": f"{layout.bodies[0][1]}@0:3"}}]}),
        lambda doc, layout: ("batch", {"ops": [
            {"tool": "replace_text", "arguments": {"doc": doc, "find": "Section", "replace": "Part", "expect": "all",
                                                   "range": f"{layout.heads[2]}..{layout.ends[2]}"}},
            {"tool": "word_format", "arguments": {"doc": doc, "targets": [layout.bodies[2][0]], "italic": True}}]}))
    assert other.ok and result.ok and _shape(result.block) == {"tries": 1, "conflicts": []}
    assert same


@pytest.mark.parametrize("first_fill", [False, True])
def test_a_dataset_table_put_alone_after_edits_side_by_side_is_prepared_once(first_fill):
    """Round 48 (R85, the team's case): after callers edited side by side, a table from a
    dataset put in alone (`from: "d:..."`, nothing else in flight) is prepared once -- 0.1.74
    looked for its own table's id in the document ("t:... is gone", three times)."""
    with opened("auto", first_fill=first_fill) as (box, session, doc, layout):
        def edit(k):
            view = session.view(context={"agent": f"c{k}"})
            for c in range(3):
                assert box.dispatch(view, "batch", {"ops": [
                    {"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": f"before:{layout.ends[k]}",
                                                                   "markdown": f"Note {k}.{c}"}},
                    {"tool": "word_set_text", "arguments": {"doc": doc, "target": layout.bodies[k][c],
                                                            "text": f"Text {k}.{c}."}}]}).ok

        threads = [threading.Thread(target=edit, args=(k,)) for k in range(3)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()
        result = box.dispatch(session.view(context={"agent": "alone"}), "word_insert_markdown",
                              {"doc": doc, "at": f"before:{layout.ends[1]}", "from": "d:rows"})
        assert result.ok and result.created[0].startswith("t:")
        assert _shape(result.block) == {"tries": 1, "conflicts": []}
