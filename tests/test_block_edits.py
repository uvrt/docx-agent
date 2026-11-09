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

from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.limits import Limits

from docx_agent import Document, blocks
from docx_agent.tools import FORMAT, GROUPS, TOOLS

import block_stress

pytestmark = pytest.mark.any_faces


def _box(mode: str = "always", tries: int = 3) -> Toolbox:
    return Toolbox([*TOOLS, block_stress.ds_rows], formats=[FORMAT], groups=GROUPS, workers=0, block_edits=mode,
                   limits=Limits(render_timeout=600, layout_timeout=600, block_tries=tries))


@contextlib.contextmanager
def opened(mode: str = "always", tries: int = 3):
    """A session on the made-up report (``block_stress.made``), its layout of ids."""
    data = block_stress.made(3)
    with _box(mode, tries) as box:
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
    ("word_format", {"targets": ["p:1A"], "bold": True}, False),
    ("word_fields", {"action": "update"}, False),
])
def test_what_a_draft_takes(tool, arguments, admitted):
    assert blocks.admits(tool, arguments) is admitted


def test_one_caller_alone_runs_as_before():
    with opened("auto") as (box, session, doc, layout):
        result = box.dispatch(session, "word_set_text", {"doc": doc, "target": layout.bodies[0][0], "text": "Alone."})
        assert result.ok and result.block is None


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


def test_an_edit_right_beside_the_paragraph_makes_the_call_prepare_again(monkeypatch):
    with opened() as (box, session, doc, layout), held(monkeypatch) as hold:
        target = layout.bodies[1][1]
        thread, slow = _in_thread(lambda: box.dispatch(session, "word_set_text",
                                                       {"doc": doc, "target": target, "text": "Held."}))
        assert hold.ready.wait(60)
        assert box.dispatch(session, "word_insert_markdown",
                            {"doc": doc, "at": f"after:{target}", "markdown": "Right after."}).ok
        hold.release()
        thread.join()
        assert slow["result"].ok and slow["result"].block["tries"] == 2


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


def test_a_value_written_by_reference_since_makes_it_prepare_again(monkeypatch):
    """The provenance item (text facts, table records) is read by every edit's bookkeeping."""
    result, other, _s, _d = _held_then(
        monkeypatch,
        lambda doc, layout: ("word_set_text", {"doc": doc, "target": layout.bodies[0][0], "text": "Held."}),
        lambda doc, layout: ("word_set_text", {"doc": doc, "target": layout.bodies[2][0],
                                               "text": "From the rows: {{d:rows/#4/name}}."}))
    assert other.ok and result.ok and result.block["tries"] == 2
    assert "customXml" in result.block["conflicts"][0]


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
        elif choice < 0.8:
            j = rng.randrange(base.count)
            copy = data[base.spans[j][0]:base.spans[j][1]]
            new = data[:begin] + copy.replace(b'paraId="', b'paraId="7', 1) + data[begin:]
        else:
            cut = rng.randrange(len(data))
            new = data[:cut] + b"<w:p>" + data[cut:]
        whole, found = blocks._read_body(new), blocks._reread_body(new, base)
        if found is None:
            continue
        reread += 1
        assert whole is not None
        assert (found.start, found.end, found.spans, found.ids) == (whole.start, whole.end, whole.spans, whole.ids)
    assert reread > 100
