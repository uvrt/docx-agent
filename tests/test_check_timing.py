"""A check says where its time went (round 47, R82), as an edit does.

``check``'s ``data.timing`` (:class:`docx_agent.tools._base.Timing`, seconds each, to the
millisecond): the parts that ran -- ``resolve`` (the ranges), ``validate``, ``copies``,
``layout``, ``facts``, ``fields`` -- and the waits apart: ``worker_wait`` (for the worker
holding the document's layout) and ``document_wait`` (taking the document back after laying
out with it let go).  Nothing else in the result changes, and it asks for nothing new.

Tool design rules: 4 (a handful of numbers, the same few names whatever the document's size;
the parts that did not run are left out, not listed as nought).  Made-up documents only."""

from __future__ import annotations

import contextlib
import time

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from docx_agent.tools import _base as base_module  # noqa: E402
from docx_agent.tools import shared as shared_module  # noqa: E402
from docx_agent.tools._base import Timing  # noqa: E402

pytestmark = pytest.mark.any_faces

WORDS = "lantern meadow river stone harbour quiet orchard copper valley signal".split()

#: A sleep measured by another clock, rounded to the millisecond: Windows' can read a little
#: short (a 0.2 s sleep as 0.199).
CLOCK = 0.01
#: Every name a check's timing may hold.
PARTS = {"resolve", "validate", "copies", "layout", "facts", "fields", "worker_wait", "document_wait"}


def _text(seed: int, words: int) -> str:
    return " ".join(WORDS[(seed + k) % len(WORDS)] for k in range(words)).capitalize() + "."


@pytest.fixture(scope="module")
def made() -> bytes:
    doc = Document.new()
    parts = ["# Report title", "An opening paragraph."]
    for c in range(1, 4):
        parts += [f"# Chapter {c}", _text(c, 90), _text(c + 1, 70), "| A | B |\n|---|---|\n| 1 | 2 |"]
    doc.insert_markdown("\n\n".join(parts))
    return doc.to_bytes()


def _open(box, data):
    session = box.session()
    doc = session.open(data, "report.docx")
    return session, doc, session.entry(doc).document


def _timing(result) -> dict:
    assert result.ok, result.to_json()
    timing = result.data["timing"]
    assert set(timing) <= PARTS, timing
    assert all(isinstance(v, float) and v >= 0 for v in timing.values()), timing
    # Its own time and the waits inside the call: never more than the call took.
    assert sum(timing.values()) <= result.duration + 0.01, (timing, result.duration)
    return timing


@pytest.mark.parametrize("workers", [0, 2])
def test_a_lone_check_says_its_parts_and_its_waits(made, workers):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=workers) as box:
        session, doc, _document = _open(box, made)
        result = box.dispatch(session, "check", {"doc": doc})
        timing = _timing(result)
        # The default include: validate, reflow (laid out in the worker holding the
        # layout), fields; the document let go while laying out, taken back after.
        assert list(timing) == ["validate", "layout", "facts", "worker_wait", "document_wait", "fields"]
        assert timing["layout"] > 0
        # The timing is the data's last key; the rest is what it was.
        assert list(result.data) == ["validate", "reflow", "fields", "timing"]


def test_ranges_and_copies_are_timed_when_given(made):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, document = _open(box, made)
        ids = [b["id"] for b in document.blocks()]
        result = box.dispatch(session, "check", {"doc": doc, "range": [f"{ids[2]}..{ids[4]}"],
                                                 "include": ["validate"]})
        timing = _timing(result)
        assert {"resolve", "validate", "layout", "facts"} <= set(timing)
        assert "fields" not in timing and "copies" not in timing      # did not run: not listed
        copied = box.dispatch(session, "check", {"doc": doc, "include": ["validate"], "copies": [
            {"blocks": f"{ids[6]}..{ids[9]}", "reference": {"block": f"{ids[2]}..{ids[5]}"}}]})
        assert "copies" in _timing(copied)


def test_a_check_answered_from_what_the_version_had_lays_nothing_out(made):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, _document = _open(box, made)
        first = _timing(box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]}))
        again = _timing(box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]}))
        assert "layout" in first and "worker_wait" in first
        # The same check of the same version: its facts kept, no worker asked.
        assert "layout" not in again and "worker_wait" not in again and "facts" in again


def test_the_wait_for_the_worker_is_apart_from_the_layout(made, monkeypatch):
    """Time before the worker's task starts -- a queue, another layout of the document in
    progress there -- is ``worker_wait``, not ``layout``; the layout's own time is ``layout``,
    not ``worker_wait``.  Both are stretched by sleeps far apart (the layout's much the
    longer), so a busy machine's slack cannot pass one off as the other."""
    import docx2svg

    stall, laying = 0.3, 2.0
    check = base_module.Engine.check
    lay_out = docx2svg.lay_out_docx

    def slow(self, *args, **kwargs):
        time.sleep(stall)
        return check(self, *args, **kwargs)

    def slow_layout(*args, **kwargs):
        time.sleep(laying)
        return lay_out(*args, **kwargs)

    monkeypatch.setattr(base_module.Engine, "check", slow)
    monkeypatch.setattr(docx2svg, "lay_out_docx", slow_layout)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, _document = _open(box, made)
        timing = _timing(box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]}))
    assert timing["layout"] >= laying - CLOCK                    # the layout's time is its own
    assert timing["worker_wait"] >= stall - CLOCK                # the wait before it is the wait
    # Had the layout been counted as waiting, the wait would hold it too: ``laying`` more.
    assert timing["worker_wait"] < stall + laying, timing


def test_taking_the_document_back_is_document_wait():
    """The time ``unlocked()`` takes to take the document back is ``document_wait``; nothing
    when it was not let go (inside a batch)."""
    stall = 0.2

    class Call:
        def __init__(self, let_go):
            self.let_go = let_go

        @contextlib.contextmanager
        def unlocked(self):
            yield self.let_go
            if self.let_go:
                time.sleep(stall)           # another call holds it

    timing = Timing()
    with shared_module._timed_unlocked(Call(True), timing)() as let_go:
        assert let_go is True
    assert timing.parts["document_wait"] >= stall - CLOCK
    held = Timing()
    with shared_module._timed_unlocked(Call(False), held)():
        pass
    assert held.parts == {}


def test_a_check_in_a_batch_lays_out_there(made):
    """Between a batch's changes the document is held: the layout is timed where it is made,
    and nothing is let go."""
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, document = _open(box, made)
        target = document.paragraphs()[3].id
        result = box.dispatch(session, "batch", {"ops": [
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": target, "text": "Changed."}},
            {"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}},
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": target, "text": "Again."}}]})
        assert result.ok, result.to_json()
        timing = result.data["ops"][1]["data"]["timing"]
        assert "layout" in timing and "document_wait" not in timing and "worker_wait" not in timing


def test_a_check_in_a_read_batch_has_its_timing(made):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, _document = _open(box, made)
        result = box.dispatch(session, "read_batch", {"ops": [
            {"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}}]})
        assert result.ok, result.to_json()
        timing = result.data["ops"][0]["data"]["timing"]
        assert set(timing) <= PARTS and "layout" in timing
