"""Reads of the committed version (round 45, R64 stage 1).

Reading calls -- ``word_read``, ``describe``, ``check``, ``render``, ``word_inspect`` -- work on a
read copy of the document's last committed version (ooxml-edit's ``Session.read_lock``, the
docx format's ``read_copy``: :func:`docx_agent.layout.read_copy`) and never wait for an edit in
flight.  The copy answers exactly as the document did at that version: the same ids, version,
refusals and facts; its layouts are the document's (the next version is laid out from them,
and an edit's reflow finds the layout a check made).  Made-up documents only."""

from __future__ import annotations

import dataclasses
import json
import threading
import time

import pytest

from ooxml_edit.tools import Result, Toolbox, integer, string

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.tools._base import word_tool

pytestmark = pytest.mark.any_faces

#: How long the slow edit holds the document.
STALL = 3


@word_tool("word_test_slow_edit", "Test only: set a paragraph's text, holding the document a while.",
           {"doc": string("Document id."), "target": string("Paragraph."), "text": string("Text."),
            "stall": integer("Seconds.", minimum=0)}, group="text", mutates=True)
def slow_edit(call, doc, target, text, stall):
    call.document.paragraph(target).set_text(text)
    time.sleep(stall)
    return Result(summary="set", changed=[target])


def made() -> bytes:
    words = "lantern meadow river stone harbour quiet orchard copper valley signal".split()
    doc = Document.new()
    parts = ["# Report title", "An opening paragraph."]
    for c in range(1, 4):
        parts.append(f"# Chapter {c}")
        for p in range(3):
            parts.append(" ".join(words[(c + p + k) % len(words)] for k in range(60)).capitalize() + ".")
        parts.append(f"- point {c}a\n- point {c}b")
        parts.append("| A | B |\n|---|---|\n| 1 | 2 |")
    doc.insert_markdown("\n\n".join(parts))
    return doc.to_bytes()


def _box(copies: bool = True) -> Toolbox:
    fmt = FORMAT if copies else dataclasses.replace(FORMAT, read_copy=None)
    return Toolbox([*TOOLS, slow_edit], formats=[fmt], groups=GROUPS, workers=0)


def _scrub(result) -> dict:
    out = result.to_json()
    out.pop("duration", None)
    return out


#: The reading calls compared, by what they ask of the first chapter (ids filled in).
def _reads(ids):
    return [("word_read", {}), ("word_read", {"range": f"{ids[2]}..{ids[5]}"}), ("describe", {}),
            ("check", {}), ("check", {"range": [f"{ids[2]}..{ids[5]}"]}),
            ("word_inspect", {"target": ids[3]}), ("find_text", {"text": "meadow"})]


def _session(box):
    session = box.session()
    doc = session.open(made(), "report.docx")
    ids = [p.id for p in session.entry(doc).document.paragraphs()]
    return session, doc, ids


def _run(copies: bool) -> list[dict]:
    with _box(copies) as box:
        session, doc, ids = _session(box)
        out = []
        steps = [("word_set_text", {"target": ids[3], "text": "A new sentence here."}),
                 *_reads(ids),
                 ("word_insert_markdown", {"at": f"after:{ids[4]}", "markdown": "Added.\n\n- one\n- two"}),
                 *_reads(ids),
                 ("word_delete", {"target": ids[4]}),
                 ("word_read", {"range": ids[4]}),             # a deleted id: refused alike
                 ("word_set_text", {"target": ids[6], "text": "Another one."}),
                 *_reads(ids)]
        for tool, arguments in steps:
            out.append(_scrub(box.dispatch(session, tool, {"doc": doc, **arguments})))
        out.append({"bytes": session.entry(doc).document.to_bytes().hex()[:0] or
                    len(session.entry(doc).document.to_bytes())})
        return out


def test_reads_of_the_committed_version_answer_as_the_document_does():
    copied, locked = _run(True), _run(False)
    assert len(copied) == len(locked)
    for one, other in zip(copied, locked):
        assert json.dumps(one, sort_keys=True, default=str) == json.dumps(other, sort_keys=True, default=str)


def test_a_check_does_not_wait_for_an_edit_in_flight():
    with _box() as box:
        session, doc, ids = _session(box)
        assert box.dispatch(session, "check", {"doc": doc}).ok
        before = session.entry(doc).version
        out: dict = {}
        thread = threading.Thread(target=lambda: out.setdefault("edit", box.dispatch(
            session, "word_test_slow_edit", {"doc": doc, "target": ids[3], "text": "Slow.", "stall": STALL})))
        thread.start()
        time.sleep(0.5)
        started = time.perf_counter()
        check = box.dispatch(session, "check", {"doc": doc, "range": [ids[6]]})
        read = box.dispatch(session, "word_read", {"doc": doc, "range": ids[3]})
        took = time.perf_counter() - started
        thread.join()
        assert check.ok and read.ok, (check.to_json(), read.to_json())
        assert took < STALL - 1, f"the reads waited {took:.1f} s"
        assert check.version == before and read.version == before
        assert "Slow." not in read.data
        assert out["edit"].ok and out["edit"].version == before + 1
        again = box.dispatch(session, "word_read", {"doc": doc, "range": ids[3]})
        assert "Slow." in again.data and again.version == before + 1


def test_the_copy_is_the_versions_bytes_ids_and_layouts():
    with _box() as box:
        session, doc, ids = _session(box)
        box.dispatch(session, "word_set_text", {"doc": doc, "target": ids[3], "text": "Changed."})
        assert box.dispatch(session, "check", {"doc": doc}).ok
        entry = session.entry(doc)
        reader = entry.reader()
        assert reader is not None and reader.version == entry.version
        copy, live = reader.document, entry.document
        assert copy is not live and copy.to_bytes() == live.to_bytes()
        assert [p.id for p in copy.paragraphs()] == [p.id for p in live.paragraphs()]
        assert copy.id_state() == live.id_state()
        # Its layouts are the document's: the check's layout is there for the next edit.
        assert copy._layouts is live._layouts and copy._layout_summaries is live._layout_summaries
        assert copy._layout_lineage == live._layout_lineage
        assert copy.font_dirs == live.font_dirs and copy.locale == live.locale


def test_an_edit_after_a_check_on_the_copy_reports_its_reflow_as_before():
    def reflow(copies):
        with _box(copies) as box:
            session, doc, ids = _session(box)
            assert box.dispatch(session, "check", {"doc": doc}).ok
            edit = box.dispatch(session, "word_set_text", {"doc": doc, "target": ids[3],
                                                           "text": "Short now."})
            return edit.checks.get("reflow") if isinstance(edit.checks, dict) else edit.checks

    first = reflow(True)
    assert first == reflow(False) and first is not None
