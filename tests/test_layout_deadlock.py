"""No wait for a layout in progress while holding the document (round 34, R17).

A lone ``check`` lays its document out with the lock let go (the layout's *flight*: other
callers asking for the same state wait for it).  A call that holds the lock -- the reads of a
``read_batch``, a ``batch``'s ops, ``word_inspect`` and other reads that lay out under the
lock -- never waits for that flight: it lays out on its own; and a flight's owner finishes
its flight before it takes the lock back.  Before, the owner waited for the lock held by a
caller waiting for its flight: every call on the document hung.

Every test here runs its calls in daemon threads joined with a deadline: a deadlock fails
the test, never hangs the run.  Made-up documents only; the layout engine is slowed down so
the calls meet while the first layout runs."""

from __future__ import annotations

import threading
import time

import pytest

from docx_agent.tools import _base

pytestmark = pytest.mark.any_faces

#: How long the slowed layout engine stalls before each layout.
STALL = 0.4
#: How long the calls of one test may take in all before it calls them deadlocked.
DEADLINE = 90.0


def report(chapters: int = 3) -> bytes:
    from docx_agent import Document

    words = "lantern meadow river stone harbour quiet orchard copper valley signal".split()
    doc = Document.new()
    parts = ["# Report title", "An opening paragraph."]
    for c in range(1, chapters + 1):
        parts.append(f"# Chapter {c}")
        for p in range(3):
            parts.append(" ".join(words[(c + p + k) % len(words)] for k in range(90)).capitalize() + ".")
    doc.insert_markdown("\n\n".join(parts))
    for paragraph in doc.paragraphs():
        if paragraph.text.startswith("Chapter "):
            paragraph.format(page_break_before=True)
    doc.insert_toc(after=doc.paragraphs()[1].id, levels=(1, 1))
    section = doc.sections()[0]
    doc.add_footer(section.id, "default", "Page ")
    doc.insert_field(f"{section.footer().paragraphs[0].id}@5", "PAGE", result="1")
    return doc.to_bytes()


@pytest.fixture(scope="module")
def made() -> bytes:
    return report()


def _toolbox(**options):
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, **options)


def _slowed(monkeypatch) -> threading.Event:
    """``Engine.lay_out`` and ``Engine.check`` (a lone check's facts, made where the layout
    is held: round 38) stalled by :data:`STALL`; the event is set when the first began."""
    started = threading.Event()
    run = _base.Engine.lay_out
    facts = _base.Engine.check

    def slow(engine, data, options, state=None):
        started.set()
        time.sleep(STALL)
        return run(engine, data, options, state)

    def slow_facts(engine, data, options, state, request):
        started.set()
        time.sleep(STALL)
        return facts(engine, data, options, state, request)

    monkeypatch.setattr(_base.Engine, "lay_out", slow)
    monkeypatch.setattr(_base.Engine, "check", slow_facts)
    return started


def _all_finish(box, session, first, then, *, lone: int = 2) -> dict:
    """``lone`` lone calls ``first`` (name, arguments) at once; once a layout began, the
    calls ``then`` at once.  Every result, by name; fails -- never hangs -- if any call is
    still running at the deadline."""
    out: dict[str, list] = {}
    lock = threading.Lock()

    def call(name, arguments):
        result = box.dispatch(session, name, arguments)
        with lock:
            out.setdefault(name, []).append(result)

    threads = [threading.Thread(target=call, args=first, daemon=True) for _ in range(lone)]
    for thread in threads:
        thread.start()
    assert _STARTED[0].wait(DEADLINE)
    later = [threading.Thread(target=call, args=c, daemon=True) for c in then]
    for thread in later:
        thread.start()
    threads += later
    ends = time.monotonic() + DEADLINE
    for thread in threads:
        thread.join(max(0.0, ends - time.monotonic()))
    hung = [t for t in threads if t.is_alive()]
    if hung:
        pytest.fail(f"deadlock: {len(hung)} of {len(threads)} calls still waiting after {DEADLINE:g} s")
    return out


_STARTED: list[threading.Event] = []


@pytest.fixture
def started(monkeypatch):
    _STARTED[:] = [_slowed(monkeypatch)]
    yield _STARTED[0]


def _ok(results) -> None:
    bad = [r.to_json() for rs in results.values() for r in rs if not r.ok]
    assert not bad, bad


def test_lone_checks_and_a_read_batch_of_check_and_inspect_all_finish(made, started):
    """The production hang: lone checks lay the version out unlocked while a read_batch of
    check and word_inspect, holding the document, asks for the same layout."""
    box = _toolbox(workers=0)
    session = box.session()
    doc = session.open(made, "report.docx")
    target = session.entry(doc).document.paragraphs()[2].id
    check = {"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}}
    inspect = {"tool": "word_inspect", "arguments": {"doc": doc, "target": target, "layout": True}}
    out = _all_finish(box, session, ("check", {"doc": doc, "include": ["reflow"]}),
                      [("read_batch", {"ops": [check, inspect]}), ("read_batch", {"ops": [inspect, check]})])
    box.close()
    _ok(out)
    pages = {r.data["reflow"]["pages"] for r in out["check"]}
    for batch in out["read_batch"]:
        answers = batch.data["ops"]
        assert not any("error" in answer for answer in answers), batch.to_json()
        pages |= {a["data"]["reflow"]["pages"] for a in answers if a["tool"] == "check"}
    assert len(pages) == 1


def test_lone_checks_and_a_lone_inspect_all_finish(made, started):
    """word_inspect alone lays out under the lock too (its state's placements)."""
    box = _toolbox(workers=0)
    session = box.session()
    doc = session.open(made, "report.docx")
    target = session.entry(doc).document.paragraphs()[2].id
    out = _all_finish(box, session, ("check", {"doc": doc, "include": ["reflow"]}),
                      [("word_inspect", {"doc": doc, "target": target, "layout": True})] * 2)
    box.close()
    _ok(out)


def test_lone_checks_and_a_batch_with_a_check_between_changes_all_finish(made, started):
    """A batch's check before and between its changes lays out holding the document."""
    box = _toolbox(workers=0)
    session = box.session()
    doc = session.open(made, "report.docx")
    body = [p.id for p in session.entry(doc).document.paragraphs() if len(p.text) > 200]
    batch = {"ops": [{"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}},
                     {"tool": "word_set_text", "arguments": {"doc": doc, "target": body[0], "text": "One."}},
                     {"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}},
                     {"tool": "word_set_text", "arguments": {"doc": doc, "target": body[1], "text": "Two."}}]}
    out = _all_finish(box, session, ("check", {"doc": doc, "include": ["reflow"]}), [("batch", batch)])
    box.close()
    _ok(out)


def test_lone_checks_and_a_fields_update_all_finish(made, started):
    """word_fields update: its work ahead (on a copy, unlocked), then the update under the
    lock, beside lone checks of the same version."""
    box = _toolbox(workers=0)
    session = box.session()
    doc = session.open(made, "report.docx")
    out = _all_finish(box, session, ("check", {"doc": doc, "include": ["reflow"]}),
                      [("word_fields", {"doc": doc, "action": "update"}),
                       ("read_batch", {"ops": [{"tool": "check", "arguments": {"doc": doc}}]})])
    box.close()
    _ok(out)


def test_many_rounds_of_checks_reads_and_edits_all_finish(made, started):
    """Nine callers at once, as a report's builders: lone checks, read_batches, inspects
    and edits, a few rounds."""
    box = _toolbox(workers=0)
    session = box.session()
    doc = session.open(made, "report.docx")
    body = [p.id for p in session.entry(doc).document.paragraphs() if len(p.text) > 200]
    target = body[0]
    reads = {"ops": [{"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}},
                     {"tool": "word_inspect", "arguments": {"doc": doc, "target": target, "layout": True}}]}
    for k in range(2):
        out = _all_finish(box, session, ("check", {"doc": doc, "include": ["reflow"]}), [
            ("read_batch", reads), ("word_inspect", {"doc": doc, "target": target, "layout": True}),
            ("word_set_text", {"doc": doc, "target": body[1 + k], "text": f"Round {k}."}),
            ("check", {"doc": doc}), ("read_batch", reads),
            ("batch", {"ops": [{"tool": "check", "arguments": {"doc": doc}},
                               {"tool": "word_set_text",
                                "arguments": {"doc": doc, "target": body[3 + k], "text": "Shorter."}}]}),
        ], lone=3)
        _ok(out)
    box.close()
