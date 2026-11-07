"""No layout under the lock (round 33, R16c; ooxml-edit 0.79.0).

* ``check`` after a batch's last change lays the document out once the batch is done, the
  document's lock let go: a read meanwhile does not wait for it.  Between a batch's changes
  it still holds the document, and says so with the way out.
* ``word_fields`` ``update`` is done *ahead* of its change on a copy, the lock let go, and
  the copy's layouts kept: the update under the lock finds every pass's layout cached.  A
  page cross-reference and the upgrade lay the document out ahead.  Where another call
  changed the document meanwhile, the work ahead is done again from there.  Same results as
  before.

Made-up documents only; the layout engine is slowed down where a test times a read
against it (``Engine.lay_out``)."""

from __future__ import annotations

import threading
import time

import pytest

from docx_agent import Document
from docx_agent.tools import _base
from docx_agent.tools.shared import BATCH_CHECK

pytestmark = pytest.mark.any_faces

#: How long the slowed layout engine stalls before each layout.
STALL = 1.0


def report(chapters: int = 3) -> bytes:
    """A made-up report: a contents, chapters on new pages, a page number in the footer."""
    words = "lantern meadow river stone harbour quiet orchard copper valley signal".split()
    doc = Document.new()
    parts = ["# Report title", "An opening paragraph."]
    for c in range(1, chapters + 1):
        parts.append(f"# Chapter {c}")
        for p in range(4):
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


class Slowed:
    """``Engine.lay_out`` stalled by :data:`STALL`: when each layout began and ended, and
    whether the document's lock was free while it ran."""

    def __init__(self, monkeypatch, entry=None) -> None:
        self.began: list[float] = []
        self.ended: list[float] = []
        self.free: list[bool] = []
        self.started = threading.Event()
        self.entry = entry
        run = _base.Engine.lay_out

        def slow(engine, data, options, state=None):
            self.began.append(time.perf_counter())
            if self.entry is not None:
                self.free.append(_free(self.entry.lock))
            self.started.set()
            time.sleep(STALL)
            try:
                return run(engine, data, options, state)
            finally:
                self.ended.append(time.perf_counter())

        facts = _base.Engine.check

        def slow_facts(engine, data, options, state, request):     # a lone check's (round 38)
            self.began.append(time.perf_counter())
            if self.entry is not None:
                self.free.append(_free(self.entry.lock))
            self.started.set()
            time.sleep(STALL)
            try:
                return facts(engine, data, options, state, request)
            finally:
                self.ended.append(time.perf_counter())

        monkeypatch.setattr(_base.Engine, "lay_out", slow)
        monkeypatch.setattr(_base.Engine, "check", slow_facts)


def _free(lock) -> bool:
    """Whether another thread can take ``lock`` now."""
    out = {}

    def attempt():
        out["got"] = lock.acquire(timeout=0.05)
        if out["got"]:
            lock.release()

    thread = threading.Thread(target=attempt)
    thread.start()
    thread.join()
    return out["got"]


def _behind(box, session, doc, slowed, slow_call):
    """Run ``slow_call`` in a thread; once its first layout began, read a paragraph.  The
    read's result, when it came back, and the slow call's result."""
    out = {}

    def run():
        out["slow"] = slow_call()

    thread = threading.Thread(target=run)
    thread.start()
    assert slowed.started.wait(60)
    first = session.entry(doc).document.paragraphs()[0].id
    out["read"] = box.dispatch(session, "word_read", {"doc": doc, "range": first})
    out["read_at"] = time.perf_counter()
    thread.join()
    return out


def test_a_check_after_the_batch_s_last_change_lays_out_unlocked(made, monkeypatch):
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        document = session.entry(doc).document
        target = [p for p in document.paragraphs() if len(p.text) > 200][2].id
        # (No layout before: a document of 20 pages or fewer gets its reflow in every edit's
        # checks, laid out with the change -- from the layout before it, when there is one.)
        slowed = Slowed(monkeypatch, session.entry(doc))
        out = _behind(box, session, doc, slowed, lambda: box.dispatch(session, "batch", {"ops": [
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": target, "text": "Shorter now."}},
            {"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}}]}))
    result = out["slow"]
    assert result.ok, result.to_json()
    # The document's lock was free while the check laid out (a read of the new version
    # meanwhile then costs only its own work).
    assert out["read"].ok and slowed.free == [True]
    check = result.data["ops"][1]
    assert check["tool"] == "check" and check["data"]["reflow"]["pages"] >= 4
    assert not any(BATCH_CHECK in w for w in result.warnings)


def test_a_check_between_a_batch_s_changes_holds_the_document_and_says_so(made, monkeypatch):
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        targets = [p.id for p in session.entry(doc).document.paragraphs() if len(p.text) > 200]
        slowed = Slowed(monkeypatch, session.entry(doc))
        result = box.dispatch(session, "batch", {"ops": [
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": targets[0], "text": "One."}},
            {"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}},
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": targets[1], "text": "Two."}}]})
    assert result.ok, result.to_json()
    assert slowed.free and not any(slowed.free)
    assert f"ops[1]: {BATCH_CHECK}" in result.warnings


@pytest.mark.parametrize("atomic", [False, True])
def test_a_check_after_an_atomic_or_partial_batch_answers_of_its_version(made, atomic):
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        target = [p for p in session.entry(doc).document.paragraphs() if len(p.text) > 200][0].id
        before = session.entry(doc).version
        result = box.dispatch(session, "batch", {"atomic": atomic, "ops": [
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": target, "text": "Short."}},
            {"tool": "check", "arguments": {"doc": doc, "include": ["reflow", "fields"]}}]})
        alone = box.dispatch(session, "check", {"doc": doc, "include": ["reflow", "fields"]})
    assert result.ok and result.version == before + 1
    check = result.data["ops"][1]["data"]
    assert check["reflow"]["pages"] == alone.data["reflow"]["pages"]
    assert check["fields"] == alone.data["fields"]


def test_word_fields_update_lays_out_ahead_unlocked(made, monkeypatch):
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        entry = session.entry(doc)
        target = [p for p in entry.document.paragraphs() if len(p.text) > 200][1].id
        assert box.dispatch(session, "word_set_text", {"doc": doc, "target": target,
                                                       "text": "A much shorter paragraph."}).ok
        slowed = Slowed(monkeypatch, entry)
        out = _behind(box, session, doc, slowed,
                      lambda: box.dispatch(session, "word_fields", {"doc": doc, "action": "update"}))
    update = out["slow"]
    assert update.ok, update.to_json()
    assert out["read"].ok and out["read_at"] < slowed.ended[0]     # read while it laid out
    # Every pass laid out ahead, on a copy, unlocked; the update under the lock found each
    # pass's layout of the bytes it wrote, and laid nothing out.
    assert slowed.free and all(slowed.free), slowed.free


def test_word_fields_update_writes_what_it_wrote_before(made):
    """The same results with and without the layout ahead (a Document's own update)."""
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        entry = session.entry(doc)
        target = [p for p in entry.document.paragraphs() if len(p.text) > 200][1].id
        assert box.dispatch(session, "word_set_text", {"doc": doc, "target": target,
                                                       "text": "A much shorter paragraph."}).ok
        expected = Document.open(entry.document.to_bytes())
        expected.update_fields()
        update = box.dispatch(session, "word_fields", {"doc": doc, "action": "update"})
        assert update.ok, update.to_json()
        assert entry.document.to_markdown() == expected.to_markdown()
        assert [(f.instruction, f.result) for f in entry.document.fields()] == \
               [(f.instruction, f.result) for f in expected.fields()]


LONG = "Edited while the fields laid out, long enough to push the chapters a page on. " * 40


def test_an_edit_while_fields_lay_out_ahead_goes_first_and_the_update_sees_it(made, monkeypatch):
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        entry = session.entry(doc)
        body = [p for p in entry.document.paragraphs() if len(p.text) > 200]
        before = entry.version
        slowed = Slowed(monkeypatch, entry)
        out = {}

        def update():
            out["update"] = box.dispatch(session, "word_fields", {"doc": doc, "action": "update"})

        thread = threading.Thread(target=update)
        thread.start()
        assert slowed.started.wait(60)
        edit = box.dispatch(session, "word_set_text", {"doc": doc, "target": body[0].id,
                                                       "text": LONG})
        edited_at = time.perf_counter()
        thread.join()
        expected = Document.open(made)
        expected.paragraph(body[0].id).set_text(LONG)
        expected.update_fields()
    assert edit.ok and edited_at < slowed.ended[0]
    assert out["update"].ok and out["update"].version == before + 2
    # Done ahead twice (the document moved on during the first), all of it unlocked.
    assert len(slowed.free) >= 2 and all(slowed.free), slowed.free
    assert [(f.instruction, f.result) for f in entry.document.fields()] == \
           [(f.instruction, f.result) for f in expected.fields()]


def test_fields_that_need_no_layout_lay_nothing_out_ahead(monkeypatch):
    doc = Document.new()
    doc.insert_markdown("Plain text.")
    doc.insert_date(f"{doc.paragraphs()[0].id}@5")
    with _toolbox(workers=0) as box:
        session = box.session()
        handle = session.open(doc.to_bytes(), "plain.docx")
        slowed = Slowed(monkeypatch)
        update = box.dispatch(session, "word_fields", {"doc": handle, "action": "update"})
    assert update.ok and slowed.began == []


def test_concurrent_field_updates_and_edits_stay_consistent(made):
    """Free-threading safe: updates laying out ahead side by side with edits, each one undo
    step and one version, every result as a lone update of the final text gives."""
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        entry = session.entry(doc)
        body = [p.id for p in entry.document.paragraphs() if len(entry.document.get(p.id).text) > 200]
        before = entry.version
        results: list = []

        def work(k: int) -> None:
            if k % 2:
                results.append(box.dispatch(session, "word_fields", {"doc": doc, "action": "update"}))
            else:
                results.append(box.dispatch(session, "word_set_text", {
                    "doc": doc, "target": body[k], "text": f"Edit {k}. " * (5 + 10 * k)}))

        threads = [threading.Thread(target=work, args=(k,)) for k in range(6)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()
        assert all(r.ok for r in results), [r.to_json() for r in results if not r.ok]
        changed = sum(1 for r in results if r.tool == "word_set_text" or r.data["changed_results"])
        assert entry.version <= before + 6 and entry.version >= before + changed - 3
        final = box.dispatch(session, "word_fields", {"doc": doc, "action": "update"})
        assert final.ok
        expected = Document.open(entry.document.to_bytes())
        expected.update_fields()
        assert [(f.instruction, f.result) for f in entry.document.fields()] == \
               [(f.instruction, f.result) for f in expected.fields()]


def test_a_page_cross_reference_lays_out_ahead_unlocked(made, monkeypatch):
    document = Document.open(made)
    chapter = next(p for p in document.paragraphs() if p.text == "Chapter 2")
    document.add_bookmark(document.range(chapter.id), "ChapterTwo")
    opening = next(p for p in document.paragraphs() if p.text == "An opening paragraph.")
    expected = Document.open(document.to_bytes())
    expected.insert_cross_reference(f"{opening.id}@{len(opening.text)}", "ChapterTwo", kind="page")
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(document.to_bytes(), "report.docx")
        slowed = Slowed(monkeypatch, session.entry(doc))
        result = box.dispatch(session, "word_fields", {
            "doc": doc, "action": "insert_cross_reference", "at": f"{opening.id}@{len(opening.text)}",
            "bookmark": "ChapterTwo", "kind": "page"})
        assert result.ok, result.to_json()
        assert slowed.free[0] is True
        assert session.entry(doc).document.paragraph(opening.id).text == expected.paragraph(opening.id).text


def test_the_upgrade_lays_out_ahead_unlocked(monkeypatch):
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).parent.parent / "tools"))
    import e6_probe

    data = e6_probe.legacy_document()
    expected = Document.open(data).upgrade_to_modern()
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(data, "legacy.docx")
        slowed = Slowed(monkeypatch, session.entry(doc))
        result = box.dispatch(session, "word_template", {"doc": doc, "action": "upgrade_to_modern"})
    assert result.ok, result.to_json()
    assert slowed.free[0] is True and not any(slowed.free[1:])
    assert result.data["mode"] == 15
    assert result.data["reflow"]["changed"] == list(expected.reflow.changed)
