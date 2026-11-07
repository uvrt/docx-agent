"""Round 38: a check's facts made in the worker holding the document's layout, the layout
kept there (affinity), and the facts the same as when the layout was read here.

A lone ``check`` hands the worker the document's bytes and the token of its last layout
(:func:`docx_agent.layout.kept_layout`); the worker -- the pool's holder of the document's
lineage -- lays out from the layout it holds and sends back only the facts
(:mod:`docx_agent.tools.check_worker`).  Every test compares with the path before (the layout
pickled here and read here), result for result, byte for byte.  Made-up documents only."""

from __future__ import annotations

import json
import os
import random
import threading
import time

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent import layout as layout_module  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from docx_agent.tools import shared as shared_module  # noqa: E402

pytestmark = pytest.mark.any_faces

WORDS = "lantern meadow river stone harbour quiet orchard copper valley signal".split()


def _text(seed: int, words: int) -> str:
    return " ".join(WORDS[(seed + k) % len(WORDS)] for k in range(words)).capitalize() + "."


def report() -> bytes:
    """Four chapters of paragraphs and a table each, a contents, a footer with the page."""
    doc = Document.new()
    parts = ["# Report title", "An opening paragraph."]
    for c in range(1, 5):
        parts.append(f"# Chapter {c}")
        for p in range(5):
            parts.append(_text(c + p, 80))
        parts.append("| Item | Amount | Note |\n| --- | --- | --- |\n"
                     + "\n".join(f"| Row {r} | {r * 7} | {_text(r, 6)} |" for r in range(1, 9)))
        parts.append(_text(c, 40))
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
    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, **options)


def _body(document) -> list[str]:
    return [p.id for p in document.paragraphs() if len(p.text) > 200]


def _chapter(document, number: int) -> str:
    """A span over chapter ``number`` (its heading to the paragraph after its table)."""
    blocks = [b["id"] for b in document.blocks()]
    heads = [k for k, b in enumerate(document.blocks()) if b["level"] is not None]
    starts = heads[1:]                       # past the title
    start = starts[number - 1]
    end = (starts[number] if number < len(starts) else len(blocks)) - 1
    return f"{blocks[start]}..{blocks[end]}"


def _session_run(box, data: bytes, steps) -> list[str]:
    """Every check's data (JSON, keys in order) after each step, on a session of ``box``."""
    session = box.session()
    doc = session.open(data, "report.docx")
    document = session.entry(doc).document
    out = []
    for step in steps:
        kind, arguments = step(document)
        result = box.dispatch(session, kind, {"doc": doc, **arguments})
        assert result.ok, result.to_json()
        if kind == "check":
            out.append(json.dumps(result.data))
        else:
            out.append(json.dumps(result.checks, sort_keys=False, default=str))
    # What saving writes into app.xml: the pages and lines of the layout checked last.
    from docx_agent.edit.properties import layout_statistics

    out.append(json.dumps(layout_statistics(document)))
    return out


def _steps():
    def check(**arguments):
        return lambda document: ("check", arguments)

    def edit(index: int, words: int):
        return lambda document: ("word_set_text", {"target": _body(document)[index],
                                                   "text": _text(index + words, words)})

    def ranged(number: int, **extra):
        return lambda document: ("check", {"range": [_chapter(document, number)], **extra})

    return [check(include=["reflow"]), check(pages=[2, 3], include=["reflow", "layout"]),
            ranged(2), edit(1, 160), check(include=["reflow"]), ranged(2, include=["layout", "reflow"]),
            edit(12, 30), check(pages=[4], include=["reflow", "layout"], min_size=11),
            check(include=["layout"], styles=["Normal"]), ranged(3, pages=[1]),
            edit(3, 300), ranged(1), check(include=["validate", "reflow", "fields"]),
            check(include=["reflow"])]


def _before_path(monkeypatch) -> None:
    """The path before round 38: the layout read in this process."""
    monkeypatch.setattr(shared_module, "check_facts", lambda *a, **k: None)


@pytest.mark.parametrize("workers", [0, 2])
def test_facts_made_where_the_layout_is_held_are_the_ones_read_here(made, monkeypatch, workers):
    with _toolbox(workers=workers) as box:
        held = _session_run(box, made, _steps())
    with monkeypatch.context() as patch:
        _before_path(patch)
        with _toolbox(workers=workers) as box:
            here = _session_run(box, made, _steps())
    assert held == here


def test_a_lone_check_never_brings_the_layout_here(made, monkeypatch):
    """(An edit of a document this short gets its reflow in its checks, laid out here: only
    the checks are counted.)"""
    loads: list = []
    real = layout_module.pickle.loads
    counting = threading.Event()

    def counted(blob, *a, **k):
        if counting.is_set():
            loads.append(len(blob))
        return real(blob, *a, **k)

    monkeypatch.setattr(layout_module.pickle, "loads", counted)
    with _toolbox(workers=2) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        document = session.entry(doc).document
        assert box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]}).ok
        for k in range(3):
            assert box.dispatch(session, "word_set_text", {"doc": doc, "target": _body(document)[k],
                                                           "text": _text(k, 50)}).ok
            counting.set()
            result = box.dispatch(session, "check", {"doc": doc, "range": [_chapter(document, 2)],
                                                     "include": ["reflow", "layout"]})
            counting.clear()
            assert result.ok and result.data["reflow"]["since_last_check"] is not None
    assert loads == []


def test_a_documents_layout_tasks_go_to_the_worker_holding_its_layout(made):
    with _toolbox(workers=2) as box:
        box.warm(wait=True)
        session = box.session()
        doc = session.open(made, "report.docx")
        document = session.entry(doc).document
        key, _ = layout_module.affinity(document)
        assert box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]}).ok
        holder = box.pool.holder(key)
        assert holder is not None and document._layout_fresh is not None
        for k in range(3):
            assert box.dispatch(session, "word_set_text", {"doc": doc, "target": _body(document)[k],
                                                           "text": _text(k, 30)}).ok
            assert box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]}).ok
            assert box.pool.holder(key) == holder
        counts = box.pool.affinity_counts
        assert counts["hit"] >= 3 and counts["elsewhere"] == 0


def test_a_holder_killed_lays_the_next_version_out_afresh_with_the_same_facts(made, monkeypatch):
    import signal

    steps = _steps()[:6]
    with monkeypatch.context() as patch:
        _before_path(patch)
        with _toolbox(workers=2) as box:
            here = _session_run(box, made, steps)
    with _toolbox(workers=2) as box:
        box.warm(wait=True)
        session = box.session()
        doc = session.open(made, "report.docx")
        document = session.entry(doc).document
        key, _ = layout_module.affinity(document)
        out = []
        for number, step in enumerate(steps):
            if number == 3:                  # the holder dies between two checks
                pid = box.pool.holder(key)
                os.kill(pid, getattr(signal, "SIGKILL", signal.SIGTERM))
                # Gone before the next call (a kill is asynchronous, on Windows too): the pool
                # then finds it dead and starts another, as after any death while idle.
                ends = time.monotonic() + 30
                while time.monotonic() < ends and any(w.pid == pid and w.alive() for w in list(box.pool._idle)):
                    time.sleep(0.05)
            kind, arguments = step(document)
            result = box.dispatch(session, kind, {"doc": doc, **arguments})
            assert result.ok, result.to_json()
            out.append(json.dumps(result.data) if kind == "check"
                       else json.dumps(result.checks, default=str))
    assert out == here[:len(steps)]


def test_many_writers_checking_at_once_all_finish_with_their_versions_facts(made):
    """Eleven threads, each editing its own paragraph and checking its chapter, on two
    workers: every call ends, and each check's page count is its version's."""
    with _toolbox(workers=2) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        document = session.entry(doc).document
        targets = _body(document)
        errors: list = []

        def writer(n: int) -> None:
            rng = random.Random(n)
            try:
                for _ in range(2):
                    edit = box.dispatch(session, "word_set_text", {
                        "doc": doc, "target": targets[n % len(targets)], "text": _text(n, rng.choice([20, 200]))})
                    assert edit.ok, edit.to_json()
                    found = box.dispatch(session, "check", {"doc": doc, "range": [_chapter(document, 1 + n % 4)]})
                    assert found.ok, found.to_json()
                    assert found.data["reflow"]["pages"] >= 5
            except BaseException as exc:  # noqa: BLE001
                errors.append(exc)

        threads = [threading.Thread(target=writer, args=(n,), daemon=True) for n in range(11)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join(300)
        assert not any(thread.is_alive() for thread in threads), "deadlock"
        assert not errors, errors
        last = box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]})
        alone = Document.open(document.to_bytes()).layout()
        assert last.data["reflow"]["pages"] == alone.page_count


def test_threads_of_one_process_share_the_layouts_held(made):
    """A thread runner (free threading's): every task in this process; a task of the
    lineage whose layout another took waits for it rather than lay out whole."""
    from ooxml_edit.tools import ThreadRunner

    with _toolbox(runner=ThreadRunner(3)) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        document = session.entry(doc).document
        results: list = []

        def one(k: int) -> None:
            results.append(box.dispatch(session, "check", {"doc": doc, "range": [_chapter(document, 1 + k % 4)]}))

        threads = [threading.Thread(target=one, args=(k,)) for k in range(6)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join(300)
        assert all(r.ok for r in results) and len(results) == 6
        assert len({r.data["reflow"]["pages"] for r in results}) == 1


# -- the summary a check compares by -------------------------------------------------------------


def test_a_summary_compares_as_the_layouts_do(made):
    rng = random.Random(38)
    document = Document.open(made)
    before = document.layout()
    for round_ in range(6):
        ids = _body(document)
        for _ in range(rng.randint(1, 3)):
            document.paragraph(rng.choice(ids)).set_text(_text(round_, rng.choice([5, 60, 240])))
        after = document.layout()
        expected = after.compare(before)
        got = layout_module.LayoutSummary(after).compare(layout_module.LayoutSummary(before),
                                                         layout_module.canon_of(document))
        assert got == expected
        before = after


def test_a_summary_made_from_the_one_before_is_the_one_made_whole(made, monkeypatch):
    """Round 39: where the layout is held, a check's summary takes the digests of the pages
    and lines docx2svg kept from the layout before (the same objects) from that layout's
    summary, and digests only the rest: the summary is the one made whole, field for field,
    and pickled without the pages it held."""
    import pickle

    from docx_agent.layout import layout_key, layout_options, lineage
    from docx_agent.tools.check_worker import check_bytes

    made_whole: list[int] = []
    digest = layout_module._line_digest

    def counted(line):
        made_whole.append(1)
        return digest(line)

    monkeypatch.setattr(layout_module, "_line_digest", counted)
    rng = random.Random(39)
    document = Document.open(made)
    options = layout_options(document, {})
    token = None
    lines = []
    for round_ in range(4):
        data = document.to_bytes()
        now = layout_key(data, options)
        state = ("held", now, token, lineage(document))
        made_whole.clear()
        out = check_bytes(data, options, state, {"summary": True})
        token = now
        summary = out["summary"]
        digested = len(made_whole)
        held = layout_module._HELD[now]
        whole = layout_module.LayoutSummary(held.extra["layout"])
        assert (summary.count, summary.known, summary.rest, summary.lines) == \
            (whole.count, whole.known, whole.rest, whole.lines)
        assert "summary before" not in held.extra
        assert pickle.loads(pickle.dumps(summary))._drawn is None
        lines.append((sum(len(page) for page in summary.lines), digested))
        ids = _body(document)
        document.paragraph(ids[-1]).set_text(_text(round_, rng.choice([5, 60])))
    # The first laid out whole; then only what the edit near the end reached.
    assert lines[0][1] == lines[0][0]
    assert all(digested < total / 2 for total, digested in lines[1:]), lines


def test_a_digest_reads_a_whole_fraction_as_its_integer():
    from fractions import Fraction

    digest = layout_module._digest
    assert digest((300, Fraction(3, 2), "a")) == digest((Fraction(300, 1), Fraction(6, 4), "a"))
    assert digest((300,)) != digest((301,)) and digest((Fraction(3, 2),)) != digest((Fraction(2, 3),))
    assert digest(((1, 2),)) == digest(((Fraction(1), 2),))


def test_keys_resolved_from_their_parents_are_the_elements_the_path_names(made):
    from docx2svg.paths import resolve_path

    document = Document.open(made)
    layout = document.layout()
    root = document.package.tree(document.package.document_part())
    keys = {key for page in layout.pages for key in page.blocks}
    assert any("/w:tc[" in key for key in keys)
    for key in sorted(keys):
        assert layout._element_of(key) is resolve_path(root, key), key
