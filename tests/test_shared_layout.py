"""Round 42: a check after a render of the same version -- and a render after a check --
lays nothing out: both name the document's layout by the same token
(:func:`docx_agent.layout.kept_layout`, the bytes' digest and the font folders and locale),
and the worker holding it answers from it (:func:`docx_agent.layout.held_layout`: ``"same"``,
no layout work).  A changing call lays its version out ahead (``"kept"``: from the version
before), so the render and the check after it find that one.  Where the worker holding the
layout is busy longer than the wait the pool allows (:func:`docx_agent.layout.affinity`), the
task goes to another worker, which lays the version out from what it holds: the case these
tests show apart.  Made-up documents only."""

from __future__ import annotations

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


def _report(chapters: int = 4, paragraphs: int = 5) -> bytes:
    doc = Document.new()
    parts = ["# Report title"]
    for c in range(1, chapters + 1):
        parts.append(f"# Chapter {c}")
        parts += [_text(c + p, 80) for p in range(paragraphs)]
        parts.append("| Item | Amount |\n| --- | --- |\n" + "\n".join(f"| Row {r} | {r * 7} |" for r in range(1, 6)))
    doc.insert_markdown("\n\n".join(parts))
    return doc.to_bytes()


@pytest.fixture
def made(monkeypatch) -> list[tuple[str, str]]:
    """``(tool, made)`` of every layout a call used, in order: ``check`` for a check's and a
    changing call's (:func:`docx_agent.layout.note_made` in the layout module), ``render``
    for a render's."""
    seen: list[tuple[str, str]] = []

    def watch(name: str, real):
        def noted(document, facts):
            if isinstance(facts, dict) and "made" in facts:
                seen.append((name, facts["made"]))
            return real(document, facts)
        return noted

    monkeypatch.setattr(layout_module, "note_made", watch("check", layout_module.note_made))
    monkeypatch.setattr(shared_module, "note_made", watch("render", shared_module.note_made))
    return seen


def _body(document) -> list[str]:
    return [p.id for p in document.paragraphs() if len(p.text) > 200]


def test_a_check_after_a_render_of_the_same_version_lays_nothing_out_and_the_reverse(made):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=2) as box:
        box.warm(wait=True)
        session = box.session()
        doc = session.open(_report(), "report.docx")
        document = session.entry(doc).document
        key, _ = layout_module.affinity(document)

        def call(tool: str, **arguments) -> list[tuple[str, str]]:
            before = len(made)
            result = box.dispatch(session, tool, {"doc": doc, **arguments} if tool not in ("batch", "read_batch")
                                  else arguments)
            assert result.ok, result.to_json()
            return made[before:]

        def edit(k: int) -> dict:
            return {"doc": doc, "target": _body(document)[k], "text": _text(k, 40 + 10 * k)}

        assert call("render", pages=[1]) == [("render", "fresh")]
        assert call("check", include=["reflow"]) == [("check", "same")]             # render, then check
        assert call("word_set_text", **edit(1)) == [("check", "kept")]              # laid out ahead
        assert call("check", include=["reflow", "layout"]) == [("check", "same")]
        assert call("render", pages=[2]) == [("render", "same")]                    # check, then render
        assert call("word_set_text", **edit(2)) == [("check", "kept")]
        assert call("render", pages=[3], fast=True, width=640) == [("render", "same")]   # fast, another width
        assert call("check", range=[_body(document)[2]]) == [("check", "same")]
        # In one call: a read_batch's render and check, a batch's edit and the check after it.
        assert call("word_set_text", **edit(3)) == [("check", "kept")]
        assert call("read_batch", ops=[{"tool": "render", "arguments": {"doc": doc, "pages": [1], "fast": True}},
                                       {"tool": "check", "arguments": {"doc": doc}}]) == \
            [("render", "same"), ("check", "same")]
        assert call("batch", ops=[{"tool": "word_set_text", "arguments": edit(4)},
                                  {"tool": "check", "arguments": {"doc": doc, "include": ["reflow"]}}]) == \
            [("check", "kept"), ("check", "same")]
        assert call("render", pages=[2]) == [("render", "same")]
        counts = box.pool.affinity_counts
        assert counts["elsewhere"] == 0 and box.pool.holder(key) is not None


def test_a_busy_holder_sends_the_check_elsewhere_which_lays_out_again(made):
    """The one case found: the worker holding the layout busy longer than the wait the pool
    allows (as long as laying the document out whole took last) -- the check goes to the
    other worker, which lays the version out from what it holds (here nothing: whole).  The
    calls after it go to that worker, the holder now, and find its layout."""
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=2) as box:
        box.warm(wait=True)
        session = box.session()
        doc = session.open(_report(), "report.docx")
        document = session.entry(doc).document
        assert box.dispatch(session, "render", {"doc": doc, "pages": [1]}).ok
        key, wait = layout_module.affinity(document)
        holder = box.pool.holder(key)
        # The holder kept busy: a long layout task with the document's affinity.
        other = _report(chapters=40, paragraphs=12)
        options = layout_module.layout_options(document, {})
        busy = threading.Thread(target=lambda: box.pool.run(layout_module.lay_out_bytes, other, options, None,
                                                            timeout=300, affinity=key))
        busy.start()
        time.sleep(0.5)
        before = len(made)
        assert box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]}).ok
        busy.join(300)
        if box.pool.affinity_counts["elsewhere"] == 0:
            pytest.skip("the long task ended before the check came (a fast machine)")
        assert made[before:] == [("check", "fresh")]
        assert box.pool.holder(key) != holder
        before = len(made)
        assert box.dispatch(session, "render", {"doc": doc, "pages": [2]}).ok
        assert made[before:] == [("render", "same")]
        assert wait is not None
