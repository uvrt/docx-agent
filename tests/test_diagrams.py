"""SmartArt (ROADMAP.md, "Charts and SmartArt"): node text, nodes added and removed, on the
diagrams Word laid out and saved (``tests/fixtures/generated/charts/smartart.docx``).

Every edit goes through the gates the charts do -- edit, save, reopen, read back; validity;
undo and redo byte for byte; docx2svg's render -- and the cached drawing is held to the
policy Word's behaviour chose (``tools/charts_probe.py``, ``smartart``): Word lays every
diagram out again from its data model on opening and rewrites the drawing on saving, so a
drawing an edit cannot keep exactly in step is dropped, with a warning, and docx2svg -- which
draws only the cache -- draws a placeholder until Word saves the document again.
"""

from __future__ import annotations

import warnings
from pathlib import Path

import pytest

from docx_agent import Document
from docx_agent.edit.charts import (DIAGRAM_POLICY, ChartEditError, DiagramDrawingDropped,
                                    DiagramDrawingError)
from docx_agent.validate import check

from test_roundtrip import entries

SMARTART = Path(__file__).parent / "fixtures" / "generated" / "charts" / "smartart.docx"


def test_reading():
    document = Document.open(SMARTART)
    first, second = document.diagrams()
    assert first.id == "d:1" and first.layout.endswith("/layout/default")
    assert first.texts == ["Plan", "Write", "Measure", "Review", "Ship"]
    assert [(n.level, n.text) for n in second.nodes] == [(0, "Goals"), (1, "Faster edits"), (1, "Fewer prompts"),
                                                         (0, "Risks"), (1, "Stale caches"), (1, "Lost formulas")]
    assert second.node(0).children[1].text == "Fewer prompts"
    assert first.drawing_part and second.drawing_part
    assert document.drawing("d:2").diagram.id == "d:2"


def _set_text_exactly(document):
    diagram = document.diagram("d:2")
    diagram.set_text(1, "Much faster edits")
    return lambda d: d.diagram("d:2").texts[1] == "Much faster edits", "Much faster edits", False


def _set_top_text_exactly(document):
    document.diagram("d:1").node(4).text = "Shipped"
    return lambda d: d.diagram("d:1").texts[4] == "Shipped", "Shipped", False


def _add_child(document):
    document.diagram("d:2").node(0).add_child("Smaller files")
    return lambda d: d.diagram("d:2").node(0).children[-1].text == "Smaller files", None, True


def _add_top_node(document):
    document.diagram("d:1").add_node("Celebrate", index=2)
    return lambda d: d.diagram("d:1").texts[2] == "Celebrate", None, True


def _remove_node(document):
    diagram = document.diagram("d:2")
    diagram.remove_node(diagram.node(3))
    return lambda d: d.diagram("d:2").texts == ["Goals", "Faster edits", "Fewer prompts"], None, True


def _apply_model(document):
    diagram = document.diagram("d:1")
    model = diagram.model
    model["nodes"][0]["t"] = "Plan well"
    del model["nodes"][-1]
    diagram.apply(model)
    return lambda d: d.diagram("d:1").texts == ["Plan well", "Write", "Measure", "Review"], None, True


EDITS = [_set_text_exactly, _set_top_text_exactly, _add_child, _add_top_node, _remove_node, _apply_model]


@pytest.mark.parametrize("edit", EDITS, ids=lambda f: f.__name__.strip("_"))
def test_diagram_edit(edit):
    data = SMARTART.read_bytes()
    document = Document.open(data)
    before = set(check(document.package))
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        read_back, shown, drops = edit(document)
    dropped = [w for w in caught if issubclass(w.category, DiagramDrawingDropped)]
    assert bool(dropped) == drops, [str(w.message) for w in caught]
    if drops:
        assert "placeholder until Word saves the document again" in str(dropped[0].message)
    edited = document.to_bytes()
    assert len(document.history._undo) == 1  # one undo step

    saved = Document.open(edited)
    assert read_back(saved)
    assert set(check(saved.package)) - before == set()
    changed = {n for n, content in entries(edited).items() if entries(data).get(n) != content}
    removed = set(entries(data)) - set(entries(edited))
    assert changed and all(n.startswith("word/diagrams/") or n in ("word/_rels/document.xml.rels",
                                                                   "[Content_Types].xml") for n in changed)
    assert all(n.startswith("word/diagrams/drawing") for n in removed)
    assert bool(removed) == drops

    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)
    while document.redo():
        pass
    assert entries(document.to_bytes()) == entries(edited)

    # docx2svg draws the cache: the exact edit's text.  A dropped drawing it cannot find, and
    # since ooxml-common 0.4.4 it no longer borrows another diagram's drawing: it draws a
    # placeholder and says so, until Word saves the document again.
    layout = document.layout()
    new = {w.code for w in layout.warnings} - {w.code for w in Document.open(data).layout().warnings}
    assert new == ({"diagram-no-cached-drawing", "drawing-not-drawn"} if drops else set())
    text = "".join(document.render_svg())
    if shown:
        assert shown in text
    if drops:
        assert sorted([text.count(">Plan<"), text.count(">Goals<")]) == [0, 1]


def test_a_dropped_drawing_is_a_placeholder_in_docx2svg():
    """With no cached drawing left on the part, docx2svg draws each diagram as a placeholder
    and says so, until Word saves the document again."""
    document = Document.open(SMARTART)
    with pytest.warns(DiagramDrawingDropped):
        document.diagram("d:1").add_node("Sixth")
        document.diagram("d:2").node(0).add_child("Smaller files")
    codes = [w.code for w in document.layout().warnings]
    assert codes.count("diagram-no-cached-drawing") == 2
    text = "".join(document.render_svg())
    assert ">Plan<" not in text and ">Goals<" not in text


def test_the_policy_is_dropping():
    assert DIAGRAM_POLICY == "drop"


def test_keeping_a_stale_drawing_is_a_choice_and_says_so():
    document = Document.open(SMARTART)
    with pytest.warns(DiagramDrawingDropped, match="no longer shows what the data model says"):
        document.diagram("d:2", on_inexact_drawing="keep").node(0).add_child("Kept stale")
    assert document.diagram("d:2").drawing_part is not None


def test_refusing_changes_nothing():
    document = Document.open(SMARTART)
    data = document.to_bytes()
    with pytest.raises(DiagramDrawingError):
        document.diagram("d:2", on_inexact_drawing="refuse").node(0).add_child("Refused")
    assert document.to_bytes() == data and not document.history.can_undo()


def test_a_refused_model_changes_nothing():
    document = Document.open(SMARTART)
    data = document.to_bytes()
    diagram = document.diagram("d:2")
    model = diagram.model
    model["nodes"].reverse()
    with pytest.raises(ChartEditError):
        diagram.apply(model)
    with pytest.raises(ChartEditError):
        diagram.apply({"nodes": "none"})
    with pytest.raises(KeyError):
        diagram.remove_node("{no-such-node}")
    with pytest.raises(KeyError):
        diagram.add_node("Orphan", parent="{no-such-node}")
    assert document.to_bytes() == data and not document.history.can_undo()


def test_a_diagram_keeps_one_node():
    document = Document.open(SMARTART)
    diagram = document.diagram("d:1")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        for _ in range(4):
            diagram.remove_node(0)
        with pytest.raises(ValueError):
            diagram.remove_node(0)
    assert diagram.texts == ["Ship"]
