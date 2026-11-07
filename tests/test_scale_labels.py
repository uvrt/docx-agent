"""A step's label in place of its value, in a Word document (round 38; ooxml-edit 0.87.0).

A scale with ``"show": "label"``: ``style_by_value`` writes each value's step's label where
the value was -- a cell reads ``High`` in the step's colours, not ``4.5`` -- and keeps, beside
the paragraph, the value it stands for and its reference, so ``check(scale=...)`` can say
later that a label is no longer its value's: the value changed at its source, the label was
edited by hand, or a label has no value kept at all.  Every document and figure is made up.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

scales = pytest.importorskip("ooxml_edit.tools.scales", reason="needs ooxml-edit 0.87 (scale labels)")
if not hasattr(scales, "SHOW"):
    pytest.skip("needs ooxml-edit 0.87 (scale labels)", allow_module_level=True)

from ooxml_edit.tools import DocumentFormat, Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

from test_tables_from_data import COLUMNS, ROWS, CLOCK, Grid, _grid_table, call, ok  # noqa: E402

#: The made-up grid, its tables read as values a check may read again.
LIVE = DocumentFormat(kind="grid", open=lambda data: Grid({}), detect=lambda data, name: False,
                      datasets={"t:": _grid_table},
                      value_sources=lambda grid: [f"t:{name}" for name in grid.tables])

STEPS = [{"min": 1, "fill": "C62828", "label": "Low"},
         {"min": 2.5, "fill": "F9A825", "label": "Middle"},
         {"min": 4, "fill": "2E7D32", "label": "High"}]

MARKDOWN = ("Alpha scored {{t:checks/Alpha/rating}} and Gamma {{t:checks/Gamma/rating}}.\n\n"
            "| Topic | Score |\n|---|---|\n| Door | 4.5 |\n| Roof | 2 |\n| Wall | n/a |\n\n"
            "Middle\n")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT, LIVE], groups=GROUPS, workers=0) as box:
        yield box


def _grid():
    return Grid({"checks": (COLUMNS, list(ROWS))})


def _session(toolbox, data=None):
    session = toolbox.session(clock=CLOCK)
    session.open(data or Document.new().to_bytes(), "report.docx")             # d1
    session.adopt(_grid(), "grid", "book.grid")                                  # d2
    if data is None:
        ok(call(toolbox, session, "word_insert_markdown", doc="d1", at="end", markdown=MARKDOWN))
    return session


def _doc(session):
    return session.entry("d1").document


def _table(session):
    return next(b["id"] for b in _doc(session).blocks() if b["kind"] == "table")


def _paragraph(session, start):
    return next(p.id for p in _doc(session).paragraphs() if p.text.startswith(start))


def labels(*targets):
    return {"name": "Score", "steps": STEPS, "show": "label", "targets": list(targets)}


def _body(doc):
    return zipfile.ZipFile(io.BytesIO(doc.to_bytes())).read("word/document.xml").decode()


def test_labels_take_the_values_place_in_their_steps_colours_in_one_undo_step(toolbox):
    session = _session(toolbox)
    table = _table(session)
    before = _doc(session).to_bytes()
    given = labels(f"{table}/col1", "{{t:checks/*/rating}}")
    preview = ok(call(toolbox, session, "style_by_value", doc="d1", scale=given, preview=True))
    assert preview.summary == "Score: would label 4 value(s), 0 already in their step, 1 skipped"
    assert _doc(session).to_bytes() == before
    result = ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    assert [(t["target"], t.get("value"), t.get("label"), t.get("skipped")) for t in result.data["targets"]] == [
        (f"{table}/c1,1", "4.5", "High", None), (f"{table}/c2,1", "2", "Low", None),
        (f"{table}/c3,1", "n/a", None, "'n/a' is not one number, nor a step's label"),
        (result.data["targets"][3]["target"], "3.5", "Middle", None),
        (result.data["targets"][4]["target"], "1.25", "Low", None)]
    doc = _doc(session)
    assert [p.text for p in doc.paragraphs()][:1] == ["Alpha scored Middle and Gamma Low."]
    body = _body(doc)
    assert '<w:rPr><w:rStyle w:val="Score3High"/></w:rPr><w:t>High</w:t>' in body
    assert '<w:shd w:val="clear" w:color="auto" w:fill="2E7D32"/></w:tcPr>' in body
    assert '<w:rPr><w:rStyle w:val="Score2Middle"/></w:rPr><w:t>Middle</w:t>' in body
    kept = [label for found in doc.text_labels().values() for label in found]
    assert [(k["text"], k["value"], k.get("source")) for k in kept] == [
        ("Middle", 3.5, "t:checks/Alpha/rating"), ("Low", 1.25, "t:checks/Gamma/rating"),
        ("High", 4.5, None), ("Low", 2, None)]
    assert doc.text_facts() == {}               # the values written by reference are labels now
    found = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert (found["right"], found["labels"], found["wrong"], found["unrecorded"]) == (4, [], [], [])
    again = ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    assert again.summary == "Score: labelled 0 value(s), 4 already in their step, 1 skipped"
    ok(call(toolbox, session, "undo", doc="d1"))             # the no-op made no step
    assert _doc(session).to_bytes() == before


def test_labels_and_their_values_survive_save_and_reopen(toolbox):
    session = _session(toolbox)
    table = _table(session)
    given = labels(f"{table}/col1", "{{t:checks/*/rating}}")
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    saved = _doc(session).to_bytes()
    item = next(name for name in zipfile.ZipFile(io.BytesIO(saved)).namelist()
                if re.fullmatch(r"customXml/item\d+\.xml", name))
    assert '"labels":[{' in zipfile.ZipFile(io.BytesIO(saved)).read(item).decode()
    reopened = _session(toolbox, saved)
    found = ok(call(toolbox, reopened, "check", doc="d1", scale=given)).data["scale"]
    assert (found["right"], found["labels"], found["unrecorded"]) == (4, [], [])
    assert Document.open(saved).text_labels() == _doc(session).text_labels()


def test_check_flags_a_changed_value_an_edited_label_and_one_with_no_value(toolbox):
    session = _session(toolbox)
    table = _table(session)
    lone = _paragraph(session, "Middle")
    ok(call(toolbox, session, "style_by_value", doc="d1",
            scale=labels(f"{table}/col1", "{{t:checks/*/rating}}")))
    header, rows = session.entry("d2").document.tables["checks"]
    rows[0] = ("Alpha", 4.25, *rows[0][2:])                                       # was 3.5
    ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
            cells=[{"row": 2, "col": 1, "text": "Lowish"}]))
    given = labels(f"{table}/col1", "{{t:checks/*/rating}}", lone)
    result = ok(call(toolbox, session, "check", doc="d1", scale=given))
    found = result.data["scale"]
    sentence = _paragraph(session, "Alpha")
    assert [(e["target"], e["shows"], e["value"], e["label"], e["why"]) for e in found["labels"]] == [
        (f"{table}/c2,1", "Lowish", "2", "Low", "edited: written 'Low'"),
        (f"{sentence}@13:19", "Middle", "4.25", "High", "its value changed: recorded 3.5, now 4.25")]
    assert found["wrong"] == [{"target": f"{sentence}@13:19", "from": "{{t:checks/*/rating}}", "value": "4.25",
                               "shows": "Middle", "step": 2, "due": 3}]
    assert found["unrecorded"] == [{"target": lone, "shows": "Middle", "label_of": 2}]
    assert (found["checked"], found["right"]) == (5, 2)
    fix = found["fix"]
    assert fix["arguments"]["scale"]["targets"] == [f"{sentence}@13:19", f"{table}/c2,1"]
    ok(call(toolbox, session, fix["tool"], **fix["arguments"]))
    assert _doc(session).paragraph(sentence).text == "Alpha scored High and Gamma Low."
    after = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert (after["right"], after["labels"], after["wrong"]) == (4, [], [])
    assert [e["target"] for e in after["unrecorded"]] == [lone] and "fix" not in after


def test_a_value_typed_over_a_label_is_the_value_again(toolbox):
    session = _session(toolbox)
    table = _table(session)
    given = labels(f"{table}/col1")
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
            cells=[{"row": 1, "col": 1, "text": "3"}]))
    found = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert found["labels"][0]["why"] == "shows a value typed over its label"
    assert found["labels"][0]["label"] == "Middle"
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    assert _doc(session).get(f"{table}/c1,1").text == "Middle"
    kept = [k for found in _doc(session).text_labels().values() for k in found]
    assert sorted(k["value"] for k in kept) == [2, 3]


def test_tracked_the_label_is_an_insertion_and_a_deletion(toolbox):
    session = _session(toolbox)
    table = _table(session)
    ok(call(toolbox, session, "word_set_tracking", doc="d1", on=True, author="Reviewer"))
    given = labels(f"{table}/col1")
    result = ok(call(toolbox, session, "style_by_value", doc="d1", scale=given))
    assert result.checks["validate"]["new"] == []
    body = _body(_doc(session))
    assert "<w:delText>4.5</w:delText>" in body and "<w:t>High</w:t>" in body
    assert ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]["right"] == 2


def test_exact_text_steps_still_style_the_text_as_it_is(toolbox):
    session = _session(toolbox)
    lone = _paragraph(session, "Middle")
    texts = {"name": "Grade", "steps": [{"value": "Low", "fill": "C62828"}, {"value": "Middle", "fill": "F9A825"},
                                        {"value": "High", "fill": "2E7D32"}], "targets": [lone]}
    result = ok(call(toolbox, session, "style_by_value", doc="d1", scale=texts))
    assert result.data["targets"] == [{"target": lone, "value": "Middle", "step": 2}]
    assert _doc(session).paragraph(lone).text == "Middle"
    assert ok(call(toolbox, session, "check", doc="d1", scale=texts)).data["scale"]["right"] == 1
