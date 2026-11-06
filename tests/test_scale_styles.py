"""Values styled by a scale in a Word document (round 36, R31; ooxml-edit 0.85.0).

``style_by_value`` reads each target's value -- a paragraph, a range, a cell, a table's
column, the values written by reference -- and gives it its step: a character style per step
(``Score 1 Low``, as Word writes one made with a colour and shading), cell shading for a cell.
``check(scale=...)`` lists the values styled with another step than their value's, or none,
with the call that fixes them.  Every document and figure is made up.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

pytest.importorskip("ooxml_edit.tools.scales", reason="needs ooxml-edit 0.85 (scales)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

from test_tables_from_data import COLUMNS, GRID, ROWS, CLOCK, Grid, call, failed, ok  # noqa: E402

STEPS = [{"min": 1, "fill": "C62828", "label": "Low"},
         {"min": 2.5, "fill": "F9A825"},
         {"min": 4, "fill": "2E7D32", "label": "High"}]

MARKDOWN = ("Alpha scored {{t:checks/Alpha/rating}} and Gamma {{t:checks/Gamma/rating}}.\n\n"
            "| Topic | Score |\n|---|---|\n| Door | 4.5 |\n| Roof | 2 |\n| Wall | n/a |\n\n"
            "3\n\nThe mean is 4,2 here.\n")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


def _session(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "report.docx")                       # d1
    session.adopt(Grid({"checks": (COLUMNS, ROWS)}), "grid", "book.grid")        # d2
    ok(call(toolbox, session, "word_insert_markdown", doc="d1", at="end", markdown=MARKDOWN))
    return session


def _doc(session):
    return session.entry("d1").document


def _ids(session):
    doc = _doc(session)
    table = next(b["id"] for b in doc.blocks() if b["kind"] == "table")
    three = next(p.id for p in doc.paragraphs() if p.text == "3")
    mean = next(p.id for p in doc.paragraphs() if p.text.startswith("The mean"))
    return table, three, mean


def scale(*targets, **more):
    return {"name": "Score", "steps": STEPS, "targets": list(targets), **more}


def _styles(doc):
    xml = zipfile.ZipFile(io.BytesIO(doc.to_bytes())).read("word/styles.xml").decode()
    return {m.group(1): m.group(0) for m in re.finditer(r'<w:style [^>]*w:styleId="(Score[^"]*)".*?</w:style>', xml)}


def _body(doc):
    return zipfile.ZipFile(io.BytesIO(doc.to_bytes())).read("word/document.xml").decode()


def test_each_value_gets_its_steps_style_in_one_undo_step(toolbox):
    session = _session(toolbox)
    table, three, mean = _ids(session)
    before = _doc(session).to_bytes()
    result = ok(call(toolbox, session, "style_by_value", doc="d1",
                     scale=scale(f"{table}/col1", "{{t:checks/*/rating}}", three, f"{mean}@12:15")))
    targets = result.data["targets"]
    assert [(t["target"], t.get("step"), t.get("skipped")) for t in targets] == [
        (f"{table}/c1,1", 3, None), (f"{table}/c2,1", 1, None),
        (f"{table}/c3,1", None, "'n/a' is not one number"),
        (targets[3]["target"], 2, None), (targets[4]["target"], 1, None),
        (three, 2, None), (f"{mean}@12:15", 3, None)]
    assert targets[3]["from"] == "{{t:checks/*/rating}}" and targets[3]["value"] == "3.5"
    assert result.summary == "Score: styled 6 value(s), 0 already in their step, 1 skipped"
    assert sorted(result.created) == ["style:Score1Low", "style:Score2", "style:Score3High"]
    assert result.data["steps"][0] == {"step": 1, "label": "Low", "style": "Score 1 Low", "fill": "C62828",
                                       "text": "FFFFFF", "contrast": 5.62}
    # As Word writes a character style made with a font colour and shading (measured).
    styles = _styles(_doc(session))
    assert styles["Score1Low"] == (
        '<w:style w:type="character" w:customStyle="1" w:styleId="Score1Low"><w:name w:val="Score 1 Low"/>'
        '<w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:color w:val="FFFFFF"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="C62828"/></w:rPr></w:style>')
    body = _body(_doc(session))
    assert '<w:shd w:val="clear" w:color="auto" w:fill="2E7D32"/></w:tcPr>' in body
    assert '<w:rPr><w:rStyle w:val="Score2"/></w:rPr><w:t>3.5</w:t>' in body
    assert '<w:rPr><w:rStyle w:val="Score3High"/></w:rPr><w:t>4,2</w:t>' in body
    ok(call(toolbox, session, "undo", doc="d1"))
    assert _doc(session).to_bytes() == before


def test_preview_changes_nothing_and_a_second_run_reuses_the_styles(toolbox):
    session = _session(toolbox)
    table, three, _ = _ids(session)
    before = _doc(session).to_bytes()
    preview = ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(f"{table}/col1"), preview=True))
    assert preview.summary.startswith("Score: would style 2 value(s)")
    assert preview.data["styles"] == [{"style": "Score 1 Low", "made": True},
                                      {"style": "Score 3 High", "made": True}]
    assert _doc(session).to_bytes() == before and preview.changed == []
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(f"{table}/col1")))
    again = ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(f"{table}/col1", three)))
    assert again.data["styles"] == [{"style": "Score 2", "made": True}]
    assert [t.get("unchanged") for t in again.data["targets"][:2]] == [True, True]
    assert again.created == ["style:Score2"]


def test_a_steps_colours_changed_update_its_style_and_say_so(toolbox):
    session = _session(toolbox)
    table, _, _ = _ids(session)
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(f"{table}/col1")))
    steps = [dict(STEPS[0]), STEPS[1], {**STEPS[2], "fill": "1B5E20"}]
    changed = ok(call(toolbox, session, "style_by_value", doc="d1",
                      scale={"name": "Score", "steps": steps, "targets": [f"{table}/c1,1"]}))
    assert changed.data["styles"] == [{"style": "Score 3 High", "changed": True}]
    assert "Score 3 High: its colours set to the step's (fill 1B5E20" in changed.warnings[0]
    assert 'w:fill="1B5E20"' in _styles(_doc(session))["Score3High"]


def test_check_reports_a_value_changed_after_it_was_styled_and_the_fix(toolbox):
    session = _session(toolbox)
    table, three, _ = _ids(session)
    given = scale(f"{table}/col1", three)
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(f"{table}/col1")))
    ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
            cells=[{"row": 1, "col": 1, "text": "1"}]))
    found = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert found["wrong"] == [{"target": f"{table}/c1,1", "from": f"{table}/col1", "value": "1", "step": 3,
                               "due": 1}]
    assert found["none"] == [{"target": three, "value": "3", "due": 2}]
    assert (found["checked"], found["right"]) == (3, 1)
    assert found["skipped"][0]["why"] == "'n/a' is not one number"
    assert found["fix"]["arguments"]["scale"]["targets"] == [f"{table}/c1,1", three]
    ok(call(toolbox, session, found["fix"]["tool"], **found["fix"]["arguments"]))
    after = ok(call(toolbox, session, "check", doc="d1", scale=given)).data["scale"]
    assert (after["right"], after["wrong"], after["none"]) == (3, [], [])


def test_check_with_scale_alone_is_only_its_fact_and_range_scopes_it(toolbox):
    session = _session(toolbox)
    table, three, _ = _ids(session)
    alone = ok(call(toolbox, session, "check", doc="d1", scale=scale(f"{table}/col1", three)))
    assert set(alone.data) == {"scale"}
    scoped = ok(call(toolbox, session, "check", doc="d1", range=[three],
                     scale=scale(f"{table}/col1", three))).data["scale"]
    assert [e["target"] for e in scoped["none"]] == [three] and scoped["checked"] == 1
    both = ok(call(toolbox, session, "check", doc="d1", include=["validate"], scale=scale(three)))
    assert "validate" in both.data and both.data["scale"]["none"]


def test_direct_colour_is_taken_off_a_styled_value(toolbox):
    session = _session(toolbox)
    _, three, _ = _ids(session)
    ok(call(toolbox, session, "word_format", doc="d1", targets=[three], color="0000FF"))
    found = ok(call(toolbox, session, "check", doc="d1", scale=scale(three))).data["scale"]
    assert found["none"]                                  # no step yet: blue text
    ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(three)))
    body = _body(_doc(session))           # the run's own colour goes; the paragraph mark's stays
    assert '<w:r><w:rPr><w:rStyle w:val="Score2"/></w:rPr><w:t>3</w:t></w:r>' in body
    assert ok(call(toolbox, session, "check", doc="d1", scale=scale(three))).data["scale"]["right"] == 1


def test_targets_that_name_nothing_are_skipped_with_the_way_out(toolbox):
    session = _session(toolbox)
    table, _, mean = _ids(session)
    result = ok(call(toolbox, session, "style_by_value", doc="d1",
                     scale=scale(table, f"{table}/col7", "p:00000000", mean, "{{d:none/*/x}}")))
    why = [t["skipped"] for t in result.data["targets"]]
    assert why[0].endswith("is a table: name its column (" + table + "/col1) or its cells")
    assert why[1] == f"{table} has grid columns 0 to 1: no column 7"
    assert why[2].startswith("no p:00000000")
    assert why[3] == "'The mean is 4,2 here.' is not one number"
    assert why[4].startswith("{{d:none/*/x}}: no value written by a reference")
    assert result.warnings[-1].startswith("nothing to change")


def test_a_refused_scale_says_how(toolbox):
    session = _session(toolbox)
    error = failed(call(toolbox, session, "style_by_value", doc="d1",
                        scale={"steps": [{"min": 3, "fill": "2E7D32"}, {"min": 1, "fill": "C62828"}],
                               "targets": ["x"]}), "invalid_arguments")
    assert "lowest first" in error.message


def test_tracked_the_styling_is_a_formatting_revision(toolbox):
    session = _session(toolbox)
    table, three, _ = _ids(session)
    ok(call(toolbox, session, "word_set_tracking", doc="d1", on=True, author="Reviewer"))
    result = ok(call(toolbox, session, "style_by_value", doc="d1", scale=scale(f"{table}/col1", three)))
    assert result.checks["validate"]["new"] == []
    body = _body(_doc(session))
    assert "<w:rPrChange" in body and "<w:tcPrChange" in body
