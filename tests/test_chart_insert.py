"""doc.insert_chart (LW3) and edit_chart's add action in a document.

The chart part and its workbook are ooxml-edit's (``ooxml_edit.charts.create``); here: the
inline drawing Word writes, the new paragraph after a block, one undo step, the document
staying valid, the workbook agreeing with the cache, tracked insertion, and the tool.
"""

from __future__ import annotations

import warnings

import pytest

from docx_agent import Document
from ooxml_edit.charts import ChartDataError
from ooxml_edit.tools import Toolbox

VOLUMES = [{"name": "Volume", "values": [120, 135, 128, 141]}]
MONTHS = ["Jan", "Feb", "Mar", "Apr"]


def report() -> Document:
    document = Document.new()
    document.insert_markdown("# Volumes\n\nVolumes rose all year.\n\nThe end.", at="end")
    return document


def test_a_chart_after_a_paragraph():
    document = report()
    body = document.paragraphs()
    edit = document.insert_chart(body[1].id, "line", MONTHS, VOLUMES, title="Monthly volumes",
                                 axis_titles={"category": "Month", "value": "Units"})
    after = document.paragraphs()
    assert len(after) == len(body) + 1 and after[2].id in edit.created
    chart = document.chart(edit.id)
    assert chart.chart_type == "line" and chart.categories == MONTHS
    assert chart.axis_title("value") == "Units" and chart.title == "Monthly volumes"
    book = chart.workbook_values()
    assert book["categories"]["values"] == MONTHS
    assert book["series"][0]["values"]["values"] == VOLUMES[0]["values"]
    assert document.validate() == []
    drawing = document.drawing(edit.id)
    assert drawing.inline and [round(v) for v in drawing.size] == [432, 252]


def test_a_chart_after_a_heading_is_not_a_heading():
    document = report()
    heading = document.paragraphs()[0]
    edit = document.insert_chart(heading.id, "column", MONTHS, VOLUMES)
    new = document.paragraph(edit.created[0])
    assert not (new.style_name or "").lower().startswith("heading")


def test_width_keeps_word_s_ratio_and_undo_restores():
    document = report()
    before = document.to_bytes()
    edit = document.insert_chart(document.paragraphs()[1].id, "pie", MONTHS, VOLUMES, width=216)
    drawing = document.drawing(edit.id)
    assert [round(v) for v in drawing.size] == [216, 126]
    document.undo()
    assert document.to_bytes() == before


def test_bad_data_changes_nothing():
    document = report()
    before = document.to_bytes()
    with pytest.raises(ChartDataError):
        document.insert_chart(document.paragraphs()[1].id, "pie", MONTHS, VOLUMES * 2)
    assert document.to_bytes() == before


def test_tracked_insertion():
    document = report()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        with document.tracking(author="Claude"):
            document.insert_chart(document.paragraphs()[1].id, "column", MONTHS, VOLUMES)
    assert any(r.kind == "insertion" for r in document.revisions())
    assert document.validate() == []


def test_the_tool_adds_a_chart_and_titles_its_axes_in_one_batch():
    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    document = report()
    anchor = document.paragraphs()[1].id
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session = box.session()
        session.open(document.to_bytes(), "report.docx")
        result = box.dispatch(session, "batch", {"ops": [
            {"tool": "edit_chart", "arguments": {"doc": "d1", "target": anchor, "action": "add",
                                                 "chart_type": "line", "categories": MONTHS,
                                                 "data": VOLUMES, "text": "Monthly volumes",
                                                 "ref": "c"}},
            {"tool": "edit_chart", "arguments": {"doc": "d1", "target": "$c",
                                                 "action": "set_axis_title", "axis": "category",
                                                 "text": "Month"}},
            {"tool": "edit_chart", "arguments": {"doc": "d1", "target": "$c",
                                                 "action": "set_axis_title", "axis": "value",
                                                 "text": "Units"}}]})
        assert result.ok, result.to_json()
        refused = box.dispatch(session, "edit_chart", {"doc": "d1", "target": anchor,
                                                        "action": "add", "chart_type": "line",
                                                        "categories": MONTHS, "data": VOLUMES,
                                                        "box": {"x": 0, "y": 0, "w": 1, "h": 1}})
        assert not refused.ok and refused.error.field == "box"
        bad = box.dispatch(session, "edit_chart", {"doc": "d1", "target": anchor, "action": "add",
                                                    "chart_type": "column", "categories": MONTHS,
                                                    "data": [{"name": "A", "values": [1]}]})
        assert not bad.ok and bad.error.code == "invalid_arguments"
