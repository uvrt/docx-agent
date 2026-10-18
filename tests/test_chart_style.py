"""A chart's look in a document: axis scales, gridlines, colours and markers, combination
charts, chart text, point labels, filled radar and bubble charts (ooxml-edit's ``Chart``
formatting, tested there) -- through the library and through ``edit_chart``.

Here: the document stays valid, the workbook agrees with the cache, every call is one undo
step, and the tool's arguments reach the library with Word's sizes.
"""

from __future__ import annotations

import pytest

from docx_agent import Document
from ooxml_edit.tools import Toolbox

C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
CATEGORIES = ["Governance", "People", "Data", "Technology", "Procure"]
CURRENT = [2.13, 2.10, 2.17, 2.33, 2.08]
TARGET = [4, 3, 4, 3, 4]


def report() -> Document:
    document = Document.new()
    document.insert_markdown("# Maturity\n\nScores per area.\n\nThe end.", at="end")
    return document


def test_the_maturity_gap_chart_in_a_document():
    document = report()
    edit = document.insert_chart(document.paragraphs()[1].id, "column", CATEGORIES,
                                 [{"name": "Current", "values": CURRENT},
                                  {"name": "Target", "values": TARGET, "type": "line"}],
                                 gridlines=False)
    chart = document.chart(edit.id)
    before = document.to_bytes()
    chart.set_axis_scale("value", minimum=0, maximum=5, major_unit=1)
    chart.format_series("Current", fill="tx2")
    chart.format_point("Current", "Procure", fill="accent2")
    chart.format_series("Target", line=None, marker="diamond", marker_size=10,
                        marker_fill="accent6")
    chart.set_text_style(size=10)
    assert chart.chart_types == ["bar", "line"]
    assert chart.scales["value"]["maximum"] == 5
    assert chart.workbook_values()["series"][1]["values"]["values"] == TARGET
    assert document.validate() == []
    for _ in range(5):
        document.undo()
    assert document.to_bytes() == before


@pytest.mark.parametrize("kind", ["radar_filled", "bubble"])
def test_the_new_kinds_in_a_document(kind):
    document = report()
    anchor = document.paragraphs()[1].id
    if kind == "bubble":
        edit = document.insert_chart(anchor, "bubble", [1, 2, 4],
                                     [{"name": "Recs", "values": [4.5, 3, 4], "size": [3, 2, 1]}])
    else:
        edit = document.insert_chart(anchor, "radar_filled", ["Speed", "Cost", "Risk"],
                                     [{"name": "Now", "values": [3, 2, 4]}])
    chart = document.chart(edit.id)
    assert chart.chart_type == ("bubble" if kind == "bubble" else "radar")
    # Word's sizes (measured): 9 pt labels.
    sizes = {n.get("sz") for n in chart._root().iter(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr") if n.get("sz")}
    assert "900" in sizes
    assert document.validate() == []


def test_the_tool_formats_a_chart_and_labels_a_matrix():
    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    document = report()
    anchor = document.paragraphs()[1].id
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session = box.session()
        session.open(document.to_bytes(), "report.docx")

        def run(**arguments):
            result = box.dispatch(session, "edit_chart", {"doc": "d1", **arguments})
            assert result.ok, result.to_json()
            return result

        run(target=anchor, action="add", chart_type="bar", categories=["Worst", "Mid", "Best"],
            data=[{"name": "Score", "values": [1.8, 2.4, 3.1]}], axis="category",
            scale={"reverse": True}, gridlines="none", ref="r")
        read = run(target="$r", action="read")
        assert read.data["scales"]["category"]["reverse"] is True
        assert read.data["scales"]["value"]["major_gridlines"] is False
        run(target=anchor, action="add", chart_type="bubble", values=[1, 2, 4],
            data=[{"name": "Recs", "values": [4.5, 3, 4], "size": [3, 2, 1]}], ref="b")
        labelled = run(target="$b", action="set_point_labels", series="Recs",
                       labels=["R1", "R2", "R3"], position="right", cells=True)
        assert labelled.data["point_labels"] == {"Recs": ["R1", "R2", "R3"]}
        run(target="$b", action="set_text_style", part="data_labels", style={"size": 8})
        refused = box.dispatch(session, "edit_chart", {"doc": "d1", "target": "$r",
                                                        "action": "set_series_type",
                                                        "series": "Score", "chart_type": "line"})
        assert not refused.ok and "horizontal bars" in refused.error.message
        assert session.entry("d1").document.validate() == []


# -- add with its look (round 5, T2) -------------------------------------------------------------


@pytest.mark.parametrize("kind, data, extra", [
    ("column", [{"name": "Now", "values": [3, 2, 4]},
                {"name": "Goal", "values": [4, 4, 4], "type": "line"}], {}),
    ("line", [{"name": "Now", "values": [3, 2, 4]}, {"name": "Goal", "values": [4, 4, 4]}], {}),
    ("pie", [{"name": "Now", "values": [3, 2, 4]}], {}),
    ("scatter", [{"name": "Now", "values": [3, 2, 4]}], {"values": [1, 2, 3]}),
    ("dot", [{"name": "Now", "values": [3, 2, 4]}, {"name": "Goal", "values": [4, 4, 4]}], {}),
])
def test_add_with_its_look_writes_the_chart_the_separate_calls_write(kind, data, extra):
    from lxml import etree

    from docx_agent.tools import FORMAT, GROUPS, TOOLS
    from ooxml_edit.charts.actions import add_style_actions

    document = report()
    anchor = document.paragraphs()[1].id
    # A colour per series, a pie's per slice; the scatter chart's one series takes one.
    look = {"data_labels": True, "colors": ["accent2", "1F4E79"][:1 if kind == "scatter" else 2],
            "style": {"size": 8, "bold": True}}
    arguments = {"doc": "d1", "target": anchor, "action": "add", "chart_type": kind,
                 "data": data, "width": 360, "gridlines": "none", **extra}
    if not extra:
        arguments["categories"] = ["Speed", "Cost", "Risk"]
    written = []
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        for together in (True, False):
            session = box.session()
            session.open(document.to_bytes(), "report.docx")
            added = box.dispatch(session, "edit_chart", {**arguments, **(look if together
                                                                          else {})})
            assert added.ok, added.to_json()
            chart_id = next(c for c in added.created if c.startswith("d:"))
            if not together:
                for action, given in add_style_actions(kind, data, **look):
                    done = box.dispatch(session, "edit_chart", {"doc": "d1", "target": chart_id,
                                                                "action": action, **given})
                    assert done.ok, done.to_json()
            edited = session.entry("d1").document
            written.append(etree.tostring(edited.chart(chart_id)._root()))
            assert edited.validate() == []
    assert written[0] == written[1]
    assert b'val="accent2"' in written[0] and b'sz="800"' in written[0]
    assert b"c:dLbls" in written[0]


# -- the dot chart (round 6, R10) ---------------------------------------------------------------


def test_a_dot_chart_on_a_fixed_scale_reads_back_as_dot():
    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    document = report()
    anchor = document.paragraphs()[1].id
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session = box.session()
        session.open(document.to_bytes(), "report.docx")

        def run(**arguments):
            result = box.dispatch(session, "edit_chart", {"doc": "d1", **arguments})
            assert result.ok, result.to_json()
            return result

        added = run(target=anchor, action="add", chart_type="dot",
                    categories=["North", "South", "East"],
                    data=[{"name": "Mean", "values": [3.2, 4.1, 2.7]}], axis="value",
                    scale={"min": 1, "max": 5}, data_labels=True, ref="dots")
        assert added.data["type"] == "dot"
        read = run(target="$dots", action="read")
        assert read.data["type"] == "dot" and read.data["types"] == ["line"]
        assert (read.data["scales"]["value"]["minimum"],
                read.data["scales"]["value"]["maximum"]) == (1, 5)
        run(target="$dots", action="set_values", series="Mean", values=[1, 2, 3])
        assert run(target="$dots", action="add_series", text="Goal",
                   values=[4, 4, 4]).data["type"] == "dot"
        edited = session.entry("d1").document
        chart = edited.chart(next(c for c in added.created if c.startswith("d:")))
        assert chart.is_dot and chart.chart_type == "line"
        assert edited.validate() == []


# -- round 9, N2: a scale's default axis; refusals that name what to give -------------------------


def test_a_scale_without_an_axis_and_what_a_refusal_names():
    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    document = report()
    anchor = document.paragraphs()[1].id
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session = box.session()
        session.open(document.to_bytes(), "report.docx")

        def call(**arguments):
            return box.dispatch(session, "edit_chart", {"doc": "d1", **arguments})

        made = call(target=anchor, action="add", chart_type="column", categories=["A", "B", "C"],
                    data=[{"name": "Score", "values": [1, 2, 3]}], scale={"max": 4}, ref="c")
        assert made.ok, made.to_json()
        assert call(target="$c", action="read").data["scales"]["value"]["maximum"] == 4
        assert call(target="$c", action="set_axis_scale", scale={"min": 1}).ok
        before = session.entry("d1").document.to_bytes()
        scatter = call(target=anchor, action="add", chart_type="scatter", values=[1, 2],
                       data=[{"name": "Y", "values": [3, 4]}], scale={"max": 5})
        assert scatter.error.message.startswith("add with scale needs axis: x or y")
        assert session.entry("d1").document.to_bytes() == before      # nothing stays
        for arguments, words in [
                ({"action": "set_value", "series": "Score", "value": 2},
                 "set_value needs category: 'A', 'B' or 'C' (or a number from 0 to 2)"),
                ({"action": "rename_series", "text": "Now"},
                 "rename_series needs series: 'Score' (or 0)"),
                ({"action": "set_values", "series": "Score"},
                 "set_values needs values: one number per category (3)"),
                ({"action": "add_category", "category": "D"},
                 "add_category needs values: one number per series (1)"),
                ({"action": "set_legend"}, "set_legend needs position: right, left, top"),
                ({"action": "set_axis_scale", "scale": {"reverse": True}},
                 "set_axis_scale needs axis: category or value"),
                ({"action": "format_point", "series": "Score", "style": {"fill": "tx2"}},
                 "format_point needs category: 'A', 'B' or 'C'")]:
            error = call(target="$c", **arguments).error
            assert error.code == "invalid_arguments" and words in error.message, error.message
        missing = call(target=anchor, action="add", categories=["A"],
                       data=[{"name": "S", "values": [1]}]).error
        assert "add needs chart_type: one of column, stacked_column" in missing.message
