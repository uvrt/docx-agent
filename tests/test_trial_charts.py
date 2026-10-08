"""``chart.workbook_values()`` (ROADMAP.md, "Trial findings", 7): what Word's Edit Data shows for
a chart's ranges, read from its embedded workbook, so an agent can check that the workbook
and the drawing agree without unzipping anything."""

from __future__ import annotations

import pytest

from conftest import CHARTS_DIR
from docx_agent import Document

CHARTS = CHARTS_DIR / "charts.docx"


def _agree(chart) -> bool:
    book = chart.workbook_values()
    categories = book["categories"]
    return (categories is None or categories["values"] == chart.categories) and \
        [s["values"]["values"] for s in book["series"]] == [s.values for s in chart.series] and \
        [s["name"]["value"] for s in book["series"]] == [s.name for s in chart.series]


@pytest.mark.parametrize("identifier", [chart.id for chart in Document.open(CHARTS).charts()])
def test_the_workbook_agrees_with_the_chart_as_word_saved_it(identifier):
    document = Document.open(CHARTS)
    chart = document.chart(identifier)
    book = chart.workbook_values()
    assert book["workbook"] == chart.workbook_part
    assert all(s["values"]["ref"].startswith("Sheet1!") for s in book["series"])
    assert _agree(chart)


def test_it_follows_every_edit_and_reads_nothing_into_the_document():
    document = Document.open(CHARTS)
    chart = document.chart("d:1")
    before = document.to_bytes()
    first = chart.workbook_values()
    assert first["categories"] == {"ref": "Sheet1!$A$2:$A$5", "values": ["Q1", "Q2", "Q3", "Q4"]}
    assert first["series"][0]["name"] == {"ref": "Sheet1!$B$1", "value": "North"}
    assert document.to_bytes() == before                       # reading changes nothing
    chart.series[0].set_value(1, 999)
    chart.add_category("Q5", [1, 2, 3])
    chart.series[1].name = "Southeast"
    book = chart.workbook_values()
    assert book["categories"]["ref"] == "Sheet1!$A$2:$A$6"
    assert book["series"][0]["values"]["values"] == [4.3, 999, 3.5, 4.5, 1]
    assert book["series"][1]["name"]["value"] == "Southeast"
    assert _agree(chart)
    document.undo()
    assert chart.workbook_values()["series"][1]["name"]["value"] == "South"


def test_a_chart_without_a_workbook_says_why():
    document = Document.open(CHARTS)
    chart = document.chart("d:1")
    root = document.package.tree(chart.part)
    root.remove(root.find("{http://schemas.openxmlformats.org/drawingml/2006/chart}externalData"))
    book = chart.workbook_values()
    assert book["workbook"] is None and book["missing"] == "the chart has no workbook"
