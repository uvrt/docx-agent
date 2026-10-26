"""S10 (review): ``edit_chart`` and ``edit_smartart`` behave the same in a Word document and
a deck.  ooxml-edit's table (``ooxml_edit.charts.cases``) is run here through this
package's tools, and through pptx-agent's in its ``tests/test_chart_cases.py``: the same
calls, the same results, but for the chart's name in a result (here ``id``)."""

from __future__ import annotations

from pathlib import Path

import pytest

from ooxml_edit.charts.cases import CHART, CHART_CASES, SMARTART_CASES, problems
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

SMARTART = Path(__file__).parent / "fixtures" / "generated" / "charts" / "smartart.docx"


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _document() -> tuple[bytes, str]:
    document = Document.new()
    document.insert_markdown("Figures.\n\nMore.", at="end")
    return document.to_bytes(), document.paragraphs()[0].id


@pytest.mark.parametrize("case", CHART_CASES, ids=[case["name"] for case in CHART_CASES])
def test_a_chart_case(box, case):
    session = box.session()
    data, anchor = _document()
    session.open(data, "chart.docx")

    def call(**arguments):
        return box.dispatch(session, "edit_chart", {"doc": "d1", **arguments})

    made = call(target=anchor, action="add", ref="c", **CHART)
    assert made.ok, made.to_json()
    for arguments in case["before"]:
        assert call(target="$c", **arguments).ok
    assert problems(case, call(target="$c", **case["call"]), key="id") == []


@pytest.mark.parametrize("case", SMARTART_CASES, ids=[case["name"] for case in SMARTART_CASES])
def test_a_smartart_case(box, case):
    session = box.session()
    session.open(SMARTART.read_bytes(), "smartart.docx")
    target = session.entry("d1").document.diagrams()[0].id
    result = box.dispatch(session, "edit_smartart", {"doc": "d1", "target": target, **case["call"]})
    assert problems(case, result) == []
