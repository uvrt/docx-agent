"""``check``'s ``column_fit`` (round 31, R6): a table column whose widest line uses well
under its width beside a column that wraps heavily, with the widths that let both fit.
Made-up label and value tables, ``Document.new()``'s package; the thresholds are
:data:`docx_agent.layout_facts.ROOMY_SHARE`, ``MIN_FREED``, ``WRAP_LINES``, ``FIT_SLACK``."""

from __future__ import annotations

import pytest
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.layout_facts import FIT_SLACK, layout_facts
from docx_agent.tools import FORMAT, GROUPS, TOOLS

VALUE = ("The courier collects the parcels from the side door each weekday morning, checks the labels "
         "against the list and signs the sheet that stays in the tray by the door. ") * 2


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def labelled(rows: list[list[str]], widths: list[float] | None = None) -> tuple[Document, str]:
    document = Document.new()
    first = document.paragraphs()[0].id
    table = document.insert_table(len(rows), len(rows[0]), after=first, data=rows, widths=widths).id
    return document, table


def lines_of(document: Document, table: str, column: int) -> list[int]:
    """How many lines each cell of ``column`` is drawn on."""
    layout = document.layout()
    counts = []
    for row in range(len(document.table(table).rows)):
        cell = document.table(table).cell(row, column)
        ids = {p.id for p in cell.paragraphs}
        counts.append(sum(1 for page in layout.pages for line in page.lines
                          if line.paragraph_id and f"p:{line.paragraph_id}" in ids))
    return counts


def test_a_label_column_beside_wrapping_values_is_a_fact_with_widths_that_fit(box):
    document, table = labelled([["Pickup", VALUE], ["Day", "Weekdays"], ["Contact", "The front desk"]])
    [entry] = layout_facts(document.layout())["column_fit"]
    roomy, wraps = entry["roomy"], entry["wraps"]
    assert entry["table"] == table and entry["pages"] == [1]
    assert roomy["column"] == 0 and wraps["column"] == 1 and roomy["width"] == wraps["width"]
    assert roomy["widest_line"] < 0.6 * roomy["width"] and roomy["longest_word"] <= roomy["widest_line"]
    assert wraps["lines"] >= 6 and wraps["longest_word"] < wraps["widest_line"]
    widths = entry["widths"]
    assert sum(widths) == pytest.approx(roomy["width"] + wraps["width"], abs=0.2)
    assert widths[0] > roomy["widest_line"] + FIT_SLACK            # its line, the slack and its margins
    assert entry["fix"] == {"tool": "word_format_table", "calls": [
        {"table": table, "scope": "column", "column": 0, "width": widths[0]},
        {"table": table, "scope": "column", "column": 1, "width": widths[1]}]}
    before = lines_of(document, table, 1)[0]
    session = box.session()
    d = session.open(document.to_bytes(), "fit.docx")
    for call in entry["fix"]["calls"]:
        result = box.dispatch(session, entry["fix"]["tool"], {"doc": d, **call})
        assert result.ok, result.error
    fitted = session.entry(d).document
    assert lines_of(fitted, table, 0) == [1, 1, 1]                  # the labels still on one line
    assert lines_of(fitted, table, 1)[0] < before                   # the value on fewer
    facts = box.dispatch(session, "check", {"doc": d, "include": ["layout"]}).data["layout"]
    assert facts["column_fit"] == []


def test_columns_that_both_wrap_or_fit_are_no_fact():
    document, _ = labelled([[VALUE, VALUE]])
    assert layout_facts(document.layout())["column_fit"] == []
    document, _ = labelled([["Pickup", "Weekdays"], ["Day", "Mornings"]])
    assert layout_facts(document.layout())["column_fit"] == []
    document, _ = labelled([[VALUE]])
    assert layout_facts(document.layout())["column_fit"] == []


def test_a_roomy_column_that_frees_little_is_no_fact():
    # The label column already narrow: under MIN_FREED to give.
    document, _ = labelled([["Pickup", VALUE]], widths=[60, 393.3])
    assert layout_facts(document.layout())["column_fit"] == []


def test_the_neighbour_that_wraps_most_gets_the_room():
    document, table = labelled([["A long first value that wraps " * 4, "Day", VALUE]])
    [entry] = layout_facts(document.layout())["column_fit"]
    assert entry["roomy"]["column"] == 1 and entry["wraps"]["column"] == 2
    assert entry["widths"][0] == pytest.approx(document.table_properties(table)["columns"][0], abs=0.1)
