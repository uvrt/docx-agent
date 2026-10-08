"""Charts (ROADMAP.md, "Charts and SmartArt"): data, titles and legend, the embedded workbook
kept in step, wherever Word holds a chart -- the body, a header, a footnote, a text box, a
group, floating.

Every edit runs on every chart of the Word-saved fixtures (``tests/fixtures/generated/
charts``, ``tools/make_chart_fixtures.py``) and goes through every gate:

* **round trip** -- edit, save, reopen, read back;
* **the workbook agrees** -- the saved document's embedded workbook opened independently
  (``tests/xlsx.py``, not ooxml-edit's reader): every formula's cells hold what its cache
  holds, tables cover the data and are named after their headers, shared-string counts are
  right;
* **validity** -- docx-agent's checks, schema order in the chart part among them, add
  nothing;
* **undo and redo** -- back to the original bytes, workbook included, and on to the edited
  ones, byte for byte;
* **render** -- docx2svg draws the edited chart (its new label among the page's text) and
  warns of nothing new.
"""

from __future__ import annotations

import io
import re
import warnings
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.edit.charts import ChartDataError, ChartDataWarning, ChartEditError, UntrackedChartEdit
from docx_agent.markdown import check as markdown_check
from docx_agent.validate import check

from test_roundtrip import entries
from xlsx import check_chart_against_workbook

CHARTS = Path(__file__).parent / "fixtures" / "generated" / "charts"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def chart_fixtures() -> list[Path]:
    return sorted(CHARTS.glob("*.docx"))


def _cases() -> list[tuple[str, str]]:
    out = []
    for path in chart_fixtures():
        for chart in Document.open(path).charts():
            out.append((path.name, chart.id))
    return out


CASES = _cases()


def check_document_charts(data: bytes, original: bytes | None = None) -> int:
    """Every chart in a saved document agrees with its embedded workbook; returns the
    formulas checked."""
    document = Document.open(data)
    before = entries(original) if original is not None else {}
    checked = 0
    for chart in document.charts():
        workbook = chart.workbook_part
        content = None if workbook is None else document.package.read(workbook)
        checked += check_chart_against_workbook(document.package.read(chart.part), content,
                                                edited=workbook is None or before.get(workbook) != content)
    return checked


def test_the_fixtures_have_every_kind_and_place():
    kinds = {}
    for path in chart_fixtures():
        for chart in Document.open(path).charts():
            kinds.setdefault("+".join(chart.chart_types), []).append(chart.id)
    assert {"bar", "line", "pie", "doughnut", "scatter", "bar+line"} <= set(kinds)
    places = Document.open(CHARTS / "chart-places.docx")
    where = {chart.id: chart._resolve().part for chart in places.charts()}
    assert where == {"d:4": "word/document.xml", "d:5/7": "word/document.xml", "d:8": "word/document.xml",
                     "d:1": "word/header1.xml", "d:2": "word/footnotes.xml"}


def test_every_original_chart_agrees_with_its_workbook():
    """The gate's own baseline: Word's fixtures are consistent before any edit."""
    for path in chart_fixtures():
        data = path.read_bytes()
        assert check_document_charts(data, data) >= 0


def test_reading():
    document = Document.open(CHARTS / "charts.docx")
    chart = document.chart("d:1")
    assert chart.chart_types == ["bar"] and chart.title == "Sales by region"
    assert chart.categories == ["Q1", "Q2", "Q3", "Q4"]
    assert [s.name for s in chart.series] == ["North", "South", "West"]
    assert chart.series[0].values == [4.3, 2.5, 3.5, 4.5]
    assert chart.axes == ["category", "value"] and chart.has_legend
    assert chart.workbook_part.startswith("word/embeddings/") and chart.workbook_part.endswith(".xlsx")
    assert document.chart("d:5").categories == [0.7, 1.8, 2.6, 3.1, 4.4]
    assert document.chart("d:6").chart_types == ["bar", "line"]
    assert document.drawing("d:1").chart.id == "d:1"
    assert document.get("d:1").kind == "chart"


def test_a_group_member_is_addressed_by_its_group_and_its_own_id():
    document = Document.open(CHARTS / "chart-places.docx")
    group = document.drawing("d:5")
    assert group.kind == "group"
    assert [(m.id, m.kind) for m in group.members] == [("d:5/6", "shape"), ("d:5/7", "chart")]
    member = document.get("d:5/7")
    assert member.kind == "chart" and member.name == "Chart 7"
    assert member.chart.series[1].name == "Paired"
    with pytest.raises(KeyError):
        document.drawing("d:5/99")
    with pytest.raises(EditError):
        document.chart("d:5")  # the group is not a chart
    with pytest.raises(EditError):
        document.chart("d:5/6")


def test_a_chart_in_a_nested_group_is_addressed_down_the_groups():
    from lxml import etree

    document = Document.open(CHARTS / "chart-places.docx")
    WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    frame = next(document.package.tree("word/document.xml").iter(f"{{{WPG}}}graphicFrame"))
    nested = etree.SubElement(frame.getparent(), f"{{{WPG}}}grpSp")
    etree.SubElement(nested, f"{{{WPG}}}cNvPr", id="9", name="Group 9")
    nested.append(frame)
    document.package.mark_dirty("word/document.xml")
    reopened = Document.open(document.to_bytes())
    assert [m.id for m in reopened.drawing("d:5").members] == ["d:5/6", "d:5/9", "d:5/9/7"]
    assert reopened.get("d:5/9").kind == "group"
    chart = reopened.chart("d:5/9/7")
    chart.series[0].set_value(0, 99)
    assert Document.open(reopened.to_bytes()).chart("d:5/9/7").series[0].values[0] == 99
    assert [c.id for c in reopened.charts()] == ["d:4", "d:5/9/7", "d:8", "d:1", "d:2"]


def test_a_drawing_that_is_not_a_chart_is_refused():
    document = Document.open(CHARTS / "smartart.docx")
    with pytest.raises(EditError):
        document.chart("d:1")
    with pytest.raises(EditError):
        Document.open(CHARTS / "charts.docx").diagram("d:1")


# ------------------------------------------------------------------------------------------
# Each edit, through every gate
# ------------------------------------------------------------------------------------------


def _numeric(chart) -> bool:
    return bool(chart.categories) and all(not isinstance(c, str) for c in chart.categories)


def _edit_value(chart):
    index = chart.point_count - 1
    chart.series[-1].set_value(index, 1234.5)
    return lambda c: c.series[-1].values[index] == 1234.5, "1234.5"


def _edit_blank(chart):
    chart.series[0].set_value(1, None)
    return lambda c: c.series[0].values[1] is None, None


def _edit_values(chart):
    wanted = [float(k) + 0.5 for k in range(chart.point_count)]
    chart.series[0].set_values(wanted)
    return lambda c: c.series[0].values == wanted, None


def _edit_category(chart):
    label = 0.25 if _numeric(chart) else "Relabelled"
    chart.set_category(0, label)
    return lambda c: c.categories[0] == label, None if _numeric(chart) else "Relabelled"


def _edit_series_name(chart):
    chart.series[-1].name = "Renamed series"
    return lambda c: c.series[-1].name == "Renamed series", "Renamed series"


def _edit_add_category(chart):
    label = 9.75 if _numeric(chart) else "Added"
    count = chart.point_count
    chart.add_category(label, [7] * len(chart.series))
    return lambda c: c.point_count == count + 1 and c.categories[-1] == label \
        and all(s.values[-1] == 7 for s in c.series), None if _numeric(chart) else "Added"


def _edit_add_first_category(chart):
    label = 0.125 if _numeric(chart) else "First"
    chart.add_category(label, {chart.series[0].name: 3}, index=0)
    return lambda c: c.categories[0] == label and c.series[0].values[0] == 3, None if _numeric(chart) else "First"


def _edit_remove_category(chart):
    second = chart.categories[1]
    chart.remove_category(0)
    return lambda c: c.categories[0] == second, None


def _edit_add_series(chart):
    count = len(chart.series)
    chart.add_series("Added series", [2] * chart.point_count)
    # A pie's or doughnut's legend names its categories, not its series.
    shown = None if chart.chart_types[0] in ("pie", "doughnut") else "Added series"
    return lambda c: len(c.series) == count + 1 and c.series[-1].name == "Added series", shown


def _edit_add_first_series(chart):
    chart.add_series("Front series", [1] * chart.point_count, index=0)
    return lambda c: c.series[0].name == "Front series", "Front series"


def _edit_remove_series(chart):
    if len(chart.series) < 2:
        chart.add_series("Doomed", [1] * chart.point_count)
    names = [s.name for s in chart.series]
    chart.remove_series(0)
    return lambda c: [s.name for s in c.series] == names[1:], None


def _edit_title(chart):
    chart.set_title("A new title")
    return lambda c: c.title == "A new title", "A new title"


def _edit_axis_title(chart):
    if "value" not in chart.axes:
        with pytest.raises(KeyError):
            chart.set_axis_title("value", "Units")
        chart.set_title("No axes")
        return lambda c: c.title == "No axes", "No axes"
    chart.set_axis_title("value", "Units")
    return lambda c: c.axis_title("value") == "Units", "Units"


def _edit_legend(chart):
    shown = chart.has_legend
    chart.set_legend(not shown)
    return lambda c: c.has_legend is not shown, None


def _edit_model(chart):
    model = chart.model
    model["series"][0]["values"][0] = 4242
    model["categories"][-1] = 8.5 if _numeric(chart) else "Through JSON"
    chart.apply(model)
    return lambda c: c.series[0].values[0] == 4242, None if _numeric(chart) else "Through JSON"


EDITS = [_edit_value, _edit_blank, _edit_values, _edit_category, _edit_series_name, _edit_add_category,
         _edit_add_first_category, _edit_remove_category, _edit_add_series, _edit_add_first_series,
         _edit_remove_series, _edit_title, _edit_axis_title, _edit_legend, _edit_model]


def _svg_text(svgs: list[str]) -> str:
    out = []
    for svg in svgs:
        root = etree.fromstring(svg.encode())
        out += [" ".join("".join(n.itertext()).split()) for n in root.iter("{http://www.w3.org/2000/svg}text")]
    return "\n".join(out)


@pytest.mark.parametrize("edit", EDITS, ids=lambda f: f.__name__[6:])
@pytest.mark.parametrize("fixture,chart_id", CASES, ids=[f"{f[:-5]}-{c}" for f, c in CASES])
def test_chart_edit(fixture, chart_id, edit):
    data = (CHARTS / fixture).read_bytes()
    document = Document.open(data)
    before = set(check(document.package))
    with warnings.catch_warnings():
        warnings.simplefilter("error")  # every fixture chart has its workbook: no cache-only edit
        read_back, shown = edit(document.chart(chart_id))
    edited = document.to_bytes()
    # One undo step per call (removing a pie's only series takes a series added first).
    steps = len(document.history._undo)
    assert steps == 1 or (edit is _edit_remove_series and steps == 2), steps

    saved = Document.open(edited)
    assert read_back(saved.chart(chart_id)), f"{edit.__name__} did not read back"
    assert check_document_charts(edited, data) > 0
    assert set(check(saved.package)) - before == set()
    changed = {name for name, content in entries(edited).items() if entries(data).get(name) != content}
    assert changed and all(n.startswith(("word/charts/", "word/embeddings/")) for n in changed), changed

    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)
    while document.redo():
        pass
    assert entries(document.to_bytes()) == entries(edited)


#: Charts docx2svg does not draw. None since docx2svg draws a chart in a footnote (related
#: from the notes part) and one inline in a text box.
NOT_DRAWN: set = set()


def _numbers(text: str) -> list[float]:
    return [float(t) for t in re.findall(r"^-?\d+(?:\.\d+)?$", text, re.M)]


@pytest.mark.parametrize("edit", [_edit_value, _edit_category, _edit_add_series, _edit_title],
                         ids=lambda f: f.__name__[6:])
@pytest.mark.parametrize("fixture,chart_id", CASES, ids=[f"{f[:-5]}-{c}" for f, c in CASES])
def test_chart_edit_renders(fixture, chart_id, edit):
    """docx2svg draws the edited chart from its caches: a new label is on the page, a value
    axis reaches a new value, and nothing is warned that the unedited document was not."""
    document = Document.open((CHARTS / fixture).read_bytes())
    warned = set(document.layout().warnings)
    chart = document.chart(chart_id)
    axes = "value" in chart.axes
    _, shown = edit(chart)
    assert set(document.layout().warnings) <= warned
    text = _svg_text(document.render_svg())
    if (fixture, chart_id) in NOT_DRAWN:
        assert shown is None or shown not in text  # pinned: the day docx2svg draws it, say so
    elif edit is _edit_value:
        assert not axes or max(_numbers(text)) >= 1234.5, (chart_id, sorted(_numbers(text))[-3:])
    elif shown is not None:
        assert shown in text, (shown, chart_id)


# ------------------------------------------------------------------------------------------
# Refusals, cache-only charts, tracking
# ------------------------------------------------------------------------------------------


def test_a_refused_edit_changes_nothing():
    document = Document.open(CHARTS / "charts.docx")
    data = document.to_bytes()
    chart = document.chart("d:1")
    with pytest.raises(ChartDataError):
        chart.series[0].set_value(0, "not a number")
    with pytest.raises(IndexError):
        chart.series[0].set_value(99, 1)
    with pytest.raises(ChartEditError):
        chart.apply({**chart.model, "types": ["line"]})
    with pytest.raises(ChartEditError):
        chart.apply({"title": 3})
    assert document.to_bytes() == data and not document.history.can_undo()


def _linked(data: bytes, chart_part: str) -> bytes:
    """``data`` with the chart's workbook linked from outside the document."""
    rels = chart_part.replace("charts/", "charts/_rels/") + ".rels"
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for info in source.infolist():
            content = source.read(info)
            if info.filename == rels:
                content = re.sub(rb'Target="[^"]+\.xlsx"', b'Target="https://example.org/book.xlsx" '
                                 b'TargetMode="External"', content)
            out.writestr(info, content)
    return buffer.getvalue()


def test_a_chart_whose_workbook_is_linked_is_edited_in_its_cache_with_a_warning():
    data = (CHARTS / "charts.docx").read_bytes()
    document = Document.open(_linked(data, Document.open(data).chart("d:2").part))
    chart = document.chart("d:2")
    assert chart.workbook_part is None
    with pytest.warns(ChartDataWarning, match="Word's Edit Data will show the old values"):
        chart.series[0].set_value(0, 5)
    assert Document.open(document.to_bytes()).chart("d:2").series[0].values[0] == 5


def test_tracked_chart_and_diagram_edits_are_applied_untracked_with_a_warning():
    """Word has no revision for a chart's data or a diagram's text (measured: its review
    counts none): tracked, the edit is applied, untracked, and a warning says so."""
    document = Document.open(CHARTS / "charts.docx")
    with document.tracking(author="Agent"):
        with pytest.warns(UntrackedChartEdit):
            document.chart("d:1").series[0].set_value(0, 9)
    edited = document.to_bytes()
    assert Document.open(edited).chart("d:1").series[0].values[0] == 9
    assert Document.open(edited).revisions() == []
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        document.chart("d:1").series[0].set_value(0, 10)  # untracked: no warning
    diagram_doc = Document.open(CHARTS / "smartart.docx")
    with diagram_doc.tracking(author="Agent"):
        with pytest.warns(UntrackedChartEdit):
            diagram_doc.diagram("d:1").set_text(0, "Planned")
    assert Document.open(diagram_doc.to_bytes()).revisions() == []


# ------------------------------------------------------------------------------------------
# Reading: to_markdown and state
# ------------------------------------------------------------------------------------------


def test_to_markdown_summarises_each_chart_in_its_drawings_comment():
    document = Document.open(CHARTS / "charts.docx")
    data = document.to_bytes()
    text = document.to_markdown()
    assert ('d:1 chart "Chart 1" [chart: bar; title: Sales by region; series: North, South, West; '
            "categories: Q1, Q2, Q3, Q4]") in text
    assert "[chart: bar + line; series: Revenue, Cost, Margin; categories: Q1, Q2, Q3, Q4]" in text
    assert document.to_bytes() == data
    places = Document.open(CHARTS / "chart-places.docx").to_markdown()
    assert '<!-- d:5 group "Group 5" [d:5/7 chart: bar; series: Grouped, Paired; categories: A, B, C] -->' in places
    assert "[chart: bar; series: Noted; categories: A, B, C]" in places  # in the footnote


def test_to_markdown_lists_a_diagrams_node_text():
    text = Document.open(CHARTS / "smartart.docx").to_markdown()
    assert "[smartart: Plan, Write, Measure, Review, Ship]" in text
    assert "[smartart: Goals [Faster edits, Fewer prompts], Risks [Stale caches, Lost formulas]]" in text


def test_to_markdown_reads_back_as_markdown():
    """The summaries are inside comments: markdown-it-py reads the projection back to the
    model it was written from (E2's reverse-parse check), a title full of ``-->`` too."""
    document = Document.open(CHARTS / "charts.docx")
    document.chart("d:1").set_title("Sales --> up <!-- and -->")
    assert "Sales - -&gt; up" in document.to_markdown()
    assert markdown_check(document) == []


def test_state_has_each_charts_and_diagrams_json():
    state = Document.open(CHARTS / "charts.docx").state()
    drawings = [d for block in state["blocks"] for d in block.get("drawings", [])]
    first = next(d for d in drawings if d["id"] == "d:1")
    assert first["chart"]["series"][0] == {"name": "North", "values": [4.3, 2.5, 3.5, 4.5]}
    assert first["chart"]["types"] == ["bar"] and first["chart"]["title"] == "Sales by region"
    places = Document.open(CHARTS / "chart-places.docx").state()
    group = next(d for block in places["blocks"] for d in block.get("drawings", []) if d["id"] == "d:5")
    assert [m["id"] for m in group["members"]] == ["d:5/6", "d:5/7"]
    assert group["members"][1]["chart"]["series"][1]["name"] == "Paired"
    smart = Document.open(CHARTS / "smartart.docx").state()
    diagram = next(d for block in smart["blocks"] for d in block.get("drawings", []) if d["id"] == "d:2")
    assert diagram["diagram"]["layout"].endswith("/vList2")
    assert [(n["lvl"], n["t"]) for n in diagram["diagram"]["nodes"]][:3] == [(0, "Goals"), (1, "Faster edits"),
                                                                          (1, "Fewer prompts")]


def test_the_state_json_applies_back():
    document = Document.open(CHARTS / "charts.docx")
    state = document.state()
    model = next(d for block in state["blocks"] for d in block.get("drawings", []) if d["id"] == "d:2")["chart"]
    model["series"][1]["name"] = "Members"
    model["categories"].append("Jul")
    model["series"][0]["values"].append(230)
    model["series"][1]["values"].append(None)
    document.chart("d:2").apply(model)
    assert document.chart("d:2").model == model
    assert check_document_charts(document.to_bytes()) > 0
    assert document.undo() and not document.history.can_undo()


# ------------------------------------------------------------------------------------------
# What a new title is written as, and in which language
# ------------------------------------------------------------------------------------------


def test_a_new_title_is_written_as_offices_chart_style_writes_one():
    """Word draws this form as its own titles look -- Aptos 14 pt, not bold, 595959 -- and
    saves it back unchanged; ooxml-edit's empty ``a:defRPr`` Word draws bold at 18 pt
    (``tools/charts_probe.py``, ``titles``)."""
    document = Document.open(CHARTS / "charts.docx")
    chart = document.chart("d:2")
    chart.set_title("Visitors and signups")
    chart.set_axis_title("value", "People")
    root = document.package.tree(chart.part)
    title = root.find(f"{C}chart/{C}title")
    defaults = title.find(f".//{A}defRPr")
    assert (defaults.get("sz"), defaults.get("b")) == ("1400", "0")
    assert defaults.find(f"{A}latin").get("typeface") == "+mn-lt"
    assert [n.get("val") for n in defaults.iter(f"{A}lumMod", f"{A}lumOff")] == ["65000", "35000"]
    assert title.find(f".//{A}r/{A}rPr").get("lang") == "en-US"
    axis = root.find(f".//{C}valAx/{C}title")
    assert axis.find(f".//{A}bodyPr").get("rot") == "-5400000"
    assert axis.find(f".//{A}defRPr").get("sz") == "1000"


def test_new_chart_text_is_in_the_documents_default_language():
    from docx_agent.edit.charts import default_language

    document = Document.open(CHARTS / "charts.docx")
    assert default_language(document) == "en-US"
    styles = document.package.tree(document.package.styles_part())
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}lang").set(f"{W}val", "nl-NL")
    document.chart("d:4").set_title("Plan en werkelijk")
    title = document.package.tree(document.chart("d:4").part).find(f"{C}chart/{C}title")
    assert title.find(f".//{A}r/{A}rPr").get("lang") == "nl-NL"
