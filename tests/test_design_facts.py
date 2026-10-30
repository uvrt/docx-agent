"""Design facts (round 16, F9): ``Document.design_facts`` and ``word_design_facts`` on a
made-up template -- headings 1-3, body, a list, captions, a table style of its own with a
header row, row bands and a first column, two sections -- and on every reading fixture.

The template is built here with the library; its table style is written by hand below
(nothing in it comes from Word or any other application)."""

from __future__ import annotations

import json

import pytest
from lxml import etree

from docx_agent import Document
from docx_agent.edit.design import KINDS
from docx_agent.tools._base import tool_names

#: A table style of this test's own: borders top, bottom and between rows; a dark header
#: row with white bold text, a light first column in bold, light odd row bands; 10 pt text.
TABLE_STYLE = (
    '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="table" '
    'w:styleId="ReportGrid"><w:name w:val="Report Grid"/><w:basedOn w:val="TableNormal"/>'
    '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr>'
    '<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="2E5C8A"/>'
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="2E5C8A"/>'
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="A6BFD9"/></w:tblBorders></w:tblPr>'
    '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>'
    '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2E5C8A"/></w:tcPr></w:tblStylePr>'
    '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr>'
    '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="DCE6F1"/></w:tcPr></w:tblStylePr>'
    '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F2F6FA"/></w:tcPr>'
    '</w:tblStylePr></w:style>')

STYLES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"


def template() -> Document:
    doc = Document.new(page="A4", title="Quarterly template")
    blocks = ["# Overview", "An opening paragraph in the body style.", "A second body paragraph.",
              "## Scope", "Text under a level-two heading.", "- first point\n- second point"]
    for k, area in enumerate(["North", "South", "East"]):
        blocks += [f"### Area {area}", "Short lead-in text for the area.",
                   f"| Item | Owner | Due |\n| --- | --- | --- |\n| Plan {k} | Team | Q{k + 1} |\n"
                   "| Review | Lead | Q4 |"]
    blocks += ["# Appendix", "Closing text."]
    doc.insert_markdown("\n\n".join(blocks))
    part = doc.package.related_parts_of_type(doc.package.document_part(), STYLES_REL)[0]
    doc.package.tree(part).append(etree.fromstring(TABLE_STYLE))
    doc = Document.open(doc.to_bytes())
    for table in doc.tables():
        doc.set_table(table.id, style="Report Grid",
                      look={"first_row": True, "first_column": True, "banded_rows": True})
        doc.insert_caption(before=table.id, label="Table", text=": Area figures")
    appendix = next(p for p in doc.paragraphs() if p.text == "Appendix")
    doc.insert_section_break(before=appendix.id)
    sections = doc.sections()
    doc.set_section(sections[-1].id, orientation="landscape", columns=2)
    doc.add_header(sections[0].id, "default", "Quarterly template")
    doc.add_header(sections[0].id, "first", "Cover header")
    doc.add_footer(sections[0].id, "default", "Footer text")
    return doc


@pytest.fixture(scope="module")
def doc() -> Document:
    return template()


@pytest.fixture(scope="module")
def facts(doc) -> dict:
    return doc.design_facts()


def _style(facts, role, name=None):
    found = [s for s in facts["styles"] if s["role"] == role and (name is None or s["style"] == name)]
    assert len(found) == 1, (role, name, facts["styles"])
    return found[0]


def test_every_kind_is_reported_in_order(facts):
    assert [k for k in facts if k in KINDS] == list(KINDS)
    assert "omitted" not in facts


def test_styles_per_role_with_usage_and_effective_look(doc, facts):
    roles = [(s["role"], s["style"], s["paragraphs"]) for s in facts["styles"]]
    assert roles == [("heading 1", "heading 1", 3), ("heading 2", "heading 2", 1), ("heading 3", "heading 3", 3),
                     ("body", "Normal", roles[3][2]), ("list", "List Bullet", 2), ("caption", "caption", 3),
                     ("table", "Normal", 27), ("header", "header", 2), ("footer", "footer", 1)]
    # The section break's paragraph, empty, took the heading style of the paragraph it was put before.
    assert _style(facts, "heading 1")["empty"] == 1 and "empty" not in _style(facts, "heading 3")
    heading = _style(facts, "heading 1")["look"]
    # The new document's Heading 1: the theme's major face, as docx2svg resolves it.
    assert heading["font"] == facts["theme"]["fonts"]["major"]["latin"]
    assert heading["keep_with_next"] is True and heading["outline_level"] == 1
    body = _style(facts, "body")
    assert body["look"]["font"] == facts["theme"]["fonts"]["minor"]["latin"]
    assert doc.paragraph(body["example"]).style_name == "Normal"
    listed = _style(facts, "list")["look"]
    assert listed["numbering"]["format"] == "bullet" and listed["indent_first"] < 0
    # Text in a cell takes the table style's 10 pt and no space after.
    table = _style(facts, "table")
    assert table["table_style"] == "Report Grid"
    assert table["look"]["size"] == 10 and table["look"]["space_after"] == 0
    assert _style(facts, "body")["look"]["size"] != 10


def test_a_style_s_look_is_docx2svg_s_for_a_paragraph_without_direct_formatting(doc, facts):
    example = doc.paragraph(_style(facts, "heading 3")["example"])
    effective = example.effective
    look = _style(facts, "heading 3")["look"]
    assert look["space_before"] == effective.space_before and look["space_after"] == effective.space_after
    run = example.runs[0].effective
    assert look["font"] == run.font and look["size"] == run.size and look["color"] == run.color


def test_a_table_style_by_place_after_conditional_formatting(doc, facts):
    [grid] = facts["table_styles"]
    assert grid["style"] == "Report Grid" and grid["tables"] == 3
    assert grid["examples"] == [t.id for t in doc.tables()]
    assert grid["options"] == {"header_row": True, "total_row": False, "first_column": True, "last_column": False,
                               "banded_rows": True, "banded_columns": False}
    places = grid["places"]
    assert set(places) == {"header_row", "band1", "band2", "first_column", "corner"}
    header = places["header_row"]
    assert header["conditions"] == ["firstRow"] and header["fill"] == "2E5C8A"
    assert header["text"]["bold"] is True and header["text"]["color"] == "FFFFFF" and header["text"]["size"] == 10
    assert header["borders"]["top"] == "single 0.5pt 2E5C8A"
    assert header["borders"]["bottom"] == "single 0.5pt A6BFD9" and header["borders"]["left"] == "none"
    assert places["band1"]["fill"] == "F2F6FA" and places["band2"]["fill"] is None
    assert "bold" not in places["band1"]["text"]
    label = places["first_column"]
    assert label["conditions"] == ["band1Horz", "firstCol"] and label["fill"] == "DCE6F1"
    assert label["text"]["bold"] is True
    # Where the header row meets the first column, the header row's formats win (they come later).
    assert places["corner"]["fill"] == "2E5C8A" and places["corner"]["text"]["bold"] is True


def test_the_page_grid_per_section_with_its_header_and_footer_variants(facts):
    first, second = facts["sections"]
    assert first["orientation"] == "portrait" and first["page"] == [595.3, 841.9]
    assert first["columns"]["count"] == 1 and first["title_page"] is True
    assert first["headers"] == {"default": "header1", "first": "header2"}
    assert first["footers"] == {"default": "footer1", "first": None}
    assert second["orientation"] == "landscape" and second["page"] == [841.9, 595.3]
    assert second["columns"]["count"] == 2 and "title_page" not in second
    assert second["headers"] == {"default": "header1"}
    assert second["linked_to_previous"] == ["header default", "footer default"]
    assert set(first["margins"]) >= {"top", "right", "bottom", "left", "header", "footer"}
    assert facts["even_and_odd_headers"] is False


def test_the_theme_palette_and_fonts(facts):
    theme = facts["theme"]
    assert list(theme["colors"])[:4] == ["dk1", "lt1", "dk2", "lt2"] and len(theme["colors"]) == 12
    assert all(len(v) == 6 for v in theme["colors"].values())
    assert theme["fonts"]["major"]["latin"] and theme["fonts"]["minor"]["latin"]


def test_a_repeated_shape_under_every_level_three_heading(doc, facts):
    [pattern] = facts["patterns"]
    assert pattern["heading_level"] == 3 and pattern["count"] == 3 and pattern["of"] == 3
    assert pattern["shape"] == ["Normal", "caption", "table Report Grid, 3 columns"]
    headings = [p.id for p in doc.paragraphs() if p.text.startswith("Area ")]
    assert [e["heading"] for e in pattern["examples"]] == headings
    tables = [t.id for t in doc.tables()]
    assert [e["blocks"] for e in pattern["examples"]] == [f"{h}..{t}" for h, t in zip(headings, tables)]
    # A span the reading tools take.
    assert "Area North" in doc.to_markdown(pattern["examples"][0]["blocks"])


def test_a_shape_under_one_heading_only_or_of_one_body_paragraph_is_no_pattern():
    doc = Document.new()
    doc.insert_markdown("# One\n\nText.\n\n# Two\n\nText.\n\n## Three\n\nText.\n\n- a\n- b")
    assert doc.design_facts(["patterns"]) == {"patterns": []}


def test_include_picks_kinds_and_caps_count_what_is_left_out(doc):
    assert list(doc.design_facts(["theme", "sections"])) == ["sections", "theme", "even_and_odd_headers"]
    capped = doc.design_facts(["styles"], caps={"styles": 3})
    assert len(capped["styles"]) == 3 and capped["omitted"] == {"styles": 6}
    with pytest.raises(ValueError, match="the kinds are styles"):
        doc.design_facts(["palette"])


def test_every_reading_fixture(reading_path):
    facts = Document.open(reading_path.read_bytes()).design_facts()
    json.dumps(facts)
    for entry in facts["styles"]:
        assert entry["paragraphs"] > 0 and "font" in entry["look"]
    for entry in facts["table_styles"]:
        assert entry["tables"] > 0


# -- the tool ----------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def box():
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def test_word_design_facts(box, doc):
    session = box.session()
    d1 = session.open(doc.to_bytes(), "template.docx")
    result = box.dispatch(session, "word_design_facts", {"doc": d1})
    assert result.ok, result.error and result.error.to_json()
    assert result.summary == "Design facts: 9 styles, 1 table styles, 2 sections, 1 patterns"
    # The Python API's paragraph keep_together is keep_lines to a model (round 25, W4).
    assert result.data == tool_names(doc.design_facts(), "look")
    assert '"keep_together"' not in json.dumps(result.data) and '"keep_lines"' in json.dumps(result.data)
    only = box.dispatch(session, "word_design_facts", {"doc": d1, "include": ["table_styles"]})
    assert list(only.data) == ["table_styles"]
    refused = box.dispatch(session, "word_design_facts", {"doc": d1, "include": ["palette"]})
    assert refused.error.code == "invalid_arguments" and "theme" in refused.error.valid_options
    assert box.tools["word_design_facts"].group == "word_style" and not box.tools["word_design_facts"].mutates


def test_word_design_facts_keeps_fewer_entries_under_a_small_result_limit(doc):
    from ooxml_edit.tools import Limits, Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, limits=Limits(max_result_chars=6000)) as toolbox:
        session = toolbox.session()
        d1 = session.open(doc.to_bytes(), "template.docx")
        result = toolbox.dispatch(session, "word_design_facts", {"doc": d1})
    assert result.ok and result.data["omitted"]["styles"] > 0
    assert len(json.dumps(result.data)) <= 6000 - 1500
    assert "left out (include fewer kinds)" in result.summary
