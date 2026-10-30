"""word_design_facts' ``declared_styles`` (production round 21, T7): every declared style,
used or not, resolved, theme colours as the RGB Word draws (``tools/theme_tint_probe.py``,
``tests/observations/theme-tints-word.json``)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_agent import Document
from docx_agent.edit.design import theme_rgb

ROOT = Path(__file__).resolve().parent.parent
SWATCHES = json.loads((ROOT / "tests" / "observations" / "theme-tints-word.json").read_text())["swatches"]


@pytest.mark.parametrize("swatch", SWATCHES, ids=lambda s: f"{s['theme']}-{s['tint']}-{s['shade']}")
def test_a_theme_colour_with_its_tint_or_shade_is_what_word_draws(swatch):
    document = Document.new()
    ours = theme_rgb(document, swatch["theme"], swatch["tint"], swatch["shade"])
    assert swatch["word"] == swatch["fill"]
    assert ours == swatch["word"]


def made() -> Document:
    document = Document.new()
    document.styles.add("Chapter title", based_on="heading 1", bold=True)
    document.styles.add("Quiet note", based_on="Normal", italic=True)
    document.styles.add("Marker", kind="character", bold=True)
    document.paragraphs()[0].set_text("Hello")
    return document


def test_declared_styles_lists_used_and_unused_styles_resolved():
    document = made()
    facts = document.design_facts(["declared_styles"])
    styles = {s["name"]: s for s in facts["declared_styles"]}
    assert styles["Normal"]["used"] and styles["Normal"]["default"]
    assert set(styles["Normal"]["look"]) >= {"font", "size", "color", "space_after"}
    chapter = styles["Chapter title"]
    assert chapter["used"] is False and chapter["based_on"] == ["heading 1", "Normal"]
    # Its colour is the theme's accent1 shaded BF, as Word draws it (the heading's w:val agrees).
    assert chapter["look"]["color"] == "0F4761" and chapter["look"]["theme_color"] == "accent1 shade BF"
    assert chapter["look"]["outline_level"] == 1 and chapter["look"]["bold"] is True
    assert styles["Quiet note"]["look"] == {"italic": True}
    assert styles["Marker"]["type"] == "character" and styles["Marker"]["look"] == {"bold": True}
    # Used styles come first.
    used = [s["used"] for s in facts["declared_styles"]]
    assert used == sorted(used, reverse=True)
    # Not in the default kinds.
    assert "declared_styles" not in document.design_facts()


#: A made-up table style: accent1 borders, a header row filled accent1 with white bold text,
#: row bands a tint of accent1, the fills' and colours' w:val decoys.
TABLE_STYLE = (
    '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="table" '
    'w:styleId="Ledger"><w:name w:val="Ledger"/><w:basedOn w:val="TableNormal"/><w:tblPr>'
    '<w:tblStyleRowBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" '
    'w:color="FF00FF" w:themeColor="accent1"/><w:bottom w:val="single" w:sz="4" w:space="0" '
    'w:color="FF00FF" w:themeColor="accent1"/></w:tblBorders></w:tblPr>'
    '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="000000" w:themeColor="background1"/></w:rPr>'
    '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00" w:themeFill="accent1"/></w:tcPr></w:tblStylePr>'
    '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00FF00" '
    'w:themeFill="accent1" w:themeFillTint="33"/></w:tcPr></w:tblStylePr></w:style>')


def test_a_table_style_lists_its_parts_and_places_with_theme_colours():
    from lxml import etree

    document = made()
    document.styles._ensure("Table Grid", "table")      # and Normal Table, its base
    document.styles._root().append(etree.fromstring(TABLE_STYLE))
    document._invalidate()
    styles = {s["name"]: s for s in document.design_facts(["declared_styles"])["declared_styles"]}
    ledger = styles["Ledger"]
    assert ledger["type"] == "table" and ledger["used"] is False and ledger["parts"] == ["firstRow", "band1Horz"]
    header, band = ledger["places"]["header_row"], ledger["places"]["band1"]
    # The fills and colours are the theme's, not the decoys beside them.
    assert (header["fill"], header["theme_fill"]) == ("156082", "accent1")
    assert header["text"] == {"color": "FFFFFF", "theme_color": "background1", "bold": True}
    assert (band["fill"], band["theme_fill"]) == ("C1E4F5", "accent1 tint 33")
    assert "single 0.5pt 156082 (accent1)" in header["borders"].values()


def test_the_tool_pages_declared_styles_by_size():
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS
    from docx_agent.tools._base import tool_names

    path = ROOT / "tests" / "fixtures" / "wordto" / "sample-10pages.docx"
    every = Document.open(path.read_bytes()).design_facts(["declared_styles"])["declared_styles"]
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session()
        d1 = session.open(path.read_bytes(), "template.docx")
        seen, cursor, pages = [], None, 0
        while True:
            arguments = {"doc": d1, "include": ["declared_styles", "theme"]}
            if cursor:
                arguments["cursor"] = cursor
            result = toolbox.dispatch(session, "word_design_facts", arguments)
            assert result.ok, result.error and result.error.to_json()
            assert len(json.dumps(result.data)) <= toolbox.limits.max_result_chars - 1500
            assert ("theme" in result.data) == (cursor is None)     # later pages: the styles only
            seen += result.data["declared_styles"]
            pages += 1
            cursor = result.next_cursor
            if cursor is None:
                break
    assert seen == tool_names(every, "look") and pages > 1     # keep_lines, as the tools say it
    assert result.summary.endswith(f"of {len(every)}")
