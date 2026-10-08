"""Formatting, styles first: styles by name (localised templates too), built-in styles
written as Word writes them, direct run and paragraph formatting, clearing it, and the
effective values read through docx2svg's resolver."""

from __future__ import annotations

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.validate import check

from test_roundtrip import entries

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@pytest.fixture
def lists_doc(markup_doc):
    return markup_doc.parent / "lists-and-styles.docx"


def xml(paragraph) -> str:
    return etree.tostring(paragraph._element).decode()


# -- styles -----------------------------------------------------------------------------------


def test_styles_are_found_by_name_in_a_localised_template(lists_doc):
    document = Document.open(lists_doc)
    styles = document.styles
    assert styles.get("Heading 1").id == "Kop1"           # w:name "heading 1", any case
    assert styles.get("normal").id == "Standaard"
    assert styles.get("Strong", "character").id == "Zwaar"
    assert styles.get("Kop1").name == "heading 1"         # the id still works
    assert styles.default("paragraph").id == "Standaard"
    paragraph = document.paragraph("p:5A00000E")
    paragraph.style = "heading 1"
    assert paragraph.style == "Kop1" and paragraph.style_name == "heading 1"
    paragraph.style = "Normal"                            # the default: no w:pStyle, as Word writes it
    assert paragraph.style is None and "pStyle" not in xml(paragraph)


def test_a_character_style_by_name_on_a_range(lists_doc):
    document = Document.open(lists_doc)
    document.anchor("Between").set_style("emphasis")
    run = next(r for r in document.paragraph("p:5A000006").runs if r.text == "Between")
    assert run.style == "Nadruk"
    assert [r.text for r in document.paragraph("p:5A000006").runs] == ["Between", " the lists."]
    document.anchor("Between").set_style(None)
    assert [r.text for r in document.paragraph("p:5A000006").runs] == ["Between the lists."]


def test_a_style_of_the_wrong_kind_or_none_is_refused(lists_doc):
    document = Document.open(lists_doc)
    with pytest.raises(EditError, match="character style"):
        document.paragraph("p:5A000006").set_style("Strong")
    with pytest.raises(EditError, match="no paragraph style"):
        document.paragraph("p:5A000006").set_style("No Such Style")


def test_a_built_in_style_is_written_as_word_writes_it(lists_doc):
    data = lists_doc.read_bytes()
    document = Document.open(data)
    document.paragraph("p:5A000006").style = "Heading 2"
    styles = document.styles
    heading = styles.get("heading 2")
    assert heading.id == "Heading2" and heading.based_on == "Standaard" and heading.link == "Heading2Char"
    linked = styles.get("Heading 2 Char", "character")
    assert linked.link == "Heading2" and linked.based_on == "Standaardalinea-lettertype"
    node = next(n for n in document.package.tree("word/styles.xml").iter(W + "style")
                if n.get(W + "styleId") == "Heading2")
    assert node.find(f"{W}next").get(W + "val") == "Standaard"
    assert node.find(f"{W}rPr/{W}color").get(W + "themeColor") == "accent1"
    assert node.find(W + "rsid") is None                  # no rsids written (decision 6)
    effective = document.paragraph("p:5A000006").effective
    assert effective.keep_with_next and effective.outline_level == 1
    assert check(Document.open(document.to_bytes()).package) == []
    document.undo()
    assert entries(document.to_bytes()) == entries(data)


def test_a_built_in_list_style_brings_its_list(lists_doc):
    document = Document.open(lists_doc)
    document.paragraph("p:5A000006").style = "List Number"
    found = document.paragraph("p:5A000006").list
    assert found.from_style and found.format == "decimal"
    assert check(Document.open(document.to_bytes()).package) == []


def test_add_and_modify_a_style(lists_doc):
    document = Document.open(lists_doc)
    result = document.styles.add("Callout", based_on="Normal", bold=True, color="1F4E79", space_after=6)
    assert result.id == "style:Callout"
    document.paragraph("p:5A000006").style = "Callout"
    effective = document.paragraph("p:5A000006").effective
    assert effective.space_after == 6
    assert document.paragraph("p:5A000006").runs[0].effective.bold
    document.styles.modify("Callout", bold=None, italic=True)
    run = document.paragraph("p:5A000006").runs[0].effective
    assert not run.bold and run.italic
    with pytest.raises(EditError, match="already"):
        document.styles.add("callout")
    assert check(Document.open(document.to_bytes()).package) == []


# -- direct formatting ------------------------------------------------------------------------


def test_format_a_range_splits_and_merges_runs(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    paragraph = document.paragraph("p:1A2B3C02")
    document.anchor("ain bo").format(bold=True, size=14, color="C00000", underline=True, highlight="yellow")
    # "bo" was bold already: it comes out like "ain ", and the two are one run.
    assert [r.text for r in paragraph.runs] == ["Pl", "ain bo", "ld", " and ", "italic", " runs."]
    run = paragraph.runs[1]
    assert (run.bold, run.size, run.color, run.underline, run.highlight) == (True, 14, "C00000", True, "yellow")
    assert paragraph.runs[2].bold and paragraph.runs[2].size is None
    document.anchor("ain").format(bold=None, size=None, color=None, underline=None, highlight=None)
    document.anchor(" bo", within="p:1A2B3C02").format(bold=None, size=None, color=None, underline=None, highlight=None)
    assert [r.text for r in paragraph.runs] == ["Plain bo", "ld", " and ", "italic", " runs."]
    assert check(Document.open(document.to_bytes()).package) == []
    document.undo(), document.undo(), document.undo()
    assert entries(document.to_bytes()) == entries(data)


def test_format_values_are_checked_before_anything_changes(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    for values in ({"bold": "yes"}, {"size": 0}, {"color": "red"}, {"highlight": "orange"},
                   {"underline": "squiggle"}, {"font": ""}, {"shadow": True}):
        with pytest.raises(EditError):
            document.anchor("bold").format(**values)
    for values in ({"alignment": "middle"}, {"space_after": -1}, {"first_line": 10, "hanging": 5},
                   {"line_spacing": ("double", 2)}):
        with pytest.raises(EditError):
            document.paragraph("p:1A2B3C02").format(**values)
    assert entries(document.to_bytes()) == entries(data)


def test_theme_colours_and_fonts(lists_doc):
    document = Document.open(lists_doc)
    document.anchor("Between").format(color="accent2", font="Georgia")
    run = document.paragraph("p:5A000006").runs[0]
    assert run.color == "accent2" and run.font == "Georgia"
    node = run._element.find(f"{W}rPr/{W}color")
    assert node.get(W + "themeColor") == "accent2" and node.get(W + "val")
    assert run.effective.font == "Georgia"


def test_paragraph_formatting(markup_doc):
    document = Document.open(markup_doc)
    paragraph = document.paragraph("p:1A2B3C02")
    paragraph.format(alignment="justify", indent_left=36, first_line=18, space_before=6, space_after=12,
                     line_spacing=1.5, keep_with_next=True, keep_together=True, page_break_before=False)
    assert paragraph.alignment == "justify" and paragraph.declared("indent_left") == 36
    assert paragraph.first_line == 18 and paragraph.line_spacing == 1.5
    effective = paragraph.effective
    assert (effective.alignment, effective.indent_left, effective.first_line) == ("justify", 36, 18)
    assert (effective.space_before, effective.space_after, effective.line_spacing) == (6, 12, 1.5)
    assert effective.keep_with_next and effective.keep_together and not effective.page_break_before
    paragraph.hanging = 9                                 # a hanging indent replaces the first-line one
    assert paragraph.first_line is None and paragraph.hanging == 9
    paragraph.line_spacing = ("exact", 14)
    assert paragraph.line_spacing == ("exact", 14) and paragraph.effective.line_rule == "exact"
    assert check(Document.open(document.to_bytes()).package) == []


def test_paragraph_run_formatting_reaches_every_run_and_the_mark(markup_doc):
    document = Document.open(markup_doc)
    paragraph = document.paragraph("p:1A2B3C02")
    paragraph.format(italic=True)
    assert all(run.italic for run in paragraph.runs)
    mark = paragraph._element.find(f"{W}pPr/{W}rPr/{W}i")
    assert mark is not None


def test_run_properties_set_directly(markup_doc):
    document = Document.open(markup_doc)
    run = document.paragraph("p:1A2B3C02").run(0)
    run.bold = False
    run.small_caps = True
    assert run.bold is False and run.small_caps is True
    assert '<w:b w:val="0"/>' in etree.tostring(run._element).decode()
    assert not run.effective.bold and run.effective.small_caps


def test_clear_direct_formatting(lists_doc):
    document = Document.open(lists_doc)
    paragraph = document.paragraph("p:5A000002")
    paragraph.format(alignment="center", space_after=3)
    paragraph.clear_direct_formatting()
    assert paragraph.alignment is None and paragraph.declared("space_after") is None
    assert [r.text for r in paragraph.runs] == ["Revenue grew ", "by ", "fourteen", " percnt",
                                                 " in the third quarter."]
    assert all(not run.effective.bold for run in paragraph.runs)


def test_clearing_keeps_styles_and_list_membership(lists_doc):
    document = Document.open(lists_doc)
    paragraph = document.paragraph("p:5A000003")
    paragraph.format(bold=True, indent_left=100)
    paragraph.clear_direct_formatting()
    assert paragraph.style == "Lijstalinea" and paragraph.list.num_id == 1


def test_clear_a_range(lists_doc):
    document = Document.open(lists_doc)
    document.anchor("fourteen").clear_formatting()
    paragraph = document.paragraph("p:5A000002")
    assert "fourteen" in [r.text for r in paragraph.runs] or "by fourteen" in [r.text for r in paragraph.runs]
    assert not any(r.effective.bold for r in paragraph.runs)


# -- effective --------------------------------------------------------------------------------


def test_effective_reads_explain_where_values_come_from(markup_doc):
    document = Document.open(markup_doc)
    heading = document.paragraph("p:1A2B3C01")
    run = heading.runs[0]
    assert run.bold is None and run.effective.bold and run.effective.size == 16
    assert "paragraph style 'Heading1'" in run.effective.explain()
    assert heading.effective.space_before == 12 and heading.effective.keep_with_next
    plain = document.paragraph("p:1A2B3C02").runs[0].effective
    assert plain.size == 11 and plain.font == "Calibri" and not plain.bold


def test_effective_follows_edits_and_undo(markup_doc):
    document = Document.open(markup_doc)
    run = document.paragraph("p:1A2B3C02").run(0)
    assert not run.effective.bold
    run.bold = True
    assert run.effective.bold
    document.undo()
    assert not document.paragraph("p:1A2B3C02").run(0).effective.bold


def test_effective_in_a_table_cell_and_a_header(markup_doc):
    document = Document.open(markup_doc)
    cell = document.paragraph("p:3C4D5E06")
    assert cell.effective.space_after == 0                 # TableGrid's paragraph spacing
    header = document.paragraph("header1/p:4D5E6F01")
    assert header.runs[0].effective.size == 11


def test_mixed_formatting_survives_on_every_fixture(docx_path):
    """set_text and replace on each fixture's paragraph with the most runs: every character
    outside the edit keeps its formatting."""
    from test_ranges import formats

    document = Document.open(docx_path)
    candidates = [p for p in document.paragraphs() if len(p.text) > 12 and "￼" not in p.text]
    if not candidates:
        pytest.skip("no paragraph to edit")
    paragraph = max(candidates, key=lambda p: (len(p.runs), len(p.text)))
    before = formats(paragraph)
    text = paragraph.text
    paragraph.set_text(text[:4] + "EDIT" + text[8:])
    after = formats(document.paragraph(paragraph.id))
    assert after[:4] == before[:4] and after[8:] == before[8:]
    word = next((w for w in text.split() if len(w) > 3 and text.count(w) == 1), None)
    if word is not None and word in document.paragraph(paragraph.id).text:
        current = document.paragraph(paragraph.id)
        start = current.text.index(word)
        before = formats(current)
        document.anchor(word, within=current.id).replace(word.upper())
        after = formats(document.paragraph(paragraph.id))
        assert after[:start] == before[:start]
        assert after[start + len(word):] == before[start + len(word):]
