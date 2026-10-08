"""Text ranges and anchors, and editing text by range: find, anchor, replace, insert,
delete -- across run, hyperlink, revision and paragraph boundaries, keeping formatting."""

from __future__ import annotations

import pytest

from docx_agent.oxml.xml import qn
from lxml import etree

from docx_agent import AmbiguousAnchor, AnchorNotFound, Document, EditError, TextRange

from test_roundtrip import entries

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def formats(paragraph) -> list[tuple[str, str]]:
    """Per character of the current view: (character, the run's properties and container)."""
    from docx_agent.edit import text as _text

    out = []
    for atom in _text.atoms(paragraph._element):
        run = atom.run
        properties = run.find(W + "rPr") if run is not None else None
        key = etree.tostring(properties, method="c14n", exclusive=True).decode() if properties is not None else ""
        parent = run.getparent().tag.rpartition("}")[2] if run is not None else ""
        out.append((atom.char, key + "|" + parent))
    return out


@pytest.fixture
def lists_doc(markup_doc):
    return markup_doc.parent / "lists-and-styles.docx"


# -- ids --------------------------------------------------------------------------------------


def test_range_ids_round_trip(markup_doc):
    document = Document.open(markup_doc)
    found = document.anchor("bold and")
    assert found.id == "p:1A2B3C02@6:14" and found.text == "bold and"
    again = document.range(found.id)
    assert again == found and again.text == "bold and"
    point = document.range("p:1A2B3C02@3")
    assert point.collapsed and point.text == ""
    across = document.range("p:1A2B3C10@2..p:1A2B3C11@3")
    assert across.text == "bookmark starts here\nand"
    assert across.paragraph_ids() == ["p:1A2B3C10", "p:1A2B3C11"]
    positional = document.range("p@body/2@2:9")
    assert positional.text == "paragra"


def test_a_paragraph_range_and_its_bounds(markup_doc):
    paragraph = Document.open(markup_doc).paragraph("p:1A2B3C02")
    assert paragraph.range().text == paragraph.text
    assert paragraph.range(6, 10).text == "bold"
    with pytest.raises(IndexError):
        paragraph.range(5, 100)


def test_ranges_follow_a_stamped_paragraph(markup_doc):
    document = Document.open(markup_doc)
    found = document.anchor("no paraId")
    assert found.id.startswith("p@body/2@")
    document.paragraph("p:1A2B3C02").set_text("Stamps the document.")
    assert found.id.startswith("p:") and found.text == "no paraId"


# -- finding ----------------------------------------------------------------------------------


def test_find_across_runs_and_containers(markup_doc):
    document = Document.open(markup_doc)
    assert [r.id for r in document.find("bold and ital")] == ["p:1A2B3C02@6:19"]
    assert document.find("the site today")[0].id == "p:1A2B3C09@6:20"     # into a hyperlink and out
    assert document.find("grew 14%")[0].text == "grew 14%"               # across an insertion
    assert document.find("grew 12%") == []                               # the deletion is not current
    assert document.find("grew 12%", view="original")[0].view == "original"
    assert document.find("Ada Lovelace")[0].id == "p:1A2B3C0C@6:18"       # a content control's content
    assert [r.id for r in document.find("A header")] == ["header1/p:4D5E6F01@0:8"]


def test_find_options(markup_doc):
    document = Document.open(markup_doc)
    assert document.find("PLAIN BOLD") == []
    assert len(document.find("PLAIN BOLD", case=False)) == 1
    assert document.find("Plain  bold") == []
    assert document.find("Plain  bold", whitespace="collapse")[0].text == "Plain bold"
    assert [r.text for r in document.find(r"\b\w+ repeats", regex=True)] == ["three repeats"] * 3
    assert document.find("repeats.\nSecond") == []
    across = document.find("repeats.\nSecond", across_paragraphs=True)
    assert len(across) == 1 and not across[0].single
    assert len(document.find("of three", within="p:2B3C4D05")) == 1
    assert len(document.find("B2", within="t:3C4D5E01")) == 1
    assert len(document.find("Page", within="footer1")) == 1


def test_anchor_refuses_ambiguity(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(AmbiguousAnchor) as raised:
        document.anchor("repeats")
    assert len(raised.value.candidates) == 3
    assert "p:2B3C4D05@" in str(raised.value)
    assert document.anchor("repeats", occurrence=1).text == "repeats"
    assert document.anchor("First of three repeats").id == "p:2B3C4D05@0:22"
    with pytest.raises(AnchorNotFound):
        document.anchor("nowhere to be found")


# -- replacing --------------------------------------------------------------------------------


def test_replace_across_runs_takes_the_first_characters_formatting(markup_doc):
    document = Document.open(markup_doc)
    paragraph = document.paragraph("p:1A2B3C02")
    before = formats(paragraph)
    document.anchor("bold and ital").replace("strong and slant")
    after = formats(paragraph)
    assert paragraph.text == "Plain strong and slantic runs."
    bold = before[6][1]
    assert after[:6] == before[:6]                          # untouched before
    assert all(f == bold for _, f in after[6:22])           # the replacement: the first character's
    assert after[22:] == before[19:]                        # untouched after, italic included


def test_replace_keeping_characters(markup_doc):
    document = Document.open(markup_doc)
    paragraph = document.paragraph("p:1A2B3C02")
    before = formats(paragraph)
    document.anchor("bold and italic").replace("bold or italic", keep="characters")
    after = formats(paragraph)
    assert paragraph.text == "Plain bold or italic runs."
    assert after[6:10] == before[6:10]                      # "bold" still bold
    assert after[14:20] == before[15:21]                    # "italic" still italic


def test_replace_in_a_hyperlink_keeps_the_link(markup_doc):
    document = Document.open(markup_doc)
    document.anchor("site").replace("page")
    paragraph = document.paragraph("p:1A2B3C09")
    link = paragraph._element.find(W + "hyperlink")
    assert "".join(link.itertext()) == "the page"
    assert document.paragraph("p:1A2B3C09").text == "Visit the page today."


def test_replace_refuses_to_cut_a_field(markup_doc):
    document = Document.open(markup_doc)
    data = document.to_bytes()
    with pytest.raises(EditError, match="field"):
        document.anchor("Page 1", within="body").replace("Sheet 2")   # into a field's result
    with pytest.raises(EditError, match="field"):
        document.replace(": Someone", "-")                            # into a simple field's
    assert entries(document.to_bytes()) == entries(data)


def test_replace_inside_one_fields_result_is_allowed(markup_doc):
    document = Document.open(markup_doc)
    document.replace("Someone", "Somebody")
    assert document.paragraph("p:1A2B3C0B").text == "Author: Somebody"


def test_replace_refuses_to_delete_a_note_reference(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(EditError, match="note"):
        document.paragraph("p:1A2B3C11").range(10, 15).delete()


def test_replace_all_is_one_step_and_counts(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    result = document.replace("three", "3")
    assert result.count == 3
    assert [p.text for p in document.paragraphs()[3:6]] == [
        "First of 3 repeats.", "Second of 3 repeats.", "Third of 3 repeats."]
    assert document.undo()
    assert entries(document.to_bytes()) == entries(data)


def test_replace_with_groups_and_a_count(markup_doc):
    document = Document.open(markup_doc)
    result = document.replace(r"(\w+) of three", r"\1 of 3", regex=True, count=2)
    assert result.count == 2
    assert [p.text for p in document.paragraphs()[3:6]] == [
        "First of 3 repeats.", "Second of 3 repeats.", "Third of three repeats."]


def test_replace_leaves_other_authors_deletions_alone(markup_doc):
    document = Document.open(markup_doc)
    document.anchor("grew 14%").replace("rose 15%")
    paragraph = document.paragraph("p:1A2B3C0D")
    assert paragraph.text == "Revenue rose 15% this year."
    assert "12" in paragraph.text_in("original")


def test_insert_and_delete_text(markup_doc):
    document = Document.open(markup_doc)
    paragraph = document.paragraph("p:1A2B3C02")
    before = formats(paragraph)
    document.insert_text("p:1A2B3C02@10", "face")             # after "bold": takes bold
    assert paragraph.text == "Plain boldface and italic runs."
    assert formats(paragraph)[10][1] == before[9][1]
    document.insert_text(paragraph.range(0, 0), ">> ")         # at the start: the first character's
    assert formats(paragraph)[0][1] == before[0][1]
    document.delete_text(document.anchor(" and italic"))
    assert paragraph.text == ">> Plain boldface runs."
    assert len(paragraph.runs) == 3


def test_editing_leaves_no_empty_run(markup_doc):
    document = Document.open(markup_doc)
    document.anchor("bold").delete()
    xml = etree.tostring(document.paragraph("p:1A2B3C02")._element)
    assert b"<w:b/>" not in xml
    for run in document.paragraph("p:1A2B3C02").runs:
        assert run.text


def test_a_range_across_paragraphs_joins_them(markup_doc):
    """As Word joins paragraphs across a deleted mark (measured, tools/e3_probe.py): the
    last paragraph keeps its element and id, the first's text comes to its start, and it
    takes the first's properties."""
    document = Document.open(markup_doc)
    found = document.find("repeats.\nSecond of", across_paragraphs=True)[0]
    first, second = found.paragraph_ids()
    first_properties = etree.tostring(document.paragraph(first)._element.find(qn("w:pPr"))) \
        if document.paragraph(first)._element.find(qn("w:pPr")) is not None else None
    result = found.replace("repeats; then the second of")
    survivor = result.renamed.get(second, second)
    assert result.removed == [result.renamed.get(first, first)]
    assert result.id == survivor
    assert document.paragraph(survivor).text == "First of three repeats; then the second of three repeats."
    properties = document.paragraph(survivor)._element.find(qn("w:pPr"))
    assert (etree.tostring(properties) if properties is not None else None) == first_properties
    with pytest.raises(KeyError):
        document.paragraph(result.renamed.get(first, first))


def test_a_range_across_a_table_or_a_section_is_refused(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(EditError, match="table"):
        document.range("p:1A2B3C14@3..p:1A2B3C15@3").delete()


def test_ranges_in_other_views_are_read_only(markup_doc):
    document = Document.open(markup_doc)
    found = document.find("grew 12%", view="original")[0]
    with pytest.raises(EditError, match="view"):
        found.replace("x")


def test_overlapping_ranges_are_refused(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(EditError, match="overlap"):
        document._replace_ranges([(document.range("p:1A2B3C02@0:8"), "a"),
                                  (document.range("p:1A2B3C02@5:10"), "b")], keep="first")


def test_every_edit_is_undone_exactly(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    document.anchor("bold and ital").replace("x")
    document.insert_text("p:1A2B3C09@0", "Go and ")
    document.find("repeats.\nSecond", across_paragraphs=True)[0].delete()
    edited = entries(document.to_bytes())
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)
    while document.redo():
        pass
    assert entries(document.to_bytes()) == edited


def test_set_text_keeps_mixed_formatting(lists_doc):
    document = Document.open(lists_doc)
    paragraph = document.paragraph("p:5A000002")
    before = formats(paragraph)
    paragraph.set_text("Revenue grew by fifteen percent in the third quarter.")
    after = formats(paragraph)
    assert paragraph.text == "Revenue grew by fifteen percent in the third quarter."
    assert after[:16] == before[:16]
    assert after[16] == before[16]                       # "f" replaces "f": bold
    assert after[-14:] == before[-14:]                   # "third quarter." red and underlined
    assert [f for c, f in after if c in "third"][-1] == [f for c, f in before if c in "third"][-1]
    assert isinstance(paragraph.range(), TextRange)


def test_text_inserted_at_a_fields_edge_stays_out_of_it(markup_doc):
    document = Document.open(markup_doc)
    document.insert_text("p:1A2B3C0B@15", "!")                 # after "Someone", a simple field's result
    paragraph = document.paragraph("p:1A2B3C0B")
    simple = paragraph._element.find(f".//{W}fldSimple")
    assert "".join(simple.itertext()) == "Someone" and paragraph.text == "Author: Someone!"
    document.insert_cross_reference("p:1A2B3C0B@0", "span")
    document.insert_text("p:1A2B3C0B@0", ">")                 # before a field that starts the paragraph
    paragraph = document.paragraph("p:1A2B3C0B")
    first_run = paragraph._element.find(W + "r")
    assert first_run.find(W + "t") is not None and first_run.find(W + "t").text == ">"
    assert paragraph.text.startswith(">A bookmark starts here")
