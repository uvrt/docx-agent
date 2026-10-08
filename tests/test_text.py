"""A paragraph's text, read across runs, fields, hyperlinks, content controls and revisions,
and written back keeping per-character formatting."""

from __future__ import annotations

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.edit import text as _text


def by_text(document: Document, start: str):
    for story in document.stories:
        for paragraph in story.paragraphs:
            if paragraph.text.startswith(start):
                return paragraph
    raise AssertionError(f"no paragraph starts {start!r}")


# -- reading ---------------------------------------------------------------------------------


@pytest.mark.parametrize("start, current, original", [
    ("Plain", "Plain bold and italic runs.", None),
    ("Visit", "Visit the site today.", None),          # hyperlink runs are text
    ("Page 1", "Page 1 of the document.", None),       # a field's result, not its instruction
    ("Author", "Author: Someone", None),               # a simple field's result
    ("Name", "Name: Ada Lovelace.", None),             # an inline content control's content
    ("Revenue", "Revenue grew 14% this year.", "Revenue grew 12% this year."),
    ("Moved from", "Moved from: ", "Moved from: travelling words"),
    ("Moved to", "Moved to: travelling words", "Moved to: "),
    ("Inside a block", "Inside a block content control.", None),
])
def test_text_in_the_current_and_original_views(markup_doc, start, current, original):
    paragraph = by_text(Document.open(markup_doc), start)
    assert paragraph.text == current
    assert paragraph.text_in("original") == (original or current)


def test_markup_view_has_both_sides(markup_doc):
    paragraph = by_text(Document.open(markup_doc), "Revenue")
    assert paragraph.text_in("markup") == "Revenue grew 1214% this year."


def test_special_characters(markup_doc):
    paragraph = by_text(Document.open(markup_doc), "Tab")
    assert paragraph.text == ("Tab\tthen line\vbreak, non‑breaking, soft­hyphen, "
                              "symbol .")
    assert by_text(Document.open(markup_doc), "\fSecond page").text == "\fSecond page."


def test_objects_are_one_character(markup_doc):
    document = Document.open(markup_doc)
    assert by_text(document, "and ends here").text == "and ends here.￼"
    footnote = document.paragraphs("footnotes")[2]
    assert footnote.text == "￼ A footnote."


def test_comment_markers_are_not_text(markup_doc):
    assert by_text(Document.open(markup_doc), "A commented").text == "A commented word."


def test_runs_include_the_ones_inside_containers(markup_doc):
    paragraph = by_text(Document.open(markup_doc), "Visit")
    assert [run.text for run in paragraph.runs] == ["Visit ", "the ", "site", " today."]
    assert paragraph.run(2).id.endswith("/r2")
    deleted = by_text(Document.open(markup_doc), "Revenue")
    assert [run.text for run in deleted.runs] == ["Revenue grew ", "14", "% this year."]


def test_unknown_view_is_refused(markup_doc):
    with pytest.raises(ValueError):
        by_text(Document.open(markup_doc), "Plain").text_in("final")


# -- writing ---------------------------------------------------------------------------------


def runs_xml(paragraph) -> list[tuple[str, str]]:
    """(run properties, text) per run, for comparing formatting."""
    out = []
    for run in _text.runs(paragraph._element):
        properties = run.find(_text.W_RPR)
        key = etree.tostring(properties).decode() if properties is not None else ""
        out.append((key, _text.run_text(paragraph._element, run)))
    return out


def test_set_text_keeps_each_characters_formatting(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "Plain")
    paragraph.set_text("Plain bold and italic words.")
    runs = runs_xml(paragraph)
    assert [text for _, text in runs] == ["Plain ", "bold", " and ", "italic", " words."]
    assert "w:b" in runs[1][0] and "w:i" in runs[3][0]


def test_a_replacement_takes_the_formatting_of_what_it_replaces(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "Plain")
    paragraph.set_text("Plain BOLDER and italic runs.")
    runs = runs_xml(paragraph)
    assert ("BOLDER" in runs[1][1]) and "w:b" in runs[1][0]


def test_an_insertion_continues_the_character_before_it(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "Plain")
    paragraph.set_text("Plain boldly and italic runs.")
    runs = runs_xml(paragraph)
    assert runs[1][1] == "boldly" and "w:b" in runs[1][0]


def test_untouched_containers_survive_an_edit(markup_doc):
    document = Document.open(markup_doc)
    link = by_text(document, "Visit")
    link.set_text("Visit the site tomorrow.")
    element = link._element
    assert element.find(_text._W + "hyperlink") is not None
    assert link.text == "Visit the site tomorrow."
    field = by_text(document, "Page 1")
    field.set_text("Page 1 of this document.")
    assert field._element.find(".//" + _text._W + "instrText").text == " PAGE "
    assert field.text == "Page 1 of this document."


def test_deleted_text_is_kept_when_the_current_text_changes(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "Revenue")
    paragraph.set_text("Revenue grew 14% last year.")
    assert paragraph.text_in("original") == "Revenue grew 12% last year."


def test_special_characters_are_written_as_elements(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "Plain")
    paragraph.set_text("Plain\tbold\vand italic runs.")
    assert paragraph.text == "Plain\tbold\vand italic runs."
    assert paragraph._element.find(".//" + _text._W + "tab") is not None
    assert paragraph._element.find(".//" + _text._W + "br") is not None


def test_deleting_everything_and_writing_into_an_empty_paragraph(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "Plain")
    paragraph.set_text("")
    assert paragraph.text == "" and paragraph.runs == []
    paragraph.set_text("Fresh")
    assert paragraph.text == "Fresh"


def test_leading_and_trailing_spaces_are_preserved(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "Plain")
    paragraph.set_text("  spaced  ")
    reopened = Document.open(document.to_bytes())
    assert any(p.text == "  spaced  " for p in reopened.paragraphs())


def test_objects_can_be_kept_but_not_created(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "and ends here")
    paragraph.set_text("and really ends here.￼")
    assert paragraph.text == "and really ends here.￼"
    with pytest.raises(ValueError):
        by_text(document, "Plain").set_text("Plain ￼")


def test_a_newline_is_refused(markup_doc):
    with pytest.raises(ValueError):
        by_text(Document.open(markup_doc), "Plain").set_text("two\nlines")


def test_setting_the_same_text_changes_nothing(markup_doc):
    document = Document.open(markup_doc)
    paragraph = by_text(document, "Plain")
    result = paragraph.set_text(paragraph.text)
    assert not result.changed
    assert document.package.dirty_parts == frozenset()
    assert not document.history.can_undo()


def test_edits_survive_save_and_reopen(docx_path):
    document = Document.open(docx_path)
    paragraphs = [p for p in document.paragraphs() if p.text.strip()]
    if not paragraphs:
        pytest.skip("no text to edit")
    target = paragraphs[len(paragraphs) // 2]
    new = target.text[: len(target.text) // 2] + " (edited) " + target.text[len(target.text) // 2:]
    result = target.set_text(new)
    reopened = Document.open(document.to_bytes())
    assert reopened.paragraph(result.id).text == new


def test_edit_error_is_a_value_error():
    assert issubclass(EditError, ValueError)
