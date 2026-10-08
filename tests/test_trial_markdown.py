"""``to_markdown`` shows what the trial's agents could not see (ROADMAP.md, "Trial findings",
4): a header's comment (``stories=``), a field's result told from typed text, and a heading
a list numbers told from one whose number is typed."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_agent import Document

DATE = "2026-10-05T09:00:00Z"
PILOT = Path(__file__).parent / "fixtures" / "generated" / "pilot" / "agreement-summary.docx"


def document() -> Document:
    made = Document.new(created=DATE)
    made.insert_markdown("# 1 Purpose\n\nWhy.[^1]\n\n# Access\n\nSee section .\n\n[^1]: A note.\n")
    made.add_bookmark(made.paragraphs()[0].range(), "RefPurpose")
    made.insert_cross_reference(made.paragraphs()[3].range(12, 12), "RefPurpose")
    made.add_header("s:body", "default", "Draft")
    made.add_comment(made.paragraphs("header1")[0].range(0, 5), "Change to Final", author="Legal")
    made.add_footer("s:body", "default", "")                   # an empty footer: not shown
    return made


def test_the_body_alone_by_default():
    made = document()
    assert "Change to Final" not in made.to_markdown(view="markup")
    assert made.to_markdown() == made.to_markdown(stories="body")


def test_all_stories_show_a_header_and_its_comment():
    made = document()
    text = made.to_markdown(view="markup", stories="all")
    assert "<!-- story: header1 (header) -->" in text
    assert "Draft{>>Legal: Change to Final<<}" in text
    assert "footer1" not in text                                 # no content: left out
    assert text.count("[^fn:1]:") == 1
    named = made.to_markdown(stories=["body", "header1"], ids=False)
    assert "See section 1 Purpose.\n\nDraft\n" in named
    with pytest.raises(KeyError, match="no story 'header9'"):
        made.to_markdown(stories=["header9"])
    with pytest.raises(ValueError, match="no range"):
        made.to_markdown("p:5FCB534E", stories="all")


def test_the_markup_view_marks_a_field_result():
    made = document()
    markup = made.to_markdown(view="markup")
    assert "See section <!-- field: REF RefPurpose \\h -->1 Purpose<!-- /field -->." in markup
    assert "<!-- field" not in made.to_markdown()                 # the final view: text only
    made.insert_page_number(made.paragraphs("footer1")[0].range(0, 0))
    footer = made.to_markdown(view="markup", stories=["footer1"])
    assert "field: PAGE" in footer and "1<!-- /field -->" in footer   # at a line's start: in the id comment


def test_a_heading_says_how_it_is_numbered():
    made = document()
    access = made.paragraphs()[2]
    access.add_to_list("number")
    text = made.to_markdown()
    assert "<!-- p:5FCB534E numbered: text bm:RefPurpose -->\n# 1 Purpose" in text
    assert f"<!-- {access.id} numbered: list -->\n# 1. Access" in text


def test_reading_changes_nothing():
    made = document()
    before = made.to_bytes()
    made.to_markdown(view="markup", stories="all")
    assert made.to_bytes() == before


def test_the_pilot_body_reads_as_before():
    """Headings without numbers carry no numbering note."""
    assert "numbered" not in Document.open(PILOT).to_markdown()
