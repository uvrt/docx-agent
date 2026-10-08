"""The end-to-end trial's small surprises (ROADMAP.md, "Trial findings", 9): the empty
paragraph a new document starts with, ``_Ref`` bookmarks, hidden bookmarks, and ``created``
as an ISO string."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

import pytest

from docx_agent import Document, EditError
from docx_agent.validate import check

TEMPLATES = Path(__file__).parent / "fixtures" / "generated" / "templates"
DATE = datetime(2026, 10, 5, 9, 0, tzinfo=timezone.utc)


def texts(document: Document) -> list[str]:
    return [p.text for p in document.paragraphs()]


# -- the empty paragraph ------------------------------------------------------------------------


@pytest.mark.parametrize("make", [lambda: Document.new(created=DATE),
                                  lambda: Document.new(template=TEMPLATES / "brand.dotx", keep_content=False)],
                         ids=["blank", "template"])
def test_markdown_at_the_end_of_a_new_document_replaces_its_empty_paragraph(make):
    document = make()
    assert texts(document) == [""]
    lone = document.paragraphs()[0].id
    result = document.insert_markdown("# Title\n\nBody.")
    assert texts(document) == ["Title", "Body."]
    assert result.removed == [lone] and result.blocks == [p.id for p in document.paragraphs()]
    assert check(document.package) == []
    document.undo()
    assert texts(document) == [""]


def test_markdown_ending_in_a_table_keeps_the_paragraph_word_needs_after_it():
    document = Document.new(created=DATE)
    document.insert_markdown("Intro.\n\n| a |\n| --- |\n| 1 |\n")
    assert texts(document) == ["Intro.", "a", "1", ""]
    assert check(document.package) == []


def test_a_new_document_with_content_or_another_place_is_left_alone():
    document = Document.new(created=DATE)
    document.insert_markdown("First.")
    document.insert_markdown("Second.")
    assert texts(document) == ["First.", "Second."]
    other = Document.new(created=DATE)
    first = other.paragraphs()[0].id
    other.insert_markdown("After.", at=f"after:{first}")
    assert texts(other) == ["", "After."]
    empty = Document.new(created=DATE)
    assert not empty.insert_markdown("<!-- only a comment -->").changed and texts(empty) == [""]


def test_tracked_it_accepts_to_the_text_and_rejects_to_the_empty_paragraph():
    untracked = Document.new(created=DATE)
    untracked.insert_markdown("# Title\n\nBody.\n\n- item\n")
    document = Document.new(created=DATE)
    original = document.to_markdown(ids=False)
    with document.tracking(author="Claude", date=DATE):
        document.insert_markdown("# Title\n\nBody.\n\n- item\n")
    assert check(document.package) == []
    accepted = Document.open(document.to_bytes())
    accepted.accept_all()
    assert accepted.to_markdown(ids=False) == untracked.to_markdown(ids=False)
    rejected = Document.open(document.to_bytes())
    rejected.reject_all()
    assert texts(rejected) == [""] and rejected.to_markdown(ids=False) == original


# -- created= as a string ----------------------------------------------------------------------


@pytest.mark.parametrize("value,expected", [("2026-10-05", datetime(2026, 10, 5, tzinfo=timezone.utc)),
                                            ("2026-10-05T09:00:00Z", DATE),
                                            ("2026-10-05T11:00:00+02:00", DATE),
                                            (datetime(2026, 10, 5, 9, 0), DATE)])
def test_created_takes_an_iso_string_as_tracking_date_does(value, expected):
    assert Document.new(created=value).properties["created"] == expected
    assert Document.new(created=value).to_bytes() == Document.new(created=expected).to_bytes()
    made = Document.new(template=TEMPLATES / "brand.dotx", created=value)
    assert made.properties["created"] == expected


def test_created_refuses_what_is_not_a_date():
    with pytest.raises(ValueError, match="ISO 8601"):
        Document.new(created="last Tuesday")


# -- bookmarks -----------------------------------------------------------------------------------


def test_a_ref_bookmark_word_makes_for_cross_references_is_taken():
    document = Document.new(created=DATE)
    document.insert_markdown("# Retention\n\nSee the retention section.")
    heading = document.paragraphs()[0]
    document.add_bookmark(heading.range(), "_Ref123456789")
    assert document.bookmark("_Ref123456789").text == "Retention"
    document.insert_cross_reference(document.paragraphs()[1].range(4, 4), "_Ref123456789")
    assert document.paragraphs()[1].text.startswith("See Retention")
    assert check(document.package) == []
    with pytest.raises(EditError, match="_Ref<digits>"):
        document.add_bookmark(heading.range(0, 3), "_Other")


def test_bookmarks_says_it_leaves_out_hidden_ones_and_lists_them_when_asked():
    from docx_agent import Document as D

    assert "_Toc" in D.bookmarks.__doc__ and "hidden=True" in D.bookmarks.__doc__
    document = Document.new(created=DATE)
    document.insert_markdown("# Retention\n\nText.")
    document.add_bookmark(document.paragraphs()[0].range(), "_Ref1")
    document.add_bookmark(document.paragraphs()[1].range(), "Visible")
    assert [b.name for b in document.bookmarks()] == ["Visible"]
    assert [b.name for b in document.bookmarks(hidden=True)] == ["_Ref1", "Visible"]


# -- an empty table of contents (trial 2, N8) -----------------------------------------------------


def test_a_table_of_contents_put_in_before_its_headings_warns_until_they_exist():
    document = Document.new(created=DATE)
    toc = document.insert_toc(before=document.paragraphs()[0].id)
    assert len(toc.warnings) == 1 and toc.warnings[0].startswith(f"{toc.id}: the table of contents has no entries")
    document.insert_markdown("# One\n\nText.\n\n## Two\n")
    updated = document.update_fields()
    assert updated.warnings == [] and "One" in document.field(toc.id).result
    assert document.update_fields([toc.id]).warnings == []


def test_update_fields_names_every_table_of_contents_still_empty():
    document = Document.new(created=DATE)
    document.insert_markdown("Text only.")
    toc = document.insert_toc(before=document.paragraphs()[0].id).id
    assert document.update_fields().warnings == [document._empty_tocs()[0]]
    assert toc in document.update_fields().warnings[0]
