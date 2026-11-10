"""Document.block_text(id): any block's plain text, whatever its kind; paragraph() and table()
on an id of another kind name that kind and what reads it (round 49, R92)."""

from __future__ import annotations

import pytest

from docx_agent import Document
from docx_agent.edit.addresses import AddressError

MD = """# Scope

Lead paragraph.

| Code | Amount |
|---|---|
| DOC-004.2 | 12 |
| DOC-007.1 | 30 |

Closing paragraph."""


def made() -> tuple[Document, dict[str, str]]:
    document = Document.new()
    document.insert_markdown(MD)
    blocks = [b["id"] for b in document.state()["blocks"]]
    ids = dict(zip(["head", "lead", "grid", "closing"], blocks))
    ids["block"] = document.insert_control(None, "rich-text", text="Inside one.\nInside two.", after=ids["lead"],
                                           tag="blockone").id
    ids["line"] = document.insert_control(document.anchor("Closing").id, "text", text="Closing", tag="lineone").id
    return document, ids


def test_a_paragraph_a_table_and_a_content_control():
    document, ids = made()
    assert document.block_text(ids["lead"]) == "Lead paragraph."
    assert document.block_text(ids["grid"]) == "Code\tAmount\nDOC-004.2\t12\nDOC-007.1\t30"
    assert document.block_text(ids["block"]) == "Inside one.\nInside two."
    assert document.block_text(ids["line"]) == "Closing"


def test_cells_rows_spans_headings_sections_stories_bookmarks_and_ranges():
    document, ids = made()
    grid = ids["grid"]
    assert document.block_text(f"{grid}/c1,0") == "DOC-004.2"
    assert document.block_text(f"{grid}/r2") == "DOC-007.1\t30"
    assert document.block_text(f"{ids['lead']}..{grid}") == ("Lead paragraph.\nInside one.\nInside two.\n"
                                                            "Code\tAmount\nDOC-004.2\t12\nDOC-007.1\t30")
    whole = document.block_text("body")
    assert document.block_text(f"h:{ids['head']}") == whole == document.block_text("s:1")
    assert whole.startswith("Scope\nLead paragraph.") and whole.endswith("Closing paragraph.")
    document.add_bookmark(grid, "Amounts")
    assert document.block_text("bm:Amounts") == document.bookmark("Amounts").text
    assert document.block_text(f"{ids['lead']}@0:4") == "Lead"


def test_a_merged_cell_once_a_cell_of_two_paragraphs_keeps_its_line_break():
    document = Document.new()
    document.insert_markdown("| A | B | C |\n|---|---|---|\n| 1 | 2 | 3 |")
    table = document.tables()[0].id
    document.merge_cells(table, (1, 0), (1, 1))
    # The merged cell once, holding both cells' paragraphs as Word's merge does.
    assert document.block_text(table) == "A\tB\tC\n1\n2\t3"
    cell = document.table(table).cell(0, 2)
    cell.paragraphs[0].insert_after("second line")
    assert document.block_text(f"{table}/c0,2") == "C\nsecond line"


def test_views_follow_tracked_changes():
    document, ids = made()
    with document.tracking(author="A", date="2026-11-10T12:00:00Z"):
        document.delete_row(ids["grid"], 2)
        document.paragraph(ids["lead"]).set_text("Lead paragraph, changed.")
    assert document.block_text(ids["grid"]) == "Code\tAmount\nDOC-004.2\t12"
    assert document.block_text(ids["grid"], view="original") == "Code\tAmount\nDOC-004.2\t12\nDOC-007.1\t30"
    assert document.block_text(ids["lead"]) == "Lead paragraph, changed."
    assert document.block_text(ids["lead"], view="original") == "Lead paragraph."
    with pytest.raises(ValueError, match="no text view"):
        document.block_text(ids["lead"], view="final")


def test_an_id_the_document_lacks_is_a_key_error():
    document, _ = made()
    with pytest.raises(KeyError, match="no paragraph or table"):
        document.block_text("p:00000000")
    with pytest.raises(KeyError, match="an id is empty"):
        document.block_text("")


# -- paragraph() and table() on another kind ----------------------------------------------------


def test_paragraph_on_a_table_or_a_control_names_its_kind_and_block_text():
    document, ids = made()
    with pytest.raises(KeyError) as found:
        document.paragraph(ids["grid"])
    assert str(found.value) == (f"{ids['grid']} is a table, not a paragraph: table('{ids['grid']}') reads it, "
                                f"block_text('{ids['grid']}') gives its text")
    with pytest.raises(AddressError) as found:
        document.paragraph(ids["block"])
    inner = [p.id for p in document.paragraphs() if p.text.startswith("Inside")]
    assert f"{ids['block']} is a content control (rich-text, block level), not a paragraph" in str(found.value)
    assert f"block_text('{ids['block']}')" in str(found.value) and found.value.options == inner
    with pytest.raises(AddressError) as found:
        document.paragraph(ids["line"])
    assert "(text, inline level)" in str(found.value) and found.value.options == [ids["closing"]]


def test_table_on_a_paragraph_a_cell_paragraph_or_a_control_names_its_kind():
    document, ids = made()
    with pytest.raises(KeyError, match=r"is a paragraph, not a table: paragraph\('.*'\) reads it, block_text"):
        document.table(ids["lead"])
    inside = document.table(ids["grid"]).cell(1, 1).paragraphs[0].id
    with pytest.raises(AddressError) as found:
        document.table(inside)
    assert f"is a paragraph in table {ids['grid']}" in str(found.value) and found.value.options == [ids["grid"]]
    with pytest.raises(AddressError, match=r"content control \(rich-text, block level\), not a table"):
        document.table(ids["block"])
    with pytest.raises(KeyError, match="no paragraph or table"):
        document.table("t:00000000")


def test_a_cell_or_span_given_to_paragraph_names_block_text_too():
    document, ids = made()
    with pytest.raises(AddressError, match=r"is a cell, .*block_text"):
        document.paragraph(f"{ids['grid']}/c1,1")
    with pytest.raises(AddressError, match=r"is a span; .*block_text"):
        document.table(f"{ids['lead']}..{ids['grid']}")


def test_to_markdown_of_a_block_control_holds_its_content():
    document, ids = made()
    markdown = document.to_markdown(range=ids["block"])
    assert "Inside one." in markdown and "Inside two." in markdown and markdown.rstrip().endswith(f"/{ids['block']} -->")
    span = document.to_markdown(range=f"{ids['lead']}..{ids['block']}")
    assert "Lead paragraph." in span and "Inside two." in span and "Code" not in span
