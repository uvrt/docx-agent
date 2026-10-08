"""E0's edits -- set a paragraph's text, insert a paragraph, delete a block -- and undo.

Undo is exact: after any edit set, undoing it (as one batch, or step by step) gives back the
original bytes of every part, and redoing gives back the edited bytes.
"""

from __future__ import annotations

import pytest

from docx_agent import Document, EditError

from test_roundtrip import entries


def edit_set(document: Document) -> int:
    """E0's edits on a document: returns how many undo steps they took."""
    paragraphs = [p for p in document.paragraphs() if p.text.strip()]
    if not paragraphs:
        document.insert_paragraph("Only paragraph.", after=document.tables()[0].id) if document.tables() else None
        return 0
    steps = 0
    first, last = paragraphs[0], paragraphs[-1]
    first.set_text(first.text.replace(" ", "  ", 1) + " (edited)")
    steps += 1
    document.insert_paragraph("Inserted after the first.", after=first.id)
    steps += 1
    document.insert_paragraph("Inserted before the last.", before=last.id)
    steps += 1
    victim = next((p for p in paragraphs[1:-1] if _can_delete(document, p)), None)
    if victim is not None:
        document.delete_block(victim.id)
        steps += 1
    return steps


def _can_delete(document: Document, paragraph) -> bool:
    try:
        from docx_agent.edit.document import _check_container_survives

        _check_container_survives(paragraph._element)
    except EditError:
        return False
    return not paragraph.ends_section and "￼" not in paragraph.text


def test_undo_step_by_step_gives_the_original_bytes(docx_path):
    data = docx_path.read_bytes()
    document = Document.open(data)
    steps = edit_set(document)
    edited = entries(document.to_bytes())
    for _ in range(steps):
        assert document.undo()
    assert entries(document.to_bytes()) == entries(data)
    assert document.package.changed_parts() == frozenset()
    for _ in range(steps):
        assert document.redo()
    assert entries(document.to_bytes()) == edited


def test_a_batch_is_one_undo_step(docx_path):
    data = docx_path.read_bytes()
    document = Document.open(data)
    with document.batch():
        steps = edit_set(document)
    if not steps:
        pytest.skip("nothing to edit")
    edited = entries(document.to_bytes())
    assert document.undo()
    assert entries(document.to_bytes()) == entries(data)
    assert not document.history.can_undo()
    assert document.redo()
    assert entries(document.to_bytes()) == edited


def test_a_failed_batch_rolls_back(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    with pytest.raises(EditError):
        with document.batch():
            document.paragraph("p:1A2B3C02").set_text("Changed.")
            document.delete_block("p:1A2B3C11")  # holds a footnote reference: refused
    assert entries(document.to_bytes()) == entries(data)
    assert document.paragraph("p:1A2B3C02").text == "Plain bold and italic runs."


def changed(data: bytes, document: Document) -> set[str]:
    original, written = entries(data), entries(document.to_bytes())
    return {name for name in original if original[name] != written[name]}


def test_only_the_edited_part_changes_once_every_paragraph_has_an_id(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    document.paragraph("p:1A2B3C02").set_text("Stamps the body's missing and invalid ids.")
    stamped = document.to_bytes()
    assert changed(data, document) == {"word/document.xml"}
    document = Document.open(stamped)
    document.paragraph("header1/p:4D5E6F01").set_text("A new header.")
    assert changed(stamped, document) == {"word/header1.xml"}


def test_paragraph_stamping_changes_only_the_edited_part(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data, stamping="paragraph")
    document.paragraph("header1/p:4D5E6F01").set_text("A new header.")
    assert changed(data, document) == {"word/header1.xml"}


def test_compatibility_mode_is_never_changed(docx_path):
    document = Document.open(docx_path)
    mode = document.compatibility_mode
    edit_set(document)
    assert Document.open(document.to_bytes()).compatibility_mode == mode


# -- insert_paragraph -------------------------------------------------------------------------


def test_insert_continues_the_paragraph_beside_it(markup_doc):
    document = Document.open(markup_doc)
    result = document.insert_paragraph("Another heading", after="p:1A2B3C01")
    paragraph = document.paragraph(result.id)
    assert paragraph.style == "Heading1"
    assert [p.id for p in document.paragraphs()][:2] == ["p:1A2B3C01", result.id]


def test_insert_with_a_style_by_name_or_id(markup_doc):
    document = Document.open(markup_doc)
    by_name = document.insert_paragraph("Styled", after="p:1A2B3C02", style="heading 1")
    assert document.paragraph(by_name.id).style == "Heading1"
    with pytest.raises(EditError):
        document.insert_paragraph("x", after="p:1A2B3C02", style="No Such Style")


def test_insert_never_copies_a_section_break(markup_doc):
    document = Document.open(markup_doc)
    first = document.paragraphs()[0]
    from docx_agent.edit.document import W_PPR, W_SECTPR

    properties = first._element.find(W_PPR)
    properties.append(properties.makeelement(W_SECTPR))
    result = document.insert_paragraph("After a section break.", after=first.id)
    assert not document.paragraph(result.id).ends_section


def test_insert_beside_a_table_and_in_a_cell(markup_doc):
    document = Document.open(markup_doc)
    table = document.tables()[0]
    after = document.insert_paragraph("After the table.", after=table.id)
    before = document.insert_paragraph("Before the table.", before=table.id)
    ids = [p.id for p in document.paragraphs()]
    cell = table.cell(1, 1).paragraphs[0]
    assert ids.index(before.id) < ids.index(cell.id) < ids.index(after.id)
    inside = document.insert_paragraph("Second in B2.", after=cell.id)
    assert [p.text for p in table.cell(1, 1).paragraphs] == ["B2", "Second in B2."]
    assert document.paragraph(inside.id).text == "Second in B2."


def test_insert_needs_exactly_one_anchor(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(EditError):
        document.insert_paragraph("x")
    with pytest.raises(EditError):
        document.insert_paragraph("x", after="p:1A2B3C01", before="p:1A2B3C02")


# -- delete_block ----------------------------------------------------------------------------


def test_delete_a_paragraph_and_a_table(markup_doc):
    document = Document.open(markup_doc)
    result = document.delete_block("p:1A2B3C02")
    assert result.removed == ["p:1A2B3C02"]
    with pytest.raises(KeyError):
        document.paragraph("p:1A2B3C02")
    table = document.delete_block("t:3C4D5E01")
    assert "p:3C4D5E06" in table.removed and document.tables() == []


@pytest.mark.parametrize("identifier, reason", [
    ("p:1A2B3C11", "footnoteReference"),
    ("p:1A2B3C13", "commentReference"),
])
def test_delete_is_refused_where_it_would_orphan_something(markup_doc, identifier, reason):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    with pytest.raises(EditError, match=reason):
        document.delete_block(identifier)
    assert entries(document.to_bytes()) == entries(data)


def test_delete_keeps_a_cell_ending_in_a_paragraph(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(EditError):
        document.delete_block("p:3C4D5E06")


def test_delete_keeps_the_other_end_of_a_bookmark(markup_doc):
    document = Document.open(markup_doc)
    document.delete_block("p:1A2B3C10")
    xml = document.package.read("word/document.xml")
    assert xml.count(b"<w:bookmarkStart") == 1 and xml.count(b"<w:bookmarkEnd") == 1
    following = document.paragraph("p:1A2B3C11")._element
    assert following[0].tag.endswith("bookmarkStart")


def test_delete_the_only_paragraph_of_a_story_is_refused(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(EditError):
        document.delete_block("header1/p:4D5E6F01")


# -- move_block ------------------------------------------------------------------------------


def test_move_a_paragraph_and_a_table_keeping_their_ids(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    document.move_block("p:1A2B3C17", after="p:1A2B3C01")
    ids = [p.id for p in document.paragraphs()]
    assert ids[:2] == ["p:1A2B3C01", "p:1A2B3C17"]
    document.paragraph("p:1A2B3C02").move(before="p:1A2B3C01")
    assert [p.id for p in document.paragraphs()][:3] == ["p:1A2B3C02", "p:1A2B3C01", "p:1A2B3C17"]
    document.move_block("t:3C4D5E01", before="p:1A2B3C02")
    assert document.paragraphs()[0].id == "p:3C4D5E02"
    from docx_agent.validate import check

    known = set(check(Document.open(data).package))
    assert set(check(Document.open(document.to_bytes()).package)) - known == set()
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)


def test_moves_that_would_break_the_document_are_refused(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(EditError):
        document.move_block("t:3C4D5E01", after="p:3C4D5E06")          # into itself
    with pytest.raises(EditError):
        document.move_block("p:3C4D5E06", after="p:1A2B3C01")          # the cell's last paragraph
    with pytest.raises(EditError):
        document.move_block("header1/p:4D5E6F01", after="p:1A2B3C01")  # across stories
