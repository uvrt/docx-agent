"""``Document.blocks()`` (production round 21, T6): the body's top-level blocks in order with
id, kind, style, outline level (as a table of contents reads it), section break and
whether ``move_blocks`` takes it."""

from __future__ import annotations

import pytest

from docx_agent import Document, EditError

DATE = "2026-10-28T09:00:00Z"


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _outline(document: Document, identifier: str, value: int) -> None:
    """A paragraph's own ``w:outlineLvl`` (0-based, as written), put in by hand."""
    from lxml import etree

    element = document.paragraph(identifier)._element
    properties = element.find(W + "pPr")
    if properties is None:
        properties = etree.SubElement(element, W + "pPr")
        element.insert(0, properties)
    etree.SubElement(properties, W + "outlineLvl").set(W + "val", str(value))
    document._invalidate()


def made() -> tuple[Document, str, str]:
    document = Document.new(created=DATE)
    first = document.paragraphs()[0]
    first.set_text("Chapter title")
    first.set_style("heading 1")
    body = first.insert_after("Some body text.")
    document.paragraph(body.id).set_style("Normal")
    sub = document.paragraph(body.id).insert_after("A sub point")
    document.paragraph(sub.id).set_style("heading 2")
    own = document.paragraph(sub.id).insert_after("Its own level")
    document.paragraph(own.id).set_style("Normal")
    _outline(document, own.id, 2)
    table = document.insert_table(2, 2, after=own.id).id
    document.set_table(table, style="Table Grid")
    control = document.insert_control(None, "rich-text", after=table, text="Inside").id
    document.insert_section_break(after=sub.id, kind="nextPage")
    return document, table, control


def test_blocks_in_order_with_kind_style_level_break_and_movable():
    document, table, control = made()
    blocks = document.blocks()
    kinds = [b["kind"] for b in blocks]
    assert kinds[:3] == ["paragraph", "paragraph", "paragraph"]
    assert "table" in kinds and "control" in kinds
    heading, body, sub = blocks[0], blocks[1], blocks[2]
    assert (heading["style"], heading["level"], heading["movable"]) == ("heading 1", 1, True)
    assert (body["style"], body["level"]) == ("Normal", None)
    assert next(b for b in blocks if b["kind"] == "paragraph" and b["level"] == 3)["style"] == "Normal"
    assert (sub["style"], sub["level"]) == ("heading 2", 2)
    breaking = [b for b in blocks if b["section_break"]]
    assert len(breaking) == 1 and breaking[0]["kind"] == "paragraph" and not breaking[0]["movable"]
    found = next(b for b in blocks if b["kind"] == "table")
    assert (found["id"], found["style"], found["level"], found["movable"]) == (table, "Table Grid", None, True)
    found = next(b for b in blocks if b["kind"] == "control")
    assert (found["id"], found["movable"], found["section_break"]) == (control, False, False)
    # Every paragraph and table id is one the API resolves; movable ones move.
    assert all(document.paragraph(b["id"]) for b in blocks if b["kind"] == "paragraph")
    assert set(blocks[0]) == {"id", "kind", "style", "level", "section_break", "movable"}


def test_a_section_break_paragraph_is_not_movable_and_move_blocks_says_so():
    document, _, _ = made()
    breaking = next(b for b in document.blocks() if b["section_break"])
    with pytest.raises(EditError, match="ends a section"):
        document.move_blocks([breaking["id"]], after=document.blocks()[0]["id"])


def test_level_is_the_outline_level_a_toc_reads():
    document, _, _ = made()
    for block in document.blocks():
        if block["kind"] == "paragraph":
            assert document.outline_level(block["id"]) == block["level"]
