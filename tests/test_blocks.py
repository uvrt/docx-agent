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
    assert (found["id"], found["movable"], found["section_break"]) == (control, True, False)
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


def test_a_style_s_outline_level_is_the_level_of_a_paragraph_stating_none():
    """``Style.outline_level`` (round 23, V6): the rule ``blocks()`` uses for a paragraph that
    states no level of its own -- built-in heading N is N, else ``w:outlineLvl`` through
    ``basedOn``; body text, character and table styles ``None``."""
    document, _, _ = made()
    document.styles.add("Chapter heading", based_on="heading 2")
    styles = {s.name: s for s in document.styles}
    assert styles["heading 1"].outline_level == 1 and styles["heading 9"].outline_level == 9
    assert styles["Chapter heading"].outline_level == 2             # a heading in its chain
    assert styles["Normal"].outline_level is None and styles["Title"].outline_level is None
    assert styles["Heading 1 Char"].outline_level is None and styles["Table Grid"].outline_level is None
    by_id = {s.id: s for s in document.styles}
    for block in document.blocks():
        if block["kind"] != "paragraph":
            continue
        element = document.paragraph(block["id"])._element
        if element.find(f"{W}pPr/{W}outlineLvl") is not None:
            continue
        named = element.find(f"{W}pPr/{W}pStyle")
        style = by_id[named.get(W + "val")] if named is not None else document.styles.default("paragraph")
        assert style.outline_level == block["level"], block


def test_outline_level_through_based_on_and_body_text():
    from lxml import etree

    document, _, _ = made()
    document.styles.add("Annex", based_on="Normal")
    document.styles.add("Annex sub", based_on="Annex")
    document.styles.add("Plain text", based_on="Annex")
    nodes = {n.get(W + "styleId"): n for n in document.styles._nodes()}
    for name, value in (("Annex", "2"), ("Plain text", "9")):
        properties = nodes[document.styles.get(name).id].find(W + "pPr")
        if properties is None:
            properties = etree.SubElement(nodes[document.styles.get(name).id], W + "pPr")
        etree.SubElement(properties, W + "outlineLvl").set(W + "val", value)
    styles = {s.name: s for s in document.styles}
    assert styles["Annex"].outline_level == 3 and styles["Annex sub"].outline_level == 3
    assert styles["Plain text"].outline_level is None                 # body text (9)


# -- a block-level content control as a block (round 23, V2) ---------------------------------------


def _with_control() -> tuple[Document, list[str], str]:
    document = Document.new(created=DATE)
    result = document.insert_markdown("First\n\nSecond\n\nThird\n")
    control = document.insert_control(None, "rich-text", after=result.blocks[0], text="Inside one\nInside two").id
    return document, result.blocks, control


def _lock(document: Document, control: str, value: str) -> None:
    from lxml import etree

    sdt = document._control(control)[1]
    etree.SubElement(sdt.find(W + "sdtPr"), W + "lock").set(W + "val", value)
    document._invalidate()


def test_a_control_is_deleted_with_its_content_like_a_table():
    document, ids, control = _with_control()
    inside = {b.id for b in document.paragraphs() if b.text.startswith("Inside")}
    result = document.delete_block(control)
    assert result.removed[0] == control and len(inside) == 2 and inside <= set(result.removed)
    assert [b["id"] for b in document.blocks()] == ids
    assert "Inside" not in document.to_markdown(ids=False)
    document.undo()
    assert control in [b["id"] for b in document.blocks()]


def test_a_control_moves_copies_and_anchors_like_a_block():
    document, ids, control = _with_control()
    moved = document.move_blocks(control, after=ids[2])
    assert moved.blocks == [control] and [b["id"] for b in document.blocks()][-1] == control
    assert document.move_block(control, before=ids[0]).id == control
    assert [b["id"] for b in document.blocks()][0] == control
    span = document.move_blocks(f"{control}..{ids[0]}", after=ids[2])
    assert span.blocks[0] == control
    other = Document.new(created=DATE)
    copied = other.copy_blocks(document, control)
    assert copied.blocks[0].startswith("cc:") and copied.copied[control] == copied.blocks[0]
    assert "Inside one" in other.to_markdown(ids=False)
    after = document.insert_markdown("After the control.", at=f"after:{control}")
    assert [b["id"] for b in document.blocks()].index(after.blocks[0]) == \
        [b["id"] for b in document.blocks()].index(control) + 1


def test_a_tracked_move_of_a_control_is_untracked_and_says_so():
    document, ids, control = _with_control()
    with document.tracking(author="Reviewer"):
        result = document.move_block(control, after=ids[2])
    assert any("moved untracked" in w for w in result.warnings)


@pytest.mark.parametrize("value", ["sdtLocked", "contentLocked", "sdtContentLocked"])
def test_a_locked_control_is_not_deleted_or_moved_and_the_lock_is_named(value):
    document, ids, control = _with_control()
    _lock(document, control, value)
    for act in (lambda: document.delete_block(control), lambda: document.move_blocks(control, after=ids[2])):
        with pytest.raises(EditError, match=f"w:lock {value}"):
            act()
    assert next(b for b in document.blocks() if b["id"] == control)["movable"] is False


def test_an_inline_control_is_not_a_block():
    document, ids, _ = _with_control()
    inline = document.insert_control(f"{ids[1]}@2", "text", text="x").id
    with pytest.raises(EditError, match="not a block"):
        document.delete_block(inline)


def test_a_tracked_deletion_of_a_control_deletes_its_content_as_tracked():
    document, ids, control = _with_control()
    with document.tracking(author="Reviewer"):
        result = document.delete_block(control)
    assert any("no revision for a control" in w for w in result.warnings)
    assert len(document.revisions(author="Reviewer")) >= 1
