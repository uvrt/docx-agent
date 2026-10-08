"""Every E1 operation on every fixture (ROADMAP.md, Phase E1, "Done when"): edit, save,
reopen and read back what the edit wrote; no validity problem added (numbering integrity,
unique ids and the rest); undo to the original bytes; redo to the edited ones.

Each operation picks its own targets from what the fixture has -- paragraphs of plain text
outside fields, consecutive paragraphs to join, a word to replace -- and returns what to
read back from the reopened document.
"""

from __future__ import annotations

import struct
import zlib

import pytest

from docx_agent import Document, EditError
from docx_agent.edit import text as _text
from docx_agent.edit.numbering import problems as numbering_problems
from docx_agent.validate import check

from test_roundtrip import entries


def png(width: int, height: int, color=(30, 140, 60)) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(color) * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def plain(document: Document, minimum: int = 12) -> list:
    """Body paragraphs with at least ``minimum`` characters, none an object or in a field."""
    out = []
    for paragraph in document.paragraphs():
        atoms = _text.atoms(paragraph._element)
        if len(atoms) >= minimum and all(a.field is None and a.char != _text.OBJECT for a in atoms):
            out.append(paragraph)
    return out


def text_of(document: Document) -> Document:
    """A paragraph to work on, even in a document without one (it is written first)."""
    if not plain(document):
        if document.paragraphs():
            document.paragraphs()[0].set_text("A paragraph written for the operation to work on.")
        else:
            document.append_paragraph("A paragraph written for the operation to work on.")
    return plain(document)[0]


def word_in(paragraph) -> str:
    words = [w.strip(".,;:()") for w in paragraph.text.split()]
    return next((w for w in words if len(w) >= 4 and paragraph.text.count(w) == 1), words[0])


# -- the operations: each edits and returns a read-back check on the reopened document --------


def op_set_text(document):
    paragraph = max(plain(document) or [text_of(document)], key=lambda p: len(p.runs))
    new = paragraph.text[:5] + "EDITED" + paragraph.text[9:]
    result = paragraph.set_text(new)
    return lambda saved: saved.paragraph(result.id).text == new


def op_insert_and_delete_text(document):
    paragraph = text_of(document)
    middle = len(paragraph.text) // 2
    document.insert_text(f"{paragraph.id}@{middle}", " [inserted] ")
    word = word_in(document.paragraph(paragraph.id))
    document.anchor(word, within=paragraph.id).delete()
    identifier = document.paragraph(paragraph.id).id
    expected = document.paragraph(identifier).text
    return lambda saved: saved.paragraph(identifier).text == expected and "[inserted]" in expected


def op_replace_all(document):
    paragraph = text_of(document)
    word = word_in(paragraph)
    try:
        result = document.replace(word, word.upper() + "#")
    except EditError:
        result = document.replace(word, word.upper() + "#", within=paragraph.id)
    count = result.count
    return lambda saved: sum(p.text.count(word.upper() + "#") for s in saved.stories for p in s.paragraphs) >= count >= 1


def op_replace_across_paragraphs(document):
    paragraphs = plain(document, 4)
    pair = next(((a, b) for a, b in zip(paragraphs, paragraphs[1:])
                 if a._element.getnext() is b._element and not a.ends_section and not b.ends_section), None)
    if pair is None:
        first = text_of(document)
        document.insert_paragraph("A second paragraph to join.", after=first.id)
        first = document.paragraph(first.id)
        pair = (first, document.paragraphs()[[p.id for p in document.paragraphs()].index(first.id) + 1])
    a, b = pair
    found = document.range(f"{a.id}@{len(a.text) - 2}..{b.id}@2")
    expected = a.text[:-2] + " | " + b.text[2:]
    result = found.replace(" | ")
    return lambda saved: saved.paragraph(result.id).text == expected


def op_format_range(document):
    paragraph = text_of(document)
    span = paragraph.range(2, min(len(paragraph.text), 9))
    span.format(bold=True, italic=True, color="1F4E79", size=13, underline="double", highlight="cyan")
    identifier = document.paragraph(paragraph.id).id

    def check_back(saved):
        runs = [r for r in saved.paragraph(identifier).runs if r.text]
        styled = [r for r in runs if r.declared("color") == "1F4E79"]
        return styled and all(r.effective.bold and r.effective.italic and r.effective.size == 13 for r in styled)

    return check_back


def op_format_paragraph(document):
    paragraph = text_of(document)
    paragraph.format(alignment="center", indent_left=18, space_after=9, line_spacing=1.25, keep_with_next=True)
    identifier = document.paragraph(paragraph.id).id

    def check_back(saved):
        effective = saved.paragraph(identifier).effective
        return (effective.alignment, effective.indent_left, effective.space_after, effective.line_spacing,
                effective.keep_with_next) == ("center", 18, 9, 1.25, True)

    return check_back


def op_paragraph_style(document):
    paragraph = text_of(document)
    paragraph.style = "Heading 2"
    identifier = document.paragraph(paragraph.id).id
    return lambda saved: saved.paragraph(identifier).style_name.casefold() == "heading 2"


def op_character_style(document):
    paragraph = text_of(document)
    paragraph.range(0, 5).set_style("Strong")
    identifier = document.paragraph(paragraph.id).id
    # Not "is it bold": Strong on a bold heading is not bold (Word's toggle rule).
    return lambda saved: saved.paragraph(identifier).runs[0].style == saved.styles.get("Strong", "character").id


def op_clear_formatting(document):
    paragraph = text_of(document)
    paragraph.format(alignment="right", bold=True)
    paragraph.clear_direct_formatting()
    identifier = document.paragraph(paragraph.id).id
    return lambda saved: saved.paragraph(identifier).declared("alignment") is None and all(
        r.declared("bold") is None for r in saved.paragraph(identifier).runs)


def op_lists(document):
    paragraphs = plain(document, 4)[:4]
    while len(paragraphs) < 3:
        document.insert_paragraph("A list item.", after=(paragraphs[-1] if paragraphs else text_of(document)).id)
        paragraphs = plain(document, 4)[:4]
    a, b, c = paragraphs[:3]
    a.add_to_list("number")
    b.add_to_list(a.list.num_id)
    b.indent_list()
    c.add_to_list("bullet")
    c.restart_numbering(4)
    b.outdent_list()
    ids = [document.paragraph(p.id).id for p in (a, b, c)]

    def check_back(saved):
        x, y, z = (saved.paragraph(i).list for i in ids)
        return (x.num_id == y.num_id and x.kind == "number" and y.level == 0 and z.kind == "bullet"
                and numbering_problems(saved) == [])

    return check_back


def op_list_continue_and_remove(document):
    paragraphs = plain(document, 4)[:3]
    while len(paragraphs) < 3:
        document.insert_paragraph("Another item.", after=(paragraphs[-1] if paragraphs else text_of(document)).id)
        paragraphs = plain(document, 4)[:3]
    a, b, c = paragraphs
    a.add_to_list("number")
    c.add_to_list("number", continue_previous=False)
    c.continue_numbering(from_id=a.id)
    a.remove_from_list()
    ids = [document.paragraph(p.id).id for p in (a, c)]
    return lambda saved: saved.paragraph(ids[0]).list is None and saved.paragraph(ids[1]).list is not None \
        and numbering_problems(saved) == []


def op_hyperlinks(document):
    paragraph = text_of(document)
    document.paragraph(paragraph.id).range(0, 4).add_bookmark("E1Target")
    external = document.paragraph(paragraph.id).range(5, 9).add_hyperlink("https://example.com/e1")
    internal = document.paragraph(paragraph.id).range(10, 12).add_hyperlink(anchor="E1Target")
    document.hyperlink(external.id).set_target("https://example.org/changed")
    document.hyperlink(internal.id).remove()
    identifier = document.paragraph(paragraph.id).id
    return lambda saved: [(h.address, h.anchor) for h in saved.paragraph(identifier).hyperlinks] == [
        ("https://example.org/changed", None)]


def op_bookmarks(document):
    paragraph = text_of(document)
    document.paragraph(paragraph.id).range(0, 6).add_bookmark("E1Mark")
    document.bookmark("E1Mark").rename("E1Renamed")
    document.insert_cross_reference(f"{document.paragraph(paragraph.id).id}@{len(document.paragraph(paragraph.id).text)}",
                                    "E1Renamed")
    expected = document.bookmark("E1Renamed").text
    document.paragraph(paragraph.id).range(1, 3).add_bookmark("E1Gone")
    document.bookmark("E1Gone").remove()
    return lambda saved: saved.bookmark("E1Renamed").text == expected and all(
        b.name != "E1Gone" for b in saved.bookmarks())


def op_pictures(document):
    paragraph = text_of(document)
    first = document.insert_picture(f"{paragraph.id}@0", png(48, 24), alt_text="Green")
    second = document.insert_picture(f"{document.paragraph(paragraph.id).id}@3", png(10, 10, (9, 9, 9)))
    document.picture(first.id).replace(png(20, 20, (200, 200, 0)))
    document.picture(first.id).resize(width=72)
    document.picture(second.id).delete()
    picture_id = first.id

    def check_back(saved):
        picture = saved.picture(picture_id)
        return picture.alt_text == "Green" and picture.size == (72.0, 36.0) and picture.image.startswith(b"\x89PNG")

    return check_back


def op_coalesce(document):
    paragraph = text_of(document)
    paragraph.range(2, 6).format(bold=True)
    paragraph = document.paragraph(paragraph.id)
    paragraph.range(2, 6).format(bold=None)
    document.coalesce_runs()
    identifier = paragraph.id
    expected = document.paragraph(identifier).text
    return lambda saved: saved.paragraph(identifier).text == expected


def op_styles_add_modify(document):
    based_on = "Normal" if document.styles.find("Normal", "paragraph") else None
    document.styles.add("E1 Callout", based_on=based_on, italic=True, space_before=4)
    document.styles.modify("E1 Callout", color="7030A0")
    paragraph = text_of(document)
    paragraph.style = "E1 Callout"
    identifier = document.paragraph(paragraph.id).id

    def check_back(saved):
        style = saved.styles.get("E1 Callout", "paragraph")
        node = next(n for n in saved.package.tree(saved.package.styles_part()).iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style")
            if n.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId") == style.id)
        from docx_agent.edit import formatting

        run = node.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
        return (saved.paragraph(identifier).style == style.id and formatting.read_run(run, "italic")
                and formatting.read_run(run, "color") == "7030A0")

    return check_back


def op_move_block(document):
    paragraphs = plain(document, 4)
    if len(paragraphs) < 2:
        first = text_of(document)
        document.insert_paragraph("A paragraph to move.", after=first.id)
        paragraphs = plain(document, 4)
    mover = next((p for p in reversed(paragraphs[1:]) if not p.ends_section and _survives(p)), None)
    if mover is None:
        mover = paragraphs[0]
    target = paragraphs[0] if mover.id != paragraphs[0].id else paragraphs[1]
    document.move_block(mover.id, before=target.id)
    ids = [document.paragraph(mover.id).id, document.paragraph(target.id).id]

    def check_back(saved):
        order = [p.id for p in saved.paragraphs()]
        return order.index(ids[0]) == order.index(ids[1]) - 1

    return check_back


def _survives(paragraph) -> bool:
    from docx_agent.edit.document import _check_container_survives

    try:
        _check_container_survives(paragraph._element)
    except EditError:
        return False
    return True


OPERATIONS = [op_set_text, op_insert_and_delete_text, op_replace_all, op_replace_across_paragraphs,
              op_format_range, op_format_paragraph, op_paragraph_style, op_character_style,
              op_clear_formatting, op_lists, op_list_continue_and_remove, op_hyperlinks, op_bookmarks,
              op_pictures, op_coalesce, op_styles_add_modify, op_move_block]


@pytest.mark.parametrize("operation", OPERATIONS, ids=lambda f: f.__name__[3:])
def test_operation(docx_path, operation):
    data = docx_path.read_bytes()
    document = Document.open(data)
    mode = document.compatibility_mode
    before = set(check(document.package))
    read_back = operation(document)
    edited = document.to_bytes()

    saved = Document.open(edited)
    assert read_back(saved), f"{operation.__name__} did not read back from the saved document"
    assert saved.compatibility_mode == mode
    assert set(check(saved.package)) - before == set()
    assert numbering_problems(saved) == []

    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)
    while document.redo():
        pass
    assert entries(document.to_bytes()) == entries(edited)
