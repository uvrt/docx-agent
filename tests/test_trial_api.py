"""The API's shape (ROADMAP.md, "Trial findings", 8): what an edit made is one attribute away
(``EditResult.object``), the collections answer as properties and as calls, ``format_cell``
says what ``set_cell`` does, and ``Picture.image`` is a property the old call still reaches."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docx_agent import Document, EditResult, Paragraph, Table

FIXTURES = Path(__file__).parent / "fixtures" / "generated"
PILOT = FIXTURES / "pilot" / "agreement-summary.docx"


def png() -> bytes:
    with zipfile.ZipFile(FIXTURES / "lists-and-styles.docx") as package:
        return package.read("word/media/image1.png")


def test_an_edit_result_gives_the_object_it_made():
    document = Document.open(PILOT)
    paragraph = document.insert_paragraph("New.", after="p:3B212964").object
    assert isinstance(paragraph, Paragraph) and paragraph.text == "New."
    table = document.insert_table(2, 2, after=paragraph.id).object
    assert isinstance(table, Table)
    assert document.insert_picture(f"{paragraph.id}@0", png(), width=20).object.size[0] == 20
    assert document.insert_markdown("More.", at="end").object.text == "More."
    assert document.paragraph("p:3B212964").set_text("Changed.").object.text == "Changed."
    with pytest.raises(AttributeError, match="doc.get"):
        EditResult("p:3B212964").object


def test_collections_answer_as_properties_and_as_calls():
    document = Document.open(PILOT)
    assert [s.name for s in document.stories()] == [s.name for s in document.stories]
    assert [s.name for s in document.styles()] == [s.name for s in document.styles]
    assert document.styles["Normal"].name == "Normal"
    document.insert_markdown("Text.[^1]\n\n[^1]: A note.\n")
    note = document.notes("footnote")[0]
    assert note.paragraph_ids() == note.paragraph_ids


def test_format_cell_is_set_cell_by_its_name():
    document = Document.open(PILOT)
    table = document.insert_table(2, 2, after="p:3B212964").id
    document.format_cell(table, 0, 0, shading="D9E2F3")
    document.table(table).format_cell(0, 1, shading="D9E2F3")
    assert document.table(table).cell(0, 0).text == ""            # formatting only, never text
    cells = [c.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr")
             for c in document.table(table)._entry()[1].element.iter(
                 "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc")]
    assert sum(1 for c in cells if c is not None and len(c.findall(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd"))) == 2


def test_picture_image_is_a_property_and_the_old_call_warns():
    document = Document.open(PILOT)
    picture = document.insert_picture("p:3B212964@0", png(), width=20).object
    assert picture.image.startswith(b"\x89PNG")
    with pytest.deprecated_call():
        assert picture.image() == bytes(picture.image)
