"""A text field whose source holds Markdown, written into the document (round 48).

A dataset's text value may carry its Markdown beside it (ooxml-edit's
``Dataset.markdown``): a workbook cell's bold and italic runs and its list lines, which the
value itself holds as characters (``•`` bullets, indentation, line breaks).  Written as
blocks -- ``word_insert_markdown``'s ``from`` naming the one field, or a reference standing
alone between blank lines in ``markdown`` -- that Markdown is what is written: list items
in the template's list styles at their levels, emphasis in its character styles.  A
reference inside a sentence writes the value's text on one line, every character as text.
A value without Markdown is written as before; a table cell from rows takes the value.

Tool design rules: 6 (the template's list styles, nested, as Markdown typed in the call
gets them); 4 (a long value counted, not echoed: its Markdown's length).  The data is made
up; the workbook side of the same is xlsx-agent's tests/test_field_markdown.py.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from ooxml_edit.tools import Dataset, Toolbox, tool

from docx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tables_from_data import GRID, _session, call, ok

if not hasattr(Dataset, "markdown_at"):
    pytest.skip("needs ooxml-edit 0.106 (Dataset.markdown)", allow_module_level=True)

#: A cell's text as the workbook holds it, and its Markdown (as xlsx-agent reads it).
LISTED = "Doors need work:\n\n• hinges\n• locks\n    • front\n    • back\n• paint"
LISTED_MD = "Doors need work:\n\n- hinges\n- **locks**\n  - front\n  - *back*\n- paint"
TYPED = "• typed\n    ◦ nested"
TYPED_MD = "- typed\n  - nested"
ROWS = [("Doors", LISTED, 3), ("Roof", TYPED, 1), ("Walls", "Plain - not a list", 2)]


@tool("ds_rich", "Keep a made-up dataset with Markdown.", {}, documents=())
def ds_rich(call):
    call.define_dataset("notes", Dataset(("topic", "body", "score"), ROWS, keys=1,
                                         markdown={(0, 1): LISTED_MD, (1, 1): TYPED_MD}),
                        source=None)


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_rich], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def session(toolbox):
    found = _session(toolbox)
    ok(call(toolbox, found, "ds_rich"))
    return found


def insert(toolbox, session, **arguments):
    return ok(call(toolbox, session, "word_insert_markdown", **{"doc": "d1", "at": "end", **arguments}))


def made(session, result):
    document = session.entry("d1").document
    return [(document.paragraph(b).style, document.paragraph(b).text) for b in result.data["blocks"]]


def typed(toolbox, session, markdown):
    """What the same Markdown typed in the call makes, in a fresh session."""
    other = _session(toolbox)
    result = ok(call(toolbox, other, "word_insert_markdown", doc="d1", at="end", markdown=markdown))
    return made(other, result)


def test_from_writes_the_fields_markdown_as_list_items_nested(toolbox, session):
    result = insert(toolbox, session, **{"from": "d:notes/#1/body"})
    assert made(session, result) == [
        (None, "Doors need work:"), ("ListBullet", "hinges"), ("ListBullet", "locks"),
        ("ListBullet2", "front"), ("ListBullet2", "back"), ("ListBullet", "paint")]
    assert made(session, result) == typed(toolbox, session, LISTED_MD)
    assert result.data["facts"] == [{"from": "d:notes/#1/body", "chars": len(LISTED_MD),
                                     "blocks": result.data["blocks"]}]
    data = session.entry("d1").document.to_bytes()
    xml = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml").decode()
    assert 'w:rStyle w:val="Strong"' in xml and 'w:rStyle w:val="Emphasis"' in xml


def test_a_reference_standing_alone_writes_its_markdown(toolbox, session):
    result = insert(toolbox, session, markdown="## Roof\n\n{{d:notes/Roof/body}}\n\nAfter.")
    assert made(session, result) == [("Heading2", "Roof"), ("ListBullet", "typed"),
                                     ("ListBullet2", "nested"), (None, "After.")]
    facts = result.data["facts"]
    assert facts[0]["from"] == "d:notes/Roof/body" and facts[0]["blocks"] == result.data["blocks"][1:3]


def test_a_reference_in_a_sentence_writes_the_text_on_one_line(toolbox, session):
    """Inline, the value's text: its lines joined by spaces, each character as text (a
    ``•`` stays a bullet character, nothing becomes a list or emphasis)."""
    result = insert(toolbox, session, markdown="See {{d:notes/Roof/body}} first.")
    assert made(session, result) == [(None, "See • typed ◦ nested first.")]
    assert result.data["facts"][0]["text"] == "• typed ◦ nested"
    result = insert(toolbox, session, markdown="{{d:notes/Doors/body}} -- as listed.")
    assert made(session, result) == [
        (None, "Doors need work: • hinges • locks • front • back • paint -- as listed.")]


def test_a_value_without_markdown_is_written_as_before(toolbox, session):
    result = insert(toolbox, session, **{"from": "d:notes/Walls/body"})
    assert made(session, result) == [(None, "Plain - not a list")]
    result = insert(toolbox, session, markdown="{{d:notes/Walls/body}}\n\nand {{d:notes/Walls/score}} more")
    assert made(session, result) == [(None, "Plain - not a list"), (None, "and 2 more")]


def test_a_table_cell_takes_the_value_itself(toolbox, session):
    """Rows placed as a table keep each value's text (its lines in the cell, line breaks
    between them), not Markdown."""
    result = ok(call(toolbox, session, "word_insert_markdown", doc="d1", at="end",
                     **{"from": "d:notes", "from_columns": ["topic", "body"]}))
    table = next(b for b in result.data["blocks"] if b.startswith("t:"))
    found = session.entry("d1").document.table(table)
    assert [found.cell(1, c).text for c in (0, 1)] == ["Doors", LISTED.replace("\n", "\x0b")]
    assert [found.cell(2, c).text for c in (0, 1)] == ["Roof", TYPED.replace("\n", "\x0b")]
