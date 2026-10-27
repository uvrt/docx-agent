"""Empty paragraphs between text (production feedback round 18, Q4): ``check`` with
``include: ["layout"]`` lists in ``layout.empty_paragraphs``, per table cell and for the
body, the blank lines left between two paragraphs of text, with the ``word_delete`` calls
that take them out; applying them (tracked too) and checking again finds nothing, and the
text is unchanged."""

from __future__ import annotations

import json

import pytest
from lxml import etree

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.layout_facts import _state, empty_paragraphs  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, json.dumps(result.to_json(), indent=1)
    return result


def _document() -> tuple[Document, str, list[str]]:
    """A made-up document: an opening line, then two blank lines (one a space only) before
    the next, a table whose cell (1, 0) reads Title, a blank line, Body text, and a closing
    line; an empty paragraph also opens the document and ends the cell (1, 1)."""
    document = Document.new()
    document.insert_markdown("Opening line.\n\nNext line.\n\n| Item | Note |\n| --- | --- |\n"
                             "| x | y |\n\nClosing line.", at="end")
    opening = document._index(document.package.document_part()).paragraphs[0].id
    blanks = [document.paragraph(opening).insert_after(" ").id]
    blanks.insert(0, document.paragraph(opening).insert_after("").id)
    document.paragraph(opening).insert_before("")
    table = document.tables()[0]
    cell = table.cell(1, 0).paragraphs[0]
    cell.insert_after("Body text")
    blanks.append(cell.insert_after("").id)
    cell.set_text("Title")
    table.cell(1, 1).paragraphs[0].insert_after("")
    return document, table.id, blanks


def _texts(document: Document) -> list[str]:
    """Every paragraph's text, empty ones left out, in order."""
    index = document._index(document.package.document_part())
    return [text for entry in index.paragraphs
            if (text := document.paragraph(entry.id).text).strip()]


def _found(box, session, doc: str, **arguments) -> list[dict]:
    data = _ok(box, session, "check", doc=doc, include=["layout"], **arguments).data
    return data["layout"]["empty_paragraphs"]


def test_the_repro_lists_the_cell_and_the_body_with_their_paragraphs(box):
    document, table, blanks = _document()
    session = box.session()
    doc = session.open(document.to_bytes(), "notes.docx")
    found = {entry["at"]: entry for entry in _found(box, session, doc)}
    assert set(found) == {"body", f"{table}/c1,0"}
    assert found[f"{table}/c1,0"] == {
        "kind": "empty_paragraphs", "at": f"{table}/c1,0", "paragraphs": [1], "ids": [blanks[2]],
        "fix": {"tool": "word_delete", "calls": [{"target": blanks[2]}]}}
    # Body: positions among its paragraphs (the empty first one is 0, never listed: it
    # leads); the two blanks are siblings, so one span deletes them.
    body = found["body"]
    assert body["paragraphs"] == [2, 3] and body["ids"] == blanks[:2]
    assert body["fix"]["calls"] == [{"target": f"{blanks[0]}..{blanks[1]}"}]


@pytest.mark.parametrize("tracked", [False, True])
def test_applying_the_fixes_leaves_nothing_and_the_text_unchanged(box, tracked):
    document, table, blanks = _document()
    before = _texts(document)
    session = box.session()
    doc = session.open(document.to_bytes(), "notes.docx")
    if tracked:
        _ok(box, session, "word_set_tracking", doc=doc, on=True, author="Reviewer")
    for entry in _found(box, session, doc):
        for call in entry["fix"]["calls"]:
            _ok(box, session, entry["fix"]["tool"], doc=doc, **call)
    assert _found(box, session, doc) == []
    edited = session.entry(doc).document
    assert _texts(edited) == before
    if tracked:                          # deleted as tracked changes, accepted they are gone
        edited.accept_all()
    gone = set(blanks)
    assert not gone & {e.id for e in edited._index(edited.package.document_part()).paragraphs}


def test_only_on_request(box):
    document, _, _ = _document()
    session = box.session()
    doc = session.open(document.to_bytes(), "notes.docx")
    assert "layout" not in _ok(box, session, "check", doc=doc).data


def test_a_table_a_section_break_or_a_page_break_before_ends_a_run_unlisted():
    document = Document.new()
    document.insert_markdown("One.\n\nTwo.\n\nThree.\n\n| A |\n| --- |\n| b |\n\nFour.", at="end")
    paragraphs = document._index(document.package.document_part()).paragraphs
    one, two, three = (paragraphs[k].id for k in range(3))
    four = paragraphs[-1].id
    # Empty, then a page break before, then empty: between One and Two.
    document.paragraph(one).insert_after("")
    breaking = document.paragraph(one).insert_after("").id
    document.paragraph(breaking).page_break_before = True
    document.paragraph(one).insert_after("")
    # Empty between Three and the table, and between the table and Four.
    document.paragraph(three).insert_after("")
    document.paragraph(four).insert_before("")
    # An empty paragraph carrying a section break, between Two and Three.
    sectioned = document.paragraph(two).insert_after("").id
    _, entry = document._resolve(sectioned)
    properties = entry.element.find(f"{{{W}}}pPr")
    if properties is None:
        properties = etree.SubElement(entry.element, f"{{{W}}}pPr")
        entry.element.insert(0, properties)
    etree.SubElement(properties, f"{{{W}}}sectPr")
    assert empty_paragraphs(document) == []


def test_two_runs_in_one_cell_are_one_entry_the_last_deleted_first():
    document = Document.new()
    document.insert_markdown("| A |\n| --- |\n| b |\n\nAfter.", at="end")
    table = document.tables()[0]
    cell = table.cell(1, 0)
    first = cell.paragraphs[0]
    third = first.insert_after("Third")
    second_blank = first.insert_after("")
    second = first.insert_after("Second")
    first_blank = first.insert_after("")
    assert [p.text for p in table.cell(1, 0).paragraphs] == ["b", "", "Second", "", "Third"]
    assert third and second
    (entry,) = empty_paragraphs(document)
    assert entry["at"] == f"{table.id}/c1,0" and entry["paragraphs"] == [1, 3]
    assert entry["fix"]["calls"] == [{"target": second_blank.id}, {"target": first_blank.id}]


@pytest.mark.parametrize("inner, state", [
    ("", "empty"),
    ("<w:r><w:t xml:space='preserve'>  </w:t></w:r>", "empty"),
    ("<w:bookmarkStart w:id='0' w:name='here'/><w:bookmarkEnd w:id='0'/>", "empty"),
    ("<w:r><w:t>x</w:t></w:r>", "text"),
    ("<w:r><w:tab/></w:r>", "text"),
    ("<w:r><w:br/></w:r>", "text"),
    ("<w:r><w:drawing/></w:r>", "text"),
    ("<w:fldSimple w:instr='PAGE'/>", "text"),
    ("<w:r><w:footnoteReference w:id='1'/></w:r>", "text"),
    ("<w:r><w:commentReference w:id='1'/></w:r>", "text"),
    ("<w:sdt><w:sdtContent/></w:sdt>", "text"),
    ("<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'/>", "text"),
    ("<w:del w:id='1' w:author='A'><w:r><w:delText>Old</w:delText></w:r></w:del>", "text"),
    ("<w:pPr><w:tabs><w:tab w:val='left' w:pos='720'/></w:tabs></w:pPr>", "empty"),
    ("<w:pPr><w:framePr w:w='2000'/></w:pPr>", "kept"),
    ("<w:pPr><w:rPr><w:del w:id='1' w:author='A'/></w:rPr></w:pPr>", None),
])
def test_what_makes_a_paragraph_empty(inner, state):
    paragraph = etree.fromstring(f"<w:p xmlns:w='{W}'>{inner}</w:p>")
    assert _state(Document.new(), paragraph) == state
