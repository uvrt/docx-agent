"""Ids: every paragraph has one, they are unique, and they hold across edits, undo and redo.

ROADMAP.md, "Addressing and stable ids", is the specification these tests hold.
"""

from __future__ import annotations

import re

import pytest

from docx_agent import Document
from docx_agent.edit import ids as _ids
from docx_agent.edit.ids import PARA_ID, TEXT_ID, valid_long_hex

from test_roundtrip import entries


def all_ids(document: Document) -> list[str]:
    return [p.id for story in document.stories for p in story.paragraphs]


def test_every_paragraph_has_a_unique_id(docx_path):
    document = Document.open(docx_path)
    ids = all_ids(document)
    assert len(ids) == len(set(ids))
    for story in document.stories:
        for paragraph in story.paragraphs:
            assert document.paragraph(paragraph.id)._element is paragraph._element


def test_id_forms(markup_doc):
    document = Document.open(markup_doc)
    ids = [p.id for p in document.paragraphs()]
    assert ids[0] == "p:1A2B3C01"
    assert ids[2] == "p@body/2"                       # no paraId: positional
    assert ids[3:6] == ["p:2B3C4D05", "p:2B3C4D05#1", "p:2B3C4D05#2"]
    assert ids[6] == "p:80000001"                     # out of range, still verbatim
    assert document.paragraphs("header1")[0].id == "header1/p:4D5E6F01"
    assert document.paragraphs("footnotes")[2].id == "footnotes/p:5E6F7003"
    table = document.tables()[0]
    assert table.id == "t:3C4D5E01"
    assert [row.id for row in table.rows] == ["tr:3C4D5E01", "tr:3C4D5E04"]
    assert table.cell(1, 0).id == "t:3C4D5E01/c0,0"   # a merged cell answers to its origin
    assert document.get("t:3C4D5E01/c1,2").text == "C2"
    assert document.get("tr:3C4D5E04").id == "tr:3C4D5E04"
    assert document.get("p:1A2B3C02/r1").text == "bold"
    assert [s.id for s in document.sections()] == ["s:body"]


def test_tables_without_row_ids_are_positional(docx_path):
    document = Document.open(docx_path)
    for number, table in enumerate(document.tables()):
        assert re.fullmatch(r"t(:[0-9A-F]{8}|@body/%d)" % number, table.id)
        for k, row in enumerate(table.rows):
            assert row.id.startswith("tr:") or row.id == f"{table.id}/r{k}"


def test_reading_never_stamps(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    all_ids(document)
    assert entries(document.to_bytes()) == entries(data)


@pytest.mark.parametrize("stamping", ["document", "paragraph"])
def test_first_edit_stamps_a_paragraph_without_a_paraid(markup_doc, stamping):
    document = Document.open(markup_doc, stamping=stamping)
    result = document.paragraph("p@body/2").set_text("Now it has one.")
    new = result.id
    assert result.renamed["p@body/2"] == new
    paragraph = document.paragraph(new)
    assert valid_long_hex(paragraph.para_id) and valid_long_hex(paragraph.text_id)
    assert paragraph.para_id not in {p.para_id for p in document.paragraphs() if p.id != new}
    assert document.paragraph("p@body/2").text == "Now it has one."   # the alias resolves


def test_stamping_is_deterministic(docx_path):
    def edit() -> bytes:
        document = Document.open(docx_path)
        paragraphs = document.paragraphs()
        if not paragraphs:
            pytest.skip("no paragraphs")
        paragraphs[-1].set_text("Edited.")
        document.insert_paragraph("Inserted.", after=paragraphs[0].id)
        return document.to_bytes()

    assert entries(edit()) == entries(edit())


def test_stamping_then_undo_gives_the_original_bytes(docx_path):
    data = docx_path.read_bytes()
    document = Document.open(data)
    paragraphs = document.paragraphs()
    if not paragraphs:
        pytest.skip("no paragraphs")
    paragraphs[0].set_text("Stamped and changed.")
    assert entries(document.to_bytes()) != entries(data)
    document.undo()
    assert entries(document.to_bytes()) == entries(data)


def test_w14_is_declared_and_ignorable_after_stamping(docx_path):
    document = Document.open(docx_path)
    paragraphs = document.paragraphs()
    if not paragraphs:
        pytest.skip("no paragraphs")
    paragraphs[0].set_text("Stamped.")
    root = document.package.tree(document.package.document_part())
    prefix = next(p for p, uri in root.nsmap.items() if uri == PARA_ID[1:].partition("}")[0])
    assert prefix == "w14" or prefix in root.get(
        "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable").split()
    ignorable = root.get("{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable")
    assert prefix in ignorable.split()
    head = document.package.read(document.package.document_part())[:2000]
    assert b"xmlns:w14=" in head and b"ns0" not in head and b"ns1" not in head


@pytest.mark.parametrize("stamping", ["document", "paragraph"])
def test_a_repeat_is_reissued_on_its_first_edit_with_the_later_repeats(markup_doc, stamping):
    document = Document.open(markup_doc, stamping=stamping)
    third = document.paragraph("p:2B3C4D05#2")
    result = document.paragraph("p:2B3C4D05#1").set_text("Second, edited.")
    assert {"p:2B3C4D05#1", "p:2B3C4D05#2"} <= set(result.renamed)
    if stamping == "paragraph":
        assert set(result.renamed) == {"p:2B3C4D05#1", "p:2B3C4D05#2"}
    assert document.paragraph("p:2B3C4D05").text == "First of three repeats."
    assert document.paragraph("p:2B3C4D05#1").text == "Second, edited."
    assert document.paragraph("p:2B3C4D05#2").text == "Third of three repeats."
    assert third.text == "Third of three repeats."
    raw = [p.para_id for p in document.paragraphs()]
    assert raw.count("2B3C4D05") == 1


def test_an_out_of_range_paraid_is_reissued(markup_doc):
    document = Document.open(markup_doc)
    result = document.paragraph("p:80000001").set_text("Now in range.")
    assert result.id != "p:80000001"
    assert valid_long_hex(document.paragraph(result.id).para_id)
    assert document.paragraph("p:80000001").text == "Now in range."


def test_a_valid_paraid_is_kept_and_its_textid_renewed(markup_doc):
    document = Document.open(markup_doc)
    before = document.paragraph("p:1A2B3C02").text_id
    result = document.paragraph("p:1A2B3C02").set_text("Plain text now.")
    assert result.id == "p:1A2B3C02" and "p:1A2B3C02" not in result.renamed
    after = document.paragraph("p:1A2B3C02").text_id
    assert after != before and valid_long_hex(after)


def test_ids_survive_unrelated_inserts_and_deletes_undo_and_redo(docx_path):
    document = Document.open(docx_path)
    paragraphs = document.paragraphs()
    if len(paragraphs) < 4:
        pytest.skip("too few paragraphs")
    held = {p.id: p.text for p in paragraphs}
    deletable = [p.id for p in paragraphs[1:-1] if _deletable(document, p.id)]
    removed = deletable[len(deletable) // 2] if deletable else None
    first = paragraphs[0].id

    document.insert_paragraph("Inserted at the start.", before=first)
    document.insert_paragraph("Inserted after the first.", after=first)
    if removed:
        document.delete_block(removed)
        held.pop(removed)

    def check() -> None:
        for identifier, text in held.items():
            assert document.paragraph(identifier).text == text, identifier

    check()
    steps = 3 if removed else 2
    for _ in range(steps):
        document.undo()
        check()
    for _ in range(steps):
        document.redo()
        check()


def _deletable(document: Document, identifier: str) -> bool:
    from docx_agent.edit.document import _REFERENCES, W_PPR, W_SECTPR, _fields_balanced

    element = document.paragraph(identifier)._element
    properties = element.find(W_PPR)
    if properties is not None and properties.find(W_SECTPR) is not None:
        return False
    if any(element.find(".//" + tag) is not None for tag in _REFERENCES):
        return False
    return _fields_balanced(element) and element.getparent().tag.endswith("}body")


def test_a_held_view_follows_its_paragraph(markup_doc):
    document = Document.open(markup_doc)
    view = document.paragraph("p@body/2")
    document.insert_paragraph("Pushes everything down.", before="p:1A2B3C01")
    assert view.text == "A paragraph with no paraId."
    view.set_text("Edited through the held view.")
    assert view.text == "Edited through the held view."
    assert not view.volatile


def test_new_paragraphs_have_durable_ids_at_once(markup_doc):
    document = Document.open(markup_doc)
    result = document.insert_paragraph("New.", after="p:1A2B3C01")
    assert result.created == [result.id] and result.id.startswith("p:")
    element = document.paragraph(result.id)._element
    assert valid_long_hex(element.get(PARA_ID)) and valid_long_hex(element.get(TEXT_ID))


def test_ids_stay_unique_after_edits(docx_path):
    document = Document.open(docx_path)
    paragraphs = document.paragraphs()
    if not paragraphs:
        pytest.skip("no paragraphs")
    with document.batch():
        for paragraph in paragraphs[::3]:
            paragraph.set_text(paragraph.text + " +")
        document.insert_paragraph("Added.", after=paragraphs[-1].id)
    ids = all_ids(document)
    assert len(ids) == len(set(ids))
    raw = [p.para_id for p in document.paragraphs() if p.para_id]
    stamped = [p for p in raw if valid_long_hex(p)]
    assert len(stamped) == len(set(stamped)) or len(set(raw)) < len(raw)


def test_document_stamping_leaves_no_paragraph_or_row_without_an_id(docx_path):
    """What Word needs to keep paraIds on saving: every paragraph and row of every story
    with a valid, unique one, every drawing with an anchorId and an editId (ROADMAP.md,
    "Durable across a Word save -- measured")."""
    from docx_agent.validate import check

    document = Document.open(docx_path)
    paragraphs = document.paragraphs()
    if not paragraphs:
        pytest.skip("no paragraphs")
    paragraphs[0].set_text(paragraphs[0].text + " (stamped)")
    for story in document.stories:
        for paragraph in story.paragraphs:
            assert not paragraph.volatile and valid_long_hex(paragraph.para_id), paragraph.id
        for table in story.tables:
            assert table.id.startswith(("t:", f"{story.name}/t:")), table.id
            for row in table.rows:
                assert row.id.rpartition("/")[2].startswith("tr:"), row.id
    assert [p for p in check(document.package) if p.code.startswith(("paraid", "textid"))] == []
    for part in document._parts():
        assert _ids.drawings_needing_ids(document.package.tree(part)) == []


def test_table_and_row_ids_follow_their_stamps(docx_path):
    document = Document.open(docx_path)
    tables = document.tables()
    if not tables:
        pytest.skip("no tables")
    old_table = tables[0].id
    old_rows = [row.id for row in tables[0].rows]
    document.paragraphs()[0].set_text("Stamps the rows too.")
    assert document.get(old_table).id.startswith("t:")
    for old in old_rows:
        assert document.get(old).id.startswith("tr:")
