"""Lists and numbering: list items made, levelled, restarted, continued and removed, with
``numbering.xml`` kept whole -- every numId resolving, no instance or definition orphaned by
an edit, no definition shared by accident."""

from __future__ import annotations

import pytest

from docx_agent import Document, EditError
from docx_agent.edit.numbering import Numbering, problems
from docx_agent.validate import check

from test_roundtrip import entries

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@pytest.fixture
def lists_doc(markup_doc):
    return markup_doc.parent / "lists-and-styles.docx"


def valid(document: Document) -> None:
    assert problems(document) == []
    reopened = Document.open(document.to_bytes())
    assert [p for p in check(reopened.package) if p.code.startswith(("numbering", "style"))] == []


def test_list_membership_is_read(lists_doc):
    document = Document.open(lists_doc)
    first = document.paragraph("p:5A000003").list
    assert (first.num_id, first.level, first.abstract_id, first.kind, first.text) == (1, 0, 0, "number", "%1.")
    assert document.paragraph("p:5A000005").list.format == "lowerLetter"
    bullet = document.paragraph("p:5A000009").list
    assert bullet.from_style and bullet.kind == "bullet"
    assert document.paragraph("p:5A000006").list is None


def test_a_new_list_gets_a_definition_of_its_own(markup_doc):
    data = markup_doc.read_bytes()
    document = Document.open(data)                       # no numbering part at all
    document.paragraph("p:1A2B3C15").add_to_list("number")
    document.paragraph("p:1A2B3C16").add_to_list("number")      # right after: joins the list
    document.paragraph("p:1A2B3C02").add_to_list("bullet")
    first, second = document.paragraph("p:1A2B3C15").list, document.paragraph("p:1A2B3C16").list
    bullet = document.paragraph("p:1A2B3C02").list
    assert first.num_id == second.num_id and first.kind == "number"
    assert bullet.kind == "bullet" and bullet.abstract_id != first.abstract_id
    # Word gives a Normal paragraph it lists the List Paragraph style (measured).
    assert document.paragraph("p:1A2B3C15").style_name == "List Paragraph"
    nsids = [n.get(W + "val") for n in Numbering(document).root.iter(W + "nsid")]
    assert len(nsids) == len(set(nsids))
    valid(document)
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)


def test_not_continuing_starts_a_separate_list(lists_doc):
    document = Document.open(lists_doc)
    document.paragraph("p:5A000006").add_to_list("number", continue_previous=False)
    found = document.paragraph("p:5A000006").list
    assert found.num_id not in (1, 2, 3, 4) and found.abstract_id not in (0, 1)
    valid(document)


def test_join_a_list_by_its_numid(lists_doc):
    document = Document.open(lists_doc)
    document.paragraph("p:5A000006").add_to_list(1, level=1)
    found = document.paragraph("p:5A000006").list
    assert (found.num_id, found.level, found.format) == (1, 1, "lowerLetter")
    with pytest.raises(EditError):
        document.paragraph("p:5A000006").add_to_list(99)


def test_levels(lists_doc):
    document = Document.open(lists_doc)
    paragraph = document.paragraph("p:5A000004")
    paragraph.indent_list()
    assert paragraph.list_level == 1 and paragraph.list.format == "lowerLetter"
    paragraph.outdent_list()
    assert paragraph.list_level == 0
    paragraph.list_level = 8
    assert paragraph.list_level == 8
    with pytest.raises(EditError):
        paragraph.list_level = 9
    with pytest.raises(EditError):
        document.paragraph("p:5A000006").indent_list()
    valid(document)


def test_restart_makes_a_new_instance_over_the_same_definition(lists_doc):
    document = Document.open(lists_doc)
    before = set(Numbering(document).nums())
    document.paragraph("p:5A000004").restart_numbering(at=5)
    found = document.paragraph("p:5A000004").list
    assert found.num_id not in before and found.abstract_id == 0
    assert document.paragraph("p:5A000005").list.num_id == found.num_id   # the items after it follow
    assert document.paragraph("p:5A000003").list.num_id == 1              # the ones before do not
    num = Numbering(document).nums()[found.num_id]
    override = num.find(W + "lvlOverride")
    assert override.get(W + "ilvl") == "0" and override.find(W + "startOverride").get(W + "val") == "5"
    valid(document)


def test_continue_numbering_joins_the_earlier_list_and_reaps_the_old_instance(lists_doc):
    data = lists_doc.read_bytes()
    document = Document.open(data)
    result = document.paragraph("p:5A000007").continue_numbering()
    assert document.paragraph("p:5A000007").list.num_id == 1
    assert document.paragraph("p:5A000008").list.num_id == 1
    assert 2 not in Numbering(document).nums()                            # restarted instance, now unused
    assert "num:2" in result.removed
    assert 4 in Numbering(document).nums()                                # unused before: left alone
    valid(document)
    document.undo()
    assert entries(document.to_bytes()) == entries(data)


def test_continue_from_a_given_paragraph(lists_doc):
    document = Document.open(lists_doc)
    document.paragraph("p:5A000006").add_to_list("number", continue_previous=False)
    document.paragraph("p:5A000006").continue_numbering(from_id="p:5A000003")
    assert document.paragraph("p:5A000006").list.num_id == 1
    valid(document)
    with pytest.raises(EditError):
        document.paragraph("p:5A00000E").continue_numbering()


def test_remove_from_a_list(lists_doc):
    document = Document.open(lists_doc)
    document.paragraph("p:5A000003").remove_from_list()
    paragraph = document.paragraph("p:5A000003")
    assert paragraph.list is None and paragraph.style is None   # List Paragraph undone too
    document.paragraph("p:5A000009").remove_from_list()          # a style's list: numId 0
    numbering = document.paragraph("p:5A000009")._element.find(f"{W}pPr/{W}numPr/{W}numId")
    assert numbering.get(W + "val") == "0" and document.paragraph("p:5A000009").list is None
    assert document.paragraph("p:5A00000A").list is not None
    valid(document)


def test_removing_a_whole_list_reaps_its_instance_and_definition(markup_doc):
    document = Document.open(markup_doc)
    document.paragraph("p:1A2B3C15").add_to_list("bullet")
    document.paragraph("p:1A2B3C15").remove_from_list()
    numbering = Numbering(document)
    assert numbering.nums() == {} and numbering.abstracts() == {}
    valid(document)


def test_changing_a_shared_definition_copies_it_first(lists_doc):
    document = Document.open(lists_doc)
    document.set_list_format("p:5A000007", format="upperRoman", text="%1)")
    restarted = document.paragraph("p:5A000007").list
    assert restarted.format == "upperRoman" and restarted.text == "%1)"
    assert restarted.abstract_id not in (0, 1)
    original = document.paragraph("p:5A000003").list
    assert original.format == "decimal" and original.text == "%1."     # the other list is untouched
    valid(document)
    nsids = [n.get(W + "val") for n in Numbering(document).root.iter(W + "nsid")]
    assert len(nsids) == len(set(nsids))


def test_changing_an_unshared_definition_edits_it_in_place(markup_doc):
    document = Document.open(markup_doc)
    document.paragraph("p:1A2B3C15").add_to_list("bullet")
    abstract = document.paragraph("p:1A2B3C15").list.abstract_id
    document.set_list_format("p:1A2B3C15", text="-", font="Arial")
    found = document.paragraph("p:1A2B3C15").list
    assert found.abstract_id == abstract and found.text == "-"
    valid(document)


def test_list_edits_add_no_problem_on_any_fixture(docx_path):
    document = Document.open(docx_path)
    before = set(check(document.package))
    paragraphs = [p for p in document.paragraphs() if p.text.strip()][:4]
    if len(paragraphs) < 3:
        pytest.skip("too few paragraphs")
    paragraphs[0].add_to_list("number")
    paragraphs[1].add_to_list("number")
    paragraphs[1].indent_list()
    paragraphs[2].add_to_list("bullet")
    paragraphs[1].restart_numbering(3)
    paragraphs[0].remove_from_list()
    assert problems(document) == []
    after = set(check(Document.open(document.to_bytes()).package))
    assert after - before == set()
