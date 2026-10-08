"""Hyperlinks (external and to bookmarks) with their relationships kept, bookmarks by name,
and cross-reference fields."""

from __future__ import annotations

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.validate import check

from test_roundtrip import entries

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


@pytest.fixture
def lists_doc(markup_doc):
    return markup_doc.parent / "lists-and-styles.docx"


def hyperlink_relationships(document: Document, part: str = "word/document.xml") -> dict[str, str]:
    return {rel.id: rel.target for rel in document.package.relationships(part).values() if rel.type == HYPERLINK}


def valid(document: Document, before: set) -> None:
    assert set(check(Document.open(document.to_bytes()).package)) - before == set()


# -- hyperlinks -------------------------------------------------------------------------------


def test_hyperlinks_are_read(lists_doc):
    document = Document.open(lists_doc)
    links = document.hyperlinks()
    assert [(link.text, link.address, link.anchor) for link in links] == [
        ("the website", "https://example.com/report", None), ("the target", None, "target")]
    assert links[0].id == "hl:p:5A00000B/0"
    assert document.hyperlink("hl:p:5A00000B/1").range().text == "the target"


def test_add_an_external_hyperlink(lists_doc):
    data = lists_doc.read_bytes()
    document = Document.open(data)
    before = set(check(document.package))
    result = document.anchor("Between the lists").add_hyperlink("https://example.org/a", tooltip="Example")
    link = document.hyperlink(result.id)
    assert (link.text, link.address, link.tooltip) == ("Between the lists", "https://example.org/a", "Example")
    run = link._locate()[2].find(W + "r")
    assert run.find(f"{W}rPr/{W}rStyle").get(W + "val") == "Hyperlink"
    assert "https://example.org/a" in hyperlink_relationships(document).values()
    valid(document, before)
    document.undo()
    assert entries(document.to_bytes()) == entries(data)


def test_a_second_link_to_the_same_address_shares_its_relationship(lists_doc):
    document = Document.open(lists_doc)
    document.anchor("Between").add_hyperlink("https://example.com/report")
    assert list(hyperlink_relationships(document).values()).count("https://example.com/report") == 1


def test_an_internal_hyperlink_needs_its_bookmark(lists_doc):
    document = Document.open(lists_doc)
    result = document.anchor("the lists").add_hyperlink(anchor="target")
    assert document.hyperlink(result.id).anchor == "target"
    with pytest.raises(KeyError):
        document.anchor("Between").add_hyperlink(anchor="missing")


def test_hyperlinks_are_refused_where_they_cannot_go(lists_doc, markup_doc):
    document = Document.open(lists_doc)
    with pytest.raises(EditError, match="already"):
        document.anchor("website").add_hyperlink("https://example.org/")
    with pytest.raises(EditError, match="scheme"):
        document.anchor("Between").add_hyperlink("example.org")
    with pytest.raises(EditError):
        document.anchor("Between").add_hyperlink("https://example.org/", anchor="target")
    other = Document.open(markup_doc)
    with pytest.raises(EditError, match="field"):
        other.anchor("Someone").add_hyperlink("https://example.org/")


def test_change_a_hyperlinks_target(lists_doc):
    document = Document.open(lists_doc)
    link = document.hyperlinks()[0]
    link.set_target("https://example.net/")
    assert link.address == "https://example.net/"
    assert "https://example.com/report" not in hyperlink_relationships(document).values()   # released
    link.set_target(anchor="target")
    assert link.address is None and link.anchor == "target"
    assert hyperlink_relationships(document) == {}
    valid(document, set())


def test_remove_a_hyperlink(lists_doc):
    data = lists_doc.read_bytes()
    document = Document.open(data)
    result = document.hyperlinks()[0].remove()
    paragraph = document.paragraph("p:5A00000B")
    assert paragraph.text == "See the website or the target."
    assert len(paragraph.hyperlinks) == 1
    assert "rStyle" not in etree.tostring(paragraph.runs[0]._element).decode()
    assert hyperlink_relationships(document) == {}
    assert result.removed == ["hl:p:5A00000B/0"]
    valid(document, set())
    document.undo()
    assert entries(document.to_bytes()) == entries(data)


def test_a_hyperlink_in_a_header_uses_the_headers_relationships(markup_doc):
    document = Document.open(markup_doc)
    document.anchor("header").add_hyperlink("https://example.org/h")
    assert "https://example.org/h" in hyperlink_relationships(document, "word/header1.xml").values()
    assert "https://example.org/h" not in hyperlink_relationships(document).values()


# -- bookmarks --------------------------------------------------------------------------------


def test_bookmarks_are_read_and_resolve_to_ranges(markup_doc, lists_doc):
    document = Document.open(markup_doc)
    span = document.bookmark("span")
    assert span.id == "bm:span" and document.bookmark("bm:SPAN").name == "span"
    assert span.range().text == "A bookmark starts here\nand ends here."
    assert Document.open(lists_doc).bookmark("target").text == "The target paragraph"


def test_add_a_bookmark(lists_doc):
    data = lists_doc.read_bytes()
    document = Document.open(data)
    result = document.anchor("Second item").add_bookmark("Second")
    assert result.id == "bm:Second"
    assert document.bookmark("Second").text == "Second item"
    marks = [int(n.get(W + "id")) for n in document.package.tree("word/document.xml").iter(W + "bookmarkStart")]
    assert len(marks) == len(set(marks)) and max(marks) == 1
    across = document.range("p:5A000003@6..p:5A000004@6")
    across.add_bookmark("Across")
    assert document.bookmark("Across").text == "item\nSecond"
    point = document.range("p:5A000006@3")
    point.add_bookmark("Here")
    assert document.bookmark("Here").range().collapsed
    valid(document, set())
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)


@pytest.mark.parametrize("name", ["1st", "_hidden", "has space", "x" * 41, "target", "TARGET"])
def test_bad_or_taken_bookmark_names_are_refused(lists_doc, name):
    document = Document.open(lists_doc)
    with pytest.raises(EditError):
        document.anchor("Between").add_bookmark(name)


def test_rename_a_bookmark_and_what_points_at_it(lists_doc):
    document = Document.open(lists_doc)
    document.insert_cross_reference("p:5A00000E@0", "target")
    result = document.bookmark("target").rename("Goal")
    assert result.renamed == {"bm:target": "bm:Goal"}
    assert document.hyperlinks()[1].anchor == "Goal"
    instructions = [n.text for n in document.package.tree("word/document.xml").iter(W + "instrText")]
    assert instructions == [" REF Goal \\h "]
    with pytest.raises(KeyError):
        document.bookmark("target")
    valid(document, set())


def test_remove_a_bookmark_says_what_still_points_at_it(lists_doc):
    document = Document.open(lists_doc)
    result = document.bookmark("target").remove()
    assert result.removed == ["bm:target"]
    assert result.warnings and "hyperlink" in result.warnings[0]
    assert document.bookmarks() == []
    assert document.paragraph("p:5A00000C").text == "The target paragraph ends here."


# -- cross-references -------------------------------------------------------------------------


def test_cross_references(lists_doc):
    document = Document.open(lists_doc)
    document.insert_cross_reference("p:5A00000E@0", "target")
    document.insert_cross_reference(document.paragraph("p:5A00000E").range(0, 0), "target", kind="page",
                                    hyperlink=False)
    paragraph = document.paragraph("p:5A00000E")
    assert paragraph.text == "1The target paragraphThe last paragraph."
    instructions = [n.text for n in paragraph._element.iter(W + "instrText")]
    assert instructions == [" PAGEREF target ", " REF target \\h "]
    valid(document, set())
    with pytest.raises(EditError):
        document.insert_cross_reference("p:5A00000E@0", "target", kind="footnote")
