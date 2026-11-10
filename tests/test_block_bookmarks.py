"""A bookmark over blocks, tables and content controls included (round 49, R91), written where
Word's Insert > Bookmark writes it for a selection of those blocks (measured in Word for Mac
16.106: the selection's range, Word's saved XML, Word reading our file back to the same
character range, and Word's re-save keeping our markers where they are)."""

from __future__ import annotations

import hashlib

import pytest
from ooxml_edit.tools import Toolbox

import block_stress
from docx_agent import Document, EditError
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.validate import check
from test_block_edits import opened as draft_opened

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

MD = """Opening line.

Second line before the grid.

| Code | Amount |
|---|---|
| K7-1 | 12 |
| K7-2 | 30 |

Third line after the grid.

Fourth line.

| Item | Note |
|---|---|
| one | first |

| Next | Grid |
|---|---|
| two | second |

Closing line."""


def made() -> tuple[Document, dict[str, str]]:
    """The made-up document the measurements used, and its blocks by a short name."""
    document = Document.new()
    document.insert_markdown(MD)
    ids = [b["id"] for b in document.state()["blocks"]]
    names = ["opening", "second", "grid", "third", "fourth", "items", "gap", "next", "closing"]
    return document, dict(zip(names, ids))


def with_controls() -> tuple[Document, dict[str, str]]:
    """A control first in the body before an empty paragraph, one before a table."""
    document = Document.new()
    document.insert_markdown("Gap line.\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\nLast line.")
    blocks = [b["id"] for b in document.state()["blocks"]]
    one = document.insert_control(None, "rich-text", text="Control one.", before=blocks[0], tag="ccone").id
    two = document.insert_control(None, "rich-text", text="Control two.", after=blocks[0], tag="cctwo").id
    document.paragraph(blocks[0]).set_text("")
    return document, {"one": one, "two": two, "gap": blocks[0], "table": blocks[1], "last": blocks[2]}


def body(document: Document) -> list:
    return list(document.package.tree(document.package.document_part()).find(W + "body"))


def element(document: Document, identifier: str):
    return document._resolve_block(identifier)[1].element


def markers(document: Document, name: str = "Span1"):
    part, start, end = document._bookmark_markers()[name]
    return start, end


def first_cell_paragraph(table):
    return next(table.iter(W + "p"))


# -- where the markers go (Word's own placement) ---------------------------------------------


def test_a_span_from_a_paragraph_through_a_table_ends_after_the_table():
    document, ids = made()
    result = document.add_bookmark(f"{ids['second']}..{ids['grid']}", "Span1")
    assert result.created == ["bm:Span1"]
    start, end = markers(document)
    paragraph, table = element(document, ids["second"]), element(document, ids["grid"])
    assert start.getparent() is paragraph and start.getprevious() is None      # before its text
    assert end.getparent() is table.getparent() and end.getprevious() is table  # beside the table
    assert end.getnext() is element(document, ids["third"])


def test_a_span_starting_on_a_table_starts_in_its_first_cell():
    document, ids = made()
    document.add_bookmark(f"{ids['grid']}..{ids['third']}", "Span1")
    start, end = markers(document)
    cell_paragraph = first_cell_paragraph(element(document, ids["grid"]))
    assert start.getparent() is cell_paragraph and start.getnext().tag == W + "r"
    third = element(document, ids["third"])
    assert end.getparent() is third and end.getprevious() is third.findall(W + "r")[-1]  # after its text


def test_a_table_followed_by_no_text_holds_the_end_after_its_last_row():
    """Word: an empty paragraph (or no paragraph) after the table puts the end inside it."""
    document, ids = made()
    document.add_bookmark(f"{ids['fourth']}..{ids['items']}", "Span1")
    _, end = markers(document)
    table = element(document, ids["items"])
    assert end.getparent() is table and end.getprevious() is table.findall(W + "tr")[-1]


def test_a_table_alone_and_two_tables_with_their_separator():
    document, ids = made()
    document.add_bookmark(ids["grid"], "Span1")
    start, end = markers(document)
    table = element(document, ids["grid"])
    assert start.getparent() is first_cell_paragraph(table) and end.getprevious() is table
    document.add_bookmark(f"{ids['items']}..{ids['next']}", "Pair")
    start, end = markers(document, "Pair")
    assert start.getparent() is first_cell_paragraph(element(document, ids["items"]))
    assert end.getprevious() is element(document, ids["next"])
    assert document.bookmark("Pair").text == "Item\nNote\none\nfirst\n\nNext\nGrid\ntwo\nsecond"


def test_a_span_of_paragraphs_marks_their_text_as_a_range_of_it_would():
    """Ending on a paragraph: the end after its text (Word's form for a selection to the end of
    its text), so the span writes what p:A@0..p:B@<length> writes."""
    document, ids = made()
    third = document.paragraph(ids["third"]).text
    other = Document.open(document.to_bytes())
    document.add_bookmark(f"{ids['second']}..{ids['third']}", "Span1")
    other.add_bookmark(other.range(f"{ids['second']}@0..{ids['third']}@{len(third)}"), "Span1")
    assert document.to_bytes() == other.to_bytes()


def test_a_content_control_is_marked_from_before_it_to_after_it():
    document, ids = with_controls()
    document.add_bookmark(ids["one"], "Span1")
    start, end = markers(document)
    control = element(document, ids["one"])
    assert start.getnext() is control and start.get(W + "displacedByCustomXml") == "next"
    assert end.getprevious() is control and end.get(W + "displacedByCustomXml") is None
    assert document.bookmark("Span1").text == "Control one."


def test_a_content_control_before_a_table_ends_in_the_tables_first_cell():
    """Word: the control's range ends where the table starts -- in its first cell, before
    its first paragraph."""
    document, ids = with_controls()
    document.add_bookmark(ids["two"], "Span1")
    _, end = markers(document)
    cell = next(element(document, ids["table"]).iter(W + "tc"))
    assert end.getparent() is cell and end.getprevious().tag == W + "tcPr"
    assert document.bookmark("Span1").text == "Control two."
    document.add_bookmark(f"{ids['two']}..{ids['table']}", "Both")
    assert document.bookmark("Both").text == "Control two.\nA\nB\n1\n2"


def test_a_bookmark_over_blocks_is_valid_and_reads_back_as_written():
    document, ids = made()
    before = set(check(document.package))
    document.add_bookmark(f"{ids['second']}..{ids['grid']}", "Span1")
    reopened = Document.open(document.to_bytes())
    assert set(check(reopened.package)) - before == set()
    assert reopened.bookmark("Span1").text == "Second line before the grid.\nCode\nAmount\nK7-1\n12\nK7-2\n30"


# -- the bookmark as an address --------------------------------------------------------------


def test_the_bookmark_addresses_its_blocks():
    document, ids = made()
    document.add_bookmark(f"{ids['second']}..{ids['grid']}", "Span1")
    markdown = document.to_markdown(range="bm:Span1")
    assert "Second line before the grid." in markdown and "| K7-2 | 30 |" in markdown
    assert "Third line" not in markdown
    assert [r.text for r in document.find("K7", within="bm:Span1")] == ["K7", "K7"]
    assert document.block_text("bm:Span1").endswith("K7-2\n30")


def test_a_heading_range_is_bookmarked_as_its_blocks():
    document = Document.new()
    document.insert_markdown("# Part one\n\nBody text.\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n# Part two\n\nMore.")
    head = document.paragraphs()[0].id
    document.add_bookmark(f"h:{head}", "Part1")
    _, end = markers(document, "Part1")
    assert end.getprevious().tag == W + "tbl"           # the range ends on its table
    assert document.bookmark("Part1").text == "Part one\nBody text.\nA\nB\n1\n2"


# -- no-ops, refusals, tracking, undo -----------------------------------------------------------


def test_the_same_span_again_changes_nothing_and_elsewhere_is_refused():
    document, ids = made()
    document.add_bookmark(f"{ids['second']}..{ids['grid']}", "Span1")
    data = document.to_bytes()
    again = document.add_bookmark(f"{ids['second']}..{ids['grid']}", "span1")
    assert not again.changed and again.warnings == ["bm:Span1 marks this text already"]
    assert document.to_bytes() == data
    with pytest.raises(EditError, match="already has a bookmark"):
        document.add_bookmark(ids["items"], "Span1")


def test_a_span_across_containers_or_reversed_is_refused_with_the_syntax():
    document, ids = made()
    inside = document.table(ids["grid"]).cell(1, 0).paragraphs[0].id
    with pytest.raises(EditError, match="a span is two block ids"):
        document.add_bookmark(f"{ids['second']}..{inside}", "Span1")
    with pytest.raises(EditError, match="ends before it starts"):
        document.add_bookmark(f"{ids['third']}..{ids['second']}", "Span1")


def test_tracked_or_not_the_same_markers_and_undo_restores_the_bytes():
    """Word tracks no bookmark: with tracking on, the markers are written alike, no revision."""
    plain, ids = made()
    data = plain.to_bytes()
    tracked = Document.open(data)
    plain.add_bookmark(f"{ids['grid']}..{ids['third']}", "Span1")
    with tracked.tracking(author="A", date="2026-11-10T12:00:00Z"):
        tracked.add_bookmark(f"{ids['grid']}..{ids['third']}", "Span1")
    assert tracked.to_bytes() == plain.to_bytes()
    tracked.undo()
    assert tracked.to_bytes() == data


# -- the tool ------------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def test_word_links_bookmarks_a_span_a_table_and_a_ref(box):
    session = box.session()
    document, ids = made()
    doc = session.open(document.to_bytes(), "spans.docx")
    result = box.dispatch(session, "word_links", {"doc": doc, "action": "add_bookmark",
                                                  "target": f"{ids['second']}..{ids['grid']}", "bookmark": "Span1"})
    assert result.ok and result.summary == "Bookmarked as Span1", result.to_json()
    result = box.dispatch(session, "word_links", {"doc": doc, "action": "add_bookmark", "target": ids["items"],
                                                  "bookmark": "Items"})
    assert result.ok, result.to_json()
    again = box.dispatch(session, "word_links", {"doc": doc, "action": "add_bookmark", "target": ids["items"],
                                                 "bookmark": "items"})
    assert again.ok and again.summary == "No change: bm:Items marks this text already"
    listed = box.dispatch(session, "word_links", {"doc": doc, "action": "list"}).data["bookmarks"]
    assert [b["name"] for b in listed] == ["Span1", "Items"]
    read = box.dispatch(session, "word_read", {"doc": doc, "range": "bm:Span1"})
    assert read.ok and "| K7-2 | 30 |" in read.data and "Second line before the grid." in read.data, read.to_json()


def test_a_hyperlink_over_blocks_is_still_refused_with_the_way_out(box):
    session = box.session()
    document, ids = made()
    doc = session.open(document.to_bytes(), "spans.docx")
    result = box.dispatch(session, "word_links", {"doc": doc, "action": "add_hyperlink",
                                                  "target": f"{ids['second']}..{ids['grid']}",
                                                  "url": "https://example.com"})
    assert not result.ok and "is a span of blocks, not text" in result.error.message
    assert "p:A@3..p:B@8" in result.error.message


def test_a_span_bookmark_on_a_draft_writes_what_it_writes_on_the_document():
    """Round 48's drafts admit a bookmark over blocks: the same result, bytes and ids."""
    out = {}
    for mode in ("off", "always"):
        with draft_opened(mode) as (box, session, doc, layout):
            document = session.entry(doc).document
            table = layout.tables[1]
            before = document.blocks_between(layout.heads[1], table)
            results = []
            for target, name in ((f"{before[1]}..{table}", "a_span"), (layout.tables[2], "a_table")):
                found = box.dispatch(session, "word_links", {"doc": doc, "action": "add_bookmark", "target": target,
                                                             "bookmark": name})
                assert found.ok, found.to_json()
                results.append(block_stress.normal(found))
            out[mode] = (results, hashlib.sha256(document.to_bytes()).hexdigest(), document.id_state(),
                         document.bookmark("a_span").text)
    assert out["off"] == out["always"]


# -- Word (opt in: pytest -m oracle) ---------------------------------------------------------------


@pytest.mark.oracle
def test_word_resaves_our_markers_where_they_are():
    import re
    import zipfile
    import io

    import oracle

    document, ids = made()
    document.add_bookmark(f"{ids['second']}..{ids['grid']}", "Span1")
    document.add_bookmark(f"{ids['fourth']}..{ids['items']}", "Inside")
    document.add_bookmark(f"{ids['grid']}..{ids['third']}", "Other")
    data = document.to_bytes()

    def placement(blob: bytes) -> list[str]:
        xml = zipfile.ZipFile(io.BytesIO(blob)).read("word/document.xml").decode()
        found = re.findall(r"<w:bookmark(Start|End)([^>]*)/>(</?w:\w+)?", xml)   # the marker, what follows it
        return [(kind, re.sub(r' w:id="\d+"', "", attributes), after) for kind, attributes, after in found]

    try:
        with oracle.session() as word:
            saved = word.resave(data, name="r49-block-bookmarks")
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))
    assert saved.ok, saved.detail
    assert placement(saved.path.read_bytes()) == placement(data)
