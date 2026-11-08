"""Round 42: ``word_insert_markdown`` at ``replace:p:A..p:B`` gives the new paragraphs the
replaced ones' paragraph settings -- style where Markdown names none, keeps, page breaks,
indents, list membership and level -- matched by kind and position, Markdown's word winning,
and says what was kept.  Before, the new paragraphs had the style map's defaults only.
Made-up documents, Document.new()'s package."""

from __future__ import annotations

import re

import pytest
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

MD = """# Plan

Intro text.

First body.

Second body.

- one
- two
    - two a
- three

After.
"""


def made() -> Document:
    document = Document.new()
    document.insert_markdown(MD, at="end")
    document.styles.add("Report Body", "paragraph", based_on="Normal", space_after=9)
    document.styles.add("Chapter Head", "paragraph", based_on="heading 2", color="1F4E79")
    ps = document.paragraphs()
    ps[2].format(keep_with_next=True, keep_together=True, page_break_before=True, indent_left=18)
    ps[3].set_style("Report Body")
    ps[3].format(keep_together=True)
    for item in ps[4:8]:
        item.format(keep_with_next=True)
    return document


def ids(document):
    return {p.text: p.id for p in document.paragraphs()}


def has(document, identifier, tag):
    element = document._resolve(identifier)[1].element
    return element.find(f"{W}pPr/{W}{tag}") is not None


def by_text(document, text):
    return next(p for p in document.paragraphs() if p.text == text)


def test_body_paragraphs_keep_their_style_keeps_page_break_and_indent():
    document = made()
    old = ids(document)
    edit = document.insert_markdown("New first.\n\nNew second.", at=f"replace:{old['First body.']}..{old['Second body.']}")
    first, second = by_text(document, "New first."), by_text(document, "New second.")
    for tag in ("keepNext", "keepLines", "pageBreakBefore", "ind"):
        assert has(document, first.id, tag), tag
    assert second.style_name == "Report Body" and has(document, second.id, "keepLines")
    assert not has(document, second.id, "pageBreakBefore")
    assert edit.kept == [
        {"block": first.id, "from": old["First body."],
         "settings": ["keep_with_next", "keep_lines", "page_break_before", "indent"]},
        {"block": second.id, "from": old["Second body."], "settings": ["style Report Body", "keep_lines"]}]


def test_what_was_lost_before_keep_settings_false():
    document = made()
    old = ids(document)
    edit = document.insert_markdown("New first.\n\nNew second.", at=f"replace:{old['First body.']}..{old['Second body.']}",
                                    keep_settings=False)
    first, second = by_text(document, "New first."), by_text(document, "New second.")
    assert edit.kept == [] and second.style_name == "Normal"
    assert not any(has(document, first.id, tag) for tag in ("keepNext", "keepLines", "pageBreakBefore", "ind"))


def test_list_items_join_the_replaced_list_at_their_levels():
    document = made()
    old = ids(document)
    nested_before = document.paragraph(old["two a"]).list
    document.insert_markdown("- uno\n- dos\n    - dos a\n- tres\n- cuatro", at=f"replace:{old['one']}..{old['three']}")
    items = [by_text(document, t) for t in ("uno", "dos", "dos a", "tres", "cuatro")]
    assert [p.style_name for p in items] == ["List Bullet", "List Bullet", "List Bullet 2", "List Bullet", "List Bullet"]
    # The nested item is in the very list (w:num) the replaced one was in, at its level.
    nested = items[2].list
    assert (nested.num_id, nested.level) == (nested_before.num_id, nested_before.level)
    assert all(has(document, p.id, "keepNext") for p in items)
    # The instance the Markdown list would have made is gone: nothing names it.
    from docx_agent.edit import numbering
    assert set(numbering.Numbering(document).nums()) <= numbering.referenced_nums(document)


def test_a_flat_list_replacing_nested_items_stays_at_their_level():
    document = made()
    old = ids(document)
    level = document.paragraph(old["two a"]).list.level
    document.insert_markdown("- deep one\n- deep two", at=f"replace:{old['two a']}")
    for text in ("deep one", "deep two"):
        assert by_text(document, text).list.level == level and by_text(document, text).style_name == "List Bullet 2"


def test_markdown_wins_on_kind_level_start_and_named_styles():
    document = made()
    old = ids(document)
    # A numbered list replacing bullets: Markdown says numbered, nothing kept for it.
    edit = document.insert_markdown("1. first\n2. second", at=f"replace:{old['one']}..{old['two']}")
    assert by_text(document, "first").list.kind == "number" and edit.kept == []
    document = made()
    old = ids(document)
    # The style map names the paragraph style: Markdown's style, the keeps kept.
    edit = document.insert_markdown("Mapped.", at=f"replace:{old['First body.']}", style_map={"paragraph": "Report Body"})
    mapped = by_text(document, "Mapped.")
    assert mapped.style_name == "Report Body" and has(document, mapped.id, "keepNext")
    assert not has(document, mapped.id, "ind")
    assert edit.kept[0]["settings"] == ["keep_with_next", "keep_lines", "page_break_before"]


def test_headings_keep_a_same_level_heading_style():
    document = made()
    plan = ids(document)["Plan"]
    document.paragraph(plan).set_style("Chapter Head")      # an outline level 2 style
    document.insert_markdown("## Renamed", at=f"replace:{plan}")
    assert by_text(document, "Renamed").style_name == "Chapter Head"
    document = made()
    plan = ids(document)["Plan"]
    document.paragraph(plan).set_style("Chapter Head")
    document.insert_markdown("# Top", at=f"replace:{plan}")     # another level: Markdown's
    assert by_text(document, "Top").style_name == "heading 1"


def test_more_new_paragraphs_take_the_last_donor_without_a_second_page_break():
    document = made()
    old = ids(document)
    document.insert_markdown("A.\n\nB.\n\nC.", at=f"replace:{old['First body.']}")
    a, b, c = (by_text(document, t) for t in ("A.", "B.", "C."))
    assert has(document, a.id, "pageBreakBefore")
    assert not has(document, b.id, "pageBreakBefore") and not has(document, c.id, "pageBreakBefore")
    assert all(has(document, p.id, "keepLines") for p in (a, b, c))


def test_tracked_replacement_keeps_settings_and_stays_valid():
    document = made()
    old = ids(document)
    with document.tracking("Writer"):
        document.insert_markdown("Tracked one.\n\n- tracked item", at=f"replace:{old['Second body.']}..{old['one']}")
    assert by_text(document, "Tracked one.").style_name == "Report Body"
    item = by_text(document, "tracked item")
    assert has(document, item.id, "keepNext") and item.style_name == "List Bullet"
    assert document.validate() == []


def test_the_tool_says_what_was_kept():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session()
        doc = session.open(made().to_bytes(), "keeps.docx")
        document = session.entry(doc).document
        old = ids(document)
        result = toolbox.dispatch(session, "word_insert_markdown",
                                  {"doc": doc, "markdown": "- uno\n- dos", "at": f"replace:{old['one']}..{old['two']}"})
        assert result.ok
        kept = result.data["kept"]
        assert [k["block"] for k in kept] == result.data["blocks"]
        assert kept[0]["settings"] == ["style List Bullet", "keep_with_next", "list level 0"]
        assert "kept the replaced paragraphs' settings" in result.summary
        # Inserting elsewhere keeps nothing and says nothing.
        result = toolbox.dispatch(session, "word_insert_markdown", {"doc": doc, "markdown": "Tail.", "at": "end"})
        assert result.ok and "kept" not in result.data


def test_the_replacement_lays_out_as_the_text_rewritten_in_place():
    """The layout check: replacing paragraphs with Markdown lays out as rewriting their text
    in place (which keeps every setting) does."""
    replaced = made()
    old = ids(replaced)
    replaced.insert_markdown("New first.\n\nNew second.\n\n- uno\n- dos\n    - dos a\n- tres",
                             at=f"replace:{old['First body.']}..{old['three']}")
    hand = made()
    old = ids(hand)
    for before, after in (("First body.", "New first."), ("Second body.", "New second."), ("one", "uno"),
                          ("two", "dos"), ("two a", "dos a"), ("three", "tres")):
        hand.paragraph(old[before]).set_text(after)
    assert _laid_out(replaced) == _laid_out(hand)


def _laid_out(document) -> list[str]:
    """The pages without the paragraphs' ids (new paragraphs have new ones)."""
    return [re.sub(r' data-docx-(?:agent-)?id="[^"]*"', "", page) for page in document.render_svg()]


@pytest.mark.parametrize("at", ["end", "after:{p}", "before:{p}"])
def test_other_places_are_written_as_before(at):
    document = made()
    p = ids(document)["First body."]
    edit = document.insert_markdown("Placed.", at=at.format(p=p))
    assert edit.kept == [] and by_text(document, "Placed.").style_name == "Normal"
