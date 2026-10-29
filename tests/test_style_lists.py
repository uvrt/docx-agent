"""Lists numbered by their style (round 23, V1): Markdown lists written in a style that numbers
itself take its numbering -- no ``w:numPr``, no indent -- and count on as Word counts them;
``restart`` gives the restarted item alone an instance of its own, Word's form; ``continue``
drops the direct numbering of items whose style numbers them.  What Word does is measured in
``feedback-round23/measurements-F.md`` (made-up documents, the Word oracle)."""

from __future__ import annotations

import warnings

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.edit import numbering as _numbering
from docx_agent.edit.numbering import Numbering, problems

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS


def _xml(text: str):
    return etree.fromstring(f'<w:root xmlns:w="{W_NS}">{text}</w:root>')[0]


def _level(ilvl: int, fmt: str, text: str, left: int, hanging: int, style: str | None) -> str:
    tied = f'<w:pStyle w:val="{style}"/>' if style else ""
    return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>{tied}'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr></w:lvl>')


def _style(style_id: str, name: str, ppr: str, based: str = "Normal") -> str:
    return (f'<w:style w:type="paragraph" w:styleId="{style_id}"><w:name w:val="{name}"/>'
            f'<w:basedOn w:val="{based}"/><w:pPr>{ppr}</w:pPr></w:style>')


def template(*, style_indent: int = 284, extra_styles: str = "") -> Document:
    """A made-up template: List Number numbers itself (level 0 of a multilevel definition, List
    Number 2 tied to its level 1), List Bullet too; Probe Unnumbered is List Number with numId 0."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        document = Document.new()
    _numbering._ensure_part(document)
    _numbering.add_abstract(document, _xml(
        '<w:abstractNum w:abstractNumId="1"><w:nsid w:val="5E1A0001"/><w:multiLevelType w:val="multilevel"/>'
        + _level(0, "decimal", "%1.", 284, 284, "ListNumber") + _level(1, "lowerLetter", "%2)", 568, 284, "ListNumber2")
        + "</w:abstractNum>"))
    _numbering.add_abstract(document, _xml(
        '<w:abstractNum w:abstractNumId="2"><w:nsid w:val="5E1A0002"/><w:multiLevelType w:val="multilevel"/>'
        + _level(0, "bullet", "•", 284, 284, "ListBullet") + "</w:abstractNum>"))
    number = _numbering.add_num(document, 1)
    bullet = _numbering.add_num(document, 2)
    indent = f'<w:ind w:left="{style_indent}" w:hanging="284"/>'
    styles = document.package.tree(document.package.styles_part())
    for node in (_style("ListNumber", "List Number", f'<w:numPr><w:numId w:val="{number}"/></w:numPr>{indent}'),
                 _style("ListNumber2", "List Number 2",
                        f'<w:numPr><w:ilvl w:val="1"/><w:numId w:val="{number}"/></w:numPr><w:ind w:left="568" w:hanging="284"/>'),
                 _style("ListBullet", "List Bullet", f'<w:numPr><w:numId w:val="{bullet}"/></w:numPr>{indent}'),
                 _style("ProbeUnnumbered", "Probe Unnumbered", '<w:numPr><w:numId w:val="0"/></w:numPr>', "ListNumber"),
                 *([_xml(extra_styles)] if extra_styles else [])):
        styles.append(_xml(node) if isinstance(node, str) else node)
    document.package.mark_dirty(document.package.styles_part())
    document.package.mark_dirty(document.package.numbering_part())
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return Document.open(document.to_bytes())


def labels(document: Document, ids: list[str]) -> list[str | None]:
    state = document.state(f"{ids[0]}..{ids[-1]}")
    out = {b["id"]: (b.get("list") or {}).get("label") for b in state["blocks"]}
    return [out.get(i) for i in ids]


def own_numbering(document: Document, identifier: str):
    return document._resolve(identifier)[1].element.find(f"{W}pPr/{W}numPr")


def items(document: Document, result) -> list[str]:
    return [i for i in result.blocks if document.list_of(i) is not None]


# -- writing -------------------------------------------------------------------------------------


def test_a_markdown_list_in_a_style_that_numbers_itself_takes_its_numbering_across_calls():
    document = template()
    nums_before = set(document.numbering.nums())
    first = document.insert_markdown("1. one\n2. two\n")
    document.insert_markdown("Between the lists.\n")
    second = document.insert_markdown("1. three\n")
    ids = first.blocks + second.blocks
    for identifier in ids:
        paragraph = document._resolve(identifier)[1].element
        assert paragraph.find(f"{W}pPr/{W}pStyle").get(W + "val") == "ListNumber"
        assert paragraph.find(f"{W}pPr/{W}numPr") is None and paragraph.find(f"{W}pPr/{W}ind") is None
        assert document.list_of(identifier).from_style
    assert labels(document, ids) == ["1.", "2.", "3."]       # Word: 3 after the interlude (measured)
    assert set(document.numbering.nums()) == nums_before     # no instance, no definition added
    assert problems(document) == []


def test_nesting_through_a_style_tied_to_the_level_and_bullets_by_style():
    document = template()
    result = document.insert_markdown("1. a\n   1. inner\n   2. inner two\n2. b\n\n- dot\n- dot two\n")
    assert all(own_numbering(document, i) is None for i in result.blocks)
    assert [document.list_of(i).level for i in result.blocks] == [0, 1, 1, 0, 0, 0]
    assert labels(document, result.blocks[:4]) == ["1.", "a)", "b)", "2."]
    assert document.to_markdown(f"{result.blocks[0]}..{result.blocks[-1]}", ids=False) == \
        "1. a\n   1. inner\n   2. inner two\n2. b\n\n- dot\n- dot two\n"


def test_a_list_starting_elsewhere_and_a_later_list_restart_on_their_first_item():
    document = template()
    document.insert_markdown("1. earlier\n")
    result = document.insert_markdown("4. four\n5. five\n\nBetween.\n\n1. again\n2. again two\n")
    listed = items(document, result)
    first, restarted = own_numbering(document, listed[0]), own_numbering(document, listed[2])
    assert first is not None and restarted is not None
    assert own_numbering(document, listed[1]) is None and own_numbering(document, listed[3]) is None
    numbering = document.numbering
    style_num = numbering._style_num("ListNumber")
    for node, start in ((first, "4"), (restarted, "1")):
        num_id = int(node.find(W + "numId").get(W + "val"))
        assert numbering.abstract_of(num_id) == numbering.abstract_of(style_num)
        override = numbering.nums()[num_id].find(f"{W}lvlOverride/{W}startOverride")
        assert override.get(W + "val") == start
    assert labels(document, listed) == ["4.", "5.", "1.", "2."]
    # Word's restart form reads back as one list each.
    assert document.to_markdown(f"{result.blocks[0]}..{result.blocks[-1]}", ids=False) == \
        "4. four\n5. five\n\nBetween.\n\n1. again\n2. again two\n"
    assert problems(document) == []


def test_a_style_that_does_not_number_itself_keeps_direct_numbering():
    tied = _style("ProbeTied", "Probe Tied", "<w:numPr><w:numId w:val=\"1\"/></w:numPr>")
    document = template(extra_styles=tied)
    for style in ("Probe Unnumbered", "Probe Tied"):    # numId 0; a level tied to List Number
        result = document.insert_markdown("1. x\n2. y\n", style_map={"ordered_list": style})
        assert all(own_numbering(document, i) is not None for i in result.blocks), style
        assert labels(document, result.blocks) == ["1.", "2."]


def test_word_s_own_list_number_2_is_its_own_list_so_nesting_stays_direct():
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        document = Document.new()
    result = document.insert_markdown("1. a\n   1. b\n2. c\n")
    a, b, c = result.blocks
    assert own_numbering(document, a) is None and own_numbering(document, c) is None
    assert own_numbering(document, b) is not None and document.list_of(b).level == 1
    assert labels(document, result.blocks) == ["1.", "a.", "2."]


# -- restart and continue --------------------------------------------------------------------------


def test_restart_on_a_style_numbered_item_overrides_that_item_alone():
    document = template()
    result = document.insert_markdown("1. one\n2. two\n\nBetween.\n\n3. three\n4. four\n5. five\n")
    listed = items(document, result)
    edit = document.restart_numbering(listed[2])
    assert own_numbering(document, listed[2]) is not None
    assert all(own_numbering(document, i) is None for i in listed[3:])
    assert labels(document, listed) == ["1.", "2.", "1.", "2.", "3."]     # Word: measured
    assert edit.warnings == []          # the level's indent is the style's
    assert problems(document) == []


def _override_level(document: Document, identifier: str):
    num_id = int(own_numbering(document, identifier).find(W + "numId").get(W + "val"))
    return document.numbering.nums()[num_id].find(f"{W}lvlOverride")


def test_restart_keeps_the_style_s_indent_where_the_level_s_differs():
    # Round 24, W7: direct numbering takes the level's indent, so the restart's instance
    # carries the level with the style's indent -- Word draws the item at the style's indent
    # (measured; a startOverride alone put it at the level's).
    document = template(style_indent=720)
    result = document.insert_markdown("1. one\n2. two\n3. three\n")
    edit = document.restart_numbering(result.blocks[1])
    assert edit.warnings == []
    override = _override_level(document, result.blocks[1])
    assert override.find(f"{W}startOverride").get(W + "val") == "1"
    level = override.find(W + "lvl")
    assert level.find(f"{W}start").get(W + "val") == "1"
    assert level.find(f"{W}pPr/{W}ind").attrib == {W + "left": "720", W + "hanging": "284"}
    assert level.find(f"{W}lvlText").get(W + "val") == "%1." and level.find(f"{W}pStyle") is not None
    assert labels(document, result.blocks) == ["1.", "1.", "2."]
    paragraph = document.paragraph(result.blocks[1]).effective
    assert (paragraph.indent_left, paragraph.hanging) == (36, 14.2)
    assert problems(document) == []


def test_a_markdown_restart_keeps_the_style_s_indent_and_its_start():
    document = template(style_indent=720)
    document.insert_markdown("1. earlier\n")
    result = document.insert_markdown("4. four\n5. five\n\nBetween.\n\n1. again\n")
    listed = items(document, result)
    for identifier, start in ((listed[0], "4"), (listed[2], "1")):
        override = _override_level(document, identifier)
        assert override.find(f"{W}startOverride").get(W + "val") == start
        # With a whole level in the override Word counts from the level's start (measured).
        assert override.find(f"{W}lvl/{W}start").get(W + "val") == start
        assert document.paragraph(identifier).effective.indent_left == 36
    assert labels(document, listed) == ["4.", "5.", "1."]
    assert document.to_markdown(f"{result.blocks[0]}..{result.blocks[-1]}", ids=False) == \
        "4. four\n5. five\n\nBetween.\n\n1. again\n"


def _direct_items(document: Document, ids: list[str]) -> None:
    """0.1.20's form: each item its own numPr over a copy of the definition."""
    numbering = Numbering(document)
    copy = _numbering.copy_abstract(document, numbering.abstract_of(numbering._style_num("ListNumber")))
    num_id = _numbering.add_num(document, copy)
    for identifier in ids:
        _numbering.set_numbering(document._resolve(identifier)[1].element, num_id, 0)


def test_continue_drops_direct_numbering_where_the_style_numbers_the_items():
    document = template()
    first = document.insert_markdown("1. one\n2. two\n\nBetween.\n\n1. three\n2. four\n")
    listed = items(document, first)
    _direct_items(document, listed[2:])
    assert labels(document, listed) == ["1.", "2.", "1.", "2."]
    document.continue_numbering(listed[2])
    assert all(own_numbering(document, i) is None for i in listed)
    assert labels(document, listed) == ["1.", "2.", "3.", "4."]
    assert problems(document) == []


def test_continue_with_nothing_to_continue_says_why():
    document = template()
    result = document.insert_markdown("1. one\n2. two\n")
    _direct_items(document, result.blocks)
    with pytest.raises(EditError, match="has direct numbering .* while its style 'List Number' numbers itself"):
        document.continue_numbering(result.blocks[0])


def test_the_lists_tool_restarts_and_continues_style_numbered_items():
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session()
        doc = session.open(template().to_bytes(), "template.docx")
        written = box.dispatch(session, "word_insert_markdown",
                               {"doc": doc, "markdown": "1. one\n2. two\n3. three\n", "at": "end"})
        assert written.ok, written.to_json()
        document = session.entry(doc).document
        listed = [b["id"] for b in document.blocks() if b["kind"] == "paragraph" and document.list_of(b["id"])]
        restarted = box.dispatch(session, "word_lists", {"doc": doc, "targets": [listed[1]], "action": "restart"})
        assert restarted.ok, restarted.to_json()
        assert labels(document, listed) == ["1.", "1.", "2."]
        continued = box.dispatch(session, "word_lists", {"doc": doc, "targets": [listed[1]], "action": "continue"})
        assert continued.ok, continued.to_json()
        assert labels(document, listed) == ["1.", "2.", "3."] and own_numbering(document, listed[1]) is None
