"""Round 36: the outline as heading-owned ranges (R32) and the guards that keep a heading's
content from changing owner unseen (R33).  Made-up documents only."""

from __future__ import annotations

import pytest
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.edit.addresses import AddressError
from docx_agent.edit.outline import content_left, headings, moves
from docx_agent.tools import FORMAT, GROUPS, TOOLS

MD = """# Opening

Opening words.

## First part

First words.

| k | v |
| --- | --- |
| 1 | 2 |

### Detail

Detail words.

## Second part

Second words.

# Closing

Closing words.
"""


def made(*, section_after: str | None = None, in_table: bool = False) -> bytes:
    document = Document.new()
    document.insert_markdown(MD, at="end")
    ids = _ids(document)
    if section_after is not None:
        document.insert_section_break(after=ids[section_after])
    if in_table:
        table = document.tables()[0]
        document.set_paragraph_style(table.cell(1, 0).paragraphs[0].id, "Heading 2")
    return document.to_bytes()


def _ids(document: Document) -> dict[str, str]:
    return {p.text: p.id for p in document.paragraphs() if p.text}


def _blocks(document: Document) -> list[str]:
    return [b["id"] for b in document.blocks()]


def _words(document: Document, text: str) -> str:
    return next(p.id for p in document.paragraphs() if p.text == text)


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def opened(box, data: bytes):
    session = box.session()
    doc = session.open(data, "outline.docx")
    document = session.entry(doc).document
    return session, doc, document, _ids(document)


def ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, result.error and result.error.to_json())
    return result


def refused(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert not result.ok, (tool, result.to_json())
    return result.error.to_json()


# -- R32: the outline -------------------------------------------------------------------------


def test_the_outline_gives_each_heading_its_range_and_section():
    document = Document.open(made(section_after="First words."))
    ids = _ids(document)
    outline = {h["text"]: h for h in document.outline()}
    assert [h["text"] for h in document.outline()] == ["Opening", "First part", "Detail", "Second part", "Closing"]
    first = outline["First part"]
    assert first["range"] == "h:" + ids["First part"] and first["level"] == 2
    # The heading, its words, the break's paragraph, the table, Detail and its words.
    assert first["blocks"] == 6 and first["last"] == _words(document, "Detail words.")
    assert first["section"] == "s:1" and first["to_section"] == "s:2"
    assert outline["Second part"]["section"] == "s:2" and "to_section" not in outline["Second part"]
    assert outline["Opening"]["blocks"] == 10 and outline["Closing"]["blocks"] == 2
    assert outline["Detail"]["blocks"] == 2 and outline["Detail"]["level"] == 3


def test_a_heading_range_reads_as_its_span_everywhere_in_the_library():
    document = Document.open(made())
    ids = _ids(document)
    span = document.heading_span("h:" + ids["First part"])
    assert span == f"{ids['First part']}..{_words(document, 'Detail words.')}"
    assert document.heading_span(ids["First part"]) == span                    # the heading's id says it too
    assert document.to_markdown("h:" + ids["First part"]) == document.to_markdown(span)
    second = f"{ids['Second part']}..{_words(document, 'Second words.')}"
    assert document.state("h:" + ids["Second part"])["blocks"] == document.state(second)["blocks"]
    assert [p.id for p in document.paragraphs()] and len(document._scope_paragraphs("h:" + ids["First part"])[1]) == 8
    # A span's ends may be headings: from the first through the end of the second's range.
    both = document.to_markdown(f"h:{ids['First part']}..h:{ids['Second part']}", ids=False)
    assert "First part" in both and "Second words." in both and "Closing" not in both
    assert document.heading_span(ids["Closing"]) == f"{ids['Closing']}..{_blocks(document)[-1]}"
    # A heading with nothing under it is just itself.
    document.delete_block(_words(document, "Closing words."))
    assert document.heading_span(ids["Closing"]) == ids["Closing"]
    # A span through a heading range's end.
    moved = document.move_blocks("h:" + ids["Detail"], after=ids["Closing"])
    assert moved.blocks[0] == ids["Detail"]


def test_only_body_paragraphs_with_an_outline_level_are_headings():
    document = Document.open(made(in_table=True))
    texts = [h["text"] for h in document.outline()]
    assert "1" not in texts                                # a heading style in a cell owns nothing
    with pytest.raises(AddressError, match="not a block of the body itself") as error:
        document.heading_span("h:" + document.tables()[0].cell(1, 0).paragraphs[0].id)
    assert error.value.options[0].startswith("h:")
    with pytest.raises(AddressError, match="is not a paragraph"):
        document.heading_span("h:" + document.tables()[0].id)
    # Body text: refused with the headings to use.
    words = _words(document, "First words.")
    with pytest.raises(AddressError, match="no outline level") as error:
        document.heading_span("h:" + words)
    assert any("First part" in option for option in error.value.options)


def test_a_title_is_a_heading_only_with_an_outline_level():
    document = Document.open(made())
    words = _words(document, "Opening words.")
    document.set_paragraph_style(words, "Title")
    assert words not in [h["id"] for h in document.outline()]
    # A paragraph stating an outline level is a heading of that level.
    _state_level(document, words, 1)
    assert {h["id"]: h["level"] for h in document.outline()}.get(words) == 2


def _state_level(document: Document, paragraph: str, value: int) -> None:
    from lxml import etree

    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    _, entry = document._resolve(paragraph)
    properties = entry.element.find(f"{{{w}}}pPr")
    level = etree.SubElement(properties, f"{{{w}}}outlineLvl")
    level.set(f"{{{w}}}val", str(value))
    document._invalidate()


def test_get_says_a_heading_range_is_one(box):
    document = Document.open(made())
    ids = _ids(document)
    with pytest.raises(AddressError, match="is a heading's range"):
        document.get("h:" + ids["First part"])


# -- R32: every tool's range takes h: ---------------------------------------------------------


def test_the_tools_take_a_heading_range(box):
    session, doc, document, ids = opened(box, made())
    first = "h:" + ids["First part"]
    read = ok(box, session, "word_read", doc=doc, range=first).data
    assert "First part" in read and "Detail words." in read and "Second part" not in read
    inspected = ok(box, session, "word_inspect", doc=doc, target=first).data
    assert len(inspected["blocks"]) == 5
    found = ok(box, session, "find_text", doc=doc, text="words", range=first).data
    assert [f["context"] for f in found] == ["First [words].", "Detail [words]."]
    checked = ok(box, session, "check", doc=doc, include=["validate"], range=[first])
    assert checked.ok
    ok(box, session, "word_format", doc=doc, targets=[first], keep_with_next=True)
    assert all(document.paragraph(p).keep_with_next for p in (ids["First part"], ids["Detail"]))
    assert not document.paragraph(ids["Second part"]).keep_with_next
    ok(box, session, "word_move", doc=doc, blocks="h:" + ids["Second part"], before=ids["First part"])
    assert [h["text"] for h in document.outline()][1:3] == ["Second part", "First part"]
    ok(box, session, "word_delete", doc=doc, target=first)
    assert [h["text"] for h in document.outline()] == ["Opening", "Second part", "Closing"]
    assert "Detail words." not in [p.text for p in document.paragraphs()]


def test_a_heading_range_naming_no_heading_is_refused_with_the_headings(box):
    session, doc, document, ids = opened(box, made())
    words = next(p.id for p in document.paragraphs() if p.text == "First words.")
    error = refused(box, session, "word_read", doc=doc, range="h:" + words)
    assert error["code"] == "invalid_arguments" and "no outline level" in error["message"]
    assert error["valid_options"][0].startswith("h:" + ids["Opening"])
    error = refused(box, session, "check", doc=doc, range=["h:" + words])
    assert error["valid_options"][0].startswith("h:")


def test_describe_names_each_headings_section_only_with_several(box):
    session, doc, _, _ = opened(box, made())
    headings_ = ok(box, session, "describe", doc=doc).data["headings"]
    assert all("section" not in h for h in headings_) and headings_[1]["blocks"] == 5
    session, doc, _, _ = opened(box, made(section_after="First words."))
    headings_ = ok(box, session, "describe", doc=doc).data["headings"]
    assert [h["section"] for h in headings_] == ["s:1", "s:1", "s:2", "s:2", "s:2"]


# -- R33: deleting a heading whose content stays ----------------------------------------------


def test_deleting_a_heading_alone_is_refused_with_its_range_and_force(box):
    session, doc, document, ids = opened(box, made())
    error = refused(box, session, "word_delete", doc=doc, target=ids["First part"])
    assert error["field"] == "target" and "would go under" in error["message"]
    assert ids["Opening"] in error["message"]
    assert error["valid_options"] == ["h:" + ids["First part"], "force"]
    assert "First part" in [p.text for p in document.paragraphs()]           # nothing changed
    # A span covering the heading but not all it owns: refused alike.
    error = refused(box, session, "word_delete", doc=doc, target=f"{ids['First part']}..{ids['First part']}")
    assert "h:" + ids["First part"] in error["valid_options"]
    # force: deleted, the warning says where its content went.
    result = ok(box, session, "word_delete", doc=doc, target=ids["First part"], force=True)
    assert any("now go under" in w and ids["Opening"] in w for w in result.warnings)
    assert "First words." in [p.text for p in document.paragraphs()]


def test_deleting_a_heading_with_nothing_under_it_or_its_whole_range_is_not_refused(box):
    session, doc, document, ids = opened(box, made())
    ok(box, session, "word_delete", doc=doc, target=f"{ids['Detail']}..{_words(document, 'Detail words.')}")
    body = document.blocks()
    last = body[-1]["id"]
    ok(box, session, "word_delete", doc=doc, target=last)          # body text: no heading involved
    ok(box, session, "word_delete", doc=doc, target=_ids(document)["Closing"])     # owns nothing now
    assert content_left(document, []) == []


# -- R33: unnumbering a paragraph its style numbers ----------------------------------------------


def _numbered_style(document: Document) -> tuple[str, str]:
    """A style that numbers its paragraphs, and a paragraph in it."""
    words = next(p.id for p in document.paragraphs() if p.text == "Second words.")
    document.add_to_list(words, "number")
    num_id = document.list_of(words).num_id
    document.styles.add("Step text", "paragraph", based_on="Normal")
    document.styles.modify("Step text", numbering={"list": num_id, "level": 0})
    document.styles.add("Step plain", "paragraph", based_on="Step text")
    document.remove_from_list(words)
    document.set_paragraph_style(words, "Step text")
    return words, num_id


def test_unnumbering_a_style_numbered_paragraph_is_refused_with_styles_that_fit(box):
    document = Document.open(made())
    words, _ = _numbered_style(document)
    session, doc, opened_document, _ = opened(box, document.to_bytes())
    assert opened_document.list_of(words) is not None
    error = refused(box, session, "word_lists", doc=doc, targets=[words], action="remove")
    assert error["field"] == "targets" and "'Step text' numbers them" in error["message"]
    assert "Normal" in error["valid_options"] and error["valid_options"][-1] == "force"
    assert opened_document.list_of(words) is not None                        # nothing changed
    ok(box, session, "word_lists", doc=doc, targets=[words], action="remove", force=True)
    assert opened_document.list_of(words) is None


def test_unnumbering_an_own_numbered_paragraph_is_not_refused(box):
    session, doc, document, ids = opened(box, made())
    words = next(p.id for p in document.paragraphs() if p.text == "Second words.")
    ok(box, session, "word_lists", doc=doc, targets=[words], action="add", kind="number")
    ok(box, session, "word_lists", doc=doc, targets=[words], action="remove")
    assert document.list_of(words) is None


# -- R33: a heading's level changed: a warning, where its content goes ------------------------


def test_demoting_a_heading_warns_where_its_content_goes(box):
    session, doc, document, ids = opened(box, made())
    result = ok(box, session, "word_format", doc=doc, targets=[ids["Second part"]], paragraph_style="Heading 3")
    assert any("demoted 2 -> 3" in w and ids["First part"] in w for w in result.warnings), result.warnings


def test_promoting_a_heading_warns_what_it_takes(box):
    session, doc, document, ids = opened(box, made())
    result = ok(box, session, "word_format", doc=doc, targets=[ids["First part"]], paragraph_style="Heading 1")
    assert any("promoted 2 -> 1" in w and "Second part" in w and "from under" in w for w in result.warnings), \
        result.warnings


def test_a_heading_made_body_text_warns_where_its_blocks_go(box):
    session, doc, document, ids = opened(box, made())
    result = ok(box, session, "word_format", doc=doc, targets=[ids["Detail"]], paragraph_style="Normal")
    assert any("no longer a heading" in w and ids["First part"] in w for w in result.warnings), result.warnings
    result = ok(box, session, "word_format", doc=doc, targets=[ids["Detail"]], paragraph_style="Heading 3")
    assert any("is now a level 3 heading" in w for w in result.warnings), result.warnings


def test_a_style_change_that_leaves_levels_alone_warns_nothing(box):
    session, doc, document, ids = opened(box, made())
    words = next(p.id for p in document.paragraphs() if p.text == "First words.")
    result = ok(box, session, "word_format", doc=doc, targets=[words], paragraph_style="Normal")
    assert not any("heading" in w for w in result.warnings)


def test_moves_compares_two_outlines():
    document = Document.open(made())
    ids = _ids(document)
    before = headings(document)
    assert moves(before, before) == []
    document.set_paragraph_style(ids["Detail"], "Heading 2")
    lines = moves(before, headings(document))
    assert len(lines) == 1 and "promoted 3 -> 2" in lines[0]
