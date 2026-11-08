"""One reading of every address (round 28, W9) and word_set_text's paragraph split
(W10, ``\\n`` in the text).  W9: a cell id, a row, a section, a bookmark and a span -- closed, open (``p:A..``
to the end of its container, ``..p:B`` from the start) or with refs at its ends -- are read
the same by every tool that takes a place; a form a parameter cannot take is refused with
the ids that would do, a malformed span with the syntax and never as an empty id.  W10:
``\\n`` writes consecutive paragraphs in the target's paragraph properties.  Made-up
documents, Document.new()'s package."""

from __future__ import annotations

import pytest
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.edit.addresses import AddressError, span_ends
from docx_agent.edit.errors import EditError
from docx_agent.tools import FORMAT, GROUPS, TOOLS

MD = """# Opening

Alpha words here.

Bravo words here.

| Name | Value |
|---|---|
| left | right |

Charlie words here.

Delta words here.
"""


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


@pytest.fixture
def session(box):
    return box.session()


def made(*, section_after: str | None = None, bookmark: str | None = None) -> bytes:
    document = Document.new()
    document.insert_markdown(MD, at="end")
    if section_after is not None:
        document.insert_section_break(after=_ids(document)[section_after])
    if bookmark is not None:
        document.add_bookmark(document.paragraph(_ids(document)[bookmark]).range(), "Spot")
    return document.to_bytes()


def _ids(document: Document) -> dict[str, str]:
    return {p.text.split()[0]: p.id for p in document.paragraphs() if p.text}


def opened(session, data: bytes) -> tuple[str, Document, dict[str, str]]:
    doc = session.open(data, "addresses.docx")
    document = session.entry(doc).document
    return doc, document, _ids(document)


def call(box, session, tool: str, /, **arguments):
    return box.dispatch(session, tool, arguments)


def ok(box, session, tool: str, /, **arguments):
    result = call(box, session, tool, **arguments)
    assert result.ok, (tool, result.error and result.error.to_json())
    return result


def refused(box, session, tool: str, code: str, /, **arguments):
    result = call(box, session, tool, **arguments)
    assert not result.ok and result.error.code == code, (tool, result.to_json())
    return result.error


# -- the parser -----------------------------------------------------------------------------


def test_span_ends_reads_closed_and_open_spans_and_leaves_text_ranges_alone():
    assert span_ends("p:A..t:B") == ("p:A", "t:B")
    assert span_ends("p:A..") == ("p:A", None)
    assert span_ends("..t:B") == (None, "t:B")
    assert span_ends(" p:A .. t:B ") == ("p:A", "t:B")
    assert span_ends("p@body/2..p@body/4") == ("p@body/2", "p@body/4")   # positional ids are blocks
    assert span_ends("p:A@3..p:B@8") is None                              # a text range's span
    assert span_ends("p:A") is None


@pytest.mark.parametrize("bad", ["..", "p:A...p:B", "p:A..p:B..p:C", "p:A@3..", "p:A@3..t:B"])
def test_a_malformed_span_says_the_syntax(bad):
    with pytest.raises(AddressError) as raised:
        span_ends(bad)
    assert "<first>..<last>" in str(raised.value) and "''" not in str(raised.value)


# -- W9: reading -----------------------------------------------------------------------------


def test_word_inspect_on_a_cell_inspects_the_cell(box, session):
    doc, document, _ = opened(session, made())
    table = document.tables()[0]
    cell = f"{table.id}/c1,1"
    data = ok(box, session, "word_inspect", doc=doc, target=cell).data
    assert data["cell"] == {"id": cell, "table": table.id, "row": 1, "column": 1,
                            "properties": document.cell_properties(table.id, 1, 1)}
    assert [b["t"] for b in data["blocks"]] == ["right"]
    assert [b["id"] for b in data["blocks"]] == [p.id for p in table.cell(1, 1).paragraphs]


def test_word_read_takes_a_cell_a_row_and_open_spans(box, session):
    doc, document, ids = opened(session, made())
    table = document.tables()[0]
    text = ok(box, session, "word_read", doc=doc, range=f"{table.id}/c1,0").data
    assert "left" in text and "right" not in text and "Name" not in text
    row = ok(box, session, "word_read", doc=doc, range=f"{table.id}/r1").data
    assert "left" in row and "right" in row and "Name" not in row
    to_end = ok(box, session, "word_read", doc=doc, range=f"{ids['Charlie']}..").data
    assert "Charlie" in to_end and "Delta" in to_end and "Bravo" not in to_end
    from_start = ok(box, session, "word_read", doc=doc, range=f"..{ids['Alpha']}").data
    assert "Opening" in from_start and "Alpha" in from_start and "Bravo" not in from_start


def test_a_span_takes_refs_at_both_ends(box, session):
    doc, _, ids = opened(session, made())
    session.entry(doc).refs.update(first=ids["Bravo"], last=ids["Charlie"])
    text = ok(box, session, "word_read", doc=doc, range="$first..$last").data
    assert "Bravo" in text and "Charlie" in text and "Alpha" not in text and "Delta" not in text
    tail = ok(box, session, "word_read", doc=doc, range="$last..").data
    assert "Charlie" in tail and "Delta" in tail and "Bravo" not in tail


@pytest.mark.parametrize("bad", ["p:A...p:B", "..", "p:A..p:B..p:C"])
def test_a_malformed_span_is_refused_with_the_syntax_never_as_an_empty_id(box, session, bad):
    doc, _, ids = opened(session, made())
    bad = bad.replace("p:A", ids["Alpha"]).replace("p:B", ids["Bravo"]).replace("p:C", ids["Charlie"])
    for tool, arguments in (("word_read", {"range": bad}), ("word_inspect", {"target": bad}),
                            ("find_text", {"text": "words", "range": bad}),
                            ("word_delete", {"target": bad}), ("word_move", {"blocks": bad, "after": ids["Delta"]})):
        error = refused(box, session, tool, "invalid_arguments", doc=doc, **arguments)
        assert "<first>..<last>" in error.message and "''" not in error.message, (tool, error.message)


def test_find_text_and_replace_text_take_a_cell_a_section_a_bookmark_and_an_open_span(box, session):
    doc, document, ids = opened(session, made(section_after="Bravo", bookmark="Charlie"))
    table = document.tables()[0]
    found = ok(box, session, "find_text", doc=doc, text="words", range=f"{ids['Charlie']}..").data
    assert [m["address"].partition("@")[0] for m in found] == [ids["Charlie"], ids["Delta"]]
    assert ok(box, session, "find_text", doc=doc, text="right", range=f"{table.id}/c1,1").data
    assert not ok(box, session, "find_text", doc=doc, text="left", range=f"{table.id}/c1,1").data
    first_section = ok(box, session, "find_text", doc=doc, text="words", range="s:1").data
    assert [m["address"].partition("@")[0] for m in first_section] == [ids["Alpha"], ids["Bravo"]]
    marked = ok(box, session, "find_text", doc=doc, text="words", range="bm:Spot").data
    assert [m["address"].partition("@")[0] for m in marked] == [ids["Charlie"]]
    ok(box, session, "replace_text", doc=doc, find="words", replace="lines", expect="all", range="s:body")
    assert [p.text for p in document.paragraphs() if p.text.endswith("here.")] == [
        "Alpha words here.", "Bravo words here.", "Charlie lines here.", "Delta lines here."]


# -- W9: writing -----------------------------------------------------------------------------


def test_word_format_takes_a_cell_a_row_a_section_and_a_bookmark(box, session):
    doc, document, ids = opened(session, made(section_after="Bravo", bookmark="Delta"))
    table = document.tables()[0]
    ok(box, session, "word_format", doc=doc, targets=[f"{table.id}/c1,1"], bold=True)
    ok(box, session, "word_format", doc=doc, targets=["s:1", "bm:Spot"], italic=True)
    bold = {p.text for p in document.paragraphs() if p.runs and p.runs[0].effective.bold}
    italic = {p.text for p in document.paragraphs() if p.runs and p.runs[0].effective.italic}
    assert bold == {"right"}
    assert italic == {"Opening", "Alpha words here.", "Bravo words here.", "Delta words here."}


def test_word_delete_and_word_move_take_open_spans(box, session):
    doc, document, ids = opened(session, made())
    ok(box, session, "word_move", doc=doc, blocks=f"..{ids['Alpha']}", after=ids["Delta"])
    assert [p.text.split()[0] for p in document.paragraphs() if p.text][-2:] == ["Opening", "Alpha"]
    ok(box, session, "word_delete", doc=doc, target=f"{ids['Delta']}..")
    assert [p.text.split()[0] for p in document.paragraphs() if p.text] == [
        "Bravo", "Name", "Value", "left", "right", "Charlie"]


def test_text_targets_take_a_cell_a_bookmark_and_a_paragraph(box, session):
    doc, document, ids = opened(session, made(bookmark="Bravo"))
    table = document.tables()[0]
    cell = f"{table.id}/c1,1"
    ok(box, session, "word_insert_text", doc=doc, target=cell, text=" side")
    assert table.cell(1, 1).text == "right side"
    ok(box, session, "word_comments", doc=doc, action="add", items=[{"target": cell, "text": "Check."}])
    assert document.comments()[0].anchor.text == "right side"
    ok(box, session, "word_notes", doc=doc, action="insert", at=ids["Alpha"], text="A note.")
    ok(box, session, "word_notes", doc=doc, action="insert", at="bm:Spot", text="Another.")
    marks = [r for r in document.state()["blocks"] if r.get("notes")]
    assert [b["t"] for b in marks] == ["Alpha words here.", "Bravo words here."]


def test_a_section_parameter_takes_any_block_and_names_its_section(box, session):
    doc, document, ids = opened(session, made(section_after="Bravo"))
    table = document.tables()[0]
    ok(box, session, "word_sections", doc=doc, action="set", section=ids["Alpha"], orientation="landscape")
    ok(box, session, "word_headers_footers", doc=doc, section=f"{table.id}/c0,0", which="header",
       action="set", text="Later pages")
    first, last = document.sections()
    assert first.orientation == "landscape" and last.orientation == "portrait"
    assert last.header() is not None and "Later pages" in last.header().paragraphs[0].text
    error = refused(box, session, "word_sections", "invalid_arguments", doc=doc, action="set",
                    section=f"{ids['Alpha']}..{ids['Bravo']}", orientation="portrait")
    assert "span" in error.message


def test_a_cell_where_a_block_is_wanted_names_its_table_and_paragraphs(box, session):
    doc, document, ids = opened(session, made())
    table = document.tables()[0]
    cell = f"{table.id}/c1,1"
    inside = table.cell(1, 1).paragraphs[0].id
    for tool, arguments in (("word_insert_markdown", {"markdown": "New.", "at": f"after:{cell}"}),
                            ("word_move", {"blocks": ids["Delta"], "after": cell}),
                            ("word_delete", {"target": cell}),
                            ("word_compare_blocks", {"blocks": [cell], "reference": {"block": ids["Alpha"]}})):
        error = refused(box, session, tool, "invalid_arguments", doc=doc, **arguments)
        assert "is a cell" in error.message and table.id in error.valid_options, (tool, error.message)
        assert inside in error.valid_options
    error = refused(box, session, "word_format_table", "invalid_arguments", doc=doc, table=inside,
                    scope="table", alignment="center")
    assert error.valid_options == [table.id]


def test_a_span_where_one_block_is_wanted_names_its_ends(box, session):
    doc, _, ids = opened(session, made())
    span = f"{ids['Alpha']}..{ids['Bravo']}"
    error = refused(box, session, "word_move", "invalid_arguments", doc=doc, blocks=ids["Delta"], after=span)
    assert "is a span" in error.message and error.valid_options == [ids["Alpha"], ids["Bravo"]]


def test_the_library_reads_the_same_forms():
    document = Document.open(made(bookmark="Charlie"))
    ids = _ids(document)
    table = document.tables()[0]
    assert document.range(f"{table.id}/c1,1").text == "right"
    assert document.range("Spot").text == "Charlie words here."
    assert document.range(ids["Alpha"]).text == "Alpha words here."
    assert document.section_of(f"{table.id}/c1,1").id == "s:body"
    with pytest.raises(AddressError, match="is a table, not text"):
        document.range(table.id)
    with pytest.raises(KeyError, match="is a cell"):          # an AddressError is a KeyError
        document.paragraph(f"{table.id}/c1,1")
    with pytest.raises(EditError, match="one container"):
        document.move_blocks(f"{ids['Alpha']}..{table.cell(1, 1).paragraphs[0].id}", after=ids["Delta"])


# -- W10: a newline splits -----------------------------------------------------------------


def test_a_newline_split_writes_each_line_as_a_paragraph_in_the_targets_properties(box, session):
    doc, document, ids = opened(session, made())
    document.format_paragraph(ids["Bravo"], alignment="center", space_after=3)
    document.paragraph(ids["Bravo"]).format(bold=True)
    style = document.paragraph(ids["Bravo"]).style
    result = ok(box, session, "word_set_text", doc=doc, target=ids["Bravo"],
                text="First line.\nSecond\vwith a break.\n\nFourth.")
    texts = [p.text for p in document.paragraphs()]
    at = texts.index("First line.")
    assert texts[at:at + 4] == ["First line.", "Second\vwith a break.", "", "Fourth."]
    assert len(result.created) == 3 and result.changed == [ids["Bravo"]]
    made_ = [document.paragraph(i) for i in result.created]
    assert [p.text for p in made_] == ["Second\vwith a break.", "", "Fourth."]
    for paragraph in made_:
        assert paragraph.style == style
        assert paragraph.effective.alignment == "center" and paragraph.effective.space_after == 3
    assert made_[0].runs[0].effective.bold      # the last run's formatting, as Enter continues it
    again = ok(box, session, "word_set_text", doc=doc, target=ids["Bravo"],
               text="First line.\nSecond\vwith a break.\n\nFourth.")
    assert again.created == [] and again.changed == []
    assert [p.text for p in document.paragraphs()] == texts


def test_without_a_newline_a_cell_target_rewrites_its_first_paragraph_as_before(box, session):
    doc, document, _ = opened(session, made())
    table = document.tables()[0]
    ok(box, session, "word_set_text", doc=doc, target=f"{table.id}/c1,1", text="R1\nR2")
    ok(box, session, "word_set_text", doc=doc, target=f"{table.id}/c1,1", text="Only\vbroken")
    assert [p.text for p in table.cell(1, 1).paragraphs] == ["Only\vbroken", "R2"]


def test_a_newline_split_applies_per_target_and_writes_a_cells_paragraphs(box, session):
    doc, document, ids = opened(session, made())
    table = document.tables()[0]
    result = ok(box, session, "word_set_text", doc=doc,                 items=[{"target": ids["Alpha"], "text": "A1\nA2"},
                       {"target": f"{table.id}/c1,1", "text": "R1\nR2"},
                       {"target": ids["Delta"], "text": "One line only"}])
    texts = [p.text for p in document.paragraphs()]
    assert texts[texts.index("A1") + 1] == "A2"
    assert [p.text for p in table.cell(1, 1).paragraphs] == ["R1", "R2"]
    assert "One line only" in texts and len(result.created) == 2
    again = ok(box, session, "word_set_text", doc=doc,                items=[{"target": f"{table.id}/c1,1", "text": "R1\nR2"}])
    assert again.created == [] and again.changed == []


def test_a_newline_split_is_tracked_and_undone_as_one_call(box, session):
    doc, document, ids = opened(session, made())
    texts = [p.text for p in document.paragraphs()]
    ok(box, session, "word_set_tracking", doc=doc, on=True, author="Reviewer")
    ok(box, session, "word_set_text", doc=doc, target=ids["Charlie"], text="Charlie words here.\nAdded.")
    assert "insertion" in {r.kind for r in document.revisions()}
    assert "Added." in [p.text for p in document.paragraphs()]
    ok(box, session, "undo", doc=doc)
    assert [p.text for p in document.paragraphs()] == texts and not document.revisions()


def test_a_newline_split_refuses_a_paragraph_that_ends_a_section():
    document = Document.open(made(section_after="Bravo"))
    ending = "p" + document.sections()[0].id[1:]          # s:<paraId>: the paragraph holding the break
    with pytest.raises(EditError, match="ends a section"):
        document.set_text_paragraphs(ending, "One\nTwo")


def test_a_newline_costs_no_parameter_and_the_strict_set_stays_within_24(box):
    # 23 then; round 40's ref on word_insert_markdown made it 24, Claude's limit.  Round 42's
    # paragraph_style (the single form's and each item's) takes word_set_text to 5 optional
    # parameters: the strict plan sends it non-strict and keeps the set within the limit;
    # word_insert_markdown's like (R47) fills the 24th place.
    definitions = box.definitions("anthropic", defer=False)
    set_text = next(d for d in definitions if d["name"] == "word_set_text")
    assert list(set_text["input_schema"]["properties"]) == ["doc", "target", "text", "paragraph_style", "items"]
    tools = {t.name: t for t in TOOLS}
    assert len(tools["word_set_text"].optional_parameters()) == 5
    assert sum(len(tools[d["name"]].optional_parameters()) for d in definitions if d.get("strict")) == 24
    assert {d["name"] for d in definitions if d.get("strict")} == {
        "replace_text", "word_changes", "word_set_tracking", "word_insert_markdown", "word_template",
        "edit_smartart"}


def test_a_reversed_span_is_refused_alike_everywhere(box, session):
    doc, _, ids = opened(session, made())
    reversed_ = f"{ids['Charlie']}..{ids['Alpha']}"
    for tool, arguments in (("word_read", {"range": reversed_}), ("word_delete", {"target": reversed_}),
                            ("word_insert_markdown", {"markdown": "New.", "at": f"replace:{reversed_}"}),
                            ("word_move", {"blocks": reversed_, "after": ids["Delta"]})):
        error = refused(box, session, tool, "refused", doc=doc, **arguments)
        assert error.message == f"{reversed_!r} ends before it starts", (tool, error.message)


def test_a_bookmark_or_story_where_a_block_is_wanted_gets_the_hint(box, session):
    doc, _, ids = opened(session, made(bookmark="Charlie"))
    for anchor in ("Spot", "bm:Spot"):
        error = refused(box, session, "word_move", "invalid_arguments", doc=doc, blocks=ids["Delta"], after=anchor)
        assert "is a bookmark" in error.message and error.valid_options == [ids["Charlie"]]
    error = refused(box, session, "word_edit_table", "invalid_arguments", doc=doc, table="Spot",
                    action="insert_row")
    assert "is a bookmark" in error.message
    error = refused(box, session, "word_insert_markdown", "invalid_arguments", doc=doc, markdown="New.",
                    at="after:body")
    assert "is a story" in error.message
