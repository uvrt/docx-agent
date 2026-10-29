"""``Document.block_of`` and ``Document.blocks_between`` (production round 22, U1): the body's
top-level block holding any id the library takes, and the inclusive span of top-level
blocks between two such ids, through the span resolver the block tools use."""

from __future__ import annotations

import pytest

from docx_agent import Document, EditError
from docx_agent.tools import block_of, blocks_between

DATE = "2026-10-28T09:00:00Z"


@pytest.fixture
def made() -> dict:
    """A heading, two body paragraphs (a note, a comment, a bookmark, a text box, a link and a
    field in them), a table with a nested table, a content control, a header."""
    document = Document.new(created=DATE)
    title = document.paragraphs()[0]
    title.set_text("Report title")
    body = title.insert_after("Some body text here.").id
    later = document.paragraph(body).insert_after("Third paragraph text.").id
    table = document.insert_table(2, 2, after=later).id
    cell = document.table(table).cell(0, 1).paragraphs[0].id
    document.paragraph(cell).set_text("cell text")
    nested = document.insert_table(1, 1, after=cell).id
    inner = document.table(nested).cell(0, 0).paragraphs[0].id
    closing = document.insert_paragraph("Closing words.", after=table).id
    control = document.insert_control(None, "rich-text", after=table, text="Inside").id
    note = document.insert_footnote(f"{body}@4", "A note.").id
    comment = document.add_comment(f"{later}@0:5", "A remark.", author="Reviewer").id
    document.add_bookmark(document.range(f"{later}@6:15"), "Marked")
    drawing = document.insert_text_box(f"{body}@2", "boxed text").id
    box = next(p.id for p in document.paragraphs() if p.text == "boxed text")
    link = document.insert_hyperlink(f"{closing}@0", "Site", url="https://example.com").id
    document.insert_page_number(f"{closing}@0")
    document.add_header("s:body", "default", "Header text")
    header = next(s for s in document.stories if s.name.startswith("header"))
    return {"doc": document, "title": title.id, "body": body, "later": later, "table": table, "cell": cell,
            "nested": nested, "inner": inner, "control": control, "closing": closing, "note": note,
            "comment": comment, "box": box, "link": link, "header": header.paragraphs[0].id,
            "header_story": header.name, "drawing": drawing}


def test_a_top_level_block_is_its_own(made):
    document = made["doc"]
    for block in document.blocks():
        assert document.block_of(block["id"]) == block["id"]


def test_nested_ids_give_the_top_level_block(made):
    document, table = made["doc"], made["table"]
    assert document.block_of(made["cell"]) == table
    assert document.block_of(made["nested"]) == table
    assert document.block_of(made["inner"]) == table
    assert document.block_of(f"{table}/c0,1") == table
    assert document.block_of(document.table(table).rows[1].id) == table
    assert document.block_of(f"{made['cell']}/r0") == table
    assert document.block_of(f"{made['cell']}@1:3") == table                       # a range: its start
    assert document.block_of(f"{made['body']}@1..{made['cell']}@2") == made["body"]
    assert document.block_of(made["box"]) == made["body"]                          # a text box's paragraph
    assert document.block_of(made["drawing"]) == made["body"]
    inside = next(p.id for p in document.paragraphs() if p.text == "Inside")
    assert document.block_of(inside) == made["control"]                            # a control's paragraph


def test_annotations_give_the_block_that_holds_them(made):
    document = made["doc"]
    assert document.block_of(made["note"]) == made["body"]                         # its reference
    assert document.block_of(made["comment"]) == made["later"]                     # its range's start
    assert document.block_of("bm:Marked") == made["later"]
    assert document.block_of("Marked") == made["later"]                            # the plain name
    assert document.block_of(made["link"]) == made["closing"]
    fields = [f.id for f in document.fields() if f.id.startswith("fld:body/")]
    assert fields and document.block_of(fields[0]) == made["closing"]
    assert document.block_of(made["control"]) == made["control"]


def test_what_is_not_in_the_body_is_none(made):
    document = made["doc"]
    assert document.block_of(made["header"]) is None
    assert document.block_of(made["header_story"]) is None
    assert document.block_of("body") is None
    assert document.block_of(document.note(made["note"]).paragraph_ids[0]) is None   # the note's own text
    assert document.block_of("s:body") is None


def test_a_section_break_is_its_paragraphs_block(made):
    document = made["doc"]
    document.insert_section_break(after=made["body"], kind="nextPage")
    section = next(s.id for s in document.sections() if s.id != "s:body")
    assert document.block_of(section) == document.block_of("p" + section[1:])


def test_an_unknown_id_is_not_found_and_a_span_is_refused(made):
    document = made["doc"]
    for unknown in ("p:DEADBEEF", "fn:99", "c#99", "bm:Nowhere", "Nowhere", "d:99", f"{made['table']}/c9,9"):
        with pytest.raises((KeyError, IndexError)):
            document.block_of(unknown)
    with pytest.raises(EditError, match="blocks_between"):
        document.block_of(f"{made['body']}..{made['later']}")
    with pytest.raises(KeyError, match="tools"):
        document.block_of("$intro")


def test_blocks_between_spans_top_level_blocks_from_nested_ends(made):
    document = made["doc"]
    assert document.blocks_between(made["note"], made["inner"]) == [made["body"], made["later"], made["table"]]
    assert document.blocks_between(made["title"], made["title"]) == [made["title"]]
    assert document.blocks_between(made["cell"], made["link"]) == [made["table"], made["control"], made["closing"]]


def test_blocks_between_refuses_a_reversed_span_as_the_span_tools_do(made):
    document = made["doc"]
    with pytest.raises(EditError, match="ends before it starts"):
        document._blocks_of(f"{made['later']}..{made['body']}")
    with pytest.raises(EditError, match="ends before it starts"):
        document.blocks_between(made["inner"], made["note"])
    with pytest.raises(EditError, match="not in the body"):
        document.blocks_between(made["header"], made["body"])


class _Call:
    """What a handler's ``call`` gives the helpers: the document, and the session's refs."""

    def __init__(self, document: Document, refs: dict[str, str]) -> None:
        self.document = document
        self.refs = refs

    def resolve_ref(self, value: str) -> str:
        return self.refs[value[1:]]


def test_the_tools_helpers_resolve_refs_first(made):
    call = _Call(made["doc"], {"cellref": made["cell"], "noteref": made["note"]})
    assert block_of(call, "$cellref") == made["table"]
    assert block_of(call, made["box"]) == made["body"]
    assert blocks_between(call, "$noteref", "$cellref") == [made["body"], made["later"], made["table"]]


def test_a_revision_and_an_endnote_give_their_block(made):
    document = made["doc"]
    with document.tracking(author="Reviewer", date=DATE):
        document.insert_text(f"{made['inner']}@0", "new ")
    revision = next(r.id for r in document.revisions())
    assert document.block_of(revision) == made["table"]
    endnote = document.insert_endnote(f"{made['closing']}@3", "An endnote.").id
    assert document.block_of(endnote) == made["closing"]


def test_the_application_resolves_refs_outside_a_call(made):
    """Round 23, V5: block_of(session, doc, id) and blocks_between(session, doc, a, b), the
    same helpers as a handler's (call, ...) forms, resolve a ref with Session.resolve_ref."""
    from ooxml_edit.tools import ToolError, Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session()
        doc = session.open(made["doc"].to_bytes(), "made.docx")
        result = toolbox.dispatch(session, "word_insert_text", {"doc": doc, "items": [
            {"target": made["cell"], "text": " more", "ref": "cell_text"},
            {"target": f"{made['body']}@0", "where": "before", "text": "So ", "ref": "opening"}]})
        assert result.ok, result.summary
        assert set(session.refs(doc)) == {"cell_text", "opening"}
        assert block_of(session, doc, "$cell_text") == made["table"]
        assert block_of(session, doc, made["box"]) == made["body"]         # an id as given
        assert blocks_between(session, doc, "$opening", "$cell_text") == [made["body"], made["later"], made["table"]]
        with pytest.raises(ToolError) as missing:
            block_of(session, doc, "$nowhere")
        assert missing.value.code == "not_found" and missing.value.valid_options == ["$cell_text", "$opening"]
        with pytest.raises(ToolError) as closed:
            block_of(session, "d9", "$opening")
        assert closed.value.code == "not_found"
        with pytest.raises(TypeError, match=r"\(session, doc, first, last\)"):
            blocks_between(session, doc, "$opening")
        toolbox.dispatch(session, "undo", {"doc": doc})                    # the refs go with it
        assert session.refs(doc) == {}
