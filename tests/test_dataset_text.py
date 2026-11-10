"""A dataset's text written by reference, and a comment on a table (round 45).

R65: ``word_insert_markdown``'s ``from`` names one text field of one row (its one name:
``markdown_from`` is refused naming it, round 47, R81) and writes its Markdown as ``markdown`` would -- paragraphs, lists in the
template's styles at their levels, emphasis -- verbatim; several fields go in one call as
references standing alone in ``markdown``; the result says which field made which blocks
and no longer echoes a long text.  ``word_set_text`` writes a field the same way.
R66: ``word_edit_table``'s ``set_cells`` takes one value per cell (``from``, a number
format after it) and references in a cell's text; round 49, R96: a cell's own ``from`` or
``text`` beside a top-level one-value ``from`` wins for its cell, the top-level value fills the
others; a top-level ``from`` of rows beside cells is refused naming what to drop for each intent.
R68: ``word_comments`` add on a table or a row: the comment covers it as Word's does when
it is selected (measured with Word 16.106: ``test_word_anchors_a_selected_table_so``).

Tool design rules: 1 (a top-level one-value ``from`` with places, and beside cells' own), 2 (every refusal names what
there is and the call that works; ``markdown_from`` refused naming ``from``), 4 (a long value counted, not echoed),
5 (several fields in one call), 6 (lists in the template's styles; Word's comment form).
The data is made up.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

pytest.importorskip("ooxml_edit.tools.textfacts", reason="needs ooxml-edit 0.33 (text facts)")

from ooxml_common.locales import Locale  # noqa: E402
from ooxml_edit.tools import Dataset, Toolbox, tool  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from docx_agent.validate import check  # noqa: E402

from test_tables_from_data import GRID, _session, call, failed, ok  # noqa: E402

NL_BE = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")

#: A made-up note in Markdown: two paragraphs, a bullet list with a nested level, a
#: numbered list -- the shapes a builder retyped.
BODY = ("Opening paragraph with *emphasis* and 41 items.\n\nSecond paragraph.\n\n"
        "- first point\n- second point\n  - detail a\n  - detail b\n- third point\n\n"
        "1. step one\n2. step two")
COLUMNS = ("topic", "body", "score", "summary")
ROWS = [("Doors", BODY, 2.456, "Doors are mostly fine."),
        ("Roof", "Short roof text.", 4, "Roof needs work.\nSoon."),
        ("Walls", "| Part | State |\n|---|---|\n| North | good |", 3.5, "Walls hold.")]


@tool("ds_notes", "Keep a made-up dataset of notes.", {}, documents=())
def ds_notes(call):
    call.define_dataset("notes", Dataset(COLUMNS, ROWS, keys=1), source=None)


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_notes], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def session(toolbox):
    found = _session(toolbox)
    ok(call(toolbox, found, "ds_notes"))
    return found


def insert(toolbox, session, **arguments):
    return call(toolbox, session, "word_insert_markdown", **{"doc": "d1", "at": "end", **arguments})


def document(session):
    return session.entry("d1").document


def body_xml(session) -> str:
    data = document(session).to_bytes()
    return zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml").decode()


# -- R65: one field as Markdown -----------------------------------------------------------------


def test_markdown_from_is_refused_as_an_unknown_argument_naming_from(toolbox, session):
    """Round 47 (R81): ``from`` is the one name.  ``markdown_from`` gets the refusal any
    unknown argument gets -- its field, the tool's arguments -- naming ``from``, not the
    nearest spelling (``markdown``, which would write the reference as text); nothing is
    written (rule 2)."""
    before = document(session).to_bytes()
    error = failed(insert(toolbox, session, markdown_from="d:notes/#1/body"), "invalid_arguments")
    assert error.message == "unknown argument 'markdown_from'; did you mean 'from'?"
    usual = failed(insert(toolbox, session, markdown_source="d:notes/#1/body"), "invalid_arguments")
    assert error.field == "markdown_from" and error.valid_options == usual.valid_options
    assert "from" in error.valid_options
    assert document(session).to_bytes() == before
    # In a batch too, as its op.
    result = call(toolbox, session, "batch", ops=[{"tool": "word_insert_markdown", "arguments": {
        "doc": "d1", "at": "end", "markdown_from": "d:notes/#1/body"}}])
    assert not result.ok and "did you mean 'from'" in result.error.message
    assert document(session).to_bytes() == before


def test_from_writes_the_fields_markdown(toolbox, session):
    result = ok(insert(toolbox, session, **{"from": "d:notes/#1/body"}))
    assert not [w for w in result.warnings if "read as" in w]
    paragraphs = document(session).paragraphs()[1:]
    assert [(p.style, p.text) for p in paragraphs] == [
        (None, "Opening paragraph with emphasis and 41 items."), (None, "Second paragraph."),
        ("ListBullet", "first point"), ("ListBullet", "second point"), ("ListBullet2", "detail a"),
        ("ListBullet2", "detail b"), ("ListBullet", "third point"),
        ("ListNumber", "step one"), ("ListNumber", "step two")]
    assert 'w:rStyle w:val="Emphasis"' in body_xml(session)
    # Where it went, and its length -- not its text again (rule 4).
    assert result.data["facts"] == [{"from": "d:notes/#1/body", "chars": len(BODY),
                                     "blocks": result.data["blocks"]}]


def test_a_field_is_the_same_markdown_as_typing_it(toolbox):
    typed, referenced = (_session(toolbox) for _ in range(2))
    for found in (typed, referenced):
        ok(call(toolbox, found, "ds_notes"))
    ok(insert(toolbox, typed, markdown=BODY))
    ok(insert(toolbox, referenced, **{"from": "d:notes/Doors/body"}))
    assert body_xml(typed) == body_xml(referenced)


def test_each_paragraph_keeps_the_field_as_its_source(toolbox, session):
    result = ok(insert(toolbox, session, **{"from": "d:notes/#1/body"}))
    kept = document(session).text_facts()
    assert set(kept) == set(result.data["blocks"])
    first = kept[result.data["blocks"][0]]
    assert first == [{"source": "d:notes/#1/body", "text": "Opening paragraph with emphasis and 41 items.",
                      "p": 0, "at": 0}]


def test_several_fields_in_one_call_each_with_its_blocks(toolbox, session):
    markdown = ("## Doors\n\n{{d:notes/Doors/body}}\n\n## Roof\n\n{{d:notes/Roof/body}}\n\n"
                "Rated {{d:notes/Roof/score:0.0}} overall.")
    result = ok(insert(toolbox, session, markdown=markdown))
    blocks = result.data["blocks"]
    assert len(blocks) == 1 + 9 + 1 + 1 + 1
    assert result.data["facts"] == [
        {"from": "d:notes/Doors/body", "chars": len(BODY), "blocks": blocks[1:10]},
        {"from": "d:notes/Roof/body", "text": "Short roof text.", "blocks": [blocks[11]]},
        {"from": "d:notes/Roof/score", "text": "4.0"}]          # a figure in a sentence: as before
    assert document(session).paragraph(blocks[11]).text == "Short roof text."


def test_tables_in_fields_line_up_with_their_separator(toolbox, session):
    result = ok(insert(toolbox, session, markdown="{{d:notes/Walls/body}}\n\n{{d:notes/Walls/body}}"))
    blocks = result.data["blocks"]
    assert [b[:2] for b in blocks] == ["t:", "p:", "t:", "p:"]       # a separator, a closing paragraph
    assert [f["blocks"] for f in result.data["facts"]] == [[blocks[0]], [blocks[2]]]


def test_a_value_that_is_not_text_is_refused_with_the_way_out(toolbox, session):
    error = failed(insert(toolbox, session, **{"from": "d:notes/#1/score"}), "invalid_arguments")
    assert error.message == ("d:notes/#1/score is a number value (2.456), not text: Markdown comes from a text "
                             "field (topic, body, summary); a figure goes into the markdown as {{d:notes/#1/score}}")
    assert error.valid_options == ["d:notes/#1/topic", "d:notes/#1/body", "d:notes/#1/summary",
                                   "markdown: '{{d:notes/#1/score}}'"]


@pytest.mark.parametrize("source, code, options", [
    ("d:nothing/#1/body", "not_found", ["d:notes"]),
    ("d:notes/#9/body", "not_found", ["d:notes/#1/body", "d:notes/#3/body"]),
    ("d:notes/#1/text", "not_found", ["d:notes/#1/topic", "d:notes/#1/body", "d:notes/#1/score",
                                      "d:notes/#1/summary"]),
    ("d:notes/Attic/body", "not_found", ["d:notes/Doors", "d:notes/Roof", "d:notes/Walls"]),
], ids=["dataset", "row", "column", "key"])
def test_an_unknown_dataset_row_or_field_names_what_there_is(toolbox, session, source, code, options):
    error = failed(insert(toolbox, session, **{"from": source}), code)
    assert error.field == "from" and error.valid_options == options
    assert len(document(session).paragraphs()) == 1                  # nothing written


def test_a_tables_options_with_one_field_are_refused(toolbox, session):
    error = failed(insert(toolbox, session, **{"from": "d:notes/#1/body", "from_columns": ["body"]}),
                   "invalid_arguments")
    assert error.field == "from_columns" and "one text field" in error.message


def test_rows_are_still_a_table(toolbox, session):
    result = ok(insert(toolbox, session, **{"from": "d:notes/#2:#3", "from_columns": ["topic", "score"]}))
    assert result.data["rows"] == 3 and "facts" not in result.data


def test_one_field_of_another_documents_table(toolbox, session):
    result = ok(insert(toolbox, session, **{"from": "t:checks/Alpha/note"}))
    assert result.data["facts"] == [{"from": "t:checks/Alpha/note", "text": "fine",
                                     "blocks": result.data["blocks"], "sources": ["G!E2"]}]


def test_set_text_writes_a_field_and_says_where(toolbox, session):
    target = document(session).paragraphs()[0].id
    result = ok(call(toolbox, session, "word_set_text", doc="d1", target=target,
                     text="{{d:notes/Doors/summary}}"))
    assert document(session).paragraphs()[0].text == "Doors are mostly fine."
    assert result.data["facts"] == [{"from": "d:notes/Doors/summary", "text": "Doors are mostly fine.",
                                     "blocks": [document(session).paragraphs()[0].id]}]
    # A field over two lines is two paragraphs, as a \n in text is.
    result = ok(call(toolbox, session, "word_set_text", doc="d1",
                     items=[{"target": target, "text": "{{d:notes/Roof/summary}}"}]))
    assert [p.text for p in document(session).paragraphs()] == ["Roof needs work.", "Soon."]
    assert result.data["facts"][0]["blocks"] == [p.id for p in document(session).paragraphs()]
    assert "text" not in result.data["facts"][0] and result.data["facts"][0]["chars"] == 22


def test_a_figure_in_a_sentence_is_reported_as_before(toolbox, session):
    target = document(session).paragraphs()[0].id
    result = ok(call(toolbox, session, "word_set_text", doc="d1", target=target,
                     text="Rated {{d:notes/Roof/score:0.0}}."))
    assert result.data["facts"] == [{"from": "d:notes/Roof/score", "text": "4.0"}]


# -- R66: one value per cell --------------------------------------------------------------------


def _table(toolbox, session) -> str:
    result = ok(insert(toolbox, session, markdown="| Topic | Score |\n|---|---|\n| a | b |\n| c | d |"))
    return next(b for b in result.data["blocks"] if b.startswith("t:"))


def test_a_cells_from_is_one_value_with_a_number_format(toolbox):
    session = _session(toolbox, locale=NL_BE)
    ok(call(toolbox, session, "ds_notes"))
    table = _table(toolbox, session)
    result = ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                     cells=[{"row": 1, "col": 1, "from": "d:notes/#1/score:0.0"},
                            {"row": 1, "col": 0, "from": "d:notes/#1/topic"},
                            {"row": 2, "col": 1, "text": "about {{d:notes/Roof/score:0.00}}"}]))
    view = document(session).table(table)
    assert [view.cell(r, c).text for r in (1, 2) for c in (0, 1)] == ["Doors", "2,5", "c", "about 4,00"]
    assert result.data["facts"] == [
        {"from": "d:notes/#1/score", "text": "2,5", "cell": f"{table}/c1,1"},
        {"from": "d:notes/#1/topic", "text": "Doors", "cell": f"{table}/c1,0"},
        {"from": "d:notes/Roof/score", "text": "4,00", "cell": f"{table}/c2,1"}]
    kept = document(session).text_facts()
    assert kept[view.cell(1, 1).paragraphs[0].id] == [{"source": "d:notes/#1/score", "format": "0.0",
                                                         "text": "2,5", "p": 0, "at": 0}]


def test_a_top_level_one_value_from_with_places_is_each_cells(toolbox, session):
    table = _table(toolbox, session)
    result = ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                     cells=[{"row": 2, "col": 1}], **{"from": "d:notes/#2/score:0.0"}))
    assert result.warnings[0] == "from d:notes/#2/score:0.0 (one value) read as the from of each of the 1 cell(s) given"
    assert document(session).table(table).cell(2, 1).text == "4.0"


def test_a_top_level_one_value_from_alone_gives_the_call_that_works(toolbox, session):
    table = _table(toolbox, session)
    error = failed(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                        **{"from": "d:notes/#2/score"}), "invalid_arguments")
    assert error.valid_options[0] == "cells: [{row: 1, col: 2, from: 'd:notes/#2/score'}]"
    assert "set_cells from fills rows" in error.message


@pytest.mark.parametrize("cell, field, code, options", [
    ({"row": 1, "col": 1}, "cells[1].text", "invalid_arguments", ["text", "from"]),
    ({"row": 1, "col": 1, "text": "x", "from": "d:notes/#1/score"}, "cells[1].from", "invalid_arguments",
     ["text", "from"]),
    ({"row": 1, "col": 1, "from": "d:notes/#7/score"}, "cells[1].from", "not_found",
     ["d:notes/#1/score", "d:notes/#3/score"]),
], ids=["neither", "both", "row"])
def test_a_cells_value_refusals_name_the_way_out(toolbox, session, cell, field, code, options):
    table = _table(toolbox, session)
    before = body_xml(session)
    error = failed(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                        cells=[{"row": 2, "col": 0, "text": "changed"}, cell]), code)
    assert error.field == field and error.valid_options == options
    assert body_xml(session) == before                               # all read before any changes


def test_a_strict_placeholder_beside_a_cells_from_is_no_text(toolbox, session):
    table = _table(toolbox, session)
    ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
            cells=[{"row": 1, "col": 1, "text": "", "from": "d:notes/#3/score:0.0"}]))
    assert document(session).table(table).cell(1, 1).text == "3.5"


def test_set_cells_text_without_references_is_unchanged(toolbox, session):
    table = _table(toolbox, session)
    result = ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                     cells=[{"row": 1, "col": 1, "text": "plain {{name}}"}]))
    assert "facts" not in result.data
    assert document(session).table(table).cell(1, 1).text == "plain {{name}}"


# -- round 49, R96: a cell's own from beside a top-level one-value from -------------------------


def test_a_cells_own_from_or_text_wins_and_the_top_level_value_fills_the_rest(toolbox, session):
    table = _table(toolbox, session)
    result = ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                     cells=[{"row": 1, "col": 1},
                            {"row": 2, "col": 1, "from": "d:notes/#3/score:0.0"},
                            {"row": 2, "col": 0, "text": "plain"},
                            {"row": 1, "col": 0, "text": "", "from": ""}],         # strict placeholders: none
                     **{"from": "d:notes/#2/score:0.0"}))
    view = document(session).table(table)
    assert [view.cell(r, c).text for r in (1, 2) for c in (0, 1)] == ["4.0", "4.0", "plain", "3.5"]
    assert result.warnings[0] == ("from d:notes/#2/score:0.0 (one value) read as the from of each of the 2 "
                                  "cell(s) given without their own; 2 kept their own")
    assert [fact["from"] for fact in result.data["facts"]] == ["d:notes/#2/score", "d:notes/#3/score",
                                                               "d:notes/#2/score"]


def test_the_top_level_value_fills_a_column_entry_with_rows(toolbox, session):
    table = _table(toolbox, session)
    ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
            cells=[{"col": 1, "rows": "1..2"}, {"row": 2, "col": 0, "from": "d:notes/#1/topic"}],
            **{"from": "d:notes/Walls/score:0.00"}))
    view = document(session).table(table)
    assert [view.cell(r, c).text for r in (1, 2) for c in (0, 1)] == ["a", "3.50", "Doors", "3.50"]


def test_every_cell_with_its_own_value_leaves_the_top_level_one_unused_said(toolbox, session):
    table = _table(toolbox, session)
    result = ok(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                     cells=[{"row": 1, "col": 1, "from": "d:notes/#1/topic"}], **{"from": "d:notes/#2/topic"}))
    assert document(session).table(table).cell(1, 1).text == "Doors"
    assert result.warnings[0] == ("from d:notes/#2/topic (one value) filled no cell: each of the 1 cell(s) "
                                  "gives its own from or text, which it keeps")


@pytest.mark.parametrize("source", ["d:notes", "d:notes/#1:#2"])
def test_a_top_level_from_of_rows_beside_cells_names_what_to_drop_for_each_intent(toolbox, session, source):
    table = _table(toolbox, session)
    before = body_xml(session)
    error = failed(call(toolbox, session, "word_edit_table", doc="d1", table=table, action="set_cells",
                        cells=[{"row": 1, "col": 1, "from": "d:notes/#1/score"}], **{"from": source}),
                   "invalid_arguments")
    assert error.message == (f"from {source} names rows, which fill the table from index, and cells are given "
                             f"too: to fill the table from {source}, drop cells; to set single cells, drop the "
                             "top-level from and give each cell its own value (cells: [{row: 1, col: 2, from: "
                             "'d:notes/#1/<column>'}])")
    assert error.field == "cells"
    assert error.valid_options == [source, "cells: [{row: 1, col: 2, from: 'd:notes/#1/<column>'}]"]
    assert body_xml(session) == before


# -- R68: a comment on a table or a row ---------------------------------------------------------


_SIMPLE = [(re.compile(r' w(?:14)?:(?:rsid\w*|paraId|textId)="[^"]*"'), ""),
           (re.compile(r"<w:(tblPr|tblGrid|tcPr|pPr|trPr|sectPr)>.*?</w:\1>"), ""),
           (re.compile(r'<w:rPr><w:rStyle w:val="[^"]*"/><w:sz w:val="(\d+)"/><w:szCs w:val="\d+"/></w:rPr>'),
            r"<REF \1>"),
           (re.compile(r"<w:rPr>.*?</w:rPr>"), "")]


def simple(xml: str) -> str:
    """The body's blocks, rows, cells, text and comment marks only; ids in order of appearance."""
    body = xml[xml.index("<w:body>"):xml.index("</w:body>")]
    for pattern, replacement in _SIMPLE:
        body = pattern.sub(replacement, body)
    numbers: dict[str, str] = {}
    return re.sub(r'w:id="(\d+)"', lambda m: f'w:id="{numbers.setdefault(m.group(1), str(len(numbers)))}"', body)


def three_tables() -> Document:
    document = Document.new()
    document.insert_markdown("Alpha before.\n\n| A | B |\n|---|---|\n| one | two |\n| three | four |\n\nMiddle one.\n\n"
                             "| C | D |\n|---|---|\n| five | six |\n\nMiddle two.\n\n| E | F |\n|---|---|\n"
                             "| seven | eight |\n| nine | ten |\n\nOmega after.")
    return document


#: Word 16.106's body after "select the text object of row 2 of table 3, add a comment;
#: add one with table 2 as its scope; select table 1, add a comment" (the oracle test below
#: measures it again), the ids renumbered in order of appearance as :func:`simple` does.
WORD_BODY = (
    '<w:body><w:p><w:r><w:t>Alpha before.</w:t></w:r></w:p>'
    '<w:tbl><w:tr><w:tc><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>A</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr><w:tc><w:p><w:r><w:t>one</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>two</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr><w:tc><w:p><w:r><w:t>three</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>four</w:t></w:r></w:p></w:tc>'
    '</w:tr></w:tbl><w:commentRangeEnd w:id="0"/>'
    '<w:p><w:r><REF 24><w:commentReference w:id="0"/></w:r><w:r><w:t>Middle one.</w:t></w:r></w:p>'
    '<w:tbl><w:tr><w:tc><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>C</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr><w:tc><w:p><w:r><w:t>five</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>six</w:t></w:r></w:p></w:tc>'
    '</w:tr></w:tbl><w:commentRangeEnd w:id="1"/>'
    '<w:p><w:r><REF 24><w:commentReference w:id="1"/></w:r><w:r><w:t>Middle two.</w:t></w:r></w:p>'
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>F</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:tr><w:tc><w:p><w:commentRangeStart w:id="2"/><w:r><w:t>seven</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:r><w:t>eight</w:t></w:r></w:p></w:tc></w:tr><w:commentRangeEnd w:id="2"/>'
    '<w:tr><w:tc><w:p><w:r><REF 24><w:commentReference w:id="2"/></w:r><w:r><w:t>nine</w:t></w:r></w:p></w:tc>'
    '<w:tc><w:p><w:r><w:t>ten</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    '<w:p><w:r><w:t>Omega after.</w:t></w:r></w:p>')


def _tables(document: Document) -> list[str]:
    return [b["id"] for b in document.state()["blocks"] if b["id"].startswith("t:")]


def commented(document: Document) -> Document:
    """The three comments of :data:`WORD_BODY`, through the tool."""
    one, two, three = _tables(document)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session()
        session.open(document.to_bytes(), "tables.docx")
        found = session.entry("d1").document
        row = found.table(three).rows[1].id
        result = ok(box.dispatch(session, "word_comments", {"doc": "d1", "action": "add", "items": [
            {"target": row, "text": "Row two of table three"}, {"target": two, "text": "Table two"},
            {"target": one, "text": "Whole table one"}]}))
        assert len(result.data["comments"]) == 3
        return Document.open(found.to_bytes())


def test_a_comment_on_a_table_or_row_is_anchored_as_word_anchors_one():
    done = commented(three_tables())
    xml = zipfile.ZipFile(io.BytesIO(done.to_bytes())).read("word/document.xml").decode()
    assert simple(xml) == WORD_BODY
    assert [(c.text, c.anchor.text) for c in done.comments()] == [
        ("Row two of table three", "seven\neight"), ("Table two", "C\nD\nfive\nsix"),
        ("Whole table one", "A\nB\none\ntwo\nthree\nfour")]
    assert check(done.package) == []


def test_a_comment_on_a_tables_last_row_ends_after_the_table():
    document = three_tables()
    one = _tables(document)[0]
    document.add_comment(document.table(one).rows[-1].id, "Last row")
    body = simple(zipfile.ZipFile(io.BytesIO(document.to_bytes())).read("word/document.xml").decode())
    assert ('<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>three</w:t></w:r></w:p>' in body
            and '</w:tr></w:tbl><w:commentRangeEnd w:id="0"/><w:p><w:r><REF 24><w:commentReference w:id="0"/>'
                '</w:r><w:r><w:t>Middle one.</w:t>' in body)


def test_a_cell_or_paragraph_target_is_text_as_before():
    document = three_tables()
    one = _tables(document)[0]
    created = document.add_comment(f"{one}/c1,0", "A cell").id
    assert document.comment(created).anchor.text == "one"
    body = simple(zipfile.ZipFile(io.BytesIO(document.to_bytes())).read("word/document.xml").decode())
    assert '<w:commentRangeStart w:id="0"/><w:r><w:t>one</w:t></w:r><w:commentRangeEnd w:id="0"/>' in body


@pytest.mark.parametrize("what", ["paragraph after the table", "row after the row", "the commented table",
                                  "the commented row"])
def test_deleting_around_a_commented_table_keeps_the_document_whole(what):
    done = commented(three_tables())
    one, _two, three = _tables(done)
    count = len(done.comments())
    if what == "paragraph after the table":
        after = done.state()["blocks"][2]["id"]
        assert done.paragraph(after).text == "Middle one."
        done.delete_block(after)
        # The reference opens the next paragraph, as Word's form has it.
        assert "<w:p><w:r><REF 24><w:commentReference" in simple(
            zipfile.ZipFile(io.BytesIO(done.to_bytes())).read("word/document.xml").decode())
    elif what == "row after the row":
        done.delete_row(three, 2)
        body = simple(zipfile.ZipFile(io.BytesIO(done.to_bytes())).read("word/document.xml").decode())
        assert '</w:tr></w:tbl><w:commentRangeEnd w:id="2"/><w:p><w:r><REF 24><w:commentReference w:id="2"/>' in body
    elif what == "the commented table":
        done.delete_block(one)
    else:
        done.delete_row(three, 1)
    assert len(done.comments()) == count
    assert check(done.package) == []
    Document.open(done.to_bytes())


@pytest.mark.oracle
def test_word_anchors_a_selected_table_so():
    """Word itself: select a row, a table, comment (the measurement of :data:`WORD_BODY`)."""
    import oracle

    if not oracle.available():
        pytest.skip("Microsoft Word or docx2svg's oracle scripts are not available")
    script = '''
on run argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file (item 1 of argv))
      set d to active document
      select (text object of row 2 of table 3 of d)
      make new Word comment at d with properties {comment text:"Row two of table three"}
      make new Word comment at d with properties {comment text:"Table two", scope:(text object of table 2 of d)}
      select (text object of table 1 of d)
      make new Word comment at d with properties {comment text:"Whole table one"}
      save as d file name (item 2 of argv) file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
end run
'''
    try:
        with oracle.session() as word:
            out = word.run_script(script, three_tables().to_bytes(), name="r68-table-comment", tag="word")
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))
    assert out.ok, out.detail
    xml = zipfile.ZipFile(io.BytesIO(out.path.read_bytes())).read("word/document.xml").decode()
    assert simple(xml) == WORD_BODY
