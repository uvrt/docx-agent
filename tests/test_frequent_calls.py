"""Calls a production run made often and had refused (round 31, R5): ``word_delete`` with
``find`` and ``target`` (find searched within the target), a ``caption`` in
``word_insert_markdown``'s ``style_map`` (a caption paragraph in Markdown), ``word_inspect``
of a cell the table lacks (the table's rows, columns and merged cells said) and
``word_format_table`` ``keep_together`` on a one-row table (``test_table_keeps.py``).
Made-up documents, ``Document.new()``'s package."""

from __future__ import annotations

import pytest
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.edit.document import NoCell
from docx_agent.tools import FORMAT, GROUPS, TOOLS

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, result.error and result.error.to_json())
    return result


def opened(box, document: Document):
    session = box.session()
    return session, session.open(document.to_bytes(), "made-up.docx")


def two_paragraphs() -> Document:
    document = Document.new()
    first = document.paragraphs()[0].id
    document.insert_paragraph("Shelf one holds the spare parts", after=first)
    document.insert_paragraph("Shelf two holds the spare tools", after=document.paragraphs()[1].id)
    return document


# -- R5a: word_delete with find and target --------------------------------------------------------


def test_find_with_a_target_is_find_within_it(box):
    session, d = opened(box, two_paragraphs())
    document = session.entry(d).document
    second = document.paragraphs()[2].id
    result = ok(box, session, "word_delete", doc=d, find="spare ", target=second)
    assert [p.text for p in document.paragraphs()[1:]] == ["Shelf one holds the spare parts",
                                                           "Shelf two holds the tools"]
    assert result.warnings == [f"find was searched within target {second}: within names where to search"]


def test_find_with_target_and_another_within_is_refused_and_neither_is_too(box):
    session, d = opened(box, two_paragraphs())
    document = session.entry(d).document
    first, second = (p.id for p in document.paragraphs()[1:])
    refused = box.dispatch(session, "word_delete", {"doc": d, "find": "spare", "target": second, "within": first})
    assert not refused.ok and refused.error.code == "invalid_arguments" and refused.error.field == "target"
    # The same block twice is one place to search.
    ok(box, session, "word_delete", doc=d, find="spare ", target=second, within=second)
    refused = box.dispatch(session, "word_delete", {"doc": d})
    assert not refused.ok and refused.error.valid_options == ["find", "target"]


def test_a_target_alone_still_deletes_the_block(box):
    session, d = opened(box, two_paragraphs())
    document = session.entry(d).document
    second = document.paragraphs()[2].id
    result = ok(box, session, "word_delete", doc=d, target=second)
    assert result.warnings == [] and len(document.paragraphs()) == 2


# -- R5b: a caption in Markdown -------------------------------------------------------------------


CAPTIONED = """Before the table.

| Item | Count |
|------|-------|
| Bolts | 12 |

Table: Parts on the shelf

: Shelf layout, seen from the door

Table: a plain paragraph, no table beside it

:no space after the colon
"""


def test_a_caption_paragraph_is_written_in_the_caption_style(box):
    session, d = opened(box, Document.new())
    document = session.entry(d).document
    result = ok(box, session, "word_insert_markdown", doc=d, at="end", markdown=CAPTIONED,
                style_map=[{"element": "caption", "style": "Caption"}])
    assert "style added: Caption (Caption), as Word writes it" in result.warnings
    body = [(p.text, p.style) for p in document.paragraphs() if p.id in result.data["blocks"]]
    assert body == [("Before the table.", None), ("Parts on the shelf", "Caption"),
                    ("Shelf layout, seen from the door", "Caption"),
                    ("Table: a plain paragraph, no table beside it", None),
                    (":no space after the colon", None)]


def test_the_caption_style_by_default_and_another_by_the_map():
    document = Document.new()
    document.styles.add("Figure note", based_on="Normal")
    document.insert_markdown(": First\n")
    document.insert_markdown(": Second\n", style_map={"caption": "Figure note"})
    styles = {p.text: p.style for p in document.paragraphs()}
    assert styles["First"] == "Caption"
    assert document.styles.find(styles["Second"]).name == "Figure note"


def test_a_caption_in_a_quote_or_list_is_text_and_reading_is_unchanged():
    document = Document.new()
    document.insert_markdown("> : quoted\n\n- : listed\n\n: Shown\n")
    texts = [p.text for p in document.paragraphs()]
    assert ": quoted" in texts and ": listed" in texts and "Shown" in texts
    # The reader shows a caption paragraph's style on its block, as before round 31.
    assert "style: caption -->\nShown" in document.to_markdown()


# -- R5c: a cell the table lacks ------------------------------------------------------------------


def merged_table() -> tuple[Document, str]:
    document = Document.new()
    first = document.paragraphs()[0].id
    table = document.insert_table(2, 3, after=first, data=[["a", "b", "c"], ["d", "e", "f"]]).id
    document.merge_cells(table, (0, 0), (1, 0))
    document.merge_cells(table, (0, 1), (0, 2))
    return document, document.tables()[0].id


def test_a_missing_cell_says_the_rows_columns_and_merged_cells(box):
    document, table = merged_table()
    with pytest.raises(NoCell) as raised:
        document.table(table).cell(1, 3)
    message = str(raised.value.args[0])
    assert message == (f"{table} has no cell at (1, 3): it has 2 row(s) (0 to 1) and 3 column(s) (0 to 2), "
                       "counted from 0; merged: (0, 0)-(1, 0), (0, 1)-(0, 2) (a merged cell's id is its "
                       "first position's)")
    session, d = opened(box, document)
    refused = box.dispatch(session, "word_inspect", {"doc": d, "target": f"{table}/c2,0"})
    assert refused.error.code == "not_found" and "it has 2 row(s) (0 to 1) and 3 column(s)" in refused.error.message
    assert refused.error.valid_options == [f"{table}/c0,0", f"{table}/c0,1", f"{table}/c1,1", f"{table}/c1,2"]
    # The table's own edits say the same.
    refused = box.dispatch(session, "word_edit_table", {"doc": d, "table": table, "action": "split",
                                                        "first": {"row": 4, "col": 0}})
    assert not refused.ok and "it has 2 row(s)" in refused.error.message
