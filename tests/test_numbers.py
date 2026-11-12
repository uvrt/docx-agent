"""Numbers in a Word document's text (round 36, R28; ooxml-edit 0.84.0).

``check`` with ``include: ["numbers"]`` reads each body paragraph's and table cell's text
(``FORMAT.number_texts``) and compares its numbers with the session's values: a dataset,
and the tables of an open document whose format lists them (here a made-up grid, as a
workbook's).  A number written by reference is left out: the document keeps it on its
paragraph (``Document.text_facts``), through edits of the text around it and a save.  A
table placed with ``from``, a table of contents and the headers are not read.

All data made up: :mod:`test_tables_from_data`'s grid and dataset.
"""

from __future__ import annotations

import dataclasses

import pytest

pytest.importorskip("ooxml_edit.tools.numbers", reason="needs ooxml-edit 0.83 (numbers)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

from test_tables_from_data import COLUMNS, GRID, ROWS, CLOCK, Grid, call, ds_put, failed, ok  # noqa: E402

#: The grid as a workbook format: its tables are value sources.
SOURCES = dataclasses.replace(GRID, value_sources=lambda grid: [f"t:{name}" for name in grid.tables])


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, SOURCES], groups=GROUPS, workers=0) as box:
        yield box


def _session(toolbox, markdown: str, **options):
    session = toolbox.session(clock=CLOCK, **options)
    session.open(Document.new().to_bytes(), "report.docx")                       # d1
    session.adopt(Grid({"checks": (COLUMNS, ROWS)}), "grid", "book.grid")        # d2
    ok(call(toolbox, session, "word_insert_markdown", doc="d1", at="end", markdown=markdown))
    return session


def numbers(toolbox, session, **arguments):
    return ok(call(toolbox, session, "check", doc="d1", include=["numbers"], **arguments)).data["numbers"]


def paragraph_with(session, text):
    return next(p for p in session.entry("d1").document.paragraphs() if text in p.text)


def test_check_offers_numbers_for_a_document(toolbox):
    sent = toolbox.definitions("anthropic", groups="core")
    check = next(d for d in sent if d["name"] == "check")
    assert "numbers" in check["input_schema"]["properties"]["include"]["items"]["enum"]


def test_paragraphs_and_cells_are_read_with_their_page(toolbox):
    session = _session(toolbox, "Alpha rated 3.4 in 2026.\n\n| check | rating |\n|---|---|\n"
                                "| Gamma | 1.3 |\n")
    found = numbers(toolbox, session)
    assert (found["unverified"], found["typed"], found["values"]) == (1, 1, 3)
    first, cell = found["texts"]
    assert first["unverified"][0]["text"] == "3.4" and first["page"] == 1
    assert first["unverified"][0]["near"][0]["from"] == "t:checks/Alpha/rating"
    assert cell["at"].endswith("/c1,1") and cell["typed"][0]["from"] == "t:checks/Gamma/rating"
    assert cell["typed"][0]["fix"]["arguments"] == {
        "doc": "d1", "find": "1.3", "replace": "{{t:checks/Gamma/rating:0.0}}", "expect": "one",
        "range": cell["at"]}
    assert found["left_aside"] == {"years": 1}


def test_the_fix_writes_by_reference_and_the_number_is_left_out_after(toolbox):
    session = _session(toolbox, "Gamma came in at 1.3 points.")
    fix = numbers(toolbox, session)["texts"][0]["typed"][0]["fix"]
    ok(call(toolbox, session, fix["tool"], **fix["arguments"]))
    after = numbers(toolbox, session)
    assert (after["typed"], after["unverified"], after["texts"]) == (0, 0, [])
    document = session.entry("d1").document
    kept = document.text_facts()[paragraph_with(session, "Gamma came").id]
    assert [(f["source"], f["text"], f["at"]) for f in kept] == [("t:checks/Gamma/rating", "1.3", 17)]
    # Kept in the file: read back from its bytes.
    again = Document.open(document.to_bytes())
    assert again.text_facts() == document.text_facts()


def test_written_values_follow_their_text_and_go_when_it_changes(toolbox):
    session = _session(toolbox, "Start.")
    target = paragraph_with(session, "Start").id
    ok(call(toolbox, session, "word_set_text", doc="d1", target=target,
            text="Alpha {{t:checks/Alpha/rating}} and Beta {{t:checks/Beta/rating}}."))
    document = session.entry("d1").document
    assert [f["at"] for f in document.text_facts()[target]] == [6, 19]
    ok(call(toolbox, session, "replace_text", doc="d1", find="Alpha", replace="Now Alpha", expect="one"))
    assert [(f["text"], f["at"]) for f in document.text_facts()[target]] == [("3.5", 10), ("4", 23)]
    ok(call(toolbox, session, "replace_text", doc="d1", find="3.5", replace="3.6", expect="one"))
    assert [f["text"] for f in document.text_facts()[target]] == ["4"]            # typed over: gone
    found = numbers(toolbox, session)
    assert found["texts"][0]["unverified"][0]["text"] == "3.6"
    ok(call(toolbox, session, "undo", doc="d1"))                                    # one step back
    assert [f["text"] for f in document.text_facts()[target]] == ["3.5", "4"]
    ok(call(toolbox, session, "word_set_text", doc="d1", target=target, text="Rewritten."))
    assert target not in document.text_facts()                                     # the entry goes


def test_replace_everywhere_keeps_each_place(toolbox):
    session = _session(toolbox, "Rated 9.9 here.\n\nAnd 9.9 there.")
    ok(call(toolbox, session, "replace_text", doc="d1", find="9.9",
            replace="{{t:checks/Gamma/rating:0.00}}", expect="all"))
    kept = session.entry("d1").document.text_facts()
    assert sorted((f["text"], f["at"]) for facts in kept.values() for f in facts) == [
        ("1.25", 4), ("1.25", 6)]
    assert numbers(toolbox, session)["texts"] == []


def test_markdown_written_by_reference_is_kept_too(toolbox):
    session = _session(toolbox, "Beta came in at **{{t:checks/Beta/rating:0.0}}** points.")
    assert numbers(toolbox, session)["texts"] == []
    kept = session.entry("d1").document.text_facts()[paragraph_with(session, "Beta came").id]
    assert kept[0]["text"] == "4.0" and kept[0]["at"] == 16


def test_a_table_placed_from_a_source_and_a_range_narrow_what_is_read(toolbox):
    session = _session(toolbox, "Intro 7.7.\n\nBody 8.8.")
    ok(call(toolbox, session, "word_insert_markdown", doc="d1", at="end", **{"from": "t:checks"}))
    found = numbers(toolbox, session)
    assert [t["unverified"][0]["text"] for t in found["texts"]] == ["7.7", "8.8"]   # the table: not read
    body = paragraph_with(session, "Body").id
    scoped = numbers(toolbox, session, range=body)
    assert [t["at"] for t in scoped["texts"]] == [body]
    assert numbers(toolbox, session, pages=[2])["texts"] == []


def test_numbers_alone_runs_no_other_check_and_a_bad_range_says_the_way_out(toolbox):
    session = _session(toolbox, "Intro 7.7.")
    data = ok(call(toolbox, session, "check", doc="d1", include=["numbers"])).data
    assert set(data) == {"numbers", "validation"}     # and what saving would refuse (round 53)
    error = failed(call(toolbox, session, "check", doc="d1", include=["numbers"], range="p:7FFFFFFF"),
                   "not_found")
    assert "give a block" in error.message


def test_deleted_text_a_header_and_a_table_of_contents_are_not_read(toolbox):
    # Whole numbers on: a table of contents' page numbers would be read if it were.
    session = _session(toolbox, "# Part one\n\nIntro 7.7.", numbers={"whole_numbers": True})
    ok(call(toolbox, session, "word_set_tracking", doc="d1", on=True, author="Reviewer"))
    ok(call(toolbox, session, "replace_text", doc="d1", find="7.7", replace="6.6", expect="one"))
    ok(call(toolbox, session, "word_set_tracking", doc="d1", on=False))
    ok(call(toolbox, session, "word_headers_footers", doc="d1", section="s:body", which="header",
            action="set", text="Header score 5.5"))
    first = session.entry("d1").document.paragraphs()[0].id
    ok(call(toolbox, session, "word_fields", doc="d1", action="insert_toc", before=first))
    ok(call(toolbox, session, "word_fields", doc="d1", action="update"))
    found = numbers(toolbox, session)
    assert [n["text"] for t in found["texts"] for n in t.get("unverified", ())] == ["6.6"]
    assert "1" in session.entry("d1").document.paragraphs()[0].text                # the TOC's page
