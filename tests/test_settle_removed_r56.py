"""Round 56 (R121): what an accept or reject removed, whose changes it settled, and a preview.

``word_changes`` reject all removed a tracked inserted table an agent had made and said
"Rejected 16 revision record(s)", naming paragraph ids only.  Now every path (ids, author,
kind, within, all; accept and reject) names each block it removed: the summary names the
table by its ``t:`` id with its rows and paragraphs counted, rows of a table that stays with
that table, paragraphs (one whose mark went: the paragraph it merged into), content controls;
``data.removed`` has them typed; the envelope's ``removed`` lists every id, rows and the
paragraphs inside a removed table too.  A call by author, kind, within or all says how many
of the records it settled were written this session -- since the document was opened, by
the records' author and date against the file as read (not by the tracking author, which an
earlier session's records share).  ``preview`` says what accept and what reject would settle
and remove, on copies: nothing changes.  Made-up documents only.

Tool-design rules: 2 (a removed table is named with the id the agent holds), 4 (the summary
names a few blocks and says where the rest are; data.removed has them all), 5 (preview
before a bulk accept or reject), 3 (preview and a call settling nothing change nothing).
"""

from __future__ import annotations

import datetime as dt

import pytest
from ooxml_edit.tools import Toolbox
from test_rows_joins_r54r import _document, _p

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

CLOCK = lambda: dt.datetime(2026, 11, 13, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
LATER = lambda: dt.datetime(2026, 11, 13, 10, 30, tzinfo=dt.timezone.utc)  # noqa: E731
ME = "Editor A"
TABLE = "| Area | Lead |\n|---|---|\n| Harbour | Ines |\n| Orchard | Tomas |"


@pytest.fixture(scope="module")
def box():
    with Toolbox(list(TOOLS), formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def ok(box, session, name, /, **arguments):
    result = box.dispatch(session, name, {"doc": "d1", **arguments})
    assert result.ok, (name, result.to_json())
    return result


def _open(box, data: bytes | None = None, clock=CLOCK):
    session = box.session(clock=clock)
    session.open(data or Document.new().to_bytes(), "plan.docx")
    return session


def _with_new_table(box):
    """A kept table, then -- tracked -- a paragraph, a new table and a paragraph after it."""
    session = _open(box)
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", at=f"after:{first}",
       markdown=f"Opening line.\n\n{TABLE}\n\nKept closing.")
    ok(box, session, "word_set_tracking", on=True, author=ME)
    made = ok(box, session, "word_insert_markdown", at="end", markdown=f"New intro.\n\n{TABLE}\n\nNew closing.")
    table = next(i for i in made.created if i.startswith("t:"))
    return session, table


def _rows_and_paragraphs(session, table: str) -> tuple[list[str], list[str]]:
    from docx_agent.edit.ids import W_P

    document = session.entry("d1").document
    part, entry = document._table_entry(table)
    inside = set(map(id, entry.element.iter(W_P)))
    paragraphs = [p.id for p in document._index(part).paragraphs if id(p.element) in inside]
    return [row_id for _, row_id in entry.rows], paragraphs


@pytest.mark.parametrize("how", [{"all": True}, {"author": ME}, {"kind": "insertion"}, "within", "ids"])
def test_a_reject_names_the_table_it_removed(box, how):
    session, table = _with_new_table(box)
    rows, paragraphs = _rows_and_paragraphs(session, table)
    if how == "within":
        how = {"within": table}
    elif how == "ids":
        listed = ok(box, session, "word_changes", action="list").data
        how = {"ids": [c["id"] for c in listed if "Harbour" in (c.get("text") or "")]}
    result = ok(box, session, "word_changes", action="reject", **how)
    assert f"removed table {table} ({len(rows)} rows, {len(paragraphs)} paragraphs)" in result.summary
    assert result.data["removed"]["tables"] == [{"id": table, "rows": len(rows), "paragraphs": len(paragraphs)}]
    # The envelope: every id, the table's rows and paragraphs too.
    assert {table, *rows, *paragraphs} <= set(result.removed)
    assert result.removed[0] == table
    assert table not in [t.id for t in session.entry("d1").document.tables()]


def test_a_broad_call_says_how_many_records_this_session_wrote(box):
    session, table = _with_new_table(box)
    result = ok(box, session, "word_changes", action="reject", all=True)
    records = result.data["records"]
    assert records > 0 and result.data["this_session"] == records
    assert result.summary.startswith(f"Rejected {records} revision record(s) (all written this session")
    # By ids the agent named them: the count is in data, the summary does not repeat it.
    session, table = _with_new_table(box)
    listed = ok(box, session, "word_changes", action="list").data
    by_id = ok(box, session, "word_changes", action="reject", ids=[listed[0]["id"]])
    assert by_id.data["this_session"] == by_id.data["records"] and "this session" not in by_id.summary


def test_records_an_earlier_session_wrote_are_not_this_ones(box):
    """The tracking author is the same in both sessions: only the records' dates against
    the file as opened tell the earlier ones apart."""
    earlier, _ = _with_new_table(box)
    data = earlier.entry("d1").document.to_bytes()
    before = len(Document.open(data).revisions())
    session = _open(box, data, clock=LATER)
    ok(box, session, "word_set_tracking", on=True, author=ME)
    ok(box, session, "word_insert_markdown", at="end", markdown="Second session line.")
    added = len(session.entry("d1").document.revisions()) - before
    result = ok(box, session, "word_changes", action="reject", author=ME)
    assert result.data["records"] == before + added
    assert result.data["this_session"] == added > 0
    assert f"({added} written this session, since the document was opened)" in result.summary


def test_an_edit_inside_an_earlier_insertion_keeps_it_the_earlier_ones(box):
    """A record an edit split (its author and date kept) was there when the file was read."""
    body = _p("Before.") + ('<w:p><w:ins w:id="5" w:author="Editor B" w:date="2026-01-02T00:00:00Z">'
                            '<w:r><w:t>Old inserted words here.</w:t></w:r></w:ins></w:p>')
    session = _open(box, _document(body))
    target = session.entry("d1").document.paragraphs()[1].id
    ok(box, session, "word_set_tracking", on=True, author=ME)
    ok(box, session, "replace_text", find="inserted", replace="added", range=target)
    revisions = session.entry("d1").document.revisions()
    earlier = [r for r in revisions if r.author == "Editor B"]
    assert len(earlier) >= 1 and len(revisions) > len(earlier)
    result = ok(box, session, "word_changes", action="accept", all=True)
    assert result.data["records"] == len(revisions)
    assert result.data["this_session"] == len(revisions) - len(earlier)


def test_accepting_a_tracked_table_deletion_names_the_table(box):
    session, _ = _with_new_table(box)
    kept = session.entry("d1").document.tables()[0].id
    rows, paragraphs = _rows_and_paragraphs(session, kept)
    ok(box, session, "word_delete", target=kept)
    result = ok(box, session, "word_changes", action="accept", within=kept)
    assert result.data["removed"]["tables"] == [{"id": kept, "rows": len(rows), "paragraphs": len(paragraphs)}]
    assert f"removed table {kept} (" in result.summary


def test_rows_removed_from_a_table_that_stays_are_named_with_it(box):
    session = _open(box)
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", at=f"after:{first}", markdown=f"Opening.\n\n{TABLE}")
    ok(box, session, "word_set_tracking", on=True, author=ME)
    table = session.entry("d1").document.tables()[0].id
    ok(box, session, "word_edit_table", table=table, action="insert_row", index=1)
    top = ok(box, session, "word_edit_table", table=table, action="insert_row", index=0, side="before")
    renamed = top.data["id"]
    assert renamed != table              # a table is named by its first row
    added = [r for _, r in session.entry("d1").document._table_entry(renamed)[1].rows][:1]
    result = ok(box, session, "word_changes", action="reject", all=True)
    rows = result.data["removed"]["rows"]
    assert len(rows) == 2 and all(r["table"] == table and r["paragraphs"] == 2 for r in rows)
    assert added[0] in [r["id"] for r in rows]
    assert "tables" not in result.data["removed"]
    assert f"of {table}" in result.summary
    # Its first row gone, the table has its old name again: said, and both names find it.
    assert result.renamed == {renamed: table}
    document = session.entry("d1").document
    assert document.get(renamed).id == document.get(table).id == table


def test_a_paragraph_whose_mark_went_names_the_one_it_merged_into(box):
    session, table = _with_new_table(box)
    document = session.entry("d1").document
    kept_closing = next(p.id for p in document.paragraphs() if p.text == "Kept closing.")
    result = ok(box, session, "word_changes", action="reject", all=True)
    merged = result.data["removed"]["merged"]
    survivor = merged[kept_closing]
    assert session.entry("d1").document.get(survivor).text == "Kept closing."
    assert f"paragraph {kept_closing} (merged into {survivor})" in result.summary


def test_the_summary_names_a_few_blocks_and_says_where_the_rest_are(box):
    session = _open(box)
    ok(box, session, "word_set_tracking", on=True, author=ME)
    ok(box, session, "word_insert_markdown", at="end",
       markdown="\n\n".join(f"{TABLE.replace('Harbour', f'Harbour {k}')}\n\nBetween {k}." for k in range(8)))
    result = ok(box, session, "word_changes", action="reject", all=True)
    assert len(result.data["removed"]["tables"]) == 8
    assert "more (data.removed)" in result.summary
    assert result.summary.count("table t:") == 6


def test_a_preview_changes_nothing_and_says_what_each_would_do(box):
    session, table = _with_new_table(box)
    entry = session.entry("d1")
    version, data = entry.version, entry.document.to_bytes()
    preview = ok(box, session, "word_changes", action="preview", all=True)
    assert entry.version == version and entry.document.to_bytes() == data
    assert preview.summary.startswith("Preview, nothing changed:")
    assert preview.data["accept"].get("removed") is None
    assert preview.data["reject"]["removed"]["tables"][0]["id"] == table
    assert preview.data["changes"] and all("id" in c for c in preview.data["changes"])
    # What it said is what the call does.
    rejected = ok(box, session, "word_changes", action="reject", all=True)
    expected = preview.data["reject"]
    assert {k: rejected.data[k] for k in ("records", "this_session", "changes_left", "removed")} == expected


def test_a_preview_takes_the_same_selection_and_refusals(box):
    session, table = _with_new_table(box)
    within = ok(box, session, "word_changes", action="preview", within=table)
    assert within.data["reject"]["removed"]["tables"][0]["id"] == table
    refused = box.dispatch(session, "word_changes", {"doc": "d1", "action": "preview"})
    assert refused.error.code == "invalid_arguments" and "ids" in refused.error.valid_options
    both = box.dispatch(session, "word_changes", {"doc": "d1", "action": "preview", "ids": ["rev:0"], "all": True})
    assert both.error.code == "invalid_arguments"
    nothing = ok(box, session, "word_changes", action="preview", author="Nobody")
    assert nothing.data["accept"]["records"] == nothing.data["reject"]["records"] == 0
    assert "remove no block" in nothing.summary


def test_settling_nothing_changes_nothing(box):
    session, _ = _with_new_table(box)
    result = ok(box, session, "word_changes", action="reject", author="Nobody")
    assert result.data["records"] == 0 and "removed" not in result.data and result.removed == []


def test_the_library_says_what_a_review_removed():
    """Document.reject_all's EditResult: removed_blocks, a content control in a removed
    paragraph among them, and ids as the caller had them before stamping."""
    para = ('<w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="Editor B" w:date="2026-01-02T00:00:00Z"/></w:rPr></w:pPr>'
            '<w:ins w:id="2" w:author="Editor B" w:date="2026-01-02T00:00:00Z"><w:r><w:t>Name: </w:t></w:r></w:ins></w:p>')
    table = ('<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr>'
             '<w:ins w:id="3" w:author="Editor B" w:date="2026-01-02T00:00:00Z"/></w:trPr><w:tc><w:tcPr/>'
             '<w:sdt><w:sdtPr><w:id w:val="78"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Cell</w:t></w:r></w:p>'
             '</w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>')
    document = Document.open(_document(_p("Before.") + para + table + _p("After.")))
    before = [p.id for p in document.paragraphs()]
    tables = [t.id for t in document.tables()]
    result = document.reject_all()
    assert result.removed_blocks["tables"] == [{"id": tables[0], "rows": 1, "paragraphs": 1}]
    assert before[1] in result.removed_blocks.get("merged", {}) or before[1] in result.removed_blocks.get(
        "paragraphs", [])
    assert result.removed[0] == tables[0] and "cc:78" in result.removed
    assert len(result.removed) == len(set(result.removed))
    assert document.tables() == []
