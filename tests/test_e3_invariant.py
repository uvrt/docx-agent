"""The tracked-change invariant (ROADMAP.md, "Accept and reject"): for every operation and
fixture, accepting every revision of the tracked edit gives the untracked edit, and
rejecting them gives the original document -- in canonical form (``canonical.py``).

Every E1 operation (``test_e1_operations.OPERATIONS``) and E3's table operations, on every
fixture: the operation untracked; the same operation inside ``Document.tracking``, saved
and reopened (so nothing rides on the session but what the file holds), then
``accept_all`` and ``reject_all``.  The original is the document as the first edit stamps
it (the whole-document stamp is not a revision).  Bookmarks and style definitions are not
revisions in Word, and rejecting keeps them: the comparison forgives a bookmark's markers
for the operations that set one, and a style definition the edit added.
"""

from __future__ import annotations

import pytest

from docx_agent import Document, EditError
from docx_agent.validate import check

from canonical import canonical, difference, without_added_styles
from test_e1_operations import OPERATIONS, plain, text_of

AUTHOR = "E3 Agent"
DATE = "2026-10-03T12:00:00Z"


def _table_of(document: Document):
    for table in document.tables():
        rows = table.rows
        grid = table.grid()
        if len(rows) >= 2 and grid and all(len(line) == len(grid[0]) and all(c is not None for c in line)
                                           for line in grid) and len(grid[0]) >= 2:
            spans = {c.id for line in grid for c in line}
            if len(spans) == len(grid) * len(grid[0]):
                return table
    paragraph = text_of(document)
    return None


def op_insert_row(document):
    table = _table_of(document)
    if table is None:
        pytest.skip("no plain table")
    document.insert_row(table.id, 0)


def op_delete_row(document):
    table = _table_of(document)
    if table is None:
        pytest.skip("no plain table")
    document.delete_row(table.id, len(table.rows) - 1)


def op_insert_column(document):
    table = _table_of(document)
    if table is None:
        pytest.skip("no plain table")
    document.insert_column(table.id, 0)


def op_delete_column(document):
    table = _table_of(document)
    if table is None:
        pytest.skip("no plain table")
    document.delete_column(table.id, 1)


def op_merge_across(document):
    table = _table_of(document)
    if table is None:
        pytest.skip("no plain table")
    document.merge_cells(table.id, (0, 0), (0, 1))


def op_merge_down(document):
    table = _table_of(document)
    if table is None:
        pytest.skip("no plain table")
    document.merge_cells(table.id, (0, 0), (1, 0))


def op_delete_paragraph(document):
    candidates = [p for p in plain(document, 4) if not p.ends_section]
    for paragraph in reversed(candidates):
        try:
            return document.delete_block(paragraph.id)
        except EditError:
            continue
    pytest.skip("no paragraph to delete")


def op_insert_paragraphs(document):
    paragraph = text_of(document)
    first = document.insert_paragraph("Inserted before.", before=paragraph.id)
    document.insert_paragraph("Inserted after.", after=paragraph.id, style="Heading 2")
    last = document.paragraphs()[-1]
    if not last.ends_section:
        document.insert_paragraph("Inserted at the end.", after=last.id)
    return first


TABLE_OPERATIONS = [op_insert_row, op_delete_row, op_insert_column, op_delete_column, op_merge_across,
                    op_merge_down]
MORE = [op_delete_paragraph, op_insert_paragraphs]
#: Operations whose result Word does not keep as revisions in part, and what the
#: comparison forgives for them.
FORGIVE_BOOKMARKS = {"bookmarks", "hyperlinks"}
#: Operations where which paragraph element survives is the edit's choice, not content.
WITHOUT_IDS = {"move_block", "replace_across_paragraphs", "merge_down", "merge_across", "delete_paragraph",
               "insert_paragraphs", "lists"}


def _name(operation) -> str:
    return operation.__name__[3:]


def _comments_described(form: dict[str, str]) -> dict[str, str]:
    return form


#: (fixture, operation) whose reject-all differs from the original by a paraId alone, recorded
#: in ROADMAP.md (Phase E6): in a new document, whose only paragraph is empty, the paragraph
#: the operation writes text into and the two typed after it are all insertions, and the
#: one left after Reject All is the last of them (which carried the old mark), not the first.
ID_ONLY = {("document", "list_continue_and_remove")}


@pytest.mark.parametrize("operation", OPERATIONS + TABLE_OPERATIONS + MORE, ids=_name)
def test_accept_all_is_the_edit_and_reject_all_the_original(docx_path, operation, request):
    data = docx_path.read_bytes()
    name = _name(operation)
    if (docx_path.stem, name) in ID_ONLY and docx_path.parent.name == "new":
        request.applymarker(pytest.mark.xfail(strict=True, reason="reject-all keeps another paraId (ID_ONLY)"))
    ids = name not in WITHOUT_IDS

    base = Document.open(data)
    base._stamp_document()
    original = canonical(base.to_bytes(), ids=ids)

    untracked = Document.open(data)
    operation(untracked)
    expected = canonical(untracked.to_bytes(), ids=ids)

    tracked = Document.open(data)
    with tracked.tracking(author=AUTHOR, date=DATE):
        operation(tracked)
    tracked_bytes = tracked.to_bytes()
    assert set(check(Document.open(tracked_bytes).package)) - set(check(Document.open(data).package)) == set()

    accepted = Document.open(tracked_bytes)
    accepted.accept(author=AUTHOR)
    assert not accepted.revisions(author=AUTHOR)
    got = without_added_styles(canonical(accepted.to_bytes(), ids=ids), expected)
    assert got == expected, "accept-all is not the untracked edit:\n" + difference(expected, got)

    rejected = Document.open(tracked_bytes)
    rejected.reject(author=AUTHOR)
    assert not rejected.revisions(author=AUTHOR)
    bookmarks = name not in FORGIVE_BOOKMARKS
    got = canonical(rejected.to_bytes(), ids=ids, bookmarks=bookmarks)
    want = original if bookmarks else canonical(base.to_bytes(), ids=ids, bookmarks=False)
    got = without_added_styles(got, want)
    assert got == want, "reject-all is not the original:\n" + difference(want, got)


def test_the_whole_edit_set_accepts_to_its_edits_and_rejects_to_the_original(docx_path):
    """E3's representative edit set (``e3_edits.py``, its tracked part) composed: each edit
    on what the ones before it made -- an insertion linked, a picture beside inserted text,
    a column inserted in a table whose rows were inserted and deleted."""
    from e3_edits import AUTHOR as SET_AUTHOR, e3_edit_set, prepare

    data = docx_path.read_bytes()
    base = Document.open(data)
    prepare(base)
    base._stamp_document()
    original = canonical(base.to_bytes(), ids=False)
    untracked = Document.open(data)
    e3_edit_set(untracked, comments=False, tracked=False)
    expected = canonical(untracked.to_bytes(), ids=False)
    tracked = Document.open(data)
    e3_edit_set(tracked, comments=False)
    tracked_bytes = tracked.to_bytes()

    accepted = Document.open(tracked_bytes)
    accepted.accept(author=SET_AUTHOR)
    got = without_added_styles(canonical(accepted.to_bytes(), ids=False), expected)
    assert got == expected, "accept-all is not the untracked edit set:\n" + difference(expected, got)

    rejected = Document.open(tracked_bytes)
    rejected.reject(author=SET_AUTHOR)
    got = without_added_styles(canonical(rejected.to_bytes(), ids=False), original)
    assert got == original, "reject-all is not the original:\n" + difference(original, got)
