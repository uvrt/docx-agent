"""Every operation tracked, on every fixture (ROADMAP.md, Phase E3, "Done when"): the
tracked edit saved, reopened and read back; no validity problem added (revision and comment
part integrity, unique ids); every new revision by the tracking author at its date, the
other authors' revisions untouched; the final view the untracked edit's; undo to the
original bytes and redo to the tracked ones; and accepting it reads back as the untracked
edit reads.
"""

from __future__ import annotations

import pytest

from docx_agent import Document
from docx_agent.validate import check

from test_e1_operations import OPERATIONS
from test_e3_invariant import MORE, TABLE_OPERATIONS
from test_roundtrip import entries

AUTHOR = "E3 Agent"
DATE = "2026-10-03T12:00:00Z"
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
#: Operations whose text in the original view is the original's: they change only text
#: and blocks (a formatting change's original view keeps today's formatting, by design).
#: Operations whose final view is not the untracked edit's: Word's final view (No Markup,
#: which docx2svg draws and the reader follows) shows a deleted cell, and cells merged by
#: ``w:cellMerge`` unmerged (docx2svg's ROADMAP, R.6).
STRUCTURAL = {"delete_column", "merge_across", "merge_down"}
TEXT_ONLY = {"set_text", "insert_and_delete_text", "replace_all", "replace_across_paragraphs", "move_block",
             "delete_paragraph", "insert_paragraphs"}


def _records(document: Document) -> list[tuple[str, str | None, str | None]]:
    """Every revision record's id, author and date -- but a ``w:tblGridChange``'s, which
    Word writes with an id alone (it goes with its table's cell records)."""
    return [(identifier, element.get(_W + "author"), element.get(_W + "date"))
            for _, identifier, element in document._revision_elements() if element.tag != _W + "tblGridChange"]


def _texts(document: Document, view: str, *, emphasis: bool = True) -> list[str]:
    lines = [line for line in document.to_markdown(view=view, ids=False).splitlines() if line.strip()]
    if not emphasis:
        # The original view shows today's run formatting (ROADMAP.md, E2's non-guarantees).
        lines = [line.replace("*", "").replace("~~", "") for line in lines]
    return lines


@pytest.mark.parametrize("operation", OPERATIONS + TABLE_OPERATIONS + MORE, ids=lambda f: f.__name__[3:])
def test_tracked_operation(docx_path, operation):
    data = docx_path.read_bytes()
    name = operation.__name__[3:]
    original = Document.open(data)
    before_problems = set(check(original.package))
    before_records = _records(original)

    untracked = Document.open(data)
    operation(untracked)

    document = Document.open(data)
    with document.tracking(author=AUTHOR, date=DATE):
        read_back = operation(document)
    edited = document.to_bytes()

    saved = Document.open(edited)
    assert set(check(saved.package)) - before_problems == set()
    records = _records(saved)
    ids = [identifier for identifier, _, _ in records]
    assert len(ids) == len(set(ids)), "revision ids repeat"
    ours = [r for r in records if r[1] == AUTHOR]
    assert all(date == DATE for _, _, date in ours)
    theirs = sorted((author, date) for _, author, date in records if author != AUTHOR)
    assert theirs == sorted((author, date) for _, author, date in before_records)
    assert {r.author for r in saved.revisions() if r.kind != "table-grid"} <= {AUTHOR} | {a for _, a, _ in before_records}

    # What the tracked document shows is what the untracked edit made.
    if name not in STRUCTURAL:
        assert _texts(saved, "final") == _texts(untracked, "final")
    if name in TEXT_ONLY:
        assert _texts(saved, "original", emphasis=False) == _texts(Document.open(data), "original", emphasis=False)

    accepted = Document.open(edited)
    accepted.accept(author=AUTHOR)
    if callable(read_back):
        assert read_back(Document.open(accepted.to_bytes())), f"{name} did not read back once accepted"

    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)
    while document.redo():
        pass
    assert entries(document.to_bytes()) == entries(edited)
