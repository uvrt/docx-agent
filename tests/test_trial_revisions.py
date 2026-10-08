"""Revision bookkeeping (ROADMAP.md, "Trial findings", 6): ``doc.changes()`` groups the
records into the edits a reviewer reads, every kind ``revisions()`` gives is documented, and
``TextRange.delete(collapse_space=True)`` leaves no double space behind a deleted sentence."""

from __future__ import annotations

import pytest

from docx_agent import Change, Document
from docx_agent.revisions.changes import REVISION_KINDS
from docx_agent.validate import check

DATE = "2026-10-05T09:00:00Z"
LATER = "2026-10-05T10:00:00Z"
TEXT = """# Remote working

Team leads should agree on core hours. Outside core hours, employees choose when they work.

Use the company VPN on every network. Report incidents to the service desk.

The allowance is EUR 300 a year.

Equipment is returned on leaving.

Questions go to the HR desk.
"""


def reviewed() -> Document:
    """Three people's changes, as the trial's selective-review task had them."""
    document = Document.new(created=DATE)
    document.insert_markdown(TEXT)
    ids = [p.id for p in document.paragraphs()]
    with document.tracking(author="Alice", date=DATE):
        document.anchor("should").replace("must")
        document.anchor(" Outside core hours, employees choose when they work.").delete()
        document.anchor("Report incidents to the service desk.").insert_after(" Report a lost device.")
    with document.tracking(author="Bob", date=DATE):
        document.anchor("EUR 300").replace("EUR 750")
        document.move_block(ids[4], after=ids[1])
    with document.tracking(author="Chen", date=LATER):
        document.insert_markdown("## Review summary\n\n- One.\n- Two.\n", at="end")
    return document


def test_changes_group_the_records_a_person_made_at_once():
    document = reviewed()
    found = [(c.kind, c.author, c.old_text, c.text) for c in document.changes()]
    assert found == [
        ("replacement", "Alice", "should", "must"),
        ("deletion", "Alice", " Outside core hours, employees choose when they work.", ""),
        ("move", "Bob", "", "Equipment is returned on leaving."),
        ("insertion", "Alice", "", " Report a lost device."),
        ("replacement", "Bob", "EUR 300", "EUR 750"),
        ("insertion", "Chen", "", "Review summary\nOne.\nTwo."),
    ]
    every = [r.id for r in document.revisions()]
    grouped = [r for c in document.changes() for r in c.revisions]
    assert sorted(grouped) == sorted(every) and len(grouped) == len(set(grouped))
    summary = document.changes(author="Chen")[0]
    kinds = {document.revision(r).kind for r in summary.revisions}
    # The marks and Word's paragraph-properties record at a story's end go with the text.
    assert {"insertion", "paragraph-mark-insertion", "paragraph-properties"} <= kinds


def test_a_change_is_accepted_or_rejected_whole():
    document = reviewed()
    replacement = document.changes(kind="replacement", author="Bob")[0]
    replacement.reject()
    assert any("EUR 300 a year" in p.text for p in document.paragraphs())
    assert not document.changes(kind="replacement", author="Bob")
    document.changes(kind="move")[0].accept()
    assert not document.changes(kind="move")
    for change in document.changes(author="Chen"):
        change.reject()
    assert not any("Review summary" in p.text for p in document.paragraphs())
    assert check(document.package) == []


def test_formatting_and_tables_are_changes_of_their_own():
    document = Document.new(created=DATE)
    document.insert_markdown("A paragraph to format.\n\nAnother.\n")
    first = document.paragraphs()[0].id
    with document.tracking(author="Dana", date=DATE):
        document.anchor("format").format(bold=True)
        document.insert_table(2, 2, after=first)
    found = [(c.kind, c.author) for c in document.changes()]
    assert ("formatting", "Dana") in found and ("insertion", "Dana") in found
    table = next(c for c in document.changes() if c.kind == "insertion")
    assert {document.revision(r).kind for r in table.revisions} >= {"row-insertion", "paragraph-mark-insertion"}
    assert isinstance(table, Change) and str(table).startswith("insertion by Dana")


def test_every_revision_kind_is_documented():
    """Each kind revisions() can give is in REVISION_KINDS and Revision.kind's docstring."""
    from docx_agent import Revision
    from docx_agent.edit.annotations import REVISION_KINDS as TAGS

    families = set(TAGS.values())
    marks = {f"{where}-{kind}" for where in ("paragraph-mark", "run-mark")
             for kind in ("insertion", "deletion", "move-from", "move-to")}
    rows = {"row-insertion", "row-deletion"}
    assert families | marks | rows == set(REVISION_KINDS)
    for kind in REVISION_KINDS:
        assert kind in Revision.kind.__doc__, kind
    document = reviewed()
    assert {r.kind for r in document.revisions()} <= set(REVISION_KINDS)


# -- delete(collapse_space=True) ----------------------------------------------------------------

SENTENCES = "Within 30 days of notice. The supplier waives the fee. The customer may end it.\n"


@pytest.mark.parametrize("anchor,expected", [
    ("The supplier waives the fee.", "Within 30 days of notice. The customer may end it."),
    ("Within 30 days of notice.", "The supplier waives the fee. The customer may end it."),
    ("The customer may end it.", "Within 30 days of notice. The supplier waives the fee."),
    ("waives ", "Within 30 days of notice. The supplier the fee. The customer may end it."),
    (" the fee", "Within 30 days of notice. The supplier waives. The customer may end it."),
])
@pytest.mark.parametrize("tracked", [False, True], ids=["untracked", "tracked"])
def test_delete_collapses_the_space_a_sentence_leaves(anchor, expected, tracked):
    document = Document.new(created=DATE)
    document.insert_markdown(SENTENCES)
    if tracked:
        with document.tracking(author="Claude", date=DATE):
            document.anchor(anchor).delete(collapse_space=True)
        assert len(document.revisions()) == 1               # the space is part of the one deletion
    else:
        document.anchor(anchor).delete(collapse_space=True)
    assert document.paragraphs()[0].text == expected
    assert check(document.package) == []


def test_delete_keeps_the_space_by_default():
    document = Document.new(created=DATE)
    document.insert_markdown(SENTENCES)
    document.anchor("The supplier waives the fee.").delete()
    assert "notice.  The customer" in document.paragraphs()[0].text
