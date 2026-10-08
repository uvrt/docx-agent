"""Moving a section (ROADMAP.md, "Trial findings", 5): ``section_blocks(heading)`` is a heading
and everything up to the next heading of its level or higher; ``move_blocks`` moves a range
or a list of blocks together, untracked or as tracked moves.  The trial's agents chained six
``move_block`` calls for one section."""

from __future__ import annotations

import pytest

from docx_agent import Document, EditError
from docx_agent.validate import check

DATE = "2026-10-05T09:00:00Z"
POLICY = """# 1 Purpose

Why the policy exists.

# 2 Access control

Who may enter.

# 3 Data retention

How long records are kept.

## 3.1 Retention periods

| Record | Years |
| --- | --- |
| Invoices | 7 |

## 3.2 Disposal

Shred paper, wipe disks.

# 4 Review

Reviewed every year.
"""


def policy() -> Document:
    document = Document.new(created=DATE)
    document.insert_markdown(POLICY)
    return document


def heading(document: Document, text: str) -> str:
    return next(p.id for p in document.paragraphs() if p.text == text)


def order(document: Document) -> list[str]:
    return [p.text for p in document.paragraphs() if p.style_name and p.style_name.startswith("heading")]


def test_section_blocks_stops_at_the_next_heading_of_its_level_or_higher():
    document = policy()
    section = document.section_blocks(heading(document, "3 Data retention"))
    texts = [document.get(i).text if i.startswith("p:") else "<table>" for i in section]
    assert texts == ["3 Data retention", "How long records are kept.", "3.1 Retention periods", "<table>",
                     "3.2 Disposal", "Shred paper, wipe disks."]
    sub = document.section_blocks(heading(document, "3.1 Retention periods"))
    assert len(sub) == 2 and sub[1].startswith("t:")
    last = document.section_blocks(heading(document, "4 Review"))
    assert [document.get(i).text for i in last] == ["4 Review", "Reviewed every year."]
    with pytest.raises(EditError, match="not a heading"):
        document.section_blocks(document.paragraphs()[1].id)


def test_move_blocks_moves_a_section_in_one_step():
    document = policy()
    before = document.to_bytes()
    section = document.section_blocks(heading(document, "3 Data retention"))
    target = document.section_blocks(heading(document, "1 Purpose"))[-1]
    result = document.move_blocks(section, after=target)
    assert order(document) == ["1 Purpose", "3 Data retention", "3.1 Retention periods", "3.2 Disposal",
                               "2 Access control", "4 Review"]
    assert result.blocks == section                               # every block keeps its id
    assert check(document.package) == []
    document.undo()                                               # one undo step
    assert document.to_bytes() == before


def test_move_blocks_takes_a_range_and_before():
    document = policy()
    first = heading(document, "2 Access control")
    last = document.section_blocks(first)[-1]
    document.move_blocks(f"{first}..{last}", before=heading(document, "4 Review"))
    assert order(document) == ["1 Purpose", "3 Data retention", "3.1 Retention periods", "3.2 Disposal",
                               "2 Access control", "4 Review"]


def test_move_blocks_refuses_a_target_inside_what_it_moves():
    document = policy()
    section = document.section_blocks(heading(document, "3 Data retention"))
    before = document.to_bytes()
    with pytest.raises(EditError, match="one of the blocks moved"):
        document.move_blocks(section, after=section[2])
    with pytest.raises(EditError, match="exactly one"):
        document.move_blocks(section)
    assert document.to_bytes() == before


def test_a_tracked_section_move_accepts_to_the_move_and_rejects_to_the_original():
    untracked = policy()
    untracked.move_blocks(untracked.section_blocks(heading(untracked, "3 Data retention")),
                          after=untracked.section_blocks(heading(untracked, "1 Purpose"))[-1])
    document = policy()
    original = document.to_markdown(ids=False)
    with document.tracking(author="Claude", date=DATE):
        document.move_blocks(document.section_blocks(heading(document, "3 Data retention")),
                             after=document.section_blocks(heading(document, "1 Purpose"))[-1])
    assert check(document.package) == []
    kinds = [c.kind for c in document.changes()]
    assert kinds == ["move"]                                      # the section, its table too: one change
    assert "{++" in document.to_markdown(view="markup")
    accepted = Document.open(document.to_bytes())
    accepted.accept_all()
    assert accepted.to_markdown(ids=False) == untracked.to_markdown(ids=False)
    rejected = Document.open(document.to_bytes())
    rejected.reject_all()
    assert rejected.to_markdown(ids=False) == original
