"""Paragraphs a deleted mark joins in the final view, and blocks wholly deleted (round 50, R99),
as Word keeps them (``tools/tracking_probe.py``, ``tests/observations/tracking-word.json``):
accepting a deleted mark keeps the *last* paragraph -- its element, its id and its
properties (which Word, tracking the deletion, set to the first one's); the final view (No
Markup) draws the same.  The view names the joined paragraph by it, ``joins`` the ones
before; :meth:`Document.blocks` says ``deleted`` for a block wholly pending deletion.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pytest

from docx_agent import Document

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import tracking_probe  # noqa: E402

OBSERVED = json.loads((ROOT / "tests" / "observations" / "tracking-word.json").read_text(encoding="utf-8"))
_JC = {"center": "center", "right": "right", None: "left"}


def _tracked_joins() -> Document:
    """The probe's document with the deletions Word made by hand made here, tracked: a mark
    deleted (Jone), a whole paragraph deleted (Kone); Lone and Mone come deleted in the file."""
    document = Document.open(tracking_probe.join_document())
    ids = {p.text.split()[0]: p.id for p in document.paragraphs() if p.text}
    end = len(document.paragraph(ids["Jone"]).text)
    with document.tracking(author="Reviewer", date="2026-11-03T10:00:00Z"):
        document.range(f"{ids['Jone']}@{end}..{ids['Jtwo']}@0").delete()
        document.delete_block(ids["Kone"])
    return document


def _accepted_by_word() -> list[tuple[str, str]]:
    """What Word's Accept All kept: each paragraph's id and alignment."""
    out = []
    for block in OBSERVED["join-accept"]:
        found = re.search(r'w14:paraId="([0-9A-F]{8})"', block)
        if found is None or "Tail end." in block:
            continue
        jc = re.search(r'<w:jc w:val="(\w+)"/>', block)
        out.append((f"p:{found.group(1)}", _JC[jc.group(1) if jc else None]))
    return out


def test_word_keeps_the_last_paragraph_and_its_properties():
    # The facts: Jone's mark deleted by Word gives Jtwo Jone's properties (a w:pPrChange);
    # accepting keeps Jtwo's element.  A deleted mark with no w:pPrChange (Lone) keeps Ltwo's.
    assert _accepted_by_word() == [("p:30000001", "center"), ("p:30000003", "right"), ("p:30000005", "right"),
                                   ("p:30000007", "right")]
    # The final view draws Lone's text on Ltwo's right-aligned line: it ends at the right
    # margin, as the right-aligned paragraphs do; Jone's centred line does not.
    final = OBSERVED["join-final"]
    assert abs(final["Lone"][1] - final["Ktwo"][1]) < 0.5 and final["Jone"][1] < final["Ktwo"][1] - 100


def test_our_tracked_deletions_write_what_word_writes():
    blocks = {re.search(r'paraId="([0-9A-F]{8})"', b).group(1): b for b in OBSERVED["join"]
              if 'paraId="3' in b}
    document = _tracked_joins()
    jtwo = document.paragraph("p:30000001")._element
    assert jtwo.find(".//{*}pPrChange") is not None and "<w:pPrChange" in blocks["30000001"]
    ktwo = document.paragraph("p:30000003")._element
    assert ktwo.find(".//{*}pPrChange") is None and "pPrChange" not in blocks["30000003"]


@pytest.mark.parametrize("built", ["ours", "file"])
def test_the_final_view_names_a_joined_paragraph_by_the_last_one(built):
    document = _tracked_joins() if built == "ours" else Document.open(tracking_probe.join_document())
    state = document.state()
    joined = [(b["id"], b["joins"], b["effective"].get("alignment") or "left") for b in state["blocks"]
              if b.get("joins")]
    expected = [("p:30000005", ["p:30000004"], "right"), ("p:30000007", ["p:30000006"], "right")]
    if built == "ours":
        expected = [("p:30000001", ["p:30000000"], "center"), ("p:30000003", ["p:30000002"], "right")] + expected
        # As Word's Accept All keeps them: the same ids, the same alignment.
        assert [(i, a) for i, _, a in joined] == _accepted_by_word()
    assert joined == expected
    markdown = document.to_markdown()
    assert "<!-- p:30000005 joins p:30000004 -->\nLone first paragraph.Ltwo second paragraph." in markdown
    assert "Mone" not in markdown and "<!-- p:30000007 joins p:30000006 -->\nMtwo next paragraph." in markdown


def test_a_joined_paragraph_is_found_by_any_of_its_ids():
    document = _tracked_joins()
    for identifier in ("p:30000000", "p:30000001"):
        assert document.to_markdown(identifier).startswith("<!-- p:30000001 joins p:30000000 -->")
        assert [b["id"] for b in document.state(identifier)["blocks"]] == ["p:30000001"]


def test_the_original_view_names_a_join_by_the_last_one_too():
    # Word's Reject All of a typed paragraph keeps the last paragraph's element (E3: "After").
    document = Document.new()
    document.insert_markdown("First.\n\nSecond.", at="end")
    first, second = (p.id for p in document.paragraphs())
    with document.tracking(author="Reviewer"):
        created = document.paragraph(first).insert_after("Typed.")
    joined = next(b for b in document.state(view="original")["blocks"] if b.get("joins"))
    assert joined["id"] == second and joined["joins"] == [created.id] and joined["t"] == "Second."


def test_a_joined_paragraph_is_placed_by_the_layout():
    state = _tracked_joins().state(layout=True)
    joined = next(b for b in state["blocks"] if b["id"] == "p:30000001")
    assert joined["placements"][0]["page"] == 1 and "unknown" not in joined


def test_blocks_flag_what_is_wholly_a_pending_deletion():
    document = _tracked_joins()
    flagged = {b["id"] for b in document.blocks() if b.get("deleted")}
    # Kone (deleted here) and Mone (deleted in the file); a mark alone deleted is not the block.
    assert flagged == {"p:30000002", "p:30000006"}
    # Absent when false: the dict keeps its shape.
    assert all("deleted" not in b for b in document.blocks() if b["id"] not in flagged)
    document.accept_all()
    assert not any(b.get("deleted") for b in document.blocks())


def test_a_paragraph_with_text_inserted_is_not_wholly_deleted():
    document = _tracked_joins()
    with document.tracking(author="Other"):
        document.anchor("Ltwo").insert_before("Kept ")
    lone = document.paragraph("p:30000004")
    with document.tracking(author="Other"):
        lone.range().insert_after(" more")
    assert not any(b.get("deleted") for b in document.blocks() if b["id"] == "p:30000004")


def test_a_table_whose_rows_are_all_deleted_is_flagged():
    document = Document.new()
    document.insert_markdown("Before.\n\n| a | b |\n|---|---|\n| 1 | 2 |\n\n| c | d |\n|---|---|\n| 3 | 4 |\n\nAfter.",
                             at="end")
    first, second = document.tables()
    with document.tracking(author="Reviewer"):
        document.delete_block(first.id)
        document.delete_row(second.id, 1)
    flags = {b["id"]: b.get("deleted", False) for b in document.blocks() if b["kind"] == "table"}
    assert flags == {first.id: True, second.id: False}
