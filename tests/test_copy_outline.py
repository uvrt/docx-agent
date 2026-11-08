"""Round 43: where a copy's list item stands in its count is compared with its sample's only
where the sample's own count begins inside the sample (a numbered list of its own, which
each copy then starts alike).  Headings numbered by the template's outline, and a list that
runs on from before the sample, compare by level and label only: their numbers are their
place in the document.  Every document here is made up."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent.edit import copy_settings  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

sys.path.insert(0, str(Path(__file__).parent))

import toc_number_cases as cases  # noqa: E402

pytestmark = pytest.mark.any_faces

PART = ("### A titled part\n\n| Name | Value |\n|---|---|\n| one | two |\n\nA body line under the table.")
STEPS = "\n\n1. Step one of the part\n2. Step two of the part\n\nA closing line of the part."


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _template() -> bytes:
    """Headings 1-3 numbered by their styles through one multilevel list tied to them
    (1, 1.1, 1.1.1), as a template numbers its outline."""
    return cases.package("", cases.styles(), cases.numbering(start=1))


def _ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, arguments, result.error and result.error.to_json())
    return result.data


def _blocks(session, doc: str) -> list[str]:
    return [b["id"] for b in session.entry(doc).document.blocks()]


def _made(box, session, sample_markdown: str, count: int, *, sections: int = 1, lead: str = "", prepare=None):
    """The sample (its own document, ``lead`` before its part; ``prepare(sample, blocks)``
    run on it before it is copied) and a report of numbered sections, ``count`` copies of
    the part under each: ``(report, sample, span, copies)``."""
    sample = session.open(_template(), "sample.docx")
    if lead:
        _ok(box, session, "word_insert_markdown", doc=sample, markdown=lead, at="end")
    before = len(_blocks(session, sample))
    _ok(box, session, "word_insert_markdown", doc=sample, markdown=sample_markdown, at="end")
    part = _blocks(session, sample)[before:]
    if prepare is not None:
        prepare(sample, _blocks(session, sample))
    span = f"{part[0]}..{part[-1]}"
    report = session.open(_template(), "report.docx")
    made = []
    for k in range(sections):
        _ok(box, session, "word_insert_markdown", doc=report, markdown=f"# Section {k + 1}\n\nAn opening line.",
            at="end")
        made += [_ok(box, session, "word_copy_from", doc=report, source_doc=sample, range=span,
                     at="end")["blocks"] for _ in range(count)]
    return report, sample, span, made


def _check(box, session, report, sample, span, made) -> dict:
    pairs = [{"blocks": f"{m[0]}..{m[-1]}", "reference": {"doc": sample, "block": span}} for m in made]
    return _ok(box, session, "check", doc=report, include=[], copies=pairs)["copies"]


def _numbers(session, doc: str, blocks: list[str]) -> list:
    """The numbers the blocks' list items show (the module's count)."""
    document = session.entry(doc).document
    items = copy_settings.list_items(document)
    return [items.get(id(document._resolve(b)[1].element), (None,) * 5)[4] for b in blocks]


def _differences(found: dict) -> list[list[dict]]:
    return [item["differences"] for item in found["items"]]


def test_a_part_titled_by_a_numbered_heading_copied_twice_follows(box):
    """The repro: the sample's title is a Heading 3 the template numbers; copied twice under
    a numbered Heading 1, each copy's title takes the next number of the outline.  Not a
    difference, and no restart offered."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART, 2)
    assert _numbers(session, report, [m[0] for m in made]) == [1, 2]
    found = _check(box, session, report, sample, span, made)
    assert found["follow"] == 2, _differences(found)
    assert "list_fix" not in found and "fix" not in found


def test_numbered_headings_compare_by_level_and_label_not_their_number(box):
    """Two titled parts per copy, three copies under each of two numbered sections: each
    copy's headings stand at the sample's outline level with its label, under another
    parent number -- alike."""
    session = box.session()
    markdown = PART + "\n\n### A second titled part\n\nAnother body line."
    report, sample, span, made = _made(box, session, markdown, 3, sections=2)
    assert _numbers(session, report, [m[0] for m in made]) == [1, 3, 5, 1, 3, 5]
    found = _check(box, session, report, sample, span, made)
    assert found["follow"] == 6, _differences(found)
    document = session.entry(report).document
    title = copy_settings.region_settings(document, [document._resolve(made[4][0])[1].element])
    values = next(iter(title.values()))
    assert (values["list_level"], values["list_format"]) == (2, "decimal %1.%2.%3")


def test_a_title_at_another_outline_level_differs_without_a_restart(box):
    """A copy's title a level up: told (its style, or its list level where only that moved),
    and never fixed by a restart."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART, 3)
    _ok(box, session, "word_format", doc=report, targets=[made[1][0]], paragraph_style="heading 2")
    _ok(box, session, "word_lists", doc=report, targets=[made[2][0]], action="level", level=1)
    found = _check(box, session, report, sample, span, made)
    assert [item["follows"] for item in found["items"]] == [True, False, False]
    styled = {d["property"]: d for d in found["items"][1]["differences"]}
    assert "paragraph_style" in styled and "list_numbering" not in styled
    level = {d["property"]: d for d in found["items"][2]["differences"]}
    assert (level["list_level"]["value"], level["list_level"]["reference"]) == (1, 2)
    assert level["list_level"]["fix"]["calls"][0]["action"] == "level"
    assert "list_numbering" not in level
    calls = found.get("list_fix", {}).get("calls", [])
    assert all(call["action"] != "restart" for call in calls)


def test_a_numbered_list_of_the_parts_own_starts_alike_in_each_copy(box):
    """The sample numbers a list of its own from 1: each copy's list starting at 1 follows;
    a copy whose list runs on from the copy before it differs, fixed by a restart -- its
    heading, numbered by the outline, never."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART + STEPS, 3)
    for blocks in made[1:]:
        _ok(box, session, "word_lists", doc=report, targets=[blocks[3]], action="restart", start=1)
    found = _check(box, session, report, sample, span, made)
    assert found["follow"] == 3, _differences(found)
    _ok(box, session, "word_lists", doc=report, targets=[made[2][3]], action="continue")
    found = _check(box, session, report, sample, span, made)
    assert [item["follows"] for item in found["items"]] == [True, True, False]
    numbering = [d for d in found["items"][2]["differences"] if d["property"] == "list_numbering"]
    assert [(d["value"], d["reference"], d["paragraphs"]) for d in numbering] == [
        ("starts at 3", "starts at 1", [made[2][3]])]
    assert found["list_fix"]["calls"] == [{"targets": [made[2][3]], "action": "restart", "start": 1}]
    for call in found["list_fix"]["calls"]:
        _ok(box, session, "word_lists", doc=report, **call)
    assert _check(box, session, report, sample, span, made)["follow"] == 3


def test_a_list_that_runs_on_from_before_the_sample_runs_on_in_the_copies(box):
    """The sample's list goes on from an item before the sample (it runs on by design):
    where each copy's list stands in its count is not compared -- a copy starting afresh or
    running on from the copy before it both follow."""
    session = box.session()
    lead = "1. An item before the part"

    def join(sample: str, blocks: list[str]) -> None:
        # The part's steps join the list begun before the part (they take its look).
        _ok(box, session, "word_format", doc=sample, targets=blocks[4:6], like=blocks[0])

    report, sample, span, made = _made(box, session, PART + STEPS, 3, lead=lead, prepare=join)
    assert _numbers(session, sample, _blocks(session, sample)[4:6]) == [2, 3]
    # The copies' lists: the first begins the report's count, the others run on.
    assert _numbers(session, report, [m[3] for m in made]) == [1, 3, 5]
    _ok(box, session, "word_lists", doc=report, targets=[made[1][3]], action="restart", start=1)
    found = _check(box, session, report, sample, span, made)
    assert found["follow"] == 3, _differences(found)
    assert "list_fix" not in found
    # Inside the copy, the list still goes on as the sample's does: a restart there is told.
    _ok(box, session, "word_lists", doc=report, targets=[made[2][4]], action="restart", start=1)
    found = _check(box, session, report, sample, span, made)
    numbering = [d for d in found["items"][2]["differences"] if d["property"] == "list_numbering"]
    assert [(d["value"], d["reference"]) for d in numbering] == [("starts at 1", "continues")]


def test_an_outline_item_is_never_offered_a_restart():
    """The rule itself: compared only where the sample's count is its own, and the copy's
    item not numbered by the outline."""
    counted = copy_settings._counted
    key = copy_settings.COUNT
    assert counted({key: "runs_on"}, {key: "own"})
    assert counted({key: "own"}, {key: "own"})
    assert not counted({key: "outline"}, {key: "own"})
    assert not counted({key: "own"}, {key: "outline"})
    assert not counted({key: "own"}, {key: "runs_on"})
    assert not counted({key: None}, {key: None})
