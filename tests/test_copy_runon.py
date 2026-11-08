"""Round 43: numbered lists that run on through the copies, as the caller states it.

One plain numbered list may count on through every copy of a sample in a document on
purpose (1-2 in the first, 3-4 in the next, ...).  The sample alone cannot show it, and in
Word's XML a copy that counts on by design looks the same as one that should have begun
again, so it is never inferred: ``check``'s ``copies`` item says ``"lists": "run_on"``.
Then the run's first copy numbers as its sample shows and each later one goes on from the
copy before it, in the same list: a restart, or a separate list, is the difference, fixed by
``word_lists`` ``continue`` joining the running list -- never a restart.  Without the field
nothing changes.  Every document here is made up."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent.edit import copy_settings  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

sys.path.insert(0, str(Path(__file__).parent))

from test_copy_outline import PART, STEPS, _blocks, _differences, _made, _numbers, _ok, _template  # noqa: E402

pytestmark = pytest.mark.any_faces


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _check(box, session, report, sample, span, made, lists=None, only=None) -> dict:
    stated = {"lists": lists} if lists is not None else {}
    pairs = [{"blocks": f"{m[0]}..{m[-1]}", "reference": {"doc": sample, "block": span}, **stated}
             for k, m in enumerate(made) if only is None or k in only]
    return _ok(box, session, "check", doc=report, include=[], copies=pairs)["copies"]


def _numbering(found: dict, k: int) -> list[tuple]:
    return [(d["value"], d["reference"], d["paragraphs"], d["fix"]) for d in found["items"][k]["differences"]
            if d["property"] == "list_numbering"]


def test_copies_whose_list_runs_on_follow_where_it_is_stated(box):
    """Copied three times, the part's list counts on (1, 3, 5): stated to run on, every copy
    follows; the later copies say which copy they go on from."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART + STEPS, 3)
    assert _numbers(session, report, [m[3] for m in made]) == [1, 3, 5]
    found = _check(box, session, report, sample, span, made, "run_on")
    assert found["follow"] == 3, _differences(found)
    assert "list_fix" not in found and "fix" not in found and "note" not in found
    assert [item.get("runs_on_from") for item in found["items"]] == [
        None, f"{made[0][0]}..{made[0][-1]}", f"{made[1][0]}..{made[1][-1]}"]


def test_without_the_field_the_same_copies_differ_as_before(box):
    """The same documents, nothing stated (or ``sample``): 0.1.69's rule -- each copy's list
    starts as the sample's, a restart offered -- and the reply as before, no new field."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART + STEPS, 3)
    default = _check(box, session, report, sample, span, made)
    assert [item["follows"] for item in default["items"]] == [True, False, False]
    assert _numbering(default, 1) == [("starts at 3", "starts at 1", [made[1][3]], {
        "tool": "word_lists", "calls": [{"targets": [made[1][3]], "action": "restart", "start": 1}]})]
    assert all("runs_on_from" not in item for item in default["items"])
    stated = _check(box, session, report, sample, span, made, "sample")
    assert stated == {**default, "kept": stated["kept"],
                      "items": [{**item, "cached": stated["items"][k]["cached"]}
                                for k, item in enumerate(default["items"])]}


def test_a_copy_that_restarts_breaks_the_run_and_its_fix_joins_it(box):
    """Stated to run on, the second copy's list restarts at 1: told (``starts at 1`` against
    ``runs on``), fixed by continuing the first copy's list -- never a restart.  The third
    copy, whose list counts on from the second's, follows.  The fix applied, the check is
    clean and the list counts 1, 3, 5 again."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART + STEPS, 3)
    _ok(box, session, "word_lists", doc=report, targets=[made[1][3]], action="restart", start=1)
    assert _numbers(session, report, [m[3] for m in made]) == [1, 1, 3]
    found = _check(box, session, report, sample, span, made, "run_on")
    assert [item["follows"] for item in found["items"]] == [True, False, True]
    join = {"targets": [made[1][3]], "action": "continue", "join": made[0][4]}
    assert _numbering(found, 1) == [("starts at 1", "runs on", [made[1][3]],
                                     {"tool": "word_lists", "calls": [join]})]
    assert found["list_fix"] == {"tool": "word_lists", "calls": [join]}
    for call in found["list_fix"]["calls"]:
        _ok(box, session, "word_lists", doc=report, **call)
    assert _numbers(session, report, [m[3] for m in made]) == [1, 3, 5]
    assert _check(box, session, report, sample, span, made, "run_on")["follow"] == 3


def test_a_copy_with_a_separate_list_is_joined_to_the_running_one(box):
    """The second copy's steps made a list of their own (a new list of Word's): a separate
    list breaks the run too; the fix names the list to join, and joins it."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART + STEPS, 3)
    steps = made[1][3:5]
    _ok(box, session, "word_lists", doc=report, targets=steps, action="remove", force=True)
    _ok(box, session, "word_lists", doc=report, targets=steps, action="add", kind="number")
    document = session.entry(report).document
    assert document.list_of(made[1][3]).num_id != document.list_of(made[0][3]).num_id
    found = _check(box, session, report, sample, span, made, "run_on")
    numbering = _numbering(found, 1)
    assert [(value, reference) for value, reference, _, _ in numbering] == [("starts at 1", "runs on")]
    assert numbering[0][3]["calls"] == [{"targets": [made[1][3]], "action": "continue", "join": made[0][4]}]
    for call in found["list_fix"]["calls"]:
        _ok(box, session, "word_lists", doc=report, **call)
    assert _numbers(session, report, [m[3] for m in made])[:2] == [1, 3]
    assert not _numbering(_check(box, session, report, sample, span, made, "run_on"), 1)


def test_the_run_s_first_copy_numbers_as_its_sample(box):
    """The first copy of a run begins as the sample's list begins: one that counts on from a
    list before it differs, and there a restart is offered (it begins the run)."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART + STEPS, 3)
    _ok(box, session, "word_insert_markdown", doc=report, markdown="1. A list item before the copies",
        at=f"before:{made[0][0]}")
    _ok(box, session, "word_format", doc=report, targets=made[0][3:5], like=_list_item_before(session, report, made))
    found = _check(box, session, report, sample, span, made, "run_on")
    numbering = _numbering(found, 0)
    assert [(value, reference) for value, reference, _, _ in numbering] == [("starts at 2", "starts at 1")]
    assert numbering[0][3]["calls"] == [{"targets": [made[0][3]], "action": "restart", "start": 1}]


def _list_item_before(session, report, made) -> str:
    blocks = _blocks(session, report)
    return blocks[blocks.index(made[0][0]) - 1]


def test_numbered_headings_are_unaffected_either_way(box):
    """Headings numbered by the template's outline compare by level and label, stated or not:
    the copies follow, and no restart nor continue is offered on them."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART, 3, sections=2)
    for lists in (None, "run_on"):
        found = _check(box, session, report, sample, span, made, lists)
        assert found["follow"] == 6, (lists, _differences(found))
        assert "list_fix" not in found
    _ok(box, session, "word_lists", doc=report, targets=[made[2][0]], action="level", level=1)
    found = _check(box, session, report, sample, span, made, "run_on")
    level = {d["property"]: d for d in found["items"][2]["differences"]}
    assert level["list_level"]["fix"]["calls"][0]["action"] == "level" and "list_numbering" not in level
    assert all(call["action"] == "level" for call in found["list_fix"]["calls"])


def test_a_run_of_one_copy_is_said(box):
    """Stated on one copy alone: it is the run's first (numbers as its sample) -- a note says
    so and how to name the whole run."""
    session = box.session()
    report, sample, span, made = _made(box, session, PART + STEPS, 3)
    found = _check(box, session, report, sample, span, made, "run_on", only={2})
    assert "copies[0]" in found["note"] and "whole run" in found["note"]
    assert _numbering(found, 0)[0][:2] == ("starts at 5", "starts at 1")


def test_a_wrong_value_is_refused_with_the_valid_ones_and_a_near_one_read(box):
    session = box.session()
    report, sample, span, made = _made(box, session, PART + STEPS, 2)
    pairs = [{"blocks": f"{m[0]}..{m[-1]}", "reference": {"doc": sample, "block": span}, "lists": "restart"}
             for m in made]
    refused = box.dispatch(session, "check", {"doc": report, "include": [], "copies": pairs})
    assert refused.error.code == "invalid_arguments" and refused.error.field == "copies[0].lists"
    assert refused.error.valid_options == ["sample", "run_on"]
    for pair in pairs:
        pair["lists"] = "runs_on"
    read = box.dispatch(session, "check", {"doc": report, "include": [], "copies": pairs})
    assert read.ok and read.data["copies"]["follow"] == 2
    assert any("read as 'run_on'" in warning for warning in read.warnings)


def test_the_handler_reads_a_placeholder_as_the_default_and_refuses_another_value():
    """Rule 1 and 2 where the handler reads the item itself."""
    from ooxml_edit.tools import ToolError

    from docx_agent.tools.scoped import _lists

    assert _lists({"lists": ""}, 0) == _lists({}, 0) == "sample"
    assert _lists({"lists": "run_on"}, 1) == "run_on"
    with pytest.raises(ToolError) as refused:
        _lists({"lists": "loop"}, 3)
    assert refused.value.field == "copies[3].lists" and refused.value.valid_options == ["sample", "run_on"]


def test_word_lists_continue_joins_the_list_named(box):
    """``join``: the list to go on from, where it is not the nearest one like it (rule 1: a
    placeholder is the default; given to another action, it is left out with a warning)."""
    session = box.session()
    doc = _new(box, session)
    _ok(box, session, "word_insert_markdown", doc=doc, markdown="1. one\n2. two\n\nA line.\n\n1. alpha\n2. beta"
        "\n\nAnother line.\n\n1. last", at="end")
    blocks = _blocks(session, doc)
    first, second, last = blocks[0], blocks[3], blocks[-1]
    document = session.entry(doc).document
    lists = [document.list_of(b).num_id for b in (first, second, last)]
    assert len(set(lists)) == 3
    result = box.dispatch(session, "word_lists", {"doc": doc, "targets": [last], "action": "continue",
                                                  "join": first})
    assert result.ok and document.list_of(last).num_id == lists[0]
    assert _numbers(session, doc, [last]) == [3]
    warned = box.dispatch(session, "word_lists", {"doc": doc, "targets": [last], "action": "level", "level": 0,
                                                  "join": first})
    assert any("join is continue's" in warning for warning in warned.warnings)


def _new(box, session) -> str:
    return session.open(_template(), "lists.docx")


def test_run_is_never_a_restart_inside_a_running_list():
    """The rule itself: a later copy's leading item goes on from the copy before it (same
    list, its count carried); a restart is held back only inside a running list."""
    lead = {copy_settings.LEAD: True, copy_settings.COUNT: "own", "list_numbering": "starts at 1"}
    run = copy_settings.Run([(0, "L", "decimal %1.", "p:A")])
    joins: dict = {}
    mine = {"list_level": 0, "list_format": "decimal %1.", copy_settings.LIST_KEY: "L",
            copy_settings.COUNT: "runs_on", "list_numbering": "starts at 3"}
    assert run.compared(mine, lead, "p:B", joins) == ("runs on", "runs on") and joins == {"p:B": "p:A"}
    restarted = {**mine, copy_settings.COUNT: "own", "list_numbering": "starts at 1"}
    assert run.compared(restarted, lead, "p:B", joins) == ("starts at 1", "runs on")
    other = {**mine, copy_settings.LIST_KEY: "M"}
    assert run.compared(other, lead, "p:B", joins) == ("starts at 3", "runs on")
    first = copy_settings.Run(None)
    assert first.compared(restarted, lead, "p:B", {}) == ("starts at 1", "starts at 1")
    assert copy_settings.settings_fix(None, "list_numbering", "starts at 3", "starts at 1", ["p:B"], "p:S",
                                      {}, {}, restart=False) is None
