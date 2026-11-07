"""``word_format``'s ``find`` (round 40): text found by the call, as ``replace_text`` finds it
-- in a block or span (``within``), once (``expect`` one, the default), the n-th match
(``occurrence``, from 1) or every one (``expect`` all) -- formatted as a range target is, so
a text and its formatting go in one call.  Every item is found before anything changes: a
text that is not there, or matches several, refuses the call with the matches listed.
Made-up text only."""

from __future__ import annotations

import datetime as dt

import pytest

from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

CLOCK = lambda: dt.datetime(2026, 11, 7, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def made(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "notes.docx")
    result = toolbox.dispatch(session, "word_insert_markdown", {
        "doc": "d1", "at": "end",
        "markdown": "Plan the river walk and the river ford.\n\nThe ford is shallow.\n\nRiver notes end here."})
    assert result.ok, result.to_json()
    return session, result.data["blocks"]


def run(toolbox, session, name, **arguments):
    return toolbox.dispatch(session, name, arguments)


def text_of(session) -> str:
    return session.entry("d1").document.to_markdown(ids=False)


def test_one_match_in_a_block_is_formatted_alone(toolbox, made):
    session, (first, second, third) = made
    result = run(toolbox, session, "word_format", doc="d1", find=[{"text": "ford", "within": second}], bold=True)
    assert result.ok, result.to_json()
    assert result.data == {"ranges": [f"{second}@4:8"]}
    assert "The **ford** is shallow." in text_of(session)
    assert "river ford." in text_of(session)                     # the first paragraph's untouched


def test_several_texts_and_a_character_style_in_one_call(toolbox, made):
    session, (first, second, third) = made
    result = run(toolbox, session, "word_format", doc="d1", character_style="Strong",
                 find=[{"text": "walk", "within": first}, {"text": "shallow"}])
    assert result.ok, result.to_json()
    assert result.summary == "Formatted 2 target(s)"
    assert "river **walk**" in text_of(session) and "**shallow**" in text_of(session)


def test_several_matches_are_refused_listing_each_with_its_occurrence(toolbox, made):
    session, (first, second, third) = made
    before = text_of(session)
    result = run(toolbox, session, "word_format", doc="d1", italic=True,
                 find=[{"text": "shallow"}, {"text": "river", "within": first}])
    assert not result.ok and result.error.code == "ambiguous"
    assert result.error.field == "find[1].occurrence"
    assert "2 places match 'river'" in result.error.message and "occurrence (1-2" in result.error.message
    assert result.error.valid_options == [f"occurrence 1: {first}@9:14 'Plan the [river] walk and the river '",
                                          f"occurrence 2: {first}@28:33 ' river walk and the [river] ford.'"]
    assert text_of(session) == before                             # the first item was not formatted


def test_occurrence_picks_one_and_expect_all_takes_each(toolbox, made):
    session, (first, second, third) = made
    result = run(toolbox, session, "word_format", doc="d1", underline=True,
                 find={"text": "river", "within": first, "occurrence": 2})
    assert result.ok and result.data["ranges"] == [f"{first}@28:33"]
    result = run(toolbox, session, "word_format", doc="d1", bold=True,
                 find=[{"text": "ford", "expect": "all"}])
    assert result.ok and len(result.data["ranges"]) == 2
    result = run(toolbox, session, "word_format", doc="d1", bold=True,
                 find=[{"text": "river", "within": first, "occurrence": 3}])
    assert result.error.code == "not_found" and "occurs 2 time(s)" in result.error.message
    assert len(result.error.valid_options) == 2
    result = run(toolbox, session, "word_format", doc="d1", bold=True,
                 find=[{"text": "river", "occurrence": 1, "expect": "all"}])
    assert result.error.code == "invalid_arguments"


def test_a_text_not_there_names_what_is(toolbox, made):
    session, (first, second, third) = made
    result = run(toolbox, session, "word_format", doc="d1", bold=True, find=[{"text": "the  Ford"}])
    assert result.error.code == "not_found" and result.error.field == "find[0].text"
    assert "written otherwise" in result.error.message
    assert result.error.valid_options == [f"{second}@0:8 'The ford'"]
    result = run(toolbox, session, "word_format", doc="d1", bold=True, find=[{"text": "bridge", "within": second}])
    assert result.error.code == "not_found" and "it holds 'The ford is shallow.'" in result.error.message


def test_targets_and_find_together_and_replace_texts_range_name(toolbox, made):
    session, (first, second, third) = made
    result = run(toolbox, session, "word_format", doc="d1", targets=third, italic=True,
                 find=[{"text": "shallow", "range": second}])
    assert result.ok, result.to_json()
    assert "find[0].range read as within" in result.warnings
    assert result.summary == "Formatted 2 target(s)"
    assert "*River notes end here.*" in text_of(session) and "*shallow*" in text_of(session)


def test_nothing_to_format(toolbox, made):
    session, _ = made
    result = run(toolbox, session, "word_format", doc="d1", bold=True)
    assert result.error.code == "invalid_arguments" and result.error.valid_options == ["targets", "find"]
    result = run(toolbox, session, "word_format", doc="d1", bold=True, targets=[], find=[])
    assert result.ok and "nothing was formatted" in result.warnings[0]


def test_found_text_in_a_batch_is_one_step_with_the_text_written(toolbox, made):
    session, (first, second, third) = made
    ops = [{"tool": "replace_text", "arguments": {"doc": "d1", "find": "shallow", "replace": "deep at noon",
                                                  "expect": "one"}},
           {"tool": "word_format", "arguments": {"doc": "d1", "find": [{"text": "deep at noon", "within": second}],
                                                 "character_style": "Strong"}}]
    result = run(toolbox, session, "batch", ops=ops, atomic=True)
    assert result.ok, result.to_json()
    assert "The ford is **deep at noon**." in text_of(session)
    assert run(toolbox, session, "undo", doc="d1").ok
    assert "The ford is shallow." in text_of(session)
    ops[1]["arguments"]["find"] = [{"text": "missing words"}]
    result = run(toolbox, session, "batch", ops=ops, atomic=True)
    assert not result.ok and "missing words" in result.error.message
    assert "The ford is shallow." in text_of(session)


def test_found_text_formats_as_a_tracked_change(toolbox, made):
    session, (first, second, third) = made
    assert run(toolbox, session, "word_set_tracking", doc="d1", on=True, author="Reviewer").ok
    result = run(toolbox, session, "word_format", doc="d1", bold=True, find=[{"text": "shallow"}])
    assert result.ok and result.checks["validate"]["new"] == []
    assert session.entry("d1").document.revisions()
