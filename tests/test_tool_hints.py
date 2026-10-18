"""Round 9 (ooxml-edit 0.47.0): a name that is no Word tool names the tool and action that
do what it says (N5), with the format's own table (``TOOL_HINTS``) for what ooxml-edit's
reading of names misses; and a batch op's ``comment`` (N4).  Made-up calls."""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.hints import tool_hints

from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.tools.format import TOOL_HINTS

from test_tools_word import AGREEMENT, CLOCK_AT


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _session(box, roles=None):
    session = box.session(clock=lambda: CLOCK_AT, **({"roles": roles} if roles else {}))
    session.open(AGREEMENT.read_bytes(), AGREEMENT.name)
    return session


@pytest.mark.parametrize("name, meant", [
    # read from the tools' names and actions
    ("word_add_comment", "word_comments with action add"),
    ("word_accept_change", "word_changes with action accept"),
    ("word_add_footnote", "word_notes with action insert"),
    ("word_insert_toc", "word_fields with action insert_toc"),
    ("word_add_row", "word_edit_table with action insert_row"),
    ("word_insert_section_break", "word_sections with action insert_break"),
    ("word_delete_paragraph", "word_delete"),
    # the format's own table
    ("word_insert_image", "word_drawings with action insert_picture"),
    ("word_insert_link", "word_links with action add_hyperlink"),   # not a header's link
    ("word_insert_table", "word_insert_markdown"),
    ("word_insert_page_break", "word_format"),                      # page_break_before
    ("word_set_margins", "word_sections with action set"),
    ("word_add_checkboxes", "word_controls with action insert"),
    ("word_track_changes", "word_set_tracking"),
])
def test_what_a_name_means(box, name, meant):
    assert [str(hint) for hint in tool_hints(name, box.tools, box.tool_hints)] == [meant]


def test_every_entry_names_a_tool_and_action_the_format_has(box):
    for name, target in TOOL_HINTS.items():
        tool, _, action = target.partition(" ")
        assert tool in box.tools, name
        assert not action or action in box.tools[tool].actions, name
        assert name not in box.tools


def test_a_batch_names_the_tool_and_action_for_an_unknown_op(box):
    session = _session(box)
    result = box.dispatch(session, "batch", {"atomic": True, "ops": [
        {"tool": "set_properties", "arguments": {"doc": "d1", "title": "Draft"},
         "note": "title first"},
        {"tool": "word_add_comment", "arguments": {"doc": "d1"}}]})
    assert result.error.message == ("ops[1]: no tool 'word_add_comment': use word_comments with "
                                    "action add")
    assert result.error.valid_options[0] == "word_comments"


def test_an_editor_is_told_the_role_accepting_needs(box):
    session = _session(box, ["edit"])
    result = box.dispatch(session, "word_accept_all_changes", {"doc": "d1"})
    assert result.error.message == ("no tool 'word_accept_all_changes': word_changes with action "
                                    "accept does that, and needs the accept role; this session "
                                    "has edit")
    assert "word_changes" not in result.error.valid_options


def test_an_action_of_another_tool(box):
    session = _session(box)
    result = box.dispatch(session, "word_comments", {"doc": "d1", "action": "accept"})
    assert result.error.message.endswith("; use word_changes with action accept")


def test_a_batch_op_comment_runs_nothing_and_is_logged(box):
    session = _session(box)
    result = box.dispatch(session, "batch", {"ops": [
        {"tool": "set_properties", "arguments": {"doc": "d1", "title": "Draft"},
         "comment": "a working title"}]})
    assert result.ok, result.error
    assert session.log[-1].comments == [{"op": 0, "comment": "a working title"}]
