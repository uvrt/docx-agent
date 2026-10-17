"""The Word tools by role (ooxml-edit 0.45.0, round 7, item 9): what each role set is sent,
what it may call, and that all roles -- the default -- is the toolbox as without roles.

Offline numbers (characters / 3.5, as test_tools_schema.py's guards; the online count is in
test_tools_online.py).  The default is unchanged: core 2,570, every definition 12,314.  A
reviewer, ``["review"]``: the core 2,024 (about 2,996 counted by the 1.48 proxy; nine
tools: save_document, describe, find_text, render, check, open_document, batch, word_read,
and word_comments, loaded in the core of a set without edit), every definition 6,805 (21
tools: the mixed tools with only their list action -- describe too for word_styles -- and
edit_chart with only read).  ``["review", "resolve"]``: 2,032 and 6,813.  ``["edit"]``:
core 1,418, every definition 10,980.  Claude's strict set for a reviewer is word_comments,
word_changes and word_notes; the default's and a builder's (``["review", "edit"]``) are
unchanged.
"""

from __future__ import annotations

import pytest

from ooxml_edit.tools import ROLES, Toolbox, adapters
from ooxml_edit.tools.adapters import anthropic_problems, openai_problems

from docx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tools_schema import COUNTED_PER_ESTIMATED, CORE_COUNTED_BUDGET, estimate
from test_tools_word import AGREEMENT, CLOCK_AT

#: A reviewer's core (``["review", "resolve"]``, the larger), estimated: a regression guard
#: as ALL_GUARD is.  Measured 2,032.
REVIEW_CORE_GUARD = 2035
#: Every definition a reviewer is sent, estimated.  Measured 6,813.
REVIEW_ALL_GUARD = 6815

REVIEW_CORE = {"save_document", "describe", "find_text", "render", "check", "open_document",
               "batch", "word_read", "word_comments"}
LISTING = {"word_changes", "word_sections", "word_fields", "word_notes", "word_links",
           "word_drawings", "word_controls"}


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _session(box, roles):
    session = box.session(clock=lambda: CLOCK_AT, roles=roles)
    session.open(AGREEMENT.read_bytes(), AGREEMENT.name)
    return session


def _sent(box, roles):
    return {d["name"]: d for d in box.definitions("anthropic", roles=roles, defer=False)}


def _enum(definition):
    return definition["input_schema"]["properties"]["action"]["enum"]


@pytest.mark.parametrize("provider", ["anthropic", "openai-responses", "openai-chat"])
def test_all_roles_sends_what_the_toolbox_sent_without_roles(box, provider):
    everything = list(box.tools.values())
    for groups in (None, "core"):
        defer = groups is None and provider in ("anthropic", "openai-responses")
        options = ({"strict_first": box.strict_first, "defer": defer}
                   if provider == "anthropic" else
                   {"groups": box.groups.values(), "namespaces": defer, "defer": defer}
                   if provider == "openai-responses" else {})
        before = adapters.definitions_for(
            provider, everything if groups is None else [t for t in everything if t.core],
            **options)
        assert box.definitions(provider, groups=groups) == before
        assert box.definitions(provider, groups=groups, roles=ROLES) == before
    assert "Role:" not in box.system_prompt()


def test_a_reviewer_is_sent_the_reading_tools_the_list_actions_and_comments(box):
    sent = _sent(box, ["review"])
    core = {d["name"] for d in box.definitions("anthropic", groups="core", roles=["review"])}
    assert core == REVIEW_CORE
    assert set(sent) == REVIEW_CORE | LISTING | {"word_inspect", "word_styles", "edit_chart",
                                                 "close_document", "read_blob"}
    for name in LISTING:
        assert _enum(sent[name]) == ["list"], name
    assert _enum(sent["word_styles"]) == ["list", "describe"]
    assert _enum(sent["edit_chart"]) == ["read"]
    assert _enum(sent["word_comments"]) == ["list", "add", "reply"]
    assert {n for n, d in sent.items() if d["strict"]} == {"word_comments", "word_changes",
                                                           "word_notes"}
    for edit_only in ("word_set_text", "word_set_tracking", "word_template", "undo",
                      "new_document", "replace_text", "word_format"):
        assert edit_only not in sent
    assert _enum(_sent(box, ["review", "resolve"])["word_comments"]) == [
        "list", "add", "reply", "resolve", "reopen", "delete"]


def test_a_builder_without_resolve_keeps_the_default_tools_and_strict_set(box):
    default, builder = _sent(box, None), _sent(box, ["review", "edit"])
    assert set(builder) == set(default)
    assert {n for n, d in builder.items() if d["strict"]} == \
        {n for n, d in default.items() if d["strict"]}
    assert _enum(builder["word_comments"]) == ["list", "add", "reply", "edit"]


@pytest.mark.parametrize("roles", [["review"], ["review", "resolve"], ["resolve"], ["edit"],
                                   ["review", "edit"]], ids="+".join)
def test_each_role_set_s_definitions_meet_each_provider_s_rules(box, roles):
    for groups in ("core", None):
        assert anthropic_problems(box.definitions("anthropic", groups=groups,
                                                  roles=roles)) == []
        assert openai_problems(box.definitions("openai-chat", groups=groups, roles=roles),
                               chat=True) == []
    assert anthropic_problems(box.definitions("anthropic", roles=roles)) == []
    assert openai_problems(box.definitions("openai-responses", roles=roles)) == []


def test_a_reviewer_s_definitions_are_within_their_guards(box):
    for roles in (["review"], ["review", "resolve"]):
        core = estimate(box.definitions("anthropic", groups="core", roles=roles))
        everything = estimate(box.definitions("anthropic", defer=False, roles=roles))
        assert core * COUNTED_PER_ESTIMATED <= CORE_COUNTED_BUDGET, (roles, core)
        assert core <= REVIEW_CORE_GUARD, (roles, core)
        assert everything <= REVIEW_ALL_GUARD, (roles, everything)


def test_a_reviewer_comments_and_lists_but_does_not_edit(box):
    session = _session(box, ["review"])
    run = lambda tool_, **arguments: box.dispatch(session, tool_, arguments)  # noqa: E731
    assert run("describe", doc="d1").ok and run("word_read", doc="d1").ok
    assert run("word_changes", doc="d1", action="list").ok
    assert run("word_fields", doc="d1", action="list").ok
    assert run("word_styles", doc="d1", action="list").ok
    added = run("word_comments", doc="d1", action="add", author="Rev",
                items=[{"find": "99.5%", "text": "Per month?", "ref": "q1"}])
    assert added.ok
    assert run("word_comments", doc="d1", action="reply", author="Rev",
               items=[{"comment": "$q1", "text": "Or per year?"}]).ok
    saved = run("save_document", doc="d1", name="reviewed.docx", format="docx")
    assert saved.ok and [o.name for o in session.take_outputs()] == ["reviewed.docx"]

    text = run("word_set_text", doc="d1", target="p:1", text="x")
    assert text.error.code == "refused" and text.error.details == {"role": "edit",
                                                                   "roles": ["review"]}
    assert text.error.message.startswith("word_set_text needs the edit role")
    accept = run("word_changes", doc="d1", action="accept", all=True)
    assert accept.error.code == "refused" and accept.error.valid_options == ["list"]
    for tool_, arguments in (("word_set_tracking", {"on": False}), ("undo", {}),
                             ("word_fields", {"action": "update"})):
        assert run(tool_, doc="d1", **arguments).error.code == "refused", tool_
    resolve = run("word_comments", doc="d1", action="resolve", items=[{"comment": "$q1"}])
    assert resolve.error.code == "refused" and "resolve role" in resolve.error.message
    closing = run("word_comments", doc="d1", action="reply", author="Rev",
                  items=[{"comment": "$q1", "text": "Done.", "resolve": True}])
    assert closing.error.code == "refused"
    listed = run("word_comments", doc="d1", action="list", open_only=True)
    assert added.created[0] in [c["id"] for c in listed.data]


def test_a_second_pass_reviewer_resolves_comments_but_does_not_edit(box):
    session = _session(box, ["review", "resolve"])
    run = lambda tool_, **arguments: box.dispatch(session, tool_, arguments)  # noqa: E731
    added = run("word_comments", doc="d1", action="add", author="Rev",
                items=[{"find": "99.5%", "text": "Per month?", "ref": "q1"}])
    assert added.ok
    assert run("word_comments", doc="d1", action="reply", author="Rev",
               items=[{"comment": "$q1", "text": "Monthly.", "resolve": True}]).ok
    assert run("word_comments", doc="d1", action="reopen", items=[{"comment": "$q1"}]).ok
    assert run("word_comments", doc="d1", action="delete", items=[{"comment": "$q1"}]).ok
    edit = run("word_comments", doc="d1", action="edit", items=[{"comment": "$q1",
                                                                 "text": "x"}])
    assert edit.error.code == "refused" and "word_comments edit needs the edit role" in \
        edit.error.message


def test_a_review_batch_applies_its_comments_and_refuses_the_edits(box):
    session = _session(box, ["review"])
    result = box.dispatch(session, "batch", {"ops": [
        {"tool": "word_comments", "arguments": {
            "doc": "d1", "action": "add", "author": "Rev",
            "items": [{"find": "99.5%", "text": "Per month?"}]}},
        {"tool": "replace_text", "arguments": {"doc": "d1", "find": "99.5%",
                                               "replace": "99.9%", "expect": "one"}}]})
    assert result.ok and result.data["failed"] == [1]
    assert "replace_text needs the edit role" in result.data["ops"][1]["error"]["message"]
    assert "99.5%" in box.dispatch(session, "word_read", {"doc": "d1"}).to_text()
