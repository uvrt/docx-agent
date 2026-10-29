"""The Word tools by role (ooxml-edit 0.45.0, round 7, item 9): what each role set is sent,
what it may call, and that all roles -- the default -- is the toolbox as without roles.

Offline numbers (characters / 3.5, as test_tools_schema.py's guards; the online count is in
test_tools_online.py).  The default is unchanged: core 2,570, every definition 12,314.  A
reviewer, ``["review"]``: the core 2,350 (about 3,478 counted by the 1.48 proxy; ten
tools: save_document, describe, find_text, render, check, open_document, batch, word_read,
and word_comments and word_changes, loaded in the core of a set without edit), every
definition 6,807 (21 tools: the mixed tools with only their list action -- describe too
for word_styles -- and edit_chart with only read).  ``["review", "resolve"]``: 2,358 and
6,815; an editor-in-chief, ``["review", "accept"]``: 2,356 and 6,812; all three 2,364 and
6,820.  ``["accept"]`` alone: core 814, every definition 1,418 (word_changes accept and
reject, word_comments for its text edits, every role's).  ``["edit"]``: core 1,418, every
definition 10,653 (word_changes is accept's now).  Claude's strict set for a reviewer is
word_comments, word_changes and word_notes; for ``["edit"]`` alone word_changes leaves it
and edit_smartart and set_properties join; the default's and a builder's
(``["review", "edit"]``) are unchanged.
"""

from __future__ import annotations

import pytest

from ooxml_edit.tools import ROLES, Toolbox, adapters
from ooxml_edit.tools.adapters import anthropic_problems, openai_problems

from docx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tools_schema import COUNTED_PER_ESTIMATED, CORE_COUNTED_BUDGET, estimate
from test_tools_word import AGREEMENT, CLOCK_AT

#: A reviewer's core (``["review", "resolve"]``, the larger), estimated: a regression guard
#: as ALL_GUARD is.  Measured 2,364 for ``["review", "resolve", "accept"]`` (2,032 for
#: review+resolve before word_changes joined a reviewer's core with the accept role); round
#: 9, N1 (render's ``sheet``, +64) 2,427; round 12, Q1 (check's ``spacing``, +35) 2,462;
#: round 15, T2 (check's ``contrast``, +37) 2,499; round 16, F11 (save_document's
#: ``clean``, +42) 2,541; round 16, F8 (check's ``layout``, ``min_size`` and ``styles``,
#: +67) 2,608; round 23, V3 (check without a deck's slides, boxes, spacing and contrast,
#: -144) 2,464.
REVIEW_CORE_GUARD = 2465
#: Every definition a reviewer is sent, estimated.  Measured 6,820 (6,813 before); round 9,
#: N1, 6,884; round 12, Q1, 6,919; round 15, T2, 6,955; round 16, F11, 6,997; F8, 7,064; F9
#: (word_design_facts, a reading tool, deferred: +170) 7,234; round 17, P3
#: (word_compare_blocks, a reading tool, deferred) and P4 (word_copy_from's description) 7,492
#: (review, resolve and accept, the largest); P6 (word_drawings' ``color``) 7,519; round 21, T7
#: (word_design_facts' ``declared_styles`` and ``cursor``) 7,560; round 23, V3 (check
#: without a deck's parameters) 7,417.
REVIEW_ALL_GUARD = 7420

REVIEW_CORE = {"save_document", "describe", "find_text", "render", "check", "open_document",
               "batch", "word_read", "word_comments", "word_changes"}
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
                                                 "close_document", "read_blob", "word_design_facts",
                                                 "word_compare_blocks"}
    for name in LISTING:
        assert _enum(sent[name]) == ["list"], name
    assert _enum(sent["word_styles"]) == ["list", "describe"]
    assert _enum(sent["edit_chart"]) == ["read"]
    assert _enum(sent["word_comments"]) == ["list", "add", "reply", "edit"]
    assert {n for n, d in sent.items() if d["strict"]} == {"word_comments", "word_changes",
                                                           "word_notes"}
    for edit_only in ("word_set_text", "word_set_tracking", "word_template", "undo",
                      "new_document", "replace_text", "word_format"):
        assert edit_only not in sent
    assert _enum(_sent(box, ["review", "resolve"])["word_comments"]) == [
        "list", "add", "reply", "resolve", "reopen", "edit", "delete"]
    chief = _sent(box, ["review", "accept"])
    assert _enum(chief["word_changes"]) == ["list", "accept", "reject"]
    assert "word_changes" not in _sent(box, ["edit"])


def test_a_builder_without_resolve_keeps_the_default_tools_and_strict_set(box):
    default, builder = _sent(box, None), _sent(box, ["review", "edit"])
    assert set(builder) == set(default)
    # Its word_changes offers list alone and so has no ids or all (round 17,
    # Tool.action_fields): the strict optional parameters it frees take edit_smartart too.
    assert {n for n, d in builder.items() if d["strict"]} == \
        {n for n, d in default.items() if d["strict"]} | {"edit_smartart"}
    assert "ids" not in builder["word_changes"]["input_schema"]["properties"]
    assert "ids" in default["word_changes"]["input_schema"]["properties"]
    assert _enum(builder["word_comments"]) == ["list", "add", "reply", "edit"]


@pytest.mark.parametrize("roles", [["review"], ["review", "resolve"], ["resolve"], ["edit"],
                                   ["review", "edit"], ["accept"], ["review", "accept"]],
                         ids="+".join)
def test_each_role_set_s_definitions_meet_each_provider_s_rules(box, roles):
    for groups in ("core", None):
        assert anthropic_problems(box.definitions("anthropic", groups=groups,
                                                  roles=roles)) == []
        assert openai_problems(box.definitions("openai-chat", groups=groups, roles=roles),
                               chat=True) == []
    assert anthropic_problems(box.definitions("anthropic", roles=roles)) == []
    assert openai_problems(box.definitions("openai-responses", roles=roles)) == []


def test_a_reviewer_s_definitions_are_within_their_guards(box):
    for roles in (["review"], ["review", "resolve"], ["review", "accept"],
                  ["review", "resolve", "accept"]):
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
    assert run("word_comments", doc="d1", action="edit", items=[{"comment": "$q1",
                                                                 "text": "Per month."}]).ok
    assert run("word_comments", doc="d1", action="delete", items=[{"comment": "$q1"}]).ok
    accept = run("word_changes", doc="d1", action="accept", all=True)
    assert accept.error.code == "refused" and "needs the accept role" in accept.error.message


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


def _tracked(box, roles):
    """A session on the agreement with one tracked change made by a builder first."""
    session = _session(box, None)
    run = lambda tool_, **arguments: box.dispatch(session, tool_, arguments)  # noqa: E731
    assert run("word_set_tracking", doc="d1", on=True, author="Builder").ok
    assert run("replace_text", doc="d1", find="99.5%", replace="99.9%", expect="one").ok
    session.roles = tuple(roles)          # the editor-in-chief takes over the session
    return session, run


def test_an_editor_in_chief_accepts_and_rejects_but_does_not_edit(box):
    session, run = _tracked(box, ["review", "accept"])
    listed = run("word_changes", doc="d1", action="list")
    assert listed.ok and listed.data
    assert run("word_changes", doc="d1", action="accept", all=True).ok
    assert not run("word_changes", doc="d1", action="list").data
    refused = run("replace_text", doc="d1", find="99.9%", replace="99%", expect="one")
    assert refused.error.code == "refused" and "needs the edit role" in refused.error.message
    assert "You may also accept and reject tracked changes." in \
        box.system_prompt(roles=["review", "accept"])


def test_accept_without_review_settles_changes_but_reads_nothing(box):
    session, run = _tracked(box, ["accept"])
    assert run("word_changes", doc="d1", action="reject", all=True).ok
    listed = run("word_changes", doc="d1", action="list")
    assert listed.error.code == "refused" and "needs the review role" in listed.error.message
    assert run("word_read", doc="d1").error.code == "refused"


def test_a_batch_holds_each_op_to_the_accept_role(box):
    session, run = _tracked(box, ["review"])
    result = run("batch", ops=[
        {"tool": "word_changes", "arguments": {"doc": "d1", "action": "list"}},
        {"tool": "word_changes", "arguments": {"doc": "d1", "action": "accept", "all": True}}])
    assert result.ok and result.data["failed"] == [1]
    assert "word_changes accept needs the accept role" in \
        result.data["ops"][1]["error"]["message"]


# -- several agents on one document (ooxml-edit 0.46.0, round 8, M1 and M2) ---------------------


def test_a_builder_a_reviewer_and_an_editor_in_chief_share_one_session(box):
    """One session, every role; each agent a view of it with its own roles and context."""
    session = _session(box, None)
    builder = session.view(roles=["review", "edit"], context={"agent": "builder-1"})
    reviewer = session.view(roles=["review"], context={"agent": "reviewer-1"})
    chief = session.view(roles=["review", "accept"], context={"agent": "chief-1"})
    assert builder.entry("d1") is reviewer.entry("d1") is session.entry("d1")
    assert box.dispatch(builder, "word_set_tracking",
                        {"doc": "d1", "on": True, "author": "Builder"}).ok
    assert box.dispatch(builder, "replace_text", {"doc": "d1", "find": "99.5%",
                                                  "replace": "99.9%", "expect": "one"}).ok
    assert box.dispatch(reviewer, "word_comments", {
        "doc": "d1", "action": "add", "author": "Rev",
        "items": [{"find": "99.9%", "text": "Per month?"}]}).ok
    refused = box.dispatch(reviewer, "word_changes", {"doc": "d1", "action": "accept",
                                                      "all": True})
    assert refused.error.code == "refused" and refused.error.details == {
        "role": "accept", "roles": ["review"]}
    batch = box.dispatch(reviewer, "batch", {"ops": [
        {"tool": "word_changes", "arguments": {"doc": "d1", "action": "list"}},
        {"tool": "replace_text", "arguments": {"doc": "d1", "find": "99.9%",
                                               "replace": "99%", "expect": "one"}}]})
    assert batch.ok and batch.data["failed"] == [1]
    # The editor-in-chief settles the builder's tracked change, and edits nothing.
    assert box.dispatch(chief, "word_changes", {"doc": "d1", "action": "list"}).data
    assert box.dispatch(chief, "word_changes", {"doc": "d1", "action": "accept",
                                                "all": True}).ok
    assert not box.dispatch(reviewer, "word_changes", {"doc": "d1", "action": "list"}).data
    assert box.dispatch(chief, "replace_text", {"doc": "d1", "find": "99.9%", "replace": "99%",
                                                "expect": "one"}).error.code == "refused"
    agents = [record.context["agent"] for record in session.log if record.context]
    assert {"builder-1", "reviewer-1", "chief-1"} == set(agents)


def test_an_editor_in_chief_s_prompt_follows_the_tools_sent(box):
    roles = ["review", "accept"]
    sent = box.definitions("anthropic", groups="core", roles=roles)
    assert "You may also accept and reject tracked changes." in \
        box.system_prompt(roles=roles, tools=sent)
    without = [d for d in sent if d.get("name") != "word_changes"]
    prompt = box.system_prompt(roles=roles, tools=without)
    assert "You may add comments" in prompt and "tracked changes" not in prompt
    assert box.system_prompt(roles=roles, groups="core") == \
        box.system_prompt(roles=roles, tools=sent)
