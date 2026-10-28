"""The Word tools against Claude itself (marked ``provider``): the loaded core's counted size,
and a request with the definitions as the adapter sends them.

Skipped unless ``ANTHROPIC_API_KEY`` is set; run with ``python -m pytest -m provider``.
``ANTHROPIC_MODEL`` picks the model (default ``claude-sonnet-5-5``, the trial model) and
``ANTHROPIC_WORKSPACE_ID``, when set, is sent as the ``anthropic-workspace-id`` header.  Plain
HTTPS, no SDK, as in ooxml-edit's online checks; neither the key nor any header is printed.

**The budget (tool roadmap, "Budgets"):** what a request loads -- the core definitions, the
hidden tool-use prompt included -- is at most 6,500 tokens as count-tokens counts them (raised from 5,500 by the user on
2026-10-11, to 6,500 on 2026-10-17).  The
offline proxy and the all-definitions guard are in the definition tests.
"""

from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request

import pytest

from ooxml_edit.tools import Toolbox
from ooxml_edit.tools.adapters import anthropic_problems

from docx_agent.tools import FORMAT, GROUPS, TOOLS

pytestmark = [
    pytest.mark.provider,
    pytest.mark.skipif(not os.environ.get("ANTHROPIC_API_KEY"),
                       reason="ANTHROPIC_API_KEY is not set"),
]

MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5")
API = "https://api.anthropic.com/v1"
#: Tokens the loaded core may add to a request, counted online.
CORE_COUNTED_BUDGET = 6500


def _post(path: str, body: dict) -> dict:
    headers = {"x-api-key": os.environ["ANTHROPIC_API_KEY"],
               "anthropic-version": "2023-06-01", "content-type": "application/json"}
    workspace = os.environ.get("ANTHROPIC_WORKSPACE_ID", "").strip()
    if workspace:
        headers["anthropic-workspace-id"] = workspace
    request = urllib.request.Request(f"{API}/{path}", data=json.dumps(body).encode(),
                                     method="POST", headers=headers)
    failure = None
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            reply = json.loads(response.read())
    except urllib.error.HTTPError as error:
        failure = f"{path}: HTTP {error.code}: {error.read().decode(errors='replace')}"
    if failure is not None:
        pytest.fail(failure, pytrace=False)   # outside the handler: no headers in a traceback
    if "usage" in reply:
        print(f"\n{path} usage: {json.dumps(reply['usage'])}")
    return reply


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def test_the_loaded_core_is_within_its_counted_budget(toolbox):
    messages = [{"role": "user", "content": "x"}]
    core = [d for d in toolbox.definitions("anthropic", groups="core") if "input_schema" in d]
    base = _post("messages/count_tokens", {"model": MODEL, "messages": messages})
    loaded = _post("messages/count_tokens", {"model": MODEL, "messages": messages,
                                             "tools": core})
    counted = loaded["input_tokens"] - base["input_tokens"]
    print(f"\n{MODEL}: the core, {len(core)} definitions, counts {counted} tokens")
    assert counted <= CORE_COUNTED_BUDGET, counted


def test_the_definitions_as_sent_are_accepted(toolbox):
    tools = toolbox.definitions("anthropic")      # core loaded, the rest deferred, strict planned
    assert anthropic_problems(tools) == []
    reply = _post("messages", {"model": MODEL, "max_tokens": 32, "tools": tools,
                               "system": toolbox.system_prompt(),
                               "output_config": {"effort": "low"},
                               "messages": [{"role": "user", "content": "Reply with: ready"}]})
    assert reply["type"] == "message"


#: A reviewer's role sets (ooxml-edit 0.45.0, round 7, item 9), an editor-in-chief's
#: (accept) among them: each its own core, cached.
REVIEW_SETS = (["review"], ["review", "resolve"], ["review", "accept"])


@pytest.mark.parametrize("roles", REVIEW_SETS, ids="+".join)
def test_a_reviewer_s_core_is_within_the_budget_and_accepted(toolbox, roles):
    messages = [{"role": "user", "content": "x"}]
    core = [d for d in toolbox.definitions("anthropic", groups="core", roles=roles)
            if "input_schema" in d]
    base = _post("messages/count_tokens", {"model": MODEL, "messages": messages})
    loaded = _post("messages/count_tokens", {"model": MODEL, "messages": messages,
                                             "tools": core})
    counted = loaded["input_tokens"] - base["input_tokens"]
    print(f"\n{MODEL}: the {'+'.join(roles)} core, {len(core)} definitions, counts {counted} "
          "tokens")
    assert counted <= CORE_COUNTED_BUDGET, counted
    tools = toolbox.definitions("anthropic", roles=roles)
    assert anthropic_problems(tools) == []
    reply = _post("messages", {"model": MODEL, "max_tokens": 32, "tools": tools,
                               "system": toolbox.system_prompt(roles=roles),
                               "output_config": {"effort": "low"},
                               "messages": [{"role": "user", "content": "Reply with: ready"}]})
    assert reply["type"] == "message"


# -- strict decoding keeps every field given -------------------------------------------------------


def _missing(given, got, where=""):
    """The dotted paths of ``given`` that ``got`` lacks: objects compared key by key, an
    array by its first item."""
    if isinstance(given, dict):
        if not isinstance(got, dict):
            return [where or "the arguments"]
        lost = []
        for key, value in given.items():
            path = f"{where}.{key}" if where else key
            lost += [path] if key not in got else _missing(value, got[key], path)
        return lost
    if isinstance(given, list) and given:
        if not isinstance(got, list) or not got:
            return [where]
        return _missing(given[0], got[0], where + "[]")
    return []


#: An item's field that names its actions by exception ("The comment (c:...), except for
#: add."; "The comments to act on (not list).").
_EXCEPT = re.compile(r"\(not (\w+)\)|except for (\w+)")


def _applies(tool, name, schema):
    """The actions of ``tool`` a field ``name`` (with ``schema``) is given with:
    :attr:`Tool.action_fields` where it names the field, else the actions its description
    starts with (``"accept, reject: ..."``, ``"list. ..."``), else every action but the one
    it says it is not for, else every action (``None``: the field is for any action)."""
    actions = tool.actions
    if name in tool.action_fields:
        return set(tool.action_fields[name])
    text = schema.get("description", "")
    lead = re.match(r"([a-z_]+(?:, [a-z_]+)*)[:.]", text)
    if lead and set(lead.group(1).split(", ")) <= set(actions):
        return set(lead.group(1).split(", "))
    unless = _EXCEPT.search(text)
    if unless and (unless.group(1) or unless.group(2)) in actions:
        return set(actions) - {unless.group(1) or unless.group(2)}
    return None


def _for_action(tool, schema, arguments, action):
    """The part of ``arguments`` (``schema``'s example) to give with ``action``: every field
    for it, an array of objects by the fields of its item that are; ``None`` when none is.
    A field for any action is given with each: a model asked for a few fields only now and
    then left out an array among them under strict decoding (``ids`` beside ``author``,
    ``items``: 2026-10-28), and kept them all when asked for the action's every field."""
    part = {}
    for key, value in arguments.items():
        prop = schema["properties"][key]
        applies = _applies(tool, key, prop)
        if applies is not None and action not in applies:
            continue
        if isinstance(value, list) and value and isinstance(value[0], dict):
            item = _for_action(tool, prop["items"], value[0], action)
            if item:
                part[key] = [item]
        else:
            part[key] = value
    return part or None


def _requests(tool, schema, arguments):
    """``arguments`` shared out over the actions ``schema`` offers, as few as cover every
    field, in their order: each request gives one action and only the fields for it, the
    required ones too (2026-10-28: a model asked for fields of two actions in one call
    sensibly made two calls, or left one action's fields out).  A tool without ``action``:
    one request with everything.  A field for none of the actions offered (a reviewer's
    word_notes offers ``list`` only) is not asked for; a tool left nothing to ask for is
    asked for its first action alone."""
    offered = (schema["properties"].get("action") or {}).get("enum")
    if not offered:
        return [arguments]
    required = set(schema.get("required", ())) - {"action"}
    done, requests = set(), []
    for action in offered:
        part = _for_action(tool, schema, {k: v for k, v in arguments.items()
                                          if k not in required and k != "action"},
                           action)
        if not part:
            continue
        new = set(_paths(part)) - done
        if not new:
            continue
        done |= new
        requests.append({**{k: arguments[k] for k in arguments if k in required},
                         "action": action, **part})
    # A tool none of whose fields is for an action offered is still called once.
    return requests or [{**{k: arguments[k] for k in arguments if k in required},
                         "action": offered[0]}]


def _paths(arguments, where=""):
    """The dotted paths of ``arguments``, an array by its first item."""
    for key, value in arguments.items():
        path = f"{where}.{key}" if where else key
        yield path
        if isinstance(value, list) and value and isinstance(value[0], dict):
            yield from _paths(value[0], path + "[]")


def _ask_for(name, schema, arguments):
    """A user message asking for one call with every argument, the optional ones listed
    first: a model writes the arguments it must give first, which is what lost trial 3's
    optional fields when they were listed before a required one.  The action, where there is
    one, comes before them all: the fields are its, and a model told the action last now and
    then made a first call without an array among them (``items``, 2026-10-28), and never did
    when told it first."""
    required = set(schema.get("required", ())) - {"action"}
    order = ([k for k in arguments if k == "action"]
             + [k for k in arguments if k not in required and k != "action"]
             + [k for k in arguments if k in required])
    lines = "\n".join(f"- {key}: {json.dumps(arguments[key])}" for key in order)
    return (f"This is a test of the tool interface; nothing is edited. Call {name} exactly "
            f"once, with exactly these arguments and values, none dropped or changed, and "
            f"nothing else:\n{lines}")


def test_strict_tools_keep_every_field_given(toolbox):
    """Each strict tool, as the adapter sends it, gets a scripted request per action whose
    fields it asks for (one for a tool without ``action``), asking for a call that uses them
    all; every field asked for must arrive (trial 3, S1)."""
    from ooxml_edit.tools.schema import example_arguments    # ooxml-edit 0.8.0

    strict = [d for d in toolbox.definitions("anthropic", defer=False) if d.get("strict")]
    assert strict
    # A reviewer's role sets send some tools strict that the default does not (round 7,
    # item 9): those too, each once, as that role set sends it.
    for roles in REVIEW_SETS:
        strict += [d for d in toolbox.definitions("anthropic", defer=False, roles=roles)
                   if d.get("strict") and d not in strict]
    lost, asked = {}, []
    for definition in strict:
        schema = definition["input_schema"]
        for arguments in _requests(toolbox.tools[definition["name"]], schema,
                                   example_arguments(schema)):
            label = definition["name"] + (f" {arguments['action']}" if "action" in arguments
                                          else "")
            asked.append(label)
            reply = _post("messages", {
                "model": MODEL, "max_tokens": 1500, "tools": [definition],
                "tool_choice": {"type": "auto"}, "output_config": {"effort": "low"},
                "messages": [{"role": "user", "content": _ask_for(
                    definition["name"], schema, arguments)}]})
            calls = [b for b in reply["content"] if b["type"] == "tool_use"]
            # The model may still split a request into several calls of the same tool (seen
            # with word_comments, round 17): the fields of all of them together are
            # checked, a field none of them has is still lost; another tool, or no call,
            # still fails.
            assert calls and {c["name"] for c in calls} == {definition["name"]}, \
                reply["content"]
            given = {}
            for call in calls:
                for key, value in call["input"].items():
                    given.setdefault(key, value)
            missing = _missing(arguments, given)
            if missing:
                lost[label] = missing
    print(f"\nstrict requests: {asked}")
    assert lost == {}, lost
