"""The Word tool definitions against each provider's rules, and the token budgets.

Offline: every canonical schema is in the common strict subset (``ooxml_edit.tools`` refuses
a tool that is not, when it is made); the Anthropic, OpenAI Responses and Chat Completions
adapters' output passes each provider's documented strict rules; no parameter is a path; the
groups are as the roadmap lists them.

**The budgets (tool roadmap, "Budgets", after T5's rationalisation).**  What binds is what a
request loads: the core, at most 5,500 tokens as Anthropic's count-tokens counts them
(``test_tools_online.py``, marked provider).  Offline, characters / 3.5 of each definition's
compact JSON stands in for it, times 1.48, the ratio of the counted to the estimated size
T5 and T5b measured on Sonnet 5.5 (the Word core: 2,384 estimated, 3,539 counted; every
definition: 11,046 and 15,544).  Every definition, deferred ones included, is held at what
the rationalisation left plus a small margin: a regression guard, not a budget.  T4 raised it
by 400 for edit_chart's add, data-label and gap-width actions (11,056 -> 11,402), which took
the place of a separate word_insert_chart; the core is unchanged.  Post-T4 lowered it:
word_insert_table and list_documents went (no model called them; a new table is Markdown),
11,402 -> 11,080 estimated, the core unchanged at 2,384.
"""

from __future__ import annotations

import json
import math

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402
from ooxml_edit.tools.adapters import anthropic_problems, openai_problems  # noqa: E402
from ooxml_edit.tools.schema import is_path_name, subset_problems, walk_properties  # noqa: E402

from docx_agent.tools import FORMAT, GROUPS, TOOLS, WORD_TOOLS  # noqa: E402

CHARS_PER_TOKEN = 3.5
COUNTED_PER_ESTIMATED = 1.48
CORE_COUNTED_BUDGET = 5500
ALL_GUARD = 11400


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def estimate(definitions: list[dict]) -> int:
    """Characters / 3.5 of the compact JSON of each definition as sent."""
    return math.ceil(sum(len(json.dumps(d, separators=(",", ":"), ensure_ascii=False))
                         for d in definitions) / CHARS_PER_TOKEN)


def test_every_canonical_schema_is_in_the_common_strict_subset():
    for tool in TOOLS:
        assert subset_problems(tool.canonical, allow_free=not tool.strict) == [], tool.name


@pytest.mark.parametrize("options", [{}, {"defer": True}], ids=["all-loaded", "deferred"])
def test_the_anthropic_definitions_keep_claudes_strict_rules(box, options):
    definitions = box.definitions("anthropic", **options)
    assert anthropic_problems(definitions) == []
    strict = [d for d in definitions if d.get("strict")]
    assert 0 < len(strict) <= 20
    if options:
        loaded = [d for d in definitions if "input_schema" in d and not d.get("defer_loading")]
        assert {d["name"] for d in loaded} == {t.name for t in TOOLS if t.group == "core"}
        assert loaded[-1].get("cache_control") and definitions[0]["type"].startswith("tool_search")


def test_the_openai_definitions_keep_the_strict_rules(box):
    responses = box.definitions("openai-responses")
    chat = box.definitions("openai-chat")
    assert openai_problems(responses) == [] and openai_problems(chat, chat=True) == []
    namespaces = box.definitions("openai-responses", namespaces=True, defer=True)
    assert openai_problems(namespaces) == []


def test_no_parameter_is_a_path_and_blobs_are_handles():
    for tool in TOOLS:
        for name, node, _ in walk_properties(tool.canonical):
            assert not is_path_name(name.split(".")[-1].rstrip("[]")), (tool.name, name)
    blobs = [(t.name, n) for t in TOOLS for n, node, _ in walk_properties(t.canonical) if "blob" in n or n == "image"]
    assert blobs and all("handle" in json.dumps(t.canonical) for t in TOOLS if t.name in {b[0] for b in blobs})


def test_the_groups_are_the_roadmaps():
    groups = {}
    for tool in WORD_TOOLS:
        groups.setdefault(tool.group, set()).add(tool.name)
    assert groups["core"] == {"word_read", "word_set_text"}        # and the shared describe
    assert groups["word_text"] == {"word_inspect", "word_insert_text", "word_delete", "word_insert_markdown"}
    assert groups["word_review"] == {"word_set_tracking", "word_changes", "word_comments"}
    assert groups["word_structure"] == {"word_move", "word_copy_from", "word_sections", "word_headers_footers",
                                        "word_fields", "word_notes", "word_links"}
    assert groups["word_objects"] == {"word_edit_table", "word_format_table",
                                      "word_drawings", "word_controls"}
    assert groups["word_style"] == {"word_format", "word_lists", "word_styles", "word_template"}
    assert all(len(names) < 10 for group, names in groups.items() if group != "core")


def test_the_loaded_core_is_within_budget_and_every_definition_within_its_guard(box):
    core = estimate(box.definitions("anthropic", groups="core"))
    every = estimate(box.definitions("anthropic", defer=False))
    assert core * COUNTED_PER_ESTIMATED <= CORE_COUNTED_BUDGET, core
    assert every <= ALL_GUARD, every


def test_descriptions_stay_short():
    for tool in WORD_TOOLS:
        assert len(tool.description.split()) <= 40, tool.name
        for name, node, _ in walk_properties(tool.canonical):
            assert len(node.get("description", "").split()) <= 15, (tool.name, name)
