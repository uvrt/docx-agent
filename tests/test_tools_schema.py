"""The Word tool definitions against each provider's rules, and the token budgets.

Offline: every canonical schema is in the common strict subset (``ooxml_edit.tools`` refuses
a tool that is not, when it is made); the Anthropic, OpenAI Responses and Chat Completions
adapters' output passes each provider's documented strict rules; no parameter is a path; the
groups are as the roadmap lists them.

**The budgets (tool roadmap, "Budgets", after T5's rationalisation).**  What binds is what a
request loads: the core, at most 6,500 tokens as Anthropic's count-tokens counts them
(raised from 5,500 by the user on 2026-10-11, to 6,500 on 2026-10-17)
(``test_tools_online.py``, marked provider).  Offline, characters / 3.5 of each definition's
compact JSON stands in for it, times 1.48, the ratio of the counted to the estimated size
T5 and T5b measured on Sonnet 5.5 (the Word core: 2,384 estimated, 3,539 counted; every
definition: 11,046 and 15,544).  Every definition, deferred ones included, is held at what
the rationalisation left plus a small margin: a regression guard, not a budget.  T4 raised it
by 400 for edit_chart's add, data-label and gap-width actions (11,056 -> 11,402), which took
the place of a separate word_insert_chart; the core is unchanged.  Post-T4 lowered it:
word_insert_table and list_documents went (no model called them; a new table is Markdown),
11,402 -> 11,080 estimated, the core unchanged at 2,384.  The radar chart type in
edit_chart's enum added 3 (11,083); the guard stays.  ooxml-edit 0.12's optional scope on
the shared undo added 33 to the core (2,384 -> 2,417, about 3,577 counted by the proxy) and
to every definition (11,116), though a Word document refuses it (it undoes document-wide);
budget and guard stay.  ooxml-edit 0.14's slide ids in render's and check's ``slides`` (the
items a string: ``s:256``, ``$ref`` or a number) added 12 to the core (2,417 -> 2,429, about
3,595 counted by the proxy) and to every definition (11,128); a document's ``pages`` stay
numbers (``"1"`` is refused, ``pages[0] must be a number``).  Budget and guard stay.  Round
3's chart formatting in the deferred edit_chart (axis scales, gridlines, series and point
colours and markers, series type, chart text, point labels; the radar_filled and bubble types)
added 646 to every definition (11,774), none to the core; the guard rises to 11,850 rather
than squeeze the descriptions.  ooxml-edit 0.18's workbook kind in the shared tools
(``describe``'s and ``find_text``'s descriptions, find_text's ``range``, ``xlsx`` in
new_document's ``kind``, ``xlsx``, ``xlsm`` and ``xltx`` in save_document's ``format``) added
64 to the core (2,471 -> 2,535, about 3,752 counted by the proxy, within 6,000) and to every
definition (11,880); the guard rises to 11,950.  ooxml-edit 0.22 (O3) put ``from`` and
``columns`` on the deferred edit_chart (about 11,933).  X1, tables from data (``from``,
``from_columns`` and ``header`` on word_insert_markdown and word_edit_table) added 194 to
every definition (12,127), none to the core; the guard rises to 12,150, the next 50, rather
than squeeze descriptions the pptx-agent twins share.  With seven optional parameters
word_insert_markdown keeps its strict place, and edit_smartart (two) no longer fits Claude's
24 beside the others: it is sent non-strict (the dispatcher still validates it), 23 optional
parameters over the strict tools.  ooxml-edit 0.26.0 (round 5, T1): ``batch`` applies its
ops one by one by default and takes ``atomic`` for all or none; its description and the new
parameter add 22 to the core (2,535 -> 2,558, about 3,786 counted) and to every definition
(12,131 -> 12,153); the guard rises to 12,155.  The strict set is unchanged: ``batch`` is
never strict.  ooxml-edit 0.32.0 (round 5, T2): a chart added with its look in one call
(``data_labels`` and ``colors`` on the deferred edit_chart, ``add`` in ``style``'s and
``part``'s descriptions, "Batch several changes.") adds 63 to every definition (12,153 ->
12,216), none to the core (2,558); the guard rises to 12,220.  The strict set is unchanged:
edit_chart left it in T4.  ooxml-edit 0.36.0 (round 6, R6): find_text's and replace_text's
``slides`` take slide ids as render's do (strings with ``orInteger``, "or ids (s:256,
$ref)"; decks only) and add 12 to the core (2,558 -> 2,570, about 3,804 counted) and to
every definition (12,216 -> 12,228); the guard rises to 12,230.  The strict set changes:
replace_text's ``slides`` items are strings now.
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

#: Round 6, R9: date formats -- ``date_format`` on word_insert_markdown and word_edit_table (a
#: table from rows, its dates as text in a pattern) and on the shared edit_chart (ooxml-edit
#: 0.38.0: add from, its category dates kept as an Excel code) -- add 84 to every definition
#: (12,228 -> 12,312), none to the core (2,570).  The guard rises to 12,315.  The strict set
#: is unchanged in its members; word_insert_markdown's and word_edit_table's schemas gain one
#: optional string each.
#: Round 6, R10: edit_chart's chart_type enum gains ``dot`` (ooxml-edit 0.40.0): +2 to every
#: definition (12,312 -> 12,314), none to the core; the guard stays 12,315.  edit_chart is
#: sent non-strict, so the strict set is unchanged.
#: Round 9, N1: render's ``sheet`` (``{columns, width}``: pages tiled into one labelled image,
#: ooxml-edit 0.49.0) adds 64 to the core (2,570 -> 2,634) and to every definition (12,314 ->
#: 12,378).  The guard rises to 12,380.  The strict set is unchanged in its members: render
#: is not sent strict to Claude; on OpenAI, where it is, ``sheet`` is one more object.
#: Round 10, item 4: the shared undo's ``scope`` says what each kind takes (ooxml-edit
#: 0.55.0): the core 2,634 -> 2,635, every definition 12,378 -> 12,379; the guard stays
#: 12,380.  The strict set is unchanged: undo is not strict, and no parameter is added.
#: Round 12, Q1: the shared check's ``spacing`` (decks: the gaps, padding and frame
#: clearances below that many pt, ooxml-edit 0.56.0; a document refuses it) adds 35 to the
#: core (2,635 -> 2,670) and to every definition (12,379 -> 12,414).  The guard rises to
#: 12,415.  The strict set is unchanged in its members: check is a reading tool, not sent
#: strict to Claude; on OpenAI, where it is, ``spacing`` is one more nullable number.
#: Round 15, T2: the shared check's ``contrast`` (decks: text below that contrast ratio with
#: what is behind it, ooxml-edit 0.59.0; a document refuses it) adds 37 to the core
#: (2,670 -> 2,707) and to every definition (12,414 -> 12,451).  The guard rises to 12,452.
#: The strict set is unchanged in its members, as with spacing.
#: Round 16, F11: the shared save_document's ``clean`` (orphaned parts, rsids, web divs and
#: an external template link taken out of the file saved, ooxml-edit 0.60.0) adds 42 to the
#: core (2,707 -> 2,749) and to every definition (12,451 -> 12,493).  The guard rises to
#: 12,494.  save_document is not sent strict to Claude; on OpenAI ``clean`` is one more
#: nullable boolean.
CHARS_PER_TOKEN = 3.5
COUNTED_PER_ESTIMATED = 1.48
CORE_COUNTED_BUDGET = 6500
ALL_GUARD = 12494


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
