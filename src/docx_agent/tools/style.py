"""Style: ``word_lists``, ``word_styles``, ``word_template``, ``word_design_facts`` and ``word_compare_blocks``
(``word_format``, the general setter, is in :mod:`.text` and in this group)."""

from __future__ import annotations

from typing import Any

import json

from ooxml_edit.tools import REVIEW, Result, ToolError, array, boolean, integer, obj, page_list, string

from ..edit.design import ALL_KINDS, CAPS, KINDS
from ..edit.formatting import RUN_PROPERTIES, read_paragraph, read_run, read_shading
from ..edit.styles import StyleError
from ._base import need, outcome, resolve, tool_names, word_tool
from .text import PARAGRAPH_FIELDS, RUN_FIELDS, paragraphs_of, split_formatting  # noqa: F401

DOC = string("Document id.")

#: What add and modify set: word_format's common fields.
STYLE_FIELDS = {k: v for k, v in {**RUN_FIELDS, **PARAGRAPH_FIELDS}.items()
                if k in ("bold", "italic", "size", "font", "color", "alignment", "space_before", "space_after",
                         "line_spacing", "indent_left", "indent_first", "keep_with_next")}
# Round 24, W6: a paragraph style's shading (behind its paragraphs), a character style's.
STYLE_FIELDS["shading"] = string("Fill: RRGGBB, theme colour, 'accent1 lighter 80%' or none.", optional=True)


# -- W10 word_lists ----------------------------------------------------------------------------


@word_tool("word_lists", "Lists: add paragraphs to a bullet or numbered list, remove them, set level, restart or continue numbering, change a level's number format.",
           {"doc": DOC,
            "targets": array(string(), "Paragraphs or spans.", min_items=1),
            "action": string("What to do.", enum=["add", "remove", "level", "restart", "continue", "format"]),
            "kind": string("add. Default bullet.", enum=["bullet", "number"], optional=True),
            "level": integer("add, level, format: level from 0.", minimum=0, maximum=8, optional=True),
            "start": integer("restart: first number (default 1); format: start value.",
                             minimum=0, maximum=32767, optional=True),
            "number_format": string("format: decimal, lowerLetter, upperRoman, bullet...",
                                    optional=True),
            "label": string("format: label, e.g. %1) or a bullet.", optional=True)},
           refs=('targets[]',), group="word_style", mutates=True)
def word_lists(call: Any, doc: str, targets: list[str], action: str, kind: str = "bullet",
               level: int | None = None, start: int | None = None, number_format: str | None = None,
               label: str | None = None) -> Result:
    document = call.document
    paragraphs = [p for address in resolve(call, targets) for p in paragraphs_of(document, address)]
    edits = []
    listless = []
    for k, paragraph in enumerate(paragraphs):
        identifier = paragraph.id
        if action == "add":
            edits.append(document.add_to_list(identifier, kind, level=level or 0))
        elif action == "remove":
            # Round 27: a paragraph in no list is as asked; one note names them all.
            edit = document.remove_from_list(identifier)
            if edit.changed:
                edits.append(edit)
            else:
                listless.append(edit.id)
        elif action == "level":
            edits.append(document.set_list_level(identifier, need(level, "level")))
        elif action == "restart":
            if k == 0:
                edits.append(document.restart_numbering(identifier, at=1 if start is None else start))
        elif action == "continue":
            if k == 0:
                edits.append(document.continue_numbering(identifier))
        else:
            if k == 0:
                edits.append(document.set_list_format(identifier, level=level, format=number_format, text=label,
                                                      start=start))
    lists = []
    for paragraph in paragraphs:
        membership = document.list_of(paragraph.id)
        lists.append({"id": paragraph.id, "list": None if membership is None else
                      {"kind": membership.kind, "level": membership.level, "num_id": membership.num_id}})
    result = outcome(edits, f"{action}: {len(paragraphs)} paragraph(s)", data=lists[:50])
    if listless:
        named = ", ".join(listless[:10]) + (f" and {len(listless) - 10} more" if len(listless) > 10 else "")
        result.warnings.append(f"{len(listless)} paragraph(s) were not in a list, left as they were: {named}")
    return result


# -- W21 word_styles ---------------------------------------------------------------------------


def _style_element(document: Any, style_id: str) -> Any:
    from ..edit.styles import _W

    part = document.package.related_parts_of_type(
        document.package.document_part(),
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
    if not part:
        return None
    root = document.package.tree(part[0])
    for node in root.iter(_W + "style"):
        if node.get(_W + "styleId") == style_id:
            return node
    return None


def describe_style(document: Any, style: Any) -> dict[str, Any]:
    from ..edit.styles import _W

    data: dict[str, Any] = {"name": style.name, "kind": style.kind, "based_on": style.based_on}
    node = _style_element(document, style.id)
    if node is not None:
        run = node.find(_W + "rPr")
        paragraph = node.find(_W + "pPr")
        declared = {}
        for name in sorted(RUN_PROPERTIES):
            value = read_run(run, name)
            if value is not None:
                declared[name] = value
        if style.kind == "paragraph":
            for name in ("alignment", "space_before", "space_after", "line_spacing", "indent_left",
                         "indent_right", "first_line", "hanging", "keep_with_next", "keep_together",
                         "page_break_before"):
                value = read_paragraph(paragraph, name)
                if value is not None:
                    declared["keep_lines" if name == "keep_together" else name] = value
        shading = read_shading(paragraph if style.kind == "paragraph" else run)
        if shading is not None:
            declared["shading"] = shading
        data["declared"] = declared
    data["usage"] = document.styles.usage(style.name, style.kind)
    return data


@word_tool("word_styles", "The style sheet: list, describe (formatting, usage), add, modify, remove (with a replacement while used), purge unused.",
           {"doc": DOC,
            "action": string("What to do.", enum=["list", "describe", "add", "modify", "remove", "purge_unused"]),
            "name": string("Style name.", optional=True),
            "kind": string("Default paragraph.",
                           enum=["paragraph", "character", "table"], optional=True),
            "based_on": string("add: based on this style.", optional=True),
            "next_style": string("add: style of the next paragraph.", optional=True),
            "replacement": string("remove: style its users get instead.", optional=True),
            "in_use_only": boolean("list: only styles in use.", optional=True),
            **STYLE_FIELDS},
           group="word_style", mutates=True, track=False,
           action_roles={"list": REVIEW, "describe": REVIEW})
def word_styles(call: Any, doc: str, action: str, name: str | None = None, kind: str | None = None,
                based_on: str | None = None, next_style: str | None = None, replacement: str | None = None,
                in_use_only: bool = False, **values: Any) -> Result:
    document = call.document
    styles = document.styles
    try:
        if action == "list":
            rows = []
            for style in styles:
                if not style.name or (kind and style.kind != kind):
                    continue
                if in_use_only and not styles.usage(style.name, style.kind).get("content"):
                    continue
                rows.append({"name": style.name, "kind": style.kind, "based_on": style.based_on})
            return Result(summary=f"{len(rows)} style(s)", data=rows[:200])
        if action == "purge_unused":
            return outcome([styles.purge_unused()], "Purged unused styles")
        name = need(name, "name")
        if action == "describe":
            return Result(summary=f"Style {name}", data=describe_style(document, styles.get(name, kind)))
        if action == "remove":
            return outcome([styles.remove(name, kind, replacement=replacement)], f"Removed style {name}")
        run, paragraph = split_formatting(values)
        formatting = {**run, **paragraph}
        if values.get("shading") is not None:
            formatting["shading"] = None if values["shading"].strip().lower() == "none" else values["shading"]
        if action == "add":
            edit = styles.add(name, kind or "paragraph", based_on=based_on, next_style=next_style, **formatting)
            return outcome([edit], f"Added style {name}")
        if not formatting:
            raise ToolError("invalid_arguments", "modify needs formatting to set",
                            valid_options=list(STYLE_FIELDS))
        return outcome([styles.modify(name, kind, **formatting)], f"Modified style {name}")
    except StyleError as error:
        raise ToolError("not_found", str(error), field="name",
                        valid_options=[s.name for s in styles if s.name and (kind is None or s.kind == kind)][:50]) \
            from None


# -- W29 word_template -------------------------------------------------------------------------


@word_tool("word_template", "Upgrade the document to Word's current compatibility mode, reporting the reflow. To save it as a template, save_document with format dotx.",
           {"doc": DOC,
            "action": string("What to do.", enum=["upgrade_to_modern"])},
           batchable=False, group="word_style", mutates=True, track=False)
def word_template(call: Any, doc: str, action: str) -> Result:
    from ._base import reflow_json

    document = call.document
    edit = document.upgrade_to_modern()
    data = {"mode": document.compatibility_mode,
            "reflow": reflow_json(edit.reflow) if edit.reflow is not None else None}
    return outcome([edit], f"Compatibility mode {document.compatibility_mode}", data=data)


# -- F9 word_design_facts ---------------------------------------------------------------------


@word_tool("word_design_facts", "Measured design facts for template analysis, no verdicts: styles in use per role with usage and effective look, table styles by place (header row, bands, label column), page grid per section, theme, repeated block shapes; or every declared style resolved.",
           {"doc": DOC,
            "include": array(string(enum=list(ALL_KINDS)), "Facts to report. Default all but declared_styles.",
                             optional=True),
            "cursor": string("next_cursor (declared_styles).", optional=True)},
           group="word_style")
def word_design_facts(call: Any, doc: str, include: list[str] | None = None, cursor: str | None = None) -> Result:
    """:meth:`Document.design_facts <docx_agent.Document.design_facts>`; a result too long
    for the call's limit keeps fewer entries per list (``omitted`` counts the rest), except
    ``declared_styles``, which is paged: as many styles as fit, ``next_cursor`` for the rest
    (a later page holds only them)."""
    document = call.document
    kinds = list(include) if include else list(KINDS)
    declared = "declared_styles" in kinds or cursor is not None
    others = [] if cursor is not None else [k for k in kinds if k != "declared_styles"]
    caps = dict(CAPS)
    limit = call.limits.max_result_chars - 1500
    data: dict[str, Any] = {}
    while others:
        data = document.design_facts(others, caps=caps)
        if len(json.dumps(data, ensure_ascii=False)) <= (limit // 2 if declared else limit) \
                or max(caps.values()) <= 1:
            break
        caps = {k: max(1, v // 2) for k, v in caps.items()}
    counts = [f"{len(data[k])} {k.replace('_', ' ')}" for k in ("styles", "table_styles", "sections", "patterns")
              if k in data]
    total, next_cursor = None, None
    if declared:
        styles = document.design_facts(["declared_styles"], caps={"declared_styles": 10 ** 6})["declared_styles"]
        page, total, _ = page_list(styles, cursor=cursor, limit=len(styles))
        start = len(styles) - len(page)
        room = limit - len(json.dumps(data, ensure_ascii=False))
        kept: list = []
        for entry in page:
            room -= len(json.dumps(entry, ensure_ascii=False)) + 2
            if room < 0 and kept:
                break
            kept.append(entry)
        data["declared_styles"] = kept
        end = start + len(kept)
        next_cursor = f"c{end}" if end < total else None
        counts.append(f"declared styles {start + 1}-{end} of {total}" if kept else f"no declared styles of {total}")
    data = tool_names(data, "look")             # a paragraph's keep_lines, as word_format says it
    left = data.get("omitted")
    return Result(summary="Design facts: " + ", ".join(counts) + (
        "; " + ", ".join(f"{n} more {k.replace('_', ' ')}" for k, n in left.items()) + " left out (include fewer kinds)"
        if left else ""), data=data, total=total, next_cursor=next_cursor)


# -- round 17, P3 word_compare_blocks ---------------------------------------------------------

#: Differences in one page of a comparison.
COMPARE_PAGE = 60


def _compared_documents(arguments: Any) -> list[str]:
    """``doc`` and, when it names another, the reference's document."""
    out = [arguments["doc"]] if isinstance(arguments.get("doc"), str) else []
    reference = arguments.get("reference")
    other = reference.get("doc") if isinstance(reference, dict) else None
    if isinstance(other, str) and other not in out:
        out.append(other)
    return out


@word_tool("word_compare_blocks", "Compare tables or paragraphs with a reference block (a sample, here or in another open document): rows, merges, widths, styles, shading, borders, direct formatting; each difference with fix calls. Measures only.",
           {"doc": DOC,
            "blocks": array(string(), "Blocks to compare, t:... or p:....", min_items=1),
            "reference": obj({"doc": string("Its document. Default doc.", optional=True),
                              "block": string("Its block.")}, "The block with the expected values."),
            "cursor": string("next_cursor of the previous page.", optional=True)},
           refs=('blocks[]',), group="word_style", documents=_compared_documents)
def word_compare_blocks(call: Any, doc: str, blocks: list[str], reference: dict,
                        cursor: str | None = None) -> Result:
    """:meth:`Document.compare_blocks <docx_agent.Document.compare_blocks>`, paged by
    difference."""
    other = reference.get("doc") or doc
    if other not in call.entries:
        raise ToolError("not_found", f"no open document {other!r}", field="reference.doc",
                        valid_options=list(call.entries)[:50])
    source = call.entries[other].document
    data = call.document.compare_blocks(resolve(call, blocks), reference["block"],
                                        reference_document=source if other != doc else None)
    differences, total, next_cursor = page_list(data["differences"], cursor=cursor, limit=COMPARE_PAGE)
    data["differences"] = differences
    fixes = sum(1 for d in differences if d["fix"])
    return Result(summary=f"Compared {len(data['compared'])} block(s) with {data['reference']}: "
                  f"{len(data['same'])} the same, {total} difference(s) ({fixes} with a fix on this page)",
                  data=data, total=total, next_cursor=next_cursor)


TOOLS = [word_lists, word_styles, word_template, word_design_facts, word_compare_blocks]
