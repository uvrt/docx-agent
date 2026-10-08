"""Style: ``word_lists``, ``word_styles`` and ``word_template`` (``word_format``, the general
setter, is in :mod:`.text` and in this group)."""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, array, boolean, integer, string

from ..edit.formatting import RUN_PROPERTIES, read_paragraph, read_run
from ..edit.styles import StyleError
from ._base import need, outcome, resolve, word_tool
from .text import PARAGRAPH_FIELDS, RUN_FIELDS, paragraphs_of, split_formatting  # noqa: F401

DOC = string("Document id.")

#: What add and modify set: word_format's common fields.
STYLE_FIELDS = {k: v for k, v in {**RUN_FIELDS, **PARAGRAPH_FIELDS}.items()
                if k in ("bold", "italic", "size", "font", "color", "alignment", "space_before", "space_after",
                         "line_spacing", "indent_left", "indent_first", "keep_with_next")}


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
    for k, paragraph in enumerate(paragraphs):
        identifier = paragraph.id
        if action == "add":
            edits.append(document.add_to_list(identifier, kind, level=level or 0))
        elif action == "remove":
            edits.append(document.remove_from_list(identifier))
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
    return outcome(edits, f"{action}: {len(paragraphs)} paragraph(s)", data=lists[:50])


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
                         "indent_right", "first_line", "hanging", "keep_with_next", "page_break_before"):
                value = read_paragraph(paragraph, name)
                if value is not None:
                    declared[name] = value
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
           group="word_style", mutates=True, track=False)
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


TOOLS = [word_lists, word_styles, word_template]
