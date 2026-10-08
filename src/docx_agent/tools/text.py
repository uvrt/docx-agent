"""Text: ``word_set_text``, ``word_insert_text``, ``word_delete``, ``word_insert_markdown`` and
the general setter ``word_format``."""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import (Result, ToolError, array, boolean, number, obj, string)

from ..edit.document import Paragraph, Table, W_P
from ..edit.ranges import TextRange
from ..edit.styles import StyleError
from ..markdown.stylemap import StyleMap
from ._base import need, one_of, outcome, remember, resolve, short, word_tool

DOC = string("Document id.")
REF = string("Ref name for $name.", optional=True)


# -- addresses ---------------------------------------------------------------------------------


def text_range(document: Any, address: str) -> TextRange:
    """A text range from an address: a range id (``p:A@4:11``, ``p:A@5``, ``p:A@3..p:B@8``),
    or a paragraph id, which is the whole paragraph's text."""
    if "@" in address and not address.startswith("p@"):
        return document.range(address)
    if address.startswith("p@") and address.count("@") > 1:
        return document.range(address)
    found = document.get(address)
    if isinstance(found, Paragraph):
        return found.range()
    raise ToolError("invalid_arguments", f"{address} is not text: give a paragraph or range id",
                    field="target")


def find_range(call: Any, find: str, within: str | None) -> TextRange:
    return call.document.anchor(find, within=resolve(call, within) if within else None)


def paragraphs_of(document: Any, address: str) -> list[Paragraph]:
    """The paragraphs an address covers: one, a span's (``p:A..p:B``, tables' cells
    included), a range's, or a table's."""
    if ".." in address and "@" not in address:
        elements, part = document._blocks_of(address)
        out = []
        index = document._index(part)
        for element in elements:
            nodes = [element] if element.tag == W_P else list(element.iter(W_P))
            for node in nodes:
                entry = index.entry_for(node)
                if entry is not None:
                    out.append(document.paragraph(entry.id))
        return out
    if "@" in address and not address.startswith("p@"):
        return [document.paragraph(i) for i in document.range(address).paragraph_ids()]
    found = document.get(address)
    if isinstance(found, Paragraph):
        return [found]
    if isinstance(found, Table):
        return [p for row in range(len(found.rows)) for col in range(found.column_count)
                for p in _cell_paragraphs(found, row, col)]
    if hasattr(found, "paragraphs"):
        paragraphs = found.paragraphs
        return list(paragraphs() if callable(paragraphs) else paragraphs)
    raise ToolError("invalid_arguments", f"{address} holds no paragraphs", field="targets")


def _cell_paragraphs(table: Table, row: int, column: int) -> list[Paragraph]:
    try:
        return list(table.cell(row, column).paragraphs)
    except IndexError:
        return []


# -- W5 word_set_text --------------------------------------------------------------------------


SET_ITEM = obj({"target": string("Paragraph, or table cell t:X/c<row>,<col>."),
                "text": string("The whole new text, plain.")})


@word_tool("word_set_text", "Rewrite whole paragraphs, keeping the formatting; tracked, only changed words become revisions. For words within a paragraph, replace_text.",
           {"doc": DOC,
            "target": string("Paragraph, or table cell t:X/c<row>,<col>.", optional=True),
            "text": string("The whole new text, plain.", optional=True),
            "items": array(SET_ITEM, "Or several paragraphs.", optional=True)},
           refs=('target', 'items[].target'), group="core", mutates=True)
def word_set_text(call: Any, doc: str, target: str | None = None, text: str | None = None,
                  items: list[dict] | None = None) -> Result:
    one_of({"target": target, "items": items}, "target", "items")
    work = items if items is not None else [{"target": target, "text": need(text, "text")}]
    edits = []
    for item in work:
        address = resolve(call, item["target"])
        found = call.document.get(address)
        if not isinstance(found, Paragraph):
            paragraphs = getattr(found, "paragraphs", None)
            if not paragraphs:
                raise ToolError("invalid_arguments", f"{address} is not a paragraph or a cell",
                                field="target")
            found = (paragraphs() if callable(paragraphs) else paragraphs)[0]
        edits.append(found.set_text(item["text"]))
    return outcome(edits, f"Set the text of {len(edits)} paragraph(s)")


# -- W6 word_insert_text -----------------------------------------------------------------------


INSERT_FIELDS = {
    "find": string("Text occurring once, to insert beside.", optional=True),
    "target": string("Or a range, position (p:A@5) or paragraph.", optional=True),
    "where": string("Default after.", enum=["before", "after"], optional=True),
    "text": string("Plain text."),
    "ref": REF,
}


@word_tool("word_insert_text", "Insert text beside words found once, or at a range or position, in the formatting there. Returns the new ranges' ids.",
           {"doc": DOC, **{k: (v if k != "text" else string("Plain text.", optional=True))
                           for k, v in INSERT_FIELDS.items()},
            "within": string("Block or span to search for find.", optional=True),
            "items": array(obj(INSERT_FIELDS), "Or several insertions.",
                           optional=True)},
           refs=('target', 'within', 'items[].target'), group="word_text", mutates=True)
def word_insert_text(call: Any, doc: str, items: list[dict] | None = None, **single: Any) -> Result:
    if items is not None and any(single.get(k) is not None for k in ("find", "target", "text")):
        raise ToolError("invalid_arguments", "give one insertion's fields or items, not both", field="items")
    work = items if items is not None else [single]
    edits, made = [], []
    for item in work:
        need(item.get("text"), "text")
        where = item.get("where") or "after"
        key = one_of(item, "find", "target")
        if key == "find":
            found = find_range(call, item["find"], item.get("within"))
        else:
            found = text_range(call.document, resolve(call, item["target"]))
        place = found.id
        edit = found.insert_before(item["text"]) if where == "before" else found.insert_after(item["text"])
        edits.append(edit)
        inserted = inserted_range(place, where, item["text"], edit.renamed)
        made.append(inserted)
        remember(call, item.get("ref"), inserted)
    result = outcome(edits, f"Inserted text at {len(edits)} place(s)", data={"ranges": made})
    return result


def inserted_range(place: str, where: str, text: str, renamed: dict[str, str]) -> str:
    """The id of text just inserted before or after the range ``place`` (its id before the
    insertion): where the range started or ended, ``len(text)`` characters on."""
    head, _, tail = place.partition("..")
    anchor = head if where == "before" or not tail else tail
    paragraph, _, span = anchor.rpartition("@")
    offsets = span.split(":")
    at = int(offsets[0] if where == "before" else offsets[-1])
    paragraph = renamed.get(paragraph, paragraph)
    return f"{paragraph}@{at}:{at + len(text)}"


# -- W7 word_delete ----------------------------------------------------------------------------


@word_tool("word_delete", "Delete text (found once, or a range) or whole blocks: paragraph, table, span p:A..t:B, picture, note.",
           {"doc": DOC,
            "find": string("Text occurring once (in within, if given).", optional=True),
            "within": string("Block or span to search for find.", optional=True),
            "target": string("Range, block, span p:A..t:B, picture d:, note fn:/en:.",
                             optional=True),
            "collapse_space": boolean("Also delete the space the deletion would leave doubled. "
                                      "Default false.", optional=True)},
           refs=('target', 'within'), group="word_text", mutates=True, exactly_one=[("find", "target")])
def word_delete(call: Any, doc: str, find: str | None = None, within: str | None = None,
                target: str | None = None, collapse_space: bool = False) -> Result:
    document = call.document
    if find is not None:
        found = find_range(call, find, within)
        text = found.text
        return outcome([found.delete(collapse_space=collapse_space)], f"Deleted {short(text, 60)!r}")
    target = resolve(call, target)
    if "@" in target and not target.startswith("p@"):
        return outcome([document.delete_text(target, collapse_space=collapse_space)], f"Deleted {target}")
    if target.startswith(("fn:", "en:")):
        return outcome([document.delete_note(target)], f"Deleted note {target}")
    if target.startswith("d:"):
        return outcome([document.picture(target).delete()], f"Deleted picture {target}")
    if ".." in target:
        elements, part = document._blocks_of(target)
        ids = [document._block_id(part, element) for element in elements]
        return outcome([document.delete_block(i) for i in ids], f"Deleted {len(ids)} block(s)")
    return outcome([document.delete_block(target)], f"Deleted {target}")


# -- W8 word_insert_markdown -------------------------------------------------------------------


STYLE_KEYS = ["paragraph", "h1", "h2", "h3", "h4", "h5", "h6", "quote", "code", "inline_code",
              "bullet", "bullet2", "bullet3", "number", "number2", "number3", "table", "table_cell",
              "footnote", "footnote_reference", "emphasis", "strong", "link"]

STYLE_MAP = array(obj({"element": string("Markdown element.", enum=STYLE_KEYS),
                       "style": string("The document's style name for it.")}),
                  "Styles differing from the defaults (h1 Heading 1, paragraph Normal...).", optional=True)


def style_map(entries: list[dict] | None) -> dict | None:
    if not entries:
        return None
    return {entry["element"]: entry["style"] for entry in entries}


def images_from(call: Any):
    def image(src: str) -> bytes:
        for blob in call.session.blobs.values():
            if src in (blob.handle, blob.name):
                if not blob.mime.startswith("image/"):
                    raise ToolError("invalid_arguments", f"{src} is a {blob.mime} blob, not an image")
                return blob.data
        raise ToolError("not_found", f"no image blob {src!r}: use a blob handle as the image source",
                        valid_options=[f"{b.handle} ({b.name})" for b in call.session.blobs.values()
                                       if b.mime.startswith("image/")])

    return image


@word_tool("word_insert_markdown", "Write Markdown as blocks in the document's styles: headings, lists, tables, links, footnotes, images ![alt](b1). Returns the new blocks' ids.",
           {"doc": DOC,
            "markdown": string("Markdown.", optional=True),
            "blob": string("Or a Markdown blob handle.", optional=True),
            "at": string("end, after:<id>, before:<id> or replace:<id>[..<id>]."),
            "style_map": STYLE_MAP,
            "key": string("Retry key: a repeat with the same key makes nothing new.", optional=True)},
           group="word_text", mutates=True, exactly_one=[("markdown", "blob")])
def word_insert_markdown(call: Any, doc: str, at: str, markdown: str | None = None,
                         blob: str | None = None, style_map: list[dict] | None = None,
                         key: str | None = None) -> Result:
    if blob is not None:
        source = call.blob(blob)
        try:
            markdown = source.data.decode("utf-8")
        except UnicodeDecodeError:
            raise ToolError("invalid_arguments", f"{blob} is not UTF-8 text", field="blob") from None
    at = _at(call, at)
    try:
        edit = call.document.insert_markdown(markdown, at=at, style_map=_style_map(style_map),
                                             images=images_from(call), fetch=None, html="refuse")
    except StyleError as error:
        raise ToolError("not_found", str(error), field="style_map",
                        valid_options=[s.name for s in call.document.styles][:50]) from None
    result = outcome([edit], f"Inserted {len(edit.blocks)} block(s)", data={"blocks": edit.blocks})
    return result


def _style_map(entries: list[dict] | None) -> Any:
    mapping = style_map(entries)
    return StyleMap.from_dict(mapping) if mapping else None


def _at(call: Any, at: str) -> str:
    """``at`` with any ``$ref`` inside resolved: ``after:$intro``."""
    if "$" not in at:
        return at
    head, _, rest = at.partition(":") if at.split(":", 1)[0] in ("after", "before", "replace") else ("", "", at)
    parts = [resolve(call, part) if part.startswith("$") else part for part in rest.split("..")]
    return (head + ":" if head else "") + "..".join(parts)


# -- W9 word_format ----------------------------------------------------------------------------


RUN_FIELDS = {
    "bold": boolean("Bold.", optional=True),
    "italic": boolean("Italic.", optional=True),
    "underline": boolean("Underline.", optional=True),
    "strike": boolean("Strikethrough.", optional=True),
    "size": number("Font size.", minimum=1, maximum=1638, optional=True),
    "font": string("Font name.", optional=True),
    "color": string("Theme colour or RRGGBB.", optional=True),
    "highlight": string("yellow, green... or none.", optional=True),
}
PARAGRAPH_FIELDS = {
    "alignment": string("Alignment.", enum=["left", "center", "right", "justify"], optional=True),
    "space_before": number("Before the paragraph.", minimum=0, maximum=1584, optional=True),
    "space_after": number("After the paragraph.", minimum=0, maximum=1584, optional=True),
    "line_spacing": number("Times single, e.g. 1.15.", minimum=0.06, maximum=132,
                           optional=True),
    "indent_left": number("Left indent.", minimum=-1584, maximum=1584, optional=True),
    "indent_right": number("Right indent.", minimum=-1584, maximum=1584, optional=True),
    "indent_first": number("First line; negative hangs.", minimum=-1584,
                           maximum=1584, optional=True),
    "keep_with_next": boolean("Keep with next.", optional=True),
    "page_break_before": boolean("Page break before.", optional=True),
}


def known_style(document: Any, name: str, kind: str, field: str) -> None:
    """``not_found`` with the document's styles of ``kind`` unless ``name`` names one, or a
    built-in style Word would add."""
    try:
        document.styles.resolve(name, kind)
    except StyleError as error:
        raise ToolError("not_found", str(error), field=field,
                        valid_options=[s.name for s in document.styles if s.name and s.kind == kind][:50]) from None


def split_formatting(values: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    run = {k: values[k] for k in RUN_FIELDS if k in values}
    paragraph = {k: values[k] for k in PARAGRAPH_FIELDS if k in values and k != "indent_first"}
    if "indent_first" in values:
        first = values["indent_first"]
        if first >= 0:
            paragraph["first_line"] = first
        else:
            paragraph["hanging"] = -first
    if "highlight" in run and run["highlight"] == "none":
        run["highlight"] = None
    return run, paragraph


@word_tool("word_format", "The general setter: styles, run and paragraph formatting on each target. A range formats its characters; a paragraph or span its paragraphs.",
           {"doc": DOC,
            "targets": array(string(), "Paragraph, range (p:A@4:11), span (p:A..p:B) or "
                             "table ids.", min_items=1),
            "paragraph_style": string("Paragraph style name.", optional=True),
            "character_style": string("Character style name.", optional=True),
            **RUN_FIELDS, **PARAGRAPH_FIELDS,
            "clear_direct": boolean("First remove direct formatting, keeping styles. Default false.",
                                    optional=True)},
           refs=('targets[]',), group="word_style", mutates=True)
def word_format(call: Any, doc: str, targets: list[str], paragraph_style: str | None = None,
                character_style: str | None = None, clear_direct: bool = False, **values: Any) -> Result:
    document = call.document
    run, paragraph = split_formatting(values)
    if not (run or paragraph or paragraph_style or character_style or clear_direct):
        raise ToolError("invalid_arguments", "say what to set: a style, a run or a paragraph property",
                        valid_options=["paragraph_style", "character_style", *RUN_FIELDS, *PARAGRAPH_FIELDS,
                                       "clear_direct"])
    for name, kind in ((paragraph_style, "paragraph"), (character_style, "character")):
        if name is not None:
            known_style(document, name, kind, "paragraph_style" if kind == "paragraph" else "character_style")
    edits = []
    try:
        for address in resolve(call, targets):
            is_range = "@" in address and not address.startswith("p@")
            paragraphs = paragraphs_of(document, address)
            if clear_direct:
                if is_range:
                    edits.append(document.clear_formatting(document.range(address)))
                else:
                    edits.extend(document.clear_formatting(p.id) for p in paragraphs)
            if paragraph_style is not None:
                edits.extend(document.set_paragraph_style(p.id, paragraph_style) for p in paragraphs)
            if is_range:
                where = document.range(address)
                if character_style is not None:
                    edits.append(document.set_character_style(where, character_style))
                if run:
                    edits.append(document.format_range(where, **run))
                if paragraph:
                    edits.extend(document.format_paragraph(p.id, **paragraph) for p in paragraphs)
            else:
                if character_style is not None:
                    edits.extend(document.set_character_style(p.range(), character_style)
                                 for p in paragraphs if p.text)
                if run or paragraph:
                    edits.extend(document.format_paragraph(p.id, **paragraph, **run) for p in paragraphs)
    except StyleError as error:
        kind = "character" if character_style is not None and paragraph_style is None else None
        raise ToolError("not_found", str(error),
                        field="character_style" if kind else "paragraph_style",
                        valid_options=[s.name for s in document.styles if kind is None
                                       or getattr(s, "kind", None) == kind][:50]) from None
    return outcome(edits, f"Formatted {len(targets)} target(s)")


TOOLS = [word_set_text, word_insert_text, word_delete, word_insert_markdown, word_format]
