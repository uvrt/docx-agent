"""Text: ``word_set_text``, ``word_insert_text``, ``word_delete``, ``word_insert_markdown`` and
the general setter ``word_format``."""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import (Result, ToolError, array, boolean, integer, number, obj, string)

from ..edit.addresses import is_text_range, span_ends
from ..edit.document import Cell, EditError, EditResult, Paragraph, Table
from ..edit.ranges import TextRange
from ..edit.styles import StyleError
from ..markdown import model as m
from ..markdown.stylemap import StyleMap
from ..markdown.write import insert_blocks
from . import positions
from ._base import cached_layout, need, note_small_text, one_of, outcome, remember, resolve, short, word_tool
from .table_data import (DATE_FORMAT, FROM_COLUMNS, HEADER, MAX_COLUMNS, MAX_ROWS, fits, no_source,
                         placed)

DOC = string("Document id.")
REF = string("Ref name for $name.", optional=True)


# -- addresses ---------------------------------------------------------------------------------


def text_range(document: Any, address: str) -> TextRange:
    """A text range from an address: a range id (``p:A@4:11``, ``p:A@5``, ``p:A@3..p:B@8``),
    or the text a paragraph, cell (``t:X/c1,1``) or bookmark holds (``Document.range``)."""
    return document.range(address)


def find_range(call: Any, find: str, within: str | None) -> TextRange:
    return call.document.anchor(find, within=resolve(call, within) if within else None)


def paragraphs_of(document: Any, address: str) -> list[Paragraph]:
    """The paragraphs an address covers: one; a span's (``p:A..p:B``, ``p:A..``; tables'
    cells included); a range's; a table's, row's or cell's; a section's, bookmark's or
    story's -- what :meth:`~docx_agent.edit.addresses.AddressOps._scope_paragraphs` reads."""
    found = None if (span_ends(address) is not None or is_text_range(address)) else document.get(address)
    if isinstance(found, Paragraph):
        return [found]
    if isinstance(found, Table):
        return [p for row in range(len(found.rows)) for col in range(found.column_count)
                for p in _cell_paragraphs(found, row, col)]
    _, entries = document._scope_paragraphs(address)
    if not entries:
        raise ToolError("invalid_arguments", f"{address} holds no paragraphs", field="targets")
    return [document.paragraph(entry.id) for entry in entries]


def _cell_paragraphs(table: Table, row: int, column: int) -> list[Paragraph]:
    try:
        return list(table.cell(row, column).paragraphs)
    except IndexError:
        return []


# -- W5 word_set_text --------------------------------------------------------------------------


SET_ITEM = obj({"target": string("Paragraph, or table cell t:X/c<row>,<col>."),
                "text": string("The whole new text, plain; \\n starts a paragraph.")})


@word_tool("word_set_text", "Rewrite whole paragraphs, keeping the formatting; tracked, only changed words become revisions. For words within a paragraph, replace_text.",
           {"doc": DOC,
            "target": string("Paragraph, or table cell t:X/c<row>,<col>.", optional=True),
            "text": string("The whole new text, plain; \\n starts a paragraph.", optional=True),
            "items": array(SET_ITEM, "Or several paragraphs.", optional=True)},
           refs=('target', 'items[].target'), group="core", mutates=True,
           facts=("text", "items[].text"))
def word_set_text(call: Any, doc: str, target: str | None = None, text: str | None = None,
                  items: list[dict] | None = None) -> Result:
    """Round 28, W10: a ``\\n`` in the text (refused before) splits it into paragraphs, the
    later ones after the target in its style (``Document.set_text_paragraphs``); ``\\v`` is
    a line break inside one."""
    one_of({"target": target, "items": items}, "target", "items")
    if items == []:         # a trimmed batch's leftover: nothing to do, said, not refused
        return Result.nothing_given("written")
    work = items if items is not None else [{"target": target, "text": need(text, "text")}]
    edits = []
    for item in work:
        edits.extend(_set_one(call, resolve(call, item["target"]), item["text"]))
    made = sum(len(edit.created) for edit in edits)
    return outcome(edits, f"Set the text of {len(work)} paragraph(s)" + (f", {made} new after them" if made else ""))


def _set_one(call: Any, address: str, text: str) -> list[Any]:
    """One target's text: a paragraph's, or a cell's first paragraph's; with ``\\n``, every
    line a paragraph -- after the paragraph, or as the cell's paragraphs (the cell's others
    go, as word_edit_table's set_cells writes a cell's lines)."""
    from .objects import _set_cell

    document = call.document
    if span_ends(address) is not None:
        raise ToolError("invalid_arguments", f"{address} is a span: word_set_text writes one paragraph or "
                        "cell per target; give items, one per paragraph", field="target")
    split = "\n" in text or "\r" in text
    found = document.get(address)
    if isinstance(found, Cell) and split:
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if [p.text for p in found.paragraphs] == lines:
            return [EditResult(found.paragraphs[0].id, changed=False)]
        return _set_cell(document, found._table, found.row, found.column, "\n".join(lines))
    if not isinstance(found, Paragraph):
        paragraphs = getattr(found, "paragraphs", None)
        if not paragraphs or isinstance(found, Table):
            raise ToolError("invalid_arguments", f"{address} is a {type(found).__name__.lower()}: word_set_text "
                            "writes a paragraph or a cell (t:X/c<row>,<col>)", field="target")
        found = (paragraphs() if callable(paragraphs) else paragraphs)[0]
    if split:
        return [document.set_text_paragraphs(found.id, text)]
    return [found.set_text(text)]


# -- W6 word_insert_text -----------------------------------------------------------------------


INSERT_FIELDS = {
    "find": string("Text occurring once, to insert beside.", optional=True),
    "target": string("Or a range, position (p:A@5) or paragraph.", optional=True),
    "where": string("Default after.", enum=["before", "after"], optional=True),
    # Round 40: the place as every placing tool names it (positions.py).
    "at": string("Or after:/before: a range, or a position.", optional=True),
    "text": string("Plain text."),
    "ref": REF,
}

_INSERT_TEXT = {"doc": DOC, **{k: (v if k != "text" else string("Plain text.", optional=True))
                               for k, v in INSERT_FIELDS.items()},
                "within": string("Block or span to search for find.", optional=True),
                "items": array(obj(INSERT_FIELDS), "Or several insertions.", optional=True)}


@word_tool("word_insert_text", "Insert text beside words found once, or at a range or position, in the formatting there. Returns the new ranges' ids.",
           _INSERT_TEXT,
           refs=('target', 'within', 'items[].target'), group="word_text", mutates=True,
           facts=("text", "items[].text"),
           aliases={**positions.aliases("word_insert_text", "text", _INSERT_TEXT, find=True),
                    **positions.aliases("word_insert_text", "text", INSERT_FIELDS, prefix="items[].",
                                        find=True)})
def word_insert_text(call: Any, doc: str, items: list[dict] | None = None, **single: Any) -> Result:
    if items is not None and any(single.get(k) is not None for k in ("find", "target", "at", "text")):
        raise ToolError("invalid_arguments", "give one insertion's fields or items, not both", field="items")
    if items == []:
        return Result.nothing_given("inserted")
    work = items if items is not None else [single]
    edits, made = [], []
    for k, item in enumerate(work):
        need(item.get("text"), "text")
        if all(item.get(name) is None for name in ("find", "target", "at")) and positions.is_place(item.get("ref")):
            raise _ref_as_place(call, "word_insert_text", item if items is not None else {"doc": doc, **item},
                                k if items is not None else None)
        key = one_of(item, "find", "target", "at")
        where = item.get("where") or "after"
        if key == "find":
            found = find_range(call, item["find"], item.get("within"))
        elif key == "target":
            found = text_range(call.document, resolve(call, item["target"]))
        else:
            where, address = _text_at(item["at"], item.get("where"))
            found = text_range(call.document, resolve(call, address))
        place = found.id
        edit = found.insert_before(item["text"]) if where == "before" else found.insert_after(item["text"])
        edits.append(edit)
        inserted = inserted_range(place, where, item["text"], edit.renamed)
        made.append(inserted)
        remember(call, item.get("ref"), inserted)
    result = outcome(edits, f"Inserted text at {len(edits)} place(s)", data={"ranges": made})
    return result


def _text_at(at: str, where: str | None) -> tuple[str, str]:
    """``word_insert_text``'s ``at``: ``(side, range)`` -- ``after:`` or ``before:`` a range,
    or a bare range or position, after it unless ``where`` says before."""
    side = positions.side_of(at)
    if side is None or side[0] == "end" or at.startswith("replace:"):
        raise ToolError("invalid_arguments", f"at {at!r}: text goes beside a range: after:<range>, "
                        "before:<range>, or a position p:<id>@<offset>", field="at")
    if at.startswith(("after:", "before:")):
        if where is not None and where != side[0]:
            raise ToolError("invalid_arguments", f"at says {side[0]} and where says {where}: give one",
                            field="where", valid_options=[side[0], where])
        return side
    return where or "after", side[1]


def _ref_as_place(call: Any, tool: str, item: dict, index: int | None) -> ToolError:
    """A ``ref`` holding an address where the call gives no place: ``ref`` names what the
    call makes ($name); the place is ``at`` (round 40)."""
    meant = {k: v for k, v in item.items() if k != "ref"}
    meant["at"] = item["ref"]
    details = {"item": meant} if index is not None else {"call": {"tool": tool, "arguments": meant}}
    return ToolError("invalid_arguments", f"ref {item['ref']!r} is an address: ref names what this call makes, "
                     f"for $name later; the place is at: write at: {item['ref']!r}",
                     field="ref" if index is None else f"items[{index}].ref", valid_options=["at"],
                     details=details)


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
            "target": string("Range, block (a control cc: too), span p:A..t:B, picture d:, note fn:/en:.",
                             optional=True),
            "collapse_space": boolean("Also delete the space the deletion would leave doubled. "
                                      "Default false.", optional=True),
            "force": boolean("Delete a heading though its content stays. Default false.", optional=True)},
           refs=('target', 'within'), group="word_text", mutates=True)
def word_delete(call: Any, doc: str, find: str | None = None, within: str | None = None,
                target: str | None = None, collapse_space: bool = False, force: bool = False) -> Result:
    """Round 36, R33: a heading deleted while blocks under it stay is refused -- they would go
    under the heading before it -- unless ``force``; ``h:<heading>`` deletes it with them."""
    document = call.document
    if find is None and target is None:
        raise ToolError("invalid_arguments", "give find (text to delete) or target (what to delete)",
                        field="find", valid_options=["find", "target"])
    note = None
    if find is not None and target is not None:
        # Round 31, R5a: find with a target is find within it (the call meant), said in a note.
        if within is not None and resolve(call, within) != resolve(call, target):
            raise ToolError("invalid_arguments", "find searches within one block or span: give within or "
                            "target with it, not both", field="target", valid_options=["within", "target"])
        within, note = target, f"find was searched within target {target}: within names where to search"
    if find is not None:
        found = find_range(call, find, within)
        text = found.text
        result = outcome([found.delete(collapse_space=collapse_space)], f"Deleted {short(text, 60)!r}")
        if note is not None:
            result.warnings.append(note)
        return result
    target = resolve(call, target)
    if "@" in target and not target.startswith(("p@", "cc@")):
        return outcome([document.delete_text(target, collapse_space=collapse_space)], f"Deleted {target}")
    if target.startswith(("fn:", "en:")):
        return outcome([document.delete_note(target)], f"Deleted note {target}")
    if target.startswith("d:"):
        return outcome([document.picture(target).delete()], f"Deleted picture {target}")
    if ".." in target:
        elements, part = document._blocks_of(target)
        ids = [document._block_id(part, element) for element in elements]
        moved = _headings_kept(document, ids, force)
        result = outcome([document.delete_block(i) for i in ids], f"Deleted {len(ids)} block(s)")
    else:
        moved = _headings_kept(document, [target], force)
        result = outcome([document.delete_block(target)], f"Deleted {target}")
    result.warnings += moved
    return result


def _headings_kept(document: Any, ids: list[str], force: bool) -> list[str]:
    """Round 36, R33: the refusal of deleting a heading whose content stays (it would go under
    the heading before), with the ways out: the heading's range ``h:``, or ``force`` -- the
    warnings saying where it went, when forced."""
    from ..edit.outline import content_left, label

    try:
        canonical = [document._resolve_block(i)[1].id for i in ids]
    except (KeyError, EditError, AttributeError):
        return []                          # delete_block refuses it, in its own words
    left = content_left(document, canonical)
    if not left:
        return []
    lines = [f"{label(found.heading)} is a heading: its {found.kept} block(s) from {found.first} (to "
             f"{found.heading.last}) would go under {label(found.under)}" for found in left]
    if force:
        return [line.replace("would go", "now go") for line in lines]
    raise ToolError("invalid_arguments", "; ".join(lines[:5]) + (f"; and {len(lines) - 5} more" if len(lines) > 5 else "")
                    + ". To delete a heading with everything under it, target its range "
                    + ", ".join(found.heading.range for found in left[:3])
                    + "; to delete the heading alone and keep its content, force true",
                    field="target", valid_options=[found.heading.range for found in left][:50] + ["force"])


# -- W8 word_insert_markdown -------------------------------------------------------------------


STYLE_KEYS = ["paragraph", "h1", "h2", "h3", "h4", "h5", "h6", "quote", "code", "inline_code",
              "bullet", "bullet2", "bullet3", "number", "number2", "number3", "table", "table_cell",
              "table_cell_center", "table_cell_right", "footnote", "footnote_reference", "emphasis", "strong", "link",
              "caption"]

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


#: Where a placing tool puts new blocks (round 40: the form every placing tool reads).
AT = string("end, after:<id>, before:<id> or replace:<id>[..<id>].")
#: A ref for the blocks a call makes: one block's id, or several's span.
BLOCKS_REF = string("Ref name for the new blocks ($name).", optional=True)

_INSERT_MARKDOWN = {"doc": DOC,
                    "markdown": string("Markdown.", optional=True),
                    "blob": string("Or a Markdown blob handle.", optional=True),
                    "from": string("Or a table of the rows of d:<id> or t:<table>.", optional=True),
                    "from_columns": FROM_COLUMNS,
                    "header": HEADER,
                    "date_format": DATE_FORMAT,
                    "at": AT,
                    "style_map": STYLE_MAP,
                    "key": string("Retry key: a repeat with the same key makes nothing new.", optional=True),
                    "ref": BLOCKS_REF}


@word_tool("word_insert_markdown", "Write Markdown as blocks in the document's styles: headings, lists, tables (empty header row: none), links, footnotes, images ![alt](b1), captions (: text). Returns the new blocks' ids.",
           _INSERT_MARKDOWN,
           group="word_text", mutates=True, exactly_one=[("markdown", "blob", "from")],
           sources=("from",), facts=("markdown",),
           aliases={"columns": "from_columns",
                    **positions.aliases("word_insert_markdown", "blocks", _INSERT_MARKDOWN)})
def word_insert_markdown(call: Any, doc: str, at: str, markdown: str | None = None,
                         blob: str | None = None, style_map: list[dict] | None = None,
                         key: str | None = None, from_columns: list[str] | None = None,
                         header: list[str] | None = None, date_format: str | None = None,
                         ref: str | None = None, **given: Any) -> Result:
    """``ref`` (round 40) names the blocks made, as a creating call's ``ref`` does: the one
    block's id, or the span from the first to the last."""
    named(call, "word_insert_markdown", ref, at)
    source = given.get("from")
    if source is not None:
        return name_blocks(call, ref, _table_from(call, source, from_columns, header, _at(call, at),
                                                  style_map, date_format))
    no_source(from_columns, header, date_format)
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
    return name_blocks(call, ref, note_small_text(call, result, edit.blocks))


def named(call: Any, tool: str, ref: str | None, at: str) -> None:
    """Refuse a ``ref`` that holds an address: it names the blocks made; the place is ``at``."""
    if ref is not None and positions.is_place(ref):
        raise ToolError("invalid_arguments", f"ref {ref!r} is an address: ref names the blocks this call makes, "
                        f"for $name later; the place is at (now {at!r}): write at: 'after:{ref}' to place them "
                        "after it", field="ref", valid_options=["at"])


def name_blocks(call: Any, ref: str | None, result: Result) -> Result:
    """``ref`` naming the blocks a call made (``data.blocks``): one's id, or the span of several."""
    blocks = (result.data or {}).get("blocks") if isinstance(result.data, dict) else None
    if ref and blocks:
        remember(call, ref, blocks[0] if len(blocks) == 1 else f"{blocks[0]}..{blocks[-1]}")
    return result


def _table_from(call: Any, source: str, from_columns: list[str] | None,
                header: list[str] | None, at: str, style_map: list[dict] | None,
                date_format: str | None = None) -> Result:
    """A new table of a source's rows (:mod:`.table_data`), written as a Markdown table is."""
    from .objects import table_json

    block = placed(call, source, from_columns, header, date_format)
    fits(block, MAX_ROWS, MAX_COLUMNS)
    cells = [[_inline(lines) for lines in row] for row in block.lines()]
    model = m.Table(cells[0] if block.header else [], cells[1:] if block.header else cells,
                    [None] * block.size[1], headed=bool(block.header))
    document = call.document
    try:
        edit = insert_blocks(document, [model], at=at, style_map=_style_map(style_map))
    except StyleError as error:
        raise ToolError("not_found", str(error), field="style_map",
                        valid_options=[s.name for s in document.styles][:50]) from None
    table = next(b for b in edit.blocks if b.startswith("t:"))
    rows, columns = block.size
    data = {"blocks": edit.blocks, **table_json(document, table, markdown=False), "from": block.provenance(0)}
    call.keep_provenance(table, block.rows, data["from"])   # kept in the file, this undo step
    return note_small_text(call, outcome([edit], f"Inserted a {rows}x{columns} table {table} from "
                                          f"{block.rows.address}", data=data), [table])


def _inline(lines: list[str]) -> list:
    """A cell's lines as literal text, hard breaks between them."""
    out: list = []
    for k, line in enumerate(lines):
        if k:
            out.append(m.Break())
        if line:
            out.append(m.Text(line))
    return out


def _style_map(entries: list[dict] | None) -> Any:
    mapping = style_map(entries)
    return StyleMap.from_dict(mapping) if mapping else None


def _at(call: Any, at: str) -> str:
    """``at`` with any ``$ref`` inside resolved: ``after:$intro``.  A ref naming several
    blocks (a span, as ``word_insert_markdown``'s ``ref`` names what it made) is after its
    last block, before its first, or replaced whole."""
    if "$" not in at:
        return at
    head, _, rest = at.partition(":") if at.split(":", 1)[0] in ("after", "before", "replace") else ("", "", at)
    parts = [resolve(call, part) if part.startswith("$") else part for part in rest.split("..")]
    joined = "..".join(parts)
    ends = span_ends(joined) if ".." in joined else None
    if ends is not None and head != "replace":
        joined = (ends[0] if head == "before" else ends[1]) or joined
    return (head + ":" if head else "") + joined


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
    "keep_lines": boolean("Keep lines together.", optional=True),
    "page_break_before": boolean("Page break before.", optional=True),
}


#: A paragraph's spacing named by its side, as a model writes it (``top: 6``; round 6, R7,
#: as pptx-agent's ``ppt_format_text`` reads it): ``space_before`` and ``space_after``.
#: ``word_format`` has no other top or bottom, so the meaning is the one.
SPACING_ALIASES = {"top": "space_before", "bottom": "space_after"}

#: word_format's names the Python API spells otherwise: ``keep_lines`` (Word's "Keep lines
#: together", ``w:keepLines``; round 25, W4) is :meth:`Paragraph.format`'s ``keep_together``.
_API_NAMES = {"keep_lines": "keep_together"}


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
    paragraph = {_API_NAMES.get(k, k): values[k] for k in PARAGRAPH_FIELDS if k in values and k != "indent_first"}
    if "indent_first" in values:
        first = values["indent_first"]
        if first >= 0:
            paragraph["first_line"] = first
        else:
            paragraph["hanging"] = -first
    if "highlight" in run and run["highlight"] == "none":
        run["highlight"] = None
    return run, paragraph


#: Round 40: text to format found by the call, as replace_text finds it (``find[]``).
FIND_ITEM = obj({"text": string("Text to format."),
                 "within": string("Block or span to search.", optional=True),
                 "occurrence": integer("Which match, from 1.", minimum=1, optional=True),
                 "expect": string("Default one: exactly one match.", enum=["one", "all"], optional=True)})


@word_tool("word_format", "The general setter: styles, run and paragraph formatting on each target. A range formats its characters; a paragraph or span its paragraphs.",
           {"doc": DOC,
            "targets": array(string(), "Paragraph, range (p:A@4:11), span (p:A..p:B) or "
                             "table ids.", or_single=True, optional=True),
            "find": array(FIND_ITEM, "Or text to find and format.", or_single=True, optional=True),
            "paragraph_style": string("Paragraph style name.", optional=True),
            "character_style": string("Character style name.", optional=True),
            **RUN_FIELDS, **PARAGRAPH_FIELDS,
            # Round 36, R30: a span's blocks on one page (Document.keep_blocks_together).
            "keep_together": boolean("Keep each target whole on one page: a span's blocks, a "
                                     "table, a paragraph's lines.", optional=True),
            "clear_direct": boolean("First remove direct formatting, keeping styles. Default false.",
                                    optional=True),
            # Round 16, F8: what check's direct_formatting fix clears -- only these.
            "clear": array(string(enum=[*RUN_FIELDS, *PARAGRAPH_FIELDS]),
                           "Direct properties to remove, back to the style's.", optional=True)},
           refs=('targets[]', 'find[].within'), group="word_style", mutates=True,
           aliases={**SPACING_ALIASES, "find[].range": "within"})
def word_format(call: Any, doc: str, targets: list[str] | None = None, paragraph_style: str | None = None,
                character_style: str | None = None, clear_direct: bool = False,
                clear: list[str] | None = None, keep_together: bool | None = None,
                find: list[dict] | None = None, **values: Any) -> Result:
    """``keep_together`` (round 36, R30): a paragraph -- one alone, or a range's or a cell's
    -- keeps its lines together (``keep_lines``, the Python API's ``keep_together``); a span,
    a table, a section or a block-level control keeps its blocks on one page
    (:meth:`~docx_agent.Document.keep_blocks_together`), with a warning where the last layout
    shows them taller than a page.

    ``find`` (round 40): text found by the call, each item's ranges formatted as a range
    target is (:func:`found_ranges`): every one is found before anything changes, so a text
    that is not there, or matches several without ``occurrence``, refuses the call whole."""
    document = call.document
    if not targets and not find:
        if targets is None and find is None:
            raise ToolError("invalid_arguments", "say what to format: targets (ids) or find (text)",
                            field="targets", valid_options=["targets", "find"])
        return Result.nothing_given("formatted", field="targets")   # a trimmed batch's leftover
    both = [name for name in clear or () if name in values]
    if both:
        raise ToolError("invalid_arguments", f"{', '.join(both)} both set and cleared: give a value or "
                        "clear it, not both", field="clear")
    run, paragraph = split_formatting(values)
    for name in clear or ():
        if name == "indent_first":
            paragraph.update(first_line=None, hanging=None)
        elif name in RUN_FIELDS:
            run[name] = None
        else:
            paragraph[_API_NAMES.get(name, name)] = None
    if not (run or paragraph or paragraph_style or character_style or clear_direct or keep_together is not None):
        raise ToolError("invalid_arguments", "say what to set: a style, a run or a paragraph property",
                        valid_options=["paragraph_style", "character_style", *RUN_FIELDS, *PARAGRAPH_FIELDS,
                                       "keep_together", "clear_direct", "clear"])
    for name, kind in ((paragraph_style, "paragraph"), (character_style, "character")):
        if name is not None:
            known_style(document, name, kind, "paragraph_style" if kind == "paragraph" else "character_style")
    edits = []
    found = found_ranges(call, find or [])
    addresses = resolve(call, list(targets or [])) + found
    # The blocks kept together, and the layout they are measured in before the change.
    kept = [] if keep_together is None else [a for a in addresses if _keeps_blocks(document, a)]
    before = cached_layout(document) if kept and keep_together else None
    # Round 36, R33: a style or a cleared outline level may change a heading's level.
    from ..edit.outline import headings, moves

    outline = headings(document) if paragraph_style is not None or clear_direct or clear else None
    try:
        for address in addresses:
            is_range = "@" in address and not address.startswith("p@")
            paragraphs = paragraphs_of(document, address)
            if keep_together is not None:
                if address in kept:
                    edits.append(_keep_blocks(document, address, keep_together, before))
                else:
                    edits.extend(document.format_paragraph(p.id, keep_together=keep_together)
                                 for p in paragraphs)
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
    result = outcome(edits, f"Formatted {len(addresses)} target(s)",
                     data={"ranges": found[:50]} if found else None)
    if outline is not None:
        renamed: dict[str, str] = {}
        for edit in edits:
            renamed.update(getattr(edit, "renamed", None) or {})
        result.warnings += moves(outline, headings(document), renamed)
    return result


def found_ranges(call: Any, items: list[dict]) -> list[str]:
    """The ranges ``word_format``'s ``find`` items name (round 40): each item's text, in its
    ``within`` (the body by default, as ``replace_text`` searches), once -- or its
    ``occurrence`` (from 1), or each match with ``expect`` ``all``.  A text found nowhere is
    refused with the places it occurs written otherwise (case, spaces) and what ``within``
    holds; one found more than once with each match, its range and its ``occurrence``."""
    from .shared import _story_ok

    document = call.document
    out: list[str] = []
    for k, item in enumerate(items):
        text, where = item["text"], f"find[{k}]"
        within = resolve(call, item["within"]) if item.get("within") else None
        occurrence, expect = item.get("occurrence"), item.get("expect") or "one"
        if occurrence is not None and expect == "all":
            raise ToolError("invalid_arguments", f"{where}: occurrence picks one match and expect all takes each: "
                            "give one of them", field=f"{where}.occurrence", valid_options=["occurrence", "expect"])
        matches = [m for m in _found(document, text, within) if within is not None or _story_ok(m.id, "body")]
        place = f" in {within}" if within else ""
        if not matches:
            raise _not_there(document, text, within, where)
        if occurrence is not None:
            if occurrence > len(matches):
                raise ToolError("not_found", f"{where}: {text!r} occurs {len(matches)} time(s){place}; no "
                                f"occurrence {occurrence}", field=f"{where}.occurrence",
                                valid_options=_listed(matches))
            out.append(matches[occurrence - 1].id)
        elif expect == "one" and len(matches) > 1:
            raise ToolError("ambiguous", f"{where}: {len(matches)} places match {text!r}{place}: give occurrence "
                            f"(1-{len(matches)}, in this order), a narrower within, or expect all",
                            field=f"{where}.occurrence", valid_options=_listed(matches))
        else:
            out.extend(m.id for m in matches)
    return out


def _found(document: Any, text: str, within: str | None, **options: Any) -> list[Any]:
    try:
        return document.find(text, within=within, **options)
    except (KeyError, ValueError) as error:
        raise ToolError("not_found", f"{within}: {error}", field="find") from None


def _listed(matches: list[Any]) -> list[str]:
    """Each match as ``occurrence n: <range> <context>``: the call can name it either way."""
    return [f"occurrence {n}: {m.id} {m.context()!r}" for n, m in enumerate(matches[:50], start=1)]


def _not_there(document: Any, text: str, within: str | None, where: str) -> ToolError:
    """The refusal of a text found nowhere: where it occurs ignoring case and spaces, and
    (within one paragraph or cell) the text that is there."""
    from .shared import _story_ok

    place = f" in {within}" if within else ""
    near = [m for m in _found(document, text, within, case=False, whitespace="collapse")
            if within is not None or _story_ok(m.id, "body")]
    message = f"{where}: {text!r} occurs nowhere{place}"
    options = []
    if near:
        message += f"; written otherwise at {len(near)} place(s): give the text as written there"
        options = [f"{m.id} {m.text!r}" for m in near[:50]]
    elif within is not None:
        try:
            held = document.range(within).text
        except Exception:  # noqa: BLE001 -- a block that holds no one text: say nothing of it
            held = None
        if held is not None:
            message += f"; it holds {short(held, 160)!r}"
    return ToolError("not_found", message, field=f"{where}.text", valid_options=options)


def _keeps_blocks(document: Any, address: str) -> bool:
    """Whether ``keep_together`` keeps ``address``'s blocks on one page (a span, a table, a
    section, a bookmark, a block-level control), not one paragraph's lines (a paragraph, a
    text range, a cell or a row: its paragraphs')."""
    if is_text_range(address) or "/" in address:
        return False
    if span_ends(address) is not None:
        return True
    try:
        return not isinstance(document.get(address), Paragraph)
    except (KeyError, ValueError):
        return True                       # a section, a bookmark: blocks_of reads them


def _keep_blocks(document: Any, address: str, on: bool, layout: Any) -> EditResult:
    try:
        result = document.keep_blocks_together(address, on)
    except EditError as error:
        raise ToolError("invalid_arguments", f"{address}: {error}; keep_together keeps the body's blocks: give "
                        "a span p:A..t:B, a table or a section", field="targets") from None
    taller = _taller_than_page(document, address, layout) if on and layout is not None else None
    if taller is not None:
        result.warnings.append(
            f"{address} is {taller[0]} pt tall and a page holds {taller[1]} pt: it cannot be kept on one page "
            "(Word starts it on a new page and breaks it where the page is full); keep shorter spans together")
    return result


def _taller_than_page(document: Any, address: str, layout: Any) -> tuple[float, float] | None:
    """``(height, room)`` in points where ``layout`` draws ``address``'s paragraphs taller
    than the body of its pages holds (:meth:`~docx_agent.layout.DocumentLayout.extent`)."""
    try:
        ids = [p.id for p in paragraphs_of(document, address)]
        extents = layout.extent(ids)
    except (KeyError, ValueError, ToolError):
        return None
    if not extents:
        return None
    height = round(sum(e.bottom - e.top for e in extents), 1)
    room = round(min(e.area[1] - e.area[0] for e in extents), 1)
    return (height, room) if height > room + 0.5 else None


TOOLS = [word_set_text, word_insert_text, word_delete, word_insert_markdown, word_format]
