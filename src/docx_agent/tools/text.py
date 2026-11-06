"""Text: ``word_set_text``, ``word_insert_text``, ``word_delete``, ``word_insert_markdown`` and
the general setter ``word_format``."""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import (Result, ToolError, array, boolean, number, obj, string)

from ..edit.addresses import is_text_range, span_ends
from ..edit.document import Cell, EditError, EditResult, Paragraph, Table
from ..edit.ranges import TextRange
from ..edit.styles import StyleError
from ..markdown import model as m
from ..markdown.stylemap import StyleMap
from ..markdown.write import insert_blocks
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
    "text": string("Plain text."),
    "ref": REF,
}


@word_tool("word_insert_text", "Insert text beside words found once, or at a range or position, in the formatting there. Returns the new ranges' ids.",
           {"doc": DOC, **{k: (v if k != "text" else string("Plain text.", optional=True))
                           for k, v in INSERT_FIELDS.items()},
            "within": string("Block or span to search for find.", optional=True),
            "items": array(obj(INSERT_FIELDS), "Or several insertions.",
                           optional=True)},
           refs=('target', 'within', 'items[].target'), group="word_text", mutates=True,
           facts=("text", "items[].text"))
def word_insert_text(call: Any, doc: str, items: list[dict] | None = None, **single: Any) -> Result:
    if items is not None and any(single.get(k) is not None for k in ("find", "target", "text")):
        raise ToolError("invalid_arguments", "give one insertion's fields or items, not both", field="items")
    if items == []:
        return Result.nothing_given("inserted")
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
            "target": string("Range, block (a control cc: too), span p:A..t:B, picture d:, note fn:/en:.",
                             optional=True),
            "collapse_space": boolean("Also delete the space the deletion would leave doubled. "
                                      "Default false.", optional=True)},
           refs=('target', 'within'), group="word_text", mutates=True)
def word_delete(call: Any, doc: str, find: str | None = None, within: str | None = None,
                target: str | None = None, collapse_space: bool = False) -> Result:
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
        return outcome([document.delete_block(i) for i in ids], f"Deleted {len(ids)} block(s)")
    return outcome([document.delete_block(target)], f"Deleted {target}")


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


@word_tool("word_insert_markdown", "Write Markdown as blocks in the document's styles: headings, lists, tables (empty header row: none), links, footnotes, images ![alt](b1), captions (: text). Returns the new blocks' ids.",
           {"doc": DOC,
            "markdown": string("Markdown.", optional=True),
            "blob": string("Or a Markdown blob handle.", optional=True),
            "from": string("Or a table of the rows of d:<id> or t:<table>.", optional=True),
            "from_columns": FROM_COLUMNS,
            "header": HEADER,
            "date_format": DATE_FORMAT,
            "at": string("end, after:<id>, before:<id> or replace:<id>[..<id>]."),
            "style_map": STYLE_MAP,
            "key": string("Retry key: a repeat with the same key makes nothing new.", optional=True)},
           group="word_text", mutates=True, exactly_one=[("markdown", "blob", "from")],
           sources=("from",), facts=("markdown",), aliases={"columns": "from_columns"})
def word_insert_markdown(call: Any, doc: str, at: str, markdown: str | None = None,
                         blob: str | None = None, style_map: list[dict] | None = None,
                         key: str | None = None, from_columns: list[str] | None = None,
                         header: list[str] | None = None, date_format: str | None = None,
                         **given: Any) -> Result:
    source = given.get("from")
    if source is not None:
        return _table_from(call, source, from_columns, header, _at(call, at), style_map,
                           date_format)
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
    return note_small_text(call, result, edit.blocks)


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


@word_tool("word_format", "The general setter: styles, run and paragraph formatting on each target. A range formats its characters; a paragraph or span its paragraphs.",
           {"doc": DOC,
            "targets": array(string(), "Paragraph, range (p:A@4:11), span (p:A..p:B) or "
                             "table ids.", min_items=1),
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
           refs=('targets[]',), group="word_style", mutates=True, aliases=SPACING_ALIASES)
def word_format(call: Any, doc: str, targets: list[str], paragraph_style: str | None = None,
                character_style: str | None = None, clear_direct: bool = False,
                clear: list[str] | None = None, keep_together: bool | None = None, **values: Any) -> Result:
    """``keep_together`` (round 36, R30): a paragraph -- one alone, or a range's or a cell's
    -- keeps its lines together (``keep_lines``, the Python API's ``keep_together``); a span,
    a table, a section or a block-level control keeps its blocks on one page
    (:meth:`~docx_agent.Document.keep_blocks_together`), with a warning where the last layout
    shows them taller than a page."""
    document = call.document
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
    addresses = resolve(call, targets)
    # The blocks kept together, and the layout they are measured in before the change.
    kept = [] if keep_together is None else [a for a in addresses if _keeps_blocks(document, a)]
    before = cached_layout(document) if kept and keep_together else None
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
    return outcome(edits, f"Formatted {len(targets)} target(s)")


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
