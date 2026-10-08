"""Objects: tables (``word_edit_table``, ``word_format_table``; a new table is Markdown,
through ``word_insert_markdown``),
pictures and other drawings (``word_drawings``, which inserts a picture too) and content
controls (``word_controls``)."""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, array, integer, number, obj, string

from ._base import need, outcome, remember, resolve, short, word_tool
from .structure import _position

DOC = string("Document id.")
TABLE = string("Table (t:...).")


def table_json(document: Any, table_id: str) -> dict[str, Any]:
    table = document.table(table_id)
    return {"id": table.id, "rows": len(table.rows), "columns": table.column_count,
            "markdown": document.to_markdown(table.id, ids=False)}


def _limit_cells(call: Any, count: int) -> None:
    if count > 5000:
        raise ToolError("limit", f"{count} cells in one call; at most 5000")


# -- W23 word_edit_table -----------------------------------------------------------------------


CELL = obj({"row": integer(minimum=0, optional=True),
            "col": integer(minimum=0, optional=True),
            "row_label": string("Or the row's first cell text.", optional=True),
            "col_label": string("Or the column's header text.", optional=True),
            "text": string("Lines are paragraphs.")})
POSITION = obj({"row": integer(minimum=0), "col": integer(minimum=0)},
               "merge, split: first cell, from 0.", optional=True)


def _locate(table: Any, cell: dict, header_row: int = 0) -> tuple[int, int]:
    def labelled(texts: list[str], label: str, what: str) -> int:
        clean = [" ".join(t.split()) for t in texts]
        hits = [k for k, t in enumerate(clean) if t == " ".join(label.split())]
        if len(hits) != 1:
            raise ToolError("not_found" if not hits else "ambiguous", f"{what} label {label!r} matches "
                            f"{len(hits)} {what}s", valid_options=[t for t in clean if t][:50])
        return hits[0]

    if cell.get("row") is not None:
        row = cell["row"]
    else:
        row = labelled([_text(table, r, 0) for r in range(len(table.rows))],
                       need(cell.get("row_label"), "cells[].row or row_label"), "row")
    if cell.get("col") is not None:
        col = cell["col"]
    else:
        col = labelled([_text(table, header_row, c) for c in range(table.column_count)],
                       need(cell.get("col_label"), "cells[].col or col_label"), "column")
    return row, col


def _text(table: Any, row: int, col: int) -> str:
    try:
        return table.cell(row, col).text
    except IndexError:
        return ""


@word_tool("word_edit_table", "Change a table's cells and structure: set cells by position or labels, insert or delete rows and columns, merge, split. Returns it as Markdown.",
           {"doc": DOC, "table": TABLE,
            "action": string("What to do.", enum=["set_cells", "insert_row", "delete_row", "insert_column",
                                                   "delete_column", "merge", "split"]),
            "cells": array(CELL, "set_cells: the cells.", optional=True),
            "index": integer("insert_*, delete_*: row or column from 0. Default the last.",
                             minimum=0, optional=True),
            "side": string("insert_*: relative to index. Default after.", enum=["after", "before"], optional=True),
            "first": POSITION, "last": POSITION,
            "rows": integer("split: into rows. Default 1.", minimum=1, maximum=63, optional=True),
            "columns": integer("split: into columns. Default 2.", minimum=1, maximum=63, optional=True)},
           refs=('table',), group="word_objects", mutates=True)
def word_edit_table(call: Any, doc: str, table: str, action: str, cells: list[dict] | None = None,
                    index: int | None = None, side: str = "after", first: dict | None = None,
                    last: dict | None = None, rows: int = 1, columns: int = 2) -> Result:
    document = call.document
    table = resolve(call, table)
    view = document.table(table)
    edits = []
    if action == "set_cells":
        cells = need(cells, "cells")
        _limit_cells(call, len(cells))
        for cell in cells:  # in order: a label set by an earlier cell is found by a later one
            current = document.table(table)
            row, col = _locate(current, cell)
            edits.extend(_set_cell(document, current, row, col, cell["text"]))
    elif action == "insert_row":
        edits.append(document.insert_row(table, index, below=side == "after"))
    elif action == "delete_row":
        edits.append(document.delete_row(table, need(index, "index")))
    elif action == "insert_column":
        col = view.column_count - 1 if index is None else index
        edits.append(document.insert_column(table, col, right=side == "after"))
    elif action == "delete_column":
        edits.append(document.delete_column(table, need(index, "index")))
    elif action == "merge":
        a, b = need(first, "first"), need(last, "last")
        edits.append(document.merge_cells(table, (a["row"], a["col"]), (b["row"], b["col"])))
    else:
        a = need(first, "first", " (the cell to split)")
        edits.append(document.split_cell(table, a["row"], a["col"], rows=rows, columns=columns))
    return outcome(edits, f"{action} on {table}", data=table_json(document, table), changed=[table])


def _set_cell(document: Any, table: Any, row: int, col: int, text: str) -> list[Any]:
    """A cell's text: its first paragraph rewritten (formatting kept), lines past the first
    as paragraphs after it, the cell's other paragraphs deleted."""
    cell = table.cell(row, col)
    paragraphs = list(cell.paragraphs)
    lines = text.split("\n")
    edits = []
    if paragraphs[0].text != lines[0]:
        edits.append(paragraphs[0].set_text(lines[0]))
    for extra in paragraphs[1:]:
        edits.append(document.delete_block(extra.id))
    last = document.table(table.id).cell(row, col).paragraphs[0].id
    for line in lines[1:]:
        edit = document.insert_paragraph(line, after=last)
        edits.append(edit)
        last = edit.id
    return edits


# -- W24 word_format_table ---------------------------------------------------------------------


BORDER_STYLES = ["none", "single", "double", "thick", "dotted", "dashed"]
BORDERS = string("Every border's style, inside ones too.", enum=BORDER_STYLES, optional=True)
_TABLE_SIDES = ("top", "bottom", "left", "right", "inside_h", "inside_v")


@word_tool("word_format_table", "Format a table, row, column or cell: style, width, alignment, shading, borders, vertical alignment, row height, header rows.",
           {"doc": DOC, "table": TABLE,
            "scope": string("What to format.", enum=["table", "row", "column", "cell"]),
            "row": integer("row, cell: from 0.", minimum=0, optional=True),
            "column": integer("column, cell: from 0.", minimum=0, optional=True),
            "style": string("table: style name.", optional=True),
            "width": number("table, column, cell: width.", minimum=1, maximum=1584, optional=True),
            "width_percent": number("table: width, % of the text width.", minimum=1, maximum=100,
                                    optional=True),
            "alignment": string("table: on the page.", enum=["left", "center", "right"], optional=True),
            "shading": string("Fill RRGGBB, or none.", optional=True),
            "borders": BORDERS,
            "vertical_alignment": string("cell, row, column: text position.", enum=["top", "center", "bottom"],
                                         optional=True),
            "height": number("row: height.", minimum=1, maximum=1584, optional=True),
            "height_rule": string("row. Default at_least.", enum=["at_least", "exact"], optional=True),
            "header_rows": integer("table: rows repeated on each page.", minimum=0, maximum=500,
                                   optional=True)},
           refs=('table',), group="word_objects", mutates=True)
def word_format_table(call: Any, doc: str, table: str, scope: str, row: int | None = None,
                      column: int | None = None, style: str | None = None, width: float | None = None,
                      width_percent: float | None = None, alignment: str | None = None,
                      shading: str | None = None, borders: str | None = None,
                      vertical_alignment: str | None = None, height: float | None = None,
                      height_rule: str = "at_least", header_rows: int | None = None) -> Result:
    document = call.document
    table = resolve(call, table)
    view = document.table(table)
    fill = None if shading == "none" else shading
    style = None if borders in (None, "none") else borders
    sides = {side: style for side in _TABLE_SIDES} if borders is not None else {}
    edits = []
    if scope == "table":
        values: dict[str, Any] = {}
        for key, value in (("style", style), ("alignment", alignment), ("header_rows", header_rows)):
            if value is not None:
                values[key] = value
        if width is not None or width_percent is not None:
            values["width"] = f"{width_percent:g}%" if width_percent is not None else width
        if shading is not None:
            values["shading"] = fill
        if borders:
            values["borders"] = sides
        if not values:
            raise ToolError("invalid_arguments", "nothing to set for the table",
                            valid_options=["style", "width", "width_percent", "alignment", "shading", "borders",
                                           "header_rows"])
        edits.append(document.set_table(table, **values))
        return outcome(edits, f"Formatted {table}", changed=[table])
    cell_values: dict[str, Any] = {}
    if shading is not None:
        cell_values["shading"] = fill
    if borders:
        cell_values["borders"] = {k: v for k, v in sides.items() if not k.startswith("inside")}
    if vertical_alignment is not None:
        cell_values["vertical_alignment"] = vertical_alignment
    if scope == "row":
        index = need(row, "row")
        if height is not None or header_rows is not None:
            row_values: dict[str, Any] = {}
            if height is not None:
                row_values.update(height=height, height_rule="exact" if height_rule == "exact" else "atLeast")
            edits.append(document.set_row(table, index, **row_values))
        positions = [(index, c) for c in range(view.column_count)]
    elif scope == "column":
        index = need(column, "column")
        if width is not None:
            edits.append(document.set_column_width(table, index, width))
        positions = [(r, index) for r in range(len(view.rows))]
    else:
        positions = [(need(row, "row"), need(column, "column"))]
        if width is not None:
            cell_values["width"] = width
    if cell_values:
        seen = set()
        for r, c in positions:
            try:
                origin = document.table(table).cell(r, c)
            except IndexError:
                continue
            if origin.id in seen:
                continue
            seen.add(origin.id)
            edits.append(document.format_cell(table, r, c, **cell_values))
    if not edits:
        raise ToolError("invalid_arguments", f"nothing to set for the {scope}",
                        valid_options=["shading", "borders", "vertical_alignment", "width", "height"])
    return outcome(edits, f"Formatted {scope} of {table}", changed=[table])


# -- W25, W26 word_drawings (inserting a picture too) -------------------------------------------

WRAPS = ["square", "tight", "top_and_bottom", "front", "behind"]


@word_tool("word_drawings", "Drawings: list; insert a picture from an image blob (width keeps the aspect ratio), a text box or a shape, inline or floating; move, resize, float or inline, set wrapping and order.",
           {"doc": DOC,
            "action": string("What to do.", enum=["list", "insert_picture", "insert_text_box", "insert_shape",
                                                   "move", "resize", "float", "inline", "set"]),
            "target": string("The drawing (d:...).", optional=True),
            "image": string("insert_picture: image blob handle, e.g. b2.", optional=True),
            "at": string("insert_*: position, or after:/before: a block (a paragraph of its own).",
                         optional=True),
            "find": string("insert_*: text occurring once to anchor after.", optional=True),
            "x": number("move, float, insert_*: offset.", minimum=-1584, maximum=1584, optional=True),
            "y": number("move, float, insert_*: offset.", minimum=-1584, maximum=1584, optional=True),
            "align": string("float, insert_*: across, instead of x; own paragraph: its alignment.",
                            enum=["left", "center", "right"],
                            optional=True),
            "against": string("What x and align measure from. Default column.",
                              enum=["column", "margin", "page"], optional=True),
            "width": number("resize, insert_*.", minimum=1, maximum=1584, optional=True),
            "height": number("resize, insert_text_box, insert_shape.", minimum=1, maximum=1584, optional=True),
            "wrap": string("float, set, insert_*: text wrapping; inserts float with it.",
                           enum=WRAPS, optional=True),
            "z_order": string("set: z-order.", enum=["front", "back", "forward", "backward"],
                              optional=True),
            "text": string("insert_text_box, insert_shape: the text inside.", optional=True),
            "preset": string("insert_shape: geometry, e.g. rect, ellipse.", optional=True),
            "alt_text": string("insert_picture (needed), set: text for screen readers.", optional=True),
            "ref": string("insert_*: ref name for $name.", optional=True),
            "key": string("Retry key: a repeat with the same key makes nothing new.", optional=True)},
           refs=('target', 'at'), group="word_objects", mutates=True)
def word_drawings(call: Any, doc: str, action: str, target: str | None = None, image: str | None = None,
                  at: str | None = None, find: str | None = None, x: float | None = None, y: float | None = None,
                  align: str | None = None, against: str | None = None, width: float | None = None,
                  height: float | None = None, wrap: str | None = None, z_order: str | None = None,
                  text: str | None = None, preset: str | None = None, alt_text: str | None = None,
                  ref: str | None = None, key: str | None = None) -> Result:
    document = call.document
    if action == "list":
        from .read import _drawings, drawing_kind

        rows = []
        for drawing in _drawings(document):
            row = {"id": drawing.id, "kind": drawing_kind(drawing), "size": [round(v, 2) for v in drawing.size],
                   "inline": drawing.inline}
            if drawing.alt_text:
                row["alt_text"] = short(drawing.alt_text, 80)
            rows.append(row)
        return Result(summary=f"{len(rows)} drawing(s)", data=rows)
    if action == "insert_picture":
        own = []
        if at and find is None and at.split(":", 1)[0] in ("after", "before"):
            # A paragraph of its own, beside the block, holding only the picture.
            side, block = at.split(":", 1)
            own.append(document.insert_paragraph("", **{side: resolve(call, block)}))
            at = f"{own[0].id}@0"
            if wrap is None and align is not None:
                own.append(document.format_paragraph(own[0].id, alignment=align))
                align = None
        return _insert_picture(call, need(image, "image"), need(alt_text, "alt_text", " for a picture"),
                               at, find, width, height, _floating(wrap, x, y, align, against, picture=True), ref,
                               own)
    if action.startswith("insert_"):
        where = _position(call, at, find)
        options = {k: v for k, v in (("x", x), ("y", y), ("width", width), ("height", height), ("wrap", wrap))
                   if v is not None}
        if action == "insert_text_box":
            edit = document.insert_text_box(where, text or "", **options)
        else:
            edit = document.insert_shape(where, preset or "rect", text=text, **options)
        remember(call, ref, edit.id)
        return outcome([edit], f"Inserted {edit.id}", data={"id": edit.id})
    target = resolve(call, need(target, "target"))
    if action == "move":
        edit = document.move_drawing(target, need(x, "x"), need(y, "y"))
    elif action == "resize":
        edit = document.resize_drawing(target, width, height)
    elif action == "float":
        edit = document.float_drawing(target, **_floating(wrap, x, y, align, against, picture=False))
    elif action == "inline":
        edit = document.inline_drawing(target)
    else:
        values = {k: v for k, v in (("wrap", wrap), ("z_order", z_order), ("x", x), ("y", y)) if v is not None}
        edits = []
        if alt_text is not None:
            picture = document.picture(target)
            with document.batch():
                picture.alt_text = alt_text
            edits.append(None)
        if values:
            edits.append(document.set_drawing(target, **values))
        if not edits:
            raise ToolError("invalid_arguments", "set needs something to set",
                            valid_options=["wrap", "z_order", "x", "y", "alt_text"])
        return outcome(edits, f"Set {target}", changed=[target])
    return outcome([edit], f"{action}: {target}", changed=[target])


def _floating(wrap: str | None, x: float | None, y: float | None, align: str | None, against: str | None, *,
              picture: bool) -> dict[str, Any] | None:
    """What a float takes: wrapping, and where across and down.  An inserted picture floats
    only when wrap is given (``None`` otherwise), against the column unless ``against`` says."""
    if picture and wrap is None:
        if any(v is not None for v in (x, y, align, against)):
            raise ToolError("invalid_arguments", "a picture floats with wrap: give wrap to place it",
                            field="wrap", valid_options=WRAPS)
        return None
    values: dict[str, Any] = {}
    if wrap is not None:
        values["wrap"] = wrap
    if against is not None or picture:
        values["horizontal_from"] = against or "column"
    if align:
        values["x_align"] = align
    elif x is not None:
        values["x"] = x
    if y is not None:
        values["y"] = y
    return values


def _insert_picture(call: Any, image: str, alt_text: str, at: str | None, find: str | None,
                    width: float | None, height: float | None, floating: dict[str, Any] | None,
                    ref: str | None, before: list | None = None) -> Result:
    if height is not None:
        raise ToolError("invalid_arguments", "a picture keeps its aspect ratio: give width", field="height")
    blob = call.blob(image)
    if not blob.mime.startswith("image/"):
        raise ToolError("invalid_arguments", f"{image} is {blob.mime}, not an image", field="image")
    document = call.document
    edit = document.insert_picture(_position(call, at, find), blob.data, width=width, alt_text=alt_text)
    edits = [*(before or []), edit]
    if floating is not None:
        edits.append(document.float_drawing(edit.id, **floating))
    remember(call, ref, edit.id)
    picture = document.picture(edit.id)
    data = {"id": edit.id, "size": [round(v, 2) for v in picture.size]}
    if before:
        data["paragraph"] = before[0].id
    return outcome(edits, f"Inserted picture {edit.id}", data=data)


# -- W27 word_controls -------------------------------------------------------------------------


@word_tool("word_controls", "Content controls: list, insert (text, drop-down, date, checkbox...), fill as typing would, or remove keeping the content.",
           {"doc": DOC,
            "action": string("What to do.", enum=["list", "insert", "fill", "remove"]),
            "control": string("fill, remove: the control (cc:...).", optional=True),
            "at": string("insert: position, or find.", optional=True),
            "find": string("insert: text occurring once; goes after it.", optional=True),
            "type": string("insert. Default text.",
                           enum=["text", "rich-text", "drop-down", "combo-box", "date", "checkbox"], optional=True),
            "items": array(string(), "insert drop-down, combo-box: the choices.", optional=True),
            "value": string("fill: text, choice, date or true/false.", optional=True),
            "tag": string("insert: a tag for programs.", optional=True),
            "title": string("insert: the title people see.", optional=True)},
           refs=('control', 'at'), group="word_objects", mutates=True)
def word_controls(call: Any, doc: str, action: str, control: str | None = None, at: str | None = None,
                  find: str | None = None, type: str = "text", items: list[str] | None = None,
                  value: str | None = None, tag: str | None = None, title: str | None = None) -> Result:
    document = call.document
    if action == "list":
        return Result(summary="Content controls", data=[
            {"id": c.id, "kind": c.kind, "title": c.alias, "tag": c.tag, "value": _plain(c.value)}
            for c in document.content_controls()][:100])
    if action == "insert":
        edit = document.insert_control(_position(call, at, find), type, items=items, tag=tag, alias=title)
        return outcome([edit], f"Inserted {edit.id}", data={"id": edit.id})
    control = resolve(call, need(control, "control"))
    if action == "remove":
        return outcome([document.remove_control(control)], f"Removed {control}")
    raw = need(value, "value")
    kind = document.content_control(control).kind
    filled: Any = raw
    if kind == "checkbox":
        filled = raw.strip().lower() in ("true", "yes", "1", "checked")
    return outcome([document.fill_control(control, filled)], f"Filled {control}", changed=[control])


def _plain(value: Any) -> Any:
    return value if isinstance(value, (str, int, float, bool)) or value is None else str(value)


TOOLS = [word_edit_table, word_format_table, word_drawings,
         word_controls]
