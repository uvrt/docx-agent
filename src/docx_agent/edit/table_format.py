"""Tables made and formatted: a new table, and a table's, row's, column's and cell's
properties -- as Word 16.106 writes each (``tools/e5_probe.py``, ``tests/observations/
e5-word.json``).

What Word writes, and docx-agent with it (lengths in points here, twips in the file):

* **A new table** is E2's (Word's Insert Table): the document's table style (the one its
  tables use most, else Table Grid), ``w:tblW`` auto, ``w:tblLook`` 04A0, the text width
  shared among the columns, a ``w:tcW`` in each cell.  A table never touches another: Word
  keeps a paragraph between two tables, and after a table that ends a cell.
* **Header rows**: ``w:trPr/w:tblHeader`` on each; **fixed layout**: ``w:tblLayout
  w:type="fixed"``; **a width in percent**: ``w:tblW w:type="pct"`` in fiftieths, the grid
  shared again in proportion (each line at the ceiling of its cumulative share), the cells'
  ``w:tcW`` left as they were; **alignment**: ``w:jc`` in the table's properties *and* every
  row's; **indent**: ``w:tblInd``; **style options**: ``w:tblLook`` with its ``w:val`` the
  bits of what it says; **cell margins**: ``w:tblCellMar`` with only the sides set.
* **A row's height**: ``w:trHeight`` -- at least with no ``w:hRule``, exactly with
  ``w:hRule="exact"``; **a row that may not break**: ``w:cantSplit``.
* **A cell's** shading ``w:shd w:val="clear" w:color="auto" w:fill=...``; borders in
  ``w:tcBorders``, **written on both cells of an edge** (the neighbour's opposite side);
  margins ``w:tcMar``; vertical alignment ``w:vAlign``; text direction
  ``w:textDirection`` (``btLr`` upward, ``tbRl`` downward), with which Word also gives the
  row ``w:cantSplit`` and, where it states no height, an at-least height of 1134 (2 cm),
  and the cell's paragraphs indents of 113 left and right.
* **A column's width**: its ``w:gridCol`` and its cells' ``w:tcW``.
* **A floating table**: ``w:tblpPr`` (``leftFromText``, ``rightFromText``,
  ``topFromText``, ``bottomFromText``, ``vertAnchor``, ``horzAnchor``, ``tblpX``/``tblpY``
  -- Word writes a positive offset **a twip more** than it was given: 100 pt is 2001 -- or
  ``tblpXSpec``/``tblpYSpec``), and ``w:tblOverlap w:val="never"`` where it may not
  overlap another.

**Tracked** (Word, tracking): a table's properties as ``w:tblPrChange`` holding the whole
old ``w:tblPr``, with ``w:tblGridChange`` holding the old grid; a cell's as ``w:tcPrChange``
holding its old ``w:tcPr``, and a paragraph's indents as ``w:pPrChange``.  A row's
properties and a table's floating Word does **not** track: they are applied untracked, and
the result says so.
"""

from __future__ import annotations

import copy
import math
from typing import TYPE_CHECKING

from ..oxml.xml import Element, insert_in_order, make, remove
from . import ids as _ids

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_TBL = _W + "tbl"
W_TR = _W + "tr"
W_TC = _W + "tc"
W_P = _W + "p"

#: A row's and a cell's records, which a property edit leaves where they are.
_ROW_RECORDS = (_W + "ins", _W + "del", _W + "trPrChange")
_CELL_RECORDS = (_W + "cellIns", _W + "cellDel", _W + "cellMerge", _W + "tcPrChange")

#: ``w:tblLook``'s flags: the attribute, its bit in ``w:val``, and whether it says "no".
LOOK = (("first_row", "firstRow", 0x0020, False), ("last_row", "lastRow", 0x0040, False),
        ("first_column", "firstColumn", 0x0080, False), ("last_column", "lastColumn", 0x0100, False),
        ("banded_rows", "noHBand", 0x0200, True), ("banded_columns", "noVBand", 0x0400, True))
#: Word's default look (Insert Table): header row, first column, banded rows.
DEFAULT_LOOK = {"first_row": True, "last_row": False, "first_column": True, "last_column": False,
                "banded_rows": True, "banded_columns": False}

SIDES = {"top": "top", "left": "left", "bottom": "bottom", "right": "right", "inside_h": "insideH",
         "inside_v": "insideV"}
OPPOSITE = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}
DIRECTIONS = {"lrTb": None, "btLr": "btLr", "tbRl": "tbRl", "up": "btLr", "down": "tbRl", "horizontal": None}
#: What Word gives a row and a cell's paragraphs with a vertical text direction (measured).
VERTICAL_ROW_HEIGHT = 1134
VERTICAL_INDENT = 113
#: Word's default distance of floating text from a floating table (Table Properties,
#: "Around": 0.32 cm) -- not measured here; the probe set its own.
FLOAT_DISTANCE = 180


class _Unset:
    def __repr__(self) -> str:
        return "unset"


UNSET = _Unset()


def twips(points: float) -> int:
    return int(round(float(points) * 20))


def points(value) -> float | None:
    try:
        return int(value) / 20
    except (TypeError, ValueError):
        return None


def _int(node: Element | None, name: str = "val", default: int | None = None) -> int | None:
    if node is None:
        return default
    try:
        return int(node.get(_W + name))
    except (TypeError, ValueError):
        return default


def own_rows(table: Element) -> list[Element]:
    out = []
    for row in table.iter(W_TR):
        owner = row.getparent()
        while owner is not None and owner.tag != W_TBL:
            owner = owner.getparent()
        if owner is table:
            out.append(row)
    return out


def row_cells(row: Element) -> list[tuple[Element, int, int]]:
    """``(cell, first grid column, span)`` of a row's own cells."""
    from .tables import _cells

    return _cells(row)


def grid_widths(table: Element) -> list[int]:
    grid = table.find(_W + "tblGrid")
    return [_int(col, "w", 0) or 0 for col in (grid.findall(_W + "gridCol") if grid is not None else [])]


def set_grid(table: Element, widths: list[int]) -> None:
    grid = table.find(_W + "tblGrid")
    if grid is None:
        grid = make("w:tblGrid")
        insert_in_order(table, grid)
    for column in grid.findall(_W + "gridCol"):
        remove(column)
    for width in widths:
        insert_in_order(grid, make("w:gridCol", **{"w:w": str(int(width))}))


def share(total: int, weights: list[int]) -> list[int]:
    """``total`` shared in proportion to ``weights``, each grid line at the ceiling of its
    cumulative share (Word's 80% of 9016 over three equal columns: 2405, 2404, 2404)."""
    whole = sum(weights) or len(weights)
    weights = weights if sum(weights) else [1] * len(weights)
    out, done, running = [], 0, 0
    for weight in weights:
        running += weight
        line = math.ceil(total * running / whole - 1e-9)
        out.append(line - done)
        done = line
    return out


def clean(properties: Element | None, tag: str, drop: tuple[str, ...]) -> Element:
    out = make(tag)
    if properties is not None:
        for child in properties:
            if isinstance(child.tag, str) and child.tag not in drop:
                out.append(copy.deepcopy(child))
    return out


def ensure(parent: Element, tag: str) -> Element:
    node = parent.find(_W + tag.partition(":")[2])
    if node is None:
        node = make(tag)
        insert_in_order(parent, node)
    return node


def drop(parent: Element | None, tag: str) -> None:
    if parent is None:
        return
    for node in parent.findall(_W + tag.partition(":")[2]):
        remove(node)


def border(tag: str, spec) -> Element:
    """A border from a spec: a style name (``"single"``), or ``{style, size (points),
    color, space (points)}``; ``"nil"`` or ``"none"`` is no border."""
    if isinstance(spec, str):
        spec = {"style": spec}
    if not isinstance(spec, dict):
        raise ValueError(f"a border is a style or a dict, not {spec!r}")
    style = spec.get("style", "single")
    node = make(tag)
    node.set(_W + "val", "nil" if style in ("none", None) else style)
    if style not in ("none", "nil", None):
        node.set(_W + "sz", str(int(round(float(spec.get("size", 0.5)) * 8))))
        node.set(_W + "space", str(int(round(float(spec.get("space", 0))))))
        node.set(_W + "color", str(spec.get("color", "auto")))
    return node


def read_border(node: Element | None) -> dict | None:
    if node is None:
        return None
    out = {"style": node.get(_W + "val")}
    if node.get(_W + "sz") is not None:
        out["size"] = _int(node, "sz", 0) / 8
    if node.get(_W + "color") is not None:
        out["color"] = node.get(_W + "color")
    if node.get(_W + "space") is not None:
        out["space"] = _int(node, "space", 0)
    if node.get(_W + "themeColor") is not None:
        out["theme_color"] = node.get(_W + "themeColor")
    return out


def shading(spec) -> Element:
    if isinstance(spec, str):
        spec = {"fill": spec}
    node = make("w:shd")
    node.set(_W + "val", spec.get("pattern", "clear"))
    node.set(_W + "color", spec.get("color", "auto"))
    node.set(_W + "fill", spec.get("fill", "auto"))
    return node


def read_shading(node: Element | None) -> dict | None:
    if node is None:
        return None
    return {"fill": node.get(_W + "fill"), "pattern": node.get(_W + "val"), "color": node.get(_W + "color")}


def width_node(tag: str, value) -> Element:
    """``auto``/``None``, points, or ``"n%"``."""
    if value is None or value == "auto":
        return make(tag, **{"w:w": "0", "w:type": "auto"})
    if isinstance(value, str) and value.strip().endswith("%"):
        percent = float(value.strip()[:-1])
        if not 0 < percent <= 100:
            raise ValueError("a width in percent is above 0 and at most 100")
        return make(tag, **{"w:w": str(int(round(percent * 50))), "w:type": "pct"})
    return make(tag, **{"w:w": str(twips(value)), "w:type": "dxa"})


def read_width(node: Element | None):
    if node is None:
        return None
    kind = node.get(_W + "type") or "dxa"
    raw = node.get(_W + "w") or "0"
    if kind == "auto":
        return "auto"
    if kind == "pct":
        text = raw.rstrip("%")
        try:
            value = float(text) if raw.endswith("%") else int(text) / 50
        except ValueError:
            return None
        return f"{value:g}%"
    if kind == "nil":
        return None
    return points(raw)


def look_node(values: dict) -> Element:
    unknown = set(values) - {name for name, *_ in LOOK}
    if unknown:
        raise ValueError(f"unknown style options {sorted(unknown)}")
    merged = {**DEFAULT_LOOK, **values}
    bits = 0
    attributes = {}
    for name, attribute, bit, negated in LOOK:
        on = bool(merged[name])
        flag = (not on) if negated else on
        attributes["w:" + attribute] = "1" if flag else "0"
        if flag:
            bits |= bit
    return make("w:tblLook", **{"w:val": f"{bits:04X}", **attributes})


def read_look(node: Element | None) -> dict:
    out = dict(DEFAULT_LOOK) if node is None else {}
    if node is None:
        return {name: False for name in out}  # no look: no conditional formatting applies
    raw = node.get(_W + "val")
    bits = int(raw, 16) if raw and all(c in "0123456789abcdefABCDEF" for c in raw) else None
    for name, attribute, bit, negated in LOOK:
        value = node.get(_W + attribute)
        if value is not None:
            flag = value in ("1", "true", "on")
        elif bits is not None:
            flag = bool(bits & bit)
        else:
            flag = False
        out[name] = (not flag) if negated else flag
    return out


# -- tracking --------------------------------------------------------------------------------


class TableTracker:
    """Snapshot a table's properties, grid and cells' properties; after an untracked edit,
    record what changed as Word records it (``w:tblPrChange`` with the whole old
    ``w:tblPr`` and ``w:tblGridChange`` with the old grid; ``w:tcPrChange`` per cell)."""

    def __init__(self, table: Element, *, rows: bool = True) -> None:
        self.table = table
        self.properties = clean(table.find(_W + "tblPr"), "w:tblPr", (_W + "tblPrChange",))
        self.grid = grid_widths(table)
        self.cells: list[tuple[Element, Element]] = [
            (cell, clean(cell.find(_W + "tcPr"), "w:tcPr", _CELL_RECORDS)) for row in own_rows(table)
            for cell, _, _ in row_cells(row)]
        #: With ``rows``, each row's properties too (the grid columns it skips).
        self.rows: list[tuple[Element, Element]] = [
            (row, clean(row.find(_W + "trPr"), "w:trPr", _ROW_RECORDS)) for row in own_rows(table)] if rows else []

    def record(self, stamp, part: str, *, table_level: bool = True, grid: bool = True) -> None:
        """Record what changed, in Word's own form (measured): once anything in the table
        changed, the old ``w:tblPr`` in a ``w:tblPrChange`` and the old grid in a
        ``w:tblGridChange`` (``table_level``), and **every** cell's old ``w:tcPr`` in a
        ``w:tcPrChange``, changed or not; a row's old ``w:trPr`` where it changed, and every
        row's -- with ``w:gridAfter`` 0 where it stated none -- where the grid did.  Word's
        own Reject All restores a table's widths and spans only from that whole record (a
        record of the changed cells alone it rebuilds from widths: measured)."""
        from ..revisions.track import _canonical

        table = self.table
        properties = table.find(_W + "tblPr")
        changed = _canonical(clean(properties, "w:tblPr", (_W + "tblPrChange",))) != _canonical(self.properties)
        grid_changed = grid_widths(table) != self.grid
        rows_changed = [(row, old) for row, old in self.rows if row.getparent() is not None and _canonical(
            clean(row.find(_W + "trPr"), "w:trPr", _ROW_RECORDS)) != _canonical(old)]
        cells_changed = any(cell.getparent() is not None and _canonical(clean(cell.find(_W + "tcPr"), "w:tcPr",
                                                                              _CELL_RECORDS)) != _canonical(old)
                            for cell, old in self.cells)
        if not (changed or grid_changed or rows_changed or cells_changed):
            return
        if table_level and properties is not None and properties.find(_W + "tblPrChange") is None:
            record = stamp.make("w:tblPrChange", part)
            record.append(copy.deepcopy(self.properties))
            insert_in_order(properties, record)
        if grid and table.find(f"{_W}tblGrid/{_W}tblGridChange") is None:
            grid = table.find(_W + "tblGrid")
            record = stamp.make("w:tblGridChange", part)
            old = make("w:tblGrid")
            for width in self.grid:
                old.append(make("w:gridCol", **{"w:w": str(width)}))
            record.append(old)
            insert_in_order(grid, record)
        snapshot = {id(row): old for row, old in self.rows}
        for row in own_rows(table):
            old = snapshot.get(id(row))
            if old is None or not (grid_changed or any(r is row for r, _ in rows_changed)):
                continue
            current = row.find(_W + "trPr")
            if current is None:
                current = make("w:trPr")
                insert_in_order(row, current)
            if current.find(_W + "trPrChange") is None:
                old = copy.deepcopy(old)
                if grid_changed and old.find(_W + "gridAfter") is None:
                    insert_in_order(old, make("w:gridAfter", **{"w:val": "0"}))
                record = stamp.make("w:trPrChange", part)
                record.append(old)
                insert_in_order(current, record)
        for cell, old in self.cells:
            if cell.getparent() is None:
                continue
            current = cell.find(_W + "tcPr")
            if current is None:
                current = make("w:tcPr")
                cell.insert(0, current)
            if current.find(_W + "tcPrChange") is not None:
                continue
            record = stamp.make("w:tcPrChange", part)
            record.append(copy.deepcopy(old))
            insert_in_order(current, record)


# -- the Document half -----------------------------------------------------------------------


class TableFormatOps:
    """Making tables and setting their properties, on :class:`docx_agent.Document`."""

    # -- reading ----------------------------------------------------------------------------

    def table_properties(self: "Document", table: str) -> dict:
        """What a table states itself (``None`` where it states nothing): ``style`` (its
        name), ``width``, ``layout``, ``alignment``, ``indent``, ``look``, ``borders``,
        ``shading``, ``cell_margins``, ``floating``, ``header_rows``, ``columns`` (the
        grid's widths) -- lengths in points, a width in percent as ``"n%"``."""
        _, entry = self._table_entry(table)
        element = entry.element
        props = element.find(_W + "tblPr")
        find = (lambda tag: props.find(_W + tag)) if props is not None else (lambda tag: None)
        style = find("tblStyle")
        style_name = None
        if style is not None:
            found = next((s for s in self.styles if s.id == style.get(_W + "val")), None)
            style_name = found.name if found is not None else style.get(_W + "val")
        layout = find("tblLayout")
        borders = find("tblBorders")
        margins = find("tblCellMar")
        rows = own_rows(element)
        header = 0
        for row in rows:
            if row.find(f"{_W}trPr/{_W}tblHeader") is None:
                break
            header += 1
        return {
            "style": style_name,
            "width": read_width(find("tblW")),
            "layout": (layout.get(_W + "type") or "autofit") if layout is not None else None,
            "alignment": find("jc").get(_W + "val") if find("jc") is not None else None,
            "indent": points(find("tblInd").get(_W + "w")) if find("tblInd") is not None else None,
            "look": read_look(find("tblLook")),
            "borders": {name: read_border(borders.find(_W + tag)) for name, tag in SIDES.items()
                        if borders.find(_W + tag) is not None} if borders is not None else None,
            "shading": read_shading(find("shd")),
            "cell_margins": {side: points(node.get(_W + "w")) for side, node in
                             ((n.tag.rpartition("}")[2], n) for n in margins)} if margins is not None else None,
            "floating": _read_floating(find("tblpPr"), find("tblOverlap")),
            "header_rows": header,
            "columns": [points(w) for w in grid_widths(element)],
        }

    def row_properties(self: "Document", table: str, row: int) -> dict:
        """What a row states itself: its height and rule, header repeat, can't-split..."""
        _, entry = self._table_entry(table)
        element = entry.rows[row][0]
        props = element.find(_W + "trPr")
        height = props.find(_W + "trHeight") if props is not None else None
        return {
            "height": points(height.get(_W + "val")) if height is not None else None,
            "height_rule": (height.get(_W + "hRule") or "atLeast") if height is not None else None,
            "cant_split": props is not None and props.find(_W + "cantSplit") is not None,
            "header": props is not None and props.find(_W + "tblHeader") is not None,
        }

    def cell_properties(self: "Document", table: str, row: int, column: int) -> dict:
        """What a cell states itself: width, shading, borders, vertical alignment, merges..."""
        cell = self._cell_element(table, row, column)
        props = cell.find(_W + "tcPr")
        find = (lambda tag: props.find(_W + tag)) if props is not None else (lambda tag: None)
        borders = find("tcBorders")
        margins = find("tcMar")
        direction = find("textDirection")
        return {
            "width": read_width(find("tcW")),
            "span": _int(find("gridSpan"), default=1),
            "vertical_merge": (find("vMerge").get(_W + "val") or "continue") if find("vMerge") is not None else None,
            "shading": read_shading(find("shd")),
            "borders": {n.tag.rpartition("}")[2]: read_border(n) for n in borders} if borders is not None else None,
            "margins": {n.tag.rpartition("}")[2]: points(n.get(_W + "w")) for n in margins}
            if margins is not None else None,
            "vertical_alignment": find("vAlign").get(_W + "val") if find("vAlign") is not None else None,
            "text_direction": direction.get(_W + "val") if direction is not None else None,
        }

    def _cell_element(self: "Document", table: str, row: int, column: int) -> Element:
        from .document import EditError

        _, entry = self._table_entry(table)
        if not 0 <= row < len(entry.rows):
            raise EditError(f"{entry.id} has no row {row}")
        for cell, start, span in row_cells(entry.rows[row][0]):
            if start <= column < start + span:
                return cell
        raise EditError(f"{entry.id} has no cell at ({row}, {column})")

    # -- a new table ------------------------------------------------------------------------

    def insert_table(self: "Document", rows: int, columns: int, *, after: str | None = None,
                     before: str | None = None, data: list[list[str]] | None = None, style="default",
                     widths: list[float] | None = None, width=None, layout: str | None = None,
                     header_rows: int = 0, alignment: str | None = None, indent: float | None = None,
                     look: dict | None = None) -> "EditResult":
        """A new table of ``rows`` x ``columns`` after or before a block (a paragraph or a
        table; in a cell, a nested table), as Word makes one: the document's table style
        (``style="default"``: the one its tables use most, else Table Grid; a name; ``None``
        for none), the width the table sits in shared among the columns unless ``widths``
        (points) says otherwise, ``data`` the cells' text (a row of strings each).
        ``width`` (points or ``"n%"``), ``layout`` (``fixed``/``autofit``), ``header_rows``,
        ``alignment``, ``indent`` (points) and ``look`` (style options) as :meth:`set_table`
        sets them.  A paragraph is put between it and a table beside it, and after it where
        it would end a cell, as Word keeps one there."""
        from . import text as _text
        from .document import EditError, EditResult

        if (after is None) == (before is None):
            raise EditError("give exactly one of after= or before=")
        if not (isinstance(rows, int) and isinstance(columns, int) and 1 <= rows <= 4096 and 1 <= columns <= 63):
            raise EditError("a table has 1 to 4096 rows and 1 to 63 columns")
        if data is not None:
            if len(data) > rows or any(len(r) > columns for r in data):
                raise EditError("data has more rows or columns than the table")
            for line in data:
                for value in line:
                    _text.check_text(str(value))
        if widths is not None and (len(widths) != columns or any(w <= 0 for w in widths)):
            raise EditError("widths: one positive width (points) per column")
        if not 0 <= header_rows <= rows:
            raise EditError("header_rows is between 0 and the number of rows")
        anchor_id = after if after is not None else before
        part, entry = self._resolve(anchor_id)
        anchor = entry.element
        container = anchor.getparent()
        style_name = None
        if style == "default":
            style_name = self._default_table_style()
        elif style is not None:
            style_name = style
            self.styles.resolve(style, "table")
        total = sum(twips(w) for w in widths) if widths is not None else self._room_for(part, anchor)
        shares = [twips(w) for w in widths] if widths is not None else share(total, [1] * columns)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._resolve(renames.get(anchor_id, anchor_id))
            anchor = entry.element
            container = anchor.getparent()
            style_id = self.styles._ensure(style_name, "table") if style_name is not None else None
            table = make("w:tbl")
            properties = make("w:tblPr")
            table.append(properties)
            if style_id is not None:
                properties.append(make("w:tblStyle", **{"w:val": style_id}))
            insert_in_order(properties, make("w:tblW", **{"w:w": "0", "w:type": "auto"}))
            insert_in_order(properties, look_node(look or {}))
            grid = make("w:tblGrid")
            table.append(grid)
            for value in shares:
                grid.append(make("w:gridCol", **{"w:w": str(value)}))
            template = anchor if anchor.tag == W_P else None
            for r in range(rows):
                row = make("w:tr")
                table.append(row)
                if r < header_rows:
                    row.append(make("w:trPr"))
                    row[0].append(make("w:tblHeader"))
                for c in range(columns):
                    cell = make("w:tc")
                    row.append(cell)
                    cell.append(make("w:tcPr"))
                    cell[0].append(make("w:tcW", **{"w:w": str(shares[c]), "w:type": "dxa"}))
                    paragraph = _cell_paragraph(template)
                    value = data[r][c] if data is not None and r < len(data) and c < len(data[r]) else ""
                    if value:
                        paragraph.append(_text.make_run(str(value), None))
                    cell.append(paragraph)
            if width is not None or layout is not None or alignment is not None or indent is not None:
                _apply_table(table, width=width if width is not None else UNSET,
                             layout=layout if layout is not None else UNSET,
                             alignment=alignment if alignment is not None else UNSET,
                             indent=indent if indent is not None else UNSET, room=self._room_for(part, anchor))
            if after is not None:
                anchor.addnext(table)
            else:
                anchor.addprevious(table)
            spacers = []
            neighbour = _next_block(table) if after is not None else _previous_block(table)
            if neighbour is not None and neighbour.tag == W_TBL:
                spacer = _cell_paragraph(None)
                (table.addnext if after is not None else table.addprevious)(spacer)
                spacers.append(spacer)
            other = _previous_block(table) if after is not None else _next_block(table)
            if other is not None and other.tag == W_TBL and other is not neighbour:
                spacer = _cell_paragraph(None)
                (table.addprevious if after is not None else table.addnext)(spacer)
                spacers.append(spacer)
            if container.tag == W_TC and _next_block(table) is None:
                spacer = _cell_paragraph(None)
                table.addnext(spacer)
                spacers.append(spacer)
            self._stamp_new(part, own_rows(table) + list(table.iter(W_P)) + spacers)
            if tracking is not None:
                from ..markdown.write import _track_table
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                _track_table(table, stamp, part)
                stamp.finish()
                for spacer in spacers:
                    self._track_new_paragraph(part, spacer, tracking)
            _ids.ensure_w14(self.package.tree(part))
            self.package.mark_dirty(part)
            self._invalidate()
            new = self._index(part).entry_for(table)
            index = self._index(part)
            made = [new.id] + [e.id for e in index.paragraphs if any(e.element is p for p in table.iter(W_P))]
            made += [index.entry_for(s).id for s in spacers]
        return EditResult(new.id, created=made, renamed=renames, blocks=[new.id])

    def _default_table_style(self: "Document") -> str:
        counts: dict[str, int] = {}
        for part in self._parts():
            for node in self.package.tree(part).iter(_W + "tblStyle"):
                value = node.get(_W + "val")
                if value:
                    counts[value] = counts.get(value, 0) + 1
        for style_id, _ in sorted(counts.items(), key=lambda kv: -kv[1]):
            found = next((s for s in self.styles if s.id == style_id and s.kind == "table"), None)
            if found is not None and found.name:
                return found.name
        return "Table Grid"

    def _room_for(self: "Document", part: str, anchor: Element) -> int:
        """The width (twips) a table beside ``anchor`` sits in: its cell's text width, or
        its section's text width (A4 with one-inch margins where none says)."""
        node = anchor.getparent()
        while node is not None and node.tag != W_TC:
            node = node.getparent()
        if node is not None:
            width = node.find(f"{_W}tcPr/{_W}tcW")
            if width is not None and (width.get(_W + "type") in (None, "dxa")) and _int(width, "w"):
                return max(_int(width, "w") - 216, 360)
            return 4000
        if part == self.package.document_part():
            entry = self._index(part).entry_for(anchor)
            try:
                section = self.section_of(entry.id) if entry is not None else self.sections()[-1]
                if section.text_width:
                    return section.text_width
            except (KeyError, IndexError, AttributeError):
                pass
        return 9026

    # -- a table's properties ---------------------------------------------------------------

    def set_table(self: "Document", table: str, *, style=UNSET, width=UNSET, layout=UNSET, alignment=UNSET,
                  indent=UNSET, look=UNSET, borders=UNSET, shading=UNSET, cell_margins=UNSET,
                  header_rows=UNSET, floating=UNSET) -> "EditResult":
        """Set a table's properties; a value of ``None`` removes what the table states.

        ``style`` a table style's name; ``width`` ``"auto"``, points or ``"n%"`` (the grid
        shared again in proportion); ``layout`` ``"fixed"`` or ``"autofit"``; ``alignment``
        ``left``, ``center`` or ``right``; ``indent`` points; ``look`` the style options
        (``first_row``, ``last_row``, ``first_column``, ``last_column``, ``banded_rows``,
        ``banded_columns``; what is not given keeps Word's default); ``borders`` a side
        (``top``, ``left``, ``bottom``, ``right``, ``inside_h``, ``inside_v``) -> a border
        (a style, or ``{style, size, color, space}``) or ``None``; ``shading`` a fill colour
        or ``{fill, pattern, color}``; ``cell_margins`` a side -> points; ``header_rows`` how
        many rows from the top repeat on each page; ``floating`` ``{x, y, x_align, y_align,
        horizontal_anchor (text, margin, page), vertical_anchor (text, margin, page), left,
        right, top, bottom (distances, points), overlap}`` or ``None`` to put it in the flow.

        Tracked, the table's properties are a ``w:tblPrChange`` (with ``w:tblGridChange``),
        as Word records them; header rows and floating Word does not track: they are applied
        untracked and a warning says so."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        if style is not UNSET and style is not None:
            self.styles.resolve(style, "table")
        if header_rows is not UNSET and not (isinstance(header_rows, int) and 0 <= header_rows <= len(entry.rows)):
            raise EditError("header_rows is between 0 and the number of rows")
        probe = copy.deepcopy(entry.element)
        try:
            _apply_table(probe, width=width, layout=layout, alignment=alignment, indent=indent, look=look,
                         borders=borders, shading=shading, cell_margins=cell_margins, floating=floating,
                         room=self._room_for(part, entry.element))
        except (ValueError, TypeError) as error:
            raise EditError(str(error)) from None
        tracking = self._active_tracking()
        warnings = []
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            element = entry.element
            # The rows too: Word writes an alignment on every row as well, and a review gives
            # the rows' back with the table's (Word records none for them: measured).
            tracker = TableTracker(element, rows=True) if tracking is not None else None
            style_id = UNSET
            if style is not UNSET:
                style_id = self.styles._ensure(style, "table") if style is not None else None
            _apply_table(element, style_id=style_id, width=width, layout=layout, alignment=alignment,
                         indent=indent, look=look, borders=borders, shading=shading, cell_margins=cell_margins,
                         room=self._room_for(part, element))
            if tracker is not None:
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                tracker.record(stamp, part)
                stamp.finish()
            if floating is not UNSET:
                _apply_table(element, floating=floating, room=0)
                if tracking is not None:
                    warnings.append("Word does not track floating a table: it is applied untracked")
            if header_rows is not UNSET:
                for k, row in enumerate(own_rows(element)):
                    properties = row.find(_W + "trPr")
                    if k < header_rows:
                        if properties is None:
                            properties = make("w:trPr")
                            insert_in_order(row, properties)
                        ensure(properties, "w:tblHeader")
                    elif properties is not None:
                        drop(properties, "w:tblHeader")
                        if not len(properties):
                            remove(properties)
                if tracking is not None:
                    warnings.append("Word does not track header rows: they are applied untracked")
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames, warnings=warnings)

    def set_column_width(self: "Document", table: str, column: int, width: float) -> "EditResult":
        """A grid column's width (points): its ``w:gridCol`` and the ``w:tcW`` of every cell
        over it (a cell spanning it changes by the difference).  Tracked: the old grid in a
        ``w:tblGridChange`` and the cells' old properties in ``w:tcPrChange``s."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        widths = grid_widths(entry.element)
        if not 0 <= column < len(widths):
            raise EditError(f"{entry.id} has no column {column}")
        if not (isinstance(width, (int, float)) and 0 < width <= 1584):
            raise EditError("a width is in points, above 0 and at most 1584")
        new = twips(width)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            element = entry.element
            tracker = TableTracker(element) if tracking is not None else None
            widths = grid_widths(element)
            difference = new - widths[column]
            widths[column] = new
            set_grid(element, widths)
            for row in own_rows(element):
                for cell, start, span in row_cells(row):
                    if start <= column < start + span:
                        properties = cell.find(_W + "tcPr")
                        node = properties.find(_W + "tcW") if properties is not None else None
                        if node is not None and node.get(_W + "type") in (None, "dxa"):
                            node.set(_W + "w", str(max(0, (_int(node, "w", 0) or 0) + difference)))
                        elif node is None or node.get(_W + "type") == "auto":
                            if properties is None:
                                properties = make("w:tcPr")
                                cell.insert(0, properties)
                            drop(properties, "w:tcW")
                            insert_in_order(properties, make("w:tcW", **{"w:w": str(sum(widths[start:start + span])),
                                                                         "w:type": "dxa"}))
            table_width = element.find(f"{_W}tblPr/{_W}tblW")
            if table_width is not None and table_width.get(_W + "type") == "dxa":
                table_width.set(_W + "w", str(max(0, (_int(table_width, "w", 0) or 0) + difference)))
            if tracker is not None:
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                tracker.record(stamp, part, table_level=False)
                stamp.finish()
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames)

    def set_row(self: "Document", table: str, row: int, *, height=UNSET, height_rule: str = "atLeast",
                cant_split=UNSET, header=UNSET) -> "EditResult":
        """A row's height (points; ``height_rule`` ``"atLeast"`` or ``"exact"``; ``None``
        removes it), whether it may break across pages (``cant_split``), and whether it
        repeats as a header row (``header``: header rows are the table's first rows).  Word
        does not track any of them (measured): tracked, they are applied untracked and a
        warning says so."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        if not 0 <= row < len(entry.rows):
            raise EditError(f"{entry.id} has no row {row}")
        if height_rule not in ("atLeast", "exact"):
            raise EditError("height_rule is 'atLeast' or 'exact'")
        if height is not UNSET and height is not None and not (isinstance(height, (int, float)) and 0 < height <= 1584):
            raise EditError("a height is in points, above 0 and at most 1584")
        warnings = []
        if header is True and row > 0 and entry.rows[row - 1][0].find(f"{_W}trPr/{_W}tblHeader") is None:
            warnings.append("Word repeats only the table's first rows as headers: the rows above are not headers")
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            element = entry.rows[row][0]
            properties = element.find(_W + "trPr")
            if properties is None:
                properties = make("w:trPr")
                insert_in_order(element, properties)
            if height is not UNSET:
                drop(properties, "w:trHeight")
                if height is not None:
                    node = make("w:trHeight")
                    if height_rule == "exact":
                        node.set(_W + "hRule", "exact")
                    node.set(_W + "val", str(twips(height)))
                    insert_in_order(properties, node)
            for value, tag in ((cant_split, "w:cantSplit"), (header, "w:tblHeader")):
                if value is UNSET:
                    continue
                drop(properties, tag)
                if value:
                    insert_in_order(properties, make(tag))
            if not len(properties):
                remove(properties)
            if tracking is not None:
                warnings.append("Word does not track a row's height, breaking or header: applied untracked")
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames, warnings=warnings)

    def format_cell(self: "Document", table: str, row: int, column: int, **values) -> "EditResult":
        """Format a cell -- its shading, borders, margins, alignment, direction, width; not
        its text, which is ``doc.table(t).cell(row, column).paragraphs[0].set_text(...)``.
        The same as :meth:`set_cell`, by the name that says what it does."""
        return self.set_cell(table, row, column, **values)

    def set_cell(self: "Document", table: str, row: int, column: int, *, shading=UNSET, borders=UNSET,
                 margins=UNSET, vertical_alignment=UNSET, text_direction=UNSET, width=UNSET) -> "EditResult":
        """A cell's formatting -- not its text (``cell.paragraphs[0].set_text``); also
        :meth:`format_cell`.  Its properties (the cell over grid position ``(row, column)``; ``None``
        removes what it states): ``shading`` a fill colour or ``{fill, pattern, color}``;
        ``borders`` a side (``top``, ``left``, ``bottom``, ``right``) -> a border or
        ``None`` -- written on the neighbour's opposite side too, as Word writes an edge;
        ``margins`` a side -> points; ``vertical_alignment`` ``top``, ``center`` or
        ``bottom``; ``text_direction`` ``lrTb`` (across), ``btLr`` (upward) or ``tbRl``
        (downward) -- vertical, the row may not break and is at least 2 cm high where it
        states no height, and the cell's paragraphs are indented 113 twips each side, as
        Word does; ``width`` points.  Tracked: ``w:tcPrChange`` on each cell changed (and
        ``w:pPrChange`` on its paragraphs), as Word records them."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        self._cell_element(table, row, column)
        if vertical_alignment is not UNSET and vertical_alignment not in (None, "top", "center", "bottom"):
            raise EditError("vertical_alignment is top, center or bottom")
        if text_direction is not UNSET and text_direction not in (None, *DIRECTIONS):
            raise EditError(f"text_direction is one of {', '.join(DIRECTIONS)}")
        if borders is not UNSET and borders is not None:
            bad = set(borders) - set(OPPOSITE)
            if bad:
                raise EditError(f"a cell's border sides are top, left, bottom and right, not {sorted(bad)}")
            for spec in borders.values():
                if spec is not None:
                    try:
                        border("w:top", spec)
                    except (ValueError, TypeError) as error:
                        raise EditError(str(error)) from None
        if margins is not UNSET and margins is not None and set(margins) - set(OPPOSITE):
            raise EditError("a cell's margins are top, left, bottom and right")
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            element = entry.element
            # Not the rows: Word changes a row a vertical text direction needs without a
            # record (measured).
            tracker = TableTracker(element, rows=False) if tracking is not None else None
            with self._track_properties([part]):
                cell = self._cell_element(entry.id, row, column)
                properties = cell.find(_W + "tcPr")
                if properties is None:
                    properties = make("w:tcPr")
                    cell.insert(0, properties)
                if shading is not UNSET:
                    drop(properties, "w:shd")
                    if shading is not None:
                        insert_in_order(properties, _shading(shading))
                if width is not UNSET:
                    drop(properties, "w:tcW")
                    if width is not None:
                        insert_in_order(properties, width_node("w:tcW", width))
                if margins is not UNSET:
                    drop(properties, "w:tcMar")
                    if margins:
                        node = make("w:tcMar")
                        for side, value in margins.items():
                            insert_in_order(node, make(f"w:{side}", **{"w:w": str(twips(value)), "w:type": "dxa"}))
                        insert_in_order(properties, node)
                if vertical_alignment is not UNSET:
                    drop(properties, "w:vAlign")
                    if vertical_alignment not in (None, "top"):
                        insert_in_order(properties, make("w:vAlign", **{"w:val": vertical_alignment}))
                if text_direction is not UNSET:
                    value = DIRECTIONS.get(text_direction) if text_direction is not None else None
                    drop(properties, "w:textDirection")
                    if value is not None:
                        insert_in_order(properties, make("w:textDirection", **{"w:val": value}))
                        _vertical_row(cell)
                if borders is not UNSET:
                    neighbours = _neighbours(element, cell)
                    for side, spec in (borders or {s: None for s in OPPOSITE}).items():
                        _set_cell_border(cell, side, spec)
                        other = neighbours.get(side)
                        if other is not None:
                            _set_cell_border(other, OPPOSITE[side], spec)
                for node in (cell, *[c for c in element.iter(W_TC)]):
                    tcpr = node.find(_W + "tcPr")
                    if tcpr is not None and not len(tcpr):
                        remove(tcpr)
            if tracker is not None:
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                tracker.record(stamp, part, table_level=False)
                stamp.finish()
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames)


# -- helpers ---------------------------------------------------------------------------------


def _shading(spec) -> Element:
    try:
        return shading(spec)
    except (AttributeError, TypeError) as error:
        raise ValueError(f"a shading is a fill colour or a dict: {error}") from None


def _cell_paragraph(template: Element | None) -> Element:
    """An empty paragraph for a new cell, in the default paragraph style (as E2 writes a
    table's cells): a heading's style carried into the cells would keep the table with
    what follows it."""
    return make("w:p")


def _next_block(node: Element) -> Element | None:
    sibling = node.getnext()
    while sibling is not None and not (isinstance(sibling.tag, str) and sibling.tag in (W_P, W_TBL, _W + "sdt")):
        sibling = sibling.getnext()
    return sibling


def _previous_block(node: Element) -> Element | None:
    sibling = node.getprevious()
    while sibling is not None and not (isinstance(sibling.tag, str) and sibling.tag in (W_P, W_TBL, _W + "sdt")):
        sibling = sibling.getprevious()
    return sibling


def _apply_table(table: Element, *, style_id=UNSET, width=UNSET, layout=UNSET, alignment=UNSET, indent=UNSET,
                 look=UNSET, borders=UNSET, shading=UNSET, cell_margins=UNSET, floating=UNSET, room: int) -> None:
    properties = table.find(_W + "tblPr")
    if properties is None:
        properties = make("w:tblPr")
        insert_in_order(table, properties)
    if style_id is not UNSET:
        drop(properties, "w:tblStyle")
        if style_id is not None:
            insert_in_order(properties, make("w:tblStyle", **{"w:val": style_id}))
    if width is not UNSET:
        node = width_node("w:tblW", width)
        drop(properties, "w:tblW")
        insert_in_order(properties, node)
        if node.get(_W + "type") == "pct":
            # Word shares the width again among the columns (measured: 80% of a 9026-twip
            # text width less the outer borders' halves).
            total = int(round((room - 10) * int(node.get(_W + "w")) / 5000))
            set_grid(table, share(total, grid_widths(table)))
        elif node.get(_W + "type") == "dxa":
            set_grid(table, share(int(node.get(_W + "w")), grid_widths(table)))
    if layout is not UNSET:
        if layout not in (None, "fixed", "autofit"):
            raise ValueError("layout is 'fixed' or 'autofit'")
        drop(properties, "w:tblLayout")
        if layout == "fixed":
            insert_in_order(properties, make("w:tblLayout", **{"w:type": "fixed"}))
    if alignment is not UNSET:
        if alignment not in (None, "left", "center", "right"):
            raise ValueError("alignment is left, center or right")
        drop(properties, "w:jc")
        if alignment not in (None, "left"):
            insert_in_order(properties, make("w:jc", **{"w:val": alignment}))
        # Word writes the alignment on every row too (measured).
        for row in own_rows(table):
            row_properties = row.find(_W + "trPr")
            if row_properties is not None:
                drop(row_properties, "w:jc")
            if alignment not in (None, "left"):
                if row_properties is None:
                    row_properties = make("w:trPr")
                    insert_in_order(row, row_properties)
                insert_in_order(row_properties, make("w:jc", **{"w:val": alignment}))
            if row_properties is not None and not len(row_properties):
                remove(row_properties)
    if indent is not UNSET:
        drop(properties, "w:tblInd")
        if indent is not None:
            insert_in_order(properties, make("w:tblInd", **{"w:w": str(twips(indent)), "w:type": "dxa"}))
    if look is not UNSET:
        drop(properties, "w:tblLook")
        if look is not None:
            insert_in_order(properties, look_node(look))
    if borders is not UNSET:
        node = properties.find(_W + "tblBorders")
        if borders is None:
            drop(properties, "w:tblBorders")
        else:
            bad = set(borders) - set(SIDES)
            if bad:
                raise ValueError(f"a table's border sides are {', '.join(SIDES)}, not {sorted(bad)}")
            if node is None:
                node = make("w:tblBorders")
                insert_in_order(properties, node)
            for side, spec in borders.items():
                drop(node, "w:" + SIDES[side])
                if spec is not None:
                    insert_in_order(node, border("w:" + SIDES[side], spec))
            if not len(node):
                remove(node)
    if shading is not UNSET:
        drop(properties, "w:shd")
        if shading is not None:
            insert_in_order(properties, _shading(shading))
    if cell_margins is not UNSET:
        drop(properties, "w:tblCellMar")
        if cell_margins:
            bad = set(cell_margins) - set(OPPOSITE)
            if bad:
                raise ValueError(f"cell margins are top, left, bottom and right, not {sorted(bad)}")
            node = make("w:tblCellMar")
            for side, value in cell_margins.items():
                insert_in_order(node, make(f"w:{side}", **{"w:w": str(twips(value)), "w:type": "dxa"}))
            insert_in_order(properties, node)
    if floating is not UNSET:
        drop(properties, "w:tblpPr")
        drop(properties, "w:tblOverlap")
        if floating:
            node, overlap = _floating(floating)
            insert_in_order(properties, node)
            if overlap is not None:
                insert_in_order(properties, overlap)


_ANCHORS = ("text", "margin", "page")
_X_ALIGN = ("left", "center", "right", "inside", "outside")
_Y_ALIGN = ("top", "center", "bottom", "inside", "outside", "inline")


def _offset(value: float) -> str:
    """Word writes a positive offset a twip more than it was given (measured)."""
    value = twips(value)
    return str(value + 1 if value > 0 else value)


def _floating(spec: dict) -> tuple[Element, Element | None]:
    if not isinstance(spec, dict):
        raise ValueError("floating is a dict of x, y, x_align, y_align, horizontal_anchor, vertical_anchor, "
                         "left, right, top, bottom, overlap")
    known = {"x", "y", "x_align", "y_align", "horizontal_anchor", "vertical_anchor", "left", "right", "top",
             "bottom", "overlap"}
    if set(spec) - known:
        raise ValueError(f"unknown floating settings {sorted(set(spec) - known)}")
    horizontal = spec.get("horizontal_anchor", "margin")
    vertical = spec.get("vertical_anchor", "text")
    if horizontal not in _ANCHORS or vertical not in _ANCHORS:
        raise ValueError(f"anchors are {', '.join(_ANCHORS)}")
    if spec.get("x_align") not in (None, *_X_ALIGN) or spec.get("y_align") not in (None, *_Y_ALIGN):
        raise ValueError("x_align is left, center, right, inside or outside; y_align top, center, bottom, "
                         "inside, outside or inline")
    node = make("w:tblpPr")
    node.set(_W + "leftFromText", str(twips(spec.get("left", FLOAT_DISTANCE / 20))))
    node.set(_W + "rightFromText", str(twips(spec.get("right", FLOAT_DISTANCE / 20))))
    if spec.get("top"):
        node.set(_W + "topFromText", str(twips(spec["top"])))
    if spec.get("bottom"):
        node.set(_W + "bottomFromText", str(twips(spec["bottom"])))
    node.set(_W + "vertAnchor", vertical)
    node.set(_W + "horzAnchor", horizontal)
    if spec.get("x_align") is not None:
        node.set(_W + "tblpXSpec", spec["x_align"])
    elif spec.get("x") is not None:
        node.set(_W + "tblpX", _offset(spec["x"]))
    if spec.get("y_align") is not None:
        node.set(_W + "tblpYSpec", spec["y_align"])
    else:
        node.set(_W + "tblpY", _offset(spec.get("y", 0.05)))
    overlap = make("w:tblOverlap", **{"w:val": "never"}) if spec.get("overlap") is False else None
    return node, overlap


def _read_floating(node: Element | None, overlap: Element | None) -> dict | None:
    if node is None:
        return None

    def offset(name: str) -> float | None:
        raw = _int(node, name)
        if raw is None:
            return None
        return (raw - 1 if raw > 0 else raw) / 20

    out = {"horizontal_anchor": node.get(_W + "horzAnchor") or "margin",
           "vertical_anchor": node.get(_W + "vertAnchor") or "margin",
           "x": offset("tblpX"), "y": offset("tblpY"),
           "x_align": node.get(_W + "tblpXSpec"), "y_align": node.get(_W + "tblpYSpec"),
           "left": points(node.get(_W + "leftFromText") or 0), "right": points(node.get(_W + "rightFromText") or 0),
           "top": points(node.get(_W + "topFromText") or 0), "bottom": points(node.get(_W + "bottomFromText") or 0),
           "overlap": overlap is None or overlap.get(_W + "val") != "never"}
    return out


def _vertical_row(cell: Element) -> None:
    """What Word does to a row and a cell given a vertical text direction (measured)."""
    row = cell.getparent()
    while row is not None and row.tag != W_TR:
        row = row.getparent()
    if row is not None:
        properties = row.find(_W + "trPr")
        if properties is None:
            properties = make("w:trPr")
            insert_in_order(row, properties)
        ensure(properties, "w:cantSplit")
        if properties.find(_W + "trHeight") is None:
            insert_in_order(properties, make("w:trHeight", **{"w:val": str(VERTICAL_ROW_HEIGHT)}))
    for paragraph in [p for p in cell if isinstance(p.tag, str) and p.tag == W_P]:
        properties = paragraph.find(_W + "pPr")
        if properties is None:
            properties = make("w:pPr")
            paragraph.insert(0, properties)
        indent = properties.find(_W + "ind")
        if indent is None:
            indent = make("w:ind")
            insert_in_order(properties, indent)
        indent.set(_W + "left", str(VERTICAL_INDENT))
        indent.set(_W + "right", str(VERTICAL_INDENT))


def _neighbours(table: Element, cell: Element) -> dict[str, Element]:
    """The cells sharing each edge of ``cell`` (where one cell does), by side."""
    rows = own_rows(table)
    placed = [(r, c, start, span) for r, row in enumerate(rows) for c, start, span in row_cells(row)]
    mine = next(((r, start, span) for r, c, start, span in placed if c is cell), None)
    if mine is None:
        return {}
    r, start, span = mine
    out: dict[str, Element] = {}
    for rr, c, s, sp in placed:
        if rr == r and s + sp == start:
            out["left"] = c
        if rr == r and s == start + span:
            out["right"] = c
        if rr == r - 1 and s == start and sp == span:
            out["top"] = c
        if rr == r + 1 and s == start and sp == span:
            out["bottom"] = c
    return out


def _set_cell_border(cell: Element, side: str, spec) -> None:
    properties = cell.find(_W + "tcPr")
    if properties is None:
        properties = make("w:tcPr")
        cell.insert(0, properties)
    node = properties.find(_W + "tcBorders")
    if node is None:
        if spec is None:
            return
        node = make("w:tcBorders")
        insert_in_order(properties, node)
    drop(node, "w:" + side)
    if spec is not None:
        insert_in_order(node, border("w:" + side, spec))
    if not len(node):
        remove(node)
