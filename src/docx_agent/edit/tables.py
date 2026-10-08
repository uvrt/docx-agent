"""Table structure: rows and columns inserted and deleted, cells merged -- untracked, and
tracked as Word tracks them or, where it does not, in the nearest form it accepts.

What Word 16.106 writes with tracking on (``tools/e3_probe.py``):

* **A row inserted**: ``w:trPr/w:ins``, each new cell's paragraph mark inserted;
* **a row deleted**: ``w:trPr/w:del``, each cell's paragraph marks and content deleted;
* **a column inserted**: the new cells, each with its paragraph mark inserted, and the grid
  widened -- no ``w:cellIns``; Word's Reject All removes such cells and their grid column;
* **a column deleted**, **cells merged** (across or down) and **a cell split**: nothing --
  Word applies them untracked.

So a deleted column is written as ``w:cellDel`` on its cells with their content deleted,
and a merge as ``w:cellDel`` on the cells merged away (their content and paragraph marks
deleted) and a ``w:tcPrChange`` on each row's first cell (its new span) -- merged down,
``w:cellMerge`` (``rest``/``cont``) on the rows' first cells -- with the merged cells'
content deleted where it was and inserted in the first cell.  A merge across records, as
Word records a change to a table, every cell's old properties and the old grid: then
Word's own Accept All and Reject All give back exactly what docx-agent's do (measured,
``tools/e5_probe.py`` and the E5 oracle; E3's form, without them, Word reviewed otherwise).
A split, a column across a span and a merge's first row deleted (E5) are written the same
way: new cells' and rows' marks inserted, every cell's old properties recorded, and a
divided vertical merge as ``w:cellMerge`` (``vMerge`` the new, ``vMergeOrig`` the old).
Each is described by a comment (opening with :data:`~docx_agent.revisions.review.
DESCRIBES`), so a reviewer sees what changed; accepting or rejecting it here removes the
comment (ROADMAP.md, decision 7).

Untracked, as Word does (``tools/e5_probe.py``): a merge gives the first cell of each row
the span, ``w:vMerge`` restart and continue down the rows, and appends the other cells'
paragraphs (all but empty ones) to the first cell in reading order; a continuation cell
keeps one empty paragraph.  The rectangle grows to whole cells.  A column goes in beside
a cell spanning its boundary (which keeps its span), and out of a cell spanning it (one
column fewer).  A split divides the cell's grid column (the other rows spanning both) or
shares its span, and divides a vertical merge or adds rows below it.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from ..oxml.xml import Element, insert_in_order, make, remove
from . import ids as _ids
from . import text as _text
from .ids import TableEntry

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_TR = _W + "tr"
W_TC = _W + "tc"
W_P = _W + "p"
W_TBL = _W + "tbl"
_CELL_RECORDS = (_W + "cellIns", _W + "cellDel", _W + "cellMerge", _W + "tcPrChange")


def _int(node: Element | None, default: int) -> int:
    try:
        return int(node.get(_W + "val")) if node is not None else default
    except (TypeError, ValueError):
        return default


def _cells(row: Element) -> list[tuple[Element, int, int]]:
    """``(cell, first grid column, span)`` of a row's cells (row-level content controls
    seen through)."""
    properties = row.find(_W + "trPr")
    column = _int(properties.find(_W + "gridBefore") if properties is not None else None, 0)
    out = []
    for cell in row.iter(W_TC):
        owner = cell.getparent()
        while owner is not None and owner.tag != W_TR:
            owner = owner.getparent()
        if owner is not row:
            continue
        span = max(1, _int(cell.find(f"{_W}tcPr/{_W}gridSpan"), 1))
        out.append((cell, column, span))
        column += span
    return out


def _vmerge(cell: Element) -> str | None:
    node = cell.find(f"{_W}tcPr/{_W}vMerge")
    if node is None:
        return None
    return node.get(_W + "val") or "continue"


def _clean_cell_properties(properties: Element | None) -> Element:
    out = make("w:tcPr")
    if properties is None:
        return out
    for child in properties:
        if isinstance(child.tag, str) and child.tag not in _CELL_RECORDS:
            out.append(copy.deepcopy(child))
    return out


def _empty_paragraph_like(paragraph: Element | None) -> Element:
    """A new empty paragraph with ``paragraph``'s properties (no section break, no revision
    record): what Word puts in a new cell."""
    from ..revisions import track as _track

    new = make("w:p")
    if paragraph is not None and paragraph.find(_W + "pPr") is not None:
        properties = copy.deepcopy(paragraph.find(_W + "pPr"))
        for child in list(properties):
            if child.tag in (_W + "sectPr", _W + "pPrChange"):
                remove(child)
            elif child.tag == _W + "rPr":
                for mark in list(child):
                    if mark.tag in _track.MARKS or mark.tag == _W + "rPrChange":
                        remove(mark)
                if not len(child):
                    remove(child)
        if len(properties):
            new.append(properties)
    return new


def _has_text(paragraph: Element) -> bool:
    return any(True for _ in paragraph.iter(_W + "r"))


class TableOps:
    """Row, column and merge edits on :class:`docx_agent.Document`."""

    def _table_entry(self: "Document", identifier: str) -> tuple[str, TableEntry]:
        from .document import EditError

        part, entry = self._resolve(identifier)
        if not isinstance(entry, TableEntry):
            raise EditError(f"{identifier} is not a table")
        return part, entry

    def _stamp_new(self: "Document", part: str, nodes: list[Element]) -> None:
        used = self._used()
        for node in nodes:
            para_id, text_id = _ids.generate(part + "\0table", used, 2)
            _ids.stamp(node, para_id, text_id)
        _ids.ensure_w14(self.package.tree(part))

    # -- rows -------------------------------------------------------------------------------

    def insert_row(self: "Document", table: str, index: int | None = None, *, below: bool = True) -> "EditResult":
        """A new row beside row ``index`` (0-based; the last by default), below it unless
        ``below=False``: its cells as that row's (properties, widths, and a vertical merge
        carried on where the new row falls inside one), each with an empty paragraph."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        rows = [row for row, _ in entry.rows]
        index = len(rows) - 1 if index is None else index
        if not 0 <= index < len(rows):
            raise EditError(f"{entry.id} has no row {index}")
        reference = rows[index]
        neighbour = rows[index + 1] if below and index + 1 < len(rows) else (rows[index - 1] if not below and index else None)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            rows = [row for row, _ in entry.rows]
            reference = rows[index]
            row = make("w:tr")
            if reference.find(_W + "tblPrEx") is not None:
                row.append(copy.deepcopy(reference.find(_W + "tblPrEx")))
            if reference.find(_W + "trPr") is not None:
                properties = copy.deepcopy(reference.find(_W + "trPr"))
                for child in list(properties):
                    if child.tag in (_W + "ins", _W + "del", _W + "trPrChange"):
                        remove(child)
                if len(properties):
                    row.append(properties)
            next_cells = {start: cell for cell, start, _ in _cells(neighbour)} if neighbour is not None else {}
            for cell, start, span in _cells(reference):
                new_cell = make("w:tc")
                properties = _clean_cell_properties(cell.find(_W + "tcPr"))
                merge = properties.find(_W + "vMerge")
                if merge is not None:
                    remove(merge)
                    inside = (_vmerge(next_cells[start]) == "continue" if below and start in next_cells else False) \
                        or (not below and _vmerge(cell) == "continue")
                    if inside:
                        insert_in_order(properties, make("w:vMerge"))
                new_cell.append(properties)
                new_cell.append(_empty_paragraph_like(cell.find(W_P)))
                row.append(new_cell)
            if below:
                reference.addnext(row)
            else:
                reference.addprevious(row)
            self._stamp_new(part, [row] + list(row.iter(W_P)))
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                properties = row.find(_W + "trPr")
                if properties is None:
                    properties = make("w:trPr")
                    insert_in_order(row, properties)
                insert_in_order(properties, stamp.make("w:ins", part))
                for paragraph in row.iter(W_P):
                    _track.mark(paragraph, "ins", stamp, part)
                stamp.finish()
            self.package.mark_dirty(part)
            story = self._story_of(part)
            new_id = ("" if story == "body" else f"{story}/") + f"tr:{row.get(_ids.PARA_ID)}"
        return EditResult(new_id, created=[new_id], renamed=renames)

    def delete_row(self: "Document", table: str, index: int) -> "EditResult":
        """Delete row ``index``; a vertical merge it starts goes on from the row below it.
        The last row is refused (delete the table)."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        rows = [row for row, _ in entry.rows]
        if not 0 <= index < len(rows):
            raise EditError(f"{entry.id} has no row {index}")
        if len(rows) == 1:
            raise EditError(f"{entry.id} has one row; delete the table instead")
        for paragraph in rows[index].iter(W_P):
            if any(True for _ in paragraph.iter(_W + "footnoteReference", _W + "endnoteReference",
                                               _W + "commentReference")):
                raise EditError(f"row {index} holds a note or comment reference")
        tracking = self._active_tracking()
        old_row_id = entry.rows[index][1]
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            rows = [row for row, _ in entry.rows]
            row = rows[index]
            following = rows[index + 1] if index + 1 < len(rows) else None
            below = {start: cell for cell, start, _ in _cells(following)} if following is not None else {}
            restarts = [below[start] for cell, start, _ in _cells(row)
                        if _vmerge(cell) == "restart" and start in below and _vmerge(below[start]) == "continue"]
            stamp = tracker = None
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp
                from .table_format import TableTracker

                stamp = Stamp(self, tracking)
                # A merge's first row deleted: Word reviews the merge's new start only from
                # its own whole-table record (measured: every cell's old properties).
                tracker = TableTracker(entry.element, rows=False) if restarts else None
            for cell in restarts:
                properties = cell.find(_W + "tcPr")
                properties.find(_W + "vMerge").set(_W + "val", "restart")
            if tracker is not None:
                tracker.record(stamp, part, table_level=False)
                _merges_as_cell_merges(entry.element, tracker, stamp, part)
            if stamp is not None:
                properties = row.find(_W + "trPr")
                if properties is None:
                    properties = make("w:trPr")
                    insert_in_order(row, properties)
                insert_in_order(properties, stamp.make("w:del", part))
                for paragraph in row.iter(W_P):
                    self._release(part, _track.delete_content(paragraph, stamp, part))
                    if _track.mark_record(paragraph) is None:
                        _track.mark(paragraph, "del", stamp, part)
                stamp.finish()
                self.package.mark_dirty(part)
                return EditResult(None, renamed=renames)
            from .document import _rehome_markers

            removed = [old_row_id]
            index_ = self._index(part)
            removed += [e.id for e in index_.paragraphs if any(e.element is p for p in row.iter(W_P))]
            _rehome_markers(row)
            from . import inline as _inline

            self._release(part, _inline.relationship_ids(row))
            remove(row)
            self.package.mark_dirty(part)
        return EditResult(None, renamed=renames, removed=removed)

    # -- columns ----------------------------------------------------------------------------

    def insert_column(self: "Document", table: str, index: int, *, right: bool = True) -> "EditResult":
        """A new column beside grid column ``index`` (0-based), to its right unless
        ``right=False``: as wide as that column, each new cell with that row's neighbour's
        properties (a vertical merge carried on) and an empty paragraph; a fixed table
        width grows by the column's.

        Across a cell that spans the boundary, the new cell goes beside that cell, which
        keeps its span -- Word's Insert Columns (measured): a row is cells, not a grid, to
        it.  A row that skips the grid where the column goes (``w:gridBefore`` or
        ``w:gridAfter``) skips one column more."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        grid = entry.element.find(_W + "tblGrid")
        columns = grid.findall(_W + "gridCol") if grid is not None else []
        if not 0 <= index < len(columns):
            raise EditError(f"{entry.id} has no column {index}")
        boundary = index + 1 if right else index
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            grid = entry.element.find(_W + "tblGrid")
            columns = grid.findall(_W + "gridCol")
            width = columns[index].get(_W + "w")
            new_column = make("w:gridCol")
            if width is not None:
                new_column.set(_W + "w", width)
            if right:
                columns[index].addnext(new_column)
            else:
                columns[index].addprevious(new_column)
            table_width = entry.element.find(f"{_W}tblPr/{_W}tblW")
            if table_width is not None and table_width.get(_W + "type") == "dxa" and width and width.isdigit():
                table_width.set(_W + "w", str(_int_attr(table_width, "w") + int(width)))
            new_paragraphs = []
            structural = any(start < boundary < start + span for row, _ in entry.rows for _, start, span in _cells(row)) \
                or any(_skips(row) != (0, 0) for row, _ in entry.rows)
            tracker = None
            if tracking is not None and structural:
                from .table_format import TableTracker

                # The grid as it was, and the rows' skips, so a review gives them back whole.
                columns[index].getparent().remove(new_column)
                tracker = TableTracker(entry.element, rows=True)
                (columns[index].addnext if right else columns[index].addprevious)(new_column)
            for row, _ in entry.rows:
                cells = _cells(row)
                before, after = _skips(row)
                if not cells or (before and boundary <= before):
                    _skip(row, "gridBefore", "wBefore", 1, width)
                    continue
                last = cells[-1][1] + cells[-1][2]
                if boundary > last or (boundary == last and after):
                    _skip(row, "gridAfter", "wAfter", 1, width)
                    continue
                crossing = next(((cell, start, span) for cell, start, span in cells if start < boundary < start + span),
                                None)
                if crossing is not None:
                    neighbour = crossing[0]
                else:
                    neighbour = next((cell for cell, start, span in cells if start <= index < start + span),
                                     cells[-1][0])
                new_cell = make("w:tc")
                properties = _clean_cell_properties(neighbour.find(_W + "tcPr"))
                span = properties.find(_W + "gridSpan")
                if span is not None:
                    remove(span)
                cell_width = properties.find(_W + "tcW")
                if cell_width is not None and width is not None and cell_width.get(_W + "type") in (None, "dxa"):
                    cell_width.set(_W + "w", width)
                new_cell.append(properties)
                paragraph = _empty_paragraph_like(neighbour.find(W_P))
                new_cell.append(paragraph)
                new_paragraphs.append(paragraph)
                if crossing is not None:
                    (crossing[0].addnext if right else crossing[0].addprevious)(new_cell)
                else:
                    at = next((cell for cell, start, _ in cells if start == boundary), None)
                    if at is not None:
                        at.addprevious(new_cell)
                    else:
                        cells[-1][0].addnext(new_cell)
            self._stamp_new(part, new_paragraphs)
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                for paragraph in new_paragraphs:
                    _track.mark(paragraph, "ins", stamp, part)
                if tracker is not None:
                    tracker.record(stamp, part, table_level=False)
                stamp.finish()
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames,
                          created=[self._paragraph_id(part, p) for p in new_paragraphs])

    def delete_column(self: "Document", table: str, index: int) -> "EditResult":
        """Delete grid column ``index``: the cells only in it, and the grid column; a cell
        spanning it and others spans one column fewer (its width less the column's), a row
        that skips it skips one fewer.  Tracked, the deleted cells are ``w:cellDel`` with
        their content deleted, a spanning cell's old span a ``w:tcPrChange``, and a comment
        says what changed (Word does not track a column's deletion: it applies it)."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        grid = entry.element.find(_W + "tblGrid")
        columns = grid.findall(_W + "gridCol") if grid is not None else []
        if not 0 <= index < len(columns):
            raise EditError(f"{entry.id} has no column {index}")
        if len(columns) == 1:
            raise EditError(f"{entry.id} has one column; delete the table instead")
        for row, _ in entry.rows:
            for cell, start, span in _cells(row):
                if start == index and span == 1 and any(True for _ in cell.iter(
                        _W + "footnoteReference", _W + "endnoteReference", _W + "commentReference")):
                    raise EditError(f"column {index} holds a note or comment reference")
        doomed_count = sum(1 for row, _ in entry.rows for cell, start, span in _cells(row)
                           if start <= index < start + span and span == 1)
        if doomed_count and all(len([c for c in _cells(row) if not (c[1] <= index < c[1] + c[2] and c[2] == 1)]) == 0
                                for row, _ in entry.rows if _cells(row)):
            raise EditError(f"{entry.id} would have no cells left; delete the table instead")
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            element = entry.element
            widths = _grid_widths(element)
            width = widths[index]
            doomed, narrowed, skipped = [], [], []
            for row, _ in entry.rows:
                before, after = _skips(row)
                cells = _cells(row)
                covering = next(((cell, start, span) for cell, start, span in cells if start <= index < start + span),
                                None)
                if covering is None:
                    skipped.append((row, "gridBefore" if index < before else "gridAfter"))
                elif covering[2] == 1:
                    doomed.append(covering[0])
                else:
                    narrowed.append(covering[0])
            spans = bool(narrowed or skipped)
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp
                from .table_format import TableTracker

                tracker = TableTracker(element, rows=True) if spans else None
                stamp = Stamp(self, tracking)
                for cell in doomed:
                    properties = cell.find(_W + "tcPr")
                    if properties is None:
                        properties = make("w:tcPr")
                        cell.insert(0, properties)
                    insert_in_order(properties, stamp.make("w:cellDel", part))
                    for paragraph in cell.iter(W_P):
                        self._release(part, _track.delete_content(paragraph, stamp, part))
                        if _track.mark_record(paragraph) is None:
                            _track.mark(paragraph, "del", stamp, part)
                if spans:
                    # The grid as it will be, the old one in a w:tblGridChange; a cell that
                    # spanned the column spans one fewer, its old span in a w:tcPrChange.
                    for cell in narrowed:
                        _narrow(cell, width)
                    for row, side in skipped:
                        _skip(row, side, "wBefore" if side == "gridBefore" else "wAfter", -1, str(width))
                    del widths[index]
                    _set_grid(element, widths)
                    tracker.record(stamp, part, table_level=False)
                stamp.finish()
                first_row = entry.rows[0][0]
                cells = _cells(first_row)
                anchor = next((cell for cell, start, span in cells if start + span == index), None)
                if anchor is None:
                    anchor = next((cell for cell, start, span in cells if not (start <= index < start + span and span == 1)),
                                  cells[0][0] if cells else None)
                if anchor is not None:
                    self._describe(part, anchor, f"column {index + 1} deleted; Word does not track deleting a "
                                                 "column, so it is written as deleted cells.")
                self.package.mark_dirty(part)
                return EditResult(renames.get(entry.id, entry.id), renamed=renames)
            from . import inline as _inline
            from .document import _rehome_markers

            removed = []
            index_ = self._index(part)
            for cell in doomed:
                removed += [e.id for e in index_.paragraphs if any(e.element is p for p in cell.iter(W_P))]
                _rehome_markers(cell)
                self._release(part, _inline.relationship_ids(cell))
                remove(cell)
            for cell in narrowed:
                _narrow(cell, width)
            for row, side in skipped:
                _skip(row, side, "wBefore" if side == "gridBefore" else "wAfter", -1, str(width))
            del widths[index]
            _set_grid(element, widths)
            table_width = element.find(f"{_W}tblPr/{_W}tblW")
            if table_width is not None and table_width.get(_W + "type") == "dxa":
                table_width.set(_W + "w", str(max(0, _int_attr(table_width, "w") - width)))
            for row, _ in entry.rows:
                if not _cells(row) and row.getparent() is not None:
                    remove(row)
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames, removed=[r for r in removed if r])

    # -- merging ----------------------------------------------------------------------------

    def _rectangle(self: "Document", entry: TableEntry, first: tuple[int, int], last: tuple[int, int]):
        """The rectangle ``first``..``last`` grown until no cell lies partly in it -- across
        a span, down a vertical merge -- and each row's cells in it."""
        from .document import EditError

        (r1, c1), (r2, c2) = first, last
        rows = [row for row, _ in entry.rows]
        if not (0 <= r1 <= r2 < len(rows)) or c1 > c2 or c1 < 0:
            raise EditError("first and last are (row, column), first above and left of last")
        layout = _layout(rows)
        while True:
            grown = False
            for r in range(r1, r2 + 1):
                for cell, start, span in _cells(rows[r]):
                    end = start + span - 1
                    if end < c1 or start > c2:
                        continue
                    if start < c1 or end > c2:
                        c1, c2 = min(c1, start), max(c2, end)
                        grown = True
                    top, bottom = layout.get(id(cell), (r, r))
                    if top < r1 or bottom > r2:
                        r1, r2 = min(r1, top), max(r2, bottom)
                        grown = True
            if not grown:
                break
        plan = []
        for r in range(r1, r2 + 1):
            inside = [(cell, start, span) for cell, start, span in _cells(rows[r]) if c1 <= start <= c2]
            if not inside or inside[0][1] != c1 or inside[-1][1] + inside[-1][2] - 1 != c2:
                raise EditError(f"row {r} does not fill the cells from column {c1} to {c2}")
            plan.append(inside)
        return (r1, c1), (r2, c2), plan

    def merge_cells(self: "Document", table: str, first: tuple[int, int], last: tuple[int, int]) -> "EditResult":
        """Merge the cells from ``first`` to ``last`` (``(row, grid column)``, inclusive)
        into one.  The rectangle grows to take in every cell it touches whole -- a cell
        spanning out of it, a vertical merge reaching beyond it -- as Word's selection of
        cells does; the result's ``warnings`` say when it grew."""
        from .document import EditResult

        part, entry = self._table_entry(table)
        first_, last_, plan = self._rectangle(entry, first, last)
        warnings = []
        if (first_, last_) != (tuple(first), tuple(last)):
            warnings.append(f"the merge takes in whole cells: ({first_[0]}, {first_[1]}) to ({last_[0]}, {last_[1]})")
        if all(len(plan_row) == 1 for plan_row in plan) and (
                len(plan) == 1 or (_vmerge(plan[0][0][0]) == "restart"
                                   and all(_vmerge(r[0][0]) == "continue" for r in plan[1:])
                                   and len({(r[0][1], r[0][2]) for r in plan}) == 1)):
            return EditResult(entry.id, changed=False, warnings=warnings)
        (r1, c1), (r2, c2) = first_, last_
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            rows = [row for row, _ in entry.rows]
            plan = [[(cell, start, span) for cell, start, span in _cells(rows[r]) if c1 <= start <= c2]
                    for r in range(r1, r2 + 1)]
            origin = plan[0][0][0]
            if tracking is None:
                self._merge_plain(part, plan, origin)
            else:
                self._merge_tracked(part, plan, origin, tracking, (r1, c1), (r2, c2), rows[r1:r2 + 1])
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames, warnings=warnings)

    def _merged_properties(self, plan_row) -> Element:
        lead = plan_row[0][0]
        properties = _clean_cell_properties(lead.find(_W + "tcPr"))
        for tag in ("gridSpan", "vMerge", "hMerge"):
            node = properties.find(_W + tag)
            if node is not None:
                remove(node)
        total = sum(span for _, _, span in plan_row)
        if total > 1:
            node = make("w:gridSpan")
            node.set(_W + "val", str(total))
            insert_in_order(properties, node)
        width = properties.find(_W + "tcW")
        widths = [cell.find(f"{_W}tcPr/{_W}tcW") for cell, _, _ in plan_row]
        if width is not None and all(w is not None and w.get(_W + "type") in (None, "dxa")
                                     and (w.get(_W + "w") or "").isdigit() for w in widths):
            width.set(_W + "w", str(sum(int(w.get(_W + "w")) for w in widths)))
        return properties

    def _merge_plain(self: "Document", part: str, plan, origin: Element) -> None:
        """Untracked, as Word merges (measured): each row's first cell takes the span (and
        the vertical merge); every other cell's paragraphs with text go to the end of the
        first cell in reading order -- row by row, left to right -- (copies, from a row's
        first cell below, which keeps its last paragraph, emptied); the other cells go."""
        from .document import _clear_revisions, _rehome_markers

        moved: list[Element] = []
        for k, plan_row in enumerate(plan):
            lead = plan_row[0][0]
            properties = self._merged_properties(plan_row)
            if len(plan) > 1:
                node = make("w:vMerge")
                if k == 0:
                    node.set(_W + "val", "restart")
                insert_in_order(properties, node)
            old = lead.find(_W + "tcPr")
            if old is not None:
                remove(old)
            lead.insert(0, properties)
            if k > 0:
                paragraphs = [p for p in lead if isinstance(p.tag, str) and p.tag == W_P]
                for paragraph in paragraphs:
                    if _has_text(paragraph):
                        copy_ = copy.deepcopy(paragraph)
                        _clear_revisions(copy_)
                        for marker in list(copy_.iter(_W + "bookmarkStart", _W + "bookmarkEnd",
                                                      _W + "commentRangeStart", _W + "commentRangeEnd")):
                            remove(marker)
                        moved.append(copy_)
                for block in [b for b in lead if isinstance(b.tag, str) and b.tag == W_TBL]:
                    moved.append(block)
                for paragraph in paragraphs[:-1]:
                    _rehome_markers(paragraph)
                    remove(paragraph)
                last = paragraphs[-1]
                for child in [c for c in last if not (isinstance(c.tag, str) and c.tag == _W + "pPr")]:
                    if isinstance(child.tag, str) and child.tag in (_W + "bookmarkStart", _W + "bookmarkEnd"):
                        continue
                    remove(child)
            for cell in [cell for cell, _, _ in plan_row][1:]:
                for block in [b for b in cell if isinstance(b.tag, str) and b.tag in (W_P, W_TBL)]:
                    if block.tag == W_TBL or _has_text(block):
                        moved.append(block)
                    else:
                        _rehome_markers(block)
                remove(cell)
        copies = [block for block in moved if block.getparent() is None and block.tag == W_P]
        for block in moved:
            origin.append(block)
        _end_with_paragraph(origin)
        if copies:
            self._stamp_new(part, copies)

    def _merge_tracked(self: "Document", part: str, plan, origin: Element, tracking, first, last,
                       rows: list[Element]) -> None:
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp

        stamp = Stamp(self, tracking)
        across = any(len(plan_row) > 1 for plan_row in plan)
        table = rows[0].getparent()
        while table is not None and table.tag != W_TBL:
            table = table.getparent()
        # Word reviews a merge across exactly only from its own whole-table record: every
        # cell's old properties, the old grid (measured: tools/e5_probe.py, mergeforms).
        from .table_format import TableTracker

        tracker = TableTracker(table) if across else None
        copies: list[Element] = []
        for k, plan_row in enumerate(plan):
            lead = plan_row[0][0]
            cells = [cell for cell, _, _ in plan_row]
            if len(cells) > 1:
                properties = lead.find(_W + "tcPr")
                if properties is None:
                    properties = make("w:tcPr")
                    lead.insert(0, properties)
                new = self._merged_properties(plan_row)
                for child in list(properties):
                    if child.tag not in _CELL_RECORDS:
                        remove(child)
                for child in new:
                    insert_in_order(properties, child)
            if len(plan) > 1:
                properties = lead.find(_W + "tcPr")
                if properties is None:
                    properties = make("w:tcPr")
                    lead.insert(0, properties)
                had = _vmerge(lead)
                existing = properties.find(_W + "vMerge")
                if existing is not None:
                    remove(existing)
                record = stamp.make("w:cellMerge", part)
                record.set(_W + "vMerge", "rest" if k == 0 else "cont")
                if had is not None:
                    record.set(_W + "vMergeOrig", "rest" if had == "restart" else "cont")
                insert_in_order(properties, record)
            for cell in (cells[1:] if k == 0 else cells):
                for paragraph in [p for p in cell if isinstance(p.tag, str) and p.tag == W_P]:
                    if _has_text(paragraph):
                        copy_ = copy.deepcopy(paragraph)
                        from .document import _clear_revisions

                        _clear_revisions(copy_)
                        for marker in list(copy_.iter(_W + "bookmarkStart", _W + "bookmarkEnd",
                                                      _W + "commentRangeStart", _W + "commentRangeEnd")):
                            remove(marker)
                        if _has_text(copy_):
                            copies.append(copy_)
                    self._release(part, _track.delete_content(paragraph, stamp, part))
            for cell in cells[1:]:
                properties = cell.find(_W + "tcPr")
                if properties is None:
                    properties = make("w:tcPr")
                    cell.insert(0, properties)
                insert_in_order(properties, stamp.make("w:cellDel", part))
                # A deleted cell's paragraph marks deleted too: Word accepts a deleted cell
                # whose content is not deleted by widening the cell before it (measured).
                for paragraph in [p for p in cell if isinstance(p.tag, str) and p.tag == W_P]:
                    if _track.mark_record(paragraph) is None:
                        _track.mark(paragraph, "del", stamp, part)
            if k > 0:
                # A continuation keeps its last paragraph: the others' marks go.
                paragraphs = [p for p in lead if isinstance(p.tag, str) and p.tag == W_P]
                for paragraph in paragraphs[:-1]:
                    if _track.mark_record(paragraph) is None:
                        _track.mark(paragraph, "del", stamp, part)
        if tracker is not None:
            tracker.record(stamp, part, table_level=False)
        if copies:
            self._stamp_new(part, copies)
            last_paragraph = [p for p in origin if isinstance(p.tag, str) and p.tag == W_P][-1]
            for copy_ in copies:
                _track.insert_paragraph_content(copy_, stamp, part)
            # Appended at the cell's end, as Word writes a paragraph typed there: the cell's
            # last mark inserted, each new paragraph but the last with its own mark
            # inserted, the last carrying the old mark's properties.
            anchor = last_paragraph
            for copy_ in copies:
                anchor.addnext(copy_)
                anchor = copy_
            sequence = [last_paragraph] + copies
            for paragraph in sequence[:-1]:
                if _track.mark_record(paragraph) is None:
                    _track.mark(paragraph, "ins", stamp, part)
            old = _track.clean_properties(last_paragraph.find(_W + "pPr"), "w:pPr")
            _track.record_paragraph_change(copies[-1], old, stamp, part)
        stamp.finish()
        how = "deleted cells" + (" and a span change" if across else "") + (" and a cell merge" if len(plan) > 1 else "")
        self._describe(part, origin, f"cells ({first[0] + 1}, {first[1] + 1}) to ({last[0] + 1}, {last[1] + 1}) "
                                     f"merged; Word does not track a merge, so it is written as {how}.")

    # -- splitting --------------------------------------------------------------------------

    def split_cell(self: "Document", table: str, row: int, column: int, *, rows: int = 1,
                   columns: int = 2) -> "EditResult":
        """Split the cell over grid position ``(row, column)`` into ``rows`` x ``columns``
        cells, as Word splits one (measured): across, a cell spanning enough grid columns
        is divided among the new cells, otherwise its grid column is divided and every other
        row's cell over it spans the new columns; down, a vertical merge is divided, and
        where the cell has too few rows new rows go below it, the row's other cells merged
        down into them.  The content stays in the first cell; the new cells hold an empty
        paragraph.  Tracked (Word does not track a split: it applies it), the new cells' and
        rows' paragraph marks are insertions, the other cells' old properties
        ``w:tcPrChange``s and the old grid a ``w:tblGridChange``, and a comment says so."""
        from .document import EditError, EditResult

        if not (isinstance(rows, int) and isinstance(columns, int) and 1 <= rows <= 63 and 1 <= columns <= 63):
            raise EditError("rows and columns are whole numbers from 1 to 63")
        part, entry = self._table_entry(table)
        if rows == 1 and columns == 1:
            return EditResult(entry.id, changed=False)
        all_rows = [r for r, _ in entry.rows]
        if not 0 <= row < len(all_rows):
            raise EditError(f"{entry.id} has no row {row}")
        target = next(((cell, start, span) for cell, start, span in _cells(all_rows[row])
                       if start <= column < start + span), None)
        if target is None:
            raise EditError(f"{entry.id} has no cell at ({row}, {column})")
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            element = entry.element
            from .table_format import TableTracker

            tracker = TableTracker(element, rows=True) if tracking is not None else None
            rows_before = [r for r, _ in entry.rows]
            cell, cell_start = next((c, s) for c, s, sp in _cells(rows_before[row]) if s <= column < s + sp)
            layout = _layout(rows_before)
            top, bottom = layout.get(id(cell), (row, row))
            # A cell merged down is split from its merge's first row.
            cell = next((c for c, s, sp in _cells(rows_before[top]) if s == cell_start), cell)
            new_cells: list[Element] = []
            if columns > 1:
                new_cells += _split_across(element, rows_before, top, bottom, cell, columns)
            heads = [cell] + [c for c in new_cells if _row_of(c) is rows_before[top]]
            new_rows: list[Element] = []
            if rows > 1:
                heads = [c for c, s, sp in _cells(rows_before[top])
                         if any(c is h for h in heads)]
                for head in heads:
                    made_rows, made_cells = _split_down(element, head, rows)
                    new_rows += made_rows
                    new_cells += made_cells
            new_paragraphs = [p for c in new_cells for p in c if isinstance(p.tag, str) and p.tag == W_P]
            for r in new_rows:
                new_paragraphs += [p for p in r.iter(W_P) if not any(p is q for q in new_paragraphs)]
            self._stamp_new(part, new_rows + new_paragraphs)
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                for new_row in new_rows:
                    properties = new_row.find(_W + "trPr")
                    if properties is None:
                        properties = make("w:trPr")
                        insert_in_order(new_row, properties)
                    insert_in_order(properties, stamp.make("w:ins", part))
                for paragraph in new_paragraphs:
                    if _track.mark_record(paragraph) is None:
                        _track.mark(paragraph, "ins", stamp, part)
                # Measured: Word's Reject All gives a split column back only without the old
                # grid recorded (it keeps the new columns with it), and a divided vertical
                # merge only with a w:cellMerge beside the cell's old properties.
                tracker.record(stamp, part, table_level=False, grid=columns == 1)
                if not new_rows:
                    _merges_as_cell_merges(element, tracker, stamp, part)
                stamp.finish()
                self._describe(part, cell, f"cell ({row + 1}, {column + 1}) split into {rows} x {columns}; Word "
                                           "does not track a split, so it is written as inserted cells and rows.")
            self.package.mark_dirty(part)
            created = [self._paragraph_id(part, p) for p in new_paragraphs]
        return EditResult(renames.get(entry.id, entry.id), renamed=renames, created=[c for c in created if c])

    def _describe(self: "Document", part: str, cell: Element, text: str) -> None:
        """A comment on a cell's first paragraph saying what structural change the tracked
        revisions in it stand for."""
        from ..revisions.review import DESCRIBES

        paragraph = cell.find(W_P)
        entry = self._index(part).entry_for(paragraph)
        if entry is None:
            self._invalidate()
            entry = self._index(part).entry_for(paragraph)
        from .ranges import TextRange

        self.add_comment(TextRange(self, entry.id, 0, entry.id, 0), f"{DESCRIBES} {text}")

    def _paragraph_id(self: "Document", part: str, paragraph: Element) -> str:
        self._invalidate()
        entry = self._index(part).entry_for(paragraph)
        return entry.id if entry is not None else ""


def _int_attr(node: Element, name: str) -> int:
    try:
        return int(node.get(_W + name) or 0)
    except ValueError:
        return 0


# -- grid helpers ----------------------------------------------------------------------------


def _grid_widths(table: Element) -> list[int]:
    grid = table.find(_W + "tblGrid")
    out = []
    for column in (grid.findall(_W + "gridCol") if grid is not None else []):
        try:
            out.append(int(column.get(_W + "w") or 0))
        except ValueError:
            out.append(0)
    return out


def _set_grid(table: Element, widths: list[int]) -> None:
    grid = table.find(_W + "tblGrid")
    if grid is None:
        grid = make("w:tblGrid")
        insert_in_order(table, grid)
    for column in grid.findall(_W + "gridCol"):
        remove(column)
    for width in widths:
        insert_in_order(grid, make("w:gridCol", **{"w:w": str(int(width))}))


def _skips(row: Element) -> tuple[int, int]:
    """A row's ``w:gridBefore`` and ``w:gridAfter``."""
    properties = row.find(_W + "trPr")
    if properties is None:
        return 0, 0
    return (_int(properties.find(_W + "gridBefore"), 0), _int(properties.find(_W + "gridAfter"), 0))


def _skip(row: Element, grid_tag: str, width_tag: str, delta: int, width: str | None) -> None:
    """Grow (or shrink) a row's ``w:gridBefore``/``w:gridAfter`` by ``delta`` columns, its
    ``w:wBefore``/``w:wAfter`` by as many columns' ``width``."""
    properties = row.find(_W + "trPr")
    if properties is None:
        properties = make("w:trPr")
        insert_in_order(row, properties)
    node = properties.find(_W + grid_tag)
    value = max(0, _int(node, 0) + delta)
    if node is not None:
        remove(node)
    measure = properties.find(_W + width_tag)
    if value:
        insert_in_order(properties, make("w:" + grid_tag, **{"w:val": str(value)}))
        if measure is not None and measure.get(_W + "type") in (None, "dxa") and width and width.isdigit():
            measure.set(_W + "w", str(max(0, _int_attr(measure, "w") + delta * int(width))))
    elif measure is not None:
        remove(measure)
    if not len(properties):
        remove(properties)


def _narrow(cell: Element, width: int) -> None:
    """A cell spanning one grid column fewer, its width less ``width``."""
    properties = cell.find(_W + "tcPr")
    span = properties.find(_W + "gridSpan")
    value = _int(span, 1) - 1
    if value > 1:
        span.set(_W + "val", str(value))
    else:
        remove(span)
    measure = properties.find(_W + "tcW")
    if measure is not None and measure.get(_W + "type") in (None, "dxa") and (measure.get(_W + "w") or "").isdigit():
        measure.set(_W + "w", str(max(0, _int_attr(measure, "w") - width)))


class _Layout(dict):
    """Each cell's (first row, last row) in the vertical merge it belongs to, by ``id()``;
    holds the cells, so their ``id()``\\ s stay theirs."""

    held: list


def _layout(rows: list[Element]) -> _Layout:
    out = _Layout()
    out.held = []
    open_: dict[int, list] = {}  # grid column -> [cells] of the merge running down it
    for r, row in enumerate(rows):
        for cell, start, span in _cells(row):
            out.held.append(cell)
            merge = _vmerge(cell)
            if merge == "continue" and start in open_:
                open_[start].append((cell, r))
            else:
                for column, group in list(open_.items()):
                    if column == start:
                        del open_[column]
                if merge in ("restart", "continue"):
                    open_[start] = [(cell, r)]
                else:
                    open_.pop(start, None)
                    out[id(cell)] = (r, r)
            group = open_.get(start)
            if group is not None:
                first = group[0][1]
                for member, _ in group:
                    out[id(member)] = (first, r)
        # A column no cell of this row continues ends its merge.
        starts = {start for _, start, _ in _cells(row)}
        for column in list(open_):
            if column not in starts:
                del open_[column]
    return out


def _row_of(cell: Element) -> Element | None:
    node = cell.getparent()
    while node is not None and node.tag != W_TR:
        node = node.getparent()
    return node


def _end_with_paragraph(cell: Element) -> None:
    blocks = [b for b in cell if isinstance(b.tag, str) and b.tag in (W_P, _W + "tbl", _W + "sdt")]
    if not blocks or blocks[-1].tag != W_P:
        cell.append(make("w:p"))


def _new_cell_like(cell: Element, span: int, width: int | None, merge: str | None) -> Element:
    """A new cell with ``cell``'s properties (no span, merge or records but those given),
    holding an empty paragraph like its first."""
    properties = _clean_cell_properties(cell.find(_W + "tcPr"))
    for tag in ("gridSpan", "vMerge", "hMerge"):
        node = properties.find(_W + tag)
        if node is not None:
            remove(node)
    _set_span_width(properties, span, width)
    if merge is not None:
        node = make("w:vMerge")
        if merge == "restart":
            node.set(_W + "val", "restart")
        insert_in_order(properties, node)
    new = make("w:tc")
    new.append(properties)
    new.append(_empty_paragraph_like(cell.find(W_P)))
    return new


def _set_span_width(properties: Element, span: int, width: int | None) -> None:
    node = properties.find(_W + "gridSpan")
    if node is not None:
        remove(node)
    if span > 1:
        insert_in_order(properties, make("w:gridSpan", **{"w:val": str(span)}))
    if width is not None:
        measure = properties.find(_W + "tcW")
        if measure is None:
            insert_in_order(properties, make("w:tcW", **{"w:w": str(width), "w:type": "dxa"}))
        elif measure.get(_W + "type") in (None, "dxa", "auto"):
            measure.set(_W + "w", str(width))
            measure.set(_W + "type", "dxa")


def _set_merge(cell: Element, merge: str | None) -> None:
    properties = cell.find(_W + "tcPr")
    if properties is None:
        if merge is None:
            return
        properties = make("w:tcPr")
        cell.insert(0, properties)
    node = properties.find(_W + "vMerge")
    if node is not None:
        remove(node)
    if merge is not None:
        node = make("w:vMerge")
        if merge == "restart":
            node.set(_W + "val", "restart")
        insert_in_order(properties, node)
    if not len(properties):
        remove(properties)


def _split_across(table: Element, rows: list[Element], top: int, bottom: int, cell: Element, n: int) -> list[Element]:
    """Split ``cell`` (and the cells under it in its vertical merge) into ``n`` across, as
    Word does: the grid divided where the new edges fall between its lines (every other
    cell over a divided column spanning the new columns), the cell's span shared where
    they fall on them.  The new cells, in document order."""
    widths = _grid_widths(table)
    lines = [0]
    for width in widths:
        lines.append(lines[-1] + width)
    start, span = next((s, sp) for c, s, sp in _cells(rows[top]) if c is cell)
    x0, x1 = lines[start], lines[start + span]
    cuts = [x0 + round((x1 - x0) * k / n) for k in range(1, n)]
    new_lines = sorted(set(lines) | set(cuts))
    if new_lines != lines:
        # Every row's cells and skips as positions, then spans over the new grid.
        positions = []
        for row in own_table_rows(table):
            before, after = _skips(row)
            cells = [(c, lines[s], lines[s + sp]) for c, s, sp in _cells(row)]
            end = lines[min(len(lines) - 1, before + sum(sp for _, _, sp in _cells(row)) + after)]
            positions.append((row, lines[before], cells, end))
        index = {x: k for k, x in enumerate(new_lines)}
        _set_grid(table, [b - a for a, b in zip(new_lines, new_lines[1:])])
        for row, left, cells, right in positions:
            before, after = _skips(row)
            if before:
                _skip(row, "gridBefore", "wBefore", index[left] - before, None)
            for c, a, b in cells:
                properties = c.find(_W + "tcPr")
                if properties is None:
                    properties = make("w:tcPr")
                    c.insert(0, properties)
                node = properties.find(_W + "gridSpan")
                if node is not None:
                    remove(node)
                if index[b] - index[a] > 1:
                    insert_in_order(properties, make("w:gridSpan", **{"w:val": str(index[b] - index[a])}))
                if not len(properties):
                    remove(properties)
            if after:
                last = cells[-1][2] if cells else left
                _skip(row, "gridAfter", "wAfter", (index[right] - index[last]) - after, None)
        lines = new_lines
    index = {x: k for k, x in enumerate(lines)}
    edges = [x0] + cuts + [x1]
    made = []
    for r in range(top, bottom + 1):
        member = cell if r == top else next((c for c, s, sp in _cells(rows[r]) if s == index[x0]), None)
        if member is None:
            continue
        merge = _vmerge(member) if bottom > top else None
        properties = member.find(_W + "tcPr")
        if properties is None:
            properties = make("w:tcPr")
            member.insert(0, properties)
        _set_span_width(properties, index[edges[1]] - index[edges[0]], edges[1] - edges[0])
        anchor = member
        for k in range(1, n):
            new = _new_cell_like(member, index[edges[k + 1]] - index[edges[k]], edges[k + 1] - edges[k], merge)
            anchor.addnext(new)
            anchor = new
            made.append(new)
    return made


def own_table_rows(table: Element) -> list[Element]:
    out = []
    for row in table.iter(W_TR):
        owner = row.getparent()
        while owner is not None and owner.tag != _W + "tbl":
            owner = owner.getparent()
        if owner is table:
            out.append(row)
    return out


def _split_down(table: Element, head: Element, m: int) -> tuple[list[Element], list[Element]]:
    """Split ``head`` (a cell in its merge's first row) into ``m`` down, as Word does: its
    vertical merge divided into ``m`` merges (rows shared out, the first ones a row more),
    and where it has fewer than ``m`` rows, new rows below its last -- each with a plain cell
    under it and every other cell of the row merged down into it."""
    rows = own_table_rows(table)
    layout = _layout(rows)
    top, bottom = layout.get(id(head), (rows.index(_row_of(head)),) * 2)
    start, span = next((s, sp) for c, s, sp in _cells(rows[top]) if c is head)
    made_rows, made_cells = [], []
    g = bottom - top + 1
    if g < m:
        last = rows[bottom]
        anchor = last
        for _ in range(m - g):
            new_row = make("w:tr")
            if last.find(_W + "tblPrEx") is not None:
                new_row.append(copy.deepcopy(last.find(_W + "tblPrEx")))
            if last.find(_W + "trPr") is not None:
                properties = copy.deepcopy(last.find(_W + "trPr"))
                for child in list(properties):
                    if child.tag in (_W + "ins", _W + "del", _W + "trPrChange", _W + "tblHeader"):
                        remove(child)
                if len(properties):
                    new_row.append(properties)
            for cell, s, sp in _cells(last):
                width = _int_attr(cell.find(f"{_W}tcPr/{_W}tcW"), "w") if cell.find(f"{_W}tcPr/{_W}tcW") is not None else None
                if s == start and sp == span:
                    new = _new_cell_like(cell, sp, width, None)
                    made_cells.append(new)
                else:
                    if _vmerge(cell) is None:
                        _set_merge(cell, "restart")
                    new = _new_cell_like(cell, sp, width, "continue")
                    made_cells.append(new)
                new_row.append(new)
            anchor.addnext(new_row)
            anchor = new_row
            made_rows.append(new_row)
        rows = own_table_rows(table)
        column_cells = [next((c for c, s, sp in _cells(rows[r]) if s == start), None) for r in range(top, top + m)]
        for c in column_cells:
            if c is not None:
                _set_merge(c, None)
        return made_rows, made_cells
    sizes = [g // m + (1 if k < g % m else 0) for k in range(m)]
    r = top
    for size in sizes:
        for k in range(size):
            c = next((c for c, s, sp in _cells(rows[r + k]) if s == start), None)
            if c is None:
                continue
            _set_merge(c, None if size == 1 else ("restart" if k == 0 else "continue"))
        r += size
    return made_rows, made_cells


def _merges_as_cell_merges(table: Element, tracker, stamp, part: str) -> None:
    """Each cell whose vertical merge a tracked edit changed: the new merge as a
    ``w:cellMerge`` (``vMerge`` the new, ``vMergeOrig`` the old), its old properties kept
    in its ``w:tcPrChange`` -- the form Word reviews a divided merge in (measured)."""
    short = {"restart": "rest", "continue": "cont"}
    for cell, old in tracker.cells:
        if cell.getparent() is None:
            continue
        before = old.find(_W + "vMerge")
        was = None if before is None else (before.get(_W + "val") or "continue")
        now = _vmerge(cell)
        if was == now:
            continue
        properties = cell.find(_W + "tcPr")
        node = properties.find(_W + "vMerge")
        if node is not None:
            remove(node)
        record = stamp.make("w:cellMerge", part)
        if now is not None:
            record.set(_W + "vMerge", short[now])
        if was is not None:
            record.set(_W + "vMergeOrig", short[was])
        insert_in_order(properties, record)
