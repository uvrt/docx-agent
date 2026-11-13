"""Objects: tables (``word_edit_table``, ``word_format_table``; a new table is Markdown,
through ``word_insert_markdown``, which with ``word_edit_table``'s ``set_cells`` also takes
a table's rows from a dataset or another document's table: :mod:`.table_data`),
pictures and other drawings (``word_drawings``, which inserts a picture too) and content
controls (``word_controls``)."""

from __future__ import annotations

import re
import time
from typing import Any

from ooxml_edit.tools import (REVIEW, Result, ToolError, array, boolean, integer, names_as, number, obj,
                              string)

from ..edit import provenance as _provenance
from ..edit.pictures import PictureError
from . import positions
from ._base import Timing, engage, need, note_small_text, outcome, remember, resolve, short, word_tool
from .structure import _position
from .originals import Watch
from .originals import add as add_original
from .table_data import DATE_FORMAT, HEADER, MAX_COLUMNS, MAX_ROWS, cell_text, fits, no_source, placed

DOC = string("Document id.")
TABLE = string("Table (t:...).")
#: A cell's id, ``t:<table>/c<row>,<column>``.
_CELL = re.compile(r"^t[:@].+/c\d+,\d+$")
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def table_json(document: Any, table_id: str, *, markdown: bool = True) -> dict[str, Any]:
    """A table's id, size and (``markdown``) its text as Markdown -- left out for a table
    written from a source, whose caller has the data."""
    table = document.table(table_id)
    kept = document.accepted_columns(table.id)[0]
    # Round 54 (lane M): the columns col positions count -- the table as accepted, as the
    # Markdown shows it; the grid columns a pending tracked deletion removes, as information.
    # Round 54 (lane R): rows too -- a row whose tracked deletion is pending not counted.
    entry = _table_entry(document, table.id)
    rows = _provenance.accepted_rows(entry) if entry is not None else list(range(len(table.rows)))
    data = {"id": table.id, "rows": len(rows), "columns": len(kept)}
    pending = [c for c in range(table.column_count) if c not in kept]
    if pending:
        data["columns_deleted_pending"] = pending
    gone = [r for r in range(len(table.rows)) if r not in rows]
    if gone:
        data["rows_deleted_pending"] = gone
    if markdown:
        data["markdown"] = document.to_markdown(table.id, ids=False)
    return data


class _Accepted:
    """Round 54 (lane M): ``word_edit_table``'s (and ``word_format_table``'s) column positions
    count the table as accepting its tracked changes leaves it -- as ``word_read``'s final
    view and render show it: a column whose tracked deletion is pending is not counted (the
    deletion is reached through ``word_changes``), and a position in a cell a pending
    tracked merge takes in is that merged cell's.  Round 54 (lane R): rows too -- a row
    whose tracked deletion is pending is not counted (``rows``: the document's rows kept, in
    order); a cell Accept All takes into the one before it (its last paragraph mark deleted)
    answers as that cell, and one Accept All removes with its text is not written.  Cell ids
    (``t:X/cR,C``) stay on the document's grid, rows and columns.  Without such a change the
    positions are the document's, as before."""

    def __init__(self, document: Any, table: str) -> None:
        self.document = document
        self.table = table
        self.kept, self.given = document.accepted_columns(table)
        self.count = document._grid_count(table)
        self.active = len(self.kept) != self.count or bool(self.given)
        entry = _table_entry(document, table)
        self.total_rows = len(entry.rows) if entry is not None else 0
        self.rows = _provenance.accepted_rows(entry) if entry is not None else []
        self.rows_active = len(self.rows) != self.total_rows
        self.lost = {(r, c) for r, c, _ in document.cells_lost_on_accept(table)}
        self.active = self.active or bool(self.lost)
        #: What the call's positions stepped over or into, said once each (rule 2).
        self.notes: list[str] = []

    def place(self, row: int, col: int, field: str) -> tuple[int, int]:
        """Position (``row``, ``col``) as (document row, grid column), as accepted."""
        if self.rows_active:
            row = self.to_row(row, f"{field}.row")
        if not self.active:
            return row, col
        return row, self.at(row, self.to_grid(col, f"{field}.col"))

    @property
    def height(self) -> int:
        return len(self.rows)

    def row_shape(self) -> str:
        skipped = [r for r in range(self.total_rows) if r not in self.rows]
        return (f"{self.table} has rows 0 to {self.height - 1}" + (
            f" as accepted (document row(s) {', '.join(map(str, skipped))} deleted by a pending tracked "
            f"change, not counted)" if skipped else ""))

    def to_row(self, row: int, field: str) -> int:
        """The document's row of position ``row``; refused past the last (rule 2)."""
        if not isinstance(row, int) or isinstance(row, bool) or not 0 <= row < self.height:
            raise ToolError("not_found", f"{field} {row}: {self.row_shape()}", field=field,
                            valid_options=[str(r) for r in range(self.height)][:50])
        found = self.rows[row]
        if found != row:
            gone = [r for r in range(found) if r not in self.rows]
            said = ", ".join(f"{r} ({n!r})" if n else str(r) for r, n in
                             ((r, _row_text(self.document, self.table, r)) for r in gone))
            note = (f"row positions count {self.table} as accepted: document row {said} has a tracked deletion "
                    f"not yet accepted and is not counted (word_changes reaches it; cell ids keep the "
                    f"document's rows)")
            if note not in self.notes:
                self.notes.append(note)
        return found

    def from_row(self, row: int) -> int:
        """Position of document row ``row`` (of the kept row before it, for one not kept)."""
        return max(0, sum(1 for r in self.rows if r <= row) - 1)

    @property
    def width(self) -> int:
        return len(self.kept)

    def shape(self) -> str:
        skipped = [c for c in range(self.count) if c not in self.kept]
        return (f"{self.table} has columns 0 to {self.width - 1}" + (
            f" as accepted (grid column(s) {', '.join(map(str, skipped))} deleted by a pending tracked "
            f"change, not counted)" if skipped else ""))

    def to_grid(self, col: int, field: str) -> int:
        """The document's grid column of position ``col``; refused past the last (rule 2)."""
        if not isinstance(col, int) or isinstance(col, bool) or not 0 <= col < self.width:
            raise ToolError("not_found", f"{field} {col}: {self.shape()}", field=field,
                            valid_options=[str(c) for c in range(self.width)][:50])
        grid = self.kept[col]
        if grid != col:
            self._skipped(col, grid)
        return grid

    def at(self, row: int, grid: int) -> int:
        """Where (document row ``row``, grid column ``grid``) is written: a cell a pending
        merge, or Accept All, takes in answers as the cell taking it in; one Accept All
        removes with its text is refused, with the reject that keeps it (rule 2)."""
        if (row, grid) in self.lost:
            raise ToolError("refused", f"({self.from_row(row)}, {self.from_grid(grid)}) of {self.table} is a cell "
                            "accepting removes (its last paragraph mark is deleted and no cell follows it in its "
                            "row): Word's Accept All takes its text out of view, word_changes accept moves the "
                            "text after the row. Reject that change first to write in the cell (check's "
                            f"accept_loses_text gives the call)", field="cells",
                            valid_options=["word_changes reject"])
        found = self.given.get((row, grid))
        if found is None:
            return grid
        try:
            merged = self.document.table(self.table).cell(row, grid)._element.find(
                f"{_W}tcPr/{_W}cellDel") is not None
        except (IndexError, KeyError):
            merged = True
        how = "a pending tracked merge" if merged else "accepting a deleted paragraph mark"
        note = (f"({self.from_row(row)}, {self.from_grid(grid)}) is in the cell {how} makes of "
                f"({self.from_row(row)}, {self.from_grid(found)}): written there")
        if note not in self.notes:
            self.notes.append(note)
        return found

    def from_grid(self, grid: int) -> int:
        """Position of grid column ``grid`` (the position of the kept column before it, for
        a column not kept)."""
        return max(0, sum(1 for c in self.kept if c <= grid) - 1)

    def labels(self, row: int) -> list[str]:
        view = self.document.table(self.table)
        return [_text(view, row, c) for c in self.kept]

    def _skipped(self, col: int, grid: int) -> None:
        gone = [c for c in range(grid) if c not in self.kept]
        names = [_deleted_header(self.document, self.table, c) for c in gone]
        said = ", ".join(f"{c} ({n!r})" if n else str(c) for c, n in zip(gone, names))
        note = (f"col positions count {self.table} as accepted: grid column {said} has a tracked deletion "
                f"not yet accepted and is not counted (word_changes reaches it); col {col} is grid column {grid}")
        if not any(n.startswith(note.split("; col ")[0]) for n in self.notes):
            self.notes.append(note)


def _loses_text(document: Any, table: str) -> list[str]:
    """Round 54 (lane R): the cells of ``table`` Word's Accept All removes with their text
    (check's ``accept_loses_text``), each said with the reject that keeps it (rule 7)."""
    from ..layout_facts import accept_loses_text

    try:
        if not document.cells_lost_on_accept(table):
            return []
    except (KeyError, ValueError):
        return []
    return [f"{e['cell']} ({e['text']!r}): Word's Accept All drops this cell and its text from view (its last "
            f"paragraph mark is a pending deletion, no cell after it in the row); word_changes accept "
            f"moves the text after the row, splitting the table; word_changes reject "
            f"{e['change'] or 'within ' + e['cell']} keeps it in the cell"
            for e in accept_loses_text(document) if e["cell"].rsplit("/c", 1)[0] == table]


def _row_text(document: Any, table: str, row: int) -> str:
    """Document row ``row``'s first cell's text (deleted text included), shortened."""
    entry = _table_entry(document, table)
    if entry is None or row >= len(entry.rows):
        return ""
    cell = entry.rows[row][0].find(_W + "tc")
    if cell is None:
        return ""
    return short(" ".join("".join(n.text or "" for n in cell.iter(_W + "t", _W + "delText")).split()), 40)


def _deleted_header(document: Any, table: str, grid: int) -> str:
    """The first row's text (deleted text included) of grid column ``grid``."""
    try:
        cell = document.table(table).cell(0, grid)._element
    except (IndexError, KeyError):
        return ""
    return " ".join("".join(node.text or "" for node in cell.iter(_W + "t", _W + "delText")).split())


def _settle_first(document: Any, doc: str, table: str, records: list[Any], message: str) -> ToolError:
    """Round 55 (R115): a refusal whose way out is accepting the pending changes in the way --
    only those, by id (rule 2): each change ``word_changes`` lists that holds one of
    ``records``."""
    listed = document._revision_elements()      # held: their id()s stay theirs
    held = {id(element): identifier for _, identifier, element in listed}
    wanted = {held[id(r)] for r in records if id(r) in held}
    ids = [c.id for c in document.changes() if wanted & set(c.revisions)]
    fix = {"tool": "word_changes", "calls": [{"action": "accept", "ids": ids}]}
    return ToolError("refused", f"{message}. Accept that change first (word_changes accept ids {ids}; reject keeps "
                     f"the cells apart instead), read {table} again, then retry",
                     valid_options=[f"word_changes accept ids {ids}", f"word_changes reject ids {ids}"],
                     details={"fix": fix, "ids": ids,
                              "call": {"tool": "word_changes", "arguments": {"doc": doc, "action": "accept",
                                                                             "ids": ids}}})


def _no_reshape_over_merge_lead(document: Any, doc: str, table: str, accepted: "_Accepted", action: str,
                                index: int) -> None:
    """Round 55 (R115): tracked, deleting a column or row that holds the first cell of a
    pending tracked merge would move the merge to another cell: refused, the merge's change
    named by id (accepting it, or rejecting it, settles it).  A column or row holding only
    cells the merge took in is deleted, their text leaving the merged cell."""
    from ..edit.tables import _cells, merge_leads
    from ..revisions.review import merge_groups

    if document._active_tracking() is None:
        return
    entry = _table_entry(document, table)
    if entry is None:
        return
    rows = [row for row, _ in entry.rows]
    if action == "delete_column":
        leads = merge_leads(entry.element)
        hit = [leads[id(cell)][1] for row in rows for cell, start, span in _cells(row)
               if start == index and span == 1 and id(cell) in leads]
        what = f"column {accepted.from_grid(index)}"
    else:
        hit = [[cell for _, cell in group] for group in merge_groups(entry.element)
               if any(cell is group[0][1] for cell in rows[index].iter(_W + "tc"))]
        what = f"row {accepted.from_row(index)}"
    if not hit:
        return
    records = [node for cells in hit for cell in cells
               for node in cell.iter(_W + "cellDel", _W + "cellMerge")]
    raise _settle_first(document, doc, table, records,
                        f"{what} of {table} holds the first cell of a pending tracked merge; deleting it would "
                        f"move the merge to another cell")


def _no_split_over_pending(document: Any, doc: str, table: str, accepted: "_Accepted", row: int,
                           grid: int) -> None:
    """A split reaching a pending change it cannot be written over.  Untracked: any pending
    deletion or merge it reaches (round 54, lane M).  Tracked (round 55, R115): a cell a
    pending tracked merge makes -- split it once that merge is settled (by id)."""
    from ..edit.tables import _cells, merge_leads

    if document._active_tracking() is None:
        _no_merge_over_pending(doc, table, accepted, grid, grid, row, row, "split")
        return
    entry = _table_entry(document, table)
    if entry is None or not 0 <= row < len(entry.rows):
        return
    leads = merge_leads(entry.element)
    cell = next((c for c, start, span in _cells(entry.rows[row][0]) if start <= grid < start + span), None)
    if cell is None or id(cell) not in leads:
        return
    records = [node for c in leads[id(cell)][1] for node in c.iter(_W + "cellDel", _W + "cellMerge")]
    raise _settle_first(document, doc, table, records,
                        f"({accepted.from_row(row)}, {accepted.from_grid(grid)}) of {table} is a cell a pending "
                        f"tracked merge makes; a split of it is written once the merge is settled")


def _corner(accepted: _Accepted, position: dict, field: str) -> tuple[int, int]:
    """A merge's or split's ``first``/``last`` as (row, grid column)."""
    return accepted.place(position["row"], position["col"], field)


def _no_merge_over_pending(doc: str, table: str, accepted: _Accepted, c1: int, c2: int, r1: int, r2: int,
                           action: str) -> None:
    """Round 54 (lane M): untracked, a merge or split reaching over a column whose tracked
    deletion is pending, or a cell a pending tracked merge takes in, would settle that change
    for some cells only: refused, with the way out (accept the table's changes first)."""
    rows_over = [r for r in range(r1, r2 + 1) if r not in accepted.rows] if accepted.rows_active else []
    if not accepted.active and not rows_over:
        return
    over = [c for c in range(c1, c2 + 1) if c not in accepted.kept]
    taken = [(r, c) for (r, c) in accepted.given if r1 <= r <= r2 and c1 <= c <= c2]
    if not over and not taken and not rows_over:
        return
    fix = {"tool": "word_changes", "calls": [{"action": "accept", "within": table}]}
    what = "a row deleted" if rows_over else "a column deleted" if over else "a cell merged"
    raise ToolError("refused", f"the {action} reaches over {what} by a tracked change not yet accepted in {table}; "
                    f"accept the table's changes first (word_changes accept within {table}), then {action}",
                    valid_options=[f"word_changes accept within {table}"],
                    details={"fix": fix, "call": {"tool": "word_changes",
                                                  "arguments": {"doc": doc, "action": "accept", "within": table}}})


def _limit_cells(call: Any, count: int) -> None:
    if count > 5000:
        raise ToolError("limit", f"{count} cells in one call; at most 5000")


# -- W23 word_edit_table -----------------------------------------------------------------------


CELL = obj({"row": integer(minimum=0, optional=True),
            "col": integer(minimum=0, optional=True),
            "row_label": string("Or the row's first cell text.", optional=True),
            "col_label": string("Or the column's header text.", optional=True),
            "text": string("Lines are paragraphs.", optional=True),
            # Round 45, R66: one value of a source, as a text reference names one.
            "from": string("Or one value: d:<id>/#12/<column>, :0.0 to round.", optional=True),
            # Round 48, R89: one text or value for a column's rows (:func:`_column_rows`).
            "rows": string("Or a column's rows: 1..40, or * (all but header rows).", optional=True)})
#: Round 49 (R90): sync reads from_columns as the source column of each table column.
SYNC_COLUMNS = array(string(), "from's columns, in order (sync: per table column, default by labels). "
                     "Default all.", optional=True)
#: Round 52 (R106): table key -> data key pairs no rule should guess; an object reads as pairs.
SYNC_PAIRS = array(obj({"table": string(), "data": string()}), "sync: key pairs matched first.",
                   or_map=True, optional=True)
POSITION = obj({"row": integer(minimum=0), "col": integer(minimum=0)},
               "merge, split: first cell, from 0.", optional=True)


def _locate(table: Any, cell: dict, header_row: int = 0, accepted: "_Accepted | None" = None) -> tuple[int, int]:
    """A cell's (row, grid column) by position or labels; with ``accepted``, ``col`` and
    the labels count the table as accepted (round 54, lane M)."""
    def labelled(texts: list[str], label: str, what: str) -> int:
        clean = [" ".join(t.split()) for t in texts]
        hits = [k for k, t in enumerate(clean) if t == " ".join(label.split())]
        if len(hits) != 1:
            raise ToolError("not_found" if not hits else "ambiguous", f"{what} label {label!r} matches "
                            f"{len(hits)} {what}s", valid_options=[t for t in clean if t][:50])
        return hits[0]

    first = accepted.kept[0] if accepted is not None and accepted.active and accepted.kept else 0
    # Round 54 (lane R): rows as accepted -- a row whose tracked deletion is pending not counted.
    rows = accepted.rows if accepted is not None and accepted.rows_active else list(range(len(table.rows)))
    if cell.get("row") is not None:
        row = accepted.to_row(cell["row"], "cells[].row") if accepted is not None and accepted.rows_active \
            else cell["row"]
    else:
        row = rows[labelled([_text(table, r, first) for r in rows],
                            need(cell.get("row_label"), "cells[].row or row_label"), "row")]
    if accepted is not None and accepted.active:
        if cell.get("col") is not None:
            col = accepted.to_grid(cell["col"], "cells[].col")
        else:
            col = accepted.kept[labelled(accepted.labels(header_row),
                                         need(cell.get("col_label"), "cells[].col or col_label"), "column")]
        return row, accepted.at(row, col)
    if cell.get("col") is not None:
        col = cell["col"]
    else:
        col = labelled([_text(table, header_row, c) for c in range(table.column_count)],
                       need(cell.get("col_label"), "cells[].col or col_label"), "column")
    return row, col


def _no_table(call: Any, doc: str, table: Any) -> ToolError:
    """Round 56 (R120): no table ``table`` in the document: its tables are the options (the first
    12, the rest counted), and a table of another open document (a workbook's ``t:stock``) is said to
    be one, with the call that writes its rows into a table here."""
    from ooxml_edit.tools import sources

    ids = [found.id for found in call.document.tables()]
    shown = ids[:12] + ([f"and {len(ids) - 12} more"] if len(ids) > 12 else [])
    listed = f"{doc}'s tables: {', '.join(shown)}" if ids else f"{doc} has no table"
    elsewhere = sources.resolving(table, call.session) if isinstance(table, str) else None
    if elsewhere is not None:
        handle = sources.qualify(elsewhere, dict(call.session.documents), call.session.formats)[0]
        kind = call.session.documents[handle].kind
        example = {"table": ids[0] if ids else "<id>", "action": "sync", "from": elsewhere}
        return ToolError("not_found", f"{table} is a table of {handle} ({kind}), not of {doc}: word_edit_table "
                         f"changes a table of {doc} ({', '.join(shown) or 'it has none'}); to keep one in step "
                         f"with its rows, e.g. {example}", field="table", valid_options=ids[:12])
    return ToolError("not_found", f"no table {table!r} in {doc}; {listed}", field="table", valid_options=ids[:12])


def _text(table: Any, row: int, col: int) -> str:
    try:
        return table.cell(row, col).text
    except IndexError:
        return ""


@word_tool("word_edit_table", "Change a table's cells and structure: set cells by position or labels, insert or delete rows and columns, merge, split, sync rows by key. Returns it as Markdown.",
           {"doc": DOC, "table": TABLE,
            "action": string("What to do.", enum=["set_cells", "insert_row", "delete_row", "insert_column",
                                                   "delete_column", "merge", "split", "sync"]),
            "cells": array(CELL, "set_cells: the cells.", optional=True),
            "from": string("set_cells, sync: rows of d:<id> or t:<table>, from row index.", optional=True),
            "from_columns": SYNC_COLUMNS,
            # Round 49 (R90): sync's key column and what becomes of rows whose key is gone.
            "key": string("sync: key column label or index. Default first.", optional=True),
            "remove": boolean("sync: delete rows of gone keys.", optional=True),
            # Round 52 (R106): keys spelled differently (ooxml_edit.tools.loose_key), and pairs.
            "loose": string("sync: also match keys loosely (case, spacing, separators, 05=5); whose "
                            "spelling stays.", enum=["table", "data"], optional=True),
            "match": SYNC_PAIRS,
            "header": HEADER,
            "date_format": DATE_FORMAT,
            "index": integer("insert_*, delete_*: row or column from 0. Default the last. "
                             "set_cells from: first row.", minimum=0, optional=True),
            # Round 40: rows added for the data, each like the last row (insert_row's).
            "grow": boolean("set_cells from: add rows the data needs.", optional=True),
            "side": string("insert_*: relative to index. Default after.", enum=["after", "before"], optional=True),
            "first": POSITION, "last": POSITION,
            "rows": integer("split: into rows. Default 1.", minimum=1, maximum=63, optional=True),
            "columns": integer("split: into columns. Default 2.", minimum=1, maximum=63, optional=True)},
           refs=('table',), group="word_objects", mutates=True, sources=("from", "cells[].from"),
           facts=("cells[].text",), aliases={"columns": names_as("columns", "from_columns")},
           # Round 51: key is sync's key column, never a retry key -- a sync repeated with the
           # same key column (of this table or another) runs again.
           retry_key=False)
def word_edit_table(call: Any, doc: str, table: str, action: str, cells: list[dict] | None = None,
                    index: int | None = None, side: str = "after", first: dict | None = None,
                    last: dict | None = None, rows: int = 1, columns: int = 2,
                    from_columns: list[str] | None = None, header: list[str] | None = None,
                    date_format: str | None = None, grow: bool | None = None, key: str | None = None,
                    remove: bool | None = None, loose: str | None = None, match: list[dict] | None = None,
                    **given: Any) -> Result:
    document = call.document
    table = resolve(call, table)
    try:
        view = document.table(table)
    except KeyError as error:
        if type(error) is not KeyError:     # an AddressError: says what the id is, with options
            raise
        raise _no_table(call, doc, table) from None
    source = given.get("from") or None  # "": a strict decoder's placeholder for none (rule 1)
    if action == "sync":
        return _sync(call, view, source, cells, index, from_columns, header, date_format, grow, key, remove,
                     loose, match)
    for name, value in (("key", key), ("remove", remove), ("loose", loose), ("match", match)):
        if value not in (None, "", False, []):
            raise ToolError("invalid_arguments", f"{name} is for sync (a table kept in step with rows by key)",
                            field=name, valid_options=["sync"])
    read_as = None
    if source is not None:
        if action != "set_cells":
            raise ToolError("invalid_arguments", f"{action} takes no from; from is for set_cells",
                            field="from", valid_options=["set_cells"])
        from . import field_text

        if field_text.names_one_value(source):
            # Round 45, R66: one value is a cell's, not rows to fill from.  Round 49, R96: a
            # cell with its own from (or text) keeps it; the others get this one.
            cells, filled = _one_value_cells(source, cells, action)
            own = len(cells) - filled
            read_as = (f"from {source} (one value) read as the from of each of the {filled} cell(s) given"
                       + (f" without their own; {own} kept their own" if own else ""))
            if not filled:
                read_as = (f"from {source} (one value) filled no cell: each of the {own} cell(s) gives its "
                           "own from or text, which it keeps")
            source = None
        elif cells is not None:
            raise _rows_beside_cells(source)
    if source is not None:
        asked = {"doc": doc, "table": table, "action": action, "from": source, "from_columns": from_columns,
                 "header": header, "date_format": date_format, "index": index}
        return _set_from(call, view, source, from_columns, header, index or 0, date_format,
                         grow=bool(grow), asked={k: v for k, v in asked.items() if v is not None})
    no_source(from_columns, header, date_format)
    if grow:
        raise ToolError("invalid_arguments", "grow adds the rows a source's data needs: it is for set_cells "
                        "with from; for rows of your own, insert_row", field="grow",
                        valid_options=["from", "insert_row"])
    kept = _Kept(document, table)
    accepted = _Accepted(document, table)
    edits = []
    watch: Watch | None = None          # round 57 (R122): the file's text a deletion removes
    facts: list[dict[str, Any]] | None = None
    fills: list[dict[str, Any]] = []
    held: set[str] = set()
    if action == "set_cells":
        cells = need(cells, "cells")
        # Round 48 (R89): an entry with rows is a cell per row of its column, one value for all.
        places, owners, fills = _column_rows(document, view, cells, accepted)
        held = _paragraphs_of(document, table) if fills else set()
        _limit_cells(call, len(places))
        values = [_cell_value(call, cell, k) for k, cell in enumerate(cells)]   # all read first, once each
        texts = [values[k][0] for k in owners]
        placed_at: list[tuple[int, int]] = []
        writes = _by_position(document, view, places, texts, accepted)
        if writes is not None:
            # Round 45 (R67): every cell by its position, each on the table -- in one pass, the
            # same edits in the same order (a thousand cells took 80 s one by one).
            for row, col, _ in writes:
                kept.cell(row, accepted.from_grid(col))
                placed_at.append((row, col))
            edits.extend(document._fill_cells(table, writes)[0])
        else:
            for cell, text in zip(places, texts):  # in order: a label set by an earlier cell is found by a later one
                current = document.table(table)
                row, col = _locate(current, cell, accepted=accepted)
                kept.cell(row, accepted.from_grid(col))
                edits.extend(_set_cell(document, current, row, col, text))
                placed_at.append((row, col))
        if fills:
            _count_fills(fills, values)
        facts = _cell_facts(call, cells, values, placed_at, lambda: kept.current_id() or table, owners,
                            accepted)
    elif action == "insert_row":
        # Round 54 (lane R): beside the row named, as accepted -- a row whose deletion is
        # pending stays on its other side.
        if accepted.rows_active:
            index = accepted.to_row(accepted.height - 1 if index is None else index, "index")
        edits.append(document.insert_row(table, index, below=side == "after"))
    elif action == "delete_row":
        index = need(index, "index")
        if accepted.rows_active:
            index = accepted.to_row(index, "index")
        _no_reshape_over_merge_lead(document, doc, table, accepted, action, index)
        watch = Watch(document)
        watch.row(table, index)
        edits.append(document.delete_row(table, index))
    elif action == "insert_column":
        # Round 54 (lane M): beside the column named, as accepted -- a column whose deletion is
        # pending stays on its other side (Word's Accept All then has the new one there).
        col = accepted.width - 1 if index is None else index
        edits.append(document.insert_column(table, accepted.to_grid(col, "index"), right=side == "after"))
        index = col
    elif action == "delete_column":
        grid = accepted.to_grid(need(index, "index"), "index")
        _no_reshape_over_merge_lead(document, doc, table, accepted, action, grid)
        watch = Watch(document)
        watch.column(table, grid)
        edits.append(document.delete_column(table, grid))
    elif action == "merge":
        a, b = need(first, "first"), need(last, "last")
        (r1, c1), (r2, c2) = _corner(accepted, a, "first"), _corner(accepted, b, "last")
        if document._active_tracking() is None:
            _no_merge_over_pending(doc, table, accepted, c1, c2, r1, r2, "merge")
        # Round 55 (R115): tracked, a merge over pending tracked changes is written in the
        # form Word's Accept All and Reject All settle exactly: a column or row whose
        # deletion is pending left out, a pending merge it reaches taken in whole.
        edits.append(document.merge_cells(table, (r1, c1), (r2, c2)))
    else:
        a = need(first, "first", " (the cell to split)")
        r1, c1 = _corner(accepted, a, "first")
        _no_split_over_pending(document, doc, table, accepted, r1, c1)
        edits.append(document.split_cell(table, r1, c1, rows=rows, columns=columns))
    kept.follow(call, action, index, side)
    # A row inserted or deleted at the top renames the table (its id is its first row's).
    now = kept.current_id() or table
    # Round 48 (R89): a column filled is said by its range and counts, not the table's text.
    data = table_json(document, now, markdown=not fills)
    if fills:
        data["filled"] = [_fill_brief(fill) for fill in fills]
    if facts is not None:
        data["facts"] = facts
    # Round 55 (R115, rule 3): a merge already made -- pending tracked too -- changes nothing.
    unchanged = action == "merge" and edits and not edits[-1].changed
    result = outcome(edits, f"{action} on {table}", data=data, changed=[] if unchanged else [now])
    if unchanged:
        result.warnings.insert(0, "nothing changed: the cells are merged already (a pending tracked merge "
                                  "counts as made)")
    if read_as:
        result.warnings.insert(0, read_as)
    if watch is not None:
        add_original(result, watch.report(call))
    result.warnings.extend(accepted.notes)
    result.warnings.extend(_loses_text(document, now))
    if fills:
        # The cells' paragraphs (changed, made and removed) are the table's: named by it, not
        # one by one (rule 4; a 431-row column listed 431 ids).
        inside = held | _paragraphs_of(document, now)
        for key in ("changed", "created", "removed"):
            kept_ids = [i for i in getattr(result, key) if i not in inside and not i.startswith(f"{now}/")]
            setattr(result, key, ([now] if key == "changed" else []) + [i for i in kept_ids if i != now])
        for fill in fills:
            result.warnings.extend(fill["notes"])
    if now != table:
        result.renamed[table] = now
    return note_small_text(call, result, [now]) if action == "set_cells" else result


def _sync(call: Any, view: Any, source: str | None, cells: Any, index: Any, from_columns: Any, header: Any,
          date_format: Any, grow: Any, key: Any, remove: Any, loose: Any = None, match: Any = None) -> Result:
    """Round 49 (R90): ``sync`` (:mod:`.table_sync`), its arguments read as meant (rule 1)."""
    from .table_sync import sync

    if source is None:
        raise ToolError("invalid_arguments", "sync takes from: the rows to keep the table in step with, "
                        "e.g. {action: 'sync', from: 'd:<id>', key: '<key column label>'}", field="from")
    notes: list[str] = []
    if cells:
        raise ToolError("invalid_arguments", "sync writes the cells from: cells are for set_cells; leave them "
                        "out (from_columns picks the columns)", field="cells", valid_options=["from_columns"])
    if header:
        raise ToolError("invalid_arguments", "sync leaves the header rows as they are: header is for set_cells "
                        "from; from_columns names the source column of each table column", field="header",
                        valid_options=["from_columns"])
    if index not in (None, 0):
        notes.append(f"index {index} not used: sync places each row by its key")
    if grow:
        notes.append("grow not needed: sync adds the rows of new keys")
    return sync(call, view, source, from_columns, key, bool(remove), date_format, notes, loose=loose or None,
                match=[(p["table"], p["data"]) for p in match or ()])


def _by_position(document: Any, view: Any, cells: list[dict], texts: list[str],
                 accepted: "_Accepted | None" = None) -> list[tuple[int, int, str]] | None:
    """``set_cells``' cells as ``(row, column, text)`` where each names its cell by position
    and the table has a cell there, and nothing needs stamping -- what
    ``Document._fill_cells`` writes as the cells one by one would be; else ``None`` (a
    label, read from text an earlier cell may set, or a cell to refuse: one by one)."""
    if any(not isinstance(c.get("row"), int) or isinstance(c.get("row"), bool) or not isinstance(c.get("col"), int)
           or isinstance(c.get("col"), bool) for c in cells) or not all(isinstance(t, str) for t in texts):
        return None
    grid = view.grid()
    at = [(c["row"], c["col"]) for c in cells]
    if accepted is not None and (accepted.active or accepted.rows_active):
        # Round 54 (lanes M, R): positions count the table as accepted; one out of range is
        # refused cell by cell (with the table's shape).
        if any(not 0 <= row < accepted.height for row, _ in at):
            return None
        at = [accepted.place(row, col, "cells[]") for row, col in at]
    for row, col in at:
        if not (0 <= row < len(grid) and 0 <= col < len(grid[row]) and grid[row][col] is not None):
            return None
    if document._stamping_due():
        return None
    return [(row, col, text) for (row, col), text in zip(at, texts)]


#: A cells entry's ``rows`` (round 48, R89): ``*`` (or ``all``), ``a..b``, ``a..``, ``..b`` or
#: ``a``, rows from 0 as ``row`` counts them, both ends included (``a-b`` and ``a:b`` too).
_ROWS = re.compile(r"^\s*(?:(?P<all>\*|all)|(?P<a>\d+)?\s*(?P<to>\.\.|-|:)?\s*(?P<b>\d+)?)\s*$", re.IGNORECASE)
#: Rows a column's result lists by number where some were not written (rule 4).
_LISTED = 20


def _header_rows(document: Any, table: str) -> int:
    """The table's header rows, as a column's body leaves them out: the rows marked to repeat
    as a header (``w:tblHeader``), else the first row where the table's look has its header
    row on (as Word draws it) -- as ``style_by_value`` reads a column (:mod:`.scales`)."""
    properties = document.table_properties(table)
    count = properties.get("header_rows") or 0
    if not count and (properties.get("look") or {}).get("first_row"):
        count = 1
    return count


def _column_rows(document: Any, view: Any, cells: list[dict],
                 accepted: "_Accepted | None" = None) -> tuple[list[dict], list[int], list[dict]]:
    """``set_cells``' cells as places, one per cell to write, each with the index of the
    entry it comes from, and the column fills (round 48, R89).

    An entry with ``rows`` is its column's cells over those rows, each given the entry's
    one value: a merged cell once (where its rows begin, or the range's first row for one
    merged from above), a row without a cell at that column left out and said.  ``*`` and an
    open start leave the header rows (:func:`_header_rows`) out; a range that names one is
    written as asked, with a note.  Its column is ``col`` or ``col_label`` (the first row's
    text), as for one cell.  Other entries are as given."""
    if not any(isinstance(c, dict) and c.get("rows") not in (None, "") for c in cells):
        return list(cells), list(range(len(cells))), []
    grid = view.grid()
    # Round 54 (lane R): rows as accepted (a row whose tracked deletion is pending not counted).
    rows = accepted.rows if accepted is not None and accepted.rows_active else list(range(len(grid)))
    height = len(rows)
    header = _header_rows(document, view.id)
    places: list[dict] = []
    owners: list[int] = []
    fills: list[dict] = []
    for k, cell in enumerate(cells):
        said = cell.get("rows")
        if said in (None, ""):          # "": a strict decoder's placeholder for none (rule 1)
            places.append(cell)
            owners.append(k)
            continue
        col = _fill_column(view, cell, k, accepted)
        grid_col = accepted.to_grid(col, f"cells[{k}].col") if accepted is not None and accepted.active else col
        first, last, notes = _fill_range(view.id, str(said), cell, k, height, header)
        written: list[tuple[int, int]] = []
        missing: list[int] = []
        seen: set[tuple[int, int]] = set()
        for row in range(first, last + 1):
            doc_row = rows[row]
            if accepted is not None and (doc_row, grid_col) in accepted.lost:
                missing.append(row)
                continue
            at = accepted.at(doc_row, grid_col) if accepted is not None and accepted.active else grid_col
            found = grid[doc_row][at] if at < len(grid[doc_row]) else None
            if found is None:
                missing.append(row)
                continue
            origin = (found.row, found.column)
            if origin in seen:
                continue                 # merged with a row above it in the range
            seen.add(origin)
            places.append({"row": row, "col": col})
            owners.append(k)
            written.append((doc_row, at))
        if missing:
            shown = ", ".join(str(r) for r in missing[:_LISTED]) + (
                f" and {len(missing) - _LISTED} more" if len(missing) > _LISTED else "")
            notes.append(f"cells[{k}]: row(s) {shown} have no cell at column {col} (left as they are)")
        if not written:
            raise ToolError("not_found", f"cells[{k}]: rows {first}..{last} have no cell at column {col}",
                            field=f"cells[{k}].rows", valid_options=[f"{header}..{height - 1}"])
        fills.append({"entry": k, "col": col, "first": first, "last": last, "header": header,
                      "before": [grid[r][c].text for r, c in written], "missing": missing,
                      "notes": notes})
    return places, owners, fills


def _paragraphs_of(document: Any, table: str) -> set[str]:
    """The ids of the table's cells' paragraphs."""
    return {p.id for line in document.table(table).grid() for cell in line if cell is not None
            for p in cell.paragraphs}


def _fill_column(view: Any, cell: dict, k: int, accepted: "_Accepted | None" = None) -> int:
    """The column a column fill writes: ``col`` or ``col_label`` -- refused with the
    table's columns (rule 2); with ``accepted``, counted as accepted (round 54, lane M)."""
    if accepted is not None and accepted.active:
        if cell.get("col") is not None and cell.get("col_label") in (None, ""):
            accepted.to_grid(cell["col"], f"cells[{k}].col")
            return cell["col"]
        if cell.get("col_label") not in (None, ""):
            return accepted.from_grid(_locate(view, {"row": 0, "col_label": cell["col_label"]},
                                              accepted=accepted)[1])
    width = view.column_count
    if cell.get("col") is not None and cell.get("col_label") in (None, ""):
        col = cell["col"]
        if not 0 <= col < width:
            raise ToolError("not_found", f"cells[{k}].col {col}: {view.id} has grid columns 0 to {width - 1}",
                            field=f"cells[{k}].col", valid_options=[str(c) for c in range(width)][:50])
        return col
    if cell.get("col_label") in (None, ""):
        raise ToolError("invalid_arguments", f"cells[{k}] gives rows but no column: give col (from 0) or "
                        f"col_label (its header text), e.g. {{col: 2, rows: '*', text: 'Yes'}}",
                        field=f"cells[{k}].col", valid_options=["col", "col_label"])
    return _locate(view, {"row": 0, "col_label": cell["col_label"]})[1]


def _fill_range(table: str, said: str, cell: dict, k: int, height: int, header: int) -> tuple[int, int, list[str]]:
    """The first and last row (both written) of a cells entry's ``rows``, and the notes."""
    body = f"{header}..{height - 1}" if header < height else f"0..{height - 1}"
    shape = (f"{table} has rows 0 to {height - 1}" + (f" ({header} header row(s))" if header else "")
             + f": rows '*' (every row below the header), '{body}', or a range of them")
    if cell.get("row") is not None or cell.get("row_label") not in (None, ""):
        raise ToolError("invalid_arguments", f"cells[{k}] gives rows and {'row' if cell.get('row') is not None else 'row_label'}: "
                        "rows names the rows; give one (for one cell, row and no rows)",
                        field=f"cells[{k}].rows", valid_options=["rows", "row", "row_label"])
    found = _ROWS.match(said)
    if found is None or not (found["all"] or found["a"] or found["b"]):
        raise ToolError("invalid_arguments", f"cells[{k}].rows {said!r} is not rows: {shape}",
                        field=f"cells[{k}].rows", valid_options=["*", body])
    notes: list[str] = []
    if found["all"]:
        first, last = header, height - 1
    else:
        a = int(found["a"]) if found["a"] else None
        b = int(found["b"]) if found["b"] else None
        if found["to"] is None:
            first = last = a
        else:
            first = header if a is None else a
            last = height - 1 if b is None else b
        if first > last:
            raise ToolError("invalid_arguments", f"cells[{k}].rows {said!r} runs backwards: {shape}",
                            field=f"cells[{k}].rows", valid_options=[f"{last}..{first}", "*"])
        if last >= height:
            raise ToolError("not_found", f"cells[{k}].rows {said!r}: no row {last}; {shape}. Rows are not "
                            "added: insert_row first, or set_cells from with grow", field=f"cells[{k}].rows",
                            valid_options=["*", f"{first}..{height - 1}" if first < height else body])
        if a is not None and a < header:
            notes.append(f"cells[{k}]: rows {first}..{last} include header row(s) {first}.."
                         f"{min(header, last + 1) - 1}, written as asked ('*' leaves them out)")
    if first > last:                    # '*' on a table of header rows only
        raise ToolError("not_found", f"cells[{k}].rows {said!r}: {table} has no rows below its "
                        f"{header} header row(s); name them to write them: rows '0..{height - 1}'",
                        field=f"cells[{k}].rows", valid_options=[f"0..{height - 1}"])
    return first, last, notes


def _count_fills(fills: list[dict], values: list[tuple[str, Any]]) -> None:
    """Each column fill's cells written, and how many held its value already (rule 3)."""
    for fill in fills:
        text = values[fill["entry"]][0]
        fill["cells"] = len(fill["before"])
        fill["already"] = sum(1 for held in fill["before"] if held == text)
        if fill["already"] == fill["cells"]:
            fill["notes"].append(f"cells[{fill['entry']}]: the {fill['cells']} cell(s) held that text already "
                                 "(nothing to change)")


def _fill_brief(fill: dict) -> dict[str, Any]:
    """What a result says of a column fill: the range written and counts (rule 4)."""
    out: dict[str, Any] = {"entry": fill["entry"], "col": fill["col"], "rows": f"{fill['first']}..{fill['last']}",
                           "cells": fill["cells"], "changed": fill["cells"] - fill["already"]}
    if fill["header"] and fill["first"] >= fill["header"]:
        out["header_rows"] = fill["header"]       # left out above the range
    if fill["missing"]:
        out["no_cell"] = len(fill["missing"])
    return out


def _one_value_cells(source: str, cells: list[dict] | None, action: str) -> tuple[list[dict], int]:
    """Round 45, R66: a top-level ``from`` naming one value with ``cells`` that give only
    places is each place's value (rule 1, read so, with a note in ``cells[].from``'s place);
    without cells it is refused with the call that works.  Round 49, R96: a cell that gives
    its own ``from`` or ``text`` keeps it (the cell's own wins for its cell), a ``rows``
    entry included; the top-level value fills the others.  The cells, and how many it
    filled."""
    example = f"cells: [{{row: 1, col: 2, from: '{source}'}}]"
    if action != "set_cells" or not cells:
        raise ToolError("invalid_arguments", f"from {source} is one value: set_cells from fills rows (d:<id>, "
                        f"d:<id>/#9:#16); one value goes in a cell: {example}", field="from",
                        valid_options=[example, source.split("/", 1)[0]])
    out, filled = [], 0
    for cell in cells:
        if cell.get("text") not in (None, "") or cell.get("from") not in (None, ""):
            out.append(cell)
        else:
            out.append({**cell, "from": source})
            filled += 1
    return out, filled


def _rows_beside_cells(source: str) -> ToolError:
    """Round 49, R96: a top-level ``from`` naming rows (a whole dataset or table, or a run of
    rows) beside ``cells``: two forms of set_cells at once, so refused, naming what to drop
    for each intent, with the call each would be."""
    base = source.split("/", 1)[0] if source.startswith("d:") else source
    one = f"{base}/#1/<column>" if source.startswith("d:") else f"{base}/<row>/<column>"
    example = f"cells: [{{row: 1, col: 2, from: '{one}'}}]"
    return ToolError("invalid_arguments", f"from {source} names rows, which fill the table from index, and cells "
                     f"are given too: to fill the table from {source}, drop cells; to set single cells, drop the "
                     f"top-level from and give each cell its own value ({example})", field="cells",
                     valid_options=[source, example])


def _cell_value(call: Any, cell: dict, k: int) -> tuple[str, dict[str, Any] | None]:
    """Round 45, R66: a cell's text -- its ``text`` (references written already), or the one
    value its ``from`` names, as a reference writes it (a number format or date pattern
    after it), and that value's record."""
    text, source = cell.get("text"), cell.get("from")
    from . import field_text

    if source == "":
        source = None                   # a strict decoder's placeholder for no from (rule 1)
    if text == "" and source is not None:
        text = None                     # a strict decoder's placeholder beside a from (rule 1)
    if source is None:
        if text is None:
            raise ToolError("invalid_arguments", f"cells[{k}] gives no text: give text, or from (one value, "
                            "d:<id>/#12/<column>)", field=f"cells[{k}].text", valid_options=["text", "from"])
        return text, None
    if text is not None:
        raise ToolError("invalid_arguments", f"cells[{k}] gives text and from: give one", field=f"cells[{k}].from",
                        valid_options=["text", "from"])
    try:
        return call.fact(source, field=f"cells[{k}].from")
    except ToolError as error:
        field_text.with_rows(call, source, f"cells[{k}].from", error)
        raise


def _cell_facts(call: Any, cells: list[dict], values: list[tuple[str, Any]], placed_at: list[tuple[int, int]],
                table: Any, owners: list[int] | None = None,
                accepted: "_Accepted | None" = None) -> list[dict[str, Any]] | None:
    """The values set_cells wrote by reference -- a cell's ``from`` and references in a
    cell's ``text`` -- each with its cell, in order; a ``from`` kept as its cell's
    paragraphs' source.  ``None`` without any (the result as before).

    Round 48 (R89): ``placed_at`` lists the cells written, ``owners`` the entry each came
    from; an entry written in many cells (``rows``) is said once, with its rows and count,
    and its value kept in each cell -- a reference in its text too, which the dispatcher's
    keeping (:func:`.textfacts.keep_written`) leaves where the same text is in many places."""
    from . import field_text

    owners = list(range(len(cells))) if owners is None else owners
    by_cell: dict[int, list[Any]] = {}
    for fact in getattr(call, "facts", None) or ():
        found = re.match(r"cells\[(\d+)\]\.text$", fact.field)
        if found:
            by_cell.setdefault(int(found.group(1)), []).append(fact)
    if not by_cell and all(record is None for _text, record in values):
        return None
    document = call.document
    view = document.table(table())     # its id now (a first row's change renames it)
    grid = view.grid()
    where: dict[int, list[tuple[int, int]]] = {}
    for k, at in zip(owners, placed_at):
        where.setdefault(k, []).append(at)
    out: list[dict[str, Any]] = []
    keeps: list[tuple[str, list[dict[str, Any]]]] = []
    for k, (text, record) in enumerate(values):
        if record is None and k not in by_cell:
            continue
        found_cells = [grid[r][c] for r, c in where.get(k, ())]
        many = cells[k].get("rows") not in (None, "")
        if many:
            def position(row: int) -> int:
                return accepted.from_row(row) if accepted is not None and accepted.rows_active else row

            rows = f"{position(found_cells[0].row)}..{position(found_cells[-1].row)}" if found_cells else ""
            col = placed_at[owners.index(k)][1]
            said = {"col": accepted.from_grid(col) if accepted is not None else col, "rows": rows,
                    "cells": len(found_cells)}
        else:
            said = {"cell": found_cells[0].id}
        paragraphs = [[p.id for p in cell.paragraphs] for cell in found_cells]
        if record is not None:
            fact = field_text.as_fact(cells[k]["from"], f"cells[{k}].from", text, record)
            for ids in paragraphs:
                keeps += _whole_paragraphs(document, record, ids)
            out.append(_said_fact(field_text.brief(fact), said))
        for fact in by_cell.get(k, ()):
            if many:
                keeps += _in_each(fact, cells[k].get("text"), paragraphs)
            out.append(_said_fact(field_text.brief(fact), said))
    if keeps:
        # Round 48: kept in one edit -- one per cell followed every value kept again.
        document.keep_text_facts_in(keeps)
    return out


def _said_fact(brief: dict[str, Any], said: dict[str, Any]) -> dict[str, Any]:
    """A fact's brief with where it went in place of ``cell``, before its sources."""
    out = {k: v for k, v in brief.items() if k in ("from", "text", "chars")}
    out.update(said)
    out.update({k: v for k, v in brief.items() if k in ("sources", "stale")})
    return out


def _whole_paragraphs(document: Any, record: dict[str, Any], ids: list[str]) -> list[tuple[str, list[dict]]]:
    """A value's record to keep as the source of each paragraph of ``ids`` with text, its
    whole text the part written (as :func:`.field_text.keep_in_paragraphs` keeps it)."""
    out = []
    for block in ids:
        if not block.startswith("p:"):
            continue
        try:
            text = document.paragraph(block).text
        except (KeyError, ValueError):
            continue
        if text:
            out.append((block, [{**record, "text": text, "p": 0, "at": 0}]))
    return out


def _in_each(fact: Any, text: Any, paragraphs: list[list[str]]) -> list[tuple[str, list[dict]]]:
    """A reference written in a column fill's text, to keep on each cell it went into: the
    paragraph of its line, at its place in the line."""
    if not isinstance(text, str) or text[fact.start:fact.end] != fact.text or not fact.text:
        return []
    line = text.count("\n", 0, fact.start)
    at = fact.start - (text.rfind("\n", 0, fact.start) + 1)
    return [(ids[line], [{**fact.record, "p": 0, "at": at}]) for ids in paragraphs if line < len(ids)]


def _set_from(call: Any, view: Any, source: str, from_columns: list[str] | None,
              header: list[str] | None, top: int, date_format: str | None = None, *,
              grow: bool = False, asked: dict[str, Any] | None = None) -> Result:
    """``set_cells`` from a source (:mod:`.table_data`): the block written from row ``top``,
    column 0, each cell's text rewritten as ``set_cells`` rewrites it.

    Round 40: with ``grow``, a block longer than the table's rows from ``top`` adds the rows
    it needs below the last, each as ``insert_row`` makes one (the last row's properties and
    cells; not a header row), in the call's undo step; ``top`` may then be the row after
    the last.  Without it the refusal gives the call with ``grow``."""
    document = call.document
    timing = Timing()            # round 45 (R67): the result says where the time went
    call.touch(timing)
    with timing.part("resolve"):
        block = placed(call, source, from_columns, header, date_format)
    started = time.perf_counter()
    # Round 54 (lane M): the block's columns are the table's as accepted.
    accepted = _Accepted(document, view.id)
    # Round 54 (lane R): and its rows (a row whose tracked deletion is pending not counted).
    height, width = accepted.height, accepted.width
    if top > height or (top == height and not grow):
        raise ToolError("not_found", f"no row {top}; {view.id} has {height}" + (
            f" (with grow, index {height} adds the rows after the last)" if top == height else ""),
                        field="index", valid_options=["grow"] if top == height else ())
    need_rows, need_columns = block.size
    short_by = need_rows - (height - top)
    if short_by > 0 and not grow:
        _no_room(block, view.id, height - top, top, short_by, asked or {})
    if grow:
        fits(block, need_rows, width, room=f"{view.id} has {width} columns")   # its columns
        fits(block, MAX_ROWS - top, MAX_COLUMNS)                               # Word's rows
    else:
        fits(block, height - top, width, room=f"{view.id} has {height - top} rows from row {top}")
    grid = _accepted_grid(view.grid(), accepted)
    if accepted.rows_active:
        grid = [_positioned(grid[r], accepted.from_row) for r in accepted.rows]
    _one_cell_each(view.id, block, grid, top, min(need_rows, height - top), need_columns, {})
    edits: list[Any] = []
    added: list[str] = []
    # Round 45 (R67): in one pass where nothing needs stamping -- the same edits, in order.
    at_once = not document._stamping_due()
    if short_by > 0:
        # The rows added are the last row's likeness: checked as the block's other rows are.
        _one_cell_each(view.id, block, [grid[-1]] * short_by, 0, short_by, need_columns, {}, first=height)
        timing.add("check", time.perf_counter() - started)
        with timing.part("rows"):
            edits, added = _grow(document, view.id, short_by, at_once=at_once)
    else:
        timing.add("check", time.perf_counter() - started)
    changed = [view.id]
    if accepted.rows_active and added:
        accepted = _Accepted(document, view.id)          # the rows added are the last ones
    with timing.part("cells"):
        writes = [(accepted.rows[row] if accepted.rows_active else row, accepted.kept[col] if accepted.active
                   else col, cell_text(lines))
                  for row, values in enumerate(block.lines(), start=top) for col, lines in enumerate(values)]
        if at_once:
            made, cells = document._fill_cells(view.id, writes)
            edits.extend(made)
            changed.extend(cells)
        else:
            for row, col, text in writes:
                current = document.table(view.id)
                edits.extend(_set_cell(document, current, row, col, text))
                changed.append(current.cell(row, col).id)
    with timing.part("record"):
        placed_from = block.provenance(top)
        _keep_placed(call, view.id, block, placed_from, top)
    started = time.perf_counter()
    data = {**table_json(document, view.id, markdown=False), "from": placed_from, "timing": timing.parts}
    # The rows added are the result's created ones (tr:...).
    result = outcome(edits, f"set_cells on {view.id} from {block.rows.address}"
                     + (f"; {len(added)} row(s) added after row {height - 1}, each like it" if added else ""),
                     data=data, changed=changed)
    below = height - (top + need_rows)
    if below > 0 or need_columns < width:
        line = (f"from filled rows {top}-{top + need_rows - 1}, columns 0-{need_columns - 1} of a "
                f"{height}x{width} table; the other cells keep their text")
        if below > 0:
            line += (f" (rows {top + need_rows}-{height - 1} below the data: delete_row index "
                     f"{top + need_rows}, {below} time(s), removes them)")
        result.warnings.append(line)
    if accepted.active and need_columns and accepted.kept[need_columns - 1] != need_columns - 1:
        result.warnings.append(f"{accepted.shape()}: the block's columns 0 to {need_columns - 1} are written "
                               "in those")
    result = note_small_text(call, result, [view.id])
    timing.add("result", time.perf_counter() - started)
    return result


def _positioned(line: list, position: Any) -> list:
    """A grid row's cells, each answering its (first) row as a position among the accepted
    rows (round 54, lane R)."""
    out = []
    for cell in line:
        if cell is None:
            out.append(None)
            continue
        found = _AcceptedCell(cell, cell.column)
        found.row = position(cell.row)
        out.append(found)
    return out


def _accepted_grid(grid: list, accepted: _Accepted) -> list:
    """``grid`` (rows of cells, as ``Table.grid``) with the columns as accepted: a column a
    pending tracked deletion removes left out, and a position a pending tracked merge takes
    in answering as the merged cell (round 54, lane M)."""
    if not accepted.active:
        return grid
    out = []
    for r, line in enumerate(grid):
        row = []
        for c in accepted.kept:
            cell = line[accepted.given.get((r, c), c)] if c < len(line) else None
            row.append(_AcceptedCell(cell, accepted.from_grid(cell.column)) if cell is not None else None)
        out.append(row)
    return out


class _AcceptedCell:
    """A cell of :func:`_accepted_grid`: its origin's column counted as accepted."""

    def __init__(self, cell: Any, column: int) -> None:
        self._cell = cell
        self.row = cell.row
        self.column = column

    def __getattr__(self, name: str) -> Any:
        return getattr(self._cell, name)


def _one_cell_each(table: str, block: Any, grid: list, top: int, count: int, columns: int,
                   seen: dict[tuple[int, int], tuple[int, int]], first: int | None = None) -> None:
    """Refuse a block over a merged cell, or where a row has no cell: rows ``top`` to
    ``top + count - 1`` of ``grid`` (named from ``first`` when they are rows to add)."""
    for k in range(count):
        row = top + k
        for col in range(columns):
            cell = grid[row][col] if col < len(grid[row]) else None
            origin = (row if first is not None else cell.row, cell.column) if cell is not None else None
            if origin is None or origin in seen:
                named = row if first is None else first + k
                what = (f"merged with ({seen[origin][0]}, {seen[origin][1]})" if origin in seen
                        else "not a cell")
                raise ToolError("refused", f"({named}, {col}) of {table} is {what}"
                                + (" in the rows grow adds (each like the last row)" if first is not None else "")
                                + f": the block from {block.rows.address} needs one cell per value; split it "
                                "first", field="from")
            seen[origin] = (row, col)


def _no_room(block: Any, table: str, room: int, top: int, short_by: int, asked: dict[str, Any]) -> None:
    """The refusal of a block longer than the table's rows from ``top``: the call with
    ``grow`` that writes it (round 40), or a row range of the source that fits."""
    from ooxml_edit.tools import sources

    header = " with the header" if block.header else ""
    fewer = sources.first_rows(block.rows.address, room - (1 if block.header else 0)) if room > 0 else None
    meant = {**asked, "grow": True}
    raise ToolError("refused", f"{block.rows.address} gives {block.size[0]:,} rows{header}; {table} has "
                    f"{room} rows from row {top}: give grow: true to add the {short_by} row(s) it needs, "
                    "each like the last row" + (f", or place fewer with {fewer}" if fewer else "")
                    + " (nothing is cut)", field="from", valid_options=["grow"] + ([fewer] if fewer else []),
                    details={"call": {"tool": "word_edit_table", "arguments": meant}})


def _grow(document: Any, table: str, count: int, *, at_once: bool = False) -> tuple[list[Any], list[str]]:
    """``count`` rows after the table's last, each as ``insert_row`` makes one -- the last
    row's properties and cells, empty -- and not a header row where the last is one.
    ``at_once`` (round 45): in one pass (``Document._rows_after_last``), the same edits."""
    if at_once:
        edits = document._rows_after_last(table, count)
        return edits, [edit.id for edit in edits if edit.created]
    edits, added = [], []
    for _ in range(count):
        last = len(document.table(table).rows) - 1
        edit = document.insert_row(table, last, below=True)
        edits.append(edit)
        added.append(edit.id)
        if _is_header_row(document, table, last + 1):
            edits.append(document.set_row(table, last + 1, header=False))
    return edits, added


def _is_header_row(document: Any, table: str, row: int) -> bool:
    entry = _table_entry(document, table)
    if entry is None or row >= len(entry.rows):
        return False
    return entry.rows[row][0].find(f"{_W}trPr/{_W}tblHeader") is not None


# -- the table's record, kept true ---------------------------------------------------------------
#
# The record kept in the document (:mod:`docx_agent.edit.provenance`) counts the table's rows
# with its tracked changes accepted: a row deleted as a tracked change is not counted.  An
# edit without ``from`` keeps it true or drops what it no longer says -- a stale address is
# worse than none (docs/TOOLS.md, "Source addresses in the saved file"):
#
# * set_cells over the block's cells (its header row included) drops those columns' entries;
# * a row inserted or deleted above the block shifts it, below it leaves it, inside it drops
#   the record (the rows placed are no longer one block);
# * a column inserted shifts the entries right of it; a column deleted drops its entry and
#   shifts those right of it;
# * merge and split drop the record;
# * a record left with no column entry is dropped.
#
# A table's id is its first row's: a row inserted or deleted at the top renames the table,
# and its record moves to the new id.


def _keep_placed(call: Any, table: str, block: Any, placed_from: dict[str, Any], top: int) -> None:
    """Keep ``set_cells`` ``from``'s record, its rows counted with tracked changes accepted;
    over a row deleted as a tracked change, drop the table's record instead."""
    entry = _table_entry(call.document, table)
    if entry is None:
        return
    # Round 54 (lane R): ``top`` is a position among the accepted rows already, as the
    # record counts them.
    call.keep_provenance(table, block.rows, dict(placed_from))


def _table_entry(document: Any, table: str) -> Any:
    """The table's index entry (its element and rows), or ``None``."""
    try:
        _part, entry = document._resolve(table)
    except KeyError:
        return None
    return entry if hasattr(entry, "rows") else None


class _Kept:
    """The table's record before an edit without ``from``, and what the edit makes of it."""

    def __init__(self, document: Any, table: str) -> None:
        self.document = document
        self.key = document._provenance_key(table)
        self.record = document.provenance().get(self.key) if self.key else None
        entry = _table_entry(document, table)
        self.element = entry.element if entry is not None else None
        self.accepted = _provenance.accepted_rows(entry) if entry is not None else []
        self.total = len(entry.rows) if entry is not None else 0
        self.dropped: set[int] = set()

    def _block(self) -> tuple[int, int]:
        """The block's first and last row (header row included), counted accepted."""
        record = self.record
        first = record.get("first_row", 0)
        top = record["header_row"] if record.get("header_row") is not None else first
        rows = max((_provenance.count(c.get("rows", "")) for c in record.get("columns") or []), default=0)
        return top, max(top, first + rows - 1)

    def cell(self, row: int, col: int) -> None:
        """``set_cells`` without ``from`` writes cell (``row``, ``col``)."""
        if self.record is None:
            return
        if row not in self.accepted:
            return
        top, last = self._block()
        if top <= self.accepted.index(row) <= last:
            self.dropped.add(col)

    def follow(self, call: Any, action: str, index: int | None, side: str) -> None:
        """Store what the edit leaves of the record, under the table's id after it."""
        if self.record is None:
            return
        after = self._after(action, index, side)
        key = self._key_now()
        if key != self.key:
            call.store_provenance(self.key, None)
        if key is not None and (after != self.record or key != self.key):
            call.store_provenance(key, after)

    def current_id(self) -> str | None:
        """The table's id after the edit, found by its element."""
        entry = self._entry_now()
        return entry.id if entry is not None else None

    def _key_now(self) -> str | None:
        """The table's durable id after the edit (found by its element), or ``None``."""
        entry = self._entry_now()
        return self.document._provenance_key(entry.id) if entry is not None else None

    def _after(self, action: str, index: int | None, side: str) -> dict[str, Any] | None:
        record = dict(self.record)
        columns = [dict(c) for c in record.get("columns") or []]
        if action == "set_cells":
            columns = [c for c in columns if c.get("col") not in self.dropped]
        elif action in ("insert_row", "delete_row"):
            top, last = self._block()
            if action == "insert_row":
                # the new row's place among the rows now, then among the accepted ones
                xml = self.total - 1 if index is None else index
                at = xml + 1 if side == "after" else xml
                now = self._entry_now()
                rows_now = _provenance.accepted_rows(now) if now is not None else []
                if at not in rows_now:
                    return None
                where = rows_now.index(at)
                if where > last:
                    return record
                if where > top:
                    return None
                step = 1
            else:
                if index not in self.accepted:
                    return record          # a row already deleted as a tracked change
                where = self.accepted.index(index)
                if where > last:
                    return record
                if where >= top:
                    return None
                step = -1
            for key in ("header_row", "first_row"):
                if record.get(key) is not None:
                    record[key] += step
            return record
        elif action == "insert_column":
            # the new column's grid index: beside index (the tool's default given), to its
            # right unless before
            at = index + 1 if side == "after" else index
            for c in columns:
                if c.get("col", 0) >= at:
                    c["col"] += 1
        elif action == "delete_column":
            columns = [c for c in columns if c.get("col") != index]
            for c in columns:
                if c.get("col", 0) > index:
                    c["col"] -= 1
        else:                      # merge, split
            return None
        if not columns:
            return None
        record["columns"] = columns
        return record

    def _entry_now(self) -> Any:
        """The table's index entry after the edit, found by its element.  On a draft only the
        table found counts as read (round 51): the others looked at on the way are not the
        call's, and another call may change them meanwhile."""
        if self.element is None:
            return None
        from ..blocks import unrecorded

        with unrecorded(self.document):
            found = None
            for story in self.document.stories:
                for table in story.tables:
                    entry = _table_entry(self.document, table.id)
                    if entry is not None and entry.element is self.element:
                        found = entry
                        break
                if found is not None:
                    break
        return _table_entry(self.document, found.id) if found is not None else None


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

#: ``look``'s names (Word's Table Design options) and :meth:`~docx_agent.Document.set_table`'s.
LOOK_NAMES = {"header_row": "first_row", "first_column": "first_column", "banded_rows": "banded_rows",
              "total_row": "last_row", "last_column": "last_column", "banded_columns": "banded_columns"}


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
            "width_percent": number("table: width, % of width_of.", minimum=1, maximum=100,
                                    optional=True),
            "width_of": string("Default text; body: paragraphs' lines; window: AutoFit Window.",
                               enum=["text", "body", "window"],
                               optional=True),
            "alignment": string("table: on the page.", enum=["left", "center", "right"], optional=True),
            "shading": string("Fill RRGGBB, or none.", optional=True),
            "borders": BORDERS,
            "vertical_alignment": string("cell, row, column: text position.", enum=["top", "center", "bottom"],
                                         optional=True),
            "height": number("row: height.", minimum=1, maximum=1584, optional=True),
            "height_rule": string("row. Default at_least.", enum=["at_least", "exact"], optional=True),
            "header_rows": integer("table: rows repeated on each page.", minimum=0, maximum=500,
                                   optional=True),
            "look": array(string(enum=list(LOOK_NAMES)), "table: style options on, the rest off.", optional=True),
            "cant_split": boolean("row (table: all): never split across pages.", optional=True),
            "keep_together": boolean("table: on one page; rows keep with next, last row not.", optional=True),
            "keep_with_next": boolean("table kept together: last row keeps with next block too.", optional=True)},
           refs=('table',), group="word_objects", mutates=True)
def word_format_table(call: Any, doc: str, table: str, scope: str, row: int | None = None,
                      column: int | None = None, style: str | None = None, width: float | None = None,
                      width_percent: float | None = None, width_of: str | None = None,
                      alignment: str | None = None, shading: str | None = None, borders: str | None = None,
                      vertical_alignment: str | None = None, height: float | None = None,
                      height_rule: str = "at_least", header_rows: int | None = None,
                      look: list[str] | None = None, cant_split: bool | None = None,
                      keep_together: bool | None = None, keep_with_next: bool | None = None) -> Result:
    document = call.document
    table = resolve(call, table)
    view = document.table(table)
    fill = None if shading == "none" else shading
    line = None if borders in (None, "none") else borders
    sides = {side: line for side in _TABLE_SIDES} if borders is not None else {}
    edits = []
    if scope == "table":
        values: dict[str, Any] = {}
        for key, value in (("style", style), ("alignment", alignment), ("header_rows", header_rows)):
            if value is not None:
                values[key] = value
        if width_of in ("body", "window") and width is not None:
            raise ToolError("invalid_arguments", "width_of is what width_percent is of: give width_percent "
                            "(or nothing, for 100), not width", field="width_of")
        if width_of == "body":
            values.update(_body_width(document, table, width_percent or 100, alignment))
        elif width_of == "window":
            pass                           # round 48: after the rest (fit_table_to_window), below
        elif width is not None or width_percent is not None:
            values["width"] = f"{width_percent:g}%" if width_percent is not None else width
        if shading is not None:
            values["shading"] = fill
        if borders:
            values["borders"] = sides
        if look is not None:
            # Round 24, W2: the style options as Table Design names them, all six written.
            values["look"] = {LOOK_NAMES[name]: name in look for name in LOOK_NAMES}
        if cant_split is not None:
            edits.extend(document.set_row(table, r, cant_split=cant_split) for r in range(len(view.rows)))
        if keep_together is False and keep_with_next:
            raise ToolError("invalid_arguments", "keep_with_next keeps a table kept together with the next block: "
                            "give keep_together true with it, or leave keep_with_next out", field="keep_with_next")
        if keep_together is not None:
            # Round 25, W2: the rows' keeps as Word reads them (Document.set_keep_together);
            # W5: with keep_with_next, the last row's too.
            edits.append(document.set_keep_together(table, keep_together, with_next=bool(keep_with_next)))
            if keep_with_next is False and not edits[-1].changed \
                    and not document.table_properties(table)["keep_together"]:
                # Round 27: not kept together, both off -- a last row keeping alone is released.
                released = document.set_table_keep_with_next(table, False)
                if released.changed:
                    edits[-1] = released
        elif keep_with_next is not None:
            edits.append(document.set_table_keep_with_next(table, keep_with_next))
        if not values and not edits and width_of != "window":
            raise ToolError("invalid_arguments", "nothing to set for the table",
                            valid_options=["style", "width", "width_percent", "alignment", "shading", "borders",
                                           "header_rows", "look", "cant_split", "keep_together", "keep_with_next"])
        if width_of == "body":
            # The columns shared again first, the cells' widths with them: an autofit
            # table is drawn at its cells' widths, which a table width alone leaves.
            from ..edit.table_format import share, twips

            grid = [twips(w) for w in document.table_properties(table)["columns"]]
            for column, new in enumerate(share(twips(values["width"]), grid)):
                edits.append(document.set_column_width(table, column, new / 20))
        if values:
            edits.append(document.set_table(table, **values))
        if width_of == "window":
            # Round 48: Word's AutoFit to Window, after a style or borders given with it
            # (the grid follows the style's borders and margins).
            renamed = {k: v for edit in edits if edit is not None for k, v in edit.renamed.items()}
            edits.append(document.fit_table_to_window(renamed.get(table, table), width_percent or 100))
        result = outcome(edits, f"Formatted {table}", changed=[table])
        if not values and not any(edit.changed for edit in edits):
            # Round 27, W7: asked for what the table already is (keep_together false on a
            # table not kept together): as asked, the warnings say why nothing changed.
            result.summary, result.changed = f"No change to {table}", []
        loose = _loose_paragraph_after(document, result.renamed.get(table, table)) if keep_with_next else None
        if keep_together:
            # Round 28, W8: from the layout, which a batch makes once, after its last op.
            _later_taller(call, result.renamed.get(table, table), before=loose)
        if loose is not None:
            result.warnings.append(loose)
        elif keep_together is False and document.table_properties(result.renamed.get(table, table))["keep_together"]:
            # Round 25, W5: what keeps the rows now is their styles' (a label style's keep).
            kept = document._table_entry(result.renamed.get(table, table))[1].id
            result.warnings.append(f"{kept} is still kept together by its rows' styles (they keep with the next); "
                                   "keep_together false removes only the direct keeps")
        return result
    if look is not None:
        raise ToolError("invalid_arguments", "look is the table's: give scope table", field="look")
    if keep_together is not None:
        raise ToolError("invalid_arguments", "keep_together is the table's: give scope table (cant_split keeps a "
                        "row on one page)", field="keep_together")
    if keep_with_next is not None:
        raise ToolError("invalid_arguments", "keep_with_next is the table's: give scope table (word_format keeps "
                        "a paragraph with the next)", field="keep_with_next")
    if cant_split is not None and scope not in ("row", "table"):
        raise ToolError("invalid_arguments", "cant_split is a row's: give scope row (or table, for every row)",
                        field="cant_split")
    cell_values: dict[str, Any] = {}
    if shading is not None:
        cell_values["shading"] = fill
    if borders:
        cell_values["borders"] = {k: v for k, v in sides.items() if not k.startswith("inside")}
    if vertical_alignment is not None:
        cell_values["vertical_alignment"] = vertical_alignment
    accepted = _Accepted(document, table)
    # Round 54 (lane R): rows as accepted, as word_edit_table counts them.
    body_rows = accepted.rows if accepted.rows_active else list(range(len(view.rows)))
    if scope == "row":
        index = accepted.to_row(need(row, "row"), "row") if accepted.rows_active else need(row, "row")
        if height is not None or header_rows is not None or cant_split is not None:
            row_values: dict[str, Any] = {}
            if height is not None:
                row_values.update(height=height, height_rule="exact" if height_rule == "exact" else "atLeast")
            if cant_split is not None:
                row_values["cant_split"] = cant_split
            edits.append(document.set_row(table, index, **row_values))
        positions = [(index, c) for c in range(view.column_count)]
    elif scope == "column":
        # Round 54 (lane M): column positions count the table as accepted.
        index = accepted.to_grid(need(column, "column"), "column") if accepted.active else need(column, "column")
        if width is not None:
            edits.append(document.set_column_width(table, index, width))
        positions = [(r, index) for r in body_rows if (r, index) not in accepted.lost]
    else:
        r = accepted.to_row(need(row, "row"), "row") if accepted.rows_active else need(row, "row")
        positions = [(r, accepted.at(r, accepted.to_grid(need(column, "column"), "column")) if accepted.active
                      else need(column, "column"))]
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
                        valid_options=["shading", "borders", "vertical_alignment", "width", "height", "cant_split"])
    result = outcome(edits, f"Formatted {scope} of {table}", changed=[table])
    result.warnings.extend(accepted.notes)
    return result


#: keep_together on a table taller than a page (round 25): Word does keep it, so does the tool.
TALLER_THAN_PAGE = "taller than a page: Word starts it on a new page and breaks it where the page is full"


def _later_taller(call: Any, table: str, *, before: str | None) -> None:
    """:func:`_taller_than_page`'s warning, said once the document is as the call leaves it
    (``call.defer``: after the handler, or after a batch's last op -- one layout for the
    batch, where laying the document out after each op cost its size per op, round 28 W8),
    in its place: before ``before`` (the warning said after it), else last."""
    def say(result: Result) -> None:
        engage(call)                        # the layout engine, with its deadline afresh
        taller = _taller_than_page(call.document, table)
        if taller is not None:
            at = result.warnings.index(before) if before in result.warnings else len(result.warnings)
            result.warnings.insert(at, taller)

    call.defer(say)


def _taller_than_page(document: Any, table: str) -> str | None:
    """The warning for a table kept together that is taller than a page, or ``None``.
    Decided from the document's layout after the edit (the one ``check`` and ``render``
    read, :meth:`~docx_agent.Document.layout`): a table kept together goes to the next page
    whole where it fits on one, so one still drawn on more than one page is taller than the
    page's text height (header rows drawn again count with it).  Where the document cannot
    be laid out, or the layout stops before the table, that is not known, and the warning
    says so instead."""
    try:
        placements = document.layout().where(table)
    except Exception:  # noqa: BLE001 -- no layout: said, not guessed
        placements = None
    if isinstance(placements, list) and placements and not any(p.story == "body" for p in placements):
        return None                         # a header's or a note's table: no page to break
    if not isinstance(placements, list) or not placements:
        return (f"{table} is kept together, but the document could not be laid out to it: whether it is "
                "taller than a page is not known (if it is, Word starts it on a new page and breaks it where "
                "the page is full)")
    if len({placement.page for placement in placements if placement.story == "body"}) > 1:
        return f"{table} is {TALLER_THAN_PAGE}"
    return None


def _loose_paragraph_after(document: Any, table: str) -> str | None:
    """The warning for a table kept with the next whose next block is a paragraph that does
    not keep with the next, a table after it (round 25, W5): Word keeps the table with that
    paragraph, but the table after it can still start the next page alone (measured,
    ``feedback-round25/measurements-J.md``) -- that paragraph keeps the chain going."""
    from ..edit import effective as _effective

    part, entry = document._table_entry(table)
    blocks = []
    node = entry.element.getnext()
    while node is not None and len(blocks) < 2:
        if node.tag in (_W_P, _W_TBL):
            blocks.append(node)
        node = node.getnext()
    if len(blocks) < 2 or blocks[0].tag != _W_P or blocks[1].tag != _W_TBL:
        return None
    if _effective.paragraph(document, blocks[0]).keep_with_next:
        return None
    index = document._index(part)
    paragraph, after = index.entry_for(blocks[0]), index.entry_for(blocks[1])
    if paragraph is None or after is None:
        return None
    return (f"{paragraph.id}, between {entry.id} and {after.id}, does not keep with the next: {after.id} can still "
            f"start a page without {entry.id}; word_format keep_with_next true on {paragraph.id} keeps them together")


_W_P = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
_W_TBL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"


def _body_width(document: Any, table: str, percent: float, alignment: str | None) -> dict[str, Any]:
    """``width_of: "body"`` (round 21, T5): the share of the body's lines beside the table
    (:meth:`~docx_agent.Document.body_lines`: the text width less the effective indents of
    the paragraph around it) as a width in points, and the table placed on those lines --
    left on their left edge, centred or right between their edges -- by its indent from
    the margin, aligned left: what Word draws at the lines' edges and keeps when the table
    is edited (round 21 measurements)."""
    lines = document.body_lines(table)
    width = round(lines["width"] * percent / 100, 2)
    if width < 1:
        raise ToolError("invalid_arguments", f"the body's lines beside {table} are {lines['width']:g} pt wide",
                        field="width_of")
    current = document.table_properties(table)["alignment"]
    where = alignment or ("left" if current in (None, "start", "both") else current)
    room = lines["width"] - width
    offset = {"left": 0, "center": room / 2, "right": room}.get(where, 0)
    return {"width": width, "indent": round(lines["left"] + offset, 2), "alignment": "left"}


# -- W25, W26 word_drawings (inserting a picture too) -------------------------------------------

WRAPS = ["square", "tight", "top_and_bottom", "front", "behind"]


#: The tool that finds a catalogue's assets, when the application configured one.
FIND_ASSET = "word_find_asset"


def drawings_tool(catalogue: Any = None) -> Any:
    """``word_drawings``; with the application's asset catalogue (:func:`docx_agent.tools.docx_format`'s
    ``assets``), the same tool taking ``asset``, an id :data:`FIND_ASSET` found, in place of
    ``image`` -- the protocol pptx-agent's ``ppt_add_picture`` takes it by
    (:mod:`ooxml_edit.tools.assets`)."""
    params = {"doc": DOC,
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
              "color": string("insert_picture: tint an SVG, theme colour or RRGGBB.", optional=True),
              "ref": string("insert_*: ref name for $name.", optional=True),
              "key": string("Retry key: a repeat with the same key makes nothing new.", optional=True)}
    if catalogue is not None:
        params["asset"] = string(f"insert_picture: or an asset id from {FIND_ASSET}.", optional=True)

    @word_tool("word_drawings", "Drawings: list; insert a picture from an image blob" + (" or catalogue asset" if catalogue is not None else "") + " (width keeps the aspect ratio), a text box or a shape, inline or floating; move, resize, float or inline, set wrapping and order.",
               params, refs=('target', 'at'), group="word_objects", mutates=True,
               action_roles={"list": REVIEW},
               aliases={"size_pt": "width", **positions.aliases("word_drawings", "picture", params, find=True)})
    def word_drawings(call: Any, doc: str, action: str, target: str | None = None, image: str | None = None,
                      at: str | None = None, find: str | None = None, x: float | None = None, y: float | None = None,
                      align: str | None = None, against: str | None = None, width: float | None = None,
                      height: float | None = None, wrap: str | None = None, z_order: str | None = None,
                      text: str | None = None, preset: str | None = None, alt_text: str | None = None,
                      ref: str | None = None, key: str | None = None, color: str | None = None,
                      asset: str | None = None) -> Result:
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
            if asset is not None:
                if image is not None:
                    raise ToolError("invalid_arguments", "give image or asset, not both", field="asset")
                from ooxml_edit.tools import assets

                data, _ = assets.fetch(catalogue, asset, call.session.limits, search_tool=FIND_ASSET)
                # An icon's text for screen readers, unless given: its id's last part, in words.
                alt_text = alt_text or " ".join(re.split(r"[-_\s]+", asset.rsplit("/", 1)[-1])).strip() or asset
            else:
                data = _image_blob(call, need(image, "image"))
            return _insert_picture(call, data, need(alt_text, "alt_text", " for a picture"),
                                   at, find, width, height, _floating(wrap, x, y, align, against, picture=True), ref,
                                   own, color=color)
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
        if not edit.changed:        # round 27: float a floating drawing, inline an inline one
            return outcome([edit], f"No change to {target}")
        return outcome([edit], f"{action}: {target}", changed=[target])

    return word_drawings


word_drawings = drawings_tool()


def find_asset_tool(catalogue: Any) -> Any:
    """``word_find_asset``: the application's asset catalogue searched by words; offered only
    with one (:func:`docx_agent.tools.docx_format`'s ``assets``), as pptx-agent's
    ``ppt_find_asset``."""
    from ooxml_edit.tools import assets, tool

    @tool(FIND_ASSET, "Search the brand asset catalogue (icons, logos) by words. Returns ids, names and "
               "keywords; place one with word_drawings insert_picture asset.",
               {"query": string("Words, e.g. arrow up."),
                "limit": integer("Default 10.", minimum=1, maximum=assets.MAX_RESULTS, optional=True)},
               group="word_objects", documents=())
    def word_find_asset(call: Any, query: str, limit: int | None = None) -> Result:
        found = assets.find(catalogue, query, limit or 10)
        names = ", ".join(item["id"] for item in found[:10])
        return Result(summary=f"{len(found)} asset(s)" + (f": {names}" if names else "; try other words"),
                      data={"assets": found})

    return word_find_asset


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


def _image_blob(call: Any, image: str) -> bytes:
    blob = call.blob(image)
    if not blob.mime.startswith("image/"):
        raise ToolError("invalid_arguments", f"{image} is {blob.mime}, not an image", field="image")
    return blob.data


def _insert_picture(call: Any, data: bytes, alt_text: str, at: str | None, find: str | None,
                    width: float | None, height: float | None, floating: dict[str, Any] | None,
                    ref: str | None, before: list | None = None, *, color: str | None = None) -> Result:
    if height is not None:
        raise ToolError("invalid_arguments", "a picture keeps its aspect ratio: give width", field="height")
    document = call.document
    where = _position(call, at, find)
    if isinstance(where, str) and _CELL.match(where):
        # A cell: the start of its first paragraph.
        row, column = (int(n) for n in where.rsplit("/c", 1)[1].split(","))
        cell = document.table(where.rsplit("/c", 1)[0]).cell(row, column)
        where = f"{cell.paragraphs[0].id}@0"
    if color is not None:
        from ooxml_edit.pictures import looks_like_svg

        if not looks_like_svg(data):
            raise ToolError("invalid_arguments", "only an SVG picture can be tinted: this one is a bitmap; "
                            "leave color out, or give an SVG", field="color")
        try:
            document._tint(color)
        except PictureError as error:
            raise ToolError("invalid_arguments", str(error), field="color") from None
    edit = document.insert_picture(where, data, width=width, alt_text=alt_text, color=color)
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
           refs=('control', 'at'), group="word_objects", mutates=True,
           action_roles={"list": REVIEW},
           aliases=positions.aliases("word_controls", "mark", ("at", "control"), find=True))
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
