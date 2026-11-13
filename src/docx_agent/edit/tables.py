"""Table structure: rows and columns inserted and deleted, cells merged -- untracked, and
tracked as Word tracks them or, where it does not, in the nearest form it accepts.

What Word 16.106 writes with tracking on (``tools/e3_probe.py``):

* **A row inserted**: ``w:trPr/w:ins``, each new cell's paragraph mark inserted;
* **a row deleted**: ``w:trPr/w:del``, each cell's paragraph marks and content deleted;
* **a column inserted**: the new cells, each with its paragraph mark inserted, and the grid
  widened -- no ``w:cellIns``; Word's Reject All removes such cells and their grid column;
* **a column deleted**, **cells merged** (across or down) and **a cell split**: nothing --
  Word applies them untracked.

So a deleted column is written as ``w:cellDel`` on its cells with their content deleted.
A merge (round 54, lane M: measured with Word 16.106, ``feedback-round54/probe-M``) keeps
each row's first cell as it is and makes the other cells ``w:cellDel`` *with their content
and paragraph marks kept*, their text inserted at the end of the first cell -- merged down,
``w:cellMerge`` (``rest``/``cont``) on the rows' first cells, a continuing row's first cell's
content deleted.  Word's Accept All gives a deleted cell that keeps content or a mark to the
cell before it (its columns and width; its content goes) and removes one whose content and
marks are all deleted; Reject All keeps the cell.  So Word's own Accept All / Reject All and docx-agent's
give the same table (every glyph and rule as Word draws it, across, down, a rectangle, a
header row, an empty cell), and the pending table stays on its grid.  (The earlier form --
the first cell's span applied, the others' content deleted, every cell's old properties and
the grid recorded -- Word accepted to a fifth grid column: 54 px off.)
A split, a column across a span and a merge's first row deleted (E5) are written with the
new cells' and rows' marks inserted, every cell's old properties recorded, and a
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
from .ids import TableEntry

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_TR = _W + "tr"
W_TC = _W + "tc"
W_P = _W + "p"
W_PPR = _W + "pPr"
W_TBL = _W + "tbl"
_CELL_RECORDS = (_W + "cellIns", _W + "cellDel", _W + "cellMerge", _W + "tcPrChange")


def _int(node: Element | None, default: int) -> int:
    try:
        return int(node.get(_W + "val")) if node is not None else default
    except (TypeError, ValueError):
        return default


def _accept_ends(row: Element) -> tuple[dict[Element, int], set[Element]]:
    """What Word's Accept All makes of a row's cells whose last paragraph mark is deleted
    (:func:`docx_agent.revisions.edges.cell_ends`): each cell taken into the one before it,
    to that cell's first grid column; and each cell removed with its text (keyed by the
    elements, which keeps their proxies)."""
    from ..revisions.edges import cell_deleted, cell_ends

    shown = [(cell, start) for cell, start, _span in _cells(row) if not cell_deleted(cell)]
    groups, gone = cell_ends([cell for cell, _ in shown])
    taken = {shown[n][0]: shown[g[0]][1] for g in groups for n in g[1:]}
    lost = {shown[n][0] for n in gone if shown[n][0] not in taken}
    return taken, lost


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


_STORIES = frozenset(_W + name for name in ("body", "hdr", "ftr", "footnote", "endnote", "comment", "txbxContent"))


def _row_fate(row: Element, author: str) -> str:
    """How a tracked deletion treats ``row`` (:meth:`TableOps._track_delete_rows`):
    ``remove`` (wholly one's own insertion), ``relocate`` (the same, holding another
    author's insertion), ``keep`` (one's own row holding what was there before: its
    content deleted, its insertion kept) or ``mark`` (``w:trPr/w:del``), for a deletion by
    ``author``."""
    from ..revisions import track as _track

    def mine(record: Element) -> bool:
        return record.get(_W + "author") == author

    record = row.find(f"{_W}trPr/{_W}ins")
    if record is None or not mine(record) or row.find(f"{_W}trPr/{_W}del") is not None:
        return "mark"
    if any(True for _ in row.iter(W_TBL)):
        return "keep"
    others = False
    for paragraph in row.iter(W_P):
        mark = _track.mark_record(paragraph)
        if mark is None or mark.tag != _W + "ins" or not mine(mark):
            return "keep"
        for run in paragraph.iter(_W + "r"):
            inserted = None
            for ancestor in run.iterancestors():
                if ancestor is paragraph:
                    break
                if ancestor.tag == _W + "ins":
                    inserted = ancestor
                    break
            if inserted is None:
                return "keep"
            others = others or not mine(inserted)
    return "relocate" if others else "remove"


def _own_inserted_cell(cell: Element, author: str) -> bool:
    """Round 57: whether ``cell`` is wholly ``author``'s pending insertion -- every paragraph
    mark inserted by them, every run inside their insertion, no nested table, no cell
    record (a deletion, a merge) on it."""
    from ..revisions import track as _track

    if any(True for _ in cell.iter(W_TBL)) or any(
            cell.find(f"{_W}tcPr/{tag}") is not None for tag in _CELL_RECORDS[:3]):
        return False
    paragraphs = list(cell.iter(W_P))
    for paragraph in paragraphs:
        mark = _track.mark_record(paragraph)
        if mark is None or mark.tag != _W + "ins" or mark.get(_W + "author") != author \
                or paragraph.find(f"{_W}pPr/{_W}rPr/{_W}del") is not None:
            return False
        for run in paragraph.iter(_W + "r"):
            inserted = next((a for a in run.iterancestors(_W + "ins", _W + "del", W_P)), None)
            if inserted is None or inserted.tag != _W + "ins" or inserted.get(_W + "author") != author:
                return False
    return bool(paragraphs)


def _paragraph_after(row: Element, going: set[int]) -> Element | None:
    """The first paragraph after ``row`` in its story, in document order, that is not in a
    row of ``going`` (rows removed outright)."""
    node = row
    while node is not None and node.tag not in _STORIES:
        for sibling in node.itersiblings():
            if not isinstance(sibling.tag, str):
                continue
            for paragraph in ([sibling] if sibling.tag == W_P else sibling.iter(W_P)):
                if not any(id(r) in going for r in paragraph.iterancestors(W_TR)):
                    return paragraph
        node = node.getparent()
    return None


def _delete_cell_mark(paragraph: Element, stamp, part: str, *, own: bool) -> None:
    """A deleted row's cell paragraph mark: deleted -- beside another author's insertion of
    it, as Word records it (measured); one's own inserted mark is left (``own``: replaced by
    the deletion, in a row one inserted and keeps)."""
    from ..revisions import track as _track

    record = _track.mark_record(paragraph)
    if record is None:
        _track.mark(paragraph, "del", stamp, part)
        return
    if record.tag != _W + "ins" or paragraph.find(f"{_W}pPr/{_W}rPr/{_W}del") is not None:
        return
    if not stamp.same(record):
        _track.mark(paragraph, "del", stamp, part)
    elif own:
        remove(record)
        _track.mark(paragraph, "del", stamp, part)


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


#: The width of a grid column added for a cell past the grid that states none, twips: what
#: Word gives it (docx2svg's ``table.UNSIZED_COLUMN``, measured).
UNSIZED_COLUMN = 360


def fit_grid(table: Element) -> bool:
    """Make ``table``'s grid fit its rows (round 53, R107) as Word writes such a table when
    it saves it (measured, docx2svg's ``make_grid_mismatch_probe.py``): a grid column added
    for each one a row reaches past it -- the part of its cell's width (``w:tcW`` in
    ``dxa``) past the grid's last line, else 360 twips -- and a row reaching fewer columns
    ends with ``w:gridAfter`` for the rest; grid columns past every row's reach go.  Whether
    it changed anything."""
    grid = table.find(_W + "tblGrid")
    if grid is None:
        grid = make("w:tblGrid")
        insert_in_order(table, grid)
    widths = _grid_widths(table)
    rows = own_table_rows(table)
    reaches = []
    for row in rows:
        before, after = _skips(row)
        cells = _cells(row)
        end = (cells[-1][1] + cells[-1][2]) if cells else before
        reaches.append((row, end, after))
    if not reaches:
        return False
    longest = max(end + after for _, end, after in reaches)
    if all(end + after == len(widths) for _, end, after in reaches):
        return False
    for row, end, after in reaches:
        for cell, start, span in _cells(row):
            if start + span <= len(widths):
                continue
            added = start + span - max(len(widths), start)
            width = cell.find(f"{_W}tcPr/{_W}tcW")
            stated = int(width.get(_W + "w")) if width is not None and width.get(_W + "type") in (None, "dxa") \
                and (width.get(_W + "w") or "").isdigit() else 0
            covered = sum(widths[start:])
            rest = stated - covered if stated > covered else UNSIZED_COLUMN * added
            widths += [UNSIZED_COLUMN] * max(start - len(widths), 0)
            widths += [rest // added + (1 if k < rest % added else 0) for k in range(added)]
    widths += [UNSIZED_COLUMN] * max(longest - len(widths), 0)
    del widths[longest:]
    _set_grid(table, widths)
    for row, end, after in reaches:
        if end + after != len(widths):
            _set_after(row, len(widths) - end, sum(widths[end:]))
    return True


def _set_after(row: Element, count: int, width: int) -> None:
    """``row``'s ``w:gridAfter`` (and ``w:wAfter``) ``count`` columns of ``width`` twips; none at 0."""
    properties = row.find(_W + "trPr")
    for tag in ("gridAfter", "wAfter"):
        node = properties.find(_W + tag) if properties is not None else None
        if node is not None:
            remove(node)
    if count <= 0:
        if properties is not None and not len(properties):
            remove(properties)
        return
    if properties is None:
        properties = make("w:trPr")
        insert_in_order(row, properties)
    insert_in_order(properties, make("w:gridAfter", **{"w:val": str(count)}))
    insert_in_order(properties, make("w:wAfter", **{"w:w": str(width), "w:type": "dxa"}))


def _spanned(row: Element) -> bool:
    """Whether a cell of ``row`` spans grid columns (``w:gridSpan`` over 1)."""
    return any(span > 1 for _cell, _start, span in _cells(row))


def _shape_row(rows: list[Element], at: int) -> int | None:
    """Round 57 (R125): the row whose cells a new row beside ``rows[at]`` -- a row with a cell
    over several grid columns -- takes: the nearest row with no such cell, the one above
    first at equal distance, a row whose deletion is pending passed over; ``None`` where
    there is none."""
    for distance in range(1, len(rows)):
        for k in (at - distance, at + distance):
            if 0 <= k < len(rows) and not _spanned(rows[k]) and _cells(rows[k]) \
                    and rows[k].find(f"{_W}trPr/{_W}del") is None:
                return k
    return None


def _row_per_column(table: Element, reference: Element) -> Element:
    """Round 57 (R125): a stand-in row with ``reference``'s properties and a cell for each
    grid column its cells cover (each with the properties of the cell over it, as wide as
    its column), for a new row where every row has a cell over several columns."""
    widths = _grid_widths(table)
    row = make("w:tr")
    for tag in ("tblPrEx", "trPr"):
        if reference.find(_W + tag) is not None:
            row.append(copy.deepcopy(reference.find(_W + tag)))
    for cell, start, span in _cells(reference):
        for column in range(start, start + span):
            new = copy.deepcopy(cell)
            for child in list(new):
                if isinstance(child.tag, str) and child.tag != _W + "tcPr":
                    remove(child)
            properties = new.find(_W + "tcPr")
            if properties is None:
                properties = make("w:tcPr")
                new.insert(0, properties)
            _set_span_width(properties, 1, widths[column] if column < len(widths) else None)
            first = cell.find(W_P)
            new.append(copy.deepcopy(first) if first is not None else make("w:p"))
            row.append(new)
    return row


def _pending_column(cell: Element) -> list[Element]:
    """The records that make ``cell`` part of a column change still pending (round 53,
    R107): its ``w:cellDel`` or ``w:cellIns``, or -- Word's form of a column inserted,
    tracked -- the inserted mark of its one empty paragraph."""
    from ..revisions import track as _track

    properties = cell.find(_W + "tcPr")
    records = [child for child in properties if child.tag in (_W + "cellDel", _W + "cellIns")] \
        if properties is not None else []
    row = cell.getparent()
    while row is not None and row.tag != W_TR:
        row = row.getparent()
    if row is not None and row.find(f"{_W}trPr/{_W}ins") is not None:
        return records                   # a row inserted, tracked: its marks are the row's
    paragraphs = [child for child in cell if isinstance(child.tag, str) and child.tag in (W_P, _W + "tbl", _W + "sdt")]
    if len(paragraphs) == 1 and paragraphs[0].tag == W_P and not _has_text(paragraphs[0]):
        record = _track.mark_record(paragraphs[0])
        if record is not None and record.tag in _track.INSERTED:
            records.append(record)
    return records


def _has_text(paragraph: Element) -> bool:
    return any(True for _ in paragraph.iter(_W + "r"))


def _whole_comments(paragraph: Element) -> set[str]:
    """The comments whose range and reference all lie in ``paragraph``."""
    def ids(tag: str) -> set[str]:
        return {node.get(_W + "id") for node in paragraph.iter(_W + tag)}

    return ids("commentRangeStart") & ids("commentRangeEnd") & ids("commentReference")


def _move_references(original: Element, copy_: Element) -> None:
    """A tracked merge's copy of a paragraph holds its comment and note references once
    (round 54): a reference written twice -- in the copy, inserted, and in the original,
    deleted -- is a document Word does not open without a dialog (measured: Word blocked
    on open).  A comment wholly in the paragraph moves with its range to the copy, as does
    a note's reference (accepting keeps them, with the merged cell); a comment reaching
    outside the paragraph stays where it is, its reference left out of the copy."""
    whole = _whole_comments(original)
    for node in list(copy_.iter(_W + "commentReference")):
        if node.get(_W + "id") not in whole:
            _drop_run_of(node)
    moved = list(original.iter(_W + "footnoteReference", _W + "endnoteReference"))
    moved += [node for node in original.iter(_W + "commentReference", _W + "commentRangeStart",
                                              _W + "commentRangeEnd") if node.get(_W + "id") in whole]
    for node in moved:
        if node.tag in (_W + "commentRangeStart", _W + "commentRangeEnd"):
            remove(node)
        else:
            _drop_run_of(node)


def _drop_run_of(node: Element) -> None:
    """``node``'s run, or ``node`` alone where its run holds more."""
    run = node.getparent()
    if run is not None and run.tag == _W + "r" and not any(
            isinstance(c.tag, str) and c.tag != _W + "rPr" and c is not node for c in run):
        remove(run)
    else:
        remove(node)


class TableOps:
    """Row, column and merge edits on :class:`docx_agent.Document`."""

    def _table_entry(self: "Document", identifier: str) -> tuple[str, TableEntry]:
        from .document import EditError

        part, entry = self._resolve(identifier)
        if not isinstance(entry, TableEntry):
            holder = entry.element.getparent()
            while holder is not None and holder.tag != _W + "tbl":
                holder = holder.getparent()
            found = self._index(part).entry_for(holder) if holder is not None else None
            if found is not None:
                from .addresses import AddressError

                raise AddressError(f"{identifier} is a paragraph in table {found.id}: give the table", [found.id])
            raise EditError(f"{identifier} is not a table")
        return part, entry

    def _stamp_new(self: "Document", part: str, nodes: list[Element]) -> None:
        used = self._used()
        for node in nodes:
            para_id, text_id = _ids.generate(part + "\0table", used, 2)
            _ids.stamp(node, para_id, text_id)
        _ids.ensure_w14(self.package.tree(part))

    # -- rows -------------------------------------------------------------------------------

    def insert_row(self: "Document", table: str, index: int | None = None, *, below: bool = True,
                   plain: bool = False) -> "EditResult":
        """A new row beside row ``index`` (0-based; the last by default), below it unless
        ``below=False``: its cells as that row's (properties, widths, and a vertical merge
        carried on where the new row falls inside one), each with an empty paragraph --
        Word's Insert Rows, which copies a cell over several columns too (measured).

        Round 57 (R125): with ``plain``, beside a row with a cell over several grid columns
        the new row takes the cells of the nearest row without one (else a cell per grid
        column), and the result's warning says so -- a row for values, one per column
        (``word_edit_table``'s)."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        rows = [row for row, _ in entry.rows]
        index = len(rows) - 1 if index is None else index
        if not 0 <= index < len(rows):
            raise EditError(f"{entry.id} has no row {index}")
        neighbour = rows[index + 1] if below and index + 1 < len(rows) else (rows[index - 1] if not below and index else None)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            rows = [row for row, _ in entry.rows]
            new_id, note = self._put_row_like(part, rows[index], neighbour, below, tracking,
                                              entry.element if plain else None)
        return EditResult(new_id, created=[new_id], renamed=renames, warnings=[note] if note else [])

    def _put_row_like(self: "Document", part: str, reference: Element, neighbour: Element | None,
                      below: bool, tracking, table: Element | None = None) -> tuple[str, str | None]:
        """:meth:`insert_row`'s change, inside its edit: the new row beside ``reference``
        (below it unless ``below`` is false), stamped, its insertion tracked under
        ``tracking``; its id, and what it says of the row's shape (or ``None``).

        Round 57 (R125): where a cell of ``reference`` spans grid columns (``w:gridSpan``)
        and ``table`` is given, the new row takes the cells of the nearest row of ``table``
        without such a cell (:func:`_shape_row`), else a cell for each grid column -- a row
        for values, one per column; Word's Insert Rows copies the span (measured), which
        made a row of one cell over three columns for four values."""
        shape, note = reference, None
        if table is not None and _spanned(reference):
            rows = own_table_rows(table)
            at = next(k for k, r in enumerate(rows) if r is reference)
            found = _shape_row(rows, at)
            if found is not None:
                shape = rows[found]
                note = (f"row {at} has a cell over several columns: the new row has row {found}'s cells, "
                        f"one per column (merge cells for a span like row {at}'s)")
            else:
                shape = _row_per_column(table, reference)
                note = (f"row {at} has a cell over several columns and every row has one: the new row has "
                        f"a cell for each of the table's {len(_grid_widths(table))} grid columns")
        row = make("w:tr")
        if shape.find(_W + "tblPrEx") is not None:
            row.append(copy.deepcopy(shape.find(_W + "tblPrEx")))
        if shape.find(_W + "trPr") is not None:
            properties = copy.deepcopy(shape.find(_W + "trPr"))
            for child in list(properties):
                if child.tag in (_W + "ins", _W + "del", _W + "trPrChange"):
                    remove(child)
                elif child.tag == _W + "tblHeader" and shape is not reference \
                        and reference.find(f"{_W}trPr/{_W}tblHeader") is None:
                    remove(child)               # a header row's shape, not its repeating
            if len(properties):
                row.append(properties)
        # A cell carries a vertical merge on where the row after the new one continues it
        # there (that row's cell over the same columns).
        after = neighbour if below else reference
        after_cells = {(start, span): cell for cell, start, span in _cells(after)} if after is not None else {}
        #: Round 53 (R107): the new cells of columns whose deletion or insertion is pending (a
        #: tracked change not yet accepted or rejected): ``(new cell, its paragraph, records)``.
        pending: list[tuple[Element, Element, list[Element]]] = []
        for cell, start, span in _cells(shape):
            new_cell = make("w:tc")
            properties = _clean_cell_properties(cell.find(_W + "tcPr"))
            merge = properties.find(_W + "vMerge")
            if merge is not None:
                remove(merge)
            if merge is not None or shape is not reference:
                if (start, span) in after_cells and _vmerge(after_cells[(start, span)]) == "continue":
                    insert_in_order(properties, make("w:vMerge"))
            new_cell.append(properties)
            paragraph = _empty_paragraph_like(cell.find(W_P))
            new_cell.append(paragraph)
            row.append(new_cell)
            records = _pending_column(cell)
            if records:
                pending.append((new_cell, paragraph, records))
        if below:
            reference.addnext(row)
        else:
            reference.addprevious(row)
        self._stamp_new(part, [row] + list(row.iter(W_P)))
        column_marks = self._follow_pending_columns(pending)
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
                if not any(paragraph is marked for marked in column_marks):
                    _track.mark(paragraph, "ins", stamp, part)
            stamp.finish()
        self.package.mark_dirty(part)
        story = self._story_of(part)
        return ("" if story == "body" else f"{story}/") + f"tr:{row.get(_ids.PARA_ID)}", note

    def _follow_pending_columns(self: "Document", pending: list[tuple[Element, Element, list[Element]]]
                                ) -> list[Element]:
        """Round 53 (R107): a new row's cells in columns whose tracked deletion or insertion
        is pending go with that change -- each a copy of its column's record, a fresh id: a
        ``w:cellDel`` (a column deleted, tracked), or its paragraph's mark inserted (Word's
        form of a column inserted, tracked).  Accepted or rejected, the new row then loses or
        keeps the cell as the column's other rows do, and the grid stays in step with them
        (a row inserted beside a column whose deletion was pending kept its cell when the
        deletion was accepted: a row a column wider than the grid).  The paragraphs given a
        mark of their column's."""
        from ..revisions import track as _track

        if not pending:
            return []
        next_id = self._next_annotation_id()
        marked: list[Element] = []
        for cell, paragraph, records in pending:
            for record in records:
                copied = copy.deepcopy(record)
                copied.set(_W + "id", str(next_id))
                next_id += 1
                if record.tag == _W + "cellDel":
                    # Round 57: deleted with the column, its mark too -- a cell w:cellDel whose
                    # mark is kept is one Accept All gives to the cell before it (a row a
                    # column wider), and reads as a pending merge (delete_row refused it).
                    mark = copy.deepcopy(record)
                    mark.tag = _W + "del"
                    mark.set(_W + "id", str(next_id))
                    next_id += 1
                    properties = paragraph.find(_W + "pPr")
                    if properties is None:
                        properties = make("w:pPr")
                        paragraph.insert(0, properties)
                    run_properties = properties.find(_W + "rPr")
                    if run_properties is None:
                        run_properties = make("w:rPr")
                        insert_in_order(properties, run_properties)
                    insert_in_order(run_properties, mark)
                    marked.append(paragraph)
                if record.tag in _track.MARKS:
                    properties = paragraph.find(_W + "pPr")
                    if properties is None:
                        properties = make("w:pPr")
                        paragraph.insert(0, properties)
                    run_properties = properties.find(_W + "rPr")
                    if run_properties is None:
                        run_properties = make("w:rPr")
                        insert_in_order(properties, run_properties)
                    insert_in_order(run_properties, copied)
                    marked.append(paragraph)
                else:
                    insert_in_order(cell.find(_W + "tcPr"), copied)
        return marked

    # -- rows deleted, tracked (round 51, R103) -----------------------------------------------

    def _track_delete_rows(self: "Document", part: str, rows: list[Element], stamp, *,
                           nested_rows: bool = False) -> list[Element]:
        """Delete ``rows`` (of one table, in document order) as tracked changes, as Word
        deletes them (measured, ``tools/rows_probe.py``, ``tests/observations/rows-word.json``):

        * a row that is wholly the author's own pending insertion -- its ``w:trPr/w:ins``, each
          cell's paragraph mark, and every run an insertion -- is **removed outright**, with no
          record left; what in it is another author's insertion stays theirs, deleted by this
          author (``w:ins`` > ``w:del``), at the start of the first paragraph after the row
          (the next row's first cell, or after the table);
        * a row the author inserted holding what is not an insertion (text or a paragraph mark
          there before) keeps its ``w:trPr/w:ins`` and gets no ``w:del``: its content and
          marks are deleted (Word's Accept All and Reject All both take the row out);
        * any other row gets ``w:trPr/w:del``, its content and marks deleted -- beside another
          author's inserted mark too.

        One's own insertions inside are removed (:func:`~docx_agent.revisions.track.
        delete_runs`).  ``nested_rows``: a deleted row's nested tables' rows are marked
        deleted too (a whole table's deletion).  Returns the rows removed outright."""
        from ..revisions import track as _track
        from . import inline as _inline
        from .document import _rehome_markers

        kinds = {id(row): _row_fate(row, stamp.author) for row in rows}
        going = {id(row) for row in rows if kinds[id(row)] in ("remove", "relocate")}
        targets: dict[int, Element | None] = {}
        for row in rows:
            if kinds[id(row)] == "relocate":
                targets[id(row)] = _paragraph_after(row, going)
                if targets[id(row)] is None:
                    kinds[id(row)] = "keep"           # nowhere for another's insertion to go
                    going.discard(id(row))
        for row in rows:
            kind = kinds[id(row)]
            if kind in ("remove", "relocate"):
                continue
            if kind == "mark":
                for target in [row] + ([r for r in row.iter(W_TR) if r is not row] if nested_rows else []):
                    properties = target.find(_W + "trPr")
                    if properties is None:
                        properties = make("w:trPr")
                        insert_in_order(target, properties)
                    if properties.find(_W + "del") is None:
                        insert_in_order(properties, stamp.make("w:del", part))
            for paragraph in row.iter(W_P):
                self._release(part, _track.delete_content(paragraph, stamp, part))
                _delete_cell_mark(paragraph, stamp, part, own=kind == "keep")
        removed: list[Element] = []
        for row in reversed(rows):                # a later row's moved content after an earlier's
            if kinds[id(row)] not in ("remove", "relocate"):
                continue
            moved: list[Element] = []
            for paragraph in row.iter(W_P):
                self._release(part, _track.delete_content(paragraph, stamp, part))
                moved += [child for child in paragraph if isinstance(child.tag, str)
                          and child.tag != _W + "pPr" and any(True for _ in child.iter(_W + "r"))]
            target = targets.get(id(row))
            if moved and target is not None:
                at = 1 if target.find(_W + "pPr") is not None else 0
                for k, node in enumerate(moved):
                    target.insert(at + k, node)
            _rehome_markers(row)
            self._release(part, _inline.relationship_ids(row))
            remove(row)
            removed.append(row)
        removed.reverse()
        return removed

    # -- a table filled in one pass (round 45, R67) -------------------------------------------
    #
    # word_edit_table's set_cells from a source wrote a block of thousands of cells as one
    # edit per row added and one or more per cell: each edit threw every index away, so the
    # next one read the whole document again (every story's paragraphs, the ids in use, the
    # table's grid) -- a 431-row block took minutes.  These two make the same changes, in the
    # same order, in one pass: the bytes and ids are those the edits one by one give (tests:
    # test_table_fill_speed.py).  Only where nothing needs stamping (Document._stamping_due),
    # where every edit's stamping stamps nothing; the caller checks, and makes the edits one
    # by one otherwise.

    def _rows_after_last(self: "Document", table: str, count: int) -> list["EditResult"]:
        """``count`` rows after the table's last, each as ``insert_row(table, last, plain=True)`` makes
        one -- and each that would repeat as a header row made a body row
        (``set_row(header=False)``), as ``word_edit_table``'s grow did, row by row: their
        edits, in order."""
        from .document import EditResult

        part, entry = self._table_entry(table)
        table_id, height = entry.id, len(entry.rows)
        tracking = self._active_tracking()
        edits: list[EditResult] = []
        with self._edit(), self._held_ids():
            last = entry.rows[-1][0]
            for k in range(count):
                new_id, note = self._put_row_like(part, last, None, True, tracking, entry.element)
                last = last.getnext()
                edits.append(EditResult(new_id, created=[new_id], renamed={}, warnings=[note] if note else []))
                if last.find(f"{_W}trPr/{_W}tblHeader") is not None:
                    self._invalidate()          # set_row finds the row in the index
                    edits.append(self.set_row(table_id, height + k, header=False))
        return edits

    def _rows_beside(self: "Document", table: str, places: list[tuple[int, bool, int]]) -> list["EditResult"]:
        """For each ``(index, below, count)`` of ``places``, in order: ``count`` rows (``plain``) as
        ``insert_row(table, index, below=below)`` makes the first and ``insert_row`` below
        the row just made each next one -- and each that would repeat as a header row made a
        body row (``set_row(header=False)``) -- as ``word_edit_table``'s ``sync`` makes them
        one by one (round 49, R90): their edits, in order.  Each ``index`` is a row's index
        when its place comes (a caller adding from the bottom up gives the table's own)."""
        from .document import EditResult

        part, entry = self._table_entry(table)
        table_id = entry.id
        rows = [row for row, _ in entry.rows]
        tracking = self._active_tracking()
        edits: list[EditResult] = []
        with self._edit(), self._held_ids():
            for index, below, count in places:
                at, side = index, below
                for _ in range(count):
                    reference = rows[at]
                    neighbour = (rows[at + 1] if side and at + 1 < len(rows)
                                 else (rows[at - 1] if not side and at else None))
                    new_id, note = self._put_row_like(part, reference, neighbour, side, tracking, entry.element)
                    at = at + 1 if side else at
                    rows.insert(at, reference.getnext() if side else reference.getprevious())
                    edits.append(EditResult(new_id, created=[new_id], renamed={}, warnings=[note] if note else []))
                    side = True
                    if rows[at].find(f"{_W}trPr/{_W}tblHeader") is not None:
                        self._invalidate()          # set_row finds the row in the index
                        edits.append(self.set_row(table_id, at, header=False))
        return edits

    def _fill_cells(self: "Document", table: str, writes: list[tuple[int, int, str]]
                    ) -> tuple[list["EditResult"], list[str]]:
        """Each ``(row, column, text)`` of ``writes`` in order, as ``word_edit_table`` writes
        a cell's text: its first paragraph's text set (formatting kept) where it differs,
        its other paragraphs deleted, each line after the first a paragraph after it.  The
        edits, and each cell's id."""
        from .document import EditResult, ParagraphEntry, cell_element
        from . import text as _text

        part, entry = self._table_entry(table)
        grid = self.table(entry.id).grid()
        rows = [row for row, _ in entry.rows]     # the same elements after every edit here
        tracking = self._active_tracking()
        edits: list[EditResult] = []
        ids: list[str] = []
        with self._edit(), self._held_ids():
            for row, column, text in writes:
                cell = grid[row][column]
                element = cell_element(rows[cell.row], cell.column)
                if element is None:
                    raise KeyError(f"{cell.id} no longer exists")
                index = self._index(part)
                paragraphs = [found for found in map(index.entry_for, element.iter(W_P))
                              if isinstance(found, ParagraphEntry)]
                lines = text.split("\n")
                first = paragraphs[0]
                if _text.paragraph_text(first.element) != lines[0]:
                    _text.check_text(lines[0])
                    nested = any(True for node in first.element.iter(W_P, W_TBL, W_TR) if node is not first.element)
                    self._write_text(part, first, lines[0], tracking, prepare=False)
                    edits.append(EditResult(first.id, renamed={}))
                    if nested:
                        self._invalidate()      # what it held may have gone with its text
                for extra in paragraphs[1:]:
                    edits.append(self.delete_block(extra.id))
                if len(lines) > 1:
                    last = self._index(part).entry_for(first.element).id
                    for line in lines[1:]:
                        made = self.insert_paragraph(line, after=last)
                        edits.append(made)
                        last = made.id
                ids.append(f"{entry.id}/c{cell.row},{cell.column}")
        return edits, ids

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
        tracking = self._active_tracking()
        # Round 45: the reference of a comment on the row before it moves on (untracked).
        from . import comments as _comments

        moving = _comments.block_references(rows[index]) if tracking is None else None
        for paragraph in rows[index].iter(W_P):
            if moving is None and any(True for _ in paragraph.iter(_W + "footnoteReference", _W + "endnoteReference",
                                                                   _W + "commentReference")):
                raise EditError(f"row {index} holds a note or comment reference")
        old_row_id = entry.rows[index][1]
        merged: list[tuple[Element, Element]] = []
        if tracking is not None:
            from ..revisions.review import merge_groups

            # Round 55 (R115): a row a pending tracked merge reaches -- the text it gave the
            # merged cell leaves it again, as had the row gone first; a row holding the
            # merge's first cell is refused.
            for group in merge_groups(entry.element):
                if any(cell is group[0][1] for cell in rows[index].iter(W_TC)):
                    raise EditError(f"row {index} holds the first cell of a pending tracked merge (its text and "
                                    "the cells merged into it); settle that merge first")
                merged += [(group[0][1], cell) for r, cell in group[1:] if r == index]
            for origin, cell in merged:
                if any(True for copy_ in _copies_of(origin, cell)
                       for _ in copy_.iter(_W + "footnoteReference", _W + "endnoteReference", _W + "commentReference")):
                    raise EditError(f"row {index} holds a note or comment reference")
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
            from .document import _rehome_markers

            removed = [old_row_id]
            if stamp is None or _row_fate(row, stamp.author) in ("remove", "relocate"):
                inside = set(map(id, row.iter(W_P)))
                removed += [e.id for e in self._index(part).paragraphs if id(e.element) in inside]
            for origin, cell in merged:
                _drop_copies(origin, cell, stamp, part)
            if stamp is not None:
                # Round 51 (R103): one's own pending row goes outright, as Word removes it.
                gone = self._track_delete_rows(part, [row], stamp)
                stamp.finish()
                self.package.mark_dirty(part)
                return EditResult(None, renamed=renames, removed=removed if gone else [])
            _comments.move_block_references(row, _comments.block_references(row) or [])
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
        if self._active_tracking() is not None:
            from ..revisions.review import merged_away

            leads, origins = merge_leads(entry.element), _merge_origins(entry.element)
            for row, _ in entry.rows:
                for cell, start, span in _cells(row):
                    if not (start == index and span == 1):
                        continue
                    if id(cell) in leads:
                        raise EditError(f"column {index} holds the first cell of a pending tracked merge (its text "
                                        "and the cells merged into it); settle that merge first")
                    if merged_away(cell) and any(True for copy_ in _copies_of(origins.get(id(cell), (None, None))[1], cell)
                                                 for _ in copy_.iter(_W + "footnoteReference", _W + "endnoteReference",
                                                                     _W + "commentReference")):
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
            # Round 57: a column wholly this author's pending insertion (insert_column, tracked,
            # and what was typed into it) goes outright, as one's own pending rows do (round 51,
            # R103) -- marked deleted, its cells stayed for a later row to copy.
            own_column = tracking is not None and bool(doomed) and not spans and all(
                _own_inserted_cell(cell, tracking.author) for cell in doomed)
            if tracking is not None and not own_column:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp
                from .table_format import TableTracker

                from ..revisions.review import column_deleted, merged_away

                tracker = TableTracker(element, rows=True) if spans else None
                stamp = Stamp(self, tracking)
                origins = _merge_origins(element)
                for cell in doomed:
                    if column_deleted(cell):
                        continue
                    if merged_away(cell):
                        # Round 55 (R115): a cell a pending tracked merge took in -- its text
                        # leaves the merged cell again, as had the column gone first; its
                        # w:cellDel stays, now with its content deleted.
                        _drop_copies(origins.get(id(cell), (None, None))[1], cell, stamp, part)
                    else:
                        properties = cell.find(_W + "tcPr")
                        if properties is None:
                            properties = make("w:tcPr")
                            cell.insert(0, properties)
                        insert_in_order(properties, stamp.make("w:cellDel", part))
                    for paragraph in cell.iter(W_P):
                        self._release(part, _track.delete_content(paragraph, stamp, part))
                        record = _track.mark_record(paragraph)
                        if record is not None and record.tag == _W + "ins" and _row_of(cell) is not None \
                                and _row_of(cell).find(f"{_W}trPr/{_W}ins") is not None:
                            # Round 55 (R115): a pending new row's cell (a split's) -- its mark
                            # deleted with the column, as a split after the deletion writes it.
                            holder = record.getparent()
                            remove(record)
                            if holder is not None and not len(holder):
                                remove(holder)
                            record = None
                        if record is None or (record.tag == _W + "ins" and record.getparent().find(_W + "del") is None):
                            # Round 57 (R125): a mark a pending insertion made (a split's new
                            # cell) is deleted beside it -- else Accept All keeps the cell, given
                            # to the one before it: a row a column wider than the others (Word's
                            # Accept All too, measured).
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
                table_id = renames.get(entry.id, entry.id)
                # Round 54 (R110, lane M): what the reader sees now, and what positions count.
                warning = (f"column {index} is deleted as a tracked change: word_read, render and word_edit_table's "
                           f"col positions count the table as accepted, without it; word_changes accept or reject "
                           f"within {table_id} settles the deletion") if doomed else None
                return EditResult(table_id, renamed=renames, warnings=[warning] if warning else [])
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

    def columns_pending_deletion(self: "Document", table: str) -> list[int]:
        """The grid columns of ``table`` that accepting its tracked changes removes: those
        only cells with a pending deletion (``w:cellDel``) cover, a deleted cell whose content
        is kept not counted -- accepting gives it to the cell before it (round 54, R110 and
        lane M)."""
        return [c for c in range(self._grid_count(table)) if c not in self.accepted_columns(table)[0]]

    def _grid_count(self: "Document", table: str) -> int:
        _part, entry = self._table_entry(table)
        grid = entry.element.find(_W + "tblGrid")
        return len(grid.findall(_W + "gridCol")) if grid is not None else 0

    def accepted_columns(self: "Document", table: str) -> tuple[list[int], dict[tuple[int, int], int]]:
        """``table`` as accepting its tracked changes leaves its columns (round 54, lane M):
        the grid columns kept, in order, and where a cell merged away is given to the cell
        before it -- ``(row, grid column)`` of each position such a cell (``w:cellDel``, its
        content kept) covers, to the grid column where the cell before it begins.  Word's
        Accept All: a deleted cell whose content is kept widens the cell before it, one
        whose content is deleted goes, and so does a grid column only such cells cover
        (measured, round 54).  Rows are counted as the document has them."""
        from ..revisions.review import _content_kept

        _part, entry = self._table_entry(table)
        count = self._grid_count(table)
        covered: dict[int, bool] = {}
        given: dict[tuple[int, int], int] = {}
        for r, (row, _) in enumerate(entry.rows):
            deleted_row = row.find(f"{_W}trPr/{_W}del") is not None
            taken, _lost = _accept_ends(row)
            before: tuple[int, int] | None = None     # the last cell kept: (start, span)
            for cell, start, span in _cells(row):
                deleted = cell.find(f"{_W}tcPr/{_W}cellDel") is not None
                if deleted and before is not None and _content_kept(cell):
                    for column in range(start, start + span):
                        given[(r, column)] = before[0]
                    deleted = False
                elif cell in taken:
                    # Round 54 (lane R): a cell Accept All takes into the one before it (that
                    # cell's last paragraph mark deleted): its positions are that cell's.
                    for column in range(start, start + span):
                        given[(r, column)] = taken[cell]
                    deleted = True
                elif cell in _lost:
                    deleted = True                     # removed with its text by Accept All
                elif not deleted:
                    before = (start, span)
                if deleted_row:
                    continue
                for column in range(start, start + span):
                    covered[column] = covered.get(column, True) and deleted
        return [c for c in range(count) if not covered.get(c, False)], given

    def cells_lost_on_accept(self: "Document", table: str) -> list[tuple[int, int, Element]]:
        """``(row, grid column, cell)`` of each cell of ``table`` Word's Accept All removes
        with its text: a cell with no next cell in its row whose last paragraph mark is a
        pending deletion (round 54, lane R; :mod:`docx_agent.revisions.edges`)."""
        _part, entry = self._table_entry(table)
        out = []
        for r, (row, _) in enumerate(entry.rows):
            if row.find(f"{_W}trPr/{_W}del") is not None:
                continue
            _taken, lost = _accept_ends(row)
            out += [(r, start, cell) for cell, start, _span in _cells(row) if cell in lost]
        return out

    def pending_structure(self: "Document", table: str) -> list[str]:
        """The tracked structure changes ``table`` holds that are not yet accepted (round 54,
        R111), each said once in words: ``column deleted`` (``w:cellDel``), ``cells
        inserted`` (``w:cellIns``), ``cells merged down`` (``w:cellMerge``), ``grid
        changed`` (``w:tblGridChange``: a merge across, a split, a column across a span).
        Nested tables' own changes are theirs, not this table's."""
        _part, entry = self._table_entry(table)
        names = {_W + "cellDel": "cells deleted", _W + "cellIns": "cells inserted",
                 _W + "cellMerge": "cells merged down", _W + "tblGridChange": "grid changed"}
        found: list[str] = []
        for node in entry.element.iter(*names):
            if next(node.iterancestors(W_TBL), None) is entry.element and names[node.tag] not in found:
                found.append(names[node.tag])
        return found

    # -- merging ----------------------------------------------------------------------------

    def _rectangle(self: "Document", entry: TableEntry, first: tuple[int, int], last: tuple[int, int],
                   pending: bool = False):
        """The rectangle ``first``..``last`` grown until no cell lies partly in it -- across
        a span, down a vertical merge -- and each row's cells in it.  ``pending`` (a tracked
        merge, round 55): a pending tracked merge counts as the one cell it makes, so a merge
        takes it in whole."""
        from .document import EditError

        (r1, c1), (r2, c2) = first, last
        rows = [row for row, _ in entry.rows]
        if not (0 <= r1 <= r2 < len(rows)) or c1 > c2 or c1 < 0:
            raise EditError("first and last are (row, column), first above and left of last")
        layout = _layout(rows)
        boxes = _pending_boxes(entry.element, rows) if pending else []
        while True:
            grown = False
            for (top, bottom, left, right) in boxes:
                if top <= r2 and bottom >= r1 and left <= c2 and right >= c1 and not (
                        r1 <= top and bottom <= r2 and c1 <= left and right <= c2):
                    r1, r2, c1, c2 = min(r1, top), max(r2, bottom), min(c1, left), max(c2, right)
                    grown = True
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
        tracking = self._active_tracking()
        first_, last_, plan = self._rectangle(entry, first, last, pending=tracking is not None)
        warnings = []
        if (first_, last_) != (tuple(first), tuple(last)):
            warnings.append(f"the merge takes in whole cells: ({first_[0]}, {first_[1]}) to ({last_[0]}, {last_[1]})")
        if tracking is not None:
            plan = _tracked_plan(plan)
        if all(len(plan_row) == 1 for plan_row in plan) and (
                len(plan) == 1 or (_vmerge(plan[0][0][0]) == "restart"
                                   and all(_vmerge(r[0][0]) == "continue" for r in plan[1:])
                                   and len({(r[0][1], r[0][2]) for r in plan}) == 1)) \
                or (tracking is not None and _merged_down_pending(plan)):
            return EditResult(entry.id, changed=False, warnings=warnings)
        (r1, c1), (r2, c2) = first_, last_
        with self._edit():
            renames = self._prepare(part, [])
            part, entry = self._table_entry(renames.get(table, table))
            rows = [row for row, _ in entry.rows]
            plan = [[(cell, start, span) for cell, start, span in _cells(rows[r]) if c1 <= start <= c2]
                    for r in range(r1, r2 + 1)]
            if tracking is not None:
                plan = _tracked_plan(plan)
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
        """Round 54 (follow-up, lane M): a merge written in the form Word accepts and rejects exactly
        (measured, Word 16.106: ``feedback-round54/probe-M``; Word's own merge is applied
        untracked).  Written so, Word's Accept All and Reject All give what docx-agent's own
        give, and while pending the table stays on its grid.  Round 55 (R115): over pending
        tracked changes too -- a cell a pending column deletion removes and a row whose
        deletion is pending are left out of the merge, a pending merge it takes in kept
        (:func:`_tracked_plan`; measured, ``feedback-round55/probe-C``):

        * each row's first cell keeps its properties (no new span): accepting the merged-away
          cells widens it;
        * every other cell of the row is ``w:cellDel`` with its content and marks kept: Word's
          Accept All gives a deleted cell whose content is kept to the cell before it (its
          columns and width) and drops the content; Reject All keeps the cell;
        * merged down, each row's first cell a ``w:cellMerge`` (``rest`` the first row,
          ``cont`` the others; the old merge in ``vMergeOrig``), a continuing row's first
          cell's content deleted;
        * every other cell's text inserted at the end of the first cell, in reading order, a
          comment or note reference moved with it once (round 54, lane T)."""
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp
        from .document import _clear_revisions

        from ..revisions.review import merged_away

        stamp = Stamp(self, tracking)
        across = any(len(plan_row) > 1 for plan_row in plan)
        copies: list[Element] = []
        for k, plan_row in enumerate(plan):
            lead = plan_row[0][0]
            # Round 55 (R115): a cell a pending tracked merge already took in keeps its
            # record; its text is in that merge's first cell, which this merge takes in.
            cells = [lead] + [cell for cell, _, _ in plan_row[1:] if not merged_away(cell)]
            if len(plan) > 1:
                properties = lead.find(_W + "tcPr")
                if properties is None:
                    properties = make("w:tcPr")
                    lead.insert(0, properties)
                record = properties.find(_W + "cellMerge")
                if record is not None:
                    # Merged down already, pending: the record says where it is now.
                    record.set(_W + "vMerge", "rest" if k == 0 else "cont")
                else:
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
                kept = cell is not lead
                for paragraph in [p for p in cell if isinstance(p.tag, str) and p.tag == W_P]:
                    if _has_text(paragraph):
                        copy_ = copy.deepcopy(paragraph)
                        _clear_revisions(copy_)
                        for marker in list(copy_.iter(_W + "bookmarkStart", _W + "bookmarkEnd",
                                                      _W + "commentRangeStart", _W + "commentRangeEnd")):
                            if marker.tag in (_W + "bookmarkStart", _W + "bookmarkEnd") \
                                    or marker.get(_W + "id") not in _whole_comments(paragraph):
                                remove(marker)
                        if _has_text(copy_):
                            _move_references(paragraph, copy_)
                            copies.append(copy_)
                    if not kept:
                        # A continuing row's first cell stays (merged into the one above):
                        # its content goes.
                        self._release(part, _track.delete_content(paragraph, stamp, part))
            for cell in cells[1:]:
                properties = cell.find(_W + "tcPr")
                if properties is None:
                    properties = make("w:tcPr")
                    cell.insert(0, properties)
                insert_in_order(properties, stamp.make("w:cellDel", part))
            if k > 0:
                # A continuation keeps its last paragraph: the others' marks go.
                paragraphs = [p for p in lead if isinstance(p.tag, str) and p.tag == W_P]
                for paragraph in paragraphs[:-1]:
                    if _track.mark_record(paragraph) is None:
                        _track.mark(paragraph, "del", stamp, part)
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
        how = "deleted cells" + (" and a cell merge" if len(plan) > 1 else "") if across else "a cell merge"
        self._describe(part, origin, f"cells ({first[0] + 1}, {first[1] + 1}) to ({last[0] + 1}, {last[1] + 1}) "
                                     f"merged as {how}: Accept All merges them, Reject All keeps them apart.")

    # -- splitting --------------------------------------------------------------------------

    def split_cell(self: "Document", table: str, row: int, column: int, *, rows: int = 1,
                   columns: int | None = None) -> "EditResult":
        """Split the cell over grid position ``(row, column)`` into ``rows`` x ``columns``
        cells, as Word splits one (measured): across, a cell spanning enough grid columns
        is divided among the new cells, otherwise its grid column is divided and every other
        row's cell over it spans the new columns; down, a vertical merge is divided, and
        where the cell has too few rows new rows go below it, the row's other cells merged
        down into them.  The content stays in the first cell; the new cells hold an empty
        paragraph.  Tracked (Word's own split is applied untracked), the new cells' and
        rows' paragraph marks are insertions, the other cells' old properties
        ``w:tcPrChange``s and the old grid a ``w:tblGridChange``, and a comment says so.

        Round 57 (R125): ``columns`` defaults to the cell's grid columns where it spans
        several (each its own cell again), else 2 -- Word's dialog offers 2 for any cell
        (measured), which on a cell over 3 columns adds a grid column and changes every other
        row.  A cell over several grid columns split into no more cells than it spans is
        divided along the grid's own lines, the first cells a column more where they do not
        share evenly; the result warns where Word, dividing the width evenly, would have
        added grid columns instead."""
        from .document import EditError, EditResult

        part, entry = self._table_entry(table)
        all_rows = [r for r, _ in entry.rows]
        target = next(((cell, start, span) for cell, start, span in _cells(all_rows[row])
                       if start <= column < start + span), None) if 0 <= row < len(all_rows) else None
        if columns is None:
            columns = target[2] if target is not None and target[2] > 1 else 2
        if not (isinstance(rows, int) and isinstance(columns, int) and 1 <= rows <= 63 and 1 <= columns <= 63):
            raise EditError("rows and columns are whole numbers from 1 to 63")
        if rows == 1 and columns == 1:
            return EditResult(entry.id, changed=False)
        if target is None:
            from .document import no_cell_message

            raise EditError(no_cell_message(entry.id, row, column, self.table(entry.id).grid()))
        warnings: list[str] = []
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
                made, uneven = _split_across(element, rows_before, top, bottom, cell, columns)
                new_cells += made
                if uneven:
                    warnings.append(f"cell ({row}, {column}) over {uneven[0]} grid columns split into {columns} "
                                    f"along the table's grid, over {', '.join(map(str, uneven[1]))} column(s) "
                                    "each: Word's Split Cells divides the width evenly, adding grid columns "
                                    "and changing every other row; the grid kept")
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
                # Round 55 (R115): a new row's cell under a cell a pending column deletion
                # removes is deleted with it (w:cellDel, its mark deleted) -- otherwise Word's
                # Accept All keeps it, the row a cell too long (measured, probe-C).
                from ..revisions.review import column_deleted

                gone_at = {start for cell_, start, _ in _cells(rows_before[top]) if column_deleted(cell_)}
                for new_row in new_rows:
                    for cell_, start, _ in _cells(new_row):
                        if start in gone_at:
                            properties = cell_.find(_W + "tcPr")
                            if properties is None:
                                properties = make("w:tcPr")
                                cell_.insert(0, properties)
                            insert_in_order(properties, stamp.make("w:cellDel", part))
                            for paragraph in cell_.iter(W_P):
                                _track.mark(paragraph, "del", stamp, part)
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
                self._describe(part, cell, f"cell ({row + 1}, {column + 1}) split into {rows} x {columns}, "
                                           "written as inserted cells and rows: Reject All takes them out.")
            self.package.mark_dirty(part)
            created = [self._paragraph_id(part, p) for p in new_paragraphs]
        return EditResult(renames.get(entry.id, entry.id), renamed=renames, created=[c for c in created if c],
                          warnings=warnings)

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


def _merge_origins(table: Element) -> dict[int, tuple[Element, Element]]:
    """Round 55 (R115): ``id()`` of each cell a pending tracked merge took in -> (the cell,
    held so its ``id()`` stays its own; the merge's first cell, where its text went)."""
    from ..revisions.review import merge_groups

    return {id(cell): (cell, group[0][1]) for group in merge_groups(table) for _, cell in group[1:]}


def merge_leads(table: Element) -> dict[int, tuple[Element, list[Element]]]:
    """Round 55 (R115): ``id()`` of each cell that leads a pending tracked merge in its row
    or down its column -> (the cell, held; the merge's cells)."""
    from ..revisions.review import merge_groups

    out: dict[int, tuple[Element, list[Element]]] = {}
    for group in merge_groups(table):
        cells = [cell for _, cell in group]
        for cell in cells:
            if cell.find(f"{_W}tcPr/{_W}cellDel") is None:
                out[id(cell)] = (cell, cells)
    return out


def _drop_copies(origin: Element | None, cell: Element, stamp, part: str) -> None:
    """Round 55 (R115): take the text a pending tracked merge copied from ``cell`` (merged
    away) out of the merge's first cell ``origin`` -- each copy is an inserted paragraph
    there with the same text; the cell's paragraph marks around it set as a merge without
    ``cell`` writes them (the cell's last mark inserted only while a copy follows it)."""
    from ..revisions import track as _track

    for found in _copies_of(origin, cell):
        paragraphs = [p for p in origin if isinstance(p.tag, str) and p.tag == W_P]
        last = found is paragraphs[-1]
        position = paragraphs.index(found)
        change = found.find(f"{_W}pPr/{_W}pPrChange")
        old = copy.deepcopy(change.find(W_PPR)) if change is not None and change.find(W_PPR) is not None else None
        remove(found)
        if not last:
            continue
        before = paragraphs[position - 1]
        record = _track.mark_record(before)
        if record is not None and record.tag in _track.INSERTED:
            holder = record.getparent()
            remove(record)
            if holder is not None and not len(holder):
                remove(holder)
        if old is not None and _all_inserted(before):
            _track.record_paragraph_change(before, old, stamp, part)


def _copies_of(origin: Element | None, cell: Element) -> list[Element]:
    """The paragraphs a pending tracked merge copied from ``cell`` into ``origin``: for each
    of the cell's paragraphs with text, the first inserted paragraph of ``origin`` (not its
    first) with the same text not yet matched."""
    from . import text as _text

    if origin is None:
        return []
    candidates = [p for p in origin if isinstance(p.tag, str) and p.tag == W_P][1:]
    candidates = [p for p in candidates if _all_inserted(p)]
    # A row's first cell merged down has its content deleted: its text as it was.
    view = "original" if cell.find(f"{_W}tcPr/{_W}cellMerge") is not None else "current"
    out: list[Element] = []
    for paragraph in [p for p in cell if isinstance(p.tag, str) and p.tag == W_P and _has_text(p)]:
        text = _text.paragraph_text(paragraph, view)
        found = next((p for p in candidates if text and _text.paragraph_text(p) == text
                      and not any(p is q for q in out)), None)
        if found is not None:
            out.append(found)
    return out


def _all_inserted(paragraph: Element) -> bool:
    """Whether every run of ``paragraph`` is inside an insertion (a merge's copy)."""
    runs = list(paragraph.iter(_W + "r"))
    return bool(runs) and all(any(a.tag in (_W + "ins", _W + "moveTo") for a in r.iterancestors())
                              for r in runs)


def _pending_boxes(table: Element, rows: list[Element]) -> list[tuple[int, int, int, int]]:
    """Round 55 (R115): each pending tracked merge of ``table`` as the rectangle it makes
    (top row, bottom row, first grid column, last grid column)."""
    from ..revisions.review import merge_groups

    held = [(r, _cells(row)) for r, row in enumerate(rows)]       # held: their id()s stay theirs
    place = {id(cell): (r, start, span) for r, cells in held for cell, start, span in cells}
    boxes = []
    for group in merge_groups(table):
        spots = [place[id(cell)] for _, cell in group if id(cell) in place]
        if spots:
            boxes.append((min(r for r, _, _ in spots), max(r for r, _, _ in spots),
                          min(s for _, s, _ in spots), max(s + n - 1 for _, s, n in spots)))
    return boxes


def _tracked_plan(plan: list) -> list:
    """Round 55 (R115): a tracked merge's rows and cells, what pending tracked changes take
    out left out -- a row whose deletion is pending (merged down: Word's Accept All removes
    it, the merge continuing over it), a cell a pending column deletion removes."""
    from ..revisions.review import column_deleted

    if len(plan) > 1:
        kept_rows = [row for row in plan if row and _row_of(row[0][0]).find(f"{_W}trPr/{_W}del") is None]
        plan = kept_rows or plan
    return [[cell for cell in row if not column_deleted(cell[0])] or row for row in plan]


def _merged_down_pending(plan: list) -> bool:
    """Whether ``plan`` is already one pending tracked merge: each row's first cell joined
    by a ``w:cellMerge`` (``rest`` then ``cont``) and the others merged away."""
    from ..revisions.review import merged_away

    def merge_of(cell: Element) -> str | None:
        record = cell.find(f"{_W}tcPr/{_W}cellMerge")
        return record.get(_W + "vMerge") if record is not None else None

    if not all(merged_away(cell) for row in plan for cell, _, _ in row[1:]):
        return False
    if len(plan) == 1:
        return len(plan[0]) > 1
    return merge_of(plan[0][0][0]) == "rest" and all(merge_of(row[0][0]) == "cont" for row in plan[1:])


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


def _split_across(table: Element, rows: list[Element], top: int, bottom: int, cell: Element,
                  n: int) -> tuple[list[Element], tuple[int, list[int]] | None]:
    """Split ``cell`` (and the cells under it in its vertical merge) into ``n`` across, as
    Word does: the grid divided where the new edges fall between its lines (every other
    cell over a divided column spanning the new columns), the cell's span shared where
    they fall on them.  Round 57 (R125): a cell over at least ``n`` grid columns is divided
    along the grid's lines instead (its columns shared out, the first cells one more), so
    no grid column is added.  The new cells, in document order; and where Word's even
    division would have added grid columns, ``(the cell's span, each new cell's)``."""
    widths = _grid_widths(table)
    lines = [0]
    for width in widths:
        lines.append(lines[-1] + width)
    start, span = next((s, sp) for c, s, sp in _cells(rows[top]) if c is cell)
    x0, x1 = lines[start], lines[start + span]
    cuts = [x0 + round((x1 - x0) * k / n) for k in range(1, n)]
    uneven = None
    if 1 < n <= span and start + span < len(lines):
        sizes = [span // n + (1 if k < span % n else 0) for k in range(n)]
        if not set(cuts) <= set(lines):
            uneven = (span, sizes)
        cuts = [lines[start + sum(sizes[:k])] for k in range(1, n)]
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
    return made, uneven


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
