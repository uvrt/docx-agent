"""Logical changes: the revision records a person made in one go, grouped (``doc.changes()``).

``doc.revisions()`` lists Word's records one by one: a replacement is a deletion and an
insertion, a new paragraph is its text and its mark (``paragraph-mark-insertion``), and Word's
form for paragraphs added at a story's end adds a ``paragraph-properties`` record on the last
one.  An agent asked for "one line per change" had to pair them itself (the end-to-end
trial).  :func:`changes` groups them as a reviewer reads them:

* **touching records by one author at one date** are one change: a deletion and an insertion
  side by side are a ``replacement``; runs, the paragraph marks between them and the
  paragraph-property records Word writes with them are one ``insertion`` or ``deletion``,
  across paragraphs;
* **a move**'s two ends (every ``w:moveFrom`` and ``w:moveTo`` of one move name) are one
  ``move``, and paragraphs moved together -- touching where they went, and where they were --
  one move, a table moved with them (a deletion and an insertion) included;
* **a table's** structure edits one change each, with the text that went or came with them
  (round 55, R116; ``_table_joins``): a column deleted or inserted, a merge, a split, a row
  inserted or deleted -- ``Change.structure`` says which -- and not every record one person
  made in the table that minute; a cell's last mark does not run on into the next cell;
* **property changes** on their own are ``formatting`` (run, paragraph, numbering), ``section``
  or ``table`` changes.

Each :class:`Change` lists its revision ids, so it is accepted or rejected whole.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from ..edit.annotations import revision_kind
from ..oxml.xml import Element

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_P = _W + "p"
W_PPR = _W + "pPr"
W_RPR = _W + "rPr"
W_TBL = _W + "tbl"
W_TR = _W + "tr"
W_TC = _W + "tc"
_CONTAINERS = {_W + "ins": "ins", _W + "del": "del", _W + "moveFrom": "del", _W + "moveTo": "ins"}
_MOVES = (_W + "moveFrom", _W + "moveTo")
_RANGE_STARTS = {_W + "moveFromRangeStart": _W + "moveFromRangeEnd", _W + "moveToRangeStart": _W + "moveToRangeEnd"}
#: What a paragraph's walk passes through to its runs.
_TRANSPARENT = {_W + name for name in ("hyperlink", "smartTag", "customXml", "fldSimple", "sdt", "sdtContent",
                                       "dir", "bdo")}
#: Run content that is text a reader sees.
_CONTENT = {_W + name for name in ("t", "delText", "tab", "br", "cr", "drawing", "pict", "object", "sym",
                                   "noBreakHyphen", "softHyphen", "footnoteReference", "endnoteReference",
                                   "ptab", "instrText", "fldChar")}

#: Every kind :meth:`Document.revisions` gives a record, and what it is.
REVISION_KINDS = {
    "insertion": "text inserted (w:ins around runs)",
    "deletion": "text deleted (w:del around runs)",
    "move-from": "text moved away: the source of a move (w:moveFrom)",
    "move-to": "text moved here: the destination of a move (w:moveTo)",
    "paragraph-mark-insertion": "a paragraph break inserted: the paragraph's mark (w:pPr/w:rPr/w:ins)",
    "paragraph-mark-deletion": "a paragraph break deleted: accepting joins the paragraph to the next",
    "paragraph-mark-move-from": "a moved paragraph's mark, where it was",
    "paragraph-mark-move-to": "a moved paragraph's mark, where it went",
    "run-mark-insertion": "a run's own insertion record (w:r/w:rPr/w:ins), as for a field's runs",
    "run-mark-deletion": "a run's own deletion record (w:r/w:rPr/w:del)",
    "run-mark-move-from": "a run's own record of a move's source (w:r/w:rPr/w:moveFrom)",
    "run-mark-move-to": "a run's own record of a move's destination (w:r/w:rPr/w:moveTo)",
    "row-insertion": "a table row inserted (w:trPr/w:ins)",
    "row-deletion": "a table row deleted (w:trPr/w:del)",
    "run-properties": "character formatting changed (w:rPrChange: the old run properties)",
    "paragraph-properties": "paragraph formatting or style changed (w:pPrChange: the old properties); Word also "
                            "writes one on the last paragraph it adds at a story's end",
    "section-properties": "a section's page setup changed (w:sectPrChange)",
    "table-properties": "a table's properties changed (w:tblPrChange)",
    "row-properties": "a row's properties changed (w:trPrChange)",
    "cell-properties": "a cell's properties changed (w:tcPrChange), a merge's or split's spans too",
    "table-grid": "a table's column grid changed (w:tblGridChange)",
    "numbering": "a paragraph's list numbering changed (w:numberingChange)",
    "cell-insertion": "a table cell inserted (w:cellIns)",
    "cell-deletion": "a table cell deleted (w:cellDel)",
    "cell-merge": "table cells merged vertically (w:cellMerge)",
}

#: The kinds :func:`changes` gives.
CHANGE_KINDS = ("insertion", "deletion", "replacement", "move", "formatting", "table", "section")


@dataclass
class Change:
    """One logical change: the records one person made in one go (:func:`changes`)."""

    document: "Document" = field(repr=False, compare=False)
    #: ``insertion``, ``deletion``, ``replacement``, ``move``, ``formatting``, ``table`` or
    #: ``section``.
    kind: str
    author: str | None
    date: str | None
    #: The revision ids it is made of, in document order (``rev:12``...).
    revisions: list[str]
    #: The text inserted (a replacement's new text, a move's text); paragraphs end in ``\\n``.
    text: str = ""
    #: The text deleted (a replacement's old text).
    old_text: str = ""
    #: The paragraphs it touches, in order.
    paragraph_ids: list[str] = field(default_factory=list)
    #: The story it is in.
    story: str = "body"
    #: What a table edit was, when the change is one: ``column deletion``, ``column
    #: insertion``, ``merge``, ``split``, ``row deletion`` or ``row insertion`` (R116).
    structure: str | None = None

    @property
    def id(self) -> str:
        """The change's id: its first revision's (``rev:12``)."""
        return self.revisions[0]

    def accept(self) -> "EditResult":
        """Accept every record of the change, one undo step."""
        return self.document.accept(list(self.revisions))

    def reject(self) -> "EditResult":
        """Reject every record of the change, one undo step."""
        return self.document.reject(list(self.revisions))

    def __str__(self) -> str:
        if self.kind == "replacement":
            what = f"{self.old_text!r} -> {self.text!r}"
        elif self.kind == "deletion":
            what = repr(self.old_text)
        elif self.kind in ("insertion", "move"):
            what = repr(self.text)
        else:
            what = ", ".join(sorted({self.document.revision(r).kind for r in self.revisions}))
        return f"{self.kind} by {self.author}: {what}"


class _Union:
    def __init__(self) -> None:
        self.parent: dict[int, int] = {}

    def find(self, key: int) -> int:
        self.parent.setdefault(key, key)
        while self.parent[key] != key:
            self.parent[key] = self.parent[self.parent[key]]
            key = self.parent[key]
        return key

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[max(ra, rb)] = min(ra, rb)


def _who(element: Element) -> tuple[str | None, str | None]:
    return element.get(_W + "author"), element.get(_W + "date")


def _is_mark(element: Element) -> bool:
    parent = element.getparent()
    return parent is not None and parent.tag == W_RPR and parent.getparent() is not None \
        and parent.getparent().tag == W_PPR


def _run_has_content(run: Element) -> bool:
    return any(isinstance(child.tag, str) and child.tag in _CONTENT for child in run)


def _tokens(paragraph: Element) -> list[tuple[str, Element]]:
    """The paragraph's content in order: ``("rec", container)`` for a revision container at
    run level, ``("text", run)`` for an untracked run a reader sees."""
    out: list[tuple[str, Element]] = []

    def walk(node: Element) -> None:
        for child in node:
            tag = child.tag
            if not isinstance(tag, str) or tag == W_PPR:
                continue
            if tag in _CONTAINERS:
                out.append(("rec", child))
            elif tag == _W + "r":
                if _run_has_content(child):
                    out.append(("text", child))
            elif tag in _TRANSPARENT:
                walk(child)

    walk(paragraph)
    return out


def _move_names(root: Element) -> dict[int, str]:
    """``id()`` of every move record -> the name of the move range it is in."""
    out: dict[int, str] = {}
    open_ranges: dict[str, list[tuple[str, str]]] = {"from": [], "to": []}
    held = list(root.iter())
    for node in held:
        tag = node.tag
        if not isinstance(tag, str):
            continue
        if tag in _RANGE_STARTS:
            side = "from" if "moveFrom" in tag else "to"
            open_ranges[side].append((node.get(_W + "id") or "", node.get(_W + "name") or ""))
        elif tag in (_W + "moveFromRangeEnd", _W + "moveToRangeEnd"):
            side = "from" if "moveFrom" in tag else "to"
            wanted = node.get(_W + "id")
            open_ranges[side] = [r for r in open_ranges[side] if r[0] != wanted]
        elif tag in _MOVES:
            side = "from" if tag == _W + "moveFrom" else "to"
            if open_ranges[side]:
                out[id(node)] = open_ranges[side][-1][1]
            elif _is_mark(node):
                # Word puts the range's start after the paragraph's properties, past the
                # mark's own record: the mark is in the range its paragraph opens.
                paragraph = node.getparent().getparent().getparent()
                start_tag = _W + ("moveFromRangeStart" if side == "from" else "moveToRangeStart")
                start = next(iter(paragraph.iter(start_tag)), None)
                if start is not None:
                    out[id(node)] = start.get(_W + "name") or ""
    return out


def _texts(element: Element) -> str:
    return "".join(node.text or "" for node in element.iter(_W + "t", _W + "delText"))


def changes(document: "Document", *, author: str | None = None, kind: str | None = None) -> list["Change"]:
    """The document's revisions grouped into logical changes, in document order (module
    docstring); filtered by ``author`` and ``kind`` (one of :data:`CHANGE_KINDS`)."""
    if kind is not None and kind not in CHANGE_KINDS:
        raise ValueError(f"kind is one of {', '.join(CHANGE_KINDS)}")
    records = document._revision_elements()
    out: list[Change] = []
    by_part: dict[str, list[tuple[str, Element]]] = {}
    for part, identifier, element in records:
        by_part.setdefault(part, []).append((identifier, element))
    for part in document._parts():
        if part in by_part:
            out += _part_changes(document, part, by_part[part])
    if author is not None:
        out = [c for c in out if c.author == author]
    if kind is not None:
        out = [c for c in out if c.kind == kind]
    return out


def record_groups(document: "Document", part: str) -> list[list[Element]]:
    """The part's revision records, grouped as :func:`changes` groups them: what accepting or
    rejecting one of a table's structure records takes with it (R116)."""
    records = [(identifier, element) for where, identifier, element in document._revision_elements()
               if where == part]
    if not records:
        return []
    by_id = dict(records)
    return [[by_id[r] for r in change.revisions] for change in _part_changes(document, part, records)]


def _part_changes(document: "Document", part: str, records: list[tuple[str, Element]]) -> list[Change]:
    root = document.package.tree(part)
    held = [element for _, element in records]  # held: keys are the proxies' id()
    key = {id(element): k for k, element in enumerate(held)}
    union = _Union()
    for k in range(len(held)):
        union.find(k)

    names = _move_names(root)

    def join(a: Element, b: Element, *, same_sign: bool = False) -> None:
        """One change, when one person made both at once -- and a move only with its own
        records, a paragraph mark only with text of its own sign."""
        if id(a) not in key or id(b) not in key or _who(a) != _who(b):
            return
        if same_sign:
            # Across a paragraph mark: what one person inserted (or moved here) in one go,
            # several moved paragraphs and a table among them too.
            if _CONTAINERS.get(a.tag) != _CONTAINERS.get(b.tag):
                return
        elif (a.tag in _MOVES) != (b.tag in _MOVES) or names.get(id(a)) != names.get(id(b)):
            return
        union.union(key[id(a)], key[id(b)])

    # A move's ends, by name.
    first_of_name: dict[str, Element] = {}
    for element in held:
        name = names.get(id(element))
        if name:
            if name in first_of_name:
                union.union(key[id(first_of_name[name])], key[id(element)])
            else:
                first_of_name[name] = element
    # Records inside another record (an insertion deleted later, a run's formatting inside an
    # insertion) go with it when one person made both at once.
    for element in held:
        parent = element.getparent()
        while parent is not None:
            if id(parent) in key:
                join(parent, element)
                break
            parent = parent.getparent()
    # A table's structure: each edit to it on its own (R116) -- a column deletion, a merge, a
    # split, a row inserted or deleted, with the text that went or came with it -- and not
    # every record one person made in the table that minute.
    labels: dict[int, str] = {}
    claimed: set[int] = set()     # paragraph breaks a merge inserted
    for table in root.iter(W_TBL):
        _table_joins(table, key, join, union, labels, claimed)
    # Paragraphs: touching records, and the marks between paragraphs.
    paragraphs = [p for p in root.iter(W_P)]
    marks: dict[int, Element] = {}
    property_changes: dict[int, Element] = {}
    for paragraph in paragraphs:
        properties = paragraph.find(W_PPR)
        if properties is None:
            continue
        mark = properties.find(W_RPR)
        if mark is not None:
            for child in mark:
                if id(child) in key and _is_mark(child):
                    marks[id(paragraph)] = child
        change = properties.find(_W + "pPrChange")
        if change is not None and id(change) in key:
            property_changes[id(paragraph)] = change
    first_rec: dict[int, Element | None] = {}
    last_rec: dict[int, Element | None] = {}
    empty: dict[int, bool] = {}
    for paragraph in paragraphs:
        tokens = _tokens(paragraph)
        empty[id(paragraph)] = not tokens
        first_rec[id(paragraph)] = tokens[0][1] if tokens and tokens[0][0] == "rec" else None
        last_rec[id(paragraph)] = tokens[-1][1] if tokens and tokens[-1][0] == "rec" else None
        for (kind_a, a), (kind_b, b) in zip(tokens, tokens[1:]):
            if kind_a == kind_b == "rec":
                join(a, b)
        mark = marks.get(id(paragraph))
        if mark is not None and last_rec[id(paragraph)] is not None and id(mark) not in claimed:
            # (A break a merge inserted after a cell's text is the merge's, not that text's.)
            join(last_rec[id(paragraph)], mark, same_sign=True)
        change = property_changes.get(id(paragraph))
        if change is not None:
            # Word's form at a story's end: the last new paragraph's properties are recorded
            # as a change, with the insertion.
            for candidate in [mark, first_rec[id(paragraph)], last_rec[id(paragraph)]]:
                if candidate is not None and _who(candidate) == _who(change) \
                        and _CONTAINERS.get(candidate.tag) == "ins":
                    join(candidate, change)
                    break
    cell_of = {id(p): next(p.iterancestors(W_TC), None) for p in paragraphs}
    for previous, paragraph in zip(paragraphs, paragraphs[1:]):
        mark = marks.get(id(previous))
        if mark is None:
            continue
        before, after = cell_of[id(previous)], cell_of[id(paragraph)]
        if before is not None and after is not None and before is not after:
            # A cell's last mark does not run on into the next cell: two cells' edits are two
            # changes (R116); a row or column inserted or deleted whole is one by its own
            # records (_table_joins).
            continue
        # The mark is the break between the two: it touches the next paragraph's start.
        start = first_rec[id(paragraph)]
        if start is not None:
            join(mark, start, same_sign=True)
        elif empty[id(paragraph)] and id(paragraph) in marks:
            join(mark, marks[id(paragraph)], same_sign=True)
        elif empty[id(paragraph)] and id(paragraph) in property_changes:
            join(mark, property_changes[id(paragraph)])
    # Groups, in document order.
    groups: dict[int, list[int]] = {}
    for k in range(len(held)):
        groups.setdefault(union.find(k), []).append(k)
    index = document._index(part)
    out: list[Change] = []
    for members in sorted(groups.values(), key=lambda m: m[0]):
        elements = [held[k] for k in members]
        label = next((labels[id(e)] for e in elements if id(e) in labels), None)
        out.append(_change(document, part, index, [records[k][0] for k in members], elements, label))
    return out


_PROPERTY_KINDS = ("row-properties", "cell-properties", "table-properties")


def _nearest(element: Element, tag: str) -> Element | None:
    return next(element.iterancestors(tag), None)


def _sign(element: Element) -> str | None:
    """``ins`` or ``del`` for an insertion or deletion record of any level (text, mark, row,
    cell), else ``None``."""
    if element.tag == _W + "cellIns":
        return "ins"
    if element.tag == _W + "cellDel":
        return "del"
    return _CONTAINERS.get(element.tag) if element.tag not in _MOVES else None


def _row_cells(row: Element) -> list[tuple[Element, int]]:
    """The row's cells with the grid column each starts at (``w:gridBefore`` and the spans
    before it)."""
    properties = row.find(_W + "trPr")
    before = properties.find(_W + "gridBefore") if properties is not None else None
    at = int(before.get(_W + "val") or 0) if before is not None else 0
    out = []
    for cell in row.iter(W_TC):
        if _nearest(cell, W_TR) is not row:
            continue
        out.append((cell, at))
        span = cell.find(f"{_W}tcPr/{_W}gridSpan")
        at += int(span.get(_W + "val") or 1) if span is not None else 1
    return out


def _visible_texts(cell: Element) -> list[str]:
    """The cell's paragraphs' text as it stands (deleted text left out), the empty ones too."""
    out = []
    for paragraph in cell.iter(W_P):
        if _nearest(paragraph, W_TC) is not cell:
            continue
        out.append("".join(t.text or "" for t in paragraph.iter(_W + "t")
                           if not any(a.tag in (_W + "del", _W + "moveFrom") for a in t.iterancestors())))
    return out


def _table_joins(table: Element, key: dict[int, int], join, union: _Union, labels: dict[int, str],
                 claimed: set[int]) -> None:
    """Join a table's structure records with what each edit made with them (R116), labelling
    the edit for :class:`Change.structure`.

    * a **row** inserted or deleted: its record, and the records of the same sign in its cells;
      rows side by side inserted (or deleted) by one person at once are one (a table inserted);
    * a **column** deleted: each deleted cell (``w:cellDel`` whose content is deleted) with its
      deleted text and marks, and the column's cells together (same grid column); a column
      inserted: each new cell (``w:cellIns``, or a cell whose last mark is inserted, outside
      an inserted row) with what is inserted in it, the column's new cells together;
    * a **merge**: a merged-away cell (``w:cellDel`` keeping its text) with the copy of its text
      inserted in the cell it went into (and the paragraph breaks inserted there); a vertical
      merge's ``w:cellMerge`` records down the column, the text deleted in the continued cells
      and inserted in the first;
    * **property** records (row, cell, table, grid) together: one restructure (a split), with
      the new cell next to a split cell and a row a split inserted (its cells continue a
      vertical merge).
    """
    inside = [e for e in table.iter() if id(e) in key and _nearest(e, W_TBL) is table]
    if not inside:
        return
    rows = [row for row in table.iter(W_TR) if _nearest(row, W_TBL) is table]
    _held = list(table.iter(W_TC))     # held: the maps below are keyed by the proxies' id()
    row_of = {id(row): k for k, row in enumerate(rows)}
    by_cell: dict[int, list[Element]] = {}
    by_row: dict[int, list[Element]] = {}
    for element in inside:
        cell = _nearest(element, W_TC)
        if cell is not None and _nearest(cell, W_TBL) is table:
            by_cell.setdefault(id(cell), []).append(element)
        row = _nearest(element, W_TR)
        if row is not None and id(row) in row_of:
            by_row.setdefault(id(row), []).append(element)

    def label(element: Element, what: str) -> None:
        labels.setdefault(id(element), what)

    # Property records: one restructure.
    properties = [e for e in inside if revision_kind(e).startswith(_PROPERTY_KINDS)]
    for a, b in zip(properties, properties[1:]):
        join(a, b)
    # Rows inserted or deleted.
    row_record: dict[int, Element] = {}
    for row in rows:
        record = next((e for e in by_row.get(id(row), []) if revision_kind(e) in ("row-insertion", "row-deletion")
                       and e.getparent().getparent() is row), None)
        if record is None:
            continue
        row_record[id(row)] = record
        label(record, "row insertion" if _sign(record) == "ins" else "row deletion")
        for element in by_row[id(row)]:
            if element is not record and _sign(element) == _sign(record):
                join(record, element)
    for upper, lower in zip(rows, rows[1:]):
        a, b = row_record.get(id(upper)), row_record.get(id(lower))
        if a is not None and b is not None and _sign(a) == _sign(b):
            join(a, b)
    # Cells.
    deleted_cells: dict[int, list[Element]] = {}      # grid column -> deleted cells' records
    new_cells: dict[int, list[Element]] = {}          # grid column -> new cells' anchors
    merge_chain: dict[int, Element] = {}              # grid column -> the cellMerge above
    for k, row in enumerate(rows):
        cells = _row_cells(row)
        previous_kept: tuple[Element, int] | None = None
        merged_away: Element | None = None
        chain_now: dict[int, Element] = {}
        for cell, column in cells:
            records = by_cell.get(id(cell), [])
            own = [e for e in records if e.getparent() is not None and e.getparent().tag == _W + "tcPr"
                   and e.getparent().getparent() is cell]
            deletions = [e for e in own if e.tag == _W + "cellDel"]
            insertion = next((e for e in own if e.tag == _W + "cellIns"), None)
            vertical = next((e for e in own if e.tag == _W + "cellMerge"), None)
            content = [e for e in records if e not in own and not revision_kind(e).startswith(_PROPERTY_KINDS)]
            last = [p for p in cell.iter(W_P) if _nearest(p, W_TC) is cell][-1:]
            mark = _mark_record(last[0]) if last else None
            new_mark = mark if mark is not None and id(mark) in key and _sign(mark) == "ins" else None
            if deletions and (any(_sign(e) == "del" for e in content) or new_mark is not None):
                # A deleted column's cell (a cell merged away earlier keeps its own record).
                deletion = deletions[-1]
                label(deletion, "column deletion")
                for element in content:
                    if _sign(element) == "del":
                        join(deletion, element)
                deleted_cells.setdefault(column, []).append(deletion)
                deletions = deletions[:-1]
            if deletions:
                # A merged-away cell: its text kept, a copy inserted in the cell it went into.
                deletion = deletions[0]
                label(deletion, "merge")
                if merged_away is not None:
                    join(merged_away, deletion)
                merged_away = deletion
                if previous_kept is not None:
                    _join_copies(deletion, by_cell.get(id(previous_kept[0]), []),
                                 {t for t in _visible_texts(cell) if t}, join, claimed)
                continue
            if any(e.tag == _W + "cellDel" for e in own) and new_mark is None:
                merged_away = None
                continue
            merged_away = None
            if vertical is not None:
                label(vertical, "merge")
                if (vertical.get(_W + "vMerge") or "") == "cont" and column in merge_chain:
                    join(merge_chain[column], vertical)
                    head = merge_chain[column]
                    moved = {_texts(e) for e in content if _sign(e) == "del" and _texts(e)}
                    for element in content:
                        if _sign(element) == "del":
                            join(vertical, element)
                    head_cell = _nearest(head, W_TC)
                    if head_cell is not None:
                        _join_copies(head, by_cell.get(id(head_cell), []), moved, join, claimed)
                    chain_now[column] = head
                else:
                    chain_now[column] = vertical
            anchor = insertion if insertion is not None else new_mark
            if anchor is not None and id(row) not in row_record:
                # A new cell: a column's, or a split's (next to the split cell).
                label(anchor, "column insertion")
                for element in content:
                    if _sign(element) == "ins":
                        join(anchor, element)
                split_cell = previous_kept[0] if previous_kept is not None else None
                beside = [e for e in by_cell.get(id(split_cell), [])
                          if revision_kind(e) == "cell-properties"] if split_cell is not None else []
                if beside:
                    labels[id(anchor)] = "split"
                    label(beside[0], "split")
                    join(beside[0], anchor)
                else:
                    new_cells.setdefault(column, []).append(anchor)
            if anchor is None:
                previous_kept = (cell, column)
            # A row a split inserted: its cells continue a vertical merge from the row above.
            record = row_record.get(id(row))
            if record is not None and _sign(record) == "ins" and k > 0:
                continuing = cell.find(f"{_W}tcPr/{_W}vMerge")
                if continuing is not None and (continuing.get(_W + "val") or "continue") == "continue":
                    above = next((c for c, at in _row_cells(rows[k - 1]) if at == column), None)
                    changed = [e for e in by_cell.get(id(above), [])
                               if revision_kind(e) == "cell-properties"] if above is not None else []
                    if changed:
                        labels[id(record)] = "split"
                        label(changed[0], "split")
                        join(changed[0], record)
        merge_chain = chain_now
    for group in list(deleted_cells.values()) + list(new_cells.values()):
        for a, b in zip(group, group[1:]):
            join(a, b)
    # A grid change carries no author or date: w:tblGridChange has only an id (ECMA-376,
    # CT_TblGridChange; Word writes it so too, measured round 54).  It belongs with the change
    # to the same table that recorded it -- its w:tblPrChange, its other property records, or
    # the latest of its row and cell records -- whose author it is listed under (R112).
    for record in inside:
        if record.tag != _W + "tblGridChange" or _who(record) != (None, None):
            continue
        mine = [e for e in inside if e is not record and e.get(_W + "author") is not None
                and revision_kind(e).startswith(("row-", "cell-", "table-"))]
        if not mine:
            continue
        table_level = [e for e in mine if e.tag == _W + "tblPrChange"]
        other_properties = [e for e in mine if revision_kind(e).startswith(_PROPERTY_KINDS)]
        partner = (table_level[-1] if table_level else other_properties[0] if other_properties
                   else max(mine, key=lambda e: e.get(_W + "date") or ""))
        union.union(key[id(partner)], key[id(record)])


def _mark_record(paragraph: Element) -> Element | None:
    properties = paragraph.find(W_PPR)
    mark = properties.find(W_RPR) if properties is not None else None
    if mark is None:
        return None
    return next((child for child in mark if child.tag in (_W + "ins", _W + "del")), None)


def _join_copies(record: Element, target: list[Element], moved: set[str], join, claimed: set[int]) -> None:
    """A merge's record with what it inserted in the cell the text went into: the copies of
    the moved text and the paragraph breaks inserted there."""
    for element in target:
        if _sign(element) != "ins":
            continue
        if revision_kind(element) == "paragraph-mark-insertion":
            claimed.add(id(element))
            join(record, element)
        elif moved and _texts(element) in moved:
            join(record, element)


def _change(document: "Document", part: str, index, identifiers: list[str], elements: list[Element],
            structure: str | None = None) -> Change:
    kinds = [revision_kind(e) for e in elements]
    # The first record that names its author (a grid change names none: R112).
    author, date = next((_who(e) for e in elements if e.get(_W + "author") is not None), _who(elements[0]))
    moved = any(k.endswith(("move-from", "move-to")) for k in kinds)
    inserted = any(k in ("insertion", "paragraph-mark-insertion", "run-mark-insertion", "row-insertion",
                         "cell-insertion") for k in kinds)
    deleted = any(k in ("deletion", "paragraph-mark-deletion", "run-mark-deletion", "row-deletion", "cell-deletion")
                  for k in kinds)
    if moved:
        kind = "move"
    elif structure in ("merge", "split"):
        kind = "table"
    elif structure == "column deletion":
        kind = "deletion"
    elif structure == "column insertion":
        kind = "insertion"
    elif inserted and deleted:
        kind = "replacement"
    elif inserted:
        kind = "insertion"
    elif deleted:
        kind = "deletion"
    elif any(k.startswith(("table-", "row-", "cell-")) for k in kinds):
        kind = "table"
    elif kinds and all(k == "section-properties" for k in kinds):
        kind = "section"
    else:
        kind = "formatting"
    paragraphs: list[Element] = []
    for element in elements:
        paragraph = element if element.tag == W_P else next((a for a in element.iterancestors(W_P)), None)
        if paragraph is not None and not any(paragraph is p for p in paragraphs):
            paragraphs.append(paragraph)
    paragraph_ids = [entry.id for p in paragraphs for entry in [index.entry_for(p)] if entry is not None]
    text = _joined([e for e in elements if e.tag in (_W + "ins", _W + "moveTo") and not _is_mark(e)
                    and e.getparent().tag != W_RPR], paragraphs)
    old = _joined([e for e in elements if e.tag in (_W + "del",) and not _is_mark(e)
                   and e.getparent().tag != W_RPR], paragraphs)
    if kind == "move":
        text = text or _joined([e for e in elements if e.tag == _W + "moveFrom" and not _is_mark(e)
                                and e.getparent().tag != W_RPR], paragraphs)
    return Change(document, kind, author, date, identifiers, text, old, paragraph_ids, document._story_of(part),
                  structure)


def _joined(containers: list[Element], paragraphs: list[Element]) -> str:
    """The containers' text, a ``\\n`` between paragraphs."""
    out: list[str] = []
    last = None
    for container in containers:
        paragraph = next((a for a in container.iterancestors(W_P)), None)
        if last is not None and paragraph is not last:
            out.append("\n")
        out.append(_texts(container))
        last = paragraph
    return "".join(out)


__all__ = ["CHANGE_KINDS", "Change", "REVISION_KINDS", "changes", "record_groups"]
