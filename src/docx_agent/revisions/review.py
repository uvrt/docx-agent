"""Reviewing tracked changes: listing them, accepting and rejecting them.

``doc.revisions(author=, kind=, within=, since=, until=)`` lists every revision record --
``w:ins``, ``w:del``, ``w:moveFrom``, ``w:moveTo``, a paragraph mark's and a row's own
records, the property changes, the cell revisions -- in every story, with its kind,
author, date, text, range and the ids it affects.  ``doc.accept(...)`` and
``doc.reject(...)`` take one id (``rev:12``), a list, a range or block id (every revision in
it), or nothing (all, after the filters); ``accept_all`` and ``reject_all`` are shorthands.
Each is one undo step.

**What accepting and rejecting do** (after PowerTools' ``RevisionProcessor``, held to Word's
own Accept All and Reject All by ``tools/e3_probe.py`` and the oracle):

* an insertion: accepted, unwrapped; rejected, removed with its content;
* a deletion: accepted, removed; rejected, unwrapped (``w:delText`` back to ``w:t``);
* a move: both ends together, with their range markers -- accepted, the source goes and
  the destination stays; rejected, the reverse;
* a paragraph mark that goes (a deletion accepted, an insertion rejected) **joins its
  paragraph to the next**: the paragraph's content moves to the start of the next one,
  which keeps its own element, id and properties (measured: Word's accept keeps the
  second paragraph's paraId and properties; the first paragraph's properties survive
  only because Word, tracking the deletion, recorded them on the second in a
  ``w:pPrChange``).  At a table's edge, accepted, as Word's Accept All leaves it (round
  54, lane A; measured by lanes J and R): before a table (or a block-level content
  control) it joins the first paragraph of the table's first cell (of the control's
  content); a cell's last paragraph joins the first of the next cell of its row, which the
  cell takes in (a row's first cell's other blocks leave the table before the row, the
  table split there; the row ends short, ``w:gridAfter``, where another row keeps the
  column); a row's last cell goes, and its text -- which Word's Accept All takes out of
  view, saving it in a broken table -- moves after the row in paragraphs of its own, the
  table split there (a clean file, user decision; the result says so).  Otherwise, with no
  paragraph after it in its container, an empty paragraph goes and one with content keeps
  its mark -- Word keeps the body's last mark;
* a property change: accepted, the record goes; rejected, the old properties return;
* a row: accepted deletion or rejected insertion removes it (the table too, when it was
  the last); a cell: ``w:cellDel`` accepted or ``w:cellIns`` rejected removes it, and a
  grid column no remaining cell covers goes with it -- a deleted cell that keeps content or a
  paragraph mark is given to the cell before it instead (its columns and width: Word's Accept
  All, measured round 54); ``w:cellMerge`` accepted merges.

A record that belongs with others is accepted or rejected with them: a move's two ends; a
row's or cell's revisions made with it (same author and date) inside it; a paragraph's
content with its mark when both were made together.  Afterwards: relationships of the
pictures and hyperlinks removed are released, numbering instances the review left unused
are removed (a style an edit added stays, as in Word: style definitions are not revisions),
a footnote or endnote whose reference the review removed goes (as a comment does), and the
comments docx-agent wrote to describe structural table changes go once those are
resolved.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..edit import inline as _inline
from ..edit import text as _text
from ..edit.annotations import REVISION_KINDS, revision_kind, revision_text
from ..oxml.xml import Element, insert_in_order, make, remove
from . import track as _track

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document, EditResult
    from ..edit.ranges import TextRange

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_P = _W + "p"
W_TBL = _W + "tbl"
W_TR = _W + "tr"
W_TC = _W + "tc"
W_PPR = _W + "pPr"
W_RPR = _W + "rPr"

_CONTAINERS = frozenset(_track.CONTAINERS)
_PROPERTY_CHANGES = frozenset(_W + name for name in (
    "rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange", "tblGridChange",
    "numberingChange"))
_CELL = frozenset(_W + name for name in ("cellIns", "cellDel", "cellMerge"))
_MOVE_MARKERS = {
    _W + "moveFromRangeStart": _W + "moveFromRangeEnd",
    _W + "moveToRangeStart": _W + "moveToRangeEnd",
}

#: The text that opens a comment docx-agent writes to describe a structural table change
#: Word has no tracked form of its own for (ROADMAP.md, decision 7).
DESCRIBES = "Tracked structural change:"


def _kind_of(element: Element) -> str:
    """``mark`` (a paragraph mark's record), ``row``, ``run`` (a content container), ``cell``,
    ``property`` or ``other``."""
    tag = element.tag
    parent = element.getparent()
    if tag in _CONTAINERS:
        if parent is not None and parent.tag == W_RPR and parent.getparent() is not None \
                and parent.getparent().tag == W_PPR:
            return "mark"
        if parent is not None and parent.tag == _W + "trPr":
            return "row"
        if parent is not None and parent.tag in (W_RPR, _W + "numPr"):
            return "other"  # a run mark (w:rPr/w:ins) or a numbering insertion: records only
        return "run"
    if tag in _CELL:
        return "cell"
    if tag in _PROPERTY_CHANGES:
        return "property"
    return "other"


def _ancestor(node: Element, tag: str) -> Element | None:
    parent = node.getparent()
    while parent is not None and parent.tag != tag:
        parent = parent.getparent()
    return parent


# -- the listing ------------------------------------------------------------------------------


@dataclass(frozen=True)
class Revision:
    """One revision record (``rev:<w:id>``), live: it resolves its element on every access."""

    document: "Document"
    id: str

    def _locate(self) -> tuple[str, Element]:
        for part, identifier, element in self.document._revision_elements():
            if identifier == self.id:
                return part, element
        raise KeyError(f"no revision {self.id!r}")

    @property
    def kind(self) -> str:
        """What the record is (``docx_agent.revisions.changes.REVISION_KINDS`` says each):

        * text: ``insertion``, ``deletion``, ``move-from`` (a move's source), ``move-to``;
        * a paragraph's mark -- the break after it -- recorded on its own:
          ``paragraph-mark-insertion``, ``paragraph-mark-deletion`` (accepting joins the
          paragraph to the next), ``paragraph-mark-move-from``, ``paragraph-mark-move-to``;
        * a run's own record (a field's runs): ``run-mark-insertion``, ``run-mark-deletion``,
          ``run-mark-move-from``, ``run-mark-move-to``;
        * a table row: ``row-insertion``, ``row-deletion``;
        * a property change, the old properties kept: ``run-properties``,
          ``paragraph-properties`` (Word also writes one on the last paragraph it adds at a
          story's end), ``section-properties``, ``table-properties``, ``row-properties``,
          ``cell-properties``, ``table-grid``, ``numbering``;
        * a cell: ``cell-insertion``, ``cell-deletion``, ``cell-merge``.

        :meth:`Document.changes` groups the records one person made at once (a replacement's
        deletion and insertion, a new paragraph's text and mark) into one change."""
        return revision_kind(self._locate()[1])

    @property
    def author(self) -> str | None:
        """Who made the change (``w:author``)."""
        return self._locate()[1].get(_W + "author")

    @property
    def date(self) -> str | None:
        """When (``w:date``, ISO 8601), or ``None``."""
        return self._locate()[1].get(_W + "date")

    @property
    def story(self) -> str:
        """The story the revision is in."""
        return self.document._story_of(self._locate()[0])

    @property
    def text(self) -> str:
        """The text the record inserts, deletes or moves (empty for the others)."""
        return revision_text(self._locate()[1])

    @property
    def paragraph_id(self) -> str | None:
        """The id of the paragraph the revision is in (``None`` for a row's or table's own)."""
        part, element = self._locate()
        paragraph = element if element.tag == W_P else _ancestor(element, W_P)
        if paragraph is None:
            return None
        entry = self.document._index(part).entry_for(paragraph)
        return entry.id if entry is not None else None

    @property
    def ids(self) -> list[str]:
        """What the record affects: its paragraph's id; for a row's or cell's record, the
        row (``tr:``) and its table (``t:``); for a table's or grid's, the table."""
        part, element = self._locate()
        index = self.document._index(part)
        out: list[str] = []
        paragraph = _ancestor(element, W_P)
        if paragraph is not None:
            entry = index.entry_for(paragraph)
            if entry is not None:
                out.append(entry.id)
        row = element if element.tag == W_TR else _ancestor(element, W_TR)
        table = _ancestor(element, W_TBL)
        if table is not None and paragraph is None:
            entry = index.entry_for(table)
            if row is not None and entry is not None:
                out += [row_id for row_element, row_id in entry.rows if row_element is row]
            if entry is not None:
                out.append(entry.id)
        return out

    @property
    def range(self) -> "TextRange | None":
        """The characters an insertion, deletion or move covers, in the ``markup`` view
        (``None`` for the other records)."""
        from ..edit.ranges import TextRange

        part, element = self._locate()
        if _kind_of(element) != "run":
            return None
        paragraph = _ancestor(element, W_P)
        entry = self.document._index(part).entry_for(paragraph) if paragraph is not None else None
        if entry is None:
            return None
        held = list(element.iter())  # held: the proxies' id()s must not be reused
        inside = {id(node) for node in held}
        atoms = _text.atoms(paragraph, "markup")
        offsets = [k for k, atom in enumerate(atoms) if id(atom.node) in inside]
        del held
        if not offsets:
            return TextRange(self.document, entry.id, 0, entry.id, 0, "markup")
        return TextRange(self.document, entry.id, offsets[0], entry.id, offsets[-1] + 1, "markup")

    def accept(self) -> "EditResult":
        """Accept this revision: :meth:`Document.accept`."""
        return self.document.accept(self.id)

    def reject(self) -> "EditResult":
        """Reject this revision: :meth:`Document.reject`."""
        return self.document.reject(self.id)

    def __repr__(self) -> str:
        return f"<Revision {self.id} {self.kind} by {self.author!r}>"


# -- the engine -------------------------------------------------------------------------------


class _Review:
    """One accept or reject over a set of records, in one part."""

    def __init__(self, document: "Document", part: str, accept: bool) -> None:
        self.document = document
        self.part = part
        self.accept = accept
        self.root = document.package.tree(part)
        self.released: list[str] = []
        self.removed_paragraphs: list[Element] = []
        self.joined: list[tuple[Element, Element]] = []
        self.tables: set[int] = set()
        self._held: list[Element] = []
        self._groups: dict[int, list[Element]] | None = None
        #: Every table's grid as the review found it, and the cells it removed: the grid
        #: columns only removed cells covered go at the end.
        self.maps = {id(table): (table, _grid_map(table)) for table in self.root.iter(W_TBL)}
        #: Round 53 (R107): the tables holding a record this review accepts or rejects: each
        #: fits its grid after it, once no structure change of it is left (:meth:`_tidy`).
        self.touched_tables: set[int] = set()
        #: Comments whose reference the review removed with the content around it: Word
        #: removes the comment then (measured: a comment on rejected inserted text goes).
        self.dropped_comments: set[str] = set()
        #: Notes whose reference the review removed (``footnote``/``endnote``, id): Word
        #: removes a note with its reference.
        self.dropped_notes: set[tuple[str, str]] = set()
        self.removed_cells: dict[int, set[int]] = {}
        #: Tables whose grid a rejected w:tblGridChange set back whole: no column is
        #: removed from it afterwards.
        self.restored_grids: set[int] = set()
        #: Tables whose cells' spans a rejected w:tcPrChange set back: their grid keeps
        #: only the lines a cell's edge is on (a split's column goes, as in Word).
        self.respanned: dict[int, Element] = {}
        #: Round 51 (R103): cells whose only paragraph's deleted mark this review accepted,
        #: by their row: a row all of whose cells are such goes (:meth:`_emptied_rows`).
        self.emptied: dict[int, tuple[Element, set[int]]] = {}
        #: Round 54 (lane A): the tables a deleted mark at a cell's end split (a part's
        #: ``id`` -> the table the review found, whose grid map holds its rows), and the rows
        #: that gave up a cell to it (they end short where the grid keeps its columns).
        self.origin: dict[int, Element] = {}
        self.grid_group: dict[int, int] = {}
        #: The blocks a row's last cell's text moved out in (its text, the paragraphs):
        #: said in the result (rule 7), since Word's Accept All shows no such text.
        self.moved_out: list[tuple[str, list[Element]]] = []
        self.edged_rows: set[int] = set()

    # -- grouping ---------------------------------------------------------------------------

    def expand(self, elements: list[Element]) -> list[Element]:
        """The records that go with ``elements``: a move's both ends and markers, a row's or
        cell's or table's records made with it, a paragraph's content with its mark."""
        out: list[Element] = []
        seen: set[int] = set()

        def add(node: Element) -> None:
            if id(node) not in seen:
                seen.add(id(node))
                out.append(node)
                self._held.append(node)

        # Held: lxml keeps one proxy per element only while it is referenced, and the
        # order is keyed by the proxies' id().
        self._all = list(self.root.iter())
        order = {id(node): k for k, node in enumerate(self._all)}
        for element in elements:
            add(element)
            kind = _kind_of(element)
            # A table's old grid (Word writes it with an id alone) goes with any record of
            # its table's structure: a table is reviewed whole.
            table = _ancestor(element, W_TBL)
            while table is not None:
                self.touched_tables.add(id(table))
                outer = table.getparent()
                table = _ancestor(outer, W_TBL) if outer is not None else None
            table = _ancestor(element, W_TBL)
            if table is not None and element.tag != _W + "tblGridChange" and (
                    kind in ("cell", "row", "mark", "property")):
                grid = table.find(f"{_W}tblGrid/{_W}tblGridChange")
                if grid is not None:
                    add(grid)
            if element.tag in (_W + "moveFrom", _W + "moveTo"):
                for node in self._move(element, order):
                    add(node)
            elif kind == "row":
                row = element.getparent().getparent()
                for node in self._records(row):
                    if _made_with(node, element):
                        add(node)
            elif kind in ("cell", "property") and element.tag in ({_W + "tcPrChange"} | _CELL):
                table = _ancestor(element, W_TBL)
                if table is not None:
                    # Round 55 (R116): the records of the edit it was part of -- a column
                    # deletion, a merge, a split -- as word_changes lists it, not every record
                    # one person made in the table that minute.
                    for node in self._edit_of(element):
                        add(node)
                    # The old grid (Word writes it with an id alone) goes with the cells'
                    # records: a table's structure is reviewed whole.
                    grid = table.find(f"{_W}tblGrid/{_W}tblGridChange")
                    if grid is not None:
                        add(grid)
            elif kind == "mark":
                paragraph = _ancestor(element, W_P)
                for node in self._records(paragraph):
                    if _kind_of(node) == "run" and _made_with(node, element):
                        add(node)
        return out

    def _edit_of(self, element: Element) -> list[Element]:
        """The records listed with ``element`` as one change (:func:`changes.record_groups`)."""
        if self._groups is None:
            from .changes import record_groups

            self._groups = {}
            for group in record_groups(self.document, self.part):
                self._held.extend(group)
                for node in group:
                    self._groups[id(node)] = group
        return self._groups.get(id(element), [element])

    def _records(self, scope: Element) -> list[Element]:
        return [node for node in scope.iter() if isinstance(node.tag, str) and node.tag in REVISION_KINDS]

    def _move(self, element: Element, order: dict[int, int]) -> list[Element]:
        """Every record and range marker of the move ``element`` belongs to."""
        name = self._move_name(element, order)
        if name is None:
            return []
        out = []
        for start_tag, end_tag in _MOVE_MARKERS.items():
            for start in self.root.iter(start_tag):
                if start.get(_W + "name") != name:
                    continue
                end = next((e for e in self.root.iter(end_tag) if e.get(_W + "id") == start.get(_W + "id")), None)
                out.append(start)
                if end is not None:
                    out.append(end)
                low = order.get(id(start), -1)
                high = order.get(id(end), low) if end is not None else low
                inside = [node for node in self._all if low <= order[id(node)] <= high]
                for node in inside:
                    if node.tag in (_W + "moveFrom", _W + "moveTo"):
                        out.append(node)
                # The marks of the paragraphs the range runs through move with it (they sit
                # before the start marker, in the paragraph's properties).
                for paragraph in {id(p): p for p in (_ancestor(n, W_P) if n.tag != W_P else n
                                                     for n in inside) if p is not None}.values():
                    record = _track.mark_record(paragraph)
                    if record is not None and record.tag in (_W + "moveFrom", _W + "moveTo"):
                        out.append(record)
        return out

    def _move_name(self, element: Element, order: dict[int, int]) -> str | None:
        position = order.get(id(element))
        if position is None:
            return None
        start_tag = _W + ("moveFromRangeStart" if element.tag == _W + "moveFrom" else "moveToRangeStart")
        end_tag = _MOVE_MARKERS[start_tag]
        best = None
        for start in self.root.iter(start_tag):
            if order.get(id(start), 10 ** 9) > position:
                continue
            end = next((e for e in self.root.iter(end_tag) if e.get(_W + "id") == start.get(_W + "id")), None)
            if end is None or order.get(id(end), -1) >= position:
                best = start.get(_W + "name")
        return best

    # -- applying ---------------------------------------------------------------------------

    def run(self, elements: list[Element]) -> None:
        elements = [e for e in elements if e.getparent() is not None]
        rows = [e for e in elements if _kind_of(e) == "row"]
        cells = [e for e in elements if _kind_of(e) == "cell"]
        runs = [e for e in elements if _kind_of(e) == "run"]
        marks = [e for e in elements if _kind_of(e) == "mark"]
        properties = [e for e in elements if _kind_of(e) == "property" and e.tag != _W + "tblGridChange"]
        grids = [e for e in elements if e.tag == _W + "tblGridChange"]
        others = [e for e in elements if _kind_of(e) == "other"]
        markers = [e for e in elements if e.tag in _MOVE_MARKERS or e.tag in _MOVE_MARKERS.values()]

        for element in rows:
            self._row(element)
        self._cells(cells)
        for element in runs:
            self._container(element)
        for element in properties:
            self._property(element)
        for element in grids:
            self._grid(element)
        for element in others:
            if element.getparent() is not None:
                self._drop(element)
        for element in markers:
            if element.getparent() is not None:
                remove(element)
        # In document order: a mark at a cell's end that takes the next cell in hands that
        # cell's last paragraph (and its mark) on to the next record (Accept All's chains).
        order = {id(node): k for k, node in enumerate(self.root.iter())}
        for element in sorted(marks, key=lambda e: order.get(id(e), -1)):
            self._mark(element)
        self._tidy()

    def _drop(self, record: Element) -> None:
        """Remove a record, and the property element it leaves empty."""
        parent = record.getparent()
        remove(record)
        if parent is not None and parent.tag in (W_RPR, _W + "trPr") and not len(parent):
            owner = parent.getparent()
            remove(parent)
            if owner is not None and owner.tag == W_PPR and not len(owner):
                remove(owner)

    def _row(self, record: Element) -> None:
        if record.getparent() is None:
            return
        row = record.getparent().getparent()
        goes = (record.tag in _track.DELETED) == self.accept
        table = _ancestor(row, W_TBL)
        if goes:
            self.released += _inline.relationship_ids(row)
            self.removed_paragraphs += list(row.iter(W_P))
            self._removing(row)
            remove(row)
            if table is not None and table.find(W_TR) is None and not any(
                    child.tag in (_W + "sdt", _W + "customXml") for child in table):
                remove(table)
        else:
            self._drop(record)

    def _cells(self, records: list[Element]) -> None:
        by_table: dict[int, tuple[Element, list[Element]]] = {}
        for record in records:
            if record.getparent() is None:
                continue
            table = _ancestor(record, W_TBL)
            if table is None:
                continue
            by_table.setdefault(id(table), (table, []))[1].append(record)
        for table, group in by_table.values():
            self.tables.add(id(table))
            removed = self.removed_cells.setdefault(id(table), set())
            for record in group:
                cell = _ancestor(record, W_TC)
                if record.tag == _W + "cellMerge":
                    if self.accept:
                        merge = record.get(_W + "vMerge")
                        value = {"rest": "restart", "cont": "continue"}.get(merge)
                    else:
                        value = {"rest": "restart", "cont": "continue"}.get(record.get(_W + "vMergeOrig") or "")
                    properties = record.getparent()
                    current = properties.find(_W + "vMerge")
                    if current is not None:
                        remove(current)
                    if value is not None:
                        node = make("w:vMerge")
                        if value == "restart":
                            node.set(_W + "val", "restart")
                        insert_in_order(properties, node)
                    remove(record)
                    continue
                goes = (record.tag == _W + "cellDel") == self.accept
                if goes and cell is not None:
                    # Round 54 (lane M): a deleted cell whose content is kept is given to the
                    # cell before it (its columns and width), its content going -- Word's
                    # Accept All (measured): a tracked merge's form.
                    before = _cell_before(cell) if self.accept and _content_kept(cell) else None
                    if before is not None:
                        _absorb(before, cell)
                    else:
                        removed.add(id(cell))
                    self._held.append(cell)
                    self._removing(cell)
                    self.released += _inline.relationship_ids(cell)
                    self.removed_paragraphs += list(cell.iter(W_P))
                    remove(cell)
                else:
                    remove(record)

    def _container(self, record: Element) -> None:
        if record.getparent() is None:
            return
        shown = record.tag in _track.INSERTED
        if shown == self.accept:
            # Keep the content: unwrap.
            if record.tag in _track.DELETED:
                _track.spell_shown(record)
            for child in list(record):
                record.addprevious(child)
            remove(record)
        else:
            self.released += _inline.relationship_ids(record)
            for node in record.iter(_W + "commentReference", _W + "footnoteReference", _W + "endnoteReference"):
                self._orphan(node)
            remove(record)

    def _orphan(self, node: Element) -> None:
        """A comment or note reference the review removes with its content: the comment or
        note goes too."""
        if node.get(_W + "id") is None:
            return
        if node.tag == _W + "commentReference":
            self.dropped_comments.add(node.get(_W + "id"))
        elif node.tag in (_W + "footnoteReference", _W + "endnoteReference"):
            self.dropped_notes.add((node.tag[len(_W):-len("Reference")], node.get(_W + "id")))

    def _removing(self, element: Element) -> None:
        for node in element.iter(_W + "commentReference", _W + "footnoteReference", _W + "endnoteReference"):
            self._orphan(node)

    def _property(self, record: Element) -> None:
        if record.getparent() is None:
            return
        owner = record.getparent()
        if self.accept:
            remove(record)
            if owner.tag in (_W + "trPr", _W + "tcPr") and not len(owner):
                remove(owner)
                return
            _remove_if_empty(owner)
            return
        old = next((child for child in record if isinstance(child.tag, str)), None)
        tag = record.tag
        keep: set[str]
        if tag == _W + "pPrChange":
            keep = {W_RPR, _W + "sectPr"}
        elif tag == _W + "rPrChange":
            keep = set(_track.MARKS)
        elif tag == _W + "trPrChange":
            keep = {_W + "ins", _W + "del"}
        elif tag == _W + "tcPrChange":
            keep = set(_CELL)
            current = owner.find(_W + "gridSpan")
            before = old.find(_W + "gridSpan") if old is not None else None
            if (current.get(_W + "val") if current is not None else "1") != \
                    (before.get(_W + "val") if before is not None else "1"):
                table = _ancestor(owner, W_TBL)
                if table is not None:
                    self.respanned[id(table)] = table
        elif tag == _W + "sectPrChange":
            # Word records only what changed (measured), and rejects by setting what it
            # recorded over what is there (measured: a recorded <w:cols w:space=.../> left
            # Word's columns at two -- it writes <w:cols w:num="1"/> itself).  So: each
            # recorded kind's attributes are set over the element's, its children (a cols'
            # columns, a notePr's settings) replace the element's; a kind recorded in its
            # "off" form was absent before (edit/sections.py) and goes; an attribute left at
            # the default docx-agent recorded for one that was absent goes too.
            from ..edit.sections import restore_kind

            remove(record)
            for child in (list(old) if old is not None else []):
                if isinstance(child.tag, str):
                    restore_kind(owner, child)
            return
        elif tag == _W + "numberingChange":
            remove(record)
            return
        else:
            keep = set()
        remove(record)
        for child in list(owner):
            if isinstance(child.tag, str) and child.tag not in keep:
                remove(child)
        for child in (list(old) if old is not None else []):
            if isinstance(child.tag, str):
                insert_in_order(owner, copy.deepcopy(child))
        if tag == _W + "trPrChange":
            # Word records a row's old state with an explicit gridAfter of 0: no skip at all.
            for node in owner.findall(_W + "gridAfter") + owner.findall(_W + "gridBefore"):
                if (node.get(_W + "val") or "0") == "0":
                    remove(node)
        if owner.tag in (_W + "trPr", _W + "tcPr") and not len(owner):
            remove(owner)
            return
        _remove_if_empty(owner)

    def _grid(self, record: Element) -> None:
        if record.getparent() is None:
            return
        grid = record.getparent()
        if self.accept:
            remove(record)
            return
        old = record.find(_W + "tblGrid")
        remove(record)
        table = grid.getparent()
        if table is not None:
            self.restored_grids.add(id(table))
        for column in grid.findall(_W + "gridCol"):
            remove(column)
        for column in (old.findall(_W + "gridCol") if old is not None else []):
            insert_in_order(grid, copy.deepcopy(column))

    def _mark(self, record: Element) -> None:
        if record.getparent() is None:
            return
        paragraph = _ancestor(record, W_P)
        goes = (record.tag in _track.DELETED) == self.accept
        if not goes or paragraph is None:
            self._drop(record)
            return
        if paragraph.getparent() is None:
            return  # gone with what held it (a note whose reference the review removed)
        following = _track.next_paragraph(paragraph)
        if following is None and self.accept and record.tag in _track.DELETED and (
                self._into_table(paragraph) or self._cell_end(paragraph, record)):
            return
        if following is None:
            cell = paragraph.getparent()
            empty = not _has_content(paragraph)
            blocks = [c for c in cell if isinstance(c.tag, str) and c.tag in (W_P, W_TBL, _W + "sdt")]
            if empty and cell.tag == W_TC and blocks == [paragraph] and record.tag in _track.INSERTED \
                    and _ancestor(cell, W_TBL) is not None:
                # Word's own form of an inserted column: rejecting the cell's inserted mark
                # removes the cell (measured), and the grid column it leaves.
                table = _ancestor(cell, W_TBL)
                self.removed_cells.setdefault(id(table), set()).add(id(cell))
                self.tables.add(id(table))
                self._held.append(cell)
                self.released += _inline.relationship_ids(cell)
                remove(cell)
                self.removed_paragraphs.append(paragraph)
                return
            # An empty paragraph whose mark goes, goes -- a section break's too (a break
            # inserted and rejected, or deleted and accepted: E4) -- unless it is all its
            # cell has.
            if empty and (len(blocks) > 1 or cell.tag != W_TC) \
                    and (paragraph.find(f"{W_PPR}/{_W}sectPr") is None or cell.tag == _W + "body"):
                self.removed_paragraphs.append(paragraph)
                self.released += _inline.relationship_ids(paragraph)
                self._removing(paragraph)
                _keep_markers(paragraph)
                remove(paragraph)
                return
            if empty and cell.tag == W_TC and blocks == [paragraph] and record.tag in _track.DELETED:
                row = _ancestor(cell, W_TR)
                if row is not None:
                    self._held += [row, cell]
                    self.emptied.setdefault(id(row), (row, set()))[1].add(id(cell))
            self._drop(record)
            return
        self._join(paragraph, following)

    def _join(self, paragraph: Element, following: Element) -> None:
        """``paragraph`` joined to ``following``: its content moves to the start of
        ``following``, which keeps its element, id and properties (Word's accept)."""
        properties = following.find(W_PPR)
        at = 1 if properties is not None else 0
        for child in [c for c in paragraph if not (isinstance(c.tag, str) and c.tag == W_PPR)]:
            following.insert(at, child)
            at += 1
        self.joined.append((paragraph, following))
        self.removed_paragraphs.append(paragraph)
        remove(paragraph)

    # -- a deleted mark at a table's edge (round 54, lane A) -----------------------------------
    #
    # What Word's Accept All makes of it (measured by lanes J and R, ``feedback-round54``;
    # :mod:`.edges` is the final view's reading of the same rules), in a clean file.

    def _into_table(self, paragraph: Element) -> bool:
        """A paragraph whose mark goes just before a table joins the first paragraph of the
        table's first cell (a nested table's, where the cell begins with one): its content
        first, under the cell paragraph's properties.  A paragraph ending a section stays.
        Before a block-level content control (or ``w:customXml``) it joins the first
        paragraph of its content, as the final view reads the content as the container's
        (docx2svg's revision probe, ``mark before a block control``: Word's Accept All)."""
        if paragraph.find(f"{W_PPR}/{_W}sectPr") is not None:
            return False
        block = _next_block(paragraph)
        if block is None or block.tag not in (W_TBL, _W + "sdt", _W + "customXml"):
            return False
        target = _first_in([block])
        if target is None:
            return False
        self._join(paragraph, target)
        return True

    def _cell_end(self, paragraph: Element, record: Element) -> bool:
        """A cell's last paragraph whose mark goes: the cell takes the next cell of its row
        in (its last paragraph joined to that cell's first, the rest after; a row's first
        cell's other blocks leave the table before the row, the table split there); a cell
        with no next cell in its row goes, and what it holds -- which Word's Accept All takes
        out of view and saves in a broken table after the row (a zero-span cell) -- goes
        after the row in paragraphs of its own, the table split there (user decision
        2026-11-13: a clean file)."""
        from ..edit.tables import _cells
        from .edges import cell_deleted

        cell = next(paragraph.iterancestors(W_TC), None)
        if cell is None or cell_deleted(cell) or _last_paragraph(cell) is not paragraph:
            return False
        row = cell.getparent()
        table = row.getparent() if row is not None else None
        if row is None or row.tag != W_TR or table is None or table.tag != W_TBL \
                or row.find(f"{_W}trPr/{_W}del") is not None:
            return False  # a cell or row inside a content control, a row being deleted
        shown = [c for c, _, _ in _cells(row) if not cell_deleted(c)]
        k = shown.index(cell)
        origin = self.origin.get(id(table), table)
        self._held += [row, cell, table]
        if k + 1 < len(shown):
            taken = shown[k + 1]
            if taken.getparent() is not row:
                return False
            # A row's first cell's blocks before its last paragraph leave the table.
            leaving = [c for c in paragraph.itersiblings(preceding=True)
                       if not (isinstance(c.tag, str) and c.tag == _W + "tcPr")][::-1] \
                if k == 0 and paragraph.getparent() is cell else []
            if not any(isinstance(c.tag, str) and c.tag in (W_P, W_TBL, _W + "sdt", _W + "customXml")
                       for c in leaving):
                leaving = []
            blocks = [c for c in taken if not (isinstance(c.tag, str) and c.tag == _W + "tcPr")]
            head = next((c for c in blocks if isinstance(c.tag, str) and c.tag in (W_P, W_TBL, _W + "sdt")), None)
            anchor = paragraph
            for child in blocks:
                anchor.addnext(child)
                anchor = child
            if head is not None and head.tag == W_P:
                self._join(paragraph, head)
            else:
                self._drop(record)
            self._held.append(taken)
            self.removed_cells.setdefault(id(origin), set()).add(id(taken))
            remove(taken)
            self.edged_rows.add(id(row))
            if leaving:
                self._out(table, row, leaving, before=True)
            return True
        # No next cell: the cell goes; what it holds goes after the row.
        self._drop(record)
        items = [c for c in cell if isinstance(c.tag, str) and c.tag != _W + "tcPr"]
        empty = [c for c in items if c.tag == W_P and not _has_content(c)]
        # Its paragraphs with nothing in them go with it; nothing moves when that is all.
        moving = [c for c in items if c not in empty] if any(
            c.tag in (W_P, W_TBL, _W + "sdt", _W + "customXml") for c in items if c not in empty) else []
        self.removed_paragraphs += [p for p in cell.iter(W_P) if not any(p in c.iter(W_P) for c in moving)]
        self._removing_kept(cell, moving)
        kept_ids = {rid for c in moving for rid in _inline.relationship_ids(c)}
        self.released += [rid for rid in _inline.relationship_ids(cell) if rid not in kept_ids]
        self.removed_cells.setdefault(id(origin), set()).add(id(cell))
        self.edged_rows.add(id(row))
        for child in moving:
            remove(child)
        remove(cell)
        if moving:
            self._out(table, row, moving, before=False)
            paragraphs = [p for c in moving for p in ([c] if c.tag == W_P else list(c.iter(W_P)))]
            self._held += paragraphs
            text = " ".join(" ".join(_text.paragraph_text(p, "current") for p in paragraphs).split())
            self.moved_out.append((text, paragraphs))
        return True

    def _removing_kept(self, cell: Element, kept: list[Element]) -> None:
        """References in ``cell`` that go with it (not in ``kept``, which moves out)."""
        inside = {id(n) for c in kept if isinstance(c.tag, str) for n in c.iter()}
        for node in cell.iter(_W + "commentReference", _W + "footnoteReference", _W + "endnoteReference"):
            if id(node) not in inside:
                self._orphan(node)

    def _out(self, table: Element, row: Element, blocks: list[Element], before: bool) -> None:
        """``blocks`` out of ``table``, before or after ``row``: the table split there
        (each part keeps the table's properties and its own grid, :meth:`_tidy`)."""
        rows = [r for r in table if isinstance(r.tag, str) and r.tag == W_TR]
        cut = rows.index(row) + (0 if before else 1)
        if cut == 0:
            for block in blocks:
                table.addprevious(block)
            return
        if cut < len(rows):
            part = make("w:tbl")
            for child in table:
                if isinstance(child.tag, str) and child.tag in (_W + "tblPr", _W + "tblGrid"):
                    copied = copy.deepcopy(child)
                    for record in list(copied.iter(_W + "tblPrChange", _W + "tblGridChange")):
                        remove(record)
                    part.append(copied)
            first = rows[cut]
            moving = [first] + list(first.itersiblings())
            for child in moving:
                part.append(child)
            table.addnext(part)
            origin = self.origin.get(id(table), table)
            self.origin[id(part)] = origin
            self.grid_group[id(part)] = id(part) if before else self.grid_group.get(id(table), id(origin))
            self._held.append(part)
        anchor = table
        for block in blocks:
            anchor.addnext(block)
            anchor = block

    def _emptied_rows(self) -> None:
        """A row whose every cell's only paragraph had its deletion accepted here, nothing
        left in it, and no record of its own left, goes -- the table with its last row: Word
        accepting its own form of a row one inserted and then deleted while it held what was
        there before (``w:trPr/w:ins`` kept, the cells' content and marks deleted: measured,
        ``tools/rows_probe.py``, ``MinePlain``)."""
        for row, cells in self.emptied.values():
            if row.getparent() is None or row.find(f"{_W}trPr/{_W}ins") is not None \
                    or row.find(f"{_W}trPr/{_W}del") is not None:
                continue
            own = [c for c in row.iter(W_TC) if _ancestor(c, W_TR) is row]
            if not own or any(id(c) not in cells or any(_has_content(p) for p in c.iter(W_P)) for c in own):
                continue
            table = _ancestor(row, W_TBL)
            self.removed_paragraphs += list(row.iter(W_P))
            self._removing(row)
            _keep_markers(row)
            remove(row)
            if table is not None and table.find(W_TR) is None and not any(
                    child.tag in (_W + "sdt", _W + "customXml") for child in table):
                remove(table)

    def _tidy(self) -> None:
        self._emptied_rows()
        for key, removed in self.removed_cells.items():
            if removed and key in self.maps and key not in self.restored_grids:
                table, grid = self.maps[key]
                # A table a deleted mark at a cell's end split: the parts a row's first cell's
                # blocks leaving it made each have a grid of their own (as Word draws and
                # saves them); the parts the text of a row's last cell moving out made keep
                # the one they were split from (Word keeps that table whole).
                groups: dict[int, list[Element]] = {}
                for part in [table] + [p for p in self._held if p.tag == W_TBL and self.origin.get(id(p)) is table]:
                    if part.getparent() is not None and not any(part is q for g in groups.values() for q in g):
                        groups.setdefault(self.grid_group.get(id(part), id(table)), []).append(part)
                for parts in groups.values():
                    _fix_grid(parts, grid, removed, self.edged_rows)
        for key, table in self.respanned.items():
            if key not in self.restored_grids and table.getparent() is not None:
                _normalise_grid(table)
        # Round 53 (R107): a table that fitted its grid fits it after the review.  A grid
        # set back whole (a w:tblGridChange rejected) that its rows no longer fit -- recorded
        # by a tracked change after a split, which records none -- is rebuilt from the cells'
        # widths, as a split's rejection rebuilds it; where it cannot be, or several tracked
        # structure changes of one table left its rows apart, the grid is fitted to them as
        # Word writes it when it saves such a table (table.fit_grid).
        from ..edit.tables import fit_grid

        for key in self.touched_tables:
            found = self.maps.get(key)
            if found is None:
                continue
            table = found[0]
            if table.getparent() is None or _fits(table) or any(
                    node.tag in _STRUCTURE for node in table.iter() if isinstance(node.tag, str)):
                continue
            if key in self.restored_grids:
                _normalise_grid(table)
            if not _fits(table):
                fit_grid(table)
        for paragraph in list(self.root.iter(W_P)):
            self.released += _prune(paragraph)


def _remove_if_empty(owner: Element) -> None:
    """Remove a run's or paragraph's (or a paragraph mark's) properties a review emptied."""
    if len(owner) or owner.tag not in (W_RPR, W_PPR):
        return
    parent = owner.getparent()
    if parent is None or parent.tag not in (_W + "r", W_P, W_PPR):
        return
    remove(owner)
    if parent.tag == W_PPR:
        _remove_if_empty(parent)


_BLOCKS = frozenset({W_P, W_TBL, _W + "sdt", _W + "customXml"})


def _next_block(node: Element) -> Element | None:
    """The block after ``node`` in its container (a paragraph, table, content control,
    ``w:customXml``, or the body's ``w:sectPr``), range markers passed over."""
    following = node.getnext()
    while following is not None:
        if isinstance(following.tag, str) and (following.tag in _BLOCKS or following.tag == _W + "sectPr"):
            return following
        following = following.getnext()
    return None


def _first_paragraph(table: Element) -> Element | None:
    """The first paragraph of ``table``'s first cell as accepting leaves the table (a row
    or cell whose deletion is pending passed over), a nested table's where the cell begins
    with one, a content control's content seen through; ``None`` where there is none."""
    from ..edit.tables import _cells
    from .edges import cell_deleted

    for row in table.iter(W_TR):
        if _ancestor(row, W_TBL) is not table or row.find(f"{_W}trPr/{_W}del") is not None:
            continue
        shown = [c for c, _, _ in _cells(row) if not cell_deleted(c)]
        if not shown:
            continue
        return _first_in(shown[0])
    return None


def _first_in(container) -> Element | None:
    for node in container:
        if not isinstance(node.tag, str) or node.tag not in _BLOCKS:
            continue
        if node.tag == W_P:
            return node
        if node.tag == W_TBL:
            return _first_paragraph(node)
        content = node.find(_W + "sdtContent") if node.tag == _W + "sdt" else node
        return _first_in(content) if content is not None else None
    return None


def _last_paragraph(container: Element) -> Element | None:
    """A cell's last block when it is a paragraph (a content control's or ``w:customXml``'s
    content seen through), as :func:`.edges.last_mark_removed` reads it."""
    for node in reversed([n for n in container if isinstance(n.tag, str)]):
        if node.tag == W_P:
            return node
        if node.tag == W_TBL:
            return None
        if node.tag in (_W + "sdt", _W + "customXml"):
            content = node.find(_W + "sdtContent") if node.tag == _W + "sdt" else node
            if content is not None and any(isinstance(n.tag, str) and n.tag in _BLOCKS for n in content):
                return _last_paragraph(content)
    return None


def _shortened(text: str, limit: int = 60) -> str:
    return text if len(text) <= limit else text[:limit - 1] + "\u2026"


def _made_with(node: Element, record: Element) -> bool:
    return node.get(_W + "author") == record.get(_W + "author") and node.get(_W + "date") == record.get(_W + "date")


def _has_content(paragraph: Element) -> bool:
    """Whether a paragraph holds anything but properties and range markers."""
    for node in paragraph.iter(_W + "r", _W + "fldSimple", "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"):
        if node is not paragraph:
            return True
    return False


def _keep_markers(paragraph: Element) -> None:
    from ..edit.document import _rehome_markers

    _rehome_markers(paragraph)


def _prune(paragraph: Element) -> list[str]:
    """Remove what a review emptied: runs with nothing but properties, revision containers,
    hyperlinks and simple fields with no run.  Returns the hyperlinks' relationship ids."""
    released: list[str] = []
    changed = True
    while changed:
        changed = False
        for node in list(paragraph.iter(_W + "r", _W + "hyperlink", *_track.CONTAINERS, _W + "smartTag",
                                        _W + "fldSimple")):
            if node.getparent() is None:
                continue
            parent = node.getparent()
            if node.tag == _W + "r":
                continue
            if parent.tag == W_RPR:
                continue  # a mark's record, not a container
            if not any(isinstance(child.tag, str) and child.tag != W_RPR for child in node):
                rid = node.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if rid:
                    released.append(rid)
                remove(node)
                changed = True
    return released


#: A table's structure changes, tracked: while one is left, its rows may be off its grid by
#: Word's design (a merge's cell merged away, w:cellDel, until accepted).
_STRUCTURE = frozenset(_W + name for name in ("cellIns", "cellDel", "cellMerge", "tcPrChange", "tblGridChange",
                                               "trPrChange"))


def _fits(table: Element) -> bool:
    """Whether every row of ``table`` reaches exactly across its grid."""
    grid_node = table.find(_W + "tblGrid")
    count = len(grid_node.findall(_W + "gridCol")) if grid_node is not None else 0
    for row, cells in _grid_map(table):
        if _ancestor(row, W_TBL) is not table:
            continue
        properties = row.find(_W + "trPr")
        after = properties.find(_W + "gridAfter") if properties is not None else None
        extra = int(after.get(_W + "val")) if after is not None and (after.get(_W + "val") or "").isdigit() else 0
        reach = (cells[-1][1] + cells[-1][2] if cells else 0) + extra
        if reach != count:
            return False
    return True


def _normalise_grid(table: Element) -> None:
    """Rebuild a table's grid from its cells' widths, after a review set their spans back
    without the old grid (a split's rejected): the grid's lines where the cells' edges are,
    each cell spanning the columns between its own.  Left as it is where a cell states no
    width in twips, a row skips grid columns, or the rows' widths disagree."""
    grid_node = table.find(_W + "tblGrid")
    if grid_node is None:
        return
    rows = [(row, cells) for row, cells in _grid_map(table) if _ancestor(row, W_TBL) is table]
    extents = []
    for row, cells in rows:
        properties = row.find(_W + "trPr")
        if properties is not None and (properties.find(_W + "gridBefore") is not None
                                       or properties.find(_W + "gridAfter") is not None):
            return
        x, placed = 0, []
        for cell, _, _ in cells:
            width = cell.find(f"{_W}tcPr/{_W}tcW")
            if width is None or width.get(_W + "type") not in (None, "dxa") or not (width.get(_W + "w") or "").isdigit():
                return
            placed.append((cell, x, x + int(width.get(_W + "w"))))
            x += int(width.get(_W + "w"))
        extents.append(placed)
    totals = {placed[-1][2] for placed in extents if placed}
    if len(totals) != 1:
        return
    lines = sorted({0} | {edge for placed in extents for _, a, b in placed for edge in (a, b)})
    index = {line: k for k, line in enumerate(lines)}
    for column in grid_node.findall(_W + "gridCol"):
        remove(column)
    for a, b in zip(lines, lines[1:]):
        insert_in_order(grid_node, make("w:gridCol", **{"w:w": str(b - a)}))
    for placed in extents:
        for cell, a, b in placed:
            properties = cell.find(_W + "tcPr")
            node = properties.find(_W + "gridSpan")
            if node is not None:
                remove(node)
            if index[b] - index[a] > 1:
                insert_in_order(properties, make("w:gridSpan", **{"w:val": str(index[b] - index[a])}))


def _grid_map(table: Element) -> list[tuple[Element, list[tuple[Element, int, int]]]]:
    """Each row with its cells' first grid columns and spans, as the table is now."""
    out = []
    for row in table.iter(W_TR):
        if _ancestor(row, W_TBL) is not table:
            continue
        properties = row.find(_W + "trPr")
        before = properties.find(_W + "gridBefore") if properties is not None else None
        column = int(before.get(_W + "val")) if before is not None and (before.get(_W + "val") or "").isdigit() else 0
        cells = []
        for cell in row.iter(W_TC):
            if _ancestor(cell, W_TR) is not row:
                continue
            span_node = cell.find(f"{_W}tcPr/{_W}gridSpan")
            span = int(span_node.get(_W + "val")) if span_node is not None and (span_node.get(_W + "val") or "").isdigit() else 1
            cells.append((cell, column, span))
            column += span
        out.append((row, cells))
    return out


#: Run content a reader sees.
_SHOWN = frozenset(_W + name for name in ("t", "tab", "br", "cr", "drawing", "pict", "object", "sym",
                                          "noBreakHyphen", "ptab"))


def _content_kept(cell: Element) -> bool:
    """Whether a deleted cell (``w:cellDel``) keeps anything: content not deleted (outside
    ``w:del`` and ``w:moveFrom``) or a paragraph mark not deleted.  Word's Accept All gives
    such a cell to the cell before it, its columns and width -- an empty cell, or one whose
    text alone is deleted, too -- and removes only a cell whose content and marks are all
    deleted (measured, round 54 lane M: ``feedback-round54/probe-M``, ``m-empty``,
    ``c-textdel``; a tracked column deletion)."""
    for node in cell.iter(*_SHOWN):
        if node.tag == _W + "t" and not node.text:
            continue
        if any(a.tag in (_W + "del", _W + "moveFrom") for a in node.iterancestors()):
            continue
        return True
    for paragraph in cell.iter(W_P):
        if next((a for a in paragraph.iterancestors(W_TC)), None) is not cell:
            continue
        record = _track.mark_record(paragraph)
        if record is None or record.tag not in _track.DELETED:
            return True
    return False


def _cell_before(cell: Element) -> Element | None:
    """The cell before ``cell`` in its row, or ``None``."""
    node = cell.getprevious()
    while node is not None and not (isinstance(node.tag, str) and node.tag == W_TC):
        node = node.getprevious()
    return node


def _absorb(before: Element, cell: Element) -> None:
    """``before`` widened by ``cell``'s grid columns and width (both in twips)."""
    def span(of: Element) -> int:
        node = of.find(f"{_W}tcPr/{_W}gridSpan")
        value = node.get(_W + "val") if node is not None else None
        return max(1, int(value)) if value and value.isdigit() else 1

    properties = before.find(_W + "tcPr")
    if properties is None:
        properties = make("w:tcPr")
        before.insert(0, properties)
    total = span(before) + span(cell)
    node = properties.find(_W + "gridSpan")
    if node is None:
        node = make("w:gridSpan")
        insert_in_order(properties, node)
    node.set(_W + "val", str(total))
    mine, theirs = properties.find(_W + "tcW"), cell.find(f"{_W}tcPr/{_W}tcW")
    if mine is not None and theirs is not None and mine.get(_W + "type") in (None, "dxa") \
            and theirs.get(_W + "type") in (None, "dxa") and (mine.get(_W + "w") or "").isdigit() \
            and (theirs.get(_W + "w") or "").isdigit():
        mine.set(_W + "w", str(int(mine.get(_W + "w")) + int(theirs.get(_W + "w"))))


def _fix_grid(tables: list[Element], grid, removed: set[int], edged: set[int] = frozenset()) -> None:
    """Remove the grid columns that only removed cells covered (in every row of ``tables``
    that covered them at all), and the rows left without a cell -- a table, left without a
    row.  ``tables``: a table, or the parts a deleted mark at a row's end split it into
    (:meth:`_Review._out`), which keep one grid.  A row in ``edged`` (it gave up a cell to a
    deleted mark at a cell's end) that ends short of the grid skips the rest of it
    (``w:gridAfter``, ``w:wAfter``), as Word saves its Accept All (measured, round 54 lane R:
    ``feedback-round54/probe-R/word/rows``)."""
    members = {id(table) for table in tables}
    rows = [(row, cells) for row, cells in grid if id(_ancestor(row, W_TBL)) in members]
    columns: dict[int, list[bool]] = {}
    for _row, cells in rows:  # a row the review removed whole covers nothing
        for cell, start, span in cells:
            for column in range(start, start + span):
                columns.setdefault(column, []).append(id(cell) in removed)
    vacated = sorted((c for c, flags in columns.items() if flags and all(flags)), reverse=True)
    # Never fewer columns than a remaining row needs: a tracked deletion across a span
    # narrowed the grid already, and its deleted cells free nothing more.
    needed = 0
    for row, cells in rows:
        properties = row.find(_W + "trPr")
        total = 0
        for tag in ("gridBefore", "gridAfter"):
            node = properties.find(_W + tag) if properties is not None else None
            total += int(node.get(_W + "val")) if node is not None and (node.get(_W + "val") or "").isdigit() else 0
        total += sum(span for cell, _, span in cells if id(cell) not in removed)
        needed = max(needed, total)
    for table in tables:
        grid_node = table.find(_W + "tblGrid")
        if grid_node is not None:
            cols = grid_node.findall(_W + "gridCol")
            for column in vacated[:max(0, len(cols) - needed)]:
                if column < len(cols):
                    remove(cols[column])
        for row in list(table.findall(W_TR)):
            if row.find(W_TC) is None and row.find(_W + "sdt") is None:
                remove(row)
        if edged and grid_node is not None:
            widths = [c.get(_W + "w") or "" for c in grid_node.findall(_W + "gridCol")]
            for row, cells in _grid_map(table):
                if id(row) in edged:
                    _skip_after(row, cells, widths)
        if table.find(W_TR) is None and table.getparent() is not None and not any(
                isinstance(child.tag, str) and child.tag in (_W + "sdt", _W + "customXml") for child in table):
            remove(table)


def _skip_after(row: Element, cells: list[tuple[Element, int, int]], widths: list[str]) -> None:
    """A row ending short of its grid skips the rest (``w:gridAfter``, and ``w:wAfter`` in
    twips where the grid states them)."""
    properties = row.find(_W + "trPr")
    after = properties.find(_W + "gridAfter") if properties is not None else None
    skipped = int(after.get(_W + "val")) if after is not None and (after.get(_W + "val") or "").isdigit() else 0
    reach = (cells[-1][1] + cells[-1][2] if cells else 0) + skipped
    if reach >= len(widths):
        return
    skipped += len(widths) - reach
    if properties is None:
        properties = make("w:trPr")
        insert_in_order(row, properties)
    if after is None:
        after = make("w:gridAfter")
        insert_in_order(properties, after)
    after.set(_W + "val", str(skipped))
    tail = widths[len(widths) - skipped:]
    width = properties.find(_W + "wAfter")
    if all(w.isdigit() for w in tail):
        if width is None:
            width = make("w:wAfter")
            insert_in_order(properties, width)
        width.set(_W + "w", str(sum(int(w) for w in tail)))
        width.set(_W + "type", "dxa")
    elif width is not None:
        remove(width)


# -- the Document half -------------------------------------------------------------------------


class RevisionOps:
    """Listing, accepting and rejecting revisions, on :class:`docx_agent.Document`."""

    def revisions(self: "Document", *, author: str | None = None, kind: str | None = None,
                  within: "str | TextRange | None" = None, since: str | None = None,
                  until: str | None = None) -> list[Revision]:
        """Every revision record, in document order (the body, then headers, footers, notes
        and comments), filtered by author, kind (a kind or its family: ``insertion`` also
        finds ``paragraph-mark-insertion`` and ``row-insertion``), place (a range, block or
        story id) and date (``since`` and ``until``, ISO 8601, inclusive)."""
        return [Revision(self, identifier) for _, identifier, _ in
                self._select_revisions(None, author=author, kind=kind, within=within, since=since, until=until)]

    def changes(self: "Document", *, author: str | None = None, kind: str | None = None) -> list:
        """The revisions grouped into the edits a reviewer reads
        (:mod:`docx_agent.revisions.changes`): a deletion and an insertion side by side by
        one author at one date are one ``replacement``; a move's two ends one ``move``; a
        new paragraph's text, its mark and Word's ``paragraph-properties`` record one
        ``insertion``; and ``deletion``, ``formatting``, ``table``, ``section``.  Each
        :class:`~docx_agent.Change` has ``kind``, ``author``, ``date``, ``text`` (inserted),
        ``old_text`` (deleted), ``revisions`` (its ids) and ``accept()``/``reject()``::

            for change in doc.changes(author="Alice"):
                print(change.kind, repr(change.old_text), "->", repr(change.text))"""
        from .changes import changes

        return changes(self, author=author, kind=kind)

    def revision(self: "Document", identifier: str) -> Revision:
        """A revision by id (``rev:12``); ``KeyError`` if none.  ``doc.revision("rev:12").accept()``."""
        found = Revision(self, identifier)
        found._locate()
        return found

    def _select_revisions(self: "Document", target, *, author=None, kind=None, within=None, since=None,
                          until=None) -> list[tuple[str, str, Element]]:
        from .stamp import normalise_date

        records = self._revision_elements()
        if target not in (None, "all"):
            wanted = target if isinstance(target, (list, tuple, set)) else [target]
            selected = []
            for item in wanted:
                if isinstance(item, str) and item.startswith("rev:"):
                    found = [r for r in records if r[1] == item]
                    if not found:
                        raise KeyError(f"no revision {item!r}")
                    selected += found
                else:
                    selected += self._revisions_within(records, item)
            records = selected
        if within is not None:
            records = self._revisions_within(records, within)
        if author is not None:
            records = [r for r in records if r[2].get(_W + "author") == author]
        if kind is not None:
            records = [r for r in records if revision_kind(r[2]) == kind or revision_kind(r[2]).endswith("-" + kind)]
        if since is not None or until is not None:
            low = normalise_date(since) if since is not None else None
            high = normalise_date(until) if until is not None else None

            def in_window(element: Element) -> bool:
                date = element.get(_W + "date")
                if date is None:
                    return False
                value = normalise_date(date)
                return (low is None or value >= low) and (high is None or value <= high)

            records = [r for r in records if in_window(r[2])]
        return records

    def _revisions_within(self: "Document", records, where) -> list:
        """The records inside a range, a block (paragraph, table), or a story."""
        from ..edit.ranges import TextRange

        if isinstance(where, str) and where in [s.name for s in self.stories]:
            part = self._part_for_story(where)
            return [r for r in records if r[0] == part]
        if isinstance(where, str) and where.startswith("t@"):
            # A table's (or its row's, cell's) volatile id -- a table whose first row has no
            # paraId -- is a block, not a text range (round 54, lane A).
            try:
                part, entry = self._resolve(where)
            except (KeyError, ValueError):
                pass
            else:
                element = getattr(entry, "element", None)
                if element is not None:
                    return [r for r in records if r[0] == part and (
                        r[2] is element or any(a is element for a in r[2].iterancestors()))]
        if isinstance(where, str) and "@" in where:
            where = self.range(where, view="markup")
        if isinstance(where, TextRange):
            part, entries = where._entries()
            out = []
            for record in records:
                if record[0] != part:
                    continue
                paragraph = _ancestor(record[2], W_P)
                if paragraph is None:
                    continue
                positions = [k for k, e in enumerate(entries) if e.element is paragraph]
                if not positions:
                    continue
                k = positions[0]
                if _kind_of(record[2]) != "run":
                    out.append(record)
                    continue
                held = list(record[2].iter())
                inside = {id(node) for node in held}
                atoms = _text.atoms(paragraph, "markup")
                offsets = [i for i, atom in enumerate(atoms) if id(atom.node) in inside]
                start = where.start if k == 0 else 0
                end = where.end if k == len(entries) - 1 else len(atoms)
                if not offsets or (offsets[0] < end and offsets[-1] + 1 > start) or start == end:
                    out.append(record)
            return out
        part, entry = self._resolve(where)
        element = entry.element
        return [r for r in records if r[0] == part and (r[2] is element or any(a is element for a in r[2].iterancestors()))]

    def accept(self: "Document", target=None, *, author: str | None = None, kind: str | None = None,
               within=None, since: str | None = None, until: str | None = None) -> "EditResult":
        """Accept revisions: one id, several, those in a range or block, or all (``None``,
        ``"all"``), after the filters.  One undo step."""
        return self._review(True, target, author=author, kind=kind, within=within, since=since, until=until)

    def reject(self: "Document", target=None, *, author: str | None = None, kind: str | None = None,
               within=None, since: str | None = None, until: str | None = None) -> "EditResult":
        """Reject revisions, as :meth:`accept` selects them."""
        return self._review(False, target, author=author, kind=kind, within=within, since=since, until=until)

    def accept_all(self: "Document") -> "EditResult":
        """Accept every revision in the document (one undo step)."""
        return self.accept()

    def reject_all(self: "Document") -> "EditResult":
        """Reject every revision in the document (one undo step)."""
        return self.reject()

    def _review(self: "Document", accept: bool, target, **filters) -> "EditResult":
        from ..edit.document import EditResult
        from ..edit.numbering import referenced_nums

        selected = self._select_revisions(target, **filters)
        if not selected:
            return EditResult(None, changed=False, count=0)
        by_part: dict[str, list[Element]] = {}
        for part, _, element in selected:
            by_part.setdefault(part, []).append(element)
        removed_ids: list[str] = []
        with self._edit():
            renames = self._prepare_many([(part, []) for part in by_part])
            before = referenced_nums(self)
            authors = self._authors_in_use()
            described: set[int] = set()
            dropped: set[str] = set()
            notes: set[tuple[str, str]] = set()
            warnings: list[str] = []
            for part, elements in by_part.items():
                index = self._index(part)
                ids_before = {id(e.element): e.id for e in index.paragraphs}
                review = _Review(self, part, accept)
                group = review.expand(elements)
                review.run(group)
                for paragraph in review.removed_paragraphs:
                    if id(paragraph) in ids_before:
                        removed_ids.append(ids_before[id(paragraph)])
                for gone, survivor in review.joined:
                    if id(gone) in ids_before and id(survivor) in ids_before:
                        self._rename({ids_before[id(gone)]: ids_before[id(survivor)]})
                if review.released:
                    self._release(part, sorted(set(review.released)))
                described |= review.tables
                if review.moved_out:
                    self._invalidate()
                    entries = {id(e.element): e.id for e in self._index(part).paragraphs}
                    for text, paragraphs in review.moved_out:
                        said = [entries[id(p)] for p in paragraphs if id(p) in entries]
                        warnings.append(
                            f"a row's last cell whose last paragraph mark was deleted went; its text "
                            f"({_shortened(text)!r}) is now after the row in {', '.join(said) or 'paragraphs of its own'}, "
                            "the table split there: Word's Accept All takes that text out of view (word_changes "
                            "reject keeps it in the cell; undo brings the change back)")
                dropped |= review.dropped_comments
                notes |= review.dropped_notes
                self.package.mark_dirty(part)
                self._invalidate()
            if notes:
                self._drop_notes(notes)
            if dropped:
                by_w_id = {r["w_id"]: r["id"] for r in self._comment_records()}
                for w_id in sorted(dropped):
                    if w_id in by_w_id:
                        self._delete_comment_now(by_w_id[w_id])
            self._describe_resolved()
            self._reap_numbering(before)
            # A text box's fallback copy holds the records its choice had: it follows.
            self._sync_text_boxes(None)
            self._tidy_people(authors)
            # A data-bound control shows its custom XML node in Word: the node follows what
            # the review left the control showing.
            self._invalidate()
            self._sync_bound_controls()
        return EditResult(None, renamed=renames, removed=removed_ids, count=len(selected), warnings=warnings)

    def _drop_notes(self: "Document", notes: set[tuple[str, str]]) -> None:
        """Remove the notes a review removed the only reference of, and release what they
        held (a link's or a picture's relationship)."""
        from ..edit import inline as _inline

        referenced: set[tuple[str, str]] = set()
        for part in self._parts():
            for node in self.package.tree(part).iter(_W + "footnoteReference", _W + "endnoteReference"):
                referenced.add((node.tag[len(_W):-len("Reference")], node.get(_W + "id")))
        for part in self._parts():
            root = self.package.tree(part)
            if root.tag not in (_W + "footnotes", _W + "endnotes"):
                continue
            for note in list(root):
                key = (note.tag[len(_W):], note.get(_W + "id"))
                if key in notes and key not in referenced and note.get(_W + "type") in (None, "normal"):
                    released = _inline.relationship_ids(note)
                    remove(note)
                    if released:
                        self._release(part, sorted(set(released)))
                    self.package.mark_dirty(part)
        self._invalidate()

    def _describe_resolved(self: "Document") -> None:
        """Remove the comments describing structural table changes once their table has no
        cell revision left."""
        comments = getattr(self, "_describing_comments", None)
        if comments is None:
            return
        for comment_id, table in comments():
            if table is None or not any(node.tag in _CELL or node.tag == _W + "tcPrChange"
                                        for node in table.iter() if isinstance(node.tag, str)):
                self._delete_comment_now(comment_id)

    def _reap_numbering(self: "Document", before: set[int]) -> None:
        """Remove numbering instances a review left unused (the ones it found in use), and
        abstract definitions only they used: what the untracked edit would have removed."""
        from ..edit.numbering import Numbering, referenced_nums

        after = referenced_nums(self)
        gone = before - after
        if not gone:
            return
        numbering = Numbering(self)
        if numbering.part is None:
            return  # numbered paragraphs with no numbering part: nothing to remove
        nums = numbering.nums()
        abstracts = {numbering.abstract_of(n) for n in gone if n in nums}
        for num_id in gone:
            if num_id in nums:
                remove(nums[num_id])
        remaining = Numbering(self)
        used = {remaining.abstract_of(n) for n in remaining.nums()}
        for abstract_id in abstracts:
            if abstract_id is not None and abstract_id not in used:
                node = remaining.abstracts().get(abstract_id)
                if node is not None:
                    remove(node)
        self.package.mark_dirty(numbering.part)
