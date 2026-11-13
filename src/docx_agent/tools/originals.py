"""Round 57 (R122): text that was in the file as opened, taken out of the final view.

Agents rebuilt a table's column (a new one inserted, the old one deleted) where ``set_cells``
would have done; tracked, the old column stayed as deleted cells and its values left the
final view without a word.  Now the edits that remove whole blocks -- ``word_edit_table``'s
``delete_row``, ``delete_column`` and ``sync`` with ``remove``, ``word_delete`` of a
paragraph, a table or a span -- say so when what they remove held text **present when the
document was opened** (``original``): in the summary, and in ``data.original_removed``
(:func:`report`): how many cells and paragraphs, a short sample with their places, the
``undo`` that takes the call back and, tracked, the ``word_changes`` reject that keeps them.
Removing only what this session made (rows and columns inserted, text written into an empty
cell) says nothing.  ``check``'s ``accept_loses_original`` (:func:`accept_loses_original`)
lists the original text a pending tracked deletion made this session removes from the
final view, with the reject that keeps it.

**Original** (lane V's ``this_session`` method, round 56): a paragraph is the file's when it
held text in the file as opened (``package.opened()``, its final view) -- found by the id the
document gave it there (``p:<paraId>``, or ``p@<story>/<n>`` without one), followed through
the renames since (the first edit stamps a paraId on each paragraph without one, a repeat is
re-issued: each kept as an alias).  A block made this session has an id the file never had.
Its text is the final view's (``current``) less what this session inserted as tracked
changes (records whose author and date the file did not have).  A cell is original when one
of its paragraphs is and that text is not empty.  An original paragraph rewritten untracked
this session counts as the file's: its place held text.
"""

from __future__ import annotations

from typing import Any

from lxml import etree

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_P = _W + "p"
W_TC = _W + "tc"
W_TR = _W + "tr"
W_TBL = _W + "tbl"
_INS = (_W + "ins", _W + "moveTo")
#: Records that take content out of view (a move's source is shown where it went).
_DELETED = (_W + "del", _W + "cellDel")
#: Story roots: an element under one is in the document; a removed one is not.
_ROOTS = frozenset(_W + name for name in ("document", "hdr", "ftr", "footnotes", "endnotes", "comments"))

#: Places a result's sample names; the rest counted.
SAMPLE = 5
#: How many listed changes a reject call names at most.
CHANGES = 10


# -- the file as opened ----------------------------------------------------------------------


def opened_facts(document: Any, part: str) -> tuple[frozenset[str], frozenset[tuple[Any, Any]]]:
    """``(ids, stamps)`` of ``part`` as the file was opened: the ids (as the document names
    them: ``p:<paraId>``, or ``p@<story>/<n>`` with none) of its paragraphs whose final
    view held text, and every revision record's ``(author, date)`` -- read once per document
    and part (what was read never changes)."""
    from ..edit.annotations import REVISION_KINDS
    from ..edit.ids import PartIndex, live_elements

    kept = document.__dict__.setdefault("_opened_facts", {})
    found = kept.get(part)
    if found is None:
        root = document.package.opened().tree(part)
        if root is None:
            found = (frozenset(), frozenset())
        else:
            index = PartIndex(document._story_of(part), root)
            # Text in the final view: a w:t with more than white space (a deletion's text is
            # w:delText) -- cheap, once per document (round 57).
            ids = frozenset(e.id for e in index.paragraphs
                            if any((t.text or "").strip() for t in e.element.iter(_W + "t")))
            stamps = frozenset((node.get(_W + "author"), node.get(_W + "date"))
                               for node in live_elements(root, frozenset(REVISION_KINDS)))
            found = (ids, stamps)
        kept[part] = found
    return found


def _live(document: Any) -> Any:
    return document.__dict__.get("_live", document)       # a read copy: the document's


class Originals:
    """Which paragraphs of a document were the file's, and which records this session wrote."""

    def __init__(self, document: Any) -> None:
        self.document = document
        self.live = _live(document)
        self._current: dict[str, set[str]] | None = None
        self._stamps: dict[str, frozenset] = {}

    def session_record(self, part: str, node: Any) -> bool:
        """Whether revision record ``node`` was written since the document was opened (a
        record whose author and date the file had was there; an edit's records carry its
        time).  Not by the part's being changed: an edit may run on a copy of the document."""
        date = node.get(_W + "date")
        if date is None:
            return False
        stamps = self._stamps.get(part)
        if stamps is None:
            stamps = self._stamps[part] = opened_facts(self.live, part)[1]
        return (node.get(_W + "author"), date) not in stamps

    def current(self, part: str) -> set[str]:
        """The ids, now, of ``part``'s paragraphs that held text when the file was opened:
        each opened id followed through the document's renames (a paragraph stamped by the
        first edit, a repeat re-issued)."""
        if self._current is None:
            self._current = {}
        found = self._current.get(part)
        if found is None:
            aliases = {**self.live.aliases, **self.document.aliases}    # a copy's renames too
            found = set()
            for identifier in opened_facts(self.live, part)[0]:
                seen = {identifier}
                while identifier in aliases and aliases[identifier] not in seen:
                    identifier = aliases[identifier]
                    seen.add(identifier)
                found.add(identifier)
            self._current[part] = found
        return found

    def paragraph(self, part: str, paragraph: Any) -> bool:
        """Whether ``paragraph`` was in the file as opened, with text."""
        entry = self.document._index(part).entry_for(paragraph)
        return entry is not None and entry.id in self.current(part)

    def text(self, part: str, node: Any) -> str:
        """The final view's text of ``node`` (a paragraph or a cell), less what this session
        inserted as tracked changes, white space runs as one space."""
        from ..edit.text import paragraph_text

        paragraphs = [node] if node.tag == W_P else list(node.iter(W_P))
        out = []
        for paragraph in paragraphs:
            ours = [n for n in paragraph.iter(*_INS) if self.session_record(part, n)]
            if ours:
                paragraph = etree.fromstring(etree.tostring(paragraph))
                for n in [n for n in paragraph.iter(*_INS) if self.session_record(part, n)]:
                    n.getparent().remove(n)
            out.append(paragraph_text(paragraph, "current"))
        return " ".join(" ".join(out).split())

    def unit(self, part: str, node: Any) -> str:
        """``node``'s original text (a cell's or a paragraph's), ``""`` when it holds none."""
        paragraphs = [node] if node.tag == W_P else list(node.iter(W_P))
        if not any(self.paragraph(part, p) for p in paragraphs):
            return ""
        return self.text(part, node)


def _attached(node: Any) -> bool:
    top = node
    while top.getparent() is not None:
        top = top.getparent()
    return top.tag in _ROOTS


def _shown(node: Any) -> str:
    from ..edit.text import paragraph_text

    paragraphs = [node] if node.tag == W_P else list(node.iter(W_P))
    return "".join(paragraph_text(p, "current") for p in paragraphs).strip()


def _short(text: str, limit: int = 60) -> str:
    return text if len(text) <= limit else text[:limit - 1] + "…"


# -- an edit's watch -----------------------------------------------------------------------


class Watch:
    """The original text an edit is about to remove (:meth:`row`, :meth:`column`,
    :meth:`block`, before the edit), and what of it left the final view (:meth:`report`,
    after)."""

    def __init__(self, document: Any) -> None:
        self.document = document
        self.originals = Originals(document)
        self.units: list[tuple[str, str, Any, Any, str]] = []      # (kind, place, part, node, text)
        self._seen: set[int] = set()

    def _add(self, kind: str, place: str, part: str, node: Any) -> None:
        if id(node) in self._seen:
            return
        text = self.originals.unit(part, node)
        if text:
            self._seen.add(id(node))
            self.units.append((kind, place, part, node, text))

    def row(self, table: str, index: int) -> None:
        from ..edit.tables import _cells

        try:
            part, entry = self.document._table_entry(table)
        except (KeyError, ValueError):
            return
        if 0 <= index < len(entry.rows):
            for cell, start, _ in _cells(entry.rows[index][0]):
                self._add("cell", f"{entry.id}/c{index},{start}", part, cell)

    def column(self, table: str, grid: int) -> None:
        from ..edit.tables import _cells

        try:
            part, entry = self.document._table_entry(table)
        except (KeyError, ValueError):
            return
        for r, (row, _) in enumerate(entry.rows):
            for cell, start, span in _cells(row):
                if start == grid and span == 1:
                    self._add("cell", f"{entry.id}/c{r},{start}", part, cell)

    def block(self, identifier: str) -> None:
        """A paragraph, a table (its cells) or a block-level control (its paragraphs)."""
        from ..edit.tables import _cells

        try:
            part, entry = self.document._resolve_block(identifier)
        except (KeyError, ValueError, AttributeError):
            return
        element = getattr(entry, "element", None)
        if element is None:
            return
        if element.tag == W_P:
            self._add("paragraph", entry.id, part, element)
        elif element.tag == W_TBL:
            for r, (row, _) in enumerate(entry.rows):
                for cell, start, _ in _cells(row):
                    self._add("cell", f"{entry.id}/c{r},{start}", part, cell)
        else:
            index = self.document._index(part)
            for paragraph in element.iter(W_P):
                found = index.entry_for(paragraph)
                self._add("paragraph", found.id if found is not None else identifier, part, paragraph)

    def lost(self) -> list[tuple[str, str, Any, Any, str]]:
        """The units watched that left the final view: removed, or their text gone from it."""
        return [u for u in self.units if not _attached(u[3]) or not _shown(u[3])]

    def report(self, call: Any) -> dict[str, Any] | None:
        """``data.original_removed`` -- ``None`` when no original text left the final view."""
        lost = self.lost()
        if not lost:
            return None
        cells = sum(1 for u in lost if u[0] == "cell")
        out: dict[str, Any] = {"cells": cells, "paragraphs": len(lost) - cells,
                               "sample": [{"at": place, "text": _short(text)} for _, place, _, _, text in lost[:SAMPLE]]}
        if len(lost) > SAMPLE:
            out["sample_left_out"] = len(lost) - SAMPLE
        doc = call.entry.doc_id
        out["undo"] = {"tool": "undo", "arguments": {"doc": doc}}
        if call.in_batch:
            out["undo_note"] = "undo takes back the whole batch: it is one step"
        tracked = [u for u in lost if _attached(u[3])]
        if tracked:
            changes = _changes_of(self.document, self.originals, tracked)
            if changes:
                out["reject"] = {"tool": "word_changes",
                                 "arguments": {"doc": doc, "action": "reject", "ids": changes[:CHANGES]}}
        return out


def _records_in(originals: Originals, units: list) -> list[Any]:
    """The deletion records this session wrote that take ``units`` out of view: in them,
    and their rows' (a row's deletion is recorded in its ``w:trPr``)."""
    out: list[Any] = []
    seen: set[int] = set()
    for _, _, part, node, _ in units:
        holders = [node]
        row = node.getparent()
        while row is not None and row.tag != W_TR:
            row = row.getparent()
        if row is not None and row.find(f"{_W}trPr/{_W}del") is not None:
            holders.append(row.find(f"{_W}trPr"))
        for holder in holders:
            for record in holder.iter(*_DELETED):
                if id(record) not in seen and originals.session_record(part, record):
                    seen.add(id(record))
                    out.append(record)
    return out


def _changes_of(document: Any, originals: Originals, units: list) -> list[str]:
    """The listed changes (``word_changes`` list's ids) holding the records that take
    ``units`` out of view, in document order."""
    records = _records_in(originals, units)
    if not records:
        return []
    wanted = {id(r) for r in records}
    ids = {identifier for _, identifier, element in document._revision_elements() if id(element) in wanted}
    return [c.id for c in document.changes() if ids & set(c.revisions)]


def said(found: dict[str, Any] | None) -> str:
    """The summary's words for :meth:`Watch.report`: ``""`` when none."""
    if not found:
        return ""
    parts = []
    if found["cells"]:
        parts.append(f"{found['cells']} cell{'' if found['cells'] == 1 else 's'}")
    if found["paragraphs"]:
        parts.append(f"{found['paragraphs']} paragraph{'' if found['paragraphs'] == 1 else 's'}")
    sample = ", ".join(f"{s['at']} {_short(s['text'], 30)!r}" for s in found["sample"][:3])
    more = len(found["sample"]) - 3 + found.get("sample_left_out", 0)
    text = (f"; {' and '.join(parts)} holding text the file had when opened left the final view "
            f"({sample}{f', {more} more' if more > 0 else ''}): undo takes the call back")
    if "reject" in found:
        text += f", or word_changes reject ids {found['reject']['arguments']['ids']} keeps them"
    return text


def add(result: Any, found: dict[str, Any] | None) -> Any:
    """``found`` (:meth:`Watch.report`) into ``result``'s summary and ``data.original_removed``."""
    if not found:
        return result
    result.summary += said(found)
    if result.data is None:
        result.data = {}
    if isinstance(result.data, dict):
        result.data["original_removed"] = found
    return result


# -- check's fact --------------------------------------------------------------------------


def accept_loses_original(document: Any) -> list[dict[str, Any]]:
    """``check``'s ``accept_loses_original``: text the file had when opened that a pending
    tracked deletion made this session takes out of the final view -- one entry per listed
    change and table (or the body's paragraphs): ``change``, ``author``, ``table`` (or
    ``paras``), ``cells`` and ``paragraphs`` counted, a ``sample`` of their places and text,
    ``fix`` the ``word_changes`` reject that keeps them.  ``[]`` at once where the body holds
    no deletion record of this session's."""
    from ..edit.tables import _cells

    part = document.package.document_part()
    originals = Originals(document)
    root = document.package.tree(part)
    records = [n for n in root.iter(*_DELETED) if originals.session_record(part, n)]
    if not records:
        return []
    index = document._index(part)
    tables: dict[int, Any] = {id(t.element): t for t in index.tables}
    units: dict[int, tuple[str, Any]] = {}            # id(node) -> (kind, node)
    for record in records:
        node = record.getparent()
        while node is not None and node.tag not in (W_TC, W_P, W_TR):
            node = node.getparent()
        if node is None:
            continue
        if node.tag == W_P:
            cell = node.getparent()
            while cell is not None and cell.tag not in (W_TC, _W + "body"):
                cell = cell.getparent()
            if cell is not None and cell.tag == W_TC:
                node = cell
        if node.tag == W_TR:
            for cell, _, _ in _cells(node):
                units.setdefault(id(cell), ("cell", cell))
        else:
            units.setdefault(id(node), ("cell" if node.tag == W_TC else "paragraph", node))
    groups: dict[tuple, dict[str, Any]] = {}
    places: dict[int, dict[int, str]] = {}
    revisions: dict[int, str] | None = None
    changes: list[Any] | None = None
    for kind, node in units.values():
        if kind == "cell" and node.find(f"{_W}tcPr/{_W}cellMerge") is not None:
            continue                    # a pending merge's: its text is shown in the merged cell
        if _shown(node) or not any(originals.paragraph(part, p) for p in ([node] if kind == "paragraph"
                                                                           else node.iter(W_P))):
            continue
        text = _deleted_text(originals, part, node)
        if not text:
            continue
        if revisions is None:
            revisions = {id(e): i for _, i, e in document._revision_elements()}
            changes = document.changes()
        mine = {revisions[id(r)] for r in _records_in(originals, [(kind, "", part, node, "")]) if id(r) in revisions}
        change = next((c for c in changes or () if mine & set(c.revisions)), None)
        place, table = _place(index, tables, places, kind, node)
        key = (change.id if change is not None else None, table)
        entry = groups.get(key)
        if entry is None:
            entry = groups[key] = {"change": key[0], "author": change.author if change is not None else None,
                                   **({"table": table} if table else {"paras": []}),
                                   "cells": 0, "paragraphs": 0, "sample": []}
            entry["fix"] = {"tool": "word_changes", "calls": [{"action": "reject", "ids": [key[0]]}]} \
                if key[0] else {"tool": "word_changes", "calls": [{"action": "reject", "within": table or place}]}
        entry["cells" if kind == "cell" else "paragraphs"] += 1
        if not table and len(entry["paras"]) < 20:
            entry["paras"].append(place)
        if len(entry["sample"]) < 3:
            entry["sample"].append({"at": place, "text": _short(text)})
    return list(groups.values())


def _deleted_text(originals: Originals, part: str, node: Any) -> str:
    """The text this session's deletion records hold in ``node``."""
    out = []
    for record in node.iter(_W + "del"):
        if originals.session_record(part, record):
            out.append("".join(t.text or "" for t in record.iter(_W + "delText", _W + "t")))
    return " ".join(" ".join(out).split())


def _place(index: Any, tables: dict[int, Any], places: dict[int, dict[int, str]], kind: str,
           node: Any) -> tuple[str, str | None]:
    """A unit's address (a cell on its table's document grid; a paragraph's id) and its
    table; ``places`` keeps each table's cells' addresses, made once."""
    from ..edit.tables import _cells

    if kind == "paragraph":
        found = index.entry_for(node)
        return (found.id if found is not None else "?"), None
    table = node.getparent()
    while table is not None and table.tag != W_TBL:
        table = table.getparent()
    entry = tables.get(id(table)) if table is not None else None
    if entry is None:
        return "?", None
    cells = places.get(id(table))
    if cells is None:
        cells = places[id(table)] = {id(cell): f"{entry.id}/c{r},{start}" for r, (row, _) in enumerate(entry.rows)
                                     for cell, start, _ in _cells(row)}
    return cells.get(id(node), entry.id), entry.id
