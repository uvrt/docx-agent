"""Reflow feedback: where docx2svg laid each paragraph out, in this API's ids.

docx2svg lays out the bytes docx-agent serialises, so every ``Line.path`` it reports names
exactly one element of the live tree.  :class:`DocumentLayout` resolves those paths with
docx2svg's own :func:`docx2svg.paths.resolve_path` (the rules ``data-docx-path`` is counted
by, exported for this) and keys every line by the id the paragraph has in the document.

The document is laid out **once** per state: :func:`convert` lays it out with
:func:`docx2svg.lay_out_docx` (:func:`lay_out_bytes`) -- every page -- and both ``layout()``
and ``render_svg()`` read that one layout, the SVG of a page drawn from it when asked for.
Each state is laid out **from the document's last layout** (round 33: docx2svg's
``previous``, only from the page an edit reaches until the pages line up again, page for
page a layout of the bytes alone), and callers asking for a state while it is laid out
**wait for that one layout** rather than lay it out again.

**Locks and waits** (round 34, R17: a wait for a layout under the document's lock hung every
call on it).  One rule: *never wait for a layout in progress while holding the document's
lock*.  In order, each taken only with those before it held, never the other way:

1. the document's lock (ooxml-edit's ``DocumentEntry.lock``, taken by ``Session.lock`` in
   document order; ``Call.unlocked()`` lets the call's go and takes them back in that order);
2. the lock over the document's layouts (:func:`_keeper`: the cache, the layouts in progress,
   the one kept for the next version; held only for a lookup or to map a layout to ids --
   never while laying out or waiting) -- one document's at a time (:func:`adopt` takes the
   copy's, lets it go, then the document's), and :data:`_KEEPERS_LOCK` only to make one;
3. :data:`_RECENT_LOCK` (taken in :func:`lay_out_bytes`, with neither of the above held);
4. a layout's drawing lock (``_Conversion.drawing``: :func:`render_svg` and
   :mod:`docx_agent.liveview` drawing a page, or a page's digest, from a layout made
   already), taken with nothing held but -- in :func:`render_svg` -- the document's lock, and
   nothing taken or waited for under it.

A layout in progress (:class:`_Flight`) is waited for only by a caller that let the
document's lock go to wait (``unlocked`` given and not yielding ``False``: a lone check, the
work ahead of a change); a caller holding it -- a batch's ops, a ``read_batch``'s reads, ``word_inspect``'s
placements, ``Document.layout()`` under the lock, ``update_fields`` -- or a caller with no
``unlocked`` at all (a library's own thread) lays out on its own, its own snapshot.  The
flight's owner ends its flight -- result or error -- *before* it takes the document's lock
back.  So a flight ends whatever any lock's holder does, and no wait closes a cycle.

What it answers (ROADMAP.md, "Reflow feedback"):

* ``where(id)`` -- one :class:`Placement` per page and text column the paragraph (or table)
  is drawn in, in points from the page's top; ``page_of(id)`` -- the page, or ``(first, last)``;
* ``blocks_on(page)`` and ``pages_of(id)`` -- the top-level blocks a page draws, and every
  page a paragraph, table or row is drawn on (round 19), from docx2svg's own page <-> block
  map (``Page.blocks``: element paths, mapped onto this API's ids);
* ``compare(before)`` -- which pages differ from an earlier layout, by a signature of what
  each page draws (glyphs, positions, rules, pictures; not the SVG text), and which blocks
  moved to another page;
* **stops, carried through**: docx2svg carries on past what it cannot lay out (round 16: an
  approximate rule, or a marked placeholder of its height, each in its coverage); where it is
  made to stop (``docx2svg.paginate.CARRY_ON`` off), a block past the stop is
  :class:`Unknown`, never guessed, and pages past either layout's stop are reported as
  unknown by ``compare``.

Ids are compared through the document's aliases, so a layout taken before an edit that
stamped or re-issued paragraphs still matches the layout after it.
"""

from __future__ import annotations

import dataclasses
import hashlib
import pickle
import re
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from fractions import Fraction
from typing import TYPE_CHECKING, Any

import docx2svg
from docx2svg.paths import path_part, resolve_path

from .edit.effective import PathFinder
from .edit.ids import ParagraphEntry, TableEntry

if TYPE_CHECKING:  # pragma: no cover
    from .edit.document import Document

#: docx2svg's device pixels (Word's 1/300-inch export grid) per point.
PX_PER_PT = Fraction(300, 72)

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
#: A key naming one of the body's children (:meth:`DocumentLayout._top_level`).
_TOP_KEY = re.compile(r"^w:body/w:([A-Za-z]+)\[(\d+)\]$")


@dataclass(frozen=True)
class Placement:
    """Where a block is drawn on one page, in one text column: 1-based ``page``, ``top``
    and ``bottom`` in points from the page's top edge, how many of its ``lines`` are
    there, the story it is drawn in (``body``, ``header``, ``footer``, ``footnotes``,
    ``endnotes``), and the 0-based text ``column`` (docx2svg's :attr:`Line.column`; 0 in a
    section of one column, and for a header's, a footer's or a text box's line)."""

    page: int
    top: float
    bottom: float
    lines: int
    story: str = "body"
    column: int = 0


@dataclass(frozen=True)
class Extent:
    """How much of one page some blocks' body content takes (round 35, R26): the 1-based
    ``page``; ``top`` and ``bottom`` in points from the page's top edge -- the first of
    their lines' tops and the last of their lines' bottoms there, as :meth:`DocumentLayout.where`
    measures a block (a line's top to its top plus its pitch; a table by its cells' lines),
    across the page's text columns (``columns``: how many they are drawn in); how many
    ``lines``; ``body`` -- the page's height less the top and bottom margins its section
    sets, in points (``w:pgSz``, ``w:pgMar``; a negative margin by its magnitude) -- and
    ``area``, ``(top, bottom)``: where docx2svg lays the body out on that page, which is
    the margins unless the page's header or footer reaches past its margin and pushes the
    body in (Word does: docx2svg's ``paginate._geometry``, measured), so ``area[1] -
    area[0]`` is the room the body really has there."""

    page: int
    top: float
    bottom: float
    lines: int
    columns: int
    body: float
    area: tuple[float, float]

    @property
    def height(self) -> float:
        """``bottom - top``."""
        return round(self.bottom - self.top, 2)


@dataclass(frozen=True)
class Unknown:
    """Where a block is cannot be said: ``reason`` is the layout's stop (``table``,
    ``drawing``...) for a block past it, with the page it stopped on, or ``not-drawn`` for a
    block docx2svg draws nothing of (a hidden paragraph, the continuation of a vertically
    merged cell, a note it does not place)."""

    reason: str
    after_page: int | None = None

    def __bool__(self) -> bool:
        return False


def coverage_facts(coverage, *, at: str | None = None, ids: dict | None = None) -> dict:
    """What a caller needs to tell "laid out and nothing wrong" from "could not lay it all
    out": docx2svg's coverage (:class:`docx2svg.coverage.Coverage`, or its ``as_dict()``),
    compact.  ``complete`` is true only when every block was laid out, every header,
    footer and text box drawn and every face found; ``substituted_fonts`` names the faces
    laid out with an open substitute (``(approximate)`` where it is not metric compatible);
    ``stop`` is where the layout stopped, ``at`` the block's id when known.  docx2svg no
    longer stops by default (round 16): what it cannot lay out it lays out by an
    approximate rule (``approximations``) or draws as a marked placeholder of its height,
    and goes on; ``placeholders`` lists those (page, reason, ``at``: the block's id where
    ``ids`` maps its element path to one, else the path), and makes the layout ``partial``.

    ``status`` tells the three apart: ``complete`` (laid out, nothing approximated),
    ``approximate`` (complete, but ``approximations`` lists places docx2svg laid out by a
    rule it has not measured -- a floating drawing in a table cell placed as no probe
    measured, say -- instead of stopping: what they touch may be off) and ``partial``
    (not complete)."""
    data = coverage if isinstance(coverage, dict) else coverage.as_dict()
    approximations = data.get("approximations") or []
    status = data.get("status") or ("partial" if not data["complete"]
                                    else "approximate" if approximations else "complete")
    facts = {"complete": data["complete"], "status": status, "pages": data["pages"],
             "blocks_laid_out": [data["blocks_laid_out"], data["blocks"]]}
    if data["estimate_source"] == "app.xml":
        facts["pages_estimated"] = data["estimated_pages"]
    stop = data.get("stop")
    if stop:
        facts["stop"] = {"page": stop["page"], "reason": stop["reason"], "at": at or stop.get("path"),
                         "message": _short(stop["message"])}
    if data["story_stops"]:
        facts["story_stops"] = [{"page": s["page"], "reason": s["reason"], "message": _short(s["message"])}
                                for s in data["story_stops"][:5]]
    if data["substituted_fonts"]:
        facts["substituted_fonts"] = [f"{s['family']} -> {s['substitute']}"
                                      + ("" if s["metric_compatible"] else " (approximate)")
                                      for s in data["substituted_fonts"]]
    if data["missing_fonts"]:
        facts["missing_fonts"] = list(data["missing_fonts"])
    if approximations:
        facts["approximations"] = [{"page": a["page"], "reason": a["reason"], "path": a.get("path"),
                                    "message": _short(a["message"])} for a in approximations[:5]]
        if len(approximations) > 5:
            facts["approximations_total"] = len(approximations)
    placeholders = data.get("placeholders") or []
    if placeholders:
        facts["placeholders"] = [{"page": p["page"], "reason": p["reason"],
                                  "at": (ids or {}).get(p.get("path")) or p.get("path"),
                                  "message": _short(p["message"])} for p in placeholders[:5]]
        if len(placeholders) > 5:
            facts["placeholders_total"] = len(placeholders)
    return facts


def relaid_facts(document: "Document", cause: dict | None) -> dict | None:
    """Why docx2svg laid a version out whole from the layout of an earlier one, in this
    API's ids (round 20): docx2svg's :attr:`~docx2svg.layout.Layout.relaid_cause` --
    ``reason`` (``styles``, ``sections``, ``fields-open``...; docx2svg's
    ``relayout.FALLBACKS``), ``field`` (the instruction of the field an edited block opens and
    another closes) and ``part`` (the part that changed) as they are -- with ``at`` the id of
    the block it names (``p:1A2B3C4D``, ``t:...``): the block in ``document`` at docx2svg's
    element path, else (a block the edit deleted) ``p:`` and its paragraph id, else the path.
    ``None`` for ``None``: the layout was made from the earlier one only where the edit
    reached, or without one."""
    if not cause:
        return None
    out = {"reason": cause.get("reason")}
    path = cause.get("at")
    if path:
        part = document.package.document_part()
        found = None
        try:
            found = _block_id(document, part, path)
        except (KeyError, ValueError):
            found = None
        if found is None and cause.get("para_id"):
            found = "p:" + str(cause["para_id"]).upper()
        out["at"] = found or path
    for key in ("field", "part"):
        if cause.get(key):
            out[key] = cause[key]
    return out


def _short(text: str, limit: int = 160) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= limit else text[:limit - 1] + "\u2026"


@dataclass(frozen=True)
class Stop:
    """Where docx2svg stopped: the 1-based page, why, and the block it stopped at."""

    page: int
    reason: str
    at: str | None


@dataclass
class Reflow:
    """What changed between two layouts."""

    #: 1-based pages whose drawing differs, among the pages both layouts know.
    changed: list[int]
    first_changed_page: int | None
    #: (before, after) page counts.
    page_count: tuple[int, int]
    #: Blocks whose first page changed: id -> (before, after).
    moved: dict[str, tuple[int, int]] = field(default_factory=dict)
    #: Pages past either layout's stop: not compared, since neither knows them.
    unknown: list[int] = field(default_factory=list)
    #: Why each changed page changed: ``content`` (what is drawn differs), ``flow`` (which
    #: blocks are on it differs), ``stories`` (only a header, footer or note differs, as
    #: when NUMPAGES changes), ``added`` or ``removed``.
    why: dict[int, str] = field(default_factory=dict)


class DocumentLayout:
    """docx2svg's layout of one state of a document, keyed by this API's ids."""

    def __init__(self, document: "Document", layout, warnings: list) -> None:
        self._document = document
        self.pages = layout.pages
        #: docx2svg's :class:`docx2svg.Warning`\\ s: everything not laid out faithfully.
        self.warnings = list(warnings)
        #: Line paths docx2svg drew that no paragraph of the document answers to.
        self.unmapped: list[str] = []
        #: Per page: [(id or None, line)].
        self._lines: list[list[tuple[str | None, object]]] = []
        resolved: dict[tuple[str, str], str | None] = {}
        # Every line's path resolved in this one state: each table's rows listed once.
        paths = PathFinder()
        for page in self.pages:
            lines = []
            for line in page.text_lines():
                key = (_part_of(document, line.path, line.part), line.path)
                if key not in resolved:
                    resolved[key] = _paragraph_id(document, *key, paths)
                    if resolved[key] is None:
                        self.unmapped.append(line.path)
                lines.append((resolved[key], line))
            self._lines.append(lines)
        self._paths = resolved
        self.stopped: Stop | None = None
        for page in self.pages:
            if page.stop is not None:
                at = _block_id(document, document.package.document_part(), page.stop.path)
                self.stopped = Stop(page.number + 1, page.stop.reason, at or page.stop.path)
                break
        #: docx2svg's :class:`docx2svg.coverage.Coverage`: how much of the document this
        #: layout covers.
        self.coverage = layout.coverage
        #: Element path -> this API's id, for each body block docx2svg drew as a placeholder.
        self._placeholder_ids: dict = {}
        for entry in self.coverage.placeholders:
            if entry.path and entry.path.startswith("w:body"):
                found = _block_id(document, document.package.document_part(), entry.path)
                if found:
                    self._placeholder_ids[entry.path] = found

    def coverage_facts(self) -> dict:
        """:func:`coverage_facts` of this layout, its stop and placeholders named by this
        API's ids."""
        return coverage_facts(self.coverage, at=self.stopped.at if self.stopped else None,
                              ids=self._placeholder_ids)

    # -- what is known -----------------------------------------------------------------------

    @property
    def page_count(self) -> int:
        """Pages docx2svg laid out (up to and including a stop's page)."""
        return len(self.pages)

    @property
    def pages_known(self) -> int | None:
        """The page count when the whole document was laid out; ``None`` when it stopped,
        since the pages after a stop are not known."""
        return None if self.stopped else len(self.pages)

    def id_for_path(self, part: str, path: str) -> str | None:
        key = (part, path)
        if key not in self._paths:
            self._paths[key] = _paragraph_id(self._document, part, path)
        return self._paths[key]

    # -- where -------------------------------------------------------------------------------

    def where(self, identifier: str) -> "list[Placement] | Unknown":
        """One placement per page and text column the paragraph or table is drawn in, in
        page order and, on a page, in the order its columns are first drawn."""
        wanted = self._wanted(identifier)
        out: list[Placement] = []
        for number, lines in enumerate(self._lines):
            columns: dict[int, list] = {}
            for key, line in lines:
                if key is not None and self._canon(key) in wanted:
                    columns.setdefault(line.column or 0, []).append(line)
            for column, mine in columns.items():
                top = min(line.top for line in mine)
                bottom = max(line.top + line.pitch for line in mine)
                story = mine[0].story or _story_kind(mine[0].path)
                out.append(Placement(number + 1, round(float(top / PX_PER_PT), 2),
                                     round(float(bottom / PX_PER_PT), 2), len(mine), story, column))
        if out:
            return out
        return self._unknown(identifier)

    def extent(self, identifiers: "str | Any", pages: "Any" = None) -> list[Extent]:
        """How much of each page the body content of some blocks takes (round 35, R26):
        one :class:`Extent` per page any of them is drawn on, in page order --
        ``identifiers`` one id or several (paragraphs, tables: a table's cells' paragraphs;
        a span's blocks given one by one), ``pages`` the 1-based pages to look on (default
        every page).  Only the body's lines count (not a header's, a footer's or a note's).
        ``[]`` for blocks nothing of which is drawn."""
        if isinstance(identifiers, str):
            identifiers = [identifiers]
        wanted: set[str] = set()
        for identifier in identifiers:
            wanted |= self._wanted(identifier)
        numbers = range(1, len(self.pages) + 1) if pages is None else sorted(
            {n for n in pages if 1 <= n <= len(self.pages)})
        out = []
        for number in numbers:
            mine = [line for key, line in self._lines[number - 1]
                    if key is not None and _in_body(line) and self._canon(key) in wanted]
            if not mine:
                continue
            top = min(line.top for line in mine)
            bottom = max(line.top + line.pitch for line in mine)
            body, area = self._body_of(number)
            out.append(Extent(number, round(float(top / PX_PER_PT), 2), round(float(bottom / PX_PER_PT), 2),
                              len(mine), len({line.column or 0 for line in mine}), body, area))
        return out

    def _body_of(self, number: int) -> tuple[float, tuple[float, float]]:
        """``(body, area)`` of a 1-based page (:class:`Extent`)."""
        page = self.pages[number - 1]
        area = (round(float(Fraction(page.column[2]) / PX_PER_PT), 2),
                round(float(Fraction(page.column[3]) / PX_PER_PT), 2))
        body = None
        info = getattr(page, "info", None)
        if info is not None:
            sections = self.__dict__.get("_section_heights")
            if sections is None:
                sections = self.__dict__["_section_heights"] = _section_bodies(self._document)
            if 0 <= info.section < len(sections):
                body = sections[info.section]
        if body is None:                        # a section with no size or margins stated
            body = round(area[1] - area[0], 2)
        return body, area

    def page_of(self, identifier: str) -> "int | tuple[int, int] | Unknown":
        """The 1-based page a block is on, or ``(first, last)`` when it spans pages."""
        placements = self.where(identifier)
        if isinstance(placements, Unknown):
            return placements
        first, last = placements[0].page, placements[-1].page
        return first if first == last else (first, last)

    def ids_on_page(self, page: int) -> list[str]:
        """The ids of the paragraphs drawn on a 1-based page, in drawing order."""
        seen: list[str] = []
        for key, _ in self._lines[page - 1]:
            if key is not None and self._canon(key) not in seen:
                seen.append(self._canon(key))
        return seen

    def blocks_on(self, page: int) -> list[str]:
        """The ids of the body's top-level blocks -- paragraphs and tables, the ids the Word
        tools take -- that the 1-based ``page`` draws something of, in document order: a
        paragraph or table split across pages is on each (docx2svg's ``Page.blocks``,
        round 19).  ``[]`` for a page the layout does not have."""
        if not 1 <= page <= len(self.pages):
            return []
        from docx2svg.paths import block_keys

        order = self._block_order()
        found: dict[str, int] = {}
        for key in self.pages[page - 1].blocks:
            if block_keys(key)[0] != key:
                continue
            identifier, element = self._key_id(key)
            if identifier is not None and identifier not in found:
                found[identifier] = order.get(element, -1)
        return sorted(found, key=found.__getitem__)

    def pages_of(self, identifier: str) -> list[int]:
        """Every 1-based page a paragraph, table, table row or cell (its id, or an alias of
        it) is drawn on, in order -- each page of one split across pages, each page a header
        row is drawn again on.
        ``[]`` for a block nothing of which is drawn (hidden, past a stop, not the body's)
        or an id the document does not have."""
        index = self._pages_index()
        return list(index.get(self._canon(identifier), []))

    def _pages_index(self) -> dict[str, list[int]]:
        index = getattr(self, "_page_ids_index", None)
        if index is None:
            index = {}
            for page in self.pages:
                for key in page.blocks:
                    identifier, _ = self._key_id(key)
                    if identifier is None:
                        continue
                    pages = index.setdefault(self._canon(identifier), [])
                    if not pages or pages[-1] != page.number + 1:
                        pages.append(page.number + 1)
            self._page_ids_index = index
        return index

    def _key_id(self, key: str) -> "tuple[str | None, Any]":
        """The id of the body block docx2svg's key names (a paragraph's, a table's, a
        row's or a cell's), and its element, cached."""
        keys = self.__dict__.setdefault("_key_ids", {})
        if key not in keys:
            document = self._document
            part = document.package.document_part()
            element = self._top_level(key)
            if element is False:
                element = resolve_path(document.package.tree(part), key)
            identifier = None
            if element is not None and element.tag in (_W + "p", _W + "tbl"):
                entry = document._index(part).entry_for(element)
                identifier = entry.id if entry is not None else None
            elif element is not None and element.tag in (_W + "tr", _W + "tc"):
                identifier = self._row_ids().get(element)
            keys[key] = (identifier, element)
        return keys[key]

    def _top_level(self, key: str) -> Any:
        """The body's child a key ``w:body/w:p[3]`` names, from the body's children listed
        once (a live view maps every page's blocks: each key resolved alone listed the body
        again); ``False`` for another kind of key, or where the body holds a child of the
        same local name in another namespace (docx2svg counts those too)."""
        found = _TOP_KEY.match(key)
        if found is None:
            return False
        table = self.__dict__.get("_top_table")
        if table is None:
            document = self._document
            root = document.package.tree(document.package.document_part())
            body = root.find(_W + "body")
            table = {}
            counts: dict[str, int] = {}
            locals_: dict[str, set] = {}
            for child in (body if body is not None else ()):
                if not isinstance(child.tag, str):
                    continue
                local = child.tag.rpartition("}")[2]
                locals_.setdefault(local, set()).add(child.tag)
                if child.tag.startswith(_W):
                    counts[local] = counts.get(local, 0) + 1
                    table[(local, counts[local])] = child
            clashes = {local for local, tags in locals_.items() if len(tags) > 1}
            table = self.__dict__["_top_table"] = (table, clashes)
        elements, clashes = table
        name, index = found.group(1), int(found.group(2))
        if name in clashes:
            return False
        return elements.get((name, index))

    def _row_ids(self) -> dict:
        """Each body table row's and cell's element -> its id (``tr:...``; a cell's
        ``<table id>/c<row>,<column>``, its column its grid origin, as :class:`Cell` counts it)."""
        rows = self.__dict__.get("_rows")
        if rows is None:
            from .edit.document import _cells, _find, _int

            document = self._document
            index = document._index(document.package.document_part())
            rows = {}
            for table in index.tables:
                for number, (row, row_id) in enumerate(table.rows):
                    rows[row] = row_id
                    column = _int(_find(row, "w:trPr/w:gridBefore"), 0)
                    for cell in _cells(row):
                        rows[cell] = f"{table.id}/c{number},{column}"
                        column += max(1, _int(_find(cell, "w:tcPr/w:gridSpan"), 1))
            self.__dict__["_rows"] = rows
        return rows

    def _block_order(self) -> dict:
        order = self.__dict__.get("_order")
        if order is None:
            document = self._document
            root = document.package.tree(document.package.document_part())
            # Keyed by the elements themselves (lxml's proxies, kept alive here), in document order.
            order = self.__dict__["_order"] = {element: k for k, element in enumerate(root.iter())}
        return order

    def _wanted(self, identifier: str) -> set[str]:
        document = self._document
        try:
            part, entry = document._resolve(identifier)
        except KeyError:
            return {self._canon(identifier)}
        if isinstance(entry, TableEntry):
            index = document._index(part)
            nested = set(map(id, entry.element.iter(_W + "p")))
            return {self._canon(e.id) for e in index.paragraphs if id(e.element) in nested}
        return {self._canon(entry.id), self._canon(identifier)}

    def _unknown(self, identifier: str) -> Unknown:
        if self.stopped is None:
            return Unknown("not-drawn")
        document = self._document
        try:
            part, entry = document._resolve(identifier)
        except KeyError:
            return Unknown("not-drawn")
        if part != document.package.document_part():
            return Unknown("not-drawn")
        stop_element = resolve_path(document.package.tree(part), self._stop_path())
        if stop_element is None:
            return Unknown(self.stopped.reason, self.stopped.page)
        index = document._index(part)
        order = {id(e.element): e.index for e in index.paragraphs}
        stop_at = min((order[id(p)] for p in stop_element.iter(_W + "p") if id(p) in order), default=None)
        mine = entry.index if isinstance(entry, ParagraphEntry) else min(
            (order[id(p)] for p in entry.element.iter(_W + "p") if id(p) in order), default=None)
        if stop_at is None or mine is None or mine >= stop_at:
            return Unknown(self.stopped.reason, self.stopped.page)
        return Unknown("not-drawn")

    def _stop_path(self) -> str:
        for page in self.pages:
            if page.stop is not None:
                return page.stop.path
        return ""

    def _canon(self, identifier: str) -> str:
        """The id an id has become: aliases followed to their end."""
        aliases = self._document._aliases
        seen = set()
        while identifier in aliases and identifier not in seen:
            seen.add(identifier)
            identifier = aliases[identifier]
        return identifier

    # -- compare -----------------------------------------------------------------------------

    def signature(self, page: int) -> str:
        """A digest of what a 1-based page draws: its lines (by paragraph id, not path),
        glyphs and positions, rules, pictures, floating drawings and stop."""
        return hashlib.sha256(repr(self._content(page - 1)).encode()).hexdigest()

    def _content(self, index: int, which: str = "all") -> tuple:
        """What page ``index`` draws: ``all`` of it, only its ``body``, or only its stories
        (headers, footers, notes)."""
        page = self.pages[index]
        lines = tuple(
            (self._canon(key) if key else line.path, line.story, line.baseline, line.top, line.pitch,
             tuple(_span(span) for span in line.spans))
            for key, line in self._lines[index]
            if which == "all" or (which == "body") == _in_body(line))
        if which == "stories":
            return lines
        rules = tuple((r.kind, r.x, r.y, r.width, r.height, r.color, r.style) for r in page.rules)
        pictures = tuple((p.relationship, p.part, p.x, p.y, p.width, p.height) for p in page.pictures)
        floats = tuple((f.x, f.y, f.width, f.height, f.behind, f.story, len(f.primitives)) for f in page.floats)
        stop = (page.stop.reason, page.stop.y, page.stop.remaining) if page.stop else None
        return (page.width_px, page.height_px, lines, rules, pictures, floats, stop)

    def compare(self, before: "DocumentLayout") -> Reflow:
        """Which pages of this layout differ from ``before``, an earlier layout of the same
        document."""
        known_before = before.stopped.page if before.stopped else len(before.pages)
        known_after = self.stopped.page if self.stopped else len(self.pages)
        limit = min(known_before, known_after)
        count = max(len(before.pages), len(self.pages))
        changed: list[int] = []
        why: dict[int, str] = {}
        for number in range(1, min(count, limit) + 1):
            if number > len(before.pages):
                changed.append(number)
                why[number] = "added"
            elif number > len(self.pages):
                changed.append(number)
                why[number] = "removed"
            elif before._content(number - 1) != self._content(number - 1):
                changed.append(number)
                if before._content(number - 1, "body") == self._content(number - 1, "body"):
                    why[number] = "stories"
                elif set(before._page_ids(number - 1)) != set(self._page_ids(number - 1)):
                    why[number] = "flow"
                else:
                    why[number] = "content"
        unknown = list(range(limit + 1, count + 1))
        first_before, first_after = before._first_pages(), self._first_pages()
        moved = {key: (first_before[key], page) for key, page in first_after.items()
                 if key in first_before and first_before[key] != page}
        return Reflow(changed=changed, first_changed_page=changed[0] if changed else None,
                      page_count=(len(before.pages), len(self.pages)), moved=moved,
                      unknown=unknown, why=why)

    def _page_ids(self, index: int) -> list[str]:
        return [self._canon(key) for key, line in self._lines[index] if key is not None and _in_body(line)]

    def _first_pages(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for index in range(len(self.pages)):
            for key in self._page_ids(index):
                out.setdefault(key, index + 1)
        return out


def _section_bodies(document: "Document") -> list[float | None]:
    """Each section's page height less its top and bottom margins, in points (``None``
    where its ``w:sectPr`` states no page height or margins), in document order -- the
    order docx2svg's ``PageInfo.section`` counts them in."""
    out: list[float | None] = []
    for section in document.sections():
        sect = section._sectPr
        size, margins = sect.find(_W + "pgSz"), sect.find(_W + "pgMar")
        try:
            height = int(size.get(_W + "h"))
            top, bottom = abs(int(margins.get(_W + "top"))), abs(int(margins.get(_W + "bottom")))
        except (AttributeError, TypeError, ValueError):
            out.append(None)
            continue
        out.append(round((height - top - bottom) / 20, 2))
    return out


def _in_body(line) -> bool:
    return line.story is None and _story_kind(line.path) == "body"


def _span(span) -> tuple:
    return (tuple(span.chars), tuple(span.xs), span.y, span.end, span.face, span.bold, span.italic,
            span.half_points, span.color, span.underline, span.strike, span.double_strike,
            span.highlight, span.shading, span.kind, span.vertical_align)


def _story_kind(path: str) -> str:
    if path.startswith("w:footnote["):
        return "footnotes"
    if path.startswith("w:endnote["):
        return "endnotes"
    if path.startswith("w:hdr"):
        return "header"
    if path.startswith("w:ftr"):
        return "footer"
    return "body"


class _Parts:
    """What :func:`docx2svg.paths.path_part` asks of a package, over ours."""

    def __init__(self, document: "Document") -> None:
        self._package = document.package
        self.main_document_part = document.package.document_part()

    def related(self, part: str, relationship_type: str) -> list[str]:
        return self._package.related_parts_of_type(part, relationship_type)


def _part_of(document: "Document", path: str, part: str | None = None) -> str:
    """The part a drawn object's path is in (:func:`docx2svg.paths.path_part`): the story's
    part a header's or footer's line names, the notes part for a note, else the main part."""
    return path_part(_Parts(document), path, part) or document.package.document_part()


def _paragraph_id(document: "Document", part: str, path: str, paths: "PathFinder | None" = None) -> str | None:
    """The id of the paragraph ``path`` names in ``part``; ``paths``: a
    :class:`~docx_agent.edit.effective.PathFinder` for many paths of one state at once."""
    tree = document.package.tree(part)
    element = paths.resolve(tree, path) if paths is not None else resolve_path(tree, path)
    if element is None or element.tag != _W + "p":
        return None
    try:
        entry = document._index(part).entry_for(element)
    except KeyError:
        return None
    return entry.id if entry is not None else None


def _block_id(document: "Document", part: str, path: str, paths: "PathFinder | None" = None) -> str | None:
    tree = document.package.tree(part)
    element = paths.resolve(tree, path) if paths is not None else resolve_path(tree, path)
    if element is None:
        return None
    entry = document._index(part).entry_for(element)
    return entry.id if entry is not None else None


# -- entry points ----------------------------------------------------------------------------

_CACHE_SIZE = 8


@dataclass
class _Conversion:
    """One state of the document laid out once: its layout in this API's ids, and what
    drawing its pages takes -- docx2svg's layout itself, the bytes and the faces it was
    measured with -- each page's SVG drawn when first asked for and tagged with ids."""

    layout: DocumentLayout
    raw: Any
    data: bytes
    fonts: Any
    options: dict
    tagged: dict[int, str] = field(default_factory=dict)
    #: Each page's SVG as docx2svg draws it (the old protocol of a ``converter``: all of
    #: them, drawn with the layout).
    svgs: dict[int, str] = field(default_factory=dict)
    #: Held while a page is drawn from this layout (its faces are written to as pages are
    #: drawn): lock 4 of the module's order.
    drawing: threading.Lock = field(default_factory=threading.Lock, repr=False, compare=False)
    #: Per 0-based page, the digest of what it draws (docx2svg's ``relayout.drawing``):
    #: what :mod:`docx_agent.liveview` compares versions by, made once.
    digests: dict[int, bytes] = field(default_factory=dict, repr=False, compare=False)


def check_options(options: dict, method: str) -> None:
    """Refuse a keyword docx2svg's :class:`~docx2svg.ConvertOptions` does not take, naming
    the ones it does (``pages`` is the render methods' own argument), before docx2svg's
    bare ``TypeError``.  Rendering always draws the final view -- every revision accepted,
    comments not shown -- by decision; reviewing is ``to_markdown(view="markup")``."""
    accepted = [f.name for f in dataclasses.fields(docx2svg.ConvertOptions) if f.name != "pages"]
    unknown = [name for name in options if name not in accepted]
    if not unknown:
        return
    hint = ""
    if any(name in ("view", "markup", "show_revisions", "revisions", "comments") for name in unknown):
        hint = ("; renders always show the final view (every revision accepted, comments hidden) -- "
                "to review revisions and comments in place, read doc.to_markdown(view=\"markup\")")
    pages = "" if method == "layout" else "pages, "
    raise TypeError(f"{method}() got unexpected keyword argument(s) {', '.join(map(repr, unknown))}; "
                    f"it takes {pages}{', '.join(accepted)}{hint}")


def cache_key(data: bytes, options: dict) -> str:
    return hashlib.sha256(data).hexdigest() + repr(sorted(options.items()))


def font_dirs_of(document: "Document | None") -> list[str]:
    """The font folders ``document`` lays out and renders with: its
    :attr:`~docx_agent.Document.font_dirs`, else ``OOXML_FONT_DIRS``
    (:func:`ooxml_common.fonts.office.user_font_dirs`), as strings."""
    from ooxml_common.fonts.office import user_font_dirs

    return [str(path) for path in user_font_dirs(getattr(document, "font_dirs", None))]


def locale_of(document: "Document | None") -> Any:
    """The regional settings ``document`` is drawn and edited in: its
    :attr:`~docx_agent.Document.locale`, else en-US (:data:`ooxml_common.locales.EN_US`)
    -- never the operating system's."""
    locale = getattr(document, "locale", None)
    if locale is None:
        from ooxml_common.locales import EN_US

        return EN_US
    return locale


def layout_options(document: "Document", options: dict) -> dict:
    """``options`` with the document's font folders (:func:`font_dirs_of`) as
    ``font_dirs`` and its locale (:func:`locale_of`) as ``locale`` unless they name their
    own: resolved here, so a layout run in a worker process -- which sees neither the
    document nor an environment variable set since it started -- uses the same folders and
    locale, and a layout with other folders or another locale is cached apart."""
    options = dict(options)
    if "font_dirs" not in options:
        options["font_dirs"] = font_dirs_of(document)
    if "locale" not in options:
        options["locale"] = locale_of(document)
    return options


def convert_bytes(data: bytes, options: dict) -> tuple:
    """docx2svg's conversion of a document's bytes: ``(layout, svgs, warnings)``, every one
    of them picklable.  What :func:`convert` runs unless the document names a
    :attr:`~docx_agent.Document.converter` -- which may run this very function in another
    process (an agent tool layer's worker pool, with a deadline) and return its answer."""
    convert_options = docx2svg.ConvertOptions(**options)
    result = docx2svg.convert_docx(data, convert_options)
    return result.layout, list(result.svgs), list(convert_options.warnings)


# A registered task of ooxml-edit's subprocess workers (``ooxml_edit.tools.worker_task``),
# which run nothing else: what a daemonic process (a Celery prefork child) lays out with.
# The mark is set as the attribute itself so that the layout does not import the tool layer.
convert_bytes.__ooxml_edit_task__ = True


# -- the layout kept for the next version (round 33, R16) -----------------------------------
#
# A document is laid out from the layout of its version before (docx2svg's ``lay_out_docx``
# with ``previous``): from the page an edit reaches until the pages line up again, the rest
# taken over -- page for page the layout of the new bytes alone (docx2svg's relayout checks;
# ``tests/test_incremental_layout.py`` here).  The layout is handed between versions as a
# pickle of the layout and its faces (:func:`_light`: without what its flow measured, most
# of a long document's pickle); the process that made it keeps it whole, with those
# measurements (:data:`_RECENT`), which the next layout in that process reuses.

#: The layouts this process made or read last, by a digest of their pickle (``laid_out``):
#: the next layout handed one back in the same process does not read it again, and finds
#: what its flow measured, which the pickle leaves out (:func:`_light`).
#:
#: A task takes its layout out (:func:`_taken`) and puts it back when it is done
#: (:func:`_remember`), so no two tasks running at once hold the same layout and the same
#: ``InstalledFonts``: laying out and drawing pages write to both, and with the GIL off
#: (Python 3.15t) two threads of an in-process runner would write at once.  A second task
#: asking for a layout another holds reads its own from the pickle.  Every use of the dict
#: holds :data:`_RECENT_LOCK`.  Two are kept: a long report's, with its measurements, is
#: about 80 MB in memory (79 pages; ``ROADMAP.md``, "Reflow feedback", round 33).
_RECENT: dict = {}
_RECENT_SIZE = 2
_RECENT_LOCK = threading.Lock()


def _light(layout):
    """``layout`` to pickle: without what its flow measured (``layout.state.memo``, which
    a layout of the next version reuses for the blocks that did not change -- most of a
    long document's pickle).  The process that made it keeps it (:data:`_RECENT`)."""
    state = getattr(layout, "state", None)
    if state is None or not getattr(state, "memo", None):
        return layout
    return dataclasses.replace(layout, state=dataclasses.replace(state, memo={}))


def _remember(blob: bytes, laid: tuple) -> None:
    """``laid`` -- ``(layout, fonts, warnings)`` -- kept as the newest under ``blob``'s
    digest, for the next task that brings ``blob`` back."""
    key = hashlib.sha256(blob).digest()
    with _RECENT_LOCK:
        _RECENT.pop(key, None)
        _RECENT[key] = laid
        while len(_RECENT) > _RECENT_SIZE:
            del _RECENT[next(iter(_RECENT))]


def _taken(blob: bytes) -> tuple:
    """``(layout, fonts, warnings)`` of a ``laid_out`` pickle, this task's alone: the one
    this process holds, taken out of :data:`_RECENT` until :func:`_remember` puts it back,
    else read from the pickle."""
    with _RECENT_LOCK:
        found = _RECENT.pop(hashlib.sha256(blob).digest(), None)
    if found is None:
        found = pickle.loads(blob)
    return found


def lay_out_bytes(data: bytes, options: dict, state: Any = None) -> tuple:
    """docx2svg's layout of a document's bytes, kept for the next version: ``(laid_out,
    facts)``, ``laid_out`` the pickle of ``(layout, fonts, warnings)`` (:func:`_light`) and
    ``facts`` -- from an earlier version's layout -- ``changed`` (the pages whose drawing
    changed), ``relaid`` (why it was laid out whole, docx2svg's
    :data:`~docx2svg.relayout.FALLBACKS` code, or ``None``) and ``relaid_cause``.

    ``state``: ``None``; a ``laid_out`` of these very bytes and options, handed back (not
    laid out again: the same pickle comes back); or ``("previous", laid_out)``, the layout
    of an earlier version, which the document is laid out again from -- with what its flow
    measured where this process still holds it (:data:`_RECENT`), else measured again
    (still faster than a whole layout: 63 pages, one paragraph edited, 6 s against 9 s; 1 s
    with the measurements).  What :func:`convert` runs -- through the
    document's :attr:`~docx_agent.Document.converter`, which may run it in another process
    (an agent tool layer's worker pool, with a deadline)."""
    if isinstance(state, (bytes, bytearray)):
        return bytes(state), {"changed": None, "relaid": None, "relaid_cause": None}
    previous = None
    if isinstance(state, tuple) and len(state) == 2 and state[0] == "previous":
        # Taken, not handed back: the new layout goes on with its fonts and measurements.
        previous = _taken(state[1])[0]
    convert_options = docx2svg.ConvertOptions(**options)
    layout, _, fonts = docx2svg.lay_out_docx(data, convert_options, previous=previous)
    laid = (layout, fonts, list(convert_options.warnings))
    blob = pickle.dumps((_light(layout), fonts, laid[2]), protocol=pickle.HIGHEST_PROTOCOL)
    _remember(blob, laid)
    if previous is None:
        return blob, {"changed": None, "relaid": None, "relaid_cause": None}
    return blob, {"changed": layout.changed_pages, "relaid": layout.relaid, "relaid_cause": layout.relaid_cause}


# A task ooxml-edit's subprocess workers may run (the mark), as :func:`convert_bytes` is.
lay_out_bytes.__ooxml_edit_task__ = True

#: The options that change what a layout draws, not where: a layout made with other values
#: of these is the same layout (docx2svg's ``relayout.options_key`` leaves them out too).
DRAWING_OPTIONS = frozenset({"pages", "width", "height", "glyph_size", "text", "selectable_text", "glyphs",
                             "warnings", "coverage"})


def layout_key(data: bytes, options: dict) -> tuple:
    """What one layout is of: the bytes' digest and the options that change it."""
    return (hashlib.sha256(data).digest(),
            repr(sorted((k, v) for k, v in options.items() if k not in DRAWING_OPTIONS)))


class _Flight:
    """A layout in progress: callers asking for the same state wait for it (:func:`convert`)."""

    def __init__(self) -> None:
        self.done = threading.Event()
        self.result: Any = None
        self.error: BaseException | None = None

    def wait(self) -> Any:
        self.done.wait()
        if self.error is not None:
            raise self.error
        return self.result


def _keeper(document: "Document") -> threading.Lock:
    """The lock over ``document``'s layouts: the cache, the layouts in progress and the
    one kept for the next version, which concurrent callers (threads; with the GIL off,
    really at once) share."""
    lock = getattr(document, "_layout_lock", None)
    if lock is None:
        with _KEEPERS_LOCK:
            lock = getattr(document, "_layout_lock", None)
            if lock is None:
                lock = threading.Lock()
                document._layout_lock = lock
                document._layout_flights = {}
                document._layout_kept = None
    return lock


_KEEPERS_LOCK = threading.Lock()


def kept_layout(document: "Document", data: bytes, options: dict) -> Any:
    """The state to hand :func:`lay_out_bytes` (or the render pipeline's layout task) for
    ``data``: the ``laid_out`` of these bytes when it is the one kept, else
    ``("previous", laid_out)`` of the document's last layout, else ``None``."""
    lock = _keeper(document)
    key = layout_key(data, options)
    with lock:
        kept = document._layout_kept
    if kept is None:
        return None
    if kept[0] == key:
        return kept[1]
    return ("previous", kept[1]) if kept[0][1] == key[1] else None


def keep_layout(document: "Document", data: bytes, options: dict, laid_out: bytes) -> None:
    """Keep ``laid_out``, the layout of ``data`` with ``options``, as the document's last:
    the one its next version is laid out from (one per document: about 5 MB for 79 pages)."""
    lock = _keeper(document)
    with lock:
        document._layout_kept = (layout_key(data, options), laid_out)


#: :func:`convert`'s ``converter`` left out: the document's own.
_DOCUMENTS = object()


def convert(document: "Document", *, unlocked=None, current=None, converter: Any = _DOCUMENTS,
            **options) -> _Conversion:
    """The document now, laid out once -- every page, so the layout answers for all of
    them -- and cached by its bytes and the options.  Both :func:`lay_out` and
    :func:`render_svg` read it, so a render and a layout of the same state cost one layout.
    docx2svg finds the faces Word draws itself, Office's cloud-font cache among them (Aptos
    Display, a new document's heading face).

    **Incremental** (round 33): laid out from the document's last layout
    (:func:`lay_out_bytes`; the render tool's layouts are kept the same way, so a check
    after a render of the same state lays nothing out), never from a stale one: docx2svg
    lays out again from the first changed block, or the whole document where it cannot
    (:data:`docx2svg.relayout.FALLBACKS`).  **Shared**: callers asking for the same state
    while it is laid out wait for that one layout -- only those that let the document's lock
    go to wait (``unlocked`` given, not yielding ``False``); a caller holding it, or with no
    ``unlocked`` (``Document.layout()``), lays out on its own
    (round 34, R17: the order is in the module's docstring).

    ``unlocked`` (a context manager factory) runs the slow part -- docx2svg's layout of the
    bytes -- inside it, where another thread may change ``document`` (an agent tool layer
    lets the document's lock go there; :func:`lay_out_unlocked`).  The bytes and options are
    taken first; after, the layout is mapped to ids on ``document`` if ``current()`` says it
    has not changed meanwhile, else on a copy opened from those bytes (the same ids).
    Either way it is the layout of the state the bytes were taken from, cached under them.
    ``converter``: who lays it out, in place of the document's (:attr:`~docx_agent.Document.converter`,
    which a tool call points at its own engine while it runs): ``None`` is docx2svg in this
    process -- an application's own thread (:mod:`docx_agent.liveview`)."""
    if "pages" in options:
        raise TypeError("convert() lays out every page; select pages when rendering")
    options = layout_options(document, options)
    data = document.to_bytes()
    key = cache_key(data, options)
    lock = _keeper(document)
    with lock:
        cached = document._layouts.get(key)
        if cached is not None:
            return cached
    outside = unlocked if unlocked is not None else _held
    # Taken with the bytes, before the lock is let go: another call on the document points
    # ``converter`` at its own engine, or back at none, when it ends (round 33, R16c).
    if converter is _DOCUMENTS:
        converter = document.converter
    with outside() as entered:
        # Round 34 (R17): only a caller that let the document's lock go waits for a layout in
        # progress (``unlocked()`` given, and not yielding ``False``: ooxml-edit's
        # ``Call.unlocked`` yields whether it let the lock go).  One that holds it -- a batch's op,
        # a read_batch's read, a layout under the lock (``Document.layout``) -- lays out on
        # its own: the layout's owner takes that lock back once it is done.
        with lock:
            cached = document._layouts.get(key)
            flight = None if cached is not None else document._layout_flights.get(key)
            leader = cached is None and flight is None
            if leader:
                flight = document._layout_flights[key] = _Flight()
        if cached is not None:
            result = None
        elif leader:
            try:
                flight.result = result = _lay_out(document, data, options, converter)
            except BaseException as error:
                flight.error = error
                raise
            finally:
                # The flight ends here, before ``outside()`` takes the document's lock back:
                # a caller holding it may be waiting for this layout (never with R17's rule,
                # but a flight never waits on the lock it may be waited for under).
                with lock:
                    document._layout_flights.pop(key, None)
                flight.done.set()
        elif unlocked is not None and entered is not False:
            result = flight.wait()
        else:
            result = _lay_out(document, data, options, converter)
    if cached is not None:
        return cached
    raw, fonts, warnings, svgs = result
    # Mapped to ids once: the first caller back maps it, under the lock over the layouts
    # (never held while waiting for anything else), and the others find it cached.
    with lock:
        cached = document._layouts.get(key)
        if cached is not None:
            return cached
        owner = document if current is None or current() else type(document).open(data)
        conversion = _Conversion(DocumentLayout(owner, raw, warnings), raw, data, fonts, options,
                                 svgs=dict(enumerate(svgs or (), start=1)))
        return keep(document, key, conversion)


@contextmanager
def _held():
    yield


def _lay_out(document: "Document", data: bytes, options: dict, converter: Any = None) -> tuple:
    """``(layout, fonts, warnings, svgs)`` of ``data``: through the document's converter --
    its :func:`lay_out_bytes` (``converter.lay_out``), from the document's last layout --
    or, for a converter of the old protocol (``converter(data, options) -> (layout, svgs,
    warnings)``), whole with every page's SVG.  The layout is this caller's own copy (the
    pickle read): the one kept for the next version is never one a caller holds.
    ``converter``: the document's when its bytes were taken."""
    if converter is not None and not hasattr(converter, "lay_out"):
        raw, svgs, warnings = converter(data, dict(options))
        return raw, None, warnings, list(svgs)
    state = kept_layout(document, data, options)
    if isinstance(state, bytes):
        blob = state                         # laid out already (a render of this state)
    else:
        run = converter.lay_out if converter is not None else lay_out_bytes
        blob, _facts = run(data, dict(options), state)
        keep_layout(document, data, options, blob)
    raw, fonts, warnings = pickle.loads(blob)
    return raw, fonts, list(warnings), None


def keep(document: "Document", key: str, conversion: _Conversion) -> _Conversion:
    """Cache ``conversion`` -- of the bytes and options ``key`` names -- on ``document``."""
    document._layouts[key] = conversion
    while len(document._layouts) > _CACHE_SIZE:
        document._layouts.pop(next(iter(document._layouts)))
    return conversion


def twin_of(document: "Document") -> dict:
    """What a copy of ``document`` needs to lay out as it does (:func:`twin`): its font
    folders, locale, stamping, converter and last layout.  Read under the document's lock."""
    lock = _keeper(document)
    with lock:
        kept = document._layout_kept
    return {"stamping": document.stamping, "font_dirs": getattr(document, "font_dirs", None),
            "locale": getattr(document, "locale", None), "converter": document.converter,
            "kept": kept}


def twin(data: bytes, like: dict) -> "Document":
    """A copy opened from ``data`` -- a document's bytes now -- that lays out as the document
    does (``like``: :func:`twin_of`), from its last layout: work on it with the document's
    lock let go, then :func:`adopt` its layouts (round 33, R16c: ``word_fields`` update)."""
    from .edit.document import Document

    copy = Document.open(data, stamping=like["stamping"])
    copy.font_dirs = like["font_dirs"]
    copy.locale = like["locale"]
    copy.converter = like["converter"]
    _keeper(copy)
    copy._layout_kept = like["kept"]
    return copy


def adopt(document: "Document", copy: "Document") -> None:
    """Keep ``copy``'s layouts (:func:`twin`) as ``document``'s: each cached under the bytes
    and options it is of, so the document finds one wherever it comes to those very bytes
    (an update of its fields writing what the copy's wrote), and the copy's last as the one
    its next version is laid out from.  Correct for any state: a layout is looked up by
    its bytes' digest."""
    with _keeper(copy):
        conversions = list(copy._layouts.items())
        kept = copy._layout_kept
    with _keeper(document):
        for key, conversion in conversions:
            if key not in document._layouts:
                keep(document, key, conversion)
        if kept is not None:
            document._layout_kept = kept


def lay_out_unlocked(document: "Document", unlocked, current) -> DocumentLayout:
    """:func:`lay_out`, docx2svg's part run inside ``unlocked()`` (:func:`convert`)."""
    return convert(document, unlocked=unlocked, current=current).layout


def lay_out(document: "Document", **options) -> DocumentLayout:
    """The document's layout now (:func:`convert`'s)."""
    return convert(document, **options).layout


_GROUP = re.compile(r'<g data-docx-path="([^"]*)"((?: data-docx-id="[^"]*")?)'
                    r'((?: data-docx-story="[^"]*" data-docx-part="([^"]*)")?)>')


def render_svg(document: "Document", pages: list[int] | None = None, **options) -> list[str]:
    """docx2svg's SVG of the pages (1-based; every page by default), every paragraph group
    carrying ``data-docx-agent-id``; from the same conversion as :func:`lay_out`."""
    conversion = convert(document, **options)
    count = len(conversion.raw.pages)
    selected = list(range(1, count + 1)) if pages is None else [n for n in pages if 1 <= n <= count]
    missing = [n for n in selected if n not in conversion.tagged and n not in conversion.svgs]
    out = []
    with conversion.drawing:
        missing = [n for n in selected if n not in conversion.tagged and n not in conversion.svgs]
        if missing:
            # Drawn from the layout as docx2svg's convert_docx draws them (page_svgs), only the
            # pages asked for.
            drawing = docx2svg.ConvertOptions(**{**conversion.options, "pages": missing})
            for number, svg in zip(missing, docx2svg.page_svgs(conversion.raw, conversion.data, drawing,
                                                               fonts=conversion.fonts)):
                conversion.svgs[number] = svg
        for number in selected:
            if number not in conversion.tagged:
                conversion.tagged[number] = _tag(document, conversion.layout, conversion.svgs.pop(number))
            out.append(conversion.tagged[number])
    return out


def _tag(document: "Document", layout: DocumentLayout, svg: str) -> str:
    """``data-docx-agent-id`` on every paragraph group of one page's SVG."""
    return tag_with(svg, lambda path, part: layout.id_for_path(_part_of(document, path, part), path))


def tag_with(svg: str, id_of: Any) -> str:
    """``data-docx-agent-id`` on every paragraph group of one page's SVG, the id
    ``id_of(path, part)`` gives for its ``data-docx-path`` and ``data-docx-part`` (``None``
    for a group of the body)."""

    def tag(match: re.Match) -> str:
        identifier = id_of(match.group(1), match.group(4))
        if identifier is None:
            return match.group(0)
        return match.group(0)[:-1] + f' data-docx-agent-id="{_escape(identifier)}">'

    return _GROUP.sub(tag, svg)


def group_ids(document: "Document", conversion: _Conversion) -> dict:
    """``(path, part) -> id`` for every paragraph group the layout's pages draw, as
    :func:`_tag` would tag it (``part`` the group's ``data-docx-part``, ``None`` in the
    body): read once from the document the layout is of, under its lock, so that pages
    drawn later -- the document changed meanwhile -- are tagged with that version's ids."""
    layout = conversion.layout
    out: dict = {}
    for page in conversion.raw.pages:
        for line in page.text_lines():
            key = (line.path, getattr(line, "part", None) if getattr(line, "story", None) else None)
            if key not in out:
                out[key] = layout.id_for_path(_part_of(document, *key), key[0])
    return out


def _escape(value: str) -> str:
    return value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
