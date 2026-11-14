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
   document order; ``Call.unlocked()`` lets the call's go and takes them back in that order)
   -- or, for a read (round 45), the lock of the read copy of its committed version
   (:func:`read_copy`; ``Session.read_lock``), which stands in its place: the copy shares
   the document's layouts, so the locks below are the document's;
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
import logging
import pickle
import re
import threading
import time
import uuid
from collections import OrderedDict
from contextlib import contextmanager
from dataclasses import dataclass, field
from fractions import Fraction
from typing import TYPE_CHECKING, Any

import docx2svg
from docx2svg.paths import path_part, resolve_path
from docx2svg.relayout import MeasureStore
from ooxml_edit.tools.results import ToolError

from .edit.effective import PathFinder
from .edit import ids as _ids
from .edit.ids import ParagraphEntry, TableEntry

if TYPE_CHECKING:  # pragma: no cover
    from .edit.document import Document

#: docx2svg's device pixels (Word's 1/300-inch export grid) per point.
PX_PER_PT = Fraction(300, 72)

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
#: A key naming one of the body's children (:meth:`DocumentLayout._top_level`).
_TOP_KEY = re.compile(r"^w:body/w:([A-Za-z]+)\[(\d+)\]$")
#: One step of a path that names a child by its place (``w:tc[2]``): resolved alike from
#: any element (:meth:`DocumentLayout._element_of`).
_INDEXED_STEP = re.compile(r"w:[A-Za-z]+\[\d+\]")


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
        #: The :func:`layout_key` it is of, where :func:`convert` made it (round 38).
        self.key: tuple | None = None
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
        drawn = self._drawn()
        found = sorted((at for key in wanted for at in drawn.get(key, ())), key=lambda at: at[:2])
        pages: dict[int, list] = {}
        for number, _, line in found:
            pages.setdefault(number, []).append(line)
        for number, lines in pages.items():
            columns: dict[int, list] = {}
            for line in lines:
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

    def _drawn(self) -> dict[str, list[tuple[int, int, Any]]]:
        """Round 45: each id (aliases followed, :meth:`_canon`) -> where its lines are drawn,
        ``(page index, place on the page, line)`` in drawing order -- what :meth:`where`
        looked for by reading every line of every page, once per id asked (a table of
        contents asked for each of its headings).  Made again when the aliases changed."""
        aliases = self._document._aliases
        kept = self.__dict__.get("_drawn_by")
        if kept is not None and kept[0] is aliases and kept[1] == aliases:
            return kept[2]
        out: dict[str, list[tuple[int, int, Any]]] = {}
        canon: dict[str, str] = {}
        for number, lines in enumerate(self._lines):
            for place, (key, line) in enumerate(lines):
                if key is None:
                    continue
                name = canon.get(key)
                if name is None:
                    name = canon[key] = self._canon(key)
                out.setdefault(name, []).append((number, place, line))
        self.__dict__["_drawn_by"] = (aliases, dict(aliases), out)
        return out

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

    def _element_of(self, key: str) -> Any:
        """The body's element docx2svg's block key names (:func:`docx2svg.paths.resolve_path`),
        cached: a key one indexed step below another (``.../w:tr[3]/w:tc[2]``) is resolved
        from that one's element -- resolving a path is step after step, so the same element
        -- not from the part's root again (round 38: a long report's rows and cells, read
        from the root each, took most of a scoped check's time)."""
        elements = self.__dict__.setdefault("_key_elements", {})
        if key in elements:
            return elements[key]
        element = self._top_level(key)
        if element is False:
            parent, _, step = key.rpartition("/")
            if parent and _INDEXED_STEP.fullmatch(step):
                above = self._element_of(parent)
                element = None if above is None else resolve_path(above, step)
            else:
                document = self._document
                element = resolve_path(document.package.tree(document.package.document_part()), key)
        elements[key] = element
        return element

    def _key_id(self, key: str) -> "tuple[str | None, Any]":
        """The id of the body block docx2svg's key names (a paragraph's, a table's, a
        row's or a cell's), and its element, cached."""
        keys = self.__dict__.setdefault("_key_ids", {})
        if key not in keys:
            document = self._document
            part = document.package.document_part()
            element = self._element_of(key)
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


#: A whole fraction as ``repr`` writes it, ``Fraction(300, 1)``: read as the integer it
#: equals (:func:`_digest`).
_WHOLE_FRACTION = re.compile(r"Fraction\((-?\d+), 1\)")


def _digest(value: Any) -> bytes:
    """A digest of ``value`` -- tuples of strings, numbers, booleans and ``None``, as a
    layout's drawing is -- equal for values ``==`` calls equal: a whole fraction is written
    as the integer it equals (``Fraction(300, 1)`` and ``300``); any other fraction's
    ``repr`` is its value's alone."""
    return hashlib.blake2b(_WHOLE_FRACTION.sub(r"\1", repr(value)).encode(), digest_size=16).digest()


def _page_digest(page: Any) -> bytes:
    """What a page draws besides its lines (:class:`LayoutSummary`)."""
    rules = tuple((r.kind, r.x, r.y, r.width, r.height, r.color, r.style) for r in page.rules)
    pictures = tuple((p.relationship, p.part, p.x, p.y, p.width, p.height) for p in page.pictures)
    floats = tuple((f.x, f.y, f.width, f.height, f.behind, f.story, len(f.primitives)) for f in page.floats)
    stop = (page.stop.reason, page.stop.y, page.stop.remaining) if page.stop else None
    return _digest((page.width_px, page.height_px, rules, pictures, floats, stop))


def _line_digest(line: Any) -> bytes:
    """What a line draws (:class:`LayoutSummary`)."""
    return _digest((line.story, line.baseline, line.top, line.pitch, tuple(_span(span) for span in line.spans)))


class LayoutSummary:
    """What :meth:`DocumentLayout.compare` reads of a layout, small enough to keep for every
    check and to send from the worker that made the layout (round 38): per page, a digest
    of what it draws besides its lines, and per line its paragraph's id (or path), whether
    it is the body's, and a digest of what it draws.  Compared with the ids followed through
    the document's aliases at the time (``canon``), it tells the same pages as ``compare``
    does on the layouts themselves."""

    def __init__(self, layout: "DocumentLayout", before: "LayoutSummary | None" = None, *,
                 keep: bool = False) -> None:
        """``before``: the summary of the layout this one was laid out from, made with
        ``keep`` -- each page and line docx2svg took over from it (the very object: a layout
        kept from the previous one keeps the pages the edit did not reach, and a page or line
        is never changed once laid out) has the digest made of it there (round 39: only what
        was laid out or drawn again is digested).  ``keep`` holds this summary's digests by
        page and line for the next, with the pages and lines themselves (not pickled)."""
        self.count = len(layout.pages)
        self.known = layout.stopped.page if layout.stopped else len(layout.pages)
        self.rest: list[bytes] = []
        self.lines: list[tuple] = []
        taken = getattr(before, "_drawn", None) or {}
        drawn: dict[int, tuple] = {}

        def digest_of(item: Any, make: Any) -> bytes:
            found = taken.get(id(item))
            value = found[1] if found is not None and found[0] is item else make(item)
            drawn[id(item)] = (item, value)
            return value

        for index, page in enumerate(layout.pages):
            self.rest.append(digest_of(page, _page_digest))
            self.lines.append(tuple((key, line.path, _in_body(line), digest_of(line, _line_digest))
                                    for key, line in layout._lines[index]))
        #: id -> (page or line, its digest), for the next summary (``keep``), else ``None``.
        self._drawn: dict[int, tuple] | None = drawn if keep else None

    def __getstate__(self) -> dict:
        # Sent to the caller without the pages and lines kept for the next summary.
        state = dict(self.__dict__)
        state["_drawn"] = None
        return state

    def _content(self, index: int, which: str, canon: Any) -> tuple:
        lines = tuple((canon(key) if key else path, drawn) for key, path, body, drawn in self.lines[index]
                      if which == "all" or (which == "body") == body)
        return (self.rest[index], lines)

    def _page_ids(self, index: int, canon: Any) -> list[str]:
        return [canon(key) for key, _, body, _ in self.lines[index] if key is not None and body]

    def _first_pages(self, canon: Any) -> dict[str, int]:
        out: dict[str, int] = {}
        for index in range(self.count):
            for key in self._page_ids(index, canon):
                out.setdefault(key, index + 1)
        return out

    def compare(self, before: "LayoutSummary", canon: Any) -> Reflow:
        """:meth:`DocumentLayout.compare` of the layouts these summarise: which pages of
        this one differ from ``before``'s, ids followed by ``canon``."""
        limit = min(before.known, self.known)
        count = max(before.count, self.count)
        changed: list[int] = []
        why: dict[int, str] = {}
        for number in range(1, min(count, limit) + 1):
            if number > before.count:
                changed.append(number)
                why[number] = "added"
            elif number > self.count:
                changed.append(number)
                why[number] = "removed"
            elif before._content(number - 1, "all", canon) != self._content(number - 1, "all", canon):
                changed.append(number)
                if before._content(number - 1, "body", canon) == self._content(number - 1, "body", canon):
                    why[number] = "stories"
                elif set(before._page_ids(number - 1, canon)) != set(self._page_ids(number - 1, canon)):
                    why[number] = "flow"
                else:
                    why[number] = "content"
        unknown = list(range(limit + 1, count + 1))
        first_before, first_after = before._first_pages(canon), self._first_pages(canon)
        moved = {key: (first_before[key], page) for key, page in first_after.items()
                 if key in first_before and first_before[key] != page}
        return Reflow(changed=changed, first_changed_page=changed[0] if changed else None,
                      page_count=(before.count, self.count), moved=moved, unknown=unknown, why=why)


def canon_of(document: "Document") -> Any:
    """The id an id has become in ``document`` now: its aliases followed to their end (as
    :meth:`DocumentLayout._canon` does on the layout's document)."""
    aliases = document._aliases

    def canon(identifier: str) -> str:
        seen = set()
        while identifier in aliases and identifier not in seen:
            seen.add(identifier)
            identifier = aliases[identifier]
        return identifier

    return canon


def summary_for(document: "Document", key: tuple, made: Any) -> LayoutSummary:
    """The :class:`LayoutSummary` of the layout of ``key`` (:func:`layout_key`), one per
    layout the document keeps (the same object for the same layout, as a cached
    :class:`DocumentLayout` is): the one kept, else ``made`` -- a summary, or a
    :class:`DocumentLayout` to summarise -- kept."""
    lock = _keeper(document)
    with lock:
        found = document._layout_summaries.get(key)
        if found is not None:
            document._layout_summaries.move_to_end(key)
            return found
    summary = made if isinstance(made, LayoutSummary) else LayoutSummary(made)
    with lock:
        found = document._layout_summaries.get(key)
        if found is not None:
            return found
        document._layout_summaries[key] = summary
        while len(document._layout_summaries) > _CACHE_SIZE:
            document._layout_summaries.popitem(last=False)
    return summary


def known_summary(document: "Document", key: tuple) -> "LayoutSummary | None":
    """The summary kept for ``key``'s layout, or ``None``."""
    with _keeper(document):
        return document._layout_summaries.get(key)


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


# -- a layout that fails (round 53, R107) ---------------------------------------------------

logger = logging.getLogger("docx_agent.layout")


class LayoutFailed(ToolError, RuntimeError):
    """Laying the document out failed (round 53, R107): ``layout_failed``, never
    ``not_found`` (which sends a model looking for an address that is there).  ``at`` is the
    block it failed at -- a table's ``t:`` id -- where docx2svg names it
    (:class:`docx2svg.LayoutError`), else ``None``; ``reason`` what went wrong.

    A version whose layout failed fails again at once, from the failure kept for it
    (:func:`convert`), and is logged once: an application's live view asking for the same
    version 900 times lays it out, and logs it, once."""

    def __init__(self, reason: str, *, at: str | None = None, path: str | None = None) -> None:
        where = f" at {at}" if at else (f" at {path}" if path else "")
        super().__init__("layout_failed", f"laying the document out failed{where}: {reason}. The document is "
                         "unchanged by it; check (validate, layout's table_grid_mismatch) says what is off",
                         valid_options=[at] if at else [], details={"at": at, "path": path, "reason": reason})
        self.at = at
        self.path = path
        self.reason = reason

    def __reduce__(self):
        return (_rebuild_failed, (self.reason, self.at, self.path))


def _rebuild_failed(reason: str, at: str | None, path: str | None) -> "LayoutFailed":
    return LayoutFailed(reason, at=at, path=path)


#: The failures a document keeps, by the layout's cache key (:func:`convert`).
_FAILURES_SIZE = 8

#: What a layout may raise that is the call's, not the layout's: passed on as raised.
_PASSED_ON = (ToolError, TimeoutError, MemoryError)


def layout_failure(document: "Document", error: BaseException) -> "LayoutFailed":
    """``error``, raised laying ``document`` out, as :class:`LayoutFailed`: the block id of the
    element docx2svg names (:class:`docx2svg.LayoutError`'s ``path``), where it resolves."""
    if isinstance(error, LayoutFailed):
        return error
    path = getattr(error, "path", None) if isinstance(error, docx2svg.LayoutError) else None
    reason = error.reason if isinstance(error, docx2svg.LayoutError) else f"{type(error).__name__}: {error}"
    at = None
    if path:
        try:
            at = _block_id(document, _part_of(document, path), path)
        except Exception:  # noqa: BLE001 -- a path the document no longer resolves: named by its path
            at = None
    return LayoutFailed(reason, at=at, path=path)


def _deterministic(error: BaseException) -> bool:
    """Whether ``error`` is docx2svg's answer for these bytes -- its :class:`docx2svg.LayoutError`,
    or raised inside docx2svg -- which laying them out again would only repeat; not a worker
    lost or another passing failure."""
    if isinstance(error, docx2svg.LayoutError):
        return True
    trace = error.__traceback__
    package = docx2svg.__file__.rsplit("/", 1)[0] if docx2svg.__file__ else None
    while trace is not None:
        if package and trace.tb_frame.f_code.co_filename.startswith(package):
            return True
        trace = trace.tb_next
    return False


def _failed(document: "Document", key: str, error: BaseException) -> "LayoutFailed":
    """``error`` as :class:`LayoutFailed`, kept as the failure of the layout ``key`` names where
    it is docx2svg's own (:func:`_deterministic`), and logged once; the failure."""
    failure = layout_failure(document, error)
    first = True
    if _deterministic(error):
        with _keeper(document):
            failures = document._layout_failures
            first = key not in failures
            failures[key] = failure
            while len(failures) > _FAILURES_SIZE:
                failures.popitem(last=False)
    if first:
        logger.warning("docx-agent: %s", failure.message, exc_info=error if error is not failure else None)
    return failure


def failed_layout(document: "Document", data: bytes, error: BaseException) -> "LayoutFailed":
    """``error``, met laying out ``data`` (``document``'s bytes) with the document's options, as
    :class:`LayoutFailed`, kept for that version and logged once (``render``'s workers)."""
    return _failed(document, cache_key(data, layout_options(document, {})), error)


def _kept_failure(document: "Document", key: str) -> "LayoutFailed | None":
    with _keeper(document):
        return document._layout_failures.get(key)


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
    comments not shown -- by decision; reviewing is ``to_markdown(view="markup")``.  The view
    is the accepted result, which can differ from Word's *No Markup* screen for pending
    table-structure changes (a tracked column deletion or merge: round 54)."""
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
#: holds :data:`_RECENT_LOCK`.  Kept within a quarter of :data:`LAYOUT_MEMORY` (round 39;
#: two before), the newest always: the layouts held by token (below) are what a tool layer's
#: tasks use, these a caller's own pickles.
_RECENT: dict = {}
_RECENT_LOCK = threading.Lock()


def _memory_budget() -> int:
    """:data:`LAYOUT_MEMORY`: ``DOCX_AGENT_LAYOUT_MEMORY`` (MB) where it is set to a
    positive number, else 256 MB."""
    import os

    try:
        value = float(os.environ.get("DOCX_AGENT_LAYOUT_MEMORY", ""))
    except ValueError:
        value = 0.0
    return int(value * 2 ** 20) if value > 0 else 256 * 2 ** 20


#: What one process -- a layout worker, or this one where it lays out itself -- keeps of the
#: layouts it made, in bytes (round 39): the layouts held by token (:data:`_HELD`) up to
#: this, the measurements kept by lineage (:data:`_MEASURES`) up to half of it, a caller's
#: own pickled layouts (:data:`_RECENT`) up to a quarter -- at most 1.75 times this in all,
#: plus the layouts tasks running at that moment have taken out.  The least recently used go
#: first; the newest of each is kept whatever its size.  ``DOCX_AGENT_LAYOUT_MEMORY`` (in MB)
#: sets it, read as the process starts (a worker gets the environment its pool starts it
#: with); 256 MB by default: about 23 ten-page documents' layouts, or five 46-page reports'.
LAYOUT_MEMORY = _memory_budget()
#: What a layout kept costs per page, in bytes: the layout, what its flow measured and what
#: was read from it since (round 39, measured with ``tracemalloc``: 1.05 MB a page on a
#: made-up 11-page document, 0.95 MB on a 46-page report).
PAGE_COST = 2 ** 20
#: What one measurement kept by lineage costs, in bytes (a block's lines; measured: 70 kB
#: on the 11-page document, 87 kB on the report).
MEASURE_COST = 96 * 2 ** 10


def layout_cost(layout: Any) -> int:
    """What keeping ``layout`` costs, in bytes: :data:`PAGE_COST` a page (one at least)."""
    return max(1, len(getattr(layout, "pages", None) or ())) * PAGE_COST


#: What this process measured in its layouts (docx2svg's ``MeasureStore``, round 37): a
#: layout handed a previous one without its measurements -- made in another worker -- takes
#: those of the blocks this process measured in an earlier version of the document, instead
#: of measuring every block again (46 pages, one paragraph edited: 0.6-0.7 s of CPU against
#: 1.9-2.5 s).  Its own lock; nothing is waited for under it.  Round 39: as many lineages as
#: half of :data:`LAYOUT_MEMORY` holds (:data:`MEASURE_COST` an entry), not two.
_MEASURES = MeasureStore(lineages=1024, entries=LAYOUT_MEMORY // 2 // MEASURE_COST)


def _fraction(numerator: int, denominator: int) -> Fraction:
    """A :class:`~fractions.Fraction` from its own terms, already in lowest terms: what a
    layout's pickle (:func:`dumps`) rebuilds each with, without the ``gcd`` the
    constructor takes."""
    return _FROM_TERMS(numerator, denominator)


_FROM_TERMS = getattr(Fraction, "_from_coprime_ints", Fraction)  # Python 3.12 and later


def _reduce_fraction(value: Fraction) -> tuple:
    return (_fraction, (value.numerator, value.denominator))


def dumps(value: Any) -> bytes:
    """``value`` (a layout, its faces and warnings) pickled, its fractions -- a long
    document's layout holds hundreds of thousands -- rebuilt from their terms as they are
    (round 37: a third less time to pickle and read back a 46-page layout)."""
    import copyreg
    import io

    buffer = io.BytesIO()
    pickler = pickle.Pickler(buffer, protocol=pickle.HIGHEST_PROTOCOL)
    pickler.dispatch_table = {**copyreg.dispatch_table, Fraction: _reduce_fraction}
    pickler.dump(value)
    return buffer.getvalue()


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
        budget = LAYOUT_MEMORY // 4
        while len(_RECENT) > 1 and sum(layout_cost(kept[0]) for kept in _RECENT.values()) > budget:
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


# -- layouts held where they were made (round 38) ----------------------------------------------
#
# A layout stays in the process that made it -- a worker, or this process when it lays out
# itself -- whole, with what its flow measured, and the next task of the document is sent
# there (ooxml-edit's pool affinity, keyed by the document's lineage): it is never pickled to
# the caller and back for the next version.  The caller names layouts by their token
# (:func:`layout_key` of the bytes and options they are of) and hands a task
# ``("held", now, kept, lineage)`` (:func:`kept_layout`): ``now`` the token of the bytes it
# hands, ``kept`` the token of the document's last layout (``None``: none yet), ``lineage``
# the document's own (:func:`lineage`); a fifth item, false, says the caller holds the
# document's lock (below).  The task takes, in order: the layout of ``now``
# (nothing to lay out), the one of ``kept``, the newest of the lineage this process holds
# with the same options (an older version: laid out again from it, as from any earlier
# version), else none (laid out whole) -- the same pages whichever it found.

class _Held:
    """A layout a process holds: ``(layout, fonts, warnings)``, and what was read from it
    since (:attr:`extra`: the document opened from its bytes, its :class:`DocumentLayout`,
    its summary), all of one state."""

    __slots__ = ("laid", "extra", "lineage", "cost")

    def __init__(self, laid: tuple, extra: dict | None = None) -> None:
        self.laid = laid
        self.extra: dict = extra if extra is not None else {}
        self.lineage: str | None = None
        #: What keeping it costs (:func:`layout_cost`).
        self.cost = layout_cost(laid[0])


#: The layouts this process holds, by token, the newest last: one per lineage, within
#: :data:`LAYOUT_MEMORY` (round 39: :func:`layout_cost`, a megabyte a page; two layouts
#: before -- with six documents built at once on two workers, every check laid its document
#: out whole).  A task takes its layout out (:func:`_take_held`) and puts the new one back
#: (:func:`_put_held`), so no two tasks running at once -- threads of one process -- hold one.
_HELD: "OrderedDict[tuple, _Held]" = OrderedDict()
_HELD_LOCK = threading.Condition()
#: Per lineage, the tasks of this process that took one of its layouts (or are making the
#: first): another task of the lineage waits for one of them, at most :data:`_HELD_WAIT`
#: s, rather than lay the document out whole (threads only: a worker process runs one task)
#: -- unless its caller holds the document's lock (round 34's rule: it never waits for a
#: layout in progress; it lays out on its own).
_HELD_TAKEN: dict[str, int] = {}
_HELD_WAIT = 5.0


def held_state(state: Any) -> bool:
    """Whether ``state`` names layouts by token (``("held", now, kept, lineage[, may_wait])``)."""
    return isinstance(state, tuple) and len(state) in (4, 5) and state[0] == "held"


def _take_held(now: tuple, kept: "tuple | None", lineage: str, may_wait: bool = True
               ) -> "tuple[_Held | None, str]":
    """The layout a task of ``now`` starts from, taken out, and how it was found:
    ``"same"`` (of ``now``), ``"kept"``, ``"lineage"`` (an older one of the document), or
    ``"fresh"`` (none: ``None``) -- only the lineage's own.  The lineage is counted as taken
    either way."""
    deadline = time.monotonic() + _HELD_WAIT
    with _HELD_LOCK:
        while True:
            found, how = None, "fresh"
            for token, name in ((now, "same"), (kept, "kept")):
                # Only the document's own (its lineage): another document of the same bytes
                # lays out afresh, as it always did.
                if token is not None and token in _HELD and _HELD[token].lineage == lineage:
                    found, how = _HELD.pop(token), name
                    break
            if found is None:
                for token in reversed(list(_HELD)):
                    if _HELD[token].lineage == lineage and token[1] == now[1]:
                        found, how = _HELD.pop(token), "lineage"
                        break
            if found is not None or not _HELD_TAKEN.get(lineage) or not may_wait:
                _HELD_TAKEN[lineage] = _HELD_TAKEN.get(lineage, 0) + 1
                return found, how
            left = deadline - time.monotonic()
            if left <= 0:
                _HELD_TAKEN[lineage] = _HELD_TAKEN.get(lineage, 0) + 1
                return None, "fresh"
            _HELD_LOCK.wait(left)


def _let_go_held(lineage: str) -> None:
    with _HELD_LOCK:
        count = _HELD_TAKEN.get(lineage, 0) - 1
        if count > 0:
            _HELD_TAKEN[lineage] = count
        else:
            _HELD_TAKEN.pop(lineage, None)
        _HELD_LOCK.notify_all()


def _put_held(token: tuple, held: _Held, lineage: str) -> None:
    """``held``, the layout of ``token``, back as its lineage's newest (the others of the
    lineage dropped), those of other lineages used longest ago dropped while all cost more
    than :data:`LAYOUT_MEMORY` (round 39)."""
    held.extra.pop("summary before", None)
    with _HELD_LOCK:
        held.lineage = lineage
        for other in [t for t, entry in _HELD.items() if entry.lineage == lineage or t == token]:
            del _HELD[other]
        _HELD[token] = held
        total = sum(entry.cost for entry in _HELD.values())
        while len(_HELD) > 1 and total > LAYOUT_MEMORY:
            total -= _HELD.popitem(last=False)[1].cost
    _let_go_held(lineage)


def held_layout(data: bytes, options: dict, state: tuple, *, changed_pages: bool = False
                ) -> "tuple[_Held, str, float]":
    """The layout of ``data`` for a ``("held", ...)`` ``state`` (the section's notes), taken
    out of this process's layouts or made -- from the one found -- and how (as
    :func:`_take_held` says; ``"same"`` made nothing) with the seconds it took.  The caller
    puts it back (:func:`_put_held`) or, failing, lets the lineage go (:func:`_let_go_held`)."""
    now, kept, lineage = state[1:4]
    found, how = _take_held(now, kept, lineage, state[4] if len(state) > 4 else True)
    if how == "same":
        assert found is not None
        return found, how, 0.0
    try:
        previous = found.laid[0] if found is not None else None
        convert_options = docx2svg.ConvertOptions(**options)
        began = time.perf_counter()
        layout, _, fonts = docx2svg.lay_out_docx(data, convert_options, previous=previous,
                                                 measures=_MEASURES, changed_pages=changed_pages)
        seconds = time.perf_counter() - began
    except BaseException:
        _let_go_held(lineage)
        raise
    made = _Held((layout, fonts, list(convert_options.warnings)))
    summary = found.extra.get("summary") if found is not None else None
    if summary is not None:
        # The summary of the layout this one was made from, for this one's (round 39:
        # :class:`LayoutSummary` digests only what was laid out again).  Dropped when the
        # task ends (:func:`_put_held`), so it never keeps the earlier layout's pages.
        made.extra["summary before"] = summary
    return made, how, seconds


def made_facts(held: _Held, how: str, seconds: float) -> dict:
    """What a task says of the layout it used: how it was found or made (:func:`_take_held`),
    the seconds it took, and -- laid out from another -- ``relaid`` and ``relaid_cause``."""
    layout = held.laid[0]
    again = how in ("kept", "lineage")
    return {"made": how, "seconds": round(seconds, 3),
            "relaid": getattr(layout, "relaid", None) if again else None,
            "relaid_cause": getattr(layout, "relaid_cause", None) if again else None}


def lay_out_bytes(data: bytes, options: dict, state: Any = None) -> tuple:
    """docx2svg's layout of a document's bytes, kept for the next version: ``(laid_out,
    facts)``, ``laid_out`` the pickle of ``(layout, fonts, warnings)`` (:func:`_light`) and
    ``facts`` -- from an earlier version's layout -- ``relaid`` (why it was laid out whole,
    docx2svg's :data:`~docx2svg.relayout.FALLBACKS` code, or ``None``) and ``relaid_cause``;
    ``changed`` is ``None`` (round 37: telling the pages whose drawing changed draws them,
    every page after an edit that adds one where a footer shows the page count, and
    nothing reads it here -- a check compares layouts itself, ``DocumentLayout.compare``).

    ``state``: ``None``; a ``laid_out`` of these very bytes and options, handed back (not
    laid out again: the same pickle comes back); or ``("previous", laid_out)``, the layout
    of an earlier version, which the document is laid out again from -- with what its flow
    measured where this process still holds it (:data:`_RECENT`), else with what this process
    measured of the same document before (:data:`_MEASURES`, round 37), else measured again.  What :func:`convert` runs -- through the
    document's :attr:`~docx_agent.Document.converter`, which may run it in another process
    (an agent tool layer's worker pool, with a deadline)."""
    if held_state(state):
        # Round 38: from the layout this process holds; the caller needs it whole.
        held, how, seconds = held_layout(data, options, state)
        try:
            layout, fonts, warnings = held.laid
            blob = dumps((_light(layout), fonts, warnings))
        except BaseException:
            _let_go_held(state[3])
            raise
        _put_held(state[1], held, state[3])
        return blob, {"changed": None, **made_facts(held, how, seconds)}
    if isinstance(state, (bytes, bytearray)):
        return bytes(state), {"changed": None, "relaid": None, "relaid_cause": None}
    previous = None
    if isinstance(state, tuple) and len(state) == 2 and state[0] == "previous":
        # Taken, not handed back: the new layout goes on with its fonts and measurements.
        previous = _taken(state[1])[0]
    convert_options = docx2svg.ConvertOptions(**options)
    # No ``changed_pages``: nothing here reads them, and telling them draws pages.
    layout, _, fonts = docx2svg.lay_out_docx(data, convert_options, previous=previous, measures=_MEASURES,
                                             changed_pages=False)
    laid = (layout, fonts, list(convert_options.warnings))
    blob = dumps((_light(layout), fonts, laid[2]))
    _remember(blob, laid)
    if previous is None:
        return blob, {"changed": None, "relaid": None, "relaid_cause": None}
    return blob, {"changed": None, "relaid": layout.relaid, "relaid_cause": layout.relaid_cause}


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


class Owner:
    """One owner of a document's layouts for the checks that wait for one (lane O, step 2):
    at most one layout of the document in flight, of its newest committed version; checks
    that come while it is laid out wait, and are all served by the next layout -- again of
    the newest version -- so that the document is laid out once for every check waiting,
    however many writers commit meanwhile.  A check is served by the first layout of a
    version at least the one the document had when it was called (:meth:`serve`).

    Kept on the document (:func:`_keeper`; a read copy shares its document's).  Its lock is
    taken with no other of the module's held, and nothing is waited for under it but its
    own condition."""

    def __init__(self) -> None:
        self._cond = threading.Condition()
        #: The version laid out now (the leader's), or ``None``.
        self.flying: Any = None
        #: The newest version laid out: what a check of an older or equal version reads.
        self.done: Any = None
        #: The newest version whose layout failed: a check of it or older is not served.
        self.failed_at = -1
        #: When a check last asked (``time.monotonic``): a relayout on commit is started
        #: only while checks ask (:meth:`recently_asked`), and once per check asking
        #: (:attr:`armed`, taken by :meth:`take_arm`): a run of edits between two checks
        #: lays out the first one's version at once, not each.
        self.asked_at = 0.0
        self.armed = False
        #: How many layouts this owner made, and how many checks it served (for tests and measurements).
        self.flights = 0
        self.served = 0
        #: Per version served, the checks reading it that have not yet had their facts made
        #: where it is held (:meth:`release`): the next layout -- which takes that worker's
        #: layout of the document on to the next version -- waits for them, at most
        #: :data:`READERS_WAIT` s, so that their facts are made from the layout they read
        #: and not laid out back to it.
        self.readers: dict[int, int] = {}

    def release(self, version: int) -> None:
        """A check served ``version`` has had its facts made (or ended without)."""
        with self._cond:
            left = self.readers.get(version, 0) - 1
            if left > 0:
                self.readers[version] = left
            else:
                self.readers.pop(version, None)
            self._cond.notify_all()

    def _read(self, version: int) -> None:
        self.served += 1
        self.readers[version] = self.readers.get(version, 0) + 1

    def recently_asked(self, seconds: float) -> bool:
        return time.monotonic() - self.asked_at <= seconds

    def take_arm(self, seconds: float) -> bool:
        """Whether a commit now lays its version out at once: a check asked within
        ``seconds`` and no commit since took that (and nothing is laid out now)."""
        with self._cond:
            if not self.armed or self.flying is not None or time.monotonic() - self.asked_at > seconds:
                return False
            self.armed = False
            return True

    def serve(self, version: int, newest: Any, lay: Any, *, ask: bool = True, wait: bool | None = None) -> Any:
        """The version (an object with ``version``: ooxml-edit's ``Committed``) whose layout
        serves a check of ``version``: the newest laid out when it is at least ``version``;
        else the one in flight when it is; else -- once the one in flight is done -- the
        newest committed then (``newest()``), laid out by this caller (``lay(target) ->
        bool``) for every caller waiting meanwhile.  ``None`` where that layout failed: the
        caller reads its own version as before.  ``ask`` false (a relayout on commit, a
        change's work ahead): not counted as a check asking, nor as a reader of what it is
        served; ``wait`` (default: ``ask``) false lays out only when nothing is in flight."""
        if wait is None:
            wait = ask
        with self._cond:
            if ask:
                self.asked_at = time.monotonic()
                self.armed = True
            while True:
                done = self.done
                if done is not None and done.version >= version:
                    if ask:
                        self._read(done.version)
                    return done
                if self.failed_at >= version:
                    return None
                if self.flying is None:
                    target = newest()
                    if target is None or target.version < version:
                        return None
                    self.flying = target
                    # The checks reading the version laid out last have their facts made
                    # first (:attr:`readers`): those tasks are queued before this layout.
                    deadline = time.monotonic() + READERS_WAIT
                    while self.readers and time.monotonic() < deadline:
                        self._cond.wait(max(0.0, deadline - time.monotonic()))
                    break
                if not wait:
                    return None
                self._cond.wait()
        ok = False
        try:
            ok = bool(lay(target))
        finally:
            with self._cond:
                self.flying = None
                self.flights += 1
                if ok and (self.done is None or target.version > self.done.version):
                    self.done = target
                elif not ok:
                    self.failed_at = max(self.failed_at, target.version)
                self._cond.notify_all()
        if not ok:
            return None
        if ask:
            with self._cond:
                self._read(target.version)
        return target


#: Per thread, the owner and version the check running on it reads (:func:`reading`).
_READING = threading.local()


def reading(owner: "Owner", version: int) -> None:
    """The check on this thread reads ``version``, served by ``owner``: it counts as one of
    the version's readers until :func:`read_done`."""
    read_done()
    _READING.held = (owner, version)


def read_done() -> None:
    """The check on this thread has had its facts made, or ends: no longer a reader."""
    held = getattr(_READING, "held", None)
    if held is not None:
        _READING.held = None
        held[0].release(held[1])


#: How long the next layout of a document waits for the checks reading its last one to have
#: their facts made (:attr:`Owner.readers`): a check that never asks is not waited for longer.
READERS_WAIT = 20.0


def own_layout(document: "Document", data: bytes, converter: Any) -> "dict | None":
    """Lay ``data`` -- a committed version of ``document`` -- out where the document's layouts
    are held (:func:`docx_agent.tools.check_worker.check_bytes` through ``converter``, a
    tool layer's engine), kept as the document's last, with its summary and the facts every
    check of it reads (:class:`Owner`'s flight): the seconds the worker spent laying it out
    and reading those facts (``{"layout", "facts"}``; noughts where it was laid out already),
    or ``None`` where it could not be.  Nothing comes back but those small facts; a layout
    that fails is kept as failed, as :func:`check_facts` keeps one."""
    if converter is None or not hasattr(converter, "check"):
        return None
    options = layout_options(document, {})
    key = layout_key(data, options)
    lock = _keeper(document)
    with lock:
        kept = document._layout_kept
        if kept is not None and isinstance(kept[1], (bytes, bytearray)):
            return None
    if known_summary(document, key) is not None:
        return {"layout": 0.0, "facts": 0.0}
    if _kept_failure(document, cache_key(data, options)) is not None:
        return None
    request = dict(_BARE_REQUEST)
    state = kept_layout(document, data, options)
    try:
        result = converter.check(data, dict(options), state, {**request, "summary": True})
    except _PASSED_ON:
        return None
    except Exception as error:  # noqa: BLE001 -- kept as this version's failure, logged once
        _failed(document, cache_key(data, options), error)
        return None
    worker = result.pop("worker", None) or {}
    keep_layout(document, data, options, key)
    note_made(document, result)
    summary_for(document, key, result.pop("summary", None))
    _keep_facts(document, (key, repr(sorted(request.items()))), result, data, options)
    return {"layout": float(worker.get("layout", 0.0)), "facts": float(worker.get("facts", 0.0))}


#: The request of a check that asks nothing but the layout's own facts (:func:`check_facts`).
_BARE_REQUEST = {"pages": None, "ranges": None, "layout": False, "min_size": None, "styles": None,
                 "keeps": None}


def _keep_facts(document: "Document", asked: tuple, result: dict, data: bytes, options: dict) -> None:
    """Keep a check's facts of one version (:func:`check_facts`), and the stats of its layout."""
    with _keeper(document):
        document._check_facts[asked] = result
        while len(document._check_facts) > _FACTS_SIZE:
            document._check_facts.popitem(last=False)
        if result.get("stats") is not None:
            stats = getattr(document, "_layout_stats", None)
            if stats is None:      # round 45: an attribute, so a read copy keeps the document's
                stats = document._layout_stats = OrderedDict()
            stats[cache_key(data, options)] = result["stats"]
            while len(stats) > _CACHE_SIZE:
                stats.popitem(last=False)


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
                document._layout_count = 0
                document._layout_tickets = {}
                # Round 38: the layouts of this document, wherever they are held, and how
                # long laying it out whole took last (the wait for its holder, :func:`affinity`).
                document._layout_lineage = uuid.uuid4().hex
                document._layout_fresh = None
                document._layout_summaries = OrderedDict()
                document._check_facts = OrderedDict()
                document._layout_failures = OrderedDict()
                # Lane O: one owner of the document's layouts for the checks that wait.
                document._layout_owner = Owner()
    return lock


_KEEPERS_LOCK = threading.Lock()


def kept_layout(document: "Document", data: bytes, options: dict, *, may_wait: bool = True) -> Any:
    """The state to hand :func:`lay_out_bytes` (or the render pipeline's layout task) for
    ``data``: ``("held", now, kept, lineage)`` (round 38: the layouts held where they were
    made, by token; the section's notes) -- ``kept`` the token of the document's last
    layout with the same options, or ``None``.  A layout kept as its pickle (a caller's
    own, :func:`keep_layout`) is handed as before: the ``laid_out`` of these bytes when it
    is the one kept, else ``("previous", laid_out)``.  Asked as a layout of ``data``
    starts: the layout is numbered then (:func:`keep_layout`).  ``may_wait`` false: the
    caller holds the document's lock, so the task never waits for another's layout in
    progress (in this process's threads, round 34's rule)."""
    lock = _keeper(document)
    key = layout_key(data, options)
    with lock:
        kept = document._layout_kept
        document._layout_count += 1
        tickets = document._layout_tickets
        tickets.pop(key, None)
        tickets[key] = document._layout_count
        while len(tickets) > _TICKETS:
            del tickets[next(iter(tickets))]
    if kept is not None and isinstance(kept[1], (bytes, bytearray)):
        if kept[0] == key:
            return kept[1]
        return ("previous", kept[1]) if kept[0][1] == key[1] else None
    token = kept[1] if kept is not None and kept[0][1] == key[1] else None
    if not may_wait:      # the caller holds the document's lock: it never waits for a layout
        return ("held", key, token, document._layout_lineage, False)
    return ("held", key, token, document._layout_lineage)


def lineage(document: "Document") -> str:
    """The document's lineage: what its layouts are held under, wherever they were made
    (round 38; a copy laid out as the document is, :func:`twin`, has the document's)."""
    _keeper(document)
    return document._layout_lineage


def affinity(document: "Document") -> "tuple[tuple, float | None]":
    """The pool's affinity key for ``document``'s layout tasks -- its lineage: they go to
    the worker holding its last layout -- and how long a task waits for that worker while it
    is busy: as long as laying the document out whole took last, what going to another
    worker may cost, so a wait never costs more than going elsewhere would (ooxml-edit
    0.90.0's rule); ``None`` -- as long as the deadline allows -- before one was timed (the
    holder is then laying it out whole, which elsewhere would cost as much)."""
    _keeper(document)
    return ("docx-agent layout", document._layout_lineage), document._layout_fresh


def note_made(document: "Document", facts: Any) -> None:
    """Remember how long a layout made whole took (``facts``' ``made`` and ``seconds``,
    :func:`made_facts`): the wait for the document's holder (:func:`affinity`)."""
    if isinstance(facts, dict) and facts.get("made") == "fresh" and facts.get("seconds") is not None:
        _keeper(document)
        document._layout_fresh = float(facts["seconds"])


#: How many layouts started are remembered by number (:func:`kept_layout`).
_TICKETS = 64


def keep_layout(document: "Document", data: bytes, options: dict, laid_out: Any) -> None:
    """Keep ``laid_out``, the layout of ``data`` with ``options``, as the document's last:
    the one its next version is laid out from -- its token (round 38: the layout stays
    where it was made) or its pickle (a caller's own, as before).
    Round 37: unless a layout started later -- of a later version, as many callers lay
    out versions at once -- is kept already: the next layout starts from the newest."""
    lock = _keeper(document)
    key = layout_key(data, options)
    with lock:
        ticket = document._layout_tickets.get(key, document._layout_count)
        kept = document._layout_kept
        if kept is not None and kept[0] != key and kept[2] > ticket:
            return
        document._layout_kept = (key, laid_out, ticket)


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
    failed = _kept_failure(document, key)
    if failed is not None:
        # Round 53 (R107): this version failed already -- at once, and not logged again.
        raise failed
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
                try:
                    flight.result = result = _lay_out(document, data, options, converter,
                                                      may_wait=unlocked is not None and entered is not False)
                except _PASSED_ON:
                    raise
                except Exception as error:
                    raise _failed(document, key, error) from error
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
            try:
                result = _lay_out(document, data, options, converter, may_wait=False)
            except _PASSED_ON:
                raise
            except Exception as error:
                raise _failed(document, key, error) from error
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
        laid = DocumentLayout(owner, raw, warnings)
        laid.key = layout_key(data, options)     # what its summary is kept under (round 38)
        conversion = _Conversion(laid, raw, data, fonts, options, svgs=dict(enumerate(svgs or (), start=1)))
        return keep(document, key, conversion)


@contextmanager
def _held():
    yield


def _lay_out(document: "Document", data: bytes, options: dict, converter: Any = None, *,
             may_wait: bool = False) -> tuple:
    """``(layout, fonts, warnings, svgs)`` of ``data``: through the document's converter --
    its :func:`lay_out_bytes` (``converter.lay_out``), from the document's last layout --
    or, for a converter of the old protocol (``converter(data, options) -> (layout, svgs,
    warnings)``), whole with every page's SVG.  The layout is this caller's own copy (the
    pickle read): the one kept for the next version is never one a caller holds.
    ``converter``: the document's when its bytes were taken; ``may_wait``: the caller let
    the document's lock go (round 38, :func:`kept_layout`)."""
    if converter is not None and not hasattr(converter, "lay_out"):
        raw, svgs, warnings = converter(data, dict(options))
        return raw, None, warnings, list(svgs)
    state = kept_layout(document, data, options, may_wait=may_wait)
    if isinstance(state, bytes):
        blob = state                         # laid out already (a render of this state)
    else:
        run = converter.lay_out if converter is not None else lay_out_bytes
        blob, facts = run(data, dict(options), state)
        # Round 38: the layout stays where it was made; kept by its token.
        keep_layout(document, data, options, state[1] if held_state(state) else blob)
        note_made(document, facts)
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
    # Round 37: the ids the document had, so the copy generates the ones it would.
    pool = document._id_pool()
    return {"stamping": document.stamping, "font_dirs": getattr(document, "font_dirs", None),
            "locale": getattr(document, "locale", None), "converter": document.converter,
            "kept": kept, "ids": (set(pool), dict(pool.cursors), document._ledger.highest_drawing),
            "lineage": document._layout_lineage, "fresh": document._layout_fresh}


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
    if like.get("lineage") is not None:          # round 38: its layouts are the document's
        copy._layout_lineage = like["lineage"]
        copy._layout_fresh = like.get("fresh")
    if like["kept"] is not None:
        copy._layout_count = like["kept"][2]   # numbered on from the document's
    if like.get("ids") is not None:
        values, cursors, highest = like["ids"]
        copy._ledger.pool = _ids.IdPool(values)
        copy._ledger.pool.cursors.update(cursors)
        copy._ledger.highest_drawing = highest
    return copy


#: Round 45: what a read copy (:func:`read_copy`) shares with the document it is of -- the
#: layouts and what is known of them, all keyed by the bytes laid out (equal bytes, equal
#: layout), under the document's one lock over them (:func:`_keeper`).  A read copy's
#: layout is the document's, and the document's next version is laid out from it.  The
#: layout a check last saw (``check``'s ``since_last_check``) is the document's too.
_SHARED_LAYOUT = ("_layout_lock", "_layout_flights", "_layout_kept", "_layout_count",
                  "_layout_tickets", "_layout_lineage", "_layout_fresh", "_layout_summaries",
                  "_check_facts", "_layout_failures", "_layouts", "_layout_stats", "_tool_seen_layout",
                  "_layout_owner",
                  # The tools' caches keyed by content or by version, set whole (a tuple) or
                  # under a lock of their own: a check's copy comparisons, the styles' digest,
                  # a render's coverage, the settings read for a like.
                  "_tool_copy_cache", "_tool_styles_digest", "_tool_render_coverage",
                  "_tool_settings_frame")
#: A document's public attributes a read copy does not take as they are: its parts and
#: history are its own, its layout engine is pointed at per call, and the count of spaces
#: written is an edit's.
_NOT_SETTINGS = ("package", "history", "converter", "typography_spaced")
#: The settings a read copy reads from its document whenever asked, so one the application
#: changes after the version was committed (a locale, a minimum text size) holds for reads
#: of it too, as it did when reads held the document's lock.
_LIVE_SETTINGS = ("font_dirs", "locale", "describe_step", "read_hint", "typography", "stamping",
                  "min_text_size")
_READER_CLASSES: dict = {}


def _live_setting(name: str) -> property:
    def get(self: Any) -> Any:
        live = self.__dict__.get("_live")
        if live is not None and hasattr(live, name):
            return getattr(live, name)
        try:
            return self.__dict__[name]
        except KeyError:
            raise AttributeError(name) from None

    def put(self: Any, value: Any) -> None:
        self.__dict__[name] = value

    return property(get, put)


def _shared_attribute(name: str) -> property:
    def get(self: Any) -> Any:
        live = self.__dict__.get("_live")
        if live is None:
            try:
                return self.__dict__[name]
            except KeyError:
                raise AttributeError(name) from None
        return getattr(live, name)

    def put(self: Any, value: Any) -> None:
        live = self.__dict__.get("_live")
        if live is None:
            self.__dict__[name] = value
        else:
            setattr(live, name, value)

    return property(get, put)


def _reader_class(kind: type) -> type:
    """The read copy's class for documents of ``kind``: ``kind`` with the layout attributes
    (:data:`_SHARED_LAYOUT`) the document's, and its bytes those it was opened from while it
    is the state it was opened as."""
    with _KEEPERS_LOCK:
        found = _READER_CLASSES.get(kind)
        if found is None:
            def to_bytes(self: Any) -> bytes:
                frozen = self.__dict__.get("_frozen")
                if frozen is not None and frozen[0] == self.history.version and frozen[1] == self._state:
                    return frozen[2]
                return kind.to_bytes(self)

            namespace: dict = {name: _shared_attribute(name) for name in _SHARED_LAYOUT}
            namespace.update({name: _live_setting(name) for name in _LIVE_SETTINGS})
            namespace["to_bytes"] = to_bytes
            namespace["__doc__"] = "A read copy of a document at one version (:func:`docx_agent.layout.read_copy`)."
            found = _READER_CLASSES[kind] = type(f"Read{kind.__name__}", (kind,), namespace)
        return found


_DRAFT_CLASSES: dict = {}


def _draft_class(kind: type) -> type:
    """A draft's class for documents of ``kind`` (round 45, lane B): ``kind`` with the
    settings read from the document it is of (:data:`_LIVE_SETTINGS`), as a read copy's,
    and :class:`docx_agent.blocks.DraftOps` recording what a call does on it.  Its layouts
    are its own."""
    from .blocks import DraftOps

    with _KEEPERS_LOCK:
        found = _DRAFT_CLASSES.get(kind)
        if found is None:
            namespace: dict = {name: _live_setting(name) for name in _LIVE_SETTINGS}
            namespace["__doc__"] = "A draft of a document at one version (:mod:`docx_agent.blocks`)."
            found = _DRAFT_CLASSES[kind] = type(f"Draft{kind.__name__}", (DraftOps, kind), namespace)
        return found


def read_copy(document: "Document"):
    """ooxml-edit's ``DocumentFormat.read_copy`` (round 45): what a copy of ``document`` at
    this version needs -- its ids (the ledger of every id it had, the renames, the version
    numbers: what :meth:`Document.id_state` holds) and its settings -- taken now, under the document's lock; and
    ``make(data)``, which opens the copy from this version's bytes with no lock held.  The
    copy answers as the document did: the same ids, version and refusals (its ids), the
    same settings (font folders, locale, stamping, how a gone id is described), and the
    document's own layouts (:data:`_SHARED_LAYOUT`: a check of the copy lays out where the
    document's layouts are, from its last one, and the document's next layout starts from
    the copy's).  Never edited: its tools only read."""
    _keeper(document)
    # Its ids as they are now: the ledger copied (what id_state would write and the copy's
    # open read back, without either), the renames, and the version numbers.
    ledger = _ids.copy_ledger(document._ledger)
    aliases = dict(document._aliases)
    version, last = document.history.version, document.history.last_version
    settings = {name: value for name, value in vars(document).items()
                if not name.startswith("_") and name not in _NOT_SETTINGS + _LIVE_SETTINGS}
    kind = type(document)
    # A changing call's layout engine, still pointed at for its checks, goes to the copy
    # its checks run on once the lock is let go (ooxml-edit's After): the document's own is
    # let go now, as the checks would have (docx_agent.tools._base.checks).
    engine = document.__dict__.pop("_checks_engine", None)
    if document.converter is engine:
        document.converter = None

    def make(data: bytes) -> "Document":
        copy = _reader_class(kind).open(data, stamping=document.stamping)
        copy.__dict__["_live"] = document
        copy.history.continue_from(version, last)
        copy._ledger = ledger
        copy._aliases.update(aliases)
        for name, value in settings.items():
            setattr(copy, name, value)
        copy.converter = None
        copy.__dict__["_handed_engine"] = engine
        copy.__dict__["_frozen"] = (copy.history.version, copy._state, data)
        return copy

    def draft(data: bytes) -> "Document":
        """Round 45 (lane B): an *editable* copy of this version (:mod:`docx_agent.blocks`):
        the same ids, version numbers and settings as the read copy, its ledger its own,
        its layouts its own (a draft is never laid out for the document), its id pool
        recording what is drawn from it."""
        from .blocks import DraftPool

        copy = _draft_class(kind).open(data, stamping=document.stamping)
        copy.__dict__["_live"] = document
        copy.history.continue_from(version, last)
        copy._ledger = _ids.copy_ledger(ledger)
        if copy._ledger.pool is None:
            copy._note_ids()        # as the document's first edit reads them
        pool = DraftPool(copy._ledger.pool)
        pool.cursors.update(copy._ledger.pool.cursors)
        copy._ledger.pool = pool
        copy._aliases.update(aliases)
        for name, value in settings.items():
            setattr(copy, name, value)
        copy.converter = None
        copy.typography_spaced = getattr(document, "typography_spaced", 0)
        return copy

    make.draft = draft  # type: ignore[attr-defined]
    return make


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
            # The copy's last is of the document's bytes now: numbered as started now.
            document._layout_count += 1
            document._layout_kept = (kept[0], kept[1], document._layout_count)


#: How many checks' facts a document keeps, by layout and request (:func:`check_facts`).
_FACTS_SIZE = 16


def check_facts(document: "Document", unlocked: Any, request: dict, *,
                timed: Any = None, also: dict | None = None) -> "tuple[dict, LayoutSummary] | None":
    """``check``'s layout facts of the document now, made in the worker holding its layout
    (round 38, :func:`docx_agent.tools.check_worker.check_bytes`, through the document's
    converter, with the lock let go: ``unlocked``) -- ``(facts, summary)``, the summary that
    of the layout the facts are of (:class:`LayoutSummary`, kept: :func:`summary_for`).  The
    same check of the same version again is answered here.  ``None`` where the converter
    cannot (a converter of the old protocol, a layout kept as its pickle): the caller lays
    the document out here, as before.

    ``also`` (lane O, step 1: one trip per check): a request that ``request`` holds with
    more besides (a check's own, ``request`` that and the pages its numbers' texts are on):
    its facts -- ``request``'s less the keys only ``request`` has -- are kept for it too,
    so the check asking it next finds them without another trip to the worker.

    ``timed`` (round 47, R82: ``timed(part, seconds)``) is told where the time went: the
    worker's ``layout`` and ``facts`` (:func:`~docx_agent.tools.check_worker.check_bytes`)
    and the rest of the wait for its task, ``worker_wait`` (the worker's queue, a layout of
    the document in progress there, the bytes there and the facts back)."""
    import copy

    converter = document.converter
    if converter is None or not hasattr(converter, "check"):
        return None
    options = layout_options(document, {})
    data = document.to_bytes()
    key = layout_key(data, options)
    asked = (key, repr(sorted(request.items())))
    lock = _keeper(document)
    with lock:
        kept = document._layout_kept
        if kept is not None and isinstance(kept[1], (bytes, bytearray)):
            return None
        found = document._check_facts.get(asked)
    summary = known_summary(document, key)
    if found is not None and summary is not None:
        read_done()
        return copy.deepcopy(found), summary
    failed = _kept_failure(document, cache_key(data, options))
    if failed is not None:
        raise failed                          # round 53: this version failed already
    with unlocked():
        state = kept_layout(document, data, options)
        started = time.perf_counter()
        try:
            result = converter.check(data, dict(options), state, {**request, "summary": summary is None})
        except _PASSED_ON:
            raise
        except Exception as error:
            raise _failed(document, cache_key(data, options), error) from error
        finally:
            read_done()             # lane O: its facts are made; the next layout may start
        waited = time.perf_counter() - started
        worker = result.pop("worker", None)
        keep_layout(document, data, options, key)
        note_made(document, result)
        if timed is not None and worker is not None:
            timed("layout", worker["layout"])
            timed("facts", worker["facts"])
            timed("worker_wait", max(0.0, waited - worker["layout"] - worker["facts"]))
    made = result.pop("summary", None)
    summary = summary if summary is not None else summary_for(document, key, made)
    _keep_facts(document, asked, result, data, options)
    if also is not None:
        extra = set(request) - set(also)
        _keep_facts(document, (key, repr(sorted(also.items()))),
                    {name: value for name, value in result.items() if name not in extra}, data, options)
    return copy.deepcopy(result), summary


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
