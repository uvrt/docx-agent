"""Reflow feedback: where docx2svg laid each paragraph out, in this API's ids.

docx2svg lays out the bytes docx-agent serialises, so every ``Line.path`` it reports names
exactly one element of the live tree.  :class:`DocumentLayout` resolves those paths with
docx2svg's own :func:`docx2svg.paths.resolve_path` (the rules ``data-docx-path`` is counted
by, exported for this) and keys every line by the id the paragraph has in the document.

The document is laid out **once** per state: :func:`convert` calls
:func:`docx2svg.convert_docx`, which returns the layout of every page and the pages' SVG
together, and both ``layout()`` and ``render_svg()`` read that one conversion.

What it answers (ROADMAP.md, "Reflow feedback"):

* ``where(id)`` -- one :class:`Placement` per page and text column the paragraph (or table)
  is drawn in, in points from the page's top; ``page_of(id)`` -- the page, or ``(first, last)``;
* ``compare(before)`` -- which pages differ from an earlier layout, by a signature of what
  each page draws (glyphs, positions, rules, pictures; not the SVG text), and which blocks
  moved to another page;
* **stops, carried through**: where docx2svg stops (a feature it cannot lay out yet), a block
  past the stop is :class:`Unknown`, never guessed, and pages past either layout's stop are
  reported as unknown by ``compare``.

Ids are compared through the document's aliases, so a layout taken before an edit that
stamped or re-issued paragraphs still matches the layout after it.
"""

from __future__ import annotations

import dataclasses
import hashlib
import re
from dataclasses import dataclass, field
from fractions import Fraction
from typing import TYPE_CHECKING

import docx2svg
from docx2svg.paths import path_part, resolve_path

from .edit.ids import ParagraphEntry, TableEntry

if TYPE_CHECKING:  # pragma: no cover
    from .edit.document import Document

#: docx2svg's device pixels (Word's 1/300-inch export grid) per point.
PX_PER_PT = Fraction(300, 72)

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


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
class Unknown:
    """Where a block is cannot be said: ``reason`` is the layout's stop (``table``,
    ``drawing``...) for a block past it, with the page it stopped on, or ``not-drawn`` for a
    block docx2svg draws nothing of (a hidden paragraph, the continuation of a vertically
    merged cell, a note it does not place)."""

    reason: str
    after_page: int | None = None

    def __bool__(self) -> bool:
        return False


def coverage_facts(coverage, *, at: str | None = None, stopped: "Stop | None" = None) -> dict:
    """What a caller needs to tell "laid out and nothing wrong" from "could not lay it all
    out": docx2svg's coverage (:class:`docx2svg.coverage.Coverage`, or its ``as_dict()``),
    compact.  ``complete`` is true only when every block was laid out, every header,
    footer and text box drawn and every face found; ``substituted_fonts`` names the faces
    laid out with an open substitute (``(approximate)`` where it is not metric compatible);
    ``stop`` is where the layout stopped, ``at`` the block's id when known.

    ``status`` tells the three apart: ``complete`` (laid out, nothing approximated),
    ``approximate`` (complete, but ``approximations`` lists places docx2svg laid out by a
    rule it has not measured -- a floating drawing in a table cell placed as no probe
    measured, say -- instead of stopping: what they touch may be off) and ``partial``
    (not complete)."""
    if coverage is None:  # a docx2svg without coverage: what the stop alone says
        facts = {"complete": stopped is None, "status": "complete" if stopped is None else "partial"}
        if stopped is not None:
            facts["stop"] = {"page": stopped.page, "reason": stopped.reason, "at": stopped.at}
        return facts
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
    return facts


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
        for page in self.pages:
            lines = []
            for line in page.text_lines():
                key = (_part_of(document, line.path, line.part), line.path)
                if key not in resolved:
                    resolved[key] = _paragraph_id(document, *key)
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
        #: layout covers (``None`` from a docx2svg that predates it).
        self.coverage = getattr(layout, "coverage", None)

    def coverage_facts(self) -> dict:
        """:func:`coverage_facts` of this layout, its stop named by this API's id."""
        return coverage_facts(self.coverage, at=self.stopped.at if self.stopped else None, stopped=self.stopped)

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


def _paragraph_id(document: "Document", part: str, path: str) -> str | None:
    element = resolve_path(document.package.tree(part), path)
    if element is None or element.tag != _W + "p":
        return None
    try:
        entry = document._index(part).entry_for(element)
    except KeyError:
        return None
    return entry.id if entry is not None else None


def _block_id(document: "Document", part: str, path: str) -> str | None:
    element = resolve_path(document.package.tree(part), path)
    if element is None:
        return None
    entry = document._index(part).entry_for(element)
    return entry.id if entry is not None else None


# -- entry points ----------------------------------------------------------------------------

_CACHE_SIZE = 8


@dataclass
class _Conversion:
    """One state of the document laid out once: its layout in this API's ids, and the SVG
    of every page, tagged with ids as they are asked for."""

    layout: DocumentLayout
    svgs: list[str]
    tagged: dict[int, str] = field(default_factory=dict)


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


def convert_bytes(data: bytes, options: dict) -> tuple:
    """docx2svg's conversion of a document's bytes: ``(layout, svgs, warnings)``, every one
    of them picklable.  What :func:`convert` runs unless the document names a
    :attr:`~docx_agent.Document.converter` -- which may run this very function in another
    process (an agent tool layer's worker pool, with a deadline) and return its answer."""
    convert_options = docx2svg.ConvertOptions(**options)
    result = docx2svg.convert_docx(data, convert_options)
    return result.layout, list(result.svgs), list(convert_options.warnings)


def convert(document: "Document", **options) -> _Conversion:
    """The document now, laid out once by :func:`docx2svg.convert_docx` -- every page, so
    the layout answers for all of them -- and cached by its bytes and the options.  Both
    :func:`lay_out` and :func:`render_svg` read it, so a render and a layout of the same
    state cost one layout.  docx2svg finds the faces Word draws itself, Office's cloud-font
    cache among them (Aptos Display, a new document's heading face)."""
    if "pages" in options:
        raise TypeError("convert() lays out every page; select pages when rendering")
    data = document.to_bytes()
    key = cache_key(data, options)
    cached = document._layouts.get(key)
    if cached is not None:
        return cached
    raw, svgs, warnings = (document.converter or convert_bytes)(data, dict(options))
    conversion = _Conversion(DocumentLayout(document, raw, warnings), list(svgs))
    document._layouts[key] = conversion
    while len(document._layouts) > _CACHE_SIZE:
        document._layouts.pop(next(iter(document._layouts)))
    return conversion


def lay_out(document: "Document", **options) -> DocumentLayout:
    """The document's layout now (:func:`convert`'s)."""
    return convert(document, **options).layout


_GROUP = re.compile(r'<g data-docx-path="([^"]*)"((?: data-docx-id="[^"]*")?)'
                    r'((?: data-docx-story="[^"]*" data-docx-part="([^"]*)")?)>')


def render_svg(document: "Document", pages: list[int] | None = None, **options) -> list[str]:
    """docx2svg's SVG of the pages (1-based; every page by default), every paragraph group
    carrying ``data-docx-agent-id``; from the same conversion as :func:`lay_out`."""
    conversion = convert(document, **options)
    count = len(conversion.svgs)
    selected = list(range(1, count + 1)) if pages is None else [n for n in pages if 1 <= n <= count]
    out = []
    for number in selected:
        if number not in conversion.tagged:
            conversion.tagged[number] = _tag(document, conversion.layout, conversion.svgs[number - 1])
        out.append(conversion.tagged[number])
    return out


def _tag(document: "Document", layout: DocumentLayout, svg: str) -> str:
    """``data-docx-agent-id`` on every paragraph group of one page's SVG."""

    def tag(match: re.Match) -> str:
        path, part = match.group(1), match.group(4)
        identifier = layout.id_for_path(_part_of(document, path, part), path)
        if identifier is None:
            return match.group(0)
        return match.group(0)[:-1] + f' data-docx-agent-id="{_escape(identifier)}">'

    return _GROUP.sub(tag, svg)


def _escape(value: str) -> str:
    return value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
