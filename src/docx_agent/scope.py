"""What ``check`` and ``render`` are scoped to by ``range`` (round 35, R18): blocks of the body.

A reviewer looks at a few sections of a long document: ``check(range=[...])`` reports the
facts about those blocks only, measured on the pages they fall on, and ``render(range=[...])``
draws those pages.  A range is any address ``word_read`` takes for its ``range`` -- a block
(``p:...``, ``t:...``), a span (``p:A..t:B``, ``p:A..`` to the body's end), a cell or row, a
section (``s:2``), a bookmark, a text range, a content control -- and several may be given.

:class:`Scope` is what they hold, read from the document once (under its lock, before any
layout):

* ``paragraphs`` -- the ids of every paragraph in them, those in tables' cells too;
* ``tables`` -- the ids of the tables holding any of those paragraphs (a cell's table: its
  table facts -- a header row not repeated, its columns' widths -- are about the cell too);
* ``drawings`` -- the ids of the drawings anchored in those paragraphs;
* ``roots`` -- the elements a problem is looked for in: each top-level block every paragraph
  of which is in scope, and the paragraphs of a block that is only partly (a cell's).

A fact is about the scope when its subject -- ``para``, ``paras``, ``ids``, ``table``,
``drawing`` or ``blocks``, whichever the entry has -- is one of those; an entry with none of
them (the pages the layout could not see whole) is about the pages, and kept on the scope's
pages.  :func:`pages_of` gives each range's pages from a layout of the document.

Only the body's blocks: a range in a header, footer or note is refused (its pages are every
page, or the note's), with the body's headings as the ways in.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:  # pragma: no cover
    from .edit.document import Document
    from .layout import DocumentLayout

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_P, _TBL = _W + "p", _W + "tbl"

#: The keys of a layout fact's entry that name what it is about (:mod:`.layout_facts`).
SUBJECT_KEYS = ("para", "paras", "ids", "table", "drawing", "blocks")

#: Ranges in one call at most: a reviewer's sections, not every block one by one.
MAX_RANGES = 20

#: Headings offered as ranges when one names nothing.
HEADINGS_OFFERED = 20


@dataclass
class Range:
    """One range as given and what it holds."""

    address: str
    paragraphs: list[str]
    blocks: list[str]
    #: The element of each paragraph, in order (for pages and problems).
    elements: list[Any] = field(default_factory=list, repr=False)


class Scope:
    """The body's blocks one or more ranges name (the module says what each set holds)."""

    def __init__(self, document: "Document", ranges: list[str]) -> None:
        main = document.package.document_part()
        found = []
        for number, address in enumerate(ranges):
            part, entries = _paragraphs_of(document, address, number)
            if part != main:
                raise _refused(document, f"range[{number}] {address!r} is in {document._story_of(part)}, "
                               "not the body: a range names the body's blocks")
            found.append((address, entries))
        self._build(document, found)

    @classmethod
    def of_ids(cls, document: "Document", spec: "list[tuple[str, list[str]]]") -> "Scope":
        """The scope :meth:`spec` describes, on ``document`` -- a copy opened from the bytes
        of the state the spec was read from (round 38: a check's facts made in the worker
        holding the layout): the same paragraphs by their ids, so the same scope."""
        scope = cls.__new__(cls)
        part = document.package.document_part()
        index = document._index(part)
        by_id = {entry.id: entry for entry in index.paragraphs}
        scope._build(document, [(address, [by_id[identifier] for identifier in ids])
                                for address, ids in spec])
        return scope

    def spec(self) -> "list[tuple[str, list[str]]]":
        """Each range's address and its paragraphs' ids: what :meth:`of_ids` rebuilds it from."""
        return [(found.address, list(found.paragraphs)) for found in self.ranges]

    def _build(self, document: "Document", found: "list[tuple[str, list]]") -> None:
        from .edit.document import _top_block

        self.document = document
        main = document.package.document_part()
        index = document._index(main)
        self.ranges: list[Range] = []
        self.paragraphs: set[str] = set()
        self.tables: set[str] = set()
        self.drawings: set[str] = set()
        held: dict[int, Any] = {}               # id(paragraph element) -> element
        tops: dict[int, Any] = {}               # id(top-level block) -> element
        for address, entries in found:
            blocks: list[str] = []
            mine: set[int] = set()
            for entry in entries:
                held[id(entry.element)] = entry.element
                self.paragraphs.add(entry.id)
                top = _top_block(entry.element)
                if top is not None and id(top) not in mine:
                    mine.add(id(top))
                    tops[id(top)] = top
                    block = index.entry_for(top)
                    if block is not None:
                        blocks.append(block.id)
                for ancestor in entry.element.iterancestors(_TBL):
                    table = index.entry_for(ancestor)
                    if table is not None:
                        self.tables.add(table.id)
            self.ranges.append(Range(address, [e.id for e in entries], blocks,
                                     [e.element for e in entries]))
        # Each top-level block wholly in scope is looked in whole; else its paragraphs.
        self.roots: list[Any] = []
        for top in tops.values():
            inside = [p for p in top.iter(_P)] if top.tag != _P else [top]
            if all(id(p) in held for p in inside):
                self.roots.append(top)
            else:
                self.roots += [p for p in inside if id(p) in held]
        if held:
            for part, identifier, frame in document._all_drawings():
                if part != main:
                    continue
                if any(id(a) in held for a in frame.iterancestors(_P)):
                    self.drawings.add(identifier)

    # -- what a fact is about ------------------------------------------------------------------

    def subjects(self) -> set[str]:
        return self.paragraphs | self.tables | self.drawings

    def holds(self, identifier: Any) -> bool:
        return isinstance(identifier, str) and identifier in self.subjects()

    def about(self, entry: Any) -> bool:
        """Whether a layout fact's entry is about the scope (the module's rule)."""
        if not isinstance(entry, dict):
            return True
        named = False
        for key in SUBJECT_KEYS:
            value = entry.get(key)
            if value is None:
                continue
            named = True
            values = value if isinstance(value, list) else [value]
            if any(self.holds(v) for v in values):
                return True
        return not named

    # -- pages -------------------------------------------------------------------------------

    def pages(self, layout: "DocumentLayout") -> list[list[int]]:
        """Each range's pages in ``layout``, in order (``[]`` for one nothing of which is
        drawn: hidden, or past where the layout stopped)."""
        out = []
        for found in self.ranges:
            pages: set[int] = set()
            for identifier in found.paragraphs:
                pages.update(layout.pages_of(identifier))
            if not pages:
                for identifier in found.blocks:
                    pages.update(layout.pages_of(identifier))
            out.append(sorted(pages))
        return out


@dataclass(frozen=True)
class ReadScope:
    """What a read of some ranges covered (round 47): the body part and its top-level blocks
    there -- ooxml-edit's ``Call.read_scope``, for whether another call's change meanwhile
    touched it (:meth:`docx_agent.blocks.Blocks.touches`)."""

    part: str
    blocks: frozenset
    #: Whether what was read depends on the layout (a check's pages, a render's page): then
    #: a change before the range's last block -- which can move it to other pages -- is one
    #: to it too.
    layout: bool = False


def read_scope(scope: "Scope", *, layout: bool = False) -> ReadScope:
    """The :class:`ReadScope` of ``scope``."""
    return ReadScope(scope.document.package.document_part(),
                     frozenset(block for found in scope.ranges for block in found.blocks), layout)


def ranges_of(value: Any) -> list[str]:
    """``range`` as a list of addresses: one string is one range (rule 1: a list's one
    item, said either way); ``[]``, ``""`` and blanks are no range (a strict decoder's
    placeholders)."""
    if value is None:
        return []
    items = [value] if isinstance(value, str) else list(value)
    out = []
    for item in items:
        if isinstance(item, str) and item.strip():
            if item.strip() not in out:
                out.append(item.strip())
    return out


def _paragraphs_of(document: "Document", address: str, number: int) -> tuple[str, list]:
    """The part and paragraphs ``address`` holds (``word_read``'s resolver), a refusal
    naming the range and the ways in where it names nothing."""
    from .edit.addresses import AddressError
    from .edit.errors import EditError
    from .edit.ids import DeletedId

    try:
        part, entries = document._scope_paragraphs(address)
    except AddressError as error:
        raise _refused(document, f"range[{number}] {address!r}: {_message(error)}", code="invalid_arguments",
                       options=list(error.options or [])) from None
    except DeletedId as error:                   # round 37: gone, and by which change
        raise _refused(document, f"range[{number}] {_message(error)}", code="not_found",
                       options=[error.replaced_by] if error.replaced_by else None) from None
    except (KeyError, IndexError) as error:
        raise _refused(document, f"range[{number}] {address!r} names nothing in the document ({_message(error)})",
                       code="not_found") from None
    except EditError as error:
        raise _refused(document, f"range[{number}] {address!r}: {error}") from None
    if not entries:
        raise _refused(document, f"range[{number}] {address!r} holds no paragraph", code="not_found")
    return part, entries


def _message(error: BaseException) -> str:
    text = error.args[0] if error.args else str(error)
    return str(text)


def _refused(document: "Document", message: str, *, code: str = "invalid_arguments",
             options: list[str] | None = None) -> Exception:
    """A refusal with the way out (rule 2): the body's headings as ranges, each section
    from its heading (``p:A..`` reads to the end; a span ``p:A..p:B`` stops before the
    next)."""
    from ooxml_edit.tools import ToolError

    if not options:
        options = headings_as_ranges(document)
    return ToolError(code, message + "; give a block, a span p:A..t:B, a cell, a section s:2 or a "
                     "heading's range h:p:A (word_read's range)", field="range",
                     valid_options=options[:50])


def headings_as_ranges(document: "Document") -> list[str]:
    """The body's first headings, each as its range with its text: ``h:p:A  Heading``
    (round 36: ``h:`` names a heading with everything under it)."""
    from .edit.outline import offered

    return offered(document, HEADINGS_OFFERED)
