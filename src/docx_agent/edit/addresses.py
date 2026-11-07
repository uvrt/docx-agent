"""One reading of the addresses the library and the tools take (round 28, W9).

Every id form ``to_markdown`` and ``state`` write names one thing, and every parameter that
takes a place reads it here, so a form one tool takes another takes too, or refuses in the
same words:

* **a block** -- a paragraph (``p:3A1F09C2``, ``p@body/7``, ``header1/p:...``), a table
  (``t:...``), a block-level content control (``cc:...``);
* **inside a table** -- a row (``tr:<paraId>``, ``t:X/r1``) or a cell (``t:X/c1,1``: grid
  row and column from 0, as :meth:`Table.cell` counts them);
* **a section** -- ``s:<paraId>``, ``s:2`` (from 1), ``s:body`` (the last);
* **a bookmark** -- ``bm:Name``, or the plain name;
* **a story** -- ``body``, ``header1``, ``footnotes``...;
* **a text range** -- ``p:A@4:11``, a position ``p:A@5``, ``p:A@3..p:B@8``;
* **a heading's range** -- ``h:p:A``, the heading and every block under it to the next
  heading of its level or higher (:mod:`.outline`), read as the span it is; ``h:`` may
  stand at a span's end (``h:p:A..h:p:B``);
* **a span of blocks** -- ``p:A..t:B``, the blocks from the first through the last, both
  in one container (the body, a cell, a header).  An end may be left open: ``p:A..`` runs
  from ``p:A`` to the last block of its container (the end of the body, for a body
  block), ``..t:B`` from the container's first block through ``t:B``.

What each kind of parameter does with them:

* **a scope** (``within``, ``word_read``'s ``range``, ``find_text``'s ``range``) takes
  every form: the paragraphs it holds (:meth:`AddressOps._scope_paragraphs`);
* **a block, or a span of blocks** (what ``word_delete`` and ``word_move`` act on) takes
  blocks and spans; a cell, row or section is refused with the ids that would do;
* **an anchor** (``after``, ``before``) takes one block; a span or a cell is refused with
  its ends, or its table and paragraphs.

A malformed span (``p:A...p:B``, ``..``, three ends) is refused with the syntax
(:data:`SPAN_SYNTAX`), never as an unknown empty id.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from lxml.etree import _Element as Element

from .ids import ParagraphEntry

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_W_P = _W + "p"

#: How a span is written, in every refusal of one.
SPAN_SYNTAX = ("a span is two block ids, <first>..<last> (p:A..t:B); p:A.. runs to the end of p:A's "
               "container (the body, a cell), ..t:B from its start")

#: A text position or range end: ``<paragraph>@<offset>`` or ``<paragraph>@<start>:<end>``
#: (not a positional block id ``p@body/7``, whose ``@`` a story name follows).
_POSITION = re.compile(r".+@\d+(?::\d+)?")

#: A cell's or row's id below a table: ``<table>/c<row>,<column>``, ``<table>/r<row>``.
_INSIDE_TABLE = re.compile(r"(?P<table>.+)/(?:c(?P<row>\d+),(?P<column>\d+)|r(?P<index>\d+))")


class AddressError(KeyError):
    """An address that is malformed, or names a thing of another kind than the parameter
    takes (a cell where a block is wanted).  A ``KeyError``, as an id the document lacks
    is, so a caller that falls back on a ``KeyError`` still does; the tools say
    ``invalid_arguments``, with ``options``: ids that would do."""

    def __init__(self, message: str, options: list[str] | None = None) -> None:
        super().__init__(message)
        self.options = list(options or [])

    def __str__(self) -> str:
        return str(self.args[0])


def span_ends(address: object) -> tuple[str | None, str | None] | None:
    """The two ends of a span of blocks (``None`` for an open one), or ``None`` when
    ``address`` is not a span of blocks -- one id, or a text range across paragraphs
    (``p:A@3..p:B@8``, which :func:`~.ranges.parse_range` reads).  A malformed span raises
    :class:`AddressError` with the syntax."""
    if not isinstance(address, str) or ".." not in address:
        return None
    first, _, last = address.partition("..")
    first, last = first.strip(), last.strip()
    if "..." in address or ".." in last or not (first or last):
        raise AddressError(f"{address!r} is not a span: {SPAN_SYNTAX}")
    positions = [bool(end and _POSITION.fullmatch(end)) for end in (first, last)]
    if any(positions):
        if all(positions):
            return None
        raise AddressError(f"{address!r} mixes a text position and a block: a range across paragraphs "
                           "is p:A@3..p:B@8, offsets at both ends; " + SPAN_SYNTAX)
    return first or None, last or None


def is_text_range(address: object) -> bool:
    """Whether ``address`` is a text range or position (``p:A@4:11``, ``p:A@5``,
    ``p:A@3..p:B@8``)."""
    if not isinstance(address, str) or "@" not in address:
        return False
    if ".." in address:
        return span_ends(address) is None
    return bool(_POSITION.fullmatch(address))


def inside_table(address: str) -> tuple[str, int, int] | tuple[str, int] | None:
    """``(table, row, column)`` for a cell id, ``(table, row)`` for a ``t:X/r<k>`` row id,
    ``None`` otherwise (``tr:`` rows are found by :meth:`Document.get`)."""
    match = _INSIDE_TABLE.fullmatch(address) if isinstance(address, str) else None
    if match is None or not match.group("table").startswith(("t:", "t@")) and "/t:" not in match.group("table"):
        return None
    if match.group("index") is not None:
        return match.group("table"), int(match.group("index"))
    return match.group("table"), int(match.group("row")), int(match.group("column"))


class AddressOps:
    """The address half of :class:`docx_agent.Document`: what a scope holds, what a span
    is, and the refusals of an address of the wrong kind."""

    # -- what an address holds ---------------------------------------------------------------

    def _scope_paragraphs(self: "Document", address: str) -> tuple[str, list[ParagraphEntry]]:
        """The part and paragraphs (document order, tables' cells included) that ``address``
        holds -- any form the module docstring lists: a story, a block or span (closed or
        open), a row, a cell, a section, a bookmark, a text range, a content control."""
        if not isinstance(address, str) or not address.strip():
            raise AddressError("an address is empty: give a block id (p:..., t:...), a cell (t:X/c1,1), "
                               "a section (s:1), a bookmark or a span (" + SPAN_SYNTAX + ")")
        address = self._expand_heading(address)             # h:p:A, a heading's range (round 36)
        part = self._part_for_story(address)
        if part is not None:
            return part, list(self._index(part).paragraphs)
        if span_ends(address) is not None:
            elements, part = self._blocks_of(address)
            return part, self._entries_in(part, elements)
        if is_text_range(address):
            found = self.range(address)
            return self._entries_named(found.paragraph_ids())
        if address.startswith(("s:", "s@")):
            return self.package.document_part(), self._entries_named(self._section_paragraphs(address))[1]
        if address.startswith("bm:") or not re.search(r"[:@/#]", address):
            return self._entries_named(self.bookmark(address).range().paragraph_ids())
        if address.startswith(("cc:", "cc@")):
            part, sdt = self._control(address)
            inner = self._entries_in(part, [sdt])
            if inner:
                return part, inner
            holder = sdt.getparent()
            while holder is not None and holder.tag != _W_P:
                holder = holder.getparent()
            return part, self._entries_in(part, [holder]) if holder is not None else []
        container = self._table_part_element(address)
        if container is not None:
            return container[0], self._entries_in(*container)
        part, entry = self._resolve(address)
        if isinstance(entry, ParagraphEntry):
            return part, [entry]
        return part, self._entries_in(part, [entry.element])

    def _entries_in(self: "Document", part: str, elements: list[Element]) -> list[ParagraphEntry]:
        nested = {id(p) for element in elements for p in element.iter(_W_P)}
        return [e for e in self._index(part).paragraphs if id(e.element) in nested]

    def _entries_named(self: "Document", identifiers: list[str]) -> tuple[str, list[ParagraphEntry]]:
        found = [self._resolve(identifier) for identifier in identifiers]
        parts = {part for part, _ in found}
        part = parts.pop() if len(parts) == 1 else self.package.document_part()
        return part, [entry for _, entry in found if isinstance(entry, ParagraphEntry)]

    def _table_part_element(self: "Document", address: str) -> tuple[str, list[Element]] | None:
        """The part and element of a row (``tr:``, ``t:X/r1``) or cell (``t:X/c1,1``) id, or
        ``None`` when ``address`` is neither form.  ``KeyError`` for one the document lacks."""
        from .document import Cell, Row

        if not (address.startswith("tr:") or "/tr:" in address or inside_table(address)):
            return None
        found = self.get(address)
        if isinstance(found, Cell):
            part, _ = found._table._entry()
            return part, [found._element]
        if isinstance(found, Row):
            part, entry = found._table._entry()
            return part, [entry.rows[found._index][0]]
        return None

    # -- spans -------------------------------------------------------------------------------

    def _span_elements(self: "Document", first: str | None, last: str | None,
                       movable: tuple[str, ...]) -> tuple[list[Element], str]:
        """The blocks of the span ``first..last`` (either end ``None``: open), in order: both
        ends of one container, the first not after the last (a reversed span is refused,
        by every tool alike)."""
        given = f"{first or ''}..{last or ''}"
        ends = []
        for end in (first, last):
            if end is not None:
                ends.append(self._resolve_block(end))
        part, anchor = ends[0]
        container = anchor.element.getparent()
        siblings = [c for c in container if isinstance(c.tag, str) and c.tag in movable]
        if len(ends) == 2:
            end_part, end = ends[1]
            if end_part != part or end.element.getparent() is not container:
                from .errors import EditError

                raise EditError(f"{given!r}: a span is blocks of one container (the body, one cell); "
                                + SPAN_SYNTAX)
            a, b = siblings.index(anchor.element), siblings.index(end.element)
        elif first is not None:
            a, b = siblings.index(anchor.element), len(siblings) - 1
        else:
            a, b = 0, siblings.index(anchor.element)
        if b < a:
            from .errors import EditError

            raise EditError(f"{given!r} ends before it starts")
        return siblings[a:b + 1], part

    # -- refusals ----------------------------------------------------------------------------

    def _refuse_inside_table(self: "Document", address: str, wanted: str) -> None:
        """An :class:`AddressError` when ``address`` is a row or cell id (``wanted``, say
        "a block", is what the parameter takes): its table and its paragraphs would do."""
        if not isinstance(address, str):
            return
        if not (address.startswith("tr:") or "/tr:" in address or inside_table(address)):
            return
        from .document import Cell, Row

        try:
            found = self.get(address)
        except (KeyError, IndexError):
            return
        if isinstance(found, Cell):
            paragraphs = [p.id for p in found.paragraphs]
            raise AddressError(f"{address} is a cell, not {wanted}: give its table {found._table.id} or a "
                               f"paragraph in it ({', '.join(paragraphs[:3])})", [found._table.id, *paragraphs])
        if isinstance(found, Row):
            raise AddressError(f"{address} is a row, not {wanted}: give its table {found._table.id}",
                               [found._table.id])

    def _unknown(self: "Document", identifier: object) -> KeyError:
        """The error for an id :meth:`_resolve` cannot find as a paragraph or table: one
        that says what the id is when it is something else."""
        if not isinstance(identifier, str) or not identifier.strip():
            return AddressError(f"an id is empty ({identifier!r}): give a block id (p:..., t:...)"
                                + (", or " + SPAN_SYNTAX if identifier == "" else ""))
        deleted = self._deleted(identifier)
        if deleted is not None:
            return deleted
        try:
            ends = span_ends(identifier)
        except AddressError as error:
            return error
        if ends is not None:
            named = [end for end in ends if end]
            return AddressError(f"{identifier} is a span; give one block here (p:A or t:B)", named)
        try:
            self._refuse_inside_table(identifier, "a paragraph or table")
        except AddressError as error:
            return error
        if identifier.startswith("h:"):
            try:
                span = self.heading_span(identifier)
            except AddressError as error:
                return error
            return AddressError(f"{identifier} is a heading's range ({span}); give one block here", span.split(".."))
        if identifier.startswith(("s:", "s@")):
            try:
                self._section_record(identifier)
            except KeyError as error:
                return error
            return AddressError(f"{identifier} is a section, not a paragraph or table: give a block in it")
        holder = None
        try:
            if is_text_range(identifier):
                holder, what = self.range(identifier).start_id, "a text range"
            elif self._part_for_story(identifier) is not None:
                return AddressError(f"{identifier} is a story, not a paragraph or table: give a block in it")
            elif identifier.startswith("bm:") or not re.search(r"[:@/#]", identifier):
                holder, what = self.bookmark(identifier).range().start_id, "a bookmark"
        except KeyError:
            holder = None
        if holder is not None:
            return AddressError(f"{identifier} is {what}, not a paragraph or table: give the paragraph "
                                f"it is in, {holder}", [holder])
        return KeyError(f"no paragraph or table {identifier!r}")
