"""The outline: the body's headings and the blocks each owns (round 36, R32), and the
guards that keep a heading's content from changing owner unseen (R33).

**A heading** is a paragraph that is a block of the body itself (not one in a table's
cell, a text box, a content control, a header, footer or note) with an outline level 1-9,
as Word reads it for a table of contents and as :meth:`Document.blocks` gives it: a
built-in heading style's number (``heading 1`` to ``heading 9``); else the level the
paragraph states (``w:outlineLvl``; 9, body text, is none); else the level its style's
``basedOn`` chain declares.  A Title, or any other style, is a heading only when the
document gives it an outline level.

**What a heading owns** -- its range -- is the heading and every top-level block after it
up to the next heading of its level or a higher one (a lower number), or the body's end:
tables, content-control blocks, deeper headings and section breaks included.  ``h:<heading>``
(``h:p:3A1F09C2``) names that range wherever a span of blocks is taken: it is the span
``<heading>..<its last block>`` (just the heading when nothing follows it).  A span's end
may be one too: ``h:p:A..h:p:B`` runs from ``p:A`` through the end of ``p:B``'s range.

:meth:`OutlineOps.outline` gives each heading with its level, text, range, last block,
block count and the section (``s:N``, counting from 1) it is in.

The guards (the tools apply them; :func:`content_left` and :func:`moves` say what they
see): deleting a heading whose range is not deleted with it leaves its content under the
heading before; changing a heading's level hands blocks from one heading to another.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Iterable

from .addresses import AddressError, span_ends

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_P, _TBL, _SDT, _CUSTOM = _W + "p", _W + "tbl", _W + "sdt", _W + "customXml"
_PPR, _SECTPR, _BODY = _W + "pPr", _W + "sectPr", _W + "body"

#: The prefix of a heading's range.
PREFIX = "h:"

#: How a heading's range is written, in every refusal of one.
SYNTAX = ("h:<heading> is the heading and every block under it, to the next heading of its level or "
          "higher (h:p:A; h:p:A..h:p:B to the end of p:B's)")

#: Headings offered in a refusal.
OFFERED = 20


@dataclass(frozen=True)
class Block:
    """One top-level block of the body as the outline reads it."""

    id: str | None
    level: int | None
    section: int


@dataclass(frozen=True)
class Heading:
    """A heading and its range: ``blocks`` of the body from ``id`` through ``last``."""

    id: str
    level: int
    text: str
    last: str
    blocks: int
    section: int
    last_section: int
    #: Its place among the body's blocks (index of the heading, and one past its last).
    start: int = 0
    end: int = 0

    @property
    def range(self) -> str:
        return PREFIX + self.id

    @property
    def span(self) -> str:
        """The range as a span of blocks: ``p:A..p:Z``, or ``p:A`` alone."""
        return self.id if self.last == self.id else f"{self.id}..{self.last}"

    def as_dict(self) -> dict[str, Any]:
        out: dict[str, Any] = {"id": self.id, "level": self.level, "text": self.text, "range": self.range,
                               "last": self.last, "blocks": self.blocks, "section": f"s:{self.section}"}
        if self.last_section != self.section:
            out["to_section"] = f"s:{self.last_section}"
        return out


def body_blocks(document: "Document") -> list[Block]:
    """Every top-level block of the body, in order: its id (``None`` for custom XML,
    which nothing addresses), its outline level and the section it is in (from 1; a
    paragraph ending a section is in the section it ends)."""
    from .document import _OutlineLevels

    part = document.package.document_part()
    index = document._index(part)
    levels = _OutlineLevels(document)
    body = document.package.tree(part).find(_BODY)
    out: list[Block] = []
    section = 1
    for node in body if body is not None else ():
        if not isinstance(node.tag, str) or node.tag not in (_P, _TBL, _SDT, _CUSTOM):
            continue
        identifier = level = None
        breaks = 0
        if node.tag in (_P, _TBL):
            entry = index.entry_for(node)
            identifier = entry.id if entry is not None else None
            if node.tag == _P:
                level = levels.effective(node)
                breaks = 1 if node.find(f"{_PPR}/{_SECTPR}") is not None else 0
        else:
            if node.tag == _SDT:
                identifier = document._block_id(part, node)
            # A break in a block-level control or custom XML ends a section too.
            breaks = sum(1 for s in node.iter(_SECTPR) if s.getparent().tag == _PPR
                         and s.getparent().getparent().tag == _P)
        out.append(Block(identifier, level, section))
        section += breaks
    return out


def headings(document: "Document", blocks: list[Block] | None = None, *, text: bool = True) -> list[Heading]:
    """The body's headings with their ranges, in order (module docstring)."""
    blocks = body_blocks(document) if blocks is None else blocks
    out = []
    for k, block in enumerate(blocks):
        if block.level is None or block.id is None:
            continue
        end = next((j for j in range(k + 1, len(blocks))
                    if blocks[j].level is not None and blocks[j].level <= block.level), len(blocks))
        last = next((j for j in range(end - 1, k, -1) if blocks[j].id is not None), k)
        words = document.paragraph(block.id).text if text else ""
        out.append(Heading(block.id, block.level, words, blocks[last].id or block.id, end - k, block.section,
                           blocks[end - 1].section, k, end))
    return out


def offered(document: "Document", limit: int = OFFERED) -> list[str]:
    """The body's first headings as ranges, each with its text: ``h:p:A  Results``."""
    out = []
    for heading in headings(document)[:limit]:
        out.append(f"{heading.range}  {' '.join(heading.text.split())[:60]}".rstrip())
    return out


class OutlineOps:
    """The outline half of :class:`docx_agent.Document`."""

    def outline(self: "Document") -> list[dict[str, Any]]:
        """The body's headings in order, each with what it owns (round 36, R32)::

            {"id": "p:3A1F09C2", "level": 2, "text": "Results", "range": "h:p:3A1F09C2",
             "last": "t:0B2C4D6E", "blocks": 7, "section": "s:2"}

        ``range`` names the heading and every block under it -- to the next heading of its
        level or higher -- wherever a span is taken (``to_markdown``, ``state``,
        ``move_blocks``, the tools' ``range`` and ``target``); ``last`` is its last block,
        ``blocks`` how many it holds, the heading too; ``section`` the section the heading
        is in, counting from 1 (``to_section`` the one its last block is in, where that is
        another).  Which paragraphs are headings: :mod:`docx_agent.edit.outline`."""
        return [heading.as_dict() for heading in headings(self)]

    def heading_span(self: "Document", identifier: str) -> str:
        """A heading's range as a span of blocks (``p:A..t:Z``, or ``p:A`` when nothing is
        under it): ``identifier`` is the heading (``p:...``) or its range (``h:p:...``).
        :class:`~docx_agent.edit.addresses.AddressError` for a paragraph that is no
        heading of the body, with the body's headings as ``options``."""
        return self._heading(identifier).span

    def _heading(self: "Document", identifier: str) -> Heading:
        given = identifier
        if isinstance(identifier, str) and identifier.startswith(PREFIX):
            identifier = identifier[len(PREFIX):]
        if not isinstance(identifier, str) or not identifier.strip():
            raise AddressError(f"{given!r} names no heading: {SYNTAX}", offered(self))
        if identifier.startswith(PREFIX) or span_ends(identifier) is not None:
            raise AddressError(f"{given!r} is not a heading's range: {SYNTAX}", offered(self))
        try:
            part, entry = self._resolve(identifier)
        except KeyError as error:
            raise AddressError(f"{given!r}: {error.args[0] if error.args else error}; {SYNTAX}",
                               offered(self)) from None
        from .ids import ParagraphEntry

        body = self.package.document_part()
        if part != body or not isinstance(entry, ParagraphEntry) or entry.element.getparent().tag != _BODY:
            where = "is not a paragraph" if not isinstance(entry, ParagraphEntry) else (
                "is not a block of the body itself (a heading in a table, text box, content control, header "
                "or note owns no blocks)")
            raise AddressError(f"{given} {where}: {SYNTAX}", offered(self))
        found = next((h for h in headings(self, text=False) if h.id == entry.id), None)
        if found is None:
            style = self.paragraph(entry.id).style_name
            raise AddressError(f"{entry.id} is not a heading (style {style!r}, no outline level): {SYNTAX}",
                               offered(self))
        return found

    def _expand_heading(self: "Document", address: Any) -> Any:
        """``address`` with a heading's range (``h:p:A``, or ``h:`` at a span's end) as the
        span of blocks it is; any other address unchanged."""
        if not isinstance(address, str) or PREFIX not in address:
            return address
        if address.startswith(PREFIX) and ".." not in address:
            return self.heading_span(address)
        if ".." not in address:
            return address
        first, _, last = address.partition("..")
        first, last = first.strip(), last.strip()
        if not (first.startswith(PREFIX) or last.startswith(PREFIX)):
            return address
        if first.startswith(PREFIX):
            first = self._heading(first).id
        if last.startswith(PREFIX):
            last = self._heading(last).last
        return f"{first}..{last}"


# -- the guards (R33) -------------------------------------------------------------------------


@dataclass(frozen=True)
class Left:
    """A heading deleted while blocks it owns stay: ``kept`` of them, the first ``first``,
    going under ``under`` (``None``: no heading before it)."""

    heading: Heading
    kept: int
    first: str
    under: Heading | None


def content_left(document: "Document", deleted: Iterable[str]) -> list[Left]:
    """The headings among ``deleted`` (top-level body block ids) whose range is not all
    deleted with them, and where what stays would go: under the nearest heading before it
    that stays (any level: its blocks join the one before)."""
    gone = set(deleted)
    blocks = body_blocks(document)
    found = headings(document, blocks)
    out = []
    for heading in found:
        if heading.id not in gone:
            continue
        kept = [b.id for b in blocks[heading.start + 1:heading.end] if b.id is not None and b.id not in gone]
        if not kept:
            continue
        under = next((h for h in reversed(found) if h.start < heading.start and h.id not in gone), None)
        out.append(Left(heading, len(kept), kept[0], under))
    return out


def label(heading: Heading | None) -> str:
    """``p:A 'Results'``, or the start of the document for ``None``."""
    if heading is None:
        return "no heading (the document's start)"
    text = " ".join(heading.text.split())
    return f"{heading.id} {text[:40]!r}" if text else heading.id


def moves(before: list[Heading], after: list[Heading], renamed: dict[str, str] | None = None,
          limit: int = 5) -> list[str]:
    """What a change of headings' levels did to the outline, one line per heading whose
    level changed, became or stopped being a heading: where the blocks it gained or lost
    now go.  ``before`` and ``after`` are :func:`headings` of the body before and after
    the change; ``renamed`` maps an id the change renamed to its new one."""
    renamed = renamed or {}
    before = [_renamed(h, renamed) for h in before]
    was = {h.id: h for h in before}
    now = {h.id: h for h in after}
    out = []
    changed = [h for h in after if h.id not in was or was[h.id].level != h.level]
    changed += [h for h in before if h.id not in now]
    for heading in changed:
        if len(out) >= limit:
            out.append(f"and {len(changed) - limit} more heading(s) changed level: describe shows the outline")
            break
        old, new = was.get(heading.id), now.get(heading.id)
        if new is None:
            under = _owner(after, heading.start)
            out.append(f"{label(heading)} is no longer a heading (was level {heading.level}): its "
                       f"{heading.blocks - 1} block(s) now go under {label(under)}" if heading.blocks > 1 else
                       f"{label(heading)} is no longer a heading (was level {heading.level})")
            continue
        if old is None:
            taken = new.blocks - 1
            before_owner = _owner(before, new.start)
            out.append(f"{label(new)} is now a level {new.level} heading: it takes the {taken} block(s) after "
                       f"it, to {new.last}, from under {label(before_owner)}" if taken else
                       f"{label(new)} is now a level {new.level} heading, with nothing under it")
            continue
        verb = "demoted" if new.level > old.level else "promoted"
        if new.end < old.end:
            lost = [h for h in after if new.end <= h.start < old.end and h.level <= new.level]
            names = ", ".join(label(h) for h in lost[:3]) + (f" and {len(lost) - 3} more" if len(lost) > 3 else "")
            out.append(f"{label(new)} {verb} {old.level} -> {new.level}: it now ends at {new.last}; "
                       f"{names or str(old.end - new.end) + ' block(s)'} after it now go under "
                       f"{label(_owner(after, new.end, before=new.start))}")
        elif new.end > old.end:
            gained = [h for h in after if old.end <= h.start < new.end and h.level == new.level + 1] or \
                     [h for h in after if old.end <= h.start < new.end]
            names = ", ".join(label(h) for h in gained[:3]) + (
                f" and {len(gained) - 3} more" if len(gained) > 3 else "")
            out.append(f"{label(new)} {verb} {old.level} -> {new.level}: it now runs to {new.last}, taking "
                       f"{names or str(new.end - old.end) + ' block(s)'} from under "
                       f"{label(_owner(before, old.end, before=old.start))}")
        else:
            parent_was = _parent(before, old)
            parent_now = _parent(after, new)
            out.append(f"{label(new)} {verb} {old.level} -> {new.level}: its blocks stay its own; it is now "
                       f"under {label(parent_now)}" + ("" if (parent_was and parent_was.id) ==
                                                       (parent_now and parent_now.id)
                                                       else f" (was {label(parent_was)})"))
    return out


def _owner(found: list[Heading], at: int, before: int | None = None) -> Heading | None:
    """The deepest heading whose range holds block ``at`` (one starting before ``before``,
    when given: not the heading that changed itself)."""
    holding = [h for h in found if h.start <= at < h.end and (before is None or h.start < before)]
    return max(holding, key=lambda h: h.start) if holding else None


def _parent(found: list[Heading], heading: Heading) -> Heading | None:
    holding = [h for h in found if h.start < heading.start < h.end and h.level < heading.level]
    return max(holding, key=lambda h: h.start) if holding else None


def _renamed(heading: Heading, renamed: dict[str, str]) -> Heading:
    if heading.id not in renamed and heading.last not in renamed:
        return heading
    from dataclasses import replace

    return replace(heading, id=renamed.get(heading.id, heading.id), last=renamed.get(heading.last, heading.last))
