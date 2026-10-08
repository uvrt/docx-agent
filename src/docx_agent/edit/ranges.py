"""Text ranges and anchors: places in the text, found by offset or by what they say.

A range is ``<paragraph id>@<start>:<end>`` -- character offsets into the paragraph's text
(ROADMAP.md, "Text ranges and anchors": ``w:t`` characters, ``\\t`` for a tab, ``\\v`` for a
line break, ``\\f`` for a page or column break, U+FFFC for a drawing or note reference, a
field's *result*), or ``<paragraph id>@<start>..<paragraph id>@<end>`` across paragraphs.
It is taken in a **view** -- ``current`` (the default: insertions in, deletions out),
``original`` or ``markup`` -- and is valid in that view only; edits take ranges in the
current view and refuse the others.  A range names a *place*: it holds its paragraphs' ids,
which survive edits, but its offsets are what they were when it was taken.

**Finding.**  :meth:`TextOps.find` returns every match as a range; :meth:`TextOps.anchor`
returns exactly one or raises :class:`AmbiguousAnchor` listing the candidates (Adeu's rule:
an ambiguous anchor is refused, never guessed).  Matching normalises nothing unless asked:
``case=False`` and ``whitespace="collapse"`` are explicit; ``regex=True`` takes a pattern;
``across_paragraphs=True`` matches over paragraph ends (spelled ``\\n``).

**Replacing** (:meth:`TextOps.replace`, :meth:`TextRange.replace`) works across run
boundaries: the runs a match spans are edited in place, the replacement takes the
formatting of the match's first character (Word's rule) -- or, with
``keep="characters"``, each character that survives keeps its own -- and text around the
match is not touched.  A match that would cut a field in two, or run from a field's result
into other text, is refused, as is one that would delete a note reference.  A match across
paragraphs is deleted and the replacement inserted: the first paragraph keeps its id and
properties and takes in what follows the match in the last one.
"""

from __future__ import annotations

import bisect
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..oxml.xml import Element, remove
from . import inline as _inline
from . import text as _text
from .ids import ParagraphEntry

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_RANGE = re.compile(r"(?P<start>.+?)@(?P<a>\d+)(?::(?P<b>\d+)|\.\.(?P<end>.+)@(?P<c>\d+))")


class AnchorNotFound(KeyError):
    """No text matches an anchor."""


class AmbiguousAnchor(LookupError):
    """More than one place matches an anchor; ``candidates`` lists them."""

    def __init__(self, text: str, candidates: list["TextRange"]) -> None:
        self.text = text
        self.candidates = candidates
        listed = "; ".join(f"{c.id} {c.context()!r}" for c in candidates[:10])
        more = f" and {len(candidates) - 10} more" if len(candidates) > 10 else ""
        super().__init__(f"{len(candidates)} places match {text!r}: {listed}{more}; "
                         "give occurrence= or within= to choose one")


class TextRange:
    """Characters ``[start, end)`` of one paragraph, or from ``start`` in one paragraph to
    ``end`` in a later one of the same story, in ``view``."""

    def __init__(self, document: "Document", start_id: str, start: int, end_id: str | None = None,
                 end: int | None = None, view: str = "current") -> None:
        if view not in _text.VIEWS:
            raise ValueError(f"no text view {view!r}")
        self._document = document
        self.start_id = start_id
        self.start = start
        self.end_id = end_id if end_id is not None else start_id
        self.end = end if end is not None else start
        self.view = view

    # -- naming ------------------------------------------------------------------------------

    @property
    def id(self) -> str:
        """The range's id: ``<paragraph id>@<start>:<end>``, or
        ``<paragraph id>@<start>..<paragraph id>@<end>`` across paragraphs."""
        start_id = self._current(self.start_id)
        end_id = self._current(self.end_id)
        if start_id == end_id:
            return f"{start_id}@{self.start}:{self.end}"
        return f"{start_id}@{self.start}..{end_id}@{self.end}"

    def _current(self, identifier: str) -> str:
        try:
            return self._document._resolve(identifier)[1].id
        except KeyError:
            return identifier

    @property
    def collapsed(self) -> bool:
        """Whether the range is empty (a place rather than text)."""
        return self.start_id == self.end_id and self.start == self.end

    @property
    def single(self) -> bool:
        """Whether the range is within one paragraph."""
        return self._current(self.start_id) == self._current(self.end_id)

    # -- reading -----------------------------------------------------------------------------

    def paragraph_ids(self) -> list[str]:
        """The ids of the paragraphs the range touches, in order."""
        return [entry.id for entry in self._entries()[1]]

    def _entries(self) -> tuple[str, list[ParagraphEntry]]:
        document = self._document
        part, first = document._resolve(self.start_id)
        part_end, last = document._resolve(self.end_id)
        if not isinstance(first, ParagraphEntry) or not isinstance(last, ParagraphEntry):
            raise KeyError(f"{self.id} does not run between paragraphs")
        if part != part_end:
            raise ValueError(f"{self.id} runs across stories")
        if last.index < first.index:
            raise ValueError(f"{self.id} ends before it starts")
        paragraphs = document._index(part).paragraphs[first.index:last.index + 1]
        return part, paragraphs

    @property
    def text(self) -> str:
        """The range's text in its view; paragraph ends are ``\\n``."""
        _, entries = self._entries()
        texts = [_text.paragraph_text(entry.element, self.view) for entry in entries]
        if len(texts) == 1:
            return texts[0][self.start:self.end]
        return "\n".join([texts[0][self.start:]] + texts[1:-1] + [texts[-1][:self.end]])

    def context(self, width: int = 20) -> str:
        """The range's text with some of what surrounds it, for telling candidates apart."""
        _, entries = self._entries()
        first = _text.paragraph_text(entries[0].element, self.view)
        last = _text.paragraph_text(entries[-1].element, self.view)
        before = first[max(0, self.start - width):self.start]
        after = last[self.end:self.end + width]
        return f"{before}[{self.text}]{after}"

    # -- editing -----------------------------------------------------------------------------

    def replace(self, text: str, *, keep: str = "first") -> "EditResult":
        """Replace the range's text: the replacement takes the formatting of its first
        character (``keep="first"``), or each surviving character keeps its own
        (``keep="characters"``).  ``doc.anchor("24 months").replace("36 months")``.

        Tracked, the whole range is recorded as deleted and the replacement as inserted,
        so take a tight range; :meth:`Paragraph.set_text` records only the words that
        changed."""
        return self._document._replace_ranges([(self, text)], keep=keep)

    def delete(self, *, collapse_space: bool = False) -> "EditResult":
        """Delete the range's text (tracked when tracking): ``doc.anchor("draft ").delete()``.

        ``collapse_space=True`` also deletes the space the deletion would leave doubled or
        stranded -- the one after the range when spaces are on both sides (a sentence in the
        middle of a paragraph), the one after it at the paragraph's start and the one
        before it at its end -- so ``doc.anchor("The supplier waives ... months.").delete(
        collapse_space=True)`` leaves "notice. The customer", not "notice.  The customer".
        Tracked, the space is part of the one deletion."""
        target = self._collapsed_space() if collapse_space else self
        return self._document._replace_ranges([(target, "")], keep="first")

    def _collapsed_space(self) -> "TextRange":
        """This range, grown by the one space its deletion would leave doubled or stranded."""
        if not self.single or self.collapsed:
            return self
        text = _text.paragraph_text(self._entries()[1][0].element, self.view)
        start, end = self.start, self.end
        if text[start:end].strip() != text[start:end]:
            return self   # the range already holds its own space
        before = text[start - 1] if start > 0 else None
        after = text[end] if end < len(text) else None
        if after == " " and (before is None or before == " "):
            end += 1
        elif before == " " and after is None:
            start -= 1
        elif before == " " and after in (".", ",", ";", ":", "!", "?", ")"):
            start -= 1
        else:
            return self
        return TextRange(self._document, self.start_id, start, self.end_id, end, view=self.view)

    def insert_before(self, text: str) -> "EditResult":
        """Insert ``text`` at the range's start, in the formatting there."""
        return TextRange(self._document, self.start_id, self.start, view=self.view).replace(text)

    def insert_after(self, text: str) -> "EditResult":
        """Insert ``text`` at the range's end, in the formatting there:
        ``doc.anchor("EUR 12,500").insert_after(", excluding VAT")``."""
        return TextRange(self._document, self.end_id, self.end, view=self.view).replace(text)

    def format(self, **values) -> "EditResult":
        """Direct run formatting on the range (``bold=True``, ``color="C00000"``...): runs
        are split at its edges, and merged again where they come out alike."""
        return self._document.format_range(self, **values)

    def set_style(self, name: str | None) -> "EditResult":
        """Apply a character style by name to the range (``None`` removes it)."""
        return self._document.set_character_style(self, name)

    def clear_formatting(self) -> "EditResult":
        """Remove the direct run formatting of the range (its character style stays)."""
        return self._document.clear_formatting(self)

    def add_hyperlink(self, url: str | None = None, *, anchor: str | None = None,
                      tooltip: str | None = None, style: bool = True) -> "EditResult":
        """Make the range a hyperlink to ``url``, or to the bookmark ``anchor``."""
        return self._document.add_hyperlink(self, url, anchor=anchor, tooltip=tooltip, style=style)

    def add_bookmark(self, name: str) -> "EditResult":
        """Bookmark the range as ``name``: ``doc.anchor("net revenue").add_bookmark("Revenue")``."""
        return self._document.add_bookmark(self, name)

    def insert_picture(self, image, **options) -> "EditResult":
        """An inline picture at the range's start."""
        return self._document.insert_picture(self, image, **options)

    def __eq__(self, other: object) -> bool:
        return isinstance(other, TextRange) and other.id == self.id and other.view == self.view

    def __hash__(self) -> int:
        return hash((self.id, self.view))

    def __repr__(self) -> str:
        return f"<TextRange {self.id} {self.text[:40]!r}>"


@dataclass
class _Match:
    range: TextRange
    replacement: str


def parse_range(document: "Document", identifier: str, view: str = "current") -> TextRange:
    """The range an id names: ``<paragraph>@<start>:<end>``, ``<paragraph>@<offset>`` (an
    empty range, a position) or ``<paragraph>@<start>..<paragraph>@<end>``."""
    match = _RANGE.fullmatch(identifier)
    if match is None:
        position = re.fullmatch(r"(.+?)@(\d+)", identifier)
        if position is None:
            raise KeyError(f"{identifier!r} is not a range")
        start_id, offset = position.group(1), int(position.group(2))
        document._resolve(start_id)
        return TextRange(document, start_id, offset, view=view)
    start_id = match.group("start")
    if match.group("b") is not None:
        return TextRange(document, start_id, int(match.group("a")), start_id, int(match.group("b")), view)
    return TextRange(document, start_id, int(match.group("a")), match.group("end"), int(match.group("c")), view)


def pattern_for(text: str, *, regex: bool = False, case: bool = True, whitespace: str | None = None) -> re.Pattern:
    """The compiled pattern a find matches with."""
    if whitespace not in (None, "collapse"):
        raise ValueError("whitespace must be None or 'collapse'")
    if regex:
        source = text
    elif whitespace == "collapse":
        source = r"\s+".join(re.escape(token) for token in text.split())
        if not source:
            raise ValueError("nothing to find")
    else:
        source = re.escape(text)
    if not text:
        raise ValueError("nothing to find")
    return re.compile(source, 0 if case else re.IGNORECASE)


class TextOps:
    """The text half of :class:`docx_agent.Document`: ranges, find, anchors, replace."""

    def range(self: "Document", identifier: str, *, view: str = "current") -> TextRange:
        """The range an id names (``p:1A2B3C4D@4:11``, ``p:...@3..p:...@8``, ``p:...@5``)."""
        return parse_range(self, identifier, view)

    def _scope(self: "Document", within: str | None) -> list[tuple[str, list[ParagraphEntry]]]:
        """The stories' paragraphs a find looks in: everything, a story, a table or one
        paragraph."""
        if within is None:
            return [(part, self._index(part).paragraphs) for part in self._parts()]
        part = self._part_for_story(within)
        if part is not None:
            return [(part, self._index(part).paragraphs)]
        part, entry = self._resolve(within)
        if isinstance(entry, ParagraphEntry):
            return [(part, [entry])]
        nested = set(map(id, entry.element.iter(_W + "p")))
        return [(part, [e for e in self._index(part).paragraphs if id(e.element) in nested])]

    def find(self: "Document", text: str, *, within: str | None = None, regex: bool = False,
             case: bool = True, whitespace: str | None = None, across_paragraphs: bool = False,
             view: str = "current") -> list[TextRange]:
        """Every place ``text`` occurs, in document order (the body, then headers, footers,
        notes and comments), as ranges in ``view``."""
        return [m.range for m in self._matches(text, None, within=within, regex=regex, case=case,
                                               whitespace=whitespace, across_paragraphs=across_paragraphs,
                                               view=view)]

    def _matches(self: "Document", text: str, replacement: str | None, *, within, regex, case, whitespace,
                 across_paragraphs, view) -> list[_Match]:
        pattern = pattern_for(text, regex=regex, case=case, whitespace=whitespace)
        out: list[_Match] = []
        for part, entries in self._scope(within):
            if across_paragraphs:
                texts = [_text.paragraph_text(e.element, view) for e in entries]
                joined = "\n".join(texts)
                starts = []
                position = 0
                for t in texts:
                    starts.append(position)
                    position += len(t) + 1
                for match in pattern.finditer(joined):
                    if match.end() == match.start():
                        continue
                    a = _locate(starts, match.start())
                    b = _locate(starts, match.end())
                    found = TextRange(self, entries[a].id, match.start() - starts[a], entries[b].id,
                                      match.end() - starts[b], view)
                    out.append(_Match(found, match.expand(replacement) if replacement is not None and regex
                                      else replacement or ""))
            else:
                for entry in entries:
                    paragraph_text = _text.paragraph_text(entry.element, view)
                    for match in pattern.finditer(paragraph_text):
                        if match.end() == match.start():
                            continue
                        found = TextRange(self, entry.id, match.start(), entry.id, match.end(), view)
                        out.append(_Match(found, match.expand(replacement) if replacement is not None and regex
                                          else replacement or ""))
        return out

    def anchor(self: "Document", text: str, *, occurrence: int | None = None, within: str | None = None,
               regex: bool = False, case: bool = True, whitespace: str | None = None,
               across_paragraphs: bool = False, view: str = "current") -> TextRange:
        """The one place ``text`` occurs.  :class:`AnchorNotFound` if it occurs nowhere,
        :class:`AmbiguousAnchor` (listing every candidate with its id) if it occurs more than
        once and ``occurrence`` (0-based, in document order) does not choose."""
        found = self.find(text, within=within, regex=regex, case=case, whitespace=whitespace,
                          across_paragraphs=across_paragraphs, view=view)
        if occurrence is not None:
            if not -len(found) <= occurrence < len(found):
                raise AnchorNotFound(f"{text!r} occurs {len(found)} times; no occurrence {occurrence}")
            return found[occurrence]
        if not found:
            raise AnchorNotFound(f"{text!r} occurs nowhere" + (f" in {within}" if within else ""))
        if len(found) > 1:
            raise AmbiguousAnchor(text, found)
        return found[0]

    def replace(self: "Document", text: str, replacement: str, *, within: str | None = None,
                regex: bool = False, case: bool = True, whitespace: str | None = None,
                across_paragraphs: bool = False, count: int | None = None,
                keep: str = "first") -> "EditResult":
        """Replace every match of ``text`` (the first ``count``) with ``replacement``
        (``\\1``, ``\\g<name>`` expanded when ``regex``), as one undo step.  Refused as a
        whole, before anything changes, if any match cuts a field or would delete a note
        reference.  ``result.count`` is the number replaced."""
        matches = self._matches(text, replacement, within=within, regex=regex, case=case,
                                whitespace=whitespace, across_paragraphs=across_paragraphs, view="current")
        if count is not None:
            matches = matches[:count]
        result = self._replace_ranges([(m.range, m.replacement) for m in matches], keep=keep)
        result.count = len(matches)
        return result

    def insert_text(self: "Document", at: str | TextRange, text: str) -> "EditResult":
        """Insert ``text`` at a position (a range id or range; a range's start), taking the
        formatting of the character before it, else the one after it."""
        where = at if isinstance(at, TextRange) else self.range(at)
        return TextRange(self, where.start_id, where.start, view=where.view).replace(text)

    def delete_text(self: "Document", what: str | TextRange, *, collapse_space: bool = False) -> "EditResult":
        """Delete a range's text (a range or its id): :meth:`TextRange.delete`."""
        where = what if isinstance(what, TextRange) else self.range(what)
        return where.delete(collapse_space=collapse_space)

    # -- the edit ----------------------------------------------------------------------------

    def _replace_ranges(self: "Document", pairs: list[tuple[TextRange, str]], *, keep: str) -> "EditResult":
        from .document import EditError, EditResult

        if not pairs:
            return EditResult(None, changed=False, count=0)
        for found, replacement in pairs:
            if found.view != "current":
                raise EditError(f"{found.id} was taken in the {found.view} view; edits take the current view")
            _text.check_text(replacement)
        # Resolve and check everything before changing anything.
        plans = []
        for found, replacement in pairs:
            part, entries = found._entries()
            first_text = _text.paragraph_text(entries[0].element)
            last_text = _text.paragraph_text(entries[-1].element)
            if found.start > len(first_text) or found.end > len(last_text):
                raise EditError(f"{found.id} is outside its paragraphs' text")
            try:
                if len(entries) == 1:
                    _inline.check_span(_text.atoms(entries[0].element), found.start, found.end)
                else:
                    _check_across(entries, found.start, found.end)
            except _inline.SpanError as error:
                raise EditError(f"{found.id}: {error}") from None
            plans.append((part, entries, found.start, found.end, replacement))
        order = {part: k for k, part in enumerate(self._parts())}
        plans.sort(key=lambda p: (order[p[0]], p[1][0].index, p[2]), reverse=True)
        for earlier, later in zip(plans[1:], plans):
            if earlier[0] == later[0] and (earlier[1][-1].index, earlier[3]) > (later[1][0].index, later[2]):
                raise EditError("the ranges overlap")

        touched: dict[str, list[ParagraphEntry]] = {}
        removed: list[str] = []
        tracking = self._active_tracking()
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp

        with self._edit():
            renames = self._prepare_many([(part, entries) for part, entries, *_ in plans])
            stamp = Stamp(self, tracking) if tracking is not None else None
            for part, entries, start, end, replacement in plans:
                if len(entries) == 1:
                    if stamp is not None:
                        released = _track.replace_span(entries[0].element, start, end, replacement, stamp, part,
                                                       keep=keep)
                    else:
                        released = _inline.replace_span(entries[0].element, start, end, replacement, keep=keep)
                    touched.setdefault(part, []).append(entries[0])
                elif stamp is not None:
                    released = _replace_across_tracked(entries, start, end, replacement, keep, stamp, part)
                    touched.setdefault(part, []).extend([entries[0], entries[-1]])
                else:
                    released = _replace_across(entries, start, end, replacement, keep=keep)
                    removed += [entry.id for entry in entries[:-1]]
                    touched.setdefault(part, []).append(entries[-1])
                self._release(part, released)
                self.package.mark_dirty(part)
            if stamp is not None:
                stamp.finish()
            for part, entries in touched.items():
                for entry in entries:
                    if entry.element.getparent() is not None:
                        self._renew_text_id(part, entry)
        removed = [renames.get(old, old) for old in removed]
        self._aliases = {old: new for old, new in self._aliases.items() if new not in removed}
        first = plans[-1][1][0] if len(plans[-1][1]) == 1 else plans[-1][1][-1]
        return EditResult(renames.get(first.id, first.id), renamed=renames, removed=removed,
                          count=len(plans))


def _locate(starts: list[int], offset: int) -> int:
    """The paragraph a joined-text offset is in.  An end offset at a paragraph's start is
    offset 0 of that paragraph: the match took the paragraph end before it."""
    return bisect.bisect_right(starts, offset) - 1


def _check_across(entries: list[ParagraphEntry], start: int, end: int) -> None:
    """A range across paragraphs is deleted and its first and last paragraphs joined: refuse
    what that would break."""
    parent = entries[0].element.getparent()
    elements = [entry.element for entry in entries]
    for element in elements:
        if element.getparent() is not parent:
            raise _inline.SpanError("the range runs across a table, a cell or a content control")
    siblings = [child for child in parent if isinstance(child.tag, str) and child.tag in (_W + "p", _W + "tbl", _W + "sdt")]
    positions = [next(k for k, s in enumerate(siblings) if s is e) for e in elements]
    if positions != list(range(positions[0], positions[0] + len(positions))):
        raise _inline.SpanError("the range runs across a table or a content control")
    for element in elements:
        properties = element.find(_W + "pPr")
        if properties is not None and properties.find(_W + "sectPr") is not None:
            raise _inline.SpanError("the range runs across a section break")
    _inline.check_span(_text.atoms(elements[0]), start, len(_text.atoms(elements[0])))
    _inline.check_span(_text.atoms(elements[-1]), 0, end)
    from .document import _fields_balanced

    for element in elements[1:-1]:
        for tag in ("footnoteReference", "endnoteReference", "commentReference"):
            if element.find(".//" + _W + tag) is not None:
                raise _inline.SpanError(f"the range holds a paragraph with a {tag}")
        if not _fields_balanced(element):
            raise _inline.SpanError("the range holds part of a field that spans paragraphs")
    for element in elements[:-1]:
        if not _fields_balanced(element):
            raise _inline.SpanError("the range holds part of a field that spans paragraphs")


def _replace_across(entries: list[ParagraphEntry], start: int, end: int, replacement: str, *,
                    keep: str) -> list[str]:
    """Delete from ``start`` in the first paragraph to ``end`` in the last, and join them as
    Word joins two paragraphs whose mark is deleted (measured, tracked and untracked): what
    is left of the first -- and ``replacement`` after it -- moves to the start of the last,
    which keeps its own element and id and takes the first one's properties."""
    from .document import _rehome_markers

    first, last = entries[0].element, entries[-1].element
    released: list[str] = []
    released += _inline.replace_span(last, 0, end, "")
    for entry in entries[1:-1]:
        released += _inline.relationship_ids(entry.element)
        _rehome_markers(entry.element)
        remove(entry.element)
    first_length = len(_text.atoms(first))
    released += _inline.replace_span(first, start, first_length, replacement, keep=keep)
    # What is left of the first paragraph moves to the start of the last, which takes its
    # properties (the section break, which neither has, aside).
    first_properties = first.find(_W + "pPr")
    last_properties = last.find(_W + "pPr")
    at = 0
    if last_properties is not None:
        remove(last_properties)
    if first_properties is not None:
        last.insert(0, first_properties)
        at = 1
    for child in list(first):
        if child is first_properties:
            continue
        last.insert(at, child)
        at += 1
    remove(first)
    released += _inline.prune(last)
    return released


def _replace_across_tracked(entries: list[ParagraphEntry], start: int, end: int, replacement: str, keep: str,
                            stamp, part: str) -> list[str]:
    """:func:`_replace_across` as Word tracks it: the text deleted (and the replacement
    inserted where the first paragraph's text ends), every paragraph mark in between
    deleted, and the last paragraph given the first one's properties, recorded in a
    ``w:pPrChange`` -- so accepting makes the untracked join, and rejecting the original."""
    from ..revisions import track as _track

    import copy

    first, last = entries[0].element, entries[-1].element
    released: list[str] = []
    released += _track.delete_span(last, 0, end, stamp, part)
    first_length = len(_text.atoms(first))
    released += _track.replace_span(first, start, first_length, replacement, stamp, part, keep=keep)
    properties = first.find(_W + "pPr")
    first_properties = copy.deepcopy(properties) if properties is not None else _text._make("w:pPr")
    for entry in [entries[0]] + entries[1:-1]:
        if entry is not entries[0]:
            released += _track.delete_content(entry.element, stamp, part)
        _track.delete_mark(entry.element, stamp, part)
    _track.take_properties(last, first_properties, stamp, part)
    return released
