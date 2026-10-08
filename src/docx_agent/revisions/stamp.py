"""Who made a tracked change, and when: the attributes every revision record carries.

A :class:`Tracking` is the mode the edit primitives consult (``Document.tracking(...)``, or
``track=True`` on one edit).  A :class:`Stamp` is one edit's view of it: it allocates
revision ids -- one above the largest annotation id in use in any story, as bookmarks and
comments are allocated, so every id in the document stays unique -- and writes the
attributes as Word 16.106 writes them (``tools/e3_probe.py``):

* ``w:ins``, ``w:del``, ``w:moveFrom``, ``w:moveTo`` and the property changes:
  ``w:id``, ``w:author``, ``w:date``, ``w16du:dateUtc``;
* the move range markers: ``w:id``, ``w:author``, ``w:date``, ``w:name``;
* ``w:tblGridChange``: ``w:id`` alone.

Dates are minutes, ISO 8601 with a ``Z``.  Word writes its local time in ``w:date`` (with
the ``Z`` all the same) and the true UTC in ``w16du:dateUtc``; docx-agent writes the UTC in
both, so its output does not depend on the machine's time zone.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..edit import ids as _ids
from ..oxml.xml import Element, make

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W16DU = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
DATE_UTC = "{%s}dateUtc" % W16DU

#: The author an edit tracked with ``track=True`` and no ``tracking()`` context is made by.
DEFAULT_AUTHOR = "docx-agent"


def now() -> str:
    """The current minute, UTC, as Word writes a revision's date."""
    moment = _dt.datetime.now(_dt.timezone.utc).replace(second=0, microsecond=0)
    return moment.strftime("%Y-%m-%dT%H:%M:00Z")


def normalise_date(value: "str | _dt.datetime | None") -> str:
    if value is None:
        return now()
    if isinstance(value, _dt.datetime):
        if value.tzinfo is None:
            value = value.replace(tzinfo=_dt.timezone.utc)
        return value.astimezone(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    text = str(value)
    try:
        parsed = _dt.datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        raise ValueError(f"{value!r} is not an ISO 8601 date") from None
    return normalise_date(parsed)


@dataclass(frozen=True)
class Tracking:
    """The tracking mode: the author and date every revision an edit writes carries."""

    author: str
    date: str
    initials: str | None = None

    @classmethod
    def make(cls, author: str | None = None, date=None, initials: str | None = None) -> "Tracking":
        """A tracking mode for ``author`` (the default author if none), ``date`` (now) and
        ``initials`` (from the author's name)."""
        author = author or DEFAULT_AUTHOR
        return cls(author, normalise_date(date), initials or _initials(author))


def _initials(author: str) -> str:
    return "".join(word[0] for word in author.split() if word[:1].isalnum()).upper()[:9] or author[:1].upper()


class Stamp:
    """One edit's revision records: their ids and attributes."""

    def __init__(self, document: "Document", tracking: Tracking) -> None:
        self.document = document
        self.tracking = tracking
        self._next = document._next_annotation_id()
        self._first = self._next
        #: Parts the edit wrote revisions into: their roots declare ``w16du``.
        self.parts: set[str] = set()

    @property
    def author(self) -> str:
        return self.tracking.author

    @property
    def date(self) -> str:
        return self.tracking.date

    def next_id(self) -> str:
        value = self._next
        self._next += 1
        return str(value)

    def make(self, tag: str, part: str | None = None, *, utc: bool = True, name: str | None = None) -> Element:
        """A revision record (``w:ins``, ``w:rPrChange``...) with its id, author and date."""
        element = make(tag)
        element.set(_W + "id", self.next_id())
        if tag == "w:tblGridChange":
            return element
        element.set(_W + "author", self.author)
        element.set(_W + "date", self.date)
        if name is not None:
            element.set(_W + "name", name)
        elif utc:
            element.set(DATE_UTC, self.date)
            if part is not None:
                self.parts.add(part)
        return element

    def same(self, element: Element) -> bool:
        """Whether a revision record is this edit's author's (Word removes one's own
        insertion outright when one deletes it: measured)."""
        return element.get(_W + "author") == self.author

    def finish(self) -> None:
        """Declare ``w16du`` on every part written into (and list it as ignorable), and list
        the author in ``people.xml`` once anything was recorded."""
        if self._next != self._first:
            self.document._ensure_person(self.author)
        # A change undone (formatting set, then cleared) leaves no record: then no person.
        self.document._tidy_people({self.author})
        self._first = self._next
        for part in self.parts:
            root = self.document.package.tree(part)
            if root is not None:
                _ids.ensure_w14(root, extra={"w16du": W16DU})
                self.document.package.mark_dirty(part)
        self.parts.clear()
