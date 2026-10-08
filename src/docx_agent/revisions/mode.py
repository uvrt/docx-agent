"""Tracking mode: what the edit primitives consult, and Word's own ``w:trackRevisions``.

``doc.tracking(author=..., date=..., initials=...)`` is a context in which every edit is
written as a revision (Aspose's ``StartTrackRevisions``, not a second API); ``track=True``
on one edit tracks that edit (as the context's author, else :data:`DEFAULT_AUTHOR`), and
``track=False`` keeps one edit untracked inside the context.  ``doc.word_tracks_changes`` is
a different thing: the ``w:trackRevisions`` setting, which tells *Word* to track the next
person's edits; the two are set independently.  (``doc.track_revisions`` and
``set_track_revisions`` are its deprecated names.)
"""

from __future__ import annotations

import functools
import warnings
from contextlib import contextmanager
from typing import TYPE_CHECKING, Iterator

from ..oxml.xml import insert_in_order, make, remove
from .stamp import Tracking

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _document_of(owner) -> "Document":
    from ..edit.document import Document

    if isinstance(owner, Document):
        return owner
    for name in ("_document", "document"):
        found = getattr(owner, name, None)
        if isinstance(found, Document):
            return found
    paragraph = getattr(owner, "_paragraph", None) or getattr(owner, "paragraph", None)
    if paragraph is not None:
        return _document_of(paragraph)
    table = getattr(owner, "_table", None) or getattr(owner, "table", None)
    if table is not None:
        return _document_of(table)
    raise TypeError(f"{owner!r} has no document")


def trackable(method):
    """Give an edit method a ``track`` keyword: ``True`` tracks it, ``False`` keeps it
    untracked, ``None`` (the default) follows ``Document.tracking``."""

    @functools.wraps(method)
    def wrapper(self, *args, track: "bool | None" = None, **kwargs):
        if track is None:
            return _with_document(method(self, *args, **kwargs), self)
        document = _document_of(self)
        if track is True:
            mode = document._tracking or Tracking.make()
        elif track is False:
            mode = False
        else:
            raise TypeError("track is True, False or None")
        document._track_stack.append(mode)
        try:
            return _with_document(method(self, *args, **kwargs), self)
        finally:
            document._track_stack.pop()

    return wrapper


def _with_document(result, owner):
    """An edit's :class:`EditResult` told its document, for ``result.object``."""
    if getattr(result, "document", 0) is None:
        try:
            result.document = _document_of(owner)
        except TypeError:
            pass
    return result


class TrackingOps:
    """The tracking mode and the ``w:trackRevisions`` setting, on :class:`docx_agent.Document`."""

    @contextmanager
    def tracking(self: "Document", author: str | None = None, date=None,
                 initials: str | None = None) -> Iterator[Tracking]:
        """Every edit inside is written as a tracked change (a revision) by ``author``
        (``date``: ISO 8601 or a datetime; ``None`` is now, to the minute, UTC), whatever
        Word's own switch says.  Nests: the inner context wins inside it.  One edit can opt
        in or out with ``track=True``/``track=False``::

            with doc.tracking(author="Claude"):
                doc.anchor("24 months").replace("36 months")
            doc.revisions(author="Claude")          # the deletion and the insertion

        This is not :attr:`word_tracks_changes`, Word's Track Changes switch for a
        person's later edits in Word; the two are independent."""
        previous = self._tracking
        self._tracking = Tracking.make(author, date, initials)
        try:
            yield self._tracking
        finally:
            self._tracking = previous

    def _active_tracking(self: "Document") -> Tracking | None:
        """The tracking an edit made now writes with, or ``None`` when it is untracked."""
        if self._track_stack:
            mode = self._track_stack[-1]
            return mode or None
        return self._tracking

    @contextmanager
    def _track_properties(self: "Document", parts) -> Iterator[None]:
        """Inside an edit: when tracking, record every paragraph-property and
        run-property change the code inside makes in ``parts`` as ``w:pPrChange`` /
        ``w:rPrChange`` (:class:`~docx_agent.revisions.track.PropertyTracker`)."""
        from . import track as _track
        from .stamp import Stamp

        tracking = self._active_tracking()
        if tracking is None:
            yield
            return
        trackers = [(part, _track.PropertyTracker([self.package.tree(part)])) for part in dict.fromkeys(parts)]
        try:
            yield
        except BaseException:
            for _, tracker in trackers:
                tracker.discard()
            raise
        stamp = Stamp(self, tracking)
        for part, tracker in trackers:
            tracker.record(stamp, part)
            self.package.mark_dirty(part)
        stamp.finish()

    @property
    def word_tracks_changes(self: "Document") -> bool:
        """Word's own Track Changes switch (``w:trackRevisions`` in the settings): whether
        *Word* records the changes a person makes after opening the file.

        It has no effect on this library's edits.  To write your own edits as tracked
        changes, use :meth:`tracking`::

            with doc.tracking(author="Claude"):
                doc.anchor("24 months").replace("36 months")

        Setting it (``doc.word_tracks_changes = True``) is :meth:`set_word_tracks_changes`."""
        part = self.package.settings_part()
        root = self.package.tree(part) if part else None
        node = root.find(_W + "trackRevisions") if root is not None else None
        return node is not None and (node.get(_W + "val") or "true").lower() not in ("0", "false", "off")

    @word_tracks_changes.setter
    def word_tracks_changes(self: "Document", value: bool) -> None:
        self.set_word_tracks_changes(value)

    def set_word_tracks_changes(self: "Document", value: bool) -> "EditResult":
        """Turn Word's own Track Changes switch (``w:trackRevisions``) on or off, as Word
        writes it (an empty element in its place in the settings' sequence; off is no
        element).  One undo step.

        This only decides whether Word tracks a person's later edits in Word; it does not
        make this library's edits tracked -- :meth:`tracking` does that.
        ``doc.set_word_tracks_changes(True)``."""
        from ..edit.document import EditError, EditResult

        part = self.package.settings_part()
        if part is None:
            raise EditError("the document has no settings part")
        if self.word_tracks_changes == bool(value):
            return EditResult(None, changed=False)
        with self._edit():
            root = self.package.tree(part)
            for node in root.findall(_W + "trackRevisions"):
                remove(node)
            if value:
                insert_in_order(root, make("w:trackRevisions"))
            self.package.mark_dirty(part)
        return EditResult(None)

    # Deprecated: the old names read as "record my edits as tracked changes", which they
    # never did (the end-to-end pilot, ROADMAP.md, "Usability").

    @property
    def track_revisions(self: "Document") -> bool:
        """Deprecated: :attr:`word_tracks_changes` (Word's own switch).  To track your own
        edits use :meth:`tracking`."""
        _deprecated("track_revisions", "word_tracks_changes")
        return self.word_tracks_changes

    @track_revisions.setter
    def track_revisions(self: "Document", value: bool) -> None:
        _deprecated("track_revisions", "word_tracks_changes")
        self.set_word_tracks_changes(value)

    def set_track_revisions(self: "Document", value: bool) -> "EditResult":
        """Deprecated: :meth:`set_word_tracks_changes` (Word's own switch).  To track your
        own edits use :meth:`tracking`."""
        _deprecated("set_track_revisions", "set_word_tracks_changes")
        return self.set_word_tracks_changes(value)


def _deprecated(old: str, new: str) -> None:
    warnings.warn(f"Document.{old} is deprecated: it is Document.{new}, Word's own Track Changes "
                  "switch for a person's later edits in Word.  To write your own edits as tracked "
                  "changes, use `with doc.tracking(author=...):`", DeprecationWarning, stacklevel=3)
