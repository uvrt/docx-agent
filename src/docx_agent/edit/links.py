"""Hyperlinks, bookmarks and cross-references.

**Hyperlinks** are ``w:hyperlink`` around runs: external ones name a relationship
(``r:id``, an ``External`` target, reused when the part already has one to the same
address), internal ones a bookmark (``w:anchor``).  Adding one splits runs at the range's
edges, wraps them, and gives them the Hyperlink character style (added as Word writes it
if the document lacks it); removing one unwraps it, takes the style off and drops the
relationship if nothing else uses it.  A hyperlink is ``hl:<paragraph id>/<n>``, the n-th
in its paragraph: positional, as runs are.

**Bookmarks** are named by their name -- ``bm:<name>`` -- because Word renumbers their
``w:id`` on every save (measured in E0, 9 -> 0).  A name starts with a letter, holds
letters, digits and underscores, is at most 40 characters and is unique in the document
regardless of case; names starting with ``_`` are Word's hidden ones and are not made
here.  A new bookmark's ``w:id`` is one above the largest annotation id in use in any
story.  Renaming one renames what points at it: internal hyperlinks and ``REF``,
``PAGEREF`` and ``NOTEREF`` field instructions.  A bookmark resolves to a text range.

**Cross-references** are ``REF`` (the bookmark's text) or ``PAGEREF`` (its page) fields
with ``\\h``, as Word inserts them, their result computed here: the text from the
bookmark, the page from docx2svg's layout (where the layout cannot say, ``1``, listed in
the result's ``unknown``: a field marked dirty makes Word ask on opening, measured in E4).
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, make, qn, remove
from . import inline as _inline
from . import text as _text
from .ids import ParagraphEntry
from .ranges import TextRange

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
W_HYPERLINK = _W + "hyperlink"
_NAME = re.compile(r"[^\W\d_][\w]{0,39}")
#: Word's own hidden cross-reference bookmarks (Insert > Cross-reference names its target
#: ``_Ref`` and nine digits): a name a caller may give too.
_REF_NAME = re.compile(r"_Ref\w{1,36}")
_SCHEME = re.compile(r"[A-Za-z][A-Za-z0-9+.\-]*:")

#: Elements whose ``w:id`` is an annotation id (bookmarks, comments, revisions, moves).
ANNOTATIONS = frozenset(_W + name for name in (
    "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "commentReference",
    "comment", "ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd",
    "moveToRangeStart", "moveToRangeEnd", "rPrChange", "pPrChange", "sectPrChange", "tblPrChange",
    "trPrChange", "tcPrChange", "tblGridChange", "numberingChange", "cellIns", "cellDel",
    "cellMerge", "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart",
    "customXmlDelRangeEnd", "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd", "permStart", "permEnd"))

_FIELD_REFERENCE = re.compile(r"(?i)(\b(?:REF|PAGEREF|NOTEREF)\s+)(\"?)([^\s\"\\]+)(\2)")


@dataclass(frozen=True)
class Hyperlink:
    """A ``w:hyperlink``, the ``index``-th in its paragraph (positional)."""

    document: "Document"
    paragraph_id: str
    index: int

    @property
    def id(self) -> str:
        """The hyperlink's id, ``hl:<paragraph id>/<index>``."""
        return f"hl:{self.document.paragraph(self.paragraph_id).id}/{self.index}"

    def _locate(self) -> tuple[str, ParagraphEntry, Element]:
        part, entry = self.document._paragraph_entry(self.paragraph_id)
        links = list(entry.element.iter(W_HYPERLINK))
        if not 0 <= self.index < len(links):
            raise KeyError(f"{self.paragraph_id} has no hyperlink {self.index}")
        return part, entry, links[self.index]

    @property
    def address(self) -> str | None:
        """The external target, or ``None`` for an internal link."""
        part, _, node = self._locate()
        rid = node.get(_R + "id")
        if rid is None:
            return None
        relationship = self.document.package.relationships(part).get(rid)
        return relationship.target if relationship is not None else None

    @property
    def anchor(self) -> str | None:
        """The bookmark an internal link goes to (or an external link's fragment)."""
        return self._locate()[2].get(_W + "anchor")

    @property
    def tooltip(self) -> str | None:
        """The hyperlink's tooltip (``w:tooltip``), or ``None``."""
        return self._locate()[2].get(_W + "tooltip")

    @property
    def text(self) -> str:
        """The hyperlink's text."""
        _, entry, node = self._locate()
        return "".join(a.char for a in _text.atoms(entry.element) if a.run is not None and _inside(a.run, node))

    def range(self) -> TextRange:
        """The text the hyperlink covers, as a :class:`TextRange`."""
        _, entry, node = self._locate()
        items = _text.atoms(entry.element)
        mine = [k for k, a in enumerate(items) if a.run is not None and _inside(a.run, node)]
        start, end = (mine[0], mine[-1] + 1) if mine else (0, 0)
        return TextRange(self.document, entry.id, start, entry.id, end)

    def set_target(self, url: str | None = None, *, anchor: str | None = None,
                   tooltip: str | None = None) -> "EditResult":
        """Point the hyperlink elsewhere: an external ``url`` or a bookmark ``anchor``,
        and its ``tooltip``.  ``link.set_target("https://example.com/")``."""
        return self.document._set_hyperlink_target(self, url, anchor, tooltip)

    def remove(self) -> "EditResult":
        """Remove the hyperlink; its text stays."""
        return self.document._remove_hyperlink(self)

    def __repr__(self) -> str:
        return f"<Hyperlink {self.id} {self.address or '#' + str(self.anchor)!r} {self.text!r}>"


def _inside(node: Element, ancestor: Element) -> bool:
    parent = node.getparent()
    while parent is not None:
        if parent is ancestor:
            return True
        parent = parent.getparent()
    return False


@dataclass(frozen=True)
class Bookmark:
    document: "Document"
    name: str

    @property
    def id(self) -> str:
        """The bookmark's id, ``bm:<name>``."""
        return f"bm:{self.name}"

    @property
    def hidden(self) -> bool:
        """Whether it is one of Word's hidden bookmarks (its name starts with ``_``)."""
        return self.name.startswith("_")

    def range(self) -> TextRange:
        """The text the bookmark marks (a range across paragraphs if it spans them)."""
        return self.document._bookmark_range(self.name)

    @property
    def text(self) -> str:
        """The text the bookmark marks."""
        return self.range().text

    def rename(self, new: str) -> "EditResult":
        """Rename the bookmark and what points at it: :meth:`Document.rename_bookmark`."""
        return self.document.rename_bookmark(self.name, new)

    def remove(self) -> "EditResult":
        """Remove the bookmark's markers (its text stays): :meth:`Document.remove_bookmark`."""
        return self.document.remove_bookmark(self.name)

    def __repr__(self) -> str:
        return f"<Bookmark {self.name!r}>"


def check_name(name: str) -> None:
    from .document import EditError

    if not isinstance(name, str) or not (_NAME.fullmatch(name) or _REF_NAME.fullmatch(name)):
        raise EditError(f"{name!r} is not a bookmark name: a letter, then letters, digits or "
                        "underscores, at most 40 characters -- or Word's hidden _Ref<digits>")


def _offset_of(paragraph: Element, marker: Element, view: str = "current") -> int:
    """How many characters of ``paragraph`` (in ``view``) precede ``marker``."""
    items = _text.atoms(paragraph, view)
    nodes = list(paragraph.iter())  # held, so each element keeps one proxy and one id()
    order = {id(node): k for k, node in enumerate(nodes)}
    at = order[id(marker)]
    return sum(1 for atom in items if order.get(id(atom.node), -1) < at)


class LinkOps:
    """The hyperlink, bookmark and cross-reference half of :class:`docx_agent.Document`."""

    # -- reading -----------------------------------------------------------------------------

    def hyperlinks(self: "Document", story: str | None = None) -> list[Hyperlink]:
        """Every hyperlink, in story and document order (one story's with ``story``)."""
        out = []
        for part in self._parts() if story is None else [self._part_for_story(story)]:
            for entry in self._index(part).paragraphs:
                count = sum(1 for _ in entry.element.iter(W_HYPERLINK))
                out += [Hyperlink(self, entry.id, k) for k in range(count)]
        return out

    def hyperlink(self: "Document", identifier: str) -> Hyperlink:
        """A hyperlink by id (``hl:<paragraph id>/<index>``); ``KeyError`` if none."""
        if not identifier.startswith("hl:"):
            raise KeyError(f"{identifier!r} is not a hyperlink id")
        paragraph, _, index = identifier[3:].rpartition("/")
        link = Hyperlink(self, paragraph, int(index))
        link._locate()
        return link

    def _bookmark_markers(self: "Document") -> dict[str, tuple[str, Element, Element | None]]:
        """name -> (part, bookmarkStart, bookmarkEnd) for every bookmark in every story."""
        out: dict[str, tuple[str, Element, Element | None]] = {}
        for part in self._parts():
            root = self.package.tree(part)
            ends = {node.get(_W + "id"): node for node in root.iter(_W + "bookmarkEnd")}
            for start in root.iter(_W + "bookmarkStart"):
                name = start.get(_W + "name") or ""
                out.setdefault(name, (part, start, ends.get(start.get(_W + "id"))))
        return out

    def bookmarks(self: "Document", *, hidden: bool = False) -> list[Bookmark]:
        """Every bookmark in document order.  Word's hidden ones -- names starting with
        ``_``: a table of contents' ``_Toc...``, a cross-reference's ``_Ref...``,
        ``_GoBack`` -- are left out, as Word's Bookmark dialog leaves them out, unless
        ``hidden=True``; :meth:`bookmark` finds one by name either way."""
        return [Bookmark(self, name) for name in self._bookmark_markers() if hidden or not name.startswith("_")]

    def bookmark(self: "Document", name: str) -> Bookmark:
        """A bookmark by name (``bm:`` prefix optional; found regardless of case)."""
        name = name[3:] if name.startswith("bm:") else name
        markers = self._bookmark_markers()
        if name in markers:
            return Bookmark(self, name)
        for existing in markers:
            if existing.casefold() == name.casefold():
                return Bookmark(self, existing)
        raise KeyError(f"no bookmark {name!r}")

    def _bookmark_range(self: "Document", name: str) -> TextRange:
        part, start, end = self._bookmark_markers()[self.bookmark(name).name]
        index = self._index(part)
        first, first_offset = self._marker_place(part, start, forward=True)
        if end is None:
            last, last_offset = first, first_offset
        else:
            last, last_offset = self._marker_place(part, end, forward=False)
        if (last.index, last_offset) < (first.index, first_offset):
            last, last_offset = first, first_offset
        del index
        return TextRange(self, first.id, first_offset, last.id, last_offset)

    def _marker_place(self: "Document", part: str, marker: Element, *, forward: bool,
                      view: str = "current") -> tuple[ParagraphEntry, int]:
        """The paragraph and offset a range marker stands at: in its paragraph, or -- at
        block level -- the start of the next paragraph (``forward``) or end of the previous."""
        index = self._index(part)
        node = marker.getparent()
        while node is not None and node.tag != _W + "p":
            node = node.getparent()
        if node is not None:
            entry = index.entry_for(node)
            return entry, _offset_of(node, marker, view)
        nodes = list(self.package.tree(part).iter())  # held: stable id()s
        order = {id(n): k for k, n in enumerate(nodes)}
        at = order[id(marker)]
        paragraphs = index.paragraphs
        if forward:
            for entry in paragraphs:
                if order[id(entry.element)] > at:
                    return entry, 0
            entry = paragraphs[-1]
            return entry, len(_text.atoms(entry.element, view))
        for entry in reversed(paragraphs):
            if order[id(entry.element)] < at:
                return entry, len(_text.atoms(entry.element, view))
        return paragraphs[0], 0

    # -- annotation ids ----------------------------------------------------------------------

    def _next_annotation_id(self: "Document") -> int:
        """One above the largest annotation id in use in any story (ROADMAP.md, "Revision
        ids": stricter than the specification, and what Word does)."""
        largest = -1
        for part in self._parts():
            for node in self.package.tree(part).iter():
                if isinstance(node.tag, str) and node.tag in ANNOTATIONS:
                    try:
                        largest = max(largest, int(node.get(_W + "id")))
                    except (TypeError, ValueError):
                        pass
        return largest + 1

    # -- hyperlinks --------------------------------------------------------------------------

    def add_hyperlink(self: "Document", where: TextRange, url: str | None = None, *, anchor: str | None = None,
                      tooltip: str | None = None, style: bool = True) -> "EditResult":
        """Make a range (in one paragraph) a hyperlink: to ``url`` (external), or to the
        bookmark ``anchor``.  ``style`` gives its runs the Hyperlink character style."""
        from .document import EditError, EditResult

        if (url is None) == (anchor is None):
            raise EditError("give a url or an anchor (a bookmark name), not both")
        if url is not None and not _SCHEME.match(url):
            raise EditError(f"{url!r} is not an address (it needs a scheme: https:, mailto:...)")
        if anchor is not None:
            self.bookmark(anchor)
        if where.view != "current" or not where.single or where.collapsed:
            raise EditError("a hyperlink is made from a non-empty range in one paragraph, in the current view")
        part, entries = where._entries()
        entry = entries[0]
        items = _text.atoms(entry.element)
        try:
            _inline.check_span(items, where.start, where.end, deleting=False)
        except _inline.SpanError as error:
            raise EditError(f"{where.id}: {error}") from None
        if any(a.field is not None for a in items[where.start:where.end]):
            raise EditError(f"{where.id} is in a field's result")
        for atom in items[where.start:where.end]:
            if atom.run is not None and any(a.tag == W_HYPERLINK for a in atom.run.iterancestors()):
                raise EditError(f"{where.id} is already (partly) a hyperlink; change or remove that one")
        if style:
            self.styles.resolve("Hyperlink", "character")
        tracking = self._active_tracking()
        if tracking is not None:
            return self._add_hyperlink_tracked(part, entry, where, url, anchor, tooltip, style, tracking)
        with self._edit():
            renames = self._prepare(part, [entry])
            runs = _inline.isolate(entry.element, where.start, where.end)
            parent = runs[0].getparent()
            if any(run.getparent() is not parent for run in runs):
                raise EditError(f"{where.id} spans runs in different containers (an insertion, a "
                                "content control); a hyperlink cannot wrap them")
            link = make("w:hyperlink")
            if url is not None:
                link.set(_R + "id", self.package.add_external_relationship(part, REL_HYPERLINK, url))
            else:
                link.set(_W + "anchor", anchor)
            if tooltip:
                link.set(_W + "tooltip", tooltip)
            link.set(_W + "history", "1")
            runs[0].addprevious(link)
            node = runs[0]
            while node is not None:
                following = node.getnext()
                link.append(node)
                if node is runs[-1]:
                    break
                node = following
            if style:
                style_id = self.styles._ensure("Hyperlink", "character")
                from . import formatting as _formatting

                for run in runs:
                    properties = _formatting.properties_of(run, "w:rPr", create=True)
                    _formatting._set_child(properties, "w:rStyle", {"w:val": style_id})
            _inline.merge_runs(runs)
            self.package.mark_dirty(part)
            index = list(entry.element.iter(W_HYPERLINK)).index(link)
        new_id = renames.get(entry.id, entry.id)
        return EditResult(f"hl:{new_id}/{index}", created=[f"hl:{new_id}/{index}"], renamed=renames)

    def _new_link(self: "Document", part: str, url: str | None, anchor: str | None, tooltip: str | None) -> Element:
        link = make("w:hyperlink")
        if url is not None:
            link.set(_R + "id", self.package.add_external_relationship(part, REL_HYPERLINK, url))
        else:
            link.set(_W + "anchor", anchor)
        if tooltip:
            link.set(_W + "tooltip", tooltip)
        link.set(_W + "history", "1")
        return link

    def _add_hyperlink_tracked(self: "Document", part, entry, where, url, anchor, tooltip, style, tracking) -> "EditResult":
        """A hyperlink made under tracking: the range's runs deleted, and a hyperlink holding
        their copies inserted after them.  Word, tracking, writes an inserted ``HYPERLINK``
        field over the text instead -- the text itself counted inserted, so its own Reject
        All loses the text (measured); this form rejects to the original."""
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp
        from .document import EditError, EditResult
        from . import formatting as _formatting

        with self._edit():
            renames = self._prepare(part, [entry])
            runs = _inline.isolate(entry.element, where.start, where.end)
            parent = runs[0].getparent()
            if any(run.getparent() is not parent for run in runs):
                raise EditError(f"{where.id} spans runs in different containers (an insertion, a "
                                "content control); a hyperlink cannot wrap them")
            copies = [copy.deepcopy(run) for run in runs]
            for run in copies:
                for node in run.iter(_W + "rPrChange"):
                    remove(node)
            if style:
                style_id = self.styles._ensure("Hyperlink", "character")
                for run in copies:
                    properties = _formatting.properties_of(run, "w:rPr", create=True)
                    _formatting._set_child(properties, "w:rStyle", {"w:val": style_id})
            place = make("w:proofErr")  # a placeholder where the link goes
            runs[-1].addnext(place)
            stamp = Stamp(self, tracking)
            released = _track.delete_runs(runs, stamp, part)
            link = self._new_link(part, url, anchor, tooltip)
            place.addprevious(link)
            remove(place)
            _track.place_outside_revisions(link, stamp)
            container = stamp.make("w:ins", part)
            link.append(container)
            for run in copies:
                container.append(run)
            stamp.finish()
            self._release(part, released)
            self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
            index = list(entry.element.iter(W_HYPERLINK)).index(link)
        new_id = renames.get(entry.id, entry.id)
        return EditResult(f"hl:{new_id}/{index}", created=[f"hl:{new_id}/{index}"], renamed=renames)

    def _replace_link_tracked(self: "Document", link: Hyperlink, url, anchor, tooltip, *, unlink: bool) -> "EditResult":
        """A hyperlink changed (``unlink`` False: given a new target) or removed under
        tracking: its runs deleted inside it, their copies inserted after it -- in a new
        hyperlink, or as plain text without the Hyperlink style."""
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp
        from .document import EditResult

        tracking = self._active_tracking()
        part, entry, node = link._locate()
        with self._edit():
            renames = self._prepare(part, [entry])
            part, entry, node = link._locate()
            runs = [run for run in _text.walk(entry.element).runs if _inside(run, node)]
            copies = [copy.deepcopy(run) for run in runs]
            hyperlink_style = self.styles.find("Hyperlink", "character")
            for run in copies:
                for record in run.iter(_W + "rPrChange"):
                    remove(record)
                if unlink and hyperlink_style is not None:
                    found = run.find(f"{_W}rPr/{_W}rStyle")
                    if found is not None and found.get(_W + "val") == hyperlink_style.id:
                        properties = found.getparent()
                        remove(found)
                        if not len(properties):
                            remove(properties)
            for child in copies:
                for wrapper in list(child.iter(_W + "ins")):
                    for inner in list(wrapper):
                        wrapper.addprevious(inner)
                    remove(wrapper)
            place = make("w:proofErr")
            node.addnext(place)
            stamp = Stamp(self, tracking)
            released = _track.delete_runs(runs, stamp, part)
            released += _inline.prune(entry.element)
            if unlink:
                parent, index = place.getparent(), place.getparent().index(place)
                remove(place)
                if copies:
                    _track.insert_nodes(parent, index, copies, stamp, part)
                result_id = renames.get(entry.id, entry.id)
            else:
                new = self._new_link(part, url if url is not None else (None if anchor is not None else link.address),
                                     anchor if anchor is not None or url is not None else link_anchor(node),
                                     tooltip if tooltip is not None else node.get(_W + "tooltip"))
                place.addprevious(new)
                remove(place)
                _track.place_outside_revisions(new, stamp)
                container = stamp.make("w:ins", part)
                new.append(container)
                for run in copies:
                    container.append(run)
                index = list(entry.element.iter(W_HYPERLINK)).index(new)
                result_id = f"hl:{renames.get(entry.id, entry.id)}/{index}"
            stamp.finish()
            self._release(part, released)
            self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
        return EditResult(result_id, renamed=renames, removed=[link.id] if unlink else [])

    def _set_hyperlink_target(self: "Document", link: Hyperlink, url, anchor, tooltip) -> "EditResult":
        from .document import EditError, EditResult

        if url is not None and anchor is not None:
            raise EditError("give a url or an anchor, not both")
        if url is not None and not _SCHEME.match(url):
            raise EditError(f"{url!r} is not an address (it needs a scheme: https:, mailto:...)")
        if anchor is not None:
            self.bookmark(anchor)
        if self._active_tracking() is not None and (url is not None or anchor is not None):
            return self._replace_link_tracked(link, url, anchor, tooltip, unlink=False)
        part, entry, node = link._locate()
        with self._edit():
            renames = self._prepare(part, [entry])
            part, entry, node = link._locate()
            old = node.get(_R + "id")
            if url is not None:
                node.set(_R + "id", self.package.add_external_relationship(part, REL_HYPERLINK, url))
                node.attrib.pop(_W + "anchor", None)
            elif anchor is not None:
                node.attrib.pop(_R + "id", None)
                node.set(_W + "anchor", anchor)
            if tooltip is not None:
                if tooltip:
                    node.set(_W + "tooltip", tooltip)
                else:
                    node.attrib.pop(_W + "tooltip", None)
            if old and old != node.get(_R + "id"):
                self._release(part, [old])
            self.package.mark_dirty(part)
        return EditResult(link.id, renamed=renames)

    def _remove_hyperlink(self: "Document", link: Hyperlink) -> "EditResult":
        from .document import EditResult

        if self._active_tracking() is not None:
            return self._replace_link_tracked(link, None, None, None, unlink=True)
        part, entry, node = link._locate()
        removed_id = link.id
        with self._edit():
            renames = self._prepare(part, [entry])
            part, entry, node = link._locate()
            rid = node.get(_R + "id")
            hyperlink_style = self.styles.find("Hyperlink", "character")
            runs = [child for child in node if child.tag == _W + "r"]
            for child in list(node):
                node.addprevious(child)
            remove(node)
            if hyperlink_style is not None:
                for run in runs:
                    style = run.find(f"{_W}rPr/{_W}rStyle")
                    if style is not None and style.get(_W + "val") == hyperlink_style.id:
                        properties = style.getparent()
                        remove(style)
                        if not len(properties):
                            remove(properties)
            _inline.merge_around(runs)
            if rid:
                self._release(part, [rid])
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames, removed=[removed_id])

    # -- bookmarks ---------------------------------------------------------------------------

    def add_bookmark(self: "Document", where: TextRange, name: str) -> "EditResult":
        """Bookmark a range (an empty one marks a place) under ``name``."""
        from .document import EditError, EditResult

        check_name(name)
        markers = self._bookmark_markers()
        if any(existing.casefold() == name.casefold() for existing in markers):
            raise EditError(f"the document already has a bookmark {name!r}")
        if where.view != "current":
            raise EditError("edits take the current view")
        part, entries = where._entries()
        first, last = entries[0], entries[-1]
        for entry, offset in ((first, where.start), (last, where.end)):
            if not 0 <= offset <= len(_text.atoms(entry.element)):
                raise EditError(f"{where.id} is outside its paragraphs' text")
        with self._edit():
            renames = self._prepare(part, [first, last])
            mark_id = str(self._next_annotation_id())
            end = make("w:bookmarkEnd")
            end.set(_W + "id", mark_id)
            parent, position = _inline.position(last.element, where.end)
            parent.insert(position, end)
            start = make("w:bookmarkStart")
            start.set(_W + "id", mark_id)
            start.set(_W + "name", name)
            parent, position = _inline.position(first.element, where.start)
            if parent is end.getparent() and position > parent.index(end):
                position = parent.index(end)
            parent.insert(position, start)
            self.package.mark_dirty(part)
        return EditResult(f"bm:{name}", created=[f"bm:{name}"], renamed=renames)

    def rename_bookmark(self: "Document", name: str, new: str) -> "EditResult":
        """Rename a bookmark, and every internal hyperlink and ``REF``/``PAGEREF``/``NOTEREF``
        field instruction that names it."""
        from .document import EditError, EditResult

        old = self.bookmark(name).name
        check_name(new)
        if any(existing.casefold() == new.casefold() and existing != old for existing in self._bookmark_markers()):
            raise EditError(f"the document already has a bookmark {new!r}")
        with self._edit():
            renames = self._prepare_many([])
            part, start, _ = self._bookmark_markers()[old]
            start.set(_W + "name", new)
            touched = {part}
            for story in self._parts():
                root = self.package.tree(story)
                for link in root.iter(W_HYPERLINK):
                    if link.get(_W + "anchor") == old:
                        link.set(_W + "anchor", new)
                        touched.add(story)
                for node in root.iter(_W + "instrText"):
                    updated = _rename_in_instruction(node.text or "", old, new)
                    if updated != node.text:
                        node.text = updated
                        touched.add(story)
                for node in root.iter(_W + "fldSimple"):
                    updated = _rename_in_instruction(node.get(_W + "instr") or "", old, new)
                    if updated != node.get(_W + "instr"):
                        node.set(_W + "instr", updated)
                        touched.add(story)
            for story in touched:
                self.package.mark_dirty(story)
        return EditResult(f"bm:{new}", renamed={f"bm:{old}": f"bm:{new}", **renames})

    def remove_bookmark(self: "Document", name: str) -> "EditResult":
        """Remove a bookmark's markers (its text stays).  What still points at it --
        internal hyperlinks, ``REF`` fields -- is listed in ``result.warnings``."""
        from .document import EditResult

        old = self.bookmark(name).name
        dangling = []
        for story in self._parts():
            root = self.package.tree(story)
            dangling += [f"a hyperlink in {self._story_of(story)}" for link in root.iter(W_HYPERLINK)
                         if link.get(_W + "anchor") == old]
            dangling += [f"a field in {self._story_of(story)}" for node in root.iter(_W + "instrText")
                         if _rename_in_instruction(node.text or "", old, "x") != (node.text or "")]
        with self._edit():
            renames = self._prepare_many([])
            part, start, end = self._bookmark_markers()[old]
            for marker in (start, end):
                if marker is not None:
                    remove(marker)
            self.package.mark_dirty(part)
        result = EditResult(None, renamed=renames, removed=[f"bm:{old}"])
        result.warnings = [f"{what} still points at the removed bookmark {old!r}" for what in dangling]
        return result

    # -- cross-references --------------------------------------------------------------------

    def insert_cross_reference(self: "Document", at: "str | TextRange", bookmark: str, *, kind: str = "text",
                               hyperlink: bool = True) -> "EditResult":
        """Insert a ``REF`` (``kind="text"``: the bookmark's text) or ``PAGEREF``
        (``kind="page"``: its page) field at a position, as Word's Insert Cross-reference
        does, with its result computed here."""
        from .document import EditError, EditResult

        if kind not in ("text", "page"):
            raise EditError("kind is 'text' or 'page'")
        target = self.bookmark(bookmark)
        where = at if isinstance(at, TextRange) else self.range(at)
        part, entries = where._entries()
        entry = entries[0]
        items = _text.atoms(entry.element)
        if not 0 <= where.start <= len(items):
            raise EditError(f"{where.id} is outside its paragraph's text")
        if 0 < where.start < len(items) and items[where.start].field is not None \
                and items[where.start - 1].field is items[where.start].field:
            raise EditError(f"{where.id} is inside a field's result")
        unknown = False
        if kind == "text":
            result_text = target.text.split("\n")[0]
        else:
            # The number Word shows on the bookmark's page (its section's start and format).
            label = self.page_label(target.range().start_id)
            if label is None:
                result_text, unknown = "1", True
            else:
                result_text = label
        result_text = result_text.replace(_text.OBJECT, "") or " "
        switch = " \\h" if hyperlink else ""
        instruction = f" {'REF' if kind == 'text' else 'PAGEREF'} {target.name}{switch} "
        with self._edit():
            renames = self._prepare(part, [entry])
            pieces = []
            begin = make("w:fldChar")
            begin.set(_W + "fldCharType", "begin")
            pieces.append([begin])
            code = make("w:instrText")
            code.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            code.text = instruction
            pieces.append([code])
            separate = make("w:fldChar")
            separate.set(_W + "fldCharType", "separate")
            pieces.append([separate])
            pieces.append(_text._spell(result_text))
            end = make("w:fldChar")
            end.set(_W + "fldCharType", "end")
            pieces.append([end])
            previous = _inline.new_run(entry.element, where.start, pieces[0])
            properties = previous.find(_W + "rPr")
            field_runs = [previous]
            for children in pieces[1:]:
                run = etree.Element(_W + "r")
                if properties is not None:
                    run.append(copy.deepcopy(properties))
                for child in children:
                    run.append(child)
                previous.addnext(run)
                previous = run
                field_runs.append(run)
            self._track_inserted(part, field_runs)
            self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames,
                          unknown=[f"PAGEREF {target.name}"] if unknown else [])


def link_anchor(node: Element) -> str | None:
    return node.get(_W + "anchor")


def _rename_in_instruction(instruction: str, old: str, new: str) -> str:
    def swap(match: re.Match) -> str:
        if match.group(3) == old:
            return f"{match.group(1)}{match.group(2)}{new}{match.group(4)}"
        return match.group(0)

    return _FIELD_REFERENCE.sub(swap, instruction)
