"""Read-only views of what E3 and E4 will edit: notes, comments, revisions, content
controls and drawings other than pictures.

E2's readers (``to_markdown``, ``state``) name these things by id, and every id they write
must resolve through :meth:`docx_agent.Document.get`; these views are what it resolves to.
They read, and change nothing.  The id forms are ROADMAP.md's ("Everything else"):

* a footnote ``fn:<w:id>``, an endnote ``en:<w:id>`` -- separators (``w:type``) are not
  notes; Word renumbers note ids on save, so these hold for the session;
* a comment ``c:<w16cid:durableId>`` when ``commentsIds.xml`` gives it one, else
  ``c#<w:id>``;
* a revision ``rev:<w:id>`` (``#k`` for the k-th repeat of an id) -- every ``w:ins``,
  ``w:del``, ``w:moveFrom``, ``w:moveTo`` and property change, in every story;
* a content control ``cc:<w:sdtPr/w:id>``, or ``cc@<story>/<n>`` (volatile) when it has none;
* a drawing ``d:<wp:docPr/@id>``, as pictures are (``#k`` for a repeat).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..oxml.xml import Element
from . import ids as _ids
from . import text as _text

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document
    from .ranges import TextRange

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
_W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
_W16CID = "{http://schemas.microsoft.com/office/word/2016/wordml/cid}"
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"

REL_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
REL_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"

#: Revision elements and the kind each records.
REVISION_KINDS = {
    _W + "ins": "insertion", _W + "del": "deletion", _W + "moveFrom": "move-from",
    _W + "moveTo": "move-to", _W + "rPrChange": "run-properties", _W + "pPrChange": "paragraph-properties",
    _W + "sectPrChange": "section-properties", _W + "tblPrChange": "table-properties",
    _W + "trPrChange": "row-properties", _W + "tcPrChange": "cell-properties",
    _W + "tblGridChange": "table-grid", _W + "numberingChange": "numbering",
    _W + "cellIns": "cell-insertion", _W + "cellDel": "cell-deletion", _W + "cellMerge": "cell-merge",
}

#: Content-control kinds, by the ``w:sdtPr`` child that says it (first match wins).
_CONTROL_KINDS = (
    (_W14 + "checkbox", "checkbox"), (_W + "dropDownList", "drop-down"), (_W + "comboBox", "combo-box"),
    (_W + "date", "date"), (_W + "picture", "picture"), (_W + "docPartObj", "building-block"),
    (_W + "docPartList", "building-block"), (_W + "group", "group"), (_W + "citation", "citation"),
    (_W + "bibliography", "bibliography"), (_W + "equation", "equation"),
    (_W15 + "repeatingSection", "repeating-section"), (_W15 + "repeatingSectionItem", "repeating-section-item"),
    (_W + "text", "text"), (_W + "richText", "rich-text"),
)


def _paragraph_of(node: Element) -> Element | None:
    while node is not None and node.tag != _W + "p":
        node = node.getparent()
    return node


def _paragraph_id(document: "Document", part: str, node: Element) -> str | None:
    paragraph = _paragraph_of(node)
    if paragraph is None:
        return None
    entry = document._index(part).entry_for(paragraph)
    return entry.id if entry is not None else None


def _blocks_text(document: "Document", part: str, container: Element, view: str = "current") -> str:
    index = document._index(part)  # holds every paragraph's proxy, so id()s are stable
    inside = {id(e) for e in container.iter(_W + "p")}
    return "\n".join(_text.paragraph_text(entry.element, view) for entry in index.paragraphs
                     if id(entry.element) in inside)


# -- notes -----------------------------------------------------------------------------------


@dataclass(frozen=True)
class Note:
    """A footnote or endnote (``fn:<id>``, ``en:<id>``)."""

    document: "Document"
    id: str

    @property
    def kind(self) -> str:
        """``"footnote"`` or ``"endnote"``."""
        return "footnote" if self.id.startswith("fn:") else "endnote"

    def _locate(self) -> tuple[str, Element]:
        for part, identifier, element in self.document._notes():
            if identifier == self.id:
                return part, element
        raise KeyError(f"no note {self.id!r}")

    @property
    def story(self) -> str:
        """The story the note's text is in (``footnotes`` or ``endnotes``)."""
        return self.document._story_of(self._locate()[0])

    @property
    def paragraph_ids(self) -> list[str]:
        """The ids of the note's own paragraphs, in its story (a property; calling it works
        too)."""
        from .document import CallableList

        part, element = self._locate()
        index = self.document._index(part)  # first: it holds the proxies the id()s compare
        inside = {id(e) for e in element.iter(_W + "p")}
        return CallableList(entry.id for entry in index.paragraphs if id(entry.element) in inside)

    def edit(self, text: str):
        """Replace the note's text (a line per paragraph): :meth:`Document.edit_note`."""
        return self.document.edit_note(self.id, text)

    def delete(self):
        """Delete the note and its reference: :meth:`Document.delete_note`."""
        return self.document.delete_note(self.id)

    def move(self, to):
        """Move the note's reference to a position: :meth:`Document.move_note`."""
        return self.document.move_note(self.id, to)

    def text_in(self, view: str = "current") -> str:
        """The note's text, without the reference mark that opens it."""
        part, element = self._locate()
        return _blocks_text(self.document, part, element, view).replace(_text.OBJECT, "", 1).strip()

    @property
    def text(self) -> str:
        """The note's text in the current view (:meth:`text_in` takes another view)."""
        return self.text_in()

    def __repr__(self) -> str:
        return f"<Note {self.id}>"


# -- comments --------------------------------------------------------------------------------


@dataclass(frozen=True)
class Comment:
    """A comment: ``c:<durableId>`` or ``c#<w:id>``.  ``doc.comments()`` lists them.

    ``comment.anchor`` is the text it is attached to, in the body or another story;
    ``comment.text`` is what the comment says::

        for comment in doc.comments():
            print(comment.id, comment.author, comment.anchor.text, "->", comment.text)
        doc.comment("c:76FCC91F").reply("Done.", author="Claude")
    """

    document: "Document"
    id: str

    def _record(self) -> dict:
        for record in self.document._comment_records():
            if record["id"] == self.id:
                return record
        raise KeyError(f"no comment {self.id!r}")

    @property
    def author(self) -> str | None:
        """The comment's author (``w:author``)."""
        return self._record()["author"]

    @property
    def initials(self) -> str | None:
        """The author's initials (``w:initials``), or ``None``."""
        return self._record()["initials"]

    @property
    def date(self) -> str | None:
        """When it was written (ISO 8601, as Word stores it), or ``None``."""
        return self._record()["date"]

    @property
    def parent(self) -> str | None:
        """The comment this one replies to (``commentsExtended``), or ``None``."""
        return self._record()["parent"]

    @property
    def done(self) -> bool:
        """Whether the thread is resolved (:meth:`resolve`, :meth:`reopen`)."""
        return self._record()["done"]

    @property
    def anchor(self) -> "TextRange | None":
        """The text the comment is attached to, as a :class:`~docx_agent.TextRange` in
        the story it covers (the body, a header, a note...): ``.text``, ``.id`` and
        ``.paragraph_ids()``.  An empty range when it comments on a place; a reply's is its
        thread's.  ``None`` if the document holds no marker for it::

            doc.comment("c:76FCC91F").anchor.text      # '24 months'
        """
        return self.document._comment_anchor(self.id)

    @property
    def content_paragraph_ids(self) -> list[str]:
        """The ids of the comment's *own* paragraphs (its text, in the comments story:
        ``comments/p:...``).  For the paragraphs it is attached to, use
        ``comment.anchor.paragraph_ids()``."""
        return list(self._record()["paragraphs"])

    @property
    def paragraph_ids(self) -> list[str]:
        """Deprecated: :attr:`content_paragraph_ids` (the comment's own paragraphs, not the
        ones it is attached to -- those are ``comment.anchor.paragraph_ids()``)."""
        import warnings

        warnings.warn("Comment.paragraph_ids is deprecated: it is Comment.content_paragraph_ids, the "
                      "comment's own paragraphs in the comments story.  The text the comment is attached "
                      "to is Comment.anchor (its .text and .paragraph_ids()).", DeprecationWarning, stacklevel=2)
        return self.content_paragraph_ids

    @property
    def text(self) -> str:
        """What the comment says (its paragraphs joined by ``\\n``)."""
        record = self._record()
        return _blocks_text(self.document, record["part"], record["element"]).replace(_text.OBJECT, "").strip()

    @property
    def replies(self) -> list["Comment"]:
        """The comments replying to this one, in order."""
        return [Comment(self.document, r["id"]) for r in self.document._comment_records() if r["parent"] == self.id]

    def reply(self, text: str, *, author: str | None = None, initials: str | None = None, date=None):
        """A reply in this comment's thread, by ``author`` (``initials``, ``date``: now by
        default): :meth:`Document.reply_to_comment`.  ``comment.reply("Done.", author="Claude")``."""
        return self.document.reply_to_comment(self.id, text, author=author, initials=initials, date=date)

    def resolve(self):
        """Mark the thread resolved: :meth:`Document.resolve_comment`."""
        return self.document.resolve_comment(self.id)

    def reopen(self):
        """Mark the thread open again: :meth:`Document.reopen_comment`."""
        return self.document.reopen_comment(self.id)

    def edit(self, text: str):
        """Replace the comment's text: :meth:`Document.edit_comment`."""
        return self.document.edit_comment(self.id, text)

    def delete(self):
        """Delete the comment and its replies: :meth:`Document.delete_comment`."""
        return self.document.delete_comment(self.id)

    def __repr__(self) -> str:
        return f"<Comment {self.id} by {self.author!r}>"


# -- revisions -------------------------------------------------------------------------------


@dataclass(frozen=True)
class Revision:
    """One revision record (``rev:<w:id>``): read-only until E3."""

    document: "Document"
    id: str

    def _locate(self) -> tuple[str, Element]:
        for part, identifier, element in self.document._revision_elements():
            if identifier == self.id:
                return part, element
        raise KeyError(f"no revision {self.id!r}")

    @property
    def kind(self) -> str:
        """``insertion``, ``deletion``, ``move-from``, ``move-to``, a property change
        (``run-properties``...), or for a paragraph mark's or row's own record
        ``paragraph-mark-insertion`` / ``row-deletion``..."""
        return revision_kind(self._locate()[1])

    @property
    def author(self) -> str | None:
        return self._locate()[1].get(_W + "author")

    @property
    def date(self) -> str | None:
        return self._locate()[1].get(_W + "date")

    @property
    def story(self) -> str:
        return self.document._story_of(self._locate()[0])

    @property
    def paragraph_id(self) -> str | None:
        part, element = self._locate()
        return _paragraph_id(self.document, part, element)

    @property
    def text(self) -> str:
        """The text the revision inserts, deletes or moves (empty for a property change)."""
        return revision_text(self._locate()[1])

    def __repr__(self) -> str:
        return f"<Revision {self.id} {self.kind} by {self.author!r}>"


def revision_kind(element: Element) -> str:
    kind = REVISION_KINDS[element.tag]
    parent = element.getparent()
    if element.tag in (_W + "ins", _W + "del", _W + "moveFrom", _W + "moveTo") and parent is not None:
        if parent.tag == _W + "rPr" and parent.getparent() is not None and parent.getparent().tag == _W + "pPr":
            return "paragraph-mark-" + kind
        if parent.tag == _W + "rPr":
            return "run-mark-" + kind
        if parent.tag == _W + "trPr":
            return "row-" + kind
    return kind


def revision_text(element: Element) -> str:
    if element.tag not in (_W + "ins", _W + "del", _W + "moveFrom", _W + "moveTo"):
        return ""
    return "".join(node.text or "" for node in element.iter(_W + "t", _W + "delText"))


# -- content controls ------------------------------------------------------------------------


@dataclass(frozen=True)
class ContentControl:
    """A ``w:sdt`` at block, run, row or cell level (``cc:<w:id>``)."""

    document: "Document"
    id: str

    def _locate(self) -> tuple[str, Element]:
        for part, identifier, element in self.document._content_controls():
            if identifier == self.id:
                return part, element
        raise KeyError(f"no content control {self.id!r}")

    @property
    def kind(self) -> str:
        """The control's kind: ``"rich-text"``, ``"text"``, ``"drop-down"``, ``"combo-box"``,
        ``"date"``, ``"checkbox"``, ``"picture"``..."""
        return control_kind(self._locate()[1])

    @property
    def level(self) -> str:
        """``block``, ``inline``, ``row`` or ``cell``: where the control sits."""
        return control_level(self._locate()[1])

    @property
    def tag(self) -> str | None:
        """The control's tag (``w:tag``), or ``None``."""
        return _sdt_value(self._locate()[1], "tag")

    @property
    def alias(self) -> str | None:
        """The control's title (``w:alias``), or ``None``."""
        return _sdt_value(self._locate()[1], "alias")

    @property
    def text(self) -> str:
        """The text the control holds now."""
        part, element = self._locate()
        if control_level(element) == "inline":
            paragraph = _paragraph_of(element)
            content = element.find(_W + "sdtContent")
            if content is None or paragraph is None:
                return ""
            nodes = list(content.iter())  # held: one proxy per element, stable id()s
            inside = {id(n) for n in nodes}
            return "".join(a.char for a in _text.atoms(paragraph) if id(a.node) in inside)
        return _blocks_text(self.document, part, element)

    @property
    def value(self):
        """A check box's state, a date control's date, a list's chosen value, else the text
        (:meth:`Document.control_value`)."""
        return self.document.control_value(self.id)

    @property
    def items(self) -> list[tuple[str, str]]:
        """A list control's items: ``(display text, value)``."""
        from .controls import items

        return items(self._locate()[1])

    @property
    def binding(self) -> dict | None:
        """``{xpath, store, prefixes}`` of a data-bound control, else ``None``."""
        from .controls import binding

        return binding(self._locate()[1])

    @property
    def lock(self) -> str | None:
        """How the control is locked (``"sdtLocked"``, ``"contentLocked"``,
        ``"sdtContentLocked"``), or ``None``."""
        from .controls import lock

        return lock(self._locate()[1])

    def fill(self, value):
        """Fill the control as typing into it would: :meth:`Document.fill_control`.
        ``doc.content_control("cc:101").fill("Final")``."""
        return self.document.fill_control(self.id, value)

    def remove(self, *, keep_content: bool = True):
        """Remove the control, keeping its content unless ``keep_content=False``:
        :meth:`Document.remove_control`."""
        return self.document.remove_control(self.id, keep_content=keep_content)

    def __repr__(self) -> str:
        return f"<ContentControl {self.id} {self.kind}>"


def control_kind(sdt: Element) -> str:
    properties = sdt.find(_W + "sdtPr")
    if properties is not None:
        for tag, kind in _CONTROL_KINDS:
            if properties.find(tag) is not None:
                return kind
    return "rich-text"


def control_level(sdt: Element) -> str:
    parent = sdt.getparent()
    while parent is not None and parent.tag in (_W + "customXml", _W + "sdtContent", _W + "smartTag"):
        parent = parent.getparent()
    tag = parent.tag if parent is not None else None
    if tag in (_W + "p", _W + "hyperlink", _W + "ins", _W + "del", _W + "moveTo", _W + "moveFrom",
               _W + "fldSimple", _W + "sdt"):
        return "inline" if sdt.find(f"{_W}sdtContent/{_W}p") is None and sdt.find(f"{_W}sdtContent/{_W}tbl") is None \
            else "block"
    if tag == _W + "tbl":
        return "row"
    if tag == _W + "tr":
        return "cell"
    return "block"


def _sdt_value(sdt: Element, name: str) -> str | None:
    node = sdt.find(f"{_W}sdtPr/{_W}{name}")
    return node.get(_W + "val") if node is not None else None


# -- drawings --------------------------------------------------------------------------------


@dataclass(frozen=True)
class Drawing:
    """A drawing that is not a picture -- a shape, text box, group, chart, diagram or
    canvas -- by its ``wp:docPr`` id (``d:<id>``)."""

    document: "Document"
    id: str

    def _locate(self) -> tuple[str, Element]:
        for part, identifier, frame in self.document._all_drawings():
            if identifier == self.id:
                return part, frame
        raise KeyError(f"no drawing {self.id!r}")

    @property
    def kind(self) -> str:
        """What the drawing is: ``"picture"``, ``"chart"``, ``"diagram"``, ``"text-box"``,
        ``"shape"``, ``"group"``..."""
        return drawing_kind(self._locate()[1])

    @property
    def name(self) -> str | None:
        """The drawing's name (``wp:docPr/@name``)."""
        return self._locate()[1].find(_WP + "docPr").get("name")

    @property
    def alt_text(self) -> str | None:
        """The drawing's alternative text (``wp:docPr/@descr``), or ``None``."""
        return self._locate()[1].find(_WP + "docPr").get("descr")

    @property
    def inline(self) -> bool:
        """Whether the drawing stands in the line (``wp:inline``) rather than floating."""
        return self._locate()[1].tag == _WP + "inline"

    @property
    def size(self) -> tuple[float, float]:
        """``(width, height)`` in points, from ``wp:extent``."""
        extent = self._locate()[1].find(_WP + "extent")
        return int(extent.get("cx")) / 12700, int(extent.get("cy")) / 12700

    @property
    def position(self) -> dict:
        """Where it is: ``{inline: True}``, or its offsets or alignments and what they are
        measured from (points)."""
        from .drawings import read_position

        return read_position(self._locate()[1])

    @property
    def wrapping(self) -> dict:
        """How text goes round it: ``wrap``, ``side``, ``distances``, ``z_order``,
        ``lock_anchor``, ``allow_overlap``, ``layout_in_cell``."""
        from .drawings import read_wrap

        return read_wrap(self._locate()[1])

    @property
    def paragraphs(self) -> list:
        """A text box's paragraphs (its story, ``d:<id>``), edited through the paragraph API."""
        return self.document.text_box_paragraphs(self.id)

    @property
    def text(self) -> str:
        """A text box's text, a line per paragraph ("" for a drawing without text)."""
        try:
            return "\n".join(p.text for p in self.paragraphs)
        except Exception:
            return ""

    @property
    def chart(self):
        """The chart this drawing holds (:meth:`Document.chart`)."""
        return self.document.chart(self.id)

    @property
    def diagram(self):
        """The SmartArt this drawing holds (:meth:`Document.diagram`)."""
        return self.document.diagram(self.id)

    @property
    def members(self) -> list:
        """A group's members, depth first, each ``d:<group>/<member>``; ``[]`` otherwise."""
        from .charts import Member, group_members

        return [Member(self.document, identifier) for identifier, _ in group_members(self._locate()[1], self.id)]

    def float(self, **values):
        """Make the drawing float: :meth:`Document.float_drawing`."""
        return self.document.float_drawing(self.id, **values)

    def make_inline(self):
        """Put a floating drawing in the line again: :meth:`Document.inline_drawing`."""
        return self.document.inline_drawing(self.id)

    def set(self, **values):
        """A floating drawing's wrapping, place and order: :meth:`Document.set_drawing`."""
        return self.document.set_drawing(self.id, **values)

    def move(self, x: float, y: float):
        """Move a floating drawing to ``(x, y)`` points: :meth:`Document.move_drawing`."""
        return self.document.move_drawing(self.id, x, y)

    def resize(self, width: float | None = None, height: float | None = None):
        """Resize the drawing (points; give one side to keep the aspect ratio):
        :meth:`Document.resize_drawing`."""
        return self.document.resize_drawing(self.id, width, height)

    def __repr__(self) -> str:
        return f"<Drawing {self.id} {self.kind}>"


def drawing_kind(frame: Element) -> str:
    """``picture``, ``chart``, ``diagram``, ``text-box``, ``group``, ``canvas``, ``shape`` or
    ``other``, from the frame's ``a:graphicData``."""
    data = frame.find(f"{_A}graphic/{_A}graphicData")
    uri = data.get("uri", "") if data is not None else ""
    if data is not None and data.find(_PIC + "pic") is not None:
        return "picture"
    if uri.endswith("/chart"):
        return "chart"
    if uri.endswith("/diagram"):
        return "diagram"
    if any(n.tag == _W + "txbxContent" for n in frame.iter()):
        return "text-box"
    if uri.endswith("wordprocessingGroup"):
        return "group"
    if uri.endswith("wordprocessingCanvas"):
        return "canvas"
    if uri.endswith("wordprocessingShape"):
        return "shape"
    return "other"


# -- the Document half -----------------------------------------------------------------------


class AnnotationOps:
    """The reading half of notes, comments, revisions, content controls and drawings, on
    :class:`docx_agent.Document`."""

    def _notes(self: "Document") -> list[tuple[str, str, Element]]:
        out = []
        for part in self._parts():
            root = self.package.tree(part)
            if root is None:
                continue
            tag = root.tag
            if tag not in (_W + "footnotes", _W + "endnotes"):
                continue
            prefix = "fn" if tag == _W + "footnotes" else "en"
            for note in root:
                if note.tag in (_W + "footnote", _W + "endnote") and note.get(_W + "type") in (None, "normal"):
                    out.append((part, f"{prefix}:{note.get(_W + 'id')}", note))
        return out

    def notes(self: "Document", kind: str | None = None) -> list[Note]:
        """Every footnote and endnote (``kind`` ``footnote`` or ``endnote`` for one sort)."""
        return [Note(self, identifier) for _, identifier, _ in self._notes()
                if kind is None or Note(self, identifier).kind == kind]

    def note(self: "Document", identifier: str) -> Note:
        """A footnote or endnote by id (``fn:2``, ``en:1``); ``KeyError`` if none."""
        found = Note(self, identifier)
        found._locate()
        return found

    def _comment_records(self: "Document") -> list[dict]:
        main = self.package.document_part()
        parts = self.package.related_parts_of_type(main, "http://schemas.openxmlformats.org/officeDocument/"
                                                   "2006/relationships/comments")
        part = parts[0] if parts and self.package.has_part(parts[0]) else None
        if part is None:
            return []
        durable: dict[str, str] = {}
        extended: dict[str, tuple[str | None, bool]] = {}
        for rel_type, ns in ((REL_COMMENTS_IDS, _W16CID), (REL_COMMENTS_EXTENDED, _W15)):
            related = self.package.related_parts_of_type(main, rel_type)
            root = self.package.tree(related[0]) if related and self.package.has_part(related[0]) else None
            if root is None:
                continue
            for node in root:
                para = (node.get(ns + "paraId") or "").upper()
                if rel_type == REL_COMMENTS_IDS:
                    if node.get(ns + "durableId"):
                        durable[para] = node.get(ns + "durableId")
                else:
                    extended[para] = ((node.get(ns + "paraIdParent") or "").upper() or None,
                                      (node.get(ns + "done") or "0") in ("1", "true"))
        records = []
        index = self._index(part)
        by_last: dict[str, str] = {}
        root = self.package.tree(part)
        for comment in root.findall(_W + "comment"):
            paragraphs = [p for p in comment.iter(_W + "p")]
            last = (paragraphs[-1].get(_ids.PARA_ID) or "").upper() if paragraphs else ""
            identifier = f"c:{durable[last]}" if last in durable else f"c#{comment.get(_W + 'id')}"
            parent_para, done = extended.get(last, (None, False))
            entries = [index.entry_for(p) for p in paragraphs]
            records.append({
                "id": identifier, "w_id": comment.get(_W + "id"), "part": part, "element": comment,
                "author": comment.get(_W + "author"), "initials": comment.get(_W + "initials"),
                "date": comment.get(_W + "date"), "parent_para": parent_para, "done": done,
                "paragraphs": [e.id for e in entries if e is not None]})
            if last:
                by_last[last] = identifier
        for record in records:
            record["parent"] = by_last.get(record.pop("parent_para") or "")
        return records

    def comments(self: "Document", *, replies: bool = True) -> list[Comment]:
        """Every comment, in ``comments.xml``'s order -- replies too (``comment.parent``
        set), unless ``replies=False``, which gives each thread's first comment only::

            for comment in doc.comments(replies=False):
                print(comment.id, comment.anchor.text, comment.text, comment.replies)

        ``comment.anchor.text`` is the text one is attached to;
        ``doc.to_markdown(view="markup")`` shows each in place, as
        ``{>>Author: text<<}<!-- c:... -->``."""
        return [Comment(self, record["id"]) for record in self._comment_records()
                if replies or record["parent"] is None]

    def comment(self: "Document", identifier: str) -> Comment:
        """A comment by id (``c:<durableId>`` or ``c#<w:id>``); ``KeyError`` if none.
        ``doc.comment("c:76FCC91F").anchor.text``."""
        found = Comment(self, identifier)
        found._record()
        return found

    def _comment_anchor(self: "Document", identifier: str, view: str = "current") -> "TextRange | None":
        """The range ``w:commentRangeStart``/``End`` of comment ``identifier`` mark (a
        reply's own markers, else its thread's), or the place of its reference when it has
        no range; ``None`` if no story holds either."""
        records = {record["id"]: record for record in self._comment_records()}
        record = records.get(identifier)
        if record is None:
            raise KeyError(f"no comment {identifier!r}")
        markers: dict[str, tuple[str, dict[str, Element]]] = {}
        for part in self._parts():
            if part == record["part"]:
                continue
            for node in self.package.tree(part).iter(_W + "commentRangeStart", _W + "commentRangeEnd",
                                                     _W + "commentReference"):
                found = markers.setdefault(node.get(_W + "id"), (part, {}))[1]
                found.setdefault(node.tag.rpartition("}")[2], node)
        # A reply Word writes has markers of its own beside its thread's; one with only a
        # reference (or none) is attached where its thread is.
        seen = set()
        fallback = None
        while record is not None and record["id"] not in seen:
            seen.add(record["id"])
            part, found = markers.get(record["w_id"], (None, {}))
            if "commentRangeStart" in found:
                return self._marker_range(part, found["commentRangeStart"], found.get("commentRangeEnd"), view)
            if fallback is None and "commentReference" in found:
                fallback = (part, found["commentReference"])
            record = records.get(record["parent"]) if record["parent"] else None
        if fallback is not None:
            return self._marker_range(fallback[0], fallback[1], fallback[1], view)
        return None

    def _marker_range(self: "Document", part: str, start: Element, end: "Element | None",
                      view: str) -> "TextRange":
        from .ranges import TextRange

        first, first_offset = self._marker_place(part, start, forward=True, view=view)
        last, last_offset = first, first_offset
        if end is not None:
            last, last_offset = self._marker_place(part, end, forward=False, view=view)
        if (last.index, last_offset) < (first.index, first_offset):
            last, last_offset = first, first_offset
        return TextRange(self, first.id, first_offset, last.id, last_offset, view=view)

    def _comment_ids(self: "Document") -> dict[str, str]:
        """``w:id`` -> comment id, for reading references and ranges in the stories."""
        return {record["w_id"]: record["id"] for record in self._comment_records()}

    def _revision_elements(self: "Document") -> list[tuple[str, str, Element]]:
        out = []
        seen: dict[str, int] = {}
        for part in self._parts():
            for node in _ids.live_elements(self.package.tree(part), frozenset(REVISION_KINDS)):
                raw = node.get(_W + "id")
                if raw is None:
                    continue
                count = seen.get(raw, 0)
                seen[raw] = count + 1
                out.append((part, f"rev:{raw}" + (f"#{count}" if count else ""), node))
        return out

    def revisions(self: "Document") -> list[Revision]:
        return [Revision(self, identifier) for _, identifier, _ in self._revision_elements()]

    def revision(self: "Document", identifier: str) -> Revision:
        found = Revision(self, identifier)
        found._locate()
        return found

    def _content_controls(self: "Document") -> list[tuple[str, str, Element]]:
        out = []
        seen: dict[str, int] = {}
        for part in self._parts():
            story = self._story_of(part)
            for number, node in enumerate(_ids.live_elements(self.package.tree(part), frozenset({_W + "sdt"}))):
                raw = _sdt_value(node, "id")
                if raw is None:
                    out.append((part, f"cc@{story}/{number}", node))
                    continue
                count = seen.get(raw, 0)
                seen[raw] = count + 1
                out.append((part, f"cc:{raw}" + (f"#{count}" if count else ""), node))
        return out

    def content_controls(self: "Document") -> list[ContentControl]:
        """Every content control (``cc:<id>``), in story and document order."""
        return [ContentControl(self, identifier) for _, identifier, _ in self._content_controls()]

    def content_control(self: "Document", identifier: str) -> ContentControl:
        """A content control by id; ``KeyError`` if none.  ``doc.content_control("cc:101").text``."""
        found = ContentControl(self, identifier)
        found._locate()
        return found

    def _all_drawings(self: "Document") -> list[tuple[str, str, Element]]:
        """``(part, id, wp:inline | wp:anchor)`` for every drawing with a ``docPr`` id; a
        picture's id is the one :meth:`pictures` gives it."""
        held = self._drawings()  # held: the proxies' id()s key the map below
        pictures = {id(frame): identifier for _, identifier, frame in held}
        out = []
        seen: dict[str, int] = {}
        for part in self._parts():
            frames = _ids.live_elements(self.package.tree(part), frozenset({_ids.WP_INLINE, _ids.WP_ANCHOR}))
            for frame in frames:
                doc_pr = frame.find(_WP + "docPr")
                raw = doc_pr.get("id") if doc_pr is not None else None
                if raw is None:
                    continue
                count = seen.get(raw, 0)
                seen[raw] = count + 1
                if id(frame) in pictures:
                    out.append((part, pictures[id(frame)], frame))
                else:
                    out.append((part, f"d:{raw}" + (f"#{count}" if count else ""), frame))
        del held
        return out

    def drawing(self: "Document", identifier: str):
        """A drawing by id: a :class:`~docx_agent.Picture` for a picture, else a
        :class:`Drawing`; a group's member (``d:<group>/<member>``) a
        :class:`~docx_agent.edit.charts.Member`."""
        if "/" in identifier:
            from .charts import Member, locate

            locate(self, identifier)
            return Member(self, identifier)
        for part, found, frame in self._all_drawings():
            if found == identifier:
                if drawing_kind(frame) == "picture" and any(i == identifier for _, i, _ in self._drawings()):
                    return self.picture(identifier)
                return Drawing(self, identifier)
        raise KeyError(f"no drawing {identifier!r}")
