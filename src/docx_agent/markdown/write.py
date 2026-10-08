"""Markdown in: ``insert_markdown`` writes Markdown as the document's own blocks (ROADMAP.md,
"The Markdown layer"; Phase E2, the write half).

The text is parsed by :func:`docx_agent.markdown.parse.parse_model` (markdown-it-py, the
dialect ``to_markdown`` writes) into the neutral model, and each construct is written with
the style :class:`~docx_agent.markdown.stylemap.StyleMap` names for it -- the same table
the reader reads -- resolved in the document by ``w:name``:

* **Paragraph styles.**  A style is found by its name, an alias or its id
  (:meth:`Styles.find`); a heading also by the outline level a style declares, as the
  reader finds one ("Kop 1" with ``w:outlineLvl`` 0 is ``#``).  A mapped style the
  document lacks is written as Word writes it on first use (the measured definitions of
  :mod:`docx_agent.edit.builtin_styles`), and the result says so.  The document's default
  paragraph style is written as Word writes it: no ``w:pStyle``.
* **Lists.**  An item is a paragraph in the level's list style (List Bullet, List Bullet
  2-5, List Number...) with its own ``w:numPr``: one list instance (``w:num``) per
  Markdown list, at ``w:ilvl`` its depth, over an abstract definition made for the
  insertion -- a copy of the document's own bullet (or numbered) definition when it has a
  nine-level one, else Word's own (E1's) -- so that nesting is ``w:ilvl`` (as the reader
  reads it) and the new lists never count on from, or disturb, the document's.  A numbered
  list restarts at its start number with a ``w:startOverride`` on its instance, Word's
  form, which docx2svg's measured counter honours (N.1-N.4).
* **Inline.**  Emphasis, strong, inline code and links are character styles (Emphasis,
  Strong, HTML Code, Hyperlink: decided).  A run has one character style, so a run with
  several marks takes the first of code, link, strong, emphasis as its style and the rest
  as direct ``w:b``/``w:i`` -- the only direct formatting written besides strikethrough
  (``w:strike``, which has no style) and a thematic break's border.  ``emphasis="direct"``
  writes ``w:i``/``w:b`` throughout instead.
* **Tables** (GFM) in the document's table style -- the one its tables use most, else Table
  Grid -- as Word makes a table (``tools/e2_probe.py``): ``w:tblW`` auto, ``tblLook``
  ``04A0``, the text width shared among the columns; the header row repeats
  (``w:tblHeader``) and a column's alignment is its cells' ``w:jc``.
* **Footnotes** (``[^label]``) are real footnotes, in the footnotes part Word makes with
  the first (and its endnotes part, both with their separators, and the settings'
  ``w:footnotePr``/``w:endnotePr``: measured).
* **Pictures** come from ``images`` (a directory a path must lie under, a mapping or a
  function) and, for ``http(s)`` addresses, only from ``fetch``; ``data:`` URIs are read.
* **Raw HTML** is refused (``html="refuse"``) or kept as literal text (``html="text"``);
  HTML comments -- ``to_markdown``'s ids among them -- are dropped.

The whole insertion is one edit (one undo step).  Tracking, it is one revision group: every
record by the same author at the same date -- paragraphs inserted with their marks, rows
with ``w:trPr/w:ins``, at a container's end in Word's typing form (the previous mark
inserted), so that accepting gives the untracked insertion and rejecting the original.
"""

from __future__ import annotations

import base64
import binascii
import os
import re
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Callable, Mapping

from ..edit import ids as _ids
from ..edit import numbering as _numbering
from ..edit import text as _text
from ..edit.errors import EditError
from ..oxml.xml import Element, append_in_order, insert_in_order, make, remove
from . import model as m
from .parse import parse_model
from .stylemap import DEFAULT, StyleMap, coerce

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
W_P = _W + "p"
W_TBL = _W + "tbl"
W_PPR = _W + "pPr"
W_SECTPR = _W + "sectPr"
W_BODY = _W + "body"
_BLOCKS = (W_P, W_TBL, _W + "sdt", _W + "customXml")

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"

#: What ``at`` may say before the colon.
PLACES = ("after", "before", "replace", "end")
#: Raw HTML: refused, or written as its text.
HTML_MODES = ("refuse", "text")
EMPHASIS_MODES = ("styles", "direct")
#: Word's form for a thematic break: what its AutoFormat makes of a ``---`` line (a bottom
#: border, single, 3/4 pt, one point from the text) -- not measured here: AppleScript's
#: typing does not trigger AutoFormat (``tools/e2_probe.py``).
RULE_BORDER = {"w:val": "single", "w:sz": "6", "w:space": "1", "w:color": "auto"}
#: Characters XML 1.0 cannot hold: written as U+FFFD, as markdown-it writes NUL.
_INVALID = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
#: The text width a table shares when the section says none (A4, one-inch margins).
_DEFAULT_TEXT_WIDTH = 9026
_MAX_LEVEL = 8

@dataclass
class Options:
    html: str = "refuse"
    emphasis: str = "styles"
    images: object = None
    fetch: Callable[[str], bytes] | None = None


@dataclass
class _Place:
    """Where the blocks go: ``container``'s children, before ``before`` (``None``: at the
    end, before a body's final ``w:sectPr``), with ``replaced`` blocks to delete."""

    part: str
    container: Element
    before: Element | None
    anchor_id: str | None
    replaced: list[str] = field(default_factory=list)
    #: The paragraph index the insertion is at (what a structural edit there renumbers).
    position: int = 0
    #: ``end`` of a body that is one empty paragraph: replacing it, only if anything is written.
    lone: bool = False


# -- the entry point -------------------------------------------------------------------------


def insert_markdown(document: "Document", md: str, *, at: str = "end", style_map: StyleMap | None = None,
                    images=None, fetch: Callable[[str], bytes] | None = None, html: str = "refuse",
                    emphasis: str = "styles") -> "EditResult":
    """See :meth:`docx_agent.Document.insert_markdown`."""
    from ..edit.document import EditResult

    if not isinstance(md, str):
        raise EditError("insert_markdown takes Markdown text")
    if html not in HTML_MODES:
        raise EditError(f"html is one of {', '.join(HTML_MODES)}")
    if emphasis not in EMPHASIS_MODES:
        raise EditError(f"emphasis is one of {', '.join(EMPHASIS_MODES)}")
    options = Options(html, emphasis, images, fetch)
    style_map = coerce(style_map)
    blocks = parse_model(md)
    notes = {b.label: b for b in blocks if isinstance(b, m.FootnoteDef)}
    blocks = [b for b in blocks if not isinstance(b, (m.FootnoteDef, m.Comment))]
    lone = None
    if at in ("end", "end:body"):
        lone = _lone_empty_paragraph(document)
        if lone is not None and blocks and not isinstance(blocks[0], m.Table):
            # Ending in a table, the empty paragraph stays as the one Word needs after it.
            at = f"before:{lone}" if isinstance(blocks[-1], m.Table) else f"replace:{lone}"
    place = _locate(document, at)
    place.lone = lone is not None
    _check(blocks, notes, options, place)
    pictures = _pictures(blocks, notes, options)

    tracking = document._active_tracking()
    try:
        return _insert(document, place, blocks, notes, options, pictures, style_map, tracking)
    except _Nothing:
        return EditResult(None, changed=False)


class _Nothing(Exception):
    """The Markdown holds nothing to write (only link definitions, comments, an empty
    quote): the edit is rolled back and reports no change."""


def _insert(document, place, blocks, notes, options, pictures, style_map, tracking) -> "EditResult":
    from ..edit.document import EditResult

    removed: list[str] = []
    lone_properties = None
    with document._edit():
        renames = document._prepare(place.part, [], place.position)
        if place.lone and place.replaced and tracking is not None:
            # Typing into a new document's empty paragraph, tracked (Word's form at a
            # container's end, E3): the new paragraphs' marks are inserted but the last
            # one's, which is the old paragraph's, its properties a change from the old
            # ones -- rejecting gives the one empty paragraph back, accepting the text.
            lone = document._resolve(place.replaced[0])[1].element
            lone_properties = _track_clean(lone.find(W_PPR))
            place.before = _next_block(lone)
            remove(lone)
            document._invalidate()
            removed.append(place.replaced[0])
            place.replaced = []
        if place.replaced and tracking is not None:
            # Tracked: the replaced blocks are deleted where they are, and the new ones go
            # after them, so that a container's last mark stays the original one.
            last = document._resolve(place.replaced[-1])[1].element
            following = _next_block(last)
            for identifier in place.replaced:
                removed += document.delete_block(identifier).removed
            place.before = following
        writer = Writer(document, place.part, style_map or DEFAULT, options, pictures, notes)
        elements = writer.blocks(blocks)
        if not elements and (not place.replaced or place.lone):
            raise _Nothing()
        _place(place, elements)
        if place.replaced and tracking is None:
            for identifier in place.replaced:
                removed += document.delete_block(identifier).removed
        if elements and _is_last(elements[-1]) and elements[-1].tag == W_TBL:
            # A container ends in a paragraph (a cell must; Word gives the body one).
            trailing = writer.paragraph([], None)
            elements[-1].addnext(trailing)
            elements.append(trailing)
        created = writer.stamp(elements)
        if tracking is not None and elements:
            writer.track(elements, tracking, old_mark=lone_properties)
        writer.finish()
        document.package.mark_dirty(place.part)
    first = created[0] if created else None
    return EditResult(first, created=created + writer.created_notes, renamed=renames, removed=removed,
                      warnings=writer.warnings, blocks=writer.top_ids)


# -- where ------------------------------------------------------------------------------------


def _locate(document: "Document", at: str) -> _Place:
    """``at``: ``end`` or ``end:<story>``; ``after:<id>``, ``before:<id>``, a bare id (after
    it); ``replace:<id>`` or ``replace:<id>..<id>``."""
    if not isinstance(at, str) or not at:
        raise EditError("at is 'end', 'end:<story>', 'after:<id>', 'before:<id>' or 'replace:<id>[..<id>]'")
    kind, _, rest = at.partition(":")
    if kind not in PLACES:
        kind, rest = "after", at
    if kind == "end":
        story = rest or "body"
        part = document._part_for_story(story)
        if part is None:
            raise EditError(f"no story {story!r}")
        root = document.package.tree(part)
        container = root.find(W_BODY) if root.find(W_BODY) is not None else root
        if root.tag in (_W + "footnotes", _W + "endnotes", _W + "comments"):
            raise EditError(f"{story} holds notes or comments: insert after one of their paragraphs")
        return _Place(part, container, None, None, position=len(document._index(part).paragraphs))
    if kind == "replace":
        first, sep, last = rest.partition("..")
        part, entry = document._resolve(first)
        elements = [entry.element]
        if sep:
            other_part, other = document._resolve(last)
            if other_part != part or other.element.getparent() is not entry.element.getparent():
                raise EditError(f"{rest!r}: a range replaced is blocks of one container")
            siblings = [c for c in entry.element.getparent() if isinstance(c.tag, str) and c.tag in _BLOCKS]
            a, b = siblings.index(entry.element), siblings.index(other.element)
            if b < a:
                a, b = b, a
            elements = siblings[a:b + 1]
        ids = [_id_of(document, part, e) for e in elements]
        return _Place(part, elements[0].getparent(), elements[0], ids[0], replaced=ids,
                      position=_position(document, part, elements[0], False))
    part, entry = document._resolve(rest)
    element = entry.element
    if kind == "after":
        return _Place(part, element.getparent(), _next_sibling(element), rest,
                      position=_position(document, part, element, True))
    return _Place(part, element.getparent(), element, rest, position=_position(document, part, element, False))


def _track_clean(properties: Element | None) -> Element:
    from ..revisions import track as _track

    return _track.clean_properties(properties, "w:pPr")


def _lone_empty_paragraph(document: "Document") -> str | None:
    """The id of the body's only block when it is an empty paragraph -- no text, runs,
    bookmarks or section break, as a new document's is -- else ``None``."""
    part = document.package.document_part()
    body = document.package.tree(part).find(W_BODY)
    if body is None:
        return None
    blocks = [child for child in body if isinstance(child.tag, str) and child.tag != W_SECTPR]
    if len(blocks) != 1 or blocks[0].tag != W_P:
        return None
    paragraph = blocks[0]
    if any(isinstance(child.tag, str) and child.tag != W_PPR for child in paragraph):
        return None
    if paragraph.find(f"{W_PPR}/{W_SECTPR}") is not None or paragraph.find(f"{W_PPR}/{_W}numPr") is not None:
        return None
    entries = document._index(part).paragraphs
    return entries[0].id if len(entries) == 1 and entries[0].element is paragraph else None


def _id_of(document: "Document", part: str, element: Element) -> str:
    for entry in document._index(part).paragraphs:
        if entry.element is element:
            return entry.id
    for entry in document._index(part).tables:
        if entry.element is element:
            return entry.id
    raise EditError("a block replaced must be a paragraph or a table")


def _position(document: "Document", part: str, element: Element, after: bool) -> int:
    for entry in document._index(part).paragraphs:
        if entry.element is element:
            return entry.index
    return document._block_position(part, element, after)


def _next_sibling(element: Element) -> Element | None:
    return element.getnext()


def _next_block(element: Element) -> Element | None:
    following = element.getnext()
    while following is not None and not (isinstance(following.tag, str) and following.tag in _BLOCKS + (W_SECTPR,)):
        following = following.getnext()
    return following


def _place(place: _Place, elements: list[Element]) -> None:
    before = place.before
    if before is None:
        last = place.container.find(W_SECTPR) if place.container.tag == W_BODY else None
        if last is not None:
            for element in elements:
                last.addprevious(element)
            return
        for element in elements:
            place.container.append(element)
        return
    for element in elements:
        before.addprevious(element)


def _is_last(element: Element) -> bool:
    following = element.getnext()
    while following is not None:
        if isinstance(following.tag, str) and following.tag in _BLOCKS:
            return False
        following = following.getnext()
    return True


# -- checks before anything changes ------------------------------------------------------------


def _walk_inline(blocks: list):
    """Every inline item of ``blocks`` (and their nested blocks), with its block."""
    for block in blocks:
        if isinstance(block, (m.Paragraph, m.Heading)):
            for item in block.inline:
                yield block, item
        elif isinstance(block, m.Quote):
            yield from _walk_inline(block.blocks)
        elif isinstance(block, m.ListBlock):
            for item in block.items:
                yield from _walk_inline(item)
        elif isinstance(block, m.Table):
            for row in [block.header] + block.rows:
                for cell in row:
                    for item in cell:
                        yield block, item
        elif isinstance(block, m.FootnoteDef):
            yield from _walk_inline(block.blocks)


def _check(blocks: list, notes: dict, options: Options, place: _Place) -> None:
    every = blocks + list(notes.values())
    if options.html == "refuse":
        for block in _walk_blocks(every):
            if isinstance(block, m.HtmlBlock):
                raise EditError(f"raw HTML is not inserted (html='text' keeps it as text): {block.text.strip()[:60]!r}")
        for _, item in _walk_inline(every):
            if isinstance(item, m.Html) and not item.text.startswith("<!--"):
                raise EditError(f"raw HTML is not inserted (html='text' keeps it as text): {item.text[:60]!r}")
    root = place.container.getroottree().getroot()
    if root.tag in (_W + "footnotes", _W + "endnotes", _W + "comments", _W + "hdr", _W + "ftr"):
        for _, item in _walk_inline(blocks):
            if isinstance(item, m.NoteRef):
                raise EditError("a footnote cannot be inserted outside the body")


def _walk_blocks(blocks: list):
    for block in blocks:
        yield block
        if isinstance(block, m.Quote):
            yield from _walk_blocks(block.blocks)
        elif isinstance(block, m.ListBlock):
            for item in block.items:
                yield from _walk_blocks(item)
        elif isinstance(block, m.FootnoteDef):
            yield from _walk_blocks(block.blocks)


# -- pictures ---------------------------------------------------------------------------------


def _pictures(blocks: list, notes: dict, options: Options) -> dict[str, bytes]:
    """Every picture's bytes, by its address, read before anything changes."""
    from ..edit.pictures import PictureError, image_format, natural_size

    out: dict[str, bytes] = {}
    for _, item in _walk_inline(blocks + list(notes.values())):
        if not isinstance(item, m.Image) or item.src in out:
            continue
        data = _read_image(item.src, options)
        try:
            image_format(data)
            natural_size(data)
        except PictureError as error:
            raise EditError(f"![{item.alt}]({item.src}): {error}") from None
        out[item.src] = data
    return out


def _read_image(src: str, options: Options) -> bytes:
    scheme = urllib.parse.urlsplit(src).scheme.lower()
    if scheme in ("http", "https", "ftp"):
        if options.fetch is None:
            raise EditError(f"{src!r} is remote: pass fetch= to read it (nothing is fetched otherwise)")
        return _bytes(options.fetch(src), src)
    if scheme == "data":
        header, _, payload = src.partition(",")
        try:
            if header.endswith(";base64"):
                return base64.b64decode(urllib.parse.unquote(payload), validate=False)
            return urllib.parse.unquote_to_bytes(payload)
        except (binascii.Error, ValueError):
            raise EditError(f"{src[:40]!r}... is not a readable data URI") from None
    source = options.images
    if source is None:
        raise EditError(f"{src!r}: pass images= (a directory, a mapping or a function) to insert pictures")
    if callable(source):
        return _bytes(source(src), src)
    if isinstance(source, Mapping):
        if src not in source:
            raise EditError(f"{src!r} is not among the images given")
        return _bytes(source[src], src)
    base = Path(os.fspath(source)).resolve()
    path = urllib.parse.unquote(urllib.parse.urlsplit(src).path) if scheme == "file" else urllib.parse.unquote(src)
    target = (base / path).resolve()
    if base != target and base not in target.parents:
        raise EditError(f"{src!r} is outside the images directory {str(base)!r}")
    if not target.is_file():
        raise EditError(f"{src!r}: no such file under {str(base)!r}")
    return target.read_bytes()


def _bytes(value, src: str) -> bytes:
    if not isinstance(value, (bytes, bytearray)):
        raise EditError(f"{src!r}: the image provider gave {type(value).__name__}, not bytes")
    return bytes(value)


# -- the writer -------------------------------------------------------------------------------


class Writer:
    """Model blocks as elements of one part, with the styles, lists, links, pictures and
    notes they need."""

    def __init__(self, document: "Document", part: str, style_map: StyleMap, options: Options,
                 pictures: dict[str, bytes], notes: dict) -> None:
        self.document = document
        self.part = part
        self.map = style_map
        self.options = options
        self.pictures = pictures
        self.notes = notes
        self.warnings: list[str] = []
        self._styles: dict[tuple, str | None] = {}
        self._abstracts: dict[str, int] = {}
        self._text_width: int | None = None
        #: Footnote label -> the w:footnote written for it.
        self._footnotes: dict[str, Element] = {}
        self._footnote_part: str | None = None
        self.created_notes: list[str] = []
        self.top_ids: list[str] = []
        #: Parts a picture was written into (their roots declare wp14).
        self._drawings: set[str] = set()
        self._quote = 0
        #: Writing a footnote's blocks: its paragraphs are footnote text, not body text.
        self._note = 0
        self._snapshot: list | None = None

    # -- styles -----------------------------------------------------------------------------

    def style(self, construct: str, level: int = 0) -> str | None:
        """The style id ``construct`` is written with (``None``: none, or the default
        paragraph style, which Word does not name)."""
        key = (construct, level)
        if key in self._styles:
            return self._styles[key]
        rule = self.map.style_for(construct, level)
        if rule is None or rule.style is None:
            self._styles[key] = None
            return None
        kind = rule.kind
        styles = self.document.styles
        found = None
        if construct == "table" and rule == DEFAULT.style_for("table"):
            # The default row: the style the document's tables use most, else Table Grid.
            found = self._table_style_in_use()
        found = found or self._find(rule.style, kind)
        if found is None and construct == "heading":
            found = self._heading_by_outline(level)
        if found is None and construct in ("paragraph", "table_cell") and rule.style.casefold() == "normal":
            # No Normal by that name (no styles at all): the document's default applies,
            # which Word writes without a w:pStyle.
            self._styles[key] = None
            return None
        if found is not None:
            style_id = found.id
        else:
            try:
                style_id = styles._ensure(rule.style, kind)
            except EditError as error:
                raise EditError(f"no {kind} style for {construct}: {error}") from None
            self._snapshot = None
            self.warnings.append(f"style added: {rule.style} ({style_id}), as Word writes it")
        if kind == "paragraph":
            default = next((s for s in self._all() if s.kind == "paragraph" and s.default), None)
            if default is not None and default.id == style_id:
                style_id = None
        self._styles[key] = style_id
        return style_id

    def _all(self) -> list:
        if self._snapshot is None:
            self._snapshot = list(self.document.styles)
        return self._snapshot

    def _find(self, name: str, kind: str):
        """:meth:`Styles.find` over one reading of the styles (it reads them all per call)."""
        styles = [s for s in self._all() if s.kind == kind]
        wanted = name.casefold()
        for test in (lambda s: (s.name or "").casefold() == wanted,
                     lambda s: any(a.casefold() == wanted for a in s.aliases),
                     lambda s: s.id == name,
                     lambda s: s.id.casefold() == wanted):
            for style in styles:
                if test(style):
                    return style
        return None

    def _heading_by_outline(self, level: int):
        from .read import Reader

        reader = Reader(self.document, "final", self.map)
        for style in self._all():
            if style.kind == "paragraph" and self.map.heading_level(style.name, reader.outline_level(style.id)) == level:
                return style
        return None

    def _table_style_in_use(self):
        """The table style the document's tables use most (Table Grid when they use none)."""
        counts: dict[str, int] = {}
        for part in self.document._parts():
            for node in self.document.package.tree(part).iter(_W + "tblStyle"):
                value = node.get(_W + "val")
                if value:
                    counts[value] = counts.get(value, 0) + 1
        for style_id, _ in sorted(counts.items(), key=lambda kv: -kv[1]):
            for style in self._all():
                if style.id == style_id and style.kind == "table":
                    return style
        return None

    # -- blocks -----------------------------------------------------------------------------

    def blocks(self, blocks: list) -> list[Element]:
        out: list[Element] = []
        for block in blocks:
            out += self.block(block)
        return out

    def block(self, block) -> list[Element]:
        if isinstance(block, m.Paragraph):
            if not block.inline:
                return []
            return [self.paragraph(block.inline, self.style("blockquote") if self._quote else self.body_style())]
        if isinstance(block, m.Heading):
            return [self.paragraph(block.inline, self.style("heading", block.level))]
        if isinstance(block, m.CodeBlock):
            return self.code(block.text)
        if isinstance(block, m.Quote):
            self._quote += 1
            try:
                return self.blocks(block.blocks)
            finally:
                self._quote -= 1
        if isinstance(block, m.ListBlock):
            return self.list(block, 0)
        if isinstance(block, m.Table):
            return [self.table(block)]
        if isinstance(block, m.Rule):
            paragraph = self.paragraph([], self.body_style())
            properties = _properties(paragraph)
            borders = append_in_order(properties, make("w:pBdr"))
            borders.append(make("w:bottom", **{k.replace(":", "__"): v for k, v in RULE_BORDER.items()}))
            return [paragraph]
        if isinstance(block, m.HtmlBlock):
            lines = block.text.rstrip("\n").split("\n")
            return [self.paragraph([m.Text(line)], self.body_style()) for line in lines]
        if isinstance(block, m.Comment):
            return []
        raise EditError(f"cannot insert {type(block).__name__}")

    def body_style(self) -> str | None:
        """A plain paragraph's style: ``paragraph`` in the body, ``footnote_text`` in a note."""
        return self.style("footnote_text") if self._note else self.style("paragraph")

    def code(self, text: str) -> list[Element]:
        style = self.style("code_block")
        body = text[:-1] if text.endswith("\n") else text
        return [self.paragraph([m.Text(line)] if line else [], style) for line in body.split("\n")]

    def list(self, block: m.ListBlock, depth: int) -> list[Element]:
        level = min(depth, _MAX_LEVEL)
        kind = "number" if block.ordered else "bullet"
        construct = "ordered_list" if block.ordered else "bullet_list"
        style = self.style(construct, depth + 1)
        num_id = _numbering.add_num(self.document, self._abstract(kind),
                                    {level: block.start} if block.ordered else None)
        out: list[Element] = []
        for item in block.items:
            first = item[0] if item else None
            rest = item[1:] if isinstance(first, m.Paragraph) else item
            inline = first.inline if isinstance(first, m.Paragraph) else []
            paragraph = self.paragraph(inline, style)
            _numbering.set_numbering(paragraph, num_id, level)
            out.append(paragraph)
            for inner in rest:
                if isinstance(inner, m.ListBlock):
                    out += self.list(inner, depth + 1)
                else:
                    out += self.block(inner)
        return out

    def _abstract(self, kind: str) -> int:
        """The abstract definition this insertion's ``kind`` lists use: a copy of the
        document's own nine-level one, else Word's own."""
        if kind in self._abstracts:
            return self._abstracts[kind]
        numbering = _numbering.Numbering(self.document)
        source = _document_definition(self.document, kind)
        if source is not None:
            abstract_id = _numbering.copy_abstract(self.document, source)
        else:
            abstract_id = max(numbering.abstracts(), default=-1) + 1
            seed = f"markdown\0{kind}\0{abstract_id}"
            node = _numbering.builtin_abstract(kind, abstract_id, _numbering._nsid(self.document, seed + "nsid"),
                                               _numbering._nsid(self.document, seed + "tmpl"))
            _numbering.add_abstract(self.document, node)
        self.document.package.mark_dirty(_numbering.Numbering(self.document).part)
        self._abstracts[kind] = abstract_id
        return abstract_id

    def paragraph(self, inline: list, style: str | None, *, align: str | None = None) -> Element:
        paragraph = make("w:p")
        if style is not None:
            properties = _properties(paragraph)
            append_in_order(properties, make("w:pStyle", **{"w:val": style}))
        if align is not None:
            append_in_order(_properties(paragraph), make("w:jc", **{"w:val": align}))
        for node in self.inline(inline):
            paragraph.append(node)
        return paragraph

    def table(self, block: m.Table) -> Element:
        rows = [block.header] + block.rows
        columns = max((len(r) for r in rows), default=0) or 1
        style = self.style("table")
        width = self.text_width()
        shares = [width // columns + (1 if k >= columns - width % columns else 0) for k in range(columns)]
        table = make("w:tbl")
        properties = append_in_order(table, make("w:tblPr"))
        if style is not None:
            append_in_order(properties, make("w:tblStyle", **{"w:val": style}))
        append_in_order(properties, make("w:tblW", **{"w:w": "0", "w:type": "auto"}))
        append_in_order(properties, make("w:tblLook", **{"w:val": "04A0", "w:firstRow": "1", "w:lastRow": "0",
                                                          "w:firstColumn": "1", "w:lastColumn": "0",
                                                          "w:noHBand": "0", "w:noVBand": "1"}))
        grid = append_in_order(table, make("w:tblGrid"))
        for share in shares:
            grid.append(make("w:gridCol", **{"w:w": str(share)}))
        aligns = list(block.aligns) + [None] * columns
        for k, row in enumerate(rows):
            tr = make("w:tr")
            table.append(tr)
            if k == 0:
                append_in_order(tr, make("w:trPr")).append(make("w:tblHeader"))
            for c in range(columns):
                cell = make("w:tc")
                tr.append(cell)
                append_in_order(cell, make("w:tcPr")).append(make("w:tcW", **{"w:w": str(shares[c]), "w:type": "dxa"}))
                inline = row[c] if c < len(row) else []
                cell.append(self.paragraph(inline, self.style("table_cell"), align=aligns[c]))
        return table

    def text_width(self) -> int:
        """The text width of the body's section (twentieths of a point)."""
        if self._text_width is None:
            self._text_width = _DEFAULT_TEXT_WIDTH
            root = self.document.package.tree(self.document.package.document_part())
            body = root.find(W_BODY)
            section = body.find(W_SECTPR) if body is not None else None
            if section is not None:
                size, margins = section.find(_W + "pgSz"), section.find(_W + "pgMar")
                try:
                    width = int(size.get(_W + "w")) - int(margins.get(_W + "left") or 0) \
                        - int(margins.get(_W + "right") or 0) - int(margins.get(_W + "gutter") or 0)
                    if width > 0:
                        self._text_width = width
                except (AttributeError, TypeError, ValueError):
                    pass
        return self._text_width

    # -- inline -----------------------------------------------------------------------------

    def inline(self, items: list) -> list[Element]:
        """Runs (and hyperlinks holding runs) for ``items``."""
        out: list[Element] = []
        link: Element | None = None
        link_href: str | None = None
        for item in items:
            href = getattr(item, "href", None)
            if isinstance(item, m.Html) and item.text.startswith("<!--"):
                continue
            nodes = self.inline_item(item)
            if not nodes:
                continue
            if href is not None and href != "":
                if link is None or href != link_href:
                    link = self._link(href)
                    link_href = href
                    out.append(link)
                for node in nodes:
                    link.append(node)
            else:
                link = link_href = None
                out += nodes
        return out

    def _link(self, href: str) -> Element:
        link = make("w:hyperlink")
        if href.startswith("#"):
            link.set(_W + "anchor", urllib.parse.unquote(href[1:]))
        else:
            link.set(_R + "id", self.document.package.add_external_relationship(self.part, REL + "hyperlink", href))
        link.set(_W + "history", "1")
        return link

    def inline_item(self, item) -> list[Element]:
        if isinstance(item, m.Text):
            text = _INVALID.sub("�", item.text.replace("\n", " "))
            if not text:
                return []
            return [_text.make_run(text, self.run_properties(item.marks, item.href is not None and item.href != ""))]
        if isinstance(item, m.Html):
            return [_text.make_run(_INVALID.sub("�", item.text.replace("\n", " ")), None)]
        if isinstance(item, m.Break):
            run = make("w:r")
            run.append(make("w:br"))
            return [run]
        if isinstance(item, m.Image):
            return [self.picture(item)]
        if isinstance(item, m.NoteRef):
            if self._footnote_part is not None and self.part == self._footnote_part:
                # A note cannot hold a note: the reference stays as text.
                return [_text.make_run(f"[^{item.label}]", None)]
            return [self.note_reference(item.label)]
        return []

    def run_properties(self, marks, link: bool) -> Element | None:
        marks = set(marks)
        properties = make("w:rPr")
        direct = self.options.emphasis == "direct"
        style = None
        if "code" in marks:
            style = self.style("code")
            marks.discard("code")
        elif link:
            style = self.style("link")
        elif "strong" in marks and not direct:
            style = self.style("strong")
            marks.discard("strong")
        elif "em" in marks and not direct:
            style = self.style("emphasis")
            marks.discard("em")
        if style is not None:
            append_in_order(properties, make("w:rStyle", **{"w:val": style}))
        if "strong" in marks:
            append_in_order(properties, make("w:b"))
            append_in_order(properties, make("w:bCs"))
        if "em" in marks:
            append_in_order(properties, make("w:i"))
            append_in_order(properties, make("w:iCs"))
        if "strike" in marks:
            append_in_order(properties, make("w:strike"))
        return properties if len(properties) else None

    def picture(self, item: m.Image) -> Element:
        from ..edit.pictures import _drawing, natural_size

        data = self.pictures[item.src]
        document = self.document
        rid = document._media(self.part, data)
        number = document._next_doc_pr()
        width, height = natural_size(data)
        limit = self.text_width() / 20.0
        if width > limit:
            width, height = limit, height * limit / width
        anchor_id, edit_id = _ids.generate(self.part + "\0picture", document._used(), 2)
        drawing = _drawing(rid, number, item.title or f"Picture {number}", item.alt or None, (width, height),
                           anchor_id, edit_id)
        run = make("w:r")
        run.append(drawing)
        self._drawings.add(self.part)
        return run

    # -- footnotes --------------------------------------------------------------------------

    def note_reference(self, label: str) -> Element:
        """A footnote reference; the footnote is written on its first reference (a second
        reference to one label is a second footnote with the same text: Word's notes have
        one reference each)."""
        definition = self.notes.get(label)
        part, root = self._footnotes_part()
        note_id = max([int(n.get(_W + "id")) for n in root.findall(_W + "footnote")
                       if (n.get(_W + "id") or "").lstrip("-").isdigit()] + [0]) + 1
        note = make("w:footnote", **{"w:id": str(note_id)})
        root.append(note)
        saved = self.part, self._quote, self._note
        self.part, self._quote, self._note = part, 0, 1
        try:
            blocks = list(definition.blocks) if definition is not None else []
            elements = self.blocks([b for b in blocks if not isinstance(b, m.Comment)])
        finally:
            self.part, self._quote, self._note = saved
        text_style = self.style("footnote_text")
        for element in elements:
            if element.tag == W_P:
                style = element.find(f"{W_PPR}/{_W}pStyle")
                if style is None and text_style is not None:
                    append_in_order(_properties(element), make("w:pStyle", **{"w:val": text_style}))
        if not elements or elements[0].tag != W_P:
            elements.insert(0, self.paragraph([], text_style))
        reference_style = self.style("footnote_reference")
        mark = make("w:r")
        if reference_style is not None:
            append_in_order(mark, make("w:rPr")).append(make("w:rStyle", **{"w:val": reference_style}))
        mark.append(make("w:footnoteRef"))
        space = _text.make_run(" ", None)
        first = elements[0]
        properties = first.find(W_PPR)
        at = 1 if properties is not None else 0
        first.insert(at, mark)
        first.insert(at + 1, space)
        for element in elements:
            note.append(element)
        self._footnotes.setdefault(label, note)
        self.created_notes.append(f"fn:{note_id}")
        run = make("w:r")
        if reference_style is not None:
            append_in_order(run, make("w:rPr")).append(make("w:rStyle", **{"w:val": reference_style}))
        run.append(make("w:footnoteReference", **{"w:id": str(note_id)}))
        return run

    def _footnotes_part(self) -> tuple[str, Element]:
        if self._footnote_part is not None:
            return self._footnote_part, self.document.package.tree(self._footnote_part)
        document = self.document
        package = document.package
        main = package.document_part()
        found = package.related_parts_of_type(main, REL + "footnotes")
        if found and package.has_part(found[0]):
            part = found[0]
        else:
            part = _new_notes_part(document, "footnote")
            if not package.related_parts_of_type(main, REL + "endnotes"):
                _new_notes_part(document, "endnote")
        self._footnote_part = part
        package.mark_dirty(part)
        return part, package.tree(part)

    # -- after placing ----------------------------------------------------------------------

    def stamp(self, elements: list[Element]) -> list[str]:
        """Give every new paragraph and row its paraId and textId; the ids, in document
        order (a table, then its cells' paragraphs), and :attr:`top_ids`."""
        document = self.document
        used = document._used()
        story = document._story_of(self.part)
        prefix = "" if story == "body" else f"{story}/"
        out: list[str] = []
        for element in elements:
            if element.tag == W_P:
                para_id, text_id = _ids.generate(self.part, used, 2)
                _ids.stamp(element, para_id, text_id)
                out.append(f"{prefix}p:{para_id}")
                self.top_ids.append(out[-1])
            else:
                first_row = None
                nested: list[str] = []
                for node in element.iter(W_P, _W + "tr"):
                    para_id, text_id = _ids.generate(self.part + ("\0row" if node.tag == _W + "tr" else ""), used, 2)
                    _ids.stamp(node, para_id, text_id)
                    if node.tag == _W + "tr":
                        first_row = first_row or para_id
                    else:
                        nested.append(f"{prefix}p:{para_id}")
                out.append(f"{prefix}t:{first_row}")
                self.top_ids.append(out[-1])
                out += nested
        _ids.ensure_w14(document.package.tree(self.part), drawings=self.part in self._drawings)
        if self._footnote_part is not None:
            notes_ids: list[str] = []
            for note in self._footnotes_list():
                for node in note.iter(W_P, _W + "tr"):
                    para_id, text_id = _ids.generate(self._footnote_part, used, 2)
                    _ids.stamp(node, para_id, text_id)
                    if node.tag == W_P:
                        notes_ids.append(f"{document._story_of(self._footnote_part)}/p:{para_id}")
            _ids.ensure_w14(document.package.tree(self._footnote_part),
                            drawings=self._footnote_part in self._drawings)
            out += notes_ids
        return out

    def _footnotes_list(self) -> list[Element]:
        root = self.document.package.tree(self._footnote_part)
        wanted = {n.split(":")[1] for n in self.created_notes}
        return [n for n in root.findall(_W + "footnote") if n.get(_W + "id") in wanted]

    def track(self, elements: list[Element], tracking, *, old_mark: Element | None = None) -> None:
        """Record the insertion as one revision group: one author, one date.  ``old_mark``:
        the properties of a paragraph the insertion replaced at its container's end, whose
        mark the last new paragraph takes (untracked, its properties recorded as a change)."""
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp

        stamp = Stamp(self.document, tracking)
        part = self.part
        for element in elements:
            if element.tag == W_P:
                _insert_content(element, stamp, part)
                _track.mark(element, "ins", stamp, part)
            else:
                _track_table(element, stamp, part)
        last = elements[-1]
        if last.tag == W_P and _track.is_last_in_container(last):
            previous = elements[0].getprevious()
            while previous is not None and not (isinstance(previous.tag, str) and previous.tag in _BLOCKS):
                previous = previous.getprevious()
            if previous is not None and previous.tag == W_P and _track.mark_record(previous) is None \
                    and previous.find(f"{W_PPR}/{W_SECTPR}") is None:
                # Word's typing form at a container's end (measured, E3): the mark before
                # the insertion is the inserted one; the last new paragraph has the old
                # mark, and its properties are recorded as a change from that paragraph's.
                record = _track.mark_record(last)
                run_properties = record.getparent()
                remove(record)
                if not len(run_properties):
                    remove(run_properties)
                _track.mark(previous, "ins", stamp, part)
                change = previous.find(f"{W_PPR}/{_W}pPrChange/{W_PPR}")
                old = _track.clean_properties(change if change is not None else previous.find(W_PPR), "w:pPr")
                _track.record_paragraph_change(last, old, stamp, part)
            elif old_mark is not None:
                record = _track.mark_record(last)
                run_properties = record.getparent()
                remove(record)
                if not len(run_properties):
                    remove(run_properties)
                _track.record_paragraph_change(last, old_mark, stamp, part)
        if self._footnote_part is not None:
            notes_part = self._footnote_part
            for note in self._footnotes_list():
                paragraphs = [c for c in note if isinstance(c.tag, str)]
                for k, element in enumerate(paragraphs):
                    if element.tag == W_P:
                        _insert_content(element, stamp, notes_part)
                        if k + 1 < len(paragraphs):
                            _track.mark(element, "ins", stamp, notes_part)
                    else:
                        _track_table(element, stamp, notes_part)
        stamp.finish()

    def finish(self) -> None:
        if self._footnote_part is not None:
            self.document.package.mark_dirty(self._footnote_part)
        numbering = _numbering.Numbering(self.document)
        if self._abstracts and numbering.part is not None:
            self.document.package.mark_dirty(numbering.part)


def _properties(paragraph: Element) -> Element:
    properties = paragraph.find(W_PPR)
    if properties is None:
        properties = make("w:pPr")
        paragraph.insert(0, properties)
    return properties


def _insert_content(paragraph: Element, stamp, part: str) -> None:
    """A new paragraph's runs inside ``w:ins``: at paragraph level in one container per
    stretch, inside a hyperlink within it (a revision container holds runs, never a
    hyperlink: E3)."""
    current: Element | None = None
    for child in list(paragraph):
        if not isinstance(child.tag, str) or child.tag == W_PPR:
            continue
        if child.tag == _W + "hyperlink":
            current = None
            inner = stamp.make("w:ins", part)
            for run in list(child):
                inner.append(run)
            child.append(inner)
            continue
        if current is None:
            current = stamp.make("w:ins", part)
            child.addprevious(current)
        current.append(child)


def _track_table(table: Element, stamp, part: str) -> None:
    from ..revisions import track as _track

    for row in table.iter(_W + "tr"):
        properties = row.find(_W + "trPr")
        if properties is None:
            properties = make("w:trPr")
            row.insert(0, properties)
            insert_in_order(row, properties)
        insert_in_order(properties, stamp.make("w:ins", part))
    for paragraph in table.iter(W_P):
        _insert_content(paragraph, stamp, part)
        _track.mark(paragraph, "ins", stamp, part)


def _document_definition(document: "Document", kind: str) -> int | None:
    """The abstract definition the document's own ``kind`` lists use most: nine levels,
    every one a bullet (or none), no level tied to a style, not a list style's."""
    numbering = _numbering.Numbering(document)
    if numbering.root is None:
        return None
    counts: dict[int, int] = {}
    for part in document._parts():
        for paragraph in document.package.tree(part).iter(W_P):
            found = paragraph.find(f"{W_PPR}/{_W}numPr/{_W}numId")
            if found is None:
                continue
            try:
                abstract = numbering.abstract_of(int(found.get(_W + "val")))
            except (TypeError, ValueError):
                continue
            if abstract is not None:
                counts[abstract] = counts.get(abstract, 0) + 1
    abstracts = numbering.abstracts()
    for abstract_id, _ in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
        node = abstracts.get(abstract_id)
        if node is None or node.find(_W + "styleLink") is not None or node.find(_W + "numStyleLink") is not None:
            continue
        levels = node.findall(_W + "lvl")
        if len(levels) < 9 or any(level.find(_W + "pStyle") is not None for level in levels):
            continue
        formats = [(level.find(_W + "numFmt").get(_W + "val") if level.find(_W + "numFmt") is not None else "decimal")
                   for level in levels[:9]]
        if kind == "bullet" and all(f == "bullet" for f in formats):
            return abstract_id
        if kind == "number" and all(f not in ("bullet", "none") for f in formats):
            return abstract_id
    return None


def _new_notes_part(document: "Document", kind: str) -> str:
    """A notes part with its separators, as Word writes it (:func:`docx_agent.edit.notes.new_notes_part`)."""
    from ..edit.notes import new_notes_part

    return new_notes_part(document, kind)


__all__ = ["HTML_MODES", "PLACES", "insert_markdown"]
