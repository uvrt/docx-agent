"""The document read into records and the Markdown model, in a view of its revisions.

The reader walks a story's blocks once and keeps, for every paragraph, a
:class:`ParagraphRecord`: its id, style, what Markdown construct it is, its list place and
number, its inline items in the view, and what the JSON state reports (runs with their
effective formatting, links, bookmarks, notes, comments, revisions, pictures, fields).
:func:`build` then groups records into the Markdown model -- lists, quotes, code blocks,
tables -- and places the ids.  Nothing here writes to the document: the tree is only
read, ids are the ones the document has now (volatile ones included, marked so), and the
effective-formatting model is docx2svg's parse of the serialised bytes.

**Views** (ROADMAP.md, "Text ranges" and "The Markdown layer"):

* ``final`` (also ``current``): insertions and moves' destinations in, deletions and
  moves' sources out, and a paragraph whose mark is deleted (or moved away) joined to the
  next as docx2svg draws it -- the joined paragraph keeps the first one's id and takes the
  last one's style and list; an empty one before a table goes; a deleted table row goes.
* ``original``: the reverse -- deletions in, insertions out, a paragraph whose mark was
  inserted joined to the next, an inserted row gone.  Formatting changes are not undone:
  the original view shows the original *text* with today's formatting.
* ``markup``: both, every revision marked in CriticMarkup (``{++inserted++}``,
  ``{--deleted--}``, a paragraph mark as ``{--¶--}``) followed by its id and author in a
  comment, and every comment as ``{>>author: text<<}`` followed by its id.

**Ids** sit in HTML comments: a block's ``<!-- p:3A1F09C2 -->`` on the line above it; a
list item's or a note's paragraph's at its end; what Markdown cannot say (a content
control, a section break, a page break, an unknown style, a drawing that is not a
picture) as a comment with an id.  A volatile id (no paraId yet, or a repeat) says
``volatile``.
"""

from __future__ import annotations

import re

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from docx2svg.linebreak import ListCounters
from docx2svg.parse.styles import Numbering, parse_numbering, parse_styles
from lxml.etree import tostring as serialize

from ..edit import effective as _effective
from ..edit import ids as _ids
from ..edit import text as _text
from ..edit.annotations import control_kind, drawing_kind
from ..edit.numbering import membership
from ..oxml.xml import Element
from . import model as m
from .stylemap import DEFAULT, StyleMap

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_V = "{urn:schemas-microsoft-com:vml}"

#: The view names ``to_markdown`` takes, and the text views (``edit.text``) they read.
VIEWS = {"final": "current", "current": "current", "original": "original", "markup": "markup"}

_TRANSPARENT = frozenset(_W + n for n in ("smartTag", "customXml", "dir", "bdo"))
_INSERTED = frozenset({_W + "ins", _W + "moveTo"})
_DELETED = frozenset({_W + "del", _W + "moveFrom"})
_CHARACTERS = {_W + "tab": "\t", _W + "ptab": "\t", _W + "noBreakHyphen": _text.NON_BREAKING_HYPHEN,
               _W + "softHyphen": _text.SOFT_HYPHEN}
_SKIPPED_RUN_CONTENT = frozenset(_W + n for n in (
    "rPr", "footnoteRef", "endnoteRef", "annotationRef", "separator", "continuationSeparator",
    "lastRenderedPageBreak", "commentReference", "instrText", "delInstrText"))


def check_view(view: str) -> str:
    if view not in VIEWS:
        raise ValueError(f"no view {view!r}; expected final, original or markup")
    return VIEWS[view]


# -- records ---------------------------------------------------------------------------------


@dataclass
class RunRecord:
    text: str
    style: str | None
    #: Direct properties the run declares (``b``, ``i``, ``color``...).
    direct: list
    effective: object
    revision: str | None = None


@dataclass
class ParagraphRecord:
    id: str
    part: str
    story: str
    elements: list
    volatile: bool
    style_id: str | None
    style_name: str | None
    #: ``paragraph``, ``heading``, ``bullet``, ``ordered``, ``quote``, ``code`` or ``rule``.
    construct: str = "paragraph"
    level: int = 0
    #: Whether no rule of the style map names the style.
    unknown_style: bool = False
    membership: object = None
    label: str | None = None
    number: int | None = None
    inline: list = field(default_factory=list)
    #: The text in the view, as the reader saw it (no comments, no markup delimiters).
    text: str = ""
    joins: list = field(default_factory=list)
    runs: list = field(default_factory=list)
    links: list = field(default_factory=list)
    bookmarks: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    comments: list = field(default_factory=list)
    revisions: list = field(default_factory=list)
    pictures: list = field(default_factory=list)
    drawings: list = field(default_factory=list)
    controls: list = field(default_factory=list)
    fields: list = field(default_factory=list)
    #: Text boxes found in the paragraph: ``(drawing id, [records])``.
    text_boxes: list = field(default_factory=list)
    #: ``(section id, type)`` when the paragraph ends a section.
    section: tuple | None = None
    #: Comment texts hoisted from the start of the line (a bookmark at the paragraph's start).
    hoisted: list = field(default_factory=list)
    #: What the paragraph's own comment adds in the markup view (a style change's revision).
    changes: list = field(default_factory=list)


@dataclass
class CellRecord:
    id: str
    row: int
    column: int
    row_span: int
    column_span: int
    records: list


@dataclass
class TableRecord:
    id: str
    part: str
    story: str
    element: Element
    style_name: str | None
    columns: int
    #: Visible rows: each a list of :class:`CellRecord` (a vertically merged cell appears
    #: once, in its first row).
    rows: list
    header_rows: int = 0
    volatile: bool = False


@dataclass
class Marker:
    """A content control's start or end (``cc``), as a block of its own."""

    id: str
    kind: str
    begin: bool
    tag: str | None = None
    alias: str | None = None


# -- the reader ------------------------------------------------------------------------------


class Reader:
    """One reading of a document in one view; caches what every paragraph asks."""

    def __init__(self, document: "Document", view: str = "final", style_map: StyleMap | None = None) -> None:
        self.document = document
        self.view_name = view if view != "current" else "final"
        self.view = check_view(view)
        self.map = style_map or DEFAULT
        self._styles = {s.id: s for s in document.styles}
        self._default_style = next((s for s in self._styles.values() if s.kind == "paragraph" and s.default), None)
        self._outline: dict[str | None, int | None] = {}
        self._counters: dict = {}
        self._numbering_model = None
        self._comments = {r["w_id"]: r for r in document._comment_records()}
        # Lists held for the reader's life: lxml keeps one proxy per element only while one
        # is referenced, and these maps are keyed by the proxies' id().
        pictures, drawings = document._drawings(), document._all_drawings()
        controls, revisions = document._content_controls(), document._revision_elements()
        self._hold = [pictures, drawings, controls, revisions]
        self._pictures = {id(frame): identifier for _, identifier, frame in pictures}
        self._drawings = {id(frame): identifier for _, identifier, frame in drawings}
        self._controls = {id(node): identifier for _, identifier, node in controls}
        self._revisions = {id(node): identifier for _, identifier, node in revisions}
        self._notes = {(identifier.partition(":")[0], identifier.partition(":")[2]): identifier
                       for _, identifier, _ in document._notes()}
        self._linked_anchors = self._anchors()

    # -- styles --------------------------------------------------------------------------------

    def style_name(self, style_id: str | None) -> str | None:
        style = self._styles.get(style_id) if style_id else None
        style = style or self._default_style
        return style.name if style else None

    def outline_level(self, style_id: str | None) -> int | None:
        """The ``w:outlineLvl`` a paragraph style declares through its ``basedOn`` chain."""
        if style_id in self._outline:
            return self._outline[style_id]
        root = self.document.package.tree(self.document.package.styles_part()) \
            if self.document.package.styles_part() else None
        nodes = {n.get(_W + "styleId"): n for n in root.findall(_W + "style")} if root is not None else {}
        current = style_id or (self._default_style.id if self._default_style else None)
        seen = set()
        level = None
        while current and current not in seen and current in nodes:
            seen.add(current)
            node = nodes[current].find(f"{_W}pPr/{_W}outlineLvl")
            if node is not None:
                try:
                    level = int(node.get(_W + "val"))
                except (TypeError, ValueError):
                    level = None
                break
            based = nodes[current].find(_W + "basedOn")
            current = based.get(_W + "val") if based is not None else None
        self._outline[style_id] = level
        return level

    # -- list numbers --------------------------------------------------------------------------

    def number(self, part: str, membership_, element: Element | None = None) -> tuple[str | None, int | None]:
        """The label and running number of the next item of ``membership_``'s list in
        ``part``, counted as Word counts -- by docx2svg's measured counter
        (``ListCounters.item``, ``parse_numbering``; its ROADMAP, "Numbering and lists --
        measured", N.1-N.4): definitions sharing a ``w:nsid`` are one list; instances over
        one definition count on from each other; a ``w:startOverride`` restarts at the
        instance's first item at that level (at an overriding ``w:lvl``'s ``w:start`` when
        there is one) and is the level's start there; ``w:lvlRestart`` and ``w:isLgl`` are
        honoured; a shallower level with no count is set to its start and counted as used.
        A text box's paragraphs are a story of their own, apart from the part's.  Called
        once per paragraph, in document order."""
        story = (part, "text-box") if element is not None and _in_text_box(element) else part
        counters = self._counters.get(story)
        if counters is None:
            counters = self._counters[story] = ListCounters()
        item = counters.item(self._numbering(), membership_.num_id, membership_.level)
        return item.label, item.number

    def _numbering(self) -> Numbering:
        """docx2svg's reading of the numbering part (``parse_numbering``), which its counter's
        public entry point, ``ListCounters.item``, counts from."""
        if self._numbering_model is None:
            package = self.document.package
            part, styles = package.numbering_part(), package.styles_part()
            self._numbering_model = parse_numbering(serialize(package.tree(part)), parse_styles(
                serialize(package.tree(styles))) if styles else None) if part else Numbering()
        return self._numbering_model

    def _anchors(self) -> set[str]:
        """Bookmark names an internal hyperlink points at: shown even when hidden (``_Toc``)."""
        out = set()
        for part in self.document._parts():
            for node in self.document.package.tree(part).iter(_W + "hyperlink"):
                if node.get(_W + "anchor"):
                    out.add(node.get(_W + "anchor"))
        return out

    # -- blocks --------------------------------------------------------------------------------

    def story(self, part: str) -> list:
        """The records of a story part's blocks (a notes or comments part: every note's)."""
        root = self.document.package.tree(part)
        tag = root.tag
        if tag in (_W + "footnotes", _W + "endnotes", _W + "comments"):
            out = []
            for child in root:
                out += self.container(child, part)
            return out
        container = root.find(_W + "body") if tag == _W + "document" else root
        return self.container(container, part)

    def container(self, container: Element, part: str) -> list:
        """A container's blocks -- paragraphs, tables, content-control markers -- in the
        view, paragraphs joined across removed marks."""
        items = self._items(container, part)
        return self._join(items, part)

    def _items(self, container: Element, part: str) -> list:
        out: list = []
        for child in container:
            tag = child.tag
            if tag == _W + "p":
                out.append(("p", child))
            elif tag == _W + "tbl":
                out.append(("tbl", child))
            elif tag == _W + "sdt":
                identifier = self._controls.get(id(child))
                kind = control_kind(child)
                tag_node = child.find(f"{_W}sdtPr/{_W}tag")
                alias = child.find(f"{_W}sdtPr/{_W}alias")
                marker = (identifier, kind, tag_node.get(_W + "val") if tag_node is not None else None,
                          alias.get(_W + "val") if alias is not None else None)
                content = child.find(_W + "sdtContent")
                out.append(("begin", marker))
                if content is not None:
                    out.extend(self._items(content, part))
                out.append(("end", marker))
            elif tag == _W + "customXml":
                out.extend(self._items(child, part))
        return out

    def _mark_removed(self, paragraph: Element) -> Element | None:
        mark = paragraph.find(f"{_W}pPr/{_W}rPr")
        if mark is None:
            return None
        gone = _DELETED if self.view == "current" else _INSERTED if self.view == "original" else ()
        return next((child for child in mark if child.tag in gone), None)

    def _join(self, items: list, part: str) -> list:
        # Groups of paragraphs that are one in the view (docx2svg's _join, both ways).
        groups: list = []
        pending: list[Element] = []
        for kind, value in items:
            if kind == "p":
                pending.append(value)
                if self._mark_removed(value) is None:
                    groups.append(("p", pending))
                    pending = []
            elif kind == "tbl":
                if pending:
                    # A paragraph whose mark is removed before a table: kept, unless it is
                    # empty and ends no section.
                    if any(_text.paragraph_text(p, self.view) or p.find(f"{_W}pPr/{_W}sectPr") is not None
                           for p in pending):
                        groups.append(("p", pending))
                    pending = []
                groups.append(("tbl", value))
            else:
                groups.append((kind, value))
        if pending:
            # At the container's end: in the original view, an inserted paragraph with
            # nothing of the original in it was not there (nothing follows for it to join).
            if self.view != "original" or any(
                    _text.paragraph_text(p, self.view) or p.find(f"{_W}pPr/{_W}sectPr") is not None for p in pending):
                groups.append(("p", pending))
        out: list = []
        for kind, value in groups:
            if kind == "p":
                out.append(self.paragraph(value, part))
            elif kind == "tbl":
                table = self.table(value, part)
                if table is not None:
                    out.append(table)
            else:
                identifier, control, tag, alias = value
                out.append(Marker(identifier, control, kind == "begin", tag, alias))
        return out

    # -- tables --------------------------------------------------------------------------------

    def table(self, element: Element, part: str) -> TableRecord | None:
        index = self.document._index(part)
        entry = index.entry_for(element)
        rows_all = entry.rows if entry is not None else []
        style = element.find(f"{_W}tblPr/{_W}tblStyle")
        style_id = style.get(_W + "val") if style is not None else None
        style_name = self._styles[style_id].name if style_id in self._styles else style_id
        grid = element.findall(f"{_W}tblGrid/{_W}gridCol")
        origins: dict[int, CellRecord] = {}
        rows: list = []
        header_rows = 0
        columns = len(grid)
        counting_header = True
        for r, (row, _) in enumerate(rows_all):
            properties = row.find(_W + "trPr")
            if properties is not None:
                if self.view == "current" and properties.find(_W + "del") is not None:
                    continue
                if self.view == "original" and properties.find(_W + "ins") is not None:
                    continue
            header = properties is not None and properties.find(_W + "tblHeader") is not None and \
                properties.find(_W + "tblHeader").get(_W + "val") not in ("0", "false", "off")
            if counting_header and header:
                header_rows += 1
            else:
                counting_header = False
            line: list = []
            column = _int(row.find(f"{_W}trPr/{_W}gridBefore"), 0)
            for tc in _cells(row):
                span = max(1, _int(tc.find(f"{_W}tcPr/{_W}gridSpan"), 1))
                merge = tc.find(f"{_W}tcPr/{_W}vMerge")
                continues = merge is not None and merge.get(_W + "val") in (None, "continue")
                if continues and column in origins:
                    origins[column].row_span += 1
                else:
                    cell = CellRecord(f"{entry.id}/c{r},{column}", r, column, 1, span, self.container(tc, part))
                    for c in range(column, column + span):
                        origins[c] = cell
                    line.append(cell)
                column += span
            columns = max(columns, column + _int(row.find(f"{_W}trPr/{_W}gridAfter"), 0))
            rows.append(line)
        if not rows:
            return None
        return TableRecord(entry.id, part, self.document._story_of(part), element, style_name, columns, rows,
                           header_rows, entry.id.startswith("t@"))

    # -- paragraphs ----------------------------------------------------------------------------

    def paragraph(self, elements: list, part: str) -> ParagraphRecord:
        index = self.document._index(part)
        first = index.entry_for(elements[0])
        last = self._as_viewed(elements[-1])
        properties = last.find(_W + "pPr")
        style = properties.find(_W + "pStyle") if properties is not None else None
        style_id = style.get(_W + "val") if style is not None else None
        record = ParagraphRecord(first.id, part, self.document._story_of(part), list(elements), first.volatile,
                                 style_id, self.style_name(style_id))
        record.joins = [index.entry_for(e).id for e in elements[1:]]
        for element in elements:
            # A joined paragraph's text follows the one before directly, as docx2svg draws it.
            _Inline(self, record, element).walk()
            self._mark_revision(record, element)
        record.inline = _tidy(record.inline)
        record.text = _view_text(record.inline)
        self._classify(record, last)
        section = properties.find(_W + "sectPr") if properties is not None else None
        if section is not None and record.story == "body":
            kind = section.find(_W + "type")
            from ..edit.document import _section_id
            record.section = (_section_id(record.id), kind.get(_W + "val") if kind is not None else "nextPage")
        return record

    def _as_viewed(self, paragraph: Element) -> Element:
        """The paragraph whose properties the view shows: in the ``original`` view, one with
        the properties a ``w:pPrChange`` recorded (its style, list and the rest) -- a stand-in
        read for its properties only; otherwise the paragraph itself."""
        change = paragraph.find(f"{_W}pPr/{_W}pPrChange") if self.view == "original" else None
        old = change.find(_W + "pPr") if change is not None else None
        if old is None:
            return paragraph
        import copy

        shadow = paragraph.makeelement(paragraph.tag, dict(paragraph.attrib))
        properties = copy.deepcopy(old)
        current = paragraph.find(_W + "pPr")
        for keep in (_W + "rPr", _W + "sectPr"):
            node = current.find(keep)
            if node is not None:
                properties.append(copy.deepcopy(node))
        shadow.append(properties)
        self._hold.append(shadow)
        return shadow

    def _mark_revision(self, record: ParagraphRecord, element: Element) -> None:
        """A paragraph mark's own insertion or deletion: recorded, and in the markup view
        written after its text; a paragraph-properties change said on its comment."""
        change = element.find(f"{_W}pPr/{_W}pPrChange")
        if change is not None:
            identifier = self._revisions.get(id(change))
            record.revisions.append(identifier)
            if self.view == "markup":
                record.changes.append(f"{identifier} paragraph-properties {change.get(_W + 'author') or ''}".strip())
        mark = element.find(f"{_W}pPr/{_W}rPr")
        if mark is None:
            return
        for child in mark:
            if child.tag in _INSERTED or child.tag in _DELETED:
                identifier = self._revisions.get(id(child))
                record.revisions.append(identifier)
                if self.view != "markup":
                    continue
                sign = "++" if child.tag in _INSERTED else "--"
                record.inline.append(m.Critic("{" + sign + "¶" + sign + "}"))
                record.inline.append(m.Html(_comment(f"{identifier} {child.get(_W + 'author') or ''}".strip())))

    def _classify(self, record: ParagraphRecord, element: Element) -> None:
        name = record.style_name
        heading = self.map.heading_level(name, self.outline_level(record.style_id))
        listed = membership(self.document, element)
        rule = self.map.paragraph_rule(name)
        if heading is not None:
            record.construct, record.level = "heading", heading
            record.unknown_style = not (rule is not None and rule.construct == "heading" and heading <= 6)
            if listed is not None and listed.format not in (None, "none"):
                record.membership = listed
                record.label, record.number = self.number(record.part, listed, element)
            return
        if listed is not None and listed.format not in (None, "none"):
            record.membership = listed
            record.construct = "bullet" if listed.format == "bullet" else "ordered"
            record.level = listed.level
            record.label, record.number = self.number(record.part, listed, element)
            record.unknown_style = rule is None
            return
        if rule is not None and rule.construct in ("bullet_list", "ordered_list"):
            record.construct = "bullet" if rule.construct == "bullet_list" else "ordered"
            record.level = rule.level - 1
            return
        if rule is not None and rule.construct == "blockquote":
            record.construct = "quote"
            return
        if rule is not None and rule.construct == "code_block":
            record.construct = "code"
            return
        record.unknown_style = rule is None
        border = element.find(f"{_W}pPr/{_W}pBdr/{_W}bottom")
        if not record.text.strip() and border is not None and border.get(_W + "val") not in (None, "nil", "none") \
                and not record.inline:
            record.construct = "rule"


def _int(node: Element | None, default: int) -> int:
    try:
        return int(node.get(_W + "val")) if node is not None else default
    except (TypeError, ValueError):
        return default


def _cells(row: Element) -> list[Element]:
    out: list[Element] = []
    for child in row:
        if child.tag == _W + "tc":
            out.append(child)
        elif child.tag in (_W + "sdt", _W + "customXml"):
            content = child.find(_W + "sdtContent") if child.tag == _W + "sdt" else child
            if content is not None:
                out.extend(_cells(content))
    return out


def _comment(text: str) -> str:
    """An HTML comment holding ``text``, which cannot end it early."""
    return "<!-- " + text.replace("--", "- -").replace(">", "&gt;") + " -->"


def comment_text(text: str) -> str:
    return text.replace("--", "- -").replace(">", "&gt;")


def _view_text(inline: list) -> str:
    """The text of the view: no comments, no CriticMarkup delimiters, no comment's words."""
    out = []
    commenting = False
    for item in inline:
        if isinstance(item, m.Critic):
            commenting = item.text == "{>>" or (commenting and item.text != "<<}")
        elif isinstance(item, m.Text) and not commenting:
            out.append(item.text)
        elif isinstance(item, m.Break):
            out.append("\v")
    return "".join(out)


def _tidy(inline: list) -> list:
    """Merge adjacent alike texts; move a comment that would start a line (after a hard
    break) before the break, where Markdown cannot mistake it for an HTML block."""
    out: list = []
    for item in inline:
        if isinstance(item, m.Text) and not item.text:
            continue
        if out and isinstance(item, m.Text) and isinstance(out[-1], m.Text) \
                and out[-1].marks == item.marks and out[-1].href == item.href:
            out[-1] = m.Text(out[-1].text + item.text, item.marks, item.href)
            continue
        out.append(item)
    for k in range(1, len(out)):
        if isinstance(out[k], m.Html) and isinstance(out[k - 1], m.Break):
            j = k
            while j > 0 and isinstance(out[j - 1], m.Break):
                out[j - 1], out[j] = out[j], out[j - 1]
                j -= 1
    return out


# -- inline ----------------------------------------------------------------------------------


class _Inline:
    """One paragraph's content in the view, as model inline items and record details."""

    def __init__(self, reader: Reader, record: ParagraphRecord, paragraph: Element) -> None:
        self.reader = reader
        self.document = reader.document
        self.record = record
        self.paragraph = paragraph
        self.view = reader.view
        self.out = record.inline
        self.href: list[str | None] = []
        self.fields: list[list] = []  # open complex fields: [mode, instruction, result, owner]
        self.simple = 0
        self._baseline = None
        self._links = list(paragraph.iter(_W + "hyperlink"))
        self._link_index = {id(node): k for k, node in enumerate(self._links)}
        self._revision: list[str | None] = []

    def walk(self) -> None:
        self._container(self.paragraph)

    # -- formatting --------------------------------------------------------------------------

    def baseline(self):
        if self._baseline is None:
            self._baseline = _effective.run(self.document, self.paragraph, None, "a")
        return self._baseline

    def _marks(self, run: Element, first: str):
        effective = _effective.run(self.document, self.paragraph, run, first or "a")
        base = self.baseline()
        marks = set()
        if effective.bold and not base.bold:
            marks.add("strong")
        if effective.italic and not base.italic:
            marks.add("em")
        if (effective.strike or effective.double_strike) and not (base.strike or base.double_strike):
            marks.add("strike")
        style_name = None
        if effective.style:
            style = self.reader._styles.get(effective.style)
            style_name = style.name if style else effective.style
            rule = self.reader.map.character_rule(style_name)
            if rule is not None and rule.construct == "code":
                marks.add("code")
        return frozenset(marks), effective, style_name

    # -- walking -------------------------------------------------------------------------------

    def _visible(self) -> bool:
        return all(f[0] == "result" for f in self.fields)

    def _container(self, node: Element) -> None:
        for child in node:
            tag = child.tag
            if not isinstance(tag, str):
                continue
            if tag == _W + "r":
                self._run(child)
            elif tag == _W + "hyperlink":
                self._hyperlink(child)
            elif tag in _TRANSPARENT:
                self._container(child)
            elif tag == _W + "fldSimple":
                instruction = (child.get(_W + "instr") or "").strip()
                self.record.fields.append({"instruction": instruction, "result": ""})
                marked = self.view == "markup" and self._visible()
                if marked:
                    self.out.append(m.Html(_comment(f"field: {instruction}")))
                self.simple += 1
                start = len(self.out)
                self._container(child)
                self.simple -= 1
                self.record.fields[-1]["result"] = m.plain_text(
                    [i for i in self.out[start:] if not isinstance(i, m.Html)])
                if marked:
                    self.out.append(m.Html(_comment("/field")))
            elif tag == _W + "sdt":
                self._inline_control(child)
            elif tag in _INSERTED or tag in _DELETED:
                self._revision_container(child)
            elif tag == _W + "bookmarkStart":
                self._bookmark(child)
            elif tag in (_M + "oMath", _M + "oMathPara"):
                if self._visible():
                    math = "".join(t.text or "" for t in child.iter(_M + "t"))
                    self.out.append(m.Html(_comment(f"equation: {math}")))

    def _hyperlink(self, node: Element) -> None:
        target = None
        rid = node.get(_R + "id")
        if rid:
            relationship = self.document.package.relationships(self.record.part).get(rid)
            target = relationship.target if relationship is not None else None
        anchor = node.get(_W + "anchor")
        if anchor:
            target = (target or "") + "#" + anchor
        index = self._link_index.get(id(node))
        start = len(self.out)
        self.href.append(target)
        self._container(node)
        self.href.pop()
        if index is not None:
            self.record.links.append({"id": f"hl:{self.record.id}/{index}", "href": target,
                                      "text": m.plain_text(self.out[start:])})

    def _inline_control(self, node: Element) -> None:
        identifier = self.reader._controls.get(id(node))
        kind = control_kind(node)
        content = node.find(_W + "sdtContent")
        self.record.controls.append({"id": identifier, "kind": kind})
        self.out.append(m.Html(_comment(f"{identifier} {kind}")))
        if content is not None:
            self._container(content)
        self.out.append(m.Html(_comment(f"/{identifier}")))

    def _revision_container(self, node: Element) -> None:
        inserted = node.tag in _INSERTED
        shown = ("current", "markup") if inserted else ("original", "markup")
        identifier = self.reader._revisions.get(id(node))
        if identifier is not None and identifier not in self.record.revisions:
            self.record.revisions.append(identifier)
        if self.view not in shown:
            return
        self._revision.append(identifier)
        if self.view != "markup":
            self._container(node)
            self._revision.pop()
            return
        sign = "++" if inserted else "--"
        self.out.append(m.Critic("{" + sign))
        self._container(node)
        self._revision.pop()
        self.out.append(m.Critic(sign + "}"))
        author = node.get(_W + "author") or ""
        moved = " moved" if node.tag in (_W + "moveTo", _W + "moveFrom") else ""
        self.out.append(m.Html(_comment(f"{identifier}{moved} {author}".strip())))

    def _bookmark(self, node: Element) -> None:
        name = node.get(_W + "name") or ""
        if not name or (name.startswith("_") and name not in self.reader._linked_anchors):
            return
        self.record.bookmarks.append(f"bm:{name}")
        self.out.append(m.Html(_comment(f"bm:{name}")))

    def _run(self, run: Element) -> None:
        items = self._run_items(run)
        text = "".join(c for kind, c, _ in items if kind == "char")
        if not items:
            return
        marks, effective, style_name = self._marks(run, text[:1])
        if effective.hidden:
            return
        properties = run.find(_W + "rPr")
        direct = [child.tag.rpartition("}")[2] for child in properties
                  if isinstance(child.tag, str) and child.tag != _W + "rStyle"] if properties is not None else []
        if text:
            self.record.runs.append(RunRecord(text, style_name, direct, effective,
                                              self._revision[-1] if self._revision else None))
        change = properties.find(_W + "rPrChange") if properties is not None else None
        if change is not None:
            identifier = self.reader._revisions.get(id(change))
            if identifier not in self.record.revisions:
                self.record.revisions.append(identifier)
        href = self.href[-1] if self.href else None
        buffer: list[str] = []

        def flush() -> None:
            if buffer:
                self.out.append(m.Text("".join(buffer), marks, href))
                buffer.clear()

        for kind, value, node in items:
            if kind == "char":
                buffer.append(value)
                continue
            flush()
            if kind == "break":
                self.out.append(m.Break())
            elif kind == "html":
                self.out.append(m.Html(value))
            elif kind == "note":
                self.out.append(m.NoteRef(value))
                self.record.notes.append(value)
            elif kind == "drawing":
                self._drawing(node, href)
            elif kind == "comment":
                self._comment_reference(node)
        flush()
        if change is not None and self.view == "markup":
            author = change.get(_W + "author") or ""
            self.out.append(m.Html(_comment(f"{self.reader._revisions.get(id(change))} formatting {author}".strip())))

    def _run_items(self, run: Element) -> list:
        """(kind, value, node) for the run's content in the view: ``char``, ``break``,
        ``html``, ``note``, ``drawing``, ``comment``."""
        out: list = []
        for child in run:
            tag = child.tag
            if not isinstance(tag, str):
                continue
            if tag == _W + "fldChar":
                kind = child.get(_W + "fldCharType")
                if kind == "begin":
                    self.fields.append(["instruction", "", child, len(self.record.fields)])
                    self.record.fields.append({"instruction": "", "result": ""})
                elif kind == "separate" and self.fields:
                    self.fields[-1][0] = "result"
                    if self.view == "markup" and self._visible():
                        instruction = self.record.fields[self.fields[-1][3]]["instruction"]
                        out.append(("html", _comment(f"field: {instruction}"), child))
                        self.fields[-1][1] = "marked"
                elif kind == "end" and self.fields:
                    marked = self.fields[-1][1] == "marked"
                    self.fields.pop()
                    if marked and self._visible():
                        out.append(("html", _comment("/field"), child))
                continue
            if tag in (_W + "instrText", _W + "delInstrText"):
                if self.fields and self.fields[-1][0] == "instruction" and \
                        (tag == _W + "instrText" or self.view != "current"):
                    self.record.fields[-1]["instruction"] = (self.record.fields[-1]["instruction"]
                                                             + (child.text or "")).strip()
                continue
            if tag == _W + "commentReference":
                if self._visible():
                    out.append(("comment", None, child))
                continue
            if tag in _SKIPPED_RUN_CONTENT or not self._visible():
                continue
            if tag == _W + "t":
                out.extend(("char", c, child) for c in child.text or "")
                self._field_result(child.text or "")
            elif tag == _W + "delText":
                if self.view != "current":
                    out.extend(("char", c, child) for c in child.text or "")
            elif tag in (_W + "br", _W + "cr"):
                kind = child.get(_W + "type")
                if kind == "page":
                    out.append(("html", _comment("page break"), child))
                elif kind == "column":
                    out.append(("html", _comment("column break"), child))
                else:
                    out.append(("break", None, child))
            elif tag in _CHARACTERS:
                out.append(("char", _CHARACTERS[tag], child))
            elif tag == _W + "sym":
                out.append(("char", _text._symbol(child), child))
            elif tag in (_W + "footnoteReference", _W + "endnoteReference"):
                prefix = "fn" if tag == _W + "footnoteReference" else "en"
                identifier = self.reader._notes.get((prefix, child.get(_W + "id")))
                if identifier is not None:
                    out.append(("note", identifier, child))
            elif tag in (_W + "drawing", _W + "pict", _MC + "AlternateContent", _W + "object"):
                out.append(("drawing", None, child))
        return out

    def _field_result(self, text: str) -> None:
        if self.fields and self.record.fields:
            self.record.fields[-1]["result"] += text

    def _comment_reference(self, node: Element) -> None:
        record = self.reader._comments.get(node.get(_W + "id"))
        if record is None:
            return
        self.record.comments.append(record["id"])
        if self.view != "markup":
            return
        text = _blocks_plain(self.document, record["part"], record["element"])
        who = ", ".join(w for w in (record["author"], "in reply" if record["parent"] else None,
                                    "resolved" if record["done"] else None) if w)
        self.out += [m.Critic("{>>"), m.Text(f"{who}: {text}" if who else text), m.Critic("<<}"),
                     m.Html(_comment(record["id"]))]

    def _drawing(self, node: Element, href: str | None) -> None:
        frames = [f for f in _ids.live_elements(node, frozenset({_ids.WP_INLINE, _ids.WP_ANCHOR}))]
        if not frames:
            # VML (w:pict, w:object): a text box's content is still text; the rest is a mark.
            for box in node.iter(_W + "txbxContent"):
                self.record.text_boxes.append((None, box))
            self.out.append(m.Html(_comment("drawing")))
            return
        for frame in frames:
            identifier = self.reader._pictures.get(id(frame)) or self.reader._drawings.get(id(frame))
            doc_pr = frame.find(_WP + "docPr")
            name = doc_pr.get("name") if doc_pr is not None else None
            alt = (doc_pr.get("descr") if doc_pr is not None else None) or ""
            kind = drawing_kind(frame)
            floating = frame.tag == _ids.WP_ANCHOR
            if id(frame) in self.reader._pictures:
                extent = frame.find(_WP + "extent")
                size = (int(extent.get("cx")) / 12700, int(extent.get("cy")) / 12700) if extent is not None else None
                self.record.pictures.append({"id": identifier, "name": name, "alt": alt, "floating": floating,
                                             "size": size})
                self.out.append(m.Image(alt, identifier or "", name, href))
            else:
                entry = {"id": identifier, "kind": kind, "name": name, "alt": alt, "floating": floating}
                label = f'{identifier} {kind}' + (f' "{name}"' if name else "")
                label += _graphic(self.document, self.record.part, frame, identifier, kind, entry)
                self.record.drawings.append(entry)
                self.out.append(m.Html(_comment(label)))
                for box in frame.iter(_W + "txbxContent"):
                    if not _in_fallback(box):
                        self.record.text_boxes.append((identifier, box))
                        break


def _graphic(document: "Document", part: str, frame: Element, identifier: str, kind: str, entry: dict) -> str:
    """A chart's or diagram's JSON into ``entry`` (``chart``, ``diagram``; a group's
    ``members``), and its one-line summary for the drawing's comment -- pptx-agent's
    outline line, ``[chart: column; title: ...; series: ...; categories: ...]``, and a
    diagram's node text, ``[smartart: Goals [Faster edits], Risks]``; in the comment, since
    the drawing stands in a paragraph's line and is not a block of its own."""
    from ..edit import charts as _charts

    summary = ""
    if kind == "chart":
        model = _charts.chart_json(document, part, frame)
        if model is not None:
            entry["chart"] = model
            summary = f" [chart: {_charts.chart_summary(model)}]"
    elif kind == "diagram":
        model = _charts.diagram_json(document, part, frame)
        if model is not None:
            entry["diagram"] = model
            summary = f" [smartart: {_charts.diagram_summary(model)}]"
    elif kind == "group":
        members = []
        for member, element in _charts.group_members(frame, identifier):
            member_kind = _charts.member_kind(element)
            properties = _charts._member_properties(element)
            item = {"id": member, "kind": member_kind,
                    "name": properties.get("name") if properties is not None else None}
            model = (_charts.chart_json(document, part, element) if member_kind == "chart" else
                     _charts.diagram_json(document, part, element) if member_kind == "diagram" else None)
            if model is not None:
                item[member_kind] = model
                text = (_charts.chart_summary if member_kind == "chart" else _charts.diagram_summary)(model)
                summary += f" [{member} {'chart' if member_kind == 'chart' else 'smartart'}: {text}]"
            members.append(item)
        entry["members"] = members
    return summary


def _in_text_box(node: Element) -> bool:
    parent = node.getparent()
    while parent is not None:
        if parent.tag == _W + "txbxContent":
            return True
        parent = parent.getparent()
    return False


def _in_fallback(node: Element) -> bool:
    parent = node.getparent()
    while parent is not None:
        if parent.tag == _MC + "Fallback":
            return True
        parent = parent.getparent()
    return False


def _blocks_plain(document: "Document", part: str, container: Element) -> str:
    index = document._index(part)
    inside = {id(e) for e in container.iter(_W + "p")}
    texts = [_text.paragraph_text(entry.element) for entry in index.paragraphs if id(entry.element) in inside]
    return " ".join(t.replace(_text.OBJECT, "").strip() for t in texts if t.replace(_text.OBJECT, "").strip())


# -- building the Markdown model -------------------------------------------------------------


@dataclass
class Options:
    ids: bool = True


def build(reader: Reader, records: list, options: Options, *, trailing: bool = False) -> list:
    """The model blocks for ``records`` (one container's, in order)."""
    out: list = []
    k = 0
    while k < len(records):
        record = records[k]
        if isinstance(record, ParagraphRecord) and record.construct in ("bullet", "ordered"):
            j = k
            while j < len(records) and isinstance(records[j], ParagraphRecord) and \
                    records[j].construct in ("bullet", "ordered"):
                j += 1
            out += _lists(reader, records[k:j], options, out)
            k = j
            continue
        if isinstance(record, ParagraphRecord) and record.construct == "quote":
            j = k
            while j < len(records) and isinstance(records[j], ParagraphRecord) and records[j].construct == "quote":
                j += 1
            inner = []
            for item in records[k:j]:
                inner += _paragraph_blocks(reader, item, options, trailing=False)
            out.append(m.Quote(inner))
            out += _after(reader, records[k:j], options)
            k = j
            continue
        if isinstance(record, ParagraphRecord) and record.construct == "code":
            j = k
            while j < len(records) and isinstance(records[j], ParagraphRecord) and records[j].construct == "code":
                j += 1
            group = records[k:j]
            if options.ids:
                out.append(m.Comment(" ".join(_label(r) for r in group)))
            out.append(m.CodeBlock("".join(_code_text(r) + "\n" for r in group)))
            out += _after(reader, group, options)
            k = j
            continue
        if isinstance(record, ParagraphRecord):
            out += _paragraph_blocks(reader, record, options, trailing=trailing)
            out += _after(reader, [record], options)
        elif isinstance(record, TableRecord):
            out += _table(reader, record, options)
        elif isinstance(record, Marker):
            if options.ids:
                if record.begin:
                    words = [record.id, record.kind] + ([f"tag: {record.tag}"] if record.tag else []) + \
                        ([f"title: {record.alias}"] if record.alias else [])
                    out.append(m.Comment(comment_text(" ".join(w for w in words if w))))
                else:
                    out.append(m.Comment(f"/{record.id}"))
        k += 1
    return out


#: A heading whose text starts with a number typed into it: "4", "4.1", "4.1.", "IV.", "A.".
_TYPED_NUMBER = re.compile(r"\s*(\d+(\.\d+)*\.?|[IVXLC]+\.|[A-Z]\.)\s")


def _label(record: ParagraphRecord) -> str:
    """What a paragraph's id comment says: the id, what joined it, volatility, an unknown
    style, comments hoisted from the start of its line."""
    words = [record.id]
    if record.joins:
        words.append("joins " + " ".join(record.joins))
    if record.volatile:
        words.append("volatile")
    if record.unknown_style and record.style_name:
        words.append(f"style: {comment_text(record.style_name)}")
    if record.construct == "heading" and record.level > 6:
        words.append(f"level {record.level}")
    if record.construct == "heading":
        if record.label:
            words.append("numbered: list")
        elif _TYPED_NUMBER.match(record.text):
            words.append("numbered: text")
    words += record.changes
    words += record.hoisted
    return " ".join(words)


def _hoist(record: ParagraphRecord, inline: list) -> list:
    """Comments at the start of the line move into the paragraph's own comment: a line that
    starts with ``<!--`` is an HTML block, not a paragraph."""
    k = 0
    while k < len(inline) and (isinstance(inline[k], m.Html) or
                               (isinstance(inline[k], m.Text) and not inline[k].text.strip())):
        k += 1
    hoisted = [i for i in inline[:k] if isinstance(i, m.Html)]
    if not hoisted:
        return inline
    record.hoisted = [h.text[5:-4].strip() for h in hoisted]
    return inline[k:]


def _strip_comments(inline: list) -> list:
    return [i for i in inline if not isinstance(i, m.Html)]


def _paragraph_blocks(reader: Reader, record: ParagraphRecord, options: Options, *, trailing: bool) -> list:
    inline = list(record.inline) if options.ids else _strip_comments(record.inline)
    if options.ids:
        inline = _hoist(record, inline)
    if record.construct == "heading":
        inline = [m.Text(" ", frozenset(), None) if isinstance(i, m.Break) else i for i in inline]
        if record.label:
            inline = [m.Text(record.label + " ")] + inline
    out: list = []
    if record.construct == "rule":
        if options.ids:
            out.append(m.Comment(_label(record)))
        out.append(m.Rule())
        return out
    empty = not m.has_content(inline)
    if options.ids and trailing and not (empty and not inline and record.construct != "heading"):
        inline = inline + ([m.Text(" ")] if inline else []) + [m.Html(_comment(_label(record)))]
    elif options.ids and empty and record.construct != "heading":
        # Nothing to write but the id: a comment on its own, not attached to what follows --
        # in a note too, where a line that is only a comment would be read as an HTML block.
        out.append(m.Comment(_label(record) + " empty"))
    elif options.ids:
        out.append(m.Comment(_label(record)))
    if record.construct == "heading":
        out.append(m.Heading(min(record.level, 6), inline))
    elif not (options.ids and empty and (not trailing or not inline)):
        out.append(m.Paragraph(inline))
    for identifier, box in record.text_boxes:
        boxed = reader.container(box, record.part)
        if options.ids:
            out.append(m.Comment(f"{identifier or 'drawing'} text-box"))
        out += build(reader, boxed, options)
        if options.ids:
            out.append(m.Comment(f"/{identifier or 'drawing'}"))
    return out


def _after(reader: Reader, records: list, options: Options) -> list:
    """What follows a group of paragraphs: a section break it ends."""
    out = []
    for record in records:
        if record.section is not None and options.ids:
            out.append(m.Comment(f"{record.section[0]} section break: {record.section[1]}"))
    return out


def _code_text(record: ParagraphRecord) -> str:
    out = []
    for item in record.inline:
        if isinstance(item, m.Text):
            out.append(item.text)
        elif isinstance(item, m.Break):
            out.append("\n")
        elif isinstance(item, m.Critic):
            out.append(item.text)
    return "".join(out)


def _lists(reader: Reader, records: list, options: Options, before: list) -> list:
    """Consecutive list paragraphs as nested Markdown lists."""
    out: list = []
    # Stack of (level, ListBlock, num_id).
    stack: list = []
    previous = next((b for b in reversed(before) if not isinstance(b, m.Comment)), None)
    for record in records:
        ordered = record.construct == "ordered"
        num_id = record.membership.num_id if record.membership is not None else None
        level = record.level
        while stack and stack[-1][0] > level:
            stack.pop()
        if stack and stack[-1][0] == level and (stack[-1][1].ordered != ordered or stack[-1][2] != num_id):
            # Another list at the same level: a sibling of the one ending.
            ended = stack.pop()[1]
            new = m.ListBlock(ordered, _start(record), alternate=not ended.alternate
                              if ended.ordered == ordered else False)
            _place(out, stack, new)
            stack.append((level, new, num_id))
        elif not stack or stack[-1][0] < level:
            alternate = False
            if not stack:
                last = out[-1] if out else previous
                alternate = isinstance(last, m.ListBlock) and last.ordered == ordered and not last.alternate
            new = m.ListBlock(ordered, _start(record), alternate=alternate)
            _place(out, stack, new)
            stack.append((level, new, num_id))
        item = _paragraph_blocks(reader, record, options, trailing=True)
        stack[-1][1].items.append(item)
        item += _after(reader, [record], options)
    return out


def _start(record: ParagraphRecord) -> int:
    return record.number if record.construct == "ordered" and record.number is not None else 1


def _place(out: list, stack: list, new: m.ListBlock) -> None:
    if stack:
        parent = stack[-1][1]
        if not parent.items:
            parent.items.append([])
        parent.items[-1].append(new)
    else:
        out.append(new)


# -- tables ----------------------------------------------------------------------------------


def _table(reader: Reader, table: TableRecord, options: Options) -> list:
    out: list = []
    if options.ids:
        words = [table.id] + (["volatile"] if table.volatile else [])
        if table.style_name:
            words.append(f"table style: {comment_text(table.style_name)}")
        out.append(m.Comment(" ".join(words)))
    gfm = _gfm(reader, table, options)
    if gfm is not None:
        out.append(gfm)
    else:
        from .html import table_html

        out.append(m.HtmlBlock(table_html(reader, table, options)))
    return out


def _gfm(reader: Reader, table: TableRecord, options: Options) -> m.Table | None:
    """The table as GFM, or ``None`` when GFM cannot say it: merged cells, a row with
    fewer cells than the grid, a cell holding anything but one plain paragraph (or none),
    a line break in a cell."""
    width = table.columns
    rows: list = []
    for line in table.rows:
        if len(line) != width or any(c.row_span != 1 or c.column_span != 1 for c in line):
            return None
        cells = []
        for cell in line:
            records = cell.records
            if len(records) > 1 or (records and not isinstance(records[0], ParagraphRecord)):
                return None
            if not records:
                cells.append([])
                continue
            record = records[0]
            if record.construct not in ("paragraph", "quote", "code") or record.text_boxes or record.section:
                return None
            inline = record.inline if options.ids else _strip_comments(record.inline)
            if any(isinstance(i, m.Break) for i in inline):
                return None
            cells.append(list(inline))
        rows.append(cells)
    aligns = []
    for column in range(width):
        found = set()
        for line in table.rows[1:] or table.rows:
            for record in line[column].records:
                if isinstance(record, ParagraphRecord):
                    found.add(_effective.paragraph(reader.document, record.elements[-1]).alignment)
        aligns.append(found.pop() if len(found) == 1 and found <= {"center", "right"} else None)
    return m.Table(rows[0], rows[1:], aligns)


# -- notes -----------------------------------------------------------------------------------


def note_definitions(reader: Reader, labels: list[str], options: Options) -> list:
    """Footnote definitions for the notes ``labels`` names, in that order, and for the notes
    they reference in turn."""
    out = []
    seen: set[str] = set()
    queue = list(labels)
    notes = {identifier: (part, element) for part, identifier, element in reader.document._notes()}
    while queue:
        label = queue.pop(0)
        if label in seen or label not in notes:
            continue
        seen.add(label)
        part, element = notes[label]
        records = reader.container(element, part)
        for record in records:
            if isinstance(record, ParagraphRecord):
                _drop_note_mark(record)
        blocks = build(reader, records, options, trailing=True)
        out.append(m.FootnoteDef(label, blocks))
        queue += [n for r in records for n in referenced_notes(r)]
    return out


def _drop_note_mark(record: ParagraphRecord) -> None:
    """A note's own reference mark (``w:footnoteRef``) is not text; nor is the space after it."""
    if record.inline and isinstance(record.inline[0], m.Text):
        record.inline[0] = m.Text(record.inline[0].text.lstrip(" "), record.inline[0].marks, record.inline[0].href)
        if not record.inline[0].text:
            record.inline.pop(0)


def referenced_notes(record) -> list[str]:
    """Notes a record references, in order (in tables and text boxes too)."""
    if isinstance(record, ParagraphRecord):
        return list(record.notes)
    if isinstance(record, TableRecord):
        return [n for line in record.rows for cell in line for r in cell.records for n in referenced_notes(r)]
    return []


def walk_records(records: list):
    """Every record, depth first: tables' cells included."""
    for record in records:
        yield record
        if isinstance(record, TableRecord):
            for line in record.rows:
                for cell in line:
                    yield from walk_records(cell.records)


__all__ = ["Reader", "ParagraphRecord", "TableRecord", "CellRecord", "Marker", "build", "note_definitions",
           "VIEWS", "check_view", "walk_records"]
