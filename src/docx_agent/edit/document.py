"""The editing API: ``Document`` -> stories -> paragraphs, tables, runs.

Every view is a *live* view over the lxml tree, never a copy: a property reads the XML on
demand, and anything the views do not model -- ``mc:AlternateContent``, ``w14``/``w15``
extensions, custom XML, permission ranges, proofing marks, math -- survives because nothing
removed it.  Views hold an *id*, not an element: they resolve it on every access (as
pptx-agent's runs do), so a view taken before an edit, an undo or a redo still answers.

Every edit is one undo step (ooxml-edit's history), checks what it can before it changes
anything, and returns an :class:`EditResult` naming what it touched: the ids it created,
renamed (a paragraph stamped or re-issued, ROADMAP.md "Addressing") and removed.
"""

from __future__ import annotations

import copy
import os
import re
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, BinaryIO, Iterator

from lxml import etree
from ooxml_edit.history import History

from ..oxml.package import WordPackage, story_name
from ..oxml.xml import REVISION_PROPERTY_TAGS, Element, qn, remove
from . import ids as _ids
from .errors import EditError
from . import text as _text
from .ids import PartIndex, ParagraphEntry, TableEntry
from . import effective as _effective
from . import inline as _inline
from . import formatting as _formatting
from .annotations import AnnotationOps
from .formatops import FormatOps
from .links import Bookmark, Hyperlink, LinkOps
from .numbering import ListMembership, ListOps
from .pictures import Picture, PictureOps
from .ranges import TextOps, TextRange
from .comments import CommentOps
from .tables import TableOps
from .table_format import UNSET, TableFormatOps
from .drawings import DrawingOps
from .charts import ChartOps
from .controls import ControlOps
from .sections import Section, SectionOps, section_id as _section_id
from .notes import NoteOps
from .fields import Field, FieldOps  # noqa: F401  (re-exported)
from .authoring import AuthoringOps
from .importing import ImportOps
from .properties import PropertyOps
from ..revisions import track as _track
from ..revisions.mode import TrackingOps, trackable
from ..revisions.review import RevisionOps
from ..revisions.stamp import Stamp

if TYPE_CHECKING:  # pragma: no cover
    from ..layout import DocumentLayout, Reflow, _Conversion

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_P = _W + "p"
W_TBL = _W + "tbl"
W_PPR = _W + "pPr"
W_SECTPR = _W + "sectPr"
W_TC = _W + "tc"
W_BODY = _W + "body"

#: Range markers: a start and an end sharing ``w:id``.
_RANGE_PAIRS = {
    _W + "bookmarkStart": _W + "bookmarkEnd",
    _W + "commentRangeStart": _W + "commentRangeEnd",
    _W + "moveFromRangeStart": _W + "moveFromRangeEnd",
    _W + "moveToRangeStart": _W + "moveToRangeEnd",
    _W + "permStart": _W + "permEnd",
}
_RANGE_ENDS = {end: start for start, end in _RANGE_PAIRS.items()}

#: What a paragraph cannot be deleted with in E0: deleting it would orphan a note, a
#: comment or half a field.  E3 and E4 handle these.
_REFERENCES = (_W + "footnoteReference", _W + "endnoteReference", _W + "commentReference")




@dataclass
class EditResult:
    """What an edit touched.  ``id`` is the edited (or created) block's id after the edit."""

    id: str | None
    created: list[str] = field(default_factory=list)
    #: Old id -> new id, for every paragraph stamped or re-issued by the edit.
    renamed: dict[str, str] = field(default_factory=dict)
    removed: list[str] = field(default_factory=list)
    #: Whether the document changed at all.
    changed: bool = True
    #: How many places a find-and-replace (or another many-place edit) changed.
    count: int | None = None
    #: What the edit did that the caller may want to know about (a reference left
    #: pointing at something removed).
    warnings: list[str] = field(default_factory=list)
    #: The top-level blocks an insertion made, in order (``insert_markdown``): ``created``
    #: lists these and what is inside them (cells' paragraphs, notes).
    blocks: list[str] = field(default_factory=list)
    #: What the edit could not compute (fields whose target is past where docx2svg's
    #: layout stops: ``update_fields``, ``insert_toc``).
    unknown: list[str] = field(default_factory=list)
    #: The pages the edit reflowed, by docx2svg's layout before and after, for an edit
    #: that reports it (``upgrade_to_modern``).
    reflow: "Reflow | None" = None
    #: A copy's id for each source id it copies (``copy_blocks``): blocks, paragraphs in
    #: them, notes and comments.
    copied: dict[str, str] = field(default_factory=dict)
    #: The document edited, set for an edit made through the API: what :attr:`object`
    #: resolves ``id`` in.
    document: "Document | None" = field(default=None, repr=False, compare=False)

    @property
    def object(self):
        """The edited or created object itself -- ``doc.get(result.id)``: a
        :class:`Paragraph`, :class:`Table`, picture, note, comment...::

            paragraph = doc.insert_paragraph("New text.", after="p:3B212964").object
            paragraph.format(bold=True)"""
        if self.id is None:
            raise AttributeError("the edit names no object (id is None)")
        if self.document is None:
            raise AttributeError(f"this result does not know its document: use doc.get({self.id!r})")
        return self.document.get(self.id)


# -- views -----------------------------------------------------------------------------------


class Paragraph:
    """A ``w:p`` anywhere in a story: the body, a table cell, a content control, a text box,
    a header, a note or a comment."""

    def __init__(self, document: "Document", identifier: str) -> None:
        self._document = document
        self._id = identifier

    def _entry(self) -> tuple[str, ParagraphEntry]:
        part, entry = self._document._resolve(self._id)
        if not isinstance(entry, ParagraphEntry):
            raise KeyError(f"{self._id} is not a paragraph")
        return part, entry

    @property
    def _element(self) -> Element:
        return self._entry()[1].element

    @property
    def id(self) -> str:
        """The paragraph's id now: after a stamp or a re-issue, the new one."""
        return self._entry()[1].id

    @property
    def story(self) -> str:
        """The story the paragraph is in: ``"body"``, ``"header1"``, ``"footnotes"``..."""
        return self._document._story_of(self._entry()[0])

    @property
    def volatile(self) -> bool:
        """Whether the id is positional (``p@...``) or a repeat (``#k``) until edited."""
        return self._entry()[1].volatile

    @property
    def para_id(self) -> str | None:
        """``w14:paraId`` as written, or ``None``."""
        return self._element.get(_ids.PARA_ID)

    @property
    def text_id(self) -> str | None:
        """Word's ``w14:textId`` on the paragraph, or ``None``."""
        return self._element.get(_ids.TEXT_ID)

    @property
    def text(self) -> str:
        """The text in the current view (insertions in, deletions out)."""
        return _text.paragraph_text(self._element)

    @text.setter
    def text(self, value: str) -> None:
        self.set_text(value)

    def text_in(self, view: str) -> str:
        """The text in ``current``, ``original`` or ``markup`` view."""
        return _text.paragraph_text(self._element, view)

    @property
    def style(self) -> str | None:
        """The declared paragraph style's id (``w:pStyle``), or ``None`` (the default).
        Set it by name (``paragraph.style = "Heading 2"``): see :meth:`set_style`."""
        node = self._element.find(qn("w:pPr"))
        node = node.find(qn("w:pStyle")) if node is not None else None
        return node.get(qn("w:val")) if node is not None else None

    @style.setter
    def style(self, name: str | None) -> None:
        self.set_style(name)

    @property
    def style_name(self) -> str | None:
        """The applied paragraph style's name (the default style's when none is declared)."""
        styles = self._document.styles
        declared = self.style
        found = styles.find(declared, "paragraph") if declared else None
        found = found or styles.default("paragraph")
        return found.name if found else None

    def set_style(self, name: str | None) -> EditResult:
        """Apply a paragraph style by name (or alias, or id): a localised template's
        "heading 1" is found whatever its id; a built-in style the document lacks is added
        as Word writes it.  ``None`` returns the paragraph to the default style."""
        return self._document.set_paragraph_style(self._id, name)

    def format(self, **values) -> EditResult:
        """Direct formatting: paragraph properties (``alignment``, ``indent_left``,
        ``space_after``, ``keep_with_next``...) and run properties (``bold``, ``size``...),
        the latter on every run and the paragraph mark.  ``None`` removes a value."""
        return self._document.format_paragraph(self._id, **values)

    def clear_direct_formatting(self, *, paragraph: bool = True) -> EditResult:
        """Remove the direct formatting of the paragraph's runs, its mark and (unless
        ``paragraph=False``) the paragraph itself; styles and list membership stay."""
        return self._document.clear_formatting(self._id, paragraph=paragraph)

    @property
    def effective(self) -> "_effective.EffectiveParagraph":
        """The paragraph's formatting as Word applies it (docx2svg's resolver)."""
        return _effective.paragraph(self._document, self._element)

    @property
    def hyperlinks(self) -> list[Hyperlink]:
        """The paragraph's hyperlinks, in order."""
        count = sum(1 for _ in self._element.iter(qn("w:hyperlink")))
        return [Hyperlink(self._document, self.id, k) for k in range(count)]

    def coalesce_runs(self) -> EditResult:
        """Merge adjacent runs whose properties differ only in revision save ids (``w:rsid*``),
        which Word leaves behind; never done implicitly, since it changes untouched bytes."""
        return self._document.coalesce_runs(self._id)

    # -- lists ---------------------------------------------------------------------------------

    @property
    def list(self) -> ListMembership | None:
        """The list the paragraph is in (instance, level, format), or ``None``."""
        return self._document.list_of(self._id)

    def add_to_list(self, kind: "str | int" = "bullet", *, level: int = 0,
                    continue_previous: bool = True) -> EditResult:
        """Make the paragraph a list item (``bullet``, ``number``, or a numId to join)."""
        return self._document.add_to_list(self._id, kind, level=level, continue_previous=continue_previous)

    @property
    def list_level(self) -> int | None:
        """The paragraph's list level (0-8), or ``None`` when it is not a list item."""
        found = self.list
        return found.level if found is not None else None

    @list_level.setter
    def list_level(self, level: int) -> None:
        self._document.set_list_level(self._id, level)

    def indent_list(self) -> EditResult:
        """One list level deeper."""
        return self._document.set_list_level(self._id, (self.list_level or 0) + 1)

    def outdent_list(self) -> EditResult:
        """One list level shallower."""
        return self._document.set_list_level(self._id, (self.list_level or 0) - 1)

    def restart_numbering(self, at: int = 1) -> EditResult:
        """Restart its list's numbering here: :meth:`Document.restart_numbering`."""
        return self._document.restart_numbering(self._id, at)

    def continue_numbering(self, from_id: str | None = None) -> EditResult:
        """Continue an earlier list's numbering: :meth:`Document.continue_numbering`."""
        return self._document.continue_numbering(self._id, from_id)

    def remove_from_list(self) -> EditResult:
        """Take the paragraph out of its list: :meth:`Document.remove_from_list`."""
        return self._document.remove_from_list(self._id)

    def declared(self, name: str):
        """A paragraph property as the paragraph itself declares it, or ``None``."""
        if name not in _formatting.PARAGRAPH_PROPERTIES:
            raise KeyError(f"no paragraph property {name!r}")
        return _formatting.read_paragraph(self._element.find(qn("w:pPr")), name)

    @property
    def ends_section(self) -> bool:
        """Whether the paragraph's mark ends a section (it holds a ``w:sectPr``)."""
        properties = self._element.find(W_PPR)
        return properties is not None and properties.find(W_SECTPR) is not None

    @property
    def runs(self) -> list["Run"]:
        """The paragraph's runs in document order, inside hyperlinks, insertions and
        controls too: ``[run.text for run in paragraph.runs]``."""
        return [Run(self, k) for k in range(len(_text.runs(self._element)))]

    def run(self, index: int) -> "Run":
        """Run ``index`` (negative counts from the end); ``IndexError`` if none."""
        count = len(_text.runs(self._element))
        if not -count <= index < count:
            raise IndexError(f"{self.id} has {count} runs; no r{index}")
        return Run(self, index % count)

    def set_text(self, value: str) -> EditResult:
        """Replace the text, keeping each surviving character's formatting; new characters
        take the formatting of what they replace or follow (ROADMAP.md, "Edit operations")."""
        return self._document._set_text(self._id, value)

    def range(self, start: int = 0, end: int | None = None, *, view: str = "current") -> TextRange:
        """Characters ``[start, end)`` of the paragraph's text (to its end by default)."""
        length = len(self.text_in(view))
        end = length if end is None else end
        if not 0 <= start <= end <= length:
            raise IndexError(f"{self.id} has {length} characters; no range {start}:{end}")
        return TextRange(self._document, self.id, start, self.id, end, view)

    def find(self, text: str, **options) -> list[TextRange]:
        """Every place ``text`` occurs in this paragraph (:meth:`Document.find`'s options)."""
        return self._document.find(text, within=self.id, **options)

    def insert_after(self, text: str = "", *, style: str | None = None) -> EditResult:
        """A new paragraph after this one: ``paragraph.insert_after("Text", style="Heading 2")``."""
        return self._document.insert_paragraph(text, after=self._id, style=style)

    def insert_before(self, text: str = "", *, style: str | None = None) -> EditResult:
        """A new paragraph before this one (:meth:`Document.insert_paragraph`)."""
        return self._document.insert_paragraph(text, before=self._id, style=style)

    def delete(self) -> EditResult:
        """Delete the paragraph (tracked when tracking): :meth:`Document.delete_block`."""
        return self._document.delete_block(self._id)

    def move(self, *, after: str | None = None, before: str | None = None) -> EditResult:
        """Move the paragraph beside another block of its story, keeping its id."""
        return self._document.move_block(self._id, after=after, before=before)

    def __repr__(self) -> str:
        text = self.text
        return f"<Paragraph {self.id} {text[:40]!r}{'...' if len(text) > 40 else ''}>"


class Run:
    """The ``index``-th run of a paragraph in the current view: positional, re-resolved on
    every access, as runs have no identity in OOXML."""

    def __init__(self, paragraph: Paragraph, index: int) -> None:
        self._paragraph = paragraph
        self._index = index

    @property
    def id(self) -> str:
        """The run's id, ``<paragraph id>/r<index>``: positional, valid until the paragraph
        is next edited."""
        return f"{self._paragraph.id}/r{self._index}"

    @property
    def _element(self) -> Element:
        runs = _text.runs(self._paragraph._element)
        if self._index >= len(runs):
            raise KeyError(f"{self.id} no longer exists")
        return runs[self._index]

    @property
    def text(self) -> str:
        """The run's text in the current view."""
        return _text.run_text(self._paragraph._element, self._element)

    @property
    def style(self) -> str | None:
        """The declared character style's id, or ``None``; set it by name."""
        node = self._element.find(qn("w:rPr"))
        node = node.find(qn("w:rStyle")) if node is not None else None
        return node.get(qn("w:val")) if node is not None else None

    @style.setter
    def style(self, name: str | None) -> None:
        self.range().set_style(name)

    def range(self) -> TextRange:
        """The run's characters as a range of its paragraph."""
        paragraph = self._paragraph._element
        element = self._element
        items = _text.atoms(paragraph)
        mine = [k for k, atom in enumerate(items) if atom.run is element]
        if not mine:
            start = next((k for k, atom in enumerate(items)
                          if atom.run is not None and _follows(atom.run, element)), len(items))
            return TextRange(self._paragraph._document, self._paragraph.id, start, self._paragraph.id, start)
        return TextRange(self._paragraph._document, self._paragraph.id, mine[0], self._paragraph.id, mine[-1] + 1)

    def format(self, **values) -> EditResult:
        """Direct formatting on this run as it is (``bold=True``, ``size=12``...)."""
        return self._paragraph._document.format_run(self._paragraph._id, self._index, **values)

    def declared(self, name: str):
        """A run property as the run itself declares it, or ``None`` (inherited)."""
        if name not in _formatting.RUN_PROPERTIES:
            raise KeyError(f"no run property {name!r}")
        return _formatting.read_run(self._element.find(qn("w:rPr")), name)

    @property
    def effective(self) -> "_effective.EffectiveRun":
        """The run's formatting as Word applies it (docx2svg's resolver)."""
        return _effective.run(self._paragraph._document, self._paragraph._element, self._element)

    def __repr__(self) -> str:
        return f"<Run {self.id} {self.text!r}>"


def _run_property(name: str):
    def getter(self: Run):
        return self.declared(name)

    def setter(self: Run, value) -> None:
        self.format(**{name: value})

    return property(getter, setter, doc=f"The run's declared ``{name}`` (``None``: inherited); "
                                        "setting it is direct formatting.")


for _name in sorted(_formatting.RUN_PROPERTIES):
    setattr(Run, _name, _run_property(_name))


def _paragraph_property(name: str):
    def getter(self: Paragraph):
        return self.declared(name)

    def setter(self: Paragraph, value) -> None:
        self.format(**{name: value})

    return property(getter, setter, doc=f"The paragraph's declared ``{name}`` (``None``: "
                                        "inherited); setting it is direct formatting.")


for _name in sorted(_formatting.PARAGRAPH_PROPERTIES):
    setattr(Paragraph, _name, _paragraph_property(_name))


def _follows(run: Element, element: Element) -> bool:
    """Whether ``run`` comes after ``element`` in document order."""
    for node in element.itersiblings():
        if node is run or any(d is run for d in node.iter()):
            return True
    parent = element.getparent()
    return parent is not None and parent.tag != W_P and _follows(run, parent)


class Table:
    """A ``w:tbl``, with a grid view that resolves ``w:gridSpan`` and ``w:vMerge``."""

    def __init__(self, document: "Document", identifier: str) -> None:
        self._document = document
        self._id = identifier

    def _entry(self) -> tuple[str, TableEntry]:
        part, entry = self._document._resolve(self._id)
        if not isinstance(entry, TableEntry):
            raise KeyError(f"{self._id} is not a table")
        return part, entry

    @property
    def id(self) -> str:
        """The table's id (``t:...``)."""
        return self._entry()[1].id

    @property
    def rows(self) -> list["Row"]:
        """The table's rows, in order."""
        return [Row(self, k) for k in range(len(self._entry()[1].rows))]

    def grid(self) -> list[list["Cell | None"]]:
        """Every grid position's cell, a merged cell answering at each position it covers
        (its origin's id); ``None`` where a row has no cell (``w:gridBefore``/``After``)."""
        out: list[list[Cell | None]] = []
        origins: dict[int, tuple[int, int]] = {}  # column -> origin of a vertical merge
        for r, (row, _) in enumerate(self._entry()[1].rows):
            line: list[Cell | None] = []
            before = _int(_find(row, "w:trPr/w:gridBefore"), 0)
            line.extend([None] * before)
            column = before
            for tc in _cells(row):
                span = max(1, _int(_find(tc, "w:tcPr/w:gridSpan"), 1))
                merge = _find(tc, "w:tcPr/w:vMerge")
                continues = merge is not None and merge.get(qn("w:val")) in (None, "continue")
                origin = origins.get(column) if continues else None
                if origin is None:
                    origin = (r, column)
                    for c in range(column, column + span):
                        origins[c] = origin
                cell = Cell(self, *origin)
                line.extend([cell] * span)
                column += span
            out.append(line)
        return out

    def insert_row(self, index: int | None = None, *, below: bool = True) -> EditResult:
        """A new row beside row ``index``: :meth:`Document.insert_row`."""
        return self._document.insert_row(self._id, index, below=below)

    def delete_row(self, index: int) -> EditResult:
        """Delete row ``index``: :meth:`Document.delete_row`."""
        return self._document.delete_row(self._id, index)

    def insert_column(self, index: int, *, right: bool = True) -> EditResult:
        """A new column beside grid column ``index``: :meth:`Document.insert_column`."""
        return self._document.insert_column(self._id, index, right=right)

    def delete_column(self, index: int) -> EditResult:
        """Delete grid column ``index``: :meth:`Document.delete_column`."""
        return self._document.delete_column(self._id, index)

    def merge(self, first: tuple[int, int], last: tuple[int, int]) -> EditResult:
        """Merge the cells from ``first`` to ``last`` (``(row, column)``, inclusive)."""
        return self._document.merge_cells(self._id, first, last)

    def split(self, row: int, column: int, *, rows: int = 1, columns: int = 2) -> EditResult:
        """Split the cell at ``(row, column)`` (:meth:`Document.split_cell`)."""
        return self._document.split_cell(self._id, row, column, rows=rows, columns=columns)

    def set(self, **values) -> EditResult:
        """Set the table's properties (:meth:`Document.set_table`)."""
        return self._document.set_table(self._id, **values)

    def set_column_width(self, column: int, width: float) -> EditResult:
        """A grid column's width in points: :meth:`Document.set_column_width`."""
        return self._document.set_column_width(self._id, column, width)

    def set_row(self, row: int, **values) -> EditResult:
        """A row's height and settings: :meth:`Document.set_row`."""
        return self._document.set_row(self._id, row, **values)

    def set_cell(self, row: int, column: int, **values) -> EditResult:
        """A cell's shading, borders, alignment... -- not its text: :meth:`Document.set_cell`."""
        return self._document.set_cell(self._id, row, column, **values)

    def format_cell(self, row: int, column: int, **values) -> EditResult:
        """A cell's formatting, as :meth:`set_cell` (its text is ``cell(r, c).paragraphs[0]``)."""
        return self._document.set_cell(self._id, row, column, **values)

    @property
    def properties(self) -> dict:
        """What the table states itself (:meth:`Document.table_properties`)."""
        return self._document.table_properties(self._id)

    @property
    def column_count(self) -> int:
        """The number of grid columns (``w:tblGrid``)."""
        part, entry = self._entry()
        grid = entry.element.find(qn("w:tblGrid"))
        return len(grid.findall(qn("w:gridCol"))) if grid is not None else 0

    def cell(self, row: int, column: int) -> "Cell":
        """The cell over grid position ``(row, column)`` (0-based); ``IndexError`` if none.
        ``table.cell(0, 1).text``."""
        grid = self.grid()
        found = grid[row][column] if 0 <= row < len(grid) and 0 <= column < len(grid[row]) else None
        if found is None:
            raise IndexError(f"{self.id} has no cell at ({row}, {column})")
        return found

    def __repr__(self) -> str:
        return f"<Table {self.id} {len(self.rows)} rows>"


class Row:
    """A table row (``table.rows[0]``); its cells are ``table.cell(row, column)``."""

    def __init__(self, table: Table, index: int) -> None:
        self._table = table
        self._index = index

    @property
    def id(self) -> str:
        """The row's id (``tr:...``)."""
        return self._table._entry()[1].rows[self._index][1]

    def __repr__(self) -> str:
        return f"<Row {self.id}>"


class Cell:
    """The cell whose grid origin is (``row``, ``column``)."""

    def __init__(self, table: Table, row: int, column: int) -> None:
        self._table = table
        self.row = row
        self.column = column

    @property
    def id(self) -> str:
        """The cell's id, ``<table id>/c<row>,<column>``."""
        return f"{self._table.id}/c{self.row},{self.column}"

    @property
    def _element(self) -> Element:
        row = self._table._entry()[1].rows[self.row][0]
        column = _int(_find(row, "w:trPr/w:gridBefore"), 0)
        for tc in _cells(row):
            if column == self.column:
                return tc
            column += max(1, _int(_find(tc, "w:tcPr/w:gridSpan"), 1))
        raise KeyError(f"{self.id} no longer exists")

    @property
    def paragraphs(self) -> list[Paragraph]:
        """The cell's paragraphs (not those of a table nested in it)."""
        part, _ = self._table._entry()
        index = self._table._document._index(part)
        nested = set(map(id, self._element.iter(W_P)))
        return [Paragraph(self._table._document, entry.id) for entry in index.paragraphs
                if id(entry.element) in nested]

    @property
    def text(self) -> str:
        """The cell's text, its paragraphs joined by ``\\n``."""
        return "\n".join(p.text for p in self.paragraphs)

    def __eq__(self, other: object) -> bool:
        return isinstance(other, Cell) and other.id == self.id

    def __hash__(self) -> int:
        return hash(self.id)

    def __repr__(self) -> str:
        return f"<Cell {self.id}>"


class CallableList(list):
    """A list that may also be called (``doc.stories`` and ``doc.stories()`` alike): the
    collections an agent reaches by property and by method, either way."""

    def __call__(self) -> "CallableList":
        return self


class Story:
    """The blocks of one part: the body, a header, a footer, the notes, the comments."""

    def __init__(self, document: "Document", name: str, part: str) -> None:
        self._document = document
        self.name = name
        self.part = part

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Every paragraph of the story in document order: inside tables, content controls
        and text boxes too."""
        return [Paragraph(self._document, entry.id)
                for entry in self._document._index(self.part).paragraphs]

    @property
    def tables(self) -> list[Table]:
        """The story's tables, in order."""
        return [Table(self._document, entry.id) for entry in self._document._index(self.part).tables]

    def __repr__(self) -> str:
        return f"<Story {self.name} {self.part}>"


# -- the document ----------------------------------------------------------------------------


#: How the first edit stamps paraIds (ROADMAP.md, "Durable across a Word save -- measured").
STAMPING = ("document", "paragraph")


class Document(TrackingOps, RevisionOps, TextOps, FormatOps, ListOps, LinkOps, PictureOps, CommentOps, TableOps,
               TableFormatOps, DrawingOps, ControlOps, SectionOps, NoteOps, FieldOps, AnnotationOps, AuthoringOps,
               PropertyOps, ImportOps, ChartOps):
    """An open Word document, editable, undoable and re-renderable.

    ``stamping`` says what the first edit stamps.  ``"document"`` (the default): every
    paragraph and table row of every story that lacks a valid, unique ``w14:paraId``, and
    every drawing without a ``wp14:anchorId`` and ``wp14:editId`` -- the only state in
    which Word keeps paraIds on saving (measured: one paragraph, row or drawing without its
    ids, in any story, or one paraId out of range, and Word writes none at all).
    ``"paragraph"``: only what an edit touches and the volatile ids it would renumber --
    fewer bytes changed, but ids then last only until Word saves the file.
    """

    #: The application's own font folders, for every layout and render of this document
    #: -- reflow, coverage, fields' page numbers, a TOC's, PNGs: a face there is laid out
    #: and drawn with.  ``None`` (the default) reads ``OOXML_FONT_DIRS``
    #: (``os.pathsep``-separated); an empty tuple means none.  Added to the folders Word
    #: uses, never in their place.  The agent tool layer sets it from its session
    #: (``Toolbox(font_dirs=...)``); a ``font_dirs`` option given to a call wins.
    font_dirs: "tuple[str, ...] | None" = None

    def __init__(self, package: WordPackage, *, stamping: str = "document") -> None:
        if stamping not in STAMPING:
            raise ValueError(f"stamping must be one of {', '.join(STAMPING)}")
        self.stamping = stamping
        self.package = package
        self.history = History(package)
        #: Old id -> new id, for every paragraph stamped or re-issued this session.
        self._aliases: dict[str, str] = {}
        self._indexes: dict[str, PartIndex] = {}
        #: docx2svg's conversions of recent states, by their bytes' hash (layout.convert).
        self._layouts: dict[str, "_Conversion"] = {}
        #: Who lays the document out: ``None`` here, through docx2svg; or a callable
        #: ``converter(data, options) -> (layout, svgs, warnings)`` --
        #: :func:`docx_agent.layout.convert_bytes`'s answer, computed elsewhere (a worker
        #: process with a deadline).  Every layout goes through it: :meth:`layout`,
        #: :meth:`render_svg`, :meth:`update_fields`, :meth:`insert_toc`, ``state(layout=True)``.
        self.converter = None
        #: Bumped on every change: what per-state caches (the effective-formatting model) key on.
        self._state = 0
        self._effective = None
        #: The tracking context (``tracking()``) and the per-edit overrides (``track=``).
        self._tracking = None
        self._track_stack: list = []

    # -- lifecycle ---------------------------------------------------------------------------

    @classmethod
    def open(cls, source: str | os.PathLike[str] | bytes | BinaryIO, *,
             stamping: str = "document") -> "Document":
        """Open a ``.docx`` (or ``.docm``, ``.dotx``...) from a path, bytes or a binary
        file.  Nothing is changed until you edit: ``Document.open("report.docx")``.

        Opening a **template** (``.dotx``, ``.dotm``) opens the template itself, to edit it,
        and warns (:class:`~docx_agent.edit.authoring.TemplateOpened`): to make a document
        from a template, as Word's File > New does, use ``Document.new(template=...)``.
        Either way :meth:`save` writes the kind of file its extension names.

        ``stamping`` decides when paragraphs without Word's ``w14:paraId`` get one
        (``"document"``: all of them, on the first edit; ROADMAP.md, "Addressing and stable
        ids")."""
        package = WordPackage.open(source)
        package.document_part()  # refuse a package that is not a Word document
        if package.kind in ("dotx", "dotm"):
            import warnings

            from .authoring import TemplateOpened

            name = os.path.basename(os.fspath(source)) if isinstance(source, (str, os.PathLike)) else "this file"
            warnings.warn(TemplateOpened(
                f"{name} is a template (.{package.kind}): editing it edits the template. To make a document "
                f"from it use Document.new(template=...); save('x.docx') writes a document either way"),
                stacklevel=2)
        return cls(package, stamping=stamping)

    @classmethod
    def new(cls, *, page=None, orientation: str | None = None, language: str | None = None, locale=None,
            title: str | None = None, author: str | None = None, created=None, template=None,
            keep_content: bool = True, attach: bool | None = None, stamping: str = "document") -> "Document":
        """A new document, in compatibility mode 15 (ROADMAP.md, "Decisions").

        **Blank** (no ``template``): what Word makes for a new blank document
        (:mod:`docx_agent.edit.blank`): ``page`` "A4" (the default), "Letter", "Legal",
        "A5", "A3" or ``(width, height)`` in points; ``orientation`` "portrait" (the
        default) or "landscape"; ``language`` the text's (``w:lang``, the default
        "en-US"); ``locale`` the margins, distances, tab stop, hyphenation zone and number
        separators Word takes from its locale (:class:`blank.Locale`; :data:`blank.METRIC`,
        measured here, for every page but Letter, which takes :data:`blank.US`).

        **From a template** (``template``: a ``.dotx``, ``.dotm``, ``.docx`` or ``.docm``,
        a path or bytes): as Word's File > New from a template (:mod:`.authoring`): the
        template's package kept, its body too unless ``keep_content`` is false, the main
        part a document's, the template attached only with ``attach=True`` (Word attaches a
        ``.dotx``; attached, Word for Mac asks for access to it on opening), the properties started
        again; ``page``, ``orientation`` and ``language`` replace the template's when
        given.  A macro-enabled template's VBA project is dropped
        (:class:`~docx_agent.edit.authoring.MacrosDropped`).

        ``title`` and ``author`` go into ``docProps/core.xml``, ``created`` (default now; a
        datetime or an ISO 8601 string, as ``tracking(date=...)`` takes) is its creation and
        modification date.  The new document is the base state undo
        returns to."""
        from .authoring import new

        return new(cls, page=page, orientation=orientation, language=language, locale=locale, title=title,
                   author=author, created=created, template=template, keep_content=keep_content, attach=attach,
                   stamping=stamping)

    def save(self, target: str | os.PathLike[str] | BinaryIO, *, validate: bool = False) -> None:
        """Write the document: ``doc.save("report-edited.docx")``.  ``docProps/app.xml``'s
        statistics are brought up to date when the body changed
        (:func:`docx_agent.edit.properties.refreshed_app`).

        **The extension decides the kind of file**: the main part's content type is set to
        a document's for ``.docx`` and ``.docm`` (macro-enabled) and to a template's for
        ``.dotx`` and ``.dotm``, as Word's Save As does -- Word refuses a file whose content
        type is another kind's, so a template opened and saved as ``.docx`` is a document.
        Nothing else differs from :meth:`to_bytes`, but for a VBA project written to a
        ``.docx`` or ``.dotx``, which is dropped from the file with a
        :class:`~docx_agent.edit.authoring.MacrosDropped` warning.  Another extension, or a
        file object, gets the package as it is.  The open document does not change.

        With ``validate=True``, :meth:`validate` runs first and an :class:`EditError`
        listing the problems is raised, and nothing written, if it finds any."""
        from .authoring import kind_for, write_as
        from .properties import REL_APP, _part, refreshed_app

        if validate:
            problems = self.validate()
            if problems:
                listed = "; ".join(str(problem) for problem in problems[:10])
                more = f" and {len(problems) - 10} more" if len(problems) > 10 else ""
                raise EditError(f"not saved: {len(problems)} validity problem(s): {listed}{more}")

        app = refreshed_app(self)
        kind = kind_for(target)
        if kind is not None and self.package.kind is not None and kind != self.package.kind:
            write_as(self, target, kind, app)
            return
        if not isinstance(target, (str, os.PathLike)):
            target.write(self.package.to_bytes({_part(self.package, REL_APP): app} if app is not None else None))
            return
        self.package.save(target, {_part(self.package, REL_APP): app} if app is not None else None)

    def to_bytes(self) -> bytes:
        """The document as ``.docx`` bytes, as :meth:`save` would write them (less the
        refreshed statistics)."""
        return self.package.to_bytes()

    def validate(self, target: str | os.PathLike[str] | None = None) -> list:
        """The structural problems Word would repair or refuse (:func:`docx_agent.validate.check`):
        schema order, relationships, content types, unique ids, paired ranges, numbering,
        style references, comment parts.  An empty list means none were found::

            problems = doc.validate()
            assert not problems, "\\n".join(map(str, problems))

        With ``target`` (a file name), also whether the package's kind fits its extension
        (``content-type-extension``: a template's content type in a ``.docx`` is a file
        Word refuses).  :meth:`save` sets the kind from the extension itself; this is for
        bytes written another way.

        Each is a :class:`docx_agent.validate.Problem` (``code``, ``part``, ``detail``).
        No edit adds one; a document may arrive with some."""
        from ..validate import check

        return check(self.package, target=target)

    @property
    def compatibility_mode(self) -> int:
        """The document's compatibility mode (15 for Word 2013 and later, 14, or 12 when it
        names none).  Edits never change it."""
        return self.package.compatibility_mode()

    # -- stories ------------------------------------------------------------------------------

    def _parts(self) -> list[str]:
        return [self.package.document_part()] + self.package.story_parts()

    def _story_of(self, part: str) -> str:
        return story_name(self.package, part)

    @property
    def stories(self) -> list[Story]:
        """Every story: the body, then headers, footers, footnotes, endnotes, comments.  A
        property (``doc.stories``); ``doc.stories()`` works too."""
        return CallableList(Story(self, self._story_of(part), part) for part in self._parts())

    def story(self, name: str = "body") -> Story:
        """A story by name (``"body"``, ``"header1"``, ``"footnotes"``...); ``KeyError`` if none."""
        for story in self.stories:
            if story.name == name:
                return story
        raise KeyError(f"no story {name!r}")

    @property
    def body(self) -> Story:
        """The body story."""
        return self.story("body")

    def paragraphs(self, story: str = "body") -> list[Paragraph]:
        """The paragraphs of a story (the body by default), in order, those in its tables
        and controls too::

            for paragraph in doc.paragraphs():
                print(paragraph.id, paragraph.text)"""
        return self.story(story).paragraphs

    def tables(self, story: str = "body") -> list[Table]:
        """The tables of a story (the body by default), in order."""
        return self.story(story).tables

    # -- addressing ---------------------------------------------------------------------------

    @property
    def aliases(self) -> dict[str, str]:
        """Every rename this session: old id -> new id."""
        return dict(self._aliases)

    def _index(self, part: str) -> PartIndex:
        index = self._indexes.get(part)
        if index is None:
            root = self.package.tree(part)
            if root is None:
                raise KeyError(f"part {part} is missing")
            index = PartIndex(self._story_of(part), root)
            self._indexes[part] = index
        return index

    def _invalidate(self) -> None:
        """Something changed: every index and cached layout may be stale."""
        self._indexes.clear()
        self._state += 1

    def _part_for_story(self, story: str) -> str | None:
        for part in self._parts():
            if self._story_of(part) == story:
                return part
        return None

    def _direct(self, identifier: str) -> tuple[str, ParagraphEntry | TableEntry] | None:
        if identifier[:2] in ("p@", "t@"):
            story = identifier[2:].rpartition("/")[0]
        elif "/" in identifier:
            story = identifier.partition("/")[0]
        else:
            story = "body"
        part = self._part_for_story(story)
        if part is None:
            return None
        entry = self._index(part).by_id.get(identifier)
        return (part, entry) if entry is not None else None

    def _resolve(self, identifier: str) -> tuple[str, ParagraphEntry | TableEntry]:
        """The part and entry an id names now.  An alias wins while its target exists, so a
        renamed paragraph is found under its old id, and -- once the rename is undone -- the
        old id names the paragraph directly again."""
        seen = set()
        target = identifier
        while target in self._aliases and target not in seen:
            seen.add(target)
            target = self._aliases[target]
        if target != identifier:
            found = self._direct(target)
            if found is not None:
                return found
        found = self._direct(identifier)
        if found is None:
            raise KeyError(f"no paragraph or table {identifier!r}")
        return found

    def paragraph(self, identifier: str) -> Paragraph:
        """A paragraph by id (``p:3B212964``, ``p@body/7``, an alias); ``KeyError`` if none.
        ``doc.paragraph("p:3B212964").set_text("New text.")``."""
        self._resolve(identifier)
        return Paragraph(self, identifier)

    def table(self, identifier: str) -> Table:
        """A table by id (``t:...``); ``KeyError`` if none."""
        self._resolve(identifier)
        return Table(self, identifier)

    def get(self, identifier: str):
        """Whatever an id names: a paragraph, run, table, row, cell, section, story,
        hyperlink (``hl:``), bookmark (``bm:``), picture or other drawing (``d:``, a group's
        member ``d:<group>/<member>``), note
        (``fn:``, ``en:``), comment (``c:``, ``c#``), revision (``rev:``) or content control
        (``cc:``, ``cc@``), field (``fld:``) -- every id ``to_markdown`` and ``state`` write."""
        kinds = (("hl:", self.hyperlink), ("bm:", self.bookmark), ("d:", self.drawing), ("fld:", self.field),
                 ("fn:", self.note), ("en:", self.note), ("c:", self.comment), ("c#", self.comment),
                 ("rev:", self.revision), ("cc:", self.content_control), ("cc@", self.content_control))
        for prefix, find in kinds:
            if identifier.startswith(prefix):
                return find(identifier)
        if ":" not in identifier and "@" not in identifier and "/" not in identifier:
            return self.story(identifier)
        if identifier == "s:body" or identifier.startswith(("s:", "s@")):
            for section in self.sections():
                if section.id == identifier or section.id == self._renamed_section(identifier):
                    return section
            raise KeyError(f"no section {identifier!r}")
        if identifier.startswith("tr:") or "/tr:" in identifier:
            for story in self.stories:
                for table in story.tables:
                    for row in table.rows:
                        if row.id == identifier:
                            return row
            raise KeyError(f"no row {identifier!r}")
        match = re.fullmatch(r"(.+)/(r\d+|c\d+,\d+)", identifier)
        if match is not None:
            base, step = match.groups()
            try:
                part, entry = self._resolve(base)
            except KeyError:
                pass
            else:
                if isinstance(entry, ParagraphEntry) and step.startswith("r"):
                    return Paragraph(self, base).run(int(step[1:]))
                if isinstance(entry, TableEntry) and step.startswith("r"):
                    return Table(self, base).rows[int(step[1:])]
                if isinstance(entry, TableEntry):
                    r, c = map(int, step[1:].split(","))
                    return Table(self, base).cell(r, c)
        part, entry = self._resolve(identifier)
        if isinstance(entry, ParagraphEntry):
            return Paragraph(self, identifier)
        return Table(self, identifier)

    def _renamed_section(self, identifier: str) -> str:
        if identifier.startswith(("s:", "s@")):
            paragraph = "p" + identifier[1:]
            target = self._aliases.get(paragraph)
            if target is not None:
                return _section_id(target)
        return identifier

    # -- stamping -----------------------------------------------------------------------------

    def _used(self) -> set[int]:
        roots = [self.package.tree(p) for p in self._parts()]
        return _ids.used_long_hex([r for r in roots if r is not None])

    def _rename(self, renames: dict[str, str]) -> None:
        for old, new in renames.items():
            self._aliases[old] = new
            # An alias that pointed at the old id follows it to the new one.
            for key, value in list(self._aliases.items()):
                if value == old:
                    self._aliases[key] = new

    def _stamp(self, part: str, entries: list[ParagraphEntry],
               used: set[int] | None = None) -> dict[str, str]:
        """Give each entry a fresh paraId and textId; record the renames as aliases."""
        if not entries:
            return {}
        used = self._used() if used is None else used
        prefix = "" if self._story_of(part) == "body" else f"{self._story_of(part)}/"
        renames: dict[str, str] = {}
        for entry in entries:
            para_id, text_id = _ids.generate(part, used, 2)
            _ids.stamp(entry.element, para_id, text_id)
            renames[entry.id] = f"{prefix}p:{para_id}"
        _ids.ensure_w14(self.package.tree(part))
        self.package.mark_dirty(part)
        self._rename(renames)
        return renames

    def _stamp_document(self) -> dict[str, str]:
        """Stamp every paragraph and row, in every story, that lacks a valid and unique
        paraId, and every drawing that lacks an anchorId and an editId: what Word needs to
        keep any of them (ROADMAP.md, "Durable across a Word save -- measured").  Returns
        the renames; nothing when the document is complete."""
        used: set[int] | None = None
        renames: dict[str, str] = {}
        for part in self._parts():
            index = self._index(part)
            paragraphs = [e for e in index.paragraphs if e.needs_stamp]
            prefix = "" if index.story == "body" else f"{index.story}/"
            rows = [(table, k, row, row_id) for table in index.tables
                    for k, (row, row_id) in enumerate(table.rows) if not row_id.startswith(prefix + "tr:")]
            drawings = _ids.drawings_needing_ids(index.root)
            if not paragraphs and not rows and not drawings:
                continue
            if used is None:
                used = self._used()
            renames.update(self._stamp(part, paragraphs, used))
            table_renames: dict[str, str] = {}
            for table, k, row, row_id in rows:
                para_id, text_id = _ids.generate(part + "\0row", used, 2)
                _ids.stamp(row, para_id, text_id)
                renames[row_id] = f"{prefix}tr:{para_id}"
                if k == 0:
                    table_renames[table.id] = f"{prefix}t:{para_id}"
            renames.update(table_renames)
            for drawing in drawings:
                anchor_id, edit_id = _ids.generate(part + "\0drawing", used, 2)
                drawing.set(_ids.ANCHOR_ID, anchor_id)
                drawing.set(_ids.EDIT_ID, edit_id)
            self._rename({key: value for key, value in renames.items() if key not in self._aliases})
            _ids.ensure_w14(self.package.tree(part), drawings=bool(drawings))
            self.package.mark_dirty(part)
        if renames:
            self._invalidate()
        return renames

    def _prepare(self, part: str, entries: list[ParagraphEntry], position: int | None = None,
                 *, including: bool = False) -> dict[str, str]:
        """Stamp before an edit, per the stamping policy: the whole document, or the
        ``entries`` an edit touches plus the volatile ids a structural edit at ``position``
        would renumber."""
        if self.stamping == "document":
            return self._stamp_document()
        stale = [e for e in entries if e.needs_stamp]
        for entry in list(stale):
            if entry.state == "repeat":
                # Re-issuing this repeat would renumber the later repeats of its paraId.
                stale += [e for e in self._index(part).paragraphs
                          if e.state == "repeat" and e.index > entry.index and e not in stale
                          and (e.raw or "").upper() == (entry.raw or "").upper()]
        if position is not None:
            stale += [e for e in self._freeze_after(part, position, including=including) if e not in stale]
        return self._stamp(part, stale)

    def _prepare_many(self, targets: list[tuple[str, list[ParagraphEntry]]]) -> dict[str, str]:
        """:meth:`_prepare` for an edit that touches paragraphs in several places."""
        if self.stamping == "document":
            return self._stamp_document()
        renames: dict[str, str] = {}
        by_part: dict[str, list[ParagraphEntry]] = {}
        for part, entries in targets:
            bucket = by_part.setdefault(part, [])
            bucket += [e for e in entries if all(e is not b for b in bucket)]
        for part, entries in by_part.items():
            renames.update(self._prepare(part, entries))
        return renames

    def _renew_text_id(self, part: str, entry: ParagraphEntry) -> None:
        """A paragraph whose text changed gets a new textId (it had one stamped, or Word's)."""
        if entry.element.get(_ids.PARA_ID) is not None:
            self._new_text_id(part, entry.element)

    def _release(self, part: str, rel_ids: list[str]) -> list[str]:
        """Drop the relationships an edit stopped using, and the parts (media) nothing uses
        any more: ooxml-edit's release and reap.  Returns the parts removed."""
        if not rel_ids:
            return []
        self.package.mark_dirty(part)
        return self.package.release(part, rel_ids)

    def coalesce_runs(self, identifier: str | None = None) -> EditResult:
        """:meth:`Paragraph.coalesce_runs` on one paragraph, or on every paragraph."""
        targets = ([self._paragraph_entry(identifier)] if identifier is not None else
                   [(part, entry) for part in self._parts() for entry in self._index(part).paragraphs])
        if not targets:
            return EditResult(None, changed=False)
        with self._edit():
            renames = self._prepare_many([(part, [entry]) for part, entry in targets])
            merged = 0
            for part, entry in targets:
                count = _inline.coalesce(entry.element)
                if count:
                    self.package.mark_dirty(part)
                merged += count
        first = targets[0][1].id
        return EditResult(renames.get(first, first), renamed=renames, changed=merged > 0 or bool(renames),
                          count=merged)

    def _fresh_ids(self, part: str) -> tuple[str, str]:
        para_id, text_id = _ids.generate(part, self._used(), 2)
        return para_id, text_id

    def _freeze_after(self, part: str, position: int, *, including: bool = False) -> list[ParagraphEntry]:
        """The volatile paragraphs a structural edit at ``position`` would renumber."""
        return [entry for entry in self._index(part).paragraphs
                if entry.volatile and (entry.index > position or (including and entry.index == position))]

    # -- edits --------------------------------------------------------------------------------

    @contextmanager
    def _edit(self) -> Iterator[None]:
        """One undo step; on failure everything is rolled back and the caches dropped."""
        outermost = not self.history.in_batch
        try:
            with self.history.batch():
                before = self._revision_id_counts() if outermost else None
                boxes = self._text_box_snapshot() if outermost else None
                yield
                if outermost:
                    self._sync_text_boxes(boxes)
                    self._unique_revision_ids(before)
        finally:
            self._invalidate()

    def _text_box_snapshot(self) -> list[tuple[str, Element, bytes]]:
        """Every text box's content as Word reads it (the ``mc:Choice``), before an edit."""
        from .drawings import _MC

        out = []
        for part in self._parts():
            for holder in self.package.tree(part).iter(_MC + "AlternateContent"):
                choice = holder.find(_MC + "Choice")
                content = next(choice.iter(_W + "txbxContent"), None) if choice is not None else None
                if content is not None:
                    out.append((part, holder, etree.tostring(content)))
        return out

    def _sync_text_boxes(self, before: list[tuple[str, Element, bytes]]) -> None:
        """After an edit: every text box whose content changed gets the same content in its
        VML fallback, as Word writes a text box (paraIds and all; a revision record copied
        gets an id of its own afterwards)."""
        from .drawings import _MC, sync_text_box

        known = {id(holder): data for _, holder, data in before} if before is not None else {}
        for part in self._parts():
            changed = False
            for holder in self.package.tree(part).iter(_MC + "AlternateContent"):
                choice = holder.find(_MC + "Choice")
                content = next(choice.iter(_W + "txbxContent"), None) if choice is not None else None
                if content is None or (before is not None and known.get(id(holder)) == etree.tostring(content)):
                    continue
                changed = sync_text_box(holder) or changed
            if changed:
                self.package.mark_dirty(part)

    def _revision_id_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for part in self._parts():
            for node in self.package.tree(part).iter(*REVISION_KINDS_ALL):
                raw = node.get(_W + "id")
                if raw is not None:
                    counts[raw] = counts.get(raw, 0) + 1
        return counts

    def _unique_revision_ids(self, before: dict[str, int]) -> None:
        """Give a revision record whose ``w:id`` an edit repeated an id of its own: an edit
        that splits a run copies its ``w:rPrChange`` with the rest of its properties (Word
        renumbers them all on saving).  Repeats the document had before are left alone."""
        after = self._revision_id_counts()
        grown = {raw for raw, count in after.items() if count > before.get(raw, 0) and count > 1}
        if not grown:
            return
        largest = self._next_annotation_id() - 1
        seen: set[str] = set()
        for part in self._parts():
            for node in self.package.tree(part).iter(*REVISION_KINDS_ALL):
                raw = node.get(_W + "id")
                if raw not in grown:
                    continue
                if raw not in seen:
                    seen.add(raw)
                    continue
                largest += 1
                node.set(_W + "id", str(largest))
                self.package.mark_dirty(part)

    def _set_text(self, identifier: str, value: str) -> EditResult:
        _text.check_text(value)
        part, entry = self._resolve(identifier)
        if not isinstance(entry, ParagraphEntry):
            raise EditError(f"{identifier} is not a paragraph")
        if _text.paragraph_text(entry.element) == value:
            return EditResult(entry.id, changed=False)
        tracking = self._active_tracking()
        with self._edit():
            stamped = entry.needs_stamp
            renames = self._prepare(part, [entry])
            element = entry.element
            if tracking is not None:
                stamp = Stamp(self, tracking)
                old = _text.paragraph_text(element)
                released = _track.replace_span(element, 0, len(old), value, stamp, part, keep="characters")
                stamp.finish()
                self._release(part, released)
            else:
                _text.set_text(element, value)
            if not stamped:
                self._new_text_id(part, element)
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames)

    def _new_text_id(self, part: str, element: Element) -> None:
        """A paragraph whose text changed gets a new ``w14:textId``, as Word gives one: the
        textId versions the text (MS-DOCX 2.6.1.1)."""
        element.set(_ids.TEXT_ID, _ids.generate(part + "\0text", self._used(), 1)[0])
        _ids.ensure_w14(self.package.tree(part))

    def insert_paragraph(self, text: str = "", *, after: str | None = None,
                         before: str | None = None, style: str | None = None) -> EditResult:
        """A new paragraph after or before a block (a paragraph or a table).

        It continues the paragraph it is inserted beside: a copy of its paragraph properties
        (without a section break or a revision record) and of its nearest run's properties
        -- the last run when inserted after, the first when before.  ``style`` (a name,
        alias or id; a built-in style the document lacks is added as Word writes it)
        replaces the copied properties with that paragraph style alone.  The new paragraph is
        stamped at once, so its id is durable from the start.
        """
        if (after is None) == (before is None):
            raise EditError("give exactly one of after= or before=")
        _text.check_text(text)
        anchor_id = after if after is not None else before
        part, entry = self._resolve(anchor_id)  # type: ignore[arg-type]
        anchor = entry.element
        style_id = self.styles.resolve(style, "paragraph")[0] if style is not None else None

        paragraph = _text._make("w:p")
        template = anchor if anchor.tag == W_P else None
        if style_id is not None:
            properties = _text._make("w:pPr")
            pstyle = _text._make("w:pStyle")
            pstyle.set(qn("w:val"), style_id)
            properties.append(pstyle)
            paragraph.append(properties)
        elif template is not None and template.find(W_PPR) is not None:
            properties = copy.deepcopy(template.find(W_PPR))
            for child in list(properties):
                if child.tag in (W_SECTPR, _W + "pPrChange"):
                    remove(child)
                elif child.tag == _W + "rPr":
                    for mark in list(child):
                        if mark.tag in REVISION_PROPERTY_TAGS:
                            remove(mark)
                    if not len(child):
                        remove(child)
            if len(properties):
                paragraph.append(properties)
        if text:
            reference = None
            if template is not None and style_id is None:
                runs = [r for r in _text.runs(template) if _text.run_text(template, r)]
                if runs:
                    reference = runs[-1] if after is not None else runs[0]
            properties = (_text.run_properties_for(template, reference)
                          if template is not None and style_id is None else None)
            paragraph.append(_text.make_run(text, properties))

        position = entry.index if isinstance(entry, ParagraphEntry) else self._block_position(part, anchor, after is not None)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [], position, including=before is not None)
            if style is not None:
                self.styles._ensure(style, "paragraph")
            para_id, text_id = self._fresh_ids(part)
            _ids.stamp(paragraph, para_id, text_id)
            if after is not None:
                anchor.addnext(paragraph)
            else:
                anchor.addprevious(paragraph)
            if tracking is not None:
                self._track_new_paragraph(part, paragraph, tracking)
            _ids.ensure_w14(self.package.tree(part))
            self.package.mark_dirty(part)
            story = self._story_of(part)
            new_id = ("" if story == "body" else f"{story}/") + f"p:{para_id}"
        return EditResult(new_id, created=[new_id], renamed=renames)

    def _track_new_paragraph(self, part: str, paragraph: Element, tracking) -> None:
        """Record a paragraph just put in place as inserted.  Followed by a block of its
        container, it is inserted with its own mark (``w:pPr/w:rPr/w:ins``).  Last in its
        container -- where Word cannot accept or reject a mark (measured) -- it is written as
        Word writes a paragraph typed at the end of the one before: that one's mark is the
        inserted one, and the new paragraph carries the old mark's properties, its own
        recorded as a change from them."""
        stamp = Stamp(self, tracking)
        _track.insert_paragraph_content(paragraph, stamp, part)
        previous = paragraph.getprevious()
        while previous is not None and not (isinstance(previous.tag, str) and previous.tag in (W_P, W_TBL)):
            previous = previous.getprevious()
        if _track.is_last_in_container(paragraph) and previous is not None and previous.tag == W_P \
                and _track.mark_record(previous) is None:
            if previous.find(f"{W_PPR}/{W_SECTPR}") is None:
                _track.mark(previous, "ins", stamp, part)
                # The old mark's properties: as they were before a change recorded on them.
                change = previous.find(f"{W_PPR}/{_W}pPrChange/{W_PPR}")
                old = _track.clean_properties(change if change is not None else previous.find(W_PPR), "w:pPr")
                _track.record_paragraph_change(paragraph, old, stamp, part)
                stamp.finish()
                return
        _track.mark(paragraph, "ins", stamp, part)
        stamp.finish()

    def append_paragraph(self, text: str = "", *, story: str = "body", style: str | None = None) -> EditResult:
        """A new paragraph at the end of a story (before the body's last section
        properties): how a story with no paragraph to insert beside gets its first."""
        _text.check_text(text)
        if story.startswith("d:"):
            # A text box's story: its paragraphs live in the part it is drawn in.
            part, container = self._text_box_content(story)
            root = self.package.tree(part)
        else:
            part = self._part_for_story(story)
            if part is None:
                raise EditError(f"no story {story!r}")
            root = self.package.tree(part)
            container = root.find(W_BODY) if root.find(W_BODY) is not None else root
        style_id = None
        if style is not None:
            style_id = self.styles.resolve(style, "paragraph")[0]
        with self._edit():
            renames = self._prepare(part, [], len(self._index(part).paragraphs))
            if style is not None:
                style_id = self.styles._ensure(style, "paragraph")
            paragraph = _text._make("w:p")
            if style_id is not None:
                properties = _text._make("w:pPr")
                pstyle = _text._make("w:pStyle")
                pstyle.set(qn("w:val"), style_id)
                properties.append(pstyle)
                paragraph.append(properties)
            if text:
                paragraph.append(_text.make_run(text, None))
            para_id, text_id = self._fresh_ids(part)
            _ids.stamp(paragraph, para_id, text_id)
            if story.startswith("d:"):
                part, container = self._text_box_content(story)
            last = container.find(W_SECTPR) if container.tag == W_BODY else None
            if last is not None:
                last.addprevious(paragraph)
            else:
                container.append(paragraph)
            tracking = self._active_tracking()
            if tracking is not None:
                self._track_new_paragraph(part, paragraph, tracking)
            _ids.ensure_w14(root)
            self.package.mark_dirty(part)
            new_id = ("" if self._story_of(part) == "body" else f"{self._story_of(part)}/") + f"p:{para_id}"
        return EditResult(new_id, created=[new_id], renamed=renames)

    def delete_block(self, identifier: str) -> EditResult:
        """Delete a paragraph or a table.

        Refused, before anything changes, where E0 cannot yet keep the document whole: a
        paragraph that ends a section (E4 joins sections), one holding a note or comment
        reference (E3/E4), one that opens or closes a field another paragraph continues, and
        a block whose removal would leave a cell, the body or a story without a paragraph to
        end on.  A bookmark, comment, move or permission range that only starts or ends in a
        deleted paragraph keeps that end, moved to the next paragraph (or the previous one).
        """
        part, entry = self._resolve(identifier)
        element = entry.element
        old_id = entry.id
        if isinstance(entry, ParagraphEntry):
            properties = element.find(W_PPR)
            if properties is not None and properties.find(W_SECTPR) is not None:
                raise EditError(f"{old_id} ends a section; deleting it would join two sections")
            for tag in _REFERENCES:
                if element.find(".//" + tag) is not None:
                    raise EditError(f"{old_id} holds a {tag.rpartition('}')[2]}; deleting it "
                                    "would orphan the note or comment")
            if not _fields_balanced(element):
                raise EditError(f"{old_id} holds part of a field that spans paragraphs")
        _check_container_survives(element)

        position = entry.index if isinstance(entry, ParagraphEntry) else self._block_position(part, element, True)
        removed_ids = [old_id]
        if isinstance(entry, TableEntry):
            nested = set(map(id, element.iter(W_P)))
            removed_ids += [e.id for e in self._index(part).paragraphs if id(e.element) in nested]
        tracking = self._active_tracking()
        if tracking is not None and not self._own_insertion(element, tracking):
            with self._edit():
                renames = self._prepare(part, [], position)
                self._track_delete_block(part, element, tracking)
                self.package.mark_dirty(part)
            return EditResult(renames.get(old_id, old_id), renamed=renames)
        with self._edit():
            renames = {old: new for old, new in self._prepare(part, [], position).items()
                       if old not in removed_ids}
            self._aliases = {old: new for old, new in self._aliases.items() if old not in removed_ids}
            if tracking is not None:
                # One's own insertion, deleted, goes outright (Word: measured); what it
                # held of one's own goes with it.
                self._release(part, _inline.relationship_ids(element))
            _rehome_markers(element)
            remove(element)
            self.package.mark_dirty(part)
        return EditResult(None, renamed=renames, removed=removed_ids)

    def _own_insertion(self, element: Element, tracking) -> bool:
        """Whether a block is wholly an insertion by the tracking author: a paragraph whose
        mark and every run are, a table whose every row is."""
        author = tracking.author
        if element.tag == W_TBL:
            rows = [row for row in element.iter(_W + "tr")]
            return bool(rows) and all(
                row.find(f"{_W}trPr/{_W}ins") is not None
                and row.find(f"{_W}trPr/{_W}ins").get(_W + "author") == author for row in rows)
        record = _track.mark_record(element)
        if record is None or record.tag != _W + "ins" or record.get(_W + "author") != author:
            return False
        for run in _text.walk(element, "markup").runs:
            owner = _track.revision_owner(run)
            if owner is None or owner.tag != _W + "ins" or owner.get(_W + "author") != author:
                return False
        return True

    def _track_delete_block(self, part: str, element: Element, tracking) -> None:
        """A block's tracked deletion.  A paragraph: its content and its mark, as Word
        deletes a whole paragraph; last in its container -- where Word cannot accept a
        deleted mark (measured: a cell's pulls the cell's first paragraph out of the
        table) -- the mark before it is the deleted one, and the paragraph takes that
        paragraph's properties, recorded as a change.  A table: every row (``w:trPr/w:del``)
        with every cell's paragraph marks and content, as Word deletes rows."""
        stamp = Stamp(self, tracking)
        if element.tag == W_TBL:
            for row in element.iter(_W + "tr"):
                properties = row.find(_W + "trPr")
                if properties is None:
                    properties = _text._make("w:trPr")
                    from ..oxml.xml import insert_in_order

                    insert_in_order(row, properties)
                from ..oxml.xml import insert_in_order

                insert_in_order(properties, stamp.make("w:del", part))
            for paragraph in element.iter(W_P):
                self._release(part, _track.delete_content(paragraph, stamp, part))
                if _track.mark_record(paragraph) is None:
                    _track.mark(paragraph, "del", stamp, part)
            stamp.finish()
            return
        own_runs = list(element.iter(_W + "r"))
        self._release(part, _track.delete_content(element, stamp, part))
        previous = element.getprevious()
        while previous is not None and not (isinstance(previous.tag, str) and previous.tag in (W_P, W_TBL)) \
                or previous is not None and previous.tag == W_P and _track.mark_record(previous) is not None \
                and _track.mark_record(previous).tag in _track.DELETED:
            # Past paragraphs whose mark is deleted already: they join on into this one.
            previous = previous.getprevious()
        if _track.is_last_in_container(element) and previous is not None and previous.tag == W_P \
                and previous.find(f"{W_PPR}/{W_SECTPR}") is None:
            properties = previous.find(W_PPR)
            previous_properties = copy.deepcopy(properties) if properties is not None else _text._make("w:pPr")
            joined_id = previous.get(_ids.PARA_ID)
            _track.delete_mark(previous, stamp, part)
            _track.take_properties(element, previous_properties, stamp, part)
            if previous.getparent() is None and joined_id and element.get(_ids.PARA_ID):
                # The mark was this author's own insertion: removing it joined that paragraph
                # into this one.  When this one's content went with the deletion (it was the
                # author's insertion too), what is left is that paragraph: it keeps its id.
                story = self._story_of(part)
                prefix = "" if story == "body" else f"{story}/"
                mine = element.get(_ids.PARA_ID)
                if not any(any(a is element for a in run.iterancestors()) for run in own_runs):
                    element.set(_ids.PARA_ID, joined_id)
                    self._rename({f"{prefix}p:{mine}": f"{prefix}p:{joined_id}"})
                else:
                    self._rename({f"{prefix}p:{joined_id}": f"{prefix}p:{mine}"})
        elif _track.mark_record(element) is None:
            _track.mark(element, "del", stamp, part)
        stamp.finish()

    def move_block(self, identifier: str, *, after: str | None = None, before: str | None = None) -> EditResult:
        """Move a paragraph or table beside another block of the same story; it keeps its
        id (and its paraIds), as Word keeps them when it moves text.  Refused, before
        anything changes, for a paragraph that ends a section, a move into the block itself,
        and one that would leave a cell or story without a final paragraph."""
        if (after is None) == (before is None):
            raise EditError("give exactly one of after= or before=")
        part, entry = self._resolve(identifier)
        target_part, target = self._resolve(after if after is not None else before)  # type: ignore[arg-type]
        if part != target_part:
            raise EditError("a block moves within its story; between stories or documents is E6's copy")
        element, anchor = entry.element, target.element
        if element is anchor:
            raise EditError("a block cannot move beside itself")
        if any(node is element for node in anchor.iterancestors()):
            raise EditError(f"{identifier} holds the block it would move beside")
        if isinstance(entry, ParagraphEntry):
            properties = element.find(W_PPR)
            if properties is not None and properties.find(W_SECTPR) is not None:
                raise EditError(f"{entry.id} ends a section; moving it would move the section break")
        _check_container_survives(element)
        if after is not None and anchor.getnext() is element or before is not None and anchor.getprevious() is element:
            return EditResult(entry.id, changed=False)
        position = min(entry.index if isinstance(entry, ParagraphEntry) else self._block_position(part, element, False),
                       target.index if isinstance(target, ParagraphEntry) else self._block_position(part, anchor, False))
        tracking = self._active_tracking()
        warnings: list[str] = []
        if tracking is not None and any(True for _ in element.iter(*_REFERENCES)):
            # A copy would repeat the note or comment reference: Word's move form needs both.
            warnings.append(f"{entry.id} holds a note or comment reference: moved untracked")
            tracking = None
        if tracking is not None and not self._own_insertion(element, tracking):
            with self._edit():
                renames = self._prepare(part, [], position)
                moved = self._track_move(part, element, anchor, after is not None, tracking)
                old_id = renames.get(entry.id, entry.id)
                story = self._story_of(part)
                prefix = "" if story == "body" else f"{story}/"
                new_id = prefix + ("p:" if moved.tag == W_P else "t:") + (
                    moved.get(_ids.PARA_ID) if moved.tag == W_P else moved.find(_W + "tr").get(_ids.PARA_ID))
                self._rename({old_id: new_id})
                renames[old_id] = new_id
                self.package.mark_dirty(part)
            return EditResult(new_id, created=[new_id], renamed=renames)
        with self._edit():
            renames = self._prepare(part, [], position)
            if after is not None:
                anchor.addnext(element)
            else:
                anchor.addprevious(element)
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames, warnings=warnings)

    def section_blocks(self, heading: str) -> list[str]:
        """The ids of a heading's section: the heading and every block after it up to the
        next heading of the same or a higher level (a lower number), in order -- what to
        give :meth:`move_blocks`, :meth:`to_markdown` (``f"{ids[0]}..{ids[-1]}"``) or
        :meth:`copy_blocks`::

            section = doc.section_blocks("p:11698B22")     # "6 Data retention" and its 6.1, 6.2...
            doc.move_blocks(section, after=doc.section_blocks("p:3A1F09C2")[-1])

        A heading is a body paragraph whose style is a heading (by name, or by the outline
        level it or its style declares, as ``to_markdown`` reads ``#``).  Headings inside
        the section at deeper levels are part of it; a section break inside it is too (its
        paragraph cannot move: :meth:`move_blocks` says so)."""
        part, entry = self._resolve(heading)
        body = self.package.document_part()
        if part != body or not isinstance(entry, ParagraphEntry) or entry.element.getparent().tag != W_BODY:
            raise EditError(f"{heading} is not a paragraph of the body itself (not in a table or a control)")
        levels = _HeadingLevels(self)
        level = levels.of(entry.element)
        if level is None:
            raise EditError(f"{entry.id} is not a heading: its style has no heading level")
        out = [entry.id]
        node = entry.element.getnext()
        while node is not None:
            if isinstance(node.tag, str) and node.tag in _MOVABLE:
                if node.tag == W_P:
                    other = levels.of(node)
                    if other is not None and other <= level:
                        break
                out.append(self._block_id(part, node))
            node = node.getnext()
        return out

    def move_blocks(self, blocks: "str | list[str]", *, after: str | None = None,
                    before: str | None = None) -> EditResult:
        """Move several blocks together beside another block, in their order, as one undo
        step: a range ``"p:A..t:B"`` (the blocks from the first through the last, of one
        container) or a list of ids (:meth:`section_blocks`'s).  Each keeps its id, as
        :meth:`move_block`; tracked, each paragraph is a tracked move (``w:moveFrom`` where
        it was, ``w:moveTo`` where it goes, Word's form) and a table a deletion and an
        insertion, as Word records them.  ``result.blocks`` are the moved blocks' ids at
        their new place, in order.  Refused, before anything changes, for a target inside
        the blocks, a paragraph that ends a section, and blocks of different stories."""
        if (after is None) == (before is None):
            raise EditError("give exactly one of after= or before=")
        elements, part = self._blocks_of(blocks)
        target_part, target = self._resolve(after if after is not None else before)  # type: ignore[arg-type]
        if target_part != part:
            raise EditError("blocks move within their story; between documents is copy_blocks")
        if any(target.element is e or any(a is e for a in target.element.iterancestors()) for e in elements):
            raise EditError(f"{after or before} is one of the blocks moved, or inside one")
        for element in elements:
            if element.tag not in (W_P, W_TBL):
                raise EditError(f"{self._block_id(part, element)} is a content control: move_blocks moves "
                                "paragraphs and tables")
            if element.tag == W_P and element.find(f"{W_PPR}/{W_SECTPR}") is not None:
                raise EditError(f"{self._block_id(part, element)} ends a section; moving it would move the "
                                "section break")
        ids = [self._block_id(part, element) for element in elements]
        moved: list[str] = []
        renamed: dict[str, str] = {}
        created: list[str] = []
        warnings: list[str] = []
        with self.batch():
            previous = None
            for identifier in ids:
                identifier = renamed.get(identifier, identifier)
                if previous is None:
                    result = self.move_block(identifier, after=after, before=before)
                else:
                    result = self.move_block(identifier, after=previous)
                renamed.update(result.renamed)
                created += result.created
                warnings += result.warnings
                previous = result.id
                moved.append(result.id)
        return EditResult(moved[0] if moved else None, created=created, renamed=renamed, warnings=warnings,
                          blocks=moved)

    def _blocks_of(self, blocks: "str | list[str]") -> tuple[list[Element], str]:
        """The block elements ``blocks`` names, in document order, and their part."""
        if isinstance(blocks, str):
            first, sep, last = blocks.partition("..")
            if not sep:
                part, entry = self._resolve(first)
                return [entry.element], part
            part, start = self._resolve(first)
            end_part, end = self._resolve(last)
            if end_part != part or start.element.getparent() is not end.element.getparent():
                raise EditError(f"{blocks!r}: a range moved is blocks of one container")
            siblings = [c for c in start.element.getparent() if isinstance(c.tag, str) and c.tag in _MOVABLE]
            a, b = siblings.index(start.element), siblings.index(end.element)
            if b < a:
                raise EditError(f"{blocks!r} ends before it starts")
            return siblings[a:b + 1], part
        if not blocks:
            raise EditError("no blocks to move")
        found = [self._resolve(identifier) for identifier in blocks]
        parts = {part for part, _ in found}
        if len(parts) != 1:
            raise EditError("the blocks moved are of one story")
        part = parts.pop()
        order = {id(node): k for k, node in enumerate(self.package.tree(part).iter())}
        elements = []
        for _, entry in found:
            if not any(entry.element is e for e in elements):
                elements.append(entry.element)
        elements.sort(key=lambda e: order[id(e)])
        return elements, part

    def _block_id(self, part: str, element: Element) -> str:
        index = self._index(part)
        found = index.entry_for(element)
        if found is not None:
            return found.id
        for _, identifier, node in self._content_controls():
            if node is element:
                return identifier
        raise EditError("not a paragraph, table or content control")

    def _track_move(self, part: str, element: Element, anchor: Element, after: bool, tracking) -> Element:
        """A block's tracked move: the paragraph stays where it was, its runs in
        ``w:moveFrom`` and its mark ``w:moveFrom``, between ``w:moveFromRangeStart`` (after its
        properties) and ``w:moveFromRangeEnd`` (after it); a copy -- a new paragraph -- goes to
        the destination, the same in ``w:moveTo``, both ranges named alike: Word's form,
        measured.  The copy holds no bookmark or comment range of the original (they stay
        with it).  A paragraph holding other revisions, and a table, move as a deletion and
        an insertion.  Returns the copy."""
        from ..oxml.xml import insert_in_order

        copy_ = copy.deepcopy(element)
        used = self._used()
        for node in copy_.iter(W_P, _W + "tr"):
            para_id, text_id = _ids.generate(part + "\0move", used, 2)
            _ids.stamp(node, para_id, text_id)
        for node in list(copy_.iter(*_MARKERS, _W + "permStart", _W + "permEnd")):
            remove(node)
        for frame in copy_.iter(_ids.WP_INLINE, _ids.WP_ANCHOR):
            doc_pr = frame.find("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr")
            if doc_pr is not None:
                doc_pr.set("id", str(self._next_doc_pr()))
        if after:
            anchor.addnext(copy_)
        else:
            anchor.addprevious(copy_)
        # Word's move form needs a paragraph after each end to join (a last mark it cannot
        # accept or reject: measured); else the move is a deletion and an insertion, each in
        # the form that has one.
        plain = element.tag == W_P and not any(
            isinstance(node.tag, str) and node.tag in REVISION_KINDS_ALL for node in element.iter()) \
            and not _track.is_last_in_container(element) and not _track.is_last_in_container(copy_)
        if not plain:
            # Deleted where it was, inserted where it goes.
            self._track_delete_block(part, element, tracking)
            stamp = Stamp(self, tracking)
            paragraphs = list(copy_.iter(W_P))
            if copy_.tag == W_TBL:
                for row in copy_.iter(_W + "tr"):
                    for record in row.findall(f"{_W}trPr/{_W}del") + row.findall(f"{_W}trPr/{_W}ins"):
                        remove(record)
                    properties = row.find(_W + "trPr")
                    if properties is None:
                        properties = _text._make("w:trPr")
                        insert_in_order(row, properties)
                    insert_in_order(properties, stamp.make("w:ins", part))
            for paragraph in paragraphs:
                _clear_revisions(paragraph)
            if copy_.tag == W_P:
                stamp.finish()
                self._track_new_paragraph(part, copy_, tracking)
            else:
                for paragraph in paragraphs:
                    _track.insert_paragraph_content(paragraph, stamp, part)
                    _track.mark(paragraph, "ins", stamp, part)
                stamp.finish()
            _ids.ensure_w14(self.package.tree(part))
            return copy_
        stamp = Stamp(self, tracking)
        name = f"move{stamp.next_id()}"
        for paragraph, kind in ((element, "moveFrom"), (copy_, "moveTo")):
            start = stamp.make(f"w:{kind}RangeStart", part, name=name)
            properties = paragraph.find(W_PPR)
            if properties is not None:
                properties.addnext(start)
            else:
                paragraph.insert(0, start)
            _track.wrap_runs(_track.content_runs(paragraph), f"w:{kind}", stamp, part)
            _track.mark(paragraph, kind, stamp, part)
            end = _text._make(f"w:{kind}RangeEnd")
            end.set(_W + "id", start.get(_W + "id"))
            paragraph.addnext(end)
        stamp.finish()
        _ids.ensure_w14(self.package.tree(part))
        return copy_

    def _block_position(self, part: str, element: Element, after: bool) -> int:
        """The paragraph index a structural edit at a table is at: its last paragraph's (an
        edit after it) or the one before its first (an edit before it)."""
        nested = [e.index for e in self._index(part).paragraphs if e.element is element
                  or any(e.element is p for p in element.iter(W_P))]
        if not nested:
            return -1
        return max(nested) if after else min(nested) - 1

    # -- history ------------------------------------------------------------------------------

    @contextmanager
    def batch(self) -> Iterator[None]:
        """Collapse every edit inside into one undo step; roll all of them back on error."""
        try:
            with self.history.batch():
                yield
        finally:
            self._invalidate()

    def undo(self) -> bool:
        """Undo the last edit (one step; a :meth:`batch` is one); ``False`` if there is none.
        Undoing every step gives back the original bytes."""
        changed = self.history.undo()
        self._invalidate()
        return changed

    def redo(self) -> bool:
        """Redo the last undone edit; ``False`` if there is none."""
        changed = self.history.redo()
        self._invalidate()
        return changed

    # -- rendering and layout -----------------------------------------------------------------

    def layout(self, **options) -> "DocumentLayout":
        """docx2svg's layout of the document as it is now, mapped to this API's ids; cached
        by the document's bytes, and laid out once with :meth:`render_svg` of the same
        state.  See :class:`docx_agent.layout.DocumentLayout`::

            before = doc.layout()
            doc.paragraph("p:3B212964").set_text("A longer text...")
            doc.layout().where("p:3B212964")    # [Placement(page=1, top=..., bottom=...)]
            doc.layout().compare(before).changed  # the pages the edit changed

        ``options`` as :meth:`render_svg`'s (no ``pages``: the layout is of every page)."""
        from ..layout import check_options, lay_out

        check_options(options, "layout")
        return lay_out(self, **options)

    # -- reading as Markdown and as JSON (E2) -------------------------------------------------

    def to_markdown(self, range: str | None = None, *, view: str = "final", ids: bool = True,
                    headers: bool = False, notes: bool = True, style_map=None, stories="body") -> str:
        """The document as Markdown with ids (:func:`docx_agent.markdown.to_markdown`): reads,
        never writes, never stamps.  ``stories="all"`` adds every header and footer with
        content and every note after the body (their comments too, in ``view="markup"``);
        a list names stories (``["body", "header1"]``).  In ``view="markup"`` field results
        are marked (``<!-- field: REF ... -->...<!-- /field -->``), and a heading's id comment
        says whether a list numbers it (``numbered: list``) or its number is typed
        (``numbered: text``)."""
        from ..markdown import to_markdown

        return to_markdown(self, range, view=view, ids=ids, headers=headers, notes=notes, style_map=style_map,
                           stories=stories)

    def insert_markdown(self, md: str, *, at: str = "end", style_map=None, images=None, fetch=None,
                        html: str = "refuse", emphasis: str = "styles") -> EditResult:
        """Write Markdown as the document's own blocks (:mod:`docx_agent.markdown.write`).

        ``at``: ``"end"`` (the body's end) or ``"end:<story>"``; ``"after:<id>"``,
        ``"before:<id>"`` (or a bare id: after it); ``"replace:<id>"`` or
        ``"replace:<id>..<id>"``, blocks of one container.  Headings, paragraphs, lists,
        quotes, code, thematic breaks, GFM tables, emphasis, strong, strikethrough, inline
        code, links (external and ``#bookmark``), hard breaks, footnotes and pictures, each in
        the style ``style_map`` names (by ``w:name``; a mapped built-in style the document
        lacks is added as Word writes it).  ``style_map`` is a dict of the styles that differ
        from the default (:class:`~docx_agent.StyleMap`, :data:`StyleMap.DEFAULT`)::

            doc.insert_markdown(md, style_map={"h1": "Report Title", "h2": "Heading 1",
                                               "paragraph": "Report Body"})

        Its keys: ``paragraph`` (body paragraphs only: table cells are ``table_cell`` and
        footnotes ``footnote``, which keep the document's defaults unless named),
        ``h1``-``h6``, ``quote``, ``code``, ``inline_code``, ``bullet``, ``bullet2``-``5``,
        ``number``, ``number2``-``5``, ``table``, ``footnote_reference``, ``emphasis``,
        ``strong``, ``link``.

        At ``"end"`` of a body that is only one empty paragraph -- a new document's -- the
        Markdown replaces that paragraph, as typing into it does (tracked, it is a deleted
        paragraph that accepting removes).  Pictures come from ``images`` -- a directory a
        path must lie under, a mapping or a function -- and remote ones only through
        ``fetch``.  ``html``: ``"refuse"`` raw HTML, or keep it as ``"text"``.  ``emphasis``:
        the Emphasis and Strong ``"styles"`` (decided), or ``"direct"`` ``w:i``/``w:b``.

        One undo step; with tracking, one revision group.  ``created`` lists every new
        paragraph and table id, ``blocks`` the top-level ones in order, ``warnings`` the
        styles added."""
        from ..markdown.write import insert_markdown

        return insert_markdown(self, md, at=at, style_map=style_map, images=images, fetch=fetch, html=html,
                               emphasis=emphasis)

    def state(self, range: str | None = None, *, view: str = "final", layout: bool = False,
              xml: bool = False, style_map=None) -> dict:
        """The structured, read-only JSON state (:func:`docx_agent.state.state`)."""
        from ..state import state

        return state(self, range, view=view, layout=layout, xml=xml, style_map=style_map)

    def render_svg(self, pages: list[int] | None = None, **options) -> list[str]:
        """The pages (1-based; all by default) as SVG strings, through docx2svg, each
        paragraph group carrying its id in ``data-docx-agent-id``::

            svgs = doc.render_svg(pages=[1])

        Renders always show the **final view**: every tracked change as if accepted,
        comments not drawn (by decision; ROADMAP.md, "How docx2svg should render
        revisions").  To review revisions and comments in place, read
        ``doc.to_markdown(view="markup")``.

        ``options`` are docx2svg's :class:`~docx2svg.ConvertOptions` (``width``,
        ``height``, ``glyph_size``, ``font_dirs``...); an unknown one raises a
        ``TypeError`` that lists them."""
        from ..layout import check_options, render_svg

        check_options(options, "render_svg")
        return render_svg(self, pages, **options)

    def render_png(self, pages: list[int] | None = None, **options) -> list[bytes]:
        """The pages (1-based; all by default) as PNG bytes, through docx2svg
        (``pip install docx-agent[png]``)::

            Path("page1.png").write_bytes(doc.render_png(pages=[1])[0])

        Like :meth:`render_svg`, the final view (revisions accepted, comments hidden);
        review with ``doc.to_markdown(view="markup")``.  ``options`` as :meth:`render_svg`'s
        (``width=`` and ``height=`` set the pixel size)."""
        import docx2svg

        from ..layout import check_options

        from ..layout import layout_options

        check_options(options, "render_png")
        convert = docx2svg.ConvertOptions(pages=pages, **layout_options(self, options))
        return docx2svg.convert_docx_to_png(self.to_bytes(), convert)

    def __repr__(self) -> str:
        return f"<Document {len(self.body.paragraphs)} paragraphs, mode {self.compatibility_mode}>"


# -- helpers ---------------------------------------------------------------------------------


#: Blocks a section is made of.
_MOVABLE = (W_P, W_TBL, _W + "sdt", _W + "customXml")


class _HeadingLevels:
    """A body paragraph's heading level (1-9), as ``to_markdown`` reads ``#``: the outline
    level it states, else its style's (by name, or the outline level the style declares)."""

    def __init__(self, document: "Document") -> None:
        from ..markdown.read import Reader

        self.reader = Reader(document)

    def of(self, paragraph: Element) -> int | None:
        properties = paragraph.find(W_PPR)
        direct = properties.find(_W + "outlineLvl") if properties is not None else None
        if direct is not None:
            try:
                value = int(direct.get(_W + "val"))
            except (TypeError, ValueError):
                value = 9
            return value + 1 if 0 <= value <= 8 else None
        style = properties.find(_W + "pStyle") if properties is not None else None
        style_id = style.get(_W + "val") if style is not None else None
        name = self.reader.style_name(style_id) if style_id else None
        return self.reader.map.heading_level(name, self.reader.outline_level(style_id))


def _find(node: Element | None, path: str) -> Element | None:
    for step in path.split("/"):
        if node is None:
            return None
        node = node.find(qn(step))
    return node


def _int(node: Element | None, default: int, attribute: str = "w:val") -> int:
    if node is None:
        return default
    try:
        return int(node.get(qn(attribute)) or "")
    except ValueError:
        return default


def _cells(row: Element) -> list[Element]:
    """A row's ``w:tc``, also inside row-level ``w:sdt``/``w:customXml``."""
    out: list[Element] = []
    for child in row:
        if child.tag == W_TC:
            out.append(child)
        elif child.tag == _W + "sdt":
            content = child.find(_W + "sdtContent")
            if content is not None:
                out.extend(_cells(content))
        elif child.tag == _W + "customXml":
            out.extend(_cells(child))
    return out


def _fields_balanced(paragraph: Element) -> bool:
    depth = 0
    for node in paragraph.iter(_W + "fldChar"):
        kind = node.get(_W + "fldCharType")
        if kind == "begin":
            depth += 1
        elif kind == "end":
            depth -= 1
            if depth < 0:
                return False
    return depth == 0


_MARKERS = frozenset(_RANGE_PAIRS) | frozenset(_RANGE_ENDS)
_BLOCK_TAGS = frozenset({W_P, W_TBL, _W + "sdt", _W + "customXml", _W + "altChunk"})


def _check_container_survives(element: Element) -> None:
    """A cell, a body or a story must still end on a paragraph once ``element`` is gone."""
    parent = element.getparent()
    if parent is None:
        raise EditError("the block has no container")
    container = parent
    while container is not None and container.tag in (_W + "sdtContent", _W + "sdt", _W + "customXml"):
        container = container.getparent()
    blocks = [child for child in _direct_blocks(container) if child is not element]
    if not blocks:
        raise EditError("deleting it would leave its story or cell with no paragraph")
    last = blocks[-1]
    if last.tag == W_TBL:
        raise EditError("deleting it would leave a table last in its story or cell; Word "
                        "needs a paragraph after it")


def _direct_blocks(container: Element) -> list[Element]:
    """The blocks of a container in order, seeing through block-level content controls and
    custom XML."""
    out: list[Element] = []
    for child in container:
        if child.tag in (W_P, W_TBL):
            out.append(child)
        elif child.tag == _W + "sdt":
            content = child.find(_W + "sdtContent")
            if content is not None:
                out.extend(_direct_blocks(content))
        elif child.tag == _W + "customXml":
            out.extend(_direct_blocks(child))
    return out


def _rehome_markers(element: Element) -> None:
    """Move the range markers whose partner is outside ``element`` to the nearest paragraph
    that stays, so no range is left with one end."""
    inside: dict[tuple[str, str], Element] = {}
    for node in element.iter():
        if node.tag in _MARKERS:
            inside[(node.tag, node.get(qn("w:id")) or "")] = node
    lonely: list[Element] = []
    for (tag, mark_id), node in inside.items():
        partner = _RANGE_PAIRS.get(tag) or _RANGE_ENDS.get(tag)
        if (partner, mark_id) not in inside:
            lonely.append(node)
    if not lonely:
        return
    following = _neighbour_paragraph(element, forward=True)
    previous = _neighbour_paragraph(element, forward=False)
    for node in lonely:
        node.tail = None
        if following is not None:
            properties = following.find(W_PPR)
            if properties is not None:
                properties.addnext(node)
            else:
                following.insert(0, node)
        elif previous is not None:
            previous.append(node)


def _neighbour_paragraph(element: Element, *, forward: bool) -> Element | None:
    body = element
    while body.getparent() is not None:
        body = body.getparent()
    paragraphs = [p for p in body.iter(W_P)]
    inside = set(map(id, element.iter(W_P)))
    try:
        anchor = next(i for i, p in enumerate(paragraphs) if id(p) in inside)
        last = max(i for i, p in enumerate(paragraphs) if id(p) in inside)
    except (StopIteration, ValueError):
        return None
    candidates = paragraphs[last + 1:] if forward else reversed(paragraphs[:anchor])
    for paragraph in candidates:
        if id(paragraph) not in inside:
            return paragraph
    return None


#: Every revision record's tag: a block holding one moves as a deletion and an insertion.
REVISION_KINDS_ALL = frozenset(_W + name for name in (
    "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange", "tblPrChange",
    "trPrChange", "tcPrChange", "tblGridChange", "numberingChange", "cellIns", "cellDel", "cellMerge"))


def _clear_revisions(paragraph: Element) -> None:
    """Make a copied paragraph's content plain: its insertions unwrapped, its deletions and
    moved-away text removed, its records dropped -- what it shows in the current view."""
    for node in list(paragraph.iter(_W + "del", _W + "moveFrom")):
        if node.getparent() is not None and node.getparent().tag != qn("w:rPr"):
            remove(node)
    for node in list(paragraph.iter(_W + "ins", _W + "moveTo")):
        parent = node.getparent()
        if parent is None:
            continue
        if parent.tag == qn("w:rPr"):
            remove(node)
            continue
        for child in list(node):
            node.addprevious(child)
        remove(node)
    for node in list(paragraph.iter(_W + "rPrChange", _W + "pPrChange")):
        remove(node)
    for node in list(paragraph.iter(_W + "del", _W + "moveFrom")):
        remove(node)


def _make_trackable() -> None:
    """Give every edit a ``track`` keyword (:func:`docx_agent.revisions.mode.trackable`)."""
    from .annotations import Note as _Note
    from .links import Hyperlink as _Hyperlink

    plan = {
        Document: ("set_paragraph_style", "set_character_style", "format_range", "format_paragraph",
                   "format_run", "clear_formatting", "insert_paragraph", "append_paragraph", "delete_block",
                   "move_block", "move_blocks", "replace", "insert_text", "delete_text", "add_to_list", "set_list_level",
                   "restart_numbering", "continue_numbering", "remove_from_list", "set_list_format",
                   "add_hyperlink", "insert_cross_reference", "insert_picture", "insert_row", "delete_row",
                   "insert_column", "delete_column", "merge_cells", "split_cell", "insert_table", "set_table",
                   "set_column_width", "set_row", "set_cell", "float_drawing", "inline_drawing", "set_drawing",
                   "move_drawing", "resize_drawing", "insert_text_box", "insert_shape", "fill_control",
                   "insert_control", "remove_control", "add_bookmark", "rename_bookmark",
                   "remove_bookmark", "coalesce_runs", "_set_text", "insert_markdown", "copy_blocks",
                   # E4: what Word tracks (a section's properties, a break, notes, fields, a
                   # table of contents, text in a story); the rest take track= and write no
                   # revision, as Word writes none for them.
                   "insert_section_break", "remove_section_break", "set_section", "add_header", "add_footer",
                   "link_to_previous", "unlink_from_previous", "remove_header", "remove_footer",
                   "insert_footnote", "insert_endnote", "edit_note", "delete_note", "move_note",
                   "insert_field", "insert_page_number", "insert_date", "insert_sequence", "insert_caption",
                   "insert_hyperlink", "insert_toc"),
        Paragraph: ("set_style", "format", "clear_direct_formatting", "add_to_list", "indent_list",
                    "outdent_list", "restart_numbering", "continue_numbering", "remove_from_list", "set_text",
                    "insert_after", "insert_before", "delete", "move", "coalesce_runs"),
        Run: ("format",),
        TextRange: ("replace", "delete", "insert_before", "insert_after", "format", "set_style",
                    "clear_formatting", "add_hyperlink", "add_bookmark", "insert_picture"),
        Table: ("insert_row", "delete_row", "insert_column", "delete_column", "merge", "split", "set",
                "set_column_width", "set_row", "set_cell"),
        _Hyperlink: ("set_target", "remove"),
        Picture: ("replace", "resize", "delete"),
        Section: ("set", "remove_break", "add_header", "add_footer", "link_to_previous", "unlink_from_previous"),
        _Note: ("edit", "delete", "move"),
    }
    for cls, names in plan.items():
        for name in names:
            setattr(cls, name, trackable(getattr(cls, name)))


_make_trackable()
