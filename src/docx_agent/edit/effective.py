"""Effective formatting: what Word applies, read through docx2svg's resolver.

docx2svg's :mod:`docx2svg.resolve` has the style cascade measured against Word --
``w:docDefaults``, table style and its conditional formats, numbering, paragraph style,
character style, direct formatting; toggle properties as Word combines them, not as
ECMA-376 says; theme fonts -- and says which level every value came from.  This module
asks it, over the document's current bytes (parsed once per document state), rather than
keep a second cascade that would drift from the measured one.

``paragraph.effective`` and ``run.effective`` return :class:`EffectiveParagraph` and
:class:`EffectiveRun`: plain values (points, names), and ``explain()``, docx2svg's account
of where each came from.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from docx2svg import model as _model
from docx2svg.parse.document import parse_package
from docx2svg.parse.properties import read_paragraph_properties, read_run_properties
from docx2svg.paths import path_part, resolve_path
from docx2svg.resolve import character_format, resolve_paragraph, resolve_run

from ..oxml.xml import Element

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_ALIGNMENTS = {"left": "left", "start": "left", "center": "center", "right": "right", "end": "right",
               "both": "justify", "distribute": "distribute"}


@dataclass(frozen=True)
class EffectiveRun:
    """A run's formatting as Word applies it."""

    bold: bool
    italic: bool
    underline: str | None
    strike: bool
    double_strike: bool
    caps: bool
    small_caps: bool
    #: Points.
    size: float
    #: The face the run's first character is drawn in (its script's slot; theme fonts
    #: resolved), or ``None`` where Word falls back to its own.
    font: str | None
    #: ``RRGGBB`` or ``auto``.
    color: str
    highlight: str | None
    vertical_align: str | None
    hidden: bool
    #: The character style applied, as declared (``None`` for the default).
    style: str | None
    _resolved: object = field(default=None, repr=False, compare=False)

    def explain(self) -> str:
        """Where every value came from, one line per property (docx2svg's account)."""
        return self._resolved.explain()


@dataclass(frozen=True)
class EffectiveParagraph:
    """A paragraph's formatting as Word applies it.  Lengths in points."""

    style: str | None
    alignment: str
    indent_left: float
    indent_right: float
    first_line: float
    hanging: float
    space_before: float
    space_after: float
    #: A multiple (``line_rule`` ``auto``) or points (``exact``, ``atLeast``).
    line_spacing: float | None
    line_rule: str
    keep_with_next: bool
    keep_together: bool
    page_break_before: bool
    widow_control: bool
    outline_level: int | None
    #: The list the paragraph is in: ``(numId, level)``, or ``None``.
    numbering: tuple[int, int] | None
    _resolved: object = field(default=None, repr=False, compare=False)

    def explain(self) -> str:
        return self._resolved.explain()


class _Model:
    """docx2svg's parse of one document state, with its paragraphs by our elements."""

    def __init__(self, document: "Document") -> None:
        self.document = document
        data = document.to_bytes()
        self.model = parse_package(data)
        self.paragraphs: dict[int, _model.Paragraph] = {}
        self._held: list[Element] = []
        package = document.package
        parts = _Parts(document)
        for block in _paragraphs(self.model.body):
            self._map(block, package.document_part())
        for relationship, blocks in self.model.stories.items():
            part = self.model.story_parts.get(relationship)
            for block in _paragraphs(blocks):
                self._map(block, part)
        for notes in (self.model.footnotes, self.model.endnotes):
            for blocks in notes.values():
                for block in _paragraphs(blocks):
                    self._map(block, path_part(parts, block.path))

    def _map(self, paragraph: _model.Paragraph, part: str | None) -> None:
        if part is None or not paragraph.path:
            return
        element = resolve_path(self.document.package.tree(part), paragraph.path)
        if element is not None and element.tag == _W + "p":
            self._held.append(element)  # one proxy per element while the map lives: stable id()s
            self.paragraphs.setdefault(id(element), paragraph)

    def paragraph(self, element: Element) -> _model.Paragraph:
        found = self.paragraphs.get(id(element))
        if found is not None and not found.joins:
            return found
        # Not drawn on its own (joined across a deleted mark, in a text box): read it here.
        properties = element.find(_W + "pPr")
        style = properties.find(_W + "pStyle") if properties is not None else None
        mark = properties.find(_W + "rPr") if properties is not None else None
        declared = read_paragraph_properties(properties)
        return _model.Paragraph(
            properties=_model.ParagraphProperties(
                style_id=style.get(_W + "val") if style is not None else None, declared=declared,
                run_properties=_model.RunProperties(declared=read_run_properties(mark))),
            table_style_id=_table_style(element),
            table_conditions=found.table_conditions if found is not None else ())


class _Parts:
    def __init__(self, document: "Document") -> None:
        self._package = document.package
        self.main_document_part = document.package.document_part()

    def related(self, part: str, relationship_type: str) -> list[str]:
        return self._package.related_parts_of_type(part, relationship_type)


def _paragraphs(blocks):
    for block in blocks or ():
        if isinstance(block, _model.Paragraph):
            yield block
        elif isinstance(block, _model.Table):
            for row in block.rows:
                for cell in row:
                    yield from _paragraphs(cell)


def _table_style(element: Element) -> str | None:
    node = element.getparent()
    while node is not None and node.tag != _W + "tbl":
        node = node.getparent()
    if node is None:
        return None
    style = node.find(f"{_W}tblPr/{_W}tblStyle")
    return style.get(_W + "val") if style is not None else ""


def model(document: "Document") -> _Model:
    """docx2svg's parse of the document as it is now, cached until the next change."""
    cached = document._effective
    if cached is None or cached[0] != document._state:
        cached = (document._state, _Model(document))
        document._effective = cached
    return cached[1]


def run(document: "Document", paragraph: Element, run_element: Element | None,
        first: str | None = None) -> EffectiveRun:
    """``run_element``'s formatting in ``paragraph``; ``first`` is its first character (read
    from the paragraph when not given), which decides the script and so the face.  With
    ``run_element`` ``None``: what a run with no properties of its own would get there --
    the paragraph's baseline, which the Markdown reader measures emphasis against."""
    state = model(document)
    owner = state.paragraph(paragraph)
    properties = run_element.find(_W + "rPr") if run_element is not None else None
    style = properties.find(_W + "rStyle") if properties is not None else None
    style_id = style.get(_W + "val") if style is not None else None
    model_run = _model.Run("", _model.RunProperties(style_id=style_id, declared=read_run_properties(properties)))
    resolved = resolve_run(state.model, owner, model_run)
    if first is None:
        from .text import run_text

        first = (run_text(paragraph, run_element) if run_element is not None else "") or "a"
    character = character_format(resolved, first[0], state.model)
    underline = resolved.get("u")
    return EffectiveRun(
        bold=character.bold, italic=character.italic,
        underline=None if underline in (None, "none") else underline,
        strike=bool(resolved.get("strike")), double_strike=bool(resolved.get("dstrike")),
        caps=bool(resolved.get("caps")), small_caps=bool(resolved.get("smallCaps")),
        size=character.half_points / 2, font=character.face,
        color=resolved.get("color") or "auto", highlight=resolved.get("highlight"),
        vertical_align=resolved.get("vertAlign"), hidden=bool(resolved.get("vanish")),
        style=style_id, _resolved=resolved)


def paragraph(document: "Document", element: Element) -> EffectiveParagraph:
    state = model(document)
    owner = state.paragraph(element)
    resolved = resolve_paragraph(state.model, owner)

    def points(key: str) -> float:
        value = resolved.get(key)
        return (value or 0) / 20

    rule = resolved.get("spacing.lineRule") or "auto"
    line = resolved.get("spacing.line")
    line_spacing = None if line is None else (line / 240 if rule == "auto" else line / 20)
    num_id = resolved.get("numPr.numId")
    numbering = (num_id, resolved.get("numPr.ilvl") or 0) if num_id else None
    style = owner.properties.style_id if owner.properties else None
    return EffectiveParagraph(
        style=style, alignment=_ALIGNMENTS.get(resolved.get("jc") or "left", resolved.get("jc")),
        indent_left=points("ind.left"), indent_right=points("ind.right"),
        first_line=points("ind.firstLine"), hanging=points("ind.hanging"),
        space_before=points("spacing.before"), space_after=points("spacing.after"),
        line_spacing=line_spacing, line_rule=rule,
        keep_with_next=bool(resolved.get("keepNext")), keep_together=bool(resolved.get("keepLines")),
        page_break_before=bool(resolved.get("pageBreakBefore")),
        widow_control=bool(resolved.get("widowControl", True)),
        outline_level=resolved.get("outlineLvl"), numbering=numbering, _resolved=resolved)
