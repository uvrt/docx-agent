"""Charts and SmartArt: the data of a chart and the nodes of a diagram, wherever Word holds one.

The editing is ooxml-edit's :mod:`ooxml_edit.charts` -- a chart part, its embedded workbook
and a diagram's data model are the same DrawingML in a Word document as in a deck, so it
lives there once, and pptx-agent hosts the same code.  What is left here is what makes it
Word's: :func:`graphic_host` tells it where a chart lives -- the drawing (``wp:inline`` or
``wp:anchor``, or a group's ``wpg:graphicFrame``), the part whose relationships name the
chart (``document.xml``, a header, a footer, the notes, with a text box's content in its
anchor's part), docx-agent's one undo step -- and what to call things in messages ("Word's
Edit Data", "the document").

**Addressing.**  A chart or diagram is a drawing, ``d:<docPr id>`` (ROADMAP.md,
"Addressing and stable ids").  A chart inside a group has no ``docPr`` of its own: it is
``d:<group's docPr id>/<its wpg:cNvPr id>`` (a nested group's members one step further
down), which Word keeps through a save as it keeps ``docPr`` ids (measured,
``tools/charts_probe.py``); a repeated ``cNvPr`` id within one group gets ``#n``, as a
repeated ``docPr`` id does.

**Caches and the workbook.**  Word draws a chart from its caches and never refreshes them
from the embedded workbook, on opening or saving, whatever ``c:autoUpdate`` says (Word
writes it back as ``0``); a save keeps the caches, the formulas and the workbook's bytes
(measured, ``tools/charts_probe.py``).  So every edit writes both, as ooxml-edit does.

**Titles.**  A new chart or axis title is written as Office's chart style writes one
(ooxml-edit's ``WORD_LOOK``, measured on Word's own new charts; :func:`title_text` before
T4): Word draws it 14 pt (an axis's 10 pt, horizontal or vertical), not bold, in the theme's
minor face at 65% of the text colour, as docx2svg does; ooxml-edit's default, an empty
``a:defRPr``, Word draws bold at 18 pt (measured).

**Tracking.**  Word has no revision for a chart's data or a diagram's text (measured: with
tracking on, its own review counts none in a document carrying either edit, and its PDF
shows them), so with tracking on an edit is applied untracked and
:class:`UntrackedChartEdit` says so, as docx-agent does for every edit Word does not track.

**The cached drawing of a diagram.**  Word lays a diagram out again from its data model on
opening and rewrites the drawing on saving, as PowerPoint does (measured: a stale drawing,
a drawing changed alone, a missing one, a node added or removed with the drawing dropped
or kept -- the PDF always shows the data model, and the saved drawing follows it).
docx2svg draws only the cache.  So a drawing an edit cannot keep exactly in step is
dropped (:data:`DIAGRAM_POLICY`), which Word does not notice and docx2svg shows as a
placeholder until Word saves the document again, rather than kept, which docx2svg would
show as the old diagram; :class:`DiagramDrawingDropped` says so.
"""

from __future__ import annotations

import warnings
from contextlib import contextmanager
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Iterator

from ooxml_edit.charts import chart as _chart
from ooxml_edit.charts import diagram as _diagram
from ooxml_edit.charts import model as _model
from ooxml_edit.charts.create import WORD_LOOK
from ooxml_edit.charts.host import GraphicHost, chart_part, diagram_parts

from ..oxml.xml import Element, make
from .errors import EditError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_WPG = "{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}"
_WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
_DGM = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
_PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"

#: What a cached diagram drawing an edit cannot keep in step becomes (``drop``, ``keep`` or
#: ``refuse``; ``ooxml_edit.charts.diagram``).  Chosen on what Word does (module docstring).
DIAGRAM_POLICY = "drop"
#: The application and the document, in ooxml-edit's messages.
APPLICATION = "Word"
DOCUMENT = "document"
#: The language of new chart text when the document's defaults name none.
FALLBACK_LANG = "en-US"

ChartDataError = _chart.ChartDataError
ChartDataWarning = _chart.ChartDataWarning
ChartModelError = _model.ChartModelError
DiagramDrawingError = _diagram.DiagramDrawingError


class UntrackedChartEdit(UserWarning):
    """A chart or diagram edit made while tracking: Word has no revision for it, so it was
    applied untracked."""


class DiagramDrawingDropped(UserWarning):
    """A diagram's cached drawing was dropped (or, by choice, kept stale) by an edit: Word
    lays the diagram out again on opening; docx2svg draws the cache it finds until Word saves
    the document again."""


class ChartEditError(EditError):
    """A chart or diagram model refused (ooxml-edit's ``ChartModelError``, word for word)."""


# -- where a chart is ------------------------------------------------------------------------


@dataclass(frozen=True)
class Member:
    """A drawing inside a group, which has no ``docPr`` of its own: ``d:<group>/<cNvPr id>``."""

    document: "Document"
    id: str

    def _locate(self) -> tuple[str, Element]:
        return locate(self.document, self.id)

    @property
    def kind(self) -> str:
        """The group member's kind: ``"picture"``, ``"chart"``, ``"diagram"``, ``"shape"``..."""
        return member_kind(self._locate()[1])

    @property
    def name(self) -> str | None:
        """The member's name, or ``None``."""
        properties = _member_properties(self._locate()[1])
        return properties.get("name") if properties is not None else None

    @property
    def chart(self) -> "Chart":
        """The chart this member holds: :meth:`Document.chart`."""
        return self.document.chart(self.id)

    @property
    def diagram(self) -> "Diagram":
        """The SmartArt diagram this member holds: :meth:`Document.diagram`."""
        return self.document.diagram(self.id)

    def __repr__(self) -> str:
        return f"<Member {self.id} {self.kind}>"


_MEMBERS = (_WPG + "graphicFrame", _WPG + "grpSp", _WPS + "wsp", _WPG + "wsp")


def _member_properties(node: Element) -> Element | None:
    for tag in (_WPG + "cNvPr", _WPS + "cNvPr"):
        found = node.find(tag)
        if found is not None:
            return found
    return None


def member_kind(node: Element) -> str:
    """``chart``, ``diagram``, ``group``, ``text-box``, ``shape``, ``picture`` or ``other``."""
    if node.tag == _WPG + "grpSp":
        return "group"
    data = node.find(f"{_A}graphic/{_A}graphicData")
    if data is not None:
        if data.find(_C + "chart") is not None:
            return "chart"
        if data.find(_DGM + "relIds") is not None:
            return "diagram"
        return "other"
    if node.tag in (_WPS + "wsp", _WPG + "wsp"):
        return "text-box" if node.find(_WPS + "txbx") is not None else "shape"
    if node.tag.endswith("}pic"):
        return "picture"
    return "other"


def group_members(group: Element, prefix: str) -> list[tuple[str, Element]]:
    """``(id, element)`` for every member of a group drawing (``wpg:wgp``, or a nested
    ``wpg:grpSp``), depth first: ``<prefix>/<cNvPr id>``."""
    holder = group.find(f"{_A}graphic/{_A}graphicData/{_WPG}wgp") if group.tag != _WPG + "grpSp" else group
    if holder is None:
        return []
    out: list[tuple[str, Element]] = []
    seen: dict[str, int] = {}
    for child in holder:
        if child.tag not in _MEMBERS and not child.tag.endswith("}pic"):
            continue
        properties = _member_properties(child)
        if properties is None:  # a picture's is one level down
            properties = child.find(f"{_PIC}nvPicPr/{_PIC}cNvPr")
        raw = properties.get("id") if properties is not None else None
        if raw is None:
            continue
        count = seen.get(raw, 0)
        seen[raw] = count + 1
        identifier = f"{prefix}/{raw}" + (f"#{count}" if count else "")
        out.append((identifier, child))
        if child.tag == _WPG + "grpSp":
            out += group_members(child, identifier)
    return out


def locate(document: "Document", identifier: str) -> tuple[str, Element]:
    """``(part, frame)`` for a drawing id: the ``wp:inline``/``wp:anchor`` of ``d:<id>``, or a
    group member's own element for ``d:<group>/<member>``."""
    base, slash, _ = identifier.partition("/")
    for part, found, frame in document._all_drawings():
        if found == base:
            if not slash:
                return part, frame
            for member, element in group_members(frame, base):
                if member == identifier:
                    return part, element
            break
    raise KeyError(f"no drawing {identifier!r}")


def default_language(document: "Document") -> str:
    """The document's default language (``w:docDefaults``' ``w:lang``), for new chart text."""
    styles = document.package.styles_part()
    root = document.package.tree(styles) if styles else None
    node = root.find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr/{_W}lang") if root is not None else None
    value = node.get(_W + "val") if node is not None else None
    return value or FALLBACK_LANG


def title_text(vertical: bool, *, lang: str | None = None) -> Element:
    """The ``c:tx`` of a new chart or axis title in the form Office's default chart style
    gives a title -- 14 pt (an axis's 10 pt), not bold, the text colour at 65%, the theme's
    minor face -- which Word draws so and saves back unchanged (measured,
    ``tools/charts_probe.py``, ``titles``).  Word's dictionary cannot add a title itself, so
    what Word writes for one it adds was not observed."""
    tx = make("c:tx")
    rich = make("c:rich")
    body = make("a:bodyPr", rot="-5400000" if vertical else "0", spcFirstLastPara="1", vertOverflow="ellipsis",
                vert="horz", wrap="square", anchor="ctr", anchorCtr="1")
    rich.append(body)
    rich.append(make("a:lstStyle"))
    paragraph = make("a:p")
    properties = make("a:pPr")
    defaults = make("a:defRPr", sz="1000" if vertical else "1400", b="0", i="0", u="none", strike="noStrike",
                    kern="1200", spc="0", baseline="0")
    fill = make("a:solidFill")
    colour = make("a:schemeClr", val="tx1")
    colour.append(make("a:lumMod", val="65000"))
    colour.append(make("a:lumOff", val="35000"))
    fill.append(colour)
    defaults.append(fill)
    for tag, face in (("a:latin", "+mn-lt"), ("a:ea", "+mn-ea"), ("a:cs", "+mn-cs")):
        defaults.append(make(tag, typeface=face))
    properties.append(defaults)
    paragraph.append(properties)
    # The language goes on the paragraph's end: the text written into the title takes its
    # run's properties from there (``replace_body_text``); an empty run's would be dropped.
    paragraph.append(make("a:endParaRPr") if lang is None else make("a:endParaRPr", lang=lang))
    rich.append(paragraph)
    tx.append(rich)
    return tx


def graphic_host(document: "Document", identifier: str) -> GraphicHost:
    """A Word drawing as ooxml-edit's charts and diagrams see it."""
    part, frame = locate(document, identifier)
    lang = default_language(document)

    @contextmanager
    def edit() -> Iterator[None]:
        tracking = document._active_tracking()
        with document._edit():
            yield
        if tracking is not None:
            warnings.warn(f"{identifier}: Word does not track a chart's data or a diagram's text: "
                          "the edit is applied untracked", UntrackedChartEdit, stacklevel=5)

    return GraphicHost(package=document.package, part=part, frame=frame, edit=edit, address=identifier,
                       application=APPLICATION, document=DOCUMENT, lang=lang,
                       title_template=lambda vertical: title_text(vertical, lang=lang), look=WORD_LOOK)


def has_chart(document: "Document", identifier: str) -> bool:
    return _kind(document, identifier) == "chart"


def has_diagram(document: "Document", identifier: str) -> bool:
    return _kind(document, identifier) == "diagram"


def _kind(document: "Document", identifier: str) -> str:
    from .annotations import drawing_kind

    _, frame = locate(document, identifier)
    return drawing_kind(frame) if frame.tag in (_WP + "inline", _WP + "anchor") else member_kind(frame)


# -- the views ---------------------------------------------------------------------------------


class Chart(_chart.Chart):
    """The chart a drawing holds -- re-resolved from the document on every call, so it
    survives undo.  Every edit is one undo step: the chart part, its caches and its embedded
    workbook together (:mod:`ooxml_edit.charts.chart`).

    Series are numbered in document order (``chart.series[0]``); categories by position.
    """

    def __init__(self, document: "Document", identifier: str) -> None:
        super().__init__(lambda: graphic_host(document, identifier))
        self.document = document
        self.id = identifier

    @property
    def model(self) -> dict[str, Any]:
        """The chart's JSON (``ooxml_edit.charts.model``): ``types``, ``title``,
        ``axis_titles``, ``legend``, ``format``, ``categories``, ``series``."""
        return self.data

    def apply(self, model: Any) -> "Chart":
        """Bring the chart to ``model`` (its JSON, edited) in one undo step; a model that
        cannot be applied is refused with :class:`ChartEditError`, nothing changed."""
        try:
            wanted = _model.canonical_chart(model, self.id)
            with self.document._edit():
                _model.apply_chart_model(self, wanted, self.id)
        except _model.ChartModelError as error:
            raise ChartEditError(str(error)) from None
        return self

    def __repr__(self) -> str:
        return f"<Chart {self.id} {'+'.join(self.chart_types) or '?'}>"


class Diagram(_diagram.Diagram):
    """The SmartArt a drawing holds -- re-resolved on every call.  Nodes are listed depth
    first and addressed by their ``modelId``."""

    def __init__(self, document: "Document", identifier: str, on_inexact_drawing: str | None = None) -> None:
        def notify(message: str) -> None:
            if "dropped" in message:
                message += (" (docx2svg draws a placeholder until Word saves the document again -- or, where "
                            "the part keeps exactly one other diagram's drawing, that drawing)")
            warnings.warn(message, DiagramDrawingDropped, stacklevel=4)

        super().__init__(lambda: graphic_host(document, identifier),
                         on_inexact_drawing=on_inexact_drawing or DIAGRAM_POLICY, notify=notify)
        self.document = document
        self.id = identifier

    def apply(self, model: Any) -> "Diagram":
        """Bring the diagram's nodes to ``model`` (its JSON, edited) in one undo step;
        refused with :class:`ChartEditError`, nothing changed."""
        try:
            wanted = _model.canonical_diagram(model, self.id)
            with self.document._edit():
                _model.apply_diagram_model(self, wanted, self.id)
        except _model.ChartModelError as error:
            raise ChartEditError(str(error)) from None
        return self

    def __repr__(self) -> str:
        return f"<Diagram {self.id} {len(self.nodes)} nodes>"


class ChartOps:
    """Charts and SmartArt, on :class:`docx_agent.Document`."""

    def _graphics(self: "Document") -> list[tuple[str, str, Element, str]]:
        """``(part, id, frame, kind)`` for every drawing and group member, in order."""
        from .annotations import drawing_kind

        out = []
        for part, identifier, frame in self._all_drawings():
            kind = drawing_kind(frame)
            out.append((part, identifier, frame, kind))
            if kind == "group":
                out += [(part, member, element, member_kind(element))
                        for member, element in group_members(frame, identifier)]
        return out

    def chart(self: "Document", identifier: str) -> Chart:
        """The chart drawing ``identifier`` (``d:<id>``, or ``d:<group>/<member>``) holds:
        its series, categories, values, titles and legend, edited in the chart's caches and
        its embedded workbook together."""
        kind = _kind(self, identifier)
        if kind != "chart":
            raise EditError(f"{identifier} is a {kind}, not a chart")
        return Chart(self, identifier)

    def diagram(self: "Document", identifier: str, *, on_inexact_drawing: str | None = None) -> Diagram:
        """The SmartArt drawing ``identifier`` holds: its nodes' text, nodes added and
        removed.  ``on_inexact_drawing`` overrides :data:`DIAGRAM_POLICY` for this view."""
        kind = _kind(self, identifier)
        if kind != "diagram":
            raise EditError(f"{identifier} is a {kind}, not a SmartArt diagram")
        return Diagram(self, identifier, on_inexact_drawing)

    def insert_chart(self: "Document", at: str, chart_type: str, categories, series, *,
                     width: float | None = None, height: float | None = None,
                     title: str | None = None, axis_titles: dict | None = None,
                     legend: str | None = "bottom", number_format: str | None = None,
                     name: str | None = None):
        """A new chart from data, as an inline drawing, as Word inserts one: the chart part
        and an embedded workbook holding the same numbers, so Edit Data opens them
        (ooxml-edit's :func:`~ooxml_edit.charts.add_chart`).

        ``at`` is a block id (``p:...``, ``t:...``) -- the chart goes in a new paragraph
        after it, which continues that paragraph's formatting (a heading's becomes
        ``Normal``) -- or a position inside a paragraph (``p:...@12``), where it goes inline
        as a picture would.  ``chart_type`` is ``column``, ``stacked_column``, ``bar``,
        ``stacked_bar``, ``line``, ``pie`` or ``scatter``; ``categories`` the labels (a
        scatter chart's x values); ``series`` ``[{"name", "values"}]``, one value per
        category.  ``width`` and ``height`` are points: Word's 432 x 252 by default, one
        given keeps that ratio.  It looks as Word's new chart of that type looks in the
        document's theme (measured: 14 pt title, 9 pt labels, legend at the bottom, a
        background fill with a light border).  Returns the drawing's ``d:<id>``; one undo
        step.  Data that cannot be charted raises ``ChartDataError`` before anything
        changes."""
        from ooxml_edit.charts import WORD_LOOK, add_chart
        from ooxml_edit.charts.create import chart_data

        from . import ids as _ids
        from . import inline as _inline
        from .document import EditResult

        chart_data(chart_type, categories, series)  # refused before anything changes
        size = _chart_size(width, height)
        with self._edit():
            renames: dict = {}
            created: list[str] = []
            if "@" in at:
                where = self.range(at)
            else:
                anchor = self.paragraph(at) if at.startswith("p:") else None
                style = None
                if anchor is not None and (anchor.style_name or "").lower().startswith(
                        ("heading", "title", "subtitle")):
                    style = "Normal"
                inserted = self.insert_paragraph(after=at, style=style)
                renames.update(inserted.renamed)
                created += inserted.created
                where = self.range(f"{inserted.id}@0")
            part, entries = where._entries()
            entry = entries[0]
            renames.update(self._prepare(part, [entry]))
            made = add_chart(self.package, part, chart_type, list(categories),
                             [dict(s) if isinstance(s, dict) else s for s in series],
                             title=title, axis_titles=axis_titles, legend=legend,
                             number_format=number_format, look=WORD_LOOK,
                             lang=default_language(self))
            number = self._next_doc_pr()
            label = name or f"Chart {number}"
            anchor_id, edit_id = _ids.generate(part + "\0chart", self._used(), 2)
            drawing = _chart_drawing(made.graphic(), number, label, size, anchor_id, edit_id)
            run = _inline.new_run(entry.element, where.start, [drawing])
            properties = run.find(_W + "rPr")
            if properties is None:
                properties = make("w:rPr")
                run.insert(0, properties)
            if properties.find(_W + "noProof") is None:
                properties.append(make("w:noProof"))
            self._track_inserted(part, [run])
            _ids.ensure_w14(self.package.tree(part), drawings=True)
            self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
        identifier = f"d:{number}"
        return EditResult(identifier, created=created + [identifier], renamed=renames)

    def charts(self: "Document") -> list[Chart]:
        """Every chart, in the body, headers, footers, notes, text boxes and groups."""
        return [Chart(self, identifier) for _, identifier, _, kind in self._graphics() if kind == "chart"]

    def diagrams(self: "Document") -> list[Diagram]:
        """Every SmartArt diagram, wherever it stands."""
        return [Diagram(self, identifier) for _, identifier, _, kind in self._graphics() if kind == "diagram"]


#: Word's new chart, measured: 5486400 x 3200400 EMU.
CHART_WIDTH = 432.0
CHART_HEIGHT = 252.0


def _chart_size(width: float | None, height: float | None) -> tuple[float, float]:
    for value in (width, height):
        if value is not None and not 1 <= float(value) <= 1584:
            raise EditError("a chart's width and height are 1-1584 pt")
    if width is None and height is None:
        return CHART_WIDTH, CHART_HEIGHT
    if height is None:
        return float(width), float(width) * CHART_HEIGHT / CHART_WIDTH
    if width is None:
        return float(height) * CHART_WIDTH / CHART_HEIGHT, float(height)
    return float(width), float(height)


def _chart_drawing(graphic: Element, number: int, name: str, size: tuple[float, float],
                   anchor_id: str, edit_id: str) -> Element:
    """``w:drawing`` with an inline chart, as Word writes one (measured: the effect extent
    one border's width right and below, an empty ``wp:cNvGraphicFramePr``), with the
    ``wp14`` ids Word needs to keep the document's paraIds, as for a picture."""
    from xml.sax.saxutils import quoteattr

    from lxml import etree

    from . import ids as _ids

    cx, cy = int(round(size[0] * 12700)), int(round(size[1] * 12700))
    xml = (f'<w:drawing xmlns:w="{_W[1:-1]}"><wp:inline xmlns:wp="{_WP[1:-1]}" '
           f'xmlns:wp14="{_ids.WP14}" distT="0" distB="0" distL="0" distR="0" '
           f'wp14:anchorId="{anchor_id}" wp14:editId="{edit_id}">'
           f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="12700" b="12700"/>'
           f'<wp:docPr id="{number}" name={quoteattr(name)}/><wp:cNvGraphicFramePr/>'
           f'</wp:inline></w:drawing>')
    drawing = etree.fromstring(xml)
    drawing[0].append(graphic)
    return drawing


# -- reading, for to_markdown and state ----------------------------------------------------------


@contextmanager
def _no_edit() -> Iterator[None]:
    raise TypeError("a read-only host")  # pragma: no cover
    yield


def chart_json(document: "Document", part: str, frame: Element) -> dict | None:
    """A chart's JSON, read without a view; ``None`` when its part is missing or unreadable."""
    host = GraphicHost(package=document.package, part=part, frame=frame, edit=_no_edit,
                       address="", application=APPLICATION)
    try:
        target = chart_part(host)
        root = document.package.tree(target) if target else None
        return _chart.chart_model(root) if root is not None else None
    except Exception:  # a chart docx-agent cannot read is still a drawing
        return None


def diagram_json(document: "Document", part: str, frame: Element) -> dict | None:
    host = GraphicHost(package=document.package, part=part, frame=frame, edit=_no_edit,
                       address="", application=APPLICATION)
    try:
        parts = diagram_parts(host)
        data = document.package.tree(parts["data"]) if parts.get("data") else None
        layout = document.package.tree(parts["layout"]) if parts.get("layout") else None
        if data is None:
            return None
        return _diagram.diagram_model(data, layout.get("uniqueId") if layout is not None else None)
    except Exception:
        return None


def chart_summary(model: dict) -> str:
    """One line: ``column; title: Sales; series: North, South; categories: Q1, Q2`` --
    pptx-agent's outline line, inside the drawing's comment."""
    types = []
    for name in model.get("types") or []:
        if name not in types:
            types.append(name)
    words = [" + ".join(types) or "chart"]
    if model.get("title"):
        words.append("title: " + _flat(model["title"]))
    names = [_flat(s.get("name") or "") for s in model.get("series") or []]
    if names:
        words.append("series: " + ", ".join(names))
    categories = [_flat("" if c is None else str(c)) for c in model.get("categories") or []]
    if categories:
        words.append("categories: " + ", ".join(categories))
    return "; ".join(words)


def diagram_summary(model: dict) -> str:
    """The node text, depth first, a node's children in brackets after it:
    ``Goals [Faster edits, Fewer prompts], Risks [Stale caches]``."""
    out = ""
    depth = 0
    first = True
    for node in model.get("nodes") or []:
        level = max(0, min(int(node.get("lvl") or 0), depth + 1)) if not first else 0
        if level > depth:
            out += " ["
        else:
            out += "]" * (depth - level) + ("" if first else ", ")
        out += _flat(node.get("t") or "")
        depth, first = level, False
    return out + "]" * depth


def _flat(text: str) -> str:
    return " ".join(str(text).split())


__all__ = ["APPLICATION", "Chart", "ChartOps", "ChartDataError", "ChartDataWarning", "ChartEditError", "ChartModelError",
           "DIAGRAM_POLICY", "Diagram", "DiagramDrawingDropped", "DiagramDrawingError", "Member",
           "UntrackedChartEdit", "chart_json", "chart_summary", "default_language", "diagram_json",
           "diagram_summary", "graphic_host", "group_members", "locate", "member_kind", "title_text"]
