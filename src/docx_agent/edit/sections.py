"""Sections and page setup, and the headers and footers each section shows (E4).

A **section** is named by the paragraph that ends it -- ``s:<paraId>``, its ``w:sectPr``
in that paragraph's properties -- and the body's last is ``s:body`` (ROADMAP.md,
"Addressing").  What Word writes for each edit was measured (``tools/e4_probe.py``,
``tests/observations/e4-word.json``) and is written so:

* **A section break** is a paragraph of its own holding the ``w:sectPr``: Word's form for
  a break put at a paragraph's start (``breaks``), the new paragraph taking the properties
  of the one it is put before (``fields``: a break before a heading is a heading).  Its
  ``w:sectPr`` is a copy of the section it splits, header and footer references
  included, and the section after the break -- the one split -- keeps its properties, gets
  the break's kind as its ``w:type`` (``nextPage`` written as none, as Word writes it) and
  gives its references to the new section before it, inheriting them from there (Word
  moves them: ``tracked``).  Untracked and tracked alike: tracked, the break's mark is an
  insertion and the split section's kind a ``w:sectPrChange``; its references stay, the
  two sections sharing the parts until the break is accepted or rejected (references are
  not revisions).
* **Removing a break** deletes the paragraph mark that holds it, as Word does: the text
  before it joins the paragraph after (an empty break paragraph simply goes) and takes the
  next section's properties and headers (``removal``).  A header or footer the removed
  section stated and the next did not is handed on to the next, which showed it already
  (it inherited it); one the next overrides goes, and its part with it.  Tracked, the mark
  is a deletion (``tracked-removal``).
* **Page setup** writes ``w:pgSz`` (``w:orient`` landscape, the sides swapped), ``w:pgMar``,
  ``w:cols`` (``w:num``, ``w:sep``, ``w:space``), ``w:vAlign``, ``w:lnNumType``,
  ``w:titlePg``, ``w:pgNumType`` (``w:fmt``, ``w:start``) and the section's notes
  (``w:footnotePr``, ``w:endnotePr``) in the schema's order (``setup``).  Tracked, a
  ``w:sectPrChange`` holds the old value of each kind that changed, as Word records only
  what changed; a kind that was absent is recorded in its "off" form (``w:titlePg
  w:val="0"``, an empty ``w:pgNumType``...) so that rejecting removes it.
* **Headers and footers** are parts of their own (``word/header<n>.xml``) whose one
  paragraph is in the Header (Footer) style, referenced from the section in Word's order
  (even, default header; even, default footer; first header, first footer: ``headers``).
  A section that states no reference of a kind shows the last one stated before it
  (``link to previous``).  Unlinking copies that story into a part of its own, as Word does
  (``headers2``: an unlinked footer starts as the one before); linking again removes the
  reference and, unused, the part.  A first-page story needs ``w:titlePg``, an even one
  the document-wide ``w:evenAndOddHeaders``: adding one sets what it needs and says so.
  Creating, linking and unlinking are not revisions (Word tracks none of them); text
  written into a story is.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.package import REL_FOOTER, REL_HEADER
from ..oxml.xml import Element, append_in_order, make, remove
from . import ids as _ids
from . import text as _text
from .errors import EditError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult, Story

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
W_P = _W + "p"
W_PPR = _W + "pPr"
W_SECTPR = _W + "sectPr"
W_BODY = _W + "body"
_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"

#: ``w:type``'s values: how a section starts.  ``nextPage`` is the default, written as none.
STARTS = ("nextPage", "continuous", "evenPage", "oddPage", "nextColumn")
#: The kinds of header and footer a section may state.
KINDS = ("default", "first", "even")
#: Word's order of the references in a ``w:sectPr`` (``headers``).
_REFERENCE_ORDER = {("header", "even"): 0, ("header", "default"): 1, ("footer", "even"): 2,
                    ("footer", "default"): 3, ("header", "first"): 4, ("footer", "first"): 5}
_REFERENCES = (_W + "headerReference", _W + "footerReference")
VERTICAL = ("top", "center", "both", "bottom")
RESTARTS = ("continuous", "eachSect", "eachPage")
LINE_RESTARTS = ("newPage", "newSection", "continuous")

#: How each kind that was absent before a tracked change is recorded in its
#: ``w:sectPrChange``: the element in its "off" form, which rejecting removes again.
_ABSENT = {
    "titlePg": lambda: make("w:titlePg", **{"w:val": "0"}),
    "vAlign": lambda: make("w:vAlign", **{"w:val": "top"}),
    "type": lambda: make("w:type", **{"w:val": "nextPage"}),
    "pgNumType": lambda: make("w:pgNumType", **{"w:fmt": "decimal"}),
    "lnNumType": lambda: make("w:lnNumType", **{"w:countBy": "0"}),
    "footnotePr": lambda: make("w:footnotePr"),
    "endnotePr": lambda: make("w:endnotePr"),
    "cols": lambda: make("w:cols", **{"w:num": "1"}),
}
#: Attributes an element may gain whose absence means a value Word can be told: a tracked
#: change that adds one records that value (Word rejects by setting what was recorded over
#: what is there, so an attribute left out of the record would stay), and rejecting drops
#: it again where it is that value.
DEFAULTS = {("cols", "num"): "1", ("cols", "sep"): "0", ("cols", "equalWidth"): "1",
            ("pgSz", "orient"): "portrait", ("pgNumType", "fmt"): "decimal", ("lnNumType", "countBy"): "0"}
#: Kinds whose meaning is their children: rejecting puts the recorded children back.
_PARENTS = frozenset({"cols", "footnotePr", "endnotePr"})


def twips(points: float) -> int:
    return int(round(float(points) * 20))


def points(value: str | int | None) -> float | None:
    if value is None:
        return None
    try:
        return int(value) / 20
    except (TypeError, ValueError):
        return None


def _int(node: Element | None, attribute: str, default: int | None = None) -> int | None:
    if node is None:
        return default
    try:
        return int(node.get(_W + attribute))
    except (TypeError, ValueError):
        return default


def _on(node: Element | None) -> bool:
    return node is not None and (node.get(_W + "val") or "true").lower() not in ("0", "false", "off")


def is_absent_marker(node: Element) -> bool:
    """Whether a kind restored from a ``w:sectPrChange`` is the "off" form a tracked change
    recorded for a kind that was absent (:data:`_ABSENT`): rejecting removes it."""
    tag = etree.QName(node).localname
    if tag not in _ABSENT or len(node):
        return False
    return _canonical(node) == _canonical(_ABSENT[tag]())


def restore_kind(sect: Element, recorded: Element) -> None:
    """Reject one recorded kind of a ``w:sectPrChange`` into ``sect`` as Word does: its
    attributes set over the element's (its children, for a kind that is its children,
    replacing them), an "off" form removing the element, a default this module recorded
    for an absent attribute dropped again."""
    tag = etree.QName(recorded).localname
    same = sect.find(recorded.tag)
    if is_absent_marker(recorded):
        if same is not None:
            remove(same)
        return
    if same is None:
        append_in_order(sect, copy.deepcopy(recorded))
        same = sect.find(recorded.tag)
    else:
        for name, value in recorded.attrib.items():
            same.set(name, value)
        if tag in _PARENTS or len(recorded):
            for child in list(same):
                remove(child)
            for child in recorded:
                if isinstance(child.tag, str):
                    same.append(copy.deepcopy(child))
    for (kind, name), value in DEFAULTS.items():
        if kind == tag and same.get(_W + name) == value:
            del same.attrib[_W + name]


# -- the view --------------------------------------------------------------------------------


class Section:
    """A section: the ``w:sectPr`` in the paragraph that ends it (``s:<paraId>``), or the
    body's last (``s:body``).  Resolved by id on every access."""

    def __init__(self, document: "Document", identifier: str, element: Element | None = None) -> None:
        self._document = document
        self.id = identifier

    @property
    def _sectPr(self) -> Element:
        return self._document._section_record(self.id)[1]

    @property
    def index(self) -> int:
        """The 0-based place among the document's sections."""
        return [s.id for s in self._document.sections()].index(self._document._canonical_section(self.id))

    @property
    def start(self) -> str:
        """How the section starts: ``nextPage`` (the default), ``continuous``, ``evenPage``,
        ``oddPage`` or ``nextColumn``."""
        node = self._sectPr.find(_W + "type")
        return node.get(_W + "val") if node is not None and node.get(_W + "val") else "nextPage"

    @property
    def page_size(self) -> tuple[int, int] | None:
        """``(width, height)`` in twips, as declared."""
        node = self._sectPr.find(_W + "pgSz")
        if node is None:
            return None
        return _int(node, "w", 0), _int(node, "h", 0)

    @property
    def page_width(self) -> float | None:
        """The page's width in points, or ``None`` when the section states none."""
        size = self.page_size
        return size[0] / 20 if size else None

    @property
    def page_height(self) -> float | None:
        """The page's height in points, or ``None`` when the section states none."""
        size = self.page_size
        return size[1] / 20 if size else None

    @property
    def orientation(self) -> str:
        """``"portrait"`` or ``"landscape"``."""
        node = self._sectPr.find(_W + "pgSz")
        return "landscape" if node is not None and node.get(_W + "orient") == "landscape" else "portrait"

    @property
    def margins(self) -> dict[str, float | None]:
        """``top``, ``right``, ``bottom``, ``left``, ``header``, ``footer`` and ``gutter``,
        in points."""
        node = self._sectPr.find(_W + "pgMar")
        return {side: points(node.get(_W + side)) if node is not None else None
                for side in ("top", "right", "bottom", "left", "header", "footer", "gutter")}

    @property
    def text_width(self) -> int | None:
        """The width between the margins (and gutter), in twips."""
        size, node = self.page_size, self._sectPr.find(_W + "pgMar")
        if size is None:
            return None
        return size[0] - (_int(node, "left", 0) or 0) - (_int(node, "right", 0) or 0) - (_int(node, "gutter", 0) or 0)

    @property
    def columns(self) -> dict:
        """``count``, ``space`` (points), ``separator`` and, for unequal columns, ``widths``
        (``[(width, space after)]`` in points)."""
        node = self._sectPr.find(_W + "cols")
        space = node.get(_W + "space") if node is not None else None
        out = {"count": _int(node, "num", 1) or 1, "space": points(space) if space else 36.0,
               "separator": node is not None and (node.get(_W + "sep") or "0").lower() in ("1", "true", "on")}
        if node is not None and node.findall(_W + "col"):
            out["widths"] = [(points(col.get(_W + "w")), points(col.get(_W + "space")) or 0.0)
                             for col in node.findall(_W + "col")]
            out["count"] = len(out["widths"])
        return out

    @property
    def line_numbering(self) -> dict | None:
        """``count_by``, ``start`` (1-based), ``distance`` (points) and ``restart``, or
        ``None`` when lines are not numbered."""
        node = self._sectPr.find(_W + "lnNumType")
        if node is None or node.get(_W + "countBy") in (None, "0"):
            return None
        return {"count_by": _int(node, "countBy", 1), "start": (_int(node, "start", 0) or 0) + 1,
                "distance": points(node.get(_W + "distance")), "restart": node.get(_W + "restart") or "newPage"}

    @property
    def vertical_alignment(self) -> str:
        """Where the text sits on the page vertically: ``"top"``, ``"center"``, ``"both"``,
        ``"bottom"``."""
        node = self._sectPr.find(_W + "vAlign")
        return node.get(_W + "val") if node is not None and node.get(_W + "val") else "top"

    @property
    def title_page(self) -> bool:
        """``w:titlePg``: whether the first page shows the ``first`` header and footer."""
        return _on(self._sectPr.find(_W + "titlePg"))

    @property
    def page_numbering(self) -> dict:
        """``format`` (``decimal`` unless stated) and ``start`` (``None``: continues)."""
        node = self._sectPr.find(_W + "pgNumType")
        return {"format": (node.get(_W + "fmt") if node is not None else None) or "decimal",
                "start": _int(node, "start")}

    def notes(self, kind: str = "footnote") -> dict:
        """The section's own ``w:footnotePr`` / ``w:endnotePr``: ``format``, ``start``,
        ``restart`` and (footnotes) ``position``; ``None`` where unstated."""
        node = self._sectPr.find(_W + ("footnotePr" if kind == "footnote" else "endnotePr"))

        def value(tag: str) -> str | None:
            child = node.find(_W + tag) if node is not None else None
            return child.get(_W + "val") if child is not None else None

        start = value("numStart")
        return {"format": value("numFmt"), "start": int(start) if start and start.isdigit() else None,
                "restart": value("numRestart"), "position": value("pos")}

    @property
    def paragraph_ids(self) -> list[str]:
        """The ids of the body's paragraphs in the section (in tables too), in order."""
        return self._document._section_paragraphs(self.id)

    def header(self, kind: str = "default") -> "Story | None":
        """The header the section shows for ``kind`` -- its own, or the one it inherits."""
        return self._document._section_story(self.id, "header", kind)

    def footer(self, kind: str = "default") -> "Story | None":
        """The story of the footer of ``kind`` (``"default"``, ``"first"``, ``"even"``) the
        section shows, or ``None``."""
        return self._document._section_story(self.id, "footer", kind)

    def stories(self) -> dict[str, dict[str, dict]]:
        """``{"header": {kind: {"story": name or None, "own": bool}}, "footer": {...}}``."""
        out: dict[str, dict[str, dict]] = {}
        for which in ("header", "footer"):
            out[which] = {}
            for kind in KINDS:
                found = self._document._section_story(self.id, which, kind)
                own = self._document._own_reference(self._sectPr, which, kind) is not None
                out[which][kind] = {"story": found.name if found is not None else None, "own": own}
        return out

    # -- edits -------------------------------------------------------------------------------

    def set(self, **values) -> "EditResult":
        """Set the section's page setup: :meth:`Document.set_section`.
        ``section.set(orientation="landscape", columns=2)``."""
        return self._document.set_section(self.id, **values)

    def remove_break(self, *, join: bool = True) -> "EditResult":
        """Remove the break that ends this section: :meth:`Document.remove_section_break`."""
        return self._document.remove_section_break(self.id, join=join)

    def add_header(self, kind: str = "default", text: str = "") -> "EditResult":
        """A header of its own for this section: :meth:`Document.add_header`."""
        return self._document.add_header(self.id, kind, text)

    def add_footer(self, kind: str = "default", text: str = "") -> "EditResult":
        """A footer of its own for this section: :meth:`Document.add_footer`."""
        return self._document.add_footer(self.id, kind, text)

    def link_to_previous(self, kind: str = "default", *, footer: bool = False) -> "EditResult":
        """Show the previous section's header (footer) instead: :meth:`Document.link_to_previous`."""
        return self._document.link_to_previous(self.id, kind, footer=footer)

    def unlink_from_previous(self, kind: str = "default", *, footer: bool = False) -> "EditResult":
        """Give the section a copy of its own: :meth:`Document.unlink_from_previous`."""
        return self._document.unlink_from_previous(self.id, kind, footer=footer)

    def __repr__(self) -> str:
        return f"<Section {self.id}>"


# -- helpers ---------------------------------------------------------------------------------


def section_id(paragraph_id: str) -> str:
    if paragraph_id.startswith("p@"):
        return "s@" + paragraph_id[2:]
    return "s:" + paragraph_id.rpartition("p:")[2]


def in_body(element: Element) -> bool:
    """Whether a paragraph is a block of the body (directly, or through block-level content
    controls and custom XML): where a section break may stand."""
    node = element.getparent()
    while node is not None:
        if node.tag == W_BODY:
            return True
        if node.tag not in (_W + "sdtContent", _W + "sdt", _W + "customXml"):
            return False
        node = node.getparent()
    return False


def _children(sect: Element) -> dict[str, Element]:
    """A ``w:sectPr``'s property children by local name (references and the change record
    aside)."""
    return {etree.QName(child).localname: child for child in sect
            if isinstance(child.tag, str) and child.tag not in _REFERENCES and child.tag != _W + "sectPrChange"}


def _canonical(node: Element | None):
    """What an element says, whatever prefixes its namespaces have where it stands."""
    if node is None or not isinstance(node.tag, str):
        return None
    return (node.tag, tuple(sorted(node.attrib.items())), (node.text or "").strip(),
            tuple(_canonical(child) for child in node if isinstance(child.tag, str)))


def _set_attributes(node: Element, values: dict[str, str | None], order: tuple[str, ...]) -> None:
    """Set (``None``: drop) ``w:`` attributes, rewriting them in ``order`` -- the order
    Word writes them in."""
    current = {etree.QName(name).localname: value for name, value in node.attrib.items()
               if name.startswith(_W)}
    others = {name: value for name, value in node.attrib.items() if not name.startswith(_W)}
    current.update(values)
    node.attrib.clear()
    for name in order:
        if current.get(name) is not None:
            node.set(_W + name, str(current.pop(name)))
    for name, value in current.items():
        if value is not None:
            node.set(_W + name, str(value))
    for name, value in others.items():
        node.set(name, value)


def _child(sect: Element, tag: str) -> Element:
    node = sect.find(_W + tag)
    if node is None:
        node = make(f"w:{tag}")
        append_in_order(sect, node)
    return node


def _drop(sect: Element, tag: str) -> None:
    for node in sect.findall(_W + tag):
        remove(node)


def _note_properties(sect: Element, tag: str, values: dict) -> None:
    """Write a section's ``w:footnotePr`` / ``w:endnotePr`` children: ``pos``, ``numFmt``,
    ``numStart``, ``numRestart`` (``None`` removes one; an empty element goes)."""
    node = sect.find(_W + tag)
    if node is None:
        node = make(f"w:{tag}")
        append_in_order(sect, node)
    for name, value in values.items():
        for old in node.findall(_W + name):
            remove(old)
        if value is not None:
            append_in_order(node, make(f"w:{name}", **{"w:val": str(value)}))
    if not len(node):
        remove(node)


_PGMAR_ORDER = ("top", "right", "bottom", "left", "header", "footer", "gutter")
_COLS_ORDER = ("num", "sep", "space", "equalWidth")
_MARGINS = {"margin_top": "top", "margin_right": "right", "margin_bottom": "bottom", "margin_left": "left",
            "header_distance": "header", "footer_distance": "footer", "gutter": "gutter"}
#: Every keyword :meth:`SectionOps.set_section` takes.
SECTION_PROPERTIES = frozenset({
    "start", "page_width", "page_height", "orientation", *_MARGINS, "columns", "column_space",
    "column_separator", "column_widths", "line_numbering", "vertical_alignment", "title_page",
    "page_number_format", "page_number_start", "footnote_format", "footnote_start", "footnote_restart",
    "footnote_position", "endnote_format", "endnote_start", "endnote_restart"})


def apply(sect: Element, values: dict) -> None:
    """Write section properties (:data:`SECTION_PROPERTIES`) into a ``w:sectPr``."""
    if "start" in values:
        start = values["start"]
        if start in (None, "nextPage"):
            _drop(sect, "type")
        else:
            _child(sect, "type").set(_W + "val", start)
    if any(k in values for k in ("page_width", "page_height", "orientation")):
        size = _child(sect, "pgSz")
        width, height = _int(size, "w", 11906), _int(size, "h", 16838)
        if values.get("page_width") is not None:
            width = twips(values["page_width"])
        if values.get("page_height") is not None:
            height = twips(values["page_height"])
        orient = size.get(_W + "orient")
        if "orientation" in values:
            landscape = values["orientation"] == "landscape"
            if landscape and width < height or not landscape and width > height:
                width, height = height, width
            orient = "landscape" if landscape else None
        _set_attributes(size, {"w": str(width), "h": str(height), "orient": orient}, ("w", "h", "orient", "code"))
    margins = {side: values[key] for key, side in _MARGINS.items() if key in values}
    if margins:
        node = _child(sect, "pgMar")
        defaults = {"top": 1440, "right": 1440, "bottom": 1440, "left": 1440, "header": 720, "footer": 720,
                    "gutter": 0}
        written = {side: node.get(_W + side) or str(defaults[side]) for side in _PGMAR_ORDER}
        written.update({side: str(twips(value or 0)) for side, value in margins.items()})
        _set_attributes(node, written, _PGMAR_ORDER)
    if any(k in values for k in ("columns", "column_space", "column_separator", "column_widths")):
        node = _child(sect, "cols")
        attributes: dict[str, str | None] = {}
        widths = values.get("column_widths")
        if "column_widths" in values:
            for col in node.findall(_W + "col"):
                remove(col)
            if widths:
                attributes["equalWidth"] = "0"
                attributes["num"] = str(len(widths))
                for width, space in widths:
                    col = make("w:col", **{"w:w": str(twips(width))})
                    if space:
                        col.set(_W + "space", str(twips(space)))
                    node.append(col)
            else:
                attributes["equalWidth"] = None
        if "columns" in values and not widths:
            count = int(values["columns"] or 1)
            attributes["num"] = str(count) if count > 1 else None
            if count <= 1:
                attributes["sep"] = None
                attributes["equalWidth"] = None
                for col in node.findall(_W + "col"):
                    remove(col)
        if "column_space" in values:
            attributes["space"] = str(twips(values["column_space"])) if values["column_space"] is not None else None
        if "column_separator" in values:
            attributes["sep"] = "1" if values["column_separator"] else None
        _set_attributes(node, attributes, _COLS_ORDER)
    if "line_numbering" in values:
        numbering = values["line_numbering"]
        _drop(sect, "lnNumType")
        if numbering:
            numbering = {} if numbering is True else dict(numbering)
            node = make("w:lnNumType")
            attributes = {"countBy": str(int(numbering.get("count_by", 1)))}
            if numbering.get("start", 1) not in (None, 1):
                attributes["start"] = str(int(numbering["start"]) - 1)
            if numbering.get("distance") is not None:
                attributes["distance"] = str(twips(numbering["distance"]))
            if numbering.get("restart") not in (None, "newPage"):
                attributes["restart"] = numbering["restart"]
            _set_attributes(node, attributes, ("countBy", "start", "distance", "restart"))
            append_in_order(sect, node)
    if "vertical_alignment" in values:
        value = values["vertical_alignment"]
        if value in (None, "top"):
            _drop(sect, "vAlign")
        else:
            _child(sect, "vAlign").set(_W + "val", value)
    if "title_page" in values:
        _drop(sect, "titlePg")
        if values["title_page"]:
            append_in_order(sect, make("w:titlePg"))
    if "page_number_format" in values or "page_number_start" in values:
        node = _child(sect, "pgNumType")
        attributes = {}
        if "page_number_format" in values:
            fmt = values["page_number_format"]
            attributes["fmt"] = None if fmt in (None, "decimal") else fmt
        if "page_number_start" in values:
            start = values["page_number_start"]
            attributes["start"] = None if start is None else str(int(start))
        _set_attributes(node, attributes, ("fmt", "start", "chapStyle", "chapSep"))
        if not node.attrib and not len(node):
            remove(node)
    for kind in ("footnote", "endnote"):
        wanted = {}
        for key, tag in (("format", "numFmt"), ("start", "numStart"), ("restart", "numRestart"),
                         ("position", "pos")):
            name = f"{kind}_{key}"
            if name in values:
                value = values[name]
                if key == "restart" and value == "continuous" or key == "start" and value == 1 \
                        or key == "position" and value == ("pageBottom" if kind == "footnote" else None):
                    value = None
                wanted[tag] = value
        if wanted:
            _note_properties(sect, f"{kind}Pr", wanted)


def check_values(values: dict) -> None:
    unknown = set(values) - SECTION_PROPERTIES
    if unknown:
        raise EditError(f"no section property {', '.join(sorted(unknown))}; the properties are "
                        f"{', '.join(sorted(SECTION_PROPERTIES))}")
    if values.get("start") not in (None, *STARTS):
        raise EditError(f"start is one of {', '.join(STARTS)}")
    if values.get("orientation") not in (None, "portrait", "landscape"):
        raise EditError("orientation is 'portrait' or 'landscape'")
    if values.get("vertical_alignment") not in (None, *VERTICAL):
        raise EditError(f"vertical_alignment is one of {', '.join(VERTICAL)}")
    for kind in ("footnote", "endnote"):
        if values.get(f"{kind}_restart") not in (None, *RESTARTS):
            raise EditError(f"{kind}_restart is one of {', '.join(RESTARTS)}")
    if values.get("footnote_position") not in (None, "pageBottom", "beneathText"):
        raise EditError("footnote_position is 'pageBottom' or 'beneathText'")
    if "endnote_position" in values:
        raise EditError("an endnote's position is the document's (set_note_settings), not a section's")
    numbering = values.get("line_numbering")
    if isinstance(numbering, dict):
        if numbering.get("restart") not in (None, *LINE_RESTARTS):
            raise EditError(f"line_numbering restart is one of {', '.join(LINE_RESTARTS)}")
        if int(numbering.get("count_by", 1)) < 1:
            raise EditError("line_numbering count_by is 1 or more")
    for key in ("page_width", "page_height", *_MARGINS, "column_space"):
        value = values.get(key)
        if value is not None and not isinstance(value, (int, float)):
            raise EditError(f"{key} is a number of points")
    for key in ("page_width", "page_height"):
        if values.get(key) is not None and values[key] <= 0:
            raise EditError(f"{key} must be positive")
    if values.get("columns") is not None and not 1 <= int(values["columns"]) <= 45:
        raise EditError("columns is 1 to 45")


def record_change(sect: Element, before: dict[str, Element], stamp, part: str) -> None:
    """Record, in the section's ``w:sectPrChange``, the old value of every kind of property
    that differs from ``before`` (Word records only what changed; a kind that was absent
    in its "off" form) -- merged into a change already there, from which a kind set back
    to its old value is dropped."""
    after = _children(sect)
    record = sect.find(_W + "sectPrChange")
    old = record.find(W_SECTPR) if record is not None else None
    recorded = _children(old) if old is not None else {}
    kinds = set(before) | set(after)
    changed = {k for k in kinds if _canonical(before.get(k)) != _canonical(after.get(k))}
    if record is not None and record.get(_W + "author") != stamp.author and changed - set(recorded):
        raise EditError("the section has another author's pending property change: accept or reject it first")
    wanted: dict[str, Element] = dict(recorded)
    for kind in changed - set(recorded):
        if kind in before:
            wanted[kind] = copy.deepcopy(before[kind])
            # An attribute the change added is recorded at the value its absence meant.
            for name in after[kind].attrib if kind in after else ():
                local = etree.QName(name).localname
                if name not in before[kind].attrib and (kind, local) in DEFAULTS:
                    wanted[kind].set(name, DEFAULTS[(kind, local)])
        elif kind in _ABSENT:
            wanted[kind] = _ABSENT[kind]()
        else:
            raise EditError(f"a tracked change cannot add {kind} where the section had none")
    for kind in list(wanted):
        now = after.get(kind)
        if now is None and is_absent_marker(wanted[kind]):
            del wanted[kind]
            continue
        trial = copy.deepcopy(now) if now is not None else None
        if trial is not None:
            holder = make("w:sectPr")
            holder.append(trial)
            restore_kind(holder, wanted[kind])
            if _canonical(holder.find(trial.tag)) == _canonical(now):
                del wanted[kind]
    if record is not None:
        remove(record)
    if not wanted:
        return
    record = stamp.make("w:sectPrChange", part)
    inner = make("w:sectPr")
    record.append(inner)
    for node in wanted.values():
        append_in_order(inner, copy.deepcopy(node))
    append_in_order(sect, record)


# -- the Document half -----------------------------------------------------------------------


class SectionOps:
    """Sections, section breaks, page setup and the headers and footers sections show, on
    :class:`docx_agent.Document`."""

    # -- reading -----------------------------------------------------------------------------

    def _section_records(self: "Document") -> list[tuple[str, Element, Element | None]]:
        """``(id, w:sectPr, paragraph that ends it or None)`` for every section, in order."""
        body_part = self.package.document_part()
        out = []
        for entry in self._index(body_part).paragraphs:
            properties = entry.element.find(W_PPR)
            sect = properties.find(W_SECTPR) if properties is not None else None
            if sect is not None and in_body(entry.element):
                out.append((section_id(entry.id), sect, entry.element))
        body = self.package.tree(body_part).find(W_BODY)
        last = body.find(W_SECTPR) if body is not None else None
        if last is not None:
            out.append(("s:body", last, None))
        return out

    def sections(self: "Document") -> list[Section]:
        """Every section of the body, in order (``s:<paraId>``..., the last ``s:body``)."""
        return [Section(self, identifier) for identifier, _, _ in self._section_records()]

    def _canonical_section(self: "Document", identifier: str) -> str:
        if identifier.startswith(("s:", "s@")) and identifier != "s:body":
            paragraph = "p" + identifier[1:]
            seen = set()
            while paragraph in self._aliases and paragraph not in seen:
                seen.add(paragraph)
                paragraph = self._aliases[paragraph]
            return section_id(paragraph)
        return identifier

    def _section_record(self: "Document", identifier: str) -> tuple[str, Element, Element | None]:
        wanted = self._canonical_section(identifier)
        for record in self._section_records():
            if record[0] in (identifier, wanted):
                return record
        raise KeyError(f"no section {identifier!r}")

    def section(self: "Document", identifier: str) -> Section:
        """A section by its id (``s:<paraId>``, ``s:body``); a paragraph id names the
        section the paragraph is in."""
        if identifier.startswith(("s:", "s@")):
            self._section_record(identifier)
            return Section(self, identifier)
        return self.section_of(identifier)

    def section_of(self: "Document", identifier: str) -> Section:
        """The section a body paragraph or table is in."""
        part, entry = self._resolve(identifier)
        if part != self.package.document_part():
            raise KeyError(f"{identifier} is not in the body")
        records = self._section_records()
        nodes = list(self.package.tree(part).iter())  # held: one proxy per element, stable id()s
        order = {id(node): k for k, node in enumerate(nodes)}
        at = order[id(entry.element)]
        for identifier_, _, paragraph in records:
            if paragraph is None or order[id(paragraph)] >= at:
                return Section(self, identifier_)
        return Section(self, records[-1][0])

    def _section_paragraphs(self: "Document", identifier: str) -> list[str]:
        records = self._section_records()
        wanted = self._section_record(identifier)[0]
        part = self.package.document_part()
        nodes = list(self.package.tree(part).iter())  # held: one proxy per element, stable id()s
        order = {id(node): k for k, node in enumerate(nodes)}
        start = -1
        for record_id, _, paragraph in records:
            end = order[id(paragraph)] if paragraph is not None else float("inf")
            if record_id == wanted:
                return [e.id for e in self._index(part).paragraphs if start < order[id(e.element)] <= end]
            start = end
        return []

    # -- headers and footers: reading --------------------------------------------------------

    def _own_reference(self: "Document", sect: Element, which: str, kind: str) -> Element | None:
        tag = _W + ("headerReference" if which == "header" else "footerReference")
        for node in sect.findall(tag):
            if (node.get(_W + "type") or "default") == kind:
                return node
        return None

    def _reference_part(self: "Document", node: Element) -> str | None:
        return self.package.related_part(self.package.document_part(), node.get(_R + "id"))

    def _inherited_reference(self: "Document", index: int, which: str, kind: str) -> Element | None:
        """The reference section ``index`` shows for ``which``/``kind``: its own, else the
        last one a section before it states (inheritance is per kind: H.2)."""
        records = self._section_records()
        for k in range(index, -1, -1):
            found = self._own_reference(records[k][1], which, kind)
            if found is not None:
                return found
        return None

    def _section_story(self: "Document", identifier: str, which: str, kind: str):
        from .document import Story

        if kind not in KINDS:
            raise EditError(f"a header or footer kind is one of {', '.join(KINDS)}")
        records = self._section_records()
        wanted = self._section_record(identifier)[0]
        index = [r[0] for r in records].index(wanted)
        node = self._inherited_reference(index, which, kind)
        part = self._reference_part(node) if node is not None else None
        if part is None or not self.package.has_part(part):
            return None
        return Story(self, self._story_of(part), part)

    @property
    def even_and_odd_headers(self: "Document") -> bool:
        """``w:evenAndOddHeaders``: document-wide, whether even pages show the ``even``
        header and footer."""
        part = self.package.settings_part()
        root = self.package.tree(part) if part else None
        return _on(root.find(_W + "evenAndOddHeaders")) if root is not None else False

    @even_and_odd_headers.setter
    def even_and_odd_headers(self: "Document", value: bool) -> None:
        self.set_even_and_odd_headers(value)

    def set_even_and_odd_headers(self: "Document", value: bool) -> "EditResult":
        """Turn ``w:evenAndOddHeaders`` on or off -- for every section of the document."""
        from .document import EditResult

        if self.even_and_odd_headers == bool(value):
            return EditResult(None, changed=False)
        with self._edit():
            self._write_even_and_odd(bool(value))
        return EditResult(None, warnings=["w:evenAndOddHeaders is document-wide: every section's even pages "
                                          + ("now show their even header and footer" if value else
                                             "now show their default header and footer")])

    def _write_even_and_odd(self: "Document", value: bool) -> None:
        part = self.package.settings_part() or self._create_settings_part()
        root = self.package.tree(part)
        for node in root.findall(_W + "evenAndOddHeaders"):
            remove(node)
        if value:
            append_in_order(root, make("w:evenAndOddHeaders"))
        self.package.mark_dirty(part)

    def _create_settings_part(self: "Document") -> str:
        """A settings part for a document that has none (no compatibility mode stated, so
        the document's mode -- 12 -- is unchanged)."""
        from ..oxml.package import REL_SETTINGS

        part = "word/settings.xml"
        data = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        self.package.add_part(part, data, f"{_WML}.settings+xml", override=True)
        self.package.add_relationship(self.package.document_part(), REL_SETTINGS, part)
        return part

    # -- section breaks ----------------------------------------------------------------------

    def insert_section_break(self: "Document", *, after: str | None = None, before: str | None = None,
                             kind: str = "nextPage") -> "EditResult":
        """A section break between two blocks of the body: a new paragraph holding the
        ``w:sectPr`` of the section it splits, as Word writes a break put at a paragraph's
        start (module docstring).  ``kind`` is how the section after the break starts
        (:data:`STARTS`).  The result's ``id`` is the new section's (the one before the
        break); ``created`` holds it and the break paragraph's id."""
        from .document import EditResult

        if (after is None) == (before is None):
            raise EditError("give exactly one of after= or before=")
        if kind not in STARTS:
            raise EditError(f"a section break's kind is one of {', '.join(STARTS)}")
        anchor_id = after if after is not None else before
        part, entry = self._resolve(anchor_id)  # type: ignore[arg-type]
        if part != self.package.document_part():
            raise EditError("a section break goes between blocks of the body")
        anchor = entry.element
        if anchor.getparent() is None or anchor.getparent().tag != W_BODY:
            raise EditError(f"{anchor_id} is not a block of the body (a section break cannot stand in a "
                            "table, a content control or a note)")
        # The block after the break: its section is the one split, its properties the
        # break paragraph's.  With none (the body's end), the body's last section is split
        # and an empty paragraph after the break keeps it a paragraph.
        following = _next_block(anchor) if after is not None else anchor
        template = following if following is not None and following.tag == W_P else (
            anchor if anchor.tag == W_P else None)
        position = entry.index if hasattr(entry, "index") else self._block_position(part, anchor, after is not None)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [], position, including=before is not None)
            records = self._section_records()
            split = records[-1] if following is None else self._section_of_block(records, following)
            sect = split[1]
            new_sect = copy.deepcopy(sect)
            for node in new_sect.findall(_W + "sectPrChange"):
                remove(node)
            paragraph = _text._make("w:p")
            properties = _break_properties(template)
            properties.append(new_sect)
            paragraph.append(properties)
            para_id, text_id = self._fresh_ids(part)
            _ids.stamp(paragraph, para_id, text_id)
            if after is not None:
                anchor.addnext(paragraph)
            else:
                anchor.addprevious(paragraph)
            created = [f"p:{para_id}"]
            filler = None
            if following is None:
                filler = _text._make("w:p")
                filler_properties = _break_properties(template)
                if len(filler_properties):
                    filler.append(filler_properties)
                filler_id, filler_text = self._fresh_ids(part)
                _ids.stamp(filler, filler_id, filler_text)
                paragraph.addnext(filler)
                created.append(f"p:{filler_id}")
            before_values = {k: copy.deepcopy(v) for k, v in _children(sect).items()}
            apply(sect, {"start": kind})
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                _track.mark(paragraph, "ins", stamp, part)
                if filler is not None:
                    _track.mark(filler, "ins", stamp, part)
                record_change(sect, before_values, stamp, part)
                stamp.finish()
            else:
                # Word gives the split section's references to the new section before it,
                # from which the split one inherits them (``tracked``).
                for node in list(sect):
                    if isinstance(node.tag, str) and node.tag in _REFERENCES:
                        remove(node)
            _ids.ensure_w14(self.package.tree(part))
            self.package.mark_dirty(part)
        new_section = section_id(created[0])
        return EditResult(new_section, created=[new_section, *created], renamed=renames)

    def _section_of_block(self: "Document", records, block: Element):
        """The section record a body block is in: the first whose end is at or after it."""
        part = self.package.document_part()
        nodes = list(self.package.tree(part).iter())  # held: one proxy per element, stable id()s
        order = {id(node): k for k, node in enumerate(nodes)}
        at = order[id(block)]
        for record in records:
            if record[2] is None or order[id(record[2])] >= at:
                return record
        return records[-1]

    def remove_section_break(self: "Document", identifier: str, *, join: bool = True) -> "EditResult":
        """Remove the break that ends section ``identifier``: its paragraph mark is deleted,
        as Word deletes a break -- the paragraph joins the one after it, whose element, id
        and properties stay (an empty break paragraph simply goes) -- and the text before it
        becomes the next section's.  A header or footer the section stated and the next
        does not is handed to the next (which showed it already); the rest go, their parts
        too when nothing else uses them.  ``join=False`` keeps the paragraph (only its
        section properties go).  Tracked, the mark is deleted as a revision."""
        from .document import EditResult

        records = self._section_records()
        wanted = self._section_record(identifier)[0]
        if wanted == "s:body":
            raise EditError("the body's last section has no break to remove")
        index = [r[0] for r in records].index(wanted)
        _, sect, paragraph = records[index]
        next_sect = records[index + 1][1]
        part = self.package.document_part()
        entry = self._index(part).entry_for(paragraph)
        following = _next_block(paragraph)
        empty = not _has_content(paragraph)
        tracking = self._active_tracking()
        if tracking is not None and (following is None or following.tag != W_P):
            raise EditError(f"{identifier}'s break is not followed by a paragraph: Word cannot join a tracked "
                            "break's mark into a table")
        with self._edit():
            renames = self._prepare(part, [entry], entry.index)
            entry = self._index(part).entry_for(paragraph)
            # Hand on what the section stated that the next does not: the next showed it.
            for node in list(sect):
                if not (isinstance(node.tag, str) and node.tag in _REFERENCES):
                    continue
                which = "header" if node.tag == _W + "headerReference" else "footer"
                kind = node.get(_W + "type") or "default"
                if self._own_reference(next_sect, which, kind) is None:
                    append_in_order(next_sect, copy.deepcopy(node))
                    _order_references(next_sect)
            removed: list[str] = []
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                if _track.delete_mark(paragraph, stamp, part) is not None:
                    removed.append(entry.id)  # one's own inserted break goes outright
                stamp.finish()
                self._reap_stories()
            else:
                released = [n.get(_R + "id") for n in sect if isinstance(n.tag, str) and n.tag in _REFERENCES]
                if empty and _survives(paragraph):
                    from .document import _rehome_markers

                    _rehome_markers(paragraph)
                    remove(paragraph)
                    removed.append(entry.id)
                elif join and following is not None and following.tag == W_P:
                    remove(sect)
                    at = 1 if following.find(W_PPR) is not None else 0
                    for child in [c for c in paragraph if not (isinstance(c.tag, str) and c.tag == W_PPR)]:
                        following.insert(at, child)
                        at += 1
                    remove(paragraph)
                    removed.append(entry.id)
                else:
                    remove(sect)
                    properties = paragraph.find(W_PPR)
                    if properties is not None and not len(properties):
                        remove(properties)
                self._release(part, [r for r in released if r])
            self.package.mark_dirty(part)
        result = EditResult(None, renamed={k: v for k, v in renames.items() if k not in removed},
                            removed=[wanted] + removed)
        if removed and following is not None and following.tag == W_P:
            survivor = self._index(part).entry_for(following)
            if survivor is not None:
                self._rename({removed[0]: survivor.id})
                result.id = survivor.id
        elif not removed:
            result.id = self._index(part).entry_for(paragraph).id
        return result

    def _reap_stories(self: "Document") -> list[str]:
        """Drop the main part's header and footer relationships no section references any
        more, and their parts."""
        main = self.package.document_part()
        rels = [rid for rid, rel in self.package.relationships(main).items()
                if rel.type in (REL_HEADER, REL_FOOTER)]
        return self._release(main, rels)

    # -- page setup --------------------------------------------------------------------------

    def set_section(self: "Document", identifier: str, **values) -> "EditResult":
        """Set a section's page setup (:data:`SECTION_PROPERTIES`, lengths in points):
        ``start``, ``page_width``, ``page_height``, ``orientation``, ``margin_top`` ...
        ``margin_left``, ``header_distance``, ``footer_distance``, ``gutter``, ``columns``,
        ``column_space``, ``column_separator``, ``column_widths`` (``[(width, space after)]``),
        ``line_numbering`` (``None`` off, or ``count_by``, ``start``, ``distance``,
        ``restart``), ``vertical_alignment``, ``title_page``, ``page_number_format``,
        ``page_number_start`` (``None`` continues), and the section's notes:
        ``footnote_format``, ``footnote_start``, ``footnote_restart``, ``footnote_position``,
        ``endnote_format``, ``endnote_start``, ``endnote_restart``.  Tracked, a
        ``w:sectPrChange``."""
        from .document import EditResult

        check_values(values)
        _, sect, paragraph = self._section_record(identifier)
        part = self.package.document_part()
        trial = copy.deepcopy(sect)
        apply(trial, values)
        if _canonical(trial) == _canonical(sect):
            return EditResult(self._canonical_section(identifier), changed=False)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [])
            _, sect, paragraph = self._section_record(identifier)
            before = {k: copy.deepcopy(v) for k, v in _children(sect).items()}
            apply(sect, values)
            if tracking is not None:
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                record_change(sect, before, stamp, part)
                stamp.finish()
            self.package.mark_dirty(part)
        section = self._canonical_section(identifier)
        if section != "s:body" and paragraph is not None:
            entry = self._index(part).entry_for(paragraph)
            section = section_id(entry.id) if entry is not None else section
        return EditResult(section, renamed=renames)

    # -- headers and footers: editing --------------------------------------------------------

    def add_header(self: "Document", section: str, kind: str = "default", text: str = "") -> "EditResult":
        """A header of its own for the section's ``kind`` (``default``, ``first``, ``even``),
        holding ``text`` in the Header style; the result's ``id`` is its story's name
        (``header3``), ``created`` its paragraph's id.  A first-page header turns on
        ``w:titlePg``, an even one the document's ``w:evenAndOddHeaders`` (said in
        ``warnings``).  Refused when the section has its own already."""
        return self._add_story(section, "header", kind, text)

    def add_footer(self: "Document", section: str, kind: str = "default", text: str = "") -> "EditResult":
        """A footer of its own for the section's ``kind`` (``default``, ``first``, ``even``),
        holding ``text``, as :meth:`add_header` writes a header::

            doc.add_footer("s:body", "default", "Confidential")"""
        return self._add_story(section, "footer", kind, text)

    def _add_story(self: "Document", identifier: str, which: str, kind: str, text: str) -> "EditResult":
        from .document import EditResult

        if kind not in KINDS:
            raise EditError(f"a header or footer kind is one of {', '.join(KINDS)}")
        _text.check_text(text)
        _, sect, _ = self._section_record(identifier)
        if self._own_reference(sect, which, kind) is not None:
            raise EditError(f"{identifier} has its own {kind} {which} already")
        style = "header" if which == "header" else "footer"
        self.styles.resolve(style, "paragraph")
        tracking = self._active_tracking()
        warnings: list[str] = []
        with self._edit():
            renames = self._prepare(self.package.document_part(), [])
            _, sect, _ = self._section_record(identifier)
            style_id = self.styles._ensure(style, "paragraph")
            part = self._new_story_part(which, style_id, text)
            self._reference(sect, which, kind, part)
            warnings += self._needs(sect, kind)
            story = self._story_of(part)
            paragraph = self.package.tree(part).find(W_P)
            if tracking is not None and text:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                _track.insert_paragraph_content(paragraph, stamp, part)
                stamp.finish()
            self.package.mark_dirty(self.package.document_part())
        paragraph_id = f"{story}/p:{paragraph.get(_ids.PARA_ID)}"
        return EditResult(story, created=[story, paragraph_id], renamed=renames, warnings=warnings)

    def _needs(self: "Document", sect: Element, kind: str) -> list[str]:
        """What showing a ``kind`` story needs, turned on: ``w:titlePg`` for a first page,
        ``w:evenAndOddHeaders`` for even pages."""
        if kind == "first" and not _on(sect.find(_W + "titlePg")):
            for node in sect.findall(_W + "titlePg"):
                remove(node)
            append_in_order(sect, make("w:titlePg"))
            return ["the section's first page now shows its first-page header and footer (w:titlePg)"]
        if kind == "even" and not self.even_and_odd_headers:
            self._write_even_and_odd(True)
            return ["w:evenAndOddHeaders is document-wide: every section's even pages now show their even "
                    "header and footer"]
        return []

    def _new_story_part(self: "Document", which: str, style_id: str | None, text: str,
                        body: list[Element] | None = None) -> str:
        """A header or footer part as Word writes one: one paragraph in the Header (Footer)
        style, with ``text`` -- or ``body``'s blocks -- related from the main part."""
        tag = "hdr" if which == "header" else "ftr"
        path = self.package.unused_part_name(f"word/{which}{{n}}.xml")
        data = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                f'<w:{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>').encode()
        self.package.add_part(path, data, f"{_WML}.{which}+xml", override=True)
        root = self.package.tree(path)
        used = self._used()
        if body is None:
            paragraph = _text._make("w:p")
            if style_id is not None:
                properties = make("w:pPr")
                properties.append(make("w:pStyle", **{"w:val": style_id}))
                paragraph.append(properties)
            if text:
                for k, line in enumerate(text.split("\n")):
                    if k:
                        root.append(paragraph)
                        paragraph = copy.deepcopy(paragraph)
                        for child in list(paragraph):
                            if child.tag != W_PPR:
                                remove(child)
                    paragraph.append(_text.make_run(line, None))
            root.append(paragraph)
        else:
            for node in body:
                root.append(node)
        for node in root.iter(W_P, _W + "tr"):
            para_id, text_id = _ids.generate(path, used, 2)
            _ids.stamp(node, para_id, text_id)
        _ids.ensure_w14(root)
        self.package.mark_dirty(path)
        self._invalidate()
        return path

    def _reference(self: "Document", sect: Element, which: str, kind: str, part: str) -> None:
        rel = REL_HEADER if which == "header" else REL_FOOTER
        rid = self.package.add_relationship(self.package.document_part(), rel, part)
        node = make(f"w:{which}Reference", **{"w:type": kind})
        node.set(_R + "id", rid)
        sect.insert(0, node)
        _order_references(sect)

    def link_to_previous(self: "Document", section: str, kind: str = "default", *,
                         footer: bool = False) -> "EditResult":
        """Make the section show the ``kind`` header (footer) of the section before it:
        its own reference goes, and its part when nothing else uses it.  On the first
        section this removes the header: there is nothing before it."""
        from .document import EditResult

        which = "footer" if footer else "header"
        if kind not in KINDS:
            raise EditError(f"a header or footer kind is one of {', '.join(KINDS)}")
        _, sect, _ = self._section_record(section)
        node = self._own_reference(sect, which, kind)
        if node is None:
            return EditResult(self._canonical_section(section), changed=False)
        story = self._story_of(self._reference_part(node)) if self._reference_part(node) else None
        with self._edit():
            renames = self._prepare(self.package.document_part(), [])
            _, sect, _ = self._section_record(section)
            node = self._own_reference(sect, which, kind)
            rid = node.get(_R + "id")
            remove(node)
            removed_parts = self._release(self.package.document_part(), [rid])
        removed = [story] if story and any(self._story_of(p) == story for p in removed_parts) else []
        return EditResult(self._canonical_section(section), renamed=renames, removed=removed)

    def remove_header(self: "Document", section: str, kind: str = "default") -> "EditResult":
        """The section's own ``kind`` header goes (:meth:`link_to_previous`)."""
        return self.link_to_previous(section, kind)

    def remove_footer(self: "Document", section: str, kind: str = "default") -> "EditResult":
        """The section's own ``kind`` footer goes (:meth:`link_to_previous` with ``footer=True``)."""
        return self.link_to_previous(section, kind, footer=True)

    def unlink_from_previous(self: "Document", section: str, kind: str = "default", *,
                             footer: bool = False) -> "EditResult":
        """Give the section a ``kind`` header (footer) of its own, a copy of the one it
        showed (Word's unlinking: ``headers2``), or an empty one in the Header style when it
        showed none.  The result's ``id`` is the new story's name."""
        from .document import EditResult

        which = "footer" if footer else "header"
        if kind not in KINDS:
            raise EditError(f"a header or footer kind is one of {', '.join(KINDS)}")
        records = self._section_records()
        wanted = self._section_record(section)[0]
        index = [r[0] for r in records].index(wanted)
        sect = records[index][1]
        if self._own_reference(sect, which, kind) is not None:
            return EditResult(self._story_of(self._reference_part(self._own_reference(sect, which, kind))),
                              changed=False)
        inherited = self._inherited_reference(index, which, kind)
        source = self._reference_part(inherited) if inherited is not None else None
        if source is None:
            return self._add_story(section, which, kind, "")
        with self._edit():
            renames = self._prepare(self.package.document_part(), [])
            records = self._section_records()
            sect = records[index][1]
            path = self.package.copy_part(source, share=lambda relationship: True)
            root = self.package.tree(path)
            used = self._used()
            for node in root.iter(W_P, _W + "tr"):
                para_id, text_id = _ids.generate(path, used, 2)
                _ids.stamp(node, para_id, text_id)
            _ids.ensure_w14(root)
            self.package.mark_dirty(path)
            self._invalidate()
            self._reference(sect, which, kind, path)
            self.package.mark_dirty(self.package.document_part())
            story = self._story_of(path)
            created = [story] + [e.id for e in self._index(path).paragraphs]
        return EditResult(story, created=created, renamed=renames)


def _order_references(sect: Element) -> None:
    """Put a ``w:sectPr``'s references first, in Word's order (``headers``)."""
    references = [node for node in sect if isinstance(node.tag, str) and node.tag in _REFERENCES]

    def key(node: Element) -> int:
        which = "header" if node.tag == _W + "headerReference" else "footer"
        return _REFERENCE_ORDER.get((which, node.get(_W + "type") or "default"), 9)

    for node in references:
        sect.remove(node)
    for k, node in enumerate(sorted(references, key=key)):
        sect.insert(k, node)


def _next_block(element: Element) -> Element | None:
    node = element.getnext()
    while node is not None:
        if isinstance(node.tag, str) and node.tag in (W_P, _W + "tbl", _W + "sdt", _W + "customXml"):
            return node
        if isinstance(node.tag, str) and node.tag == W_SECTPR:
            return None
        node = node.getnext()
    return None


def _previous_block(element: Element) -> Element | None:
    node = element.getprevious()
    while node is not None:
        if isinstance(node.tag, str) and node.tag in (W_P, _W + "tbl", _W + "sdt", _W + "customXml"):
            return node
        node = node.getprevious()
    return None


def _survives(paragraph: Element) -> bool:
    from .document import _check_container_survives

    try:
        _check_container_survives(paragraph)
    except EditError:
        return False
    return True


def _has_content(paragraph: Element) -> bool:
    return any(isinstance(child.tag, str) and child.tag != W_PPR for child in paragraph)


def _break_properties(template: Element | None) -> Element:
    """A break paragraph's properties: the paragraph's it is put before (Word gives a break
    before a heading the heading's style: ``fields``), without its section break, list
    membership or revision records."""
    from ..oxml.xml import REVISION_PROPERTY_TAGS

    properties = make("w:pPr")
    source = template.find(W_PPR) if template is not None else None
    if source is not None:
        properties = copy.deepcopy(source)
        for child in list(properties):
            if child.tag in (W_SECTPR, _W + "pPrChange", _W + "numPr"):
                remove(child)
            elif child.tag == _W + "rPr":
                for mark in list(child):
                    if mark.tag in REVISION_PROPERTY_TAGS:
                        remove(mark)
                if not len(child):
                    remove(child)
    return properties


__all__ = ["KINDS", "SECTION_PROPERTIES", "STARTS", "Section", "SectionOps"]
