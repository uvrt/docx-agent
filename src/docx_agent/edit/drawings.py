"""Drawings: a picture floated and put inline again, a floating drawing's position, wrapping,
order and locks, text boxes and shapes made as Word makes them, groups moved and resized --
as Word 16.106 writes each (``tools/e5_probe.py``, ``tests/observations/e5-word.json``).

What Word writes, and docx-agent with it:

* **Floating a picture** (Word's ``convert to shape``): ``wp:anchor`` with ``distT``/``distB``
  0 and ``distL``/``distR`` 114300 (9 pt), ``simplePos="0"``, ``relativeHeight`` one step
  (1024) above every drawing's in the document -- Word's first is 251658240 --
  ``behindDoc="0"``, ``locked="0"``, ``layoutInCell="1"``, ``allowOverlap="1"``,
  ``wp:simplePos`` at 0, 0, ``wp:positionH`` against the column and ``wp:positionV``
  against the paragraph (offsets where it stood), the extent and effect extent,
  **``wp:wrapNone``** (in front of the text), the ``docPr``, frame properties and graphic;
  its ``wp14:anchorId`` kept.  **Inline again**: ``wp:inline`` with every distance 0.
* **Wrapping**: ``wp:wrapSquare wrapText="bothSides"`` (or ``left``, ``right``,
  ``largest``); ``wp:wrapTight`` and ``wp:wrapThrough`` with Word's polygon round the frame
  (0,0 0,21000 21300,21000 21300,0, ``edited="0"``); ``wp:wrapTopAndBottom``; in front of
  the text ``wp:wrapNone`` with ``behindDoc="0"``, behind it ``behindDoc="1"``.  Distances
  in ``distT``/``distB``/``distL``/``distR`` (EMU).
* **Position**: ``wp:posOffset`` (EMU, as given) or ``wp:align`` against ``relativeFrom``
  -- ``page``, ``margin``, ``column``, ``character``, ``leftMargin``, ``rightMargin``,
  ``insideMargin``, ``outsideMargin`` across; ``page``, ``margin``, ``paragraph``, ``line``,
  ``topMargin``, ``bottomMargin``, ``insideMargin``, ``outsideMargin`` down.
* **Order**: to the front, ``relativeHeight`` a step above the highest; to the back, 1025
  below the lowest (measured).  **Lock**: ``locked="1"``; **overlap**: ``allowOverlap``.
* **Resizing** a floating drawing adds ``wp14:sizeRelH``/``sizeRelV`` with a percentage of
  0 (measured).
* **A text box** (Word's Insert Text Box): in ``mc:AlternateContent``, a ``wps:wsp`` with
  ``txBox="1"``, no fill and a 0.5 pt black line, its ``w:txbxContent``, ``wps:bodyPr``
  with the default insets, wrapped square; the fallback a VML ``v:shape`` of
  ``#_x0000_t202`` holding the same ``w:txbxContent`` (paraIds and all) and ``w10:wrap``.
  Its paragraphs are edited through the paragraph API by their ids; after every edit the
  fallback's copy is made the same as what Word reads (``Document._edit``).
* **A shape** (Word's ``make new shape``): a ``wps:wsp`` with the preset geometry, the
  theme's style references (line ``accent1`` shaded, fill ``accent1``), an effect extent
  of a point for its line, its fallback a VML ``v:rect``, ``v:oval`` or ``v:roundrect``
  (other presets have none).

**Tracked**: Word tracks a new text box or shape as an inserted run, and nothing else here:
floating, wrapping, moving, ordering and resizing are applied untracked, and the result
says so.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, insert_in_order, make, remove
from . import ids as _ids

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
W10 = "urn:schemas-microsoft-com:office:word"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_WP = "{%s}" % WP
_WP14 = "{%s}" % WP14
_A = "{%s}" % A
_WPS = "{%s}" % WPS
_MC = "{%s}" % MC
_V = "{%s}" % V
EMU = 12700
#: Word's first ``relativeHeight`` and its step (measured: 251658240, then +1024 each).
FIRST_HEIGHT = 251658240
HEIGHT_STEP = 1024
#: Word's default horizontal distance of a floating drawing from the text (9 pt).
SIDE_DISTANCE = 114300
#: Word's tight and through polygon round a rectangular frame (measured).
POLYGON = ((0, 0), (0, 21000), (21300, 21000), (21300, 0), (0, 0))

WRAPS = {"square": "wrapSquare", "tight": "wrapTight", "through": "wrapThrough",
         "top_and_bottom": "wrapTopAndBottom", "front": "wrapNone", "none": "wrapNone", "behind": "wrapNone"}
SIDES = {"both": "bothSides", "left": "left", "right": "right", "largest": "largest"}
HORIZONTAL = ("page", "margin", "column", "character", "leftMargin", "rightMargin", "insideMargin",
              "outsideMargin")
VERTICAL = ("page", "margin", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin")
X_ALIGN = ("left", "center", "right", "inside", "outside")
Y_ALIGN = ("top", "center", "bottom", "inside", "outside")
#: Presets the VML fallback can draw, and its element for each.
VML_PRESETS = {"rect": "rect", "ellipse": "oval", "roundRect": "roundrect"}


def emu(points: float) -> int:
    return int(round(float(points) * EMU))


def _pt(value) -> float | None:
    try:
        return int(value) / EMU
    except (TypeError, ValueError):
        return None


def frame_of(drawing_frame: Element) -> Element:
    return drawing_frame


def graphic_shape(frame: Element) -> Element | None:
    """The element the frame's graphic draws: ``pic:pic``, ``wps:wsp``, ``wpg:wgp``..."""
    data = frame.find(f"{_A}graphic/{_A}graphicData")
    if data is None:
        return None
    return next((child for child in data if isinstance(child.tag, str)), None)


def _xfrm_ext(frame: Element) -> Element | None:
    """The ``a:ext`` of the shape's own transform (a picture's, a shape's, a group's)."""
    shape = graphic_shape(frame)
    if shape is None:
        return None
    for path in (".//{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr/" + _A + "xfrm/" + _A + "ext",
                 _WPS + "spPr/" + _A + "xfrm/" + _A + "ext", "{%s}grpSpPr/%sxfrm/%sext" % (WPG, _A, _A)):
        found = shape.find(path)
        if found is not None:
            return found
    return None


def alternate_content(frame: Element) -> Element | None:
    """The ``mc:AlternateContent`` whose choice holds the frame, if any."""
    node = frame.getparent()
    while node is not None and node.tag != _W + "r":
        if node.tag == _MC + "AlternateContent":
            return node
        node = node.getparent()
    return None


def fallback_shape(frame: Element) -> Element | None:
    """The VML shape of the frame's fallback (``v:shape``, ``v:rect``, ``v:group``...)."""
    holder = alternate_content(frame)
    fallback = holder.find(_MC + "Fallback") if holder is not None else None
    if fallback is None:
        return None
    pict = fallback.find(_W + "pict")
    if pict is None:
        return None
    return next((child for child in pict if isinstance(child.tag, str) and child.tag.startswith(_V)
                 and child.tag != _V + "shapetype"), None)


def _style(shape: Element) -> dict[str, str]:
    out = {}
    for item in (shape.get("style") or "").split(";"):
        key, _, value = item.partition(":")
        if key.strip():
            out[key.strip()] = value.strip()
    return out


def _set_style(shape: Element, values: dict[str, str]) -> None:
    shape.set("style", ";".join(f"{k}:{v}" for k, v in values.items()))


def _points_text(value: float) -> str:
    """A VML length as Word writes one: whole inches in inches (144 pt is ``2in``), the rest
    in points (measured)."""
    text = f"{value:.2f}".rstrip("0").rstrip(".")
    if text in ("", "-0", "0"):
        return "0"
    if abs(value) >= 72 and abs(value / 72 - round(value / 72)) < 1e-9:
        return f"{int(round(value / 72))}in"
    return f"{text}pt"


def all_heights(document: "Document") -> list[int]:
    out = []
    for part in document._parts():
        for node in document.package.tree(part).iter(_WP + "anchor"):
            try:
                out.append(int(node.get("relativeHeight")))
            except (TypeError, ValueError):
                pass
    return out


def read_position(frame: Element) -> dict:
    """A drawing's place: ``inline``, or ``{x, y, x_align, y_align, horizontal_from,
    vertical_from}`` (points)."""
    if frame.tag == _WP + "inline":
        return {"inline": True}
    out: dict = {"inline": False}
    for axis, tag in (("x", "positionH"), ("y", "positionV")):
        node = frame.find(_WP + tag)
        out[f"{'horizontal' if axis == 'x' else 'vertical'}_from"] = node.get("relativeFrom") if node is not None else None
        offset = node.find(_WP + "posOffset") if node is not None else None
        align = node.find(_WP + "align") if node is not None else None
        out[axis] = _pt(offset.text) if offset is not None else None
        out[f"{axis}_align"] = align.text if align is not None else None
    return out


def read_wrap(frame: Element) -> dict:
    if frame.tag == _WP + "inline":
        return {"wrap": "inline"}
    kinds = {v: k for k, v in WRAPS.items() if k not in ("none", "behind")}
    wrap = next((child for child in frame if isinstance(child.tag, str) and child.tag.startswith(_WP + "wrap")), None)
    name = kinds.get(wrap.tag[len(_WP):], "front") if wrap is not None else "front"
    if name == "front" and frame.get("behindDoc") in ("1", "true"):
        name = "behind"
    side = {v: k for k, v in SIDES.items()}.get(wrap.get("wrapText")) if wrap is not None else None
    return {"wrap": name, "side": side,
            "distances": {side_: _pt(frame.get("dist" + key)) for side_, key in
                          (("top", "T"), ("bottom", "B"), ("left", "L"), ("right", "R"))},
            "z_order": int(frame.get("relativeHeight") or 0), "lock_anchor": frame.get("locked") in ("1", "true"),
            "allow_overlap": frame.get("allowOverlap") not in ("0", "false"),
            "layout_in_cell": frame.get("layoutInCell") not in ("0", "false")}


class DrawingOps:
    """Floating drawings, text boxes and shapes, on :class:`docx_agent.Document`."""

    def _frame(self: "Document", identifier: str) -> tuple[str, Element]:
        from .document import EditError

        for part, found, frame in self._all_drawings():
            if found == identifier:
                return part, frame
        raise EditError(f"no drawing {identifier!r}")

    # -- floating and inline ----------------------------------------------------------------

    def float_drawing(self: "Document", identifier: str, **values) -> "EditResult":
        """Make an inline drawing float, as Word's ``convert to shape`` does: in front of
        the text, against the column and the paragraph where it stood (0, 0 here: where
        the layout put it is not the file's), a step above every other drawing; then
        ``values`` as :meth:`set_drawing` takes them.  Word does not track it."""
        from .document import EditError, EditResult

        part, frame = self._frame(identifier)
        if frame.tag != _WP + "inline":
            raise EditError(f"{identifier} floats already")
        self._check_drawing_values(values)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare_many([])
            part, frame = self._frame(identifier)
            anchor = make("wp:anchor")
            for key, value in (("distT", "0"), ("distB", "0"), ("distL", str(SIDE_DISTANCE)),
                               ("distR", str(SIDE_DISTANCE)), ("simplePos", "0"),
                               ("relativeHeight", str(max(all_heights(self), default=FIRST_HEIGHT - HEIGHT_STEP)
                                                      + HEIGHT_STEP)),
                               ("behindDoc", "0"), ("locked", "0"), ("layoutInCell", "1"), ("allowOverlap", "1")):
                anchor.set(key, value)
            for name in (_ids.ANCHOR_ID, _ids.EDIT_ID):
                if frame.get(name) is not None:
                    anchor.set(name, frame.get(name))
            anchor.append(make("wp:simplePos", x="0", y="0"))
            for tag, origin in (("wp:positionH", "column"), ("wp:positionV", "paragraph")):
                position = make(tag, relativeFrom=origin)
                offset = make("wp:posOffset")
                offset.text = "0"
                position.append(offset)
                anchor.append(position)
            for child in frame:
                if not isinstance(child.tag, str):
                    continue
                if child.tag == _WP + "docPr":
                    if anchor.find(_WP + "effectExtent") is None:
                        anchor.append(make("wp:effectExtent", l="0", t="0", r="0", b="0"))
                    anchor.append(make("wp:wrapNone"))
                anchor.append(copy.deepcopy(child))
            frame.addprevious(anchor)
            remove(frame)
            _ids.ensure_w14(self.package.tree(part), drawings=True)
            if values:
                self._apply_drawing(anchor, values)
            self.package.mark_dirty(part)
        warnings = ["Word does not track floating a drawing: it is applied untracked"] if tracking else []
        return EditResult(identifier, renamed=renames, warnings=warnings)

    def inline_drawing(self: "Document", identifier: str) -> "EditResult":
        """Put a floating drawing in the line again, as Word's ``convert to inline shape``
        does: ``wp:inline``, every distance 0, its ids kept.  Word does not track it."""
        from .document import EditError, EditResult

        part, frame = self._frame(identifier)
        if frame.tag != _WP + "anchor":
            raise EditError(f"{identifier} is inline already")
        if alternate_content(frame) is not None:
            raise EditError(f"{identifier} is a shape or text box: Word puts only pictures in the line")
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare_many([])
            part, frame = self._frame(identifier)
            inline = make("wp:inline", distT="0", distB="0", distL="0", distR="0")
            for name in (_ids.ANCHOR_ID, _ids.EDIT_ID):
                if frame.get(name) is not None:
                    inline.set(name, frame.get(name))
            for tag in ("extent", "effectExtent", "docPr", "cNvGraphicFramePr"):
                node = frame.find(_WP + tag)
                if node is not None:
                    inline.append(copy.deepcopy(node))
            graphic = frame.find(_A + "graphic")
            if graphic is not None:
                inline.append(copy.deepcopy(graphic))
            frame.addprevious(inline)
            remove(frame)
            self.package.mark_dirty(part)
        warnings = ["Word does not track putting a drawing in the line: it is applied untracked"] if tracking else []
        return EditResult(identifier, renamed=renames, warnings=warnings)

    # -- a floating drawing's settings ------------------------------------------------------

    _DRAWING_KEYS = frozenset({"wrap", "side", "distances", "x", "y", "x_align", "y_align", "horizontal_from",
                               "vertical_from", "z_order", "lock_anchor", "allow_overlap", "layout_in_cell",
                               "width", "height"})

    def _check_drawing_values(self: "Document", values: dict) -> None:
        from .document import EditError

        unknown = set(values) - self._DRAWING_KEYS
        if unknown:
            raise EditError(f"unknown drawing settings {sorted(unknown)}")
        if "wrap" in values and values["wrap"] not in WRAPS:
            raise EditError(f"wrap is one of {', '.join(WRAPS)}")
        if "side" in values and values["side"] not in SIDES:
            raise EditError(f"side is one of {', '.join(SIDES)}")
        if values.get("horizontal_from") not in (None, *HORIZONTAL):
            raise EditError(f"horizontal_from is one of {', '.join(HORIZONTAL)}")
        if values.get("vertical_from") not in (None, *VERTICAL):
            raise EditError(f"vertical_from is one of {', '.join(VERTICAL)}")
        if values.get("x_align") not in (None, *X_ALIGN) or values.get("y_align") not in (None, *Y_ALIGN):
            raise EditError("x_align is left, center, right, inside or outside; y_align top, center, bottom, "
                            "inside or outside")
        if "x" in values and values.get("x_align") is not None or "y" in values and values.get("y_align") is not None:
            raise EditError("give an offset or an alignment on each axis, not both")
        if "distances" in values and (not isinstance(values["distances"], dict)
                                      or set(values["distances"]) - {"top", "bottom", "left", "right"}):
            raise EditError("distances is a dict of top, bottom, left and right (points)")
        z = values.get("z_order")
        if z is not None and z not in ("front", "back", "forward", "backward") and not isinstance(z, int):
            raise EditError("z_order is front, back, forward, backward or a relativeHeight")
        for key in ("width", "height"):
            if key in values and not (isinstance(values[key], (int, float)) and 0 < values[key] <= 1584):
                raise EditError("sizes are in points, up to 1584")

    def set_drawing(self: "Document", identifier: str, **values) -> "EditResult":
        """A floating drawing's settings: ``wrap`` (``square``, ``tight``, ``through``,
        ``top_and_bottom``, ``front``/``none``, ``behind``), ``side`` (``both``, ``left``,
        ``right``, ``largest``), ``distances`` (``{top, bottom, left, right}`` points),
        ``x``/``y`` (points) or ``x_align``/``y_align`` against ``horizontal_from`` and
        ``vertical_from``, ``z_order`` (``front``, ``back``, ``forward``, ``backward`` or a
        ``relativeHeight``), ``lock_anchor``, ``allow_overlap``, ``layout_in_cell``, and
        ``width``/``height`` (points; one alone keeps the aspect ratio; an inline drawing's
        size too) -- a group's members scale with it.  Word does not track any of them:
        tracked, they are applied untracked and the result says so."""
        from .document import EditError, EditResult

        part, frame = self._frame(identifier)
        self._check_drawing_values(values)
        if frame.tag == _WP + "inline" and set(values) - {"width", "height"}:
            raise EditError(f"{identifier} is inline: float it first (float_drawing)")
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare_many([])
            part, frame = self._frame(identifier)
            self._apply_drawing(frame, values)
            self.package.mark_dirty(part)
        warnings = ["Word does not track a drawing's place, wrapping, order or size: applied untracked"] \
            if tracking else []
        return EditResult(identifier, renamed=renames, warnings=warnings)

    def move_drawing(self: "Document", identifier: str, x: float, y: float) -> "EditResult":
        """Move a floating drawing to ``(x, y)`` points from where it is measured from."""
        return self.set_drawing(identifier, x=x, y=y)

    def resize_drawing(self: "Document", identifier: str, width: float | None = None,
                       height: float | None = None) -> "EditResult":
        """Resize a drawing to ``width`` and/or ``height`` points (give one to keep its aspect
        ratio): ``doc.resize_drawing("d:3", width=200)``."""
        values = {k: v for k, v in (("width", width), ("height", height)) if v is not None}
        return self.set_drawing(identifier, **values)

    def _apply_drawing(self: "Document", frame: Element, values: dict) -> None:
        if "wrap" in values or "side" in values:
            current = read_wrap(frame)
            kind = values.get("wrap", current["wrap"])
            side = values.get("side", current.get("side") or "both")
            for child in [c for c in frame if isinstance(c.tag, str) and c.tag.startswith(_WP + "wrap")]:
                remove(child)
            tag = WRAPS[kind]
            node = make("wp:" + tag)
            if tag in ("wrapSquare", "wrapTight", "wrapThrough"):
                node.set("wrapText", SIDES[side])
            if tag in ("wrapTight", "wrapThrough"):
                polygon = make("wp:wrapPolygon", edited="0")
                for k, (x, y) in enumerate(POLYGON):
                    polygon.append(make("wp:start" if k == 0 else "wp:lineTo", x=str(x), y=str(y)))
                node.append(polygon)
            insert_in_order(frame, node)
            frame.set("behindDoc", "1" if kind == "behind" else "0")
            self._sync_fallback_wrap(frame, kind, side)
        if "distances" in values:
            for side, key in (("top", "distT"), ("bottom", "distB"), ("left", "distL"), ("right", "distR")):
                if side in values["distances"]:
                    frame.set(key, str(emu(values["distances"][side])))
        for axis, tag, origin_key in (("x", "positionH", "horizontal_from"), ("y", "positionV", "vertical_from")):
            align_key = f"{axis}_align"
            if axis not in values and align_key not in values and origin_key not in values:
                continue
            node = frame.find(_WP + tag)
            if node is None:
                node = make("wp:" + tag, relativeFrom="column" if axis == "x" else "paragraph")
                insert_in_order(frame, node)
            if values.get(origin_key) is not None:
                node.set("relativeFrom", values[origin_key])
            if axis in values or values.get(align_key) is not None:
                for child in list(node):
                    remove(child)
                if values.get(align_key) is not None:
                    child = make("wp:align")
                    child.text = values[align_key]
                else:
                    child = make("wp:posOffset")
                    child.text = str(emu(values[axis]))
                node.append(child)
        if "z_order" in values:
            heights = sorted(h for h in all_heights(self) if h != int(frame.get("relativeHeight") or 0))
            order = values["z_order"]
            mine = int(frame.get("relativeHeight") or FIRST_HEIGHT)
            if order == "front":
                value = max(heights + [mine - HEIGHT_STEP]) + HEIGHT_STEP
            elif order == "back":
                value = min(heights + [mine]) - HEIGHT_STEP - 1
            elif order == "forward":
                above = [h for h in heights if h > mine]
                value = above[0] + 1 if above else mine
            elif order == "backward":
                below = [h for h in heights if h < mine]
                value = below[-1] - 1 if below else mine
            else:
                value = int(order)
            frame.set("relativeHeight", str(max(0, value)))
        if "lock_anchor" in values:
            frame.set("locked", "1" if values["lock_anchor"] else "0")
        if "allow_overlap" in values:
            frame.set("allowOverlap", "1" if values["allow_overlap"] else "0")
        if "layout_in_cell" in values:
            frame.set("layoutInCell", "1" if values["layout_in_cell"] else "0")
        if "width" in values or "height" in values:
            extent = frame.find(_WP + "extent")
            old = (int(extent.get("cx")), int(extent.get("cy")))
            width = emu(values["width"]) if "width" in values else None
            height = emu(values["height"]) if "height" in values else None
            if width is None:
                width = int(round(old[0] * height / old[1])) if old[1] else old[0]
            if height is None:
                height = int(round(old[1] * width / old[0])) if old[0] else old[1]
            extent.set("cx", str(width))
            extent.set("cy", str(height))
            inner = _xfrm_ext(frame)
            if inner is not None:
                inner.set("cx", str(width))
                inner.set("cy", str(height))
            if frame.tag == _WP + "anchor":
                for tag, child, attribute in (("sizeRelH", "pctWidth", "relativeFrom"),
                                              ("sizeRelV", "pctHeight", "relativeFrom")):
                    if frame.find(_WP14 + tag) is None:
                        node = make("wp14:" + tag, relativeFrom="margin")
                        value = make("wp14:" + child)
                        value.text = "0"
                        node.append(value)
                        insert_in_order(frame, node)
        if frame.tag == _WP + "anchor":
            self._sync_fallback_style(frame)

    # -- the VML fallback -------------------------------------------------------------------

    def _sync_fallback_style(self: "Document", frame: Element) -> None:
        """Bring a fallback shape's size, place and order in line with the drawing."""
        shape = fallback_shape(frame)
        if shape is None:
            return
        style = _style(shape)
        extent = frame.find(_WP + "extent")
        style["width"] = _points_text(int(extent.get("cx")) / EMU)
        style["height"] = _points_text(int(extent.get("cy")) / EMU)
        style["z-index"] = frame.get("relativeHeight") or style.get("z-index", "0")
        relative = {"column": "text", "paragraph": "text", "character": "char", "line": "line"}
        for axis, tag, margin in (("horizontal", "positionH", "margin-left"), ("vertical", "positionV", "margin-top")):
            node = frame.find(_WP + tag)
            if node is None:
                continue
            origin = node.get("relativeFrom") or ""
            style[f"mso-position-{axis}-relative"] = relative.get(origin, origin.replace("Margin", "-margin-area")
                                                                  if origin.endswith("Margin") else origin)
            offset, align = node.find(_WP + "posOffset"), node.find(_WP + "align")
            if offset is not None:
                style[margin] = _points_text(int(offset.text) / EMU)
                style[f"mso-position-{axis}"] = "absolute"
            elif align is not None:
                style[f"mso-position-{axis}"] = align.text
        for key, attribute in (("mso-wrap-distance-left", "distL"), ("mso-wrap-distance-right", "distR"),
                               ("mso-wrap-distance-top", "distT"), ("mso-wrap-distance-bottom", "distB")):
            if frame.get(attribute) is not None:
                style[key] = _points_text(int(frame.get(attribute)) / EMU)
        _set_style(shape, style)

    def _sync_fallback_wrap(self: "Document", frame: Element, kind: str, side: str) -> None:
        shape = fallback_shape(frame)
        if shape is None:
            return
        for node in shape.findall("{%s}wrap" % W10):
            remove(node)
        vml = {"square": "square", "tight": "tight", "through": "through", "top_and_bottom": "topAndBottom"}
        if kind in vml:
            node = etree.SubElement(shape, "{%s}wrap" % W10)
            node.set("type", vml[kind])
            if side != "both":
                node.set("side", side)

    # -- text boxes and shapes --------------------------------------------------------------

    def insert_text_box(self: "Document", at, text: str = "", *, width: float = 144, height: float = 72,
                        x: float = 0, y: float = 0, wrap: str = "square", name: str | None = None) -> "EditResult":
        """A text box anchored at a position (a range or its id; its start), as Word's
        Insert Text Box makes one: against the column and the paragraph at ``(x, y)``
        points, ``width`` x ``height``, wrapped square, no fill and a 0.5 pt black line;
        ``text`` its paragraphs (a line each).  The result's ``id`` is the drawing's
        (``d:<id>``); ``created`` lists it and the text box's paragraphs, which the
        paragraph API edits.  Tracked, the box is an inserted run, as Word records one."""
        return self._insert_shape(at, "rect", text, width, height, x, y, wrap, name, text_box=True)

    def insert_shape(self: "Document", at, preset: str = "rect", *, width: float = 72, height: float = 72,
                     x: float = 0, y: float = 0, wrap: str = "front", text: str | None = None,
                     name: str | None = None) -> "EditResult":
        """A DrawingML shape of a preset geometry (``rect``, ``ellipse``, ``roundRect``,
        ``triangle``... -- every ``ST_ShapeType`` name, ooxml-common's table) anchored at a
        position, as Word's ``make new shape`` makes one: the theme's accent fill and line,
        in front of the text unless ``wrap`` says otherwise, at ``(x, y)`` points against the
        column and the paragraph.  ``text`` puts text in it (centred, as Word writes a
        shape's text)."""
        return self._insert_shape(at, preset, text, width, height, x, y, wrap, name, text_box=False)

    def _insert_shape(self: "Document", at, preset: str, text: str | None, width: float, height: float, x: float,
                      y: float, wrap: str, name: str | None, *, text_box: bool) -> "EditResult":
        from ooxml_common.drawingml.presets import PRESETS

        from . import inline as _inline
        from . import text as _text
        from .document import EditError, EditResult
        from .ranges import TextRange

        if preset not in PRESETS:
            raise EditError(f"{preset!r} is not a preset geometry")
        if wrap not in WRAPS:
            raise EditError(f"wrap is one of {', '.join(WRAPS)}")
        for value in (width, height):
            if not (isinstance(value, (int, float)) and 0 < value <= 1584):
                raise EditError("sizes are in points, up to 1584")
        lines = (text or "").split("\n") if (text is not None and (text or text_box)) else None
        for line in lines or []:
            _text.check_text(line)
        where = at if isinstance(at, TextRange) else self.range(at)
        if where.view != "current":
            raise EditError("edits take the current view")
        part, entries = where._entries()
        entry = entries[0]
        if any(True for node in entry.element.iterancestors(_W + "txbxContent")):
            raise EditError("a text box or shape is not anchored in a text box")
        items = _text.atoms(entry.element)
        if not 0 <= where.start <= len(items):
            raise EditError(f"{where.id} is outside its paragraph's text")
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [entry])
            number = self._next_doc_pr()
            used = self._used()
            anchor_id, edit_id = _ids.generate(part + "\0shape", used, 2)
            label = name or (f"Text Box {number}" if text_box else f"Shape {number}")
            paragraphs = []
            if lines is not None:
                for line in lines:
                    paragraph = make("w:p")
                    if not text_box:
                        paragraph.append(make("w:pPr"))
                        paragraph[0].append(make("w:jc", **{"w:val": "center"}))
                    if line:
                        paragraph.append(_text.make_run(line, None))
                    para_id, text_id = _ids.generate(part + "\0box", used, 2)
                    _ids.stamp(paragraph, para_id, text_id)
                    paragraphs.append(paragraph)
            run_children = _shape_xml(number, label, preset, emu(width), emu(height), emu(x), emu(y), wrap,
                                      max(all_heights(self), default=FIRST_HEIGHT - HEIGHT_STEP) + HEIGHT_STEP,
                                      anchor_id, edit_id, paragraphs, text_box)
            run = _inline.new_run(entry.element, where.start, run_children)
            _inline.hoist(run)
            no_proof = make("w:noProof")
            properties = run.find(_W + "rPr")
            if properties is None:
                properties = make("w:rPr")
                run.insert(0, properties)
            insert_in_order(properties, no_proof)
            self._track_inserted(part, [run])
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                choice = run.find(f".//{_MC}Choice//{_W}txbxContent")
                if choice is not None:
                    stamp = Stamp(self, tracking)
                    for paragraph in choice.findall(_W + "p"):
                        _track.insert_paragraph_content(paragraph, stamp, part)
                    stamp.finish()
                    self._copy_to_fallback(run)
            root = self.package.tree(part)
            _ids.ensure_w14(root, drawings=True, extra={"wps": WPS, "wp14": WP14})
            self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
            self._invalidate()
            index = self._index(part)
            made = [f"d:{number}"] + [index.entry_for(p).id for p in paragraphs if index.entry_for(p) is not None]
        return EditResult(f"d:{number}", created=made, renamed=renames)

    def _copy_to_fallback(self: "Document", node: Element) -> None:
        for holder in node.iter(_MC + "AlternateContent"):
            sync_text_box(holder)

    # -- text box stories -------------------------------------------------------------------

    def text_box_paragraphs(self: "Document", identifier: str) -> list:
        """The paragraphs of a text box (``d:<id>``), as :class:`~docx_agent.Paragraph`\\ s."""
        from .document import EditError, Paragraph

        part, frame = self._frame(identifier)
        content = frame.find(f".//{_W}txbxContent")
        if content is None:
            raise EditError(f"{identifier} is not a text box")
        index = self._index(part)
        inside = {id(p) for p in content.iter(_W + "p")}
        return [Paragraph(self, entry.id) for entry in index.paragraphs if id(entry.element) in inside]

    def _text_box_content(self: "Document", identifier: str) -> tuple[str, Element]:
        from .document import EditError

        part, frame = self._frame(identifier)
        content = frame.find(f".//{_W}txbxContent")
        if content is None:
            raise EditError(f"{identifier} is not a text box")
        return part, content


def sync_text_box(holder: Element) -> bool:
    """Make an ``mc:AlternateContent``'s fallback text box hold what its choice's does (the
    same paragraphs, paraIds and all, as Word writes it).  Whether it changed anything."""
    choice = holder.find(f"{_MC}Choice")
    fallback = holder.find(f"{_MC}Fallback")
    if choice is None or fallback is None:
        return False
    source = next(choice.iter(_W + "txbxContent"), None)
    target = next(fallback.iter(_W + "txbxContent"), None)
    if source is None or target is None:
        return False
    if etree.tostring(source) == etree.tostring(target):
        return False
    for child in list(target):
        remove(child)
    for child in source:
        target.append(copy.deepcopy(child))
    return True


def _shape_xml(number: int, name: str, preset: str, cx: int, cy: int, x: int, y: int, wrap: str, height: int,
               anchor_id: str, edit_id: str, paragraphs: list[Element], text_box: bool) -> list[Element]:
    """The run content Word writes for a shape or a text box: ``mc:AlternateContent`` with the
    ``wps`` drawing and, where VML can draw it, its fallback."""
    from xml.sax.saxutils import quoteattr

    wrap_tag = WRAPS[wrap]
    behind = "1" if wrap == "behind" else "0"
    wrap_xml = f"<wp:{wrap_tag}/>" if wrap_tag in ("wrapNone", "wrapTopAndBottom") else (
        f'<wp:{wrap_tag} wrapText="bothSides"/>' if wrap_tag == "wrapSquare" else
        f'<wp:{wrap_tag} wrapText="bothSides"><wp:wrapPolygon edited="0">'
        + "".join(f'<wp:{"start" if k == 0 else "lineTo"} x="{px}" y="{py}"/>' for k, (px, py) in enumerate(POLYGON))
        + f"</wp:wrapPolygon></wp:{wrap_tag}>")
    effect = "0" if text_box else "12700"
    if text_box:
        properties = ('<a:noFill/><a:ln w="6350"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>')
        style = ""
        anchor = "t"
    else:
        properties = ""
        style = ('<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr>'
                 '</a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0">'
                 '<a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/>'
                 "</a:fontRef></wps:style>")
        anchor = "ctr"
    content = ""
    if paragraphs:
        content = "<w:txbxContent>" + "".join(etree.tostring(p, encoding="unicode") for p in paragraphs) \
                  + "</w:txbxContent>"
    txbx = f"<wps:txbx>{content}</wps:txbx>" if content else ""
    body = (f'<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" '
            f'wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" spcCol="0" rtlCol="0" '
            f'fromWordArt="0" anchor="{anchor}" anchorCtr="0" forceAA="0" compatLnSpc="1">'
            '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr>')
    choice = (
        f'<mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="{SIDE_DISTANCE}" '
        f'distR="{SIDE_DISTANCE}" simplePos="0" relativeHeight="{height}" behindDoc="{behind}" locked="0" '
        f'layoutInCell="1" allowOverlap="1" wp14:anchorId="{anchor_id}" wp14:editId="{edit_id}">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="{effect}" t="{effect}" r="{effect}" b="{effect}"/>'
        f'{wrap_xml}<wp:docPr id="{number}" name={quoteattr(name)}/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="{WPS}"><wps:wsp><wps:cNvSpPr{" txBox=" + chr(34) + "1" + chr(34) if text_box else ""}/>'
        f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom>{properties}</wps:spPr>{style}{txbx}{body}'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>")
    fallback = ""
    vml = "shape" if text_box else VML_PRESETS.get(preset)
    if vml is not None:
        css = (f"position:absolute;margin-left:{_points_text(x / EMU)};margin-top:{_points_text(y / EMU)};"
               f"width:{_points_text(cx / EMU)};height:{_points_text(cy / EMU)};z-index:{height};"
               "visibility:visible;mso-wrap-style:square;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;"
               "mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;"
               "mso-position-horizontal-relative:text;mso-position-vertical:absolute;"
               f"mso-position-vertical-relative:text;v-text-anchor:{'top' if text_box else 'middle'}")
        textbox = f"<v:textbox>{content}</v:textbox>" if content else ""
        wrap_vml = {"wrapSquare": "square", "wrapTight": "tight", "wrapThrough": "through",
                    "wrapTopAndBottom": "topAndBottom"}.get(wrap_tag)
        w10 = f'<w10:wrap type="{wrap_vml}"/>' if wrap_vml else ""
        if text_box:
            shape = ('<v:shapetype w14:anchorId="{a}" id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
                     'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" '
                     'o:connecttype="rect"/></v:shapetype>').format(a=anchor_id)
            shape += (f'<v:shape id={quoteattr(name)} o:spid="_x0000_s{1024 + number}" type="#_x0000_t202" '
                      f'style="{css}" filled="f" strokeweight=".5pt"><v:fill o:detectmouseclick="t"/>'
                      f"{textbox}{w10}</v:shape>")
        else:
            shape = (f'<v:{vml} w14:anchorId="{anchor_id}" id={quoteattr(name)} o:spid="_x0000_s{1024 + number}" '
                     f'style="{css}" fillcolor="#4472c4 [3204]" strokecolor="#09101d [484]" strokeweight="1pt">'
                     f"{textbox}{w10}</v:{vml}>")
        fallback = f"<mc:Fallback><w:pict>{shape}</w:pict></mc:Fallback>"
    xml = (f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:w="{W}" xmlns:wp="{WP}" xmlns:wp14="{WP14}" '
           f'xmlns:a="{A}" xmlns:wps="{WPS}" xmlns:v="{V}" xmlns:o="{O}" xmlns:w10="{W10}" xmlns:w14="{W14}">'
           f"{choice}{fallback}</mc:AlternateContent>")
    return [etree.fromstring(xml)]
