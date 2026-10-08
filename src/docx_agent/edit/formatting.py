"""Direct formatting: run and paragraph properties as declared, read and written.

Styles come first (:mod:`docx_agent.edit.styles`); this is the second level, what Word calls
direct formatting.  Reads are **declared** values -- ``None`` when the run or paragraph does
not say, so the value is inherited -- and the effective values come from docx2svg's resolver
(:mod:`docx_agent.edit.effective`).  Writes go to one declared level: a run's ``w:rPr``, a
paragraph's ``w:pPr``, a paragraph mark's ``w:pPr/w:rPr``, or a style's.  ``None`` removes
the declaration (inherit again).

Run properties (keyword: value):

* ``bold``, ``italic``, ``strike``, ``double_strike``, ``caps``, ``small_caps`` -- ``True``,
  ``False`` (written as Word writes an explicit off, ``w:val="0"``) or ``None``; bold and
  italic write their complex-script twins (``w:bCs``, ``w:iCs``) as Word does;
* ``underline`` -- ``True`` (single), ``False`` (none) or an ``ST_Underline`` value;
* ``size`` -- points (``w:sz`` and ``w:szCs``, in half points);
* ``font`` -- a face for the ASCII, high-ANSI and complex-script slots (``w:rFonts``), the
  slots' theme references removed;
* ``color`` -- ``RRGGBB``, ``auto``, or a theme colour (``accent1``, ``text1``...) written as
  ``w:themeColor`` with the theme's value in ``w:val``, as Word writes it;
* ``highlight`` -- an ``ST_HighlightColor`` name (``yellow``...);
* ``vertical_align`` -- ``superscript``, ``subscript`` or ``baseline``.

Paragraph properties: ``alignment`` (``left``, ``center``, ``right``, ``justify``,
``distribute``), ``indent_left``, ``indent_right``, ``first_line``, ``hanging`` (points;
a first-line indent and a hanging one exclude each other), ``space_before``,
``space_after`` (points), ``line_spacing`` (a multiple, ``1.5``; or ``("exact", pt)`` /
``("at_least", pt)``), ``keep_with_next``, ``keep_together``, ``page_break_before``,
``widow_control`` (``True``/``False``/``None``).
"""

from __future__ import annotations

import re
from typing import Callable

from ..oxml.xml import Element, append_in_order, make, qn, remove

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


class FormattingError(ValueError):
    """A formatting value Word would not accept, refused before anything changed."""


TOGGLES = {
    "bold": ("w:b", "w:bCs"),
    "italic": ("w:i", "w:iCs"),
    "strike": ("w:strike",),
    "double_strike": ("w:dstrike",),
    "caps": ("w:caps",),
    "small_caps": ("w:smallCaps",),
}
RUN_PROPERTIES = frozenset(TOGGLES) | {"underline", "size", "font", "color", "highlight", "vertical_align"}

UNDERLINES = frozenset({
    "single", "words", "double", "thick", "dotted", "dottedHeavy", "dash", "dashedHeavy",
    "dashLong", "dashLongHeavy", "dotDash", "dashDotHeavy", "dotDotDash", "dashDotDotHeavy",
    "wave", "wavyHeavy", "wavyDouble", "none",
})
HIGHLIGHTS = frozenset({
    "black", "blue", "cyan", "green", "magenta", "red", "yellow", "white", "darkBlue",
    "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray",
    "none",
})
THEME_COLORS = frozenset({
    "dark1", "light1", "dark2", "light2", "accent1", "accent2", "accent3", "accent4",
    "accent5", "accent6", "hyperlink", "followedHyperlink", "background1", "text1",
    "background2", "text2",
})
VERTICAL_ALIGNS = frozenset({"superscript", "subscript", "baseline"})
_HEX = re.compile(r"[0-9A-Fa-f]{6}")

ALIGNMENTS = {"left": "left", "center": "center", "right": "right", "justify": "both",
              "distribute": "distribute"}
_ALIGNMENT_NAMES = {"left": "left", "start": "left", "center": "center", "right": "right",
                    "end": "right", "both": "justify", "distribute": "distribute"}
PARAGRAPH_TOGGLES = {
    "keep_with_next": "w:keepNext",
    "keep_together": "w:keepLines",
    "page_break_before": "w:pageBreakBefore",
    "widow_control": "w:widowControl",
}
PARAGRAPH_PROPERTIES = frozenset(PARAGRAPH_TOGGLES) | {
    "alignment", "indent_left", "indent_right", "first_line", "hanging", "space_before",
    "space_after", "line_spacing",
}

#: What clearing direct formatting keeps of a run's properties: the character style, and
#: what is not formatting (language, proofing, script, equation, revision record).
KEPT_RUN = frozenset(_W + name for name in (
    "rStyle", "lang", "noProof", "rtl", "cs", "oMath", "rPrChange", "ins", "del", "moveFrom", "moveTo"))
#: ... and of a paragraph's: its style, list membership, section break, frame, revision
#: record and the table-style and HTML bookkeeping Word maintains.
KEPT_PARAGRAPH = frozenset(_W + name for name in (
    "pStyle", "numPr", "sectPr", "pPrChange", "cnfStyle", "divId", "framePr", "rPr"))


# -- validation ------------------------------------------------------------------------------


def check_run(values: dict) -> None:
    for name, value in values.items():
        if name not in RUN_PROPERTIES:
            raise FormattingError(f"no run property {name!r}; expected one of {', '.join(sorted(RUN_PROPERTIES))}")
        if value is None:
            continue
        if name in TOGGLES and not isinstance(value, bool):
            raise FormattingError(f"{name} is True, False or None")
        if name == "underline" and not (isinstance(value, bool) or value in UNDERLINES):
            raise FormattingError(f"underline is True, False, None or one of {', '.join(sorted(UNDERLINES))}")
        if name == "size" and not (isinstance(value, (int, float)) and not isinstance(value, bool)
                                   and 1 <= value <= 1638):
            raise FormattingError("size is in points, 1 to 1638")
        if name == "font" and not (isinstance(value, str) and value.strip() and len(value) <= 31):
            raise FormattingError("font is a face name of at most 31 characters")
        if name == "color" and not (isinstance(value, str) and (_HEX.fullmatch(value) or value == "auto"
                                                                or value in THEME_COLORS)):
            raise FormattingError(f"color is RRGGBB, 'auto' or a theme colour ({', '.join(sorted(THEME_COLORS))})")
        if name == "highlight" and value not in HIGHLIGHTS:
            raise FormattingError(f"highlight is one of {', '.join(sorted(HIGHLIGHTS))}")
        if name == "vertical_align" and value not in VERTICAL_ALIGNS:
            raise FormattingError("vertical_align is superscript, subscript or baseline")


def check_paragraph(values: dict) -> None:
    for name, value in values.items():
        if name not in PARAGRAPH_PROPERTIES:
            raise FormattingError(f"no paragraph property {name!r}; expected one of "
                                  f"{', '.join(sorted(PARAGRAPH_PROPERTIES))}")
        if value is None:
            continue
        if name in PARAGRAPH_TOGGLES and not isinstance(value, bool):
            raise FormattingError(f"{name} is True, False or None")
        if name == "alignment" and value not in ALIGNMENTS:
            raise FormattingError(f"alignment is one of {', '.join(ALIGNMENTS)}")
        if name in ("indent_left", "indent_right") and not _number(value, -1584, 1584):
            raise FormattingError(f"{name} is in points, -1584 to 1584")
        if name in ("first_line", "hanging", "space_before", "space_after") and not _number(value, 0, 1584):
            raise FormattingError(f"{name} is in points, 0 to 1584")
        if name == "line_spacing":
            if isinstance(value, tuple):
                if len(value) != 2 or value[0] not in ("exact", "at_least") or not _number(value[1], 0.05, 1584):
                    raise FormattingError("line_spacing is a multiple, or ('exact' | 'at_least', points)")
            elif not _number(value, 0.06, 132):
                raise FormattingError("line_spacing is a multiple (1.0, 1.5...), or ('exact' | 'at_least', points)")
    if values.get("first_line") is not None and values.get("hanging") is not None:
        raise FormattingError("a paragraph has a first-line indent or a hanging one, not both")


def _number(value, low, high) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and low <= value <= high


# -- reading ---------------------------------------------------------------------------------


def _on(node: Element | None) -> bool | None:
    if node is None:
        return None
    return (node.get(_W + "val") or "true").lower() not in ("0", "false", "off")


def read_run(properties: Element | None, name: str):
    """A run property as ``properties`` (a ``w:rPr``) declares it, or ``None``."""
    if properties is None:
        return None
    if name in TOGGLES:
        return _on(properties.find(qn(TOGGLES[name][0])))
    if name == "underline":
        node = properties.find(_W + "u")
        if node is None:
            return None
        value = node.get(_W + "val") or "single"
        return False if value == "none" else (True if value == "single" else value)
    if name == "size":
        node = properties.find(_W + "sz")
        try:
            return int(node.get(_W + "val")) / 2 if node is not None else None
        except (TypeError, ValueError):
            return None
    if name == "font":
        node = properties.find(_W + "rFonts")
        return node.get(_W + "ascii") or node.get(_W + "hAnsi") if node is not None else None
    if name == "color":
        node = properties.find(_W + "color")
        if node is None:
            return None
        return node.get(_W + "themeColor") or node.get(_W + "val")
    if name == "highlight":
        node = properties.find(_W + "highlight")
        return node.get(_W + "val") if node is not None else None
    if name == "vertical_align":
        node = properties.find(_W + "vertAlign")
        return node.get(_W + "val") if node is not None else None
    raise KeyError(name)


def read_paragraph(properties: Element | None, name: str):
    """A paragraph property as ``properties`` (a ``w:pPr``) declares it, or ``None``."""
    if properties is None:
        return None
    if name in PARAGRAPH_TOGGLES:
        return _on(properties.find(qn(PARAGRAPH_TOGGLES[name])))
    if name == "alignment":
        node = properties.find(_W + "jc")
        return _ALIGNMENT_NAMES.get(node.get(_W + "val"), node.get(_W + "val")) if node is not None else None
    if name in ("indent_left", "indent_right", "first_line", "hanging"):
        node = properties.find(_W + "ind")
        if node is None:
            return None
        names = {"indent_left": ("left", "start"), "indent_right": ("right", "end"),
                 "first_line": ("firstLine",), "hanging": ("hanging",)}[name]
        for attribute in names:
            if node.get(_W + attribute) is not None:
                return _points(node.get(_W + attribute))
        return None
    if name in ("space_before", "space_after"):
        node = properties.find(_W + "spacing")
        value = node.get(_W + name[6:]) if node is not None else None
        return _points(value) if value is not None else None
    if name == "line_spacing":
        node = properties.find(_W + "spacing")
        if node is None or node.get(_W + "line") is None:
            return None
        rule = node.get(_W + "lineRule") or "auto"
        line = int(node.get(_W + "line"))
        if rule == "auto":
            return line / 240
        return ("exact" if rule == "exact" else "at_least", line / 20)
    raise KeyError(name)


def _points(twips: str | None) -> float | None:
    try:
        return int(twips) / 20
    except (TypeError, ValueError):
        return None


# -- writing ---------------------------------------------------------------------------------


def properties_of(owner: Element, tag: str, create: bool) -> Element | None:
    """``owner``'s ``w:rPr`` or ``w:pPr`` (``tag``), created in schema order if asked."""
    node = owner.find(qn(tag))
    if node is None and create:
        node = append_in_order(owner, make(tag))
    return node


def _set_child(parent: Element, tag: str, attributes: dict[str, str] | None) -> Element | None:
    """Replace ``parent``'s ``tag`` child with one carrying ``attributes``; ``None``
    removes it.  An existing child keeps its place and the attributes it is not given."""
    node = parent.find(qn(tag))
    if attributes is None:
        if node is not None:
            remove(node)
        return None
    if node is None:
        node = append_in_order(parent, make(tag))
    for name, value in attributes.items():
        if value is None:
            node.attrib.pop(qn(name), None)
        else:
            node.set(qn(name), value)
    return node


def write_run(owner: Element, values: dict, theme_color: Callable[[str], str | None] | None = None) -> None:
    """Write run properties on ``owner`` (a ``w:r``, a ``w:style``, or a ``w:pPr`` for its
    mark), creating its ``w:rPr`` as needed and removing it if left empty (on a run)."""
    check_run(values)
    properties = properties_of(owner, "w:rPr", create=any(v is not None for v in values.values()))
    if properties is None:
        return
    for name, value in values.items():
        if name in TOGGLES:
            for tag in TOGGLES[name]:
                if value is None:
                    _set_child(properties, tag, None)
                else:
                    node = _set_child(properties, tag, {})
                    if value:
                        node.attrib.pop(qn("w:val"), None)
                    else:
                        node.set(qn("w:val"), "0")
        elif name == "underline":
            if value is None:
                _set_child(properties, "w:u", None)
            else:
                _set_child(properties, "w:u", {"w:val": "single" if value is True else
                                               ("none" if value is False else value)})
        elif name == "size":
            half_points = None if value is None else str(int(round(value * 2)))
            _set_child(properties, "w:sz", None if value is None else {"w:val": half_points})
            _set_child(properties, "w:szCs", None if value is None else {"w:val": half_points})
        elif name == "font":
            node = properties.find(_W + "rFonts")
            if value is None:
                if node is not None:
                    for attribute in ("ascii", "hAnsi", "cs", "asciiTheme", "hAnsiTheme", "cstheme"):
                        node.attrib.pop(_W + attribute, None)
                    if not node.attrib:
                        remove(node)
            else:
                _set_child(properties, "w:rFonts", {"w:ascii": value, "w:hAnsi": value, "w:cs": value,
                                                    "w:asciiTheme": None, "w:hAnsiTheme": None,
                                                    "w:cstheme": None})
        elif name == "color":
            if value is None:
                _set_child(properties, "w:color", None)
            elif value in THEME_COLORS:
                rgb = (theme_color(value) if theme_color else None) or "000000"
                _set_child(properties, "w:color", {"w:val": rgb, "w:themeColor": value,
                                                   "w:themeTint": None, "w:themeShade": None})
            else:
                _set_child(properties, "w:color", {"w:val": value.upper() if value != "auto" else value,
                                                   "w:themeColor": None, "w:themeTint": None,
                                                   "w:themeShade": None})
        elif name == "highlight":
            _set_child(properties, "w:highlight", None if value is None else {"w:val": value})
        elif name == "vertical_align":
            _set_child(properties, "w:vertAlign", None if value is None else {"w:val": value})
    if not len(properties):
        remove(properties)


def write_paragraph(paragraph_properties_owner: Element, values: dict) -> None:
    """Write paragraph properties on a ``w:p`` (its ``w:pPr``) or a ``w:style``."""
    check_paragraph(values)
    owner = paragraph_properties_owner
    properties = properties_of(owner, "w:pPr", create=any(v is not None for v in values.values()))
    if properties is None:
        return
    for name, value in values.items():
        if name in PARAGRAPH_TOGGLES:
            tag = PARAGRAPH_TOGGLES[name]
            if value is None:
                _set_child(properties, tag, None)
            else:
                node = _set_child(properties, tag, {})
                if value:
                    node.attrib.pop(qn("w:val"), None)
                else:
                    node.set(qn("w:val"), "0")
        elif name == "alignment":
            _set_child(properties, "w:jc", None if value is None else {"w:val": ALIGNMENTS[value]})
        elif name in ("indent_left", "indent_right", "first_line", "hanging"):
            attribute = {"indent_left": "w:left", "indent_right": "w:right",
                         "first_line": "w:firstLine", "hanging": "w:hanging"}[name]
            twips = None if value is None else str(int(round(value * 20)))
            node = properties.find(_W + "ind")
            if node is None and twips is None:
                continue
            changes = {attribute: twips}
            if name == "indent_left":
                changes["w:start"] = None
            elif name == "indent_right":
                changes["w:end"] = None
            elif name == "first_line" and twips is not None:
                changes["w:hanging"] = None
            elif name == "hanging" and twips is not None:
                changes["w:firstLine"] = None
            node = _set_child(properties, "w:ind", changes)
            if not node.attrib:
                remove(node)
        elif name in ("space_before", "space_after"):
            twips = None if value is None else str(int(round(value * 20)))
            node = properties.find(_W + "spacing")
            if node is None and twips is None:
                continue
            side = name[6:]
            node = _set_child(properties, "w:spacing", {f"w:{side}": twips, f"w:{side}Autospacing": None})
            if not node.attrib:
                remove(node)
        elif name == "line_spacing":
            node = properties.find(_W + "spacing")
            if value is None:
                if node is not None:
                    node.attrib.pop(_W + "line", None)
                    node.attrib.pop(_W + "lineRule", None)
                    if not node.attrib:
                        remove(node)
                continue
            if isinstance(value, tuple):
                line, rule = str(int(round(value[1] * 20))), ("exact" if value[0] == "exact" else "atLeast")
            else:
                line, rule = str(int(round(value * 240))), "auto"
            _set_child(properties, "w:spacing", {"w:line": line, "w:lineRule": rule})
    if owner.tag == _W + "p" and not len(properties):
        remove(properties)


def clear_run(run: Element) -> bool:
    """Remove a run's direct formatting (:data:`KEPT_RUN` stays); returns whether it had any."""
    properties = run.find(_W + "rPr")
    if properties is None:
        return False
    changed = False
    for child in list(properties):
        if isinstance(child.tag, str) and child.tag not in KEPT_RUN:
            remove(child)
            changed = True
    if not len(properties):
        remove(properties)
    return changed


def clear_paragraph(paragraph: Element) -> bool:
    """Remove a paragraph's direct paragraph formatting and its mark's (:data:`KEPT_PARAGRAPH`
    and the mark's :data:`KEPT_RUN` stay)."""
    properties = paragraph.find(_W + "pPr")
    if properties is None:
        return False
    changed = False
    for child in list(properties):
        if isinstance(child.tag, str) and child.tag not in KEPT_PARAGRAPH:
            remove(child)
            changed = True
    mark = properties.find(_W + "rPr")
    if mark is not None:
        for child in list(mark):
            if isinstance(child.tag, str) and child.tag not in KEPT_RUN:
                remove(child)
                changed = True
        if not len(mark):
            remove(mark)
    if not len(properties):
        remove(properties)
    return changed
