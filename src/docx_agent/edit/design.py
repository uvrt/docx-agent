"""Design facts of a document, for analysing a template: measured, with ids, no verdicts.

:meth:`Document.design_facts <docx_agent.Document.design_facts>` reports what a document's
design is made of, as Word draws it -- the counterpart of pptx-agent's slide design facts
for a Word document as a whole.  What the facts mean (a house style to follow, a template
to rebuild) is the application's to say.

Each kind (:data:`KINDS`):

* ``styles`` -- every paragraph style in use, **per role**, with its usage count and its
  look after inheritance (docx2svg's resolver: ``w:docDefaults``, the ``w:basedOn`` chain,
  the table style for text in cells, toggles as Word combines them; no direct formatting).
  A paragraph's role is decided in this order: the story it is in (``header``, ``footer``,
  ``footnote``, ``endnote``); in a table cell, ``table``; a heading (``heading 1`` to
  ``heading 9``, as ``word_read`` reads ``#``); the ``Title`` and ``Subtitle`` styles;
  ``caption`` (the Caption style or one based on it); ``toc`` (the TOC styles); ``list``
  (numbered or bulleted); else ``body``.  Comments are not counted.  The same style in two
  roles is two entries: in a cell the table style is part of its look.  ``empty`` counts
  the paragraphs of an entry with no text and no drawing (spacers, section breaks).
* ``table_styles`` -- every table style in use, with its look as the tables using it draw
  it: a sample of five rows and four columns under the tables' most common style options
  (``w:tblLook``) and the paragraph style their cells hold most, each place -- the header
  row, the first row band, the second, the label (first) column, the last row and column
  when the options turn them on -- with its fill, its four borders after the conflicts
  between neighbours, and its text, all from docx2svg's table resolver and cascade, the
  conditional formats it applies there listed.
* ``sections`` -- the page grid of each section: size, orientation, margins, text columns,
  where it starts, and the header and footer each variant shows (``first`` only where the
  section has a title page, ``even`` only where the document has even and odd headers).
* ``theme`` -- the theme's name, colour palette (``dk1`` ... ``folHlink``) and fonts
  (major and minor; East Asian and complex-script faces where it names them).
* ``declared_styles`` (only when asked for) -- every paragraph, character and table style
  the document declares, used or not (``used``), with its ``basedOn`` chain and its look
  after inheritance -- what differs from the default paragraph style's, which is given
  whole -- and for a table style the conditional formats it declares and each place's fill,
  borders and text (:func:`_declared_styles`).
* ``patterns`` -- repeated block shapes under headings, as samples to copy.  The **shape**
  of a heading is the blocks after it up to the next heading of any level (an empty
  paragraph is never a heading): a paragraph is its style's name, consecutive paragraphs of
  one style counting once and empty ones not at all; a table is ``table <style>, <n> columns``; a content control ``control``.  A
  pattern is one shape under two or more headings of one level, the shape a table or two
  entries or more; each lists how many headings of that level have it, of how many, and a
  few examples (heading id and the span of blocks).

**Colours** are what Word draws: where the element that gives a colour, a shading's fill or
a border names a theme colour (``w:themeColor``, ``w:themeFill``, with a tint or shade),
the theme's colour (:func:`theme_rgb`, measured) -- Word does not draw the ``w:val`` beside
it -- and ``theme_color`` / ``theme_fill`` (or the border's text) says which.

Long lists are capped (:data:`CAPS`); ``omitted`` counts what was left out.
"""

from __future__ import annotations

from collections import Counter
from typing import TYPE_CHECKING, Any

from docx2svg import model as _model
from docx2svg.parse.document import cell_conditions
from docx2svg.paths import resolve_path
from docx2svg.table import edge_winner, resolve_table, table_style_id

from . import effective as _effective

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

#: The facts reported by default, in the order they are reported.
KINDS = ("styles", "table_styles", "sections", "theme", "patterns")

#: Every kind there is: :data:`KINDS` and ``declared_styles``, reported only when asked for.
ALL_KINDS = (*KINDS, "declared_styles")

#: How many entries of each list are reported; ``omitted`` counts the rest.
CAPS = {"styles": 60, "table_styles": 12, "sections": 20, "patterns": 12, "declared_styles": 300}

#: Roles in the order they are listed.
ROLES = ("title", "subtitle", *(f"heading {n}" for n in range(1, 10)), "body", "list", "caption",
         "table", "toc", "header", "footer", "footnote", "endnote")

#: The sample a table style's look is read from: rows and columns.
SAMPLE_ROWS, SAMPLE_COLUMNS = 5, 4

#: The colour slots of a theme, in Word's order.
THEME_SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
               "hlink", "folHlink")

#: ECMA-376's default style options, which docx2svg applies to a table without ``w:tblLook``.
_DEFAULT_LOOK = {"firstRow": True, "lastRow": False, "firstColumn": True, "lastColumn": False,
                 "noHBand": False, "noVBand": True}


def design_facts(document: "Document", include=None, caps: dict | None = None) -> dict[str, Any]:
    """The facts of ``include`` (default all of :data:`KINDS`; ``declared_styles`` only when
    named) as JSON-ready data."""
    kinds = [k for k in (KINDS if include is None else ALL_KINDS) if include is None or k in include]
    unknown = [k for k in (include or ()) if k not in ALL_KINDS]
    if unknown:
        raise ValueError(f"no design facts {', '.join(map(repr, unknown))}: the kinds are {', '.join(ALL_KINDS)}")
    limits = {**CAPS, **(caps or {})}
    out: dict[str, Any] = {}
    omitted: dict[str, int] = {}
    builders = {"styles": _styles, "table_styles": _table_styles, "sections": _sections,
                "patterns": _patterns, "declared_styles": _declared_styles}
    for kind in kinds:
        if kind == "theme":
            out["theme"] = _theme(document)
            continue
        items = builders[kind](document)
        limit = limits[kind]
        if len(items) > limit:
            omitted[kind] = len(items) - limit
        out[kind] = items[:limit]
    if "sections" in out:
        out["even_and_odd_headers"] = document.even_and_odd_headers
    if omitted:
        out["omitted"] = omitted
    return out


# -- values ------------------------------------------------------------------------------------


def _pt(value: float | None) -> float | None:
    if value is None:
        return None
    value = round(value, 2)
    return int(value) if value == int(value) else value


def _border(border) -> str:
    """``single 0.5pt 000000``, ``double 1.5pt auto``, or ``none``."""
    if not border or border[0] in (None, "nil", "none"):
        return "none"
    return f"{border[0]} {_pt(border[1] / 8)}pt {border[3] or 'auto'}"


def _fill(shading) -> str | None:
    if not shading or shading[0] == "nil":
        return None
    fill = shading[2]
    return None if fill in (None, "auto") else fill.upper()


def _text_look(document: "Document", run: Any) -> dict[str, Any]:
    out: dict[str, Any] = {"font": run.font, "size": _pt(run.size)}
    out.update(_color(document, run._resolved, "color", run.color))
    for name in ("bold", "italic", "caps", "small_caps"):
        if getattr(run, name):
            out[name] = True
    if run.underline:
        out["underline"] = run.underline
    return out


def _look(document: "Document", paragraph: Any, run: Any) -> dict[str, Any]:
    out = _text_look(document, run)
    out["alignment"] = paragraph.alignment
    out["space_before"] = _pt(paragraph.space_before)
    out["space_after"] = _pt(paragraph.space_after)
    if paragraph.line_spacing is not None:
        out["line_spacing"] = (_pt(paragraph.line_spacing) if paragraph.line_rule == "auto"
                               else f"{_pt(paragraph.line_spacing)}pt {paragraph.line_rule}")
    out["indent_left"] = _pt(paragraph.indent_left)
    if paragraph.indent_right:
        out["indent_right"] = _pt(paragraph.indent_right)
    first = paragraph.first_line - paragraph.hanging
    if first:
        out["indent_first"] = _pt(first)
    for name in ("keep_with_next", "keep_together", "page_break_before"):
        if getattr(paragraph, name):
            out[name] = True
    if paragraph.outline_level is not None and paragraph.outline_level < 9:
        out["outline_level"] = paragraph.outline_level + 1
    if paragraph.numbering is not None:
        out["numbering"] = _numbering(document, *paragraph.numbering)
    resolved = paragraph._resolved
    shading = resolved.get("shd")
    if _fill(shading) is not None:
        out.update(_color(document, resolved, "shd", _fill(shading), name="shading"))
    borders = {}
    for side in ("top", "left", "bottom", "right", "between"):
        border = resolved.get(f"pBdr.{side}")
        if border and border[0] not in (None, "nil", "none"):
            themed = _color(document, resolved, f"pBdr.{side}", border[3] or "auto")
            borders[side] = _border((border[0], border[1], border[2], themed["color"])) + (
                f" ({themed['theme_color']})" if "theme_color" in themed else "")
    if borders:
        out["borders"] = borders
    return out


# -- theme colours -----------------------------------------------------------------------------

#: The attributes naming a theme colour: on ``w:color`` and a border, and on ``w:shd`` (its fill).
_THEME_ATTRIBUTES = {"color": ("themeColor", "themeTint", "themeShade"),
                     "shd": ("themeFill", "themeFillTint", "themeFillShade")}


def theme_rgb(document: "Document", name: str, tint: str | None = None, shade: str | None = None) -> str | None:
    """A Word theme colour (``w:themeColor``'s ``accent1``, ``text1``, ``background2``...;
    the text and background names through the settings' ``w:clrSchemeMapping``) with its
    ``w:themeTint`` / ``w:themeShade`` (hex bytes), as ``RRGGBB``; ``None`` for a name the
    theme does not define.

    Measured (Word 16, text and shading alike; ``feedback-round21/measurements-B.md`` and
    ``measurements-D.md``): **Word draws the theme colour, not the ``w:val`` / ``w:fill``
    written beside it**, its tint or shade applied by
    :func:`ooxml_common.themetint.word_theme_tint` -- Windows' integer HLS on a scale of
    1530, every swatch measured to the level; with both, the tint alone counts."""
    from ooxml_common.themetint import word_theme_tint

    from .formatops import _SCHEME

    if name not in _SCHEME:
        return None
    base = document._theme_color(name)
    if base is None:
        return None
    return word_theme_tint(base, tint, shade)


def _declaring(document: "Document", resolved: Any, key: str) -> Any:
    """The element that gave ``resolved``'s ``key`` (a ``w:color``, ``w:shd`` or border side),
    from docx2svg's account of where it came from; ``None`` where it cannot be found."""
    origin = resolved.origin(key) if hasattr(resolved, "origin") else None
    if origin is None:
        return None
    root = document.styles._root()
    if root is None:
        return None
    paragraph_key = key == "shd" or key.startswith("pBdr.")
    local = key.split(".")[-1] if key.startswith("pBdr.") else key
    if origin.level == "docDefaults":
        holder = root.find(f"{_W}docDefaults/{_W}{'pPrDefault' if paragraph_key else 'rPrDefault'}")
    elif origin.style_id or origin.declared_in:
        holder = document.styles._node(origin.declared_in or origin.style_id)
        if holder is not None and origin.detail:
            holder = next((c for c in holder.findall(_W + "tblStylePr") if c.get(_W + "type") == origin.detail),
                          None)
    else:
        return None
    if holder is None:
        return None
    path = f"{_W}pPr/{_W}pBdr/{_W}{local}" if key.startswith("pBdr.") else (
        f"{_W}pPr/{_W}shd" if key == "shd" else f"{_W}rPr/{_W}{key}")
    return holder.find(path)


def _themed_element(document: "Document", node: Any, attributes: tuple[str, str, str]) -> tuple[str, str] | None:
    """``(RRGGBB, "accent1 shade BF")`` for an element naming a theme colour, else ``None``."""
    if node is None or not node.get(_W + attributes[0]):
        return None
    name, tint, shade = (node.get(_W + a) for a in attributes)
    rgb = theme_rgb(document, name, tint, shade)
    if rgb is None:
        return None
    label = " ".join([name] + [f"{kind} {value.upper()}" for kind, value in (("tint", tint), ("shade", shade))
                               if value])
    return rgb, label


def _color(document: "Document", resolved: Any, key: str, value: str | None, name: str = "color") -> dict[str, Any]:
    """``{name: RRGGBB}`` -- the theme's colour where the declaring element names one, with
    ``theme_<name>`` saying which -- else ``{name: value}``."""
    themed = _themed_element(document, _declaring(document, resolved, key),
                             _THEME_ATTRIBUTES["shd" if key == "shd" else "color"])
    if themed is None:
        return {name: value}
    return {name: themed[0], f"theme_{name}": themed[1]}


def _numbering(document: "Document", num_id: int, level: int) -> dict[str, Any]:
    found = _effective.model(document).model.numbering.get(num_id, {}).get(level)
    out: dict[str, Any] = {"num_id": num_id, "level": level}
    if found is not None:
        out["format"] = found.format or "decimal"
        if found.text is not None:
            out["label"] = found.text
    return out


# -- styles ------------------------------------------------------------------------------------


def _ancestor(node: Any, tag: str) -> Any:
    node = node.getparent()
    while node is not None and node.tag != tag:
        node = node.getparent()
    return node


class _Names:
    """Paragraph style ids -> the style applied (the default for none or an unknown one)."""

    def __init__(self, document: "Document") -> None:
        self.styles = {s.id: s for s in document.styles if s.kind == "paragraph"}
        self.by_name = {(s.name or "").lower(): s for s in self.styles.values()}
        self.default = document.styles.default("paragraph")

    def applied(self, style_id: str | None) -> Any:
        return self.styles.get(style_id) or self.default if style_id else self.default

    def chain(self, style: Any) -> list[str]:
        names, seen = [], set()
        while style is not None and style.id not in seen:
            seen.add(style.id)
            names.append((style.name or "").lower())
            style = self.styles.get(style.based_on) if style.based_on else None
        return names


def _style_id(paragraph: Any) -> str | None:
    properties = paragraph.find(_W + "pPr")
    style = properties.find(_W + "pStyle") if properties is not None else None
    return style.get(_W + "val") if style is not None else None


def _empty(node: Any) -> bool:
    """Whether a paragraph holds no text and no drawing: a spacer, or a section break's."""
    return (not "".join(t.text or "" for t in node.iter(_W + "t")).strip()
            and next(node.iter(_W + "drawing", _W + "pict", _W + "object"), None) is None)


def _role(document: "Document", story: str, node: Any, style: Any, names: _Names, levels: Any) -> str | None:
    if story.startswith("header"):
        return "header"
    if story.startswith("footer"):
        return "footer"
    if story in ("footnotes", "endnotes"):
        return story[:-1]
    if story != "body":
        return None                       # comments
    if _ancestor(node, _W + "tbl") is not None:
        return "table"
    level = levels.of(node)
    if level is not None:
        return f"heading {level}"
    chain = names.chain(style)
    if chain and chain[0] in ("title", "subtitle"):
        return chain[0]
    if "caption" in chain:
        return "caption"
    if chain and (chain[0].startswith("toc ") or chain[0] == "toc heading"):
        return "toc"
    properties = node.find(_W + "pPr")
    direct = properties.find(_W + "numPr") if properties is not None else None
    if direct is not None:
        num = direct.find(_W + "numId")
        return "list" if num is not None and num.get(_W + "val") not in (None, "0") else "body"
    return "list" if _effective.paragraph(document, node).numbering is not None else "body"


def _styles(document: "Document") -> list[dict[str, Any]]:
    from .document import _HeadingLevels

    names = _Names(document)
    levels = _HeadingLevels(document)
    found: dict[tuple[str, str | None], dict[str, Any]] = {}
    for part in document._parts():
        story = document._story_of(part)
        for entry in document._index(part).paragraphs:
            node = entry.element
            style = names.applied(_style_id(node))
            role = _role(document, story, node, style, names, levels)
            if role is None:
                continue
            key = (role, style.id if style is not None else None)
            item = found.get(key)
            if item is None:
                item = found[key] = {"count": 0, "example": entry.id, "tables": Counter()}
            item["count"] += 1
            if _empty(node):
                item["empty"] = item.get("empty", 0) + 1
            if role == "table":
                table = _ancestor(node, _W + "tbl")
                declared = table.find(f"{_W}tblPr/{_W}tblStyle")
                item["tables"][declared.get(_W + "val") if declared is not None else ""] += 1
    out = []
    for (role, style_id), item in found.items():
        style = names.styles.get(style_id)
        table_style = item["tables"].most_common(1)[0][0] if item["tables"] else None
        paragraph, run = _effective.style_look(document, style_id, table_style=table_style,
                                               conditions=("wholeTable",) if table_style is not None else ())
        entry: dict[str, Any] = {"role": role, "style": style.name if style is not None else None,
                                 "paragraphs": item["count"], "example": item["example"]}
        if item.get("empty"):
            entry["empty"] = item["empty"]
        if table_style is not None:
            entry["table_style"] = _table_style_name(document, table_style)
        entry["look"] = _look(document, paragraph, run)
        out.append(entry)
    out.sort(key=lambda e: (ROLES.index(e["role"]), -e["paragraphs"], e["style"] or ""))
    return out


def _table_style_name(document: "Document", style_id: str) -> str | None:
    """The name of the table style a table naming ``style_id`` gets (``""``: none named)."""
    sheet = _effective.model(document).model
    applied = table_style_id(sheet, _model.Table(style_id=style_id or None))
    style = sheet.styles.styles.get(applied) if applied and sheet.styles is not None else None
    return (style.name or style.style_id) if style is not None else None


# -- table styles ------------------------------------------------------------------------------


def _model_tables(blocks, out: list) -> None:
    for block in blocks or ():
        if isinstance(block, _model.Table):
            out.append(block)
            for row in block.rows:
                for cell in row:
                    _model_tables(cell, out)


def _cell_styles(table: _model.Table, counts: Counter) -> None:
    for row in table.rows:
        for cell in row:
            for block in cell:
                if isinstance(block, _model.Paragraph):
                    counts[block.properties.style_id if block.properties else None] += 1


def _table_styles(document: "Document") -> list[dict[str, Any]]:
    state = _effective.model(document)
    sheet = state.model
    groups: dict[str | None, dict[str, Any]] = {}
    parts = [(document.package.document_part(), sheet.body)]
    parts += [(state.model.story_parts.get(rel), blocks) for rel, blocks in sheet.stories.items()]
    for part, blocks in parts:
        if part is None:
            continue
        tables: list = []
        _model_tables(blocks, tables)
        tree = document.package.tree(part)
        index = document._index(part)
        for table in tables:
            applied = table_style_id(sheet, table)
            group = groups.setdefault(applied, {"tables": 0, "looks": Counter(), "cells": Counter(), "ids": []})
            group["tables"] += 1
            group["looks"][tuple(sorted((table.look or _DEFAULT_LOOK).items()))] += 1
            _cell_styles(table, group["cells"])
            element = resolve_path(tree, table.path)
            entry = index.entry_for(element) if element is not None else None
            if entry is not None and len(group["ids"]) < 3:
                group["ids"].append(entry.id)
    out = []
    for style_id, group in groups.items():
        style = sheet.styles.styles.get(style_id) if style_id and sheet.styles is not None else None
        look = dict(group["looks"].most_common(1)[0][0])
        cell_style = group["cells"].most_common(1)[0][0] if group["cells"] else None
        entry: dict[str, Any] = {"style": (style.name or style.style_id) if style is not None else None,
                                 "tables": group["tables"], "examples": group["ids"],
                                 "options": _options(look)}
        if style_id is not None:
            entry["places"] = _places(document, style_id, look, cell_style)
        out.append(entry)
    out.sort(key=lambda e: (-e["tables"], e["style"] or ""))
    return out


def _options(look: dict) -> dict[str, bool]:
    """``w:tblLook`` as Word's Table Design options name them."""
    return {"header_row": look["firstRow"], "total_row": look["lastRow"], "first_column": look["firstColumn"],
            "last_column": look["lastColumn"], "banded_rows": not look["noHBand"],
            "banded_columns": not look["noVBand"]}


def _places(document: "Document", style_id: str, look: dict, cell_style: str | None) -> dict[str, Any]:
    """The sample table's places under ``look``: each one's conditional formats, fill, borders
    and text, from docx2svg's table resolver and cascade."""
    sheet = _effective.model(document).model
    rows, columns = SAMPLE_ROWS, SAMPLE_COLUMNS
    conditions = tuple(tuple(cell_conditions(look, r, c, rows, columns) for c in range(columns)) for r in range(rows))
    paragraph = _model.Paragraph(properties=_model.ParagraphProperties(style_id=cell_style))
    sample = _model.Table(style_id=style_id, rows=tuple(tuple((paragraph,) for _ in range(columns))
                                                        for _ in range(rows)),
                          grid=(1440,) * columns, row_properties=({},) * rows,
                          cell_properties=(({},) * columns,) * rows, look=look, conditions=conditions)
    resolved = resolve_table(sheet, sample)
    first_row = 1 if look["firstRow"] else 0
    first_column = 1 if look["firstColumn"] else 0
    wanted = []
    if look["firstRow"]:
        wanted.append(("header_row", 0, first_column))
    band = "band" if not look["noHBand"] else "body"
    wanted.append((f"{band}1" if band == "band" else "body", first_row, first_column))
    if band == "band":
        wanted.append(("band2", first_row + 1, first_column))
    if look["firstColumn"]:
        wanted.append(("first_column", first_row, 0))
    if look["lastColumn"]:
        wanted.append(("last_column", first_row, columns - 1))
    if look["lastRow"]:
        wanted.append(("total_row", rows - 1, first_column))
    if look["firstRow"] and look["firstColumn"]:
        wanted.append(("corner", 0, 0))
    out = {}
    for name, r, c in wanted:
        out[name] = _place(document, resolved, r, c, conditions[r][c], cell_style, style_id)
    return out


def _place(document: "Document", resolved: Any, r: int, c: int, conditions: tuple, cell_style: str | None,
           style_id: str) -> dict[str, Any]:
    row = resolved.rows[r]
    cell = row.cells[c]
    above = resolved.rows[r - 1].cells[c] if r > 0 else None
    below = resolved.rows[r + 1].cells[c] if r + 1 < len(resolved.rows) else None
    top = cell.borders.get("top") if above is None else edge_winner(
        (above.borders.get("bottom"), above.own.get("bottom")), (cell.borders.get("top"), cell.own.get("top")))[0]
    bottom = cell.borders.get("bottom") if below is None else edge_winner(
        (cell.borders.get("bottom"), cell.own.get("bottom")), (below.borders.get("top"), below.own.get("top")))[0]
    paragraph, run = _effective.style_look(document, cell_style, table_style=style_id, conditions=conditions)
    text = _text_look(document, run)
    if paragraph.alignment != "left":
        text["alignment"] = paragraph.alignment
    themed = _table_theme_colors(document, style_id, conditions)
    out = {"conditions": [k for k in conditions if k != "wholeTable"], "fill": _fill(cell.properties.get("shd"))}
    if out["fill"] in themed["fill"]:
        out["fill"], out["theme_fill"] = themed["fill"][out["fill"]]
    out["borders"] = {side: _themed_border(border, themed["border"]) for side, border in
                      (("top", top), ("bottom", bottom), ("left", cell.left_edge[0]), ("right", cell.right_edge[0]))}
    out["text"] = text
    return out


def _table_theme_colors(document: "Document", style_id: str,
                        conditions: tuple = ()) -> dict[str, dict[str, tuple[str, str]]]:
    """A table style's theme colours at a place, by the ``w:fill`` / ``w:color`` value
    written beside them (``{"fill": {...}, "border": {...}}``, value -> ``(RRGGBB, label)``):
    docx2svg's table resolver reports the value, and this names the theme colour of the
    element holding it -- over the ``basedOn`` chain, root first, each style's own
    properties and then the place's conditional formats in ``conditions``' order, a later
    one taking a value over (the order the resolver merges them in)."""
    out: dict[str, dict[str, tuple[str, str]]] = {"fill": {}, "border": {}}
    chain, current = [], style_id
    while current and current not in chain:
        node = document.styles._node(current)
        if node is None:
            break
        chain.append(current)
        based = node.find(_W + "basedOn")
        current = based.get(_W + "val") if based is not None else None
    for current in reversed(chain):
        node = document.styles._node(current)
        holders = [child for child in node if child.tag in (_W + "tblPr", _W + "tcPr")]
        parts = {part.get(_W + "type"): part for part in node.findall(_W + "tblStylePr")}
        for condition in conditions:
            if condition in parts:
                holders.append(parts[condition])
        for holder in holders:
            for element in holder.iter(_W + "shd"):
                themed = _themed_element(document, element, _THEME_ATTRIBUTES["shd"])
                if themed is not None and element.get(_W + "fill"):
                    out["fill"][element.get(_W + "fill").upper()] = themed
            for borders in (*holder.iter(_W + "tblBorders"), *holder.iter(_W + "tcBorders")):
                for element in borders:
                    themed = _themed_element(document, element, _THEME_ATTRIBUTES["color"])
                    if themed is not None and element.get(_W + "color"):
                        out["border"][element.get(_W + "color").upper()] = themed
    return out


def _themed_border(border, themed: dict[str, tuple[str, str]]) -> str:
    """:func:`_border`, the colour the theme's where the style's element names a theme colour."""
    if border and border[0] not in (None, "nil", "none") and (border[3] or "").upper() in themed:
        rgb, label = themed[border[3].upper()]
        return _border((border[0], border[1], border[2], rgb)) + f" ({label})"
    return _border(border)


# -- declared styles ---------------------------------------------------------------------------

#: ``w:tblStylePr`` types -> the style options that turn them on.
_PART_OPTIONS = {"firstRow": "firstRow", "lastRow": "lastRow", "firstCol": "firstColumn", "lastCol": "lastColumn"}


def _style_usage(document: "Document") -> set[str]:
    """The ids of the styles content applies -- paragraphs, runs and tables in every story,
    the default paragraph style for a paragraph naming none."""
    used: set[str] = set()
    default = document.styles.default("paragraph")
    for part in document._parts():
        tree = document.package.tree(part)
        if tree is None:
            continue
        for node in tree.iter(_W + "pStyle", _W + "rStyle", _W + "tblStyle"):
            if node.get(_W + "val"):
                used.add(node.get(_W + "val"))
        if default is not None and any(p.find(f"{_W}pPr/{_W}pStyle") is None for p in tree.iter(_W + "p")):
            used.add(default.id)
    return used


def _declared_styles(document: "Document") -> list[dict[str, Any]]:
    """Every paragraph, character and table style the document declares, used or not, in
    its order: its type, whether content applies it (``used``), its ``basedOn`` chain
    (``based_on``, nearest first) and its look after inheritance as a paragraph, a run or a
    table of it draws it -- the same resolution as the styles in use.  To keep the list
    short, a look leaves out what is the same as the default paragraph style's look (that
    style's entry, ``default: true``, has all of it); a table style lists the conditional
    formats its chain declares (``parts``) and, per place (as ``table_styles`` names them),
    its fill, the borders drawn (sides with the same border joined: ``"top,bottom"``) and
    the text where it differs from the default paragraph style's.  Used styles come first,
    then paragraph, character and table styles, each in the document's order.  Kept for
    the document's state: a tool reading it a page at a time resolves it once."""
    cached = document.__dict__.get("_declared_styles")
    if cached is not None and cached[0] == document._state:
        return list(cached[1])
    out = _resolve_declared(document)
    document.__dict__["_declared_styles"] = (document._state, out)
    return list(out)


def _resolve_declared(document: "Document") -> list[dict[str, Any]]:
    styles = {s.id: s for s in document.styles}
    used = _style_usage(document)
    default = document.styles.default("paragraph")
    paragraph, run = _effective.style_look(document, default.id if default is not None else None)
    base = _look(document, paragraph, run)
    out = []
    for style in styles.values():
        if style.kind not in ("paragraph", "character", "table"):
            continue
        entry: dict[str, Any] = {"name": style.name or style.id, "type": style.kind, "used": style.id in used}
        if style.default:
            entry["default"] = True
        chain, current = [], style.based_on
        while current and current in styles and current not in chain and current != style.id:
            chain.append(current)
            current = styles[current].based_on
        if chain:
            entry["based_on"] = [styles[c].name or c for c in chain]
        if style.kind == "paragraph":
            paragraph, run = _effective.style_look(document, style.id)
            look = _look(document, paragraph, run)
            entry["look"] = look if default is not None and style.id == default.id else _differences(look, base)
        elif style.kind == "character":
            _, run = _effective.style_look(document, None, character_style=style.id)
            entry["look"] = _differences(_text_look(document, run), base)
        else:
            parts = _table_parts(document, [style.id, *chain])
            entry["parts"] = parts
            options = {option: part in parts for part, option in _PART_OPTIONS.items()}
            options.update(noHBand=not ({"band1Horz", "band2Horz"} & set(parts)), noVBand=True)
            places = {}
            for name, place in _places(document, style.id, options, None).items():
                compact: dict[str, Any] = {}
                for key in ("fill", "theme_fill"):
                    if place.get(key):
                        compact[key] = place[key]
                drawn: dict[str, list[str]] = {}
                for side, border in place["borders"].items():
                    if border != "none":
                        drawn.setdefault(border, []).append(side)
                if drawn:
                    compact["borders"] = {",".join(sides): border for border, sides in drawn.items()}
                text = _differences(place["text"], base)
                if text:
                    compact["text"] = text
                places[name] = compact
            entry["places"] = places
        out.append(entry)
    # Used styles first, then by type: a capped list keeps them.
    out.sort(key=lambda e: (not e["used"], ("paragraph", "character", "table").index(e["type"])))
    return out


def _differences(look: dict[str, Any], base: dict[str, Any]) -> dict[str, Any]:
    """``look`` less what is the same in ``base``; a flag ``base`` has and ``look`` lacks is
    given as ``False``, a theme colour's name kept with its colour."""
    out = {key: value for key, value in look.items() if base.get(key) != value
           and not (key.startswith("theme_") and look.get(key[6:]) == base.get(key[6:]))}
    for key in ("color", "shading"):
        if key in out and f"theme_{key}" in look:
            out[f"theme_{key}"] = look[f"theme_{key}"]
    for key, value in base.items():
        if value is True and key not in look:
            out[key] = False
    return out


def _table_parts(document: "Document", chain: list[str]) -> list[str]:
    """The conditional formats (``w:tblStylePr`` types) a table style's chain declares."""
    found: list[str] = []
    for style_id in chain:
        node = document.styles._node(style_id)
        for part in node.findall(_W + "tblStylePr") if node is not None else ():
            kind = part.get(_W + "type")
            if kind and kind not in found:
                found.append(kind)
    return found


# -- sections ----------------------------------------------------------------------------------


def _sections(document: "Document") -> list[dict[str, Any]]:
    even = document.even_and_odd_headers
    out = []
    for section in document.sections():
        columns = section.columns
        item: dict[str, Any] = {
            "id": section.id, "start": section.start, "orientation": section.orientation,
            "page": [_pt(section.page_width), _pt(section.page_height)],
            "margins": {k: _pt(v) for k, v in section.margins.items() if v is not None},
            "columns": {"count": columns["count"], "space": _pt(columns["space"])}}
        if columns.get("separator"):
            item["columns"]["separator"] = True
        if columns.get("widths"):
            item["columns"]["widths"] = [[_pt(w), _pt(s)] for w, s in columns["widths"]]
        if section.title_page:
            item["title_page"] = True
        stories = section.stories()
        linked = []
        for which in ("header", "footer"):
            shown = {}
            for kind in ("default", "first", "even"):
                if (kind == "first" and not section.title_page) or (kind == "even" and not even):
                    continue
                value = stories[which][kind]
                shown[kind] = value["story"]
                if not value["own"] and section.index > 0:
                    linked.append(f"{which} {kind}")
            item[which + "s"] = shown
        if linked:
            item["linked_to_previous"] = linked
        out.append(item)
    return out


# -- theme -------------------------------------------------------------------------------------


def _theme(document: "Document") -> dict[str, Any] | None:
    from .importing import _theme_root

    root = _theme_root(document)
    state = _effective.model(document).model
    if root is None and not state.theme_colors:
        return None
    out: dict[str, Any] = {}
    if root is not None and root.get("name"):
        out["name"] = root.get("name")
    out["colors"] = {slot: state.theme_colors[slot] for slot in THEME_SLOTS if slot in state.theme_colors}
    scheme = state.font_scheme
    if scheme is not None:
        fonts = {}
        for name in ("major", "minor"):
            faces = getattr(scheme, name)
            entry = {"latin": faces.latin}
            if faces.east_asian:
                entry["east_asian"] = faces.east_asian
            if faces.complex_script:
                entry["complex_script"] = faces.complex_script
            fonts[name] = entry
        out["fonts"] = fonts
    return out


# -- patterns ----------------------------------------------------------------------------------


def _patterns(document: "Document") -> list[dict[str, Any]]:
    from ..oxml.xml import qn

    body = document.package.document_part()
    index = document._index(body)
    blocks = []   # (level or None, token or None, id or None)
    for block in document.blocks():
        identifier = block["id"]
        if block["kind"] == "paragraph":
            entry = index.by_id.get(identifier)
            if entry is None or _empty(entry.element):
                blocks.append((None, None, identifier))
                continue
            blocks.append((block["level"], block["style"] or "", identifier))
        elif block["kind"] == "table" and identifier in index.by_id:
            node = index.by_id[identifier].element
            declared = node.find(f"{qn('w:tblPr')}/{qn('w:tblStyle')}")
            name = _table_style_name(document, declared.get(qn("w:val")) if declared is not None else "")
            columns = len(node.findall(f"{qn('w:tblGrid')}/{qn('w:gridCol')}"))
            blocks.append((None, f"table {name or 'no style'}, {columns} columns", identifier))
        else:
            blocks.append((None, "control", identifier))
    shapes: dict[tuple[int, tuple], list] = {}
    headings: Counter = Counter()
    for k, (level, _token, identifier) in enumerate(blocks):
        if level is None:
            continue
        headings[level] += 1
        shape: list[str] = []
        last = None
        for j in range(k + 1, len(blocks)):
            if blocks[j][0] is not None:
                break
            token = blocks[j][1]
            if token is None:
                continue
            if not shape or token != shape[-1]:
                shape.append(token)
            last = blocks[j][2]
        if shape and (len(shape) >= 2 or any(t.startswith("table ") for t in shape)):
            shapes.setdefault((level, tuple(shape)), []).append((identifier, last))
    out = []
    for (level, shape), found in shapes.items():
        if len(found) < 2:
            continue
        out.append({"heading_level": level, "shape": list(shape), "count": len(found), "of": headings[level],
                    "examples": [{"heading": h, "blocks": f"{h}..{last}"} for h, last in found[:3]]})
    out.sort(key=lambda e: (-e["count"], e["heading_level"], e["shape"]))
    return out
