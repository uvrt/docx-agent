"""Blocks compared with a reference block: measured, never judged (round 17, P3).

The Word counterpart of pptx-agent's slide comparison against a reference: a table (or a
paragraph) copied from a sample and filled is compared with the sample -- in this document
or another one -- and every place it differs is listed with the value the reference has and,
where a Word tool can make the change, the calls that make it.  Whether a difference is
intended is the caller's to decide.

**Tables**, as they state themselves (not as drawn):

* ``header_rows`` -- the rows repeated at the top of each page (``w:tblHeader``);
* ``columns`` -- the grid's columns; ``rows`` -- the body rows, the header rows apart: a copy
  filled with more items than the sample differs here by design (no fix);
* ``merges`` -- per row, its merged cells: ``c1..c2`` a cell spanning grid columns 1 to 2,
  ``c3 up`` a cell merged with the one above;
* ``widths`` -- the grid's column widths, points, each equal to the reference's within
  :data:`WIDTH_TOLERANCE`; ``width`` -- the table's preferred width (points, ``"n%"`` or
  ``"auto"``), the same tolerance;
* ``table_style``, ``table_alignment`` (on the page), ``table_shading`` and ``table_borders``;
* per cell: ``shading`` (its fill, ``RRGGBB``), ``borders`` and ``vertical_alignment``
  (``top`` where the cell states none: ECMA-376's ``w:vAlign`` default);
* per paragraph in a cell, and for a paragraph block: ``paragraph_style`` (by name, so
  another document's styles compare) and its direct formatting, in ``word_format``'s names
  -- ``alignment``, ``space_before``, ``space_after``, ``line_spacing``, ``indent_left``,
  ``indent_right``, ``indent_first``, ``keep_with_next``, ``keep_lines``, ``page_break_before`` (what its
  ``w:pPr`` states) and ``bold``, ``italic``, ``underline``, ``strike``, ``size``,
  ``font``, ``color``, ``highlight`` (what its runs state, the value most of its text has;
  ``None`` where nothing is stated, the style's then).

**Corresponding parts.**  Rows by the reference's structure: the reference's header rows one
to one, then body row *i* to the reference's body row *i*, or its last body row when the
copy has more (a sample with one body row is the pattern for all of them).  Cells by grid
column; a cell's paragraph *k* to the reference cell's paragraph *k*, or its last.

**Fixes** are the calls (each one call's arguments but ``doc``) that give the block the
reference's value, in the order listed: ``word_format_table`` (``header_rows``, a column's
``width``, ``style``, ``alignment``, ``shading``, ``borders`` where the reference's are one
style on every side as that tool writes it), ``word_edit_table`` ``split`` and ``merge``
(a row's horizontal merges, split first; the cells split off take the merged cell's
properties, as Word's Split Cells, so they are compared at their places), ``word_format`` (``paragraph_style``, a
direct value, or ``clear`` where the reference states none).  None where no tool can write
the reference's value: the row and column counts, a vertical merge, a table width in
``auto``, an exact line spacing, a border with its own colour or width.

**Width tolerance.**  Word stores lengths in twentieths of a point (twips) and a width
reaches a document by several routes -- typed in centimetres or inches, set in points by a
tool, spread over a grid by a builder -- each rounding to whole twips (1 cm is 566.93 of
them), so two widths meant to be the same can differ by a twip or two.  Half a point (10
twips) absorbs that rounding and is well below what a reader can see on the page (0.67 px
at 96 dpi); a column dragged by hand moves by more.
"""

from __future__ import annotations

from collections import Counter
from typing import TYPE_CHECKING, Any

from . import effective as _effective
from . import text as _text
from .table_format import grid_widths, own_rows, read_border, read_shading, read_width, row_cells

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: Two widths, points, equal within this (the module's "Width tolerance").
WIDTH_TOLERANCE = 0.5

#: Every property compared, in the order reported.
PROPERTIES = ("kind", "header_rows", "columns", "rows", "merges", "widths", "width", "table_style",
              "table_alignment", "table_shading", "table_borders", "shading", "borders", "vertical_alignment",
              "paragraph_style", "alignment", "space_before", "space_after", "line_spacing",
              "indent_left", "indent_right", "indent_first", "keep_with_next", "keep_lines",
              "page_break_before", "bold", "italic", "underline", "strike", "size", "font", "color", "highlight")

#: A paragraph's direct properties: word_format's name -> the w:pPr child and attributes
#: that state it (no attributes: the element itself).
_PARAGRAPH = {
    "alignment": ("jc", ()),
    "space_before": ("spacing", ("before", "beforeLines", "beforeAutospacing")),
    "space_after": ("spacing", ("after", "afterLines", "afterAutospacing")),
    "line_spacing": ("spacing", ("line",)),
    "indent_left": ("ind", ("left", "start")),
    "indent_right": ("ind", ("right", "end")),
    "indent_first": ("ind", ("firstLine", "hanging", "firstLineChars", "hangingChars")),
    "keep_with_next": ("keepNext", ()),
    "keep_lines": ("keepLines", ()),
    "page_break_before": ("pageBreakBefore", ()),
}
#: A run's direct properties: word_format's name -> the w:rPr children that state it.
_RUN = {
    "bold": ("b", "bCs"), "italic": ("i", "iCs"), "underline": ("u",), "strike": ("strike",),
    "size": ("sz", "szCs"), "font": ("rFonts",), "color": ("color",), "highlight": ("highlight",),
}
#: The border styles word_format_table writes (its ``borders``), each at half a point, in
#: automatic colour, with no space (:func:`docx_agent.edit.table_format.border`).
_TOOL_BORDERS = ("single", "double", "thick", "dotted", "dashed")
_CELL_SIDES = ("top", "left", "bottom", "right")
_TABLE_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")


class _Block:
    """One block read for comparing: its document, id and element."""

    def __init__(self, document: "Document", identifier: str) -> None:
        part, entry = document._resolve(identifier)
        self.document = document
        self.part = part
        self.id = entry.id
        self.element = entry.element
        self.kind = "table" if self.element.tag == _W + "tbl" else "paragraph"
        self._index = document._index(part)

    def paragraph_id(self, element) -> str | None:
        entry = self._index.entry_for(element)
        return entry.id if entry is not None else None


def compare_blocks(document: "Document", blocks: list[str], reference: str,
                   reference_document: "Document | None" = None) -> dict[str, Any]:
    """Each of ``blocks`` (paragraph or table ids of ``document``) compared with
    ``reference`` (a block of ``reference_document``, else of ``document``):
    ``{"reference", "compared", "same", "differences"}``, each difference
    ``{"block", "property", "cells" | "rows" | "paragraphs", "value", "reference", "fix"}``
    (the module says what each is)."""
    source = _Block(reference_document or document, reference)
    out: list[dict] = []
    same: list[str] = []
    compared: list[str] = []
    for identifier in blocks:
        block = _Block(document, identifier)
        compared.append(block.id)
        found: list[dict] = []
        if block.kind != source.kind:
            found.append({"property": "kind", "value": block.kind, "reference": source.kind, "fix": None})
        elif block.kind == "table":
            found += _tables(block, source)
        else:
            found += _paragraphs(block, source.document, [(block.element, source.element)])
        if not found:
            same.append(block.id)
        order = {name: k for k, name in enumerate(PROPERTIES)}
        found.sort(key=lambda d: order.get(d["property"], len(order)))
        out += [{"block": block.id, **entry} for entry in found]
    return {"reference": source.id, "compared": compared, "same": same, "differences": out}


# -- tables ------------------------------------------------------------------------------------


def _header_rows(rows: list) -> int:
    count = 0
    for row in rows:
        node = row.find(f"{_W}trPr/{_W}tblHeader")
        if node is None or (node.get(_W + "val") or "1").lower() in ("0", "false", "off"):
            break
        count += 1
    return count


def _row_map(rows: int, reference_rows: int, header: int) -> list[int]:
    """Per row of a copy, the reference's row it corresponds to (the module's rule)."""
    out = []
    body = max(reference_rows - header, 0)
    for r in range(rows):
        if r < header or body == 0:
            out.append(min(r, reference_rows - 1))
        else:
            out.append(header + min(r - header, body - 1))
    return out


def _origins(rows: list) -> list[dict[int, tuple[Any, int, bool]]]:
    """Per row, ``{first grid column: (cell, span, merged up)}``."""
    out = []
    for row in rows:
        line = {}
        for cell, column, span in row_cells(row):
            merge = cell.find(f"{_W}tcPr/{_W}vMerge")
            up = merge is not None and merge.get(_W + "val") in (None, "continue")
            line[column] = (cell, span, up)
        out.append(line)
    return out


def _merges(line: dict) -> list[str]:
    out = []
    for column in sorted(line):
        _, span, up = line[column]
        if span > 1:
            out.append(f"c{column}..c{column + span - 1}")
        if up:
            out.append(f"c{column} up")
    return out


def _side(node) -> str:
    border = read_border(node)
    if border is None:
        return "unset"
    if border["style"] in ("nil", "none"):
        return "none"
    words = [border["style"], f"{border.get('size', 0):g}pt", border.get("color", "auto")]
    if border.get("space"):
        words.append(f"space {border['space']}")
    if border.get("theme_color"):
        words.append(border["theme_color"])
    return " ".join(words)


def _borders(container, sides: tuple[str, ...]) -> dict[str, str] | None:
    if container is None:
        return None
    found = {side: _side(container.find(_W + side)) for side in sides}
    found = {side: value for side, value in found.items() if value != "unset"}
    return found or None


def _border_fix(reference: dict | None, sides: tuple[str, ...]) -> str | None:
    """``word_format_table``'s ``borders`` value that writes ``reference``, or ``None``."""
    if reference is None or set(reference) != set(sides):
        return None
    values = set(reference.values())
    if values == {"none"}:
        return "none"
    if len(values) == 1:
        style, _, rest = next(iter(values)).partition(" ")
        if style in _TOOL_BORDERS and rest == "0.5pt auto":
            return style
    return None


def _fill(node) -> str | None:
    shading = read_shading(node)
    if shading is None:
        return None
    fill = shading.get("fill")
    return None if fill in (None, "auto") else fill.upper()


def _close(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= WIDTH_TOLERANCE
    return a == b


def _style_name(document: "Document", style_id: str | None) -> str | None:
    if style_id is None:
        return None
    found = next((s for s in document.styles if s.id == style_id), None)
    return found.name if found is not None else style_id


def _has_style(document: "Document", name: str, kind: str) -> bool:
    try:
        document.styles.resolve(name, kind)
        return True
    except Exception:  # noqa: BLE001 -- a style the document neither has nor would add
        return False


def _tables(block: _Block, source: _Block) -> list[dict]:
    document = block.document
    out: list[dict] = []
    rows, reference_rows = own_rows(block.element), own_rows(source.element)
    if not rows or not reference_rows:
        return out

    def table(name: str, value, reference, fix_values: dict | None) -> None:
        if value == reference:
            return
        fix = ({"tool": "word_format_table", "calls": [{"table": block.id, "scope": "table", **fix_values}]}
               if fix_values is not None else None)
        out.append({"property": name, "value": value, "reference": reference, "fix": fix})

    header, reference_header = _header_rows(rows), _header_rows(reference_rows)
    table("header_rows", header, reference_header, {"header_rows": reference_header})
    widths = [round(w / 20, 2) for w in grid_widths(block.element)]
    reference_widths = [round(w / 20, 2) for w in grid_widths(source.element)]
    table("columns", len(widths), len(reference_widths), None)
    body, reference_body = len(rows) - reference_header, len(reference_rows) - reference_header
    if body != reference_body:
        out.append({"property": "rows", "value": max(body, 0), "reference": reference_body, "fix": None})
    mapping = _row_map(len(rows), len(reference_rows), reference_header)
    grid, reference_grid = _origins(rows), _origins(reference_rows)
    _table_merges(block, out, grid, reference_grid, mapping)
    differing = [k for k in range(min(len(widths), len(reference_widths)))
                 if not _close(widths[k], reference_widths[k])]
    if differing:
        out.append({"property": "widths", "value": widths, "reference": reference_widths,
                    "columns": differing,
                    "fix": {"tool": "word_format_table", "calls": [
                        {"table": block.id, "scope": "column", "column": k, "width": reference_widths[k]}
                        for k in differing]} if len(widths) == len(reference_widths) else None})
    props, reference_props = _table_props(block), _table_props(source)
    width, reference_width = props["width"], reference_props["width"]
    if not _close(width, reference_width):
        fix = ({"width": reference_width} if isinstance(reference_width, (int, float)) and reference_width >= 1
               else {"width_percent": float(reference_width[:-1])}
               if isinstance(reference_width, str) and reference_width.endswith("%") else None)
        table("width", width, reference_width, fix)
    style = reference_props["style"]
    table("table_style", props["style"], style,
          {"style": style} if style is not None and _has_style(document, style, "table") else None)
    alignment = reference_props["alignment"]
    table("table_alignment", props["alignment"], alignment,
          {"alignment": alignment} if alignment in ("left", "center", "right") else None)
    shading = reference_props["shading"]
    table("table_shading", props["shading"], shading, {"shading": shading or "none"})
    borders = reference_props["borders"]
    fix_borders = _border_fix(borders, _TABLE_SIDES)
    table("table_borders", props["borders"], borders, {"borders": fix_borders} if fix_borders else None)
    _cells(block, source, out, grid, reference_grid, mapping)
    return out


def _table_props(block: _Block) -> dict[str, Any]:
    properties = block.element.find(_W + "tblPr")

    def find(tag: str):
        return properties.find(_W + tag) if properties is not None else None

    style = find("tblStyle")
    jc = find("jc")
    alignment = jc.get(_W + "val") if jc is not None else None
    alignment = {"start": "left", "end": "right"}.get(alignment, alignment)
    return {"style": _style_name(block.document, style.get(_W + "val")) if style is not None else None,
            "width": read_width(find("tblW")), "alignment": alignment,
            "shading": _fill(find("shd")), "borders": _borders(find("tblBorders"), _TABLE_SIDES)}


def _table_merges(block: _Block, out: list, grid: list, reference_grid: list, mapping: list) -> None:
    groups: dict[tuple, list[int]] = {}
    for r, line in enumerate(grid):
        value, reference = _merges(line), _merges(reference_grid[mapping[r]])
        if value != reference:
            groups.setdefault((tuple(value), tuple(reference)), []).append(r)
    for (value, reference), rows in groups.items():
        calls: list[dict] | None = []
        for r in rows:
            found = _merge_calls(block.id, r, grid[r], reference_grid[mapping[r]])
            if found is None:
                calls = None
                break
            calls += found
        out.append({"property": "merges", "rows": rows, "value": list(value), "reference": list(reference),
                    "fix": {"tool": "word_edit_table", "calls": calls} if calls else None})


def _merge_calls(table: str, r: int, line: dict, reference: dict) -> list[dict] | None:
    """``word_edit_table`` calls that give row ``r`` the reference row's horizontal merges:
    a split of each merged cell the reference does not have, then a merge of each the row
    lacks; ``None`` where a vertical merge is involved or the grids differ."""
    if any(up for _, _, up in line.values()) or any(up for _, _, up in reference.values()):
        return None
    mine = {(c, span) for c, (_, span, _) in line.items()}
    theirs = {(c, span) for c, (_, span, _) in reference.items()}
    if sum(span for _, span in mine) != sum(span for _, span in theirs):
        return None
    calls = [{"table": table, "action": "split", "first": {"row": r, "col": c}, "rows": 1, "columns": span}
             for c, span in sorted(mine - theirs) if span > 1]
    calls += [{"table": table, "action": "merge", "first": {"row": r, "col": c},
               "last": {"row": r, "col": c + span - 1}} for c, span in sorted(theirs - mine) if span > 1]
    return calls


def _cells(block: _Block, source: _Block, out: list, grid: list, reference_grid: list, mapping: list) -> None:
    """Each cell's shading, borders and vertical alignment, and its paragraphs."""
    differences: dict[tuple, list[tuple[int, int]]] = {}
    pairs: list[tuple] = []
    for r, line in enumerate(grid):
        reference_line = reference_grid[mapping[r]]
        for column, (cell, span, up) in sorted(line.items()):
            if up:
                continue        # its content is the merge's first cell's
            match = _reference_cell(reference_line, column)
            if match is None:
                continue
            # A merged cell the reference has as single cells: the merges fix splits it, and
            # the cells split off take its properties (as Word's Split Cells), so each is
            # compared with the reference's cell at its place.
            split = span > 1 and all(reference_line.get(c, (None, 0, False))[1] == 1
                                     for c in range(column, column + span))
            for place in range(column, column + span) if split else (column,):
                against = reference_line[place][0] if split else match
                for name, value, reference in _cell_values(cell, against):
                    if value != reference:
                        differences.setdefault((name, _key(value), _key(reference)), []).append((r, place))
            paragraphs = cell.findall(_W + "p")
            reference_paragraphs = match.findall(_W + "p")
            if reference_paragraphs:
                pairs += [(p, reference_paragraphs[min(k, len(reference_paragraphs) - 1)])
                          for k, p in enumerate(paragraphs)]
    columns = {c for line in grid for c in line}
    for (name, value, reference), cells in differences.items():
        value, reference = _unkey(value), _unkey(reference)
        out.append({"property": name, "cells": _cell_refs(block.id, cells, grid, columns),
                    "value": value, "reference": reference,
                    "fix": _cell_fix(block, name, reference, cells, grid)})
    out += _paragraphs(block, source.document, pairs)


def _reference_cell(line: dict, column: int):
    if column in line:
        return line[column][0]
    for start, (cell, span, _) in line.items():
        if start <= column < start + span:
            return cell
    return None


def _cell_values(cell, reference):
    def read(tc):
        properties = tc.find(_W + "tcPr")
        find = (lambda tag: properties.find(_W + tag)) if properties is not None else (lambda tag: None)
        valign = find("vAlign")
        return {"shading": _fill(find("shd")), "borders": _borders(find("tcBorders"), _CELL_SIDES),
                # Not stated is top (ECMA-376's w:vAlign default), which word_format_table writes.
                "vertical_alignment": (valign.get(_W + "val") if valign is not None else None) or "top"}

    mine, theirs = read(cell), read(reference)
    for name in ("shading", "borders", "vertical_alignment"):
        yield name, mine[name], theirs[name]


def _key(value):
    return tuple(sorted(value.items())) if isinstance(value, dict) else value


def _unkey(value):
    return dict(value) if isinstance(value, tuple) else value


def _cell_refs(table: str, cells: list[tuple[int, int]], grid: list, columns: set) -> list[str]:
    """``t:X/cR,C`` per cell, ``t:X/cR,*`` for every cell of a row."""
    by_row: dict[int, list[int]] = {}
    for r, c in cells:
        by_row.setdefault(r, []).append(c)
    out = []
    for r, found in sorted(by_row.items()):
        own = [c for c, (_, _, up) in grid[r].items() if not up]
        if len(own) > 1 and sorted(found) == sorted(own):
            out.append(f"{table}/c{r},*")
        else:
            out += [f"{table}/c{r},{c}" for c in sorted(found)]
    return out


def _cell_fix(block: _Block, name: str, reference, cells: list[tuple[int, int]], grid: list) -> dict | None:
    if name == "shading":
        values = {"shading": reference or "none"}
    elif name == "vertical_alignment":
        if reference not in ("top", "center", "bottom"):
            return None
        values = {"vertical_alignment": reference}
    else:
        style = _border_fix(reference, _CELL_SIDES)
        if style is None:
            return None
        values = {"borders": style}
    by_row: dict[int, list[int]] = {}
    for r, c in cells:
        by_row.setdefault(r, []).append(c)
    calls = []
    for r, found in sorted(by_row.items()):
        own = [c for c, (_, _, up) in grid[r].items() if not up]
        if len(own) > 1 and sorted(found) == sorted(own) and not any(up for _, _, up in grid[r].values()):
            calls.append({"table": block.id, "scope": "row", "row": r, **values})
        else:
            calls += [{"table": block.id, "scope": "cell", "row": r, "column": c, **values} for c in sorted(found)]
    return {"tool": "word_format_table", "calls": calls}


# -- paragraphs --------------------------------------------------------------------------------


def _paragraphs(block: _Block, reference_document: "Document", pairs: list[tuple]) -> list[dict]:
    """The paragraph style and direct formatting of each ``(paragraph, reference)``, the
    reference's in ``reference_document``."""
    document = block.document
    groups: dict[tuple, list[str]] = {}
    for paragraph, reference in pairs:
        identifier = block.paragraph_id(paragraph)
        if identifier is None:
            continue
        mine = _paragraph_values(document, paragraph)
        theirs = _paragraph_values(reference_document, reference)
        for name in mine:
            if mine[name] != theirs[name]:
                groups.setdefault((name, _key(mine[name]), _key(theirs[name])), []).append(identifier)
    out = []
    for (name, value, reference), paragraphs in groups.items():
        value, reference = _unkey(value), _unkey(reference)
        out.append({"property": name,
                    "paragraphs": paragraphs, "value": value, "reference": reference,
                    "fix": _paragraph_fix(document, name, reference, paragraphs)})
    return out


def _paragraph_values(document: "Document", paragraph) -> dict[str, Any]:
    properties = paragraph.find(_W + "pPr")
    style = properties.find(_W + "pStyle") if properties is not None else None
    values: dict[str, Any] = {"paragraph_style": _style_name(document, style.get(_W + "val") if style is not None
                                                             else _default_style(document))}
    declared = {name for name, (tag, attributes) in _PARAGRAPH.items()
                if properties is not None and (node := properties.find(_W + tag)) is not None
                and (not attributes or any(node.get(_W + a) is not None for a in attributes))}
    effective = _effective.paragraph(document, paragraph) if declared else None
    for name in _PARAGRAPH:
        values[name] = _paragraph_value(effective, name) if name in declared else None
    values.update(_run_values(document, paragraph))
    return values


def _default_style(document: "Document") -> str | None:
    found = next((s for s in document.styles if getattr(s, "kind", None) == "paragraph"
                  and getattr(s, "default", False)), None)
    return found.id if found is not None else None


def _paragraph_value(effective, name: str):
    if name == "indent_first":
        return round(effective.first_line - effective.hanging, 2)
    if name == "line_spacing":
        if effective.line_spacing is None:
            return None
        return (round(effective.line_spacing, 2) if effective.line_rule == "auto"
                else f"{effective.line_rule} {round(effective.line_spacing, 2):g}pt")
    value = getattr(effective, "keep_together" if name == "keep_lines" else name)
    return round(value, 2) if isinstance(value, float) else value


def _run_values(document: "Document", paragraph) -> dict[str, Any]:
    walked = _text.walk(paragraph)
    texts: dict[int, tuple[Any, int]] = {}
    for atom in walked.atoms:
        if atom.run is not None and atom.kind == "text":
            run, count = texts.get(id(atom.run), (atom.run, 0))
            texts[id(atom.run)] = (run, count + 1)
    counts: dict[str, Counter] = {name: Counter() for name in _RUN}
    for run, count in texts.values():
        rpr = run.find(_W + "rPr")
        declared = {name for name, tags in _RUN.items()
                    if rpr is not None and any(rpr.find(_W + tag) is not None for tag in tags)}
        effective = _effective.run(document, paragraph, run, None) if declared else None
        for name in _RUN:
            value = _run_value(effective, name) if name in declared else None
            counts[name][value] += count
    return {name: (max(found.items(), key=lambda item: (item[1], item[0] is None))[0] if found else None)
            for name, found in counts.items()}


def _run_value(effective, name: str):
    value = getattr(effective, name)
    if name == "underline":
        return value if value not in (None, "none") else "none"
    if name == "size":
        return round(value, 2)
    return value


def _paragraph_fix(document: "Document", name: str, reference, paragraphs: list[str]) -> dict | None:
    call: dict[str, Any] = {"targets": paragraphs}
    if name == "paragraph_style":
        if reference is None or not _has_style(document, reference, "paragraph"):
            return None
        call["paragraph_style"] = reference
    elif reference is None:
        call["clear"] = [name]
    elif name == "line_spacing" and not isinstance(reference, (int, float)):
        return None                     # exact or at-least: word_format sets a multiple only
    elif name == "underline":
        if reference not in ("single", "none"):
            return None
        call["underline"] = reference == "single"
    elif name == "alignment":
        word = {"both": "justify", "start": "left", "end": "right"}.get(reference, reference)
        if word not in ("left", "center", "right", "justify"):
            return None
        call["alignment"] = word
    elif name == "highlight":
        call["highlight"] = reference
    elif name == "color":
        if not (isinstance(reference, str) and len(reference) == 6):
            return None
        call["color"] = reference
    else:
        call[name] = reference
    return {"tool": "word_format", "calls": [call]}


__all__ = ["PROPERTIES", "WIDTH_TOLERANCE", "compare_blocks"]
