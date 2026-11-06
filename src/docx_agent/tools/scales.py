"""Values styled by a scale in a Word document (production round 36, R31): the docx side of
``style_by_value`` and ``check``'s ``scale`` (:mod:`ooxml_edit.tools.scales` reads the scale,
chooses each value's step and reports; this module reads the values and writes the steps).

**Targets** (:func:`scale_values`, the docx ``DocumentFormat.scale_values``):

* a paragraph (``p:A``): its text is the value, styled whole;
* a text range in one paragraph (``p:A@4:5``): its characters;
* a cell (``t:X/c1,2``): its text, the cell shaded and its text styled;
* a table's column (``t:X/col2``, a grid column from 0): every body cell of it -- the rows
  marked to repeat as a header (``w:tblHeader``) left out, and the first row where none is
  marked and the table's look has its header row on (as Word draws it);
* values written by reference (``{{d:scores/*/score}}``): each value a text tool wrote from
  a matching reference and still where it was written (:meth:`docx_agent.Document.text_facts`)
  -- its characters, or its cell where it is the cell's whole text.

**Styles** (:func:`write`): one character style per step, named from the scale
(``Score 1 Low``), made once and reused.  As Word writes a character style made with a font
colour and shading (measured in Word for Mac 16.106, round 36: a made-up document, the style
made and applied by script, saved): ``w:type="character"``, ``w:customStyle="1"``, its id the
name without spaces (``Score1Low``), based on Default Paragraph Font, and in its ``w:rPr``
the text colour (``w:color``) and the fill (``w:shd w:val="clear" w:color="auto"``); applied
as ``w:rStyle`` on the value's runs.  A style the document has already by that name is
reused; where its colours are not the step's, they are set (every value in it follows: a
warning says so).  A cell is shaded as Word shades one (``w:shd`` in its ``w:tcPr``, the same
form), and its text gets the step's style.  Direct colour, shading and highlight on a styled
value's runs are taken off: they would hide the step.

**The step a value has now**: the scale's style its text's runs carry (by the style's name:
the scale's name and the step's number, whatever label follows); ``exact`` where every run
carries it, the style holds the step's colours, no run overrides them and a cell's fill is the
step's.
"""

from __future__ import annotations

from typing import Any, Mapping

from ooxml_edit.tools import Result, ScaleValue, ToolError
from ooxml_edit.tools import scales as _scales

from ..edit import text as _text
from ..edit import formatting as _formatting
from ..edit.document import Cell, Paragraph, Table
from ..edit.ranges import TextRange
from ..oxml.xml import remove

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_TC = _W + "tc"
#: A run's direct properties that would hide a step's colours.
_HIDING = ("color", "shd", "highlight")


# -- reading --------------------------------------------------------------------------------------


def scale_values(call: Any, scale: _scales.Scale, scope: Mapping[str, Any]) -> list[ScaleValue]:
    """``DocumentFormat.scale_values`` for a Word document (the module's notes); ``scope``:
    ``check``'s ``range`` and ``pages``."""
    document = call.document
    held = _held(call, scope)
    styles = _style_names(document)
    out: list[ScaleValue] = []
    for target in scale.targets:
        try:
            found = _read_target(document, scale, target, styles)
        except ToolError as error:
            found = [ScaleValue(target, skip=error.message)]
        for value in found:
            if held is not None and value.skip is None and not held(value):
                continue
            out.append(value)
    return out


def _held(call: Any, scope: Mapping[str, Any]) -> Any:
    """A test of whether a value is in ``scope`` (``None``: everything is), and each value's
    page where the layout is read (``pages``)."""
    from ..scope import Scope, ranges_of
    from ._base import engage, release, resolve
    from .shared import _pages_of
    from .textfacts import _layout

    pages = _pages_of(call, scope.get("slides"), scope.get("pages"))
    ranges = ranges_of(resolve(call, scope.get("range")))
    if pages is None and not ranges:
        return None
    paragraphs = Scope(call.document, ranges).paragraphs if ranges else None
    layout = None
    if pages is not None:
        engage(call)
        try:
            layout = _layout(call, required=True)
        finally:
            release(call.entries.values())

    def held(value: ScaleValue) -> bool:
        ids = value.handle[2] if value.handle else []
        if paragraphs is not None and not set(ids) & paragraphs:
            return False
        if layout is not None:
            on = sorted({n for p in ids for n in layout.pages_of(p)})
            if not set(on) & set(pages):
                return False
            if on:
                value.where = {"page": on[0]}
        return True

    return held


def _style_names(document: Any) -> dict[str, str]:
    return {style.id: style.name or style.id for style in document.styles if style.kind == "character"}


def _read_target(document: Any, scale: _scales.Scale, target: str, styles: dict[str, str]) -> list[ScaleValue]:
    column = _scales.column_target(target)
    if column is not None:
        return _column(document, scale, target, *column, styles)
    wanted = _scales.reference_target(target)
    if wanted is not None:
        return _referenced(document, scale, target, wanted, styles)
    if "@" in target and not target.startswith("p@"):
        try:
            where = document.range(target)
        except (KeyError, ValueError) as error:
            raise ToolError("not_found", f"no range {target}: {_said(error)}") from None
        if where.start_id != where.end_id:
            raise ToolError("invalid_arguments", f"{target} runs across paragraphs: a value is in one "
                            "paragraph (p:A@4:5)")
        if where.end <= where.start:
            raise ToolError("invalid_arguments", f"{target} is a position, not characters: give "
                            "p:A@<start>:<end>")
        paragraph = document.paragraph(where.start_id)
        return [_text_value(document, scale, target, None, paragraph, where.start, where.end, styles)]
    try:
        found = document.get(target)
    except (KeyError, ValueError, IndexError) as error:
        raise ToolError("not_found", f"no {target}: {_said(error)}; a target is a paragraph, a range "
                        "p:A@4:5, a cell t:X/c1,2, a column t:X/col2 or {{d:...}} values") from None
    if isinstance(found, Paragraph):
        return [_text_value(document, scale, target, None, found, 0, None, styles)]
    if isinstance(found, Cell):
        return [_cell_value(document, scale, found, None, styles)]
    if isinstance(found, Table):
        raise ToolError("invalid_arguments", f"{target} is a table: name its column ({target}/col1) "
                        "or its cells")
    raise ToolError("invalid_arguments", f"{target} is a {type(found).__name__.lower()}: a target is a "
                    "paragraph, a range, a cell, a table's column or {{d:...}} values")


def _said(error: BaseException) -> str:
    return str(error.args[0]) if error.args else str(error)


def _runs_over(element: Any, start: int, end: int | None) -> tuple[str, list[Any]]:
    atoms = _text.atoms(element)
    chosen = atoms[start:end]
    text = "".join(a.char for a in chosen).replace(_text.OBJECT, " ")
    runs: list[Any] = []
    for atom in chosen:
        if atom.char.strip() and atom.run is not None and atom.run not in runs:
            runs.append(atom.run)
    return text, runs


def _now(document: Any, scale: _scales.Scale, runs: list[Any], styles: dict[str, str],
         fill: str | None = None, cell: bool = False) -> tuple[int | None, bool, str | None]:
    """The step the runs carry, whether exactly, and what differs."""
    found: list[int | None] = []
    direct = False
    for run in runs:
        properties = run.find(_W + "rPr")
        node = properties.find(_W + "rStyle") if properties is not None else None
        name = styles.get(node.get(_W + "val") or "") if node is not None else None
        found.append(scale.step_of_style(name))
        if properties is not None and any(properties.find(_W + tag) is not None for tag in _HIDING):
            direct = True
    steps = [n for n in found if n is not None]
    if not steps:
        return None, True, None
    now = max(set(steps), key=steps.count)
    step = scale.steps[now]
    if len(set(found)) > 1:
        return now, False, "styled in part"
    held = _style_colours(document, scale, step)
    if held is not None:
        return now, False, held
    if direct:
        return now, False, "direct colour on its text"
    if cell and (fill or "").upper() != step.fill:
        return now, False, f"cell fill {fill or 'none'}"
    return now, True, None


def _style_colours(document: Any, scale: _scales.Scale, step: _scales.Step) -> str | None:
    """What a step's style says other than the step (``None``: its colours)."""
    style = document.styles.find(scale.style_name(step), "character")
    if style is None:
        return None
    fill, text = _colours_of(document, style.id)
    if fill != step.fill or text != step.text:
        return f"style {style.name} fill {fill or 'none'}, text {text or 'auto'}"
    return None


def _colours_of(document: Any, style_id: str) -> tuple[str | None, str | None]:
    node = document.styles._node(style_id)
    properties = node.find(_W + "rPr") if node is not None else None
    if properties is None:
        return None, None
    shading, colour = properties.find(_W + "shd"), properties.find(_W + "color")
    fill = shading.get(_W + "fill") if shading is not None else None
    text = colour.get(_W + "val") if colour is not None else None
    return (fill.upper() if fill else None), (text.upper() if text else None)


def _text_value(document: Any, scale: _scales.Scale, target: str, given: str | None, paragraph: Any,
                start: int, end: int | None, styles: dict[str, str]) -> ScaleValue:
    element = paragraph._element
    text, runs = _runs_over(element, start, end)
    if not text.strip():
        return ScaleValue(target, given=given, skip=f"{target} holds no text")
    now, exact, differs = _now(document, scale, runs, styles)
    stop = len(_text.atoms(element)) if end is None else end
    address = paragraph.id if start == 0 and end is None else f"{paragraph.id}@{start}:{stop}"
    return ScaleValue(address, text.strip(), given=given, now=now, exact=exact, differs=differs,
                      handle=("range", (paragraph.id, start, stop), [paragraph.id]))


def _cell_value(document: Any, scale: _scales.Scale, cell: Cell, given: str | None,
                styles: dict[str, str]) -> ScaleValue:
    element = cell._element
    paragraphs = cell.paragraphs
    texts, runs = [], []
    for paragraph in paragraphs:
        text, own = _runs_over(paragraph._element, 0, None)
        texts.append(text)
        runs += own
    text = "\n".join(texts).strip()
    if not text:
        return ScaleValue(cell.id, given=given, skip=f"{cell.id} is empty")
    shading = element.find(f"{_W}tcPr/{_W}shd")
    fill = shading.get(_W + "fill") if shading is not None else None
    fill = None if fill in (None, "auto") else fill.upper()
    now, exact, differs = _now(document, scale, runs, styles, fill, cell=True)
    if now is None and fill is not None:
        match = next((s for s in scale.steps if s.fill == fill), None)
        if match is not None:              # shaded as a step, its text not styled
            now, exact, differs = match.index, False, "text not in the step's style"
    return ScaleValue(cell.id, text, given=given, now=now, exact=exact, differs=differs,
                      handle=("cell", cell.id, [p.id for p in paragraphs]))


def _column(document: Any, scale: _scales.Scale, target: str, table_id: str, column: int,
            styles: dict[str, str]) -> list[ScaleValue]:
    try:
        table = document.table(table_id)
        table.id                                   # resolves it
    except (KeyError, ValueError) as error:
        raise ToolError("not_found", f"no table {table_id}: {_said(error)}") from None
    width = table.column_count
    if column >= width:
        raise ToolError("invalid_arguments", f"{table_id} has grid columns 0 to {width - 1}: no column "
                        f"{column}")
    properties = document.table_properties(table.id)
    header = properties.get("header_rows") or 0
    if not header and (properties.get("look") or {}).get("first_row"):
        header = 1
    grid = table.grid()
    out: list[ScaleValue] = []
    seen: set[str] = set()
    for r in range(header, len(grid)):
        cell = grid[r][column] if column < len(grid[r]) else None
        if cell is None or cell.id in seen:
            continue
        seen.add(cell.id)
        if cell.row != r:                           # merged from a row above
            continue
        out.append(_cell_value(document, scale, cell, target, styles))
    if not out:
        raise ToolError("not_found", f"{target}: no body cells ({header} header row(s) of {len(grid)})")
    return out


def _referenced(document: Any, scale: _scales.Scale, target: str, wanted: str,
                styles: dict[str, str]) -> list[ScaleValue]:
    from .textfacts import _cell_id

    facts = document.text_facts()
    main = document.package.document_part()
    index = document._index(main)
    order = {entry.id: k for k, entry in enumerate(index.paragraphs)}
    out: list[tuple[int, int, ScaleValue]] = []
    for paragraph_id, found in facts.items():
        for fact in found:
            if not isinstance(fact.get("source"), str) or not _scales.reference_matches(wanted, fact["source"]):
                continue
            try:
                paragraph = document.paragraph(paragraph_id)
                element = paragraph._element
            except (KeyError, ValueError):
                continue
            at, written = int(fact["at"]), str(fact.get("text") or "")
            cell = next(element.iterancestors(_TC), None)
            value = None
            if cell is not None:
                address = _cell_id(index, cell)
                if address is not None:
                    found_cell = document.get(address)
                    if isinstance(found_cell, Cell) and found_cell.text.strip() == written.strip():
                        value = _cell_value(document, scale, found_cell, target, styles)
            if value is None:
                value = _text_value(document, scale, f"{paragraph.id}@{at}:{at + len(written)}", target,
                                    paragraph, at, at + len(written), styles)
            out.append((order.get(paragraph.id, len(order)), at, value))
    if not out:
        raise ToolError("not_found", f"{target}: no value written by a reference like it is still in the "
                        "document's text")
    seen: set[str] = set()
    values = []
    for _, _, value in sorted(out, key=lambda item: (item[0], item[1])):
        if value.target not in seen:
            seen.add(value.target)
            values.append(value)
    return values


# -- writing --------------------------------------------------------------------------------------


def write(call: Any, scale: _scales.Scale, plan: list[tuple[ScaleValue, _scales.Step]],
          preview: bool) -> Result:
    """Write each value's step (the module's notes); with ``preview``, only say which styles
    would be made or changed."""
    from ._base import outcome

    document = call.document
    edits: list[Any] = []
    styles: list[dict[str, Any]] = []
    warnings: list[str] = []
    needed = sorted({step.index for _, step in plan})
    ids: dict[int, str] = {}
    for index in needed:
        step = scale.steps[index]
        name = scale.style_name(step)
        style = document.styles.find(name)
        if style is not None and style.kind != "character":
            raise ToolError("invalid_arguments", f"{name!r} is a {style.kind} style here: give the scale "
                            "another name", field="scale")
        if style is None:
            styles.append({"style": name, "made": True})
            if not preview:
                edits.append(_make_style(document, name, step))
        else:
            fill, text = _colours_of(document, style.id)
            if fill != step.fill or text != step.text:
                uses = sum(document.styles.usage(style.name, "character").values())
                styles.append({"style": name, "changed": True})
                warnings.append(f"{name}: its colours set to the step's (fill {step.fill}, text {step.text}); "
                                f"its {uses} use(s) follow")
                if not preview:
                    edits.append(document.styles.modify(style.name, "character", color=step.text,
                                                        shading=step.fill))
            else:
                styles.append({"style": name, "reused": True})
        if not preview:
            ids[index] = document.styles.find(name, "character").id
    if not preview:
        for value, step in plan:
            edits.extend(_apply(document, value, step, ids[step.index]))
    result = outcome(edits, "")
    result.warnings.extend(warnings)
    result.data = {"steps": [s.brief(style=scale.style_name(s)) for s in scale.steps], "styles": styles}
    return result


def _make_style(document: Any, name: str, step: _scales.Step) -> Any:
    """A character style as Word writes one made with a font colour and shading (the
    module's notes): no ``w:qFormat``, unlike a style made with ``word_styles``."""
    edit = document.styles.add(name, "character", based_on="Default Paragraph Font", color=step.text,
                               shading=step.fill)
    node = document.styles._node(document.styles.find(name, "character").id)
    for flag in node.findall(_W + "qFormat"):
        remove(flag)
    return edit


def _apply(document: Any, value: ScaleValue, step: _scales.Step, style_id: str) -> list[Any]:
    kind, where, paragraphs = value.handle

    def styled(owner: Any) -> None:
        properties = _formatting.properties_of(owner, "w:rPr", create=True)
        for tag in _HIDING:
            for node in properties.findall(_W + tag):
                remove(node)
        _formatting._set_child(properties, "w:rStyle", {"w:val": style_id})

    if kind == "range":
        paragraph, start, end = where
        return [document._format_range(TextRange(document, paragraph, start, paragraph, end), styled,
                                       marks=False)]
    cell = document.get(where)
    edits = [document.set_cell(cell._table.id, cell.row, cell.column, shading=step.fill)]
    for paragraph in cell.paragraphs:
        if paragraph.text.strip():
            edits.append(document._format_range(paragraph.range(), styled, marks=False))
    return edits


def style_by_value(call: Any, doc: str, scale: Any, preview: bool | None = None) -> Result:
    """``style_by_value``'s Word handler (its handler in :mod:`.shared`)."""
    return _scales.apply(call, scale, bool(preview), write)
