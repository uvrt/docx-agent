"""Tables from data (X1): a source's rows as table text, for ``word_insert_markdown``'s
``from`` (a new table) and ``word_edit_table``'s ``set_cells`` ``from`` (an existing one).

A source is a dataset (``d:<id>``) or a table of another open document (``t:<name>``),
read with ``call.source`` (:mod:`ooxml_edit.tools.sources`) as ``edit_chart`` reads its
rows.  ``from_columns`` picks the source's columns by name, in the order placed (default:
every column, in the source's order); every row is placed, in the source's order.  Above
the rows goes a header row: the column names as the source has them, unless ``header``
gives its texts, or ``[]`` for none.  A value is written as a chart's category label is
(:meth:`ooxml_edit.tools.sources.Rows.label`): text as it is; a date in the session
locale's day order (``3/1/2026`` en-US, ``1/3/2026`` nl-BE) -- a Python date, and the ISO
8601 text of a ``date``, ``datetime`` or ``time`` column (a workbook's date cells, a
derive's dates), so one document never mixes ``1/3/2026`` and ``2026-01-20``; a number as
the General format writes it with the session locale's decimal separator, a logical value ``TRUE`` or
``FALSE``, a blank as an empty cell; a value's line breaks are line breaks in the cell's
one paragraph.  ``date_format`` writes the dates of the columns placed in a pattern of
Excel's date tokens instead (``d mmmm yyyy`` is ``2 March 2026``, ``2 maart 2026`` in
nl-BE; ``dd/mm/yyyy`` ``02/03/2026``; month and day names measured in Excel for Mac,
2026-10-16 -- :mod:`ooxml_edit.dates`), as fixed text: the pattern's separators in every
locale.  Nothing is cut: a block over Word's 63 columns, 4,096 rows or 5,000 cells
per call, or bigger than the table it is written into, is refused.

``d:`` is also how a Word document names a drawing (``d:3``); a ``from`` only ever names a
dataset (whose id starts with a letter or ``_``), and :func:`source_hint` says so when a
drawing's id is given.
"""

from __future__ import annotations

import re
from typing import Any

from ooxml_edit.tools import ToolError, array, string
from ooxml_edit.tools import sources

from ..edit.text import LINE_BREAK

#: The most a table from data takes: columns (Word's), rows (the header included) and cells
#: (one call's, as ``set_cells`` takes).
MAX_ROWS, MAX_COLUMNS, MAX_CELLS = 4096, 63, 5000

FROM_COLUMNS = array(string(), "from's columns, in order. Default all.", optional=True)
HEADER = array(string(), "from's header texts; [] none. Default column names.", optional=True)
DATE_FORMAT = string("Dates of from as e.g. d mmmm yyyy or dd/mm/yyyy.", optional=True)

_LINES = re.compile(r"\r\n|[\r\n\v]")
_INVALID = re.compile("[\x00-\x08\x0c\x0e-\x1f￾￿]")
_DRAWING = re.compile(r"^d:[0-9]+(#[0-9]+)?$")


class Placed:
    """A source's rows as table text: the header (if any), then a row per source row."""

    def __init__(self, rows: Any, at: list[int], header: list[str], locale: Any,
                 date_format: str | None = None) -> None:
        self.rows, self.at, self.header, self.locale = rows, at, header, locale
        self.date_format = date_format
        #: The block's size: the header row and a row per source row; a column per pick.
        self.size = ((1 if header else 0) + len(rows.dataset.rows), len(at))

    def lines(self) -> list[list[list[str]]]:
        """The block's cells, row by row (the header first), each as its lines."""
        out = [[_lines(text) for text in self.header]] if self.header else []
        return out + [[_lines(_text(self.rows, row, column, self.locale, self.date_format))
                       for column in self.at]
                      for row in range(len(self.rows.dataset.rows))]

    def provenance(self, top: int) -> dict[str, Any]:
        """``data.from`` for the block written with its first row at table row ``top``."""
        order = list(range(len(self.rows.dataset.rows)))
        return {**self.rows.describe(), "header_row": top if self.header else None,
                "first_row": top + (1 if self.header else 0),
                "columns": [{"col": col, **self.rows.provenance(column, order)}
                            for col, column in enumerate(self.at)]}


def no_source(from_columns: Any, header: Any, date_format: Any = None) -> None:
    """Refuse ``from_columns``, ``header`` or ``date_format`` without a ``from``."""
    for name, value in (("from_columns", from_columns), ("header", header),
                        ("date_format", date_format)):
        if value is not None:
            raise ToolError("invalid_arguments", f"{name} is for from; give from too", field=name)


def placed(call: Any, source: str, from_columns: list[str] | None,
           header: list[str] | None, date_format: str | None = None) -> Placed:
    """The source's rows and the columns picked; refusals name the argument.  ``date_format``
    writes the dates of the columns picked in that pattern
    (:func:`ooxml_edit.tools.sources.date_pattern`), refused when none of them holds a
    date."""
    rows = call.source(source, field="from")
    names = from_columns if from_columns is not None else list(rows.dataset.columns)
    if not names:
        raise ToolError("invalid_arguments", "from_columns names no column; leave it out for "
                        "every column", field="from_columns")
    at = [rows.column(name, field=f"from_columns[{i}]") for i, name in enumerate(names)]
    if len(set(at)) != len(at):
        raise ToolError("invalid_arguments", "a column is named twice", field="from_columns")
    if header is None:
        header = [rows.dataset.columns[column] for column in at]
    elif header and len(header) != len(at):
        raise ToolError("invalid_arguments", f"{len(header)} header texts for {len(at)} columns; "
                        "[] for no header row", field="header")
    if date_format is not None:
        sources.date_pattern(date_format, call.locale)
        count = len(rows.dataset.rows)
        if not any(rows.date(row, column) is not None for column in at for row in range(count)):
            raise sources.no_dates(rows, at)
    return Placed(rows, at, list(header), call.locale, date_format)


def _text(rows: Any, row: int, column: int, locale: Any, date_format: str | None = None) -> str:
    """A source value as cell text, as a chart's category label writes it (a date in
    ``date_format`` when given); a blank is empty; anything else (not a finite number) is
    refused, naming the cell."""
    value = rows.dataset.rows[row][column]
    if value is None or value == "":
        return ""
    text = rows.label(row, column, locale, date_format)
    if text is None:
        raise ToolError("invalid_arguments", f"{rows.where(row, column)} is "
                        f"{sources.what(value)}: a cell takes text, a number, a logical value or "
                        "a blank", field="from")
    return text


def _lines(text: str) -> list[str]:
    """A cell's text as its lines; characters XML cannot hold as U+FFFD."""
    return [_INVALID.sub("�", line) for line in _LINES.split(text)]


def cell_text(lines: list[str]) -> str:
    """A cell's lines as one paragraph's text, line breaks between them."""
    return LINE_BREAK.join(lines)


def fits(block: Placed, max_rows: int, max_columns: int, room: str | None = None) -> None:
    """Refuse a block bigger than ``max_rows`` x ``max_columns``, saying by how much: a
    table's limits (``limit``), or (``room``) the room an existing table has (``refused``)."""
    address = block.rows.address
    need_rows, need_columns = block.size
    code = "refused" if room else "limit"
    if need_rows > max_rows:
        header = " with the header" if block.header else ""
        # A row range of the source places fewer: offered when the source can take one.
        fewer = sources.first_rows(address, max_rows - (1 if block.header else 0))
        fix = (f"{room}, or insert_row first" if room
               else f"a table takes at most {max_rows:,}: filter the source or split it across "
                    "tables") + (f", or place fewer with {fewer}" if fewer else "")
        raise ToolError(code, f"{address} gives {need_rows:,} rows{header}; {fix} (nothing is "
                        "cut)", field="from", valid_options=[fewer] if fewer else None)
    if need_columns > max_columns:
        fix = f"the table has {max_columns}" if room else f"a table takes at most {max_columns}"
        raise ToolError(code, f"{address}: {need_columns} columns picked; {fix}: pick fewer in "
                        "from_columns", field="from_columns")
    if need_rows * need_columns > MAX_CELLS:
        raise ToolError("limit", f"{address} gives {need_rows:,}x{need_columns} = "
                        f"{need_rows * need_columns:,} cells; at most {MAX_CELLS:,} in one call: "
                        "filter the source or pick fewer columns (nothing is cut)", field="from")


def source_hint(document: Any, value: str) -> str | None:
    """The docx format's ``source_hint``: a ``from`` given a drawing's id (``d:3``, the
    form Word documents name drawings by) is told it is a drawing, not a dataset."""
    if not isinstance(value, str) or not _DRAWING.match(value):
        return None
    try:
        document.drawing(value)
    except (KeyError, ValueError):
        return (f"{value} is a drawing's id (none here), not a dataset: a source is a dataset "
                "d:<id>, its id starting with a letter, or a table t:<name>")
    return (f"{value} is a drawing in this document, not a dataset: a source is a dataset "
            "d:<id>, its id starting with a letter, or a table t:<name>")
