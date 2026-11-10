"""``word_edit_table``'s ``sync`` (round 49, R90): a table kept in step with a source's rows by
a key column, in one call.

The table's key column (``key``: a column's header label or its index; default the first
column taken from the source) holds one key per data row.  Each source row is matched to the
table row with its key (text compared with runs of white space as one space):

* a source key with no row is **added** -- a row as ``insert_row`` makes one beside a data row
  (its properties, cells and paragraph formatting, empty; a header row's likeness made a body
  row, as ``grow`` does), placed in **key order** where the table's data rows are in key order
  already (numbers within keys compared as numbers: ``K-9`` before ``K-10``), else in **the
  source's order** (after the row of the nearest key before it in the source; above the first
  data row when none is) -- the result says which;
* a matched row's cells that differ from the source's value are **rewritten** as ``set_cells``
  rewrites a cell; equal cells are not touched (rule 3), the key cell never;
* a table key not in the source is **missing**: listed, the row left as it is -- or, with
  ``remove``, deleted (``delete_row``); never deleted silently.

The columns taken (``from_columns``): one source column per table column, in the table's
order from its first, ``""`` for a column left as it is; by default each table column whose
header label is a source column's name (the case aside).  The header rows (marked to repeat,
else the first where the table's look has its header row on; else the first row when its
labels named the columns) and the total row (the look's last row on) are left as they are,
as are rows with no key (an empty key cell) -- counted, not matched.  A row deleted as a
tracked change is not a data row.  Values are written as ``set_cells`` ``from`` writes them
(``date_format`` too).  Duplicate keys, in the table or the source, are refused naming them.
With tracking on, every change is tracked (rows inserted and deleted, text written).

Where the values came from is kept as ``set_cells`` keeps it: when every data row of the table
is a source row afterwards, the table's record of the source (its rows in the table's order);
else each cell written keeps its value's source, and a table record the sync made untrue is
dropped (a note says so).  The result counts what was added, changed and missing or removed,
each listing up to :data:`LISTED` keys (rule 4), and says where the time went."""

from __future__ import annotations

import re
import time
from typing import Any

from ooxml_edit.tools import Result, ToolError

from ..edit import provenance as _provenance
from ._base import Timing, note_small_text, outcome
from .table_data import cell_text, placed

#: Keys each list of the result names at most (rule 4); the rest are counted.
LISTED = 20

_PARTS = re.compile(r"(\d+)")


def _norm(text: str) -> str:
    return " ".join(text.split())


def natural(key: str) -> tuple:
    """A key's sort order: its digit runs as numbers, the rest as text, the case aside."""
    return tuple((0, int(part), "") if part.isdigit() else (1, 0, part.casefold())
                 for part in _PARTS.split(key) if part)


def _capped(keys: list[str]) -> dict[str, Any]:
    out: dict[str, Any] = {"count": len(keys), "keys": keys[:LISTED]}
    if len(keys) > LISTED:
        out["more"] = len(keys) - LISTED
    return out


class _Plan:
    """What a sync reads before it changes anything."""

    def __init__(self) -> None:
        self.pairs: list[tuple[int, int]] = []     # (table column, index into block.at)
        self.key_col = 0
        self.key_pick = 0                           # the key's index into block.at
        self.body = 0                               # the first body row
        self.end = 0                                # the row after the last body row
        self.by_key: dict[str, int] = {}            # table key -> row
        self.order: list[str] = []                  # table keys, in table order
        self.keyless: list[int] = []                # body rows with no key
        self.mapped_by = "labels"


def sync(call: Any, view: Any, source: str, from_columns: list[str] | None, key: Any,
         remove: bool, date_format: str | None, notes: list[str]) -> Result:
    """``word_edit_table`` ``sync``: see the module doc."""
    from .objects import _Kept, _header_rows, _set_cell, _table_entry

    document = call.document
    timing = Timing()
    call.touch(timing)
    table = view.id
    with timing.part("resolve"):
        rows = call.source(source, field="from")
        plan = _Plan()
        header = _header_rows(document, table)
        grid = view.grid()
        height, width = len(grid), view.column_count
        fields = _mapping(rows, grid, header, width, from_columns, table, plan)
        block = placed(call, source, [rows.dataset.columns[c] for c in fields], [], date_format)
        texts = [[cell_text(lines) for lines in line] for line in block.lines()]
    started = time.perf_counter()
    _key_column(plan, key, grid, header, rows, fields, table)
    _body(document, table, plan, grid, header, height, rows, fields)
    entry = _table_entry(document, table)
    accepted = set(_provenance.accepted_rows(entry)) if entry is not None else set(range(height))
    _table_keys(plan, grid, accepted, table)
    source_keys = _source_keys(texts, plan.key_pick, block.rows.address)
    blank = [i for i, found in source_keys if found is None]
    if blank:
        shown = ", ".join(f"#{i + 1}" for i in blank[:LISTED]) + (
            f" and {len(blank) - LISTED} more" if len(blank) > LISTED else "")
        notes.append(f"{len(blank)} row(s) of {block.rows.address} with a blank key not placed: {shown}")
    # What differs: cells of matched rows, keys to add, keys gone.
    writes: list[tuple[int, int, str, int, int]] = []   # (row, col, text, source row, pick)
    changed_keys: dict[str, list[str]] = {}
    added: list[int] = []
    for i, found in source_keys:
        if found is None:
            continue
        row = plan.by_key.get(found)
        if row is None:
            added.append(i)
            continue
        for col, pick in plan.pairs:
            if col == plan.key_col:
                continue
            new = texts[i][pick]
            cell = grid[row][col]
            if cell is None or (cell.row, cell.column) != (row, col):
                if cell is not None and cell.text == new:
                    continue
                raise ToolError("refused", f"({row}, {col}) of {table} is "
                                f"{'not a cell' if cell is None else 'merged with (%d, %d)' % (cell.row, cell.column)}"
                                f": sync writes one value per cell; split it first", field="from")
            if cell.text != new:
                writes.append((row, col, new, i, pick))
                changed_keys.setdefault(found, []).append(_label(rows, fields, pick))
    wanted = {k for _, k in source_keys if k is not None}
    gone = [k for k in plan.order if k not in wanted]
    placement = "key order" if _sorted(plan.order) else "source order"
    slots = _slots(plan, source_keys, added, placement)
    timing.add("compare", time.perf_counter() - started)
    if not writes and not added and not (gone and remove):
        said = _said(block, plan, grid, header, rows, fields, [], source_keys, {}, [], False, placement, [])
        return _unchanged(call, view, block, plan, gone, {k: said[k] for k in ("key", "mapped")}, notes, timing)
    # The changes: rows added (anchors from the bottom, so the rows above keep their
    # indices), then the cells written at the rows' new places, then rows deleted.
    kept = _Kept(document, table)
    old_key = kept.key
    edits: list[Any] = []
    at_once = not document._stamping_due()
    with timing.part("rows"):
        final = _layout(height, slots)
        new_rows = _insert(document, kept, table, slots, at_once, edits)
    now = kept.current_id() or table
    with timing.part("cells"):
        cell_writes = [(final[("old", r)], c, text) for r, c, text, _, _ in writes]
        for anchor in slots:
            for i in anchor[2]:
                cell_writes += [(final[("new", i)], col, texts[i][pick]) for col, pick in plan.pairs
                                if texts[i][pick] != ""]
        cell_writes.sort(key=lambda w: (w[0], w[1]))
        if at_once and cell_writes:
            edits.extend(document._fill_cells(now, cell_writes)[0])
        else:
            for row, col, text in cell_writes:
                current = document.table(now)
                edits.extend(_set_cell(document, current, row, col, text))
    removed_rows: list[str] = []
    taken_out: list[int] = []            # rows deleted outright (a tracked deletion keeps its row)
    with timing.part("rows_removed"):
        if remove and gone:
            for row in sorted((final[("old", plan.by_key[k])] for k in gone), reverse=True):
                current = kept.current_id() or now
                made = document.delete_row(current, row)
                edits.append(made)
                if made.removed:
                    taken_out.append(row)
                removed_rows += [i for i in made.removed if _is_row(i)]
    now = kept.current_id() or now
    if taken_out:
        final = {k: v - sum(1 for row in taken_out if row < v) for k, v in final.items()}
    with timing.part("record"):
        record_note = _keep_sources(call, kept, old_key, now, block, plan, height, writes,
                                    slots, final, texts)
    started = time.perf_counter()
    data: dict[str, Any] = {"id": now, "rows": len(document.table(now).rows),
                            "columns": document.table(now).column_count}
    data.update(_said(block, plan, grid, header, rows, fields, added, source_keys, changed_keys, gone,
                      remove, placement, writes))
    data["timing"] = timing.parts
    summary = (f"sync of {now} with {block.rows.address}: {len(added)} added, {len(changed_keys)} changed"
               + ("" if not gone else f", {len(gone)} removed" if remove
                  else f", {len(gone)} missing (left; remove: true deletes them)"))
    result = outcome(edits, summary, data=data, changed=[now])
    # The rows' paragraphs and cells are the table's: named by it (rule 4); rows added and
    # removed by their ids.
    result.changed = [now]
    result.created = [i for i in new_rows]
    result.removed = removed_rows
    if now != table:
        result.renamed[table] = now
    result.warnings.extend(notes)
    if record_note:
        result.warnings.append(record_note)
    result = note_small_text(call, result, [now])
    timing.add("result", time.perf_counter() - started)
    return result


def _is_row(identifier: str) -> bool:
    return identifier.split("/")[-1].startswith("tr:")


def _labels(grid: list, row: int, width: int) -> list[str]:
    """Row ``row``'s text per grid column, a merged cell's at its first column only."""
    out = []
    for col in range(width):
        cell = grid[row][col] if row < len(grid) and col < len(grid[row]) else None
        out.append(_norm(cell.text) if cell is not None and cell.column == col else "")
    return out


def _mapping(rows: Any, grid: list, header: int, width: int, from_columns: list[str] | None,
             table: str, plan: _Plan) -> list[int]:
    """The source columns taken, in pick order; ``plan.pairs`` each table column with its pick."""
    columns = list(rows.dataset.columns)
    fields: list[int] = []
    given = [c if isinstance(c, str) else "" for c in from_columns] if from_columns is not None else None
    if given is not None and any(given):
        plan.mapped_by = "from_columns"
        if len(given) > width:
            raise ToolError("invalid_arguments", f"from_columns names {len(given)} columns; {table} has {width}: "
                            "one source column per table column, from the first, '' for one left as it is",
                            field="from_columns", valid_options=columns[:50])
        for col, name in enumerate(given):
            if not name:
                continue
            found = rows.column(name, field=f"from_columns[{col}]")
            if found in fields:
                raise ToolError("invalid_arguments", f"from_columns names {name!r} twice: one table column "
                                "per source column", field=f"from_columns[{col}]")
            plan.pairs.append((col, len(fields)))
            fields.append(found)
        return fields
    labels = _labels(grid, header - 1 if header else 0, width)
    for col, label in enumerate(labels):
        if not label:
            continue
        hits = [k for k, name in enumerate(columns) if _norm(str(name)) == label]
        if not hits:
            hits = [k for k, name in enumerate(columns) if _norm(str(name)).casefold() == label.casefold()]
        if len(hits) != 1:
            continue
        if hits[0] in fields:
            raise ToolError("ambiguous", f"two columns of {table} have the label {label!r}: say which takes "
                            f"{columns[hits[0]]!r} with from_columns (one source column per table column, "
                            "'' for one left as it is)", field="from_columns", valid_options=columns[:50])
        plan.pairs.append((col, len(fields)))
        fields.append(hits[0])
    if not fields:
        example = [str(columns[k]) if k < len(columns) else "" for k in range(min(width, len(columns)))]
        raise ToolError("not_found", f"no header label of {table} ({', '.join(repr(t) for t in labels if t) or 'none'}) "
                        f"is a column of {rows.address} ({', '.join(repr(str(c)) for c in columns[:20])}): "
                        f"give from_columns, one source column per table column, e.g. {example}",
                        field="from_columns", valid_options=[str(c) for c in columns[:50]],
                        details={"from_columns": example})
    return fields


def _label(rows: Any, fields: list[int], pick: int) -> str:
    """How a result names a column changed: the source column's name."""
    return str(rows.dataset.columns[fields[pick]])


def _key_column(plan: _Plan, key: Any, grid: list, header: int, rows: Any, fields: list[int], table: str) -> None:
    mapped = {col: pick for col, pick in plan.pairs}
    width = len(grid[0]) if grid else 0
    labels = _labels(grid, header - 1 if header else 0, width)
    named = [f"{col}: {labels[col] or rows.dataset.columns[fields[pick]]}" for col, pick in plan.pairs]
    if key is None or key == "":
        col = 0 if 0 in mapped else plan.pairs[0][0]
    else:
        said = _norm(str(key))
        col = None
        hits = [c for c in mapped if labels[c].casefold() == said.casefold()]
        if len(hits) == 1:
            col = hits[0]
        elif said.isdigit() or isinstance(key, int):
            col = int(said)
        else:
            hits = [c for c, pick in mapped.items() if str(rows.dataset.columns[fields[pick]]).casefold() == said.casefold()]
            col = hits[0] if len(hits) == 1 else None
        if col is None or col not in mapped:
            raise ToolError("not_found", f"key {key!r} is not a column sync takes from the source; the columns "
                            f"taken: {', '.join(named)}", field="key", valid_options=named)
    plan.key_col = col
    plan.key_pick = mapped[col]


def _body(document: Any, table: str, plan: _Plan, grid: list, header: int, height: int, rows: Any,
          fields: list[int]) -> None:
    """The body's rows: below the header rows (or a first row of labels), above a total row."""
    body = header
    if not header and height:
        first = grid[0][plan.key_col] if plan.key_col < len(grid[0]) else None
        name = _norm(str(rows.dataset.columns[fields[plan.key_pick]])).casefold()
        if plan.mapped_by == "labels" or (first is not None and _norm(first.text).casefold() == name):
            body = 1
    end = height
    if (document.table_properties(table).get("look") or {}).get("last_row") and height - 1 > body:
        end = height - 1
    plan.body, plan.end = body, end


def _table_keys(plan: _Plan, grid: list, accepted: set[int], table: str) -> None:
    duplicate: dict[str, list[int]] = {}
    for row in range(plan.body, plan.end):
        if row not in accepted:
            continue                 # deleted as a tracked change
        cell = grid[row][plan.key_col] if plan.key_col < len(grid[row]) else None
        text = _norm(cell.text) if cell is not None and cell.row == row else ""
        if not text:
            plan.keyless.append(row)
            continue
        if text in plan.by_key:
            duplicate.setdefault(text, [plan.by_key[text]]).append(row)
            continue
        plan.by_key[text] = row
        plan.order.append(text)
    if duplicate:
        shown = "; ".join(f"{k!r} in rows {', '.join(map(str, r))}" for k, r in list(duplicate.items())[:10])
        raise ToolError("refused", f"{table} has duplicate keys in column {plan.key_col}: {shown}"
                        + (f" and {len(duplicate) - 10} more" if len(duplicate) > 10 else "")
                        + ": a key names one row; make them unique (set_cells) or delete the extra rows "
                          "(delete_row), or sync by another key column", field="key",
                        valid_options=list(duplicate)[:50])


def _source_keys(texts: list[list[str]], pick: int, address: str) -> list[tuple[int, str | None]]:
    """Each source row's key (``None`` for a blank one); duplicates refused naming them."""
    out: list[tuple[int, str | None]] = []
    seen: dict[str, list[int]] = {}
    for i, line in enumerate(texts):
        found = _norm(line[pick].replace("\v", " "))
        out.append((i, found or None))
        if found:
            seen.setdefault(found, []).append(i)
    duplicate = {k: v for k, v in seen.items() if len(v) > 1}
    if duplicate:
        shown = "; ".join(f"{k!r} in rows {', '.join('#%d' % (r + 1) for r in v)}"
                          for k, v in list(duplicate.items())[:10])
        raise ToolError("refused", f"{address} has duplicate keys: {shown}"
                        + (f" and {len(duplicate) - 10} more" if len(duplicate) > 10 else "")
                        + ": a key names one row; derive a dataset with one row per key, or sync by another "
                          "key column", field="from", valid_options=list(duplicate)[:50])
    return out


def _sorted(keys: list[str]) -> bool:
    """Whether the table's keys are in key order (two at least)."""
    if len(keys) < 2:
        return False
    order = [natural(k) for k in keys]
    return all(a <= b for a, b in zip(order, order[1:]))


def _slots(plan: _Plan, source_keys: list[tuple[int, str | None]], added: list[int],
           placement: str) -> list[tuple[int, bool, list[int]]]:
    """Where the rows added go: ``(row, below, source rows)`` -- beside table row ``row``
    (below it unless not ``below``), in that order; one slot per row."""
    if not added:
        return []
    key_of = dict(source_keys)
    data_rows = [plan.by_key[k] for k in plan.order]
    slots: dict[tuple[int, bool], list[int]] = {}
    if not data_rows:
        # No data row: after the body's last row (the header's, or a row with no key).
        anchor = max(plan.end - 1, 0)
        slots[(anchor, True)] = list(added)
    elif placement == "key order":
        for i in sorted(added, key=lambda i: natural(key_of[i])):
            mine = natural(key_of[i])
            after = next((r for k, r in zip(plan.order, data_rows) if natural(k) > mine), None)
            place = (after, False) if after is not None else (data_rows[-1], True)
            slots.setdefault(place, []).append(i)
    else:
        last: int | None = None
        for i, found in source_keys:
            if found is None:
                continue
            if found in plan.by_key:
                last = plan.by_key[found]
            elif i in added:
                place = (last, True) if last is not None else (data_rows[0], False)
                slots.setdefault(place, []).append(i)
    return [(row, below, rows) for (row, below), rows in sorted(slots.items(), key=lambda s: s[0][0])]


def _layout(height: int, slots: list[tuple[int, bool, list[int]]]) -> dict[tuple[str, int], int]:
    """Each row's index after the rows are added: ``("old", row)`` and ``("new", source row)``."""
    before = {row: rows for row, below, rows in slots if not below}
    after = {row: rows for row, below, rows in slots if below}
    out: dict[tuple[str, int], int] = {}
    k = 0
    for row in range(height):
        for i in before.get(row, ()):
            out[("new", i)] = k
            k += 1
        out[("old", row)] = k
        k += 1
        for i in after.get(row, ()):
            out[("new", i)] = k
            k += 1
    return out


def _insert(document: Any, kept: Any, table: str, slots: list[tuple[int, bool, list[int]]], at_once: bool,
            edits: list[Any]) -> list[str]:
    """The rows added, slot by slot from the bottom: each as ``insert_row`` makes it beside
    the slot's row, the next below the one before, a header row's likeness made a body row
    -- in one pass where nothing needs stamping (``Document._rows_beside``), the same edits."""
    from .objects import _is_header_row

    places = [(row, below, len(rows)) for row, below, rows in reversed(slots)]
    if at_once:
        made = document._rows_beside(table, places)
        edits.extend(made)
        return [edit.id for edit in made if edit.created]
    created: list[str] = []
    current = table
    for row, below, count in places:
        index, side = row, below
        for _ in range(count):
            edit = document.insert_row(current, index, below=side)
            edits.append(edit)
            created.append(edit.id)
            current = kept.current_id() or current      # a row added above the first renames it
            index, side = (index + 1 if side else index), True
            if _is_header_row(document, current, index):
                edits.append(document.set_row(current, index, header=False))
    return created


def _keep_sources(call: Any, kept: Any, old_key: str | None, now: str, block: Any, plan: _Plan, height: int,
                  writes: list, slots: list, final: dict, texts: list[list[str]]) -> str | None:
    """Where the values came from (see the module doc); a note when a record is dropped."""
    from ooxml_edit.tools import textfacts

    from .objects import _table_entry

    document = call.document
    entry = _table_entry(document, now)
    if entry is None:
        return None
    rows = block.rows
    accepted = _provenance.accepted_rows(entry)
    view = document.table(now)
    grid = view.grid()
    end = len(grid) - (height - plan.end)        # above the total row
    index_of = {}
    for i, line in enumerate(texts):
        found = _norm(line[plan.key_pick].replace("\v", " "))
        if found:
            index_of[found] = i
    body = [r for r in range(plan.body, end) if r in accepted]
    order: list[int] = []
    for row in body:
        cell = grid[row][plan.key_col]
        found = index_of.get(_norm(cell.text)) if cell is not None and cell.row == row else None
        if found is None:
            order = []
            break
        order.append(found)
    if order and len(order) == len(texts):
        first = accepted.index(body[0])
        saved = {**rows.describe(), "header_row": None, "first_row": first,
                 "columns": [{"col": col, **rows.provenance(block.at[pick], order)} for col, pick in plan.pairs]}
        if old_key is not None and old_key != kept._key_now():
            call.store_provenance(old_key, None)
        call.keep_provenance(now, rows, saved)
        return None
    note = None
    if kept.record is not None and old_key is not None:
        call.store_provenance(old_key, None)
        note = (f"the table's record of where its values came from ({kept.record.get('source') or 'a source'}) "
                "no longer holds for every row: dropped; each cell written keeps its value's source")
    base, first_row = _base(rows.address)
    keeps: list[tuple[str, list[dict]]] = []
    cells: list[tuple[int, int, int, int]] = [(final[("old", r)], c, i, pick) for r, c, _, i, pick in writes]
    for _, _, added in slots:
        cells += [(final[("new", i)], col, i, pick) for i in added for col, pick in plan.pairs if texts[i][pick]]
    for row, col, i, pick in cells:
        column = block.at[pick]
        text = texts[i][pick]
        address = f"{base}/#{first_row + i}/{textfacts.quote(str(rows.dataset.columns[column]))}"
        record = textfacts.record(rows, i, column, address, text, None)
        found = grid[row][col]
        for paragraph in found.paragraphs if found is not None else ():
            if paragraph.id.split("/")[-1].startswith("p:") and paragraph.text:
                keeps.append((paragraph.id, [{**record, "text": paragraph.text, "p": 0, "at": 0}]))
    if keeps:
        document.keep_text_facts_in(keeps)
    return note


def _base(address: str) -> tuple[str, int]:
    """A source's address without a row range, and its first row's number (from 1)."""
    found = re.match(r"^(.*?)/#(\d+)(?::#?\d*)?$", address)
    if found:
        return found.group(1), int(found.group(2))
    return address, 1


def _said(block: Any, plan: _Plan, grid: list, header: int, rows: Any, fields: list[int], added: list[int],
          source_keys: list, changed_keys: dict[str, list[str]], gone: list[str], remove: bool,
          placement: str, writes: list) -> dict[str, Any]:
    key_of = dict(source_keys)
    columns = rows.dataset.columns
    out: dict[str, Any] = {
        "key": {"col": plan.key_col, "column": str(columns[fields[plan.key_pick]])},
        "mapped": {str(col): str(columns[fields[pick]]) for col, pick in plan.pairs},
        "added": {**_capped([key_of[i] for i in added]), "placed": placement},
        "changed": {"count": len(changed_keys), "cells": len(writes),
                    "keys": {k: v for k, v in list(changed_keys.items())[:LISTED]}},
        ("removed" if remove else "missing"): _capped(gone),
        "unchanged": len(plan.order) - len(changed_keys) - len(gone),
    }
    if len(changed_keys) > LISTED:
        out["changed"]["more"] = len(changed_keys) - LISTED
    if not added:
        out["added"].pop("placed")
    if plan.keyless:
        out["no_key_rows"] = len(plan.keyless)
    return out


def _unchanged(call: Any, view: Any, block: Any, plan: _Plan, gone: list[str], said: dict[str, Any],
               notes: list[str], timing: Timing) -> Result:
    """Rule 3: the table holds the source's values already -- nothing written."""
    data: dict[str, Any] = {"id": view.id, "rows": len(view.rows), "columns": view.column_count, **said,
                            "added": {"count": 0, "keys": []}, "changed": {"count": 0, "cells": 0, "keys": {}},
                            "missing": _capped(gone), "unchanged": len(plan.order) - len(gone),
                            "timing": timing.parts}
    if plan.keyless:
        data["no_key_rows"] = len(plan.keyless)
    result = Result(summary=f"sync of {view.id} with {block.rows.address}: nothing to change"
                    + (f"; {len(gone)} missing (left; remove: true deletes them)" if gone else ""), data=data)
    result.warnings.append("the table holds the source's values already: nothing changed")
    result.warnings.extend(notes)
    return result
