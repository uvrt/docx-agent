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
labels named the columns) and the total rows (the look's last row on; any row holding a
formula field, ``=SUM(ABOVE)``) are left as they are, as are rows with no key (an empty key
cell) -- counted, not matched.  A row deleted as a tracked change is not a data row.  Values
are written as ``set_cells`` ``from`` writes them (``date_format`` too).  Keys are compared
with their case, white space runs aside: a table key and a source key that would match
loosely are said in a note.

Keys spelled differently (round 52, R106).  ``match`` pairs a table key with a source key
(``[{table, data}]``, or one object ``{table key: data key}``) for keys no rule should guess:
applied first; a key it names that is not one is refused with the keys there are (a pair
written the other way round is read as meant, with a note).  Then exact matches; then, with
``loose``, the keys left on both sides by :func:`ooxml_edit.tools.loose_key` (case, white
space, the separators ``. - _ /`` between runs, numbers as numbers: ``K7-12.3`` =
``k7_12_03``).  A key that would loosely match more than one key on the other side is
refused before any change, naming them (the way out: ``match``, or the data fixed).  A row
matched by ``match`` or loosely keeps the table's spelling of its key (``loose: "table"``,
and without ``loose``) or takes the source's (``loose: "data"``: the key cell rewritten, a
tracked change when tracking is on).  The result's ``loose_matches`` lists the loose pairs
(table key: source key), ``paired`` counts ``match``'s.  (xlsx-agent's filter ``loose`` is
another, smaller rule -- white space and case in a cell's text -- for matching text, not keys.)
Duplicate keys, in the table or the source, are refused naming them; rows to add like a row
whose one cell spans two columns taken are refused before any change, with the way out.
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

from ooxml_edit.tools import Result, ToolError, loose_key

from ..edit import provenance as _provenance
from ._base import Timing, note_small_text, outcome
from .table_data import cell_text, placed

#: Keys each list of the result names at most (rule 4); the rest are counted.
LISTED = 20

_PARTS = re.compile(r"(\d+)")
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


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
        self.totals: list[int] = []                 # body rows computed by a formula field
        self.mapped_by = "labels"


def sync(call: Any, view: Any, source: str, from_columns: list[str] | None, key: Any,
         remove: bool, date_format: str | None, notes: list[str], *, loose: str | None = None,
         match: list[tuple[str, str]] = ()) -> Result:
    """``word_edit_table`` ``sync``: see the module doc."""
    from .objects import _Accepted, _accepted_grid, _Kept, _header_rows, _set_cell, _table_entry

    document = call.document
    timing = Timing()
    call.touch(timing)
    table = view.id
    with timing.part("resolve"):
        rows = call.source(source, field="from")
        plan = _Plan()
        header = _header_rows(document, table)
        # Round 54 (lane M): the table's columns as accepted (a pending tracked column
        # deletion not counted); written back at the document's grid columns.
        as_accepted = _Accepted(document, table)
        grid = _accepted_grid(view.grid(), as_accepted)
        height, width = len(grid), as_accepted.width
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
    # Which table row each source key names: pairs given, then exact, then loose (R106).
    row_of, paired, loosely = _matches(plan, source_keys, match, loose, block.rows.address, table, notes)
    respell = {t: d for t, d in [*paired, *loosely]} if loose == "data" else {}
    # What differs: cells of matched rows, keys to add, keys gone.
    writes: list[tuple[int, int, str, int, int]] = []   # (row, col, text, source row, pick)
    changed_keys: dict[str, list[str]] = {}
    added: list[int] = []
    table_key = {row: k for k, row in plan.by_key.items()}
    for i, found in source_keys:
        if found is None:
            continue
        row = row_of.get(found)
        if row is None:
            added.append(i)
            continue
        named = table_key[row] if table_key[row] not in respell else found
        for col, pick in plan.pairs:
            if col == plan.key_col and table_key[row] not in respell:
                continue
            new = texts[i][pick]
            cell = grid[row][col]
            if cell is None or (cell.row, cell.column) != (row, col):
                if cell is not None and cell.text == new:
                    continue
                if cell is not None and cell.row == row:
                    _merged_refusal(f"row {row} of {table} (key {named!r})", row, grid, plan, rows, fields, table)
                raise ToolError("refused", f"({row}, {col}) of {table} is "
                                f"{'not a cell' if cell is None else 'merged with (%d, %d)' % (cell.row, cell.column)}"
                                f": sync writes one value per cell; split it first", field="from")
            if cell.text != new and (col != plan.key_col or _norm(cell.text) != found):
                writes.append((row, col, new, i, pick))
                changed_keys.setdefault(named, []).append(_label(rows, fields, pick))
    taken = set(row_of.values())
    gone = [k for k in plan.order if plan.by_key[k] not in taken]
    placement = "key order" if _sorted(plan.order) else "source order"
    slots = _slots(plan, source_keys, added, placement, row_of)
    _merged_likeness(slots, grid, plan, rows, fields, table)
    key_of = dict(source_keys)
    if not loose:
        _note_loose(gone, [key_of[i] for i in added], remove, notes)
    timing.add("compare", time.perf_counter() - started)
    if not writes and not added and not (gone and remove):
        said = _said(block, plan, grid, header, rows, fields, [], source_keys, {}, [], False, placement, [],
                     loose, paired, loosely)
        return _unchanged(call, view, block, plan, gone,
                          {k: said[k] for k in ("key", "mapped", "loose_matches", "paired") if k in said},
                          notes, timing)
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
        if as_accepted.active:
            cell_writes = [(row, as_accepted.kept[col], text) for row, col, text in cell_writes]
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
                                    slots, final, texts, row_of)
    started = time.perf_counter()
    data: dict[str, Any] = {"id": now, "rows": _accepted_count(document, now),
                            "columns": len(document.accepted_columns(now)[0])}
    data.update(_said(block, plan, grid, header, rows, fields, added, source_keys, changed_keys, gone,
                      remove, placement, writes, loose, paired, loosely))
    data["timing"] = timing.parts
    summary = (f"sync of {now} with {block.rows.address}: {len(added)} added, {len(changed_keys)} changed"
               + ("" if not gone else f", {len(gone)} removed" if remove
                  else f", {len(gone)} missing (left; remove: true deletes them)")
               + (f"; {len(loosely)} key(s) matched loosely" if loosely else ""))
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


def _accepted_count(document: Any, table: str) -> int:
    """The table's rows as word_edit_table counts them (round 54, lane R): a row whose
    tracked deletion is pending not counted."""
    from .objects import _table_entry

    entry = _table_entry(document, table)
    return len(_provenance.accepted_rows(entry)) if entry is not None else len(document.table(table).rows)


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
    # A row holding a formula field (=SUM(ABOVE), any = field) computes over the table: a
    # total row, never matched, added to or removed; those at the bottom end the body.
    _part, entry = document._resolve(table)
    plan.totals = [r for r in range(body, end) if _formula_row(entry.rows[r][0])]
    while end - 1 > body and end - 1 in plan.totals:
        end -= 1
    plan.body, plan.end = body, end


def _formula_row(row: Any) -> bool:
    """Whether a table row holds a formula field: a ``w:fldSimple`` or a field's instruction
    text starting with ``=``."""
    for node in row.iter(_W + "fldSimple", _W + "instrText"):
        said = node.get(_W + "instr") if node.tag == _W + "fldSimple" else node.text
        if said and said.lstrip().startswith("="):
            return True
    return False


def _table_keys(plan: _Plan, grid: list, accepted: set[int], table: str) -> None:
    duplicate: dict[str, list[int]] = {}
    for row in range(plan.body, plan.end):
        if row not in accepted or row in plan.totals:
            continue                 # deleted as a tracked change, or a total row
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
        # Rows as positions count them (round 54, lane R: a pending deleted row not counted).
        order = sorted(accepted)
        shown = "; ".join(f"{k!r} in rows {', '.join(str(order.index(x)) for x in r)}"
                          for k, r in list(duplicate.items())[:10])
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


def _matches(plan: _Plan, source_keys: list[tuple[int, str | None]], match: Any, loose: str | None,
             address: str, table: str, notes: list[str]
             ) -> tuple[dict[str, int], list[tuple[str, str]], list[tuple[str, str]]]:
    """Which table row each source key names (round 52, R106): ``match``'s pairs first, then
    the keys equal as they are, then (``loose``) the keys left equal by :func:`loose_key`.
    Returns the rows by source key, the pairs given and the loose pairs (table key, source
    key).  Refuses, before any change, a pair naming a key that is not one, a key paired
    twice, and a key that would loosely match more than one key on the other side."""
    wanted = [k for _, k in source_keys if k is not None]
    in_source = set(wanted)
    row_of: dict[str, int] = {}
    paired: list[tuple[str, str]] = []
    used: set[str] = set()              # table keys matched
    for n, (said_table, said_data) in enumerate(match):
        t, d = _norm(str(said_table)), _norm(str(said_data))
        if t not in plan.by_key and d in plan.by_key and t in in_source:
            notes.append(f"match pair {n + 1} ({said_table!r}: {said_data!r}) read the other way round: "
                         f"{d!r} is the table's key, {t!r} the source's")
            t, d = d, t
        if t not in plan.by_key or d not in in_source:
            bad, keys, where = (t, plan.order, f"a key of {table}") if t not in plan.by_key else (
                d, wanted, f"a key of {address}")
            near = [k for k in keys if loose_key(k) == loose_key(bad)]
            raise ToolError("not_found", f"match pair {n + 1}: {bad!r} is not {where} (key column "
                            f"{plan.key_col}); pairs are {{table: table key, data: source key}}"
                            + (f"; loosely equal: {', '.join(map(repr, near[:10]))}" if near else ""),
                            field=f"match[{n}]", valid_options=(near + [k for k in keys if k not in near])[:50])
        if t in used or d in row_of:
            raise ToolError("invalid_arguments", f"match pair {n + 1}: {t if t in used else d!r} is paired "
                            "twice: a key names one row", field=f"match[{n}]")
        used.add(t)
        row_of[d] = plan.by_key[t]
        if t != d:
            paired.append((t, d))
    for k in wanted:
        if k not in row_of and k in plan.by_key and k not in used:
            row_of[k] = plan.by_key[k]
            used.add(k)
    loosely: list[tuple[str, str]] = []
    if not loose:
        return row_of, paired, loosely
    tables: dict[tuple, list[str]] = {}
    datas: dict[tuple, list[str]] = {}
    for k in plan.order:
        if k not in used and loose_key(k):
            tables.setdefault(loose_key(k), []).append(k)
    for k in wanted:
        if k not in row_of and loose_key(k):
            datas.setdefault(loose_key(k), []).append(k)
    clashes: list[str] = []
    for form, ours in tables.items():
        theirs = datas.get(form)
        if not theirs:
            continue
        if len(ours) > 1 or len(theirs) > 1:
            clashes.append(f"{', '.join(map(repr, ours))} (table) / {', '.join(map(repr, theirs))} (source)")
            continue
        row_of[theirs[0]] = plan.by_key[ours[0]]
        loosely.append((ours[0], theirs[0]))
    if clashes:
        raise ToolError("ambiguous", f"{len(clashes)} key(s) loosely match more than one key on the other side: "
                        + "; ".join(clashes[:10]) + (f" and {len(clashes) - 10} more" if len(clashes) > 10 else "")
                        + ": pair them with match ({table key: source key}) or make the keys unique; "
                          "nothing was changed", field="loose", valid_options=["match"])
    return row_of, paired, loosely


def _note_loose(gone: list[str], added: list[str], remove: bool, notes: list[str]) -> None:
    """Without ``loose``: say which table keys gone and source keys added would match loosely
    -- the way out (rule 2)."""
    forms: dict[tuple, list[str]] = {}
    for k in gone:
        forms.setdefault(loose_key(k), []).append(k)
    pairs = [(forms[loose_key(k)][0], k) for k in added if len(forms.get(loose_key(k)) or ()) == 1]
    if not pairs:
        return
    shown = ", ".join(f"{a!r} / {b!r}" for a, b in pairs[:LISTED]) + (
        f" and {len(pairs) - LISTED} more" if len(pairs) > LISTED else "")
    notes.append(f"{len(pairs)} table key(s) match a source key loosely (case, spaces, separators, leading zeros) "
                 f"but were compared exactly (the source's added, the table's {'removed' if remove else 'missing'}): "
                 f"{shown} (table / source); loose: 'table' matches them, keeping the table's spelling")


def _merged_likeness(slots: list, grid: list, plan: _Plan, rows: Any, fields: list[int], table: str) -> None:
    """Refuse, before any change, rows to add like a row whose one cell spans two or more of
    the columns taken: a value each would go in one cell (rule 2: the way out)."""
    for row, _below, _rows in slots:
        _merged_refusal(f"new rows are made like row {row} of {table}", row, grid, plan, rows, fields, table)


def _merged_refusal(what: str, row: int, grid: list, plan: _Plan, rows: Any, fields: list[int], table: str) -> None:
    """Refuse where one cell of ``row`` spans two or more columns taken: which, and the
    ``from_columns`` that takes only the first of them (rule 2)."""
    by_cell: dict[tuple[int, int], list[int]] = {}
    for col, _pick in plan.pairs:
        cell = grid[row][col] if col < len(grid[row]) else None
        if cell is not None:
            by_cell.setdefault((cell.row, cell.column), []).append(col)
    for origin, cols in by_cell.items():
        if len(cols) < 2:
            continue
        picks = dict(plan.pairs)
        names = [str(rows.dataset.columns[fields[picks[c]]]) for c in cols]
        example = ["" for _ in range(max(c for c, _ in plan.pairs) + 1)]
        for c, pick in plan.pairs:
            if c not in cols[1:]:
                example[c] = str(rows.dataset.columns[fields[pick]])
        raise ToolError("refused", f"{what}: its cell ({origin[0]}, {origin[1]}) is merged across columns "
                        f"{', '.join(map(str, cols))}, each taken ({', '.join(names)}), one value each: take only "
                        f"one of them -- '' for the others in from_columns, e.g. {example} -- or split the cell "
                        "first (nothing was changed)", field="from_columns", valid_options=names,
                        details={"from_columns": example})


def _sorted(keys: list[str]) -> bool:
    """Whether the table's keys are in key order (two at least)."""
    if len(keys) < 2:
        return False
    order = [natural(k) for k in keys]
    return all(a <= b for a, b in zip(order, order[1:]))


def _slots(plan: _Plan, source_keys: list[tuple[int, str | None]], added: list[int],
           placement: str, row_of: dict[str, int]) -> list[tuple[int, bool, list[int]]]:
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
            if found in row_of:
                last = row_of[found]
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
                  writes: list, slots: list, final: dict, texts: list[list[str]],
                  row_of: dict[str, int]) -> str | None:
    """Where the values came from (see the module doc); a note when a record is dropped."""
    from ooxml_edit.tools import textfacts

    from .objects import _Accepted, _accepted_grid, _table_entry

    document = call.document
    entry = _table_entry(document, now)
    if entry is None:
        return None
    rows = block.rows
    accepted = _provenance.accepted_rows(entry)
    view = document.table(now)
    grid = _accepted_grid(view.grid(), _Accepted(document, now))   # columns as accepted (lane M)
    end = len(grid) - (height - plan.end)        # above the total row
    # Each row's source row, by where the rows are now: matched rows, then the rows added.
    index_of: dict[int, int] = {}
    for i, line in enumerate(texts):
        found = _norm(line[plan.key_pick].replace("\v", " "))
        if found in row_of:
            index_of[final[("old", row_of[found])]] = i
        elif ("new", i) in final:
            index_of[final[("new", i)]] = i
    body = [r for r in range(plan.body, end) if r in accepted]
    order: list[int] = []
    for row in body:
        found = index_of.get(row)
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
          placement: str, writes: list, loose: str | None = None, paired: list = (),
          loosely: list = ()) -> dict[str, Any]:
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
    if loose:
        out["loose_matches"] = {"count": len(loosely), "pairs": dict(loosely[:LISTED])}
        if len(loosely) > LISTED:
            out["loose_matches"]["more"] = len(loosely) - LISTED
    if paired:
        out["paired"] = len(paired)
    if plan.keyless:
        out["no_key_rows"] = len(plan.keyless)
    if plan.totals:
        out["total_rows"] = len(plan.totals)
    return out


def _unchanged(call: Any, view: Any, block: Any, plan: _Plan, gone: list[str], said: dict[str, Any],
               notes: list[str], timing: Timing) -> Result:
    """Rule 3: the table holds the source's values already -- nothing written."""
    data: dict[str, Any] = {"id": view.id, "rows": _accepted_count(call.document, view.id),
                            "columns": len(call.document.accepted_columns(view.id)[0]), **said,
                            "added": {"count": 0, "keys": []}, "changed": {"count": 0, "cells": 0, "keys": {}},
                            "missing": _capped(gone), "unchanged": len(plan.order) - len(gone),
                            "timing": timing.parts}
    for name in ("loose_matches", "paired"):     # after the counts, as a sync that changes says them
        if name in data:
            data[name] = data.pop(name)
    if plan.keyless:
        data["no_key_rows"] = len(plan.keyless)
    if plan.totals:
        data["total_rows"] = len(plan.totals)
    loosely = (said.get("loose_matches") or {}).get("count")
    result = Result(summary=f"sync of {view.id} with {block.rows.address}: nothing to change"
                    + (f"; {len(gone)} missing (left; remove: true deletes them)" if gone else "")
                    + (f"; {loosely} key(s) matched loosely" if loosely else ""), data=data)
    result.warnings.append("the table holds the source's values already: nothing changed")
    result.warnings.extend(notes)
    return result
