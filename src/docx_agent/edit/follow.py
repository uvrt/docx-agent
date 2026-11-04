"""Copies made to follow a revised sample, in one edit (round 31, R8).

A table (or a paragraph) copied from a sample and filled with its own values drifts from the
sample when the sample is revised: :mod:`.compare` lists the differences, and this module
makes them go away -- each copy takes the sample's formatting and keeps its own text and
values.  :meth:`Document.follow_sample <docx_agent.Document.follow_sample>` is one undo step
for every copy; ``preview`` reports the same without changing anything.

**Made from a sample.**  docx-agent keeps no record of where a table was copied from: a copy
(``word_copy_from``, Word's own copy and paste, a Markdown table formatted by hand) has ids of
its own, and nothing in the file ties it to its sample.  So the copies are named, a list of
block ids (``word_read``'s); a copy that is not like its sample (another kind of block, a
table with other columns) is reported, not forced.

**Addresses** as the other tools take them: a cell or row address (``t:X/c1,2``, ``t:X/r1``)
or a paragraph in a table names that table (for the sample, and for copies of a table); an
empty copy address, a copy named twice and the sample itself among its copies are ignored
with a note -- nothing to do is no error.  An address that names nothing is refused with the
document's tables and their rows and columns, those shaped like the sample first.

**Corresponding parts** are :mod:`.compare`'s: rows by the sample's structure (its header rows
one to one, body row *i* to the sample's body row *i*, its last body row the pattern for any
more), cells by grid column, a cell's paragraphs aligned by style and text.

**What a table copy takes from its sample** -- the sample's own properties, as its XML states
them, not a reading of them:

* the table's properties (``w:tblPr``): style, width, alignment, indent, borders, shading,
  layout, cell margins and spacing, style options, floating -- but its alt text (title and
  description), which is the copy's own;
* the grid's column widths, where the copy has as many columns as the sample;
* each row's properties (``w:trPr``): height, may not break, header row, alignment, hidden --
  but its tracked insertion or deletion and its cells before or after the grid (``gridBefore``
  and the like, its structure);
* each cell's properties (``w:tcPr``) from the sample's cell over the same grid columns: width,
  borders, shading, margins, vertical alignment, text direction, no wrap, fit text -- but its
  merges (``gridSpan``, ``vMerge``, ``hMerge``) and change records; a cell the sample has
  merged otherwise follows nothing;
* each paired paragraph's properties (``w:pPr``): style, alignment, spacing, indents, keeps,
  borders, shading, tabs, the paragraph mark's formatting -- but its list membership
  (``w:numPr``: the sample's list is another list, and joining it would continue its
  numbering), its section break and change records;
* each paired paragraph's run formatting, **by property**: for each property of a run (bold,
  size, font, colour, character style, ...), the value most of the sample paragraph's text
  has.  Where the copy's text mostly has another, the copy's runs with that value take the
  sample's; runs that differ from the rest of their paragraph (a word in bold) keep theirs.
  So the copy's text and its own emphasis stay; the look of the whole follows.

**Never changed**: text, the number of rows and columns, merges, list membership.  Rows are a
copy's values; a column added or removed, or cells merged or split, would move them.  Those
differences are reported per copy (``not_followed``), each with why and, where a Word tool
makes the change, the calls (``word_compare_blocks``'s fixes).  What is reported is what
:mod:`.compare` still finds after following, less the body row count (a copy's own by design).

**A sample in another open document**: styles are taken by name -- a style this document
lacks is added from Word's built-ins where Word has one by that name, otherwise the copy keeps
its own and the result says so.

**Tracked**, as each edit records itself (:mod:`.table_format`): the table's properties and grid
as ``w:tblPrChange`` and ``w:tblGridChange``, each cell's as ``w:tcPrChange``, paragraphs'
and runs' as ``w:pPrChange`` and ``w:rPrChange``.  Word does not track a row's height,
breaking or header (measured, E5): those are applied untracked and a warning says so.
"""

from __future__ import annotations

import copy
from collections import Counter
from typing import TYPE_CHECKING, Any, Callable

from ..oxml.xml import Element, insert_in_order, make, remove
from . import compare as _compare
from . import text as _text
from .table_format import TableTracker, grid_widths, own_rows, set_grid

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: What a copy keeps of its own in each kind of properties: the sample's are taken for the rest.
KEEP_TABLE = frozenset(_W + t for t in ("tblCaption", "tblDescription", "tblPrChange"))
KEEP_ROW = frozenset(_W + t for t in ("ins", "del", "trPrChange", "cnfStyle", "gridBefore", "gridAfter",
                                      "wBefore", "wAfter"))
KEEP_CELL = frozenset(_W + t for t in ("gridSpan", "vMerge", "hMerge", "cellIns", "cellDel", "cellMerge",
                                       "tcPrChange", "cnfStyle"))
KEEP_PARAGRAPH = frozenset(_W + t for t in ("numPr", "sectPr", "pPrChange", "cnfStyle", "rPr"))
#: A run's (and a paragraph mark's) revision marks and change record: never formatting.
KEEP_RUN = frozenset(_W + t for t in ("ins", "del", "moveFrom", "moveTo", "rPrChange"))

#: The table properties by the names ``word_format_table`` and :mod:`.compare` use.
TABLE_NAMES = {"tblStyle": "style", "tblW": "width", "jc": "alignment", "tblInd": "indent",
               "tblBorders": "borders", "shd": "shading", "tblLayout": "layout",
               "tblCellMar": "cell_margins", "tblCellSpacing": "cell_spacing", "tblLook": "look",
               "tblpPr": "floating", "tblOverlap": "floating", "bidiVisual": "right_to_left",
               "tblStyleRowBandSize": "look", "tblStyleColBandSize": "look"}

#: Why each kind of difference left after following was not followed.
WHY = {
    "kind": "the copy is not the same kind of block as the sample",
    "columns": "the copy's columns hold its values: columns are not added or removed",
    "widths": "the copy has another number of columns than the sample",
    "merges": "cells are not merged or split: it would move the copy's text",
    "missing_paragraph": "a line of the sample's text the copy's cell lacks: text is the copy's own",
    "extra_paragraph": "a line no sample paragraph matches by text or style: it keeps its formatting",
    "paragraph_style": "this document has no such style and Word has no built-in one by its name",
    "table_style": "this document has no such style and Word has no built-in one by its name",
    "list": "list membership is not followed: the sample's list is another list (word_lists)",
}
#: Why, for anything else left: the cell it is in is merged differently from the sample's.
WHY_OTHER = "the cell is merged differently from the sample's"

#: Differences reported per copy at most; ``more`` counts the rest.
NOT_FOLLOWED_CAP = 20
#: Differences reported in one reply at most, over every copy (the rest counted in ``more``
#: and named by the reply's ``left_out``).
REPLY_CAP = 100


class _Styles:
    """The copy document's style id for a sample's style id: the same document's ids as they
    are, another's by name (a built-in one added where this document lacks it)."""

    def __init__(self, document: "Document", source: "Document") -> None:
        self.document = document
        self.source = source
        self.missing: dict[str, str] = {}      # sample style name -> kind, not found here

    def __call__(self, style_id: str | None, kind: str) -> str | None:
        if style_id is None or self.source is self.document:
            return style_id
        name = _compare._style_name(self.source, style_id)
        try:
            return self.document.styles._ensure(name, kind)
        except Exception:  # noqa: BLE001 -- a style the document neither has nor would add
            self.missing[name] = kind
            return None


def follow(document: "Document", sample: str, copies: list[str], *,
           sample_document: "Document | None" = None, preview: bool = False) -> tuple[Any, dict]:
    """Each of ``copies`` made to follow ``sample`` (the module's rules): ``(EditResult,
    report)``, the report ``{"sample", "preview", "copies": [...]}`` with one entry per copy
    (:func:`_report`).  One undo step; with ``preview`` nothing changes."""
    from .document import EditError, EditResult

    source = sample_document or document
    notes: list[str] = []
    sample_entry = _locate(source, sample, None, notes, "sample")
    sample_id = sample_entry.id
    kind = "table" if sample_entry.element.tag == _W + "tbl" else "paragraph"
    resolved: list[tuple[str, Any]] = []
    seen: set[int] = set()
    for identifier in copies:
        if not isinstance(identifier, str) or not identifier.strip():
            notes.append("an empty copy address was ignored")
            continue
        entry = _locate(document, identifier.strip(), kind, notes, "copies", like=sample_entry.element)
        if source is document and entry.element is sample_entry.element:
            notes.append(f"{identifier} is the sample itself: nothing to follow")
            continue
        if id(entry.element) in seen:
            notes.append(f"{identifier} was named twice: followed once")
            continue
        seen.add(id(entry.element))
        resolved.append((identifier, entry))
    if not resolved and not copies:
        raise EditError("copies names at least one block made from the sample; tables like it: "
                        + (", ".join(_shapes(document, sample_entry.element)) or "none"))
    tracking = document._active_tracking()
    reports: list[dict] = []
    warnings: list[str] = list(dict.fromkeys(notes))

    class _Preview(Exception):
        pass

    def run() -> dict[str, str]:
        targets = []
        for _, entry in resolved:
            part = _part_of(document, entry.id)
            index = document._index(part)
            paragraphs = [e for e in index.paragraphs if _inside(e.element, entry.element)]
            targets.append((part, paragraphs))
        renames = document._prepare_many(targets)
        styles = _Styles(document, source)
        sample_element = source._resolve(sample_id)[1].element
        for asked, entry in resolved:
            identifier = renames.get(entry.id, entry.id)
            part, current = document._resolve(identifier)
            element = current.element
            log = _Log()
            kind = "table" if element.tag == _W + "tbl" else "paragraph"
            sample_kind = "table" if sample_element.tag == _W + "tbl" else "paragraph"
            if kind == sample_kind == "table":
                tracker = TableTracker(element, rows=True) if tracking is not None else None
                with document._track_properties([part]):
                    rows = _follow_table(document, source, element, sample_element, styles, log)
                if tracker is not None:
                    from ..revisions.stamp import Stamp

                    stamp = Stamp(document, tracking)
                    tracker.record(stamp, part)
                    stamp.finish()
                rows()          # untracked, as Word does not track them
                if tracking is not None and log.rows and not warnings:
                    warnings.append("Word does not track a row's height, breaking or header: applied untracked")
            elif kind == sample_kind:
                with document._track_properties([part]):
                    _follow_paragraph(document, source, element, sample_element, styles, log)
            document.package.mark_dirty(part)
            reports.append({"asked": asked, "id": identifier, "part": part, "log": log,
                            "styles": dict(styles.missing)})
            styles.missing.clear()
        # What is still different, as word_compare_blocks finds it.
        document._invalidate()
        for entry in reports:
            index = document._index(entry["part"])
            entry["lists"] = [found.id for p in entry["log"].list
                              if (found := index.entry_for(p)) is not None]
            found = document.compare_blocks([entry["id"]], sample_id,
                                            reference_document=source if source is not document else None)
            entry["left"] = [d for d in (_textless(document, d) for d in found["differences"]) if d is not None]
        return renames

    renames: dict[str, str] = {}
    if preview:
        try:
            with document.history.savepoint(), document._edit():
                renames = run()
                raise _Preview
        except _Preview:
            pass
        document._invalidate()          # the rolled-back trees are the ones before
    else:
        with document._edit():
            renames = run()
    copies_out = [_report(entry, preview) for entry in reports]
    room = REPLY_CAP
    for entry in copies_out:
        shown = entry["not_followed"]
        if len(shown) > room:
            entry["more"] = entry.get("more", 0) + len(shown) - room
            entry["not_followed"] = shown[:room]
        room -= len(entry["not_followed"])
    left_out = sum(entry.get("more", 0) for entry in copies_out)
    report: dict[str, Any] = {"sample": sample_id, "preview": preview, "copies": copies_out}
    if left_out:
        report["left_out"] = (f"{left_out} more difference(s) not followed, counted in each copy's 'more': "
                              "word_compare_blocks on that copy lists them all, paged")
    first = reports[0]["id"] if reports else None
    changed = not preview and any(entry["changed"] for entry in copies_out)
    result = EditResult(first, renamed={} if preview else renames, changed=changed,
                        count=sum(1 for entry in copies_out if entry["changed"]), warnings=warnings)
    return result, report


def _shapes(document: "Document", like: Element | None, limit: int = 20) -> list[str]:
    """The document's tables as ``t:X (rows x columns)``, those with ``like``'s columns first."""
    found = []
    for table in document.tables():
        element = document._resolve(table.id)[1].element
        found.append((like is not None and len(grid_widths(element)) != len(grid_widths(like)),
                      f"{table.id} ({len(own_rows(element))} rows x {len(grid_widths(element))} columns)"))
    found.sort(key=lambda item: item[0])
    out = [text for _, text in found[:limit]]
    if len(found) > limit:
        out.append(f"{len(found) - limit} more (word_read lists them)")
    return out


def _locate(document: "Document", identifier: str, kind: str | None, notes: list[str], field: str, *,
            like: Element | None = None):
    """The block ``identifier`` names, as the other tools take it: a cell or row address
    (``t:X/c1,2``, ``t:X/r1``), or a paragraph in a table, where a table is wanted (or for
    the sample), names its table -- with a note.  A refusal names the tables there are."""
    from .addresses import AddressError
    from .document import EditError
    from .ids import TableEntry

    address = identifier.strip()
    if address.startswith("t:") and "/" in address and address.split("/", 1)[1][:1] in ("c", "r"):
        table = address.split("/", 1)[0]
        notes.append(f"{address} is in table {table}: the table follows")
        address = table
    try:
        part, entry = document._resolve(address)
    except (KeyError, AddressError) as error:
        message = error.args[0] if error.args else str(error)
        raise EditError(f"{field}: {message}; the tables here: "
                        + (", ".join(_shapes(document, like)) or "none")) from None
    if not isinstance(entry, TableEntry) and kind != "paragraph":
        holder = entry.element.getparent()
        while holder is not None and holder.tag != _W + "tbl":
            holder = holder.getparent()
        found = document._index(part).entry_for(holder) if holder is not None else None
        if found is not None:
            notes.append(f"{identifier} is a paragraph in table {found.id}: the table follows")
            entry = found
    return entry


def _textless(document: "Document", difference: dict) -> dict | None:
    """A run property's difference without the paragraphs that have no text: their mark took
    the sample's run formatting, which a comparison (reading text) does not see."""
    if difference["property"] not in _compare._RUN or not difference.get("paragraphs"):
        return difference
    kept = [p for p in difference["paragraphs"]
            if _text.paragraph_text(document._resolve(p)[1].element).strip()]
    if not kept:
        return None
    if difference.get("fix"):
        difference["fix"] = {**difference["fix"], "calls": [{**call, "targets": kept}
                                                            for call in difference["fix"]["calls"]]}
    return {**difference, "paragraphs": kept}


def _part_of(document: "Document", identifier: str) -> str:
    return document._resolve(identifier)[0]


def _inside(node: Element, ancestor: Element) -> bool:
    while node is not None:
        if node is ancestor:
            return True
        node = node.getparent()
    return False


class _Log:
    """What following changed in one copy."""

    def __init__(self) -> None:
        self.table: list[str] = []
        self.widths = False
        self.rows = 0
        self.cells = 0
        self.paragraphs = 0
        self.runs = 0
        #: Paragraphs in a list where their sample paragraph is in none, or the other way.
        self.list: list[Element] = []

    def changed(self) -> bool:
        return bool(self.table or self.widths or self.rows or self.cells or self.paragraphs or self.runs)


def _report(entry: dict, preview: bool) -> dict:
    """One copy's entry: ``copy`` (its id), ``changed``, what changed (``table`` the table's
    properties by name, ``widths``, and how many ``rows``, ``cells``, ``paragraphs`` and
    ``runs``), and ``not_followed``: each difference left, with ``why`` (and ``fix``, the calls
    that make it, where a tool does)."""
    log: _Log = entry["log"]
    out: dict[str, Any] = {"copy": entry["id"], "changed": log.changed()}
    if entry["asked"] != entry["id"]:
        out["asked"] = entry["asked"]
    if log.table:
        out["table"] = log.table
    if log.widths:
        out["widths"] = True
    for name in ("rows", "cells", "paragraphs", "runs"):
        if getattr(log, name):
            out[name] = getattr(log, name)
    left = []
    for difference in entry["left"]:
        name = difference["property"]
        if name in ("rows", "paragraphs"):
            continue            # a copy's own row count; a cell's line count, said by its lines
        item = {k: v for k, v in difference.items() if k != "block"}
        why = WHY.get(name, WHY_OTHER)
        if name in ("paragraph_style", "table_style") and difference.get("reference") not in entry["styles"]:
            why = WHY_OTHER
        item["why"] = why
        left.append(item)
    if entry["lists"]:
        left.append({"property": "list", "paragraphs": entry["lists"], "why": WHY["list"], "fix": None})
    out["not_followed"] = left[:NOT_FOLLOWED_CAP]
    if len(left) > NOT_FOLLOWED_CAP:
        out["more"] = len(left) - NOT_FOLLOWED_CAP
    return out


# -- tables ------------------------------------------------------------------------------------


def _take(target: Element, source: Element | None, keep: frozenset, tag: str,
          translate: Callable[[Element], Element | None] | None = None) -> tuple[Element | None, list[str]]:
    """``target``'s properties ``tag`` (a child of ``target``) made the sample's (``source``'s
    children), keeping its own of ``keep``: the local names that changed, and the element
    (``None`` where it is left empty)."""
    old = target.find(_W + tag.partition(":")[2])
    new = make(tag)
    for child in (source if source is not None else ()):
        if not isinstance(child.tag, str) or child.tag in keep:
            continue
        taken = copy.deepcopy(child)
        if translate is not None:
            taken = translate(taken)
            if taken is None:
                continue
        insert_in_order(new, taken)
    for child in (old if old is not None else ()):
        if isinstance(child.tag, str) and child.tag in keep:
            insert_in_order(new, copy.deepcopy(child))
    before = _children(old, keep)
    after = _children(new, keep)
    names = sorted({name for name in set(before) | set(after) if before.get(name) != after.get(name)})
    if not names:
        return old, []
    if old is not None:
        position = list(target).index(old)
        remove(old)
    else:
        position = None
    if len(new):
        if position is not None:
            target.insert(position, new)
        else:
            insert_in_order(target, new)
        return new, names
    return None, names


def _key(node: Element) -> tuple:
    """An element's value: its name, attributes and children, whatever namespaces its
    document declares (another document's root declares others)."""
    return (node.tag, tuple(sorted(node.attrib.items())), (node.text or "").strip(),
            tuple(_key(child) for child in node if isinstance(child.tag, str)))


def _children(properties: Element | None, keep: frozenset) -> dict[str, tuple]:
    out: dict[str, tuple] = {}
    for child in (properties if properties is not None else ()):
        if isinstance(child.tag, str) and child.tag not in keep:
            name = child.tag.rpartition("}")[2]
            out[name] = out.get(name, ()) + (_key(child),)
    return out


def _style_translator(styles: _Styles, tag: str, kind: str) -> Callable[[Element], Element | None]:
    def translate(node: Element) -> Element | None:
        if node.tag != _W + tag:
            return node
        found = styles(node.get(_W + "val"), kind)
        if found is None:
            return None
        node.set(_W + "val", found)
        return node

    return translate


def _follow_table(document: "Document", source: "Document", table: Element, sample: Element,
                  styles: _Styles, log: _Log) -> Callable[[], None]:
    """The table's, grid's, cells' and paragraphs' properties (tracked by the caller); the
    rows' returned as a call to make after the tracked record."""
    translate = _style_translator(styles, "tblStyle", "table")
    old_style = table.find(f"{_W}tblPr/{_W}tblStyle")
    missing_before = len(styles.missing)
    _, names = _take(table, sample.find(_W + "tblPr"), KEEP_TABLE, "w:tblPr", translate)
    if len(styles.missing) > missing_before and old_style is not None \
            and table.find(f"{_W}tblPr/{_W}tblStyle") is None:
        # The sample's style is not here: the copy keeps its own.
        properties = table.find(_W + "tblPr")
        if properties is None:
            properties = make("w:tblPr")
            insert_in_order(table, properties)
        insert_in_order(properties, copy.deepcopy(old_style))
        names = [n for n in names if n != "tblStyle"]
    log.table = list(dict.fromkeys(TABLE_NAMES.get(n, n) for n in names))
    widths, sample_widths = grid_widths(table), grid_widths(sample)
    same_columns = len(widths) == len(sample_widths)
    if same_columns and widths != sample_widths:
        set_grid(table, sample_widths)
        log.widths = True
    rows, sample_rows = own_rows(table), own_rows(sample)
    if not rows or not sample_rows:
        return lambda: None
    header = _compare._header_rows(sample_rows)
    mapping = _compare._row_map(len(rows), len(sample_rows), header)
    if same_columns:
        grid, sample_grid = _compare._origins(rows), _compare._origins(sample_rows)
        for r, line in enumerate(grid):
            sample_line = sample_grid[mapping[r]]
            for column, (cell, span, up) in sorted(line.items()):
                match = sample_line.get(column)
                if match is None or match[1] != span:
                    continue            # merged otherwise: compare reports it
                sample_cell = match[0]
                if _take(cell, sample_cell.find(_W + "tcPr"), KEEP_CELL, "w:tcPr")[1]:
                    log.cells += 1
                if up:
                    continue
                paragraphs = cell.findall(_W + "p")
                sample_paragraphs = sample_cell.findall(_W + "p")
                if not paragraphs or not sample_paragraphs:
                    continue
                pairs, _, _ = _compare._align(document, paragraphs, source, sample_paragraphs)
                for paragraph, reference in pairs:
                    _follow_paragraph(document, source, paragraph, reference, styles, log)

    def follow_rows() -> None:
        for r, row in enumerate(rows):
            if _take(row, sample_rows[mapping[r]].find(_W + "trPr"), KEEP_ROW, "w:trPr")[1]:
                log.rows += 1

    return follow_rows


# -- paragraphs --------------------------------------------------------------------------------


def _follow_paragraph(document: "Document", source: "Document", paragraph: Element, sample: Element,
                      styles: _Styles, log: _Log) -> None:
    """A paragraph's properties and its runs' formatting from ``sample``'s (the module's rules)."""
    sample_properties = sample.find(_W + "pPr")
    translate = _style_translator(styles, "pStyle", "paragraph")
    old_style = paragraph.find(f"{_W}pPr/{_W}pStyle")
    missing_before = len(styles.missing)
    properties, names = _take(paragraph, sample_properties, KEEP_PARAGRAPH, "w:pPr", translate)
    if len(styles.missing) > missing_before and old_style is not None:
        properties = paragraph.find(_W + "pPr")
        if properties is None:
            properties = make("w:pPr")
            paragraph.insert(0, properties)
        if properties.find(_W + "pStyle") is None:
            insert_in_order(properties, copy.deepcopy(old_style))
            names = [n for n in names if n != "pStyle"]
    has_list = paragraph.find(f"{_W}pPr/{_W}numPr") is not None
    sample_list = sample.find(f"{_W}pPr/{_W}numPr") is not None
    if has_list != sample_list:
        log.list.append(paragraph)
    # The paragraph mark's formatting.
    run_translate = _style_translator(styles, "rStyle", "character")
    mark_source = sample_properties.find(_W + "rPr") if sample_properties is not None else None
    owner = paragraph.find(_W + "pPr")
    if owner is None and mark_source is not None and any(
            isinstance(c.tag, str) and c.tag not in KEEP_RUN for c in mark_source):
        owner = make("w:pPr")
        paragraph.insert(0, owner)
    mark_changed = False
    if owner is not None:
        _, mark = _take(owner, mark_source, KEEP_RUN, "w:rPr", run_translate)
        mark_changed = bool(mark)
        if not len(owner):
            remove(owner)
    if names or mark_changed:
        log.paragraphs += 1
    log.runs += _follow_runs(paragraph, sample, styles)


def _run_weights(paragraph: Element) -> list[tuple[Element, int]]:
    """The paragraph's runs in order, each with how many characters of text it has."""
    walked = _text.walk(paragraph)
    counts: Counter = Counter()
    for atom in walked.atoms:
        if atom.run is not None:
            counts[id(atom.run)] += 1
    return [(run, counts[id(run)]) for run in walked.runs]


def _run_values(run: Element, styles: _Styles | None = None) -> dict[str, tuple]:
    """Per property of the run, ``(value, element)``: the sample's styles as this document's."""
    properties = run.find(_W + "rPr")
    out: dict[str, tuple] = {}
    for child in (properties if properties is not None else ()):
        if not isinstance(child.tag, str) or child.tag in KEEP_RUN:
            continue
        name = child.tag.rpartition("}")[2]
        if name == "rStyle" and styles is not None:
            found = styles(child.get(_W + "val"), "character")
            if found is None:
                continue
            child = copy.deepcopy(child)
            child.set(_W + "val", found)
        out[name] = (_key(child), child)
    return out


def _majority(values: list[tuple[dict[str, tuple], int]], name: str) -> tuple | None:
    """The ``(value, element)`` of ``name`` most of the text has (``None``: not stated); a tie
    goes to not stated, then to the earlier run's."""
    counts: Counter = Counter()
    first: dict[Any, int] = {}
    held: dict[Any, tuple] = {}
    for k, (found, weight) in enumerate(values):
        entry = found.get(name)
        value = entry[0] if entry is not None else None
        counts[value] += weight
        first.setdefault(value, k)
        if entry is not None:
            held.setdefault(value, entry)
    if not counts:
        return None
    best = max(counts, key=lambda value: (counts[value], value is None, -first[value]))
    return held.get(best)


def _follow_runs(paragraph: Element, sample: Element, styles: _Styles) -> int:
    """The copy's runs' formatting toward the sample's, property by property; how many runs
    changed."""
    sample_runs = [(_run_values(run, styles), weight) for run, weight in _run_weights(sample)]
    runs = [(run, _run_values(run), weight) for run, weight in _run_weights(paragraph)]
    if not sum(weight for _, weight in sample_runs):
        return 0
    if not sum(weight for _, _, weight in runs):
        # No text yet: the paragraph mark's formatting is what text typed there takes.
        holder = paragraph.find(_W + "pPr")
        if holder is None:
            holder = make("w:pPr")
            paragraph.insert(0, holder)
        runs.append((holder, _run_values(holder), 1))
    names = {name for found, _ in sample_runs for name in found} | {name for _, found, _ in runs for name in found}
    changed: set[int] = set()
    for name in sorted(names):
        wanted = _majority(sample_runs, name)
        mine = _majority([(found, weight) for _, found, weight in runs], name)
        wanted_value = wanted[0] if wanted is not None else None
        mine_value = mine[0] if mine is not None else None
        if wanted_value == mine_value:
            continue
        for run, found, _ in runs:
            if (found[name][0] if name in found else None) != mine_value:
                continue        # its own emphasis
            properties = run.find(_W + "rPr")
            for node in (properties.findall(_W + name) if properties is not None else []):
                remove(node)
            if wanted is not None:
                if properties is None:
                    properties = make("w:rPr")
                    insert_in_order(run, properties)
                insert_in_order(properties, copy.deepcopy(wanted[1]))   # rStyle already this document's
            if properties is not None and not len(properties):
                remove(properties)
            if run.tag == _W + "pPr" and not len(run):
                remove(run)
            changed.add(id(run))
    for holder, _, _ in runs:
        if holder.tag == _W + "pPr" and not len(holder) and holder.getparent() is not None:
            remove(holder)
    return len(changed)


__all__ = ["KEEP_CELL", "KEEP_PARAGRAPH", "KEEP_ROW", "KEEP_TABLE", "NOT_FOLLOWED_CAP", "REPLY_CAP", "WHY", "follow"]
