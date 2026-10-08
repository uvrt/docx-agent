"""Inline primitives: what every text and formatting edit is built from.

A paragraph's characters (:mod:`docx_agent.edit.text`'s text map, in the current view) are
addressed by Python string offsets.  On that map:

* :func:`replace_span` replaces the characters ``[start, end)`` with new text -- an insert
  when the span is empty, a delete when the text is -- writing new characters into the
  ``w:t`` beside them, so no run is created or split for text that has text next to it.
  ``keep="first"`` gives the whole replacement the formatting of the span's first character
  (Word's rule for a replacement); ``keep="characters"`` diffs the old and new text, as
  ``set_text`` does, so a character that survives keeps its own formatting.
* :func:`isolate` splits runs at a span's edges and returns the runs that hold exactly the
  span, for an edit that changes runs as a whole (formatting, a hyperlink around them);
  :func:`position` splits at one offset and says where something inserted there goes.
* :func:`merge_runs` merges what such an edit split, once it is done: adjacent runs among
  the touched ones whose properties came out the same, so no edit leaves redundant runs.
  Runs it did not touch are never merged (``coalesce_runs`` does that, on request).
* :func:`prune` removes what an edit emptied: a run with nothing but properties, and a
  hyperlink or an insertion that no longer holds a run.

Refused, before anything changes (:func:`check_span`): a span that cuts a field in two or
runs from a field's result into ordinary text, and one that would delete a note reference.
"""

from __future__ import annotations

import copy

from lxml import etree

from ..oxml.xml import XML_SPACE, Element, qn, remove
from . import text as _text
from .text import Atom

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
W_R = _W + "r"
W_T = _W + "t"
W_RPR = _W + "rPr"
W_PPR = _W + "pPr"

#: Run content that a merge may join: text and the characters spelled as elements.
_MERGEABLE = frozenset(_W + name for name in (
    "rPr", "t", "tab", "br", "cr", "noBreakHyphen", "softHyphen", "sym", "ptab"))
#: Containers an edit may empty, and remove once empty.
_PRUNABLE = frozenset(_W + name for name in ("hyperlink", "ins", "moveTo", "smartTag"))
#: Objects whose deletion would orphan what they refer to.
_REFERENCES = frozenset({_W + "footnoteReference", _W + "endnoteReference"})


class SpanError(ValueError):
    """A span an edit cannot take as it is."""


def check_span(items: list[Atom], start: int, end: int, *, deleting: bool = True) -> None:
    """Refuse a span that cuts a field, or would delete a note reference."""
    if not 0 <= start <= end <= len(items):
        raise SpanError(f"offsets {start}:{end} are outside the paragraph's {len(items)} characters")
    owners = {id(atom.field) if atom.field is not None else None for atom in items[start:end]}
    if len(owners) > 1:
        raise SpanError("the span cuts a field, or runs from a field's result into other text")
    if deleting:
        for atom in items[start:end]:
            if atom.kind == "element" and atom.node.tag in _REFERENCES:
                raise SpanError("the span holds a note reference; deleting it would orphan the note")


def replace_span(paragraph: Element, start: int, end: int, new: str, *, keep: str = "first") -> list[str]:
    """Replace the current view's characters ``[start, end)`` with ``new``.  Returns the
    relationship ids of the drawings and objects it removed (for the caller to release)."""
    if keep not in ("first", "characters"):
        raise ValueError("keep must be 'first' or 'characters'")
    _text.check_text(new)
    items = _text.atoms(paragraph)
    check_span(items, start, end, deleting=True)
    if _text.OBJECT in new:
        raise SpanError("new text cannot create an object (U+FFFC)")
    old = "".join(atom.char for atom in items[start:end])
    if old == new:
        return []
    deleted: set[int] = set()
    inserts: list[tuple[int | None, str, str]] = []
    if not old:
        anchor = _text._anchor(items, start, end)
        if anchor[0] is not None and anchor[1] in ("before", "after") and not _text._plain(items[anchor[0]]):
            # Only a field's result beside the insertion: a run of its own, outside the field.
            new_run(paragraph, start, _text._spell(new))
            return []
        inserts.append(anchor + (new,))
    elif keep == "first":
        deleted.update(range(start, end))
        if new:
            inserts.append(_text._anchor(items, start, end) + (new,))
    else:
        for tag, i1, i2, j1, j2 in _text.coarse_opcodes(old, new):
            if tag == "equal":
                continue
            deleted.update(range(start + i1, start + i2))
            if j2 > j1:
                inserts.append(_text._anchor(items, start + i1, start + i2) + (new[j1:j2],))
    released = [rid for k in deleted if items[k].kind == "element" for rid in relationship_ids(items[k].node)]
    _text._apply(paragraph, items, deleted, inserts)
    released += prune(paragraph)
    return released


def relationship_ids(node: Element) -> list[str]:
    """Every relationship id an element (a drawing, an object) spells, in ``r:`` attributes."""
    out = []
    for element in node.iter():
        if not isinstance(element.tag, str):
            continue
        for name, value in element.attrib.items():
            if name.startswith(_R):
                out.append(value)
    return out


def prune(paragraph: Element) -> list[str]:
    """Remove runs left with nothing but properties, and hyperlinks and insertions left
    without a run; returns the relationship ids of the hyperlinks removed."""
    released: list[str] = []
    changed = True
    while changed:
        changed = False
        for node in list(paragraph.iter(W_R, *_PRUNABLE)):
            if node.getparent() is None:
                continue
            if node.tag == W_R or node.getparent().tag in (W_RPR, _W + "trPr", _W + "numPr"):
                continue  # a paragraph mark's or row's revision record, not a container
            if not any(isinstance(child.tag, str) and child.tag != _W + "rPr" for child in node):
                rid = node.get(_R + "id")
                if rid:
                    released.append(rid)
                remove(node)
                changed = True
    return released


# -- splitting -------------------------------------------------------------------------------


def split_before(atom: Atom) -> bool:
    """Make ``atom`` the first content of its run: split its ``w:t`` at it, and the run
    before it.  Returns whether anything was split."""
    run = atom.run
    if run is None:
        return False  # an equation: it is its own boundary
    node = atom.node
    split = False
    if atom.kind in ("text", "deleted") and atom.offset > 0:
        text = node.text or ""
        left = _copy_text(node, text[:atom.offset])
        node.addprevious(left)
        _set_text(node, text[atom.offset:])
        split = True
    before = []
    for child in run:
        if child is node:
            break
        if isinstance(child.tag, str) and child.tag != W_RPR:
            before.append(child)
    if before:
        left_run = run.makeelement(run.tag, dict(run.attrib))
        properties = run.find(W_RPR)
        if properties is not None:
            left_run.append(copy.deepcopy(properties))
        for child in before:
            left_run.append(child)
        run.addprevious(left_run)
        split = True
    return split


def isolate(paragraph: Element, start: int, end: int) -> list[Element]:
    """Split runs so that ``[start, end)`` is exactly the content of whole runs; return those
    runs in document order (none for an empty span)."""
    items = _text.atoms(paragraph)
    check_span(items, start, end, deleting=False)
    if start == end:
        return []
    if end < len(items):
        split_before(items[end])
        items = _text.atoms(paragraph)
    split_before(items[start])
    items = _text.atoms(paragraph)
    out: list[Element] = []
    for atom in items[start:end]:
        if atom.run is not None and (not out or out[-1] is not atom.run):
            out.append(atom.run)
    return out


def position(paragraph: Element, offset: int) -> tuple[Element, int]:
    """Where an element inserted at character ``offset`` goes: ``(parent, index)``, after
    splitting the run there.  At the end it follows the paragraph's last run; in a
    paragraph without text, it goes after the paragraph's properties and any markers."""
    items = _text.atoms(paragraph)
    if not 0 <= offset <= len(items):
        raise SpanError(f"offset {offset} is outside the paragraph's {len(items)} characters")
    if offset < len(items):
        split_before(items[offset])
        items = _text.atoms(paragraph)
        atom = items[offset]
        anchor = atom.run if atom.run is not None else atom.node
        if atom.field is not None and (offset == 0 or items[offset - 1].field is not atom.field):
            # The first character of a field's result: what goes here goes before the field.
            anchor = atom.field if atom.field.tag == _W + "fldSimple" else atom.field.getparent()
        parent = anchor.getparent()
        return parent, parent.index(anchor)
    runs = _text.runs(paragraph)
    if runs:
        last = runs[-1]
        while last.getparent() is not None and last.getparent().tag == _W + "fldSimple":
            last = last.getparent()  # after a simple field, not inside its result
        parent = last.getparent()
        return parent, parent.index(last) + 1
    return paragraph, len(paragraph)


# -- merging ---------------------------------------------------------------------------------


def _signature(run: Element, *, rsids: bool = True) -> bytes:
    properties = run.find(W_RPR)
    rpr = etree.tostring(properties, method="c14n") if properties is not None else b""
    attributes = sorted((k, v) for k, v in run.attrib.items() if rsids or not k.startswith(_W + "rsid"))
    return repr(attributes).encode() + b"\0" + rpr


def _mergeable(run: Element) -> bool:
    return run.tag == W_R and all(not isinstance(child.tag, str) or child.tag in _MERGEABLE for child in run)


def merge_runs(runs: list[Element]) -> None:
    """Merge adjacent runs among ``runs`` (and the runs they were split from) whose
    properties are the same: what an edit that split runs leaves behind."""
    touched = {id(run) for run in runs}
    seen: set[int] = set()
    for run in runs:
        if id(run) in seen or run.getparent() is None:
            continue
        seen.add(id(run))
        # Merge into the previous sibling while it is a touched twin, then the next ones.
        current = run
        previous = current.getprevious()
        while (previous is not None and id(previous) in touched and _twins(previous, current)):
            _join(previous, current)
            current = previous
            previous = current.getprevious()
        following = current.getnext()
        while following is not None and id(following) in touched and _twins(current, following):
            seen.add(id(following))
            _join(current, following)
            following = current.getnext()


def merge_around(runs: list[Element]) -> None:
    """:func:`merge_runs` over ``runs`` and their immediate neighbours: the pieces a split
    left on either side of the runs an edit changed."""
    widened: list[Element] = []
    for run in runs:
        for candidate in (run.getprevious(), run, run.getnext()):
            if candidate is not None and candidate.tag == W_R and all(candidate is not w for w in widened):
                widened.append(candidate)
    merge_runs(widened)


def _twins(left: Element, right: Element) -> bool:
    return (left.getnext() is right and not (left.tail or "").strip() and _mergeable(left)
            and _mergeable(right) and _signature(left) == _signature(right))


def _join(left: Element, right: Element) -> None:
    for child in list(right):
        if isinstance(child.tag, str) and child.tag == W_RPR:
            continue
        left.append(child)
    remove(right)
    _join_texts(left)


def _join_texts(run: Element) -> None:
    previous = None
    for child in list(run):
        if child.tag == W_T and previous is not None and previous.tag == W_T and _same_attributes(previous, child):
            _set_text(previous, (previous.text or "") + (child.text or ""))
            remove(child)
            continue
        previous = child


def _same_attributes(a: Element, b: Element) -> bool:
    return {k: v for k, v in a.attrib.items() if k != XML_SPACE} == {k: v for k, v in b.attrib.items() if k != XML_SPACE}


def coalesce(paragraph: Element) -> int:
    """Merge adjacent runs whose properties differ only in revision save ids (``w:rsid*``):
    the explicit ``coalesce_runs``.  Returns how many runs were merged away."""
    merged = 0
    for run in list(_text.runs(paragraph)):
        if run.getparent() is None:
            continue
        following = run.getnext()
        while (following is not None and following.tag == W_R and not (run.tail or "").strip()
               and _mergeable(run) and _mergeable(following)
               and _signature(run, rsids=False) == _signature(following, rsids=False)):
            _join(run, following)
            merged += 1
            following = run.getnext()
    return merged


def _copy_text(node: Element, text: str) -> Element:
    copy_ = node.makeelement(node.tag, {k: v for k, v in node.attrib.items() if k != XML_SPACE})
    _set_text(copy_, text)
    return copy_


def _set_text(node: Element, text: str) -> None:
    node.text = text
    if text != text.strip() or "  " in text:
        node.set(XML_SPACE, "preserve")


def new_run(paragraph: Element, offset: int, children: list[Element]) -> Element:
    """A run holding ``children``, inserted at ``offset`` with the formatting new text there
    would take (the character before it, else the one after, else the paragraph mark's)."""
    items = _text.atoms(paragraph)
    reference = None
    if offset > 0 and items[offset - 1].run is not None:
        reference = items[offset - 1].run
    elif offset < len(items) and items[offset].run is not None:
        reference = items[offset].run
    properties = _text.run_properties_for(paragraph, reference)
    run = etree.Element(W_R)
    if properties is not None:
        run.append(properties)
    for child in children:
        run.append(child)
    parent, index = position(paragraph, offset)
    parent.insert(index, run)
    return run


def hoist(run: Element) -> None:
    """Put a run that a position placed inside a content control's content just outside
    the control (before it at the content's start, else after it): a control or drawing
    made there is not the control's content, which a fill replaces."""
    while True:
        parent = run.getparent()
        if parent is None or parent.tag != _W + "sdtContent":
            return
        sdt = parent.getparent()
        before = parent.index(run) == 0
        (sdt.addprevious if before else sdt.addnext)(run)
