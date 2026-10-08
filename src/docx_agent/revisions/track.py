"""The edit primitives in tracking mode: each writes the revision form Word writes for it.

Measured on Word 16.106 for Mac (``tools/e3_probe.py``, ``tests/observations/e3-word.json``;
ROADMAP.md, "Phase E3"):

* **Inserted text** is a run of its own inside ``w:ins``, beside the text it follows; text
  inserted into one's own insertion joins it, into another author's splits it.
* **Deleted text** is its runs inside ``w:del``, ``w:t`` turned ``w:delText`` and
  ``w:instrText`` ``w:delInstrText`` (a field's begin, separate and end runs with them).
  Deleting one's own insertion removes it outright, as Word does.  A replacement is the
  deletion, then the insertion.
* **A paragraph mark** inserted or deleted is ``w:pPr/w:rPr/w:ins|w:del``.  A deleted mark
  joins its paragraph to the next on accepting; Word, tracking a deletion across a mark,
  gives the next paragraph the first one's properties with a ``w:pPrChange``, so that
  accepting makes what deleting untracked makes (the first paragraph's properties).
* **Formatting** is the new properties with the *complete* old ones in ``w:rPrChange`` or
  ``w:pPrChange``; formatting reverted to the old drops the record.
* **Moves** are ``w:moveFrom`` / ``w:moveTo`` with range markers sharing a name; the
  paragraph's mark moves with it; the destination is a new paragraph.
* **Rows** are ``w:trPr/w:ins|w:del`` with each cell's paragraph marks and content.

Property changes are recorded by snapshot (:class:`PropertyTracker`): the edit runs as it
would untracked, over paragraphs and runs tagged with a temporary attribute (which a split
run's copy keeps, and which keeps runs of different origins from merging), and every one
whose properties came out different gets its record.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from lxml import etree

from ..edit import inline as _inline
from ..edit import text as _text
from ..oxml.xml import Element, append_in_order, insert_in_order, make, qn, remove
from .stamp import Stamp

if TYPE_CHECKING:  # pragma: no cover
    pass

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
W_P = _W + "p"
W_R = _W + "r"
W_INS = _W + "ins"
W_DEL = _W + "del"
W_MOVE_FROM = _W + "moveFrom"
W_MOVE_TO = _W + "moveTo"
W_PPR = _W + "pPr"
W_RPR = _W + "rPr"

#: Containers whose content is inserted (current view) or deleted (original view).
INSERTED = (W_INS, W_MOVE_TO)
DELETED = (W_DEL, W_MOVE_FROM)
CONTAINERS = INSERTED + DELETED
#: Marks a paragraph's ``w:rPr`` holds as revisions.
MARKS = frozenset(CONTAINERS)

#: Run content spelled differently inside a deletion.
_TO_DELETED = {_W + "t": _W + "delText", _W + "instrText": _W + "delInstrText"}
_TO_SHOWN = {value: key for key, value in _TO_DELETED.items()}


# -- containers -----------------------------------------------------------------------------


def revision_owner(node: Element, stop: Element | None = None) -> Element | None:
    """The innermost revision container (``w:ins``, ``w:del``, ``w:moveFrom``,
    ``w:moveTo``) ``node`` is in, below ``stop``; ``None`` when there is none."""
    parent = node.getparent()
    while parent is not None and parent is not stop and parent.tag != W_P:
        if parent.tag in CONTAINERS:
            return parent
        parent = parent.getparent()
    return None


def _same_container(left: Element, right: Element) -> bool:
    return (left.tag == right.tag and left.get(_W + "author") == right.get(_W + "author")
            and left.get(_W + "date") == right.get(_W + "date"))


def join_neighbours(container: Element) -> Element:
    """Merge ``container`` into an adjacent revision container of the same kind, author
    and date (one revision, as Word writes a contiguous change).  Returns the survivor."""
    previous = container.getprevious()
    if previous is not None and _same_container(previous, container) and not (previous.tail or "").strip():
        for child in list(container):
            previous.append(child)
        remove(container)
        container = previous
    following = container.getnext()
    if following is not None and _same_container(container, following):
        for child in list(following):
            container.append(child)
        remove(following)
    return container


def spell_deleted(run: Element) -> None:
    for child in run:
        if isinstance(child.tag, str) and child.tag in _TO_DELETED:
            child.tag = _TO_DELETED[child.tag]


def spell_shown(node: Element) -> None:
    for child in node.iter(*_TO_SHOWN):
        child.tag = _TO_SHOWN[child.tag]


def relationship_ids(node: Element) -> list[str]:
    return _inline.relationship_ids(node)


# -- deleting -------------------------------------------------------------------------------


def delete_runs(runs: list[Element], stamp: Stamp, part: str) -> list[str]:
    """Track the deletion of ``runs`` (in document order): each goes inside a ``w:del``
    where it stands -- in a hyperlink, a field's result, another author's insertion -- one
    ``w:del`` per stretch of sibling runs; a run inside one's own insertion is removed.
    Returns the relationship ids of what was removed outright (pictures, for release)."""
    released: list[str] = []
    current: Element | None = None
    for run in runs:
        if run.getparent() is None:
            continue
        owner = revision_owner(run)
        if owner is not None and owner.tag in DELETED:
            continue  # already deleted (or moved away)
        if owner is not None and owner.tag == W_INS and stamp.same(owner):
            released += relationship_ids(run)
            parent = run.getparent()
            remove(run)
            while parent is not None and parent.tag != W_P and not any(isinstance(c.tag, str) for c in parent):
                # What the run leaves empty goes too: its insertion, a hyperlink (released).
                if parent.get(_R + "id"):
                    released.append(parent.get(_R + "id"))
                grand = parent.getparent()
                remove(parent)
                parent = grand
            current = None
            continue
        spell_deleted(run)
        previous = run.getprevious()
        if current is not None and previous is current:
            current.append(run)
            continue
        current = stamp.make("w:del", part)
        run.addprevious(current)
        current.append(run)
        join_neighbours(current)
        current = current if current.getparent() is not None else None
    return released


def delete_span(paragraph: Element, start: int, end: int, stamp: Stamp, part: str) -> list[str]:
    """Track the deletion of the current view's characters ``[start, end)``."""
    if start >= end:
        return []
    runs = _inline.isolate(paragraph, start, end)
    return delete_runs(runs, stamp, part)


def content_runs(paragraph: Element) -> list[Element]:
    """Every run of the paragraph in the current view, field characters and instructions
    included: what deleting the whole paragraph deletes."""
    return [run for run in _text.walk(paragraph, "current").runs]


def delete_content(paragraph: Element, stamp: Stamp, part: str) -> list[str]:
    """Track the deletion of a paragraph's whole content (not its mark)."""
    return delete_runs(content_runs(paragraph), stamp, part)


def wrap_runs(runs: list[Element], tag: str, stamp: Stamp, part: str) -> list[Element]:
    """Wrap each stretch of sibling ``runs`` in a new revision container ``tag``
    (``w:moveFrom``, ``w:moveTo``, ``w:ins``) where it stands.  Moved-away text stays
    ``w:t``, as Word writes it (measured).  Returns the containers."""
    out: list[Element] = []
    current: Element | None = None
    for run in runs:
        if run.getparent() is None:
            continue
        if current is not None and run.getprevious() is current:
            current.append(run)
            continue
        current = stamp.make(tag, part)
        run.addprevious(current)
        current.append(run)
        out.append(current)
    return out


# -- inserting ------------------------------------------------------------------------------


def _split_container(container: Element, index: int, stamp: Stamp) -> Element:
    """Split a revision container before its ``index``-th child: the rest goes into a copy
    placed after it -- a revision of its own, with an id of its own, as Word splits one.
    Returns the copy."""
    rest = container.makeelement(container.tag, dict(container.attrib))
    rest.set(_W + "id", stamp.next_id())
    for child in list(container)[index:]:
        rest.append(child)
    container.addnext(rest)
    return rest


def insert_nodes(parent: Element, index: int, nodes: list[Element], stamp: Stamp, part: str) -> Element | None:
    """Track the insertion of run-level ``nodes`` at ``parent[index]``: inside a new
    ``w:ins`` -- or, at a place inside one's own insertion, as part of it; inside another
    author's insertion or a deletion, the container is split around the new one.  Returns
    the ``w:ins`` written (``None`` when they joined one's own)."""
    owner = parent if parent.tag in CONTAINERS else revision_owner(parent)
    if owner is not None and owner.tag == W_INS and stamp.same(owner):
        for k, node in enumerate(nodes):
            parent.insert(index + k, node)
        return None
    while owner is not None:
        # Climb out of every revision container: split each at the insertion point.
        if parent is owner:
            if 0 < index < len(parent):
                rest = _split_container(owner, index, stamp)
                parent, index = rest.getparent(), rest.getparent().index(rest)
            else:
                grand = owner.getparent()
                parent, index = grand, grand.index(owner) + (1 if index >= len(owner) else 0)
        else:
            # The point is inside something (a hyperlink) inside the container: leave it as
            # it is -- the new insertion goes in that something, which the container owns.
            break
        owner = parent if parent.tag in CONTAINERS else revision_owner(parent)
    container = stamp.make("w:ins", part)
    parent.insert(index, container)
    for node in nodes:
        container.append(node)
    return join_neighbours(container)


def place_outside_revisions(node: Element, stamp: Stamp) -> None:
    """Move a paragraph-level element (a hyperlink) that sits inside a revision container
    out of it, splitting the container around it: a ``w:ins`` or ``w:del`` holds runs, not
    hyperlinks (Word drops a hyperlink it finds in one: measured, a re-save keeps its text
    and loses the link)."""
    parent = node.getparent()
    while parent is not None and parent.tag in CONTAINERS:
        index = parent.index(node)
        if 0 < index < len(parent) - 1:
            rest = _split_container(parent, index + 1, stamp)
            parent.addnext(node)
        elif index == 0:
            parent.addprevious(node)
        else:
            parent.addnext(node)
        if not any(isinstance(c.tag, str) for c in parent):
            remove(parent)
        parent = node.getparent()


def insert_text(paragraph: Element, offset: int, text: str, stamp: Stamp, part: str,
                reference: Element | None = None) -> Element:
    """Track the insertion of ``text`` at the current view's ``offset``, formatted as the
    run ``reference`` (else as plain insertion there would be).  Returns the new run."""
    items = _text.atoms(paragraph)
    if reference is None:
        reference = _reference_run(items, offset)
    properties = _text.run_properties_for(paragraph, reference)
    run = _text.make_run(text, properties)
    parent, index = _inline.position(paragraph, offset)
    insert_nodes(parent, index, [run], stamp, part)
    return run


def _reference_run(items, offset: int) -> Element | None:
    """The run new text at ``offset`` takes its formatting from: the character before it
    when that is text, else the one after, else either (as untracked insertion does)."""
    before = items[offset - 1] if offset > 0 else None
    after = items[offset] if offset < len(items) else None
    for test in (lambda a: a.kind == "text" and not a.in_field, lambda a: a.kind == "text", lambda a: True):
        for atom in (before, after):
            if atom is not None and atom.run is not None and test(atom):
                return atom.run
    return None


def replace_span(paragraph: Element, start: int, end: int, new: str, stamp: Stamp, part: str, *,
                 keep: str = "first") -> list[str]:
    """Tracked :func:`docx_agent.edit.inline.replace_span`: the deletion of ``[start, end)``
    and the insertion of ``new`` after it -- with ``keep="characters"``, of each changed
    stretch only, as ``set_text``'s diff finds them."""
    items = _text.atoms(paragraph)
    _inline.check_span(items, start, end, deleting=True)
    if _text.OBJECT in new:
        raise _inline.SpanError("new text cannot create an object (U+FFFC)")
    old = "".join(atom.char for atom in items[start:end])
    if old == new:
        return []
    if keep == "first" or not old:
        stretches = [(start, end, new)]
    else:
        stretches = [(start + i1, start + i2, new[j1:j2])
                     for tag, i1, i2, j1, j2 in _text.coarse_opcodes(old, new) if tag != "equal"]
    released: list[str] = []
    for a, b, text in reversed(stretches):
        items = _text.atoms(paragraph)
        reference = None
        if b > a:
            first = items[a]
            reference = first.run if first.kind == "text" else _reference_run(items, a)
        if not old and a < len(items) and a > 0 and not _text._plain(items[a - 1]) and _text._plain(items[a]):
            reference = items[a].run
        released += delete_span(paragraph, a, b, stamp, part)
        if text:
            if reference is None or reference.getparent() is None:
                reference = _reference_run(_text.atoms(paragraph), a)
            insert_text(paragraph, a, text, stamp, part, reference)
    released += _inline.prune(paragraph)
    return released


# -- paragraph marks ------------------------------------------------------------------------


def mark(paragraph: Element, kind: str, stamp: Stamp, part: str, *, name: str | None = None) -> Element:
    """Record the paragraph's mark as inserted, deleted or moved (``ins``, ``del``,
    ``moveFrom``, ``moveTo``)."""
    properties = paragraph.find(W_PPR)
    if properties is None:
        properties = make("w:pPr")
        paragraph.insert(0, properties)
    run_properties = properties.find(W_RPR)
    if run_properties is None:
        run_properties = make("w:rPr")
        insert_in_order(properties, run_properties)
    record = stamp.make(f"w:{kind}", part)
    insert_in_order(run_properties, record)
    return record


def mark_record(paragraph: Element) -> Element | None:
    """The paragraph mark's own revision record, if it has one."""
    found = paragraph.find(f"{_W}pPr/{_W}rPr")
    if found is None:
        return None
    return next((child for child in found if child.tag in MARKS), None)


def clean_properties(properties: Element | None, tag: str) -> Element:
    """A copy of a ``w:pPr`` or ``w:rPr`` without what is not formatting: a paragraph's
    mark properties, section break and change record; a run's revision marks and change
    record."""
    out = make(tag)
    if properties is None:
        return out
    skip = {W_RPR, _W + "sectPr", _W + "pPrChange"} if tag == "w:pPr" else set(MARKS) | {_W + "rPrChange"}
    for child in properties:
        if isinstance(child.tag, str) and child.tag not in skip:
            out.append(copy.deepcopy(child))
    return out


def _canonical(element: Element) -> bytes:
    return etree.tostring(element, method="c14n")


def record_paragraph_change(paragraph: Element, old: Element, stamp: Stamp, part: str) -> None:
    """Give ``paragraph`` a ``w:pPrChange`` holding ``old`` (its properties before an edit),
    unless they are the same -- or the change it already has records them, in which case
    a change back to them drops the record."""
    properties = paragraph.find(W_PPR)
    new = clean_properties(properties, "w:pPr")
    existing = properties.find(_W + "pPrChange") if properties is not None else None
    if existing is not None:
        original = existing.find(W_PPR)
        if original is not None and _canonical(clean_properties(original, "w:pPr")) == _canonical(new):
            remove(existing)
            if not len(properties):
                remove(properties)
        return
    if _canonical(old) == _canonical(new):
        return
    if properties is None:
        properties = make("w:pPr")
        paragraph.insert(0, properties)
    record = stamp.make("w:pPrChange", part)
    record.append(copy.deepcopy(old))
    append_in_order(properties, record)


def record_run_change(owner: Element, old: Element, stamp: Stamp, part: str) -> None:
    """:func:`record_paragraph_change` for run properties: a run's, or a paragraph mark's
    (``owner`` is then the ``w:pPr``)."""
    properties = owner.find(W_RPR)
    new = clean_properties(properties, "w:rPr")
    existing = properties.find(_W + "rPrChange") if properties is not None else None
    if existing is not None:
        original = existing.find(W_RPR)
        if original is not None and _canonical(clean_properties(original, "w:rPr")) == _canonical(new):
            remove(existing)
            if not len(properties):
                remove(properties)
                if owner.tag == W_PPR and not len(owner):
                    remove(owner)
        return
    if _canonical(old) == _canonical(new):
        return
    if properties is None:
        properties = make("w:rPr")
        insert_in_order(owner, properties)
    record = stamp.make("w:rPrChange", part)
    record.append(copy.deepcopy(old))
    append_in_order(properties, record)


#: The temporary mark :class:`PropertyTracker` puts on paragraphs and runs: an attribute in
#: no namespace, so no declaration is left behind once it is taken off.
TAG = "_docx_agent_tracking"


class PropertyTracker:
    """Snapshot the formatting of every paragraph and run under some roots, let an edit run
    untracked, then record what it changed (:func:`record_paragraph_change`,
    :func:`record_run_change`)."""

    def __init__(self, roots: list[Element]) -> None:
        self.roots = roots
        self.paragraphs: dict[str, tuple[Element, Element]] = {}
        self.runs: dict[str, Element] = {}
        counter = 0
        for root in roots:
            for node in root.iter(W_P, W_R):
                counter += 1
                key = str(counter)
                node.set(TAG, key)
                if node.tag == W_P:
                    properties = node.find(W_PPR)
                    mark_properties = properties.find(W_RPR) if properties is not None else None
                    self.paragraphs[key] = (clean_properties(properties, "w:pPr"),
                                            clean_properties(mark_properties, "w:rPr"))
                else:
                    self.runs[key] = clean_properties(node.find(W_RPR), "w:rPr")

    def record(self, stamp: Stamp, part: str) -> None:
        for root in self.roots:
            for node in list(root.iter(W_P, W_R)):
                key = node.attrib.pop(TAG, None)
                if key is None:
                    continue
                if node.tag == W_P and key in self.paragraphs:
                    old_paragraph, old_mark = self.paragraphs[key]
                    record_paragraph_change(node, old_paragraph, stamp, part)
                    properties = node.find(W_PPR)
                    if properties is not None or len(old_mark):
                        if properties is None:
                            properties = make("w:pPr")
                            node.insert(0, properties)
                        record_run_change(properties, old_mark, stamp, part)
                        if not len(properties):
                            remove(properties)
                elif node.tag == W_R and key in self.runs:
                    owner = revision_owner(node)
                    # Formatting one's own insertion is part of the insertion; deleted text
                    # is not formatted.
                    if owner is None or (owner.tag in INSERTED and not stamp.same(owner)):
                        record_run_change(node, self.runs[key], stamp, part)

    def discard(self) -> None:
        for root in self.roots:
            for node in root.iter(W_P, W_R):
                node.attrib.pop(TAG, None)


# -- whole paragraphs -------------------------------------------------------------------------


def insert_paragraph_content(paragraph: Element, stamp: Stamp, part: str) -> None:
    """Record a new paragraph's runs as inserted (in one ``w:ins`` after its properties and
    any markers)."""
    runs = [child for child in paragraph if isinstance(child.tag, str) and child.tag != W_PPR]
    if not runs:
        return
    container = stamp.make("w:ins", part)
    runs[0].addprevious(container)
    for run in runs:
        container.append(run)


def same_properties(left: Element | None, right: Element | None) -> bool:
    return _canonical(clean_properties(left, "w:pPr")) == _canonical(clean_properties(right, "w:pPr"))


def delete_mark(paragraph: Element, stamp: Stamp, part: str) -> Element | None:
    """Record a paragraph's mark as deleted -- or, when it is one's own insertion, take it
    out at once (Word removes one's own insertion outright): the paragraph's content goes
    to the start of the next paragraph, which is returned (``None`` when nothing joined)."""
    record = mark_record(paragraph)
    if record is None or (record.tag == W_INS and not stamp.same(record)):
        # Another author's inserted mark gets our deletion beside it, as Word records one.
        mark(paragraph, "del", stamp, part)
        return None
    if record.tag != W_INS:
        return None  # deleted or moved away already
    following = next_paragraph(paragraph)
    if following is None:
        return None
    run_properties = record.getparent()
    remove(record)
    if not len(run_properties):
        properties = run_properties.getparent()
        remove(run_properties)
        if not len(properties):
            remove(properties)
    at = 1 if following.find(W_PPR) is not None else 0
    for child in [c for c in paragraph if not (isinstance(c.tag, str) and c.tag == W_PPR)]:
        following.insert(at, child)
        at += 1
    remove(paragraph)
    return following


def take_properties(target: Element, source: Element, stamp: Stamp, part: str) -> None:
    """Give ``target`` the paragraph properties of ``source`` (a paragraph or its
    ``w:pPr``; its section break and mark aside), recording the old ones in a
    ``w:pPrChange``: how Word, tracking a deletion across a paragraph mark, makes accepting
    it keep the first paragraph's properties."""
    old = clean_properties(target.find(W_PPR), "w:pPr")
    new = clean_properties(source if source.tag == W_PPR else source.find(W_PPR), "w:pPr")
    if _canonical(old) == _canonical(new):
        return
    properties = target.find(W_PPR)
    if properties is None:
        properties = make("w:pPr")
        target.insert(0, properties)
    for child in list(properties):
        if child.tag not in (W_RPR, _W + "sectPr", _W + "pPrChange"):
            remove(child)
    for child in new:
        insert_in_order(properties, copy.deepcopy(child))
    record_paragraph_change(target, old, stamp, part)


def is_last_in_container(paragraph: Element) -> bool:
    """Whether no block follows the paragraph in its container (a cell, the body, a story):
    where Word cannot accept or reject its mark (measured: a cell's last paragraph's mark
    pulls the cell's first paragraph out of the table; the body's last stays)."""
    following = paragraph.getnext()
    while following is not None:
        if isinstance(following.tag, str) and following.tag in (W_P, _W + "tbl", _W + "sdt", _W + "customXml"):
            return False
        following = following.getnext()
    return True


def next_paragraph(paragraph: Element) -> Element | None:
    """The paragraph right after this one in its container, or ``None`` (a table, the end)."""
    following = paragraph.getnext()
    while following is not None:
        if isinstance(following.tag, str):
            if following.tag == W_P:
                return following
            if following.tag in (_W + "tbl", _W + "sdt", _W + "customXml", _W + "sectPr"):
                return None
        following = following.getnext()
    return None
