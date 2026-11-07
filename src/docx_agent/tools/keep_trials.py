"""What a range's keep would do to the pages, laid out (round 40).

``check(range=[...])`` says of a range of whole blocks whether it is drawn on one page and
whether it is kept so (:func:`.scoped.range_keeps`), and where a keep would close its
breaks, the call that sets it (``fix``: ``word_format`` ``keep_together``).  What that call
does to the pages it did not say: Word moves the range whole to the next page, which can
leave much of the page before it blank and push what follows down.  An agent that sets the
keep, sees the gap in the next render and releases it, is offered the keep again by the
next check.

Here each such keep is **tried**: the range kept, alone -- the document's bytes with only
that call's change (:meth:`~docx_agent.Document.keep_blocks_together`, on a copy opened
from the bytes checked) -- laid out from the layout checked (docx2svg's incremental
layout: from the page the change reaches until the pages line up again; page for page the
layout of those bytes alone), in the worker holding that layout
(:mod:`.check_worker`).  Nothing estimated: ``if_kept`` says where the range then starts
and ends, whether it is on one page, how much of the page before it is left blank (points
and share of the body's height there, and what is blank there now), the document's page
count now and then, and the page what follows starts on now and then.

The other side, for a range kept already that starts a page (its first block the first
drawn there, not a page break or a section start that starts it anyway): ``if_released``,
the same facts with its keep released (``word_format`` ``keep_together: false``), and that
call as its ``fix`` -- so an agent sees both sides and decides once.

Each range is tried alone, whatever the others' keeps would do with it.  At most
:data:`TRIALS` trials a check, and within :data:`TRIAL_SECONDS` (the first is always made; rule 4): a range past
them says so (``effect_left_out``), and how to ask for it.
"""

from __future__ import annotations

import time
from fractions import Fraction
from typing import Any

#: Trial layouts one check makes at most (rule 4).
TRIALS = 4
#: Seconds of trial layouts one check spends at most: none is started past them (a check's
#: task has the layout deadline, 30 s by default, for everything it does).
TRIAL_SECONDS = 6.0

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_P = _W + "p"
_TBL = _W + "tbl"


def asked() -> dict[str, float]:
    """What a check asks of the worker (its ``keeps``): the caps, read here -- in the
    caller's process -- so that they are the caller's."""
    return {"trials": TRIALS, "seconds": TRIAL_SECONDS}


def keep_effects(previous: Any, data: bytes, options: dict, layout: Any, scope: Any,
                 reported: list[dict[str, Any]], *, trials: int = TRIALS,
                 seconds: float = TRIAL_SECONDS) -> None:
    """``if_kept`` on each reported range with a keep ``fix``, ``if_released`` on each kept
    range that starts a page, in place -- each from a trial layout of the document with that
    one change, laid out from ``previous`` (docx2svg's layout of ``data``, which ``layout``
    maps to ids); at most ``trials``, none started past ``seconds``."""
    tried = 0
    began = time.monotonic()
    index = None
    for found, entry in zip(scope.ranges, reported):
        on = _wanted(layout, found, entry)
        if on is None:
            continue
        if tried >= trials or (tried and time.monotonic() - began > seconds):
            entry["effect_left_out"] = (f"not laid out: a check tries at most {trials:g} keeps (within "
                                        f"{seconds:g} s); check this range alone (range: "
                                        f"[{found.address!r}]) for what {'keeping' if on else 'releasing'} "
                                        "it does")
            continue
        if index is None:
            index = _Index(layout)
        tried += 1
        try:
            effect = _trial(previous, data, options, layout, index, found, on)
        except Exception as error:  # noqa: BLE001 -- a trial never fails the check it is in
            entry["effect_left_out"] = (f"not laid out: the trial layout failed ({type(error).__name__}); "
                                        "make the call and check again")
            continue
        if effect is not None:
            entry["if_kept" if on else "if_released"] = effect


def left_out(layout: Any, scope: Any, reported: list[dict[str, Any]], why: str) -> None:
    """``effect_left_out`` on each reported range a trial would be made for, where none is
    made (``why``: how to ask for it)."""
    for found, entry in zip(scope.ranges, reported):
        if _wanted(layout, found, entry) is not None:
            entry["effect_left_out"] = why


def _wanted(layout: Any, found: Any, entry: dict) -> "bool | None":
    """``True`` -- try the keep -- for a range with a keep ``fix``; ``False`` -- try releasing
    it -- for one kept, on one page, that starts a page only because of its keep (its first
    block the first drawn on a page after the first, with no page break or section start
    before it); else ``None``."""
    if entry.get("fix") is not None:
        return True
    if not (entry.get("kept") and entry.get("one_page")) or not entry.get("pages"):
        return None
    start = entry["pages"][0]
    if start < 2 or not found.blocks:
        return None
    on_page = layout.blocks_on(start)
    if not on_page or on_page[0] != layout._canon(found.blocks[0]):
        return None
    from .scoped import _starts_page

    document = layout._document
    element = _element(document, found.blocks[0])
    before = _block_before(element) if element is not None else None
    if before is None or _starts_page(document, before, element):
        return None
    return False


def _element(document: Any, identifier: str) -> Any:
    try:
        return document._resolve_block(identifier)[1].element
    except Exception:  # noqa: BLE001 -- a block the copy cannot name is not tried
        return None


def _block_before(element: Any) -> Any:
    """The paragraph or table just before ``element`` among its siblings, else ``None``."""
    node = element.getprevious()
    while node is not None and node.tag not in (_P, _TBL):
        node = node.getprevious()
    return node


class _Index:
    """The body's top-level blocks as docx2svg keys them in the layout checked: each
    block's key by its element, and the keys in the order they are first drawn."""

    def __init__(self, layout: Any) -> None:
        from docx2svg.paths import block_keys

        self.key_of: dict[int, str] = {}
        self.order: list[str] = []
        seen: set[str] = set()
        for page in layout.pages:
            for key in page.blocks:
                if key in seen or block_keys(key)[0] != key:
                    continue
                seen.add(key)
                _, element = layout._key_id(key)
                if element is not None and element is not False:
                    self.key_of.setdefault(id(element), key)
                self.order.append(key)


def _trial(previous: Any, data: bytes, options: dict, layout: Any, index: _Index, found: Any,
           on: bool) -> "dict[str, Any] | None":
    import docx2svg

    from ..edit.copies import blocks_of
    from ..layout import _MEASURES

    document = layout._document
    try:
        blocks = blocks_of(document, found.address)
    except (KeyError, ValueError):
        return None
    keys = [index.key_of.get(id(element)) for _, element in blocks]
    if not keys or None in keys:
        return None
    mine = set(keys)
    last = max(index.order.index(key) for key in keys)
    after = next((key for key in index.order[last + 1:] if key not in mine), None)
    # The change alone, on a copy of the bytes checked: what the call ``fix`` names writes
    # (Document.keep_blocks_together's own changes, read on the document checked, whose
    # styles are resolved already), each on the same element of the copy.
    changes, _ = document._keep_changes(blocks, on)
    if not changes:
        return None
    convert_options = docx2svg.ConvertOptions(**options)
    trial, _, _ = docx2svg.lay_out_docx(_changed(data, document, changes), convert_options, previous=previous,
                                        measures=_MEASURES, keep=False, changed_pages=False)
    now = _Pages(layout.pages, keys, after)
    then = _Pages(trial.pages, keys, after)
    if now.start is None or then.start is None:
        return None
    effect: dict[str, Any] = {"starts": then.start, "ends": then.end, "one_page": then.start == then.end}
    # The page before where the range starts kept: blank at its foot when kept (on: the
    # trial's; off: now's), and the other side's there.
    kept_start, free_start = (then.start, now.start) if on else (now.start, then.start)
    if kept_start > free_start and kept_start > 1:
        page = kept_start - 1
        kept_pages, free_pages = (trial.pages, layout.pages) if on else (layout.pages, trial.pages)
        kept_room, free_room = _room(kept_pages[page - 1]), _room(free_pages[page - 1])
        room_then, room_now = (kept_room, free_room) if on else (free_room, kept_room)
        effect["blank"] = {"page": page, "pt": room_then[0], "share": room_then[1],
                           "now_pt": room_now[0], "now_share": room_now[1]}
    effect["document_pages"] = [len(layout.pages), len(trial.pages)]
    if after is not None and now.after is not None and then.after is not None:
        effect["after"] = [now.after, then.after]
    if not on:
        effect["fix"] = {"tool": "word_format", "calls": [{"targets": [found.address], "keep_together": False}]}
    effect["said"] = _said(effect, on)
    return effect


def _changed(data: bytes, document: Any, changes: list[tuple[Any, dict]]) -> bytes:
    """``data`` with each ``(paragraph, values)`` of ``changes`` -- elements of
    ``document``, opened from ``data`` -- written on the same paragraph of a copy."""
    from ..edit.formatting import write_paragraph
    from ..oxml.package import WordPackage

    package = WordPackage.open(data)
    part = package.document_part()
    root = package.tree(part)
    for element, values in changes:
        steps = []
        node = element
        while node.getparent() is not None:
            parent = node.getparent()
            steps.append(parent.index(node))
            node = parent
        target = root
        for step in reversed(steps):
            target = target[step]
        write_paragraph(target, values)
    package.mark_dirty(part)
    return package.to_bytes()


class _Pages:
    """Where some top-level blocks (their keys) are drawn in one layout's pages: the first
    and last 1-based page, and the first page of the block ``after`` them."""

    def __init__(self, pages: Any, keys: list[str], after: "str | None") -> None:
        wanted = set(keys)
        drawn = [page.number + 1 for page in pages if wanted.intersection(page.blocks)]
        self.start = drawn[0] if drawn else None
        self.end = drawn[-1] if drawn else None
        self.after = None
        if after is not None:
            self.after = next((page.number + 1 for page in pages if after in page.blocks), None)


def _room(page: Any) -> tuple[float, float]:
    """The blank at the foot of a page's body: from the bottom of its last body line to the
    bottom of where the body may be drawn (or to the first note line drawn above it), in
    points, and as a share of that height (two decimals)."""
    from ..layout import PX_PER_PT, _in_body

    top, bottom = Fraction(page.column[2]), Fraction(page.column[3])
    lowest = top
    for line in page.text_lines():
        if _in_body(line):
            lowest = max(lowest, Fraction(line.top + line.pitch))
        elif line.story is None and line.path.startswith(("w:footnote", "w:endnote")):
            bottom = min(bottom, Fraction(line.top))
    height = bottom - top
    blank = max(Fraction(0), bottom - lowest)
    return round(float(blank / PX_PER_PT), 1), round(float(blank / height), 2) if height > 0 else 0.0


def _said(effect: dict[str, Any], on: bool) -> str:
    start, end = effect["starts"], effect["ends"]
    out = ("keeping it whole (this range alone): " if on else "releasing its keep (this range alone): ")
    out += f"starts on page {start}" + (", on one page" if start == end else f", over pages {start}-{end}")
    blank = effect.get("blank")
    if blank is not None:
        out += (f"; page {blank['page']} {'left' if on else 'then'} {round(blank['share'] * 100)}% blank "
                f"({blank['pt']:g} pt), now {round(blank['now_share'] * 100)}%")
    before, after_pages = effect["document_pages"]
    out += ("; the document stays at " + _pages(before) if before == after_pages
            else f"; the document goes from {before} to {_pages(after_pages)}")
    moved = effect.get("after")
    if moved is not None:
        out += (f"; what follows stays on page {moved[0]}" if moved[0] == moved[1]
                else f"; what follows moves from page {moved[0]} to {moved[1]}")
    return out


def _pages(count: int) -> str:
    return f"{count} page" + ("" if count == 1 else "s")
