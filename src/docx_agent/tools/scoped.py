"""``check`` and ``render`` on some blocks (round 35, R18), and copies against their
reference (R19): the parts of the shared tools' Word handlers (:mod:`.shared`) that read a
range or a copy.

**Ranges** (:mod:`docx_agent.scope`): ``check(range=[...])`` reports the facts about the
blocks -- the layout's measured on the pages they fall on, ``validate``'s problems found
in them (:func:`docx_agent.validate.problems_in`) -- and each range's pages;
``render(range=[...])`` draws those pages.  The layout is the document's one, shared and
incremental (:mod:`docx_agent.layout`), laid out with the document's lock let go where the
call may (a lone call), else on its own snapshot: never a wait for a layout under the lock.

**Copies** (:mod:`docx_agent.edit.copies`): each copy compared with its reference, kept by
what both sides hold -- the digests of their blocks (ids and XML) and of each document's
styles and theme -- not by the document's version, so an edit elsewhere in either document
finds the comparison kept -- and (round 42) by the list items before the copy, which decide
the numbers its lists show (:func:`docx_agent.edit.copies.list_context`).
:data:`COPY_CACHE` comparisons are kept per document, the least recently used dropped first.
The reply's ``fix`` is one ``word_follow_sample`` call per reference block; ``list_fix``
(R49) the ``word_lists`` calls of every copy's list differences, alike ones joined (rule 5),
at most :data:`LIST_CALLS` (``more`` counts the rest).
"""

from __future__ import annotations

import copy as _copy
import threading
from collections import OrderedDict
from typing import Any

from ooxml_edit.tools import ToolError

from ..edit import copies as _copies
from ..edit.copy_settings import LIST_SETTINGS
from ..edit.ids import DeletedId
from ..scope import Scope, headings_as_ranges

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_P = _W + "p"

#: Comparisons kept per document (each a few kilobytes: a copy's differences and example text).
COPY_CACHE = 128

#: ``word_lists`` calls in a reply's ``list_fix`` at most (the rest counted: ``more``).
LIST_CALLS = 40

#: Copies in one check at most (a chapter's, or a reviewer's few chapters).
MAX_COPIES = 40

#: Differences listed per copy (the copy's own are counted, not listed); the rest counted
#: in ``more``.
DIFFERENCES_PER_COPY = 20
#: Differences listed in one reply over every copy; past it, counted.
DIFFERENCES_PER_REPLY = 100
#: Places with example text listed per copy, and per reply.
EXAMPLES_PER_COPY = 10
EXAMPLES_PER_REPLY = 50

_CACHE_LOCK = threading.Lock()


# -- ranges ------------------------------------------------------------------------------------


#: Pages a range's ``extent`` lists (rule 4); the rest are counted in ``extent_more``.
EXTENT_PAGES = 12


def range_pages(scope: Scope, layout: Any, extent: bool = False) -> tuple[list[dict[str, Any]], list[int]]:
    """Each range as reported -- ``{"range", "blocks", "pages"}``, with ``extent``
    (:func:`range_extent`) -- and every page any of them is on."""
    out = []
    every: set[int] = set()
    for found, pages in zip(scope.ranges, scope.pages(layout)):
        reported = {"range": found.address, "blocks": len(found.blocks), "pages": pages}
        if extent and pages:
            reported.update(range_extent(layout, found.paragraphs, pages))
            reported.update(range_keeps(scope.document, found, layout))
        out.append(reported)
        every.update(pages)
    return out, sorted(every)


def range_extent(layout: Any, paragraphs: list[str], pages: list[int]) -> dict[str, Any]:
    """How much of its pages a range's body content takes (round 35, R26;
    :meth:`docx_agent.layout.DocumentLayout.extent`): ``extent``, per page ``{"page",
    "top", "bottom", "body"}`` in points -- the top of its first line and the bottom of
    its last there, from the page's top edge, and the page's height less its section's
    top and bottom margins -- with ``area`` (``[top, bottom]``: where the body may be
    drawn) only on a page whose header or footer reaches past its margin and pushes the
    body in, and ``columns`` only where it is drawn in more than one; ``height``, its
    ``bottom - top`` summed over every page.  At most :data:`EXTENT_PAGES` pages are
    listed; ``extent_more`` counts the rest (each block or a shorter span as its own range
    lists them)."""
    extents = layout.extent(paragraphs, pages)
    listed = []
    for found in extents[:EXTENT_PAGES]:
        entry: dict[str, Any] = {"page": found.page, "top": found.top, "bottom": found.bottom, "body": found.body}
        if found.area[1] - found.area[0] < found.body - 0.01:
            entry["area"] = list(found.area)
        if found.columns > 1:
            entry["columns"] = found.columns
        listed.append(entry)
    out: dict[str, Any] = {"extent": listed, "height": round(sum(e.bottom - e.top for e in extents), 2)}
    if len(extents) > EXTENT_PAGES:
        out["extent_more"] = len(extents) - EXTENT_PAGES
    return out


#: Page breaks a range's ``breaks`` lists (rule 4); the rest are counted in ``breaks_more``.
BREAKS_LISTED = 10


def range_keeps(document: Any, found: Any, layout: Any) -> dict[str, Any]:
    """Whether a range of whole body blocks is drawn on one page and set to stay so (round
    36, R30: the caller's spans, nothing stored in the document): ``one_page``; ``kept``
    -- every block holds the keeps ``word_format`` ``keep_together`` writes
    (:meth:`~docx_agent.Document.keep_blocks_together`); and where it is not on one page,
    ``breaks`` -- each page break in it: ``{"after", "before", "page"}`` between two
    blocks (``page_break``: the later one starts a page anyway -- a page break, a section
    start -- which no keep undoes), ``{"inside", "pages"}`` within one (at most
    :data:`BREAKS_LISTED`, ``breaks_more`` the rest) -- and ``fix`` (that call) where its
    height fits the body of a page and a keep would close a break, else
    ``taller_than_page`` (it cannot be kept: Word starts it on a new page and breaks it
    where the page is full).  ``{}`` for a range that is not whole blocks (a
    cell, a text range) or not drawn whole."""
    from ..edit.copies import blocks_of

    try:
        blocks = blocks_of(document, found.address)
    except (KeyError, ValueError):
        return {}
    held = {id(element) for element in found.elements}
    for _, element in blocks:
        inside = [element] if element.tag == _P else list(element.iter(_P))
        if not all(id(paragraph) in held for paragraph in inside):
            return {}
    drawn = [layout.pages_of(identifier) for identifier, _ in blocks]
    if not blocks or not all(drawn):
        return {}
    changes, _ = document._keep_changes(blocks, True)
    out: dict[str, Any] = {"one_page": len({p for pages in drawn for p in pages}) == 1, "kept": not changes}
    if out["one_page"]:
        return out
    breaks: list[dict[str, Any]] = []
    for k, ((identifier, element), pages) in enumerate(zip(blocks, drawn)):
        if k and max(drawn[k - 1]) < min(pages):
            gap = {"after": blocks[k - 1][0], "before": identifier, "page": min(pages)}
            if _starts_page(document, blocks[k - 1][1], element):
                gap["page_break"] = True
            breaks.append(gap)
        if len(pages) > 1:
            breaks.append({"inside": identifier, "pages": pages})
    out["breaks"] = breaks[:BREAKS_LISTED]
    if len(breaks) > BREAKS_LISTED:
        out["breaks_more"] = len(breaks) - BREAKS_LISTED
    extents = layout.extent(found.paragraphs)
    height = sum(e.bottom - e.top for e in extents)
    room = min((e.area[1] - e.area[0] for e in extents), default=0.0)
    if extents and height > room + 0.5:
        out["taller_than_page"] = True
    elif not out["kept"] and not all(b.get("page_break") for b in breaks):
        out["fix"] = {"tool": "word_format", "calls": [{"targets": [found.address], "keep_together": True}]}
    return out


def _starts_page(document: Any, before: Any, element: Any) -> bool:
    """Whether a block starts a page whatever keeps it: a page break before it (its style's
    or its own), a page break at its start or at the end of the paragraph before it, or a
    section break there."""
    from ..edit import effective as _effective
    from ..layout_facts import _page_break_at

    if element.tag == _P and (_effective.paragraph(document, element).page_break_before
                              or _page_break_at(element, start=True)):
        return True
    if before.tag == _P and (_page_break_at(before, start=False) or before.find(f"{_W}pPr/{_W}sectPr") is not None):
        return True
    return False


def problems_in_scope(document: Any, scope: Scope, problems: list) -> tuple[list, int]:
    """The problems found in the scope's blocks, and how many are elsewhere."""
    from ..validate import problems_in

    inside = problems_in(problems, document.package.document_part(), scope.roots)
    return inside, len(problems) - len(inside)


# -- copies ------------------------------------------------------------------------------------


def _documents(call: Any, item: dict, number: int) -> tuple[str, Any]:
    reference = item.get("reference")
    other = reference.get("doc") if isinstance(reference, dict) else None
    other = other or call.entry.doc_id
    if other not in call.entries:
        raise ToolError("not_found", f"copies[{number}].reference.doc {other!r} is not open here; open it "
                        "first (open_document) or name one of these", field="copies",
                        valid_options=list(call.entries)[:50])
    return other, call.entries[other]


def _blocks(document: Any, address: Any, where: str) -> list[tuple[str, Any]]:
    from ..edit.addresses import AddressError
    from ..edit.errors import EditError

    if isinstance(address, list):
        address = [a for a in address if isinstance(a, str) and a.strip()]
        if len(address) == 1:
            address = address[0]
        elif address:
            found: list[tuple[str, Any]] = []
            for item in address:
                found += [b for b in _blocks(document, item, where) if all(b[1] is not f[1] for f in found)]
            return found
    if not isinstance(address, str) or not address.strip():
        raise ToolError("invalid_arguments", f"{where} is empty: a block (p:..., t:...) or a span p:A..t:B",
                        field="copies", valid_options=headings_as_ranges(document)[:50])
    try:
        found = _copies.blocks_of(document, address.strip())
    except AddressError as error:
        raise ToolError("invalid_arguments", f"{where} {address!r}: {error.args[0] if error.args else error}",
                        field="copies", valid_options=list(error.options or headings_as_ranges(document))[:50]) \
            from None
    except DeletedId as error:                   # round 37: gone, and by which change
        raise ToolError("not_found", f"{where}: {error}", field="copies",
                        valid_options=[error.replaced_by] if error.replaced_by else
                        headings_as_ranges(document)[:50]) from None
    except (KeyError, IndexError, EditError) as error:
        raise ToolError("not_found", f"{where} {address!r} names no block of the body "
                        f"({error.args[0] if error.args else error}); give a block or a span p:A..t:B, "
                        "or a heading's range h:p:A", field="copies",
                        valid_options=headings_as_ranges(document)[:50]) from None
    if not found:
        raise ToolError("not_found", f"{where} {address!r} holds no paragraph or table", field="copies",
                        valid_options=headings_as_ranges(document)[:50])
    return found


def _styles_digest(entry: Any) -> str:
    """The document's styles digest, kept for its version (read once per version)."""
    document = entry.document
    kept = getattr(document, "_tool_styles_digest", None)
    if kept is not None and kept[0] == entry.version:
        return kept[1]
    found = _copies.styles_digest(document)
    document._tool_styles_digest = (entry.version, found)
    return found


def _cache(document: Any) -> "OrderedDict[tuple, dict]":
    with _CACHE_LOCK:
        cache = getattr(document, "_tool_copy_cache", None)
        if cache is None:
            cache = document._tool_copy_cache = OrderedDict()
        return cache


def compared(call: Any, copy_blocks: list, other: str, other_entry: Any, reference_blocks: list
             ) -> tuple[dict, bool]:
    """One copy's comparison (:func:`docx_agent.edit.copies.compare_copy`), from the cache
    when neither side's blocks nor either document's styles changed: ``(result, kept)``."""
    document = call.document
    key = (_copies.digest(document, copy_blocks), _styles_digest(call.entry), other,
           _copies.digest(other_entry.document, reference_blocks), _styles_digest(other_entry),
           # Round 42 (R49): the numbers the lists show, which the items before decide.
           _copies.list_context(document, copy_blocks), _copies.list_context(other_entry.document, reference_blocks))
    cache = _cache(document)
    with _CACHE_LOCK:
        found = cache.get(key)
        if found is not None:
            cache.move_to_end(key)
            return found, True
    found = _copies.compare_copy(document, copy_blocks, other_entry.document, reference_blocks)
    with _CACHE_LOCK:
        cache[key] = found
        cache.move_to_end(key)
        while len(cache) > COPY_CACHE:
            cache.popitem(last=False)
    return found, False


def compare_copies(call: Any, copies: list) -> dict[str, Any]:
    """``check``'s ``copies``: each copy compared with its reference, the reply within its
    budget (rule 4: what is left out is counted, and how to read it said)."""
    if not isinstance(copies, list):
        copies = [copies]
    items = [c for c in copies if c not in (None, {}, "")]        # a strict decoder's placeholders
    if len(items) > MAX_COPIES:
        raise ToolError("limit", f"{len(items)} copies: at most {MAX_COPIES} per check; split them over "
                        "two checks (one per chapter)", field="copies")
    doc = call.entry.doc_id
    out: list[dict[str, Any]] = []
    groups: dict[tuple[str, str], list[str]] = {}
    room, examples_room = DIFFERENCES_PER_REPLY, EXAMPLES_PER_REPLY
    kept_count = 0
    list_calls: list[dict] = []
    for number, item in enumerate(items):
        if not isinstance(item, dict) or not isinstance(item.get("reference"), dict):
            raise ToolError("invalid_arguments", f"copies[{number}] is {{\"blocks\": the copy's block or span, "
                            "\"reference\": {\"doc\": its document (default doc), \"block\": its block or "
                            "span}}", field="copies")
        reference = item["reference"]
        other, other_entry = _documents(call, item, number)
        copy_blocks = _blocks(call.document, item.get("blocks"), f"copies[{number}].blocks")
        reference_blocks = _blocks(other_entry.document, reference.get("block"),
                                   f"copies[{number}].reference.block")
        found, kept = compared(call, copy_blocks, other, other_entry, reference_blocks)
        kept_count += kept
        entry: dict[str, Any] = {"copy": _said(item.get("blocks")), "reference": _said(reference.get("block"))}
        if other != doc:
            entry["doc"] = other
        entry["follows"] = found["follows"]
        listed = [{k: v for k, v in d.items()} for d in found["differences"] if d["kind"] != "own"]
        own = len(found["differences"]) - len(listed)
        shown = listed[:min(DIFFERENCES_PER_COPY, room)]
        room -= len(shown)
        entry["differences"] = shown
        if own:
            entry["own"] = own
        if len(listed) > len(shown):
            entry["more"] = len(listed) - len(shown)
        texts = found["example_text"]
        shown_texts = texts[:min(EXAMPLES_PER_COPY, examples_room)]
        examples_room -= len(shown_texts)
        if texts:
            entry["example_text"] = _copy.deepcopy(shown_texts)
            if len(texts) > len(shown_texts):
                entry["example_text_more"] = len(texts) - len(shown_texts)
        entry["cached"] = kept
        for copy_id, reference_id in found["pairs"]:
            # word_follow_sample keeps a copy's lists (R49: word_lists makes them, list_fix).
            if any(d["kind"] == "format" and d.get("block") == copy_id and d["property"] not in LIST_SETTINGS
                   for d in found["differences"]):
                groups.setdefault((other, reference_id), []).append(copy_id)
        for difference in found["differences"]:
            if difference["property"] in LIST_SETTINGS and difference.get("fix"):
                list_calls += difference["fix"]["calls"]
        out.append(_copy.deepcopy(entry))
    report: dict[str, Any] = {
        "compared": len(out), "follow": sum(1 for e in out if e["follows"]),
        "with_example_text": sum(1 for e in out if e.get("example_text")),
        "kept": kept_count, "items": out}
    if groups:
        calls = []
        for (other, reference_id), copy_ids in groups.items():
            sample: dict[str, Any] = {"block": reference_id}
            if other != doc:
                sample["doc"] = other
            calls.append({"sample": sample, "copies": list(dict.fromkeys(copy_ids))})
        # Rule 5: one call per reference block makes every copy of it follow (preview first).
        report["fix"] = {"tool": "word_follow_sample", "calls": calls}
    if list_calls:
        # Rule 5: every copy's list differences in one batch of word_lists calls, alike ones joined.
        merged = _merged(list_calls)
        report["list_fix"] = {"tool": "word_lists", "calls": merged[:LIST_CALLS]}
        if len(merged) > LIST_CALLS:
            report["list_fix"]["more"] = len(merged) - LIST_CALLS
    more = sum(e.get("more", 0) for e in out)
    more_examples = sum(e.get("example_text_more", 0) for e in out)
    if more or more_examples:
        report["left_out"] = (f"{more} difference(s) and {more_examples} place(s) with example text past "
                              "the reply's budget, counted per copy ('more', 'example_text_more'): check fewer "
                              "copies at once, or word_compare_blocks on one copy's block for all its "
                              "differences, paged")
    return report


def _merged(calls: list[dict]) -> list[dict]:
    """``word_lists`` calls with the same arguments but their targets made one, in order --
    those of an action that changes each target (``level``, ``add``, ``remove``); the
    others (``format``, ``restart``, ``continue``: the first target's list) kept apart."""
    out: list[dict] = []
    joined: dict[str, dict] = {}
    for call in calls:
        if call.get("action") not in ("level", "add", "remove"):
            out.append(_copy.deepcopy(call))
            continue
        key = repr(sorted((k, v) for k, v in call.items() if k != "targets"))
        found = joined.get(key)
        if found is None:
            found = joined[key] = _copy.deepcopy(call)
            out.append(found)
        else:
            found["targets"] += [t for t in call["targets"] if t not in found["targets"]]
    return out


def _said(address: Any) -> Any:
    if isinstance(address, list):
        return [a for a in address if isinstance(a, str)]
    return address.strip() if isinstance(address, str) else address
