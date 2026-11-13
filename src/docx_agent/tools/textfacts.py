"""Text written by reference, and the texts a check of numbers reads (round 36, R28).

**Keeping what was written by reference.**  The dispatcher writes ``{{d:by_topic/door/mean}}``
as the value before a text tool runs (:mod:`ooxml_edit.tools.textfacts`; ``call.facts``
says where in each argument).  :func:`keep_written` then finds each value in the paragraphs
the call wrote -- the result's changed, created and renamed blocks, ``replace_text``'s
paragraphs, ``word_set_text``'s targets -- by the text around it in the argument (spaces and
Markdown's marks aside), and keeps it on its paragraph
(:meth:`docx_agent.Document.keep_text_facts`), in the call's undo step.  A value found
nowhere, or in two places alike, is not kept: a figure is listed as typed rather than kept
under a wrong place.  ``replace_text`` writes its replacement at every match, so each
place keeps the value.

**The texts a check of numbers reads** (:func:`number_texts`, the docx
``DocumentFormat.number_texts``): the body's paragraphs, one text each, and each table
cell's paragraphs as one text (``t:X/c1,2``), in document order, as ``find_text`` reads
them -- deleted text left out, a note's or a drawing's mark read as a space.  Left out: the
tables placed with ``from`` (they keep their record), a table of contents' entries, the
headers, footers and notes.  The spans kept as written by reference are marked.  ``range``
narrows it to blocks, as ``check`` resolves a range (:class:`docx_agent.scope.Scope`);
``pages`` to the pages the layout puts them on; each text says its page.
"""

from __future__ import annotations

import re
from typing import Any, Mapping

from ooxml_edit.tools import NumberText, ToolError

from ..edit import text as _text
from ..edit.document import Cell, Paragraph, Table, _cells, _find, _int
from ..edit.fields import _toc_paragraphs, scan

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_P, _TC, _TR, _TBL = _W + "p", _W + "tc", _W + "tr", _W + "tbl"

#: How far around a value its argument's text is compared with a paragraph's.
CONTEXT = 40

#: Characters Markdown writes that a paragraph's text does not have.
_MARKS = set("*_`|#>\\")


# -- keeping what was written ----------------------------------------------------------------------


def _argument(arguments: Mapping[str, Any], field: str) -> Any:
    """The argument a fact's ``field`` names (``items[0].text``)."""
    node: Any = arguments
    for name, index in re.findall(r"([^.\[\]]+)|\[(\d+)\]", field):
        try:
            node = node[int(index)] if index else node[name]
        except (KeyError, IndexError, TypeError):
            return None
    return node


def _squeezed(text: str) -> str:
    """``text`` without spaces of any kind and Markdown's marks: what is compared."""
    return "".join(ch for ch in text if not ch.isspace() and ch not in _MARKS)


def _common(a: str, b: str) -> int:
    count = 0
    for x, y in zip(a, b):
        if x != y:
            break
        count += 1
    return count


def _touched(call: Any, returned: Any, arguments: Mapping[str, Any]) -> list[str]:
    """The paragraphs the call wrote: its result's blocks and the targets it was given,
    tables and cells as their paragraphs."""
    named: list[str] = []
    for key in ("changed", "created"):
        named += [str(v) for v in getattr(returned, key, None) or ()]
    named += [str(v) for v in (getattr(returned, "renamed", None) or {}).values()]
    data = getattr(returned, "data", None)
    if isinstance(data, Mapping):
        named += [str(v) for v in data.get("paragraphs") or () if isinstance(v, str)]
    for target in [arguments.get("target")] + [item.get("target") for item in arguments.get("items") or ()
                                               if isinstance(item, Mapping)]:
        if isinstance(target, str):
            named.append(target)
    document = call.document
    out: list[str] = []
    for name in dict.fromkeys(named):
        name = name.split("@")[0] if "@" in name and ".." not in name else name
        try:
            found = document.get(name)
        except (KeyError, ValueError, IndexError):
            continue
        if isinstance(found, Paragraph):
            out.append(found.id)
        elif isinstance(found, Cell):
            out += [p.id for p in found.paragraphs]
        elif isinstance(found, Table):
            for line in found.grid():
                for cell in line:
                    if cell is not None:
                        out += [p.id for p in cell.paragraphs]
    return list(dict.fromkeys(out))


def keep_written(call: Any, returned: Any, arguments: Mapping[str, Any]) -> int:
    """Keep each value the call wrote by reference on the paragraph it went into (the
    module's notes); how many places were kept."""
    facts = list(getattr(call, "facts", None) or ())
    if not facts or call.entry is None:
        return 0
    document = call.document
    paragraphs = []
    for identifier in _touched(call, returned, arguments):
        try:
            paragraphs.append((identifier, document.paragraph(identifier).text))
        except (KeyError, ValueError):
            continue
    taken: set[tuple[str, int]] = set()
    kept: dict[str, list[dict[str, Any]]] = {}
    count = data_count(returned)
    for fact in facts:
        argument = _argument(arguments, fact.field)
        if not isinstance(argument, str) or argument[fact.start:fact.end] != fact.text or not fact.text:
            continue
        before = argument[max(0, fact.start - CONTEXT):fact.start].rpartition("\n")[2]
        after = argument[fact.end:fact.end + CONTEXT].partition("\n")[0]
        left, right = _squeezed(before)[::-1], _squeezed(after)
        found = []
        for identifier, text in paragraphs:
            at = text.find(fact.text)
            while at >= 0:
                end = at + len(fact.text)
                if not (at and text[at - 1].isdigit()) and not (end < len(text) and text[end].isdigit()) \
                        and (identifier, at) not in taken:
                    a = _common(_squeezed(text[max(0, at - CONTEXT):at])[::-1], left)
                    b = _common(_squeezed(text[end:end + CONTEXT]), right)
                    found.append((a + b, a == len(left) and b == len(right), identifier, at))
                at = text.find(fact.text, at + 1)
        if not found:
            continue
        found.sort(key=lambda f: -f[0])
        whole = [f for f in found if f[1]]
        copies = count if fact.field == "replace" else 1
        if whole:
            chosen = whole[:copies] if len(whole) <= copies or copies > 1 else []
        else:
            chosen = [found[0]] if len(found) == 1 or found[0][0] > found[1][0] else []
        for _score, _whole, identifier, at in chosen:
            taken.add((identifier, at))
            kept.setdefault(identifier, []).append({**fact.record, "p": 0, "at": at})
    places = 0
    for identifier, found_facts in kept.items():
        if document.keep_text_facts(identifier, found_facts):
            places += len(found_facts)
    return places


def data_count(returned: Any) -> int:
    """How many places a call wrote its replacement in (``replace_text``'s ``count``)."""
    data = getattr(returned, "data", None)
    count = data.get("count") if isinstance(data, Mapping) else None
    return count if isinstance(count, int) and count > 0 else 1


# -- the texts a check of numbers reads ------------------------------------------------------------


def number_texts(call: Any, scope: Mapping[str, Any]) -> list[NumberText]:
    """``DocumentFormat.number_texts`` for a Word document (the module's notes)."""
    from ..scope import Scope, ranges_of
    from ._base import engage, release, resolve
    from .shared import _pages_of

    document = call.document
    pages = _pages_of(call, scope.get("slides"), scope.get("pages"))
    ranges = ranges_of(resolve(call, scope.get("range")))
    held = Scope(document, ranges) if ranges else None
    main = document.package.document_part()
    index = document._index(main)
    toc = {id(p) for found in scan(main, document.package.tree(main)) if found.keyword == "TOC"
           for p in _toc_paragraphs(found)}
    placed = set()
    for key in document.provenance():
        found = document._direct(key) if key.rpartition("/")[2].startswith("t:") else None
        if found is not None:
            placed.add(id(found[1].element))
    written = document.text_facts()
    groups: dict[Any, list[tuple[Any, str, tuple[tuple[int, int], ...]]]] = {}
    order: list[Any] = []
    for entry in index.paragraphs:
        element = entry.element
        if id(element) in toc or any(id(t) in placed for t in element.iterancestors(_TBL)):
            continue
        if held is not None and entry.id not in held.paragraphs:
            continue
        text = _text.paragraph_text(element).replace(_text.OBJECT, " ")
        spans = tuple((f["at"], f["at"] + len(f["text"])) for f in written.get(entry.id, ()))
        cell = next(element.iterancestors(_TC), None)
        key = id(cell) if cell is not None else id(element)
        if key not in groups:
            groups[key] = []
            order.append((key, cell))
        groups[key].append((entry, text, spans))
    # Each text with digits, read before the layout: where the call lets the document go
    # while it is laid out, another call may change it meanwhile.
    found: list[tuple[str, str, tuple[tuple[int, int], ...], list[str]]] = []
    for key, cell in order:
        members = groups[key]
        joined = "\n".join(text for _, text, _ in members)
        if not any(ch.isdigit() for ch in joined):
            continue
        at = members[0][0].id if cell is None else _cell_id(index, cell)
        if at is None:
            continue
        marked: list[tuple[int, int]] = []
        offset = 0
        for _, text, own in members:
            marked += [(offset + a, offset + b) for a, b in own]
            offset += len(text) + 1
        found.append((at, joined, tuple(marked), [entry.id for entry, _, _ in members]))
    # Round 57 (R124): only those texts' pages, asked for where the layout is held (a check
    # of a 300-page report pulled the whole layout here to name each text's first page:
    # seconds of the interpreter every call shares).
    engage(call)
    try:
        layout = _layout(call, [i for *_, ids in found for i in ids], required=pages is not None)
    finally:
        release(call.entries.values())
    out: list[NumberText] = []
    for at, joined, marked, ids in found:
        where: dict[str, Any] = {}
        if layout is not None:
            on = sorted({n for identifier in ids for n in layout.pages_of(identifier)})
            if pages is not None and not set(on) & set(pages):
                continue
            if on:
                where["page"] = on[0]
        out.append(NumberText(at, joined, marked, {"range": at}, where))
    return out


class _PagesOf:
    """The pages of some blocks, as :meth:`~docx_agent.layout.DocumentLayout.pages_of` says
    them, read where the layout is held (round 57, R124)."""

    def __init__(self, pages: dict[str, list[int]]) -> None:
        self._pages = pages

    def pages_of(self, identifier: str) -> list[int]:
        return list(self._pages.get(identifier, ()))


def _layout(call: Any, ids: list[str], required: bool) -> Any:
    """The document's layout -- anything with ``pages_of`` for ``ids`` -- laid out as
    ``check`` lays it out (its lock let go where the call may): the pages of ``ids`` read in
    the worker holding the layout where it can (:func:`docx_agent.layout.check_facts`),
    else the layout itself; ``None`` where it failed, unless ``required`` (``pages``
    given): refused."""
    from ..layout import check_facts, lay_out_unlocked

    entry, document = call.entry, call.document
    version = entry.version
    try:
        if call.may_unlock:
            request = {"pages": None, "ranges": None, "layout": False, "min_size": None, "styles": None,
                       "keeps": None, "pages_of": sorted(set(ids))}
            made = check_facts(document, call.unlocked, request)
            if made is not None:
                return _PagesOf(made[0].get("pages_of") or {})
            return lay_out_unlocked(document, call.unlocked, lambda: entry.version == version)
        return document.layout()
    except ToolError as error:
        if required:
            raise ToolError(error.code, f"numbers on pages need the layout: {error.message}; leave pages "
                            "out to read the whole body, or give range", field="pages") from None
        call.warn(f"numbers: the layout failed ({error.code}): no pages said")
        return None


def _cell_id(index: Any, cell: Any) -> str | None:
    """A ``w:tc``'s id, ``t:X/c<row>,<column>`` (its grid column), or ``None``."""
    row = next(cell.iterancestors(_TR), None)
    table = next(cell.iterancestors(_TBL), None)
    found = index.entry_for(table) if table is not None else None
    if row is None or found is None:
        return None
    number = next((k for k, (element, _) in enumerate(found.rows) if element is row), None)
    if number is None:
        return None
    column = _int(_find(row, "w:trPr/w:gridBefore"), 0)
    for tc in _cells(row):
        if tc is cell:
            return f"{found.id}/c{number},{column}"
        column += max(1, _int(_find(tc, "w:tcPr/w:gridSpan"), 1))
    return None
