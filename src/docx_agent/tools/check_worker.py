"""``check``'s layout facts made where the layout is held (round 38).

A check's facts about the layout -- the page count, the coverage, each range's pages and
extent, the layout facts of the pages asked for -- were read in the caller's process from a
layout pickled to it from the worker that made it, and mapped to ids there: about 0.4 s to
read the pickle and as much again to map it, for every check, in the one interpreter every
agent's calls share.  :func:`check_bytes` runs in the worker holding the document's layout
(:mod:`docx_agent.layout`, "layouts held where they were made") and sends back only the
facts: it maps the layout to ids on a copy of the document opened from the very bytes laid
out (ids are the bytes' own: ``w14:paraId``, or a paragraph's place), so the facts are the
ones the caller would have read, for the version it checked.  The caller keeps what stays
with the document: the problems (``validate``), the fields, the comparison with the last
check (from a small :class:`~docx_agent.layout.LayoutSummary`, its ids followed through the
document's aliases there).
"""

from __future__ import annotations

from typing import Any

from ..layout import DocumentLayout, LayoutSummary, _in_body, _put_held, held_layout, made_facts


def _opened(held: Any, data: bytes, options: dict) -> DocumentLayout:
    """The layout of ``held`` mapped to ids on a copy of the document opened from ``data``,
    made once per layout held (kept with it)."""
    found = held.extra.get("layout")
    if found is None:
        from ..edit.document import Document, WordPackage

        # Opened as the package it is (a template too, without Document.open's warning: the
        # caller opened it already, and said so).
        owner = Document(WordPackage.open(data))
        owner.font_dirs = options.get("font_dirs")
        owner.locale = options.get("locale")
        layout, _, warnings = held.laid
        found = held.extra["layout"] = DocumentLayout(owner, layout, warnings)
    return found


def check_bytes(data: bytes, options: dict, state: tuple, request: dict) -> dict:
    """In a worker: the layout of ``data`` -- the one held, or laid out from the one found
    (:func:`docx_agent.layout.held_layout`) -- and the facts ``request`` asks of it:

    * always ``reflow``: ``pages``, ``complete``, ``stopped``, ``coverage`` (as ``check``
      reports them) and ``warnings`` (``layout_warnings``' text); ``stats``, the pages and
      body lines ``app.xml`` is saved with;
    * ``summary`` (asked): the layout's :class:`~docx_agent.layout.LayoutSummary`;
    * ``ranges`` (each range's address and paragraph ids, :meth:`docx_agent.scope.Scope.spec`):
      ``range`` (each as reported, with its extent) and ``on`` (every page they are on);
      with ``keeps`` (round 40), each range's keep tried (:mod:`.keep_trials`);
    * ``layout`` (asked): the layout facts of ``pages`` -- the ranges' pages among them
      when ranges are given -- with ``min_size`` and ``styles``.

    Plus how the layout was found or made (:func:`docx_agent.layout.made_facts`)."""
    held, how, seconds = held_layout(data, options, state)
    try:
        out = made_facts(held, how, seconds)
        out.update(_facts(held, data, options, request))
    finally:
        _put_held(state[1], held, state[3])    # a layout of these bytes either way
    return out


# A task ooxml-edit's subprocess workers may run (the mark, as the layout's tasks set it).
check_bytes.__ooxml_edit_task__ = True  # type: ignore[attr-defined]


def _facts(held: Any, data: bytes, options: dict, request: dict) -> dict:
    from ..layout_facts import layout_facts
    from ..scope import Scope
    from ._base import short
    from .scoped import range_pages

    layout = _opened(held, data, options)
    reflow: dict[str, Any] = {"pages": layout.page_count, "complete": layout.pages_known is not None}
    if layout.stopped is not None:
        reflow["stopped"] = {"page": layout.stopped.page, "reason": layout.stopped.reason,
                             "at": layout.stopped.at}
    reflow["coverage"] = layout.coverage_facts()
    out: dict[str, Any] = {
        "reflow": reflow,
        "warnings": [short(str(w), 160) for w in layout.warnings[:10]] if layout.warnings else None,
        "stats": ({"pages": layout.pages_known,
                   "lines": sum(1 for page in layout.pages for line in page.text_lines() if _in_body(line))}
                  if layout.pages_known is not None else {})}
    if request.get("summary"):
        summary = held.extra.get("summary")
        if summary is None:
            summary = held.extra["summary"] = LayoutSummary(layout, held.extra.get("summary before"), keep=True)
        out["summary"] = summary
    pages = request.get("pages")
    scoped_pages = pages
    scope = None
    if request.get("ranges"):
        scope = Scope.of_ids(layout._document, request["ranges"])
        out["range"], on = range_pages(scope, layout, extent=True)
        out["on"] = on
        if request.get("keeps"):
            # Round 40: what each range's keep would do, laid out from this very layout.
            from .keep_trials import keep_effects

            caps = request["keeps"]
            keep_effects(held.laid[0], data, options, layout, scope, out["range"],
                         trials=caps["trials"], seconds=caps["seconds"])
        scoped_pages = on if pages is None else [p for p in on if p in pages]
    if request.get("layout"):
        out["layout"] = layout_facts(layout, pages=scoped_pages, min_size=request.get("min_size"),
                                     styles=request.get("styles"), scope=scope)
    return out
