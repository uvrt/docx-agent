"""The Word handlers of the shared tools (``ooxml_edit.tools.shared``): one definition each
for every format, dispatched by the document's kind.

S1 ``open_document``, S4 ``close_document``, S5 ``undo`` and ``batch`` are
the session's own (``shared.SESSION_TOOLS``); this module adds docx's ``new_document``,
``save_document``, ``find_text``, ``replace_text``, ``render``, ``check``, ``edit_chart`` (its
``read`` action too), ``edit_smartart``, ``set_properties`` and ``describe`` (its handler,
``word_describe``, is in :mod:`.read`)."""

from __future__ import annotations

import time
from contextlib import contextmanager
from typing import Any, Iterator

from ooxml_edit.tools import Result, ToolError, shared, tiles, worker_task
from ooxml_edit.tools import render as pipeline
from ooxml_edit.tools.render import png_size

from ..edit.document import Document
from ..edit.properties import refreshed_app
from ..layout import (affinity, canon_of, check_facts, font_dirs_of, held_state, keep_layout, kept_layout,
                      lay_out_unlocked, layout_key, layout_options, locale_of, note_made, summary_for)
from ..layout_facts import CAP, empty_paragraphs, layout_facts, numbering_skips_level, punctuation_spacing
from ..scope import MAX_RANGES, Scope, ranges_of, read_scope
from .render import RENDERER
from . import keep_trials
from .keep_trials import left_out
from .scoped import compare_copies, problems_in_scope, range_pages
from ._base import (Timing, cached_layout, outcome, page, reflow_json, remember, resolve, short, validate_delta,
                    wrap)

KIND = "docx"
PAGE_SIZES = ("A4", "Letter", "Legal", "A5", "A3")
DEFAULT_WIDTH = 1000


def _pages_of(call: Any, slides: list | None, pages: list[int] | None) -> list[int] | None:
    """``pages``, or ``slides`` read as pages: a model that renders decks too often names a
    document's pages ``slides`` (10 of trial 3's 18 Word runs lost their first render to
    it).  Given both, ``pages`` is the document's and ``slides`` is ignored.  The result's
    warnings say which."""
    if slides is None:
        return pages
    if pages is not None:
        call.warn("slides ignored: a document has pages")
        return pages
    call.warn("slides read as pages: a document has pages")
    out = []
    for index, item in enumerate(slides):
        # A deck's slides may be ids (s:256); a document's pages are numbers only.
        if isinstance(item, str) and not item.strip().isdigit():
            raise ToolError("invalid_arguments", f"slides[{index}] is {item!r}: a document has pages, "
                            "numbered from 1 (pages: [1])", field="pages")
        out.append(int(item))
    return out


def _decks_only(name: str, use: str) -> ToolError:
    return ToolError("invalid_arguments", f"{name} is for decks; for a document use {use}", field=name)


def handler(name: str, *, mutates: bool = False, track: bool = True):
    """``shared.handler(name, kind="docx")`` around a handler wrapped as every Word tool is."""

    def make(fn):
        return shared.handler(name, kind=KIND)(wrap(fn, mutates=mutates, track=track))

    return make


# -- S2 new_document ---------------------------------------------------------------------------


@handler("new_document")
def new_document(call: Any, kind: str, template_blob: str | None = None, size: str | None = None,
                 title: str | None = None, author: str | None = None, name: str | None = None) -> Result:
    if size is not None and size not in PAGE_SIZES:
        raise ToolError("invalid_arguments", f"a document's size is a page size, not {size!r}", field="size",
                        valid_options=list(PAGE_SIZES))
    template = call.blob(template_blob).data if template_blob else None
    document = Document.new(page=size, title=title, author=author, created=call.now(), template=template,
                            keep_content=False, locale=call.locale)
    doc_id = call.session.adopt(document, KIND, name or "new.docx", source=template_blob or "new",
                                size=len(template or b""))
    entry = call.session.entry(doc_id)
    return Result(summary=f"Made {entry.name} as {doc_id}", created=[doc_id],
                  data={**entry.describe(), "summary": _summary(document)})


def _summary(document: Document) -> dict[str, Any]:
    from .format import summary

    return summary(document)


# -- S3 save_document --------------------------------------------------------------------------


def serialise(document: Document, kind: str) -> bytes:
    """The document as a file of ``kind`` (docx, docm, dotx, dotm), as ``save`` writes it:
    the main part's content type set to the kind's (the rest of the package as it is) and
    ``app.xml``'s statistics refreshed."""
    from ..edit.authoring import KINDS, content_types_as, has_macros
    from ..edit.properties import REL_APP, _part
    from ..oxml.package import MAIN_CONTENT_TYPES

    package = document.package
    if not KINDS[kind].macro_enabled and has_macros(package):
        raise ToolError("refused", f"the document holds macros, which a .{kind} cannot carry")
    replacements = {}
    if package.kind != kind:
        replacements["[Content_Types].xml"] = content_types_as(package, MAIN_CONTENT_TYPES[kind])
    app = refreshed_app(document)
    if app is not None:
        replacements[_part(package, REL_APP)] = app
    return package.to_bytes(replacements or None)


@handler("save_document")
def save_document(call: Any, doc: str, name: str, format: str, clean: bool = False) -> Result:
    document = call.document
    if format in ("pptx", "potx", "outline"):
        raise ToolError("invalid_arguments", f"{format} is a deck format; a document saves as docx, dotx or "
                        "markdown", field="format", valid_options=["docx", "dotx", "markdown"])
    if format not in ("docx", "dotx", "markdown"):
        raise ToolError("invalid_arguments", f"a document saves as docx, dotx or markdown, not {format}",
                        field="format", valid_options=["docx", "dotx", "markdown"])
    if format == "markdown":
        if clean:
            raise ToolError("invalid_arguments", "clean is for a docx or dotx; a Markdown copy carries none "
                            "of it: leave clean out", field="clean")
        data = document.to_markdown(ids=False, stories="all").encode("utf-8")
        return Result(summary=f"Saved {name} (Markdown) for the application",
                      data=call.output(name, "markdown", data))
    extension = name.lower().rpartition(".")[2] if "." in name else ""
    if extension in ("docx", "dotx", "docm", "dotm") and extension[:3] != format[:3]:
        raise ToolError("invalid_arguments", f"{name} names a .{extension} but format is {format}: Word refuses a "
                        "file whose kind and extension disagree", field="name")
    kind = format
    if document.package.kind in ("docm", "dotm") and extension in ("docm", "dotm"):
        kind = extension
    data = serialise(document, kind)
    cleaned = None
    if clean:
        from ..edit.cleaning import clean as clean_bytes

        data, cleaned = clean_bytes(data)
    # Round 38: the output carries the document's id state, for opening these bytes again.
    result = Result(summary=f"Saved {name} for the application",
                    data=call.output(name, format, data, state=True))
    if cleaned is not None:
        # What clean took out (round 16, F11): counts and bytes per kind, the file's sizes.
        result.data["cleaned"] = cleaned
        before, after = cleaned["bytes"]["before"], cleaned["bytes"]["after"]
        result.summary += f", cleaned: {before:,} -> {after:,} bytes"
    result.data["coverage"] = _coverage(call)
    return result


def _coverage(call: Any) -> dict[str, Any]:
    """How much of the document docx2svg could lay out (``coverage_facts``): a check
    that passed on a complete layout is told from one that could not see everything."""
    try:
        return call.document.layout().coverage_facts()
    except ToolError as error:
        return {"error": f"{error.code}: {error.message}"}


# -- S6 find_text ------------------------------------------------------------------------------


def _story_ok(address: str, stories: str) -> bool:
    return stories == "all" or not ("/" in address.split("@")[0] and not address.startswith("p@body/"))


@handler("find_text")
def find_text(call: Any, doc: str, text: str, regex: bool = False, slides: list | None = None,
              range: str | None = None, stories: str = "body", cursor: str | None = None) -> Result:
    if slides is not None:
        raise _decks_only("slides", "range or stories")
    found = call.document.find(text, within=resolve(call, range), regex=regex)
    rows = [{"address": m.id, "kind": "text", "text": short(m.text, 80), "context": m.context()}
            for m in found if _story_ok(m.id, stories)]
    shown, total, next_cursor = page(rows, cursor, call.limits.max_list_items)
    result = Result(summary=f"{total} match(es)", data=shown, next_cursor=next_cursor)
    result.total = total
    return result


# -- S7 replace_text ---------------------------------------------------------------------------


@handler("replace_text", mutates=True)
def replace_text(call: Any, doc: str, find: str, replace: str, expect: str, regex: bool = False,
                 slides: list | None = None, range: str | None = None, stories: str = "body") -> Result:
    if slides is not None:
        raise _decks_only("slides", "range or stories")
    document = call.document
    within = resolve(call, range)
    matches = [m for m in document.find(find, within=within, regex=regex) if _story_ok(m.id, stories)]
    if not matches:
        raise ToolError("not_found", f"{find!r} occurs nowhere" + (f" in {within}" if within else ""))
    if expect == "one" and len(matches) > 1:
        raise ToolError("ambiguous", f"{len(matches)} places match {find!r}; narrow it with range or more "
                        "words", valid_options=[f"{m.id} {m.context()!r}" for m in matches[:50]])
    edits = []
    if stories == "all" or len(matches) == len(document.find(find, within=within, regex=regex)):
        edit = document.replace(find, replace, within=within, regex=regex, count=1 if expect == "one" else None)
        edits.append(edit)
        count = edit.count
    else:
        # Body only, while other stories match too: replace the body's matches one by one,
        # last first, so earlier ranges stay where they are.
        count = 0
        for paragraph in dict.fromkeys(m.id.rsplit("@", 1)[0] for m in matches):
            edits.append(document.replace(find, replace, within=paragraph, regex=regex,
                                          count=1 if expect == "one" else None))
            count += edits[-1].count or 0
    paragraphs = list(dict.fromkeys(m.id.rsplit("@", 1)[0] for m in matches))
    result = outcome(edits, f"Replaced {count} occurrence(s)", data={"count": count, "paragraphs": paragraphs[:50]})
    result.changed = [result.renamed.get(p, p) for p in paragraphs]
    return result


# -- S8 render ---------------------------------------------------------------------------------


@worker_task  # a subprocess worker (a daemonic parent's) runs registered tasks only
def render_pages(data: bytes, pages: list[int], width: int, font_dirs: list[str] | None = None,
                 locale: Any = None) -> tuple[list[bytes], int, dict | None]:
    """In a worker: the pages as PNGs at ``width`` pixels, the page count, and the layout's
    coverage (``coverage_facts``).  ``font_dirs`` is
    the application's font folders, resolved by the caller (the worker sees neither the
    session nor a variable set since it started); ``None`` leaves the worker's default.
    ``locale`` is the session's (``call.locale``, an :class:`ooxml_common.locales.Locale`:
    picklable), handed over the same way; ``None`` is en-US."""
    import docx2svg

    from ..layout import coverage_facts, locale_of

    locale = locale_of(None) if locale is None else locale
    options = docx2svg.ConvertOptions(pages=pages, width=width, font_dirs=font_dirs, locale=locale)
    pngs = docx2svg.convert_docx_to_png(data, options)
    count = (len(docx2svg.convert_docx_to_svg(data, docx2svg.ConvertOptions(font_dirs=font_dirs,
                                                                             locale=locale)))
             if len(pngs) < len(pages) else -1)
    return list(pngs), count, coverage_facts(options.coverage)


def draw(call: Any, pages: list[int], width: int, relaid: list | None = None, fast: bool = False
         ) -> tuple[dict, int | None, dict | None]:
    """``{page: (png, meta, cached, approximate)}`` for the pages the document has among
    ``pages``, drawn ``width`` px wide; the page count when the document was laid out
    (``None`` when every page was held already); and the layout's coverage facts then.
    ``relaid``, a list, is given why the document was laid out whole where a layout of the
    version before was handed over (:func:`docx_agent.layout.relaid_facts`).  ``fast``
    (round 39): pages without an exact image held drawn approximately, for about a third of
    the CPU (:meth:`.render.PageRenderer.rasterise_fast`), ``approximate`` then true.

    ooxml-edit's render pipeline draws them (:func:`ooxml_edit.tools.render.draw`, with the
    document's :data:`.render.RENDERER`): an image is kept by its content in the toolbox's
    store (``call.images``), this version's key of each page on the document, so a repeat
    is answered there.  The rest: one worker lays the document out, told which images the
    store holds -- a page an edit left as it was is keyed, found and not drawn again -- and
    the pages left are rasterised over the workers at once, in chunks.  The layout comes
    back with the pages and is kept on the document by the bytes' digest (the latest only):
    the next call drawing other pages of the same bytes hands it to its worker instead of
    laying the document out again (36 pages: about 1.5 s a call); a call after an edit
    hands it over as the layout of an earlier version, and the worker lays out again only
    the pages the edit reaches (round 19: :class:`.render.PageRenderer`)."""
    document = call.document
    fonts = tuple(font_dirs_of(document))
    locale = locale_of(document)
    options = {"font_dirs": list(fonts), "locale": locale}
    snapshot: list = []

    def kept(data: bytes) -> Any:
        # The document's kept layout (one: the latest; round 33, shared with check's,
        # docx_agent.layout.kept_layout): of these bytes, so that drawing another page of a
        # version -- or a page after a check of it -- does not lay it out again; else of an
        # earlier version with the same font folders and locale, so that the worker lays
        # out again only what the edits since reach.  Round 38: named by its token, in the
        # worker holding it (the affinity below).
        snapshot.append(data)
        state = kept_layout(document, data, options, may_wait=call.may_unlock)
        snapshot.append(state)
        return state

    def on_facts(facts: dict, alone: bool) -> None:
        state = snapshot[1]
        if held_state(state):
            keep_layout(document, snapshot[0], options, state[1])
            note_made(document, facts)
        elif facts.get("laid_out") is not None:
            keep_layout(document, snapshot[0], options, facts["laid_out"])

    key, wait = affinity(document)
    drawn, facts = pipeline.draw(call, RENDERER, pages, width, ids={page: page for page in pages},
                                 fonts=fonts, locale=locale, state=kept, on_facts=on_facts,
                                 affinity=key, affinity_wait=wait, fast=fast)
    if not facts:
        return drawn, None, None
    if relaid is not None and facts[0].get("relaid_cause"):
        # Laid out whole from the layout of the version before (round 20): why, in block ids.
        from ..layout import relaid_facts

        relaid.append(relaid_facts(document, facts[0]["relaid_cause"]))
    return drawn, facts[0]["count"], facts[0]["coverage"]


@handler("render")
def render(call: Any, doc: str, slides: list | None = None, pages: list[int] | None = None,
           width: int | None = None, sheet: dict | None = None, range: Any = None,
           fast: bool | None = None) -> Result:
    fast = bool(fast)
    ranges = ranges_of(resolve(call, range))
    if ranges:
        # Round 35, R18: the pages some blocks fall on -- one image, or a sheet of them.
        pages, scoped = _range_pages(call, ranges, _pages_of(call, slides, pages))
        if sheet is None and len(pages) > call.limits.max_images_per_call:
            call.warn(f"the range is on {len(pages)} pages, more than {call.limits.max_images_per_call} "
                      "images a call: drawn as one sheet")
            sheet = {}
        result = (_render_sheet(call, pages, width, sheet, fast) if sheet is not None
                  else _render_pages(call, pages, width, fast))
        result.data["range"] = scoped
        return result
    if sheet is not None:
        return _render_sheet(call, _pages_of(call, slides, pages), width, sheet, fast)
    return _render_pages(call, _pages_of(call, slides, pages), width, fast)


#: What a render's result says of its fast images (round 39): which facts are approximate.
APPROXIMATE = ["pixels"]


def _approximate(result: Result, pages: list[int]) -> Result:
    """``result`` saying which of its images are approximate (``render`` with ``fast``):
    ``approximate`` names what is (the pixels: text drawn as outlines, the layout and page
    breaks exact), and the summary the pages and the way to the exact image."""
    if pages:
        result.data["approximate"] = list(APPROXIMATE)
        result.data["approximate_pages"] = pages
        result.summary += (f"; fast: page(s) {', '.join(map(str, pages))} approximate (pixels: text as "
                           "outlines; layout and page breaks exact). Render without fast for the exact image")
    return result


def _render_pages(call: Any, pages: list[int] | None, width: int | None, fast: bool = False) -> Result:
    """``render`` of ``pages`` (default the first), each its own image."""
    pages = list(dict.fromkeys(pages or [1]))
    width = width or DEFAULT_WIDTH
    if len(pages) > call.limits.max_images_per_call:
        raise ToolError("limit", f"at most {call.limits.max_images_per_call} pages per call; sheet: {{}} tiles "
                        f"up to {tiles.MAX_TILES} into one image", field="pages")
    call.check_images(len(pages))        # before drawing: a refused call costs nothing
    version = call.entry.version
    fonts = tuple(font_dirs_of(call.document))
    locale = locale_of(call.document)
    relaid: list = []
    drawn, count, coverage = draw(call, pages, width, relaid, fast)
    if len(drawn) < len(pages):
        raise ToolError("not_found", f"the document has {count} page(s)", field="pages",
                        valid_options=list(range(1, count + 1))[:50])
    if coverage is not None:
        call.document._tool_render_coverage = ((version, fonts, locale), coverage)
    images = []
    for number in pages:
        png = drawn[number][0]
        w, h = png_size(png)
        label = f"page {number}" + (" (approximate)" if drawn[number][3] else "")
        images.append(call.image(png, w, h, label=label).describe())
    result = Result(summary=f"Rendered page(s) {', '.join(map(str, pages))} (final view: changes accepted, "
                    "comments hidden)", data={"pages": pages,
                                              "cached": [p for p in pages if drawn[p][2]]})
    render_version, coverage = getattr(call.document, "_tool_render_coverage", (None, None))
    if coverage is not None and render_version == (version, fonts, locale):
        result.data["coverage"] = coverage
    if relaid:
        # The document was laid out whole although a layout of the version before was kept:
        # what caused it (a styles change, an edited section break, a field one block opens
        # and another closes...), to explain a slow refresh.
        result.data["relaid"] = relaid[0]
    return _approximate(result, [p for p in pages if drawn[p][3]])


def _range_pages(call: Any, ranges: list[str], pages: list[int] | None) -> tuple[list[int], list[dict]]:
    """The pages ``ranges`` fall on (with ``pages``, those of them), and each range as
    reported (``{"range", "blocks", "pages"}``), from the document's layout: laid out with
    its lock let go where the call may, else on its own snapshot (round 34's rule)."""
    if len(ranges) > MAX_RANGES:
        raise ToolError("limit", f"{len(ranges)} ranges: at most {MAX_RANGES} per render; a span p:A..t:B "
                        "covers many blocks", field="range")
    document, entry = call.document, call.entry
    scope = Scope(document, ranges)
    if pages is None:
        call.read_scope = read_scope(scope)     # what the result is of (round 47)
    version = entry.version
    layout = (lay_out_unlocked(document, call.unlocked, lambda: entry.version == version) if call.may_unlock
              else document.layout())
    scoped, on = range_pages(scope, layout)
    if not on:
        raise ToolError("not_found", "the range's blocks are drawn on no page (hidden, or past where the "
                        "layout stopped)", field="range",
                        valid_options=list(range(1, layout.page_count + 1))[:50])
    if pages is not None:
        wanted = [p for p in on if p in pages]
        if not wanted:
            raise ToolError("invalid_arguments", f"pages {pages} hold none of the range's blocks: they are on "
                            f"pages {on}; leave pages out to draw them", field="pages", valid_options=on[:50])
        on = wanted
    return on, scoped


def _render_sheet(call: Any, pages: list[int] | None, width: int | None, sheet: dict,
                  fast: bool = False) -> Result:
    """``render`` with ``sheet``: the pages tiled into one image, in order, each labelled
    ``page 3`` (:mod:`ooxml_edit.tools.tiles`); ``pages`` left out is the first
    ``tiles.MAX_TILES``, or all there are.  ``sheet.width`` is each tile's, ``width``
    standing in for it.  Tiles too wide or too tall for the image's edge are drawn narrower
    -- the widest that fit (:func:`tiles.fit`) -- and ``sheet.shrunk`` says from what width
    and why: before anything is drawn, at the shortest page the sections declare, for the
    pages named or the page count a render of this version found; again once the pages are
    drawn and their heights known (drawn again narrower then: the layout is kept, so that
    is rasterising alone).  Only tiles too tall even at the narrowest are refused, naming
    how many pages fit.  The tiles are the images ``render`` keeps (:func:`draw`: those held
    are not drawn again); ``cached_tiles`` lists them.  The sheet is cached whole, one
    entry, so a sheet of 30 does not push the single renders out."""
    limits = call.limits
    named = pages is not None
    pages = list(dict.fromkeys(pages)) if named else list(range(1, tiles.MAX_TILES + 1))
    fonts = tuple(font_dirs_of(call.document))
    locale = locale_of(call.document)
    version = call.entry.version
    known = pages
    if not named:
        render_version, coverage = getattr(call.document, "_tool_render_coverage", (None, None))
        count = coverage.get("pages") if coverage is not None and render_version == (
            version, fonts, locale) else None
        known = pages[:count] if isinstance(count, int) and count > 0 else None
    tiling = tiles.tiling_of(sheet, len(known or pages), name="sheet", width=width, what="pages",
                             max_edge=limits.max_image_edge)
    if known is not None:
        shortest = _shortest_ratio(call.document)
        if shortest is not None:
            tiling = tiles.fit(tiling, [shortest] * len(known), name="sheet", what="pages",
                               max_edge=limits.max_image_edge)
    call.check_images(1)                 # before drawing: a refused call costs nothing
    entry = call.entry
    key = ("sheet", version, tuple(pages) if named else None, tiling.columns, tiling.width,
           fonts, locale) + (("fast",) if fast else ())
    made = entry.render_cache.get(key)
    cached = made is not None
    tiles_cached = None
    if made is None:
        drawn, count, coverage = draw(call, pages, tiling.width, fast=fast)
        if len(drawn) < len(pages) and named:
            raise ToolError("not_found", f"the document has {count} page(s)", field="pages",
                            valid_options=list(range(1, count + 1))[:50])
        pages = [p for p in pages if p in drawn]
        drawn_as = tiling
        tiling = tiles.tiling_of(sheet, len(pages), name="sheet", width=width, what="pages",
                                 max_edge=limits.max_image_edge)
        sizes = [png_size(drawn[p][0]) for p in pages]
        tiling = tiles.fit(tiling, [h / w for w, h in sizes], name="sheet", what="pages",
                           max_edge=limits.max_image_edge)
        if named and tiling.width > drawn_as.width:
            tiling = drawn_as            # drawn a pixel narrower than needed: kept
        elif tiling.width != drawn_as.width:     # taller pages, or fewer than the first 30
            drawn, _, again = draw(call, pages, tiling.width, fast=fast)
            coverage = coverage if again is None else again
        tiles_cached = [p for p in pages if drawn[p][2]]
        rough = [p for p in pages if drawn[p][3]]
        # The snapshot is taken (the tiles); the worker draws the sheet from it.
        png, w, h = pipeline.draw_tiles(call, RENDERER, [(drawn[p][0], f"page {p}") for p in pages],
                                        tiling, name="sheet", what="pages", fonts=fonts)
        if coverage is not None:
            call.document._tool_render_coverage = ((version, fonts, locale), coverage)
        made = (png, w, h, pages, tiling, rough)
        entry.render_cache.put(key, made)
    png, w, h, pages, tiling, rough = made
    call.image(png, w, h, label=f"sheet of pages {pages[0]}-{pages[-1]}" if len(pages) > 1
               else f"sheet of page {pages[0]}")
    data = {"pages": pages, "sheet": {"columns": tiling.columns, "width": tiling.width},
            "cached": cached}
    if tiling.shrunk:
        data["sheet"]["shrunk"] = tiling.shrunk
    if tiles_cached:
        data["cached_tiles"] = tiles_cached
    result = Result(summary=f"Rendered {len(pages)} page(s) as one sheet, {tiling.columns} column(s) of "
                    f"{tiling.width} px{tiles.narrowed_note(tiling)}, labelled with page numbers (final "
                    "view: changes accepted, comments hidden)", data=data)
    render_version, coverage = getattr(call.document, "_tool_render_coverage", (None, None))
    if coverage is not None and render_version == (version, fonts, locale):
        result.data["coverage"] = coverage
    return _approximate(result, rough)


def _shortest_ratio(document: Document) -> float | None:
    """The height over the width of the shortest page the document's sections declare
    (``w:pgSz``), or ``None`` when a section declares no size: a sheet too tall even with
    such pages needs narrower tiles whatever the pages are."""
    ratios = []
    for section in document.sections():
        size = section.page_size
        if not size or not size[0] or not size[1]:
            return None
        ratios.append(size[1] / size[0])
    return min(ratios) if ratios else None


# -- S9 check ----------------------------------------------------------------------------------


@handler("check")
def check(call: Any, doc: str, slides: list | None = None, pages: list[int] | None = None,
          include: list[str] | None = None, boxes: bool | None = None,
          spacing: float | None = None, contrast: float | None = None,
          min_size: float | None = None, styles: list[str] | None = None,
          range: Any = None, copies: list | None = None) -> Result:
    """The document's facts; ``range`` (round 35, R18) scopes them to some blocks
    (:mod:`.scoped`), ``copies`` (R19) compares copies with their reference.

    Round 47 (R82): ``data.timing`` says where the time went, as an edit's does
    (:class:`~._base.Timing`): the parts that ran -- ``resolve`` (the ranges), ``validate``,
    ``copies``, ``layout`` (laying the document out), ``facts`` (reading the layout's facts
    and comparing with the last check), ``fields`` -- and, apart, the waits: ``worker_wait``
    (for the worker holding the document's layout: its queue, a layout of the document in
    progress there, the bytes there and the facts back) and ``document_wait`` (taking the
    document back after laying out with it let go).  The wait for the document before the
    call ran is the result's duration less these."""
    timing = Timing()
    pages = _pages_of(call, slides, pages)
    if boxes is not None:
        raise _decks_only("boxes", "include")
    if spacing is not None:                   # a slide's gaps and padding (round 12, Q1)
        raise _decks_only("spacing", "include")
    if contrast is not None:                  # a slide's text against its fill (round 15, T2)
        raise _decks_only("contrast", "include")
    wanted = list(include or ["validate", "reflow", "fields"])
    deck_only = [k for k in wanted if k in ("fit", "collisions", "facts", "design")]
    if deck_only:
        raise ToolError("invalid_arguments", f"{', '.join(deck_only)} are deck facts; a document reports "
                        "validate, reflow, fields and layout", field="include",
                        valid_options=["validate", "reflow", "fields", "layout"])
    document, entry = call.document, call.entry
    if min_size is None:
        # The application's minimums (round 17, P6b: docx_format(min_text_size=...)).
        min_size = getattr(document, "min_text_size", None)
    if (min_size is not None or styles is not None) and "layout" not in wanted:
        wanted.append("layout")               # what they measure (round 16, F8)
    # Round 35, R18: the facts about some blocks, on the pages they fall on -- read under the
    # lock, before the layout lets it go.
    with timing.part("resolve") if range is not None else _untimed():
        ranges = ranges_of(resolve(call, range))
        if len(ranges) > MAX_RANGES:
            raise ToolError("limit", f"{len(ranges)} ranges: at most {MAX_RANGES} per check; a span p:A..t:B "
                            "covers many blocks", field="range")
        scope = Scope(document, ranges) if ranges else None
    if scope is not None and pages is None:
        # What the result is of, for whether a change meanwhile touched it (round 47).
        call.read_scope = read_scope(scope)
    if scope is not None and "layout" not in wanted:
        wanted.append("layout")               # the facts about the blocks are the layout's
    # Round 40: each range's keep tried (.keep_trials).
    keeps = scope is not None and (KEEP_TRIALS or "keeps" in wanted)
    report: dict[str, Any] = {}
    if "validate" in wanted:
        with timing.part("validate"):
            _validate(document, entry, scope, report)
    if copies:
        # Round 35, R19: under the lock, on the version validate read.
        with timing.part("copies"):
            report["copies"] = compare_copies(call, copies)
    if "reflow" in wanted or "layout" in wanted:
        started, inner = time.perf_counter(), sum(timing.parts.values())
        reflow, layout = _layout_facts(call, pages, seen="reflow" in wanted, scope=scope,
                                       layout_wanted="layout" in wanted, min_size=min_size, styles=styles,
                                       keeps=keeps, timing=timing)
        # Round 38: made where the layout is held -- the ranges' pages and the layout facts
        # came back with the reflow's (a dict); else the layout itself, read here.
        made = layout if isinstance(layout, dict) else None
        scoped_pages = pages
        if scope is not None:
            if made is not None:
                report["range"], on = made["range"], made["on"]
                scoped_pages = on if pages is None else [p for p in on if p in pages]
                if pages is not None and not scoped_pages:
                    call.warn(f"pages {pages} hold none of the range's blocks (on pages {on}): no layout facts")
                since = reflow.get("since_last_check")
                if since is not None:
                    since["changed"] = [p for p in since["changed"] if p in scoped_pages]
            elif layout is not None:
                report["range"], on = range_pages(scope, layout, extent=True)
                if keeps:
                    # Round 40: a keep is tried only where the layout is held.
                    left_out(layout, scope, report["range"], KEEPS_NOT_HERE)
                scoped_pages = on if pages is None else [p for p in on if p in pages]
                if pages is not None and not scoped_pages:
                    call.warn(f"pages {pages} hold none of the range's blocks (on pages {on}): no layout facts")
                since = reflow.get("since_last_check")
                if since is not None:
                    since["changed"] = [p for p in since["changed"] if p in scoped_pages]
            else:
                report["range"] = [{"range": r.address, "blocks": len(r.blocks), "pages": None}
                                   for r in scope.ranges]
        if "reflow" in wanted:
            report["reflow"] = reflow
        if "layout" in wanted:
            # Measured on the layout reflow reports and render draws: never a second one.
            # Where the layout failed, the facts read from the document alone (round 18, Q4;
            # round 29, W12).
            report["layout"] = (made["layout"] if made is not None else
                                layout_facts(layout, pages=scoped_pages, min_size=min_size, styles=styles,
                                             scope=scope)
                                if layout is not None else
                                {"error": reflow["error"],
                                 "empty_paragraphs": [e for e in empty_paragraphs(document)
                                                      if scope is None or scope.about(e)][:CAP],
                                 "numbering_skips_level": [e for e in numbering_skips_level(document)
                                                           if scope is None or scope.about(e)][:CAP],
                                 "punctuation_spacing": [e for e in punctuation_spacing(document)
                                                         if scope is None or scope.about(e)][:CAP]})
        # The rest of the layout's facts here, beside what the worker timed (layout, facts, waits).
        inner = sum(timing.parts.values()) - inner
        timing.add("facts", max(0.0, time.perf_counter() - started - inner))
    if "fields" in wanted:
        with timing.part("fields"):
            report["fields"] = _field_facts(document)
    if "app" in wanted:
        report["app_findings"] = None
    summary = "Checked " + ", ".join(wanted)
    if scope is not None:
        summary += f" in {len(scope.ranges)} range(s)"
    if "copies" in report:
        found = report["copies"]
        summary += (f"; copies: {found['follow']} of {found['compared']} follow their reference, "
                    f"{found['with_example_text']} with its example text")
    report["timing"] = timing.parts
    return Result(summary=summary, data=report)


def _validate(document: Document, entry: Any, scope: Any, report: dict[str, Any]) -> None:
    """``check``'s ``validate``, into ``report``: the problems (in ``scope``'s blocks, with
    those elsewhere counted) and the delta against those the document had when opened."""
    found = list(document.validate())          # once: the delta reads the same problems
    if scope is None:
        report["validate"] = {**validate_delta(entry, found), "problems": [str(p) for p in found][:50]}
        return
    inside, elsewhere = problems_in_scope(document, scope, found)
    delta = validate_delta(entry, found)
    named = {str(p) for p in inside}
    report["validate"] = {"new": [p for p in delta["new"] if p in named], "fixed": delta["fixed"],
                          "baseline": delta["baseline"], "problems": [str(p) for p in inside][:50],
                          "elsewhere": elsewhere}


@contextmanager
def _untimed() -> Iterator[None]:
    yield


def _timed_unlocked(call: Any, timing: Timing) -> Any:
    """``call.unlocked``, the time it takes to take the document back after counted as
    ``document_wait`` (round 47, R82) when it let the document go."""
    @contextmanager
    def unlocked() -> Iterator[bool]:
        with call.unlocked() as let_go:
            yield let_go
            back = time.perf_counter()
        if let_go:
            timing.add("document_wait", time.perf_counter() - back)
    return unlocked


#: Whether a check with ranges tries each range's keep by default (round 40, .keep_trials).
KEEP_TRIALS = True
#: What a range a keep is not tried for says, where the check read the layout in this process (round 40).
KEEPS_NOT_HERE = ("not laid out: this check read the layout here, not where it is held (a check in a batch "
                  "that holds the document); check this range alone, after the batch, for what its keep "
                  "does to the pages")


#: What a check laid out between a batch's changes says (round 33, R16c): the way out.
BATCH_CHECK = ("check laid the document out between the batch's changes, holding it: other calls waited. "
               "Put check after the batch's last change (or call it alone): there it lets the document go "
               "while it lays out")


def _layout_facts(call: Any, pages: list[int] | None, seen: bool = True, *, scope: Any = None,
                  layout_wanted: bool = False, min_size: Any = None,
                  styles: list[str] | None = None, keeps: bool = False,
                  timing: Timing | None = None) -> tuple[dict[str, Any], Any]:
    """``reflow``'s facts, and the layout they are of (``None`` where it failed); ``seen``:
    the layout a later check's ``since_last_check`` is measured from (not for ``layout``
    alone, which reports no reflow).

    Round 38: a lone check's facts are made in the worker holding the document's layout
    (:func:`docx_agent.layout.check_facts`) -- with the ranges' pages (``scope``) and the
    layout facts (``layout_wanted``, ``min_size``, ``styles``) -- and the second value is
    then those, a dict ``{"range", "on", "layout"}``: the layout never comes here.

    ``timing`` (round 47, R82) is told the ``layout`` -- the worker's, else laying out here
    -- and the waits (:func:`check_facts`, :func:`_timed_unlocked`)."""
    timing = timing if timing is not None else Timing()
    unlocked = _timed_unlocked(call, timing)
    document = call.document
    entry = call.entry
    version = entry.version
    made = None
    try:
        if call.may_unlock:
            # Laid out with the lock let go (a check outside a batch): of this version, the
            # one validate and fields read and the result reports.
            request = {"pages": list(pages) if pages is not None else None,
                       "ranges": scope.spec() if scope is not None else None,
                       "layout": layout_wanted, "min_size": min_size, "styles": styles,
                       "keeps": keep_trials.asked() if keeps and scope is not None else None}
            made = check_facts(document, unlocked, request, timed=timing.add)
            if made is None:
                started, waited = time.perf_counter(), timing.parts.get("document_wait", 0.0)
                layout = lay_out_unlocked(document, unlocked, lambda: entry.version == version)
                timing.add("layout", time.perf_counter() - started
                           - (timing.parts.get("document_wait", 0.0) - waited))
            else:
                layout = None
        else:
            laid_out = call.in_batch and cached_layout(document, fetch=False) is None
            with timing.part("layout"):
                layout = document.layout()
            if laid_out:
                # Between a batch's changes (round 33, R16c): after the last one, a check
                # runs once the batch is done and lets the document go while it lays out.
                call.warn(BATCH_CHECK)
    except ToolError as error:
        return {"error": f"{error.code}: {error.message}"}, None
    if made is not None:
        result, now = made
        facts: dict[str, Any] = result["reflow"]
        warnings = result["warnings"]
    else:
        assert layout is not None
        facts = {"pages": layout.page_count, "complete": layout.pages_known is not None}
        if layout.stopped is not None:
            facts["stopped"] = {"page": layout.stopped.page, "reason": layout.stopped.reason,
                                "at": layout.stopped.at}
        # The whole of it: blocks laid out, header and footer stops, faces substituted or
        # missing -- so "the check passed" is told from "the check could not see everything".
        facts["coverage"] = layout.coverage_facts()
        key = layout.key if layout.key is not None else layout_key(
            document.to_bytes(), layout_options(document, {}))
        now = summary_for(document, key, layout)
        warnings = [short(str(w), 160) for w in layout.warnings[:10]] if layout.warnings else None
    before = getattr(document, "_tool_seen_layout", None)
    if before is not None and before is not now:
        # Compared by summaries (round 38), ids followed through the document's aliases now.
        reflow = reflow_json(now.compare(before, canon_of(document)))
        if pages:
            reflow["changed"] = [p for p in reflow["changed"] if p in pages]
        facts["since_last_check"] = reflow
    if seen:
        document._tool_seen_layout = now
    if warnings:
        facts["layout_warnings"] = warnings
    if made is not None:
        result.pop("reflow")
        return facts, result
    return facts, layout


def _field_facts(document: Document) -> dict[str, Any]:
    from ..edit.fields import scan

    # Each part's fields read once (round 33: a Field's properties read its part again each,
    # which made a contents of 500 entries cost seconds a check, under the lock).
    fields = [(f"fld:{document._story_of(part)}/{found.number}", found)
              for part in document._parts() for found in scan(part, document.package.tree(part))]
    counts: dict[str, int] = {}
    for _, found in fields:
        counts[found.keyword] = counts.get(found.keyword, 0) + 1
    facts: dict[str, Any] = {"counts": counts}
    empty = document._empty_tocs(found=fields)
    if empty:
        facts["empty_toc"] = empty
    missing = []
    # The bookmarks read once, found as Document.bookmark finds one (a "bm:" prefix, any case).
    markers = document._bookmark_markers()
    folded = {name.casefold() for name in markers}
    for identifier, found in fields:
        if found.keyword in ("REF", "PAGEREF", "NOTEREF"):
            words = found.instruction.split()
            name = words[1] if len(words) > 1 else ""
            bare = name[3:] if name.startswith("bm:") else name
            if bare not in markers and bare.casefold() not in folded:
                missing.append(f"{identifier}: {found.keyword} to a missing bookmark {name!r}")
    if missing:
        facts["missing_bookmarks"] = missing
    facts["note"] = "results are caches: after moving headings or text, word_fields update"
    return facts


# -- S10-S12 charts and SmartArt ---------------------------------------------------------------


def chart_json(chart: Any) -> dict[str, Any]:
    """What an ``edit_chart`` result holds (ooxml-edit's ``chart_json``), the chart by its id."""
    from ooxml_edit.charts.actions import chart_json as shared_json

    return shared_json(chart, "id")


@handler("edit_chart", mutates=True)
def edit_chart(call: Any, doc: str, target: str, action: str, chart_type: str | None = None,
               categories: list[str] | None = None, data: list[dict] | None = None, box: dict | None = None,
               width: float | None = None, number_format: str | None = None, series: str | None = None,
               category: str | None = None, values: list[float] | None = None, value: float | None = None,
               text: str | None = None, axis: str | None = None, position: str | None = None,
               ref: str | None = None, scale: dict | None = None, gridlines: str | None = None,
               style: dict | None = None, part: str | None = None, labels: list | None = None,
               cells: bool | None = None, secondary: bool | None = None, data_labels: bool | None = None,
               colors: list | None = None) -> Result:
    from ooxml_edit.charts import ChartDataError
    from ooxml_edit.charts.actions import STYLE_ACTIONS

    try:
        if action == "add":
            return _add_chart(call, resolve(call, target), chart_type, categories, data, box, width, values,
                              number_format, text, position, ref, axis, scale, gridlines, data_labels, colors,
                              style, part)
        if action in STYLE_ACTIONS:
            return _style_chart(call, target, action, series=series, category=category, axis=axis,
                                scale=scale, gridlines=gridlines, style=style, part=part, labels=labels,
                                cells=cells, position=position, chart_type=chart_type, secondary=secondary)
        return _edit_chart(call, target, action, number_format=number_format, series=series,
                           category=category, values=values, value=value, text=text, axis=axis,
                           position=position)
    except (ChartDataError, IndexError) as error:
        raise ToolError("invalid_arguments", str(error).strip("'\"")) from None


def _style_chart(call: Any, target: str, action: str, *, series: str | None, category: str | None,
                 **arguments: Any) -> Result:
    """The formatting actions (ooxml-edit's ``apply_style_action``)."""
    from ooxml_edit.charts.actions import (apply_style_action, chart_style_json, resolve_category,
                                           resolve_series)

    chart = call.document.chart(target)
    chosen = None if series is None else resolve_series(chart, series, action).index
    point = None if category is None or chosen is None else resolve_category(chart, category, action)
    apply_style_action(chart, action, series=chosen, category=point, **arguments)
    chart = call.document.chart(target)
    data = chart_json(chart)
    data.update(chart_style_json(chart))
    return Result(summary=f"{action} on {target}", changed=[target], data=data)


def _add_chart(call: Any, target: str, chart_type: str | None, categories: list[str] | None,
               data: list[dict] | None, box: dict | None, width: float | None, values: list[float] | None,
               number_format: str | None, text: str | None, position: str | None, ref: str | None,
               axis: str | None = None, scale: dict | None = None, gridlines: str | None = None,
               data_labels: bool | None = None, colors: list | None = None, style: dict | None = None,
               part: str | None = None) -> Result:
    """``action: "add"``: a new inline chart in a paragraph of its own after ``target``, with the look its
    arguments give (ooxml-edit's ``add_plan`` / ``finish_add``: the same chart as the separate calls)."""
    if box is not None:
        raise ToolError("invalid_arguments", "box is for decks; in a document give width", field="box")
    from ooxml_edit.charts.actions import add_plan, finish_add

    plan = add_plan(chart_type, categories, data, values, position=position, axis=axis, scale=scale,
                    data_labels=data_labels, colors=colors, style=style, part=part)
    with call.document._edit():  # the chart and its look: one undo step
        edit = call.document.insert_chart(target, chart_type, plan["labels"], [dict(entry) for entry in data],
                                          width=width, title=text, legend=plan["legend"],
                                          number_format=number_format,
                                          category_order=plan["category_order"],
                                          gridlines=gridlines != "none")
        finish_add(call.document.chart(edit.id), chart_type, plan, axis=axis, gridlines=gridlines)
    remember(call, ref, edit.id)
    chart = call.document.chart(edit.id)
    data = chart_json(chart)
    paragraphs = [identifier for identifier in edit.created if identifier.startswith("p:")]
    if paragraphs:
        data["paragraph"] = paragraphs[0]  # where a caption goes after
    return outcome([edit], f"Inserted a {chart_type} chart {edit.id} in {data.get('paragraph', target)}",
                   data=data)


def _edit_chart(call: Any, target: str, action: str, **arguments: Any) -> Result:
    """``read`` and the data actions (ooxml-edit's ``apply_data_action``)."""
    from ooxml_edit.charts.actions import apply_data_action, chart_read_json

    chart = call.document.chart(target)
    if action == "read":
        return Result(summary=f"Read {target}", data=chart_read_json(chart, "id"))
    apply_data_action(chart, action, **arguments)
    chart = call.document.chart(target)
    return Result(summary=f"{action} on {target}", changed=[target], data=chart_json(chart))


@handler("edit_smartart", mutates=True)
def edit_smartart(call: Any, doc: str, target: str, action: str, node: int | None = None,
                  text: str | None = None) -> Result:
    from ooxml_edit.charts.actions import apply_smartart_action, smartart_json

    apply_smartart_action(call.document.diagram(target), action, node=node, text=text)
    return Result(summary=f"{action} on {target}", changed=[target],
                  data=smartart_json(call.document.diagram(target)))


# -- S16 describe ------------------------------------------------------------------------------


def _describe(call: Any, doc: str) -> Result:
    from .read import word_describe

    return word_describe(call, doc)


describe = handler("describe")(_describe)


# -- S13 set_properties ------------------------------------------------------------------------


@handler("set_properties", mutates=True, track=False)
def set_properties(call: Any, doc: str, title: str | None = None, author: str | None = None,
                   language: str | None = None, subject: str | None = None) -> Result:
    values = {k: v for k, v in (("title", title), ("author", author), ("language", language),
                                ("subject", subject)) if v is not None}
    if not values:
        raise ToolError("invalid_arguments", "give a property to set",
                        valid_options=["title", "author", "language", "subject"])
    edit = call.document.set_properties(**values)
    properties = call.document.properties
    return outcome([edit], f"Set {', '.join(values)}",
                   data={k: properties.get(k) for k in ("title", "author", "language", "subject")})


# -- style_by_value (round 36, R31) -------------------------------------------------------------


def _style_by_value(call: Any, doc: str, scale: Any, preview: bool | None = None) -> Result:
    from .scales import style_by_value as run

    return run(call, doc, scale, preview)


style_by_value = handler("style_by_value", mutates=True)(_style_by_value)


#: ``read_batch`` (ooxml-edit 0.78.0; round 32, R15): many reads in one call, each with its
#: own budget; a reader's tool, in the core.  ``style_by_value`` (round 36, R31) is listed
#: in ``TOOLS`` after every Word tool: the strict plan tries tools in order, and the ones
#: before it keep their places.
TOOLS = [new_document, save_document, describe, find_text, replace_text, render, check, edit_chart,
         edit_smartart, set_properties, *shared.SESSION_TOOLS, shared.READ_BATCH]
