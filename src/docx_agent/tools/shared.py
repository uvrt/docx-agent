"""The Word handlers of the shared tools (``ooxml_edit.tools.shared``): one definition each
for every format, dispatched by the document's kind.

S1 ``open_document``, S4 ``close_document``, S5 ``undo`` and ``batch`` are
the session's own (``shared.SESSION_TOOLS``); this module adds docx's ``new_document``,
``save_document``, ``find_text``, ``replace_text``, ``render``, ``check``, ``edit_chart`` (its
``read`` action too), ``edit_smartart``, ``set_properties`` and ``describe`` (its handler,
``word_describe``, is in :mod:`.read`)."""

from __future__ import annotations

import hashlib
import re
import struct
from typing import Any

from ooxml_edit.tools import Result, ToolError, shared, tiles, worker_task

from ..edit.document import Document
from ..edit.properties import refreshed_app
from ..layout import font_dirs_of, lay_out_unlocked, locale_of
from ..layout_facts import layout_facts
from ._base import need, outcome, page, reflow_json, remember, resolve, short, validate_delta, wrap

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
    result = Result(summary=f"Saved {name} for the application", data=call.output(name, format, data))
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
    coverage (``coverage_facts``; ``None`` from a docx2svg without it).  ``font_dirs`` is
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
    coverage = getattr(options, "coverage", None)
    return list(pngs), count, coverage_facts(coverage) if coverage is not None else None


@worker_task
def rasterise_tiles(svg: str, font_dirs: list[str] | None = None) -> bytes:
    """In a worker: the PNG of a contact sheet's SVG (``ooxml_edit.tools.tiles.tiled_svg``:
    the pages' PNGs at their own size, their labels), drawn once at the SVG's own size."""
    from docx2svg.png import svg_to_png

    return svg_to_png(svg, font_dirs=font_dirs)


def png_size(data: bytes) -> tuple[int, int]:
    return struct.unpack(">II", data[16:24])


# The render tool's tasks.  A document is laid out whole, so one worker lays it out once and
# writes the pages' SVGs (lay_out_pages); the pages whose image is not held already are
# then rasterised over all the workers at once (rasterise_pages), in chunks.  Both work
# from the bytes they are given (docx2svg opens them itself); their PNGs are render_pages'
# byte for byte: the same docx2svg passes and the same rasteriser arguments.


@worker_task
def warm_up() -> None:
    """A new worker's first task, before any call's: the renderer's imports, the
    rasteriser's own load of the system's fonts (once per process), and the stamp of the
    renderers' code -- what the first render would pay for."""
    import docx2svg  # noqa: F401
    from docx2svg.png import svg_to_png

    _code_stamp()
    svg_to_png('<svg xmlns="http://www.w3.org/2000/svg" width="64" height="16">'
               '<text x="0" y="12" font-family="sans-serif" font-size="10">Aa</text></svg>')


#: The ``data-docx-*`` attributes of a page's SVG: where its elements come from (paths,
#: ids), which the rasteriser never reads -- left out of the key, so the ids an edit gives
#: every paragraph (``w14:paraId``) do not make every page new.
_DOCX_DATA = re.compile(r' data-docx-[a-z-]+="[^"]*"')


def _code_stamp() -> list:
    from ooxml_edit.tools.rendercache import code_stamp

    return code_stamp("docx2svg", "ooxml_common", "resvg_py")


@worker_task
def lay_out_pages(data: bytes, pages: list[int], width: int, font_dirs: list[str] | None = None,
                  locale: Any = None, known: list[str] = (), laid_out: bytes | None = None) -> dict:
    """In a worker: the document laid out once, and each of ``pages`` that it has as
    ``(page, key, svg)`` -- ``key`` the image's content key (the SVG, and everything else the
    pixels depend on: the font files the rasteriser is handed, with their times and sizes,
    the embedded faces by digest, the rasteriser's options, the renderers' code), ``svg``
    ``None`` where the key is in ``known`` (the caller holds that image).  Also ``count``
    (the pages there are), ``coverage`` (as :func:`render_pages`) and ``raster``: what
    :func:`rasterise_pages` hands the rasteriser, the embedded faces as bytes.

    ``laid_out`` is a layout of these very bytes, font folders and locale that an earlier
    call returned as ``"laid_out"`` (pickled: the layout and the faces it was measured
    with, :func:`docx2svg.lay_out_docx`); given, the document is not laid out again --
    the same pages, byte for byte, for the cost of unpickling it.  Without it, the layout
    is made here and returned so, for the caller to hand the next call."""
    import os
    import pickle
    import tempfile

    import docx2svg
    from docx2svg.fonts import rasteriser_files
    from ooxml_common.fonts.rasterdb import file_stamps, resvg_system_dirs
    from ooxml_edit.tools import content_key

    from ..layout import coverage_facts, locale_of

    locale = locale_of(None) if locale is None else locale
    options = docx2svg.ConvertOptions(pages=pages, width=width, font_dirs=font_dirs, locale=locale)
    fresh = None
    if laid_out is not None:
        layout, fonts = pickle.loads(laid_out)
    else:
        layout, _, fonts = docx2svg.lay_out_docx(data, options)
        fresh = pickle.dumps((layout, fonts), protocol=pickle.HIGHEST_PROTOCOL)
    count = len(layout.pages)
    coverage = getattr(layout, "coverage", None)
    svgs = docx2svg.page_svgs(layout, data, options, device_size=True, fonts=fonts)
    numbers = [page for page in pages if 1 <= page <= count]
    # The rasteriser's arguments, as docx2svg's convert_docx_to_png makes them.
    faces = docx2svg.drawn_faces(layout)
    files: list = []
    embedded: list = []
    sans = None
    missing = fonts is None
    with tempfile.TemporaryDirectory(prefix="docx2svg-embedded-") as directory:
        if fonts is not None:
            found, lacking, _ = rasteriser_files(fonts, faces, directory)
            missing = bool(lacking)
            for path in found:
                if os.path.dirname(path) == directory:
                    with open(path, "rb") as handle:
                        embedded.append((os.path.basename(path), handle.read()))
                    files.append(["embedded", os.path.basename(path)])
                else:
                    files.append(path)
            for family in ("Arial", "Helvetica", "Calibri"):
                face = fonts.drawing_face(family, record=False)
                if face is not None and not face.source.startswith("embedded:"):
                    sans = family
                    if face.source not in files:
                        files.append(face.source)
                    break
    skip = not missing
    dirs = docx2svg.user_font_dirs(options)
    system = None
    if not skip:
        found_dirs = resvg_system_dirs()
        system = file_stamps((), found_dirs) if found_dirs is not None else "fontconfig"
    plain = [path for path in files if isinstance(path, str)]
    fingerprint = content_key("docx-agent raster", _code_stamp(), skip, sans, file_stamps(plain, dirs),
                              files, [(name, content_key(blob)) for name, blob in embedded], system)
    known = set(known)
    out = []
    for page, svg in zip(numbers, svgs):
        key = content_key("docx-agent png", _DOCX_DATA.sub("", svg), fingerprint)
        out.append((page, key, None if key in known else svg))
    return {"pages": out, "count": count,
            "coverage": coverage_facts(coverage) if coverage is not None else None,
            "raster": {"files": files, "embedded": embedded, "dirs": dirs, "skip": skip,
                       "sans": sans},
            "laid_out": fresh}


@worker_task
def rasterise_pages(svgs: list[str], raster: dict) -> list[bytes]:
    """In a worker: the PNG of each page SVG of :func:`lay_out_pages`, with its ``raster``
    arguments (the embedded faces written to a temporary folder for as long as it takes)."""
    import os
    import tempfile

    from docx2svg.png import svg_to_png

    with tempfile.TemporaryDirectory(prefix="docx2svg-embedded-") as directory:
        written = {}
        for name, blob in raster["embedded"]:
            path = written[name] = os.path.join(directory, name)
            with open(path, "wb") as handle:
                handle.write(blob)
        files = [written[item[1]] if isinstance(item, list) else item for item in raster["files"]]
        return [svg_to_png(svg, backend="auto", font_dirs=raster["dirs"], font_files=files,
                           skip_system_fonts=raster["skip"], sans_serif_family=raster["sans"])
                for svg in svgs]


def draw(call: Any, pages: list[int], width: int) -> tuple[dict, int | None, dict | None]:
    """``{page: (png, cached)}`` for the pages the document has among ``pages``, drawn
    ``width`` px wide; the page count when the document was laid out (``None`` when every
    page was held already); and the layout's coverage facts then.

    An image is kept by its content in the toolbox's store (``call.images``); this version's
    key of each page is remembered on the document, so a repeat is answered here.  The rest:
    one worker lays the document out (:func:`lay_out_pages`), told which images the store
    holds -- a page an edit left as it was is keyed, found and not drawn again -- and the pages left are rasterised over the workers at once, in chunks
    (:func:`rasterise_pages`).  The bytes are the version's snapshot, taken under the lock;
    the workers work with the lock let go.  The layout comes back with the pages and is
    kept on the document by the bytes' digest (the latest only): the next call drawing
    other pages of the same bytes hands it to its worker instead of laying the document
    out again (36 pages: about 1.5 s a call)."""
    entry, version, limits, store = call.entry, call.entry.version, call.limits, call.images
    fonts = tuple(font_dirs_of(call.document))
    locale = locale_of(call.document)
    drawn: dict = {}
    todo = []
    for number in dict.fromkeys(pages):
        key = entry.render_keys.get((version, number, width, fonts, locale))
        found = store.get(key) if key is not None else None
        if found is not None:
            drawn[number] = (found[0], True)
        else:
            todo.append(number)
    if not todo:
        return drawn, None, None
    data = entry.snapshot()
    held = store.entries()           # what the worker is told is drawn already
    # The layout of these bytes an earlier call made, kept on the document (one: the
    # latest), so that drawing another page of a version does not lay it out again.
    layout_key = (hashlib.sha256(data).digest(), fonts, locale)
    kept = getattr(call.document, "_tool_render_layout", None)
    laid_out = kept[1] if kept is not None and kept[0] == layout_key else None
    with call.unlocked():
        laid = call.run(lay_out_pages, data, todo, width, list(fonts), locale, list(held), laid_out,
                        timeout=limits.render_timeout_for(len(todo)))
        fresh = [(page, key, svg) for page, key, svg in laid["pages"] if svg is not None]
        parallel = min(call.parallelism, len(fresh)) or 1
        chunks = [fresh[start::parallel] for start in range(parallel)]
        chunks = [chunk for chunk in chunks if chunk]
        results = call.run_all(rasterise_pages, [([svg for _, _, svg in chunk], laid["raster"])
                                                 for chunk in chunks],
                               timeouts=[limits.render_timeout_for(len(chunk)) for chunk in chunks])
    if laid.get("laid_out") is not None:
        call.document._tool_render_layout = (layout_key, laid["laid_out"])
    pngs: dict = {}
    failed = None
    for chunk, result in zip(chunks, results):
        if isinstance(result, BaseException):
            failed = failed or result
            continue
        for (page, key, _), png in zip(chunk, result):
            store.put(key, png)
            pngs[key] = png
    for page, key, svg in laid["pages"]:
        png = pngs.get(key) if svg is not None else held[key][0]
        if png is None:
            continue
        entry.render_keys.put((version, page, width, fonts, locale), key)
        drawn[page] = (png, svg is None)
    if failed is not None:
        raise failed                     # what the other chunks drew is kept
    return drawn, laid["count"], laid["coverage"]


@handler("render")
def render(call: Any, doc: str, slides: list | None = None, pages: list[int] | None = None,
           width: int | None = None, sheet: dict | None = None) -> Result:
    if sheet is not None:
        return _render_sheet(call, _pages_of(call, slides, pages), width, sheet)
    pages = list(dict.fromkeys(_pages_of(call, slides, pages) or [1]))
    width = width or DEFAULT_WIDTH
    if len(pages) > call.limits.max_images_per_call:
        raise ToolError("limit", f"at most {call.limits.max_images_per_call} pages per call; sheet: {{}} tiles "
                        f"up to {tiles.MAX_TILES} into one image", field="pages")
    call.check_images(len(pages))        # before drawing: a refused call costs nothing
    entry, version = call.entry, call.entry.version
    fonts = tuple(font_dirs_of(call.document))
    locale = locale_of(call.document)
    drawn, count, coverage = draw(call, pages, width)
    if len(drawn) < len(pages):
        raise ToolError("not_found", f"the document has {count} page(s)", field="pages",
                        valid_options=list(range(1, count + 1))[:50])
    if coverage is not None:
        call.document._tool_render_coverage = ((version, fonts, locale), coverage)
    images = []
    for number in pages:
        png = drawn[number][0]
        w, h = png_size(png)
        images.append(call.image(png, w, h, label=f"page {number}").describe())
    result = Result(summary=f"Rendered page(s) {', '.join(map(str, pages))} (final view: changes accepted, "
                    "comments hidden)", data={"pages": pages,
                                              "cached": [p for p in pages if drawn[p][1]]})
    render_version, coverage = getattr(call.document, "_tool_render_coverage", (None, None))
    if coverage is not None and render_version == (version, fonts, locale):
        result.data["coverage"] = coverage
    return result


def _render_sheet(call: Any, pages: list[int] | None, width: int | None, sheet: dict) -> Result:
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
           fonts, locale)
    made = entry.render_cache.get(key)
    cached = made is not None
    tiles_cached = None
    if made is None:
        drawn, count, coverage = draw(call, pages, tiling.width)
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
            drawn, _, again = draw(call, pages, tiling.width)
            coverage = coverage if again is None else again
        tiles_cached = [p for p in pages if drawn[p][1]]
        with call.unlocked():
            # The snapshot is taken (the tiles); the worker draws the sheet from it.
            views = [(drawn[p][0], *png_size(drawn[p][0]), f"page {p}") for p in pages]
            tiles.check_size(tiling, [h for _, _, h, _ in views], name="sheet", what="pages",
                             max_edge=limits.max_image_edge)
            svg, w, h = tiles.tiled_svg(views, tiling)
            png = call.run(rasterise_tiles, svg, list(fonts))
        if coverage is not None:
            call.document._tool_render_coverage = ((version, fonts, locale), coverage)
        made = (png, w, h, pages, tiling)
        entry.render_cache.put(key, made)
    png, w, h, pages, tiling = made
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
    return result


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
          min_size: float | None = None, styles: list[str] | None = None) -> Result:
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
    if (min_size is not None or styles is not None) and "layout" not in wanted:
        wanted.append("layout")               # what they measure (round 16, F8)
    document, entry = call.document, call.entry
    report: dict[str, Any] = {}
    if "validate" in wanted:
        problems = [str(p) for p in document.validate()]
        report["validate"] = {**validate_delta(entry), "problems": problems[:50]}
    if "reflow" in wanted or "layout" in wanted:
        reflow, layout = _layout_facts(call, pages, seen="reflow" in wanted)
        if "reflow" in wanted:
            report["reflow"] = reflow
        if "layout" in wanted:
            # Measured on the layout reflow reports and render draws: never a second one.
            report["layout"] = (layout_facts(layout, pages=pages, min_size=min_size, styles=styles)
                                if layout is not None else {"error": reflow["error"]})
    if "fields" in wanted:
        report["fields"] = _field_facts(document)
    if "app" in wanted:
        report["app_findings"] = None
    return Result(summary="Checked " + ", ".join(wanted), data=report)


def _layout_facts(call: Any, pages: list[int] | None, seen: bool = True) -> tuple[dict[str, Any], Any]:
    """``reflow``'s facts, and the layout they are of (``None`` where it failed); ``seen``:
    the layout a later check's ``since_last_check`` is measured from (not for ``layout``
    alone, which reports no reflow)."""
    document = call.document
    entry = call.entry
    version = entry.version
    try:
        if call.may_unlock:
            # Laid out with the lock let go (a check outside a batch): of this version, the
            # one validate and fields read and the result reports.
            layout = lay_out_unlocked(document, call.unlocked, lambda: entry.version == version)
        else:
            layout = document.layout()
    except ToolError as error:
        return {"error": f"{error.code}: {error.message}"}, None
    facts: dict[str, Any] = {"pages": layout.page_count, "complete": layout.pages_known is not None}
    if layout.stopped is not None:
        facts["stopped"] = {"page": layout.stopped.page, "reason": layout.stopped.reason, "at": layout.stopped.at}
    # The whole of it: blocks laid out, header and footer stops, faces substituted or
    # missing -- so "the check passed" is told from "the check could not see everything".
    facts["coverage"] = layout.coverage_facts()
    before = getattr(document, "_tool_seen_layout", None)
    if before is not None and before is not layout:
        reflow = reflow_json(layout.compare(before))
        if pages:
            reflow["changed"] = [p for p in reflow["changed"] if p in pages]
        facts["since_last_check"] = reflow
    if seen:
        document._tool_seen_layout = layout
    if layout.warnings:
        facts["layout_warnings"] = [short(str(w), 160) for w in layout.warnings[:10]]
    return facts, layout


def _field_facts(document: Document) -> dict[str, Any]:
    fields = document.fields()
    counts: dict[str, int] = {}
    for field in fields:
        counts[field.keyword] = counts.get(field.keyword, 0) + 1
    facts: dict[str, Any] = {"counts": counts}
    empty = document._empty_tocs()
    if empty:
        facts["empty_toc"] = empty
    missing = []
    for field in fields:
        if field.keyword in ("REF", "PAGEREF", "NOTEREF"):
            words = field.instruction.split()
            name = words[1] if len(words) > 1 else ""
            try:
                document.bookmark(name)
            except KeyError:
                missing.append(f"{field.id}: {field.keyword} to a missing bookmark {name!r}")
    if missing:
        facts["missing_bookmarks"] = missing
    facts["note"] = "results are caches: after moving headings or text, word_fields update"
    return facts


# -- S10-S12 charts and SmartArt ---------------------------------------------------------------


#: What ``add`` needs, said when it is missing.
_ADD_NEEDS = {"chart_type": "one of " + ", ".join(shared.CHART_TYPES),
              "data": "each series as {name, values}, one value per category",
              "categories": "the category labels, in order"}


def _chart_need(chart: Any, action: str, name: str, given: Any) -> Any:
    """``given``, or ``invalid_arguments`` saying what ``action`` needs ``name`` to be on this
    chart: the series or categories (with their positions), the axes, the numbers."""
    if given is not None:
        return given
    from ooxml_edit.charts.actions import choices

    if action == "add":
        expected = _ADD_NEEDS[name]
    elif name == "series":
        expected = choices([s.name for s in chart.series])
    elif name == "category" and action != "add_category":
        expected = choices(_labels(chart))
    elif name == "axis":
        expected = " or ".join(chart.axes) or "the chart has no axes"
    elif name == "values":
        expected = (f"one number per series ({len(chart.series)})" if action == "add_category"
                    else f"one number per category ({len(chart.categories)})")
    elif name == "value":
        expected = ("the gap as a whole % of a bar's width, 0-500" if action == "set_gap_width"
                    else "the number")
    elif name == "position":
        expected = "right, left, top, bottom or none"
    else:
        expected = {"add_category": "the new category's label",
                    "add_series": "the new series' name"}.get(action, "the new name")
    raise ToolError("invalid_arguments", f"{action} needs {name}: {expected}", field=name)


def _series(chart: Any, which: str | None, action: str = "this action") -> Any:
    which = _chart_need(chart, action, "series", which)
    if which.isdigit() and int(which) < len(chart.series):
        return chart.series[int(which)]
    try:
        return chart.series_named(which)
    except (KeyError, ValueError):
        raise ToolError("not_found", f"no series {which!r}", field="series",
                        valid_options=[s.name for s in chart.series]) from None


def _labels(chart: Any) -> list[str]:
    """The categories as a model reads and names them: ISO dates for dates kept as serials
    (``chart.category_dates``), else each label as text."""
    from ooxml_edit.charts.actions import category_labels

    return category_labels(chart)


def _category(chart: Any, which: str | None, action: str = "this action") -> int:
    which = _chart_need(chart, action, "category", which)
    labels = _labels(chart)
    if which in labels:
        return labels.index(which)
    if which.isdigit() and int(which) < len(labels):
        return int(which)
    hint = "; its categories are dates, named as YYYY-MM-DD" if chart.category_dates is not None else ""
    raise ToolError("not_found", f"no category {which!r}{hint}", field="category", valid_options=labels)


def chart_json(chart: Any) -> dict[str, Any]:
    data = {"id": chart.address if hasattr(chart, "address") else None,
            "type": "dot" if chart.is_dot else chart.chart_type,
            "title": chart.title, "categories": _labels(chart),
            "series": [{"name": s.name, "values": list(s.values)} for s in chart.series],
            "legend": chart.has_legend}
    if chart.category_dates is not None:
        # Dates kept as serials: listed as ISO dates, and how the chart shows them.
        data["category_format"] = chart.category_format
    return data


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
        return _edit_chart(call, target, action, number_format, series, category, values, value, text, axis,
                           position)
    except (ChartDataError, IndexError) as error:
        raise ToolError("invalid_arguments", str(error).strip("'\"")) from None


def _style_chart(call: Any, target: str, action: str, *, series: str | None, category: str | None,
                 **arguments: Any) -> Result:
    """The formatting actions (ooxml-edit's ``apply_style_action``)."""
    from ooxml_edit.charts.actions import apply_style_action, chart_style_json

    chart = call.document.chart(target)
    chosen = None if series is None else _series(chart, series, action).index
    point = None if category is None else _category(chart, category, action)
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
    arguments give (``add_style_actions``: the same chart as the separate calls)."""
    if box is not None:
        raise ToolError("invalid_arguments", "box is for decks; in a document give width", field="box")
    chart_type = _chart_need(None, "add", "chart_type", chart_type)
    data = _chart_need(None, "add", "data", data)
    if chart_type in ("scatter", "bubble"):
        if values is None or categories is not None:
            raise ToolError("invalid_arguments", f"a {chart_type} chart takes its x values in values, not "
                            "categories", field="values")
        labels = list(values)
    else:
        if values is not None:
            raise ToolError("invalid_arguments", "values are a scatter or bubble chart's x values; give each "
                            "series' values in data", field="values")
        labels = list(_chart_need(None, "add", "categories", categories))
    if len(labels) * max(len(data), 1) > 10_000:
        raise ToolError("limit", "at most 10,000 chart values", field="data")
    if position not in (None, "none", "right", "left", "top", "bottom"):
        raise ToolError("invalid_arguments", f"{position} is a label position; a legend is right, left, top, "
                        "bottom or none", field="position")
    from ooxml_edit.charts.actions import (add_style_actions, apply_add_style, apply_style_action,
                                           scale_axis)

    look = add_style_actions(chart_type, data, data_labels=data_labels, colors=colors, style=style, part=part)
    legend = None if position == "none" else (position or "default")
    reverse = bool((scale or {}).get("reverse")) and axis == "category"
    rest = {k: v for k, v in (scale or {}).items() if not (reverse and k == "reverse")}
    with call.document._edit():  # the chart and its look: one undo step
        edit = call.document.insert_chart(target, chart_type, labels, [dict(entry) for entry in data],
                                          width=width, title=text, legend=legend, number_format=number_format,
                                          category_order="reversed" if reverse else "normal",
                                          gridlines=gridlines != "none")
        chart = call.document.chart(edit.id)
        if rest:  # no axis: the chart's one value axis (refused, and undone, if it has more)
            apply_style_action(chart, "set_axis_scale", scale=rest,
                               axis=axis or scale_axis(chart, rest, action="add with scale"))
        if gridlines not in (None, "major", "none"):
            apply_style_action(chart, "set_gridlines", axis="y" if chart_type in ("scatter", "bubble")
                               else "value", gridlines=gridlines)
        apply_add_style(chart, look)
    remember(call, ref, edit.id)
    chart = call.document.chart(edit.id)
    data = chart_json(chart)
    paragraphs = [identifier for identifier in edit.created if identifier.startswith("p:")]
    if paragraphs:
        data["paragraph"] = paragraphs[0]  # where a caption goes after
    return outcome([edit], f"Inserted a {chart_type} chart {edit.id} in {data.get('paragraph', target)}",
                   data=data)


def _edit_chart(call: Any, target: str, action: str, number_format: str | None, series: str | None,
                category: str | None, values: list[float] | None, value: float | None, text: str | None,
                axis: str | None, position: str | None) -> Result:
    chart = call.document.chart(target)
    if action == "read":
        return _read_chart(chart, target)

    def want(name: str, given: Any) -> Any:
        return _chart_need(chart, action, name, given)

    if action == "set_values":
        _series(chart, series, action).set_values(want("values", values))
    elif action == "set_value":
        _series(chart, series, action).set_value(_category(chart, category, action), want("value", value))
    elif action == "add_category":
        chart.add_category(want("category", category), want("values", values))
    elif action == "remove_category":
        chart.remove_category(_category(chart, category, action))
    elif action == "rename_category":
        chart.set_category(_category(chart, category, action), want("text", text))
    elif action == "add_series":
        chart.add_series(want("text", text), want("values", values))
    elif action == "remove_series":
        chart.remove_series(_series(chart, series, action).index)
    elif action == "rename_series":
        _series(chart, series, action).set_name(want("text", text))
    elif action == "set_title":
        chart.set_title(text)
    elif action == "set_axis_title":
        chart.set_axis_title(want("axis", axis), text)
    elif action in ("show_data_labels", "hide_data_labels"):
        chosen = None if series is None else [_series(chart, series, action).index]
        chart.set_data_labels(action == "show_data_labels", number_format=number_format, series=chosen)
    elif action == "set_gap_width":
        given = want("value", value)
        if given != int(given):
            raise ToolError("invalid_arguments", "a gap width is a whole percentage, 0-500", field="value")
        chart.set_gap_width(int(given))
    else:
        where = want("position", position)
        chart.set_legend(where != "none", {"right": "r", "left": "l", "top": "t", "bottom": "b"}.get(where, "r"))
    chart = call.document.chart(target)
    return Result(summary=f"{action} on {target}", changed=[target], data=chart_json(chart))


def _read_chart(chart: Any, target: str) -> Result:
    """``edit_chart`` with ``action: "read"``: the chart as drawn and its workbook."""
    data = chart_json(chart)
    data["number_format"] = chart.number_format
    from ooxml_edit.charts.actions import chart_style_json

    data.update(chart_style_json(chart))
    try:
        data["workbook"] = chart.workbook_values()
    except Exception as error:  # noqa: BLE001 -- a chart without a readable workbook
        data["workbook"] = f"unreadable: {error}"
    return Result(summary=f"Read {target}", data=data)


@handler("edit_smartart", mutates=True)
def edit_smartart(call: Any, doc: str, target: str, action: str, node: int | None = None,
                  text: str | None = None) -> Result:
    diagram = call.document.diagram(target)
    if action == "set_text":
        diagram.set_text(need(node, "node"), need(text, "text"))
    elif action == "add_node":
        diagram.add_node(text or "")
    elif action == "remove_node":
        diagram.remove_node(need(node, "node"))
    else:
        diagram.add_node(text or "", parent=diagram.node(need(node, "node")))
    diagram = call.document.diagram(target)
    return Result(summary=f"{action} on {target}", changed=[target],
                  data={"nodes": [short(t, 80) for t in diagram.texts]})


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


TOOLS = [new_document, save_document, describe, find_text, replace_text, render, check, edit_chart,
         edit_smartart, set_properties, *shared.SESSION_TOOLS]
