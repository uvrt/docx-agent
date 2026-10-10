"""The Word handlers of the shared tools (``ooxml_edit.tools.shared``): one definition each
for every format, dispatched by the document's kind.

S1 ``open_document``, S4 ``close_document``, S5 ``undo`` and ``batch`` are
the session's own (``shared.SESSION_TOOLS``); this module adds docx's ``new_document``,
``save_document``, ``find_text``, ``replace_text``, ``render``, ``check``, ``edit_chart`` (its
``read`` action too), ``edit_smartart``, ``set_properties`` and ``describe`` (its handler,
``word_describe``, is in :mod:`.read`)."""

from __future__ import annotations

import struct
from typing import Any

from ooxml_edit.tools import Result, ToolError, shared

from ..edit.document import Document
from ..edit.properties import refreshed_app
from ..layout import font_dirs_of
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
    return slides


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
                            keep_content=False)
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
def save_document(call: Any, doc: str, name: str, format: str) -> Result:
    document = call.document
    if format in ("pptx", "potx", "outline"):
        raise ToolError("invalid_arguments", f"{format} is a deck format; a document saves as docx, dotx or "
                        "markdown", field="format", valid_options=["docx", "dotx", "markdown"])
    if format == "markdown":
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
    result = Result(summary=f"Saved {name} for the application", data=call.output(name, format, data))
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


def render_pages(data: bytes, pages: list[int], width: int, font_dirs: list[str] | None = None
                 ) -> tuple[list[bytes], int, dict | None]:
    """In a worker: the pages as PNGs at ``width`` pixels, the page count, and the layout's
    coverage (``coverage_facts``; ``None`` from a docx2svg without it).  ``font_dirs`` is
    the application's font folders, resolved by the caller (the worker sees neither the
    session nor a variable set since it started); ``None`` leaves the worker's default."""
    import docx2svg

    from ..layout import coverage_facts

    options = docx2svg.ConvertOptions(pages=pages, width=width, font_dirs=font_dirs)
    pngs = docx2svg.convert_docx_to_png(data, options)
    count = (len(docx2svg.convert_docx_to_svg(data, docx2svg.ConvertOptions(font_dirs=font_dirs)))
             if len(pngs) < len(pages) else -1)
    coverage = getattr(options, "coverage", None)
    return list(pngs), count, coverage_facts(coverage) if coverage is not None else None


def png_size(data: bytes) -> tuple[int, int]:
    return struct.unpack(">II", data[16:24])


@handler("render")
def render(call: Any, doc: str, slides: list | None = None, pages: list[int] | None = None,
           width: int | None = None) -> Result:
    pages = list(dict.fromkeys(_pages_of(call, slides, pages) or [1]))
    width = width or DEFAULT_WIDTH
    if len(pages) > call.limits.max_images_per_call:
        raise ToolError("limit", f"at most {call.limits.max_images_per_call} pages per call", field="pages")
    entry = call.entry
    # The font folders, resolved here and handed to the worker; they key the cache, so
    # renders with different folders do not mix.
    fonts = tuple(font_dirs_of(call.document))
    wanted = [p for p in pages if entry.render_cache.get((entry.version, p, width, fonts)) is None]
    if wanted:
        pngs, count, coverage = call.run(render_pages, call.document.to_bytes(), wanted, width, list(fonts),
                                         timeout=call.limits.render_timeout)
        if len(pngs) < len(wanted):
            raise ToolError("not_found", f"the document has {count} page(s)", field="pages",
                            valid_options=list(range(1, count + 1))[:50])
        for number, png in zip(wanted, pngs):
            entry.render_cache.put((entry.version, number, width, fonts), png)
        if coverage is not None:
            call.document._tool_render_coverage = ((entry.version, fonts), coverage)
    images = []
    for number in pages:
        png = entry.render_cache.get((entry.version, number, width, fonts))
        w, h = png_size(png)
        images.append(call.image(png, w, h, label=f"page {number}").describe())
    result = Result(summary=f"Rendered page(s) {', '.join(map(str, pages))} (final view: changes accepted, "
                    "comments hidden)", data={"pages": pages})
    version, coverage = getattr(call.document, "_tool_render_coverage", (None, None))
    if coverage is not None and version == (entry.version, fonts):
        result.data["coverage"] = coverage
    return result


# -- S9 check ----------------------------------------------------------------------------------


@handler("check")
def check(call: Any, doc: str, slides: list | None = None, pages: list[int] | None = None,
          include: list[str] | None = None, boxes: bool | None = None) -> Result:
    pages = _pages_of(call, slides, pages)
    if boxes is not None:
        raise _decks_only("boxes", "include")
    wanted = include or ["validate", "reflow", "fields"]
    deck_only = [k for k in wanted if k in ("fit", "collisions", "facts", "design")]
    if deck_only:
        raise ToolError("invalid_arguments", f"{', '.join(deck_only)} are deck facts; a document reports "
                        "validate, reflow and fields", field="include", valid_options=["validate", "reflow", "fields"])
    document, entry = call.document, call.entry
    report: dict[str, Any] = {}
    if "validate" in wanted:
        problems = [str(p) for p in document.validate()]
        report["validate"] = {**validate_delta(entry), "problems": problems[:50]}
    if "reflow" in wanted:
        report["reflow"] = _layout_facts(call, pages)
    if "fields" in wanted:
        report["fields"] = _field_facts(document)
    if "app" in wanted:
        report["app_findings"] = None
    return Result(summary="Checked " + ", ".join(wanted), data=report)


def _layout_facts(call: Any, pages: list[int] | None) -> dict[str, Any]:
    document = call.document
    try:
        layout = document.layout()
    except ToolError as error:
        return {"error": f"{error.code}: {error.message}"}
    facts: dict[str, Any] = {"pages": layout.page_count, "complete": layout.pages_known is not None}
    if layout.stopped is not None:
        facts["stopped"] = {"page": layout.stopped.page, "reason": layout.stopped.reason, "at": layout.stopped.at}
    # The whole of it: blocks laid out, header and footer stops, faces substituted or
    # missing -- so "the check passed" is told from "the check could not see everything".
    facts["coverage"] = layout.coverage_facts()
    seen = getattr(document, "_tool_seen_layout", None)
    if seen is not None and seen is not layout:
        reflow = reflow_json(layout.compare(seen))
        if pages:
            reflow["changed"] = [p for p in reflow["changed"] if p in pages]
        facts["since_last_check"] = reflow
    document._tool_seen_layout = layout
    if layout.warnings:
        facts["layout_warnings"] = [short(str(w), 160) for w in layout.warnings[:10]]
    return facts


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


def _series(chart: Any, which: str | None) -> Any:
    which = need(which, "series")
    if which.isdigit() and int(which) < len(chart.series):
        return chart.series[int(which)]
    try:
        return chart.series_named(which)
    except (KeyError, ValueError):
        raise ToolError("not_found", f"no series {which!r}", field="series",
                        valid_options=[s.name for s in chart.series]) from None


def _category(chart: Any, which: str | None) -> int:
    which = need(which, "category")
    labels = [str(c) for c in chart.categories]
    if which in labels:
        return labels.index(which)
    if which.isdigit() and int(which) < len(labels):
        return int(which)
    raise ToolError("not_found", f"no category {which!r}", field="category", valid_options=labels)


def chart_json(chart: Any) -> dict[str, Any]:
    return {"id": chart.address if hasattr(chart, "address") else None, "type": chart.chart_type,
            "title": chart.title, "categories": [str(c) for c in chart.categories],
            "series": [{"name": s.name, "values": list(s.values)} for s in chart.series],
            "legend": chart.has_legend}


@handler("edit_chart", mutates=True)
def edit_chart(call: Any, doc: str, target: str, action: str, chart_type: str | None = None,
               categories: list[str] | None = None, data: list[dict] | None = None, box: dict | None = None,
               width: float | None = None, number_format: str | None = None, series: str | None = None,
               category: str | None = None, values: list[float] | None = None, value: float | None = None,
               text: str | None = None, axis: str | None = None, position: str | None = None,
               ref: str | None = None) -> Result:
    from ooxml_edit.charts import ChartDataError

    try:
        if action == "add":
            return _add_chart(call, resolve(call, target), chart_type, categories, data, box, width, values,
                              number_format, text, position, ref)
        return _edit_chart(call, target, action, number_format, series, category, values, value, text, axis,
                           position)
    except (ChartDataError, IndexError) as error:
        raise ToolError("invalid_arguments", str(error).strip("'\"")) from None


def _add_chart(call: Any, target: str, chart_type: str | None, categories: list[str] | None,
               data: list[dict] | None, box: dict | None, width: float | None, values: list[float] | None,
               number_format: str | None, text: str | None, position: str | None, ref: str | None) -> Result:
    """``action: "add"``: a new inline chart in a paragraph of its own after ``target``."""
    if box is not None:
        raise ToolError("invalid_arguments", "box is for decks; in a document give width", field="box")
    chart_type = need(chart_type, "chart_type")
    data = need(data, "data")
    if chart_type == "scatter":
        if values is None or categories is not None:
            raise ToolError("invalid_arguments", "a scatter chart takes its x values in values, not categories",
                            field="values")
        labels = list(values)
    else:
        if values is not None:
            raise ToolError("invalid_arguments", "values are a scatter chart's x values; give each series' "
                            "values in data", field="values")
        labels = list(need(categories, "categories"))
    if len(labels) * max(len(data), 1) > 10_000:
        raise ToolError("limit", "at most 10,000 chart values", field="data")
    legend = None if position == "none" else (position or "default")
    edit = call.document.insert_chart(target, chart_type, labels, [dict(entry) for entry in data], width=width,
                                      title=text, legend=legend, number_format=number_format)
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
    if action == "set_values":
        _series(chart, series).set_values(need(values, "values"))
    elif action == "set_value":
        _series(chart, series).set_value(_category(chart, category), need(value, "value"))
    elif action == "add_category":
        chart.add_category(need(category, "category"), need(values, "values"))
    elif action == "remove_category":
        chart.remove_category(_category(chart, category))
    elif action == "rename_category":
        chart.set_category(_category(chart, category), need(text, "text"))
    elif action == "add_series":
        chart.add_series(need(text, "text"), need(values, "values"))
    elif action == "remove_series":
        chart.remove_series(_series(chart, series).index)
    elif action == "rename_series":
        _series(chart, series).set_name(need(text, "text"))
    elif action == "set_title":
        chart.set_title(text)
    elif action == "set_axis_title":
        chart.set_axis_title(need(axis, "axis"), text)
    elif action in ("show_data_labels", "hide_data_labels"):
        chosen = None if series is None else [_series(chart, series).index]
        chart.set_data_labels(action == "show_data_labels", number_format=number_format, series=chosen)
    elif action == "set_gap_width":
        given = need(value, "value")
        if given != int(given):
            raise ToolError("invalid_arguments", "a gap width is a whole percentage, 0-500", field="value")
        chart.set_gap_width(int(given))
    else:
        where = need(position, "position")
        chart.set_legend(where != "none", {"right": "r", "left": "l", "top": "t", "bottom": "b"}.get(where, "r"))
    chart = call.document.chart(target)
    return Result(summary=f"{action} on {target}", changed=[target], data=chart_json(chart))


def _read_chart(chart: Any, target: str) -> Result:
    """``edit_chart`` with ``action: "read"``: the chart as drawn and its workbook."""
    data = chart_json(chart)
    data["number_format"] = chart.number_format
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
