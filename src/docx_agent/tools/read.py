"""Reading: the document at a glance (``word_describe``, the shared ``describe``'s Word
handler, made a tool in :mod:`.shared`), ``word_read`` (Markdown with ids,
paged) and ``word_inspect`` (a block's exact formatting)."""

from __future__ import annotations

from collections import Counter
from typing import Any

from ooxml_edit.tools import Result, ToolError, boolean, page_text, string

from ..edit.document import _MOVABLE, W_BODY, W_P, _HeadingLevels
from ..markdown.read import _TYPED_NUMBER
from ._base import resolve, short, word_tool

DOC = string("Document id.")
CURSOR = string("next_cursor of the previous page.", optional=True)

#: Leave room in a page for the envelope around the Markdown.
_PAGE_MARGIN = 1500


def heading_tree(document: Any) -> list[dict[str, Any]]:
    """The body's headings in order (LW6's tree, composed): id, level, text, how it is
    numbered (``list`` or typed ``text``), and ``blocks``: how many blocks its section holds,
    itself and everything up to the next heading at its level or above."""
    levels = _HeadingLevels(document)
    body = document.package.document_part()
    index = document._index(body)
    blocks: list[tuple[int | None, str | None]] = []   # (heading level or None, id) per body block
    root = document.package.tree(body).find(W_BODY)
    for node in root:
        if not isinstance(node.tag, str) or node.tag not in _MOVABLE:
            continue
        if node.tag == W_P:
            entry = index.entry_for(node)
            blocks.append((levels.of(node), entry.id if entry is not None else None))
        else:
            blocks.append((None, None))
    out = []
    for k, (level, identifier) in enumerate(blocks):
        if level is None or identifier is None:
            continue
        end = next((j for j in range(k + 1, len(blocks)) if blocks[j][0] is not None and blocks[j][0] <= level),
                   len(blocks))
        paragraph = document.paragraph(identifier)
        text = paragraph.text
        item: dict[str, Any] = {"id": identifier, "level": level, "text": short(text, 100), "blocks": end - k}
        if paragraph.list is not None:
            item["numbered"] = "list"
        elif _TYPED_NUMBER.match(text):
            item["numbered"] = "text"
        out.append(item)
    return out


def _pages(call: Any) -> Any:
    try:
        layout = call.document.layout()
    except ToolError as error:
        return f"{error.code}: {error.message}"
    known = layout.pages_known
    return known if known is not None else f"at least {layout.page_count} (layout stopped)"


def _sections(document: Any) -> list[dict[str, Any]]:
    out = []
    for section in document.sections():
        width, height = section.page_width, section.page_height
        item = {"id": section.id, "start": section.start, "orientation": section.orientation,
                "page": [width, height], "margins": {k: section.margins[k] for k in
                                                     ("top", "right", "bottom", "left")},
                "columns": section.columns.get("count", 1)}
        stories = section.stories()
        for which in ("header", "footer"):
            shown = {kind: value["story"] for kind, value in stories[which].items() if value["story"]}
            if shown:
                item[which + "s"] = shown
        out.append(item)
    return out


def _styles_in_use(document: Any) -> list[dict[str, Any]]:
    counts: Counter = Counter()
    for story in document.stories:
        for paragraph in story.paragraphs:
            counts[paragraph.style_name] += 1
    return [{"name": name, "paragraphs": n} for name, n in counts.most_common(40) if name]


def word_describe(call: Any, doc: str) -> Result:
    """The shared ``describe``'s Word handler: the document at a glance."""
    document = call.document
    entry = call.entry
    comments = document.comments()
    threads = [c for c in comments if c.parent is None]
    revisions = document.revisions()
    fields = Counter(f.keyword for f in document.fields())
    data: dict[str, Any] = {
        "name": entry.name, "pages": _pages(call), "compatibility_mode": document.compatibility_mode,
        "properties": {k: v for k, v in document.properties.items() if k in ("title", "author", "subject")},
        "tracking": ({"author": entry.tracking["author"]} if entry.tracking and entry.tracking.get("on")
                     else None),
        "headings": heading_tree(document)[:80],
        "sections": _sections(document),
        "styles_in_use": _styles_in_use(document),
        "comments": {"threads": len(threads), "open": sum(1 for c in threads if not c.done),
                     "by_author": dict(Counter(c.author for c in comments))},
        "revisions": {"records": len(revisions), "changes": len(document.changes()),
                      "by_author": dict(Counter(r.author for r in revisions))},
        "fields": dict(fields),
        "tables": [{"id": t.id, "rows": len(t.rows), "columns": t.column_count,
                    "first_row": short(" | ".join(_cell_text(t, 0, k) for k in range(t.column_count)), 80)}
                   for t in document.tables()][:30],
        "drawings": [{"id": d.id, "kind": drawing_kind(d), **({"alt_text": short(d.alt_text, 60)} if d.alt_text else {})}
                     for d in _drawings(document)][:30],
        "notes": dict(Counter(n.kind for n in document.notes())),
        "controls": len(document.content_controls()),
        "problems": [str(p) for p in entry.baseline_problems][:20],
    }
    if len(data["headings"]) == 80:
        data["headings_note"] = "the first 80 headings; word_read a range for the rest"
    return Result(summary=f"Described {entry.name}", data=data)


def _cell_text(table: Any, row: int, column: int) -> str:
    try:
        return table.cell(row, column).text
    except IndexError:
        return ""


def drawing_kind(drawing: Any) -> str:
    return getattr(drawing, "kind", None) or "picture"


def _drawings(document: Any) -> list[Any]:
    seen = []
    for picture in document.pictures():
        seen.append(document.drawing(picture.id))
    known = {d.id for d in seen}
    for chart in document.charts():
        identifier = getattr(chart, "id", None)
        if identifier and identifier not in known:
            seen.append(document.drawing(identifier))
            known.add(identifier)
    return seen


@word_tool("word_read", "The document as Markdown with an id comment before each block. Paged: pass next_cursor back. view markup shows tracked changes, comments and fields in place.",
           {"doc": DOC,
            "range": string("A block or span p:A..t:B. Default the whole body.", optional=True),
            "view": string("final: changes accepted; markup: changes, comments, fields in place. Default final.",
                           enum=["final", "markup", "original"], optional=True),
            "stories": string("all adds headers, footers, notes. Default body.",
                              enum=["body", "all"], optional=True),
            "cursor": CURSOR},
           refs=('range',), group="core")
def word_read(call: Any, doc: str, range: str | None = None, view: str = "final",
              stories: str = "body", cursor: str | None = None) -> Result:
    if range is not None and stories == "all":
        raise ToolError("invalid_arguments", "stories=all reads whole stories: give no range with it",
                        field="range")
    entry = call.entry
    range = resolve(call, range)
    key = ("markdown", range, view, stories)
    read = lambda: call.document.to_markdown(range, view=view, stories=stories)  # noqa: E731
    # Inside a batch the version moves only at its end: read the state as it is now.
    markdown = read() if call.in_batch else entry.cached(entry.check_cache, key, read)
    limit = max(2000, call.limits.max_result_chars - _PAGE_MARGIN)
    text, next_cursor = page_text(markdown, cursor=cursor, limit=limit)
    result = Result(summary=f"Read {range or 'the document'} ({view})", data=text,
                    next_cursor=next_cursor)
    result.total = len(markdown)
    return result


@word_tool("word_inspect", "Exact formatting of blocks: style, runs, effective font, size, spacing, indents, list, revisions; with layout, where each is placed.",
           {"doc": DOC,
            "target": string("A block or span p:A..t:B."),
            "layout": boolean("Also each block's page, top and bottom. Default false.",
                              optional=True)},
           refs=('target',), group="word_text")
def word_inspect(call: Any, doc: str, target: str, layout: bool = False) -> Result:
    state = call.document.state(resolve(call, target), layout=layout)
    data = {key: state[key] for key in ("styles", "blocks", "notes", "comments", "revisions")
            if state.get(key)}
    return Result(summary=f"Inspected {target}", data=data)


TOOLS = [word_read, word_inspect]
