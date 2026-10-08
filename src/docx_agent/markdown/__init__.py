"""The Markdown layer: ``to_markdown`` with ids and views, and ``insert_markdown``
(ROADMAP.md, "The Markdown layer"; Phase E2).

``Document.to_markdown(range=None, *, view="final", ids=True, headers=False, notes=True)``
projects the document as CommonMark plus GFM tables, strikethrough and footnotes, every
block carrying its id in an HTML comment.  It reads and never writes: not one byte of the
document changes, and no id is stamped (a paragraph without a paraId shows its volatile
positional id, marked ``volatile``).

The modules: :mod:`.stylemap` is the one table both directions read; :mod:`.read` walks
the document into records and the neutral :mod:`.model`; :mod:`.render` writes the model
as text; :mod:`.parse` reads Markdown back with markdown-it-py into a normalised AST --
the reverse-parse check, and the model :mod:`.write` (``insert_markdown``) writes into the
document with the same style map.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from . import model
from .parse import differences, model_ast, parse_ast
from .read import (Marker, Options, ParagraphRecord, Reader, TableRecord, build, note_definitions,
                   referenced_notes, walk_records)
from .render import render
from .stylemap import DEFAULT, Rule, StyleMap

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def project(document: "Document", range: str | None = None, *, view: str = "final", ids: bool = True,
            headers: bool = False, notes: bool = True, style_map: StyleMap | None = None,
            reader: Reader | None = None, stories="body") -> list:
    """The model blocks ``to_markdown`` writes (see :func:`to_markdown`).  ``reader``: one
    :class:`Reader` of the document as it is now, to read several ranges with (it keeps
    each story's records; the document must not change while it is used)."""
    reader = reader or Reader(document, view, style_map)
    options = Options(ids=ids)
    blocks: list = []
    labels: list[str] = []
    extra, every_note = _stories(document, stories)
    if extra and range is not None:
        raise ValueError("stories= reads whole stories: give no range with it")
    headers = headers and not extra
    for part, records, kind in select(reader, range):
        if kind == "notes":
            labels += [identifier for where, identifier, _ in document._notes() if where == part]
            continue
        if kind == "story" and ids:
            blocks.append(model.Comment(f"story: {document._story_of(part)}"))
        blocks += build(reader, records, options)
        labels += [n for r in walk_records(records) for n in referenced_notes(r)]
        labels += [n for r in walk_records(records) if isinstance(r, ParagraphRecord)
                   for _, box in r.text_boxes for n in _box_notes(reader, r, box)]
    if headers and range is None:
        for part in document.package.story_parts():
            root = document.package.tree(part)
            if root is not None and root.tag in (_W + "hdr", _W + "ftr"):
                records = reader.story(part)
                if ids:
                    kind = "header" if root.tag == _W + "hdr" else "footer"
                    blocks.append(model.Comment(f"story: {document._story_of(part)} ({kind})"))
                blocks += build(reader, records, options)
                labels += [n for r in walk_records(records) for n in referenced_notes(r)]
    for part in extra:
        root = document.package.tree(part)
        records = reader.story(part)
        if ids:
            kind = {_W + "hdr": "header", _W + "ftr": "footer"}.get(root.tag, "story")
            blocks.append(model.Comment(f"story: {document._story_of(part)} ({kind})"))
        blocks += build(reader, records, options)
        labels += [n for r in walk_records(records) for n in referenced_notes(r)]
    if every_note:
        labels += [identifier for where, identifier, _ in document._notes() if where in every_note]
    if notes or every_note:
        blocks += note_definitions(reader, _unique(labels), options)
    return blocks


#: What ``stories=`` takes besides a list of story names.
STORIES = ("body", "all")


def _stories(document: "Document", stories) -> tuple[list[str], set[str]]:
    """``stories``: the header and footer parts to write after the body, and the notes parts
    whose every note is written.  ``"body"``: none; ``"all"``: every header and footer with
    content, and every note; a list names stories (``"body"``, ``"header1"``, ``"footnotes"``...)."""
    if stories in (None, "body"):
        return [], set()
    if isinstance(stories, str) and stories != "all":
        stories = [stories]
    parts = document.package.story_parts()
    if stories == "all":
        wanted = [p for p in parts if _has_content(document, p)]
    else:
        names = list(stories)
        unknown = [n for n in names if n != "body" and n not in {document._story_of(p) for p in parts}]
        if unknown:
            raise KeyError(f"no story {unknown[0]!r}: the stories are body, "
                           + ", ".join(document._story_of(p) for p in parts))
        wanted = [p for p in parts if document._story_of(p) in names]
    extra = [p for p in wanted if document.package.tree(p).tag in (_W + "hdr", _W + "ftr")]
    notes = {p for p in wanted if document.package.tree(p).tag in (_W + "footnotes", _W + "endnotes")}
    return extra, notes


def _has_content(document: "Document", part: str) -> bool:
    """Whether a story holds anything a reader sees: text, a drawing, a field, a note."""
    root = document.package.tree(part)
    if root is None:
        return False
    if root.tag in (_W + "footnotes", _W + "endnotes"):
        return any(n.get(_W + "type") in (None, "normal") for n in root)
    if any((t.text or "").strip() for t in root.iter(_W + "t")):
        return True
    return any(True for _ in root.iter(_W + "drawing", _W + "pict", _W + "fldChar", _W + "fldSimple",
                                       _W + "commentReference"))


def _box_notes(reader: Reader, record: ParagraphRecord, box) -> list[str]:
    return [n for r in walk_records(reader.container(box, record.part)) for n in referenced_notes(r)]


def _unique(labels: list[str]) -> list[str]:
    seen: set[str] = set()
    return [x for x in labels if not (x in seen or seen.add(x))]


def to_markdown(document: "Document", range: str | None = None, *, view: str = "final", ids: bool = True,
                headers: bool = False, notes: bool = True, style_map: StyleMap | None = None,
                stories="body") -> str:
    """The document (or ``range`` of it) as Markdown.

    ``range``: ``None`` for the body; a story (``"body"``, ``"header1"``, ``"footnotes"``,
    ``"comments"``); a section (``"s:<id>"``, ``"s:body"`` for the last); a block id
    (``"p:3A1F09C2"``, ``"t:..."``, ``"cc:..."``, a paragraph inside a table names its
    table); or ``"<id>..<id>"``, the blocks from the first through the second.
    ``view``: ``final`` (accepted), ``original`` (rejected) or ``markup`` (CriticMarkup).
    ``ids``: the id comments (off: plain Markdown).  ``headers``: the headers and footers
    after the body.  ``notes``: footnote and endnote definitions for the notes referenced.
    ``stories``: ``"body"`` (the default), ``"all"`` -- the body, then every header and
    footer with content (each after a ``<!-- story: header1 (header) -->`` comment), then
    every footnote and endnote -- or a list of story names.  In the markup view the
    comments of those stories show where they are attached, as the body's do.

    In the ``markup`` view a field's result is marked where it is
    (``<!-- field: REF RefIncident \\h -->section 4<!-- /field -->``), so computed text
    (cross-references, page numbers, a table of contents) is told from typed text.  A heading
    a list numbers says so in its id comment (``numbered: list``, its number written before
    its text as Word draws it); one whose number is typed into its text says
    ``numbered: text`` -- renumbering it is editing its text.
    """
    return render(project(document, range, view=view, ids=ids, headers=headers, notes=notes,
                          style_map=style_map, stories=stories))


def check(document: "Document", **options) -> list[str]:
    """The reverse-parse check: where markdown-it reads ``to_markdown``'s text differently
    from the model it was written from (empty when they agree)."""
    blocks = project(document, **options)
    text = render(blocks)
    return differences(model_ast(blocks), parse_ast(text))


# -- ranges ----------------------------------------------------------------------------------


def select(reader: Reader, range: str | None) -> list[tuple[str, list, str]]:
    """``(part, records, kind)`` to write: ``kind`` is ``body``, ``story`` or ``notes``."""
    document = reader.document
    body = document.package.document_part()
    if range is None:
        return [(body, reader.story(body), "body")]
    if ".." in range:
        first, _, last = range.partition("..")
        records = _story(reader, _part_of(document, first))
        a, b = _position(document, records, first), _position(document, records, last)
        if _part_of(document, first) != _part_of(document, last):
            raise KeyError(f"{range!r} spans two stories")
        if b < a:
            a, b = b, a
        return [(_part_of(document, first), records[a:b + 1], "range")]
    if range.startswith(("s:", "s@")):
        records = reader.story(body)
        sections = [k for k, r in enumerate(records) if isinstance(r, ParagraphRecord) and r.section is not None]
        wanted = document.get(range).id
        if wanted == "s:body":
            start = sections[-1] + 1 if sections else 0
            return [(body, records[start:], "range")]
        for n, k in enumerate(sections):
            if records[k].section[0] == wanted:
                start = sections[n - 1] + 1 if n else 0
                return [(body, records[start:k + 1], "range")]
        raise KeyError(f"no section {range!r}")
    if ":" not in range and "@" not in range and "/" not in range:
        story = document.story(range)
        root = document.package.tree(story.part)
        if root.tag in (_W + "footnotes", _W + "endnotes"):
            return [(story.part, [], "notes")]
        return [(story.part, reader.story(story.part), "story" if story.part != body else "body")]
    part = _part_of(document, range)
    records = _story(reader, part)
    k = _position(document, records, range)
    return [(part, records[k:k + 1], "range")]


def _story(reader: Reader, part: str) -> list:
    """A story's records, read once per reader."""
    cache = reader.__dict__.setdefault("_stories", {})
    if part not in cache:
        cache[part] = reader.story(part)
    return cache[part]


def _part_of(document: "Document", identifier: str) -> str:
    if identifier.startswith("cc"):
        for part, found, _ in document._content_controls():
            if found == identifier:
                return part
        raise KeyError(f"no content control {identifier!r}")
    return document._resolve(identifier)[0]


def _canonical(document: "Document", identifier: str) -> str:
    if identifier.startswith("cc"):
        return identifier
    part, entry = document._resolve(identifier)
    return entry.id


def _position(document: "Document", records: list, identifier: str) -> int:
    """The index of the top-level record holding ``identifier``."""
    wanted = _canonical(document, identifier)
    for k, record in enumerate(records):
        if isinstance(record, Marker):
            if record.id == wanted:
                return k
            continue
        ids = set()
        for inner in walk_records([record]):
            if isinstance(inner, ParagraphRecord):
                ids.add(inner.id)
                ids.update(inner.joins)
            elif isinstance(inner, TableRecord):
                ids.add(inner.id)
        if wanted in ids:
            return k
    raise KeyError(f"{identifier!r} is not a block of its story in this view")


__all__ = ["DEFAULT", "Rule", "StyleMap", "check", "project", "select", "to_markdown"]
