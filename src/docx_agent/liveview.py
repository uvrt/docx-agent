"""An application's live view of a document, from the layout the tools share (round 35, R27).

A page that shows a document as it is written -- each page as SVG, redrawn when an edit
changes it -- needs to know which pages to draw again and what each page shows.  It asks
the layout the tools already made of that version (:mod:`docx_agent.layout`: one per
version, incremental, shared by ``check``, ``render`` and the Python API), never a layout
of its own::

    from docx_agent.liveview import page_svg, view

    shown = view(session, "d1")                          # first call: every page changed
    for number in shown.changed:
        drawn = page_svg(session, "d1", number, shown.token)
        put(number, drawn.svg, drawn.images)             # pictures by name: img/3fa1....png
    ...
    shown = view(session, "d1", since=shown.token)       # after edits: the pages they changed

* :func:`view` -- the pages in order, each with the ids of the body's top-level blocks it
  shows and its size; the pages whose drawing changed since ``since`` (every page without
  one); a new ``token``, naming this version's layout.  ``stale`` (with a ``note``) when
  ``since`` names no layout kept for the document: every page is listed then.
* :func:`page_svg` -- one page of the version a token names, as SVG, drawn from that
  layout even when the document moved on since: ``text="outlines"`` draws text as glyph
  paths (``glyphs``: one set kept across pages, docx2svg's ``ConvertOptions.glyphs``);
  ``images="external"`` gives pictures back by name (:func:`externalise_images`, the
  names pptx2svg's ``liveview`` gives), ``"embedded"`` keeps them in the SVG.

Pages changed are those whose drawing differs from the same page of the token's version
(docx2svg's ``relayout.drawing``: the SVG less its ``data-docx-*`` attributes), and those
it did not have: what docx2svg's ``Layout.changed_pages`` says for a layout made from the
one before.  Pages equal in the layout (and every part but the document's text the same)
are not drawn to tell.

**Locks** (the order is :mod:`docx_agent.layout`'s; round 34's rule: never wait for a
layout in progress holding the document's lock).  ``view(session, doc)`` takes the
document's lock to read its bytes and version, and lets it go while the version is laid out
-- waiting there for a layout another call has in progress -- then takes it back to map the
pages to ids.  Called by a thread that holds the lock already, it lays out on its own
instead of waiting.  ``view(document)`` (a document outside a session: no lock) never
waits; either way the version's layout is the one cached for all.  Pages are compared, and
:func:`page_svg` draws, with no document lock held (only the layout's drawing lock, under
which nothing is waited for).  Laid out in this process (the application's thread), never
through a tool call's worker engine.
"""

from __future__ import annotations

import base64
import hashlib
import io
import re
import secrets
import zipfile
from collections import OrderedDict
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Any, Iterator, Mapping, MutableMapping

import docx2svg

from .layout import _Conversion, _keeper, cache_key, convert, group_ids, tag_with

__all__ = ["PageSvg", "UnknownToken", "View", "ViewPage", "embed_images", "externalise_images",
           "page_svg", "view"]

#: Tokens -- versions' layouts -- kept per document for ``since`` and :func:`page_svg`, the
#: least recently used dropped first.  Each holds its layout (about 80 MB at 79 pages while
#: the layout cache has it too; ``ROADMAP.md``, "Reflow feedback").
VIEW_TOKENS = 4

#: :func:`page_svg`'s ``text`` and ``images``.
TEXT_MODES = ("text", "outlines")
IMAGE_MODES = ("external", "embedded")


@dataclass(frozen=True)
class ViewPage:
    """One page: its 1-based ``number``, the ids of the body's top-level blocks it draws
    something of, in document order (:meth:`~docx_agent.layout.DocumentLayout.blocks_on`),
    and its size in points."""

    number: int
    blocks: tuple[str, ...]
    width: float
    height: float


@dataclass(frozen=True)
class View:
    """:func:`view`'s answer."""

    #: This version's layout: ``since`` for the next view, and :func:`page_svg`'s ``token``.
    token: str
    #: Every page, in order.
    pages: list[ViewPage]
    #: The 1-based pages whose drawing changed since ``since`` -- every page without one, or
    #: with a stale one.  Pages ``since``'s version had past the last are not listed: the
    #: page count says they are gone.
    changed: list[int]
    #: The token the pages were compared with, as given.
    since: str | None = None
    #: ``since`` names no layout kept for this document (another document's, one from
    #: before a restart, or older than the last :data:`VIEW_TOKENS`): every page is in
    #: ``changed``, and ``note`` says so.
    stale: bool = False
    note: str | None = None
    #: The session's version of the document laid out (``None`` outside a session).
    version: int | None = None

    @property
    def page_count(self) -> int:
        return len(self.pages)


@dataclass(frozen=True)
class PageSvg:
    """:func:`page_svg`'s answer: the page's SVG, and -- with ``images="external"`` -- the
    pictures it references, by name; ``warnings``, what was not drawn faithfully (text kept
    as text in ``outlines``, say)."""

    page: int
    token: str
    svg: str
    images: dict[str, bytes] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


class UnknownToken(LookupError):
    """A token :func:`page_svg` was given that names no layout kept for the document."""


@dataclass
class _Seen:
    """A version a view was given: its layout, its pages, its groups' ids, its parts."""

    conversion: _Conversion
    pages: list[ViewPage]
    ids: dict
    #: A digest of every part but the main document's and ``docProps``: what a page drawn
    #: from equal layouts may still draw differently by (a picture's bytes, a face).
    parts: bytes


# -- view ----------------------------------------------------------------------------------------


def view(where: Any, *args: str, since: str | None = None) -> View:
    """The document's pages now, which changed since ``since``, and a new token (the module
    says how): ``view(document)`` or ``view(session, doc)`` -- ``doc`` an open Word
    document's id (``"d1"``)."""
    document, entry, _ = _target(where, args, "view", 0)
    if entry is None:
        conversion = convert(document, converter=None)
        token, seen = _seen(document, conversion)
        version = None
    else:
        with entry.lock:
            version = entry.version
            conversion = convert(document, unlocked=lambda: _let_go(entry.lock),
                                 current=lambda: entry.version == version, converter=None)
            # The pages mapped to ids under the lock: the layout's document is this version.
            token, seen = _seen(document, conversion)
    # Compared with no lock held: drawing a page to tell takes only its layout's.
    if since is None:
        return View(token, seen.pages, [p.number for p in seen.pages], version=version)
    if since == token:
        return View(token, seen.pages, [], since, version=version)
    with _keeper(document):
        old = _tokens(document).get(since)
        if old is not None:
            _tokens(document).move_to_end(since)
    if old is None:
        return View(token, seen.pages, [p.number for p in seen.pages], since, stale=True, version=version,
                    note=f"since {since!r} names no layout kept for this document (another document's, one "
                    f"from before a restart, or older than the last {VIEW_TOKENS}): every page is listed "
                    "as changed")
    return View(token, seen.pages, _changed(old, seen), since, version=version)


def _seen(document: Any, conversion: _Conversion) -> tuple[str, _Seen]:
    """The token of ``conversion`` and what a view keeps of it, made once per version."""
    lock = _keeper(document)
    with lock:
        nonce = document.__dict__.setdefault("_view_nonce", secrets.token_hex(4))
        token = f"{nonce}-{hashlib.sha256(cache_key(conversion.data, conversion.options).encode()).hexdigest()[:24]}"
        seen = _tokens(document).get(token)
        if seen is not None and seen.conversion is not conversion:
            seen = None                       # the same version laid out again: keep the new one
    if seen is None:
        layout = conversion.layout
        pages = [ViewPage(page.number + 1, tuple(layout.blocks_on(page.number + 1)), page.width_pt, page.height_pt)
                 for page in conversion.raw.pages]
        seen = _Seen(conversion, pages, group_ids(layout._document, conversion), _parts(conversion.data))
    with lock:
        tokens = _tokens(document)
        tokens[token] = seen
        tokens.move_to_end(token)
        while len(tokens) > VIEW_TOKENS:
            tokens.popitem(last=False)
    return token, seen


def _tokens(document: Any) -> "OrderedDict[str, _Seen]":
    """The document's kept tokens (under its layouts' lock, :func:`_keeper`)."""
    tokens = document.__dict__.get("_view_tokens")
    if tokens is None:
        tokens = document.__dict__["_view_tokens"] = OrderedDict()
    return tokens


def _changed(old: _Seen, new: _Seen) -> list[int]:
    """The 1-based pages of ``new`` whose drawing differs from ``old``'s page of the same
    number, and those ``old`` did not have."""
    if old.conversion is new.conversion:
        return []
    before, after = old.conversion.raw.pages, new.conversion.raw.pages
    same_parts = old.parts == new.parts
    out = []
    for index, page in enumerate(after):
        if index >= len(before):
            out.append(index + 1)
        elif same_parts and page == before[index]:
            continue
        elif _digest(old.conversion, index) != _digest(new.conversion, index):
            out.append(index + 1)
    return out


def _digest(conversion: _Conversion, index: int) -> bytes:
    """What page ``index`` draws, digested (docx2svg's ``relayout.drawing``), made once:
    the layout's own where docx2svg made it when it compared versions."""
    from docx2svg import relayout

    with conversion.drawing:
        found = conversion.digests.get(index)
        if found is None:
            state = getattr(conversion.raw, "state", None)
            found = (getattr(state, "drawings", None) or {}).get(index)
            if found is None:
                found = relayout.drawing(conversion.raw.pages[index], conversion.data, conversion.fonts)
            conversion.digests[index] = found
    return found


def _parts(data: bytes) -> bytes:
    digest = hashlib.sha256()
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        for info in sorted(package.infolist(), key=lambda i: i.filename):
            name = info.filename
            if name == "word/document.xml" or name.startswith("docProps/"):
                continue
            digest.update(name.encode() + b"\0" + package.read(info) + b"\0")
    return digest.digest()


# -- one page --------------------------------------------------------------------------------------


def page_svg(where: Any, *args: Any, text: str = "text", images: str = "external",
             glyphs: dict | None = None) -> PageSvg:
    """One page of the version ``token`` names, as SVG: ``page_svg(document, page, token)``
    or ``page_svg(session, doc, page, token)``.  Its paragraph groups carry
    ``data-docx-agent-id`` (that version's ids), as :meth:`~docx_agent.Document.render_svg`'s.

    ``text``: ``"text"`` (``<text>``, drawn with the viewer's fonts) or ``"outlines"`` (the
    glyphs' paths: no font needed to view it); with ``glyphs`` -- a dict kept across pages
    and versions -- the glyph definitions go there, each once, and the page references them
    (docx2svg's ``glyph_sheet`` of it goes once in the page showing them).  ``images``:
    ``"external"`` -- each picture referenced by name, its bytes in the answer's
    ``images`` (:func:`externalise_images`) -- or ``"embedded"``.

    Raises :class:`UnknownToken` for a token no :func:`view` of the document gave or no
    longer kept, ``IndexError`` for a page the version does not have, ``ValueError`` for a
    ``text`` or ``images`` it does not take."""
    if text not in TEXT_MODES:
        raise ValueError(f"text must be one of {', '.join(TEXT_MODES)}, not {text!r}")
    if images not in IMAGE_MODES:
        raise ValueError(f"images must be one of {', '.join(IMAGE_MODES)}, not {images!r}")
    if glyphs is not None and text != "outlines":
        raise ValueError("glyphs is the glyph set of outlined pages: give text=\"outlines\" with it")
    document, _, (page, token) = _target(where, args, "page_svg", 2)
    with _keeper(document):
        seen = _tokens(document).get(token)
        if seen is not None:
            _tokens(document).move_to_end(token)
    if seen is None:
        raise UnknownToken(f"token {token!r} names no layout kept for this document (another document's, "
                           f"one from before a restart, or older than the last {VIEW_TOKENS}); call view() for a "
                           "current token and the pages it changed")
    count = len(seen.pages)
    if not isinstance(page, int) or not 1 <= page <= count:
        raise IndexError(f"page {page!r}: the version {token!r} names has {count} page(s), 1 to {count}")
    conversion = seen.conversion
    options = docx2svg.ConvertOptions(**{**conversion.options, "pages": [page], "text": text, "glyphs": glyphs})
    with conversion.drawing:
        svg = docx2svg.page_svgs(conversion.raw, conversion.data, options, fonts=conversion.fonts)[0]
    svg = tag_with(svg, lambda path, part: seen.ids.get((path, part)))
    store: dict[str, bytes] = {}
    if images == "external":
        svg = externalise_images(svg, store)
    return PageSvg(page, token, svg, store, [str(w) for w in options.warnings])


# -- the target ------------------------------------------------------------------------------------


def _target(where: Any, args: tuple, name: str, rest: int) -> tuple[Any, Any, tuple]:
    """``(document, entry, the other arguments)`` of ``(document, ...)`` -- a Document, no
    entry -- or ``(session, doc, ...)``, a session's open Word document; ``rest``: how many
    arguments follow."""
    from .edit.document import Document

    tail = ", page, token" if rest else ""
    if isinstance(where, Document) and len(args) == rest:
        return where, None, tuple(args)
    if isinstance(where, Document) or not callable(getattr(where, "entry", None)) or len(args) != rest + 1:
        raise TypeError(f"{name} takes (document{tail}) or (session, doc{tail})")
    entry = where.entry(args[0])            # not_found names the open documents
    if not isinstance(entry.document, Document):
        raise TypeError(f"{args[0]} is a {entry.kind} document; {name} draws a Word document's pages")
    return entry.document, entry, tuple(args[1:])


@contextmanager
def _let_go(lock: Any) -> Iterator[bool]:
    """The document's lock let go while the version is laid out (the layout's ``unlocked``):
    yields whether it was.  A thread that held it before :func:`view` took it holds it still
    after one release: it is not let go, and the layout is this caller's own (never a wait
    under the lock); a lock that cannot say whether it is held counts as held."""
    lock.release()
    try:
        owned = getattr(lock, "_is_owned", None)
        yield owned is not None and not owned()
    finally:
        lock.acquire()


# -- pictures by name (the names and forms pptx2svg.liveview gives) -------------------------------

#: A picture as docx2svg embeds it: ``href`` or ``xlink:href``, a media type, base64.
_DATA_HREF = re.compile(r'((?:xlink:)?href)="data:([^";,]+);base64,([A-Za-z0-9+/=]*)"')
_SNIFF = ((b"\x89PNG\r\n\x1a\n", "png"), (b"\xff\xd8\xff", "jpg"), (b"GIF87a", "gif"), (b"GIF89a", "gif"),
          (b"BM", "bmp"), (b"II*\x00", "tif"), (b"MM\x00*", "tif"))
_MIME = {"png": "image/png", "jpg": "image/jpeg", "gif": "image/gif", "bmp": "image/bmp", "tif": "image/tiff",
         "webp": "image/webp", "svg": "image/svg+xml"}
_EXTENSION = {mime: extension for extension, mime in _MIME.items()}
_NAMED_HREF = re.compile(r'((?:xlink:)?href)="(?!data:|#)([^"]*)"')
_IMAGE_BY_NAME = re.compile(r'<image\b[^>]*?\bhref="(?!data:|#)([^"]+)"')


def _format(data: bytes, mime: str) -> str:
    for magic, extension in _SNIFF:
        if data.startswith(magic):
            return extension
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    head = data[:512].lstrip()
    if head.startswith(b"<svg") or (head.startswith(b"<?xml") and b"<svg" in data[:4096]):
        return "svg"
    if mime in _EXTENSION:
        return _EXTENSION[mime]
    return re.sub(r"[^A-Za-z0-9.+-]", "_", mime.partition("/")[2] or "bin")


def externalise_images(svg: str, store: MutableMapping[str, bytes], prefix: str = "img/") -> str:
    """``svg`` with every embedded picture (``href="data:...;base64,..."``) referenced by
    name instead, its bytes put in ``store`` under that name: ``prefix``, the first 32 hex
    digits of the bytes' SHA-256, ``.`` and the format the bytes are (else the media type's
    subtype) -- the names pptx2svg's ``liveview.externalise_images`` gives, so one picture
    on several pages, or in a deck and a document, is one entry."""
    if "base64," not in svg:
        return svg

    def replace(match: re.Match) -> str:
        attribute, mime, encoded = match.groups()
        data = base64.b64decode(encoded)
        name = f"{prefix}{hashlib.sha256(data).hexdigest()[:32]}.{_format(data, mime)}"
        store.setdefault(name, data)
        return f'{attribute}="{name}"'

    return _DATA_HREF.sub(replace, svg)


def embed_images(svg: str, images: Mapping[str, bytes]) -> str:
    """``svg`` with every picture :func:`externalise_images` named embedded again; a name
    ``images`` lacks raises ``KeyError`` naming it."""

    def replace(match: re.Match) -> str:
        attribute, name = match.groups()
        data = images.get(name)
        if data is None:
            return match.group(0)
        extension = name.rpartition(".")[2]
        return f'{attribute}="data:{_MIME.get(extension, "image/" + extension)};base64,' \
               f'{base64.b64encode(data).decode("ascii")}"'

    embedded = _NAMED_HREF.sub(replace, svg) if images else svg
    left = _IMAGE_BY_NAME.search(embedded)
    if left is not None:
        raise KeyError(f"the SVG references the picture {left.group(1)!r} and images has no such entry; pass "
                       "the images page_svg gave")
    return embedded
