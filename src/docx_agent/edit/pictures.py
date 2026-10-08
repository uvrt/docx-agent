"""Pictures: inline pictures inserted and replaced, resized, described and deleted.

A picture is ``d:<wp:docPr/@id>`` (unique in the document; a repeat is ``d:<id>#<k>``).
Inserting one writes the ``w:drawing`` Word writes for an inline picture -- ``wp:inline``
with its extent, a ``wp:docPr`` whose id is one above the largest in any story, the
``wp14:anchorId`` and ``wp14:editId`` Word needs to keep the document's paraIds (measured
in E0: one drawing without them and Word writes no paraId at all), and a ``pic:pic``
stretching an ``a:blip`` over it -- in a run of its own at the position, formatted as text
there would be.

**Media** is shared by content: a picture whose bytes a part under ``word/media`` already
holds reuses that part (ooxml-edit's ``find_part_with_bytes``), and a relationship to it is
reused too.  Replacing or deleting a picture releases its relationship, and the media part
goes when nothing else uses it (ooxml-edit's ``release`` and ``reap``).

**Size.**  Without a size, a picture is its natural size: pixels at the density it states,
or 96 pixels to the inch where it states none (Word's assumption for a picture without one;
see ROADMAP.md, Phase E1).  With one dimension given, the other keeps the aspect ratio.
"""

from __future__ import annotations

import copy

import os
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, remove
from . import ids as _ids
from . import inline as _inline
from . import text as _text
from .ranges import TextRange


class ImageBytes(bytes):
    """A picture's bytes; calling them (the old ``picture.image()``) gives them too."""

    def __call__(self) -> bytes:
        import warnings

        warnings.warn("Picture.image is a property now: picture.image, not picture.image()",
                      DeprecationWarning, stacklevel=2)
        return bytes(self)

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_WP = "{%s}" % WP
_A = "{%s}" % A
_PIC = "{%s}" % PIC
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
EMU_PER_POINT = 12700
#: Word's density for a picture that states none (pixels per inch).
DEFAULT_DPI = 96.0

#: Formats by signature: (extension, content type).
_FORMATS = (
    (b"\x89PNG\r\n\x1a\n", "png", "image/png"),
    (b"\xff\xd8", "jpeg", "image/jpeg"),
    (b"GIF87a", "gif", "image/gif"),
    (b"GIF89a", "gif", "image/gif"),
    (b"BM", "bmp", "image/bmp"),
)


class PictureError(ValueError):
    pass


def image_format(data: bytes) -> tuple[str, str]:
    for signature, extension, content_type in _FORMATS:
        if data.startswith(signature):
            return extension, content_type
    raise PictureError("not a PNG, JPEG, GIF or BMP picture")


def natural_size(data: bytes) -> tuple[float, float]:
    """A picture's natural size in points: its pixels at its stated density, else 96 dpi."""
    from ooxml_common.imagemeta import _pixels_and_density

    found = _pixels_and_density(data)
    if not found or not found[0] or not found[1]:
        raise PictureError("the picture's size cannot be read")
    width, height, density_x, density_y = found
    if not _states_density(data):
        density_x = density_y = DEFAULT_DPI
    # A density stated in pixels per metre, as Word takes it: whole dots per inch (a
    # 300 dpi PNG states 11,811 per metre, 299.9994 dpi, and Word draws it at 300).
    density_x, density_y = round(density_x) or DEFAULT_DPI, round(density_y) or DEFAULT_DPI
    return width * 72.0 / density_x, height * 72.0 / density_y


def _states_density(data: bytes) -> bool:
    """Whether the file states a density (PNG ``pHYs`` in metres, JFIF with units, BMP)."""
    if data.startswith(b"\x89PNG"):
        position = 8
        while position + 8 <= len(data):
            length = int.from_bytes(data[position:position + 4], "big")
            kind = data[position + 4:position + 8]
            if kind == b"pHYs":
                return data[position + 16] == 1 if position + 17 <= len(data) else False
            if kind == b"IDAT":
                return False
            position += 12 + length
        return False
    if data.startswith(b"\xff\xd8"):
        return len(data) > 13 and data[6:11] == b"JFIF\x00" and data[13] in (1, 2)
    if data.startswith(b"BM"):
        return True
    return False


def _read(image: "bytes | str | os.PathLike[str]") -> bytes:
    if isinstance(image, (bytes, bytearray)):
        return bytes(image)
    with open(os.fspath(image), "rb") as handle:
        return handle.read()


class Picture:
    """A drawing that shows a picture (``pic:pic``), by its ``wp:docPr`` id."""

    def __init__(self, document: "Document", identifier: str) -> None:
        self.document = document
        self.id = identifier

    def __eq__(self, other: object) -> bool:
        return isinstance(other, Picture) and other.id == self.id and other.document is self.document

    def __hash__(self) -> int:
        return hash(self.id)

    def _locate(self) -> tuple[str, Element]:
        for part, identifier, frame in self.document._drawings():
            if identifier == self.id:
                return part, frame
        raise KeyError(f"no picture {self.id}")

    @property
    def name(self) -> str | None:
        """The picture's name (``wp:docPr/@name``)."""
        return self._locate()[1].find(_WP + "docPr").get("name")

    @property
    def alt_text(self) -> str | None:
        """The picture's alternative text (``wp:docPr/@descr``), or ``None``; settable."""
        return self._locate()[1].find(_WP + "docPr").get("descr")

    @alt_text.setter
    def alt_text(self, value: str | None) -> None:
        self.document._picture_edit(self, alt_text=value)

    @property
    def size(self) -> tuple[float, float]:
        """``(width, height)`` in points, from ``wp:extent``."""
        extent = self._locate()[1].find(_WP + "extent")
        return int(extent.get("cx")) / EMU_PER_POINT, int(extent.get("cy")) / EMU_PER_POINT

    @property
    def inline(self) -> bool:
        """Whether the picture stands in the line rather than floating."""
        return self._locate()[1].tag == _WP + "inline"

    @property
    def media(self) -> str | None:
        """The part the picture shows."""
        part, frame = self._locate()
        blip = frame.find(f".//{_A}blip")
        return self.document.package.related_part(part, blip.get(_R + "embed")) if blip is not None else None

    @property
    def image(self) -> "ImageBytes | None":
        """The image's bytes (``picture.image``), or ``None`` for a linked image.  It was a
        method: ``picture.image()`` still gives the bytes, with a ``DeprecationWarning``."""
        media = self.media
        return ImageBytes(self.document.package.read(media)) if media else None

    def replace(self, image: "bytes | str | os.PathLike[str]", *, fit: str = "frame") -> "EditResult":
        """Show another picture in this frame.  ``fit="frame"`` keeps the frame's size;
        ``"width"`` keeps its width and takes the new picture's aspect ratio; ``"natural"``
        sizes it as inserting it would."""
        return self.document._replace_picture(self, _read(image), fit)

    def resize(self, width: float | None = None, height: float | None = None) -> "EditResult":
        """Points; one dimension alone keeps the aspect ratio."""
        return self.document._picture_edit(self, size=(width, height))

    def delete(self) -> "EditResult":
        """Delete the picture (tracked when tracking)."""
        return self.document._delete_picture(self)

    @property
    def position(self) -> dict:
        """A floating picture's position (where it is measured from and its offsets)."""
        from .drawings import read_position

        return read_position(self._locate()[1])

    @property
    def wrapping(self) -> dict:
        """A floating picture's wrapping (``wrap`` and its distances)."""
        from .drawings import read_wrap

        return read_wrap(self._locate()[1])

    def float(self, **values) -> "EditResult":
        """Float the picture (:meth:`Document.float_drawing`)."""
        return self.document.float_drawing(self.id, **values)

    def make_inline(self) -> "EditResult":
        """Put a floating picture in the line again: :meth:`Document.inline_drawing`."""
        return self.document.inline_drawing(self.id)

    def set(self, **values) -> "EditResult":
        """A floating picture's place, wrapping, order and size (:meth:`Document.set_drawing`)."""
        return self.document.set_drawing(self.id, **values)

    def __repr__(self) -> str:
        return f"<Picture {self.id} {self.name!r}>"


class PictureOps:
    """The picture half of :class:`docx_agent.Document`."""

    def _drawings(self: "Document") -> list[tuple[str, str, Element]]:
        """``(part, id, wp:inline | wp:anchor)`` for every picture, in document order."""
        out = []
        seen: dict[str, int] = {}
        for part in self._parts():
            root = self.package.tree(part)
            for frame in _ids.live_elements(root, frozenset({_ids.WP_INLINE, _ids.WP_ANCHOR})):
                if frame.find(f".//{_PIC}pic") is None:
                    continue
                doc_pr = frame.find(_WP + "docPr")
                raw = doc_pr.get("id") if doc_pr is not None else None
                if raw is None:
                    continue
                count = seen.get(raw, 0)
                seen[raw] = count + 1
                out.append((part, f"d:{raw}" + (f"#{count}" if count else ""), frame))
        return out

    def pictures(self: "Document") -> list[Picture]:
        """Every picture (``d:<id>``), in story and document order."""
        return [Picture(self, identifier) for _, identifier, _ in self._drawings()]

    def picture(self: "Document", identifier: str) -> Picture:
        """A picture by id (``d:3``); ``KeyError`` if none.  ``doc.picture("d:3").alt_text``."""
        found = Picture(self, identifier)
        found._locate()
        return found

    def _next_doc_pr(self: "Document") -> int:
        largest = 0
        for part in self._parts():
            for node in self.package.tree(part).iter(_WP + "docPr"):
                try:
                    largest = max(largest, int(node.get("id")))
                except (TypeError, ValueError):
                    pass
        return largest + 1

    def _media(self: "Document", part: str, data: bytes) -> str:
        """A relationship from ``part`` to a media part holding ``data``: an existing part
        with the same bytes, or a new one."""
        extension, content_type = image_format(data)
        existing = self.package.find_part_with_bytes(data, "word/media")
        if existing is None:
            target = self.package.unused_part_name(f"word/media/image{{n}}.{extension}")
            existing = self.package.add_part(target, data, content_type)
        return self.package.add_relationship(part, REL_IMAGE, existing)

    def insert_picture(self: "Document", at: "str | TextRange", image: "bytes | str | os.PathLike[str]", *,
                       width: float | None = None, height: float | None = None,
                       alt_text: str | None = None, name: str | None = None) -> "EditResult":
        """An inline picture at a position (a range id or range; its start), from bytes or a
        file.  ``width``/``height`` in points (natural size by default; one alone keeps the
        aspect ratio)."""
        from .document import EditError, EditResult

        data = _read(image)
        try:
            image_format(data)
            size = _fit(natural_size(data), width, height)
        except PictureError as error:
            raise EditError(str(error)) from None
        where = at if isinstance(at, TextRange) else self.range(at)
        if where.view != "current":
            raise EditError("edits take the current view")
        part, entries = where._entries()
        entry = entries[0]
        items = _text.atoms(entry.element)
        if not 0 <= where.start <= len(items):
            raise EditError(f"{where.id} is outside its paragraph's text")
        if 0 < where.start < len(items) and items[where.start].field is not None \
                and items[where.start - 1].field is items[where.start].field:
            raise EditError(f"{where.id} is inside a field's result")
        with self._edit():
            renames = self._prepare(part, [entry])
            rid = self._media(part, data)
            number = self._next_doc_pr()
            label = name or f"Picture {number}"
            used = self._used()
            anchor_id, edit_id = _ids.generate(part + "\0picture", used, 2)
            drawing = _drawing(rid, number, label, alt_text, size, anchor_id, edit_id)
            run = _inline.new_run(entry.element, where.start, [drawing])
            self._track_inserted(part, [run])
            _ids.ensure_w14(self.package.tree(part), drawings=True)
            self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
        return EditResult(f"d:{number}", created=[f"d:{number}"], renamed=renames)

    def _replace_picture(self: "Document", picture: Picture, data: bytes, fit: str) -> "EditResult":
        from .document import EditError, EditResult

        if fit not in ("frame", "width", "natural"):
            raise EditError("fit is 'frame', 'width' or 'natural'")
        try:
            image_format(data)
            natural = natural_size(data)
        except PictureError as error:
            raise EditError(str(error)) from None
        part, frame = picture._locate()
        blip = frame.find(f".//{_A}blip")
        if blip is None or blip.get(_R + "embed") is None:
            raise EditError(f"{picture.id} shows no embedded picture")
        tracking = self._active_tracking()
        if tracking is not None and not _own_insertion(frame, tracking.author):
            return self._replace_picture_tracked(picture, data, fit, natural, tracking)
        with self._edit():
            renames = self._prepare_many([])
            part, frame = picture._locate()
            blip = frame.find(f".//{_A}blip")
            old = blip.get(_R + "embed")
            blip.set(_R + "embed", self._media(part, data))
            if fit != "frame":
                width = picture.size[0]
                size = natural if fit == "natural" else _fit(natural, width, None)
                _set_size(frame, size)
            if old != blip.get(_R + "embed"):
                self._release(part, [old])
            self.package.mark_dirty(part)
        return EditResult(picture.id, renamed=renames)

    def _track_inserted(self: "Document", part: str, runs: list[Element]) -> None:
        """When tracking: put runs an edit just placed inside a ``w:ins`` where they stand."""
        tracking = self._active_tracking()
        if tracking is None or not runs:
            return
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp

        stamp = Stamp(self, tracking)
        parent = runs[0].getparent()
        index = parent.index(runs[0])
        for run in runs:
            remove(run)
        _track.insert_nodes(parent, index, runs, stamp, part)
        stamp.finish()

    def _isolate_drawing(self: "Document", part: str, frame: Element) -> tuple[Element, Element]:
        """``(paragraph, run)``: the run holding only the frame's drawing (split out)."""
        paragraph = frame
        while paragraph is not None and paragraph.tag != _W + "p":
            paragraph = paragraph.getparent()
        atoms = _text.atoms(paragraph)
        inside = {id(node) for node in frame.iterancestors()}
        offset = next(k for k, atom in enumerate(atoms) if id(atom.node) in inside or atom.node is frame)
        runs = _inline.isolate(paragraph, offset, offset + 1)
        return paragraph, runs[0]

    def _replace_picture_tracked(self: "Document", picture: Picture, data: bytes, fit: str, natural, tracking) -> "EditResult":
        """A picture replaced under tracking: the old one deleted, a new one inserted after
        it (Word tracks a picture as text: measured, an inserted picture is a run in
        ``w:ins``)."""
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp
        from .document import EditResult

        with self._edit():
            renames = self._prepare_many([])
            part, frame = picture._locate()
            paragraph, run = self._isolate_drawing(part, frame)
            new_run = copy.deepcopy(run)
            new_frame = next(new_run.iter(_WP + "inline", _WP + "anchor"))
            blip = new_frame.find(f".//{_A}blip")
            blip.set(_R + "embed", self._media(part, data))
            number = self._next_doc_pr()
            for node in (new_frame.find(_WP + "docPr"), new_frame.find(f".//{_PIC}cNvPr")):
                if node is not None:
                    node.set("id", str(number))
            used = self._used()
            anchor_id, edit_id = _ids.generate(part + "\0picture", used, 2)
            new_frame.set(_ids.ANCHOR_ID, anchor_id)
            new_frame.set(_ids.EDIT_ID, edit_id)
            if fit != "frame":
                width = picture.size[0]
                _set_size(new_frame, natural if fit == "natural" else _fit(natural, width, None))
            stamp = Stamp(self, tracking)
            released = _track.delete_runs([run], stamp, part)
            holder = run.getparent()
            _track.insert_nodes(holder.getparent(), holder.getparent().index(holder) + 1, [new_run], stamp, part)
            stamp.finish()
            self._release(part, released)
            entry = self._index(part).entry_for(paragraph)
            if entry is not None:
                self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
        return EditResult(f"d:{number}", created=[f"d:{number}"], renamed=renames, removed=[picture.id])

    def _picture_edit(self: "Document", picture: Picture, *, alt_text=..., size=None) -> "EditResult":
        from .document import EditError, EditResult

        part, frame = picture._locate()
        if size is not None:
            width, height = size
            if width is None and height is None:
                raise EditError("give a width, a height or both")
            for value in (width, height):
                if value is not None and not (isinstance(value, (int, float)) and 0 < value <= 1584):
                    raise EditError("sizes are in points, up to 1584 (22 inches)")
        with self._edit():
            renames = self._prepare_many([])
            part, frame = picture._locate()
            if alt_text is not ...:
                for node in (frame.find(_WP + "docPr"), frame.find(f".//{_PIC}cNvPr")):
                    if node is None:
                        continue
                    if alt_text:
                        node.set("descr", alt_text)
                    else:
                        node.attrib.pop("descr", None)
            if size is not None:
                _set_size(frame, _fit(picture.size, size[0], size[1]))
            self.package.mark_dirty(part)
        return EditResult(picture.id, renamed=renames)

    def _delete_picture(self: "Document", picture: Picture) -> "EditResult":
        from .document import EditResult

        tracking = self._active_tracking()
        if tracking is not None:
            from ..revisions import track as _track
            from ..revisions.stamp import Stamp

            with self._edit():
                renames = self._prepare_many([])
                part, frame = picture._locate()
                paragraph, run = self._isolate_drawing(part, frame)
                stamp = Stamp(self, tracking)
                released = _track.delete_runs([run], stamp, part)
                stamp.finish()
                self._release(part, released)
                entry = self._index(part).entry_for(paragraph)
                if entry is not None:
                    self._renew_text_id(part, entry)
                self.package.mark_dirty(part)
            return EditResult(None, renamed=renames)
        with self._edit():
            renames = self._prepare_many([])
            part, frame = picture._locate()
            drawing = frame.getparent()
            run = drawing.getparent()
            while drawing is not None and drawing.tag != _W + "drawing":
                drawing = drawing.getparent()
            container = drawing.getparent()
            if container is not None and container.tag.endswith("}Choice"):
                drawing = container.getparent()  # the mc:AlternateContent holding it
            released = _inline.relationship_ids(drawing)
            run = drawing.getparent()
            paragraph = run
            while paragraph is not None and paragraph.tag != _W + "p":
                paragraph = paragraph.getparent()
            remove(drawing)
            if paragraph is not None:
                _inline.prune(paragraph)
                entry = self._index(part).entry_for(paragraph)
                if entry is not None:
                    self._renew_text_id(part, entry)
            self._release(part, released)
            self.package.mark_dirty(part)
        return EditResult(None, renamed=renames, removed=[picture.id])


def _own_insertion(frame: Element, author: str) -> bool:
    """Whether a picture is inside an insertion by ``author``: replacing or deleting it
    changes that insertion, not the document as it was."""
    for node in frame.iterancestors():
        if node.tag == _W + "ins":
            return node.get(_W + "author") == author
        if node.tag == _W + "p":
            return False
    return False


def _fit(natural: tuple[float, float], width: float | None, height: float | None) -> tuple[float, float]:
    if width is None and height is None:
        return natural
    if width is not None and height is not None:
        return width, height
    if width is not None:
        return width, natural[1] * width / natural[0]
    return natural[0] * height / natural[1], height


def _emu(points: float) -> str:
    return str(int(round(points * EMU_PER_POINT)))


def _set_size(frame: Element, size: tuple[float, float]) -> None:
    extent = frame.find(_WP + "extent")
    extent.set("cx", _emu(size[0]))
    extent.set("cy", _emu(size[1]))
    ext = frame.find(f".//{_PIC}spPr/{_A}xfrm/{_A}ext")
    if ext is not None:
        ext.set("cx", _emu(size[0]))
        ext.set("cy", _emu(size[1]))


def _drawing(rid: str, number: int, name: str, alt_text: str | None, size: tuple[float, float],
             anchor_id: str, edit_id: str) -> Element:
    from xml.sax.saxutils import quoteattr

    cx, cy = _emu(size[0]), _emu(size[1])
    descr = f" descr={quoteattr(alt_text)}" if alt_text else ""
    xml = (
        f'<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<wp:inline xmlns:wp="{WP}" xmlns:wp14="{_ids.WP14}" distT="0" distB="0" distL="0" distR="0" '
        f'wp14:anchorId="{anchor_id}" wp14:editId="{edit_id}">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{number}" name={quoteattr(name)}{descr}/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{A}" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}">'
        f'<pic:pic xmlns:pic="{PIC}"><pic:nvPicPr><pic:cNvPr id="{number}" name={quoteattr(name)}{descr}/>'
        f'<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr>'
        f'<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        f'</a:graphicData></a:graphic></wp:inline></w:drawing>'
    )
    return etree.fromstring(xml)
