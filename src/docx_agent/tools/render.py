"""Rendering pages: docx2svg behind the tool layer's render pipeline.

:data:`RENDERER` is the document's :class:`ooxml_edit.tools.render.Renderer`: the pipeline
(:func:`ooxml_edit.tools.render.draw`, its worker tasks, the cache keys, the PNG checks) is
ooxml-edit's, the same for every format; this module says what is a document's.  A document
is laid out whole -- text flows from one page into the next -- so one worker lays it out
once and writes the pages' SVGs (:meth:`PageRenderer.lay_out`); the pages whose image is not
held already are then rasterised over all the workers at once, in chunks
(:meth:`PageRenderer.rasterise`).  Both work from the bytes they are given (docx2svg opens
them itself); their PNGs are ``render_pages``' byte for byte (:mod:`.shared`: docx2svg's
``convert_docx_to_png``): the same docx2svg passes and the same rasteriser arguments.
"""

from __future__ import annotations

import re
from typing import Any

from ooxml_edit.tools import render, worker_task
from ooxml_edit.tools.render import Layout, RendererBase, View

#: The ``data-docx-*`` attributes of a page's SVG: where its elements come from (paths,
#: ids), which the rasteriser never reads -- left out of the key, so the ids an edit gives
#: every paragraph (``w14:paraId``) do not make every page new.
_DOCX_DATA = re.compile(r' data-docx-[a-z-]+="[^"]*"')


class PageRenderer(RendererBase):
    """docx2svg as the render pipeline calls it (:class:`ooxml_edit.tools.render.Renderer`):
    a view is a 1-based page number, and the document is laid out whole.

    The ``state`` a task is handed is a layout of these very bytes, font folders and locale
    that an earlier task returned as its facts' ``"laid_out"`` (pickled: the layout and the
    faces it was measured with, :func:`docx2svg.lay_out_docx`); given, the document is not
    laid out again -- the same pages, byte for byte, for the cost of unpickling it.
    Without it, the layout is made and returned so, for the caller to hand the next call.
    ``("previous", laid_out)`` is such a layout of an **earlier version** of the document
    (round 19): the document is laid out again from it -- from the page the first changed
    block is on until the pages line up with it again (``lay_out_docx``'s ``previous``) --
    the same pages, byte for byte, as a layout of the bytes alone; the pages an edit left
    as they were then key to the images held, and are not drawn again.
    The facts are also ``count`` (the pages there are), ``coverage`` (the layout's
    coverage facts, :func:`docx_agent.layout.coverage_facts`), and from an earlier
    version's layout ``changed`` (the pages whose drawing changed), ``relaid`` (why the
    whole document was laid out again, or ``None``) and ``relaid_cause`` (what caused it:
    docx2svg's ``Layout.relaid_cause``, element paths as docx2svg writes them; the tool maps
    them to block ids, :func:`docx_agent.layout.relaid_facts`)."""

    name = "docx-agent"
    whole_document = True
    #: The code that draws the pixels: a new install is a new key.
    modules = ("docx2svg", "ooxml_common", "resvg_py")

    def lay_out(self, data: bytes, items: list, width: int, font_dirs: "list[str]",
                locale: Any, state: Any) -> Layout:
        import os
        import pickle
        import tempfile

        import docx2svg
        from docx2svg.fonts import rasteriser_files
        from ooxml_common.fonts.rasterdb import file_stamps
        from ooxml_edit.tools import content_key

        from ..layout import coverage_facts, locale_of

        locale = locale_of(None) if locale is None else locale
        options = docx2svg.ConvertOptions(pages=items, width=width, font_dirs=font_dirs,
                                          locale=locale)
        fresh = None
        previous = None
        if isinstance(state, tuple) and len(state) == 2 and state[0] == "previous":
            previous = _loaded(state[1])[0]
            state = None
        if state is not None:
            layout, fonts = _loaded(state)
        else:
            layout, _, fonts = docx2svg.lay_out_docx(data, options, previous=previous)
            fresh = pickle.dumps((_light(layout), fonts), protocol=pickle.HIGHEST_PROTOCOL)
            _remember(fresh, (layout, fonts))
        count = len(layout.pages)
        svgs = docx2svg.page_svgs(layout, data, options, device_size=True, fonts=fonts)
        numbers = [page for page in items if 1 <= page <= count]
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
        system = None if skip else render.system_fonts()
        plain = [path for path in files if isinstance(path, str)]
        fingerprint = render.fingerprint(self, skip, sans, file_stamps(plain, dirs), files,
                                         [(name, content_key(blob)) for name, blob in embedded],
                                         system)
        return Layout([View(page, svg) for page, svg in zip(numbers, svgs)], fingerprint,
                      raster={"files": files, "embedded": embedded, "dirs": dirs, "skip": skip,
                              "sans": sans},
                      facts={"count": count, "coverage": coverage_facts(layout.coverage),
                             "laid_out": fresh, "changed": getattr(layout, "changed_pages", None),
                             "relaid": getattr(layout, "relaid", None) if previous is not None else None,
                             "relaid_cause": getattr(layout, "relaid_cause", None) if previous is not None
                             else None})

    def key_text(self, svg: str) -> str:
        return _DOCX_DATA.sub("", svg)

    def rasterise(self, views: "list[View]", raster: Any) -> "list[tuple[bytes, Any]]":
        """Each page's PNG, with the layout's ``raster`` arguments (the embedded faces
        written to a temporary folder for as long as it takes)."""
        import os
        import tempfile

        from docx2svg.png import svg_to_png

        with tempfile.TemporaryDirectory(prefix="docx2svg-embedded-") as directory:
            written = {}
            for name, blob in raster["embedded"]:
                path = written[name] = os.path.join(directory, name)
                with open(path, "wb") as handle:
                    handle.write(blob)
            files = [written[item[1]] if isinstance(item, list) else item
                     for item in raster["files"]]
            return [(svg_to_png(view.svg, backend="auto", font_dirs=raster["dirs"],
                                font_files=files, skip_system_fonts=raster["skip"],
                                sans_serif_family=raster["sans"]), None)
                    for view in views]

    def rasterise_tiles(self, svg: str, font_dirs: "list[str] | None") -> bytes:
        """A contact image's SVG (the pages' PNGs at their own size, their labels) drawn
        once, at its own size, with the system's fonts where docx2svg takes them."""
        from docx2svg.png import svg_to_png

        return svg_to_png(svg, font_dirs=font_dirs)

    def warm_up(self) -> None:
        """What a worker's first layout and raster would pay for once (per process): docx2svg's
        imports and its scan of the font folders (:func:`docx2svg.warm`, for the folders
        ``OOXML_FONT_DIRS`` names), and the rasteriser's own load of the system's fonts: one
        line of text."""
        import docx2svg
        from docx2svg.png import svg_to_png

        docx2svg.warm()

        svg_to_png('<svg xmlns="http://www.w3.org/2000/svg" width="64" height="16">'
                   '<text x="0" y="12" font-family="sans-serif" font-size="10">Aa</text></svg>')


#: The layouts this process made or read last, by a digest of their pickle (``laid_out``):
#: the next call handing one back to the same process does not read it again, and finds
#: what the layout's flow measured, which the pickle leaves out (:func:`_light`).
_RECENT: dict = {}
_RECENT_SIZE = 2


def _light(layout):
    """``layout`` to pickle: without what its flow measured (``layout.state.memo``, which
    a layout of the next version reuses for the blocks that did not change -- most of a
    long document's pickle).  The process that made it keeps it (:data:`_RECENT`); a layout
    read in another is laid out again from without it: the same pages, measured again."""
    import dataclasses

    state = getattr(layout, "state", None)
    if state is None or not getattr(state, "memo", None):
        return layout
    return dataclasses.replace(layout, state=dataclasses.replace(state, memo={}))


def _remember(blob: bytes, laid: tuple) -> None:
    import hashlib

    key = hashlib.sha256(blob).digest()
    _RECENT.pop(key, None)
    _RECENT[key] = laid
    while len(_RECENT) > _RECENT_SIZE:
        _RECENT.pop(next(iter(_RECENT)))


def _loaded(blob: bytes) -> tuple:
    """``(layout, fonts)`` of a ``laid_out`` pickle: the one this process holds, else read."""
    import hashlib
    import pickle

    found = _RECENT.get(hashlib.sha256(blob).digest())
    if found is None:
        found = pickle.loads(blob)
        _remember(blob, found)
    return found


#: The document's renderer, handed to the pipeline's worker tasks.
RENDERER = PageRenderer()


@worker_task
def warm_up() -> None:
    """A new worker's first task, before any call's (:func:`ooxml_edit.tools.render.warm_up`):
    what the first render would pay for, and the stamp of the renderers' code."""
    render.warm_up(RENDERER)
