"""Round 39: ``render``'s ``fast`` -- the same layout, the pages' text drawn as outlines in one
raster pass (about a third of an exact page's CPU), the result saying which images are
approximate -- and the render lane under it: an exact render is still the serial renderer's
image byte for byte, and a fast render never returns a worse image than one held."""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")
pytest.importorskip("resvg_py")

from ooxml_edit.tools import InProcess, Limits, Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from docx_agent.tools.render import RENDERER  # noqa: E402
from docx_agent.tools.shared import render_pages  # noqa: E402


def _pages(count: int = 3) -> bytes:
    """``count`` pages, a heading on each, whatever the fonts measure: each paragraph after
    the first starts a page, in Arial (every runner lays it out)."""
    from ooxml_edit.xml import qn

    document = Document.new()
    theme = document.package.tree("word/theme/theme1.xml")
    for latin in theme.iter(qn("a:latin")):
        latin.set("typeface", "Arial")
    document.package.mark_dirty("word/theme/theme1.xml")
    for number in range(count):
        made = document.append_paragraph(f"Part {number + 1}: a made-up note on lanterns and meadows.")
        if number:
            document.format_paragraph(made.id, page_break_before=True)
    return document.to_bytes()


def _box(**options) -> Toolbox:
    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, font_dirs=[],
                   limits=Limits(image_budget=None), **options)


def _ok(box, session, tool, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, result.error and result.error.to_json()
    return result


@pytest.mark.parametrize("workers", [0, 2], ids=["in-process", "workers"])
def test_a_fast_render_says_it_is_approximate_and_an_exact_one_is_unchanged(workers):
    data = _pages()
    expected, count, _ = render_pages(data, [1, 2, 3], 300, [])
    with _box(workers=workers) as box:
        session = box.session()
        d = session.open(data, "note.docx")
        fast = _ok(box, session, "render", doc=d, pages=[1, 2], width=300, fast=True)
        assert fast.data["approximate"] == ["pixels"] and fast.data["approximate_pages"] == [1, 2]
        assert [image.label for image in fast.images] == ["page 1 (approximate)", "page 2 (approximate)"]
        assert "fast: page(s) 1, 2 approximate" in fast.summary and "without fast" in fast.summary
        # The same layout: the same size of image, the same pages.
        assert [(i.width, i.height) for i in fast.images] == [(i.width, i.height) for i in
                                                             _ok(box, session, "render", doc=d, pages=[1, 2],
                                                                 width=300).images]
        exact = _ok(box, session, "render", doc=d, pages=[1, 2, 3], width=300)
        assert [image.data for image in exact.images] == expected            # byte for byte
        assert "approximate" not in exact.data and exact.data["cached"] == [1, 2]
        # Exact images held: a fast render gets them, and says nothing is approximate.
        again = _ok(box, session, "render", doc=d, pages=[2, 3], width=300, fast=True)
        assert [image.data for image in again.images] == expected[1:]
        assert "approximate" not in again.data and again.data["cached"] == [2, 3]


def test_fast_is_cheaper_pixels_of_the_same_svg():
    """The renderer's fast pixels: the page's SVG with its text outlined, one raster pass."""
    assert RENDERER.fast is True
    data = _pages(1)
    with _box(runner=InProcess(1)) as box:
        session = box.session()
        d = session.open(data, "note.docx")
        fast = _ok(box, session, "render", doc=d, pages=[1], width=300, fast=True)
        sheet = _ok(box, session, "render", doc=d, pages=[1], sheet={"width": 200}, fast=True)
        assert sheet.data["approximate"] == ["pixels"] and sheet.data["approximate_pages"] == [1]
    png = fast.images[0].data
    assert png[:8] == b"\x89PNG\r\n\x1a\n"
