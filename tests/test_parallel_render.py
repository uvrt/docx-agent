"""Round 10, O1 and O3 for documents: the document laid out once, its pages rasterised over
the workers at once, and each image kept by its content, so a page an edit left as it was
is not drawn again.  The images must be what the serial renderer draws, byte for byte
(``render_pages``: docx2svg's ``convert_docx_to_png``)."""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")
pytest.importorskip("resvg_py")

from ooxml_edit.tools import InProcess, Limits, Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from docx_agent.tools.shared import lay_out_pages, render_pages  # noqa: E402


def _pages(count: int = 4) -> bytes:
    """``count`` pages whatever the fonts measure: each paragraph after the first starts a
    page, in Arial (every runner lays it out)."""
    from ooxml_edit.xml import qn

    document = Document.new()
    theme = document.package.tree("word/theme/theme1.xml")
    for latin in theme.iter(qn("a:latin")):
        latin.set("typeface", "Arial")
    document.package.mark_dirty("word/theme/theme1.xml")
    for number in range(count):
        made = document.append_paragraph(f"Page {number + 1} of a made-up note.")
        if number:
            document.format_paragraph(made.id, page_break_before=True)
    return document.to_bytes()


def _box(**options) -> Toolbox:
    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, font_dirs=[],
                   limits=options.pop("limits", Limits(image_budget=None)), **options)


def _ok(box, session, tool, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, result.error and result.error.to_json()
    return result


def test_pages_rasterised_over_the_workers_are_the_serial_renderers_byte_for_byte():
    data = _pages()
    expected, count, _ = render_pages(data, [1, 2, 3, 4], 300, [])
    with _box(workers=2) as box:
        session = box.session()
        d = session.open(data, "note.docx")
        result = _ok(box, session, "render", doc=d, pages=[1, 2, 3, 4], width=300)
    assert [image.data for image in result.images] == expected
    assert result.data["cached"] == [] and "coverage" in result.data


def test_a_repeat_is_answered_from_the_store_and_an_edit_redraws_only_its_page():
    data = _pages()
    with _box(workers=2) as box:
        session = box.session()
        d = session.open(data, "note.docx")
        first = _ok(box, session, "render", doc=d, pages=[1, 2, 3, 4], width=300)
        again = _ok(box, session, "render", doc=d, pages=[1, 2, 3, 4], width=300)
        assert again.data["cached"] == [1, 2, 3, 4]
        _ok(box, session, "replace_text", doc=d, find="Page 3 of", replace="Page three of",
            expect="one")
        after = _ok(box, session, "render", doc=d, pages=[1, 2, 3, 4], width=300)
        expected = render_pages(session.entry(d).document.to_bytes(), [1, 2, 3, 4], 300, [])[0]
    assert after.data["cached"] == [1, 2, 4]
    assert [image.data for image in after.images] == expected
    assert after.images[2].data != first.images[2].data


def test_a_page_past_the_end_is_not_found_and_a_sheet_reuses_the_tiles():
    data = _pages(3)
    with _box(workers=2) as box:
        session = box.session()
        d = session.open(data, "note.docx")
        missing = box.dispatch(session, "render", {"doc": d, "pages": [2, 9]})
        assert not missing.ok and missing.error.code == "not_found"
        _ok(box, session, "render", doc=d, pages=[1, 2], width=200)
        sheet = _ok(box, session, "render", doc=d, pages=[1, 2, 3],
                    sheet={"columns": 3, "width": 200})
    assert sheet.data["cached"] is False and sheet.data["cached_tiles"] == [1, 2]


class _Recording(InProcess):
    """Runs tasks in this process, and keeps their names."""

    def __init__(self) -> None:
        super().__init__(1)
        self.ran: list[str] = []

    def run(self, fn, *args, timeout, **kwargs):
        self.ran.append(fn.__name__)
        return super().run(fn, *args, timeout=timeout, **kwargs)


def test_refusals_cost_no_render():
    data = _pages(3)
    runner = _Recording()
    with _box(limits=Limits(image_budget=1), runner=runner) as box:
        session = box.session()
        d = session.open(data, "note.docx")
        over = box.dispatch(session, "render", {"doc": d, "pages": [1, 2]})
        assert not over.ok and over.error.code == "limit"
        # Thirty named pages of 300 px in four columns: too tall at any page this
        # document's sections declare -- said before anything is laid out.
        tall = box.dispatch(session, "render", {"doc": d, "pages": list(range(1, 31)),
                                                "sheet": {"columns": 4, "width": 300}})
        assert not tall.ok and tall.error.code == "limit" and "pages" in tall.error.message
        assert runner.ran == []


def test_lay_out_pages_leaves_out_the_svg_of_an_image_held():
    data = _pages(2)
    first = lay_out_pages(data, [1, 2], 300, [])
    assert first["count"] == 2 and all(svg for _, _, svg in first["pages"])
    key = first["pages"][0][1]
    again = lay_out_pages(data, [1, 2], 300, [], known=[key])
    assert again["pages"][0] == (1, key, None)
    assert again["pages"][1][2] == first["pages"][1][2]
