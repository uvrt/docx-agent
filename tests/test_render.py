"""Rendering and reflow feedback through docx2svg, and the identity contract with it.

The loop only works if the ids in the picture are the ids the API answers to, and if
``where``/``page_of``/``compare`` describe the pages docx2svg draws.  These tests pin both,
since they span two repositories and would otherwise drift silently.
"""

from __future__ import annotations

import re

import pytest
from lxml import etree

from docx_agent import Document
from docx_agent.layout import Unknown

from test_edit import edit_set

_GROUPS = re.compile(r"<g data-docx-path=\"(?![^\"]*separator\")[^\"]*\"(?: data-docx-id=\"([^\"]*)\")?"
                     r"(?: data-docx-story=\"[^\"]*\" data-docx-part=\"[^\"]*\")?"
                     r"(?: data-docx-agent-id=\"([^\"]*)\")?>\s*<g data-docx-line=")


def groups(svg: str) -> list[tuple[str | None, str | None]]:
    """(data-docx-id, data-docx-agent-id) of every paragraph group (a note separator,
    which docx2svg draws as a line group in some documents, is no paragraph of the API's)."""
    return [(m.group(1), m.group(2)) for m in _GROUPS.finditer(svg)]


def check_rendered_ids(document: Document) -> None:
    svgs = document.render_svg()
    layout = document.layout()
    assert len(svgs) == layout.page_count
    for number, svg in enumerate(svgs, start=1):
        found = groups(svg)
        for raw, agent in found:
            assert agent is not None, f"page {number}: a paragraph group without an id"
            paragraph = document.paragraph(agent)
            if raw is not None:
                assert paragraph.para_id == raw
        assert {agent for _, agent in found} == set(layout.ids_on_page(number))


def test_every_drawn_paragraph_carries_its_id(docx_path):
    check_rendered_ids(Document.open(docx_path))


def test_ids_are_carried_after_edits(docx_path):
    document = Document.open(docx_path)
    before = document.layout()
    edit_set(document)
    check_rendered_ids(document)
    after = document.layout()
    assert after.unmapped == []
    # An edit must not make docx2svg stop earlier, or warn more.
    assert {w.code for w in after.warnings} <= {w.code for w in before.warnings}
    assert (after.stopped is None) or (before.stopped is not None)


def test_where_agrees_with_the_svg(markup_doc):
    document = Document.open(markup_doc)
    layout = document.layout()
    svgs = document.render_svg()
    for paragraph in document.paragraphs():
        where = layout.where(paragraph.id)
        pages = {n for n, svg in enumerate(svgs, start=1)
                 if f'data-docx-agent-id="{paragraph.id}"' in svg}
        if isinstance(where, Unknown):
            assert pages == set()
        else:
            assert {p.page for p in where} == pages
            for placement in where:
                assert 0 < placement.top < placement.bottom < 842  # A4, points


def test_where_in_points_matches_the_drawn_baselines(markup_doc):
    document = Document.open(markup_doc)
    placement = document.layout().where("p:1A2B3C01")[0]
    svg = document.render_svg(pages=[1])[0]
    match = re.search(r'data-docx-agent-id="p:1A2B3C01">\s*<g data-docx-line="\d+" '
                      r'data-docx-baseline="(\d+)"', svg)
    baseline_pt = int(match.group(1)) * 72 / 300
    assert placement.top < baseline_pt <= placement.bottom


def test_page_of_tables_stories_and_spans(markup_doc):
    layout = Document.open(markup_doc).layout()
    assert layout.page_of("t:3C4D5E01") == 1
    assert layout.page_of("p:1A2B3C17") == 2
    assert layout.page_of("header1/p:4D5E6F01") == (1, 2)       # drawn on every page
    assert layout.where("footnotes/p:5E6F7003")[0].story == "footnotes"
    assert isinstance(layout.where("p:3C4D5E05"), Unknown)      # a merged cell's continuation


def _svgs(document: Document) -> list[str]:
    return document.render_svg()


def _differing_pages(before: list[str], after: list[str]) -> list[int]:
    """Pages whose SVG differs, ignoring paths and ids (an insertion renumbers the paths
    after it, and an edit may stamp a paraId): an independent account of what changed."""
    def strip(svg: str) -> str:
        return re.sub(r' data-docx-(?:agent-)?(?:id|path)="[^"]*"', "", svg)

    count = max(len(before), len(after))
    return [n + 1 for n in range(count)
            if n >= len(before) or n >= len(after) or strip(before[n]) != strip(after[n])]


def test_compare_flags_exactly_the_page_a_one_word_edit_changed(long_doc):
    document = Document.open(long_doc)
    before_layout, before_svgs = document.layout(), _svgs(document)
    page = 9
    target = next(document.paragraph(i) for i in before_layout.ids_on_page(page)
                  if before_layout.page_of(i) == page and "Duis aute" in document.paragraph(i).text)
    target.set_text(target.text.replace("Duis aute", "Dues aute", 1))
    reflow = document.layout().compare(before_layout)
    assert reflow.changed == [page]
    assert reflow.changed == _differing_pages(before_svgs, _svgs(document))
    assert reflow.why == {page: "content"} and reflow.moved == {}
    assert reflow.page_count == (36, 36)


def test_compare_follows_a_reflow(markup_doc):
    path = markup_doc.parent.parent / "wordto" / "sample-10pages.docx"
    document = Document.open(path)
    before_layout, before_svgs = document.layout(), _svgs(document)
    second = before_layout.ids_on_page(1)[1]
    document.insert_paragraph("A new paragraph that pushes everything after it down. " * 6,
                              after=second)
    after = document.layout()
    reflow = after.compare(before_layout)
    assert reflow.changed == [1, 2, 3] == _differing_pages(before_svgs, _svgs(document))
    assert reflow.why == {1: "flow", 2: "flow", 3: "flow"}
    assert reflow.moved
    for identifier, (old, new) in reflow.moved.items():
        assert before_layout.where(identifier)[0].page == old
        assert after.where(identifier)[0].page == new


def test_compare_after_undo_is_empty(long_doc):
    document = Document.open(long_doc)
    before = document.layout()
    document.paragraphs()[5].set_text("Changed.")
    document.undo()
    assert document.layout().compare(before).changed == []


def test_a_stop_is_carried_through(markup_doc):
    """A frame other than a drop cap stops docx2svg's layout: what is past it is unknown."""
    document = Document.open(markup_doc)
    before = document.layout()
    element = document.paragraph("p:1A2B3C09")._element
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    properties = etree.Element(w + "pPr")
    element.insert(0, properties)
    etree.SubElement(properties, w + "framePr", {w + "w": "2000", w + "wrap": "around",
                                                 w + "hAnchor": "page", w + "x": "2000"})
    document.package.mark_dirty("word/document.xml")
    document._invalidate()
    layout = document.layout()
    assert layout.stopped.reason == "frame" and layout.stopped.at == "p:1A2B3C09"
    assert layout.pages_known is None
    assert layout.where("p:1A2B3C17") == Unknown("frame", 1)
    assert layout.where("p:1A2B3C02")[0].page == 1
    assert any(w.code == "layout-stopped:frame" for w in layout.warnings)
    reflow = layout.compare(before)
    assert reflow.unknown == [2] and reflow.changed == [1]


def test_png(markup_doc):
    pytest.importorskip("resvg_py")
    pngs = Document.open(markup_doc).render_png(pages=[2])
    assert len(pngs) == 1 and pngs[0].startswith(b"\x89PNG")


def test_a_render_and_a_layout_lay_the_document_out_once(markup_doc, monkeypatch):
    import docx2svg

    calls = []
    real = docx2svg.convert_docx
    monkeypatch.setattr(docx2svg, "convert_docx", lambda *a, **k: calls.append(1) or real(*a, **k))
    monkeypatch.setattr(docx2svg, "convert_docx_to_layout", None)
    monkeypatch.setattr(docx2svg, "convert_docx_to_svg", None)
    document = Document.open(markup_doc)
    document.render_svg(pages=[2])
    document.layout()
    document.render_svg()
    assert len(calls) == 1
    document.paragraph("p:1A2B3C02").set_text("Changed.")
    layout = document.layout()
    document.render_svg(pages=[1])
    assert len(calls) == 2
    assert layout.page_count == 2


def test_where_says_which_text_column(markup_doc):
    """A body of two text columns: a paragraph that breaks from the first column into the
    second is placed twice on its page, once per column."""
    document = Document.open(markup_doc)
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    body = document.package.tree("word/document.xml").find(w + "body")
    section = body.find(w + "sectPr")
    columns = section.find(w + "cols")
    if columns is None:
        columns = etree.SubElement(section, w + "cols")
        from docx_agent.oxml.xml import insert_in_order

        insert_in_order(section, columns)
    columns.set(w + "num", "2")
    columns.set(w + "space", "720")
    document.package.mark_dirty("word/document.xml")
    document._invalidate()
    long = document.paragraph("p:1A2B3C02")
    long.set_text("Words that fill a column and go on into the next one. " * 60)
    placements = document.layout().where(long.id)
    assert {p.column for p in placements} == {0, 1}
    assert len({(p.page, p.column) for p in placements}) == len(placements)
    first = document.layout().where("p:1A2B3C01")
    assert [p.column for p in first] == [0]


# -- after E1's edits -----------------------------------------------------------------------


def test_e1_edits_render_with_their_ids(docx_path):
    from e1_edits import e1_edit_set
    from docx_agent.validate import check

    document = Document.open(docx_path)
    before = document.layout()
    problems_before = set(check(document.package))
    expected = e1_edit_set(document)
    check_rendered_ids(document)
    after = document.layout()
    assert after.unmapped == []
    # The cross-reference the edits added is a field docx2svg draws as cached: that warning
    # is the edit's own obstacle; nothing else may be new.
    assert {w.code for w in after.warnings} - {"field-cached:REF"} <= {w.code for w in before.warnings}
    assert (after.stopped is None) or (before.stopped is not None)
    for key in ("heading", "big"):
        placed = after.where(expected.ids[key])
        assert not isinstance(placed, Unknown) or after.stopped is not None
    assert set(check(Document.open(document.to_bytes()).package)) - problems_before == set()


def test_reflow_of_e1_edits_is_sensible(markup_doc):
    """A style change on page 1 moves what follows; a same-length replacement on the last
    page changes that page alone."""
    path = markup_doc.parent.parent / "wordto" / "sample-10pages.docx"
    document = Document.open(path)
    first_layout = document.layout()
    target = first_layout.ids_on_page(1)[2]
    document.paragraph(target).style = "Title"
    after = document.layout()
    reflow = after.compare(first_layout)
    assert reflow.changed and reflow.changed[0] == 1
    assert reflow.why[1] in ("content", "flow")
    for identifier, (old, new) in reflow.moved.items():
        assert first_layout.where(identifier)[0].page == old and after.where(identifier)[0].page == new

    before = document.layout()
    page = before.page_count
    candidate = next(i for i in before.ids_on_page(page) if before.page_of(i) == page
                     and len(document.paragraph(i).text) > 40)
    word = next(w for w in document.paragraph(candidate).text.split() if len(w) >= 5 and w.isalpha())
    document.anchor(word, within=candidate, occurrence=0).replace(word[::-1])
    reflow = document.layout().compare(before)
    assert reflow.changed == [page] and reflow.why == {page: "content"} and reflow.moved == {}


def test_a_picture_pushes_text_down(markup_doc):
    from e1_edits import png

    document = Document.open(markup_doc)
    before = document.layout()
    top = before.where("p:1A2B3C09")[0].top
    document.insert_picture("p:1A2B3C02@0", png(96, 96), width=200)
    after = document.layout()
    assert after.where("p:1A2B3C09")[0].top > top + 100 or after.where("p:1A2B3C09")[0].page > 1


# -- the converter: laying out elsewhere ---------------------------------------------------------


def test_a_converter_lays_out_every_state_and_its_answer_is_the_same_as_here():
    from concurrent.futures import ProcessPoolExecutor
    import multiprocessing

    from docx_agent.layout import convert_bytes

    document = Document.new()
    document.insert_markdown("# One\n\nText.\n\n## Two\n\nMore text.")
    here = document.layout()
    seen = []
    with ProcessPoolExecutor(1, mp_context=multiprocessing.get_context("spawn")) as pool:
        def converter(data, options):
            seen.append(options)
            return pool.submit(convert_bytes, data, options).result()

        document.converter = converter
        document._layouts.clear()
        there = document.layout()
        assert [there.signature(n) for n in range(1, there.page_count + 1)] == \
            [here.signature(n) for n in range(1, here.page_count + 1)]
        toc = document.insert_toc(before=document.paragraphs()[0].id)
        assert "One" in document.field(toc.id).result and len(seen) >= 2
    document.converter = None
