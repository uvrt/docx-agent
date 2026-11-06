"""Round 35, R26: how much of its pages a range takes, in ``check(range=...)``.

Per range, on each page it is drawn on: the top and bottom of its body content in points (what
``DocumentLayout.where`` measures, line tops to line tops plus pitch), the page's body height
(the page less its section's top and bottom margins), and -- only where a header or footer
reaches past its margin and pushes the body in, as docx2svg lays it out -- the body's area.
A block's own extent: the block named as its own range (rule 1: the address forms ``range``
takes already), or ``DocumentLayout.extent`` in the Python API.  Made-up documents only.
"""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.layout import PX_PER_PT  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS, scoped  # noqa: E402

pytestmark = pytest.mark.any_faces

WORDS = "lantern meadow river stone harbour quiet orchard copper valley signal".split()


def _text(seed: int, words: int = 70) -> str:
    return " ".join(WORDS[(seed + k) % len(WORDS)] for k in range(words)).capitalize() + "."


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _check(box, session, **arguments):
    result = box.dispatch(session, "check", arguments)
    assert result.ok, result.error and result.error.to_json()
    return result


def _record() -> tuple[Document, list[str]]:
    """A heading, an opening paragraph, a table and an answer paragraph, after enough text
    that the table starts near a page's foot and goes on to the next page."""
    doc = Document.new()
    parts = [_text(k) for k in range(8)] + ["## A record", _text(3, 40),
                                             "| Item | Owner |\n|---|---|\n"
                                             + "\n".join(f"| Item {k} | {WORDS[k % 10]} |" for k in range(40)),
                                             _text(5, 50)]
    doc.insert_markdown("\n\n".join(parts))
    blocks = [b["id"] for b in doc.blocks()][-4:]
    return doc, blocks


def test_a_range_reports_its_extent_on_each_page(box):
    doc, record = _record()
    session = box.session()
    d1 = session.open(doc.to_bytes(), "record.docx")
    span = f"{record[0]}..{record[-1]}"
    reported = _check(box, session, doc=d1, range=[span], include=["layout"]).data["range"][0]
    document = session.entry(d1).document
    layout = document.layout()
    assert len(reported["pages"]) == 2, reported
    first, last = layout.where(record[0])[0], layout.where(record[-1])[-1]
    extent = reported["extent"]
    assert [e["page"] for e in extent] == reported["pages"]
    assert extent[0]["top"] == first.top and extent[-1]["bottom"] == last.bottom
    # The page's body: its height less the section's top and bottom margins.
    margins = document.sections()[0].margins
    body = round(document.sections()[0].page_height - margins["top"] - margins["bottom"], 2)
    assert all(e["body"] == body for e in extent)
    assert all("area" not in e and "columns" not in e for e in extent)    # nothing pushed it
    assert reported["height"] == round(sum(e["bottom"] - e["top"] for e in extent), 2)
    assert "extent_more" not in reported


def test_each_block_named_as_its_own_range_has_its_own_extent(box):
    """Per block: the blocks given as ranges beside the span (rule 1: no new field)."""
    doc, record = _record()
    session = box.session()
    d1 = session.open(doc.to_bytes(), "record.docx")
    ranges = [f"{record[0]}..{record[-1]}", *record]
    reported = _check(box, session, doc=d1, range=ranges).data["range"]
    layout = session.entry(d1).document.layout()
    for block, entry in zip(record, reported[1:]):
        placements = [p for p in layout.where(block) if p.story == "body"]
        assert [(e["page"], e["top"], e["bottom"]) for e in entry["extent"]] == \
            [(p.page, p.top, p.bottom) for p in placements]
    table = reported[3]
    assert table["blocks"] == 1 and len(table["extent"]) == 2                # the table over two pages


def test_a_header_taller_than_its_margin_pushes_the_body_and_says_so(box):
    doc, record = _record()
    section = doc.sections()[0]
    doc.add_header(section.id, "default", "Header line")
    last = section.header().paragraphs[0].id
    for k in range(6):
        last = doc.insert_paragraph(f"Header line {k}", after=last).id
    session = box.session()
    d1 = session.open(doc.to_bytes(), "record.docx")
    reported = _check(box, session, doc=d1, range=[record[1]]).data["range"][0]
    raw = session.entry(d1).document.layout().pages
    [entry] = reported["extent"]
    page = raw[entry["page"] - 1]
    # docx2svg's own layout: the body starts below the header, not at the margin.
    top, bottom = (round(float(page.column[k] / PX_PER_PT), 2) for k in (2, 3))
    assert entry["area"] == [top, bottom]
    assert top > section.margins["top"] + 1
    assert bottom - top < entry["body"]
    assert entry["top"] >= top


def test_each_page_has_its_own_sections_body():
    """A landscape section: its pages' body is its page height less its margins."""
    doc = Document.new()
    doc.insert_markdown("\n\n".join(_text(k) for k in range(4)))
    paragraphs = doc.paragraphs()
    doc.insert_section_break(after=paragraphs[1].id, kind="nextPage")
    first = doc.sections()[0]
    doc.set_section(first.id, orientation="landscape")
    layout = doc.layout()
    sections = doc.sections()
    before = layout.extent(paragraphs[0].id)
    after = layout.extent(paragraphs[3].id)
    for found, section in ((before, sections[0]), (after, sections[1])):
        margins = section.margins
        assert found[0].body == round(section.page_height - margins["top"] - margins["bottom"], 2)
    assert before[0].body != after[0].body


def test_columns_are_counted_and_the_extent_spans_them():
    doc = Document.new()
    doc.insert_markdown("\n\n".join(_text(k, 90) for k in range(3)))
    doc.set_section(doc.sections()[0].id, columns=2)
    layout = doc.layout()
    ids = [p.id for p in doc.paragraphs()]
    [found] = layout.extent(ids)
    assert found.columns == 2
    placements = [p for i in ids for p in layout.where(i)]
    assert found.top == min(p.top for p in placements) and found.bottom == max(p.bottom for p in placements)


def test_what_is_not_drawn_has_no_extent():
    doc = Document.new()
    doc.insert_markdown("One.\n\nTwo.")
    layout = doc.layout()
    assert layout.extent("p:00000000") == []
    assert layout.extent(doc.paragraphs()[0].id)[0].lines == 1


def test_a_long_range_lists_its_first_pages_and_counts_the_rest(box, monkeypatch):
    """Rule 4: the extent's pages are capped; the rest are counted, never cut silently."""
    doc, record = _record()
    monkeypatch.setattr(scoped, "EXTENT_PAGES", 1)
    session = box.session()
    d1 = session.open(doc.to_bytes(), "record.docx")
    reported = _check(box, session, doc=d1, range=[f"{record[0]}..{record[-1]}"]).data["range"][0]
    assert len(reported["extent"]) == 1 and reported["extent_more"] == 1
    layout = session.entry(d1).document.layout()
    assert reported["height"] == round(sum(e.height for e in layout.extent(
        session.entry(d1).document.blocks_between(record[0], record[-1]))), 2)


def test_the_top_level_key_table_resolves_as_docx2svg_does():
    """``blocks_on``'s body keys from one listing of the body: the same elements docx2svg's
    resolver finds."""
    from docx2svg.paths import resolve_path

    doc, _ = _record()
    layout = doc.layout()
    tree = doc.package.tree(doc.package.document_part())
    for page in layout.pages:
        for key in page.blocks:
            found = layout._top_level(key)
            if found is not False:
                assert found is resolve_path(tree, key), key
