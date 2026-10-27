"""Pages and blocks both ways, in this API's ids (round 19, R3): ``DocumentLayout.blocks_on``
and ``DocumentLayout.pages_of``, from docx2svg's own page <-> block map (``Page.blocks``).

Every document here is made up: the 36-page ``long_doc`` and one written in Markdown with
a paragraph longer than a page and a table of 70 rows.
"""

from __future__ import annotations

import pytest

from docx_agent import Document


@pytest.fixture(scope="module")
def crossing() -> Document:
    document = Document.new()
    long_text = " ".join(f"Sentence {k} of a paragraph that goes on past the foot of its page." for k in range(140))
    rows = "\n".join(f"| Row {k} | plain | cell {k} |" for k in range(1, 70))
    document.insert_markdown(f"# A long paragraph\n\n{long_text}\n\n| Name | Kind | Note |\n|---|---|---|\n{rows}\n\n"
                             "After the table.\n")
    return document


@pytest.fixture(scope="module")
def long_layout(long_doc):
    document = Document.open(long_doc)
    return document, document.layout()


def _body_pages(layout, identifier: str) -> list[int]:
    """The pages a paragraph's body lines are on, as ``ids_on_page`` sees them."""
    return [number for number in range(1, layout.page_count + 1) if identifier in layout.ids_on_page(number)]


def test_pages_of_a_paragraph_are_where_its_lines_are(long_layout):
    document, layout = long_layout
    paragraphs = [p.id for p in document.paragraphs()]
    assert layout.page_count > 30
    for identifier in paragraphs:
        pages = layout.pages_of(identifier)
        assert pages == _body_pages(layout, identifier), identifier
        if pages:
            first, last = pages[0], pages[-1]
            assert layout.page_of(identifier) == (first if first == last else (first, last))


def test_blocks_on_a_page_are_its_top_level_blocks_in_order(long_layout):
    document, layout = long_layout
    # The body's top-level blocks, in document order: paragraphs outside tables, and tables.
    in_tables = {p.id for t in document.tables() for row in range(len(t.rows))
                 for column in range(t.column_count) for p in _cell_paragraphs(t, row, column)}
    top_level = [p.id for p in document.paragraphs() if p.id not in in_tables]
    tables = {t.id for t in document.tables()}
    seen: list[str] = []
    for number in range(1, layout.page_count + 1):
        blocks = layout.blocks_on(number)
        assert blocks, number
        for identifier in blocks:
            assert number in layout.pages_of(identifier)
            assert identifier in tables or identifier in top_level, identifier
        # In document order, a block split across two pages last on one and first on the next.
        paragraphs = [b for b in blocks if b not in tables]
        assert paragraphs == sorted(paragraphs, key=top_level.index)
        seen.extend(b for b in blocks if b not in seen)
    drawn = [p for p in top_level if layout.pages_of(p)]
    assert [b for b in seen if b not in tables] == drawn
    assert layout.blocks_on(0) == [] and layout.blocks_on(layout.page_count + 1) == []


def _cell_paragraphs(table, row: int, column: int) -> list:
    try:
        return table.cell(row, column).paragraphs
    except KeyError:
        return []


def test_a_block_across_pages_is_on_each(crossing):
    layout = crossing.layout()
    long_paragraph = next(p for p in crossing.paragraphs() if p.text.startswith("Sentence 0"))
    assert len(layout.pages_of(long_paragraph.id)) >= 2
    table = crossing.tables()[0]
    pages = layout.pages_of(table.id)
    assert len(pages) >= 2
    for number in pages:
        assert table.id in layout.blocks_on(number)
    # A row is on its page; a cell's paragraph on its row's.
    row = table.rows[40]
    row_pages = layout.pages_of(row.id)
    assert len(row_pages) == 1 and row_pages[0] in pages
    assert layout.pages_of(table.cell(40, 0).paragraphs[0].id) == row_pages
    assert layout.pages_of(table.cell(40, 2).id) == row_pages
    assert layout.pages_of("p:0BADF00D") == []


def test_ids_follow_an_edit_through_their_aliases(crossing):
    document = Document.open(crossing.to_bytes())
    first = document.paragraphs()[0]
    before = document.layout().pages_of(first.id)
    document.paragraphs()[-1].set_text("After the table, edited.")   # stamps every paragraph
    assert document.layout().pages_of(first.id) == before
