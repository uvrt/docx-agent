#!/usr/bin/env python3
"""Write the sample documents docx-agent installs, ``src/docx_agent/samples/*.docx``.

``accept_loses_text.docx`` (round 55, R119)
    The shape check's ``accept_loses_text`` names, which docx-agent's own edits never make
    (a tracked ``set_cells`` keeps a cell's last paragraph mark): two tables of two rows and
    two columns, each with one cell whose last paragraph's mark is a pending tracked deletion
    by "Editor B" --

    * table 1, row 1's last cell (``c0,1``): "Kept." then "Last.", whose mark is deleted --
      a row's last cell;
    * table 2, the table's last cell (``c1,1``): the same two paragraphs -- the table's last
      cell.

    Word's Accept All takes each of those cells, with its text, out of view (measured
    round 55: the saved file keeps the text, after the row); docx-agent's ``word_changes``
    accept moves the text into paragraphs after the row and splits the table there; reject
    keeps it in the cell.

The recipe, for a document of one's own: in the cell's last ``w:p``, put
``<w:pPr><w:rPr><w:del w:id=".." w:author=".." w:date=".."/></w:rPr></w:pPr>`` -- a tracked
deletion of that paragraph's mark -- in a row's last cell.

    python tools/make_samples.py
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from docx_agent import Document

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_W = "{%s}" % W
OUT = Path(__file__).resolve().parents[1] / "src" / "docx_agent" / "samples"
AUTHOR = "Editor B"
DATE = "2026-11-13T09:00:00Z"


_COUNT = [0]


def _p(text: str, deleted_mark: int | None = None) -> str:
    # Paragraph ids as Word writes them (8 hex digits below 0x80000000), so ids stay put.
    _COUNT[0] += 1
    ids = f'w14:paraId="1A55{_COUNT[0]:04X}" w14:textId="2B66{_COUNT[0]:04X}"'
    mark = (f'<w:pPr><w:rPr><w:del w:id="{deleted_mark}" w:author="{AUTHOR}" w:date="{DATE}"/></w:rPr></w:pPr>'
            if deleted_mark is not None else "")
    return f'<w:p {ids}>{mark}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


# Borders on the table itself (the blank document defines no table style to take them from).
_BORDERS = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                   for side in ("top", "left", "bottom", "right", "insideH", "insideV"))


def _table(at: tuple[int, int], record: int) -> str:
    cells = [[_p(f"r{r + 1}c{c + 1}")] for r in range(2) for c in range(2)]
    cells[at[0] * 2 + at[1]] = [_p("Kept."), _p("Last.", deleted_mark=record)]
    rows = "".join(f'<w:tr w14:paraId="3C77{record:03X}{r}" w14:textId="4D88{record:03X}{r}">' + "".join('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
                                      + "".join(cells[r * 2 + c]) + "</w:tc>" for c in range(2)) + "</w:tr>"
                   for r in range(2))
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' + _BORDERS + '</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' + rows + "</w:tbl>")


def accept_loses_text() -> bytes:
    _COUNT[0] = 0
    body = (_p("A row's last cell whose last paragraph mark is a pending tracked deletion:")
            + _table((0, 1), 901)
            + _p("The table's last cell, the same:")
            + _table((1, 1), 902)
            + _p("End."))
    document = Document.new()
    part = document.package.document_part()
    container = document.package.tree(part).find(_W + "body")
    section = container.find(_W + "sectPr")
    for child in list(container):
        if child is not section:
            container.remove(child)
    for child in list(etree.fromstring(f'<w:body xmlns:w="{W}" xmlns:w14="{W14}">{body}</w:body>')):
        section.addprevious(child)
    document.package.mark_dirty(part)
    return document.to_bytes()


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "accept_loses_text.docx").write_bytes(accept_loses_text())
    print("wrote", OUT / "accept_loses_text.docx")
