"""What E5's edits write, held to what Word wrote for the same edits (``tools/e5_probe.py``,
``tests/observations/e5-word.json``): each probe document is rebuilt here, edited through
the API as the probe edited it in Word, and compared block by block.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pytest
from lxml import etree

from docx_agent import Document
from docx_agent.validate import check

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import e5_probe  # noqa: E402

OBSERVED = json.loads((ROOT / "tests" / "observations" / "e5-word.json").read_text(encoding="utf-8"))
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
#: Dutch Word's id for Table Grid, and English Word's.
STYLE_IDS = {"Tabelraster": "TableGrid"}


def plain(xml: str) -> str:
    xml = re.sub(r' xmlns(:\w+)?="[^"]*"', "", xml)
    xml = re.sub(r' (w14:paraId|w14:textId|w:rsid\w*|r:id|mc:Ignorable|wp14:editId)="[^"]*"', "", xml)
    xml = re.sub(r' w:(author|date)="[^"]*"| w16du:dateUtc="[^"]*"', "", xml)
    xml = re.sub(r' w:id="\d+"', "", xml)
    # Word drops xml:space where the text neither starts nor ends with a space.
    xml = re.sub(r'<w:t xml:space="preserve">([^<\s][^<]*[^<\s]|[^<\s])</w:t>', r"<w:t>\1</w:t>", xml)
    for dutch, english in STYLE_IDS.items():
        xml = xml.replace(f'w:val="{dutch}"', f'w:val="{english}"')
    return xml


def blocks(document: Document) -> list[str]:
    body = document.package.tree(document.package.document_part()).find(_W + "body")
    return [plain(etree.tostring(child, encoding="unicode")) for child in body]


def observed(probe: str) -> list[str]:
    return [plain(block) for block in OBSERVED[probe]["blocks"]]


def _tables(found: list[str]) -> dict[str, str]:
    """Each table block by the key its first cell names (``Head r1c1`` -> ``Head``)."""
    out = {}
    for block in found:
        if block.startswith("<w:tbl>"):
            match = re.search(r"<w:t>(\w+) r1c1", block)
            if match:
                out[match.group(1)] = block
    return out


# -- table properties --------------------------------------------------------------------------

TABLE_EDITS = {
    "Head": lambda d, t: d.set_row(t, 0, header=True),
    "Fixed": lambda d, t: d.set_table(t, layout="fixed"),
    "Pct": lambda d, t: d.set_table(t, width="80%"),
    "Center": lambda d, t: d.set_table(t, alignment="center"),
    "Right": lambda d, t: d.set_table(t, alignment="right"),
    "Indent": lambda d, t: d.set_table(t, indent=36),
    "Look": lambda d, t: d.set_table(t, look={"first_row": False, "first_column": False, "last_row": True,
                                              "last_column": True}),
    "Exact": lambda d, t: d.set_row(t, 0, height=30, height_rule="exact"),
    "AtLeast": lambda d, t: d.set_row(t, 1, height=40),
    "CantSplit": lambda d, t: d.set_row(t, 0, cant_split=True),
    "Shade": lambda d, t: (d.set_cell(t, 1, 1, shading={"fill": "00FFFF", "pattern": "solid"}),
                           d.set_cell(t, 0, 0, shading="FFFF00")),
    "Border": lambda d, t: d.set_cell(t, 1, 1, borders={"top": "double", "left": {"style": "single", "size": 2.25,
                                                                                  "color": "FF0000"}}),
    "Margin": lambda d, t: (d.set_cell(t, 1, 1, margins={"top": 5, "left": 12}), d.set_table(t, cell_margins={"left": 9})),
    "VAlign": lambda d, t: (d.set_cell(t, 1, 1, vertical_alignment="center"),
                            d.set_cell(t, 0, 0, vertical_alignment="bottom")),
    "TextDir": lambda d, t: (d.set_cell(t, 1, 1, text_direction="btLr"), d.set_cell(t, 0, 0, text_direction="tbRl")),
    "Widths": lambda d, t: d.set_column_width(t, 1, 150),
    "Float": lambda d, t: d.set_table(t, floating={"x": 100, "y": 20, "horizontal_anchor": "page",
                                                   "vertical_anchor": "text", "left": 9, "right": 12, "top": 3,
                                                   "bottom": 6, "overlap": False}),
}
#: What Word wrote beside the edit that is not the edit's: shading a cell gave the cell
#: above the one shaded solid a bottom border of the table's own (Table Grid's inside line).
WORD_ONLY = {"Shade": [(r'<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>',
                        "")]}


@pytest.fixture(scope="module")
def table_probe():
    return e5_probe.tables_document()


@pytest.mark.parametrize("key", list(TABLE_EDITS))
def test_table_properties_are_written_as_word_writes_them(table_probe, key, tracked=False):
    document = Document.open(table_probe)
    tables = {re.search(r"(\w+) r1c1", t.cell(0, 0).text).group(1): t.id for t in document.tables()}
    if tracked:
        with document.tracking(author="E5", date="2026-10-04T12:00:00Z"):
            TABLE_EDITS[key](document, tables[key])
    else:
        TABLE_EDITS[key](document, tables[key])
    ours = _tables(blocks(document))[key]
    word = _tables(observed("tables-tracked" if tracked else "tables"))[key]
    for pattern, replacement in WORD_ONLY.get(key, []):
        word = word.replace(pattern, replacement, 1)
    assert ours == word
    assert check(document.package) == []


# -- merges and splits -----------------------------------------------------------------------

MERGE_EDITS = {
    "MergeAcross": lambda d, t: d.merge_cells(t, (0, 0), (0, 1)),
    "MergeDown": lambda d, t: d.merge_cells(t, (0, 0), (1, 0)),
    "MergeRect": lambda d, t: d.merge_cells(t, (0, 0), (1, 1)),
    "SplitCols": lambda d, t: d.split_cell(t, 1, 1, rows=1, columns=2),
    "Unspan": lambda d, t: d.split_cell(t, 0, 0, rows=1, columns=2),
    "Unmerge": lambda d, t: d.split_cell(t, 0, 0, rows=3, columns=1),
    "ColRightOfSpan": lambda d, t: d.insert_column(t, 0, right=True),
    "RowInMerge": lambda d, t: d.insert_row(t, 1, below=True),
    "RowDelMerge": lambda d, t: d.delete_row(t, 1),
    "RowDelFirst": lambda d, t: d.delete_row(t, 0),
    "RowInSpan": lambda d, t: d.insert_row(t, 0, below=True),
    "ColDelPlain": lambda d, t: d.delete_column(t, 1),
}
#: Where Word's form and ours part, and why (each recorded in ROADMAP.md): splitting a cell
#: down, Word also gives both rows an at-least height of half the row's (130 twips here),
#: from its own layout; docx-agent writes no height.
NOT_WORDS = {
    "SplitRows": "Word gives the split rows an at-least height from its layout",
    "SplitBoth": "Word also widens the grid by 33 twips and gives the rows a height from its layout",
}


@pytest.fixture(scope="module")
def merge_probe():
    return e5_probe.merges_document()


@pytest.mark.parametrize("key", list(MERGE_EDITS))
def test_merges_and_splits_are_written_as_word_writes_them(merge_probe, key):
    document = Document.open(merge_probe)
    tables = {re.search(r"(\w+) r1c1", "".join(t._entry()[1].element.itertext())).group(1): t.id
              for t in document.tables()}
    MERGE_EDITS[key](document, tables[key])
    ours = _tables(blocks(document)).get(key)
    word = _tables(observed("merges")).get(key)
    if key == "RowDelFirst":  # the first row's text went with it: find the table by its second
        ours = next(b for b in blocks(document) if "RowDelFirst r2c2" in b)
        word = next(b for b in observed("merges") if "RowDelFirst r2c2" in b)
    assert ours == word
    assert check(document.package) == []


def test_split_rows_is_words_but_for_the_heights_word_takes_from_its_layout(merge_probe):
    document = Document.open(merge_probe)
    tables = {re.search(r"(\w+) r1c1", "".join(t._entry()[1].element.itertext())).group(1): t.id
              for t in document.tables()}
    for key, rows, columns in (("SplitRows", 2, 1), ("SplitBoth", 2, 2)):
        document.split_cell(tables[key], 1, 1, rows=rows, columns=columns)
        ours = _tables(blocks(document))[key]
        word = _tables(observed("merges"))[key].replace('<w:trPr><w:trHeight w:val="130"/></w:trPr>', "")
        if key == "SplitBoth":
            word = word.replace('<w:gridCol w:w="1033"/>', '<w:gridCol w:w="1000"/>')
        assert ours == word, key


# -- drawings ----------------------------------------------------------------------------------

DRAWING_EDITS = {
    "Float": {},
    "Square": {"wrap": "square"},
    "Tight": {"wrap": "tight"},
    "Through": {"wrap": "through"},
    "TopBottom": {"wrap": "top_and_bottom"},
    "Front": {"wrap": "front"},
    "Behind": {"wrap": "behind"},
    "Sides": {"wrap": "square", "side": "left", "distances": {"top": 3, "bottom": 6, "left": 9, "right": 12}},
    "Page": {"horizontal_from": "page", "x": 100, "vertical_from": "page", "y": 200},
    "Aligned": {"horizontal_from": "margin", "x_align": "center", "vertical_from": "margin", "y_align": "bottom"},
    "Column": {"horizontal_from": "column", "x": 20, "vertical_from": "paragraph", "y": 10},
    "Character": {"horizontal_from": "character", "x": 5, "vertical_from": "line", "y": 2},
    "Inside": {"horizontal_from": "page", "x_align": "inside", "vertical_from": "topMargin", "y_align": "top"},
    "Front2": {"wrap": "front", "z_order": "front"},
    "Back2": {"wrap": "front", "z_order": "back"},
    "Lock": {"lock_anchor": True},
    "Overlap": {"allow_overlap": False},
    "Size": {"width": 144, "height": 72},
}
#: Where Word's anchor holds what the edit did not set: the offsets where its layout put
#: the picture (docx-agent floats it at 0, 0 from the column and the paragraph), and, on
#: the tight case alone, behindDoc="1" (Word's quirk: its through case, the same steps, 0).
_LAYOUT = re.compile(r"<wp:posOffset>-?\d+</wp:posOffset>")


def _anchor(block: str) -> str:
    found = re.search(r"<wp:(anchor|inline) .*?</wp:(anchor|inline)>", block)
    text = re.sub(r"<a:graphic>.*</a:graphic>", "<a:graphic/>", found.group(0))  # the picture: the probe's
    return re.sub(r' relativeHeight="\d+"', "", re.sub(r' wp14:anchorId="[^"]*"', "", text))


@pytest.fixture(scope="module")
def drawing_probe():
    return e5_probe.drawings_document()


@pytest.mark.parametrize("key", list(DRAWING_EDITS) + ["Inline"])
def test_drawings_are_floated_as_word_floats_them(drawing_probe, key):
    document = Document.open(drawing_probe)
    pictures = {document.picture(p.id).alt_text.split()[-1]: p.id for p in document.pictures()}
    if key == "Inline":
        document.float_drawing(pictures[key])
        document.inline_drawing(pictures[key])
    else:
        document.float_drawing(pictures[key], **DRAWING_EDITS[key])
    ours = _anchor(next(b for b in blocks(document) if f'descr="Case {key}"' in b))
    word = _anchor(next(b for b in observed("drawings") if f'descr="Case {key}"' in b))
    if key in ("Float", "Square", "Tight", "Through", "TopBottom", "Front", "Behind", "Sides", "Front2", "Back2",
               "Lock", "Overlap", "Size"):
        ours, word = _LAYOUT.sub("<wp:posOffset/>", ours), _LAYOUT.sub("<wp:posOffset/>", word)
    if key == "Tight":
        word = word.replace('behindDoc="1"', 'behindDoc="0"')
    assert ours == word
    assert check(document.package) == []


def test_order_is_words(drawing_probe):
    """To the front: a step above the highest; to the back: 1025 below the lowest."""
    document = Document.open(drawing_probe)
    ids = [p.id for p in document.pictures()]
    for identifier in ids:
        document.float_drawing(identifier)
    heights = [document.drawing(i).wrapping["z_order"] for i in ids]
    assert heights == [251658240 + 1024 * k for k in range(len(ids))]  # Word's, in the order floated
    document.set_drawing(ids[3], z_order="back")
    assert document.drawing(ids[3]).wrapping["z_order"] == 251658240 - 1025
    document.set_drawing(ids[0], z_order="front")
    assert document.drawing(ids[0]).wrapping["z_order"] == max(heights) + 1024


def test_text_box_is_words(drawing_probe):
    """A text box as Word's Insert Text Box writes one, but for its size and place and the
    fallback's binary drawing cache (``o:gfxdata``, Word's own)."""
    document = Document.open(drawing_probe)
    paragraph = next(p for p in document.paragraphs() if p.text.startswith("Shape anchor"))
    made = document.insert_text_box(f"{paragraph.id}@0", "The end.", width=144, height=144)
    ours = next(b for b in blocks(document) if 'txBox="1"' in b)
    word = next(b for b in observed("drawings") if 'txBox="1"' in b)

    def shape(text: str) -> str:
        text = re.sub(r"<w:txbxContent>.*?</w:txbxContent>", "<w:txbxContent/>", text)  # its paragraphs aside
        text = re.sub(r' (o:gfxdata|o:spid|w14:anchorId|wp14:anchorId|id|name)="[^"]*"', "", text)
        text = re.sub(r' relativeHeight="\d+"|z-index:\d+;', "", text)
        return text[text.index("<mc:AlternateContent"):text.index("</mc:AlternateContent>")]

    assert shape(ours) == shape(word)
    assert [p.text for p in document.drawing(made.id).paragraphs] == ["The end."]
    assert check(document.package) == []
