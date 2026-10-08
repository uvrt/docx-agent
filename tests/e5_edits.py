"""E5's operations, for any document (ROADMAP.md, Phase E5): each edits a document and
returns a check that reads the edit back from a saved and reopened copy.

Tables (a new one with header rows, its properties, rows and cells; merges and splits and
columns across them; a nested table; a floating table), drawings (a picture floated, wrapped,
positioned, ordered, locked and put inline again; a text box and a shape, their text edited
through the paragraph API) and content controls (each kind inserted and filled, a block
one, one removed).  Each works on paragraphs of the document's own body
(``e4_edits.enough`` writes them where a document has too few).
"""

from __future__ import annotations

import datetime

from docx_agent import Document
from docx_agent.validate import grid_problems

from e4_edits import enough
from e1_edits import png

DATA = [["E5 head one", "E5 head two", "E5 head three"], ["a1", "a2", "a3"], ["b1", "b2", "b3"]]
BOX = "E5 text box line"
SHAPE_TEXT = "E5 shape text"
CONTROL = "E5 control value"


def _table(document: Document, data=DATA, **options) -> str:
    paragraphs = enough(document, 2)
    return document.insert_table(len(data), len(data[0]), after=paragraphs[0].id, data=data, **options).id


def _grid_ok(document: Document) -> bool:
    return all(grid_problems(document._table_entry(t.id)[1].element) == [] for s in document.stories
               for t in s.tables)


def _table_with(document: Document, text: str):
    for table in document.tables():
        if text in "".join(table._entry()[1].element.itertext()):
            return table
    return None


def op_insert_table(document: Document, header: int = 1):
    _table(document, header_rows=header)

    def check(saved: Document) -> bool:
        table = _table_with(saved, "E5 head two")
        return (table is not None and table.properties["header_rows"] == header and table.cell(1, 2).text == "a3"
                and len(table.properties["columns"]) == 3 and _grid_ok(saved))

    return check


def op_table_properties(document: Document):
    table = _table(document)
    document.set_table(table, width="80%", alignment="center", layout="fixed",
                       look={"first_row": True, "banded_rows": False}, borders={"top": "double"},
                       cell_margins={"left": 6})
    document.set_row(table, 0, height=24, height_rule="exact", cant_split=True)
    document.set_cell(table, 1, 1, shading="FFFF00", borders={"bottom": {"style": "single", "size": 1.5,
                                                                         "color": "C00000"}},
                      vertical_alignment="center", text_direction="btLr", margins={"top": 3})
    document.set_column_width(table, 0, 100)

    def check(saved: Document) -> bool:
        found = _table_with(saved, "E5 head two")
        props = found.properties
        row = saved.row_properties(found.id, 0)
        cell = saved.cell_properties(found.id, 1, 1)
        below = saved.cell_properties(found.id, 2, 1)
        return (props["width"] == "80%" and props["alignment"] == "center" and props["layout"] == "fixed"
                and props["look"]["banded_rows"] is False and props["borders"]["top"]["style"] == "double"
                and props["columns"][0] == 100 and row == {"height": 24, "height_rule": "exact", "cant_split": True,
                                                         "header": False}
                and cell["shading"]["fill"] == "FFFF00" and cell["vertical_alignment"] == "center"
                and cell["text_direction"] == "btLr" and cell["borders"]["bottom"]["color"] == "C00000"
                and below["borders"]["top"]["color"] == "C00000" and _grid_ok(saved))

    return check


def op_merge_split(document: Document, down: bool = True):
    data = [[f"m{r}{c}" for c in range(4)] for r in range(4)]
    table = _table(document, data)
    document.merge_cells(table, (0, 0), (1, 1) if down else (0, 1))
    document.split_cell(table, 2, 2, rows=2 if down else 1, columns=2)
    table = document.tables()[[t.id for t in document.tables()].index(_table_with(document, "m33").id)].id
    document.insert_column(table, 0, right=True)
    document.delete_column(table, 2)
    document.merge_cells(table, (3, 0), (3, 1))

    def check(saved: Document) -> bool:
        found = _table_with(saved, "m33")
        if found is None or not _grid_ok(saved):
            return False
        text = "".join(found._entry()[1].element.itertext())
        merged = found.cell(0, 0)
        wanted = ["m00", "m01", "m10", "m11"] if down else ["m00", "m01"]
        kept = ("m00", "m11", "m33") if down else ("m00", "m01", "m33")  # across: column 1 deleted
        return all(t in text for t in kept) and merged.text.split("\n")[:len(wanted)] == wanted

    return check


def op_nested_table(document: Document):
    table = _table(document, [["E5 outer one", "E5 outer two"], ["o1", "o2"]])
    cell = document.table(table).cell(1, 1)
    inner = document.insert_table(2, 2, after=cell.paragraphs[0].id, data=[["n1", "n2"], ["n3", "n4"]])

    def check(saved: Document) -> bool:
        outer = _table_with(saved, "E5 outer two")
        element = outer._entry()[1].element
        nested = element.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl")
        return len(nested) == 1 and "n4" in "".join(nested[0].itertext()) and _grid_ok(saved) \
            and inner.id in [t.id for t in saved.tables()]

    return check


def op_floating_table(document: Document):
    table = _table(document)
    document.set_table(table, floating={"x": 36, "y": 6, "horizontal_anchor": "margin", "vertical_anchor": "text",
                                        "right": 9, "bottom": 6, "overlap": False})

    def check(saved: Document) -> bool:
        floating = _table_with(saved, "E5 head two").properties["floating"]
        return floating is not None and floating["x"] == 36 and floating["y"] == 6 and floating["overlap"] is False

    return check


def op_float_picture(document: Document, overlap: bool = False):
    paragraphs = enough(document, 2)
    picture = document.insert_picture(f"{paragraphs[1].id}@0", png(32, 16, (180, 40, 40)), width=72, alt_text="E5 float").id
    document.float_drawing(picture, wrap="square", side="both", distances={"left": 6, "right": 6})
    document.set_drawing(picture, horizontal_from="margin", x_align="right", vertical_from="paragraph", y=4,
                         z_order="front", lock_anchor=True, allow_overlap=overlap)

    def check(saved: Document) -> bool:
        found = saved.drawing(picture)
        position, wrapping = found.position, found.wrapping
        return (position["x_align"] == "right" and position["horizontal_from"] == "margin" and position["y"] == 4
                and wrapping["wrap"] == "square" and wrapping["lock_anchor"] and wrapping["allow_overlap"] == overlap
                and wrapping["distances"]["left"] == 6)

    return check


def op_picture_round_trip(document: Document):
    paragraphs = enough(document, 2)
    picture = document.insert_picture(f"{paragraphs[1].id}@0", png(32, 16, (180, 40, 40)), width=48, alt_text="E5 inline").id
    document.float_drawing(picture, wrap="behind", x=12, y=0)
    document.set_drawing(picture, width=60)
    document.inline_drawing(picture)

    def check(saved: Document) -> bool:
        found = saved.drawing(picture)
        return found.inline and abs(found.size[0] - 60) < 0.01

    return check


def op_text_box(document: Document, anchor: int = 0):
    paragraphs = enough(document, 2)
    box = document.insert_text_box(f"{paragraphs[anchor].id}@0", "E5 box to edit\nSecond line", width=160, height=60,
                                   x=200, y=10).id
    document.drawing(box).paragraphs[0].set_text(BOX)
    document.append_paragraph("E5 box appended", story=box)

    def check(saved: Document) -> bool:
        found = saved.drawing(box)
        return found.kind == "text-box" and found.text == f"{BOX}\nSecond line\nE5 box appended" \
            and found.wrapping["wrap"] == "square"

    return check


def op_shape(document: Document, anchor: int = 0):
    paragraphs = enough(document, 2)
    shape = document.insert_shape(f"{paragraphs[anchor].id}@0", "ellipse", width=80, height=40, x=320, y=0,
                                  text=SHAPE_TEXT).id
    plain = document.insert_shape(f"{paragraphs[anchor].id}@0", "triangle", width=40, height=40, x=420, y=0,
                                  wrap="behind").id
    document.set_drawing(plain, x_align="right", horizontal_from="margin")

    def check(saved: Document) -> bool:
        return saved.drawing(shape).text == SHAPE_TEXT and saved.drawing(plain).kind == "shape" \
            and saved.drawing(plain).wrapping["wrap"] == "behind"

    return check


def op_controls(document: Document):
    paragraphs = enough(document, 3)
    a, b = paragraphs[0], paragraphs[1]
    text = document.insert_control(f"{a.id}@0", "text", tag="E5Text").id
    rich = document.insert_control(f"{a.id}@0", "rich-text").id
    drop = document.insert_control(f"{a.id}@0", "drop-down", items=["Alpha", ("Beta", "B")]).id
    combo = document.insert_control(f"{a.id}@0", "combo-box", items=["Red", "Green"]).id
    date = document.insert_control(f"{a.id}@0", "date").id
    box = document.insert_control(f"{a.id}@0", "checkbox").id
    block = document.insert_control(None, "rich-text", text="E5 block", after=b.id).id
    gone = document.insert_control(f"{b.id}@0", "text", text="E5 removed").id
    document.fill_control(text, CONTROL)
    document.fill_control(rich, "E5 rich")
    document.fill_control(drop, "B")
    document.fill_control(combo, "Purple")
    document.fill_control(date, datetime.date(2026, 10, 4))
    document.fill_control(box, True)
    document.fill_control(block, "E5 block one\nE5 block two")
    document.remove_control(gone)

    def check(saved: Document) -> bool:
        values = {c.id: c for c in saved.content_controls()}
        return (values[text].text == CONTROL and values[rich].text == "E5 rich" and values[drop].text == "Beta"
                and values[drop].value == "B" and values[combo].text == "Purple"
                and values[date].value == datetime.date(2026, 10, 4) and values[date].text == "4-10-2026"
                and values[box].value is True and values[block].text == "E5 block one\nE5 block two"
                and gone not in values and "E5 removed" in saved.paragraph(b.id).text)

    return check


OPERATIONS = [op_insert_table, op_table_properties, op_merge_split, op_nested_table, op_floating_table,
              op_float_picture, op_picture_round_trip, op_text_box, op_shape, op_controls]


def e5_edit_set(document: Document) -> dict:
    """E5's representative edit set, composed (for the Word oracle): a table with a header
    row; a merge and a split; a nested table; a floating picture; a text box; a shape;
    content controls of each kind, filled.
    The set is held to docx2svg's pages, so it keeps out of what docx2svg does not model,
    each found by the oracle and proposed to docx2svg (ROADMAP.md, Phase E5): a drawing that
    may not overlap (Word moves it clear), a vertical merge whose row Word splits at a page's
    foot (docx2svg moves the row whole), a header row at a page's foot (Word moves it whole
    and repeats it; docx2svg splits it), and a text box wrapped beside the tables (Word
    places them beside it differently) -- the box and shape go in the last paragraph."""
    checks = {"insert_table": op_insert_table(document, header=0)}
    for operation in (op_nested_table, op_controls):
        checks[operation.__name__[3:]] = operation(document)
    checks["merge_split"] = op_merge_split(document, down=False)
    checks["text_box"] = op_text_box(document, anchor=-1)
    checks["shape"] = op_shape(document, anchor=-1)
    checks["float_picture"] = op_float_picture(document, overlap=True)
    return checks
