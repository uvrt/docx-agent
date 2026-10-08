"""E5 tracked (ROADMAP.md, "Accept and reject"): for every E5 edit Word tracks -- or that
decision 7 writes in the nearest form Word reviews (merges, splits, columns across spans)
-- on every fixture, accepting every revision gives the untracked edit and rejecting them
the document before it, in canonical form.  Each case prepares what it edits untracked (a
table with spans and merges, a control bound to a custom XML part), then edits it once
untracked and once inside ``Document.tracking``, saved and reopened.

What Word does not track is applied untracked with a warning (a row's height, a table's or
a drawing's floating, a drawing's place, wrapping, order and size, a control's wrapper):
held separately below.
"""

from __future__ import annotations

import pytest

from docx_agent import Document
from docx_agent.validate import check, grid_problems

from canonical import canonical, difference, without_added_styles
from e4_edits import enough

AUTHOR = "E5 Agent"
DATE = "2026-10-04T12:00:00Z"
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _spanned(document: Document) -> str:
    """A 4 x 4 table with a cell over two columns, one merged down two rows and a 2 x 2 merge."""
    paragraphs = enough(document, 2)
    data = [[f"s{r}{c}" for c in range(4)] for r in range(4)]
    table = document.insert_table(4, 4, after=paragraphs[0].id, data=data).id
    document.merge_cells(table, (0, 0), (0, 1))
    document.merge_cells(table, (1, 3), (2, 3))
    return _find(document)


def _find(document: Document) -> str:
    for table in document.tables():
        if "s33" in "".join(table._entry()[1].element.itertext()):
            return table.id
    raise AssertionError("the prepared table is gone")


def _bound(document: Document) -> str:
    """A text control bound to a custom XML part's node."""
    paragraph = enough(document, 2)[0]
    made = document.insert_control(f"{paragraph.id}@0", "text", text="E5 bound before")
    _bind(document, made.id)
    return made.id


def _bind(document: Document, identifier: str) -> None:
    store = "{6C3C8BC8-F283-45AE-878A-BAB7291F4D35}"
    package = document.package
    main = package.document_part()
    item = package.unused_part_name("customXml/item{n}.xml")
    props = item.replace("item", "itemProps")
    package.add_part(item, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                           b'<root xmlns="urn:docx-agent:e5"><name>E5 bound before</name></root>')
    package.add_part(props, ('<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n'
                             f'<ds:datastoreItem ds:itemID="{store}" xmlns:ds="http://schemas.openxmlformats.org/'
                             'officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>').encode(),
                     "application/vnd.openxmlformats-officedocument.customXmlProperties+xml", override=True)
    package.add_relationship(main, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
                             item)
    package.add_relationship(item, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                                   "customXmlProps", props)
    _, sdt = document._control(identifier)
    properties = sdt.find(_W + "sdtPr")
    from docx_agent.oxml.xml import insert_in_order, make

    insert_in_order(properties, make("w:dataBinding", **{"w:prefixMappings": "xmlns:ns0='urn:docx-agent:e5'",
                                                          "w:xpath": "/ns0:root[1]/ns0:name[1]",
                                                          "w:storeItemID": store}))
    document.package.mark_dirty(document.package.document_part())


CASES = {
    "insert_table": (None, lambda d, _: d.insert_table(3, 3, after=enough(d, 2)[0].id,
                                                       data=[["t1", "t2", "t3"], ["t4", "t5", "t6"]], header_rows=1)),
    "table_properties": (_spanned, lambda d, t: d.set_table(_find(d), width="80%", alignment="center",
                                                            layout="fixed", borders={"top": "double"},
                                                            look={"first_row": False}, cell_margins={"left": 9})),
    "cell_properties": (_spanned, lambda d, t: d.set_cell(_find(d), 2, 2, shading="FFFF00", vertical_alignment="bottom",
                                                          borders={"top": "double", "left": None},
                                                          margins={"top": 4})),
    "column_width": (_spanned, lambda d, t: d.set_column_width(_find(d), 2, 120)),
    "merge_across": (_spanned, lambda d, t: d.merge_cells(_find(d), (3, 1), (3, 2))),
    "merge_rectangle": (_spanned, lambda d, t: d.merge_cells(_find(d), (2, 0), (3, 1))),
    "merge_grown": (_spanned, lambda d, t: d.merge_cells(_find(d), (0, 1), (1, 1))),
    "split_across": (_spanned, lambda d, t: d.split_cell(_find(d), 3, 2, rows=1, columns=2)),
    "split_down": (_spanned, lambda d, t: d.split_cell(_find(d), 3, 0, rows=2, columns=1)),
    "split_merge": (_spanned, lambda d, t: d.split_cell(_find(d), 1, 3, rows=2, columns=1)),
    "split_both": (_spanned, lambda d, t: d.split_cell(_find(d), 2, 1, rows=2, columns=3)),
    "insert_column_span": (_spanned, lambda d, t: d.insert_column(_find(d), 0, right=True)),
    "delete_column_span": (_spanned, lambda d, t: d.delete_column(_find(d), 1)),
    "insert_row_merge": (_spanned, lambda d, t: d.insert_row(_find(d), 1, below=True)),
    "delete_row_merge": (_spanned, lambda d, t: d.delete_row(_find(d), 1)),
    "text_box": (None, lambda d, _: d.insert_text_box(f"{enough(d, 2)[0].id}@0", "E5 tracked box\nTwo", x=200)),
    "shape": (None, lambda d, _: d.insert_shape(f"{enough(d, 2)[0].id}@0", "ellipse", text="E5 tracked shape")),
    "fill_control": (_bound, lambda d, c: d.fill_control(c, "E5 bound after")),
}
_MODERN = ("commentsExtended", "commentsIds", "commentsExtensible")


def _forgiving(form: dict[str, str]) -> dict[str, str]:
    """Without the modern comment parts: a describing comment gives a document's old-style
    comments theirs, as Word does on saving, and they stay when it goes (E3)."""
    out = {k: v for k, v in form.items() if not any(f"word/{m}.xml" == k for m in _MODERN)}
    for key in ("[Content_Types].xml", "word/_rels/document.xml.rels"):
        if key in out:
            out[key] = "\n".join(line for line in out[key].splitlines() if not any(m in line for m in _MODERN))
    return out


@pytest.mark.parametrize("case", list(CASES))
def test_accept_all_is_the_edit_and_reject_all_the_original(docx_path, case):
    prepare, edit = CASES[case]
    data = docx_path.read_bytes()
    base = Document.open(data)
    target = prepare(base) if prepare else None
    base._stamp_document()
    start = base.to_bytes()
    original = _forgiving(canonical(start, ids=False))

    untracked = Document.open(start)
    edit(untracked, target)
    expected = _forgiving(canonical(untracked.to_bytes(), ids=False))

    tracked = Document.open(start)
    with tracked.tracking(author=AUTHOR, date=DATE):
        edit(tracked, target)
    tracked_bytes = tracked.to_bytes()
    assert set(check(Document.open(tracked_bytes).package)) - set(check(Document.open(start).package)) == set()
    assert tracked.revisions(author=AUTHOR), "nothing was tracked"

    accepted = Document.open(tracked_bytes)
    accepted.accept(author=AUTHOR)
    assert not accepted.revisions(author=AUTHOR)
    got = without_added_styles(_forgiving(canonical(accepted.to_bytes(), ids=False)), expected)
    assert got == expected, "accept-all is not the untracked edit:\n" + difference(expected, got)
    for table in accepted.tables():
        assert grid_problems(table._entry()[1].element) == []

    rejected = Document.open(tracked_bytes)
    rejected.reject(author=AUTHOR)
    assert not rejected.revisions(author=AUTHOR)
    got = without_added_styles(_forgiving(canonical(rejected.to_bytes(), ids=False)), original)
    assert got == original, "reject-all is not the original:\n" + difference(original, got)


def _bound_xml(document: Document) -> bytes:
    return next(document.package.read(name) for name in document.package.part_names
                if name.startswith("customXml/item") and b"urn:docx-agent:e5" in document.package.read(name))


def test_a_bound_control_follows_the_review():
    """Rejecting a tracked fill of a bound control puts the old text back in the control
    and in its custom XML node; accepting keeps the new in both."""
    document = Document.open("tests/fixtures/samplelib/sample-simple.docx")
    control = _bound(document)
    with document.tracking(author=AUTHOR, date=DATE):
        document.fill_control(control, "E5 bound after")
    data = document.to_bytes()
    assert b"E5 bound after" in _bound_xml(Document.open(data))
    rejected = Document.open(data)
    rejected.reject_all()
    assert rejected.content_control(control).text == "E5 bound before"
    assert b"E5 bound before" in _bound_xml(rejected)
    accepted = Document.open(data)
    accepted.accept_all()
    assert accepted.content_control(control).text == "E5 bound after"
    assert b"E5 bound after" in _bound_xml(accepted)


UNTRACKED_IN_WORD = {
    "row": lambda d: d.set_row(d.insert_table(2, 2, after=enough(d, 1)[0].id).id, 0, height=30, cant_split=True),
    "floating table": lambda d: d.set_table(d.insert_table(2, 2, after=enough(d, 1)[0].id).id,
                                            floating={"x": 10, "y": 10}),
    "header rows": lambda d: d.set_table(d.insert_table(2, 2, after=enough(d, 1)[0].id).id, header_rows=1),
    "float a picture": lambda d: d.float_drawing(_picture(d), wrap="square"),
    "a drawing's place": lambda d: (lambda p: (d.float_drawing(p), d.set_drawing(p, x=10, z_order="back")))(_picture(d)),
    "a control": lambda d: d.insert_control(f"{enough(d, 1)[0].id}@0", "checkbox"),
}


def _picture(document: Document) -> str:
    from e1_edits import png

    return document.insert_picture(f"{enough(document, 1)[0].id}@0", png(8, 8), track=False).id


@pytest.mark.parametrize("what", list(UNTRACKED_IN_WORD))
def test_what_word_does_not_track_is_applied_and_said(what):
    """Tracked, an edit Word applies untracked is applied untracked here too, and the
    result's warnings say so."""
    document = Document.open("tests/fixtures/samplelib/sample-simple.docx")
    with document.tracking(author=AUTHOR, date=DATE):
        result = UNTRACKED_IN_WORD[what](document)
    results = result if isinstance(result, tuple) else (result,)
    assert any("Word" in w and ("untracked" in w or "no revision" in w) for r in results if r is not None
               for w in r.warnings), what


def test_a_vertical_text_direction_tracks_the_cell_and_not_the_row():
    """Word records a cell's text direction (``w:tcPrChange``) and its paragraphs' indents
    (``w:pPrChange``), and changes the row's height and breaking without a record
    (measured): rejecting gives the cell back, the row keeps them."""
    document = Document.open("tests/fixtures/samplelib/sample-simple.docx")
    table = document.insert_table(2, 2, after=enough(document, 1)[0].id).id
    with document.tracking(author=AUTHOR, date=DATE):
        document.set_cell(table, 0, 0, text_direction="btLr")
    rejected = Document.open(document.to_bytes())
    rejected.reject_all()
    table = rejected.tables()[-1].id if rejected.tables()[-1].id == table else table
    assert rejected.cell_properties(table, 0, 0)["text_direction"] is None
    assert rejected.row_properties(table, 0)["cant_split"] is True
