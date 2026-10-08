"""E5's edits in detail: what each writes, what it refuses, and what it keeps -- drawings
(groups moved and resized, a text box's fallback kept in step, the VML style following a
move), content controls (each kind's fill, locks, bindings), tables (refusals, reading).
"""

from __future__ import annotations

import datetime
import sys
from pathlib import Path

import pytest

from docx_agent import Document, EditError
from docx_agent.validate import check

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import e5_probe  # noqa: E402

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"


def _group_document() -> bytes:
    """A floating group of two rectangles, as Word writes one (``wpg:wgp``)."""
    ids = e5_probe.Ids(0x3E000000)
    member = ('<wps:wsp><wps:cNvPr id="{n}" name="Rect {n}"/><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="{x}" y="0"/>'
              '<a:ext cx="500000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
              '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>')
    group = (
        '<w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="3E100001" wp14:editId="3E100002">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        '<wp:extent cx="1100000" cy="500000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '<wp:docPr id="7" name="Group 7"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="{WPG}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="1100000" cy="500000"/><a:chOff x="0" y="0"/><a:chExt cx="1100000" cy="500000"/></a:xfrm>'
        '</wpg:grpSpPr>' + member.format(n=8, x=0) + member.format(n=9, x=600000) +
        '</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>')
    body = (f'<w:p w14:paraId="{ids()}" w14:textId="77777777"><w:r><w:t>Group here.</w:t></w:r>{group}</w:p>'
            + e5_probe.para(ids, "After the group."))
    namespaces = (e5_probe.DRAWING_NS + ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
                  f' xmlns:wpg="{WPG}"')
    return e5_probe.package(body, namespaces=namespaces)


def test_a_group_is_read_moved_and_resized():
    document = Document.open(_group_document())
    group = document.drawing("d:7")
    assert group.kind == "group" and not group.inline and group.size == (1100000 / 12700, 500000 / 12700)
    document.move_drawing("d:7", 36, 12)
    document.resize_drawing("d:7", width=1100000 / 12700 * 2)
    moved = Document.open(document.to_bytes()).drawing("d:7")
    assert moved.position["x"] == 36 and moved.position["y"] == 12
    assert moved.size == pytest.approx((2200000 / 12700, 1000000 / 12700))
    frame = moved._locate()[1]
    ext = frame.find(".//{%s}grpSpPr/{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm" % WPG)
    # The group's box grows, its members' child coordinates stay: they scale with it.
    assert ext[1].get("cx") == "2200000" and ext[3].get("cx") == "1100000"
    assert check(moved.document.package) == []


def test_a_text_box_fallback_follows_every_edit():
    document = Document.open(e5_probe.drawings_document())
    paragraph = document.paragraphs()[1]
    box = document.insert_text_box(f"{paragraph.id}@0", "First line\nSecond line").id
    document.drawing(box).paragraphs[1].set_text("Second, edited")
    with document.tracking(author="E5", date="2026-10-04T12:00:00Z"):
        document.drawing(box).paragraphs[0].set_text("First, tracked")
    holder = next(document.package.tree(document.package.document_part()).iter(_MC + "AlternateContent"))
    choice = holder.find(f"{_MC}Choice").find(f".//{_W}txbxContent")
    fallback = holder.find(f"{_MC}Fallback").find(f".//{_W}txbxContent")
    strip = lambda e: [("".join(p.itertext())) for p in e.findall(_W + "p")]  # noqa: E731
    assert strip(choice) == strip(fallback)
    assert check(document.package) == []  # the copy's revision ids are its own
    accepted = Document.open(document.to_bytes())
    accepted.accept_all()
    holder = next(accepted.package.tree(accepted.package.document_part()).iter(_MC + "AlternateContent"))
    assert holder.find(_MC + "Fallback").find(f".//{_W}ins") is None


def test_a_vml_fallback_follows_a_move():
    document = Document.open(e5_probe.drawings_document())
    paragraph = document.paragraphs()[1]
    box = document.insert_text_box(f"{paragraph.id}@0", "Moved", width=144, height=72).id
    document.set_drawing(box, x=72, y=36, wrap="top_and_bottom", horizontal_from="page")
    holder = next(document.package.tree(document.package.document_part()).iter(_MC + "AlternateContent"))
    shape = holder.find(f"{_MC}Fallback/{_W}pict/{{urn:schemas-microsoft-com:vml}}shape")
    style = shape.get("style")
    assert "margin-left:1in" in style and "margin-top:36pt" in style and "mso-position-horizontal-relative:page" in style
    assert shape.find("{urn:schemas-microsoft-com:office:word}wrap").get("type") == "topAndBottom"


def test_drawing_settings_are_checked_first():
    document = Document.open(e5_probe.drawings_document())
    picture = document.pictures()[0].id
    with pytest.raises(EditError):
        document.set_drawing(picture, wrap="square")  # inline: float it first
    with pytest.raises(EditError):
        document.float_drawing(picture, wrap="sideways")
    with pytest.raises(EditError):
        document.float_drawing(picture, x=3, x_align="left")
    with pytest.raises(EditError):
        document.inline_drawing(picture)
    with pytest.raises(EditError):
        document.insert_shape(f"{document.paragraphs()[0].id}@0", "notAShape")
    document.float_drawing(picture)
    assert document.drawing(picture).wrapping["wrap"] == "front"


# -- content controls --------------------------------------------------------------------------


@pytest.fixture()
def controls():
    return Document.open(e5_probe.controls_document())


def test_each_kind_reads_and_fills(controls):
    assert controls.content_control("cc:1003").value == "A"
    controls.fill_control("cc:1003", "Beta")
    controls.fill_control("cc:1004", "Green")
    controls.fill_control("cc:1005", "2026-12-24")
    controls.fill_control("cc:1006", True)
    saved = Document.open(controls.to_bytes())
    assert saved.content_control("cc:1003").text == "Beta" and saved.content_control("cc:1003").value == "B"
    assert saved.content_control("cc:1004").value == "G"
    assert saved.content_control("cc:1005").text == "24-12-2026"
    assert saved.content_control("cc:1005").value == datetime.date(2026, 12, 24)
    assert saved.content_control("cc:1006").text == "☒" and saved.content_control("cc:1006").value is True
    controls.fill_control("cc:1006", False)
    assert controls.content_control("cc:1006").text == "☐"


def test_a_fill_is_refused_where_word_refuses_it(controls):
    with pytest.raises(EditError):
        controls.fill_control("cc:1003", "Gamma")  # a drop-down's items only
    with pytest.raises(EditError):
        controls.fill_control("cc:1006", "yes")
    with pytest.raises(EditError):
        controls.fill_control("cc:1001", "two\nlines")
    _, sdt = controls._control("cc:1001")
    from docx_agent.oxml.xml import insert_in_order, make

    insert_in_order(sdt.find(_W + "sdtPr"), make("w:lock", **{"w:val": "contentLocked"}))
    with pytest.raises(EditError):
        controls.fill_control("cc:1001", "locked")
    _, sdt = controls._control("cc:1002")
    insert_in_order(sdt.find(_W + "sdtPr"), make("w:lock", **{"w:val": "sdtLocked"}))
    with pytest.raises(EditError):
        controls.remove_control("cc:1002")


def test_a_bound_fill_writes_the_node_and_refuses_a_missing_one(controls):
    controls.fill_control("cc:1007", "Bound filled")
    assert b"<name>Bound filled</name>" in controls.package.read("customXml/item1.xml")
    _, sdt = controls._control("cc:1008")
    sdt.find(f"{_W}sdtPr/{_W}dataBinding").set(_W + "xpath", "/ns0:root[1]/ns0:missing[1]")
    with pytest.raises(EditError):
        controls.fill_control("cc:1008", "nowhere")


def test_a_placeholder_goes_with_a_fill(controls):
    _, sdt = controls._control("cc:1001")
    from docx_agent.oxml.xml import insert_in_order, make

    insert_in_order(sdt.find(_W + "sdtPr"), make("w:showingPlcHdr"))
    controls.fill_control("cc:1001", "Real value")
    _, sdt = controls._control("cc:1001")
    assert sdt.find(f"{_W}sdtPr/{_W}showingPlcHdr") is None


def test_controls_inserted_at_a_controls_edge_are_not_nested(controls):
    paragraph = controls.paragraphs()[1]
    made = [controls.insert_control(f"{paragraph.id}@0", kind, items=items)
            for kind, items in (("text", None), ("rich-text", None), ("combo-box", ["x"]))]
    for result in made:
        _, sdt = controls._control(result.id)
        assert sdt.getparent().tag == _W + "p"
    assert check(controls.package) == [p for p in check(Document.open(e5_probe.controls_document()).package)]


def test_removing_a_control_keeps_or_drops_its_content(controls):
    controls.remove_control("cc:1001")
    assert "Plain value" in controls.paragraphs()[1].text
    controls.remove_control("cc:1100", keep_content=False)
    assert all("Block one" not in p.text for p in controls.paragraphs())
    assert "cc:1001" not in [c.id for c in controls.content_controls()]


# -- tables ------------------------------------------------------------------------------------


def test_table_edits_are_checked_first():
    document = Document.open("tests/fixtures/samplelib/sample-simple.docx")
    paragraph = document.paragraphs()[0]
    with pytest.raises(EditError):
        document.insert_table(0, 2, after=paragraph.id)
    with pytest.raises(EditError):
        document.insert_table(2, 2, after=paragraph.id, widths=[72])
    table = document.insert_table(2, 2, after=paragraph.id).id
    for bad in ({"layout": "stretchy"}, {"alignment": "middle"}, {"width": "120%"}, {"borders": {"diagonal": "single"}},
                {"floating": {"x": 1, "horizontal_anchor": "sky"}}, {"look": {"zebra": True}}):
        with pytest.raises(EditError):
            document.set_table(table, **bad)
    with pytest.raises(EditError):
        document.set_cell(table, 0, 0, vertical_alignment="middle")
    with pytest.raises(EditError):
        document.set_row(table, 5, height=10)
    with pytest.raises(EditError):
        document.split_cell(table, 0, 0, rows=0)


def test_a_table_never_touches_another():
    document = Document.open("tests/fixtures/samplelib/sample-simple.docx")
    paragraph = document.paragraphs()[0]
    first = document.insert_table(2, 2, after=paragraph.id).id
    second = document.insert_table(2, 2, after=first).id
    element = document.table(first)._entry()[1].element
    following = element.getnext()
    assert following.tag == _W + "p" and following.getnext() is document.table(second)._entry()[1].element
    cell = document.table(second).cell(0, 0)
    nested = document.insert_table(1, 1, after=cell.paragraphs[0].id).id
    assert document.table(nested)._entry()[1].element.getnext().tag == _W + "p"
    assert check(document.package) == check(Document.open("tests/fixtures/samplelib/sample-simple.docx").package)
