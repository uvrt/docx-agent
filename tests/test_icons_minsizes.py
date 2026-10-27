"""Round 17, P6: an icon from the application's asset catalogue, inline or floating, tinted
as Word draws a recoloured SVG (measured: agent-work feedback-round17/measurements-B.md), and
the application's minimum text sizes by kind, which check and the builders both read.
Made-up icons and documents."""

from __future__ import annotations

import io
import json
import re
import struct
import zipfile

import pytest

pytest.importorskip("ooxml_edit.pictures", reason="needs ooxml-edit 0.68 (ooxml_edit.pictures)")
pytest.importorskip("resvg_py", reason="an SVG's PNG fallback is drawn by docx2svg[png]")

from ooxml_edit.tools import FolderCatalogue, Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.edit.errors import EditError  # noqa: E402
from docx_agent.edit.minimums import MinimumError, text_minimums  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS, docx_format  # noqa: E402

ICON = (b'<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">'
        b'<circle cx="12" cy="12" r="10" fill="#000000"/><path fill="#FFFFFF" d="M11 6h2v8h-2zM11 16h2v2h-2z"/>'
        b'</svg>')
WIDE = b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 20"><rect width="40" height="20"/></svg>'


def _parts(document: Document) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(document.to_bytes())) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _blank() -> tuple[Document, str]:
    document = Document.new()
    first = document.paragraphs()[0].id
    document.paragraph(first).set_text("Icon: ")
    return document, first


# -- the library: an SVG picture, tinted ---------------------------------------------------------


def test_an_svg_picture_is_stored_as_word_stores_one():
    document, first = _blank()
    made = document.insert_picture(f"{first}@6", ICON, alt_text="A made-up icon")
    picture = document.picture(made.id)
    assert picture.size == (18.0, 18.0)                   # 24 px at 96 dpi, as Word places it
    parts = _parts(document)
    assert sorted(n for n in parts if n.startswith("word/media/")) == ["word/media/image1.png",
                                                                     "word/media/image1.svg"]
    assert parts["word/media/image1.svg"] == ICON
    assert struct.unpack(">II", parts["word/media/image1.png"][16:24]) == (48, 48)   # twice its size
    assert b'<Default Extension="svg" ContentType="image/svg+xml"/>' in parts["[Content_Types].xml"]
    xml = parts["word/document.xml"].decode()
    assert re.search(r'<wp:docPr id="\d+" name="Graphic \d+"', xml)
    assert '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip' in xml
    rels = parts["word/_rels/document.xml.rels"].decode()
    assert re.search(r'Type="[^"]+/relationships/image" Target="media/image1.svg"', rels)
    assert document.validate() == []
    wide = document.insert_picture(f"{first}@0", WIDE, width=20, alt_text="A bar")
    assert document.picture(wide.id).size == (20.0, 10.0)


def test_a_tint_is_a_theme_colour_with_its_class_or_a_hex_colour():
    document, first = _blank()
    document.insert_picture(f"{first}@6", ICON, width=12, alt_text="Icon", color="accent2")
    svg = next(v for k, v in _parts(document).items() if k.endswith(".svg")).decode()
    accent2 = document._theme_color("accent2")
    assert f".MsftOfcThm_Accent2_Fill_v2 {{\n fill:#{accent2}; \n}}" in svg
    assert svg.count('class="MsftOfcThm_Accent2_Fill_v2"') == 2 and svg.count(f'fill="#{accent2}"') == 2
    other, first = _blank()
    other.insert_picture(f"{first}@6", ICON, alt_text="Icon", color="1f4e79")
    svg = next(v for k, v in _parts(other).items() if k.endswith(".svg")).decode()
    assert "MsftOfcThm" not in svg and svg.count('fill="#1F4E79"') == 2
    text, first = _blank()
    text.insert_picture(f"{first}@6", ICON, alt_text="Icon", color="Text 1")
    assert "MsftOfcThm_Text1_Fill_v2" in next(v for k, v in _parts(text).items() if k.endswith(".svg")).decode()


def test_what_cannot_be_tinted_or_read_is_refused():
    from test_tools_word import png as make_png

    document, first = _blank()

    with pytest.raises(EditError, match="only an SVG picture can be tinted"):
        document.insert_picture(f"{first}@0", make_png(), alt_text="x", color="accent1")
    with pytest.raises(EditError, match="theme colour .* or RRGGBB"):
        document.insert_picture(f"{first}@0", ICON, alt_text="x", color="purple-ish")
    with pytest.raises(EditError, match="document type"):
        document.insert_picture(f"{first}@0", b'<!DOCTYPE svg><svg xmlns="http://www.w3.org/2000/svg"/>',
                                alt_text="x")
    with pytest.raises(EditError, match="states no size"):
        document.insert_picture(f"{first}@0", b'<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>',
                                alt_text="x")


# -- the tools: the catalogue ---------------------------------------------------------------------


@pytest.fixture
def icons(tmp_path):
    folder = tmp_path / "icons"
    (folder / "icon").mkdir(parents=True)
    (folder / "icon" / "alert-circle.svg").write_bytes(ICON)
    (folder / "index.json").write_text(json.dumps([
        {"id": "icon/alert-circle", "file": "icon/alert-circle.svg", "name": "Alert circle",
         "keywords": ["alert", "warning", "circle"]}]))
    return FolderCatalogue(folder)


def test_without_a_catalogue_nothing_is_sent_for_it():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        names = {d["name"]: d for d in box.definitions("anthropic", defer=False)}
    assert "word_find_asset" not in names
    assert "asset" not in names["word_drawings"]["input_schema"]["properties"]
    assert "color" in names["word_drawings"]["input_schema"]["properties"]


def test_an_icon_from_the_catalogue_in_a_cell_tinted_and_floating(icons):
    with Toolbox(TOOLS, formats=[docx_format(assets=icons)], groups=GROUPS) as box:
        names = {d["name"]: d for d in box.definitions("anthropic", defer=False)}
        assert "asset" in names["word_drawings"]["input_schema"]["properties"] and "word_find_asset" in names
        session = box.session()
        sample = Document.new()
        sample.insert_markdown("| Item | Note |\n| --- | --- |\n| One | a |", at="end")
        d = session.open(sample.to_bytes(), "icons.docx")
        table = session.entry(d).document.tables()[0].id
        found = box.dispatch(session, "word_find_asset", {"query": "warning"})
        assert found.ok and found.data["assets"][0]["id"] == "icon/alert-circle"
        # The team's call, as written: size_pt read as width, a cell as the place.
        result = box.dispatch(session, "word_drawings", {
            "doc": d, "action": "insert_picture", "asset": "icon/alert-circle", "at": f"{table}/c1,0",
            "size_pt": 12, "color": "accent2"})
        assert result.ok, result.error and result.error.to_json()
        document = session.entry(d).document
        picture = document.picture(result.data["id"])
        assert picture.size == (12.0, 12.0) and picture.alt_text == "alert circle"
        assert document.table(table).cell(1, 0).paragraphs[0].text.startswith("￼")
        svg = next(v for k, v in _parts(document).items() if k.endswith(".svg")).decode()
        assert "MsftOfcThm_Accent2_Fill_v2" in svg
        floating = box.dispatch(session, "word_drawings", {
            "doc": d, "action": "insert_picture", "asset": "icon/alert-circle", "at": f"{table}/c1,1",
            "width": 24, "wrap": "square", "align": "right", "alt_text": "Warning"})
        assert floating.ok and not document.drawing(floating.data["id"]).inline
        for arguments, field in (({"asset": "icon/none"}, "asset"), ({"asset": "icon/alert-circle", "color": "x"},
                                                                       "color")):
            refused = box.dispatch(session, "word_drawings", {"doc": d, "action": "insert_picture",
                                                              "at": f"{table}/c1,0", **arguments})
            assert not refused.ok and refused.error.field == field
        assert document.validate() == []


# -- minimum text sizes by kind -------------------------------------------------------------------


def test_the_setting_is_read_once_for_everything():
    assert text_minimums(None) == {}
    assert set(text_minimums(9)) == {"body", "table", "footnote", "endnote", "header", "footer", "caption"}
    assert text_minimums({"table": 8, "caption": None}) == {"table": 8.0}
    for bad in ({"chart": 9}, {}, 0, "9", True):
        with pytest.raises(MinimumError):
            text_minimums(bad)
    with pytest.raises(MinimumError):
        docx_format(min_text_size={"slides": 10})


def _box_session(box, markdown: str):
    session = box.session()
    sample = Document.new()
    d = session.open(sample.to_bytes(), "sizes.docx")
    result = box.dispatch(session, "word_insert_markdown", {"doc": d, "at": "end", "markdown": markdown})
    assert result.ok, result.error and result.error.to_json()
    return session, d, result


def test_builders_and_check_report_text_below_its_kinds_minimum():
    markdown = "Body text.\n\n| Item | Note |\n| --- | --- |\n| One | a |"
    with Toolbox(TOOLS, formats=[docx_format(min_text_size={"table": 30})], groups=GROUPS) as box:
        session, d, result = _box_session(box, markdown)
        small = result.data["small_text"]
        assert small and {e["kind"] for e in small} == {"table"} and all(e["min"] == 30 for e in small)
        assert any(w.startswith("small_text:") for w in result.warnings)
        checked = box.dispatch(session, "check", {"doc": d, "include": ["layout"]})
        assert checked.ok
        layout = checked.data["layout"]
        assert layout["small_text"] and {e["kind"] for e in layout["small_text"]} == {"table"}
        assert all(e["min"] == 30 for e in layout["small_text"])
        # The same paragraphs, whichever reads it.
        assert {e["para"] for e in layout["small_text"]} == {e["para"] for e in small}
    # A caption has a minimum of its own.
    with Toolbox(TOOLS, formats=[docx_format(min_text_size={"caption": 30})], groups=GROUPS) as box:
        session, d, result = _box_session(box, "Body text.")
        assert "small_text" not in result.data
        body = session.entry(d).document.paragraphs()[-1].id
        assert box.dispatch(session, "word_fields", {"doc": d, "action": "insert_caption", "after": body,
                                                     "label": "Table", "text": "Made-up items"}).ok
        found = box.dispatch(session, "check", {"doc": d, "include": ["layout"]}).data["layout"]["small_text"]
        assert [e["kind"] for e in found] == ["caption"]


def test_without_the_setting_nothing_changes():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session, d, result = _box_session(box, "Body text.\n\n| a | b |\n| --- | --- |\n| 1 | 2 |")
        assert "small_text" not in result.data and not any(w.startswith("small_text") for w in result.warnings)
        layout = box.dispatch(session, "check", {"doc": d, "include": ["layout"]}).data["layout"]
        assert "small_text" not in layout
        # check's own min_size is one number for every kind, as before: no "min" per entry.
        numbered = box.dispatch(session, "check", {"doc": d, "min_size": 30}).data["layout"]["small_text"]
        assert numbered and all("min" not in e for e in numbered)
