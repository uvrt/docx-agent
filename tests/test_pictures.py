"""Inline pictures: inserted from bytes or a file, media shared by content and reaped when
unused, replaced keeping the frame, resized, described, deleted -- with unique ``docPr`` ids
and the ``wp14`` ids Word needs to keep the document's paraIds."""

from __future__ import annotations

import struct
import zlib

import pytest

from docx_agent import Document, EditError
from docx_agent.edit.ids import valid_long_hex
from docx_agent.validate import check

from test_roundtrip import entries

WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
WP14 = "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}"


def png(width: int, height: int, color=(200, 30, 30), dpi: int | None = None) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(color) * width for _ in range(height))
    out = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    if dpi:
        per_metre = int(round(dpi / 0.0254))
        out += chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    return out + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


@pytest.fixture
def lists_doc(markup_doc):
    return markup_doc.parent / "lists-and-styles.docx"


def media(document: Document) -> list[str]:
    return sorted(name for name in document.package.part_names if name.startswith("word/media/"))


def test_pictures_are_read(lists_doc):
    document = Document.open(lists_doc)
    [picture] = document.pictures()
    assert picture.id == "d:1" and picture.alt_text == "A blue square" and picture.inline
    assert picture.size == (36.0, 36.0)
    assert picture.media == "word/media/image1.png"


def test_insert_a_picture_at_its_natural_size(markup_doc, tmp_path):
    data = markup_doc.read_bytes()
    document = Document.open(data)
    before = set(check(document.package))
    result = document.insert_picture("p:1A2B3C17@4", png(96, 48), alt_text="Red")
    picture = document.picture(result.id)
    assert picture.size == (72.0, 36.0)                  # 96 dpi where the file states none (measured)
    assert document.paragraph("p:1A2B3C17").text == "The ￼last paragraph."
    frame = picture._locate()[1]
    assert valid_long_hex(frame.get(WP14 + "anchorId")) and valid_long_hex(frame.get(WP14 + "editId"))
    assert frame.find(WP + "docPr").get("descr") == "Red"
    path = tmp_path / "blue.png"
    path.write_bytes(png(300, 150, (0, 0, 255), dpi=300))
    second = document.insert_picture(document.paragraph("p:1A2B3C17").range(0, 0), path)
    assert document.picture(second.id).size == (72.0, 36.0)   # the stated 300 dpi (measured)
    ids = [int(n.get("id")) for n in document.package.tree("word/document.xml").iter(WP + "docPr")]
    assert len(ids) == len(set(ids))
    assert set(check(Document.open(document.to_bytes()).package)) - before == set()
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)


def test_one_dimension_keeps_the_aspect_ratio(markup_doc):
    document = Document.open(markup_doc)
    result = document.insert_picture("p:1A2B3C17@0", png(40, 20), width=100)
    assert document.picture(result.id).size == (100.0, 50.0)
    document.picture(result.id).resize(height=10)
    assert document.picture(result.id).size == (20.0, 10.0)


def test_media_is_shared_by_content_and_reaped_when_unused(lists_doc):
    document = Document.open(lists_doc)
    original = document.pictures()[0].image
    copy_ = document.insert_picture("p:5A00000E@0", original)
    assert media(document) == ["word/media/image1.png"]       # the same bytes: the same part
    document.picture(copy_.id).delete()
    assert media(document) == ["word/media/image1.png"]       # still used by the first
    document.pictures()[0].replace(png(8, 8))
    assert media(document) == ["word/media/image2.png"]       # the old one reaped
    assert document.pictures()[0].size == (36.0, 36.0)        # the frame kept
    assert set(check(Document.open(document.to_bytes()).package)) == set()


def test_replace_fitting_the_width(lists_doc):
    document = Document.open(lists_doc)
    document.pictures()[0].replace(png(40, 20), fit="width")
    assert document.pictures()[0].size == (36.0, 18.0)


def test_alt_text_and_delete(lists_doc):
    data = lists_doc.read_bytes()
    document = Document.open(data)
    picture = document.pictures()[0]
    picture.alt_text = "A different description"
    assert picture.alt_text == "A different description"
    picture.delete()
    assert document.pictures() == [] and media(document) == []
    assert document.paragraph("p:5A00000D").text == "A picture:  and text after it."
    assert set(check(Document.open(document.to_bytes()).package)) == set()
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)


def test_not_a_picture_is_refused(markup_doc):
    document = Document.open(markup_doc)
    with pytest.raises(EditError):
        document.insert_picture("p:1A2B3C17@0", b"not an image")
    with pytest.raises(EditError, match="field"):
        document.insert_picture("p:1A2B3C0B@10", png(4, 4))   # inside a simple field's result


def test_a_picture_in_a_header(markup_doc):
    document = Document.open(markup_doc)
    document.insert_picture("header1/p:4D5E6F01@0", png(10, 10))
    rels = document.package.relationships("word/header1.xml")
    assert any(rel.type.endswith("/image") for rel in rels.values())
    assert set(check(Document.open(document.to_bytes()).package)) - set(check(Document.open(markup_doc).package)) == set()


def test_pictures_on_every_fixture(docx_path):
    document = Document.open(docx_path)
    before = set(check(document.package))
    paragraph = next((p for p in document.paragraphs() if p.text.strip()), None)
    if paragraph is None:
        pytest.skip("no text")
    result = document.insert_picture(paragraph.range(0, 0), png(20, 10))
    for picture in document.pictures()[:2]:
        picture.replace(png(5, 5, (0, 200, 0)))
    document.picture(result.id).delete()
    after = set(check(Document.open(document.to_bytes()).package))
    assert after - before == set()
