"""Round 17, P5: a table's alt text (``w:tblCaption``, ``w:tblDescription``) and a content
control's tag and title (``w:tag``, ``w:alias``) in ``word_read``'s block comments and in
``describe``, so a sample block can be found by what names it without reading the XML."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

CONSTRUCTS = Path(__file__).parent / "fixtures" / "generated" / "markdown" / "constructs.docx"
AGREEMENT = Path(__file__).parent / "fixtures" / "generated" / "pilot" / "agreement-summary.docx"
_LOOK = ('<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" '
         'w:noHBand="0" w:noVBand="1"/>')


def _patched(tmp_path: Path) -> Path:
    """The constructs fixture with alt text on its first table (a title and a description),
    a title only on its second, and a tag and title on the inline text control cc:301."""
    source = zipfile.ZipFile(CONSTRUCTS)
    xml = source.read("word/document.xml").decode()
    assert xml.count(_LOOK) >= 2
    xml = xml.replace(_LOOK, _LOOK + '<w:tblCaption w:val="sample:record-table"/>'
                      '<w:tblDescription w:val="One row per item, &#10;a total at the foot."/>', 1)
    first = xml.index("sample:record-table")
    second = xml.index(_LOOK, first)
    xml = xml[:second] + _LOOK + '<w:tblCaption w:val="Second -- table"/>' + xml[second + len(_LOOK):]
    xml = xml.replace('<w:sdtPr><w:id w:val="301"/>',
                      '<w:sdtPr><w:alias w:val="Name"/><w:tag w:val="sample:name"/><w:id w:val="301"/>', 1)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = xml.encode() if item.filename == "word/document.xml" else source.read(item.filename)
            target.writestr(item, data)
    path = tmp_path / "tagged.docx"
    path.write_bytes(out.getvalue())
    return path


def test_table_comments_carry_the_alt_text(tmp_path):
    document = Document.open(_patched(tmp_path))
    first, second, third = document.tables()[:3]
    markdown = document.to_markdown()
    assert (f"<!-- {first.id} table style: Table Grid alt: sample:record-table "
            "description: One row per item, a total at the foot. -->") in markdown
    # "--" cannot be in a comment: written as comment text is, as the style is.
    assert f"<!-- {second.id} table style: Table Grid alt: Second - - table -->" in markdown
    assert f"<!-- {third.id} table style: Table Grid -->" in markdown
    assert first.alt_text == {"alt": "sample:record-table", "description": "One row per item, \na total at the foot."}
    assert third.alt_text == {}


def test_an_inline_content_control_carries_its_tag_and_title(tmp_path):
    markdown = Document.open(_patched(tmp_path)).to_markdown()
    assert "Name: <!-- cc:301 text tag: sample:name title: Name -->Grace Hopper<!-- /cc:301 -->" in markdown
    # A block content control as before.
    assert "<!-- cc:303 rich-text tag: clause title: Clause -->" in markdown


def test_describe_lists_alt_text_and_tagged_content_control(tmp_path):
    path = _patched(tmp_path)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session = box.session()
        d = session.open(path.read_bytes(), path.name)
        result = box.dispatch(session, "describe", {"doc": d})
        assert result.ok
        tables = result.data["tables"]
        assert tables[0]["alt"] == "sample:record-table"
        assert tables[0]["description"].startswith("One row per item")
        assert tables[1]["alt"] == "Second -- table" and "description" not in tables[1]
        assert "alt" not in tables[2]
        assert result.data["tagged"] == [
            {"id": "cc:301", "kind": "text", "tag": "sample:name", "title": "Name"},
            {"id": "cc:303", "kind": "rich-text", "tag": "clause", "title": "Clause"}]
        # Without any, the key is not there.
        plain = session.open(AGREEMENT.read_bytes(), AGREEMENT.name)
        assert "tagged" not in box.dispatch(session, "describe", {"doc": plain}).data
