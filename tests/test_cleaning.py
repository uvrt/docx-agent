"""save_document's clean (round 16, F11): what Word keeps but never draws, taken out of the
file saved -- on made-up packages built here (Word's measurement: ``docx_agent.edit.cleaning``)."""

from __future__ import annotations

import io
import re
import zipfile

from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.edit.cleaning import clean
from docx_agent.tools import FORMAT, GROUPS, TOOLS

TEMPLATE = "file:///C:/Users/someone/Templates/Normal.dotm"


def _parts(data: bytes) -> dict[str, bytes]:
    archive = zipfile.ZipFile(io.BytesIO(data))
    return {name: archive.read(name) for name in archive.namelist()}


def _zip(parts: dict[str, bytes]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return out.getvalue()


def bloated(*, drawn_div: bool = False) -> bytes:
    document = Document.new()
    first = document.paragraphs()[0]
    first.text = "Heading"
    first.style = "Heading 1"
    document.paragraph(first.id).insert_after("Body text in the first section.")
    parts = _parts(document.to_bytes())
    body = parts["word/document.xml"].decode()
    body = body.replace("<w:p ", '<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3" ')
    body = body.replace("<w:r>", '<w:r w:rsidRPr="00D4E5F6">')
    body = re.sub(r"(<w:pStyle [^>]*/>)", r'\1<w:divId w:val="7"/>', body)
    parts["word/document.xml"] = body.encode()
    settings = parts["word/settings.xml"].decode()
    settings = settings.replace("<w:defaultTabStop", '<w:attachedTemplate r:id="rId1"/><w:defaultTabStop', 1)
    settings = settings.replace("<m:mathPr>", '<w:rsids><w:rsidRoot w:val="00A1B2C3"/><w:rsid w:val="00A1B2C3"/>'
                                '<w:rsid w:val="00D4E5F6"/></w:rsids><m:mathPr>', 1)
    parts["word/settings.xml"] = settings.encode()
    parts["word/_rels/settings.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/'
        'package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        f'relationships/attachedTemplate" Target="{TEMPLATE}" TargetMode="External"/></Relationships>').encode()
    margin = "1440" if drawn_div else "0"
    web = parts["word/webSettings.xml"].decode().replace(
        "<w:optimizeForBrowser/>", f'<w:divs><w:div w:id="7"><w:marLeft w:val="{margin}"/><w:marRight w:val="0"/>'
        '<w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divs><w:optimizeForBrowser/>', 1)
    parts["word/webSettings.xml"] = web.encode()
    parts["[trash]/0000.dat"] = b"\x01" * 4096
    return _zip(parts)


def test_clean_takes_out_what_word_never_draws_and_says_how_much():
    data = bloated()
    out, report = clean(data)
    parts = _parts(out)
    assert "[trash]/0000.dat" not in parts
    assert report["orphaned_parts"] == {"count": 1, "bytes": 4096, "parts": ["[trash]/0000.dat"]}
    assert report["attached_template"]["target"] == TEMPLATE
    assert b"attachedTemplate" not in parts["word/settings.xml"]
    assert b"attachedTemplate" not in parts["word/_rels/settings.xml.rels"]
    assert report["rsids"]["attributes"] == len(re.findall(rb"w:rsid\w*=", _parts(data)["word/document.xml"]))
    assert report["rsids"]["listed"] == 3 and report["rsids"]["bytes"] > 0
    assert not any(re.search(rb"w:rsid", value) for name, value in parts.items() if name.endswith(".xml"))
    assert report["web_divs"]["count"] == 1
    assert b"w:divs" not in parts["word/webSettings.xml"] and b"divId" not in parts["word/document.xml"]
    assert report["bytes"] == {"before": len(data), "after": len(out)}
    reopened = Document.open(out)
    assert [p.text for p in reopened.paragraphs()] == ["Heading", "Body text in the first section."]


def test_a_division_word_draws_is_kept_and_a_clean_file_cleans_to_itself():
    """A division with a margin is drawn (measured: 8 of 9 pages differ without it), so the
    tree stays whole; cleaning what is clean changes no byte and reports only the sizes."""
    out, report = clean(bloated(drawn_div=True))
    assert "web_divs" not in report and b"w:divs" in _parts(out)["word/webSettings.xml"]
    plain = Document.new().to_bytes()
    again, nothing = clean(plain)
    assert again == plain and nothing == {"bytes": {"before": len(plain), "after": len(plain)}}


def test_save_document_cleans_the_file_and_leaves_the_document():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session()
        d = session.open(bloated(), "old.docx")
        result = box.dispatch(session, "save_document", {"doc": d, "name": "new.docx", "format": "docx",
                                                         "clean": True})
        assert result.ok and result.data["cleaned"]["orphaned_parts"]["count"] == 1
        assert "cleaned:" in result.summary
        saved = session.take_outputs()[-1].data
        assert "[trash]/0000.dat" not in _parts(saved)
        assert session.entry(d).document.package.has_part("[trash]/0000.dat")
        refused = box.dispatch(session, "save_document", {"doc": d, "name": "new.md", "format": "markdown",
                                                          "clean": True})
        assert not refused.ok and refused.error.field == "clean"
