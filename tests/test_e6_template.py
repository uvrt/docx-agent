"""E6: a document from a template, and a document saved as one (ROADMAP.md, Phase E6), held to
what Word did with File > New from a template and Save As template
(``tests/observations/e6-word.json``: ``template-dotx``, ``template-docx``, ``as-template``).

Templates: ``tests/fixtures/generated/templates`` (``tools/make_template_fixtures.py``) and
the corpus's Dutch templates used as templates.
"""

from __future__ import annotations

import io
import json
import warnings
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import pytest
from lxml import etree

from docx_agent import Document
from docx_agent.edit.authoring import TemplateOpened
from docx_agent.edit.authoring import MacrosDropped, ModeUpgraded
from docx_agent.validate import check

from test_roundtrip import entries

HERE = Path(__file__).parent
TEMPLATES = HERE / "fixtures" / "generated" / "templates"
BRAND = TEMPLATES / "brand.dotx"
OBSERVATIONS = json.loads((HERE / "observations" / "e6-word.json").read_text(encoding="utf-8"))
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CREATED = datetime(2026, 10, 4, 9, 30, tzinfo=timezone.utc)
DOCUMENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
TEMPLATE_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
TEMPLATE_SOURCES = [BRAND, TEMPLATES / "brand-macros.dotm", TEMPLATES / "brand-mode14.dotx",
                    HERE / "fixtures" / "generated" / "markdown" / "dutch-template.docx",
                    HERE / "fixtures" / "generated" / "lists-and-styles.docx"]


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def quietly(**options) -> Document:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return Document.new(**options)


def _texts(document: Document) -> list[str]:
    return [p.text for p in document.paragraphs()]


def test_word_keeps_the_templates_body_styles_sections_and_headers():
    """What Word did: the body, the styles, the section (Letter, its top margin), the
    header and footer, the list, ``evenAndOddHeaders`` -- all kept; the main part a
    document's."""
    word = OBSERVATIONS["template-dotx"]
    assert len([b for b in word["body"] if b.startswith("<w:p")]) == 4
    assert "Template heading" in word["body"][0] and 'w:w="12240"' in word["body"][-1]
    assert ["/word/document.xml", DOCUMENT_TYPE] in word["package"]["content_types"]
    assert "<w:evenAndOddHeaders/>" in word["settings"] and "<w:attachedTemplate" in word["settings"]
    assert "Brand Body" in word["styles"] and 'w:val="C00000"' in word["styles"]["heading 1"]
    assert OBSERVATIONS["template-docx"]["properties"]["docProps/app.xml"].count("<Template>Normal.dotm") == 1
    assert "<w:attachedTemplate" not in OBSERVATIONS["template-docx"]["settings"]


def test_from_a_dotx_as_word_makes_one():
    with pytest.warns(TemplateOpened):
        template = Document.open(BRAND)
    document = Document.new(template=BRAND, author="Claude", created=CREATED)
    assert document.package.kind == "docx"
    assert document.package.content_type(document.package.document_part()) == DOCUMENT_TYPE
    assert _texts(document) == _texts(template)
    assert document.compatibility_mode == 15
    assert check(document.package) == []
    # Styles, numbering, header, footer and the section kept.
    for name in ("Brand Body", "Brand Mark", "heading 1"):
        assert document.styles.find(name) is not None
    styles = etree.fromstring(parts(document.to_bytes())["word/styles.xml"])
    heading = next(s for s in styles.findall(W + "style") if s.find(W + "name").get(W + "val") == "heading 1")
    assert heading.find(f"{W}rPr/{W}color").get(W + "val") == "C00000"
    section = document.sections()[0]
    assert section.header("default") is not None and section.footer("default") is not None
    assert document.paragraphs()[2].list is not None
    root = etree.fromstring(parts(document.to_bytes())["word/document.xml"])
    size = root.find(f"{W}body/{W}sectPr/{W}pgSz")
    assert size.get(W + "w") == "12240"
    settings = etree.fromstring(parts(document.to_bytes())["word/settings.xml"])
    assert settings.find(W + "evenAndOddHeaders") is not None


def test_a_template_is_attached_only_when_asked():
    """Word attached the .dotx (an external relationship to its path) and named it in
    app.xml; docx-agent does so with ``attach=True`` only -- attached, Word for Mac asks for
    access to the template on opening (the oracle's export blocked on it)."""
    assert "<w:attachedTemplate" in OBSERVATIONS["template-dotx"]["settings"]
    document = Document.new(template=BRAND, attach=True)
    data = parts(document.to_bytes())
    settings = etree.fromstring(data["word/settings.xml"])
    assert settings.find(W + "attachedTemplate") is not None
    rels = etree.fromstring(data["word/_rels/settings.xml.rels"])
    target = rels[0].get("Target")
    assert target.startswith("file:///") and target.endswith("/brand.dotx") and rels[0].get("TargetMode") == "External"
    assert b"<Template>brand.dotx</Template>" in data["docProps/app.xml"]
    for document in (Document.new(template=BRAND), Document.new(template=BRAND.read_bytes()),
                     Document.new(template=HERE / "fixtures" / "generated" / "markdown" / "dutch-template.docx")):
        data = parts(document.to_bytes())
        assert etree.fromstring(data["word/settings.xml"]).find(W + "attachedTemplate") is None
        assert b"<Template>Normal.dotm</Template>" in data["docProps/app.xml"]
    with pytest.raises(ValueError):
        Document.new(template=BRAND.read_bytes(), attach=True)


def test_properties_start_again_as_word_starts_them():
    """Word kept the title, subject and keywords, made the author and last modifier its
    user, revision 1, created and modified now, and kept the company."""
    word = OBSERVATIONS["template-dotx"]["properties"]["docProps/core.xml"]
    assert "<dc:title>Template Title</dc:title>" in word and "<cp:revision>1</cp:revision>" in word
    assert "<dc:creator>{user}</dc:creator>" in word
    document = Document.new(template=BRAND, author="Claude", created=CREATED)
    properties = document.properties
    assert properties["title"] == "Template Title" and properties["subject"] == "Template subject"
    assert properties["keywords"] == "template keywords"
    assert properties["author"] == properties["last_modified_by"] == "Claude"
    assert properties["revision"] == 1 and properties["created"] == properties["modified"] == CREATED
    assert b"<Company>Template Company</Company>" in parts(document.to_bytes())["docProps/app.xml"]
    assert Document.new(template=BRAND, title="Mine").properties["title"] == "Mine"
    assert "author" not in Document.new(template=BRAND).properties


def test_keep_content_false_drops_the_body_keeps_the_rest():
    document = Document.new(template=BRAND, keep_content=False)
    paragraphs = document.paragraphs()
    assert len(paragraphs) == 1 and paragraphs[0].text == "" and paragraphs[0].style is None
    assert not paragraphs[0].volatile
    assert document.sections()[0].header("default") is not None
    assert document.styles.find("Brand Body") is not None
    assert check(document.package) == []


def test_keep_content_false_drops_notes_and_comments():
    template = HERE / "fixtures" / "generated" / "ids-and-markup.docx"
    document = Document.new(template=template, keep_content=False)
    assert check(document.package) == []
    assert document.comments() == []
    assert [n for n in document.notes()] == []


def test_a_macro_enabled_template_gives_a_docx_without_its_macros():
    with pytest.warns(MacrosDropped):
        document = Document.new(template=TEMPLATES / "brand-macros.dotm")
    names = document.package.part_names
    assert not any("vba" in name.lower() for name in names)
    assert document.package.kind == "docx"
    assert b"vbaProject" not in parts(document.to_bytes())["[Content_Types].xml"]
    # Not attached: its macros would come back from it (Word asks first; the oracle's
    # export blocked on the question).
    assert "word/_rels/settings.xml.rels" not in names
    assert check(document.package) == []


def test_a_template_below_mode_15_gives_a_converted_document():
    with pytest.warns(ModeUpgraded):
        document = Document.new(template=TEMPLATES / "brand-mode14.dotx")
    assert document.compatibility_mode == 15
    assert check(document.package) == []


def test_page_orientation_and_language_replace_the_templates():
    document = Document.new(template=BRAND, page="A4", orientation="landscape", language="de-DE")
    root = etree.fromstring(parts(document.to_bytes())["word/document.xml"])
    size = root.find(f"{W}body/{W}sectPr/{W}pgSz")
    assert (size.get(W + "w"), size.get(W + "h"), size.get(W + "orient")) == ("16838", "11906", "landscape")
    styles = etree.fromstring(parts(document.to_bytes())["word/styles.xml"])
    assert styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}lang").get(W + "val") == "de-DE"


@pytest.mark.parametrize("template", TEMPLATE_SOURCES, ids=lambda p: p.name)
def test_every_template_gives_a_valid_editable_document(template, tmp_path):
    document = quietly(template=template, created=CREATED)
    assert document.package.kind == "docx" and document.compatibility_mode == 15
    assert check(document.package) == []
    original = document.to_bytes()
    first = document.paragraphs()[0]
    document.insert_paragraph("Added after the template's first paragraph.", after=first.id)
    document.insert_markdown("## A heading\n\nText.\n")
    edited = document.to_bytes()
    document.save(tmp_path / "made.docx")
    saved = Document.open(tmp_path / "made.docx")
    assert "Added after the template's first paragraph." in _texts(saved)
    assert check(saved.package) == []
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(original)
    while document.redo():
        pass
    assert entries(document.to_bytes()) == entries(edited)
    assert document.layout().pages


def test_save_as_template_changes_the_main_content_type_only(tmp_path):
    """Word's Save As template changed the main part's content type (and nothing else of
    the document's own)."""
    word = OBSERVATIONS["as-template"]["package"]["content_types"]
    assert ["/word/document.xml", TEMPLATE_TYPE] in word
    document = Document.new(created=CREATED)
    document.insert_markdown("# Letterhead\n\nBody.\n")
    before = document.to_bytes()
    document.save_as_template(tmp_path / "mine.dotx")
    assert document.to_bytes() == before
    saved = parts((tmp_path / "mine.dotx").read_bytes())
    ours = parts(before)
    assert set(saved) == set(ours)
    changed = {name for name in saved if saved[name] != ours[name]}
    assert changed <= {"[Content_Types].xml", "docProps/app.xml"}
    with pytest.warns(TemplateOpened):
        template = Document.open(tmp_path / "mine.dotx")
    assert template.package.kind == "dotx"
    again = Document.new(template=tmp_path / "mine.dotx")
    assert again.package.kind == "docx" and [p.text for p in again.paragraphs()] == [
        p.text for p in document.paragraphs()]
