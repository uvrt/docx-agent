"""E6: a new document (ROADMAP.md, Phase E6): what ``Document.new()`` writes, held to what Word
16.106 made for a new blank document (``tests/observations/e6-word.json``, recorded by
``tools/e6_probe.py``), its locale and page parameters, its properties, and the gates --
valid, saved, reopened and read back, undone and redone, laid out.

The new document is also in the corpus every phase's suite runs on (``conftest.py``,
``new/document``), so E1-E5's edit sets, ``insert_markdown`` and the oracle hold for it.
"""

from __future__ import annotations

import io
import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import pytest
from lxml import etree

from docx_agent import Document
from docx_agent.edit import blank
from docx_agent.edit.properties import read_app
from docx_agent.validate import check

from test_roundtrip import entries

OBSERVATIONS = json.loads((Path(__file__).parent / "observations" / "e6-word.json").read_text(encoding="utf-8"))
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CREATED = datetime(2026, 10, 4, 9, 30, tzinfo=timezone.utc)


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def tree(document: Document, part: str):
    return etree.fromstring(parts(document.to_bytes())[part])


# -- Word's new document --------------------------------------------------------------------------


def test_every_part_is_the_one_word_writes():
    """The probe compared ``Document.new()`` with Word's new document part by part (ids,
    rsids, dates, authors and the localised names normalised): all the same."""
    compared = OBSERVATIONS["new-compared"]
    assert compared and set(compared.values()) == {"same"}, compared


def test_the_package_is_words():
    word = OBSERVATIONS["new"]["package"]
    ours = parts(Document.new().to_bytes())
    assert sorted(ours) == word["parts"]
    types = etree.fromstring(ours["[Content_Types].xml"])
    overrides = sorted((n.get("PartName"), n.get("ContentType")) for n in types if n.get("PartName"))
    assert overrides == [tuple(pair) for pair in word["content_types"]]
    rels = etree.fromstring(ours["word/_rels/document.xml.rels"])
    assert sorted((n.get("Type").rsplit("/", 1)[1], n.get("Target")) for n in rels) == \
        [tuple(pair) for pair in word["relationships"]]


def _settings_facts(data: bytes) -> str:
    text = data.decode("utf-8")
    text = re.sub(r' xmlns:\w+="[^"]*"', "", text)
    text = re.sub(r"<w:rsids>.*?</w:rsids>|<w15:docId [^>]*/>|<w14:docId [^>]*/>", "", text)
    return re.sub(r"^<\?xml[^>]*>\s*", "", text)


def test_the_settings_are_words_with_its_locale_and_language():
    """With this Word's language and locale, the settings are Word's but for the window's
    view, which Word writes and docx-agent leaves out (a file without one opens in Print
    Layout)."""
    ours = _settings_facts(parts(blank.build(language="nl-BE", locale=blank.METRIC))["word/settings.xml"])
    word = OBSERVATIONS["new"]["settings"].replace('<w:view w:val="normal"/>', "")
    assert ours == word


def test_the_body_is_one_empty_paragraph_and_words_section():
    word = OBSERVATIONS["new"]["body"]
    document = Document.new()
    body = tree(document, "word/document.xml").find(W + "body")
    paragraph, section = list(body)
    assert paragraph.tag == W + "p" and len(paragraph) == 0
    assert paragraph.get("{http://schemas.microsoft.com/office/word/2010/wordml}paraId")
    ours = re.sub(r' xmlns:\w+="[^"]*"', "", etree.tostring(section, encoding="unicode"))
    assert ours == word[1]
    assert word[0] == "<w:p/>"


def test_the_styles_are_the_ones_a_new_document_has():
    document = Document.new()
    names = [s.name for s in document.styles]
    for name in ("Normal", "Default Paragraph Font", "Normal Table", "No List", "Title", "Subtitle", "Quote",
                 "List Paragraph", "Intense Emphasis", "Intense Quote", "Intense Reference", "Heading 1 Char",
                 "Title Char"):
        assert name in names
    assert [f"heading {n}" for n in range(1, 10)] == [n for n in names if n.startswith("heading ")]
    assert document.styles.default("paragraph").id == "Normal"
    assert document.styles.find("heading 1").id == "Heading1"
    styles = tree(document, "word/styles.xml")
    assert len(styles.findall(f"{W}latentStyles/{W}lsdException")) == 376
    defaults = styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr")
    assert defaults.find(W + "sz").get(W + "val") == "24"
    assert defaults.find(W + "lang").get(W + "val") == "en-US"


def test_the_theme_is_the_office_theme():
    theme = tree(Document.new(), "word/theme/theme1.xml")
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    assert theme.find(f"{a}themeElements/{a}fontScheme/{a}majorFont/{a}latin").get("typeface") == "Aptos Display"
    assert theme.find(f"{a}themeElements/{a}fontScheme/{a}minorFont/{a}latin").get("typeface") == "Aptos"
    assert theme.find(f"{a}themeElements/{a}clrScheme/{a}accent1/{a}srgbClr").get("val") == "156082"


def test_it_is_mode_15_and_valid():
    document = Document.new()
    assert document.compatibility_mode == 15
    assert check(document.package) == []
    assert document.package.kind == "docx"


# -- parameters ---------------------------------------------------------------------------------


def _section(document: Document) -> dict:
    return document.sections()[0].properties


@pytest.mark.parametrize("page, size", [("A4", (11906, 16838)), ("Letter", (12240, 15840)),
                                        ("A5", (8391, 11906)), ((300, 400), (6000, 8000))])
def test_page_sizes(page, size):
    document = Document.new(page=page)
    section = tree(document, "word/document.xml").find(f"{W}body/{W}sectPr/{W}pgSz")
    assert (int(section.get(W + "w")), int(section.get(W + "h"))) == size
    assert section.get(W + "orient") is None


def test_landscape_puts_the_longer_side_across():
    section = tree(Document.new(orientation="landscape"), "word/document.xml").find(f"{W}body/{W}sectPr/{W}pgSz")
    assert (section.get(W + "w"), section.get(W + "h"), section.get(W + "orient")) == ("16838", "11906", "landscape")


def test_the_locale_by_default_and_given():
    metric = tree(Document.new(), "word/document.xml").find(f"{W}body/{W}sectPr/{W}pgMar")
    assert metric.get(W + "top") == "1417" and metric.get(W + "header") == "708"
    settings = tree(Document.new(), "word/settings.xml")
    assert settings.find(W + "hyphenationZone").get(W + "val") == "425"
    assert settings.find(W + "decimalSymbol").get(W + "val") == ","
    letter = Document.new(page="Letter")
    margins = tree(letter, "word/document.xml").find(f"{W}body/{W}sectPr/{W}pgMar")
    assert margins.get(W + "top") == "1440" and margins.get(W + "header") == "720"
    settings = tree(letter, "word/settings.xml")
    assert settings.find(W + "hyphenationZone") is None
    assert settings.find(W + "defaultTabStop").get(W + "val") == "720"
    assert settings.find(W + "listSeparator").get(W + "val") == ","
    custom = Document.new(locale=blank.Locale(margin=1000, hyphenation_zone=None))
    assert tree(custom, "word/document.xml").find(f"{W}body/{W}sectPr/{W}pgMar").get(W + "left") == "1000"


def test_language():
    document = Document.new(language="nl-NL")
    assert tree(document, "word/styles.xml").find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}lang").get(
        W + "val") == "nl-NL"
    assert tree(document, "word/settings.xml").find(W + "themeFontLang").get(W + "val") == "nl-NL"


@pytest.mark.parametrize("value", ["B5", (10, 10), "landscape"])
def test_bad_page_sizes_are_refused(value):
    with pytest.raises(ValueError):
        Document.new(page=value)
    with pytest.raises(ValueError):
        Document.new(orientation="sideways")


def test_title_author_and_dates():
    document = Document.new(title="Quarterly report", author="Claude", created=CREATED)
    properties = document.properties
    assert properties["title"] == "Quarterly report"
    assert properties["author"] == properties["last_modified_by"] == "Claude"
    assert properties["created"] == properties["modified"] == CREATED
    assert properties["revision"] == 1
    assert properties["statistics"] == {"pages": 1, "words": 0, "characters": 0, "lines": 0, "paragraphs": 0,
                                        "characters_with_spaces": 0}
    assert "author" not in Document.new().properties


def test_deterministic_given_the_date():
    assert Document.new(created=CREATED).to_bytes() == Document.new(created=CREATED).to_bytes()


# -- the gates ------------------------------------------------------------------------------------


def test_edit_save_reopen_undo_redo(tmp_path):
    document = Document.new(created=CREATED)
    original = document.to_bytes()
    first = document.paragraphs()[0]
    first.set_text("A first paragraph.")
    document.insert_paragraph("A heading", after=first.id, style="Heading 1")
    document.insert_markdown("Some *emphasis* and a list:\n\n- one\n- two\n")
    edited = document.to_bytes()
    path = tmp_path / "new.docx"
    document.save(path)
    saved = Document.open(path)
    assert [p.text for p in saved.paragraphs()][:2] == ["A first paragraph.", "A heading"]
    assert saved.paragraphs()[1].style_name == "heading 1"
    assert check(saved.package) == []
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(original)
    while document.redo():
        pass
    assert entries(document.to_bytes()) == entries(edited)


def test_the_layout_is_whole_headings_included():
    """Aptos Display, the heading face, is one of Office's cloud fonts: docx2svg searches
    Office's cloud-font cache itself, so where Word has downloaded it the layout goes to the
    end."""
    from docx2svg.fonts import cloud_font_dirs

    document = Document.new()
    document.insert_markdown("# A heading in Aptos Display\n\nBody text in Aptos.\n")
    layout = document.layout()
    if not any(Path(d).name == "Aptos Display" for d in cloud_font_dirs()):
        pytest.skip("Office has not downloaded Aptos Display on this machine")
    assert layout.stopped is None and layout.page_count == 1
    for paragraph in document.paragraphs():
        assert layout.where(paragraph.id)


def test_save_brings_apps_statistics_up_to_date(tmp_path):
    document = Document.new()
    document.insert_markdown("One two three.\n\nFour five.\n")
    document.layout()
    document.save(tmp_path / "counted.docx")
    stats = read_app(Document.open(tmp_path / "counted.docx").package)
    assert stats == {"pages": 1, "words": 5, "characters": 21, "lines": 2, "paragraphs": 2,
                     "characters_with_spaces": 24}
    # Without a layout of the state saved, the page and line counts are left as they were.
    document.insert_markdown("Six.\n")
    document.save(tmp_path / "counted2.docx")
    stats = read_app(Document.open(tmp_path / "counted2.docx").package)
    assert stats["words"] == 6 and stats["lines"] == 0 and stats["pages"] == 1
    # The edits never touch app.xml: to_bytes() is the package as it is.
    assert read_app(Document.open(document.to_bytes()).package)["words"] == 0


def test_a_saved_document_that_did_not_change_keeps_its_app_part(tmp_path):
    source = Path(__file__).parent / "fixtures" / "samplelib" / "sample-simple.docx"
    document = Document.open(source)
    document.save(tmp_path / "same.docx")
    assert parts((tmp_path / "same.docx").read_bytes())["docProps/app.xml"] == \
        parts(source.read_bytes())["docProps/app.xml"]
