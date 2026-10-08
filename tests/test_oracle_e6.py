"""Word itself on E6's authoring, opt-in (``pytest -m oracle``; ROADMAP.md, Phase E6, "Done
when").

* **A new document is Word's**: Word's own new document with text typed and ``Document.new()``
  with the same text (in this Word's language and locale) are drawn the same by docx2svg,
  line for line and glyph for glyph, and Word's PDFs of the two have the same characters in
  the same faces, sizes and colours.
* **Everything made exports unprompted**: new documents (blank, Letter, landscape, with
  Markdown), documents from every template, copied content (every style policy), and every
  fixture upgraded to mode 15.  (E1-E5's edit sets on a new document export through their
  own oracle suites: the new document is in the corpus, ``new/document``.)
* **Copied blocks keep their look** with Keep Source Formatting: in Word's PDF each copied
  paragraph's characters have the faces, sizes, weights and colours they have in the
  source's PDF.
* **Re-saving keeps everything**: Word's re-save of a new document keeps every style
  definition docx-agent wrote (Word gives built-in styles its interface's ids -- recorded
  below); of a document from a template, its styles, body, header and footer; of copied
  content, its text, notes, comments, renamed bookmarks and separate lists; of an upgraded
  document, mode 15 and the compatibility settings.
"""

from __future__ import annotations

import io
import re
import sys
import warnings
import zipfile
from collections import Counter
from pathlib import Path

import pytest
from lxml import etree

import oracle
from conftest import fixture_id, fixture_paths
from docx_agent import Document
from docx_agent.edit import blank

from test_oracle_e3 import _plain

pytestmark = pytest.mark.oracle

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import e6_probe  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TEMPLATES = ROOT / "tests" / "fixtures" / "generated" / "templates"
TYPED = ["The first paragraph Word typed into its own new document.",
         "A second paragraph, long enough to run across the width of the page more than once, so that "
         "where its lines break, and how far apart they are, says whether the two documents are laid out "
         "alike: the same face, the same size, the same spacing after and between the lines.",
         "A third and last one."]
TYPED_SCRIPT = e6_probe._script("set d to make new document\n      insert text \"" + "\" & return & \"".join(TYPED)
                                + "\" at end of text object of d")
MARKDOWN = "# A heading\n\nBody text with *emphasis* and **strength**.\n\n- one\n- two\n\n> A quote.\n"


@pytest.fixture(scope="module")
def word():
    try:
        with oracle.session() as session:
            yield session
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _export(word, data: bytes, name: str) -> Path:
    outcome = word.export_pdf(data, name=name)
    assert outcome, f"Word did not export {name}: {outcome.outcome} {outcome.detail}"
    return outcome.path


def _resave(word, data: bytes, name: str) -> bytes:
    outcome = word.resave(data, name=name)
    assert outcome, f"Word did not save {name}: {outcome.outcome} {outcome.detail}"
    return outcome.path.read_bytes()


def _text(pdf: Path) -> str:
    return " ".join(" ".join(oracle.pdf_pages(pdf)).split())


# -- a new document is Word's ----------------------------------------------------------------------


def _ours_typed() -> bytes:
    document = Document.new(language="nl-BE", locale=blank.METRIC)
    first = document.paragraphs()[0]
    first.set_text(TYPED[0])
    previous = first.id
    for text in TYPED[1:]:
        previous = document.insert_paragraph(text, after=previous).id
    return document.to_bytes()


def _drawn(data: bytes) -> list[tuple]:
    """What docx2svg draws on each page, without the paragraphs' ids (Word wrote none)."""
    layout = Document.open(data).layout()
    assert layout.stopped is None
    out = []
    for index, page in enumerate(layout.pages):
        content = layout._content(index)
        lines = tuple(line[1:] for line in content[2])
        out.append((content[0], content[1], lines) + content[3:])
    return out


def test_a_new_document_is_drawn_as_words_own(word):
    made = word.run_script(TYPED_SCRIPT, None, name="e6-oracle-new", tag="typed")
    assert made, f"Word did not make its document: {made.outcome} {made.detail}"
    theirs = made.path.read_bytes()
    ours = _ours_typed()
    assert [p.text for p in Document.open(theirs).paragraphs()] == TYPED
    assert _drawn(ours) == _drawn(theirs)
    mine = oracle.pdf_chars(_export(word, ours, "e6-new-ours"))
    words = oracle.pdf_chars(_export(word, theirs, "e6-new-words"))
    assert len(mine) == len(words) == 1
    strip = lambda chars: [(c.char, c.size, c.color, c.font.split("+")[-1]) for c in chars if c.char.strip()]  # noqa: E731
    assert strip(mine[0]) == strip(words[0])


# -- everything made exports ------------------------------------------------------------------------


def _new_documents() -> dict[str, bytes]:
    out = {"blank": Document.new().to_bytes(), "letter": Document.new(page="Letter").to_bytes(),
           "landscape": Document.new(orientation="landscape", language="nl-NL").to_bytes()}
    document = Document.new(title="E6 oracle", author="E6 Agent")
    document.insert_markdown(MARKDOWN, at=f"replace:{document.paragraphs()[0].id}")
    out["markdown"] = document.to_bytes()
    return out


@pytest.mark.parametrize("name", ["blank", "letter", "landscape", "markdown"])
def test_word_exports_new_documents(word, name):
    data = _new_documents()[name]
    pdf = _export(word, data, f"e6-new-{name}")
    if name == "markdown":
        text = _text(pdf)
        for visible in ("A heading", "emphasis", "one", "two", "A quote."):
            assert visible in text


TEMPLATE_SOURCES = sorted(TEMPLATES.glob("*.dot*")) + [
    ROOT / "tests" / "fixtures" / "generated" / "markdown" / "dutch-template.docx",
    ROOT / "tests" / "fixtures" / "generated" / "lists-and-styles.docx"]


def _from_template(path: Path, keep: bool = True) -> bytes:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        document = Document.new(template=path, keep_content=keep, author="E6 Agent")
    document.insert_markdown("## Added from a template\n\nText added.\n")
    return document.to_bytes()


@pytest.mark.parametrize("keep", [True, False], ids=["content", "empty"])
@pytest.mark.parametrize("path", TEMPLATE_SOURCES, ids=lambda p: p.name)
def test_word_exports_documents_from_templates(word, path, keep):
    data = _from_template(path, keep)
    text = _text(_export(word, data, f"e6-template-{path.stem}-{keep}"))
    assert "Added from a template" in text and "Text added." in text
    if keep and path.parent == TEMPLATES:
        assert "Template heading" in text and "Template header" in text


def _copied(policy: str) -> bytes:
    source = Document.open(e6_probe.copy_document("source"))
    target = Document.open(e6_probe.copy_document("destination"))
    paragraphs = source.paragraphs()
    target.copy_blocks(source, f"{paragraphs[0].id}..{paragraphs[-1].id}", styles=policy)
    return target.to_bytes()


@pytest.mark.parametrize("policy", ["use_destination", "keep_source", "merge"])
def test_word_exports_copied_content(word, policy):
    text = _text(_export(word, _copied(policy), f"e6-copy-{policy}"))
    for visible in ("Source heading", "Source brand paragraph", "Source list item two", "Source footnote.",
                    "Source value", "Cell two", "Source last paragraph.", "Destination end."):
        assert visible in text, visible


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_word_exports_every_fixture_copied_into_a_new_document(word, path):
    source = Document.open(path)
    body = source.package.tree(source.package.document_part()).find(W + "body")
    paragraphs = [p for p in source.paragraphs() if p._element.getparent() is body]
    if not paragraphs:
        pytest.skip("no paragraph in the body")
    target = Document.new()
    target.copy_blocks(source, f"{paragraphs[0].id}..{paragraphs[-1].id}")
    _export(word, target.to_bytes(), f"e6-copied-{path.stem}")


def _look(chars: list, text: str) -> list[tuple]:
    found = oracle.find_in_pdf(chars, text)
    assert found, text
    return [(c.char, c.size, c.color, re.sub(r"^[A-Z]{6}\+", "", c.font), c.stroked) for c in found
            if c.char.strip()]


def test_copied_blocks_keep_their_look_in_word(word):
    """Keep Source Formatting: each copied paragraph is drawn by Word as the source draws it."""
    source = e6_probe.copy_document("source")
    theirs = oracle.pdf_chars(_export(word, source, "e6-look-source"))
    ours = oracle.pdf_chars(_export(word, _copied("keep_source"), "e6-look-copy"))
    for text in ("Source heading", "Source brand paragraph with accented words.", "Source list item one",
                 "Source only style.", "Source last paragraph."):
        assert _look(ours, text) == _look(theirs, text), text


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_word_exports_every_fixture_upgraded(word, path):
    document = Document.open(path)
    result = document.upgrade_to_modern()
    if not result.changed:
        pytest.skip("already in mode 15")
    data = document.to_bytes()
    _export(word, data, f"e6-upgraded-{path.stem}")
    saved = Document.open(_resave(word, data, f"e6-upgraded-{path.stem}"))
    assert saved.compatibility_mode == 15


# -- re-saving -----------------------------------------------------------------------------------


def _styles(data: bytes) -> dict[str, etree._Element]:
    root = etree.fromstring(_parts(data)["word/styles.xml"])
    return {s.find(W + "name").get(W + "val"): s for s in root.findall(W + "style")}


#: English Word's names of the linked styles, as Dutch Word renames them on saving.
LOCALISED = {"Heading 1 Char": "Kop 1 Char", "Title Char": "Titel Char", "Subtitle Char": "Ondertitel Char",
             "Quote Char": "Citaat Char", "Intense Quote Char": "Duidelijk citaat Char",
             **{f"Heading {n} Char": f"Kop {n} Char" for n in range(2, 10)}}


def _definition(style: etree._Element, ids: dict[str, str]) -> str:
    node = etree.fromstring(etree.tostring(style))
    node.attrib.pop(W + "styleId", None)
    for child in list(node):
        if child.tag in (W + "rsid", W + "name"):
            node.remove(child)
        elif child.tag in (W + "basedOn", W + "next", W + "link"):
            child.set(W + "val", ids.get(child.get(W + "val"), child.get(W + "val")))
    return etree.tostring(node, method="c14n").decode()


def test_word_keeps_every_style_a_new_document_has(word):
    """Word's re-save changes no style definition docx-agent wrote: only ids (its interface's)
    and the names of linked styles (Dutch Word's ``Kop 1 Char``) -- recorded."""
    ours = _new_documents()["markdown"]
    theirs = _resave(word, ours, "e6-new-resave")
    mine, words = _styles(ours), _styles(theirs)
    my_names = {s.get(W + "styleId"): name for name, s in mine.items()}
    their_names = {s.get(W + "styleId"): name for name, s in words.items()}
    for name, style in mine.items():
        their_name = name if name in words else LOCALISED.get(name, name)
        assert their_name in words, name
        mapped_mine = {k: LOCALISED.get(v, v) for k, v in my_names.items()}
        assert _definition(style, mapped_mine) == _definition(words[their_name], their_names), name
    saved = Document.open(theirs)
    assert saved.compatibility_mode == 15
    assert [p.text for p in saved.paragraphs()] == [p.text for p in Document.open(ours).paragraphs()]
    assert [p.id for p in saved.paragraphs()] == [p.id for p in Document.open(ours).paragraphs()]


def test_word_keeps_a_document_from_a_template(word):
    data = _from_template(TEMPLATES / "brand.dotx")
    saved = Document.open(_resave(word, data, "e6-template-resave"))
    assert saved.package.kind == "docx"
    assert [p.text for p in saved.paragraphs()] == [p.text for p in Document.open(data).paragraphs()]
    names = {s.name for s in saved.styles}
    assert {"Brand Body", "Brand Mark"} <= names
    section = saved.sections()[0]
    assert section.header("default").paragraphs[0].text == "Template header"
    assert section.footer("default").paragraphs[0].text == "Template footer"


def test_word_keeps_copied_content(word):
    data = _copied("use_destination")
    saved = Document.open(_resave(word, data, "e6-copy-resave"))
    ours = Document.open(data)
    assert [p.text for p in saved.paragraphs()] == [p.text for p in ours.paragraphs()]
    assert sorted(b.name for b in saved.bookmarks()) == ["Shared", "Shared_1"]
    assert sorted(c.text for c in saved.comments()) == ["Destination comment", "Source comment"]
    assert [n.text.strip() for n in saved.notes("footnote")] == ["Destination footnote.", "Source footnote."]
    numbers = [b["list"]["label"] for b in saved.state()["blocks"] if "list item" in b.get("t", "")]
    assert numbers == ["1.", "2.", "1.", "2."]
    assert [p.id for p in saved.paragraphs()] == [p.id for p in ours.paragraphs()]


def test_an_attached_template_word_may_read_opens_unprompted(word):
    """Attached (``attach=True``), a template inside Word's sandbox -- the Office group
    container -- opens without a question; outside it Word asks for access, which is why
    docx-agent attaches none by default."""
    staged = oracle.ORACLE_DIR / "e6-attached-brand.dotx"
    staged.write_bytes((TEMPLATES / "brand.dotx").read_bytes())
    document = Document.new(template=staged, attach=True)
    assert "Template heading" in _text(_export(word, document.to_bytes(), "e6-attached"))
