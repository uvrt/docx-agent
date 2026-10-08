"""Word itself on E4's structure, opt-in (``pytest -m oracle``; ROADMAP.md, Phase E4, "Done
when").

* Every fixture with E4's edit set (``e4_edits.e4_edit_set``: a section break, page
  numbers in a footer, a first-page header, a footnote, an endnote, a page reference and a
  table of contents) **exports from Word unprompted**, its PDF showing the notes, the
  header and the table of contents.
* **A section break inserted mid-document** gives docx2svg and Word the same page count
  and the same section on each page (each section's footer names it).
* **Headers and footers show on the right pages**: a first-page header on the first page,
  an even footer on even pages, the default one on the others -- in Word's PDF and in
  docx2svg's pages alike.
* **Tables of contents and page references agree with Word**: Word opens the document,
  updates every field itself and saves it; the page numbers it computed are ours, entry by
  entry, wherever docx2svg's layout reaches (every difference would be listed).
* **A re-save keeps everything**: sections, headers and footers, notes, fields and the
  table of contents; what Word changes is recorded (``test_a_word_resave_keeps_the_structure``).
* Tracked, Word's own **Accept All and Reject All** read as ours.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

import oracle
from conftest import fixture_paths, fixture_id
from docx_agent import Document

from e4_edits import DATE, ENDNOTE, EVEN, FIRST, FOOTNOTE, e4_edit_set, enough

pytestmark = pytest.mark.oracle

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import e4_probe  # noqa: E402
from e3_probe import review_script  # noqa: E402

AUTHOR = "E4 Oracle"
TRACK_DATE = "2026-10-04T12:00:00Z"
#: Word updating every field of a document, as F9 on everything does.
UPDATE = e4_probe._script(["repeat with f in (get fields of d)\n  update field f\nend repeat"])


@pytest.fixture(scope="module")
def word():
    try:
        with oracle.session() as session:
            yield session
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))


def _text(pdf: Path) -> str:
    return " ".join(" ".join(oracle.pdf_pages(pdf)).split())


def _edited(path) -> tuple[Document, dict, bytes]:
    document = Document.open(path)
    made = e4_edit_set(document)
    return document, made, document.to_bytes()


def _toc(document: Document) -> list[tuple[str, str]]:
    """A table of contents' entries: (text, page) each."""
    field = next((f for f in document.fields() if f.keyword == "TOC"), None)
    if field is None:
        return []
    out = []
    for line in field.result.split("\n"):
        text, _, page = line.rpartition("\t")
        out.append((" ".join(text.split()), page.strip()))
    return out


def _pagerefs(document: Document) -> dict[str, str]:
    """Every page reference outside a table of contents: bookmark -> result."""
    out = {}
    for field in document.fields():
        if field.keyword == "PAGEREF" and not field.instruction.split()[1].startswith("_Toc"):
            out[field.instruction.split()[1]] = field.result.strip()
    return out


# -- export -----------------------------------------------------------------------------------


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_word_exports_the_structure_and_shows_it(word, path):
    document, made, data = _edited(path)
    outcome = word.export_pdf(data, name=f"e4-{path.stem}")
    assert outcome, f"Word did not export {path.name}: {outcome.outcome} {outcome.detail}"
    text = _text(outcome.path)
    for visible in (FOOTNOTE, ENDNOTE, FIRST):
        assert visible in text, visible
    for heading in made["headings"]:
        wanted = document.paragraph(heading).text.replace("\ufffc", "")
        assert " ".join(wanted.split())[:30] in text


# -- sections on pages ----------------------------------------------------------------------


def _sectioned(path) -> bytes:
    """The fixture with a next-page break inserted mid-document and every section's own
    footer naming it."""
    document = Document.open(path)
    paragraphs = enough(document, 4)
    document.insert_section_break(after=paragraphs[len(paragraphs) // 2].id, kind="nextPage")
    for k, section in enumerate(document.sections()):
        if document._own_reference(section._sectPr, "footer", "default") is None:
            document.add_footer(section.id, "default", f"E4 section {k} footer")
        else:
            story = section.footer("default")
            document.append_paragraph(f"E4 section {k} footer", story=story.name)
    return document.to_bytes()


def _sections_in_pdf(pdf: Path) -> list[int | None]:
    out = []
    for page in oracle.pdf_pages(pdf):
        found = re.findall(r"E4 section (\d+) footer", " ".join(page.split()))
        out.append(int(found[-1]) if found else None)
    return out


def _sections_drawn(data: bytes) -> tuple[list[int | None], bool]:
    layout = Document.open(data).layout()
    out = []
    for page in layout._lines[:layout.pages_known or len(layout._lines)]:
        text = " ".join("".join("".join(span.chars) for span in line.spans) for _, line in page)
        found = re.findall(r"E4 section (\d+) footer", " ".join(text.split()))
        out.append(int(found[-1]) if found else None)
    return out, layout.stopped is not None


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_a_section_break_gives_word_and_docx2svg_the_same_pages(word, path):
    data = _sectioned(path)
    outcome = word.export_pdf(data, name=f"e4-sections-{path.stem}")
    assert outcome, f"Word did not export {path.name}: {outcome.outcome} {outcome.detail}"
    words = _sections_in_pdf(outcome.path)
    drawn, stopped = _sections_drawn(data)
    if not stopped:
        assert len(drawn) == len(words)
    assert drawn == words[:len(drawn)]


def test_headers_and_footers_show_on_the_right_pages(word):
    document = Document.open(e4_probe.headings_document())
    document.add_header("s:body", "first", FIRST)
    document.add_footer("s:body", "default", "E4 odd footer")
    document.add_footer("s:body", "even", EVEN)
    footer = document.section("s:body").footer("default")
    document.insert_page_number(f"{footer.paragraphs[0].id}@0")
    data = document.to_bytes()
    outcome = word.export_pdf(data, name="e4-header-pages")
    assert outcome, outcome.detail
    pages = [" ".join(p.split()) for p in oracle.pdf_pages(outcome.path)]
    layout = Document.open(data).layout()
    drawn = [" ".join(" ".join("".join("".join(s.chars) for s in line.spans) for _, line in page).split())
             for page in layout._lines]
    assert len(drawn) == len(pages) and len(pages) >= 4
    for number, (words, ours) in enumerate(zip(pages, drawn), start=1):
        for text in (words, ours):
            assert (FIRST in text) == (number == 1)
            assert (EVEN in text) == (number % 2 == 0)
            assert ("E4 odd footer" in text) == (number % 2 == 1 and number > 1)
        if number > 1 and number % 2 == 1:
            assert f"{number}E4oddfooter" in words.replace(" ", "")


# -- fields: Word's own update ----------------------------------------------------------------


def _updated(word, data: bytes, name: str) -> Document:
    outcome = word.run_script(UPDATE, data, name=name, tag="updated", timeout=200)
    assert outcome, f"Word did not update {name}: {outcome.outcome} {outcome.detail}"
    return Document.open(outcome.path.read_bytes())


def _known_pages(document: Document) -> int:
    layout = document.layout()
    return layout.pages_known or (layout.stopped.page - 1 if layout.stopped else len(layout.pages))


def _agreement(ours: Document, words: Document) -> tuple[list, list]:
    """(agreeing, differing) TOC entries and page references, where docx2svg's layout
    reaches."""
    known = _known_pages(ours)
    mine, theirs = _toc(ours), _toc(words)
    assert [t for t, _ in mine] == [t for t, _ in theirs], (mine, theirs)
    agree, differ = [], []
    for (text, page), (_, word_page) in zip(mine, theirs):
        if page and page.isdigit() and int(page) > known:
            continue
        (agree if page == word_page else differ).append((text, page, word_page))
    references, word_references = _pagerefs(ours), _pagerefs(words)
    for name, page in references.items():
        (agree if page == word_references.get(name) else differ).append((name, page, word_references.get(name)))
    return agree, differ


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_toc_and_page_references_are_words(word, path):
    document, made, data = _edited(path)
    page = document.paragraph(made["headings"][-1])
    page.range(0, min(4, len(page.text))).add_bookmark("E4Far")
    first = document.paragraph(made["headings"][0])
    document.insert_cross_reference(f"{first.id}@{len(first.text)}", "E4Far", kind="page")
    data = document.to_bytes()
    words = _updated(word, data, f"e4-toc-{path.stem}")
    agree, differ = _agreement(Document.open(data), words)
    assert agree and not differ, differ


@pytest.mark.parametrize("name", ["long", "probe"])
def test_a_long_toc_is_words(word, name):
    """Many headings over many pages: sample-long's 18 chapters (36 pages), and the probe
    document of three heading levels with a section restarting its numbers in lower Roman."""
    if name == "long":
        document = Document.open(ROOT / "tests" / "fixtures" / "samplelib" / "sample-long.docx")
    else:
        document = Document.open(e4_probe.headings_document())
        for k, level in e4_probe.heading_levels().items():
            document.paragraph(document.paragraphs()[k - 1].id).style = f"heading {level}"
        chapter3 = next(p for p in document.paragraphs() if p.text == "Chapter 3")
        broken = document.insert_section_break(before=chapter3.id)
        document.set_section("s:body", page_number_format="lowerRoman", page_number_start=1)
        assert broken.id
    first = next(e for e in document._index(document.package.document_part()).paragraphs)
    document.insert_toc(before=first.id)
    data = document.to_bytes()
    words = _updated(word, data, f"e4-toc-{name}")
    agree, differ = _agreement(Document.open(data), words)
    assert len(agree) >= 10 and not differ, differ


# -- re-save ----------------------------------------------------------------------------------


def _structure(document: Document) -> dict:
    return {
        "sections": [(s.start, s.orientation, s.columns["count"], s.title_page) for s in document.sections()],
        "stories": [[(which, kind, " ".join("\n".join(p.text for p in (getattr(s, which)(kind) or
                                                                        type("E", (), {"paragraphs": []})).paragraphs).split()))
                     for which in ("header", "footer") for kind in ("default", "first", "even")]
                    for s in document.sections()],
        "notes": sorted(n.text for n in document.notes()),
        "fields": sorted(f.keyword for f in document.fields() if not f.instruction.split()[1:2] == ["_Toc"]
                         and f.keyword != "PAGEREF"),
        "toc": _toc(document),
    }


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_a_word_resave_keeps_the_structure(word, path):
    """Word's re-save keeps the sections, the stories each shows, the notes, the fields and
    the table of contents (its entries and page numbers); what it changes otherwise is
    E0-E3's record (rsids, its own ids, note and bookmark numbers)."""
    document, made, data = _edited(path)
    outcome = word.resave(data, name=f"e4-{path.stem}")
    assert outcome, f"Word did not save {path.name}: {outcome.outcome} {outcome.detail}"
    saved = Document.open(outcome.path.read_bytes())
    ours = _structure(Document.open(data))
    theirs = _structure(saved)
    assert theirs["sections"] == ours["sections"]
    assert theirs["stories"] == ours["stories"]
    assert theirs["notes"] == ours["notes"]
    assert theirs["toc"] == ours["toc"]
    assert theirs["fields"] == ours["fields"]


# -- tracked ----------------------------------------------------------------------------------


def _tracked(path) -> bytes:
    document = Document.open(path)
    with document.tracking(author=AUTHOR, date=TRACK_DATE):
        paragraphs = enough(document, 4)
        document.insert_section_break(after=paragraphs[1].id, kind="continuous")
        document.set_section("s:body", margin_top=54, columns=2)
        c = document.paragraph(paragraphs[2].id)
        document.insert_footnote(f"{c.id}@{len(c.text)}", FOOTNOTE)
        d = document.paragraph(paragraphs[3].id)
        document.insert_date(f"{d.id}@0", date=DATE)
        document.paragraph(paragraphs[0].id).style = "Heading 1"
        document.insert_toc(before=document.paragraph(paragraphs[0].id).id)
    return document.to_bytes()


def _final(document: Document) -> list[str]:
    """The final view's Markdown, notes named by their order (Word renumbers note ids on
    saving: E0)."""
    lines = document.to_markdown(view="final", ids=False).splitlines()
    lines = [re.sub(r"\(d:\d+[^)]*\)", "(picture)", line) for line in lines if line.strip()]
    names: dict[str, str] = {}
    for line in lines:
        for found in re.findall(r"\[\^((?:fn|en):\d+)\]", line):
            names.setdefault(found, f"note{len(names)}")
    return [re.sub(r"\[\^((?:fn|en):\d+)\]", lambda m: f"[^{names[m.group(1)]}]", line) for line in lines]


@pytest.mark.parametrize("accept", [True, False], ids=["accept-all", "reject-all"])
@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_words_review_of_tracked_structure_reads_as_ours(word, path, accept):
    data = _tracked(path)
    action = "accept" if accept else "reject"
    outcome = word.run_script(review_script(action), data, name=f"e4-{path.stem}", tag=action, timeout=200)
    assert outcome, f"Word did not {action} {path.name}: {outcome.outcome} {outcome.detail}"
    theirs = Document.open(outcome.path.read_bytes())
    ours = Document.open(data)
    ours.accept_all() if accept else ours.reject_all()
    if not accept and not any(p.text for p in Document.open(path).paragraphs()):
        # A body with no paragraph (samplelib's blank), or only an empty one (E6's new
        # document): Word neither accepts nor rejects the
        # last mark of what was inserted (E3, E2); its Reject All here joins the break into
        # the body's section, whose properties keep the break's kind and the tracked change
        # (continuous, two columns, 54 pt).  Recorded, and asserted so a Word that changes
        # is noticed; ours gives the original.
        assert [(s.start, s.columns["count"], s.margins["top"]) for s in theirs.sections()] == \
            [("continuous", 2, 54.0)]
        return
    assert _final(theirs) == _final(ours)
    assert [(s.start, s.columns["count"], s.margins["top"]) for s in theirs.sections()] == \
        [(s.start, s.columns["count"], s.margins["top"]) for s in ours.sections()]


# -- last: a field marked dirty ------------------------------------------------------------


def test_a_dirty_field_makes_word_ask(word):
    """Why a result the layout cannot say is left unmarked: a page reference marked dirty
    (``w:dirty``) makes Word ask on opening whether to update the document's fields -- the
    export blocks on the question (measured; as it would for ``w:updateFields``)."""
    document = Document.open(e4_probe.plain_document())
    last = document.paragraphs()[-1]
    last.range(0, 9).add_bookmark("Far")
    result = document.insert_field(f"{document.paragraphs()[0].id}@0", " PAGEREF Far \\h ")
    plain = word.export_pdf(document.to_bytes(), name="e4-undirty", timeout=90)
    assert plain, f"Word did not export the unmarked field: {plain.outcome} {plain.detail}"
    document.field(result.id)._found().begin.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}dirty", "true")
    document.package.mark_dirty(document.package.document_part())
    outcome = word.export_pdf(document.to_bytes(), name="e4-dirty", timeout=60)
    assert not outcome and outcome.outcome == "rejected"
    # Last in the module on purpose: the Word the recovery quits may hold the question
    # into its next launches (seen: every save after it blocked in the same session, and a
    # Word left hanging after the session).  So this Word (the session's, under the lock)
    # is made to go now.
    import subprocess
    import time

    word.recover()
    for _ in range(30):
        if not oracle.word_running():
            break
        time.sleep(1)
    if oracle.word_running():
        subprocess.run(["pkill", "-9", "-x", "Microsoft Word"], check=False)
        time.sleep(3)
    assert not oracle.word_running()
