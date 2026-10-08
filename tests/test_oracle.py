"""Word itself, opt-in (``pytest -m oracle``): every fixture with E0's edits opens and
exports to PDF unprompted, and a re-save through Word keeps what docx-agent wrote.

The re-save is how "Word accepted it as it is" is told apart from "Word silently repaired
it": the test diffs what Word wrote back -- every paraId docx-agent stamped kept, every
paragraph's text, the compatibility mode -- and prints what else it changed.

Word is one instance per machine: the tests hold the machine-wide lock (tests/oracle.py)
and skip, rather than wait or share, when Word is already in use.
"""

from __future__ import annotations

import io
import zipfile

import pytest

import oracle
from docx_agent import Document

from test_edit import edit_set

pytestmark = pytest.mark.oracle


@pytest.fixture(scope="module")
def word():
    try:
        with oracle.session() as session:
            yield session
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))


def edited(path) -> Document:
    document = Document.open(path)
    edit_set(document)
    return document


def test_word_exports_every_fixture_with_e0_edits(word, docx_path):
    document = edited(docx_path)
    outcome = word.export_pdf(document.to_bytes(), name=docx_path.stem)
    assert outcome, f"Word did not export {docx_path.name}: {outcome.outcome} {outcome.detail}"
    pages = oracle.pdf_pages(outcome.path)
    assert pages
    # Word's PDF separates justified words with tabs and lines with CRLF.
    text = " ".join(" ".join(pages).split())
    if any(p.text.strip() for p in Document.open(docx_path).paragraphs()):
        for marker in ("(edited)", "Inserted after the first.", "Inserted before the last."):
            assert marker in text, marker


def test_word_keeps_what_docx_agent_wrote(word, docx_path):
    document = edited(docx_path)
    data = document.to_bytes()
    outcome = word.resave(data, name=docx_path.stem)
    assert outcome, f"Word did not save {docx_path.name}: {outcome.outcome} {outcome.detail}"
    saved = Document.open(outcome.path.read_bytes())

    assert saved.compatibility_mode == document.compatibility_mode
    for story in document.stories:
        ours = [(p.para_id, p.text) for p in story.paragraphs if not separator(p)]
        theirs = [(p.para_id, p.text) for p in saved.story(story.name).paragraphs if not separator(p)] \
            if story.name in {s.name for s in saved.stories} else []
        if not ours:
            continue
        assert [t for _, t in theirs] == [t for _, t in ours], story.name
        if document.history.can_undo():
            # The document was stamped whole: Word keeps every paraId.
            assert [i for i, _ in theirs] == [i for i, _ in ours], story.name
    print(f"\n{docx_path.name}: {summary(data, outcome.path.read_bytes())}")


def separator(paragraph) -> bool:
    """A paragraph of a separator note (``w:footnote w:type="separator"``...): Word writes
    those itself on saving, with new paraIds (measured), so their ids are not durable."""
    note = paragraph._element.getparent()
    return note is not None and note.get(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type") is not None


def summary(ours: bytes, words: bytes) -> str:
    """What Word changed in a re-save: parts added, removed and rewritten."""
    def parts(data: bytes) -> dict[str, bytes]:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            return {name: archive.read(name) for name in archive.namelist()}

    before, after = parts(ours), parts(words)
    added = sorted(set(after) - set(before))
    removed = sorted(set(before) - set(after))
    rewritten = sorted(name for name in set(before) & set(after) if before[name] != after[name])
    return f"added {added}; removed {removed}; rewritten {rewritten}"


# -- E1 -----------------------------------------------------------------------------------------


def e1_edited(path):
    from e1_edits import e1_edit_set

    document = Document.open(path)
    expected = e1_edit_set(document)
    return document, expected


def test_word_exports_every_fixture_with_e1_edits(word, docx_path):
    """No repair prompt; the replaced text, the heading, the list numbers and the bullet's
    cross-reference in Word's PDF; the styles visible in how it draws them."""
    from e1_edits import BIG, HEADING

    document, expected = e1_edited(docx_path)
    outcome = word.export_pdf(document.to_bytes(), name=f"e1-{docx_path.stem}")
    assert outcome, f"Word did not export {docx_path.name}: {outcome.outcome} {outcome.detail}"
    text = " ".join(" ".join(oracle.pdf_pages(outcome.path)).split())
    for marker in expected.texts:
        assert marker in text, marker
    pages = oracle.pdf_chars(outcome.path)
    heading = oracle.find_in_pdf(pages, HEADING)
    big = oracle.find_in_pdf(pages, BIG)
    body = oracle.find_in_pdf(pages, "E1 alpha item")
    assert heading and big and body
    assert heading[0].size > body[0].size                     # Heading 2: 13 pt against the body's
    assert abs(big[0].size - 20) < 0.2 and big[0].color == (192, 0, 0)   # direct: 20 pt, C00000
    swapped = oracle.find_in_pdf(pages, "E1swap")
    # Strong is bold, unless the paragraph's style is bold too (Word's toggle rule): what the
    # resolver says the run is, Word must draw.
    found = document.anchor("E1swap")
    run = next(r for r in document.paragraph(found.start_id).runs if "E1swap" in r.text)
    drawn_bold = swapped[0].weight >= 600 or "bold" in swapped[0].font.lower() or swapped[0].stroked
    assert swapped and drawn_bold == run.effective.bold
    print(f"\n{docx_path.name}: heading {heading[0].size} pt {heading[0].font}, body {body[0].size} pt, "
          f"big {big[0].size} pt {big[0].color}, E1swap weight {swapped[0].weight} {swapped[0].font}")


def test_word_keeps_what_e1_wrote(word, docx_path):
    """Re-saved by Word: every paragraph's text and paraId kept; and what Word changed of our
    styles, lists, hyperlinks, bookmarks and pictures, printed (ROADMAP.md, Phase E1)."""
    document, expected = e1_edited(docx_path)
    data = document.to_bytes()
    outcome = word.resave(data, name=f"e1-{docx_path.stem}")
    assert outcome, f"Word did not save {docx_path.name}: {outcome.outcome} {outcome.detail}"
    saved = Document.open(outcome.path.read_bytes())
    assert saved.compatibility_mode == document.compatibility_mode
    for story in document.stories:
        ours = [(p.para_id, p.text) for p in story.paragraphs if not separator(p)]
        if not ours or story.name not in {s.name for s in saved.stories}:
            continue
        theirs = [(p.para_id, p.text) for p in saved.story(story.name).paragraphs if not separator(p)]
        assert [t for _, t in theirs] == [t for _, t in ours], story.name
        assert [i for i, _ in theirs] == [i for i, _ in ours], story.name

    report = e1_resave_report(document, saved, expected)
    assert report["lists"]["ours"] == report["lists"]["word"]
    assert report["hyperlinks"]["ours"] == report["hyperlinks"]["word"]
    assert set(report["bookmarks"]["ours"]) <= set(report["bookmarks"]["word"])
    assert report["pictures"]["ours"] == report["pictures"]["word"]
    print(f"\n{docx_path.name}: {report['styles']}; ids {report['style_ids']}; bookmarks added by Word "
          f"{sorted(set(report['bookmarks']['word']) - set(report['bookmarks']['ours']))}; "
          f"docPr {report['pictures']['docpr']}; numbering {report['numbering']}")


def e1_resave_report(ours: Document, words: Document, expected) -> dict:
    """What Word kept and changed of the styles, lists, hyperlinks, bookmarks and pictures."""
    from lxml import etree

    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    def definitions(document: Document) -> dict[str, tuple[str, str]]:
        root = document.package.tree(document.package.styles_part())
        out = {}
        for node in root.findall(w + "style"):
            name = node.find(w + "name").get(w + "val")
            parts = []
            for tag in ("pPr", "rPr"):
                child = node.find(w + tag)
                if child is not None:
                    child = etree.fromstring(etree.tostring(child))
                    for numbering in child.iter(w + "numId"):
                        numbering.set(w + "val", "#")
                    parts.append(etree.tostring(child, method="c14n", exclusive=True).decode())
            out[name] = (node.get(w + "styleId"), "".join(parts))
        return out

    ours_styles, word_styles = definitions(ours), definitions(words)
    written = [name for name in ours_styles if ours.styles.find(name) and name in (
        "heading 2", "Heading 2 Char", "Strong", "List Paragraph", "Hyperlink")]
    changed = sorted(name for name in written if name in word_styles and ours_styles[name][1] != word_styles[name][1])
    missing = sorted(name for name in written if name not in word_styles)
    renamed = {name: (ours_styles[name][0], word_styles[name][0]) for name in written
               if name in word_styles and ours_styles[name][0] != word_styles[name][0]}

    def lists(document: Document):
        out = []
        for identifier in expected.ids["items"].split(","):
            found = document.paragraph(identifier).list
            out.append(None if found is None else (found.kind, found.format, found.level, found.text))
        return out

    def starts(document: Document):
        numbering = document.numbering
        values = []
        for identifier in expected.ids["items"].split(",")[2:3]:
            found = document.paragraph(identifier).list
            num = numbering.nums()[found.num_id]
            override = num.find(f"{w}lvlOverride/{w}startOverride")
            values.append(override.get(w + "val") if override is not None else None)
        return values

    def links(document: Document):
        return [(h.text, h.address, h.anchor) for h in document.hyperlinks()]

    def pictures(document: Document):
        return [(p.alt_text, tuple(round(v, 1) for v in p.size)) for p in document.pictures()]

    return {
        "styles": {"changed": changed, "missing": missing},
        "style_ids": renamed,
        "lists": {"ours": lists(ours), "word": lists(words)},
        "numbering": {"restart": (starts(ours), starts(words))},
        "hyperlinks": {"ours": links(ours), "word": links(words)},
        "bookmarks": {"ours": [b.name for b in ours.bookmarks(hidden=True)],
                      "word": [b.name for b in words.bookmarks(hidden=True)]},
        "pictures": {"ours": pictures(ours), "word": pictures(words),
                     "docpr": ([p.id for p in ours.pictures()], [p.id for p in words.pictures()])},
    }
