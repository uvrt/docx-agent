"""Word itself on E3's tracked edits and comments, opt-in (``pytest -m oracle``; ROADMAP.md,
Phase E3, "Done when").

For every fixture with E3's edit set (``e3_edits.py``):

* Word opens it and exports it unprompted (its final view: ``tests/oracle.py`` exports a
  document with markup in Word's *No Markup* view);
* Word's own **Accept All**, saved, reads as our ``accept_all`` does -- and its **Reject
  All** as our ``reject_all`` -- compared through the final view's Markdown (text,
  constructs by style *name*, lists, emphasis, links, tables), which a Word save leaves
  comparable where its XML is not (rsids, its own style ids, merged runs);
* a Word re-save keeps our revisions' author and date, every comment and reply with its
  durable id, thread and resolution;
* docx2svg's drawing of the tracked document is Word's final-view export, page by page.

The structural table forms (a column deleted, cells merged down and across) are held to
Word's Accept All and Reject All on one table fixture; where Word differs (its Reject All
of a merge across), the difference is asserted as recorded, so a Word that changes is
noticed.
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

import pytest

import oracle
from docx_agent import Document

from e3_edits import AUTHOR, DATE, e3_edit_set

pytestmark = pytest.mark.oracle

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
from e3_probe import review_script  # noqa: E402


@pytest.fixture(scope="module")
def word():
    try:
        with oracle.session() as session:
            yield session
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))


def _tracked(path) -> tuple[Document, bytes]:
    document = Document.open(path)
    e3_edit_set(document)
    return document, document.to_bytes()


def _final(document: Document) -> list[str]:
    """The final view as Markdown lines, without ids; comments are not in it."""
    lines = document.to_markdown(view="final", ids=False).splitlines()
    return [re.sub(r"\(d:\d+[^)]*\)", "(picture)", line) for line in lines if line.strip()]


def _ours(data: bytes, accept: bool) -> Document:
    document = Document.open(data)
    if accept:
        document.accept_all()
    else:
        document.reject_all()
    return document


def _words(word, data: bytes, accept: bool, name: str) -> Document:
    action = "accept" if accept else "reject"
    outcome = word.run_script(review_script(action), data, name=f"e3-{name}", tag=action, timeout=200)
    assert outcome, f"Word did not {action} {name}: {outcome.outcome} {outcome.detail}"
    return Document.open(outcome.path.read_bytes())


def test_word_exports_every_tracked_fixture(word, docx_path):
    document, data = _tracked(docx_path)
    outcome = word.export_pdf(data, name=f"e3-{docx_path.stem}")
    assert outcome, f"Word did not export {docx_path.name}: {outcome.outcome} {outcome.detail}"
    assert oracle.pdf_pages(outcome.path)


@pytest.mark.parametrize("accept", [True, False], ids=["accept-all", "reject-all"])
def test_words_review_reads_as_ours(word, docx_path, accept):
    document, data = _tracked(docx_path)
    ours = _ours(data, accept)
    words = _words(word, data, accept, docx_path.stem)
    assert words.revisions(author=AUTHOR) == []
    # Comments are not revisions: both keep them.
    assert sorted(c.text for c in words.comments()) == sorted(c.text for c in ours.comments())
    assert _final(words) == _final(ours)


def test_a_word_resave_keeps_revisions_and_comments(word, docx_path):
    document, data = _tracked(docx_path)
    outcome = word.resave(data, name=f"e3-{docx_path.stem}")
    assert outcome, f"Word did not save {docx_path.name}: {outcome.outcome} {outcome.detail}"
    saved = Document.open(outcome.path.read_bytes())
    ours = Document.open(data)
    mine = ours.revisions(author=AUTHOR)
    theirs = saved.revisions(author=AUTHOR)
    assert mine and len(theirs) >= 1
    assert {r.date for r in theirs} == {DATE}
    # Word keeps every kind of revision we wrote -- but drops a formatting change that
    # changes nothing (bold on a heading bold already), and its record with it (recorded).
    assert {r.kind for r in theirs} <= {r.kind for r in mine}
    assert {r.kind for r in mine} - {r.kind for r in theirs} <= {"run-properties"}

    def threads(document: Document) -> list[tuple]:
        records = document.comments()
        return sorted((c.id, c.author, c.text, c.parent, c.done) for c in records)

    assert threads(saved) == threads(ours)
    # Word re-saves a hyperlink that holds a revision as a HYPERLINK field (and writes the
    # element again once the revision is accepted, as its Accept All above shows):
    # recorded; the link's text and address are kept.
    assert _unlinked(_final(saved)) == _unlinked(_final(ours))
    xml = saved.package.read(saved.package.document_part()).decode("utf-8")
    assert 'HYPERLINK "https://example.com/e3"' in xml or "https://example.com/e3" in str(
        [r.target for r in saved.package.relationships(saved.package.document_part()).values()])


def _unlinked(lines: list[str]) -> list[str]:
    return [re.sub(r"(?<!!)\[([^\]]*)\]\([^)]*\)", r"\1", line) for line in lines]


def _page_characters(word, data: bytes, name: str) -> tuple[list[Counter], list[Counter], bool]:
    """Each page's characters as Word's PDF has them and as docx2svg draws them (whitespace,
    hyphenation points and symbol-font characters aside), and whether docx2svg stopped."""
    outcome = word.export_pdf(data, name=name)
    assert outcome, f"Word did not export {name}: {outcome.outcome} {outcome.detail}"
    words = [Counter(_plain(page)) for page in oracle.pdf_pages(outcome.path)]
    layout = Document.open(data).layout()
    known = layout.pages_known or len(layout._lines)
    drawn = [Counter(_plain("".join("".join(c for span in line.spans for c in span.chars) for _, line in page)))
             for page in layout._lines[:known]]
    return words, drawn, layout.stopped is not None


def test_docx2svg_draws_words_final_view(word, docx_path):
    """docx2svg's drawing of the tracked document is Word's export of its final view: the
    same characters on each page it lays out -- compared as each page's multiset of
    characters, since the PDF's text runs table cells, headers and notes in an order of its
    own -- with no difference the unedited document does not have already."""
    document, data = _tracked(docx_path)
    words, drawn, stopped = _page_characters(word, data, f"e3-{docx_path.stem}")
    if not stopped:
        assert len(drawn) == len(words)
    base_words, base_drawn, _ = _page_characters(word, docx_path.read_bytes(), f"unedited-{docx_path.stem}")
    allowed = Counter()
    for a, b in zip(base_words, base_drawn):
        allowed += (a - b) + (b - a)
    found = Counter()
    for k, page in enumerate(drawn):
        found += (words[k] - page) + (page - words[k])
    assert not found - allowed, dict(found - allowed)


def _plain(text: str) -> str:
    """What both draw as text: no soft hyphens, object marks or private-use characters
    (symbol-font bullets), and the hyphen docx2svg draws for a non-breaking one."""
    text = text.replace("\u2011", "-")
    return re.sub("[\\s\u00ad\ufffc\ufffe\u2022\ue000-\uf8ff-]", "", text)


# -- the structural table forms ------------------------------------------------------------------

#: What Word's Accept All and Reject All make of each structural form, against ours:
#: ``same``, or the difference recorded (ROADMAP.md, Phase E3).
STRUCTURAL = {
    ("delete_column", True): "same", ("delete_column", False): "same",
    ("merge_down", True): "same", ("merge_down", False): "same",
    # E3's form differed both ways (Word's Accept All left the merged cell a column wider,
    # its Reject All mangled the row); E5's -- the deleted cells' marks deleted, every cell's
    # old properties and the old grid recorded, as Word records a table (tools/e5_probe.py,
    # mergeforms) -- Word reviews as docx-agent does.
    ("merge_across", True): "same", ("merge_across", False): "same",
}


def _structural(path, operation: str) -> bytes:
    document = Document.open(path)
    table = document.tables()[0]
    with document.tracking(author=AUTHOR, date=DATE):
        if operation == "delete_column":
            document.delete_column(table.id, 2)
        elif operation == "merge_down":
            document.merge_cells(table.id, (1, 0), (2, 0))
        else:
            document.merge_cells(table.id, (3, 0), (3, 1))
    return document.to_bytes()


@pytest.mark.parametrize("operation,accept", sorted(STRUCTURAL), ids=lambda v: str(v))
def test_words_review_of_the_structural_forms(word, operation, accept):
    from conftest import FIXTURE_DIR

    path = FIXTURE_DIR / "samplelib" / "sample-simple.docx"
    data = _structural(path, operation)
    export = word.export_pdf(data, name=f"e3-{operation}")
    assert export, f"Word did not export it: {export.outcome} {export.detail}"
    ours = _ours(data, accept)
    words = _words(word, data, accept, operation)
    # Word keeps the comment that describes the change; ours goes with the change.
    assert [c.text for c in words.comments()] and all(
        c.text.startswith("Tracked structural change:") for c in words.comments())
    assert ours.comments() == []
    same = _final(words) == _final(ours)
    assert ("same" if same else "differs") == STRUCTURAL[(operation, accept)], \
        "\n".join(_final(words)) + "\n----\n" + "\n".join(_final(ours))
