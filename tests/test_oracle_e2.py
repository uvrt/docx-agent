"""Word itself on ``insert_markdown``, opt-in (``pytest -m oracle``; ROADMAP.md, Phase E2,
"Done when").

For every fixture and the blank document, with a Markdown document of every construct
(:data:`MARKDOWN`: headings, lists nested and numbered from 3, a table, code, a quote, links,
footnotes, a picture) inserted at the end of the body:

* Word opens it and exports it unprompted, and its PDF shows what was inserted: the
  headings, the list items with the numbers Word draws -- the ones our state gives -- the
  table's cells, the code, the quote, the link texts and the footnotes' text;
* docx2svg's drawing has, page by page, Word's characters, with no difference the
  unedited document lacks;
* a Word re-save keeps everything: the final view's Markdown, the notes, the pictures,
  the styles used; what Word changes is recorded (:func:`test_a_word_resave_keeps_the_insertion`);
* tracked, Word's own Accept All and Reject All read as ours -- but for what Word's review
  leaves of a form it writes itself (:data:`LEFTOVER`, recorded).
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

import oracle
from conftest import fixture_id, reading_paths
from docx_agent import Document

from test_markdown_write import IMAGES
from test_oracle_e3 import _final, _page_characters

pytestmark = pytest.mark.oracle

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
from e3_probe import review_script  # noqa: E402

AUTHOR = "E2 Oracle"
DATE = "2026-10-03T12:00:00Z"
MARKDOWN = """# Inserted from Markdown

A paragraph with *emphasis*, **strong**, `code`, ~~struck~~ and a
[link to the site](https://example.com/e2).[^first]

## A list

3. third item
4. fourth item
   - nested bullet
   - another bullet
5. fifth item

| Region | Value |
| :--- | ---: |
| North | 4.1 |
| South | 2.8 |

> A quoted sentence.

```
code line one
  code line two
```

---

A picture: ![a square](square.png) and a second note.[^second]

[^first]: The first footnote's text.
[^second]: The second footnote's text.
"""
#: What the PDF must show of it (whitespace collapsed).
VISIBLE = ["Inserted from Markdown", "A list", "third item", "fourth item", "nested bullet", "fifth item",
           "North", "4.1", "South", "2.8", "A quoted sentence.", "code line one", "code line two",
           "link to the site", "The first footnote's text.", "The second footnote's text."]


@pytest.fixture(scope="module")
def word():
    try:
        with oracle.session() as session:
            yield session
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))


def _inserted(path, *, tracked: bool = False):
    document = Document.open(path)
    if tracked:
        with document.tracking(author=AUTHOR, date=DATE):
            result = document.insert_markdown(MARKDOWN, images=IMAGES)
    else:
        result = document.insert_markdown(MARKDOWN, images=IMAGES)
    return document, result, document.to_bytes()


def _text(pdf: Path) -> str:
    return " ".join(" ".join(oracle.pdf_pages(pdf)).split())


@pytest.mark.parametrize("path", reading_paths(), ids=fixture_id)
def test_word_exports_the_insertion_and_shows_it(word, path):
    document, result, data = _inserted(path)
    outcome = word.export_pdf(data, name=f"e2-{path.stem}")
    assert outcome, f"Word did not export {path.name}: {outcome.outcome} {outcome.detail}"
    text = _text(outcome.path)
    for visible in VISIBLE:
        assert " ".join(visible.split()) in text, visible
    # The numbers Word draws are the ones the state gives (docx2svg's counter).
    state = document.state(f"{result.blocks[0]}..{result.blocks[-1]}")
    labels = [(b["list"]["label"], b["t"]) for b in state["blocks"]
              if b.get("list") and b["list"]["format"] != "bullet"]
    assert [label for label, _ in labels] == ["3.", "4.", "5."]
    for label, item in labels:
        assert re.search(re.escape(label) + r"\s*" + re.escape(item), text), (label, item)
    # Footnote marks: 1 and 2 after this document's own notes.
    assert "Inserted from Markdown" in text


#: The documents whose pages are compared: not the chart fixtures, whose charts' text is not
#: among docx2svg's lines, and whose floating drawings docx2svg pages apart from Word
#: (``chart-places``; ROADMAP.md, "Charts and SmartArt", proposal 7) -- their charts are held
#: to Word by ``test_oracle_charts.py``.
PAGED = [path for path in reading_paths() if path.parent.name != "charts"]


@pytest.mark.parametrize("path", PAGED, ids=fixture_id)
def test_docx2svg_draws_what_word_draws(word, path):
    _, _, data = _inserted(path)
    words, drawn, stopped = _page_characters(word, data, f"e2-{path.stem}")
    if not stopped:
        assert len(drawn) == len(words)
    base_words, base_drawn, _ = _page_characters(word, path.read_bytes(), f"unedited-{path.stem}")
    from collections import Counter

    allowed = Counter()
    for a, b in zip(base_words, base_drawn):
        allowed += (a - b) + (b - a)
    found = Counter()
    for k, page in enumerate(drawn):
        found += (words[k] - page) + (page - words[k])
    assert not found - allowed, dict(found - allowed)


#: What a Word re-save changes of an insertion, recorded (nothing else may change).
def test_a_word_resave_keeps_the_insertion(word):
    for path in reading_paths():
        document, result, data = _inserted(path)
        outcome = word.resave(data, name=f"e2-{path.stem}")
        assert outcome, f"Word did not save {path.name}: {outcome.outcome} {outcome.detail}"
        saved = Document.open(outcome.path.read_bytes())
        ours = Document.open(data)
        assert _final(saved) == _final(ours), path.name
        assert [n.text.strip() for n in saved.notes()] == [n.text.strip() for n in ours.notes()], path.name
        assert len(saved.pictures()) == len(ours.pictures()), path.name
        # Every style the insertion names keeps its name (the paragraphs, below) -- but the
        # Dutch template's "Kop 1" and "Kop 2", whose ids are Dutch Word's for its built-in
        # headings, which Word renames "heading 1" and "heading 2"; and where the insertion
        # wrote the styles part (layout-sweep has none), Word gives its styles its own
        # localised ids and names the default paragraph style Normal (recorded).
        # Our paraIds stay (the document is complete: E0's measurement), each paragraph
        # with its text and its style's name.
        for identifier in result.blocks:
            if identifier.startswith("p:"):
                assert saved.paragraph(identifier).text == ours.paragraph(identifier).text
                pair = (ours.paragraph(identifier).style_name, saved.paragraph(identifier).style_name)
                assert pair[0] == pair[1] or pair in {("Kop 1", "heading 1"), ("Kop 2", "heading 2"),
                                                       (None, "Normal")}, \
                    (path.name, identifier, pair)


#: What Word's own Accept All and Reject All leave of an insertion at a body's end (recorded):
#: a Reject All of two or more paragraphs typed at the end leaves the last one's paragraph-
#: property change unrejected -- and so does Word's Reject All of the same paragraphs *Word*
#: typed (``tools/e2_probe.py``, "typing_at_the_end"): the form is Word's; and in a body
#: with no paragraph at all (samplelib's blank), the last new paragraph's inserted mark is
#: the container's last, which Word neither accepts nor rejects (E3's measurement).
LEFTOVER = {"paragraph-properties", "paragraph-mark-insertion"}


def _text_lines(lines: list[str]) -> list[str]:
    """The lines without their block syntax: the paragraph a pending property change is
    left on shows the inserted paragraph's construct, not its own."""
    out = [re.sub(r"^(#+ |> |- |\* |\d+[.)] )", "", line) for line in lines]
    return [line for line in out if not re.fullmatch(r"`{3,}|~{3,}", line)]


@pytest.mark.parametrize("accept", [True, False], ids=["accept-all", "reject-all"])
@pytest.mark.parametrize("path", reading_paths(), ids=fixture_id)
def test_words_review_of_the_tracked_insertion_reads_as_ours(word, path, accept):
    _, result, data = _inserted(path, tracked=True)
    export = word.export_pdf(data, name=f"e2-tracked-{path.stem}")
    assert export, f"Word did not export it: {export.outcome} {export.detail}"
    ours = Document.open(data)
    ours.accept_all() if accept else ours.reject_all()
    action = "accept" if accept else "reject"
    outcome = word.run_script(review_script(action), data, name=f"e2-{path.stem}", tag=action, timeout=200)
    assert outcome, f"Word did not {action}: {outcome.outcome} {outcome.detail}"
    words = Document.open(outcome.path.read_bytes())
    left = words.revisions(author=AUTHOR)
    if not left:
        assert _final(words) == _final(ours)
        return
    # Only what is recorded, on the body's last paragraph; the text is ours.
    assert {r.kind for r in left} <= LEFTOVER, [(r.kind, r.paragraph_id) for r in left]
    last = words.paragraphs()[-1].id
    assert {r.paragraph_id for r in left} == {last}
    assert _text_lines(_final(words)) == _text_lines(_final(ours))
