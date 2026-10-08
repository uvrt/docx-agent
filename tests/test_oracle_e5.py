"""Word itself on E5's tables, drawings and content controls, opt-in (``pytest -m oracle``;
ROADMAP.md, Phase E5, "Done when").

* Every fixture with E5's edit set (``e5_edits.e5_edit_set``: a table with a header row, a
  merge and a split; a nested table; a floating picture; a text box; a shape; a control of
  each kind, filled) **exports from Word unprompted**, its PDF showing the tables', the text
  box's, the shape's and the controls' text.
* **Word's pages are docx2svg's**: each page's characters in Word's PDF are the ones
  docx2svg draws there, wherever its layout reaches, with no difference the unedited
  document lacks -- for the edit set on every fixture and for the table probe with every
  property edit.
* **Every content-control fill shows in Word's PDF**, a data-bound one too -- and a bound
  control whose cache disagrees with its node shows the node's, as Word does.
* **A re-save keeps everything**: the tables' cells, spans and merges, header rows and
  floating; the drawings' kind, place, wrapping, size and text; the controls' kinds and
  values and the bound node.
* Tracked, **Word's own Accept All and Reject All read as ours** for every table operation,
  the merge across among them (in the form Word reviews: E3's, with every cell of its row
  and the grid recording their old state), a text box, a shape and a control's fill.
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

import pytest

import oracle
from conftest import fixture_id, fixture_paths
from docx_agent import Document

from e5_edits import BOX, CONTROL, SHAPE_TEXT, e5_edit_set
from test_e5_tracked import AUTHOR, CASES, DATE
from test_oracle_e3 import _plain

pytestmark = pytest.mark.oracle

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import e5_probe  # noqa: E402
from e3_probe import review_script  # noqa: E402
from test_e5_word_forms import TABLE_EDITS  # noqa: E402


@pytest.fixture(scope="module")
def word():
    try:
        with oracle.session() as session:
            yield session
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))


def _text(pdf: Path) -> str:
    return " ".join(" ".join(oracle.pdf_pages(pdf)).split())


def _edited(path) -> tuple[Document, bytes]:
    document = Document.open(path)
    document.checks = e5_edit_set(document)
    return document, document.to_bytes()


# -- export -----------------------------------------------------------------------------------


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_word_exports_the_edit_set_and_shows_it(word, path):
    document, data = _edited(path)
    outcome = word.export_pdf(data, name=f"e5-{path.stem}")
    assert outcome, f"Word did not export {path.name}: {outcome.outcome} {outcome.detail}"
    text = _text(outcome.path).replace(" ", "")  # narrow cells' text comes out letter-spaced
    for visible in ("E5 head two", "m33", "n4", BOX, SHAPE_TEXT, CONTROL, "E5 rich", "Beta", "Purple", "4-10-2026",
                    "E5 block two"):
        assert visible.replace(" ", "") in text, visible


# -- pages ------------------------------------------------------------------------------------


def _page_characters(word, data: bytes, name: str) -> tuple[list[Counter], list[Counter], bool]:
    outcome = word.export_pdf(data, name=name)
    assert outcome, f"Word did not export {name}: {outcome.outcome} {outcome.detail}"
    words = [Counter(_plain(page)) for page in oracle.pdf_pages(outcome.path)]
    layout = Document.open(data).layout()
    known = layout.pages_known or len(layout._lines)
    drawn = []
    for page in layout._lines[:known]:
        drawn.append(Counter(_plain("".join("".join(c for span in line.spans for c in span.chars) for _, line in page))))
    return words, drawn, layout.stopped is not None


#: Where Word's pages and docx2svg's part by a line at a page's foot -- each edit group
#: alone agrees on these fixtures (tables, controls, the floating picture, the text box and
#: shape, each exported apart); together the small differences add up to a line or two
#: changing pages (recorded in ROADMAP.md, Phase E5).
A_LINE_APART = {"ids-and-markup", "sample-long", "sample-simple"}


def _same_pages(word, data: bytes, base: bytes, name: str, stem: str | None = None) -> None:
    """The same pages with the same characters on each, but what the unedited document
    already differs by; on :data:`A_LINE_APART` the same pages and characters, a line's
    characters on another page at most."""
    words, drawn, stopped = _page_characters(word, data, name)
    base_words, base_drawn, _ = _page_characters(word, base, f"{name}-unedited")
    allowed = Counter()
    for a, b in zip(base_words, base_drawn):
        allowed += (a - b) + (b - a)
    if not stopped:
        assert len(drawn) == len(words), (len(drawn), len(words))
        total = sum(words, Counter())
        mine = sum(drawn, Counter())
        assert not ((total - mine) + (mine - total)) - allowed
    found = {}
    for k, page in enumerate(drawn):
        moved = (words[k] - page) + (page - words[k]) - allowed
        if moved:
            found[k] = sum(moved.values())
    if stem in A_LINE_APART:
        assert all(count <= 120 for count in found.values()), found
    else:
        assert not found, found


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_docx2svg_draws_words_pages(word, path):
    """The edit set's pages: each page's characters as Word's PDF has them and as docx2svg
    draws them (as E3's oracle compares them), wherever docx2svg's layout reaches."""
    _, data = _edited(path)
    _same_pages(word, data, path.read_bytes(), f"e5-pages-{path.stem}", path.stem)


def test_table_properties_give_word_and_docx2svg_the_same_pages(word):
    """The table probe with every property edit E5 writes (header rows, fixed layout, a
    width in percent, alignment, indent, style options, heights, shading, borders, margins,
    vertical alignment, text direction, a column's width, floating)."""
    document = Document.open(e5_probe.tables_document())
    tables = {re.search(r"(\w+) r1c1", "".join(t._entry()[1].element.itertext())).group(1): t.id
              for t in document.tables()}
    for key, edit in TABLE_EDITS.items():
        edit(document, tables[key])
    _same_pages(word, document.to_bytes(), e5_probe.tables_document(), "e5-table-probe")


def test_merges_and_splits_give_word_and_docx2svg_the_same_pages(word):
    from test_e5_word_forms import MERGE_EDITS

    document = Document.open(e5_probe.merges_document())
    tables = {re.search(r"(\w+) r1c1", "".join(t._entry()[1].element.itertext())).group(1): t.id
              for t in document.tables()}
    for key, edit in MERGE_EDITS.items():
        edit(document, tables[key])
    _same_pages(word, document.to_bytes(), e5_probe.merges_document(), "e5-merge-probe")


# -- content controls ---------------------------------------------------------------------------


def test_every_fill_shows_in_words_pdf_bound_ones_too(word):
    document = Document.open(e5_probe.controls_document())
    fills = {"cc:1001": "Filled plain", "cc:1002": "Filled rich", "cc:1003": "Beta", "cc:1004": "Filled combo",
             "cc:1005": "2026-12-24", "cc:1007": "Filled bound", "cc:1008": "Filled bound date"}
    for identifier, value in fills.items():
        document.fill_control(identifier, value)
    document.fill_control("cc:1006", True)
    document.fill_control("cc:1100", "Filled block one\nFilled block two")
    outcome = word.export_pdf(document.to_bytes(), name="e5-controls")
    assert outcome, f"Word did not export it: {outcome.outcome} {outcome.detail}"
    text = _text(outcome.path)
    for visible in ("Filled plain", "Filled rich", "Beta", "Filled combo", "24-12-2026", "Filled bound",
                    "Filled bound date", "Filled block one", "Filled block two", "☒"):
        assert visible in text, visible


def test_word_shows_a_bound_controls_node_not_its_cache(word):
    """Why a fill writes the node: a bound control whose cached content disagrees with its
    node shows the node's in Word (the probe's ``Stale cache``)."""
    outcome = word.export_pdf(e5_probe.controls_document(), name="e5-controls-stale")
    assert outcome, outcome.detail
    text = _text(outcome.path)
    assert "Bound from XML" in text and "Stale cache" not in text


# -- re-save ----------------------------------------------------------------------------------


def _structure(document: Document) -> dict:
    tables = []
    for story in document.stories:
        for table in story.tables:
            grid = table.grid()
            spans = [[(c.row, c.column) if c is not None else None for c in row] for row in grid]
            texts = [[" ".join(c.text.split()) if c is not None else None for c in row] for row in grid]
            props = table.properties
            tables.append((spans, texts, props["header_rows"], props["floating"] is not None))
    drawings = []
    for part, identifier, frame in document._all_drawings():
        found = document.drawing(identifier)
        wrapping = found.wrapping
        drawings.append((getattr(found, "kind", "picture"), found.inline, wrapping.get("wrap"),
                         tuple(round(v, 1) for v in found.size), getattr(found, "text", "") if hasattr(found, "kind")
                         else ""))
    controls = [(c.kind, c.text) for c in document.content_controls()]
    return {"tables": tables, "drawings": sorted(drawings), "controls": controls}


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_a_word_resave_keeps_tables_drawings_and_controls(word, path):
    document, data = _edited(path)
    outcome = word.resave(data, name=f"e5-{path.stem}")
    assert outcome, f"Word did not save {path.name}: {outcome.outcome} {outcome.detail}"
    saved = Document.open(outcome.path.read_bytes())
    # Every edit reads back from Word's file as from ours (Word renumbers drawings' and
    # controls' nothing: their ids are kept).
    assert {name: check(saved) for name, check in document.checks.items()} == \
        {name: True for name in document.checks}
    ours, theirs = _structure(Document.open(data)), _structure(saved)
    assert theirs["tables"] == ours["tables"]
    assert theirs["drawings"] == ours["drawings"]
    assert theirs["controls"] == ours["controls"]


def test_a_word_resave_keeps_a_bound_fill(word):
    document = Document.open(e5_probe.controls_document())
    document.fill_control("cc:1007", "Kept bound")
    outcome = word.resave(document.to_bytes(), name="e5-bound-resave")
    assert outcome, outcome.detail
    saved = Document.open(outcome.path.read_bytes())
    assert saved.content_control("cc:1007").text == "Kept bound"
    xml = next(saved.package.read(n) for n in saved.package.part_names if n.startswith("customXml/item")
               and b"urn:docx-agent:e5" in saved.package.read(n))
    assert b"<name>Kept bound</name>" in xml


# -- tracked: Word's review against ours ------------------------------------------------------------


def _review_view(document: Document) -> dict:
    """What a review leaves, as compared: the final view's Markdown (ids and drawings'
    numbers aside) and every table's grid of texts."""
    lines = [re.sub(r"\((d|cc):[^)]*\)", "", line) for line in document.to_markdown(view="final", ids=False).splitlines()
             if line.strip()]
    tables = [[[" ".join(c.text.split()) if c is not None else None for c in row] for row in table.grid()]
              for story in document.stories for table in story.tables]
    return {"lines": lines, "tables": tables}


#: Where Word's review and ours part: ``same`` everywhere (decision 7's forms are the ones
#: Word reviews); a case Word reviews otherwise would be recorded here with why.
REVIEWED = {name: "same" for name in CASES}


@pytest.mark.parametrize("accept", [True, False], ids=["accept-all", "reject-all"])
@pytest.mark.parametrize("case", list(CASES))
def test_words_review_of_tracked_e5_edits_reads_as_ours(word, case, accept):
    path = ROOT / "tests" / "fixtures" / "samplelib" / "sample-simple.docx"
    prepare, edit = CASES[case]
    document = Document.open(path)
    target = prepare(document) if prepare else None
    with document.tracking(author=AUTHOR, date=DATE):
        edit(document, target)
    data = document.to_bytes()
    action = "accept" if accept else "reject"
    outcome = word.run_script(review_script(action), data, name=f"e5-{case}", tag=action, timeout=200)
    assert outcome, f"Word did not {action} {case}: {outcome.outcome} {outcome.detail}"
    theirs = Document.open(outcome.path.read_bytes())
    ours = Document.open(data)
    ours.accept_all() if accept else ours.reject_all()
    same = _review_view(theirs) == _review_view(ours)
    assert ("same" if same else "differs") == REVIEWED[case], (_review_view(theirs), _review_view(ours))


def test_word_is_gone():
    """Last: no Word left running, no lock file of ours left behind."""
    assert not list(oracle.ORACLE_DIR.glob("~$*"))
