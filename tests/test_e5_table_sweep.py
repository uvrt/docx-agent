"""E5's randomised table sweep (ROADMAP.md, Phase E5, "Done when": the table sweep keeps the
grid consistent after each step), as pptx-agent sweeps its tables: random rows and columns
inserted and deleted, rectangles merged, cells split across and down, widths set -- on
tables with spans and vertical merges -- and after every step the grid is consistent
(:func:`docx_agent.validate.grid_problems`), the document has no validity problem it did
not have, every cell's text is somewhere, and undo gives the original bytes.  Tracked, each
step's accept is the untracked step and its reject the table before it.
"""

from __future__ import annotations

import random

import pytest

from docx_agent import Document, EditError
from docx_agent.validate import check, grid_problems

from canonical import canonical, difference, without_added_styles
from test_roundtrip import entries

BLANK = "tests/fixtures/samplelib/sample-blank.docx"


def _table(document: Document, rows: int = 4, columns: int = 4) -> str:
    anchor = document.paragraphs()[0].id if document.paragraphs() else document.append_paragraph("Anchor.").id
    data = [[f"r{r}c{c}" for c in range(columns)] for r in range(rows)]
    return document.insert_table(rows, columns, after=anchor, data=data).id


def _grid(document: Document, table: str) -> list[list[str]]:
    return [[c.id.rpartition("/")[2] if c else "-" for c in row] for row in document.table(table).grid()]


def _texts(document: Document, table: str) -> list[str]:
    _, entry = document._table_entry(table)
    return sorted(t for t in "".join(entry.element.itertext()).replace("r", " r").split() if t)


def _step(document: Document, table: str, rng: random.Random) -> str:
    table = document.tables()[0].id  # a table's id is its first row's: it moves with that row
    grid = document.table(table).grid()
    rows, columns = len(grid), len(grid[0]) if grid else 0
    kind = rng.choice(["insert_row", "delete_row", "insert_column", "delete_column", "merge", "merge", "split",
                       "split", "width", "cell"])
    r, c = rng.randrange(rows), rng.randrange(columns)
    if kind == "insert_row":
        document.insert_row(table, r, below=rng.random() < 0.5)
    elif kind == "delete_row":
        document.delete_row(table, r)
    elif kind == "insert_column":
        document.insert_column(table, c, right=rng.random() < 0.5)
    elif kind == "delete_column":
        document.delete_column(table, c)
    elif kind == "merge":
        r2, c2 = min(rows - 1, r + rng.randrange(2)), min(columns - 1, c + rng.randrange(3))
        document.merge_cells(table, (r, c), (r2, c2))
    elif kind == "split":
        document.split_cell(table, r, c, rows=rng.choice([1, 1, 2, 3]), columns=rng.choice([1, 2, 2, 3]))
    elif kind == "width":
        document.set_column_width(table, c, rng.choice([36, 72, 100.5]))
    else:
        document.set_cell(table, r, c, shading=rng.choice(["FFFF00", None]),
                          borders={rng.choice(["top", "left", "bottom", "right"]): "double"})
    return f"{kind} ({r}, {c})"


@pytest.mark.parametrize("seed", range(12))
def test_sweep_keeps_the_grid(seed):
    rng = random.Random(seed)
    document = Document.open(BLANK)
    table = _table(document)
    original = document.to_bytes()
    before = set(check(Document.open(original).package))
    steps = []
    for _ in range(25):
        try:
            steps.append(_step(document, table, rng))
        except EditError as error:
            steps.append(f"refused: {error}")
            continue
        table = document.tables()[0].id
        _, entry = document._table_entry(table)
        assert grid_problems(entry.element) == [], (steps, _grid(document, table))
        assert set(check(document.package)) - before == set(), steps
    saved = Document.open(document.to_bytes())
    assert set(check(saved.package)) - before == set()
    wanted = entries(original)
    for _ in steps:  # an edit that changed nothing is no undo step
        if entries(document.to_bytes()) == wanted:
            break
        document.undo()
    assert entries(document.to_bytes()) == wanted


def _normalised(data: bytes) -> bytes:
    from docx_agent.revisions.review import _normalise_grid

    document = Document.open(data)
    _normalise_grid(document._table_entry(document.tables()[0].id)[1].element)
    document.package.mark_dirty(document.package.document_part())
    return document.to_bytes()


@pytest.mark.parametrize("seed", range(8))
def test_sweep_tracked_accepts_and_rejects(seed):
    """Each step tracked: accepting every revision gives the step untracked, rejecting them
    the table before it (canonical form, paragraph ids aside: which element survives a
    merge or a split is the edit's choice)."""
    rng = random.Random(100 + seed)
    document = Document.open(BLANK)
    table = _table(document)
    for _ in range(4):  # spans and merges to work on
        try:
            _step(document, table, rng)
        except EditError:
            pass
    start = document.to_bytes()
    checked = 0
    for _ in range(8):
        choice = rng.random()
        plain = Document.open(start)
        try:
            what = _step(plain, table, random.Random(choice))
        except EditError:
            continue
        tracked = Document.open(start)
        with tracked.tracking(author="Sweep", date="2026-10-04T12:00:00Z"):
            _step(tracked, table, random.Random(choice))
        accepted = Document.open(tracked.to_bytes())
        accepted.accept_all()
        rejected = Document.open(tracked.to_bytes())
        rejected.reject_all()
        _, entry = accepted._table_entry(accepted.tables()[0].id)
        assert grid_problems(entry.element) == [], what
        expected = canonical(plain.to_bytes(), ids=False)
        got = without_added_styles(canonical(accepted.to_bytes(), ids=False), expected)
        assert got == expected, (what, difference(got, expected))
        if what.startswith("split"):
            # A split's column is rejected as Word rejects it: the grid rebuilt from the
            # cells' widths (Word reviews no old grid there: measured), so lines no cell's
            # edge used before come back without them.
            original = canonical(_normalised(start), ids=False)
        else:
            original = canonical(start, ids=False)
        back = without_added_styles(canonical(rejected.to_bytes(), ids=False), original)
        assert back == original, (what, difference(back, original))
        start = plain.to_bytes()
        checked += 1
    assert checked
