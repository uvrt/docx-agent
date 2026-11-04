"""Round 32, R12 and R14: a table of contents' entries for numbered headings, and its right
tab stop next to a ``toc N`` style's own, as Word's own update writes them (measured:
``tools/toc_numbers_probe.py``, ``tests/observations/toc-numbers-word.json``)."""

from __future__ import annotations

import io
import json
import sys
from pathlib import Path

import pytest

from docx_agent import Document

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "tools"))

import toc_number_cases  # noqa: E402
from toc_numbers_probe import entries  # noqa: E402

OBSERVED = json.loads((Path(__file__).parent / "observations" / "toc-numbers-word.json").read_text())["cases"]
CASES = {**toc_number_cases.CASES, **toc_number_cases.CASES2, **toc_number_cases.CASES3}


def updated(name: str) -> list[dict]:
    document = Document.open(CASES[name])
    document.update_fields()
    out = io.BytesIO()
    document.save(out)
    return entries(out.getvalue())


@pytest.mark.parametrize("name", sorted(OBSERVED))
def test_the_entries_are_words(name):
    """Every entry's tab stops and runs as Word wrote them: the number, a tab run and the
    text (or the number, a space and the text in one run), the stop after the number, and
    a right dot stop only where the toc style has no stop of its own."""
    assert updated(name) == OBSERVED[name]


def test_the_number_s_tab_run_takes_the_mark_s_formatting():
    document = Document.open(CASES["n-tab"])
    document.update_fields()
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    entry = next(p._element for p in document.paragraphs() if p.style_name == "toc 1")
    tab = [r for r in entry.find(w + "hyperlink") if r.find(w + "tab") is not None][0]
    assert [c.tag for c in tab.find(w + "rPr")] == [w + "noProof"]
    assert document.validate() == []


def test_an_update_is_stable():
    """A second update writes the same entries (the numbers counted once per update)."""
    document = Document.open(CASES["n-wide"])
    document.update_fields()
    first = [p.text for p in document.paragraphs()]
    document.update_fields()
    assert [p.text for p in document.paragraphs()] == first
