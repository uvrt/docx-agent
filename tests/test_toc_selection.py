"""Which paragraphs a table of contents lists, and at which level, held to Word (production
round 21, T1): ``tools/toc_probe.py``'s made-up documents, each with a ``TOC`` field of
``\\o``, ``\\t`` and ``\\u`` alone and combined, updated here and compared with the entries
Word wrote for the same document (``tests/observations/toc-word.json``)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest
from ooxml_common.locales import Locale

from docx_agent import Document
from docx_agent.edit.fields import _heading_text, toc_options

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import toc_probe  # noqa: E402

OBSERVED = json.loads((ROOT / "tests" / "observations" / "toc-word.json").read_text(encoding="utf-8"))
#: The Dutch Mac's regional settings, which the measuring Word ran under.
DUTCH = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy", language="nl-NL",
               list_separator=";")
#: Dutch Word matches ``\t``'s names against its own display names: a built-in heading is
#: "Kop 2" there, "Heading 2" in English Word (``t-heading-en``).  This API matches the
#: names the styles part gives (English Word's), so these differ from what Dutch Word did.
LOCALISED_NAMES = {"t-heading-en-name-nl", "t-heading-nl-name-nl", "t-overrides-style-level"}


def listed(case: str) -> list[list]:
    """What an update lists for the case: ``[level, text]`` per entry (the selection only;
    one case is updated whole below, page numbers and all)."""
    observed = OBSERVED[case]
    document = Document.open(toc_probe.make(observed["body"], observed["instruction"]))
    if observed["locale"] == "nl-NL":
        document.locale = DUTCH
    return [[level, _heading_text(entry.element)] for entry, level in document._toc_entries(observed["instruction"])]


@pytest.mark.parametrize("case", sorted(OBSERVED))
def test_the_entries_are_words(case):
    if case in LOCALISED_NAMES:
        pytest.xfail("Word matched its own language's display names for built-in styles")
    assert listed(case) == OBSERVED[case]["entries"]


def test_the_options_read_as_word_reads_them():
    assert toc_options(' TOC \\o "1-3" \\h \\z \\u ') == {"outline": (1, 3), "own": True, "styles": []}
    assert toc_options(" TOC \\h \\z ")["outline"] == (1, 9)            # neither \o nor \t: 1-9
    assert toc_options(' TOC \\t "A,1" ')["outline"] is None
    assert toc_options(' TOC \\t " A , 2 ,B" \\t "C,0,D,10,E,x" ')["styles"] == [
        ("A", 2), ("B", 1), ("C", 1), ("D", 1), ("E", 1)]
    assert toc_options(' TOC \\t "A;2" ', ";")["styles"] == [("A", 2)]
    assert toc_options(' TOC \\t "A;2" ')["styles"] == [("A;2", 1)]     # en-US: one name


def test_an_update_writes_the_entries_word_wrote():
    observed = OBSERVED["o19-u-t-outline"]
    document = Document.open(toc_probe.make(observed["body"], observed["instruction"]))
    document.locale = DUTCH
    document.update_fields()
    assert toc_probe.entries(document.to_bytes()) == observed["entries"]


def test_a_document_without_locale_splits_on_en_us_comma():
    document = Document.open(toc_probe.make("base", ' TOC \\h \\z \\t "Section title,1,Report title,2" '))
    document.update_fields()
    assert toc_probe.entries(document.to_bytes()) == [
        [1, "Chapter alpha"], [2, "Report name"], [1, "Section own level three"], [1, "Chapter beta"]]
