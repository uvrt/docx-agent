"""List numbers as Word counts them, against docx2svg's recording of Word.

docx2svg measured how Word counts list items (its ROADMAP, "Numbering and lists --
measured", N.1-N.4): ``tools/make_numbering_probe.py`` builds the probe, 43 cases in five
families, and ``tests/fixtures/numbering-observations.json`` records the lines Word drew
for it, each a list label followed by its item's text (``7.B0``: instance B, level 0).
Both are read here from docx2svg's checkout (beside this one, or ``DOCX2SVG_REPO``), so
the two projects cannot drift: the labels the JSON state gives, and the numbers
``to_markdown`` writes, are Word's.
"""

from __future__ import annotations

import io
import json
import os
import re
import sys
from collections import Counter
from pathlib import Path

import pytest

from docx_agent import Document

REPO = Path(os.environ.get("DOCX2SVG_REPO") or Path(__file__).resolve().parents[2] / "docx2svg")
OBSERVATIONS = REPO / "tests" / "fixtures" / "numbering-observations.json"

if not OBSERVATIONS.exists():  # pragma: no cover
    pytest.skip(f"docx2svg's checkout is not at {REPO}", allow_module_level=True)

sys.path.insert(0, str(REPO / "tools"))
import make_numbering_probe as probe  # noqa: E402

DATA = json.loads(OBSERVATIONS.read_text(encoding="utf-8"))["documents"]
#: An item's text in the probe: its instance's letter and its level.
ITEM = re.compile(r"[A-Z][0-8]")


def _word(setting: str, case) -> list[str]:
    """``label+text`` of every list item Word drew for ``case``, in drawing order (a line
    may hold two: a table's cells side by side)."""
    lines = DATA[f"numbering-{setting}"]["word"][case.key] or []
    return [found.group(0) for _, _, _, text in lines if not text.startswith("Case")
            for found in re.finditer(r"(.+?\.)([A-Z][0-8])", text)]


def _ours(document) -> dict[int, list[str]]:
    """``label+text`` of every list item the state gives, per case, in reading order: a
    note's items after the paragraph referencing it, a text box's after its anchor, a
    header's with the case whose section it heads."""
    body = document.state()
    notes = {note["id"]: note["blocks"] for note in body["notes"]}
    headers = {story["name"]: document.state(story["name"])["blocks"]
               for story in body["document"]["stories"] if story["kind"] in ("hdr", "ftr")}
    out: dict[int, list[str]] = {}
    case = None

    def paragraphs(blocks):
        for block in blocks:
            if block.get("type") == "paragraph":
                yield block
                for box in block.get("text_boxes", []) or []:
                    yield from paragraphs(box.get("blocks", []))
                for note in block.get("notes", []) or []:
                    yield from paragraphs(notes.get(note, []))
            elif block.get("type") == "table":
                for row in block.get("rows", []):
                    for cell in row:
                        yield from paragraphs(cell.get("blocks", []))
            elif block.get("type") == "content_control":
                yield from paragraphs(block.get("blocks", []))

    def add(blocks):
        nonlocal case
        for block in paragraphs(blocks):
            text = (block.get("t") or "").strip()
            found = re.match(r"^Case (\d+)$", text)
            if found:
                case = int(found.group(1))
                continue
            listed = block.get("list")
            if listed and listed.get("label") and ITEM.fullmatch(text) and case is not None:
                out.setdefault(case, []).append(listed["label"] + text)

    for block in body["blocks"]:
        add([block])
        section = block.get("section")
        if section is not None or block is body["blocks"][-1]:
            for sec in body["document"]["sections"]:
                if sec["id"] == (section or {}).get("id", "s:body"):
                    for name in set((sec.get("headers") or {}).values()) | set((sec.get("footers") or {}).values()):
                        add(headers.get(name, []))
    return out


@pytest.fixture(scope="module", params=list(probe.SETTINGS))
def probe_document(request):
    document = Document.open(io.BytesIO(probe.build(request.param)))
    return request.param, document


def test_the_labels_are_words(probe_document):
    setting, document = probe_document
    ours = _ours(document)
    wrong = {}
    groups: dict = {}
    story = set()
    for number, case in enumerate(probe.CASES):
        # A header is drawn on every page it heads: its cases compare as one set.  Notes,
        # cells side by side and text boxes are drawn where Word puts them, not in reading
        # order: the story family's cases compare as multisets.
        key = "header" if "header" in case.key else case.key
        if case.family == "story":
            story.add(key)
        word, mine = groups.setdefault(key, ([], []))
        word += _word(setting, case)
        mine += ours.get(number, [])
    for key, (word, mine) in groups.items():
        same = (set(word) == set(mine) if key == "header" else
                Counter(word) == Counter(mine) if key in story else mine == word)
        if not same:
            wrong[key] = (word, mine)
    assert not wrong, json.dumps(wrong, indent=1)


def test_markdown_writes_words_numbers(probe_document):
    setting, document = probe_document
    markdown = document.to_markdown(ids=False)
    numbers = [int(n) for n in re.findall(r"^\s*(\d+)[.)] [A-Z][0-8]$", markdown, flags=re.M)]
    expected = []
    for case in probe.CASES:
        if case.family == "story":
            continue
        for line in _word(setting, case):
            label = line[:-2]
            if re.fullmatch(r"\d+\.", label):
                expected.append(int(label[:-1]))
    # Every decimal, single-number label of the body's non-story cases appears, in order,
    # among the numbers the Markdown writes for its ordered items.
    it = iter(numbers)
    assert all(any(n == e for n in it) for e in expected)
