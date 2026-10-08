"""Word itself on the golden transcripts' outputs, opt-in (``pytest -m oracle``): every file
the tools wrote for the trial's Word tasks -- and the batch of 30 tracked edits and 10
comments -- exports from Word unprompted, showing the text the task asked for (in Word's
final view where the document has revisions or comments)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

import goldens_replay  # noqa: E402
import oracle  # noqa: E402

pytestmark = pytest.mark.oracle

HERE = Path(__file__).parent
GOLDENS = HERE / "goldens"

#: Text Word must show in each output's final view.
EXPECTED = {
    "w1-template-report": ["Annual Water Quality Report 2025", "Executive summary", "Nitrate (mg/l)"],
    "w2-review-comments": ["initial term is 36 months", "99.9%", "Final"],
    "w3-figures-update": ["EUR 57.1 million", "Q4 was the strongest quarter"],
    "w4-restructure": ["4 Data retention", "6 Incident response"],
    "w5-selective-review": ["Review summary", "Team leads must agree"],
    "w6-merge-section": ["Working at height", "stepladders on uneven ground"],
    "batch-30-edits": ["(Reviewed: point 1.)", "Note 1:"],
}


@pytest.fixture(scope="module")
def word():
    try:
        with oracle.session() as session:
            yield session
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))


def _inputs(transcript: dict) -> Path:
    if "fixture" in transcript:
        return (HERE / "fixtures" / transcript["fixture"]).parent
    return GOLDENS / "trial" / "_inputs" / transcript["task"] / "input"


@pytest.mark.parametrize("task", sorted(EXPECTED))
def test_word_opens_every_golden_output(word, task):
    transcript = json.loads((GOLDENS / "transcripts" / f"{task}.json").read_text(encoding="utf-8"))
    _, outputs = goldens_replay.run(transcript, _inputs(transcript))
    outcome = word.export_pdf(outputs[-1].data, name=f"tools-{task}")
    assert outcome, f"Word did not open it: {outcome.outcome} {outcome.detail}"
    text = " ".join(" ".join(oracle.pdf_pages(outcome.path)).split())
    missing = [want for want in EXPECTED[task] if want not in text]
    assert not missing, (missing, text[:400])
