"""Golden transcripts: the trial's Word tasks w1-w6 done with the tools only, replayed.

Each transcript (``goldens/transcripts``) is the tool calls that do one task of the end-to-end
trial -- the calls an agent with the tools makes, no Python -- with every call's normalised
result.  Replaying one on the task's inputs must give the same results, an output that is
byte for byte the one recorded (the session clock is fixed), and an output that passes the
trial's own check (``goldens/trial/_grading/<task>/check.py``, recovered unchanged).  So
"the tools can do every trial task without code" is a regression test.

Byte for byte holds where Office's faces are installed (the Mac the transcripts were recorded
on): the layout -- page counts, reflow, coverage, field page numbers -- is in the results and
in the saved file.  Every other runner, CI's included, replays them too and compares what does
not depend on the layout: each result's ok, summary, ids, warnings, the schema of its data
(``data_keys``) and the digest of its data less what the layout measures (``data_sha_core``),
as ``goldens_replay`` splits them.  So a result that gains or loses a key fails on every
runner, not only on a Mac with Office.  ``tools/refresh_goldens.py`` re-records them (CONTRIBUTING.md,
"Golden transcripts").

``batch-30-edits`` is one ``batch`` call writing 30 tracked insertions and 10 comments
anchored on them by refs.
"""

from __future__ import annotations

import functools
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

import goldens_replay  # noqa: E402
from conftest import OFFICE_FACES_ABSENT  # noqa: E402

HERE = Path(__file__).parent
GOLDENS = HERE / "goldens"
TRANSCRIPTS = sorted((GOLDENS / "transcripts").glob("*.json"))


def inputs_of(transcript: dict) -> Path:
    if "fixture" in transcript:
        return (HERE / "fixtures" / transcript["fixture"]).parent
    return GOLDENS / "trial" / "_inputs" / transcript["task"] / "input"


@functools.lru_cache(maxsize=None)
def replayed(path: Path) -> tuple[dict, list, list]:
    """The transcript, its replayed results (normalised) and outputs: one replay for both
    tests below."""
    transcript = json.loads(path.read_text(encoding="utf-8"))
    results, outputs = goldens_replay.replay(transcript, inputs_of(transcript))
    return transcript, results, outputs


def passes_its_check(transcript: dict, output, tmp_path: Path) -> None:
    check = GOLDENS / "trial" / "_grading" / transcript["task"] / "check.py"
    if not check.exists():
        return
    target = tmp_path / output.name
    target.write_bytes(output.data)
    proc = subprocess.run([sys.executable, str(check), str(target)], capture_output=True, text=True,
                          encoding="utf-8", env={**os.environ, "PYTHONUTF8": "1"}, timeout=600)
    assert proc.stdout, proc.stderr
    report = json.loads(proc.stdout)
    failed = [r for r in report["results"] if not r["ok"]]
    assert not failed and report["passed"] == report["total"], failed


@pytest.mark.parametrize("path", TRANSCRIPTS, ids=[p.stem for p in TRANSCRIPTS])
def test_a_golden_transcript_replays_to_the_same_results_and_bytes(path, tmp_path):
    if OFFICE_FACES_ABSENT:
        # The transcripts' layout facts (pages, reflow, coverage, field page numbers, the page
        # count saved in app.xml) were recorded on macOS with Office's faces; a runner
        # without them lays out with substitutes, or stops.  The core replay below holds the
        # rest there.
        pytest.skip("Office's faces are not all installed here: the core replay compares what "
                    "does not depend on them")
    transcript, results, outputs = replayed(path)
    mismatches = [(index, call["tool"], call["expect"], got)
                  for index, (call, got) in enumerate(zip(transcript["calls"], results))
                  if got != call["expect"]]
    assert mismatches == [], mismatches[:2]
    assert outputs, "the transcript saved nothing"
    output = outputs[-1]
    assert output.name == transcript["output"]["name"]
    assert hashlib.sha256(output.data).hexdigest() == transcript["output"]["sha256"]
    assert output.validate["new"] == []
    passes_its_check(transcript, output, tmp_path)


@pytest.mark.any_faces
@pytest.mark.parametrize("path", TRANSCRIPTS, ids=[p.stem for p in TRANSCRIPTS])
def test_a_golden_transcript_replays_to_the_same_schema_and_core_anywhere(path, tmp_path):
    """Needs no Office face: what the layout measures (``goldens_replay``'s split) is left
    out, everything else -- each result's data keys included -- must be as recorded."""
    transcript, results, outputs = replayed(path)
    mismatches = goldens_replay.core_mismatches(transcript, results)
    assert mismatches == [], mismatches[:2]
    assert outputs, "the transcript saved nothing"
    output = outputs[-1]
    assert output.name == transcript["output"]["name"]
    assert output.validate["new"] == []
    if not goldens_replay.writes_layout(transcript):  # else its page numbers are this runner's
        passes_its_check(transcript, output, tmp_path)


def test_the_trial_tasks_are_done_by_tools_alone():
    names = {json.loads(p.read_text(encoding="utf-8"))["task"] for p in TRANSCRIPTS}
    assert {f"w{n}" for n in range(1, 7)} <= {name[:2] for name in names}
    for path in TRANSCRIPTS:
        for call in json.loads(path.read_text(encoding="utf-8"))["calls"]:
            assert call["expect"]["ok"], (path.stem, call["tool"])


def test_one_batch_writes_30_tracked_edits_and_10_comments_as_one_step():
    from docx_agent import Document

    path = GOLDENS / "transcripts" / "batch-30-edits.json"
    transcript = json.loads(path.read_text(encoding="utf-8"))
    _, outputs = goldens_replay.run(transcript, inputs_of(transcript))
    document = Document.open(outputs[-1].data)
    insertions = [r for r in document.revisions() if r.author == "Claude"]
    assert len(document.changes(author="Claude")) == 30 and all(r.kind == "insertion" for r in insertions)
    assert len(document.comments()) == 10 and document.validate() == []
    batch = next(call for call in transcript["calls"] if call["tool"] == "batch")
    assert len(batch["expect"]["created"]) == 10 and len(batch["expect"]["refs"]) == 30
