"""Golden transcripts: the trial's Word tasks w1-w6 done with the tools only, replayed.

Each transcript (``goldens/transcripts``) is the tool calls that do one task of the end-to-end
trial -- the calls an agent with the tools makes, no Python -- with every call's normalised
result.  Replaying one on the task's inputs must give the same results, an output that is
byte for byte the one recorded (the session clock is fixed), and an output that passes the
trial's own check (``goldens/trial/_grading/<task>/check.py``, recovered unchanged).  So
"the tools can do every trial task without code" is a regression test.

``batch-30-edits`` is one ``batch`` call writing 30 tracked insertions and 10 comments
anchored on them by refs.
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

import goldens_replay  # noqa: E402

HERE = Path(__file__).parent
GOLDENS = HERE / "goldens"
TRANSCRIPTS = sorted((GOLDENS / "transcripts").glob("*.json"))


def inputs_of(transcript: dict) -> Path:
    if "fixture" in transcript:
        return (HERE / "fixtures" / transcript["fixture"]).parent
    return GOLDENS / "trial" / "_inputs" / transcript["task"] / "input"


@pytest.mark.parametrize("path", TRANSCRIPTS, ids=[p.stem for p in TRANSCRIPTS])
def test_a_golden_transcript_replays_to_the_same_results_and_bytes(path, tmp_path):
    transcript = json.loads(path.read_text(encoding="utf-8"))
    replayed, outputs = goldens_replay.run(json.loads(path.read_text(encoding="utf-8")), inputs_of(transcript))
    assert replayed["_mismatches"] == [], replayed["_mismatches"][:2]
    assert outputs, "the transcript saved nothing"
    output = outputs[-1]
    assert output.name == transcript["output"]["name"]
    assert hashlib.sha256(output.data).hexdigest() == transcript["output"]["sha256"]
    assert output.validate["new"] == []
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
