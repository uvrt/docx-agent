"""Edits on drafts are serializable (round 45, lane B): ``block_stress.run`` -- callers in
threads editing one made-up report at once, their own chapters and each other's, every
admitted edit prepared on a draft -- replayed one call at a time in the order the changes were
committed: every result, the bytes and the id state the same.  Small here; long runs:
``python tests/block_stress.py --seeds 0-99 --threads 8 --calls 16``."""

from __future__ import annotations

import pytest

import block_stress

pytestmark = pytest.mark.any_faces


@pytest.mark.parametrize("seed, first_fill", [(0, False), (1, True)])
def test_concurrent_edits_equal_their_commit_order_run_alone(seed, first_fill):
    found = block_stress.run(seed, threads=4, calls=8, chapters=3, first_fill=first_fill)
    assert found["same"], found
    assert found["drafted"] > 0


@pytest.mark.parametrize("seed", [0, 1])
def test_concurrent_tracked_edits_equal_their_commit_order_run_alone(seed):
    """Round 50 (R97): every edit a tracked change -- revision ids drawn on drafts, drawn
    again when each patch applies -- the same as one call at a time, every id once."""
    found = block_stress.run(seed, threads=4, calls=8, chapters=3, first_fill=bool(seed % 2), tracked=True)
    assert found["same"], found
    assert found["drafted"] > 0 and found["revisions"] > 0 and not found["repeated_ids"]


@pytest.mark.slow
@pytest.mark.parametrize("seed", range(2, 8))
def test_many_callers_long(seed):
    found = block_stress.run(seed, threads=8, calls=16, chapters=4)
    assert found["same"], found


def test_a_call_that_drew_ids_then_failed_says_the_version_it_ran_on(monkeypatch):
    """0.1.98 (ooxml-edit 0.112.1): the replay's flake.  A batch that drew ids, then failed,
    runs again under the document's lock (the ids it drew there are spent); its result said
    the version read once the lock was let go -- another caller's change, made after it -- so
    the replay ran it after that change, and the ids that change drew came out otherwise."""
    import threading

    from ooxml_edit.tools import DocumentEntry

    from docx_agent import Document

    data = block_stress.made(3)
    layout = block_stress.Layout(Document.open(data))
    with block_stress._box("always") as box:
        session = box.session(clock=block_stress.CLOCK)
        doc = session.open(data, "report.docx")
        block_stress.track(box, session, doc)
        failing = threading.local()
        others: list = []
        leave = DocumentEntry.leave

        def leave_then_another_changes_it(self):
            leave(self)
            if getattr(failing, "now", False):
                failing.now = False
                other = threading.Thread(target=lambda: others.append(box.dispatch(
                    session.view(context={"agent": "other"}), "word_insert_markdown",
                    {"doc": doc, "at": f"after:{layout.bodies[1][0]}", "markdown": "Another caller's paragraph."})))
                other.start()
                other.join()

        monkeypatch.setattr(DocumentEntry, "leave", leave_then_another_changes_it)
        before = session.entry(doc).version
        failing.now = True
        failed = box.dispatch(session.view(context={"agent": "me"}), "batch", {"atomic": True, "ops": [
            {"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": "end", "markdown": "| A | B |\n|---|---|\n| V | 4 |"}},
            {"tool": "word_set_text", "arguments": {"doc": doc, "target": "p:7FFFFFF0", "text": "x"}}]})
        assert not failed.ok and failed.block and failed.block.get("lock")
        assert others and others[0].ok and others[0].version == before + 1
        assert failed.version == before
