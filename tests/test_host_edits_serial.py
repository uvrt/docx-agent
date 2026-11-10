"""The application's own edits under ``session.lock`` (round 48): committed as the lock goes, so
the next edit on a draft drafts at once, and serializable beside callers' edits on drafts --
``host_stress.run`` replayed one call at a time in commit order: every result, the bytes and
the id state the same.  Small here; long runs: ``python tests/host_stress.py --seeds 0-19
--threads 8 --calls 16``."""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Toolbox
from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

import host_stress

pytestmark = pytest.mark.any_faces


@pytest.mark.parametrize("seed", [0, 1])
def test_the_applications_edits_beside_drafts_equal_their_commit_order_run_alone(seed):
    found = host_stress.run(seed, threads=4, calls=8, chapters=3, first_fill=seed % 2 == 1)
    assert found["same"], found
    assert found["host"] > 0 and found["drafted"] > 0 and found["outside"] == 0


@pytest.mark.slow
@pytest.mark.parametrize("seed", range(2, 6))
def test_the_applications_edits_beside_many_callers_long(seed):
    found = host_stress.run(seed, threads=8, calls=16, chapters=4)
    assert found["same"], found


def _notes() -> bytes:
    document = Document.new()
    document.paragraphs()[0].set_text("Made-up notes")
    for n in range(4):
        document.append_paragraph(f"Paragraph {n} of the notes.")
    return document.to_bytes()


def test_an_edit_after_the_applications_own_drafts_at_once():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0, block_edits="always") as box:
        session = box.session()
        doc = session.open(_notes(), "notes.docx")
        entry = session.entry(doc)
        ids = [p.id for p in entry.document.paragraphs()]
        assert box.dispatch(session, "word_set_text", {"doc": doc, "target": ids[1], "text": "First."}).ok
        before = entry.version
        with session.lock([doc]) as [locked]:
            locked.document.paragraph(ids[2]).insert_after("Added by the application.")
            locked.document.paragraph(ids[3]).set_text("Set by the application.")
        assert entry.version == before + 2 and entry.committed.version == entry.version
        with session.read_lock([doc]) as [reader]:
            assert "Added by the application." in [p.text for p in reader.document.paragraphs()]
        result = box.dispatch(session, "word_set_text", {"doc": doc, "target": ids[4], "text": "After."})
        assert result.ok and result.block["tries"] == 1 and result.block["conflicts"] == []
        with session.lock([doc]):
            pass
        assert entry.version == result.version
