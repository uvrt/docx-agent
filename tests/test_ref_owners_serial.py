"""Refs by several callers stay serializable (round 47): ``ref_stress.run`` -- the serial-replay
stress of ``block_stress`` with inserts naming what they make (a thread's own names, and
names every thread gives: the first to define one owns it, the others are refused before
anything is made and may use it) and places and targets given as ``$name`` -- replayed one call
at a time in commit order: every result, the bytes and the id state the same.  Small here;
long runs: ``python tests/ref_stress.py --seeds 0-19 --threads 8 --calls 16``."""

from __future__ import annotations

import pytest

from ooxml_edit.tools import Toolbox
from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

import ref_stress

pytestmark = pytest.mark.any_faces


@pytest.mark.parametrize("seed", [0, 3])
def test_refs_by_several_callers_equal_their_commit_order_run_alone(seed):
    found = ref_stress.run(seed, threads=4, calls=8, chapters=3, first_fill=seed % 2 == 1)
    assert found["same"], found
    assert found["drafted"] > 0 and found["defined"] > 0 and found["taken"] > 0


@pytest.mark.slow
@pytest.mark.parametrize("seed", range(4, 8))
def test_refs_by_many_callers_long(seed):
    found = ref_stress.run(seed, threads=8, calls=16, chapters=4)
    assert found["same"], found


def test_a_name_another_agent_gave_is_refused_before_anything_is_made():
    document = Document.new()
    document.paragraphs()[0].set_text("Made-up notes")
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session()
        doc = session.open(document.to_bytes(), "notes.docx")
        one, two = session.view(context={"agent": "one"}), session.view(context={"agent": "two"})
        made = box.dispatch(one, "word_insert_markdown", {"doc": doc, "at": "end", "markdown": "First", "ref": "table"})
        assert made.ok, made.to_json()
        state = session.entry(doc).document.id_state()
        refused = box.dispatch(two, "word_insert_markdown", {"doc": doc, "at": "end", "markdown": "Second",
                                                             "ref": "table"})
        assert not refused.ok and refused.error.field == "ref" and refused.error.valid_options == ["table_2"]
        assert "agent=one" in refused.error.message
        assert session.entry(doc).document.id_state() == state         # no id drawn
        used = box.dispatch(two, "word_insert_markdown", {"doc": doc, "at": "after:$table", "markdown": "Second",
                                                          "ref": "table_2"})
        assert used.ok, used.to_json()
        assert session.read_refs(doc) == {"table": made.refs["table"], "table_2": used.refs["table_2"]}
        assert session.read_resolve_ref(doc, "$table_2") == used.refs["table_2"]
