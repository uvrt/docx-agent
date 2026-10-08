"""Layout in the worker pool, under its deadline (roadmap, "Safety and robustness": Word layout
or reflow 30 s, ``update_fields`` 30 s).  Every layout a Word tool asks for -- describe's page
count, the reflow in checks, ``update_fields``' passes, a TOC's page numbers, saving's
statistics -- goes through ``Document.converter`` to a worker process; past the deadline the
worker is killed, the call reports ``timeout``, and a changing call is rolled back."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Limits, Toolbox  # noqa: E402

from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

LONG = Path(__file__).parent / "fixtures" / "samplelib" / "sample-long.docx"
CLOCK = lambda: dt.datetime(2026, 10, 7, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731


def test_a_reflow_check_on_a_36_page_document_stays_within_its_deadline():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session = box.session(clock=CLOCK)
        d = session.open(LONG.read_bytes(), LONG.name)
        first = box.dispatch(session, "check", {"doc": d, "include": ["reflow"]})
        assert first.ok and first.data["reflow"]["pages"] == 36
        edit = box.dispatch(session, "word_insert_text", {"doc": d, "target": "p@body/3",
                                                          "text": " A sentence that reflows the first page."})
        assert edit.ok and edit.checks["reflow"] == "stale: call check"   # 36 pages: over the 20 for every edit
        again = box.dispatch(session, "check", {"doc": d, "include": ["reflow"]})
        assert again.ok and again.data["reflow"]["since_last_check"]["changed"][:1] == [1]
        assert again.duration < box.limits.layout_timeout, again.duration
        assert session.entry(d).document.converter is None


def test_a_layout_past_its_deadline_is_a_timeout_and_changes_nothing():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, limits=Limits(layout_timeout=0.05)) as box:
        session = box.session(clock=CLOCK)
        d = session.open(LONG.read_bytes(), LONG.name)
        document = session.entry(d).document
        before = document.to_bytes()
        toc = box.dispatch(session, "word_fields", {"doc": d, "action": "insert_toc", "before": "p@body/0"})
        assert not toc.ok and toc.error.code == "timeout"
        assert document.to_bytes() == before                   # rolled back whole
        facts = box.dispatch(session, "check", {"doc": d, "include": ["reflow"]})
        assert facts.ok and facts.data["reflow"]["error"].startswith("timeout")
