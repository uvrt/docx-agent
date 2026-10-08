"""``examples/review_comments.py`` on the pilot's input (ROADMAP.md, "Usability
(end-to-end pilot)"): every comment acted on as a tight tracked change, replied to and
resolved, nothing else changed, the result valid."""

from __future__ import annotations

import runpy
from pathlib import Path

from docx_agent import Document
from conftest import FIXTURE_DIR

EXAMPLE = Path(__file__).parent.parent / "examples" / "review_comments.py"
PILOT = FIXTURE_DIR / "generated" / "pilot" / "agreement-summary.docx"


def test_review_comments_example(tmp_path):
    target = tmp_path / "reviewed.docx"
    runpy.run_path(str(EXAMPLE), run_name="example")["review"](PILOT, target)
    original, reviewed = Document.open(PILOT), Document.open(target)
    assert reviewed.validate() == []

    final = reviewed.to_markdown(ids=False)
    for text in ("for a period of 36 months,", "EUR 12,500, excluding VAT,", "at most 2.5% per year",
                 "at least five working days in advance", "procurement team (procurement@example.com)."):
        assert text in final, text
    # Rejecting every change gives the original text back: nothing else was touched.
    reviewed.reject_all()
    assert reviewed.to_markdown(view="final") == original.to_markdown(view="final")

    reviewed = Document.open(target)
    revisions = reviewed.revisions()
    assert {r.author for r in revisions} == {"Claude"}
    assert sorted(r.text for r in revisions if r.kind == "deletion") == ["24 months", "3%", "48 hours"]
    threads = reviewed.comments(replies=False)
    assert len(threads) == 5 and all(c.done for c in threads)
    assert all([reply.author for reply in c.replies] == ["Claude"] for c in threads)
