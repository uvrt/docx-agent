"""Formatting changes on what is one's own pending insertion (round 50, R98), as Word records
them (``tools/tracking_probe.py``, ``tests/observations/tracking-word.json``):

* paragraphs typed one after another at a container's end keep no ``w:pPrChange`` but on the
  last one, which holds the container's old mark -- the change recorded on a mark passes on
  with it when the paragraph before gets an inserted mark;
* formatting given in the call that inserts records nothing;
* a later call that formats what one inserted earlier is recorded, as Word records it, and so
  is formatting of what was there before.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest
from lxml import etree
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

ROOT = Path(__file__).resolve().parent.parent
OBSERVED = json.loads((ROOT / "tests" / "observations" / "tracking-word.json").read_text(encoding="utf-8"))
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
AUTHOR = "Writer"


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _base(*texts: str) -> bytes:
    document = Document.new()
    document.insert_markdown("\n\n".join(texts), at="end")
    return document.to_bytes()


def _opened(box, data: bytes):
    session = box.session()
    doc = session.open(data, "plan.docx")
    ok(box, session, "word_set_tracking", doc=doc, on=True, author=AUTHOR)
    return session, doc, session.entry(doc).document


def ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, result.error and result.error.to_json())
    return result


def _changes(document: Document) -> list[str]:
    """The text of each body paragraph holding a ``w:pPrChange``."""
    return [p.text for p in document.paragraphs() if p._element.find(f"{W}pPr/{W}pPrChange") is not None]


def _xml(document: Document) -> str:
    return etree.tostring(document.package.tree(document.package.document_part())).decode()


def test_word_keeps_a_property_change_only_on_the_last_paragraph_typed_at_the_end():
    # The measured fact this follows: three paragraphs typed at the end, each styled.
    for name in ("end-text-first", "end-style-first"):
        typed = [b for b in OBSERVED[name] if "typed." in b]
        assert len(typed) == 3
        assert all("pPrChange" not in b and "<w:ins " in b.split("</w:pPr>")[0] for b in typed[:2])
        assert "<w:pPrChange" in typed[2] and "<w:rPr><w:ins" not in typed[2]
        # The paragraph the typing started from: its mark inserted, nothing else recorded.
        assert "pPrChange" not in next(b for b in OBSERVED[name] if "Existing last." in b)


def test_sections_appended_at_the_end_record_one_property_change_as_word_does(box):
    session, doc, document = _opened(box, _base("Existing text."))
    for k in range(6):
        ok(box, session, "word_insert_markdown", doc=doc, at="end", markdown=f"## Part {k}\n\n- item {k}")
    # Before round 50: one w:pPrChange per call, each on a paragraph the call inserted.
    assert _changes(document) == ["item 5"]
    paragraphs = document.paragraphs()
    for paragraph in paragraphs[:-1]:
        # Every mark but the last is an insertion of the author's (the first: Word's typing
        # form, the mark of the paragraph the insertion follows).
        mark = paragraph._element.find(f"{W}pPr/{W}rPr/{W}ins")
        assert mark is not None and mark.get(W + "author") == AUTHOR
    # The last records its properties as a change from the old mark's -- the existing
    # paragraph's, as it was.
    old = paragraphs[-1]._element.find(f"{W}pPr/{W}pPrChange/{W}pPr")
    assert old is not None and not len(old)
    kinds = {r.kind for r in document.revisions()}
    assert kinds == {"insertion", "paragraph-mark-insertion", "paragraph-properties"}


def test_appended_sections_reject_to_the_original_and_accept_to_the_text(box):
    session, doc, document = _opened(box, _base("Existing text."))
    for k in range(3):
        ok(box, session, "word_insert_markdown", doc=doc, at="end", markdown=f"## Part {k}\n\n1. step {k}")
    accepted = Document.open(document.to_bytes())
    accepted.accept_all()
    assert [p.text for p in accepted.paragraphs()] == ["Existing text.", "Part 0", "step 0", "Part 1", "step 1",
                                                       "Part 2", "step 2"]
    assert accepted.paragraphs()[1].style_name.lower() == "heading 2" and not accepted.revisions()
    rejected = Document.open(document.to_bytes())
    rejected.reject_all()
    assert [p.text for p in rejected.paragraphs()] == ["Existing text."]
    assert rejected.paragraphs()[0]._element.find(f"{W}pPr/{W}pStyle") is None


def test_a_change_the_last_paragraph_had_passes_on_with_its_mark():
    # A paragraph whose formatting another author changed (tracked) is the body's last; text
    # appended after it, tracked: the record moves to the new last paragraph, which holds the
    # old mark, and still says what the properties were before either edit.
    document = Document.new()
    document.insert_markdown("Lead.\n\nLast one.", at="end")
    last = document.paragraphs()[-1]
    with document.tracking(author="Other", date="2026-11-01T10:00:00Z"):
        document.paragraph(last.id).format(alignment="center")
    assert _changes(document) == ["Last one."]
    with document.tracking(author=AUTHOR, date="2026-11-02T10:00:00Z"):
        document.paragraph(last.id).insert_after("Appended.", style="Heading 1")
    assert _changes(document) == ["Appended."]
    old = document.paragraphs()[-1]._element.find(f"{W}pPr/{W}pPrChange/{W}pPr")
    assert old is not None and old.find(f"{W}jc") is None and old.find(f"{W}pStyle") is None
    rejected = Document.open(document.to_bytes())
    rejected.reject_all()
    assert [p.text for p in rejected.paragraphs()] == ["Lead.", "Last one."]
    assert rejected.paragraphs()[-1]._element.find(f"{W}pPr/{W}jc") is None


def test_formatting_given_with_the_insertion_records_none(box):
    session, doc, document = _opened(box, _base("Existing text.", "Closing text."))
    first = document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", doc=doc, at=f"after:{first}", markdown="Styled.\n\n- a\n- b",
       style_map=[{"element": "paragraph", "style": "Quote"}])
    ok(box, session, "word_set_text", doc=doc, target=first, text="Existing text.\nSplit off.",
       paragraph_style="Normal")
    assert "pPrChange" not in _xml(document) and "rPrChange" not in _xml(document)


def test_a_later_call_formatting_ones_own_insertion_is_recorded_as_word_records_it(box):
    # Measured: Word records a style, an alignment and bold set on a paragraph it had just
    # typed (``format``: Stwonew, its mark and text the user's own insertions).
    typed = next(b for b in OBSERVED["format"] if "Stwonew" in b)
    assert '<w:ins w:id="' in typed and "<w:pPrChange" in typed and "<w:rPrChange" in typed
    session, doc, document = _opened(box, _base("Existing text.", "Closing text."))
    first = document.paragraphs()[0].id
    created = ok(box, session, "word_insert_markdown", doc=doc, at=f"after:{first}", markdown="New paragraph.")
    new = created.created[0] if created.created else document.paragraphs()[1].id
    ok(box, session, "word_format", doc=doc, targets=[new], alignment="center")
    assert _changes(document) == ["New paragraph."]


def test_formatting_existing_text_is_recorded_as_before(box):
    session, doc, document = _opened(box, _base("Existing text.", "Closing text."))
    ok(box, session, "word_insert_markdown", doc=doc, at="end", markdown="Appended.")
    ok(box, session, "word_format", doc=doc, targets=[document.paragraphs()[0].id], alignment="right")
    assert _changes(document) == ["Existing text."]
    assert re.search(r"<w:jc w:val=\"right\"/>", _xml(document))
