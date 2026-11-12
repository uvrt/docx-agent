"""Round 53, R109: ``check`` lists what ``save_document`` would refuse.

The save gate compares the document's validation problems with those it had when opened
(ooxml-edit's :mod:`~ooxml_edit.tools.savegate`); every ``check`` result carries the same
comparison as ``validation`` -- whatever its ``include`` or ``range`` -- and ``validate``'s
own delta is read through it, so the three never disagree.  A made-up "Shelf" document is
edited into three new problems by a test tool (paragraph styles the document lacks: what a
copy without its styles or a hand-made patch leaves), then fixed with ``word_format``.
"""

from __future__ import annotations

import datetime as dt

import pytest

pytest.importorskip("ooxml_edit.tools.savegate", reason="needs ooxml-edit 0.111.0")

from lxml import etree  # noqa: E402
from ooxml_edit.tools import Result, Toolbox, string, tool  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

pytestmark = pytest.mark.any_faces

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CLOCK = lambda: dt.datetime(2026, 11, 12, tzinfo=dt.timezone.utc)  # noqa: E731
STYLES = ["ShelfTop", "ShelfMiddle", "ShelfLow"]
NEW = sorted(f"[style-missing] word/document.xml: pStyle {name}" for name in STYLES)


@tool("name_styles", "Test only: give paragraphs a paragraph style the document lacks.",
      {"doc": string("Document id."), "paragraphs": string("Comma-separated p: ids.")},
      kind="docx", mutates=True)
def name_styles(call, doc, paragraphs):
    document = call.document
    for ref, style in zip(paragraphs.split(","), STYLES):
        paragraph = document._resolve(ref)[1].element
        properties = paragraph.find(W + "pPr")
        if properties is None:
            properties = etree.Element(W + "pPr")
            paragraph.insert(0, properties)
        node = properties.find(W + "pStyle")
        if node is None:
            node = etree.Element(W + "pStyle")
            properties.insert(0, node)
        node.set(W + "val", style)
    document._invalidate()
    return Result(summary="styled", changed=True)


@pytest.fixture
def box():
    with Toolbox([*TOOLS, name_styles], formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _ok(box, session, name, /, **arguments):
    result = box.dispatch(session, name, arguments)
    assert result.ok, (name, result.error and result.error.to_json())
    return result


def _shelf(box):
    """The made-up document: three places, each a heading, a table and an outcome line."""
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "shelf.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    md = "# Shelf\n\nWhat the shelf holds, by place.\n\n"
    for place in ("Top", "Middle", "Low"):
        md += (f"## {place}\n\n| item | count | notes |\n|---|---|---|\n| jar | 3 | glass |\n"
               f"| tin | 5 | partly done |\n\nOutcome: {place.lower()} shelf partly done.\n\n")
    _ok(box, session, "word_insert_markdown", doc="d1", at=f"after:{first}", markdown=md)
    document = session.entry("d1").document
    outcomes = [p.id for p in document.paragraphs() if p.text.startswith("Outcome")]
    intro = next(p.id for p in document.paragraphs() if p.text.startswith("What the shelf"))
    return session, outcomes, intro


def _save(box, session):
    return box.dispatch(session, "save_document", {"doc": "d1", "name": "shelf.docx",
                                                   "format": "docx"})


@pytest.mark.parametrize("asked", [{"include": ["layout"]}, {"range": "intro"},
                                   {"include": ["fields"]}, {}])
def test_check_lists_the_problems_the_save_refuses_whatever_it_was_asked(box, asked):
    session, outcomes, intro = _shelf(box)
    _ok(box, session, "name_styles", doc="d1", paragraphs=",".join(outcomes))
    arguments = {"doc": "d1", **({"range": intro} if "range" in asked else asked)}
    checked = _ok(box, session, "check", **arguments)
    refused = _save(box, session)
    assert refused.error.code == "refused"
    fact = checked.data["validation"]
    assert sorted(fact["new"]) == NEW and fact["baseline"] == 0
    assert fact["new"] == refused.error.valid_options      # the same words, the same order
    assert "saving would refuse 3 new validation problem(s)" in checked.summary


def test_the_fix_with_the_tools_clears_it_and_the_save_passes(box):
    session, outcomes, _ = _shelf(box)
    _ok(box, session, "name_styles", doc="d1", paragraphs=",".join(outcomes))
    fixed = _ok(box, session, "word_format", doc="d1", targets=outcomes, paragraph_style="Normal")
    assert fixed.checks["validate"]["new"] == []
    checked = _ok(box, session, "check", doc="d1", include=["layout"])
    assert checked.data["validation"] == {"new": [], "baseline": 0}
    assert "saving would refuse" not in checked.summary
    saved = _save(box, session)
    assert saved.ok and saved.data["validate"]["new"] == []


def test_validate_the_calls_checks_and_validation_agree(box):
    session, outcomes, intro = _shelf(box)
    edited = _ok(box, session, "name_styles", doc="d1", paragraphs=",".join(outcomes))
    checked = _ok(box, session, "check", doc="d1")
    assert edited.checks["validate"]["new"] == checked.data["validate"]["new"] \
        == checked.data["validation"]["new"]
    # scoped: validate names those in the range (none here), validation the whole document
    scoped = _ok(box, session, "check", doc="d1", range=intro)
    assert scoped.data["validate"]["new"] == [] and scoped.data["validate"]["elsewhere"] == 3
    assert scoped.data["validation"]["new"] == checked.data["validation"]["new"]


def test_the_problems_are_found_once_per_version(box, monkeypatch):
    session, outcomes, _ = _shelf(box)
    found = []
    real = Document.validate

    def counted(self, *args, **kwargs):
        found.append(1)
        return real(self, *args, **kwargs)

    monkeypatch.setattr(Document, "validate", counted)
    _ok(box, session, "name_styles", doc="d1", paragraphs=",".join(outcomes))
    assert len(found) == 1                                  # the call's checks
    for asked in ([], ["validate"], ["layout"]):
        _ok(box, session, "check", doc="d1", **({"include": asked} if asked else {}))
    assert len(found) == 1                                  # kept for the version
    assert _save(box, session).error.code == "refused"
    assert len(found) == 2                                  # the gate finds them afresh


def test_a_check_inside_a_batch_sees_the_batchs_own_edit(box):
    session, outcomes, _ = _shelf(box)
    _ok(box, session, "name_styles", doc="d1", paragraphs=",".join(outcomes))
    _ok(box, session, "check", doc="d1")                    # kept for this version
    result = _ok(box, session, "batch", ops=[
        {"tool": "word_format", "arguments": {"doc": "d1", "targets": outcomes,
                                              "paragraph_style": "Normal"}},
        {"tool": "check", "arguments": {"doc": "d1", "include": ["validate"]}}])
    inside = result.data["ops"][1]["data"]
    assert inside["validation"]["new"] == [] and inside["validate"]["new"] == []
    assert _save(box, session).ok
