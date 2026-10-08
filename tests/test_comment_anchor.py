"""``Comment.anchor``: the text a comment is attached to, not the comment's own paragraphs
(the end-to-end pilot read ``paragraph_ids`` as the former; ROADMAP.md, "Usability")."""

from __future__ import annotations

import pytest

from docx_agent import Document, TextRange
from conftest import FIXTURE_DIR

PILOT = FIXTURE_DIR / "generated" / "pilot" / "agreement-summary.docx"


def test_anchor_is_the_commented_text():
    document = Document.open(PILOT)
    anchors = {c.id: c.anchor for c in document.comments()}
    assert {k: a.text for k, a in anchors.items()} == {
        "c:76FCC91F": "24 months", "c:4C9866FD": "EUR 12,500", "c:22CB2567": "at most 3% per year",
        "c:5E432F44": "48 hours", "c:777AEE05": "procurement team"}
    first = anchors["c:76FCC91F"]
    assert isinstance(first, TextRange) and first.paragraph_ids() == ["p:3B212964"]
    assert document.range(first.id).text == "24 months"


def test_a_reply_is_attached_where_its_thread_is():
    document = Document.open(PILOT)
    reply = document.reply_to_comment("c:4C9866FD", "Added.", author="Claude").id
    assert document.comment(reply).anchor.text == "EUR 12,500"


def test_a_comment_on_a_place_has_an_empty_anchor():
    document = Document.open(PILOT)
    at = document.anchor("starting on")
    place = TextRange(document, at.start_id, at.start)
    created = document.add_comment(place, "Here.").id
    anchor = document.comment(created).anchor
    assert anchor.collapsed and anchor.start == at.start and anchor.text == ""


def test_content_paragraph_ids_and_the_deprecated_alias():
    comment = Document.open(PILOT).comment("c:76FCC91F")
    assert comment.content_paragraph_ids == ["comments/p:6CAAAF88"]
    with pytest.warns(DeprecationWarning, match="Comment.anchor"):
        assert comment.paragraph_ids == comment.content_paragraph_ids


def test_comments_without_replies_lists_each_thread_once():
    document = Document.open(PILOT)
    reply = document.comment("c:76FCC91F").reply("Changed.", author="Claude").id
    assert reply in [c.id for c in document.comments()]
    roots = document.comments(replies=False)
    assert len(roots) == 5 and reply not in [c.id for c in roots]
    assert [r.id for r in document.comment("c:76FCC91F").replies] == [reply]
