"""E4's tracked operations, on every fixture (ROADMAP.md, Phase E4; the invariant of
"Accept and reject"): tracked, saved and reopened, no validity problem added and every new
revision ours; accepting them all gives the untracked edit and rejecting them the original,
in canonical form.

What the comparison forgives, each as Word does it (``tests/canonical.py``'s rules and):

* what is not a revision in Word and stays when the edit is rejected: style definitions,
  bookmarks (a table of contents' ``_Toc`` ones, a cross-reference's target), notes parts
  left with only their separators, a header part a tracked break shared;
* a tracked section break leaves the split section's header and footer references where
  they are (references are not revisions) where the untracked break moves them to the new
  section: the sections show the same stories either way, which is checked instead;
* notes are compared by their order, not their ids (a tracked move is a copy of the note).
"""

from __future__ import annotations

import re

import pytest

from docx_agent import Document
from docx_agent.validate import check

from canonical import canonical, difference, without_added_definitions
from e4_edits import TRACKED

AUTHOR = "E4 Agent"
DATE = "2026-10-04T12:00:00Z"
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _name(operation) -> str:
    return operation.__name__[3:]


def _notes_in_order(form: dict[str, str]) -> dict[str, str]:
    """Footnote and endnote ids renumbered by their references' order in the body."""
    form = dict(form)
    body = form.get("word/document.xml", "")
    for kind in ("footnote", "endnote"):
        order: dict[str, str] = {}
        for found in re.findall(rf"}}{kind}Reference \{{[^}}]*\}}id='(-?\d+)'", body):
            order.setdefault(found, f"n{len(order)}")
        if not order:
            continue
        body = re.sub(rf"(}}{kind}Reference \{{[^}}]*\}}id=')(-?\d+)'",
                      lambda m: m.group(1) + order.get(m.group(2), "x") + "'", body)
        part = f"word/{kind}s.xml"
        if part in form:
            text = re.sub(rf"(}}{kind} \{{[^}}]*\}}id=')(\d+)'",
                          lambda m: m.group(1) + order.get(m.group(2), "gone") + "'", form[part])
            # A note no reference names any more (a rejected copy) is not in the document.
            lines = text.split("\n")
            out, skip = [], None
            for line in lines:
                indent = len(line) - len(line.lstrip())
                if skip is not None and indent > skip:
                    continue
                skip = None
                if f"}}{kind} " in line and "id='gone'" in line:
                    skip = indent
                    continue
                out.append(line)
            # Notes by their number, as the references order them.
            form[part] = "\n".join(out)
    form["word/document.xml"] = body
    return form


def _without_references(form: dict[str, str]) -> dict[str, str]:
    form = dict(form)
    form["word/document.xml"] = "\n".join(line for line in form["word/document.xml"].split("\n")
                                          if "}headerReference" not in line and "}footerReference" not in line)
    return form


def _shown(document: Document) -> list[tuple]:
    """What every section shows: its stories' text, per kind."""
    out = []
    for section in document.sections():
        for which in ("header", "footer"):
            for kind in ("default", "first", "even"):
                story = getattr(section, which)(kind)
                out.append((which, kind, "\n".join(p.text for p in story.paragraphs) if story else None))
    return out


def _without_empty_paragraphs(text: str) -> str:
    """A canonical document part without its body's paragraphs that hold no run."""
    lines = text.split("\n")
    out: list[str] = []
    k = 0
    while k < len(lines):
        line = lines[k]
        indent = len(line) - len(line.lstrip())
        if line.strip().split(" ")[0].endswith("main}p"):
            end = k + 1
            while end < len(lines) and len(lines[end]) - len(lines[end].lstrip()) > indent:
                end += 1
            if not any(b.strip().split(" ")[0].endswith(("main}r", "main}sectPr")) for b in lines[k + 1:end]):
                k = end
                continue
        out.append(line)
        k += 1
    return "\n".join(out)


def _forgiving(form: dict[str, str], against: dict[str, str]) -> dict[str, str]:
    form = _notes_in_order(without_added_definitions(form, against))
    if _without_empty_paragraphs(against.get("word/document.xml", "")) == against.get("word/document.xml") \
            and not re.search(r"main\}p( |$)", against.get("word/document.xml", ""), re.M):
        # A body with no paragraph at all (samplelib's blank): the last paragraph an edit
        # adds keeps its mark, which Word neither accepts nor rejects (E2's record).
        form = {**form, "word/document.xml": _without_empty_paragraphs(form["word/document.xml"])}
    return form


@pytest.mark.parametrize("operation", TRACKED, ids=_name)
def test_accept_all_is_the_edit_and_reject_all_the_original(docx_path, operation):
    data = docx_path.read_bytes()
    name = _name(operation)
    if name == "remove_break" and len(Document.open(data).sections()) < 2:
        # The operation would insert the break it removes: one's own insertion, removed,
        # goes outright (no revision); test_e4_word_forms holds a tracked removal.
        pytest.skip("one section: nothing to remove")
    base = Document.open(data)
    base._stamp_document()
    original = _notes_in_order(canonical(base.to_bytes(), ids=False, bookmarks=False))

    untracked = Document.open(data)
    operation(untracked)
    expected = _notes_in_order(canonical(untracked.to_bytes(), ids=False, bookmarks=False))

    tracked = Document.open(data)
    with tracked.tracking(author=AUTHOR, date=DATE):
        read_back = operation(tracked)
    tracked_bytes = tracked.to_bytes()
    saved = Document.open(tracked_bytes)
    assert set(check(saved.package)) - set(check(Document.open(data).package)) == set()
    ours = [r for r in saved.revisions() if r.author == AUTHOR]
    assert ours and all(r.date == DATE for r in ours)

    accepted = Document.open(tracked_bytes)
    accepted.accept(author=AUTHOR)
    assert not accepted.revisions(author=AUTHOR)
    got = _forgiving(canonical(accepted.to_bytes(), ids=False, bookmarks=False), expected)
    want = expected
    if name == "section_break":
        assert _shown(accepted) == _shown(untracked)
        got, want = _without_references(got), _without_references(expected)
    assert got == want, "accept-all is not the untracked edit:\n" + difference(want, got)
    assert read_back(accepted) or name in ("notes",)

    rejected = Document.open(tracked_bytes)
    rejected.reject(author=AUTHOR)
    assert not rejected.revisions(author=AUTHOR)
    got = _forgiving(canonical(rejected.to_bytes(), ids=False, bookmarks=False), original)
    assert got == original, "reject-all is not the original:\n" + difference(original, got)
