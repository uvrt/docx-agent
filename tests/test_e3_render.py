"""E3's representative edit set (``e3_edits.py``: tracked edits of every kind, comments
with a reply and a resolution) on every fixture: what renders and what reads.

* the document stays valid (revision and comment part integrity included);
* docx2svg draws it -- the final view, unmarked -- with the API's ids, no warning but the
  ones the fixture had, and the edits' text where the final view has it;
* ``to_markdown`` reads back as written in every view, names only ids that resolve, and
  its final view is what docx2svg draws, paragraph by paragraph;
* the JSON state lists the edit set's revisions and comments.
"""

from __future__ import annotations

from docx_agent import Document
from docx_agent.markdown import check as markdown_check
from docx_agent.markdown.read import ParagraphRecord, Reader, walk_records
from docx_agent.validate import check

from e3_edits import AUTHOR, COMMENT, REPLY, e3_edit_set
from test_markdown import _COMPUTED, _expected, _normal, ids_in
from test_render import check_rendered_ids


def test_e3_edits_render_and_read(docx_path):
    document = Document.open(docx_path)
    before = document.layout()
    problems_before = set(check(document.package))
    expected = e3_edit_set(document)
    saved = Document.open(document.to_bytes())
    assert set(check(saved.package)) - problems_before == set()

    check_rendered_ids(saved)
    after = saved.layout()
    assert after.unmapped == []
    assert {w.code for w in after.warnings} <= {w.code for w in before.warnings}
    assert (after.stopped is None) or (before.stopped is not None)
    drawn_text = "".join(
        "".join(c for span in line.spans if span.kind == "text" for c in span.chars)
        for page in after._lines for _, line in page)
    if after.stopped is None:
        for text in expected.final:
            assert _normal(text) in _normal(drawn_text), text
    for text in expected.gone:
        assert _normal(text) not in _normal(drawn_text) or _normal(text) in _normal(
            saved.to_markdown(view="final", ids=False)), text

    for view in ("final", "original", "markup"):
        assert markdown_check(saved, view=view, headers=True) == []
        for identifier in ids_in(saved.to_markdown(view=view, headers=True)):
            saved.get(identifier)
    markup = saved.to_markdown(view="markup")
    assert COMMENT in markup and REPLY in markup and AUTHOR in markup

    drawn: dict[str, list[str]] = {}
    for page in after._lines:
        for identifier, line in page:
            if identifier is not None:
                drawn.setdefault(identifier, []).append(
                    "".join(c for span in line.spans if span.kind == "text" for c in span.chars))
    reader = Reader(saved, "final")
    for part in saved._parts():
        if saved._story_of(part).startswith(("header", "footer")):
            continue
        for record in walk_records(reader.story(part)):
            if not isinstance(record, ParagraphRecord) or record.id not in drawn:
                continue
            if any(_COMPUTED.match(f["instruction"]) for f in record.fields):
                continue
            assert _expected(record).fullmatch(_normal("".join(drawn[record.id]))), \
                (record.id, record.text, "".join(drawn[record.id]))

    state = saved.state(view="markup")
    authors = {revision["author"] for revision in state["revisions"]}
    assert AUTHOR in authors
    comment_texts = " ".join(c["t"] for c in state["comments"])
    assert COMMENT in comment_texts
