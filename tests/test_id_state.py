"""A document saved and opened again with its id state (round 38): what the file cannot say --
ids deleted after the last surviving one of their sequence, who deleted them -- comes back
from a short string kept beside the bytes, and Word's bytes are not touched."""

from __future__ import annotations

import threading

import pytest

from docx_agent import Document
from docx_agent.edit import ids as _ids
from docx_agent.edit.ids import DeletedId, IdStateError


def every_id(document: Document) -> set[str]:
    out: set[str] = set()
    for story in document.stories:
        for paragraph in story.paragraphs:
            out.add(paragraph.id)
        for table in story.tables:
            out.add(table.id)
            out.update(row.id for row in table.rows)
    return out


def _deleted_at_the_end() -> tuple[Document, list[str]]:
    """Blocks added, the latest ones deleted: no surviving id of the sequence comes after
    them, so the file alone cannot tell they were ever there."""
    document = Document.new()
    document.insert_markdown("\n\n".join(f"Line {k}." for k in range(8)))
    gone = [p.id for p in document.paragraphs()[-4:]]
    for identifier in gone:
        document.delete_block(identifier)
    return document, gone


# -- the limit the file alone leaves ------------------------------------------------------------


def test_without_the_state_the_latest_deleted_ids_come_back_as_before():
    document, gone = _deleted_at_the_end()
    reopened = Document.open(document.to_bytes())
    assert reopened.history.version == 0
    made = [reopened.append_paragraph(f"New {k}.").id for k in range(6)]
    assert set(made) & set(gone)            # the stated limit, unchanged without a state


def test_with_the_state_no_deleted_id_is_given_again_and_naming_one_says_deleted():
    document, gone = _deleted_at_the_end()
    data, state = document.to_bytes(), document.id_state()
    version = document.history.version
    reopened = Document.open(data, ids=state)
    made = [reopened.append_paragraph(f"New {k}.").id for k in range(30)]
    assert not set(made) & set(gone)
    assert len(set(made)) == len(made)
    for identifier in gone:
        with pytest.raises(DeletedId) as caught:
            reopened.paragraph(identifier)
        assert caught.value.earlier and caught.value.version is not None
        message = str(caught.value)
        assert f"{identifier} was deleted in the change that made version {caught.value.version}" in message
        assert "before the document was saved and opened again" in message
        assert caught.value.version <= version


def test_the_bytes_are_the_same_with_or_without_the_state():
    document, _ = _deleted_at_the_end()
    data = document.to_bytes()
    document.id_state()
    assert document.to_bytes() == data
    assert Document.open(data, ids=document.id_state()).to_bytes() == Document.open(data).to_bytes()


def test_no_value_the_document_ever_had_is_generated_after_the_reopen():
    document = Document.new()
    document.insert_markdown("\n\n".join(f"Line {k}." for k in range(10)))
    for paragraph in document.paragraphs()[:5]:
        document.paragraph(paragraph.id).set_text(paragraph.text + " again")   # textIds renewed
    for paragraph in document.paragraphs()[-3:]:
        document.delete_block(paragraph.id)
    pool = set(document._id_pool())
    reopened = Document.open(document.to_bytes(), ids=document.id_state())
    before = set(reopened._id_pool())
    assert pool <= before
    reopened.insert_markdown("\n\n".join(f"More {k}." for k in range(40)))
    new = set(reopened._id_pool()) - before
    assert new and not new & pool


def test_versions_go_on_from_the_saved_one():
    document, _ = _deleted_at_the_end()
    document.undo()
    document.redo()
    saved, last = document.history.version, document.history.last_version
    reopened = Document.open(document.to_bytes(), ids=document.id_state())
    assert reopened.history.version == saved
    reopened.append_paragraph("Next.")
    assert reopened.history.version == last + 1
    assert not reopened.undo() or reopened.history.version == saved


def test_a_state_survives_several_reopens_and_this_session_s_deletions_join_it():
    document, gone = _deleted_at_the_end()
    second = Document.open(document.to_bytes(), ids=document.id_state())
    survivor = second.paragraphs()[-1].id
    second.delete_block(survivor)
    third = Document.open(second.to_bytes(), ids=second.id_state())
    made = [third.append_paragraph(f"New {k}.").id for k in range(20)]
    assert not set(made) & {*gone, survivor}
    for identifier in (*gone, survivor):
        with pytest.raises(DeletedId, match="before the document was saved and opened again"):
            third.paragraph(identifier)
    # Opened without a change, the state it gives is the one it took.
    assert Document.open(third.to_bytes(), ids=third.id_state()).id_state() == third.id_state()


def test_a_replaced_block_still_names_its_replacement():
    document = Document.new()
    document.insert_markdown("One.\n\nTwo.")
    old = document.paragraphs()[-1].id
    new = document.insert_markdown("Two, again.", at=f"replace:{old}").blocks[0]
    reopened = Document.open(document.to_bytes(), ids=document.id_state())
    with pytest.raises(DeletedId) as caught:
        reopened.paragraph(old)
    assert caught.value.replaced_by == new and f"{old} was replaced with {new}" in str(caught.value)


def test_a_deleted_drawing_number_is_not_given_again_after_the_reopen():
    from tests.test_retired_ids import png

    document = Document.new()
    anchor = document.append_paragraph("Picture here.").id
    first = document.insert_picture(f"{anchor}@0", png(), alt_text="One").id
    document.picture(first).delete()
    reopened = Document.open(document.to_bytes(), ids=document.id_state())
    second = reopened.insert_picture(f"{anchor}@0", png(), alt_text="Two").id
    assert second != first
    with pytest.raises(DeletedId, match=f"{first} was deleted"):
        reopened.drawing(first)


def test_a_positional_id_held_from_before_the_first_edit_still_resolves():
    plain = Document.new()
    plain.insert_markdown("One.\n\nTwo.\n\nThree.")
    # A file whose paragraphs have no paraId: the first edit stamps them, keeping aliases.
    data = plain.to_bytes()
    stripped = Document.open(data)
    for root in stripped._roots():
        for node in root.iter(_ids.W_P):
            node.attrib.pop(_ids.PARA_ID, None)
            node.attrib.pop(_ids.TEXT_ID, None)
    for part in stripped._parts():
        stripped.package.mark_dirty(part)
    document = Document.open(stripped.to_bytes())
    held = document.paragraphs()[2].id
    assert held.startswith("p@")
    document.paragraph(document.paragraphs()[0].id).set_text("First.")
    document.delete_block(document.paragraphs()[0].id)
    reopened = Document.open(document.to_bytes(), ids=document.id_state())
    assert reopened.paragraph(held).text == "Three."


# -- refusals ------------------------------------------------------------------------------------


def test_a_state_of_other_bytes_is_refused_with_the_way_out():
    document, _ = _deleted_at_the_end()
    state = document.id_state()
    document.append_paragraph("Later.")
    with pytest.raises(IdStateError) as caught:
        Document.open(document.to_bytes(), ids=state)
    message = str(caught.value)
    assert "of other bytes" in message and "without ids=" in message and "Session.save" in message
    assert isinstance(caught.value, ValueError)


@pytest.mark.parametrize("state", ["nonsense", "docx-ids.1.AAAA", "docx-ids.9.xyz", 42])
def test_a_state_that_is_none_is_refused_with_the_way_out(state):
    document, _ = _deleted_at_the_end()
    with pytest.raises(IdStateError, match="without ids="):
        Document.open(document.to_bytes(), ids=state)


def test_an_empty_state_is_no_state():
    document, gone = _deleted_at_the_end()
    reopened = Document.open(document.to_bytes(), ids="")
    assert reopened.history.version == 0


# -- size ---------------------------------------------------------------------------------------


def test_the_state_of_a_long_day_s_edits_stays_small():
    """A 300-paragraph report rewritten all day: 1,500 text edits, 600 blocks deleted, 600
    added.  The state stays a few kilobytes."""
    document = Document.new()
    document.insert_markdown("\n\n".join(f"Paragraph {k} of the report." for k in range(300)))
    for round_ in range(5):
        for paragraph in document.paragraphs()[:300]:
            document.paragraph(paragraph.id).set_text(f"Paragraph, pass {round_}.")
    for round_ in range(6):
        for paragraph in document.paragraphs()[-100:]:
            document.delete_block(paragraph.id)
        document.insert_markdown("\n\n".join(f"Added {round_}.{k}." for k in range(100)))
    state = document.id_state()
    assert len(state) < 8000, len(state)
    reopened = Document.open(document.to_bytes(), ids=state)
    assert set(reopened._id_pool()) >= set(document._id_pool())


# -- the tools ----------------------------------------------------------------------------------


@pytest.fixture
def box():
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _report() -> bytes:
    document = Document.new()
    document.insert_markdown("# Part one\n\n" + "\n\n".join(f"Text {k}." for k in range(6)))
    return document.to_bytes()


def test_tools_save_and_open_under_the_same_doc_id_keep_who_deleted(box):
    session = box.session()
    doc = session.open(_report(), "report.docx")
    writer = session.view(context={"agent": "writer-2"})
    target = session.entry(doc).document.paragraphs()[-1].id
    deleted = box.dispatch(writer, "word_delete", {"doc": doc, "target": target})
    assert deleted.ok, deleted.to_json()
    saved = session.save(doc)
    assert saved.doc == doc and saved.state and saved.version == deleted.version
    session.close(doc)

    later = box.session()
    again = later.open(saved.data, saved.name, doc_id=saved.doc, state=saved.state)
    assert again == doc
    made = box.dispatch(later, "word_insert_markdown", {"doc": doc, "markdown": "New text.", "at": "end"})
    assert made.ok and target not in made.data["blocks"]
    assert made.version == deleted.version + 1
    result = box.dispatch(later.view(context={"agent": "writer-1"}), "word_set_text",
                          {"doc": doc, "target": target, "text": "Mine."})
    error = result.to_json()["error"]
    assert error["code"] == "not_found"
    assert (f"{target} was deleted by word_delete (agent=writer-2), the change that made version "
            f"{deleted.version}, before the document was saved and opened again") in error["message"]


def test_the_save_tool_s_output_carries_the_state(box):
    session = box.session()
    doc = session.open(_report(), "report.docx")
    target = session.entry(doc).document.paragraphs()[-1].id
    assert box.dispatch(session, "word_delete", {"doc": doc, "target": target}).ok
    saved = box.dispatch(session, "save_document", {"doc": doc, "name": "out.docx", "format": "docx"})
    assert saved.ok, saved.to_json()
    assert "state" not in str(saved.to_json())
    output = session.take_outputs()[-1]
    assert output.state
    later = box.session()
    reopened = later.open(output.data, output.name, doc_id=doc, state=output.state)
    blocks = box.dispatch(later, "word_insert_markdown", {"doc": reopened, "markdown": "Again.", "at": "end"})
    assert target not in blocks.data["blocks"]
    markdown = box.dispatch(session, "save_document", {"doc": doc, "name": "out.md", "format": "markdown"})
    assert markdown.ok and session.take_outputs()[-1].state is None


def test_tools_refuse_a_state_of_other_bytes_and_open_without_one_as_before(box):
    from ooxml_edit.tools import StateRefused

    session = box.session()
    doc = session.open(_report(), "report.docx")
    saved = session.save(doc)
    assert box.dispatch(session, "word_insert_markdown", {"doc": doc, "markdown": "More.", "at": "end"}).ok
    newer = session.save(doc)
    later = box.session()
    with pytest.raises(StateRefused, match="of other bytes"):
        later.open(newer.data, "report.docx", doc_id=doc, state=saved.state)
    assert later.open(newer.data, "report.docx", doc_id=doc) == doc
    later.close(doc)
    assert later.open(newer.data, "report.docx", doc_id=doc, state="") == doc


def test_writers_on_threads_after_a_reopen_never_take_a_retired_id(box):
    session = box.session()
    doc = session.open(_report(), "report.docx")
    handed: list[str] = []
    for k in range(6):
        made = box.dispatch(session, "word_insert_markdown", {"doc": doc, "markdown": f"Early {k}.", "at": "end"})
        handed.extend(made.data["blocks"])
        assert box.dispatch(session, "word_delete", {"doc": doc, "target": made.data["blocks"][0]}).ok
    saved = session.save(doc)
    later = box.session()
    later.open(saved.data, saved.name, doc_id=doc, state=saved.state)
    writers, rounds = 4, 8
    found: list[list[str]] = [[] for _ in range(writers)]
    errors: list[str] = []
    start = threading.Barrier(writers)

    def run(k: int) -> None:
        view = later.view(context={"agent": f"writer-{k}"})
        start.wait()
        try:
            for n in range(rounds):
                made = box.dispatch(view, "word_insert_markdown",
                                    {"doc": doc, "markdown": f"Writer {k} block {n}.", "at": "end"})
                assert made.ok, made.to_json()
                found[k].extend(made.data["blocks"])
                assert box.dispatch(view, "word_delete", {"doc": doc, "target": made.data["blocks"][0]}).ok
        except Exception as error:  # noqa: BLE001 -- any error is the failure
            errors.append(repr(error))

    threads = [threading.Thread(target=run, args=(k,)) for k in range(writers)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    assert errors == []
    every = [identifier for ids in found for identifier in ids]
    assert len(every) == len(set(every)) == writers * rounds
    assert not set(every) & set(handed)
