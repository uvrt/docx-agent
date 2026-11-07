"""Round 37: an id is never given to another block.

A block's id, once it named a block in the document's session, is never generated again for
another block -- in any part, after undo and redo, after a copy or a replacement, from any
thread -- and naming it fails as *deleted*, with the change that removed it, never reaching
another block.  A document saved and opened again goes on after the ids its file still has.
"""

from __future__ import annotations

import struct
import threading
import zlib

import pytest

from docx_agent import Document
from docx_agent.edit import ids as _ids
from docx_agent.edit.ids import DeletedId


def png(width: int = 12, height: int = 8, color=(40, 120, 200)) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(color) * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def every_id(document: Document) -> set[str]:
    """Every block id in every story, now."""
    out: set[str] = set()
    for story in document.stories:
        for paragraph in story.paragraphs:
            out.add(paragraph.id)
        for table in story.tables:
            out.add(table.id)
            out.update(row.id for row in table.rows)
    return out


def hex_of(identifier: str) -> str:
    return identifier.rpartition(":")[2]


# -- the example --------------------------------------------------------------------------------


def test_a_deleted_paragraph_s_id_is_not_given_to_the_next_one():
    document = Document.new()
    document.append_paragraph("a")
    document.append_paragraph("b")
    last = document.paragraphs()[-1].id
    document.delete_block(last)
    made = document.append_paragraph("c").id
    assert made != last
    with pytest.raises(DeletedId) as caught:
        document.paragraph(last)
    message = str(caught.value)
    assert f"{last} was deleted" in message
    assert "version 3" in message and "to_markdown" in message
    assert caught.value.version == 3 and caught.value.identifier == last


def test_every_form_of_a_deleted_id_says_deleted():
    document = Document.new()
    document.insert_markdown("# Heading one\n\nFirst.\n\nSecond.\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\nAfter.")
    heading, first, second, table = (block["id"] for block in document.blocks()[0:4])
    row = document.table(table).rows[0].id
    document.delete_block(second)
    document.delete_block(table)
    for address in (second, f"{second}@0:3", f"{first}..{second}", f"{table}/c0,0", f"{table}/r0", row):
        with pytest.raises(KeyError) as caught:
            document.get(address)
        assert "was deleted" in str(caught.value), address
    with pytest.raises(KeyError, match="was deleted"):
        document.to_markdown(f"{first}..{second}")
    with pytest.raises(KeyError, match="was deleted"):
        document.move_block(second, after=heading)
    with pytest.raises(KeyError, match="was deleted"):
        document.block_of(second)


def test_a_made_up_id_is_still_not_found_not_deleted():
    document = Document.new()
    document.append_paragraph("a")
    document.delete_block(document.paragraphs()[-1].id)
    with pytest.raises(KeyError) as caught:
        document.paragraph("p:0BADF00D")
    assert not isinstance(caught.value, DeletedId)
    assert "no paragraph or table" in str(caught.value)


def test_many_appends_after_deletions_never_take_a_deleted_id():
    document = Document.new()
    document.insert_markdown("\n\n".join(f"Line {k}." for k in range(30)))
    gone = [p.id for p in document.paragraphs()[5:25]]
    for identifier in gone:
        document.delete_block(identifier)
    had = set(gone)
    for k in range(300):
        made = document.append_paragraph(f"New {k}.").id
        assert made not in had
        had.add(made)
    document.insert_markdown("| A | B |\n|---|---|\n" + "\n".join("| x | y |" for _ in range(40)))
    assert not every_id(document) & set(gone)


def test_no_value_any_state_had_is_generated_again():
    """Nor a textId, renewed by an edit of the text: the pool is every value, any kind."""
    document = Document.new()
    document.insert_markdown("One.\n\nTwo.\n\nThree.")
    had: set[int] = set()

    def values() -> set[int]:
        return _ids.used_long_hex([document.package.tree(p) for p in document._parts()])

    for k in range(40):
        had |= values()
        target = document.paragraphs()[k % 3]
        target.set_text(f"Text {k}.")
        new = document.append_paragraph(f"More {k}.").id
        assert int(hex_of(new), 16) not in had
        document.delete_block(new)


# -- undo, redo, aliases ------------------------------------------------------------------------


def test_undo_brings_a_deleted_block_back_with_its_id_and_redo_deletes_it_again():
    document = Document.new()
    document.insert_markdown("One.\n\nTwo.\n\nThree.")
    two = document.paragraphs()[1].id
    document.delete_block(two)
    with pytest.raises(DeletedId):
        document.paragraph(two)
    document.undo()
    assert document.paragraph(two).text == "Two."
    document.redo()
    with pytest.raises(DeletedId, match="version 2"):
        document.paragraph(two)
    after = document.to_bytes()
    document.undo()
    document.redo()
    assert document.to_bytes() == after


def test_an_undone_insertion_s_id_is_not_given_to_the_next_insertion():
    document = Document.new()
    made = document.append_paragraph("First try.").id
    document.undo()
    again = document.append_paragraph("Second try.").id
    assert again != made
    with pytest.raises(DeletedId, match="was deleted"):
        document.paragraph(made)
    document.undo()
    document.redo()
    assert document.paragraph(again).text == "Second try."


def test_an_alias_from_stamping_still_resolves(markup_doc):
    document = Document.open(markup_doc)
    volatile = next(p.id for p in document.paragraphs() if p.volatile)
    result = document.paragraph(volatile).set_text("Stamped.")
    assert volatile in result.renamed
    assert document.paragraph(volatile).text == "Stamped."
    document.delete_block(result.renamed[volatile])
    with pytest.raises(KeyError):
        document.paragraph(volatile)


def test_the_same_edits_write_the_same_bytes():
    def run() -> bytes:
        document = Document.new()
        document.insert_markdown("One.\n\nTwo.\n\nThree.")
        document.delete_block(document.paragraphs()[2].id)
        document.append_paragraph("Four.")
        document.undo()
        document.append_paragraph("Five.")
        return document.to_bytes()

    assert run() == run()


# -- replacements and copies --------------------------------------------------------------------


def test_a_replaced_block_names_the_one_that_replaced_it():
    document = Document.new()
    document.insert_markdown("One.\n\nTwo.\n\nThree.")
    two = document.paragraphs()[1].id
    new = document.insert_markdown("Two again.", at=f"replace:{two}").blocks[0]
    assert new != two
    with pytest.raises(DeletedId) as caught:
        document.paragraph(two)
    assert caught.value.replaced_by == new
    assert f"{two} was replaced with {new}" in str(caught.value)


def test_replacing_two_copies_in_one_batch_gives_neither_an_id_either_had():
    document = Document.new()
    document.insert_markdown("Model.\n\n| Name | Value |\n|---|---|\n| a | 1 |\n| b | 2 |\n\nEnd.")
    model = document.blocks()[1]["id"]
    first = document.copy_blocks(document, [model], at="end").blocks[0]
    second = document.copy_blocks(document, [model], at="end").blocks[0]
    before = every_id(document)
    old = {first, second} | {row.id for row in document.table(first).rows + document.table(second).rows}
    with document.batch():
        one = document.copy_blocks(document, [model], at=f"replace:{first}")
        two = document.copy_blocks(document, [model], at=f"replace:{second}")
    made = set(one.created) | set(two.created)
    assert not made & before
    assert len(made) == len(one.created) + len(two.created)
    for identifier in old:
        with pytest.raises(DeletedId):
            document.get(identifier)
    with pytest.raises(DeletedId) as caught:
        document.table(first)
    assert caught.value.replaced_by == one.blocks[0]


def test_a_copy_from_another_document_does_not_keep_an_id_deleted_here():
    source = Document.new()
    source.insert_markdown("Kept.\n\nOther.")
    document = Document.open(source.to_bytes())
    kept = document.paragraphs()[1].id
    document.delete_block(kept)
    copied = document.copy_blocks(source, [kept], at="end").blocks[0]
    assert copied != kept
    with pytest.raises(DeletedId):
        document.paragraph(kept)
    fresh = Document.new()
    assert fresh.copy_blocks(source, [kept], at="end").blocks[0] == kept   # kept where never used


# -- drawings -----------------------------------------------------------------------------------


def test_a_deleted_drawing_s_number_is_not_given_to_the_next():
    document = Document.new()
    anchor = document.append_paragraph("Picture here.").id
    first = document.insert_picture(f"{anchor}@0", png(), alt_text="One").id
    document.picture(first).delete()
    second = document.insert_picture(f"{anchor}@0", png(), alt_text="Two").id
    assert second != first
    with pytest.raises(DeletedId, match=f"{first} was deleted"):
        document.drawing(first)


# -- a document saved and opened again ----------------------------------------------------------


def test_reopened_the_sequence_goes_on_after_the_last_id_the_file_has():
    document = Document.new()
    document.insert_markdown("\n\n".join(f"Line {k}." for k in range(12)))
    gone = [p.id for p in document.paragraphs()[3:9]]
    for identifier in gone:
        document.delete_block(identifier)
    reopened = Document.open(document.to_bytes())
    made = [reopened.append_paragraph(f"New {k}.").id for k in range(20)]
    assert not set(made) & set(gone)
    assert not set(made) & (every_id(reopened) - set(made))


def test_generate_starts_after_the_last_member_in_use():
    first = _ids.generate("part", set(), 6)
    values = [int(v, 16) for v in first]
    # None of the sequence in use: its first members, as before round 37.
    assert _ids.generate("part", {12345}, 2) == first[:2]
    # The third and fifth in use, the others deleted: after the fifth, never the holes.
    assert _ids.generate("part", {values[2], values[4]}, 1) == first[5:6]
    # A pool remembers where the seed got to, and gives what a plain set would.
    pool = _ids.IdPool({values[2], values[4]})
    assert _ids.generate("part", pool, 1) == first[5:6]
    assert pool.cursors["part"] > 0
    assert _ids.generate("part", pool, 1) == _ids.generate("part", set(values), 1)


# -- the tools ----------------------------------------------------------------------------------


@pytest.fixture
def box():
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _report(blocks: int = 6) -> bytes:
    document = Document.new()
    document.insert_markdown("# Part one\n\n" + "\n\n".join(f"Text {k}." for k in range(blocks))
                             + "\n\n# Part two\n\n" + "\n\n".join(f"More {k}." for k in range(blocks)))
    return document.to_bytes()


def test_a_tool_naming_a_deleted_id_says_which_call_deleted_it(box):
    session = box.session()
    doc = session.open(_report(), "report.docx")
    one = session.view(context={"agent": "writer-1"})
    two = session.view(context={"agent": "writer-2"})
    target = session.entry(doc).document.paragraphs()[3].id
    deleted = box.dispatch(two, "word_delete", {"doc": doc, "target": target})
    assert deleted.ok, deleted.to_json()
    made = box.dispatch(two, "word_insert_markdown", {"doc": doc, "markdown": "Other text.", "at": "end"})
    assert target not in made.data["blocks"]
    for tool, arguments in (("word_set_text", {"doc": doc, "target": target, "text": "Mine."}),
                            ("word_read", {"doc": doc, "range": target}),
                            ("word_inspect", {"doc": doc, "target": target}),
                            ("word_format", {"doc": doc, "targets": [target], "bold": True}),
                            ("word_delete", {"doc": doc, "target": target})):
        result = box.dispatch(one, tool, arguments)
        assert not result.ok, tool
        error = result.to_json()["error"]
        assert error["code"] == "not_found", (tool, error)
        assert f"{target} was deleted by word_delete (agent=writer-2), the change that made version " \
               f"{deleted.version}" in error["message"], (tool, error)
        assert "word_read" in error["message"]


def test_check_render_and_read_batch_say_deleted(box):
    session = box.session()
    doc = session.open(_report(), "report.docx")
    target = session.entry(doc).document.paragraphs()[3].id
    assert box.dispatch(session, "word_delete", {"doc": doc, "target": target}).ok
    checked = box.dispatch(session, "check", {"doc": doc, "range": [target]})
    assert not checked.ok and "was deleted by word_delete" in checked.to_json()["error"]["message"]
    rendered = box.dispatch(session, "render", {"doc": doc, "range": [target], "width": 200})
    assert not rendered.ok and "was deleted by word_delete" in rendered.to_json()["error"]["message"]
    read = box.dispatch(session, "read_batch", {"ops": [
        {"tool": "word_read", "arguments": {"doc": doc, "range": target}}]})
    assert "was deleted by word_delete" in str(read.to_json())
    batch = box.dispatch(session, "batch", {"ops": [
        {"tool": "word_set_text", "arguments": {"doc": doc, "target": target, "text": "x"}}]})
    assert "was deleted by word_delete" in str(batch.to_json())


def test_writers_on_threads_deleting_and_adding_never_share_an_id(box):
    """Several callers at once, each adding blocks and deleting its own: no id is ever handed
    out twice, and each caller's deleted ids say deleted, by it."""
    session = box.session()
    doc = session.open(_report(), "report.docx")
    writers = 4
    rounds = 12
    handed: list[list[str]] = [[] for _ in range(writers)]
    errors: list[str] = []
    start = threading.Barrier(writers)

    def run(k: int) -> None:
        view = session.view(context={"agent": f"writer-{k}"})
        start.wait()
        try:
            for n in range(rounds):
                made = box.dispatch(view, "word_insert_markdown",
                                    {"doc": doc, "markdown": f"Writer {k} block {n}.\n\nSecond {n}.", "at": "end"})
                assert made.ok, made.to_json()
                handed[k].extend(made.data["blocks"])
                gone = made.data["blocks"][0]
                deleted = box.dispatch(view, "word_delete", {"doc": doc, "target": gone})
                assert deleted.ok, deleted.to_json()
                again = box.dispatch(view, "word_set_text", {"doc": doc, "target": gone, "text": "x"})
                assert not again.ok and f"by word_delete (agent=writer-{k})" in again.to_json()["error"]["message"]
        except Exception as error:  # noqa: BLE001 -- any error is the failure
            errors.append(repr(error))

    threads = [threading.Thread(target=run, args=(k,)) for k in range(writers)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    assert errors == []
    every = [identifier for ids in handed for identifier in ids]
    assert len(every) == len(set(every)) == writers * rounds * 2


def test_a_rebuilt_table_of_contents_keeps_its_lines_ids_and_retires_the_lines_it_drops():
    document = Document.new()
    document.insert_markdown("# One\n\nText.\n\n# Two\n\nText.\n\n# Three\n\nText.")
    document.insert_toc(before=document.paragraphs()[0].id)
    lines = [p.id for p in document.paragraphs() if (p.style or "").lower().startswith("toc")]
    before = document.to_bytes()
    document.update_fields()
    assert document.to_bytes() == before                     # nothing changed: the same bytes
    three = next(p.id for p in document.paragraphs() if p.text == "Three")
    document.delete_block(three)
    document.update_fields()
    kept = [p.id for p in document.paragraphs() if (p.style or "").lower().startswith("toc")]
    assert kept == lines[:len(kept)] and len(kept) == len(lines) - 1
    added = document.insert_markdown("# Four", at="end").blocks[0]
    document.update_fields()
    now = [p.id for p in document.paragraphs() if (p.style or "").lower().startswith("toc")]
    assert now[:-1] == kept and now[-1] not in lines and added not in lines
