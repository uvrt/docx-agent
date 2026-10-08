"""Tracked changes in detail: what each edit writes, listing and filtering revisions,
accepting and rejecting some of them, the join rule, comments and their parts, Word's
``w:trackRevisions`` setting, table structure -- each undone exactly.
"""

from __future__ import annotations

import datetime as _dt

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.revisions.review import DESCRIBES
from docx_agent.validate import check

from test_roundtrip import entries

AUTHOR = "E3 Agent"
DATE = "2026-10-03T12:00:00Z"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"


@pytest.fixture
def styled(constructs_doc):
    return Document.open(constructs_doc.parent.parent / "lists-and-styles.docx")


@pytest.fixture
def table_doc():
    from conftest import FIXTURE_DIR

    return Document.open(FIXTURE_DIR / "samplelib" / "sample-simple.docx")


def _xml(element) -> str:
    return etree.tostring(element).decode()


def _revenue(document: Document):
    return next(p for p in document.paragraphs() if p.text.startswith("Revenue"))


# -- what an edit writes -----------------------------------------------------------------------


def test_a_replacement_is_a_deletion_then_an_insertion(styled):
    paragraph = _revenue(styled)
    with styled.tracking(author=AUTHOR, date=DATE):
        styled.anchor("percnt", within=paragraph.id).replace("percent")
    xml = _xml(styled.paragraph(paragraph.id)._element)
    assert "<w:delText" in xml and "percnt</w:delText>" in xml
    deletion, insertion = [r for r in styled.revisions(author=AUTHOR)]
    assert (deletion.kind, deletion.text) == ("deletion", "percnt")
    assert (insertion.kind, insertion.text) == ("insertion", "percent")
    assert insertion.date == DATE and insertion.author == AUTHOR
    assert insertion.range.text == "percent" and insertion.range.view == "markup"
    assert insertion.ids == [paragraph.id]
    assert styled.paragraph(paragraph.id).text_in("original").count("percnt") == 1
    # Word writes the true UTC beside the date (w16du:dateUtc) and lists the author.
    assert "dateUtc" in xml
    people = styled.package.tree(styled._comment_part("people"))
    assert [p.get(W15 + "author") for p in people] == [AUTHOR]


def test_deleting_ones_own_insertion_removes_it(styled):
    paragraph = _revenue(styled)
    with styled.tracking(author=AUTHOR, date=DATE):
        styled.insert_text(f"{paragraph.id}@0", "NEWTEXT ")
        styled.range(f"{styled.paragraph(paragraph.id).id}@2:5").delete()
    revisions = styled.revisions(author=AUTHOR)
    assert [(r.kind, r.text) for r in revisions] == [("insertion", "NEXT ")]


def test_deleting_another_authors_insertion_nests_a_deletion(markup_doc):
    document = Document.open(markup_doc)
    theirs = next(r for r in document.revisions() if r.kind == "insertion" and r.author != AUTHOR)
    found = theirs.range
    with document.tracking(author=AUTHOR, date=DATE):
        current = document.paragraph(found.start_id)
        text = theirs.text
        at = current.text.index(text)
        document.range(f"{current.id}@{at}:{at + 1}").delete()
    record = document.revision(theirs.id)._locate()[1]
    assert record.find(W + "del") is not None and record.find(W + "del").get(W + "author") == AUTHOR


def test_formatting_records_the_whole_old_run_properties(styled):
    paragraph = _revenue(styled)
    with styled.tracking(author=AUTHOR, date=DATE):
        styled.anchor("third", within=paragraph.id).format(bold=True)
    change = next(styled.paragraph(paragraph.id)._element.iter(W + "rPrChange"))
    old = change.find(W + "rPr")
    assert old.find(W + "color") is not None and old.find(W + "b") is None
    # Formatting set back drops the record.
    with styled.tracking(author=AUTHOR, date=DATE):
        styled.anchor("third", within=paragraph.id).format(bold=None)
    assert not list(styled.paragraph(paragraph.id)._element.iter(W + "rPrChange"))
    assert styled.revisions(author=AUTHOR) == []


def test_a_style_and_a_list_are_paragraph_property_changes(styled):
    paragraph = _revenue(styled)
    with styled.tracking(author=AUTHOR, date=DATE):
        paragraph.style = "Heading 2"
        styled.paragraph(paragraph.id).add_to_list("number")
    kinds = [r.kind for r in styled.revisions(author=AUTHOR)]
    assert kinds == ["paragraph-properties"]
    change = next(styled.paragraph(paragraph.id)._element.iter(W + "pPrChange"))
    assert change.find(f"{W}pPr/{W}pStyle") is None  # the old properties: none


def test_an_inserted_paragraph_has_its_mark_inserted_and_a_last_one_takes_the_mark_before(styled):
    paragraph = _revenue(styled)
    with styled.tracking(author=AUTHOR, date=DATE):
        middle = styled.insert_paragraph("Inserted in the middle.", after=paragraph.id)
        last = styled.paragraphs()[-1]
        end = styled.insert_paragraph("Inserted at the end.", after=last.id)
    assert styled.paragraph(middle.id)._element.find(f"{W}pPr/{W}rPr/{W}ins") is not None
    # Last in the body: Word cannot accept or reject that mark (measured), so the mark
    # before it is the inserted one, as when Word types a paragraph at the end.
    assert styled.paragraph(end.id)._element.find(f"{W}pPr/{W}rPr/{W}ins") is None
    assert styled.paragraph(last.id)._element.find(f"{W}pPr/{W}rPr/{W}ins") is not None


def test_a_move_is_two_ranges_with_one_name(styled):
    paragraphs = styled.paragraphs()
    mover, target = paragraphs[5], paragraphs[1]
    mover_id = mover.id
    with styled.tracking(author=AUTHOR, date=DATE):
        result = styled.move_block(mover_id, before=target.id)
    root = styled.package.tree(styled.package.document_part())
    names = {n.get(W + "name") for n in root.iter(W + "moveFromRangeStart", W + "moveToRangeStart")}
    assert len(names) == 1 and next(iter(names)).startswith("move")
    # The id the caller held follows the moved paragraph to where it went.
    assert styled.paragraph(mover_id).id == result.id
    assert result.renamed == {mover_id: result.id}
    kinds = sorted(r.kind for r in styled.revisions(author=AUTHOR))
    assert kinds == ["move-from", "move-to", "paragraph-mark-move-from", "paragraph-mark-move-to"]


# -- listing and choosing -----------------------------------------------------------------------


def test_revisions_filter_by_author_kind_place_and_date(markup_doc):
    document = Document.open(markup_doc)
    theirs = document.revisions()
    paragraph = next(p for p in document.paragraphs() if len(p.text) > 20 and p.text_in("markup") == p.text)
    with document.tracking(author=AUTHOR, date=DATE):
        document.insert_text(f"{paragraph.id}@0", "Tracked ")
        created = document.insert_paragraph("A tracked paragraph.", after=paragraph.id)
    mine = document.revisions(author=AUTHOR)
    assert len(document.revisions()) == len(theirs) + len(mine)
    assert {r.kind for r in document.revisions(author=AUTHOR, kind="insertion")} == {
        "insertion", "paragraph-mark-insertion"}
    assert [r.id for r in document.revisions(within=created.id)] == \
        [r.id for r in mine if created.id in r.ids]
    assert document.revisions(since="2026-10-03T11:59:00Z", author=AUTHOR) == mine
    assert document.revisions(until="2026-10-02T00:00:00Z", author=AUTHOR) == []
    within = document.revisions(within=f"{document.paragraph(paragraph.id).id}@0:3")
    assert within and all(r.author == AUTHOR for r in within)


def test_accept_and_reject_one_revision_and_undo(styled):
    paragraph = _revenue(styled)
    with styled.tracking(author=AUTHOR, date=DATE):
        styled.anchor("percnt", within=paragraph.id).replace("percent")
    data = styled.to_bytes()
    deletion, insertion = styled.revisions(author=AUTHOR)
    styled.accept(insertion.id)
    assert [r.kind for r in styled.revisions(author=AUTHOR)] == ["deletion"]
    assert "percnt" in styled.paragraph(paragraph.id).text_in("markup")
    styled.revision(styled.revisions(author=AUTHOR)[0].id).reject()
    assert styled.paragraph(paragraph.id).text.count("percnt") == 1
    assert "percent" in styled.paragraph(paragraph.id).text
    styled.undo()
    styled.undo()
    assert entries(styled.to_bytes()) == entries(data)


def test_a_move_is_accepted_whole(styled):
    paragraphs = styled.paragraphs()
    mover, target = paragraphs[5], paragraphs[1]
    text = mover.text
    with styled.tracking(author=AUTHOR, date=DATE):
        styled.move_block(mover.id, before=target.id)
    source = next(r for r in styled.revisions(author=AUTHOR) if r.kind == "move-from")
    styled.accept(source.id)
    assert styled.revisions(author=AUTHOR) == []
    texts = [p.text for p in styled.paragraphs()]
    assert texts.count(text) == 1 and texts.index(text) == 1
    root = styled.package.tree(styled.package.document_part())
    assert not list(root.iter(W + "moveFromRangeStart", W + "moveToRangeEnd"))


def test_accepting_a_deleted_mark_joins_into_the_next_paragraph(styled):
    """Measured (tools/e3_probe.py, JoinX): accepting a deleted mark keeps the next
    paragraph's element, id and properties; the first paragraph's text comes to its start."""
    paragraphs = styled.paragraphs()
    first, second = paragraphs[5], paragraphs[6]
    first_id = first.id
    with styled.tracking(author=AUTHOR, date=DATE):
        styled.range(f"{first.id}@{len(first.text)}..{second.id}@0").delete()
    # Word, tracking, gives the second the first's properties as a recorded change.
    record = next(r for r in styled.revisions(author=AUTHOR) if r.kind == "paragraph-mark-deletion")
    joined_text = first.text + second.text
    second_id = second.id
    styled.accept(author=AUTHOR)
    assert styled.paragraph(second_id).text == joined_text
    assert styled._direct(first_id) is None  # the first paragraph's element is gone
    assert styled.paragraph(first_id).id == second_id  # its id an alias of the joined paragraph
    assert record


def test_rejecting_ours_leaves_theirs(markup_doc):
    document = Document.open(markup_doc)
    theirs = [(r.id, r.kind, r.author) for r in document.revisions()]
    paragraph = next(p for p in document.paragraphs() if len(p.text) > 12)
    with document.tracking(author=AUTHOR, date=DATE):
        document.paragraph(paragraph.id).set_text("Entirely new text for this paragraph.")
    document.reject(author=AUTHOR)
    assert [(r.id, r.kind, r.author) for r in document.revisions()] == theirs


# -- comments -----------------------------------------------------------------------------------


def test_a_comment_in_the_modern_parts(styled):
    paragraph = _revenue(styled)
    result = styled.add_comment(styled.anchor("fourteen", within=paragraph.id), "Check this figure.",
                                author="Ann Reviewer", initials="AR", date=DATE)
    assert result.id.startswith("c:")
    comment = styled.comment(result.id)
    assert (comment.author, comment.initials, comment.date, comment.text) == (
        "Ann Reviewer", "AR", DATE, "Check this figure.")
    package = styled.package
    main = package.document_part()
    kinds = {r.type.rpartition("/")[2] for r in package.relationships(main).values()}
    assert {"comments", "commentsExtended", "commentsIds", "commentsExtensible", "people"} <= kinds
    assert styled.styles.find("annotation text").id == "CommentText"
    assert styled.styles.find("annotation reference").id == "CommentReference"
    xml = _xml(styled.paragraph(paragraph.id)._element)
    assert xml.index("commentRangeStart") < xml.index(">four<") < xml.index(">teen<") < xml.index("commentRangeEnd")
    reference = next(styled.paragraph(paragraph.id)._element.iter(W + "commentReference")).getparent()
    assert reference.find(f"{W}rPr/{W}sz").get(W + "val") == "22"  # the text's size, as Word writes
    assert not [p for p in check(Document.open(styled.to_bytes()).package) if p.code.startswith("comment")]


def test_replies_resolution_editing_and_deletion_undo_exactly(styled):
    paragraph = _revenue(styled)
    original = styled.to_bytes()
    first = styled.add_comment(styled.paragraph(paragraph.id).range(0, 7), "First.", author="Ann", date=DATE)
    reply = styled.reply_to_comment(first.id, "A reply.", author="Bob", date=DATE)
    assert styled.comment(reply.id).parent == first.id
    assert [c.id for c in styled.comment(first.id).replies] == [reply.id]
    styled.resolve_comment(reply.id)  # resolving a reply resolves its thread
    assert styled.comment(first.id).done
    styled.reopen_comment(first.id)
    assert not styled.comment(first.id).done
    last_para = styled.comment(first.id)._record()["element"].findall(f".//{W}p")[-1].get(
        "{http://schemas.microsoft.com/office/word/2010/wordml}paraId")
    styled.edit_comment(first.id, "First, edited.\nWith a second paragraph.")
    assert styled.comment(first.id).text == "First, edited.\nWith a second paragraph."
    assert styled.comment(first.id)._record()["element"].findall(f".//{W}p")[-1].get(
        "{http://schemas.microsoft.com/office/word/2010/wordml}paraId") == last_para
    edited = styled.to_bytes()
    assert not [p for p in check(Document.open(edited).package) if p.code.startswith("comment")]
    styled.delete_comment(first.id)
    assert styled.comments() == []
    # The last comment gone, the comment parts go, as Word writes a document without one.
    assert styled._comment_part("comments") is None and styled._comment_part("people") is None
    styled.undo()
    assert entries(styled.to_bytes()) == entries(edited)
    while styled.undo():
        pass
    assert entries(styled.to_bytes()) == entries(original)


def test_comments_are_not_revisions(styled):
    paragraph = _revenue(styled)
    with styled.tracking(author=AUTHOR, date=DATE):
        result = styled.add_comment(styled.paragraph(paragraph.id).range(0, 7), "Noted.")
    assert styled.revisions() == []
    assert styled.comment(result.id).author == AUTHOR and styled.comment(result.id).date == DATE


# -- settings ----------------------------------------------------------------------------------


def test_word_tracking_is_a_setting_of_its_own(styled):
    data = styled.to_bytes()
    assert not styled.word_tracks_changes
    styled.word_tracks_changes = True
    root = styled.package.tree(styled.package.settings_part())
    names = [etree.QName(c).localname for c in root]
    assert "trackRevisions" in names and names.index("trackRevisions") < names.index("defaultTabStop")
    assert styled.word_tracks_changes
    with styled.tracking(author=AUTHOR):
        pass
    styled.word_tracks_changes = False
    assert root.find(W + "trackRevisions") is None
    styled.undo()
    styled.undo()
    assert entries(styled.to_bytes()) == entries(data)


def test_a_tracking_date_is_utc_to_the_minute(styled):
    when = _dt.datetime(2026, 10, 3, 14, 30, tzinfo=_dt.timezone(_dt.timedelta(hours=2)))
    with styled.tracking(author=AUTHOR, date=when):
        _revenue(styled).set_text("Revenue grew.")
    assert {r.date for r in styled.revisions(author=AUTHOR)} == {"2026-10-03T12:30:00Z"}
    _revenue(styled).set_text("Revenue grew a lot.", track=True)
    assert {r.author for r in styled.revisions()} == {AUTHOR, "docx-agent"}
    with styled.tracking(author=AUTHOR, date=DATE):
        _revenue(styled).set_text("Untracked inside a tracking context.", track=False)


# -- tables ------------------------------------------------------------------------------------


def _grid_ok(document: Document, table_id: str) -> bool:
    table = document.table(table_id)
    widths = {len(line) for line in table.grid()}
    columns = len(table._entry()[1].element.find(W + "tblGrid").findall(W + "gridCol"))
    return widths == {columns}


def test_rows_and_columns_untracked(table_doc):
    table = table_doc.tables()[0]
    data = table_doc.to_bytes()
    rows, columns = len(table.rows), len(table.grid()[0])
    table_doc.insert_row(table.id, 0)
    table_doc.insert_column(table_doc.tables()[0].id, 1)
    table = table_doc.tables()[0]
    assert (len(table.rows), len(table.grid()[0])) == (rows + 1, columns + 1) and _grid_ok(table_doc, table.id)
    table_doc.delete_column(table.id, 2)
    table_doc.delete_row(table.id, len(table.rows) - 1)
    table = table_doc.tables()[0]
    assert (len(table.rows), len(table.grid()[0])) == (rows, columns) and _grid_ok(table_doc, table.id)
    while table_doc.undo():
        pass
    assert entries(table_doc.to_bytes()) == entries(data)


def test_merges_untracked(table_doc):
    table = table_doc.tables()[0]
    texts = [c.text for c in table.grid()[0][:2]]
    table_doc.merge_cells(table.id, (0, 0), (0, 1))
    table = table_doc.tables()[0]
    grid = table.grid()
    assert grid[0][0] == grid[0][1] and grid[0][0].text == "\n".join(texts)
    table_doc.merge_cells(table.id, (1, 2), (2, 2))
    grid = table_doc.tables()[0].grid()
    assert grid[1][2] == grid[2][2]
    # E5: a rectangle a cell lies partly in grows to take it in whole, and says so.
    result = table_doc.merge_cells(table.id, (0, 1), (1, 1))
    assert result.warnings and "(0, 0) to (1, 1)" in result.warnings[0]
    grid = table_doc.tables()[0].grid()
    assert grid[0][0] == grid[0][1] == grid[1][0] == grid[1][1]
    assert _grid_ok(table_doc, table.id)


def test_a_tracked_column_insert_is_words_form_and_rejects_cleanly(table_doc):
    table = table_doc.tables()[0]
    data = table_doc.to_bytes()
    with table_doc.tracking(author=AUTHOR, date=DATE):
        table_doc.insert_column(table.id, 0)
    table = table_doc.tables()[0]
    root = table._entry()[1].element
    assert not list(root.iter(W + "cellIns"))  # Word writes none: the new cells' marks are inserted
    assert len(table_doc.revisions(author=AUTHOR, kind="paragraph-mark-insertion")) == len(table.rows)
    table_doc.reject(author=AUTHOR)
    from canonical import canonical

    base = Document.open(data)
    base._stamp_document()
    assert canonical(table_doc.to_bytes()) == canonical(base.to_bytes())


def test_a_tracked_merge_is_described_and_the_comment_goes_with_it(table_doc):
    table = table_doc.tables()[0]
    with table_doc.tracking(author=AUTHOR, date=DATE):
        table_doc.merge_cells(table.id, (0, 0), (1, 0))
    described = [c for c in table_doc.comments() if c.text.startswith(DESCRIBES)]
    assert len(described) == 1 and "merged" in described[0].text
    kinds = {r.kind for r in table_doc.revisions(author=AUTHOR)}
    assert {"cell-merge", "deletion"} <= kinds
    table_doc.accept(author=AUTHOR)
    assert table_doc.comments() == [] and table_doc.revisions() == []
    grid = table_doc.tables()[0].grid()
    assert grid[0][0] == grid[1][0]


def test_old_tracking_setting_names_are_deprecated_aliases(styled):
    with pytest.warns(DeprecationWarning, match=r"doc\.tracking"):
        assert styled.track_revisions is False
    with pytest.warns(DeprecationWarning, match="word_tracks_changes"):
        styled.track_revisions = True
    assert styled.word_tracks_changes
    with pytest.warns(DeprecationWarning, match="set_word_tracks_changes"):
        styled.set_track_revisions(False)
    assert not styled.word_tracks_changes
    # Neither name makes this library's own edits tracked: only tracking() does.
    styled.word_tracks_changes = True
    styled.paragraphs()[0].set_text("Untracked all the same.")
    assert styled.revisions() == []
