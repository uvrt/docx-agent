"""The Word tools (``docx_agent.tools``) one by one: the happy path, the error codes with their
valid options, undo, idempotence, tracking mode, refs and batches, and the checks."""

from __future__ import annotations

import datetime as dt
import json
import struct
import zlib
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

FIXTURES = Path(__file__).parent / "fixtures"
AGREEMENT = FIXTURES / "generated" / "pilot" / "agreement-summary.docx"
CHARTS = FIXTURES / "generated" / "charts" / "charts.docx"
SMARTART = FIXTURES / "generated" / "charts" / "smartart.docx"
BRAND = FIXTURES / "generated" / "templates" / "brand.dotx"
CLOCK_AT = dt.datetime(2026, 10, 7, 9, 0, tzinfo=dt.timezone.utc)


def png(width: int = 40, height: int = 20) -> bytes:
    row = b"\x00" + bytes((40, 90, 160)) * width

    def chunk(kind: bytes, body: bytes) -> bytes:
        return (struct.pack(">I", len(body)) + kind + body
                + struct.pack(">I", zlib.crc32(kind + body) & 0xFFFFFFFF))

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(row * height))
            + chunk(b"IEND", b""))


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


@pytest.fixture
def session(box):
    return box.session(clock=lambda: CLOCK_AT)


def open_doc(session, path: Path = AGREEMENT) -> str:
    return session.open(path.read_bytes(), path.name)


def ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, result.error and result.error.to_json())
    return result


def fails(box, session, tool: str, code: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert not result.ok and result.error.code == code, (tool, result.to_json())
    return result.error


def document(session, doc: str = "d1") -> Document:
    return session.entry(doc).document


def _empty_batches_record_nothing() -> bool:
    from ooxml_edit.history import History

    class Value:
        def snapshot(self):
            return 0

        def restore(self, snapshot):
            pass

    history = History(Value())
    with history.batch():
        pass
    return history.version == 0


# -- the format and the definitions --------------------------------------------------------------


def test_every_word_tool_is_docx_or_shared_and_in_a_group():
    names = {t.name for t in TOOLS}
    word = [t for t in TOOLS if t.name.startswith("word_")]
    # W1-W27 and W29 (W28 waits for LE3), less word_describe (the shared describe), and
    # word_review_changes and word_insert_picture (actions of word_changes, word_drawings),
    # and word_insert_table (post-T4: a new table is Markdown)
    assert len(word) == 24
    assert "describe" in names and not names & {"word_describe", "word_review_changes",
                                                "word_insert_picture", "read_chart",
                                                "word_insert_table", "list_documents"}
    assert "word_insert_chart" not in names
    groups = {g.name for g in GROUPS} | {"core"}
    assert all(t.group in groups for t in TOOLS)
    core = sorted(t.name for t in TOOLS if t.group == "core")
    assert {"describe", "word_read", "word_set_text"} <= set(core) and "word_format" not in core


def test_the_format_detects_opens_and_validates_from_bytes():
    data = AGREEMENT.read_bytes()
    assert FORMAT.detect(data, "x.docx") and FORMAT.detect(data, "no-extension")
    assert not FORMAT.detect(b"%PDF-1.7", "x.pdf")
    assert FORMAT.problems(FORMAT.open(data)) == []
    assert "Word documents" in FORMAT.prompt and "house" not in FORMAT.prompt.lower()


# -- reading -------------------------------------------------------------------------------------


def test_describe_lists_headings_sections_comments_and_pages(box, session):
    d = open_doc(session)
    data = ok(box, session, "describe", doc=d).data
    assert data["pages"] == 1 and data["headings"][0] == {
        "id": "p:4C32E1FB", "level": 1, "text": "Supplier agreement summary", "blocks": 11}
    assert data["sections"][0]["id"] == "s:body" and data["comments"]["threads"] >= 1


def test_read_pages_with_a_cursor_and_reads_a_range(box, session):
    d = session.open(Path(FIXTURES / "samplelib" / "sample-long.docx").read_bytes(), "long.docx")
    first = ok(box, session, "word_read", doc=d)
    assert first.next_cursor and first.total > len(first.data)
    second = ok(box, session, "word_read", doc=d, cursor=first.next_cursor)
    assert second.data and second.data != first.data
    one = ok(box, session, "word_read", doc=d, range="p@body/3")
    assert one.data.startswith("<!-- p@body/3") and one.next_cursor is None
    fails(box, session, "word_read", "invalid_arguments", doc=d, cursor="nope")


def test_inspect_gives_effective_formatting(box, session):
    d = open_doc(session)
    data = ok(box, session, "word_inspect", doc=d, target="p:16425160").data
    assert data["blocks"][0]["effective"]["sz"] > 0


# -- text ----------------------------------------------------------------------------------------


def test_set_text_tracked_records_only_the_changed_words(box, session):
    d = open_doc(session)
    ok(box, session, "word_set_tracking", doc=d, on=True, author="Claude")
    paragraph = document(session).paragraph("p:3B212964")
    result = ok(box, session, "word_set_text", doc=d, target="p:3B212964",
                text=paragraph.text.replace("24 months", "36 months"))
    assert result.changed == ["p:3B212964"]
    changes = document(session).changes(author="Claude")
    assert [(c.kind, c.old_text, c.text) for c in changes] == [("replacement", "24", "36")]
    assert result.checks["validate"]["new"] == []


def test_set_text_takes_items_and_refuses_both_forms(box, session):
    d = open_doc(session)
    ok(box, session, "word_set_text", doc=d, items=[{"target": "p:16425160", "text": "Scope of work"},
                                                    {"target": "p:12972045", "text": "Prices"}])
    assert [document(session).paragraph(i).text for i in ("p:16425160", "p:12972045")] == ["Scope of work", "Prices"]
    fails(box, session, "word_set_text", "invalid_arguments", doc=d, target="p:16425160", text="x",
          items=[{"target": "p:16425160", "text": "y"}])


def test_insert_text_returns_the_inserted_range_and_names_it(box, session):
    d = open_doc(session)
    result = ok(box, session, "word_insert_text", doc=d, find="EUR 12,500", text=", excluding VAT", ref="vat")
    rng = result.data["ranges"][0]
    assert document(session).range(rng).text == ", excluding VAT" and result.refs == {"vat": rng}
    ok(box, session, "word_format", doc=d, targets=["$vat"], bold=True)
    assert document(session).range(rng).text == ", excluding VAT"


def test_an_anchor_found_twice_is_ambiguous_with_every_candidate(box, session):
    d = open_doc(session)
    error = fails(box, session, "word_insert_text", "ambiguous", doc=d, find="the", text="x")
    assert len(error.valid_options) > 2 and all("@" in option for option in error.valid_options)
    fails(box, session, "word_insert_text", "not_found", doc=d, find="no such words", text="x")


def test_delete_text_blocks_and_spans(box, session):
    d = open_doc(session)
    ok(box, session, "word_delete", doc=d, find="starting on 1 January 2027", collapse_space=True)
    assert "1 January" not in document(session).paragraph("p:3B212964").text
    before = len(document(session).paragraphs())
    blocks = ok(box, session, "word_insert_markdown", doc=d, markdown="One.\n\nTwo.", at="end").data["blocks"]
    ok(box, session, "word_delete", doc=d, target=f"{blocks[0]}..{blocks[-1]}")
    assert len(document(session).paragraphs()) == before
    error = fails(box, session, "word_delete", "refused", doc=d, target="p:73D22116..p:7C9523C2")
    assert "comment" in error.message


def test_insert_markdown_in_the_documents_styles_with_a_style_map(box, session):
    d = open_doc(session)
    result = ok(box, session, "word_insert_markdown", doc=d, at="after:p:3B212964",
                markdown="## Extra\n\nOne **bold** point.\n\n- a\n- b\n",
                style_map=[{"element": "paragraph", "style": "Normal"}])
    assert len(result.data["blocks"]) == 4
    error = fails(box, session, "word_insert_markdown", "invalid_arguments", doc=d, at="end",
                  markdown="x", style_map=[{"element": "h7", "style": "Normal"}])
    assert "h1" in error.valid_options


def test_insert_markdown_from_a_blob_with_an_image_blob(box, session):
    d = open_doc(session)
    image = session.add_blob(png(), "logo.png")
    md = session.add_blob(f"Logo: ![Company logo]({image})".encode(), "part.md")
    result = ok(box, session, "word_insert_markdown", doc=d, blob=md, at="end")
    assert document(session).pictures()[-1].alt_text == "Company logo" and result.data["blocks"]


def test_format_sets_run_and_paragraph_properties_in_points(box, session):
    d = open_doc(session)
    ok(box, session, "word_format", doc=d, targets=["p:3B212964"], size=12, space_after=9, indent_first=-18,
       alignment="justify")
    paragraph = document(session).paragraph("p:3B212964")
    assert (paragraph.space_after, paragraph.hanging, paragraph.alignment) == (9, 18, "justify")
    ok(box, session, "word_format", doc=d, targets=["p:3B212964@4:12"], italic=True, color="accent1")
    error = fails(box, session, "word_format", "not_found", doc=d, targets=["p:3B212964"],
                  paragraph_style="No Such Style")
    assert "Normal" in error.valid_options
    fails(box, session, "word_format", "invalid_arguments", doc=d, targets=["p:3B212964"])


# -- review --------------------------------------------------------------------------------------


def test_tracking_mode_dates_revisions_by_the_session_clock(box, session):
    d = open_doc(session)
    fails(box, session, "word_set_tracking", "invalid_arguments", doc=d, on=True)
    ok(box, session, "word_set_tracking", doc=d, on=True, author="Claude", word_switch=True)
    ok(box, session, "replace_text", doc=d, find="3%", replace="2.5%", expect="one")
    revisions = document(session).revisions(author="Claude")
    assert revisions and {r.date for r in revisions} == {"2026-10-07T09:00:00Z"}
    assert document(session).word_tracks_changes
    ok(box, session, "word_set_tracking", doc=d, on=False)
    ok(box, session, "replace_text", doc=d, find="99.5%", replace="99.9%", expect="one")
    assert len(document(session).revisions(author="Claude")) == len(revisions)


def test_changes_and_reviewing_by_author_and_by_id(box, session):
    d = open_doc(session)
    ok(box, session, "word_set_tracking", doc=d, on=True, author="Claude")
    ok(box, session, "replace_text", doc=d, find="24 months", replace="36 months", expect="one")
    ok(box, session, "replace_text", doc=d, find="3%", replace="2.5%", expect="one")
    listing = ok(box, session, "word_changes", doc=d, action="list", author="Claude")
    assert listing.checks == {} and listing.version == 2            # list changes nothing
    listed = listing.data
    assert [c["kind"] for c in listed] == ["replacement", "replacement"]
    review = ok(box, session, "word_changes", doc=d, action="reject", ids=[listed[0]["id"]])
    assert review.data["records"] == 2 and review.data["changes_left"] == 1
    ok(box, session, "word_changes", doc=d, action="accept", author="Claude")
    assert document(session).revisions() == [] and "2.5%" in document(session).paragraph("p:4B50F4CD").text
    fails(box, session, "word_changes", "invalid_arguments", doc=d, action="accept")
    fails(box, session, "word_changes", "invalid_arguments", doc=d, action="list", all=True)
    fails(box, session, "word_changes", "invalid_arguments", doc=d, action="accept", all=True,
          detail="records")


def test_comments_add_reply_resolve_in_items_with_refs(box, session):
    d = open_doc(session)
    added = ok(box, session, "word_comments", doc=d, action="add", author="Claude",
               items=[{"find": "99.5%", "text": "Per month?", "ref": "q1"},
                      {"target": "p:3B212964@92:101", "text": "Term?", "ref": "q2"}])
    assert len(added.created) == 2 and set(added.refs) == {"q1", "q2"}
    ok(box, session, "word_comments", doc=d, action="reply", author="Claude",
       items=[{"comment": "$q1", "text": "Monthly.", "resolve": True}])
    threads = {c["id"]: c for c in ok(box, session, "word_comments", doc=d, action="list").data}
    assert threads[added.refs["q1"]]["done"] and threads[added.refs["q1"]]["replies"][0]["text"] == "Monthly."
    assert not threads[added.refs["q2"]]["done"]
    fails(box, session, "word_comments", "not_found", doc=d, action="resolve", items=[{"comment": "$nope"}])


@pytest.mark.skipif(not _empty_batches_record_nothing(), reason="needs ooxml-edit's empty-batch fix (PR #7)")
def test_a_list_action_records_no_undo_step(box, session):
    d = open_doc(session)
    version = session.entry(d).version
    ok(box, session, "word_comments", doc=d, action="list")
    ok(box, session, "word_fields", doc=d, action="list")
    assert session.entry(d).version == version
    fails(box, session, "undo", "refused", doc=d)


# -- structure -----------------------------------------------------------------------------------


def test_move_a_section_after_another_section(box, session):
    d = open_doc(session)
    ok(box, session, "word_move", doc=d, section_heading="p:73D22116", after_section="p:16425160")
    texts = [p.text for p in document(session).paragraphs()]
    assert texts.index("Contacts") == texts.index("Scope") + 2
    fails(box, session, "word_move", "invalid_arguments", doc=d, section_heading="p:73D22116")


def test_copy_from_another_document_maps_styles(box, session):
    d1 = open_doc(session)
    b = session.add_blob(AGREEMENT.read_bytes(), "second.docx")
    d2 = ok(box, session, "open_document", blob=b).created[0]
    result = ok(box, session, "word_copy_from", doc=d1, source_doc=d2, section_heading="p:73D22116",
                at="end", style_map=[{"source": "heading 2", "destination": "Heading 1"}])
    heading = result.data["copied"]["p:73D22116"]
    assert document(session).paragraph(heading).style_name.lower() == "heading 1"
    fails(box, session, "word_copy_from", "invalid_arguments", doc=d1, source_doc=d1, range="p:73D22116", at="end")


def test_sections_break_and_page_setup_in_points(box, session):
    d = open_doc(session)
    broke = ok(box, session, "word_sections", doc=d, action="insert_break", after="p:3B212964", start="nextPage")
    ok(box, session, "word_sections", doc=d, action="set", section="s:body", orientation="landscape",
       margins={"left": 54, "right": 54})
    sections = ok(box, session, "word_sections", doc=d, action="list").data
    assert len(sections) == 2 and sections[1]["orientation"] == "landscape"
    assert sections[1]["margins"]["left"] == 54 and broke.data["section"] == sections[0]["id"]


def test_headers_and_footers_with_page_x_of_y(box, session):
    d = open_doc(session)
    footer = ok(box, session, "word_headers_footers", doc=d, section="s:body", which="footer", action="set",
                page_number="page_x_of_y").data["story"]
    assert document(session).paragraphs(footer)[0].text == "Page 1 of 1"
    # Trial 3: "Confidential" with page_x_of_y gave "Confidential1 of 5".
    ok(box, session, "word_headers_footers", doc=d, section="s:body", which="footer", action="set",
       text="Confidential", page_number="page_x_of_y")
    assert [p.text for p in document(session).paragraphs(footer)] == ["Confidential | Page 1 of 1"]
    ok(box, session, "word_headers_footers", doc=d, section="s:body", which="footer", action="set",
       text="Draft - ", page_number="page_x_of_y")
    assert [p.text for p in document(session).paragraphs(footer)] == ["Draft - Page 1 of 1"]
    ok(box, session, "word_headers_footers", doc=d, section="s:body", which="footer", action="set",
       text="Page ", page_number="after_text")
    assert [p.text for p in document(session).paragraphs(footer)] == ["Page 1"]
    ok(box, session, "word_headers_footers", doc=d, section="s:body", which="footer", action="set", text="Draft")
    assert [p.text for p in document(session).paragraphs(footer)] == ["Draft"]
    ok(box, session, "word_headers_footers", doc=d, section="s:body", which="footer", action="remove")
    # \n makes paragraphs, in a new story as in an existing one (it was refused with code unit).
    header = ok(box, session, "word_headers_footers", doc=d, section="s:body", which="header", action="set",
                text="Acme Ltd\nBoard paper", page_number="page_x_of_y").data["story"]
    assert [p.text for p in document(session).paragraphs(header)] == ["Acme Ltd", "Board paper | Page 1 of 1"]
    ok(box, session, "word_headers_footers", doc=d, section="s:body", which="header", action="set",
       text="One\nTwo\nThree")
    assert [p.text for p in document(session).paragraphs(header)] == ["One", "Two", "Three"]


def test_word_read_refuses_a_range_with_every_story_as_invalid_arguments(box, session):
    d = open_doc(session)
    first = document(session).paragraphs()[0].id
    error = fails(box, session, "word_read", "invalid_arguments", doc=d, range=first, stories="all")
    assert error.field == "range"


def test_fields_toc_update_and_cross_reference(box, session):
    d = open_doc(session)
    ok(box, session, "word_links", doc=d, action="add_bookmark", target="p:12972045", bookmark="Pricing")
    ok(box, session, "word_fields", doc=d, action="insert_cross_reference", find="EUR 12,500",
       bookmark="Pricing", kind="text")
    toc = ok(box, session, "word_fields", doc=d, action="insert_toc", before="p:16425160", levels="1-2")
    assert toc.warnings == []
    listed = ok(box, session, "word_fields", doc=d, action="list").data
    assert {f["keyword"] for f in listed} >= {"TOC", "REF"}
    ok(box, session, "word_fields", doc=d, action="update")
    fails(box, session, "word_fields", "invalid_arguments", doc=d, action="insert_toc", before="p:16425160",
          levels="3-1")


def test_an_empty_toc_is_a_warning_and_a_fact(box, session):
    d = ok(box, session, "new_document", kind="docx").created[0]
    first = document(session, d).paragraphs()[0].id
    toc = ok(box, session, "word_fields", doc=d, action="insert_toc", before=first)
    assert toc.warnings and "no entries" in toc.warnings[0]
    facts = ok(box, session, "check", doc=d, include=["fields"]).data["fields"]
    assert facts["empty_toc"]


def test_notes_and_links(box, session):
    d = open_doc(session)
    note = ok(box, session, "word_notes", doc=d, action="insert", find="99.5%", text="Measured monthly.", ref="n1")
    ok(box, session, "word_notes", doc=d, action="edit", note="$n1", text="Measured each month.")
    assert [n["text"] for n in ok(box, session, "word_notes", doc=d, action="list").data] == ["Measured each month."]
    ok(box, session, "word_links", doc=d, action="add_hyperlink", find="cloud hosting", url="https://example.com/")
    links = ok(box, session, "word_links", doc=d, action="list").data["hyperlinks"]
    assert links[0]["url"] == "https://example.com/"
    ok(box, session, "word_notes", doc=d, action="delete", note=note.created[0])


# -- objects -------------------------------------------------------------------------------------


def test_tables_insert_edit_by_label_and_format(box, session):
    d = open_doc(session)
    # A new table is Markdown (word_insert_table went post-T4: no model called it).
    made = ok(box, session, "word_insert_markdown", doc=d, at="after:p:3B212964",
              markdown="| Item | Fee |\n|---|---|\n| Hosting | 12,500 |\n| Support | 1,000 |\n")
    table = made.created[0]
    assert table.startswith("t:")
    edited = ok(box, session, "word_edit_table", doc=d, table=table, action="set_cells",
                cells=[{"row_label": "Support", "col_label": "Fee", "text": "1,500"}])
    assert "1,500" in edited.data["markdown"]
    ok(box, session, "word_edit_table", doc=d, table=table, action="insert_row")
    ok(box, session, "word_format_table", doc=d, table=table, scope="table", width_percent=80, alignment="center")
    ok(box, session, "word_format_table", doc=d, table=table, scope="row", row=0, shading="D9E2F3", height=20)
    ok(box, session, "word_format_table", doc=d, table=table, scope="column", column=1, width=120)
    assert document(session).table_properties(table)["alignment"] == "center"
    error = fails(box, session, "word_edit_table", "not_found", doc=d, table=table, action="set_cells",
                  cells=[{"row_label": "Nope", "col": 1, "text": "x"}])
    assert "Hosting" in error.valid_options


def test_picture_from_a_blob_floating_and_drawings(box, session):
    d = open_doc(session)
    image = session.add_blob(png(80, 40), "logo.png")
    made = ok(box, session, "word_drawings", doc=d, action="insert_picture", image=image, at="p:3B212964@0",
              width=72, alt_text="Logo", wrap="square", align="right", against="margin", ref="logo")
    assert made.refs == {"logo": made.data["id"]}
    picture = made.data["id"]
    assert made.data["size"] == [72.0, 36.0]
    ok(box, session, "word_drawings", doc=d, action="resize", target=picture, width=144)
    listed = ok(box, session, "word_drawings", doc=d, action="list").data
    assert listed[0]["alt_text"] == "Logo" and not listed[0]["inline"]
    box_made = ok(box, session, "word_drawings", doc=d, action="insert_text_box", find="99.5%", text="Note",
                  width=100, height=40)
    assert box_made.data["id"].startswith("d:")
    text = session.add_blob(b"# not an image", "x.md")
    fails(box, session, "word_drawings", "invalid_arguments", doc=d, action="insert_picture", image=text,
          at="p:3B212964@0", alt_text="x")
    fails(box, session, "word_drawings", "not_found", doc=d, action="insert_picture", image="b99",
          at="p:3B212964@0", alt_text="x")
    fails(box, session, "word_drawings", "invalid_arguments", doc=d, action="insert_picture", image=image,
          at="p:3B212964@0")                                     # a picture needs alt text
    fails(box, session, "word_drawings", "invalid_arguments", doc=d, action="insert_picture", image=image,
          at="p:3B212964@0", alt_text="x", align="left")         # placing it needs wrap


def test_content_controls(box, session):
    d = open_doc(session)
    made = ok(box, session, "word_controls", doc=d, action="insert", find="Scope", type="drop-down",
              items=["Draft", "Final"], title="Status")
    ok(box, session, "word_controls", doc=d, action="fill", control=made.data["id"], value="Final")
    listed = ok(box, session, "word_controls", doc=d, action="list").data
    assert listed[0]["title"] == "Status" and listed[0]["value"] == "Final"


# -- style ---------------------------------------------------------------------------------------


def test_lists_add_restart_and_format(box, session):
    d = open_doc(session)
    ok(box, session, "word_lists", doc=d, targets=["p:3B212964", "p:4B50F4CD"], action="add", kind="number")
    data = ok(box, session, "word_lists", doc=d, targets=["p:4B50F4CD"], action="level", level=1).data
    assert data[0]["list"]["level"] == 1
    ok(box, session, "word_lists", doc=d, targets=["p:3B212964"], action="format", number_format="upperRoman")


def test_styles_add_modify_describe_remove(box, session):
    d = open_doc(session)
    ok(box, session, "word_styles", doc=d, action="add", name="Note Body", based_on="Normal", size=9, italic=True,
       space_after=4)
    described = ok(box, session, "word_styles", doc=d, action="describe", name="Note Body").data
    assert described["declared"]["size"] == 9 and described["declared"]["italic"] is True
    ok(box, session, "word_styles", doc=d, action="modify", name="Note Body", size=10)
    ok(box, session, "word_format", doc=d, targets=["p:3B212964"], paragraph_style="Note Body")
    fails(box, session, "word_styles", "refused", doc=d, action="remove", name="Note Body")
    ok(box, session, "word_styles", doc=d, action="remove", name="Note Body", replacement="Normal")
    error = fails(box, session, "word_styles", "not_found", doc=d, action="describe", name="Nope")
    assert error.valid_options


def test_a_template_is_saved_by_save_document(box, session):
    d = open_doc(session)
    result = ok(box, session, "save_document", doc=d, name="brand-copy.dotx", format="dotx")
    out = session.take_outputs()[-1]
    assert result.data["format"] == "dotx" and out.name == "brand-copy.dotx"
    with pytest.warns(UserWarning, match="template"):
        assert Document.open(out.data).package.kind == "dotx"
    fails(box, session, "word_template", "invalid_arguments", doc=d, action="save_as_template", name="x.dotx")


def test_upgrade_reports_the_reflow(box, session):
    d = session.open((FIXTURES / "generated" / "mode14.docx").read_bytes(), "old.docx")
    data = ok(box, session, "word_template", doc=d, action="upgrade_to_modern").data
    assert data["mode"] == 15 and "changed" in data["reflow"]


# -- the shared tools' Word handlers -------------------------------------------------------------


def test_new_document_from_a_template_blob_and_save_formats(box, session):
    template = session.add_blob(BRAND.read_bytes(), "brand.dotx")
    d = ok(box, session, "new_document", kind="docx", template_blob=template, title="T", author="A",
           name="report.docx").created[0]
    assert document(session, d).properties["title"] == "T"
    ok(box, session, "word_insert_markdown", doc=d, markdown="# Report\n\nText.", at="end")
    ok(box, session, "save_document", doc=d, name="report.docx", format="docx")
    ok(box, session, "save_document", doc=d, name="report.md", format="markdown")
    outputs = session.take_outputs()
    assert Document.open(outputs[0].data).package.kind == "docx"
    assert outputs[1].data.decode().startswith("# Report")
    fails(box, session, "save_document", "invalid_arguments", doc=d, name="x.pptx", format="pptx")
    fails(box, session, "save_document", "invalid_arguments", doc=d, name="x.dotx", format="docx")
    fails(box, session, "new_document", "invalid_arguments", kind="docx", size="16:9")


def test_save_refuses_a_new_validation_problem(box, session, monkeypatch):
    d = open_doc(session)
    from docx_agent.validate import Problem

    monkeypatch.setattr(Document, "validate", lambda self, target=None: [Problem("x", "word/document.xml", "bad")])
    error = fails(box, session, "save_document", "refused", doc=d, name="a.docx", format="docx")
    assert error.valid_options and session.take_outputs() == []


def test_find_and_replace_one_or_all_body_or_all_stories(box, session):
    d = open_doc(session)
    found = ok(box, session, "find_text", doc=d, text="months")
    assert found.total >= 2 and found.data[0]["address"].startswith("p:")
    fails(box, session, "find_text", "invalid_arguments", doc=d, text="x", slides=[1])
    fails(box, session, "replace_text", "ambiguous", doc=d, find="months", replace="weeks", expect="one")
    result = ok(box, session, "replace_text", doc=d, find="months", replace="weeks", expect="all")
    assert result.data["count"] == found.total
    fails(box, session, "replace_text", "not_found", doc=d, find="months", replace="x", expect="all")


def test_render_pages_within_limits_and_cached_by_version(box, session):
    d = open_doc(session)
    first = ok(box, session, "render", doc=d, pages=[1], width=400)
    assert first.images[0].width == 400 and first.images[0].data.startswith(b"\x89PNG")
    again = ok(box, session, "render", doc=d, pages=[1], width=400)
    assert again.images[0].data == first.images[0].data
    error = fails(box, session, "render", "not_found", doc=d, pages=[9])
    assert error.valid_options == [1]
    # slides on a document is read as pages, with a warning (trial 3: 10 of 18 Word runs
    # lost their first render to it); given both, pages is used.
    as_pages = ok(box, session, "render", doc=d, slides=[1], width=400)
    assert as_pages.images[0].data == first.images[0].data
    assert as_pages.warnings == ["slides read as pages: a document has pages"]
    both = ok(box, session, "render", doc=d, slides=[2], pages=[1], width=400)
    assert both.data["pages"] == [1] and both.warnings == ["slides ignored: a document has pages"]


def test_check_reports_validate_reflow_and_fields(box, session):
    d = open_doc(session)
    data = ok(box, session, "check", doc=d).data
    assert data["validate"]["new"] == [] and data["reflow"]["pages"] == 1 and "counts" in data["fields"]
    fails(box, session, "check", "invalid_arguments", doc=d, include=["collisions"])
    assert ok(box, session, "check", doc=d, slides=[1]).warnings == [
        "slides read as pages: a document has pages"]


def test_charts_and_smartart(box, session):
    d = session.open(CHARTS.read_bytes(), "charts.docx")
    data = ok(box, session, "edit_chart", doc=d, target="d:1", action="set_value", series="South", category="Q2",
              value=9.5).data
    assert data["series"][1]["values"][1] == 9.5
    reading = ok(box, session, "edit_chart", doc=d, target="d:1", action="read")
    assert reading.checks == {} and not reading.changed              # read changes nothing
    read = reading.data
    assert read["workbook"]["series"][1]["values"]["values"][1] == 9.5
    error = fails(box, session, "edit_chart", "not_found", doc=d, target="d:1", action="set_values", series="East",
                  values=[1, 2, 3, 4])
    assert "North" in error.valid_options
    s = session.open(SMARTART.read_bytes(), "smartart.docx")
    target = document(session, s).diagrams()[0]
    nodes = ok(box, session, "edit_smartart", doc=s, target=target.address if hasattr(target, "address") else "d:1",
               action="set_text", node=0, text="First").data["nodes"]
    assert nodes[0] == "First"


def test_set_properties(box, session):
    d = open_doc(session)
    data = ok(box, session, "set_properties", doc=d, title="New title", language="en-GB").data
    assert data["title"] == "New title" and data["language"] == "en-GB"


def test_undo_gives_back_the_original_bytes_and_redo_the_edit(box, session):
    d = open_doc(session)
    original = document(session).to_bytes()
    ok(box, session, "word_set_tracking", doc=d, on=True, author="Claude")
    ok(box, session, "word_insert_markdown", doc=d, markdown="More.", at="end")
    edited = document(session).to_bytes()
    ok(box, session, "undo", doc=d)
    assert document(session).to_bytes() == original
    ok(box, session, "undo", doc=d, redo=True)
    assert document(session).to_bytes() == edited
    # A document undoes document-wide: every edit shares its body part, so no scope.
    error = fails(box, session, "undo", "invalid_arguments", doc=d, scope="1")
    assert error.field == "scope" and document(session).to_bytes() == edited


# -- batch and checks ----------------------------------------------------------------------------


def test_a_failing_op_rolls_the_whole_batch_back(box, session):
    d = open_doc(session)
    original = document(session).to_bytes()
    result = box.dispatch(session, "batch", {"ops": [
        {"tool": "word_insert_text", "arguments": {"doc": d, "find": "EUR 12,500", "text": "!", "ref": "x"}},
        {"tool": "word_delete", "arguments": {"doc": d, "find": "no such text"}}]})
    assert not result.ok and result.error.details["op"] == 1
    assert document(session).to_bytes() == original and session.entry(d).refs == {}


def test_batch_refuses_tracking_changes_and_templates(box, session):
    d = open_doc(session)
    for op in ({"tool": "word_set_tracking", "arguments": {"doc": d, "on": True, "author": "C"}},
               {"tool": "word_template", "arguments": {"doc": d, "action": "upgrade_to_modern"}}):
        result = box.dispatch(session, "batch", {"ops": [op]})
        assert not result.ok and result.error.code == "invalid_arguments"


def test_checks_report_the_reflow_once_the_document_was_laid_out(box, session):
    d = open_doc(session)
    blind = ok(box, session, "word_insert_markdown", doc=d, markdown="One.", at="end")
    assert blind.checks["reflow"] is None and blind.checks["validate"]["new"] == []
    ok(box, session, "describe", doc=d)
    seen = ok(box, session, "word_insert_markdown", doc=d, markdown="Two.\n\n" * 3, at="end")
    assert seen.checks["reflow"]["changed"] == [1]


def test_a_long_document_reports_stale_reflow_beyond_twenty_pages(box, session):
    d = session.open((FIXTURES / "samplelib" / "sample-long.docx").read_bytes(), "long.docx")
    ok(box, session, "describe", doc=d)
    edited = ok(box, session, "word_insert_text", doc=d, target="p@body/3", text=" More.")
    assert edited.checks["reflow"] == "stale: call check"


def test_results_fit_the_size_cap(box, session):
    d = session.open((FIXTURES / "samplelib" / "sample-long.docx").read_bytes(), "long.docx")
    for name, arguments in (("word_read", {"doc": d}), ("describe", {"doc": d}),
                            ("find_text", {"doc": d, "text": "e"})):
        result = box.dispatch(session, name, arguments)
        text = json.dumps(result.to_json(), ensure_ascii=False)
        assert result.ok and len(text) <= box.limits.max_result_chars + 500, (name, len(text))


def test_a_creating_call_retried_with_its_key_makes_nothing_new(box, session):
    d = open_doc(session)
    table = "| a | b |\n|---|---|\n| 1 | 2 |\n"
    first = ok(box, session, "word_insert_markdown", doc=d, at="after:p:3B212964", markdown=table, key="t1")
    again = ok(box, session, "word_insert_markdown", doc=d, at="after:p:3B212964", markdown=table, key="t1")
    assert again.created == first.created and "already done" in again.summary
    assert len(document(session).tables()) == 1


@pytest.mark.skipif(not _empty_batches_record_nothing(), reason="needs ooxml-edit's empty-batch fix (PR #7)")
def test_setting_the_same_text_again_changes_nothing_and_records_no_step(box, session):
    d = open_doc(session)
    text = document(session).paragraph("p:16425160").text
    version = session.entry(d).version
    ok(box, session, "word_set_text", doc=d, target="p:16425160", text=text)
    assert session.entry(d).version == version


def test_a_picture_in_a_paragraph_of_its_own_centred(box, session):
    d = open_doc(session)
    image = session.add_blob(png(60, 40), "chart.png")
    made = ok(box, session, "word_drawings", doc=d, action="insert_picture", image=image,
              at="after:p:3B212964", width=144, align="center", alt_text="A chart")
    paragraph = made.data["paragraph"]
    document_ = document(session)
    ids = [p.id for p in document_.paragraphs()]
    assert ids[ids.index("p:3B212964") + 1] == paragraph
    assert document_.paragraph(paragraph).text == "\ufffc" and made.data["size"] == [144.0, 96.0]
    inspected = ok(box, session, "word_inspect", doc=d, target=paragraph).data
    assert inspected["blocks"][0]["effective"]["alignment"] == "center"
    ok(box, session, "undo", doc=d)
    assert paragraph not in [p.id for p in document(session).paragraphs()]


# -- coverage: "the check passed" against "the check could not see everything" ---------------


def _vml_first() -> bytes:
    """A document whose first paragraph holds a VML shape, which docx2svg does not lay out:
    its layout stops there, on any machine, whatever faces it has."""
    import io
    import zipfile

    w = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
         'xmlns:v="urn:schemas-microsoft-com:vml"')
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    body = ('<w:p><w:r><w:pict><v:rect style="position:absolute;width:100pt;height:20pt" fillcolor="#4472c4"/>'
            '</w:pict></w:r><w:r><w:t>A shape.</w:t></w:r></w:p><w:p><w:r><w:t>After it.</w:t></w:r></w:p>')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" '
            'ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'))
        archive.writestr("_rels/.rels", (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship '
            f'Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>'))
        archive.writestr("word/document.xml", f'<w:document {w}><w:body>{body}<w:sectPr/></w:body></w:document>')
    return buffer.getvalue()


def test_check_render_and_save_say_how_much_was_laid_out(box, session):
    d = open_doc(session)
    reflow = ok(box, session, "check", doc=d, include=["reflow"]).data["reflow"]
    coverage = reflow["coverage"]
    assert isinstance(coverage["complete"], bool) and coverage["pages"] == reflow["pages"]
    # A complete coverage has no stop; a body stop is the reflow's stop too.
    assert ("stop" in coverage) == ("stopped" in reflow)
    assert coverage["complete"] is False or ("stop" not in coverage and "missing_fonts" not in coverage)
    rendered = ok(box, session, "render", doc=d, pages=[1], width=400)
    assert rendered.data["coverage"] == coverage
    saved = ok(box, session, "save_document", doc=d, name="agreement.docx", format="docx")
    assert saved.data["coverage"] == coverage


def test_a_layout_that_stops_is_not_complete(box, session):
    d = session.open(_vml_first(), "shape.docx")
    reflow = ok(box, session, "check", doc=d, include=["reflow"]).data["reflow"]
    coverage = reflow["coverage"]
    assert coverage["complete"] is False
    assert coverage["stop"]["reason"] == "drawing" and coverage["stop"]["page"] == 1
    assert coverage["stop"]["at"] == reflow["stopped"]["at"] and coverage["stop"]["at"].startswith("p")
    assert coverage["blocks_laid_out"] == [0, 2]
    assert ok(box, session, "save_document", doc=d, name="shape.docx", format="docx").data["coverage"] == coverage
