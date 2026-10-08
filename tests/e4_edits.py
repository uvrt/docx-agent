"""E4's operations, for any document (ROADMAP.md, Phase E4): each edits a document and
returns a check that reads the edit back from a saved and reopened copy.

Sections (a break inserted, page setup, a break removed), headers and footers (created,
edited through the paragraph API, page numbers in them, unlinked and linked again), notes
(inserted, edited, moved, deleted), fields (cross-references, a date, a caption, a page
number, a hyperlink field, updated) and a table of contents.  A document with too little
in it for one gets what it needs written first (``test_e1_operations.text_of``).
"""

from __future__ import annotations

import datetime

from docx_agent import Document, EditError
from docx_agent.edit import text as _text

DATE = datetime.date(2026, 10, 4)
HEADER = "E4 header marker"
FIRST = "E4 first-page header"
EVEN = "E4 even-page footer"
FOOTNOTE = "E4 footnote marker"
ENDNOTE = "E4 endnote marker"


def body_paragraphs(document: Document, minimum: int = 12) -> list:
    """Paragraphs of the body itself (not in a table or control), with text and nothing in
    a field or an object, that end no section."""
    out = []
    for paragraph in document.paragraphs():
        element = paragraph._element
        atoms = _text.atoms(element)
        if (len(atoms) >= minimum and all(a.field is None and a.char != _text.OBJECT for a in atoms)
                and not paragraph.ends_section and element.getparent().tag.endswith("}body")):
            out.append(paragraph)
    return out


def enough(document: Document, count: int = 3) -> list:
    """At least ``count`` plain body paragraphs, written at the body's end where missing."""
    while len(body_paragraphs(document)) < count:
        document.append_paragraph(f"A paragraph written for E4's operations, number {len(body_paragraphs(document)) + 1}.")
    return body_paragraphs(document)


def op_section_break(document: Document):
    paragraphs = enough(document)
    middle = paragraphs[len(paragraphs) // 2]
    count = len(document.sections())
    result = document.insert_section_break(after=middle.id, kind="continuous")
    new = result.id

    def check(saved: Document) -> bool:
        sections = saved.sections()
        ids = [s.id for s in sections]
        return (len(sections) == count + 1 and new in ids
                and sections[ids.index(new) + 1].start == "continuous"
                and saved.paragraph(result.created[1]).ends_section)

    return check


def op_page_setup(document: Document):
    enough(document, 1)
    section = document.sections()[-1].id
    values = dict(orientation="landscape", margin_top=54, margin_left=90, gutter=18, columns=2, column_space=24,
                  column_separator=True, vertical_alignment="center", line_numbering={"count_by": 5},
                  title_page=True, page_number_format="lowerRoman", page_number_start=5,
                  footnote_format="lowerRoman", footnote_restart="eachSect", footnote_position="beneathText")
    result = document.set_section(section, **values)

    def check(saved: Document) -> bool:
        found = saved.section(result.id)
        return (found.orientation == "landscape" and found.page_width > found.page_height
                and found.margins["top"] == 54 and found.margins["left"] == 90 and found.margins["gutter"] == 18
                and found.columns == {"count": 2, "space": 24.0, "separator": True}
                and found.vertical_alignment == "center" and found.line_numbering["count_by"] == 5
                and found.title_page and found.page_numbering == {"format": "lowerRoman", "start": 5}
                and found.notes("footnote") == {"format": "lowerRoman", "start": None, "restart": "eachSect",
                                                "position": "beneathText"})

    return check


def op_remove_break(document: Document):
    if len(document.sections()) < 2:
        paragraphs = enough(document)
        document.insert_section_break(after=paragraphs[0].id, kind="nextPage")
    count = len(document.sections())
    first = document.sections()[0].id
    document.remove_section_break(first)
    return lambda saved: len(saved.sections()) == count - 1 and first not in [s.id for s in saved.sections()]


def _story_text(story) -> str:
    return "\n".join(p.text for p in story.paragraphs) if story is not None else ""


def op_headers(document: Document):
    enough(document, 1)
    section = document.sections()[-1]
    sid = section.id
    if section.header("default") is not None and document._own_reference(section._sectPr, "header", "default") is not None:
        story = section.header("default")
        story.paragraphs[0].set_text(HEADER)
    else:
        document.add_header(sid, "default", "E4 header to edit")
        story = document.section(sid).header("default")
        story.paragraphs[0].set_text(HEADER)
    if document._own_reference(document.section(sid)._sectPr, "footer", "default") is None:
        document.add_footer(sid, "default", "Page ")
    footer = document.section(sid).footer("default")
    paragraph = footer.paragraphs[-1]
    document.insert_page_number(f"{paragraph.id}@{len(paragraph.text)}")
    if document._own_reference(document.section(sid)._sectPr, "header", "first") is None:
        document.add_header(sid, "first", FIRST)
    if document._own_reference(document.section(sid)._sectPr, "footer", "even") is None:
        document.add_footer(sid, "even", EVEN)

    def check(saved: Document) -> bool:
        found = saved.section(sid)
        return (HEADER in _story_text(found.header("default")) and FIRST in _story_text(found.header("first"))
                and EVEN in _story_text(found.footer("even")) and found.title_page and saved.even_and_odd_headers
                and any(f.keyword == "PAGE" for f in saved.fields(found.footer("default").name)))

    return check


def op_link_unlink(document: Document):
    paragraphs = enough(document)
    document.insert_section_break(after=paragraphs[0].id)
    first, second = document.sections()[0].id, document.sections()[1].id
    if document._own_reference(document.section(first)._sectPr, "header", "default") is None:
        document.add_header(first, "default", HEADER)
    unlinked = document.unlink_from_previous(second, "default")
    document.story(unlinked.id).paragraphs[0].set_text("E4 unlinked header")
    copied = document.section(second).header("default").name
    document.link_to_previous(second, "default")

    def check(saved: Document) -> bool:
        ids = [s.id for s in saved.sections()]
        index = ids.index(second) if second in ids else 1
        shown = saved.sections()[index].header("default")
        return shown is not None and shown.name == saved.sections()[index - 1].header("default").name \
            and copied not in [s.name for s in saved.stories]

    return check


def op_notes(document: Document):
    paragraphs = enough(document, 3)
    a, b, c = paragraphs[0], paragraphs[1], paragraphs[-1]
    footnote = document.insert_footnote(f"{a.id}@{len(a.text)}", "E4 footnote to edit")
    document.insert_endnote(f"{b.id}@{len(document.paragraph(b.id).text)}", ENDNOTE)
    document.edit_note(footnote.id, FOOTNOTE + "\nIts second paragraph.")
    gone = document.insert_footnote(f"{b.id}@0", "E4 footnote deleted")
    document.delete_note(gone.id)
    moved = document.move_note(footnote.id, f"{c.id}@{len(document.paragraph(c.id).text)}")
    target = document.paragraph(c.id).id

    def check(saved: Document) -> bool:
        notes = {n.text for n in saved.notes()}
        reference = saved._note_reference(moved.id)
        paragraph = saved._index(reference[0]).entry_for(reference[1].getparent().getparent()) if reference else None
        return (FOOTNOTE + "\nIts second paragraph." in notes and ENDNOTE in notes and "E4 footnote deleted" not in notes
                and paragraph is not None and paragraph.id == target)

    return check


def op_fields(document: Document):
    paragraphs = enough(document, 3)
    a, b, c = paragraphs[0], paragraphs[1], paragraphs[-1]
    document.paragraph(a.id).range(0, 4).add_bookmark("E4Target")
    expected_ref = document.bookmark("E4Target").text
    document.insert_cross_reference(f"{b.id}@{len(document.paragraph(b.id).text)}", "E4Target", kind="text")
    document.insert_cross_reference(f"{b.id}@0", "E4Target", kind="page")
    document.insert_date(f"{c.id}@0", "d MMMM yyyy", date=DATE)
    caption = document.insert_caption(after=c.id, label="Figure", text=": an E4 caption")
    page = document.insert_page_number(f"{document.paragraph(a.id).id}@{len(document.paragraph(a.id).text)}")
    document.insert_hyperlink(f"{c.id}@{len(document.paragraph(c.id).text)}", " E4 link",
                                     "https://example.com/e4")
    document.update_fields(date=DATE)
    label = document.page_label(document.paragraph(a.id).id)
    page_text = document.field(page.id).result

    def check(saved: Document) -> bool:
        results = {f.keyword: f.result for f in saved.fields() if f.keyword in ("REF", "DATE", "SEQ")}
        return (results.get("REF") == expected_ref and results.get("DATE") == "4 October 2026"
                and results.get("SEQ") == "1" and saved.paragraph(caption.id).text == "Figure 1: an E4 caption"
                and any(h.address == "https://example.com/e4" for h in saved.hyperlinks())
                and (label is None or page_text == label))

    return check


def op_toc(document: Document):
    paragraphs = enough(document, 3)
    for k, paragraph in enumerate(paragraphs[:3]):
        document.paragraph(paragraph.id).style = "Heading 1" if k != 1 else "Heading 2"
    headings = [document.paragraph(p.id).id for p in paragraphs[:3]]
    first = next(b for b in document._index(document.package.document_part()).paragraphs
                 if b.element.getparent().tag.endswith("}body"))
    document.insert_toc(before=first.id)
    labels = {h: document.page_label(h) for h in headings}

    def check(saved: Document) -> bool:
        toc = next(f for f in saved.fields() if f.keyword == "TOC")
        lines = toc.result.split("\n")
        pages = [line.rpartition("\t")[2] for line in lines]
        wanted = [labels[h] for h in headings]
        return len(lines) >= 3 and all(w is None or w in pages for w in wanted)

    return check


OPERATIONS = [op_section_break, op_page_setup, op_remove_break, op_headers, op_link_unlink, op_notes, op_fields,
              op_toc]
#: The operations Word tracks (ROADMAP.md, Phase E4): a section's properties, a break, a
#: note, a field, a table of contents, a header's text.
TRACKED = [op_section_break, op_page_setup, op_remove_break, op_notes, op_fields, op_toc]


def e4_edit_set(document: Document, *, track: bool = False) -> dict:
    """The whole set composed, for the render tests and the oracle: a break, page numbers
    in a footer, a first-page header, notes, fields and a table of contents."""
    paragraphs = enough(document, 4)
    out: dict = {}
    for k, paragraph in enumerate(paragraphs[:3]):
        document.paragraph(paragraph.id).style = "Heading 1" if k != 1 else "Heading 2"
    out["headings"] = [document.paragraph(p.id).id for p in paragraphs[:3]]
    out["break"] = document.insert_section_break(after=document.paragraph(paragraphs[1].id).id, kind="nextPage").id
    last = document.sections()[-1].id
    if document._own_reference(document.section(last)._sectPr, "footer", "default") is None:
        document.add_footer(last, "default", "Page ")
    footer = document.section(last).footer("default")
    document.insert_page_number(f"{footer.paragraphs[-1].id}@{len(footer.paragraphs[-1].text)}")
    if document._own_reference(document.section(last)._sectPr, "header", "first") is None:
        document.add_header(last, "first", FIRST)
    c = document.paragraph(paragraphs[2].id)
    out["footnote"] = document.insert_footnote(f"{c.id}@{len(c.text)}", FOOTNOTE).id
    d = document.paragraph(paragraphs[3].id)
    out["endnote"] = document.insert_endnote(f"{d.id}@{len(d.text)}", ENDNOTE).id
    first = next(b for b in document._index(document.package.document_part()).paragraphs
                 if b.element.getparent().tag.endswith("}body"))
    out["toc"] = document.insert_toc(before=first.id).id
    document.update_fields(date=DATE)
    return out


__all__ = ["EditError", "OPERATIONS", "TRACKED", "e4_edit_set"]
