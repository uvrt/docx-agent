"""What E4's edits write, held to what Word wrote for the same edits
(``tools/e4_probe.py``, ``tests/observations/e4-word.json``): each probe document is
rebuilt here, edited through the API as the probe edited it in Word, and compared -- part
by part where the forms are Word's, and the table of contents' page numbers to Word's own.
"""

from __future__ import annotations

import datetime
import json
import re
import sys
from pathlib import Path

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.validate import check

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import e4_probe  # noqa: E402

OBSERVED = json.loads((ROOT / "tests" / "observations" / "e4-word.json").read_text(encoding="utf-8"))
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
#: Dutch Word's ids for the built-in styles the probes used, and English Word's.
STYLE_IDS = {"Koptekst": "Header", "Voettekst": "Footer", "Voetnoottekst": "FootnoteText",
             "Voetnootmarkering": "FootnoteReference", "Eindnoottekst": "EndnoteText",
             "Eindnootmarkering": "EndnoteReference", "Bijschrift": "Caption", "Kop1": "Heading1",
             "Kop2": "Heading2", "Kop3": "Heading3", **{f"Inhopg{n}": f"TOC{n}" for n in range(1, 10)}}


def plain(xml: str) -> str:
    """A part or block as compared: no namespace declarations, ids, rsids or relationship
    ids; style ids in English; bookmarks' numbers aside."""
    xml = re.sub(r' xmlns(:\w+)?="[^"]*"', "", xml)
    xml = re.sub(r' (w14:paraId|w14:textId|w:rsid\w*|r:id|mc:Ignorable)="[^"]*"', "", xml)
    xml = re.sub(r'_Toc\d+', "_Toc", xml)
    xml = re.sub(r'(bookmark(?:Start|End)) w:id="\d+"', r"\1", xml)
    for dutch, english in STYLE_IDS.items():
        xml = xml.replace(f'w:val="{dutch}"', f'w:val="{english}"')
    return xml


def xml(element) -> str:
    return plain(etree.tostring(element, encoding="unicode"))


def observed_section(probe: str, index: int) -> str:
    """The ``index``-th ``w:sectPr`` of a probe's body as Word wrote it, its references
    aside (Word's AppleScript made empty headers and footers on asking for one)."""
    found = re.findall(r"<w:sectPr>.*?</w:sectPr>(?!</w:sectPrChange)", "".join(OBSERVED[probe]["blocks"]))
    text = found[index]
    return plain(re.sub(r"<w:(header|footer)Reference [^>]*/>", "", text))


# -- sections ----------------------------------------------------------------------------------


def test_page_setup_is_written_as_word_writes_it():
    document = Document.open(e4_probe.sectioned_document())
    second = document.sections()[1].id
    document.set_section(second, orientation="landscape", margin_top=54, margin_left=90, gutter=18, columns=2,
                         column_space=24, column_separator=True, vertical_alignment="center",
                         line_numbering={"count_by": 5}, title_page=True, page_number_format="lowerRoman",
                         page_number_start=5)
    ours = xml(document.section(second)._sectPr).replace('<w:type w:val="nextPage"/>', "")
    assert ours == observed_section("setup", 1)


def test_a_break_before_a_heading_is_a_paragraph_of_its_own_with_the_headings_properties():
    document = Document.open(e4_probe.headings_document())
    for k, level in e4_probe.heading_levels().items():
        document.paragraph(document.paragraphs()[k - 1].id).style = f"heading {level}"
    chapter2 = [k for k, level in e4_probe.heading_levels().items() if level == 1][1]
    target = document.paragraphs()[chapter2 - 1]
    result = document.insert_section_break(before=target.id, kind="continuous")
    ours = xml(document.paragraph(result.created[1])._element)
    word = next(plain(b) for b in OBSERVED["fields"]["blocks"] if "<w:sectPr>" in b and "<w:p " in b)
    assert ours == word
    assert document.sections()[-1].start == "continuous"


def test_section_breaks_of_each_kind_split_the_section_as_word_does():
    document = Document.open(e4_probe.plain_document())
    paragraphs = document.paragraphs()
    document.insert_section_break(after=paragraphs[8].id, kind="evenPage")
    document.insert_section_break(before=paragraphs[5].id, kind="continuous")
    document.insert_section_break(after=paragraphs[2].id, kind="nextPage")
    starts = [s.start for s in document.sections()]
    # Word: the break paragraphs' sections keep what they split, the last kinds are the breaks'.
    assert starts == ["nextPage", "nextPage", "continuous", "evenPage"]
    word = [plain(s) for s in re.findall(r"<w:sectPr>.*?</w:sectPr>", "".join(OBSERVED["breaks"]["blocks"]))]
    ours = [xml(s._sectPr) for s in document.sections()]
    assert ours == word
    assert set(check(Document.open(document.to_bytes()).package)) == set()


def test_a_break_moves_the_split_sections_headers_to_the_new_section():
    """Word gives the new section (before the break) the references; the one split inherits
    them (``tracked``: the break paragraph's sectPr holds all six)."""
    document = Document.open(e4_probe.sectioned_document())
    document.add_header("s:body", "default", "Body header")
    document.insert_section_break(after=document.paragraphs()[9].id)
    sections = document.sections()
    assert document._own_reference(sections[-2]._sectPr, "header", "default") is not None
    assert document._own_reference(sections[-1]._sectPr, "header", "default") is None
    assert sections[-1].header("default").paragraphs[0].text == "Body header"


def test_removing_a_break_keeps_the_next_sections_properties_and_drops_the_headers_it_overrides():
    """``removal``: the break between sections 1 and 2 deleted -- the text joins section 2's
    properties (its left margin, its header); section 1's header part goes."""
    document = Document.open(e4_probe.sectioned_document())
    first, second = document.sections()[0].id, document.sections()[1].id
    document.set_section(first, margin_left=36)
    document.set_section(second, margin_left=108)
    document.add_header(first, "default", "Header one")
    document.add_header(second, "default", "Header two")
    headers = len([p for p in document.package.part_names if "header" in p])
    paragraph = document.paragraph(document.section(first).paragraph_ids[-1])
    document.remove_section_break(first)
    sections = document.sections()
    assert len(sections) == 2 and sections[0].margins["left"] == 108
    assert sections[0].header("default").paragraphs[0].text == "Header two"
    assert len([p for p in document.package.part_names if "header" in p]) == headers - 1
    joined = [p.text for p in document.paragraphs()]
    word = [re.sub(r"<[^>]+>", "", b) for b in OBSERVED["removal"]["blocks"][:-1]]
    assert [t[:24] for t in joined] == [t[:24] for t in word]
    del paragraph


def test_a_tracked_section_change_records_what_changed():
    document = Document.open(e4_probe.headings_document())
    with document.tracking(author="A", date="2026-10-04T12:00:00Z"):
        document.set_section("s:body", margin_top=54, columns=2)
    record = document.section("s:body")._sectPr.find(_W + "sectPrChange")
    old = xml(record.find(_W + "sectPr"))
    word = re.search(r"<w:sectPrChange [^>]*>(<w:sectPr>.*?</w:sectPr>)", "".join(OBSERVED["tracked"]["blocks"]))
    word_old = plain(word.group(1))
    # Word records the margins whole, as here; of the columns only the number (here the old
    # element whole, so rejecting restores its spacing too).
    assert re.search(r"<w:pgMar [^>]*/>", old).group(0) == re.search(r"<w:pgMar [^>]*/>", word_old).group(0)
    assert old.count("<w:cols") == 1 and "<w:cols" in word_old
    document.reject_all()
    assert document.section("s:body").margins["top"] == 72 and document.section("s:body").columns["count"] == 1


def test_a_tracked_break_is_an_inserted_mark_holding_the_section():
    document = Document.open(e4_probe.headings_document())
    target = document.paragraphs()[18]
    with document.tracking(author="A", date="2026-10-04T12:00:00Z"):
        result = document.insert_section_break(before=target.id)
    paragraph = document.paragraph(result.created[1])._element
    assert paragraph.find(f"{_W}pPr/{_W}rPr/{_W}ins") is not None
    assert paragraph.find(f"{_W}pPr/{_W}sectPr") is not None
    document.reject_all()
    assert len(document.sections()) == 1


def test_section_values_are_checked_before_anything_changes():
    document = Document.open(e4_probe.sectioned_document())
    for values in ({"orientation": "sideways"}, {"start": "nextChapter"}, {"columns": 0}, {"colour": "red"},
                   {"vertical_alignment": "middle"}, {"footnote_restart": "never"}, {"page_width": -1}):
        with pytest.raises(EditError):
            document.set_section("s:body", **values)
    with pytest.raises(EditError):
        document.remove_section_break("s:body")
    with pytest.raises(EditError):
        document.insert_section_break(after=document.paragraphs()[0].id, kind="nextChapter")
    assert document.package.dirty_parts == frozenset()


# -- headers and footers --------------------------------------------------------------------


def test_a_header_part_is_word_s_and_its_references_are_in_word_s_order():
    document = Document.open(e4_probe.sectioned_document())
    first = document.sections()[0].id
    document.add_header(first, "default", "Default header one")
    document.add_header(first, "first", "First header one")
    document.add_header(first, "even", "Even header one")
    for kind in ("even", "default", "first"):
        document.add_footer(first, kind, "")
    sect = document.section(first)._sectPr
    order = [(n.tag.rpartition("}")[2], n.get(_W + "type")) for n in sect
             if n.tag.endswith("Reference")]
    word = re.findall(r"<w:(header|footer)Reference w:type=\"(\w+)\"", OBSERVED["headers"]["blocks"][3])
    assert order == [(f"{which}Reference", kind) for which, kind in word]
    story = document.section(first).header("default")
    root = document.package.tree(story.part)
    observed = next(v for k, v in OBSERVED["headers"].items() if k.startswith("word/header") and "Default header one" in v)
    assert plain(etree.tostring(root, encoding="unicode")) == plain(re.sub(r"<\?xml[^>]*\?>\s*", "", observed))
    assert document.section(first).title_page and document.even_and_odd_headers


def test_unlinking_copies_the_story_before_and_linking_again_removes_it():
    """``headers2``: an unlinked footer starts as a copy of the one before; a header linked
    again loses its part."""
    document = Document.open(e4_probe.sectioned_document())
    sections = [s.id for s in document.sections()]
    document.add_footer(sections[0], "default", "Footer one")
    result = document.unlink_from_previous(sections[1], "default", footer=True)
    assert document.story(result.id).paragraphs[0].text == "Footer one"
    assert document.section(sections[1]).footer("default").name == result.id
    unlinked = document.unlink_from_previous(sections[2], "default")
    assert unlinked.id.startswith("header")  # no header before it: an empty one
    parts = set(document.package.part_names)
    document.link_to_previous(sections[2], "default")
    assert set(document.package.part_names) == parts - {f"word/{unlinked.id}.xml"}
    assert document.section(sections[2]).header("default") is None


def test_page_number_fields_are_word_s_forms():
    document = Document.open(e4_probe.sectioned_document())
    first = document.sections()[0].id
    document.add_footer(first, "default", "")
    document.add_footer(first, "first", "")
    document.add_footer(first, "even", "")
    for kind, which in (("PAGE", "default"), ("NUMPAGES", "first"), ("SECTIONPAGES", "even")):
        story = document.section(first).footer(which)
        document.insert_page_number(f"{story.paragraphs[0].id}@0", kind)
    texts = {}
    for which in ("default", "first", "even"):
        root = document.package.tree(document.section(first).footer(which).part)
        texts[which] = plain(etree.tostring(root, encoding="unicode"))
    observed = {k: plain(re.sub(r"<\?xml[^>]*\?>\s*", "", v)) for k, v in OBSERVED["headers2"].items()
                if k.startswith("word/footer")}
    # The cached number is where each was computed (ours: the first page showing the story;
    # Word's: where it was put in), the form the same.
    digits = lambda text: re.sub(r"<w:t>\d+</w:t>", "<w:t>n</w:t>", text)  # noqa: E731
    for which in ("default", "first", "even"):
        assert digits(texts[which]) in {digits(v) for v in observed.values()}, texts[which]


# -- notes ------------------------------------------------------------------------------------


def test_a_note_is_written_as_word_writes_one():
    document = Document.open(e4_probe.sectioned_document())
    paragraph = document.paragraphs()[1]
    note = document.insert_footnote(f"{paragraph.id}@{len(paragraph.text)}", "No Reference")
    ours = plain(etree.tostring(document._note_element(note.id)[1], encoding="unicode"))
    word = re.search(r'<w:footnote w:id="1">.*?</w:footnote>', OBSERVED["notes"]["word/footnotes.xml"]).group(0)
    # Word gives the note's text and mark the Dutch interface's language; that aside, alike.
    word = re.sub(r"<w:rPr><w:lang w:val=\"nl-NL\"/></w:rPr>", "", plain(word))
    assert ours == word
    reference = plain(etree.tostring(document._note_reference(note.id)[1].getparent(), encoding="unicode"))
    assert reference == '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'
    names = set(document.package.part_names)
    assert {"word/footnotes.xml", "word/endnotes.xml"} <= names


def test_note_numbering_is_the_sections():
    document = Document.open(e4_probe.sectioned_document())
    second = document.sections()[1].id
    document.set_section(second, footnote_format="lowerRoman", footnote_restart="eachSect",
                         footnote_position="beneathText")
    document.set_section("s:body", footnote_format="upperLetter", footnote_start=3)
    word = re.findall(r"<w:footnotePr>.*?</w:footnotePr>", "".join(OBSERVED["notes"]["blocks"]))[0]
    assert xml(document.section(second)._sectPr.find(_W + "footnotePr")) == plain(word)
    word = re.findall(r"<w:footnotePr>.*?</w:footnotePr>", "".join(OBSERVED["notes2"]["blocks"]))[0]
    assert xml(document.section("s:body")._sectPr.find(_W + "footnotePr")) == plain(word)


def test_notes_are_inserted_edited_moved_and_deleted():
    document = Document.open(e4_probe.sectioned_document())
    a, b, c = document.paragraphs()[1], document.paragraphs()[5], document.paragraphs()[9]
    one = document.insert_footnote(f"{a.id}@{len(a.text)}", "One")
    two = document.insert_footnote(f"{b.id}@{len(b.text)}", "Two")
    end = document.insert_endnote(f"{c.id}@3", "End")
    document.edit_note(one.id, "One, edited\nand a second paragraph")
    assert document.note(one.id).text == "One, edited\nand a second paragraph"
    document.move_note(two.id, f"{c.id}@0")
    assert document._note_reference(two.id)[1].getparent().getparent() is document.paragraph(c.id)._element
    document.delete_note(end.id)
    assert [n.id for n in document.notes()] == [one.id, two.id]
    assert set(check(Document.open(document.to_bytes()).package)) == set()
    with pytest.raises(EditError):
        document.set_note_settings("footnote", position="beneathText")
    document.set_note_settings("endnote", position="sectEnd")
    settings = document.package.tree(document.package.settings_part())
    assert settings.find(f"{_W}endnotePr/{_W}pos").get(_W + "val") == "sectEnd"


def test_a_tracked_note_deletion_keeps_the_note_until_accepted():
    document = Document.open(e4_probe.sectioned_document())
    a = document.paragraphs()[1]
    note = document.insert_footnote(f"{a.id}@{len(a.text)}", "Kept until accepted")
    with document.tracking(author="A", date="2026-10-04T12:00:00Z"):
        document.delete_note(note.id)
    assert document.note(note.id).text == "Kept until accepted"
    document.accept_all()
    assert document.notes() == []


# -- fields -----------------------------------------------------------------------------------


def test_field_forms_are_word_s():
    document = Document.open(e4_probe.headings_document())
    for k, level in e4_probe.heading_levels().items():
        document.paragraph(document.paragraphs()[k - 1].id).style = f"heading {level}"
    paragraphs = document.paragraphs()
    last = len(paragraphs)
    date = document.insert_date(f"{paragraphs[last - 1].id}@0", "d MMMM yyyy", date=datetime.date(2026, 10, 4))
    caption = document.insert_caption(after=paragraphs[last - 3].id, text=": a figure")
    observed = "".join(OBSERVED["fields"]["blocks"])
    ours = xml(document.paragraph(paragraphs[last - 1].id)._element)
    word_date = re.search(r'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE.*?'
                          r'w:fldCharType="end"/></w:r>', observed).group(0)
    assert word_date in ours
    word_caption = re.search(r'<w:r><w:t xml:space="preserve">Figure </w:t></w:r><w:fldSimple.*?</w:fldSimple>'
                             r'<w:r><w:t>: a figure</w:t></w:r>', observed).group(0)
    caption_xml = xml(document.paragraph(caption.id)._element)
    assert word_caption in caption_xml
    assert '<w:pStyle w:val="Caption"/>' in caption_xml
    # Field ids are positional, as run ids are: the caption before it renumbered the date.
    assert next(f for f in document.fields() if f.keyword == "DATE").result == "4 October 2026"
    del date


def test_a_hyperlink_field_is_written_as_word_saves_one():
    document = Document.open(e4_probe.plain_document())
    paragraph = document.paragraphs()[0]
    document.insert_field(f"{paragraph.id}@{len(paragraph.text)}", 'HYPERLINK "https://example.com/e4"')
    link = document.paragraphs()[0].hyperlinks[0]
    assert link.address == "https://example.com/e4" and link.text == "https://example.com/e4"
    assert '<w:rStyle w:val="Hyperlink"/>' in xml(document.paragraphs()[0]._element)


def test_sequence_numbers_count_in_order_and_restart():
    document = Document.open(e4_probe.plain_document())
    paragraphs = document.paragraphs()
    first = document.insert_caption(after=paragraphs[3].id, label="Table", text=": first")
    second = document.insert_caption(after=paragraphs[1].id, label="Table", text=": second, earlier")
    document.update_fields()
    assert document.paragraph(second.id).text == "Table 1: second, earlier"
    assert document.paragraph(first.id).text == "Table 2: first"
    restart = document.insert_field(f"{paragraphs[6].id}@0", " SEQ Table \\r 7 ")
    document.update_fields()
    assert document.field(restart.id).result == "7"


def test_a_page_reference_shows_the_number_word_shows():
    """A section restarting at 5 in lower Roman: a PAGEREF to its text is ``v``, not the
    page's place."""
    document = Document.open(e4_probe.sectioned_document())
    second = document.sections()[1]
    document.set_section(second.id, page_number_format="lowerRoman", page_number_start=5)
    target = document.paragraph(second.paragraph_ids[1])
    target.range(0, 7).add_bookmark("InSecond")
    document.insert_cross_reference(f"{document.paragraphs()[0].id}@0", "InSecond", kind="page")
    field = next(f for f in document.fields() if f.keyword == "PAGEREF")
    assert field.result == "v"


def test_a_field_past_where_the_layout_stops_is_unknown_and_left():
    """Its result is left and the field listed as unknown -- never marked dirty, which makes
    Word ask on opening (``test_oracle_e4``)."""
    from docx_agent.layout import Unknown

    document = Document.open(e4_probe.plain_document())
    paragraph = document.paragraphs()[-1]
    paragraph.range(0, 9).add_bookmark("Far")
    original = document.layout

    class Stopped:
        stopped = True
        pages = []

        def where(self, identifier):
            return Unknown("drawing", 1)

    document.layout = lambda **options: Stopped()  # type: ignore[method-assign]
    result = document.insert_field(f"{document.paragraphs()[0].id}@0", " PAGEREF Far \\h ")
    assert result.unknown == [result.id]
    assert not document.field(result.id).dirty and document.field(result.id).result == ""
    reference = document.insert_cross_reference(f"{document.paragraphs()[1].id}@0", "Far", kind="page")
    assert reference.unknown == ["PAGEREF Far"]
    assert "dirty" not in etree.tostring(document.package.tree("word/document.xml")).decode()
    document.layout = original  # type: ignore[method-assign]
    updated = document.update_fields()
    assert updated.unknown == [] and document.field(result.id).result == "1"


# -- the table of contents --------------------------------------------------------------------


def _toc_document(margins: dict) -> Document:
    document = Document.open(e4_probe.headings_document())
    for k, level in e4_probe.heading_levels().items():
        document.paragraph(document.paragraphs()[k - 1].id).style = f"heading {level}"
    if margins:
        document.set_section("s:body", **margins)
    return document


def _entries(blocks: list[str]) -> list[tuple[str, ...]]:
    return [tuple(re.findall(r"<w:t>([^<]*)</w:t>", b)) for b in blocks if "Inhopg" in b or 'w:val="TOC' in b]


def test_the_table_of_contents_is_word_s_and_its_page_numbers_are_word_s():
    """``tocfield``: Word's own table of contents of the probe, updated by Word -- every
    entry, every page number."""
    document = _toc_document({"margin_left": 90, "margin_right": 36})
    result = document.insert_toc(before=document.paragraphs()[0].id)
    assert result.unknown == []
    ours = [xml(document.paragraph(i)._element) for i in result.created[1:]]
    word = [plain(b) for b in OBSERVED["tocfield"]["blocks"] if "Inhopg" in b]
    word = [w.replace(" \\* MERGEFORMAT", "") for w in word]
    assert ours == word
    # The field's end opens the paragraph it was put before, as Word's does.
    after = xml(document.paragraphs()[len(ours)]._element).replace(' xml:space="preserve"', "")
    assert after == plain(OBSERVED["tocfield"]["blocks"][len(word)])


def test_the_toc_styles_are_word_s():
    document = Document.open(e4_probe.nine_document())
    for level in range(1, 10):
        document.paragraph(document.paragraphs()[2 * level - 1].id).style = f"heading {level}"
    result = document.insert_toc(before=document.paragraphs()[0].id, levels=(1, 9))
    assert [document.paragraph(i).style_name for i in result.created[1:]] == [f"toc {n}" for n in range(1, 10)]
    assert _entries([xml(document.paragraph(i)._element) for i in result.created[1:]]) == \
        _entries(OBSERVED["tocstyles"]["blocks"])


def test_a_toc_is_rebuilt_when_headings_change_and_its_numbers_follow_the_layout():
    document = _toc_document({})
    result = document.insert_toc(before=document.paragraphs()[0].id)
    toc = result.id
    heading = next(p for p in document.paragraphs() if p.text == "Part 2.2")
    heading.set_text("Part 2.2, renamed")
    # Push Chapter 3 a page on.
    chapter3 = next(p for p in document.paragraphs() if p.text == "Chapter 3")
    chapter3.format(page_break_before=True)
    document.update_fields()
    lines = document.field(toc).result.split("\n")
    assert "Part 2.2, renamed\t3" in lines
    labels = {p.text: document.page_label(p.id) for p in document.paragraphs() if p.style_name and
              p.style_name.startswith("heading")}
    assert all(f"{text}\t{labels[text]}" in lines for text in labels)
    assert set(check(Document.open(document.to_bytes()).package)) == set()


def test_a_tracked_toc_is_one_insertion_with_its_links_outside_the_revisions():
    document = _toc_document({})
    with document.tracking(author="A", date="2026-10-04T12:00:00Z"):
        result = document.insert_toc(before=document.paragraphs()[0].id)
    paragraphs = [document.paragraph(i)._element for i in result.created[1:]]
    assert all(p.find(f"{_W}pPr/{_W}rPr/{_W}ins") is not None for p in paragraphs)
    assert all(link.getparent().tag == _W + "p" for p in paragraphs for link in p.iter(_W + "hyperlink"))
    assert not [n for p in paragraphs for n in p.iter(_W + "ins") if n.find(_W + "hyperlink") is not None]
    count = len(document.paragraphs())
    document.reject_all()
    assert len(document.paragraphs()) == count - len(paragraphs)
    assert not [f for f in document.fields() if f.keyword == "TOC"]


def test_a_tracked_break_removal_deletes_its_mark_and_accepting_joins_the_sections():
    """``tracked-removal``: Word deletes the break's paragraph mark as a revision, the
    section properties staying on it until accepted."""
    document = Document.open(e4_probe.sectioned_document())
    first = document.sections()[0].id
    paragraph = document.paragraph(document.section(first).paragraph_ids[-1])
    with document.tracking(author="A", date="2026-10-04T12:00:00Z"):
        document.remove_section_break(first)
    element = paragraph._element
    assert element.find(f"{_W}pPr/{_W}rPr/{_W}del") is not None and element.find(f"{_W}pPr/{_W}sectPr") is not None
    word = next(b for b in OBSERVED["tracked-removal"]["blocks"] if "<w:del " in b)
    assert re.search(r"<w:pPr><w:rPr><w:del [^>]*/></w:rPr><w:sectPr>", word)
    accepted = Document.open(document.to_bytes())
    accepted.accept_all()
    assert len(accepted.sections()) == 2
    document.reject_all()
    assert len(document.sections()) == 3
