"""``to_markdown``, construct by construct, on E2's generated fixtures and the corpus: what
each Word construct is written as, in each view, and the ranges a projection can cover."""

from __future__ import annotations

import pytest

from docx_agent import Document
from docx_agent.markdown import DEFAULT, Rule, check, parse_ast


@pytest.fixture(scope="module")
def constructs(constructs_doc):
    return Document.open(constructs_doc).to_markdown(headers=True)


def block(markdown: str, identifier: str) -> str:
    """The lines from ``identifier``'s comment up to the next blank line."""
    start = markdown.index(f"<!-- {identifier}")
    end = markdown.find("\n\n", start)
    return markdown[start:end if end >= 0 else None]


# -- headings --------------------------------------------------------------------------------


def test_headings_by_style(constructs):
    assert block(constructs, "p:6B000002") == "<!-- p:6B000002 -->\n# Headings and text"
    assert block(constructs, "p:6B000004") == "<!-- p:6B000004 -->\n## A second level"
    assert block(constructs, "p:6B000005") == "<!-- p:6B000005 -->\n### A third level"
    # Deeper than Markdown goes: level 6, the style and level said.
    assert block(constructs, "p:6B000006") == "<!-- p:6B000006 style: heading 7 level 7 -->\n###### A seventh level"
    # A style no rule names: a paragraph, the style said.
    assert block(constructs, "p:6B000001") == "<!-- p:6B000001 style: Title -->\nMarkdown constructs"


def test_a_dutch_template_has_headings_by_outline_level_and_by_name(dutch_doc):
    markdown = Document.open(dutch_doc).to_markdown()
    # "Kop 1" and "Kop 2" are found by their w:outlineLvl, "heading 3" (id Kop3, no
    # outline level) by its name; neither the ids nor the Dutch names matter.
    assert "<!-- p:5E000002 style: Kop 1 -->\n# Inleiding" in markdown
    assert "<!-- p:5E000004 style: Kop 2 -->\n## Achtergrond" in markdown
    assert "<!-- p:5E000005 -->\n### Details" in markdown
    assert "<!-- p:5E000001 style: Titel -->\nEen Nederlands sjabloon" in markdown
    # Styles found by their w:name whatever their localised ids.
    assert "Tekst met *nadruk* en **zwaar**." in markdown
    assert "> <!-- p:5E000006 -->\n> Een citaat." in markdown
    assert "- Eerste punt <!-- p:5E000007 -->\n- Tweede punt <!-- p:5E000008 -->" in markdown
    assert "1. Stap een <!-- p:5E000009 -->\n2. Stap twee <!-- p:5E00000A -->" in markdown
    assert "<!-- p:5E00000B -->\n```\nx = 1\n```" in markdown


def test_the_corpus_dutch_template_maps_heading_1_by_name(markup_doc):
    markdown = Document.open(markup_doc.parent / "lists-and-styles.docx").to_markdown()
    assert "<!-- p:5A000001 -->\n# Lijsten en stijlen" in markdown


# -- inline ----------------------------------------------------------------------------------


def test_emphasis_strong_strike_and_code(constructs):
    # Direct bold and italic read as the Strong and Emphasis styles do: from the effective
    # formatting, against the paragraph's own.
    assert ("Plain, **bold**, *italic*, ~~struck~~, **strong**, *emphasised* and `inline code`."
            in block(constructs, "p:6B000003"))


def test_emphasis_is_measured_against_the_paragraph(markup_doc):
    style = Document.open(markup_doc.parent.parent / "docx2svg" / "style-document.docx").to_markdown()
    # A heading's bold is its style's, not emphasis; a quote's italic likewise.
    assert "# Where the faces come from" in style and "**Where" not in style
    assert "> A quote is italic, so its emphasis comes out upright." in style


def test_mixed_formatting_nests(markup_doc):
    markdown = Document.open(markup_doc.parent / "lists-and-styles.docx").to_markdown()
    assert "Revenue grew by **four*teen*** percnt in the third quarter." in markdown


def test_special_characters_are_escaped(constructs):
    assert ("Stars \\* and \\_underscores\\_, a \\[bracket\\], \\<tag>, \\`ticks\\`, & an entity \\&copy;, "
            "a \\{++critic++} brace, a back\\\\slash and a pipe |.") in constructs
    assert "<!-- p:6B000008 -->\n\\# not a heading" in constructs
    assert "<!-- p:6B000009 -->\n1\\. not a list" in constructs
    ast = parse_ast(constructs)
    texts = [node[1][0][1] for node in ast if node[0] == "p" and node[1] and node[1][0][0] == "t"]
    assert "# not a heading" in texts and "1. not a list" in texts
    assert any(t.startswith("Stars * and _underscores_, a [bracket], <tag>, `ticks`, & an entity &copy;")
               for t in texts)


def test_breaks_tabs_and_hidden_text(constructs):
    assert block(constructs, "p:6B00000A") == "<!-- p:6B00000A -->\nindented by a tab"
    assert block(constructs, "p:6B00000B") == "<!-- p:6B00000B -->\nLine one\\\nline two."   # hidden text left out
    assert block(constructs, "p:6B000042") == "<!-- p:6B000042 page break -->\nAfter a page break."


def test_links_and_bookmarks(constructs):
    assert ("An [external link](https://example.com/markdown), an [internal one](#notes) and a note.[^fn:1] "
            "Another.[^en:1]") in constructs
    # The bookmark the internal link goes to, where it starts.
    assert "<!-- p:6B000023 bm:notes -->\nNotes are written at the end.[^fn:2]" in constructs


def test_footnotes_and_endnotes(constructs):
    tail = constructs[constructs.index("[^fn:1]: "):]
    assert tail == ("[^fn:1]: A first footnote. <!-- footnotes/p:1D000001 -->\n\n"
                    "[^en:1]: An endnote. <!-- endnotes/p:2E000001 -->\n\n"
                    "[^fn:2]: A second footnote, <!-- footnotes/p:1D000002 -->\n\n"
                    "    in two paragraphs. <!-- footnotes/p:1D000003 -->\n")
    ast = parse_ast(constructs)
    notes = [node for node in ast if node[0] == "footnote"]
    assert [n[1] for n in notes] == ["fn:1", "en:1", "fn:2"]
    assert len([b for b in notes[2][2] if b[0] == "p"]) == 2


def test_notes_off_and_the_notes_story(constructs_doc):
    document = Document.open(constructs_doc)
    assert "[^fn:1]:" not in document.to_markdown(notes=False)
    notes = document.to_markdown("footnotes")
    assert notes.startswith("[^fn:1]: A first footnote.") and "[^en:1]" not in notes


def test_pictures_carry_their_id_and_alt_text(constructs):
    assert 'A picture ![A blue square](d:1 "Picture 1") inline.' in constructs
    assert '![A red square](d:2 "Floating")Text beside a floating picture.' in constructs
    images = [i for node in parse_ast(constructs) if node[0] == "p" for i in node[1] if i[0] == "img"]
    assert ("img", "A blue square", "d:1", "Picture 1", None) in images


def test_text_boxes_follow_their_paragraph(constructs):
    assert block(constructs, "p:6B000026") == '<!-- p:6B000026 d:3 text-box "Text Box 3" -->\nText beside a text box.'
    assert "<!-- d:3 text-box -->\n\n<!-- p:6B0000F1 -->\nInside a text box.\n\n<!-- /d:3 -->" in constructs


def test_fields_show_their_cached_result(constructs):
    assert block(constructs, "p:6B000027") == "<!-- p:6B000027 -->\nPage 1, see Notes are written at the end.."


def test_content_controls(constructs):
    assert ("Name: <!-- cc:301 text -->Grace Hopper<!-- /cc:301 -->, agreed: "
            "<!-- cc:302 checkbox -->☒<!-- /cc:302 -->") in constructs
    assert ("<!-- cc:303 rich-text tag: clause title: Clause -->\n\n<!-- p:6B00003F -->\nA clause in a control.\n\n"
            "<!-- p:6B000040 -->\nIts second paragraph.\n\n<!-- /cc:303 -->") in constructs


def test_sections_and_headers(constructs):
    assert "<!-- p:6B000041 -->\nBefore a section break.\n\n<!-- s:6B000041 section break: continuous -->" in constructs
    assert "<!-- story: header1 (header) -->\n\n<!-- header1/p:7A0000A1 -->\nConstructs, the header" in constructs
    assert "<!-- story: footer1 (footer) -->" in constructs
    assert "header1" not in Document.open(constructs_doc_path()).to_markdown()


def constructs_doc_path():
    from conftest import MARKDOWN_DIR

    return MARKDOWN_DIR / "constructs.docx"


# -- blocks ----------------------------------------------------------------------------------


def test_quotes_and_code_blocks(constructs):
    assert ("> <!-- p:6B00000C -->\n> A quotation.\n>\n> <!-- p:6B00000D -->\n> Its second paragraph."
            in constructs)
    assert "<!-- p:6B00000E p:6B00000F -->\n```\ndef answer():\n    return 42\n```" in constructs


def test_a_bordered_empty_paragraph_is_a_thematic_break(constructs):
    assert block(constructs, "p:6B000010") == "<!-- p:6B000010 -->\n---"


def test_lists_nest_and_number(constructs):
    assert ("- Bullet one <!-- p:6B000012 -->\n  - Nested bullet <!-- p:6B000013 -->\n"
            "    - Deeper bullet <!-- p:6B000014 -->\n- Bullet two <!-- p:6B000015 -->\n\n"
            # Another list right after: another marker, so Markdown keeps them apart.
            "* Another list <!-- p:6B000016 -->\n* Its second item <!-- p:6B000017 -->") in constructs
    assert ("1. Step one <!-- p:6B000019 -->\n2. Step two <!-- p:6B00001A -->\n"
            "   1. Sub-step <!-- p:6B00001B -->\n3. Step three <!-- p:6B00001C -->") in constructs
    # A w:startOverride restarts the instance at 5, as Word numbers it.
    assert "5. Five <!-- p:6B00001E -->\n6. Six <!-- p:6B00001F -->" in constructs
    # A list from the style's numbering.
    assert "- From the style <!-- p:6B000020 -->" in constructs
    lists = [node for node in parse_ast(constructs) if node[0] == "list"]
    assert [(n[1], n[2], len(n[3])) for n in lists] == [(False, 1, 2), (False, 1, 2), (True, 1, 3), (True, 5, 2),
                                                        (False, 1, 1)]


def test_restarted_lists_in_the_corpus(markup_doc):
    markdown = Document.open(markup_doc.parent / "lists-and-styles.docx").to_markdown()
    assert ("1. First item <!-- p:5A000003 -->\n2. Second item <!-- p:5A000004 -->\n"
            "   1. A sub-item <!-- p:5A000005 -->") in markdown
    assert "1. Restarted at one <!-- p:5A000007 -->\n2. Continues the restart <!-- p:5A000008 -->" in markdown


def test_a_restart_after_an_edit_numbers_as_word_does(markup_doc):
    document = Document.open(markup_doc.parent / "lists-and-styles.docx")
    document.restart_numbering("p:5A000008", at=7)  # E1's oracle: Word's PDF shows 7.
    markdown = document.to_markdown()
    # Another instance: a list of its own, kept apart from the one before by its delimiter.
    assert "1. Restarted at one <!-- p:5A000007 -->\n\n7) Continues the restart <!-- p:5A000008 -->" in markdown
    assert check(document) == []


def test_gfm_tables(constructs):
    assert block(constructs, "t:6B000029") == (
        "<!-- t:6B000029 table style: Table Grid -->\n| Item | Count |\n| --- | ---: |\n| Pens | 12 |\n"
        "| a \\| b | 500 |")
    table = next(node for node in parse_ast(constructs) if node[0] == "table")
    assert table[3] == (None, "right")
    assert table[2][1][0] == (("t", "a | b", frozenset(), None),)


def test_tables_gfm_cannot_hold_are_html(constructs, markup_doc):
    html = block(constructs, "t:6B000033")
    assert '<table data-id="t:6B000033">' in html
    assert "<table><tr><td>Inner</td><td>cells</td></tr></table>" in html  # nested
    assert '<td data-id="t:6B000033/c1,0"><!-- p:6B00003A --><p>Two</p><!-- p:6B00003B --><p>paragraphs</p></td>' in html
    merged = Document.open(markup_doc).to_markdown()
    assert '<td data-id="t:3C4D5E01/c0,0" rowspan="2">' in merged
    assert '<td data-id="t:3C4D5E01/c0,1" colspan="2">' in merged


def test_the_corpus_tables_are_gfm(markup_doc):
    markdown = Document.open(markup_doc.parent.parent / "wordto" / "sample-with-table.docx").to_markdown()
    assert "| Format | Extension | Open Standard | Max File Size |\n| --- | --- | --- | --- |" in markdown
    assert markdown.count("<!-- t@body/") == 3


def test_empty_paragraphs_keep_their_id(markup_doc):
    markdown = Document.open(markup_doc.parent.parent / "wordto" / "sample-with-table.docx").to_markdown()
    assert "<!-- p@body/27 volatile empty -->\n\n<!-- p@body/28 volatile -->\n## Table 2" in markdown


# -- views -----------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def review(review_doc):
    document = Document.open(review_doc)
    return {view: document.to_markdown(view=view) for view in ("final", "original", "markup")}


def test_insertions_and_deletions(review):
    assert "The fee is twelve euros." in review["final"]
    assert "The fee is ten euros." in review["original"]
    assert "The fee is {--ten--}<!-- rev:101 Alice -->{++twelve++}<!-- rev:102 Alice --> euros." in review["markup"]
    assert "Bob carefully read it." in review["final"] and "Bob read it." in review["original"]


def test_moves(review):
    assert "Moved from here: .\n" in review["final"] and "Moved to here: the moved words." in review["final"]
    assert "Moved from here: the moved words." in review["original"] and "Moved to here: ." in review["original"]
    assert "{--the moved words--}<!-- rev:105 moved Alice -->" in review["markup"]
    assert "{++the moved words++}<!-- rev:107 moved Alice -->" in review["markup"]


def test_deleted_paragraph_marks_join(review):
    # As docx2svg draws the final view: one paragraph, the first one's id.
    assert ("<!-- p:7C000006 joins p:7C000007 -->\n"
            "This paragraph's mark is deleted, so it joins the next one in the final view.") in review["final"]
    assert "<!-- p:7C000007 -->\nthe next one in the final view." in review["original"]
    assert "so it joins {--¶--}<!-- rev:108 Bob -->" in review["markup"]
    # A paragraph deleted whole: gone from the final view (joined, empty, into the next).
    assert "A whole deleted paragraph" not in review["final"]
    assert "<!-- p:7C000009 joins p:7C00000A -->\nText after the deleted paragraph." in review["final"]
    # One inserted whole: gone from the original view.
    assert "A whole inserted paragraph" not in review["original"]
    assert "{++A whole inserted paragraph.++}<!-- rev:109 Alice -->{++¶++}<!-- rev:110 Alice -->" in review["markup"]
    # An empty paragraph whose mark is deleted, before a table: not drawn, not written.
    assert "p:7C00000F" not in review["final"] and "<!-- p:7C00000F empty -->" in review["original"]


def test_table_rows(review):
    assert "| Ben | 4 |" in review["final"] and "Ann" not in review["final"]
    assert "| Ann | 3 |" in review["original"] and "Ben" not in review["original"]
    assert "| {--Ann--}<!-- rev:116 Bob --> | {--3--}<!-- rev:117 Bob --> |" in review["markup"]


def test_property_changes_are_said_in_the_markup_view(review):
    assert "Now **bold**<!-- rev:113 formatting Bob --> by a revision." in review["markup"]
    assert "<!-- p:7C00000C rev:114 paragraph-properties Alice -->\n## Made a heading" in review["markup"]
    assert "rev:113" not in review["final"]


def test_comments_in_the_markup_view_only(review):
    assert ("A disputed{>>Alice: Is this right?<<}<!-- c:1A2B3C4D -->"
            "{>>Bob, in reply: Yes, it is.<<}<!-- c:2B3C4D5E --> clause.") in review["markup"]
    assert "A settled{>>Bob, resolved: Done.<<}<!-- c#3 --> point." in review["markup"]
    assert "A disputed clause." in review["final"] and "{>>" not in review["final"] + review["original"]


def test_the_markup_view_of_the_corpus(markup_doc):
    markdown = Document.open(markup_doc).to_markdown(view="markup")
    assert "Revenue grew {--12--}<!-- rev:201 Reviewer -->{++14++}<!-- rev:202 Reviewer -->% this year." in markdown
    assert "A commented{>>Reviewer: A comment.<<}<!-- c#1 --> word." in markdown


def test_unknown_view_is_refused(review_doc):
    with pytest.raises(ValueError):
        Document.open(review_doc).to_markdown(view="current-ish")
    assert Document.open(review_doc).to_markdown(view="current") == Document.open(review_doc).to_markdown()


# -- ranges ----------------------------------------------------------------------------------


def test_a_block_and_a_range_of_blocks(constructs_doc):
    document = Document.open(constructs_doc)
    assert document.to_markdown("p:6B000004", notes=False) == "<!-- p:6B000004 -->\n## A second level\n"
    span = document.to_markdown("p:6B000004..p:6B000006")
    assert span.count("<!-- p:") == 3 and span.startswith("<!-- p:6B000004")
    # A list item's range is the list item; a cell's paragraph names its table.
    assert document.to_markdown("p:6B00001E") == "5. Five <!-- p:6B00001E -->\n"
    assert document.to_markdown("p:6B00002E").startswith("<!-- t:6B000029")
    # A content control by its id; notes referenced in a range come with it.
    assert document.to_markdown("cc:303") == "<!-- cc:303 rich-text tag: clause title: Clause -->\n"
    assert document.to_markdown("p:6B000023").endswith(
        "[^fn:2]: A second footnote, <!-- footnotes/p:1D000002 -->\n\n    in two paragraphs. "
        "<!-- footnotes/p:1D000003 -->\n")


def test_a_section_and_a_story(constructs_doc):
    document = Document.open(constructs_doc)
    first = document.to_markdown("s:6B000041", notes=False)
    last = document.to_markdown("s:body", notes=False)
    assert first.startswith("<!-- p:6B000001 style: Title -->") and first.rstrip().endswith("continuous -->")
    assert last.startswith("<!-- p:6B000042 page break -->") and "The end." in last
    assert document.to_markdown("header1") == (
        "<!-- story: header1 -->\n\n<!-- header1/p:7A0000A1 -->\nConstructs, the header\n")


def test_ranges_resolve_old_ids_after_a_stamp(markup_doc):
    document = Document.open(markup_doc)
    document.paragraph("p@body/2").set_text("Stamped now.")
    renamed = document.aliases["p@body/2"]
    assert document.to_markdown("p@body/2") == f"<!-- {renamed} -->\nStamped now.\n"


# -- the style map ---------------------------------------------------------------------------


def test_the_style_map_can_be_overridden(markup_doc):
    document = Document.open(markup_doc.parent.parent / "docx2svg" / "style-document.docx")
    default = document.to_markdown()
    assert "<!-- p@body/13 volatile style: Code -->\ncode = resolve(style)" in default
    mapped = document.to_markdown(style_map=DEFAULT.with_rules(Rule("code_block", "HTML Preformatted",
                                                                    reads=("Code",))))
    assert ("<!-- p@body/13 volatile p@body/14 volatile p@body/15 volatile -->\n```\ncode = resolve(style)\n"
            "exact = 14 pt\ncontextual = True\n```") in mapped
    assert check(document, style_map=DEFAULT.with_rules(Rule("code_block", "HTML Preformatted", reads=("Code",)))) == []


def test_the_map_answers_both_directions():
    assert DEFAULT.style_for("heading", 2).style == "heading 2"
    assert DEFAULT.style_for("bullet_list", 7).style == "List Bullet 5"
    assert DEFAULT.style_for("blockquote").style == "Quote"
    assert DEFAULT.style_for("code_block").style == "HTML Preformatted"
    assert DEFAULT.style_for("code").style == "HTML Code"
    assert DEFAULT.style_for("emphasis").style == "Emphasis" and DEFAULT.style_for("strong").style == "Strong"
    assert DEFAULT.paragraph_rule("HEADING 3").construct == "heading"
    assert DEFAULT.heading_level("Kop 1", 0) == 1 and DEFAULT.heading_level("heading 4", None) == 4
    assert DEFAULT.heading_level("Body Text", None) is None
    with pytest.raises(ValueError):
        DEFAULT.with_rules(Rule("no-such-construct", "Normal"))
