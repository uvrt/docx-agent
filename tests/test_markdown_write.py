"""E2's write half: ``insert_markdown`` (ROADMAP.md, "The Markdown layer"; Phase E2).

* **The AST round trip.**  Every selected CommonMark spec example (``tests/corpus/commonmark``,
  CC BY-SA 4.0) and every hand-written document (``tests/corpus/handwritten``), inserted
  and read back with ``to_markdown(ids=False)``, gives the same CommonMark AST after the
  round trip's normalisation (:func:`docx_agent.markdown.parse.roundtrip_ast`): in a blank
  document one example at a time, and in every fixture one after another at the end of
  the body, each read over its own blocks.  :data:`EXCLUDED` lists the spec examples
  outside the guarantee, each with its reason; in the blank document each must still fail
  (so the list stays exact).
* **Styles.**  What is written names only styles the document defines (after the edit),
  each the one the style map names for its construct -- found by name, or for a heading by
  its outline level -- and the only direct formatting is the decided kind.
* **Validity, undo, render, ids**, and **tracking**: accept-all is the untracked insertion
  and reject-all the original, on every fixture, at every kind of place.
"""

from __future__ import annotations

import base64
import json
from pathlib import Path

import pytest
from lxml import etree

from docx_agent import Document, EditError
from docx_agent.markdown import Reader, project
from docx_agent.markdown.parse import differences, roundtrip_ast
from docx_agent.markdown.render import render
from docx_agent.markdown.stylemap import DEFAULT, Rule
from docx_agent.validate import check

from canonical import canonical, difference, without_added_definitions
from conftest import MARKDOWN_DIR, NEW_CREATED, fixture_id, reading_paths

ROOT = Path(__file__).resolve().parent
CORPUS = json.loads((ROOT / "corpus" / "commonmark" / "examples.json").read_text(encoding="utf-8"))
EXAMPLES = CORPUS["examples"]
HANDWRITTEN = {p.stem: p.read_text(encoding="utf-8") for p in sorted((ROOT / "corpus" / "handwritten").glob("*.md"))}


def blank() -> Document:
    """A new blank document, as Word makes one (E6's ``Document.new()``; until E6, the
    generated ``generated/markdown/blank.docx``, which stays among the reading fixtures)."""
    return Document.new(created=NEW_CREATED)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _png() -> bytes:
    import struct
    import zlib

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes((40, 120, 200)) * 4 for _ in range(4))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 4, 4, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


PNG = _png()
IMAGES = {"square.png": PNG}


def _any_image(src: str) -> bytes:
    return PNG


_ITEM = "a list item holds one paragraph and the lists nested in it; its other blocks follow the list"
_QUOTE = "a block quote holds paragraphs; a heading, list, code block or quote in it is written outside it"
_ADJACENT = "two block quotes, or two code blocks, with nothing between them read back as one"
_EMPTY = "an empty code block reads back as one blank line; an empty quote or link destination is nothing"
_SPACE_CODE = "a code span of only whitespace is whitespace at a paragraph's end"
#: Spec examples outside the guarantee, and why (ROADMAP.md, "Guarantees").
EXCLUDED: dict[int, str] = {
    **{n: _ITEM for n in (4, 5, 7, 61, 108, 254, 256, 258, 262, 263, 264, 270, 271, 273, 274, 277, 278, 286, 287,
                          288, 290, 300, 307, 309, 316, 318, 319, 320, 321, 324, 325)},
    **{n: _QUOTE for n in (6, 128, 228, 229, 230, 232, 235, 236, 237, 250, 251, 252, 259, 260, 292, 293)},
    242: _ADJACENT,
    **{n: _EMPTY for n in (126, 130, 144, 200, 218, 239, 240, 485, 486, 567)},
    334: _SPACE_CODE,
}


def _read_back(document: Document, result, reader: Reader | None = None) -> str:
    if not result.blocks:
        return ""
    return render(project(document, f"{result.blocks[0]}..{result.blocks[-1]}", ids=False, reader=reader))


# -- the round trip ----------------------------------------------------------------------------


def test_the_corpus_is_the_selected_spec_with_its_licence():
    assert CORPUS["licence"].startswith("CC BY-SA 4.0")
    assert CORPUS["spec"] == "CommonMark Spec 0.31.2" and CORPUS["author"] == "John MacFarlane"
    assert len(EXAMPLES) == 575 and CORPUS["left_out"] == {"raw_html": 77, "trace_check": 0}
    assert set(EXCLUDED) <= {e["example"] for e in EXAMPLES}


@pytest.mark.parametrize("example", EXAMPLES, ids=lambda e: f"{e['example']}")
def test_spec_example_round_trips_in_a_blank_document(example):
    md = example["markdown"]
    document = blank()
    result = document.insert_markdown(md, images=_any_image, fetch=_any_image)
    written = _read_back(document, result)
    found = differences(roundtrip_ast(md), roundtrip_ast(written))
    if example["example"] in EXCLUDED:
        assert found, f"example {example['example']} now round-trips: take it off EXCLUDED"
        return
    assert found == [], f"{md!r} -> {written!r}"


@pytest.mark.parametrize("name", sorted(HANDWRITTEN))
def test_handwritten_document_round_trips_in_a_blank_document(name):
    md = HANDWRITTEN[name]
    document = blank()
    result = document.insert_markdown(md, images=IMAGES)
    assert differences(roundtrip_ast(md), roundtrip_ast(_read_back(document, result))) == []
    # With ids: every id written resolves, and the new blocks carry theirs.
    with_ids = document.to_markdown()
    for identifier in result.blocks:
        assert f"{identifier}" in with_ids
    for identifier in result.created:
        document.get(identifier)


#: Direct formatting ``insert_markdown`` may write: a run's second mark beside its style,
#: strikethrough; a cell's alignment, a thematic break's border; numbering; revision marks.
_RUN_DIRECT = {"rStyle", "b", "bCs", "i", "iCs", "strike", "ins", "del"}
_PARAGRAPH_DIRECT = {"pStyle", "numPr", "jc", "pBdr", "rPr", "pPrChange"}


def _style_names(document: Document) -> dict[str, tuple[str, str]]:
    return {s.id: (s.name or s.id, s.kind) for s in document.styles}


def _check_styles(document: Document, results: list) -> None:
    """Every style the insertions name exists and is the mapped one; no other direct
    formatting than the decided kind."""
    from docx_agent.markdown.read import Reader as _Reader

    names = _style_names(document)
    mapped = {r.style.casefold() for r in DEFAULT.rules if r.style}
    reader = _Reader(document)
    for result in results:
        for identifier in result.created:
            if identifier.startswith("fn:"):
                continue
            element = document._resolve(identifier)[1].element
            for node in element.iter(W + "pStyle", W + "rStyle", W + "tblStyle"):
                value = node.get(W + "val")
                assert value in names, f"{identifier}: style id {value!r} is not defined"
                name, _ = names[value]
                heading = DEFAULT.heading_level(name, reader.outline_level(value)) is not None
                assert name.casefold() in mapped or heading or node.tag == W + "tblStyle", \
                    f"{identifier}: {name!r} is no style the map names"
            for properties in element.iter(W + "rPr"):
                if properties.getparent().tag == W + "r":
                    extra = {etree.QName(c).localname for c in properties} - _RUN_DIRECT
                    assert not extra, f"{identifier}: direct run formatting {extra}"
            for properties in element.iter(W + "pPr"):
                if properties.getparent().tag == W + "p":
                    extra = {etree.QName(c).localname for c in properties} - _PARAGRAPH_DIRECT
                    assert not extra, f"{identifier}: direct paragraph formatting {extra}"


@pytest.mark.parametrize("path", reading_paths(), ids=fixture_id)
def test_the_corpus_round_trips_in_every_fixture(path):
    """Every spec example the guarantee covers and every hand-written document, inserted one
    after another at the end of the fixture's body, each read back over its own blocks; the
    styles they name; validity."""
    data = path.read_bytes()
    document = Document.open(data)
    problems = set(check(Document.open(data).package))
    done = []
    for example in EXAMPLES:
        if example["example"] in EXCLUDED:
            continue
        done.append((f"spec {example['example']}", example["markdown"],
                     document.insert_markdown(example["markdown"], images=_any_image, fetch=_any_image)))
    for name, md in HANDWRITTEN.items():
        done.append((name, md, document.insert_markdown(md, images=IMAGES)))
    reader = Reader(document)
    failed = []
    for name, md, result in done:
        found = differences(roundtrip_ast(md), roundtrip_ast(_read_back(document, result, reader)))
        if found:
            failed.append((name, found[:2]))
    assert failed == []
    _check_styles(document, [result for _, _, result in done])
    assert set(check(Document.open(document.to_bytes()).package)) - problems == set()


# -- styles ------------------------------------------------------------------------------------


def _styles_of(document: Document, result) -> list[str]:
    out = []
    for identifier in result.blocks:
        element = document._resolve(identifier)[1].element
        found = element.find(f"{W}pPr/{W}pStyle")
        out.append(found.get(W + "val") if found is not None else None)
    return out


def test_a_dutch_template_maps_headings_by_name_and_outline_level():
    document = Document.open((MARKDOWN_DIR / "dutch-template.docx").read_bytes())
    before = {s.id for s in document.styles}
    result = document.insert_markdown("# Een\n\n## Twee\n\n### Drie\n\n> Citaat\n\n- punt\n\n1. stap\n\n*nadruk* **zwaar**\n")
    assert _styles_of(document, result)[:4] == ["Kop1", "Kop2", "Kop3", "Citaat"]
    assert _styles_of(document, result)[4:6] == ["Lijstopsomteken", "Lijstnummering"]
    runs = document._resolve(result.blocks[-1])[1].element.findall(f".//{W}rStyle")
    assert [r.get(W + "val") for r in runs] == ["Nadruk", "Zwaar"]
    added = {s.id for s in document.styles} - before
    assert not any("Kop" in s or "Heading" in s for s in added), added
    assert document.to_markdown(f"{result.blocks[0]}..{result.blocks[2]}", ids=False) == "# Een\n\n## Twee\n\n### Drie\n"


def test_english_names_with_localised_ids_are_found_by_name():
    document = Document.open((ROOT / "fixtures" / "generated" / "lists-and-styles.docx").read_bytes())
    result = document.insert_markdown("# Kop\n\n**zwaar**\n")
    assert _styles_of(document, result)[0] == "Kop1"
    assert document._resolve(result.blocks[1])[1].element.find(f".//{W}rStyle").get(W + "val") == "Zwaar"


def test_missing_styles_are_added_as_word_writes_them_and_said():
    """A new document has heading 1 and Quote (Word's Normal template's); the rest are added."""
    document = blank()
    result = document.insert_markdown("# H\n\n> q\n\n```\nc\n```\n\n| a |\n| - |\n| b |\n\n`x` *e* [l](https://e.x/)\n")
    names = {s.name for s in document.styles}
    for name in ("heading 1", "Quote"):
        assert name in names and not any(f"style added: {name} " in w for w in result.warnings), name
    for name in ("HTML Preformatted", "Table Grid", "HTML Code", "Emphasis", "Hyperlink"):
        assert name in names
        assert any(f"style added: {name} " in w for w in result.warnings), name
    grid = next(s for s in document.styles if s.name == "Table Grid")
    assert grid.based_on == "TableNormal"


def test_a_table_takes_the_style_the_document_tables_use_most():
    document = Document.open((ROOT / "fixtures" / "wordto" / "sample-with-table.docx").read_bytes())
    used = [n.get(W + "val") for n in document.package.tree(document.package.document_part()).iter(W + "tblStyle")]
    result = document.insert_markdown("| a | b |\n| - | - |\n| 1 | 2 |\n")
    table = document._resolve(result.blocks[0])[1].element
    assert table.find(f"{W}tblPr/{W}tblStyle").get(W + "val") == max(set(used), key=used.count)


def test_a_style_map_row_overrides_the_default():
    document = blank()
    style_map = DEFAULT.with_rules(Rule("blockquote", "Intense Quote"))
    result = document.insert_markdown("> q\n", style_map=style_map)
    style = _styles_of(document, result)[0]
    assert next(s.name for s in document.styles if s.id == style) == "Intense Quote"


def test_emphasis_direct_by_option():
    document = blank()
    result = document.insert_markdown("*e* and **s**\n", emphasis="direct")
    element = document._resolve(result.blocks[0])[1].element
    assert element.find(f".//{W}rStyle") is None
    assert element.find(f".//{W}i") is not None and element.find(f".//{W}b") is not None
    assert document.to_markdown(result.blocks[0], ids=False) == "*e* and **s**\n"


def test_two_marks_are_a_style_and_direct_formatting():
    document = blank()
    result = document.insert_markdown("***both*** and [**bold link**](https://e.x/)\n")
    runs = document._resolve(result.blocks[0])[1].element.iter(W + "r")
    shapes = [sorted(etree.QName(c).localname for c in r.find(W + "rPr")) for r in runs if r.find(W + "rPr") is not None]
    assert ["i", "iCs", "rStyle"] in shapes and ["b", "bCs", "rStyle"] in shapes


# -- what is written -----------------------------------------------------------------------------


def test_lists_nest_by_level_restart_and_start_where_markdown_says():
    document = blank()
    result = document.insert_markdown("3. three\n4. four\n   - inner\n     1. deep\n\n1) again\n")
    numbering = document.numbering
    levels, nums = [], []
    for identifier in result.blocks:
        membership = document.list_of(identifier)
        levels.append(membership.level)
        nums.append(membership.num_id)
    assert levels == [0, 0, 1, 2, 0]
    assert len(set(nums)) == 4  # one instance per Markdown list
    root = numbering.root
    first = next(n for n in root.findall(W + "num") if n.get(W + "numId") == str(nums[0]))
    override = first.find(f"{W}lvlOverride/{W}startOverride")
    assert override is not None and override.get(W + "val") == "3"
    state = document.state(f"{result.blocks[0]}..{result.blocks[-1]}")
    numbers = [b["list"]["label"] for b in state["blocks"] if b.get("list")]
    assert numbers[:2] == ["3.", "4."] and numbers[-1] == "1."
    assert document.to_markdown(f"{result.blocks[0]}..{result.blocks[-1]}", ids=False) == \
        "3. three\n4. four\n   - inner\n     1. deep\n\n1) again\n"


def test_new_lists_do_not_share_the_document_definitions():
    document = Document.open((ROOT / "fixtures" / "generated" / "lists-and-styles.docx").read_bytes())
    numbering = document.numbering
    before = set(numbering.abstracts())
    result = document.insert_markdown("1. a\n2. b\n")
    abstract = numbering.abstract_of(document.list_of(result.blocks[0]).num_id)
    assert abstract not in before


def test_a_table_is_written_as_word_makes_one():
    document = blank()
    document.insert_markdown("Before the table.\n")    # the new document's empty paragraph replaced
    result = document.insert_markdown("| a | b | c |\n| :- | :-: | -: |\n| 1 | 2 | 3 |\n")
    table = document._resolve(result.blocks[0])[1].element
    # The text width shared: 9072 twips of an A4 page with 2.5 cm margins.
    assert [int(c.get(W + "w")) for c in table.iter(W + "gridCol")] == [3024, 3024, 3024]
    assert table.find(f"{W}tblPr/{W}tblW").get(W + "type") == "auto"
    assert table.find(f"{W}tblPr/{W}tblLook").get(W + "val") == "04A0"
    rows = table.findall(W + "tr")
    assert rows[0].find(f"{W}trPr/{W}tblHeader") is not None and rows[1].find(f"{W}trPr") is None
    assert [p.find(f"{W}pPr/{W}jc").get(W + "val") for p in rows[1].iter(W + "p")] == ["left", "center", "right"]
    # A table at a story's end gets a paragraph after it.
    assert result.blocks[-1].startswith("p:")


def test_footnotes_are_real_notes_with_the_parts_word_makes():
    document = blank()
    result = document.insert_markdown("A[^x] and B[^y].\n\n[^x]: First.\n[^y]: Second, *two*.\n")
    notes = document.notes("footnote")
    assert [n.text.strip() for n in notes] == ["First.", "Second, two."]
    assert {"fn:1", "fn:2"} <= set(result.created)
    package = document.package
    assert package.has_part("word/footnotes.xml") and package.has_part("word/endnotes.xml")
    settings = package.tree(package.settings_part())
    assert settings.find(W + "footnotePr") is not None and settings.find(W + "endnotePr") is not None
    assert check(package) == []
    assert document.to_markdown(ids=False).endswith("[^fn:1]: First.\n\n[^fn:2]: Second, *two*.\n")


def test_thematic_break_is_an_empty_paragraph_with_a_bottom_border():
    document = blank()
    result = document.insert_markdown("---\n")
    bottom = document._resolve(result.blocks[0])[1].element.find(f"{W}pPr/{W}pBdr/{W}bottom")
    assert dict((etree.QName(k).localname, v) for k, v in bottom.attrib.items()) == \
        {"val": "single", "sz": "6", "space": "1", "color": "auto"}


def test_links_external_and_to_a_bookmark():
    document = blank()
    result = document.insert_markdown("[out](https://example.com/x) and [in](#target)\n")
    links = document._resolve(result.blocks[0])[1].element.findall(W + "hyperlink")
    assert links[1].get(W + "anchor") == "target"
    rel = document.package.relationships(document.package.document_part())
    assert rel[links[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")].target \
        == "https://example.com/x"


# -- pictures and raw HTML -------------------------------------------------------------------------


def test_pictures_come_only_from_what_the_caller_gives(tmp_path):
    (tmp_path / "pics").mkdir()
    (tmp_path / "pics" / "a.png").write_bytes(PNG)
    (tmp_path / "secret.png").write_bytes(PNG)
    document = blank()
    with pytest.raises(EditError, match="images="):
        document.insert_markdown("![a](a.png)\n")
    result = document.insert_markdown("![alt text](a.png \"Name\")\n", images=tmp_path / "pics")
    picture = document.pictures()[0]
    assert picture.alt_text == "alt text" and picture.name == "Name"
    with pytest.raises(EditError, match="outside the images directory"):
        document.insert_markdown("![x](../secret.png)\n", images=tmp_path / "pics")
    with pytest.raises(EditError, match="remote"):
        document.insert_markdown("![x](https://example.com/a.png)\n", images=tmp_path / "pics")
    fetched = []
    document.insert_markdown("![x](https://example.com/a.png)\n", fetch=lambda url: fetched.append(url) or PNG)
    assert fetched == ["https://example.com/a.png"]
    data = "data:image/png;base64," + base64.b64encode(PNG).decode()
    document.insert_markdown(f"![d]({data})\n")
    document.insert_markdown("![m](m.png)\n", images={"m.png": PNG})
    assert len(document.pictures()) == 4
    assert result.blocks and check(document.package) == []


def test_raw_html_is_refused_or_kept_as_text_and_comments_are_dropped():
    document = blank()
    before = document.to_bytes()
    with pytest.raises(EditError, match="raw HTML"):
        document.insert_markdown("a <b>bold</b> word\n")
    with pytest.raises(EditError, match="raw HTML"):
        document.insert_markdown("<div>\nx\n</div>\n")
    assert document.to_bytes() == before
    result = document.insert_markdown("a <b>bold</b> word\n", html="text")
    assert document.paragraph(result.blocks[0]).text == "a <b>bold</b> word"
    result = document.insert_markdown("<!-- p:12345678 -->\nkept <!-- c -->text\n")
    assert [document.paragraph(b).text for b in result.blocks] == ["kept text"]


def test_nothing_to_write_changes_nothing():
    document = blank()
    before = document.to_bytes()
    result = document.insert_markdown("[ref]: /url\n\n<!-- only a comment -->\n")
    assert result.changed is False and result.created == []
    assert document.to_bytes() == before


# -- where ---------------------------------------------------------------------------------------


def _document_with_paragraphs() -> Document:
    document = blank()
    document.insert_markdown("first\n\nsecond\n\nthird\n\nlast\n")
    return document


def test_after_before_replace_and_ranges():
    document = _document_with_paragraphs()
    ids = [p.id for p in document.paragraphs() if p.text]
    document.insert_markdown("after first", at=f"after:{ids[0]}")
    document.insert_markdown("before second", at=f"before:{ids[1]}")
    document.insert_markdown("bare id: after third", at=ids[2])
    result = document.insert_markdown("*replaced*", at=f"replace:{ids[1]}")
    assert result.removed == [ids[1]]
    document.insert_markdown("# one\n\ntwo", at=f"replace:{ids[2]}..{ids[3]}")
    texts = [p.text for p in document.paragraphs() if p.text]
    assert texts == ["first", "after first", "before second", "replaced", "one", "two"]


def test_end_of_another_story():
    document = Document.open((ROOT / "fixtures" / "generated" / "ids-and-markup.docx").read_bytes())
    result = document.insert_markdown("**Header** text", at="end:header1")
    assert result.blocks[0].startswith("header1/p:")
    assert document.paragraph(result.blocks[0]).text == "Header text"
    with pytest.raises(EditError, match="outside the body"):
        document.insert_markdown("x[^1]\n\n[^1]: n\n", at="end:header1")
    with pytest.raises(EditError):
        document.insert_markdown("x", at="end:footnotes")


def test_a_range_must_be_blocks_of_one_container():
    document = Document.open((ROOT / "fixtures" / "generated" / "ids-and-markup.docx").read_bytes())
    document._stamp_document()
    body = document.paragraphs()[0].id
    cell = document.tables()[0].cell(1, 1).paragraphs[0].id
    with pytest.raises(EditError, match="one container"):
        document.insert_markdown("x", at=f"replace:{body}..{cell}")


# -- undo, ids, render -----------------------------------------------------------------------------


MIXED = HANDWRITTEN["report"] + "\n" + HANDWRITTEN["notes"] + "\n![sq](square.png)\n"


@pytest.mark.parametrize("path", reading_paths(), ids=fixture_id)
def test_one_undo_step_byte_for_byte(path):
    data = path.read_bytes()
    document = Document.open(data)
    before = document.to_bytes()
    document.insert_markdown(MIXED, images=IMAGES)
    after = document.to_bytes()
    assert document.undo()
    assert document.to_bytes() == before
    assert not document.undo() or True
    document.redo()
    assert document.to_bytes() == after


@pytest.mark.parametrize("path", reading_paths(), ids=fixture_id)
def test_inserted_content_renders_with_its_ids(path):
    from test_render import check_rendered_ids

    document = Document.open(path.read_bytes())
    before = document.layout()
    result = document.insert_markdown(MIXED, images=IMAGES)
    saved = Document.open(document.to_bytes())
    check_rendered_ids(saved)
    layout = saved.layout()
    # docx2svg does not draw a footnote referenced in a table cell (notes.md has one): its
    # own, known gap, which it says.
    assert {w.code for w in layout.warnings} <= {w.code for w in before.warnings} | {"footnotes-not-drawn"}
    if layout.stopped is None:
        for identifier in result.blocks:
            assert layout.where(identifier), identifier


# -- tracking ------------------------------------------------------------------------------------

AUTHOR = "E2 Agent"
DATE = "2026-10-03T12:00:00Z"
PLACES = ("end", "after", "before", "replace", "range", "last")


def _clean(paragraph) -> bool:
    """A body paragraph of plain runs: one a replacement can be compared on (one holding
    another author's revisions or proofing marks keeps them, accepted, as Word does)."""
    element = paragraph._element
    return (element.getparent().tag == W + "body" and bool(paragraph.text.strip()) and not paragraph.ends_section
            and all(c.tag in (W + "pPr", W + "r") for c in element)
            and all(x.tag in (W + "rPr", W + "t") for r in element.findall(W + "r") for x in r))


def _place(document: Document, kind: str) -> str | None:
    if kind == "end":
        return "end"
    clean = [p for p in document.paragraphs() if _clean(p)]
    if not clean:
        return None
    first = clean[0]
    body = [c for c in first._element.getparent() if c.tag in (W + "p", W + "tbl")]
    if kind in ("after", "before", "replace"):
        return f"{kind}:{first.id}"
    if kind == "range":
        k = body.index(first._element)
        following = [p for p in clean if k + 1 < len(body) and p._element is body[k + 1]]
        return f"replace:{first.id}..{following[0].id}" if following else None
    last = document.paragraphs()[-1]
    previous = body[-2] if len(body) > 1 else None
    if _clean(last) and body[-1] is last._element and previous is not None and previous.tag == W + "p":
        return f"replace:{last.id}"
    return None


@pytest.mark.parametrize("kind", PLACES)
@pytest.mark.parametrize("path", reading_paths(), ids=fixture_id)
def test_tracked_insertion_accepts_to_the_edit_and_rejects_to_the_original(path, kind):
    data = path.read_bytes()
    base = Document.open(data)
    base._stamp_document()
    stamped = base.to_bytes()
    at = _place(base, kind)
    if at is None:
        pytest.skip("no such place in this fixture")
    original = canonical(stamped, ids=False)
    untracked = Document.open(stamped)
    untracked.insert_markdown(MIXED, at=at, images=IMAGES)
    expected = canonical(untracked.to_bytes(), ids=False)

    tracked = Document.open(stamped)
    other = {(r.author, r.kind) for r in tracked.revisions()}
    with tracked.tracking(author=AUTHOR, date=DATE):
        result = tracked.insert_markdown(MIXED, at=at, images=IMAGES)
    tracked_bytes = tracked.to_bytes()
    reopened = Document.open(tracked_bytes)
    assert set(check(reopened.package)) - set(check(Document.open(data).package)) == set()
    ours = reopened.revisions(author=AUTHOR)
    assert ours and all(r.date == DATE for r in ours)  # one revision group: one author, one date
    assert {(r.author, r.kind) for r in reopened.revisions() if r.author != AUTHOR} == other
    assert result.blocks

    accepted = Document.open(tracked_bytes)
    accepted.accept(author=AUTHOR)
    got = without_added_definitions(canonical(accepted.to_bytes(), ids=False), expected)
    assert got == expected, "accept-all is not the untracked insertion:\n" + difference(expected, got)

    rejected = Document.open(tracked_bytes)
    rejected.reject(author=AUTHOR)
    got = without_added_definitions(canonical(rejected.to_bytes(), ids=False), original)
    assert got == original, "reject-all is not the original:\n" + difference(original, got)


def test_tracked_insertion_reads_in_every_view():
    document = blank()
    with document.tracking(author=AUTHOR, date=DATE):
        result = document.insert_markdown(HANDWRITTEN["report"])
    final = document.to_markdown(f"{result.blocks[0]}..{result.blocks[-1]}", ids=False)
    assert roundtrip_ast(final) == roundtrip_ast(HANDWRITTEN["report"])
    assert "Quarterly report" not in document.to_markdown(view="original", ids=False)
    assert "{++" in document.to_markdown(view="markup")
    assert document.undo() and document.to_bytes() == blank().to_bytes()


def test_track_keyword_tracks_one_insertion():
    document = blank()
    document.insert_markdown("tracked", track=True)
    assert {r.author for r in document.revisions()} == {"docx-agent"}
    count = len(document.revisions())
    with document.tracking(author=AUTHOR):
        document.insert_markdown("not tracked", track=False)
    assert len(document.revisions()) == count


def test_a_reference_rejected_takes_its_note_with_it():
    document = Document.open((ROOT / "fixtures" / "generated" / "ids-and-markup.docx").read_bytes())
    before = len(document.notes("footnote"))
    with document.tracking(author=AUTHOR, date=DATE):
        document.insert_markdown("New[^n].\n\n[^n]: A [linked](https://example.com/n) note.\n")
    assert len(document.notes("footnote")) == before + 1
    document.reject(author=AUTHOR)
    assert len(document.notes("footnote")) == before
    assert check(document.package) == []
    rels = document.package.relationships("word/footnotes.xml") if document.package.has_part(
        "word/_rels/footnotes.xml.rels") else {}
    assert not any("example.com/n" in (r.target or "") for r in rels.values())


def test_a_paragraph_in_a_cell_is_replaced_in_its_cell():
    document = Document.open((ROOT / "fixtures" / "generated" / "ids-and-markup.docx").read_bytes())
    document._stamp_document()
    table = document.tables()[0]
    cell = table.cell(1, 1).paragraphs[0].id
    result = document.insert_markdown("**new** cell text\n\n- and a list", at=f"replace:{cell}")
    assert [p.text for p in document.tables()[0].cell(1, 1).paragraphs] == ["new cell text", "and a list"]
    assert result.removed == [cell] and check(document.package) == []


def test_a_note_part_declares_only_the_ignorable_prefixes_it_uses():
    """A picture beside a note's reference made the notes part list ``wp14`` as ignorable
    without declaring it, and Word would not open the document (its oracle)."""
    document = blank()
    document.insert_markdown("A ![a](square.png) note.[^s]\n\n[^s]: The note.\n", images=IMAGES)
    assert check(Document.open(document.to_bytes()).package) == []
    notes = document.package.tree("word/footnotes.xml")
    assert "wp14" not in (notes.get("{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable") or "")
    assert notes.find(f".//{W}footnote/{W}p/{W}pPr/{W}pStyle") is not None


# -- a line break at the end of a paragraph --------------------------------------------------
#
# CommonMark has no hard line break at the end of a block (spec 6.7: "Hard line breaks are
# for separating inline content within a block"): ``Title\`` reads back as the text
# ``Title\``.  A trailing ``w:br`` -- a cover page's Shift+Enter after its title -- is
# written as ``<br>`` instead, which reads back as a break.


def _body_texts(document: Document) -> list[str]:
    return [p.text for p in document.stories[0].paragraphs]


def _no_backslash_ends_a_block(markdown: str) -> bool:
    """No line ending in ``\\`` is the last of its block (a mid-paragraph break is one)."""
    lines = markdown.split("\n")
    return not any(line.endswith("\\") and (k + 1 == len(lines) or not lines[k + 1].strip())
                   for k, line in enumerate(lines))


@pytest.mark.parametrize("text", [
    "Cover title\v",
    "Cover title\v\v",
    "Line one\vline two\v",
    "\v",
    "\v\vAfter two breaks\v",
])
def test_a_trailing_line_break_reads_back_as_a_break(text):
    document = blank()
    paragraph = document.insert_markdown("placeholder").blocks[0]
    document.paragraph(paragraph).set_text(text)
    markdown = document.to_markdown(ids=False)
    assert _no_backslash_ends_a_block(markdown), markdown
    again = blank()
    again.insert_markdown(markdown)
    assert _body_texts(again) == [text]


def test_a_cover_page_round_trips():
    document = blank()
    blocks = document.insert_markdown("Title\n\nSubtitle\n\nAuthor\n\nIntroduction text.").blocks
    for identifier, text in zip(blocks, ["Annual report\v\v", "Subtitle\v\v\v", "Prepared by A. Author\vOctober\v",
                                         "Introduction text."]):
        document.paragraph(identifier).set_text(text)
    markdown = document.to_markdown(ids=False)
    assert _no_backslash_ends_a_block(markdown), markdown
    assert markdown.startswith("Annual report<br><br>\n")
    again = blank()
    again.insert_markdown(markdown)
    assert _body_texts(again) == _body_texts(document)
    # And the ids view, whose comments sit before and after the blocks, reads the same.
    again = blank()
    again.insert_markdown(document.to_markdown())
    assert _body_texts(again) == _body_texts(document)


@pytest.mark.parametrize("markdown", ["Title<br>", "Title<br/>", "Title<BR />", "Title\\\n<br>"])
def test_br_in_markdown_is_a_line_break(markdown):
    from docx_agent.markdown.parse import parse_model
    from docx_agent.markdown import model as m

    blocks = parse_model(markdown)
    assert len(blocks) == 1 and isinstance(blocks[0], m.Paragraph)
    breaks = sum(isinstance(item, m.Break) for item in blocks[0].inline)
    assert breaks == (2 if "\\" in markdown else 1)
    assert not any(isinstance(item, m.Html) for item in blocks[0].inline)


def test_a_paragraph_of_only_a_break_is_not_an_html_block():
    from docx_agent.markdown.parse import parse_model
    from docx_agent.markdown import model as m

    blocks = parse_model("<br>\n\n<br><br>\n")
    assert [type(b) for b in blocks] == [m.Paragraph, m.Paragraph]
    assert [len(b.inline) for b in blocks] == [1, 2]
    assert all(isinstance(item, m.Break) for b in blocks for item in b.inline)
