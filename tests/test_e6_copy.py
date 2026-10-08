"""E6: copying blocks between documents (ROADMAP.md, Phase E6), held to what Word did when it
copied a source into a destination whose styles, list, footnote, comment, bookmark,
picture, content control and paraIds all collide (``tools/e6_probe.py``:
``copy-formatted``, ``copy-insert-file``, ``paste-*``, ``copy-formatted-tracked`` in
``tests/observations/e6-word.json``), and to the gates: valid, saved, reopened and read back,
undone and redone, laid out, and -- tracked -- accept-all the untracked copy, reject-all the
destination before it.  The sweep copies every fixture into a new document and into another
fixture.
"""

from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_agent import Document, copy_blocks
from docx_agent.validate import check

from canonical import canonical, difference, without_added_styles
from conftest import fixture_id, fixture_paths
from test_roundtrip import entries

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "tools"))
import e6_probe  # noqa: E402

OBSERVATIONS = json.loads((HERE / "observations" / "e6-word.json").read_text(encoding="utf-8"))
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
AUTHOR = "E6 Agent"
DATE = "2026-10-04T12:00:00Z"


def source() -> Document:
    return Document.open(e6_probe.copy_document("source"))


def destination() -> Document:
    return Document.open(e6_probe.copy_document("destination"))


def everything(document: Document) -> str:
    paragraphs = document.paragraphs()
    return f"{paragraphs[0].id}..{paragraphs[-1].id}"


def copied(policy: str = "use_destination", **options) -> tuple[Document, object]:
    target = destination()
    result = target.copy_blocks(source(), everything(source()), at="end", styles=policy, **options)
    return target, result


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def styles_by_name(document: Document) -> dict[str, etree._Element]:
    root = etree.fromstring(parts(document.to_bytes())["word/styles.xml"])
    return {s.find(W + "name").get(W + "val"): s for s in root.findall(W + "style")}


def _paragraph(document: Document, text: str):
    return next(p for p in document.paragraphs() if p.text.startswith(text))


def _word_blocks(probe: str) -> dict[str, etree._Element]:
    """Word's copies, by their text: the observation's blocks after the destination's end."""
    out = {}
    seen = False
    for block in OBSERVATIONS[probe]["body"]:
        xml = re.sub(r"<(/?)(\w+):", r"<\1\2_", block)
        xml = re.sub(r" (\w+):(\w+)=", r" \1_\2=", xml)
        node = etree.fromstring(xml)
        text = "".join(t.text or "" for t in node.iter("w_t"))
        if text.startswith("Destination end."):
            seen = True
            continue
        if seen and text:
            out[text[:20]] = node
    return out


def _tags(node, path: str) -> list[str]:
    found = node.find(path)
    return [] if found is None else sorted(child.tag.split("_", 1)[1] for child in found)


# -- what Word does, held ----------------------------------------------------------------------


def test_word_used_the_destinations_styles_and_imported_the_rest():
    word = OBSERVATIONS["copy-formatted"]
    assert word["styles"]["Brand"].count("7030A0") == 1          # the destination's Brand
    assert "Georgia" in word["styles"]["Source Only"]             # imported, with the source's default face
    assert 'w:sz w:val="22"' in word["styles"]["Source Only"]
    assert "Georgia" not in word["styles"]["Hyperlink"]           # a character style: not folded
    for probe in ("copy-insert-file", "paste-default", "paste-destination"):
        assert set(OBSERVATIONS[probe]["styles"]) == set(word["styles"]), probe


def test_use_destination_as_word():
    target, result = copied()
    styles = styles_by_name(target)
    assert styles["Brand"].find(f"{W}rPr/{W}color").get(W + "val") == "7030A0"
    assert _paragraph(target, "Source brand").style_name == "Brand"
    imported = styles["Source Only"]
    assert imported.find(f"{W}rPr/{W}rFonts").get(W + "ascii") == "Georgia"
    assert imported.find(f"{W}rPr/{W}sz").get(W + "val") == "22"
    assert imported.find(f"{W}rPr/{W}color").get(W + "val") == "ED7D31"
    assert styles["Hyperlink"].find(f"{W}rPr/{W}rFonts") is None
    assert set(styles) == set(OBSERVATIONS["copy-formatted"]["styles"]) - {
        "annotation text", "annotation reference", "Tekst opmerking Char", "(docDefaults)"} | {"Hyperlink"}
    assert any("Source Only" in w for w in result.warnings)
    assert check(target.package) == []


def test_keep_source_writes_the_look_direct_as_word():
    """Keep Source Formatting: no style imported, each paragraph in the default style with
    the source's look -- paragraph, mark and runs -- direct, by the same properties Word wrote."""
    target, _ = copied("keep_source")
    before = set(styles_by_name(destination()))
    assert set(styles_by_name(target)) == before
    word = _word_blocks("paste-keep")
    for paragraph in target.paragraphs():
        if not paragraph.text.startswith(("Source heading", "Source brand", "Source list", "Source only")):
            continue
        assert paragraph.style is None
        theirs = word[paragraph.text[:20]]
        ours = etree.fromstring(re.sub(r" xmlns:\w+=\"[^\"]*\"", "", etree.tostring(
            paragraph._element, encoding="unicode")).replace("w:", "w_").replace(" w14_", " w14:"),
            etree.XMLParser(recover=True))
        assert _tags(ours, "w_pPr") == _tags(theirs, "w_pPr"), paragraph.text
        assert _tags(ours, "w_pPr/w_rPr") == _tags(theirs, "w_pPr/w_rPr"), paragraph.text
        assert _tags(ours, "w_r/w_rPr") == _tags(theirs, "w_r/w_rPr"), paragraph.text
    heading = _paragraph(target, "Source heading")._element
    assert heading.find(f"{W}r/{W}rPr/{W}color").get(W + "val") == "C00000"
    assert heading.find(f"{W}r/{W}rPr/{W}rFonts").get(W + "ascii") == "Georgia"
    assert heading.find(f"{W}pPr/{W}outlineLvl").get(W + "val") == "0"
    note = _paragraph(target, "Source note")._element
    assert note.find(f".//{W}footnoteReference/../{W}rPr/{W}vertAlign").get(W + "val") == "superscript"
    assert check(target.package) == []


def test_merge_keeps_bold_italic_and_underline_as_word():
    target, _ = copied("merge")
    word = _word_blocks("paste-merge")
    for text in ("Source brand paragraph",):
        ours = _paragraph(target, text)._element
        theirs = word[text[:20]]
        assert ours.find(W + "pPr/" + W + "pStyle") is None
        assert sorted(c.tag[len(W):] for c in ours.find(f"{W}r/{W}rPr")) == _tags(theirs, "w_r/w_rPr")
    # Word's Merge Formatting made the heading a Normal paragraph; docx-agent keeps a
    # heading's level (ROADMAP.md, "Trial findings"): the destination's heading style.
    assert _paragraph(target, "Source heading").style_name == "heading 1"
    accented = next(r for r in _paragraph(target, "Source brand")._element.iter(W + "r")
                    if "".join(t.text or "" for t in r.iter(W + "t")) == "accented")
    assert sorted(c.tag[len(W):] for c in accented.find(W + "rPr")) == ["b", "u"]
    assert set(styles_by_name(target)) == set(styles_by_name(destination()))


def test_lists_are_kept_apart_as_word_keeps_them():
    """Word gave the source's list a new instance over a copy of its abstract definition,
    its nsid kept; the destination's list counts on alone."""
    word = OBSERVATIONS["copy-formatted"]["word/numbering.xml"]
    assert word.count("<w:abstractNum ") == 2 and 'w:val="5A5A0001"' in word
    target, _ = copied()
    numbering = etree.fromstring(parts(target.to_bytes())["word/numbering.xml"])
    assert len(numbering.findall(W + "abstractNum")) == 2 and len(numbering.findall(W + "num")) == 2
    assert {n.get(W + "val") for n in numbering.iter(W + "nsid")} == {"5A5A0001", "6B6B0002"}
    ours = _paragraph(target, "Source list item one").list
    theirs = _paragraph(target, "Destination list item one").list
    assert ours.num_id != theirs.num_id
    labels = [b["list"]["label"] for b in target.state()["blocks"] if b.get("t", "").startswith("Source list item")]
    assert labels == ["1.", "2."]


def test_continue_joins_the_destinations_list():
    target = destination()
    item = _paragraph(target, "Destination list item two")
    src = source()
    first = _paragraph(src, "Source list item one")
    last = _paragraph(src, "Source list item two")
    target.copy_blocks(src, f"{first.id}..{last.id}", at=f"after:{item.id}", lists="continue")
    labels = [b["list"]["label"] for b in target.state()["blocks"] if "list item" in b.get("t", "")]
    assert labels == ["1.", "2.", "3.", "4."]
    assert len(etree.fromstring(parts(target.to_bytes())["word/numbering.xml"]).findall(W + "num")) == 1


def test_media_is_shared_by_content_and_drawing_ids_reissued():
    """Word: both pictures one image1.png; the copy's docPr id re-issued."""
    word = OBSERVATIONS["copy-formatted"]
    assert list(word["media"]) == ["word/media/image1.png"]
    target, _ = copied()
    assert [n for n in target.package.part_names if n.startswith("word/media/")] == ["word/media/image1.png"]
    ids = [int(n.get("id")) for n in etree.fromstring(parts(target.to_bytes())["word/document.xml"]).iter(
        WP + "docPr")]
    assert len(ids) == len(set(ids)) == 2 and 1 in ids


def test_notes_and_comments_come_along_renumbered():
    word = OBSERVATIONS["copy-formatted"]
    assert '<w:footnote w:id="2">' in word["word/footnotes.xml"] and "Source footnote." in word["word/footnotes.xml"]
    assert "Source comment" in word["word/comments.xml"]
    target, result = copied()
    assert result.copied["fn:1"] == "fn:2"
    footnotes = etree.fromstring(parts(target.to_bytes())["word/footnotes.xml"])
    assert "Source footnote." in "".join(footnotes.itertext())
    comments = target.comments()
    assert sorted(c.text for c in comments) == ["Destination comment", "Source comment"]
    ids = {c.id for c in comments}
    assert len(ids) == 2
    assert "word/commentsExtended.xml" in target.package.part_names
    assert check(target.package) == []


def test_bookmarks_are_renamed_or_dropped_as_word():
    """Word dropped the copy's colliding bookmark; ``rename`` (the default) keeps it under a
    new name, ``drop`` does as Word."""
    word = OBSERVATIONS["copy-formatted"]["body"]
    assert sum(block.count('w:name="Shared"') for block in word) == 1
    target, result = copied()
    names = sorted(b.name for b in target.bookmarks())
    assert names == ["Shared", "Shared_1"]
    assert any("Shared_1" in w for w in result.warnings)
    dropped, _ = copied(bookmarks="drop")
    assert sorted(b.name for b in dropped.bookmarks()) == ["Shared"]
    assert check(target.package) == [] and check(dropped.package) == []


def test_a_renamed_bookmarks_links_follow():
    src = Document.new()
    first = src.paragraphs()[0]
    first.set_text("Target text here.")
    src.paragraph(first.id).range().add_bookmark("Mark")
    link = src.insert_paragraph("See the target.", after=first.id)
    src.anchor("target", within=link.id).add_hyperlink(anchor="Mark")
    reference = src.insert_paragraph("Page ", after=link.id)
    src.insert_cross_reference(f"{reference.id}@5", "Mark", kind="page")
    target = Document.new()
    target.paragraphs()[0].set_text("Mine.")
    target.paragraph(target.paragraphs()[0].id).range().add_bookmark("Mark")
    target.copy_blocks(src, everything(src))
    xml = parts(target.to_bytes())["word/document.xml"].decode()
    assert 'w:anchor="Mark_1"' in xml and "PAGEREF Mark_1" in xml and xml.count('w:name="Mark"') == 1


def test_ids_are_kept_unless_they_collide():
    """Word re-issued every paraId in the document; docx-agent only the copies' that collide,
    and the content control's id (as Word)."""
    target, result = copied()
    before = {p.id for p in destination().paragraphs()}
    assert before <= {p.id for p in target.paragraphs()}
    assert result.copied["p:7B000001"] != "p:7B000001"                 # collided
    assert result.copied["p:7B00000D"] == "p:7B00000D"                 # did not
    assert result.copied["t:7B00000A"] == "t:7B00000A"
    assert "sdt#5" in result.copied and result.copied["sdt#5"] != "sdt#5"
    raw = [p.get(W14 + "paraId") for p in etree.fromstring(parts(target.to_bytes())["word/document.xml"]).iter(W + "p")]
    assert len(raw) == len(set(raw))
    assert result.blocks[0] == result.id and len(result.blocks) == 11


def test_tracked_as_word_tracks_a_paste():
    """Word's tracked copy: every mark and run an insertion, the row too."""
    word = "".join(OBSERVATIONS["copy-formatted-tracked"]["body"])
    assert word.count("<w:ins ") > 20 and "<w:trPr><w:ins " in word
    target = destination()
    target._stamp_document()
    start = target.to_bytes()
    untracked = Document.open(start)
    untracked.copy_blocks(source(), everything(source()))
    expected = canonical(untracked.to_bytes(), ids=False)
    tracked = Document.open(start)
    with tracked.tracking(author=AUTHOR, date=DATE):
        tracked.copy_blocks(source(), everything(source()))
    data = tracked.to_bytes()
    assert check(Document.open(data).package) == []
    assert len(tracked.revisions(author=AUTHOR)) > 10
    accepted = Document.open(data)
    accepted.accept(author=AUTHOR)
    got = canonical(accepted.to_bytes(), ids=False)
    assert got == expected, difference(expected, got)
    rejected = Document.open(data)
    rejected.reject(author=AUTHOR)
    original = canonical(start, ids=False)
    got = without_added_styles(canonical(rejected.to_bytes(), ids=False), original)
    # The copy's notes, comments, list and media stay, unreferenced, as an insertion's do.
    got = {k: v for k, v in got.items() if k in original}
    assert {k: v for k, v in got.items() if not k.startswith(("word/footnotes", "word/comments", "word/numbering",
                                                             "[Content", "word/_rels", "word/people"))} == \
        {k: v for k, v in original.items() if not k.startswith(("word/footnotes", "word/comments", "word/numbering",
                                                               "[Content", "word/_rels", "word/people"))}


@pytest.mark.parametrize("policy", ["use_destination", "keep_source", "merge"])
def test_one_undo_step_and_saved_read_back(policy, tmp_path):
    target = destination()
    original = target.to_bytes()
    result = target.copy_blocks(source(), everything(source()), styles=policy)
    edited = target.to_bytes()
    target.save(tmp_path / "copied.docx")
    saved = Document.open(tmp_path / "copied.docx")
    texts = [p.text for p in saved.paragraphs()]
    for paragraph in source().paragraphs():
        assert paragraph.text in texts
    assert [saved.get(i) is not None for i in result.blocks].count(True) == len(result.blocks)
    assert check(saved.package) == []
    assert target.undo() and entries(target.to_bytes()) == entries(original)
    assert target.redo() and entries(target.to_bytes()) == entries(edited)


def test_places_and_refusals():
    src = source()
    one = _paragraph(src, "Source heading").id
    target = destination()
    anchor = _paragraph(target, "Destination heading").id
    made = target.copy_blocks(src, one, at=f"before:{anchor}")
    assert target.paragraphs()[0].text == "Source heading" and made.blocks == [target.paragraphs()[0].id]
    made = target.copy_blocks(src, [one, _paragraph(src, "Source last").id], at=f"replace:{anchor}")
    assert "Destination heading" not in [p.text for p in target.paragraphs()]
    assert copy_blocks(src, one, to=target, at="end").changed
    with pytest.raises(Exception):
        target.copy_blocks(target, anchor)
    with pytest.raises(Exception):
        target.copy_blocks(src, one, styles="whatever")
    with pytest.raises(Exception):
        target.copy_blocks(src, [])


def test_the_copy_lays_out():
    target, result = copied()
    layout = target.layout()
    for identifier in result.blocks:
        assert layout.where(identifier)


# -- every fixture, both ways ---------------------------------------------------------------------

SOURCES = fixture_paths()


def _top_level(document: Document) -> str | None:
    body = document.package.tree(document.package.document_part()).find(W + "body")
    blocks = [c for c in body if c.tag in (W + "p", W + "tbl")]
    ids = []
    for element in (blocks[0], blocks[-1]) if blocks else ():
        entry = next((e for e in document._index(document.package.document_part()).paragraphs
                      if e.element is element), None) or next(
            (t for t in document._index(document.package.document_part()).tables if t.element is element), None)
        if entry is None:
            return None
        ids.append(entry.id)
    return f"{ids[0]}..{ids[1]}" if ids else None


@pytest.mark.parametrize("policy", ["use_destination", "keep_source"])
@pytest.mark.parametrize("path", SOURCES, ids=fixture_id)
def test_every_fixture_copies_into_a_new_document_and_another(path, policy):
    src = Document.open(path.read_bytes())
    span = _top_level(src)
    if span is None:
        pytest.skip("the body has no paragraph or table")
    texts = [p.text for p in src.paragraphs() if p.text.strip()]
    for target in (Document.new(), Document.open(HERE / "fixtures" / "generated" / "lists-and-styles.docx")):
        before = set(check(target.package))
        original = target.to_bytes()
        result = target.copy_blocks(src, span, styles=policy)
        assert set(check(target.package)) - before == set(), path.name
        got = [p.text for p in target.paragraphs()]
        for text in texts:
            assert text in got
        assert target.layout().pages
        assert target.undo() and entries(target.to_bytes()) == entries(original)
        assert result.blocks


def _spans(document: Document, identifier: str) -> list[tuple]:
    layout = document.layout()
    wanted = layout._canon(identifier)
    out = []
    for page in layout._lines:
        for key, line in page:
            if key and layout._canon(key) == wanted:
                out += [("".join(s.chars), s.face, s.bold, s.italic, s.half_points, s.color, s.underline)
                        for s in line.spans]
    return out


def test_copied_blocks_keep_their_look_in_docx2svg():
    """Keep Source Formatting: docx2svg draws each copied paragraph in the faces, sizes,
    weights and colours it draws the source's in."""
    src = source()
    target, result = copied("keep_source")
    for paragraph in src.paragraphs():
        if not paragraph.text.startswith(("Source heading", "Source brand", "Source list", "Source only",
                                          "Source last")):
            continue
        copy = result.copied[paragraph.id]
        assert _spans(target, copy) == _spans(src, paragraph.id), paragraph.text


def test_use_destination_takes_the_destinations_look_in_docx2svg():
    """Use Destination Styles: a style the destination has gives the copy its look there."""
    target, result = copied()
    mine = _paragraph(target, "Destination brand")
    theirs = _paragraph(target, "Source brand")
    looks = lambda spans: {(face, bold, italic, size, color) for _, face, bold, italic, size, color, _ in spans}  # noqa: E731
    assert looks(_spans(target, theirs.id)) <= looks(_spans(target, mine.id))
