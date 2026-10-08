"""E2's reading gates, held on every fixture in every view: ``to_markdown`` and ``state``
change no byte; the Markdown parses, reads back as the model it was written from, and
names only ids that resolve; the final view's text is what docx2svg draws."""

from __future__ import annotations

import json
import re

import pytest
from lxml import etree

from docx_agent import Document
from docx_agent.markdown import check, parse_ast
from docx_agent.markdown.read import ParagraphRecord, Reader, TableRecord, walk_records

from test_roundtrip import entries

VIEWS = ("final", "original", "markup")

#: Every id form the projection writes (ROADMAP.md, "Addressing").
ID = re.compile(r"(?<![\w:@#/-])((?:[A-Za-z0-9]+/)?(?:p|t|tr|s)[:@][^\s>\"]+|(?:cc|d|bm|fn|en|c|rev|hl)[:@#][^\s>\"\])]+)")
COMMENT = re.compile(r"<!--(.*?)-->", re.DOTALL)
STORY = re.compile(r"<!-- story: (\S+)")


def ids_in(markdown: str) -> set[str]:
    """Every id the Markdown names: in comments, ``data-id`` attributes, image sources and
    footnote labels."""
    found: set[str] = set()
    for body in COMMENT.findall(markdown):
        body = re.sub(r"(?:table )?style: .*", "", body)  # a style's name is not an id
        found.update(ID.findall(body))
    found.update(re.findall(r'data-id="([^"]+)"', markdown))
    found.update(re.findall(r"!\[[^\]]*\]\((d:[^\s)]+)", markdown))
    found.update(re.findall(r"\[\^([^\]]+)\]", markdown))
    found.update(STORY.findall(markdown))
    return found


def snapshot(document: Document) -> dict[str, bytes]:
    """Every XML part as its tree serialises now: a reader that changed a tree without
    marking it dirty would show here, not in ``to_bytes``."""
    out = {}
    for name in document.package.part_names:
        if name.endswith((".xml", ".rels")):
            tree = document.package.tree(name)
            if tree is not None:
                out[name] = etree.tostring(tree)
    return out


def test_reading_changes_no_byte(reading_path):
    data = reading_path.read_bytes()
    document = Document.open(data)
    before = snapshot(document)
    for view in VIEWS:
        document.to_markdown(view=view, headers=True)
        document.to_markdown(view=view, ids=False)
        json.dumps(document.state(view=view, xml=True))
    json.dumps(document.state(layout=True))
    assert snapshot(document) == before
    assert document.package.changed_parts() == frozenset()
    assert not document.history.can_undo()
    assert document.aliases == {}
    assert entries(document.to_bytes()) == entries(data)


@pytest.mark.parametrize("view", VIEWS)
def test_markdown_reads_back_as_written(reading_path, view):
    document = Document.open(reading_path)
    assert check(document, view=view, headers=True) == []
    assert check(document, view=view, ids=False) == []


@pytest.mark.parametrize("view", VIEWS)
def test_every_id_resolves(reading_path, view):
    document = Document.open(reading_path)
    markdown = document.to_markdown(view=view, headers=True)
    parse_ast(markdown)
    ids = ids_in(markdown)
    for identifier in sorted(ids):
        document.get(identifier)  # raises KeyError for an id that does not resolve
    # Every block of the body is named: its paragraphs (or the paragraph a view joined them
    # into) and tables.  A GFM table's cells are named by the cell ids its id implies.
    top = Reader(document, view).story(document.package.document_part())
    named = {r.id for r in top if isinstance(r, (ParagraphRecord, TableRecord))}
    assert named - ids == set()
    for record in top:
        if isinstance(record, ParagraphRecord) and record.joins:
            assert f"{record.id} joins {' '.join(record.joins)}" in markdown


def test_ids_off_writes_no_comment(reading_path):
    markdown = Document.open(reading_path).to_markdown(view="markup", ids=False)
    assert "<!--" not in markdown and "data-id" not in markdown


def test_volatile_ids_are_marked(markup_doc):
    markdown = Document.open(markup_doc).to_markdown()
    assert "<!-- p@body/2 volatile -->" in markdown
    assert "<!-- p:2B3C4D05#1 volatile -->" in markdown
    assert "<!-- p:2B3C4D05 -->" in markdown


# -- the final view is what docx2svg draws ---------------------------------------------------

#: Fields docx2svg computes for its drawing; Markdown shows the result Word cached.
_COMPUTED = re.compile(r"^(PAGE|NUMPAGES|SECTIONPAGES)\b")


def _normal(text: str) -> str:
    """Characters both sides draw: no whitespace (a tab is a jump, a line's end a break),
    no soft hyphen or object mark, a non-breaking hyphen as the hyphen docx2svg draws, and
    no Symbol-font private-use character (drawn by docx2svg as a glyph of that font)."""
    text = text.replace("‑", "-")
    return re.sub(r"[\s­￼-]", "", text)


def _expected(record: ParagraphRecord) -> re.Pattern:
    """The reader's text with each note reference a mark docx2svg numbers, and a note's own
    paragraph's leading mark."""
    from docx_agent.markdown import model as m

    pieces = [re.escape(_normal(i.text)) if isinstance(i, m.Text) else "[0-9ivxlcdm*]+"
              for i in record.inline if isinstance(i, (m.Text, m.NoteRef))]
    lead = "(?:[0-9ivxlcdm*]+)?" if record.story in ("footnotes", "endnotes") else ""
    return re.compile(lead + "".join(pieces))


def test_final_view_is_what_docx2svg_draws(reading_path):
    document = Document.open(reading_path)
    layout = document.layout()
    drawn: dict[str, list[str]] = {}
    for page in layout._lines:
        for identifier, line in page:
            if identifier is not None:
                drawn.setdefault(identifier, []).append(
                    "".join(c for span in line.spans if span.kind == "text" for c in span.chars))
    reader = Reader(document, "final")
    compared = 0
    for part in document._parts():
        if document._story_of(part).startswith(("header", "footer")):
            continue  # drawn once per page, PAGE computed: held by the body's rule below
        for record in walk_records(reader.story(part)):
            if not isinstance(record, ParagraphRecord) or record.id not in drawn:
                continue
            if any(_COMPUTED.match(f["instruction"]) for f in record.fields):
                continue
            assert _expected(record).fullmatch(_normal("".join(drawn[record.id]))), \
                (record.id, record.text, "".join(drawn[record.id]))
            compared += 1
    assert compared or not any(drawn)


def test_an_edited_document_reads_back(docx_path):
    """After E1's edit set -- stamped ids, a new heading, lists restarted at 7, links, a
    cross-reference, direct formatting, a picture -- the projection still reads back as
    written, and names only ids that resolve."""
    from e1_edits import e1_edit_set

    document = Document.open(docx_path)
    e1_edit_set(document)
    for view in VIEWS:
        assert check(document, view=view, headers=True) == []
    markdown = document.to_markdown(headers=True)
    for identifier in ids_in(markdown):
        document.get(identifier)
    assert "## E1 Heading Marker" in markdown
    assert "\n\n7) E1 gamma item <!-- " in markdown  # the restart: a list of its own, at 7
