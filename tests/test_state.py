"""The structured, read-only JSON state: its schema, its ids, and what it says about
styles, formatting, lists, tables, pictures, links, revisions, comments and layout."""

from __future__ import annotations

import base64
import json

from lxml import etree

from docx_agent import Document
from docx_agent.state import SCHEMA, VERSION

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def blocks(state: dict):
    """Every block, depth first: cells', controls' and text boxes' too."""
    def walk(items):
        for item in items:
            yield item
            for line in item.get("rows", []):
                for cell in line:
                    yield from walk(cell["blocks"])
            yield from walk(item.get("blocks", []))
            for box in item.get("text_boxes", []):
                yield from walk(box["blocks"])
    yield from walk(state["blocks"])
    for note in state["notes"]:
        yield from walk(note["blocks"])


def ids(state: dict) -> set[str]:
    out: set[str] = set()
    for item in blocks(state):
        out.add(item["id"])
        out.update(item.get("joins", []))
        for key in ("bookmarks", "notes", "comments", "revisions"):
            out.update(item.get(key, []))
        for key in ("links", "pictures", "drawings", "controls"):
            out.update(x["id"] for x in item.get(key, []))
        for line in item.get("rows", []):
            out.update(cell["id"] for cell in line)
        if item.get("section"):
            out.add(item["section"]["id"])
    out.update(n["id"] for n in state["notes"])
    out.update(c["id"] for c in state["comments"])
    out.update(r["id"] for r in state["revisions"])
    out.update(s["id"] for s in state["document"]["sections"])
    return out


def test_every_id_in_the_state_resolves(reading_path):
    document = Document.open(reading_path)
    for view in ("final", "original", "markup"):
        state = document.state(view=view)
        assert state["schema"] == SCHEMA and state["version"] == VERSION == 1 and state["view"] == view
        json.dumps(state)
        for identifier in ids(state):
            document.get(identifier)


def test_paragraphs_styles_and_effective_formatting(dutch_doc):
    state = Document.open(dutch_doc).state()
    heading = next(b for b in state["blocks"] if b["id"] == "p:5E000002")
    assert (heading["style"], heading["markdown"], heading["level"]) == ("Kop 1", "heading", 1)
    assert heading["effective"]["sz"] == 16.0 and heading["effective"]["b"] is True
    text = next(b for b in state["blocks"] if b["id"] == "p:5E000003")
    assert text["t"] == "Tekst met nadruk en zwaar."
    assert [r for r in text["runs"] if r["t"] == "zwaar"] == [{"t": "zwaar", "style": "Strong", "b": True}]
    assert [r for r in text["runs"] if r["t"] == "nadruk"] == [{"t": "nadruk", "style": "Emphasis", "i": True}]
    # Styles by name, with their localised ids.
    assert state["styles"]["Kop 1"] == {"id": "Kop1", "kind": "paragraph", "based_on": "Normal"}
    assert state["styles"]["Strong"]["id"] == "Zwaar"
    assert state["styles"]["Normal"]["id"] == "Standaard"


def test_direct_formatting_is_declared_and_effective(constructs_doc):
    state = Document.open(constructs_doc).state()
    plain = next(b for b in state["blocks"] if b["id"] == "p:6B000003")
    bold = next(r for r in plain["runs"] if r["t"] == "bold")
    assert bold == {"t": "bold", "direct": ["b"], "b": True}
    code = next(r for r in plain["runs"] if r["t"] == "inline code")
    assert code["style"] == "HTML Code" and code["font"] == "Courier New" and code["sz"] == 10.0


def test_lists(constructs_doc):
    state = Document.open(constructs_doc).state()
    five = next(b for b in state["blocks"] if b["id"] == "p:6B00001E")
    assert five["markdown"] == "ordered" and five["level"] == 1
    assert five["list"] == {"num_id": 4, "level": 0, "format": "decimal", "label": "5.", "number": 5,
                            "from_style": False}
    sub = next(b for b in state["blocks"] if b["id"] == "p:6B00001B")
    assert sub["list"]["label"] == "a." and sub["level"] == 2
    styled = next(b for b in state["blocks"] if b["id"] == "p:6B000020")
    assert styled["list"]["from_style"] is True and styled["markdown"] == "bullet"


def test_tables_with_spans(markup_doc):
    state = Document.open(markup_doc).state()
    table = next(b for b in state["blocks"] if b["type"] == "table")
    assert (table["id"], table["columns"], len(table["rows"])) == ("t:3C4D5E01", 3, 2)
    first = table["rows"][0]
    assert [(c["id"], c["row_span"], c["column_span"]) for c in first] == [
        ("t:3C4D5E01/c0,0", 2, 1), ("t:3C4D5E01/c0,1", 1, 2)]
    assert [c["id"] for c in table["rows"][1]] == ["t:3C4D5E01/c1,1", "t:3C4D5E01/c1,2"]
    assert first[0]["blocks"][0]["t"] == "Merged down"


def test_nested_tables_and_controls(constructs_doc):
    state = Document.open(constructs_doc).state()
    outer = next(b for b in state["blocks"] if b["id"] == "t:6B000033")
    nested = outer["rows"][0][1]["blocks"][0]
    assert nested["type"] == "table" and nested["rows"][0][0]["blocks"][0]["t"] == "Inner"
    control = next(b for b in state["blocks"] if b["type"] == "content_control")
    assert (control["id"], control["kind"], control["tag"], control["alias"]) == ("cc:303", "rich-text", "clause",
                                                                                   "Clause")
    assert [b["t"] for b in control["blocks"]] == ["A clause in a control.", "Its second paragraph."]
    inline = next(b for b in state["blocks"] if b["id"] == "p:6B00003E")
    assert inline["controls"] == [{"id": "cc:301", "kind": "text"}, {"id": "cc:302", "kind": "checkbox"}]


def test_links_pictures_notes_and_fields(constructs_doc):
    state = Document.open(constructs_doc).state()
    links = next(b for b in state["blocks"] if b["id"] == "p:6B000022")
    assert links["links"] == [
        {"id": "hl:p:6B000022/0", "href": "https://example.com/markdown", "t": "external link"},
        {"id": "hl:p:6B000022/1", "href": "#notes", "t": "internal one"}]
    assert links["notes"] == ["fn:1", "en:1"]
    picture = next(b for b in state["blocks"] if b["id"] == "p:6B000025")["pictures"][0]
    assert picture == {"id": "d:2", "name": "Floating", "alt": "A red square", "floating": True, "size": [36.0, 36.0]}
    box = next(b for b in state["blocks"] if b["id"] == "p:6B000026")
    assert box["drawings"][0]["kind"] == "text-box" and box["text_boxes"][0]["blocks"][0]["t"] == "Inside a text box."
    fields = next(b for b in state["blocks"] if b["id"] == "p:6B000027")["fields"]
    assert fields == [{"instruction": "PAGE", "result": "1"},
                      {"instruction": "REF notes \\h", "result": "Notes are written at the end."}]
    assert [n["id"] for n in state["notes"]] == ["fn:1", "en:1", "fn:2"]
    assert [b["t"] for b in state["notes"][2]["blocks"]] == [" A second footnote,", "in two paragraphs."]
    section = next(b for b in state["blocks"] if b["id"] == "p:6B000041")["section"]
    assert section == {"id": "s:6B000041", "type": "continuous"}
    assert state["document"]["sections"][1]["headers"] == {"default": "header1"}


def test_revisions_and_comments(review_doc):
    document = Document.open(review_doc)
    final, original = document.state(), document.state(view="original")
    joined = next(b for b in final["blocks"] if b["id"] == "p:7C000006")
    assert joined["joins"] == ["p:7C000007"] and joined["t"].endswith("joins the next one in the final view.")
    kinds = {r["id"]: (r["kind"], r["author"]) for r in final["revisions"]}
    assert kinds["rev:101"] == ("deletion", "Alice") and kinds["rev:102"] == ("insertion", "Alice")
    assert kinds["rev:105"] == ("move-from", "Alice") and kinds["rev:108"] == ("paragraph-mark-deletion", "Bob")
    assert kinds["rev:113"] == ("run-properties", "Bob") and kinds["rev:114"] == ("paragraph-properties", "Alice")
    assert kinds["rev:118"] == ("row-deletion", "Bob") and kinds["rev:121"] == ("row-insertion", "Alice")
    runs = next(b for b in final["blocks"] if b["id"] == "p:7C000002")["runs"]
    assert {"t": "twelve", "rev": "rev:102"} in runs
    assert {"t": "ten", "rev": "rev:101"} in next(b for b in original["blocks"] if b["id"] == "p:7C000002")["runs"]
    comments = {c["id"]: c for c in final["comments"]}
    assert comments["c:2B3C4D5E"]["parent"] == "c:1A2B3C4D" and comments["c:2B3C4D5E"]["t"] == "Yes, it is."
    assert comments["c#3"]["done"] is True and comments["c#3"]["paragraphs"] == ["p:7C00000E"]
    table = next(b for b in final["blocks"] if b["type"] == "table")
    assert [[c["blocks"][0]["t"] for c in line] for line in table["rows"]] == [["Name", "Score"], ["Ben", "4"]]
    assert table["header_rows"] == 1


def test_layout_placements_and_raw_xml(markup_doc):
    state = Document.open(markup_doc).state(layout=True, xml=True)
    assert state["layout"] == {"page_count": 2, "pages_known": 2, "stopped": None}
    last = next(b for b in state["blocks"] if b["id"] == "p:1A2B3C17")
    assert last["placements"][0]["page"] == 2 and last["placements"][0]["story"] == "body"
    table = next(b for b in state["blocks"] if b["type"] == "table")
    assert table["placements"][0]["page"] == 1
    continuation = table["rows"][1]
    assert all("placements" in cell["blocks"][0] or "unknown" in cell["blocks"][0] for cell in continuation)
    element = etree.fromstring(base64.b64decode(last["xml"][0]))
    assert element.tag == W + "p" and "".join(element.itertext()) == "The last paragraph."
    assert "xml" not in Document.open(markup_doc).state()["blocks"][0]


def test_a_range(constructs_doc):
    state = Document.open(constructs_doc).state("p:6B000022..p:6B000023")
    assert [b["id"] for b in state["blocks"]] == ["p:6B000022", "p:6B000023"]
    assert [n["id"] for n in state["notes"]] == ["fn:1", "en:1", "fn:2"]
    assert state["range"] == "p:6B000022..p:6B000023"


def test_comments_carry_the_text_they_are_attached_to(review_doc):
    comments = {c["id"]: c for c in Document.open(review_doc).state()["comments"]}
    assert comments["c:1A2B3C4D"]["anchor"] == {"id": "p:7C00000D@2:10", "t": "disputed"}
    # A reply with only a reference is attached where its thread is.
    assert comments["c:2B3C4D5E"]["anchor"] == comments["c:1A2B3C4D"]["anchor"]
