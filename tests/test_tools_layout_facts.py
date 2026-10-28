"""check's layout facts for documents (round 16, F8): measured from the layout reflow uses.

Every document here is made up: Document.new()'s package with a body written for the case,
its paragraphs given paraIds so that their ids stay put across the fixes the facts offer.
The cases hold whatever faces the machine has: a row taller than a page splits, a
keep-with-next paragraph before a page break stays at its page's foot, a picture wider
than the page runs past the margin, on any metrics.  Word for Mac 16.106 drew made-up
probes of the keep-with-next and row cases as docx2svg does (CHANGELOG, round 16, F8).
"""

from __future__ import annotations

import io
import itertools
import struct
import zipfile
import zlib

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

_ids = itertools.count(0x100)


def p(text: str = "", rpr: str = "", ppr: str = "", runs: str | None = None) -> str:
    """A paragraph with a paraId of its own: one run of ``text`` (or ``runs`` as given)."""
    if runs is None:
        runs = (f'<w:r>{f"<w:rPr>{rpr}</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'
                if text else "")
    return (f'<w:p w14:paraId="{next(_ids):08X}" w14:textId="77777777">'
            f'{f"<w:pPr>{ppr}</w:pPr>" if ppr else ""}{runs}</w:p>')


def table(rows: list[list[str]], *, width: int = 9000, header: bool = False, indent: int | None = None) -> str:
    """A bordered table of ``rows`` of cell texts; ``header``: its first row repeats."""
    cols = len(rows[0])
    grid = f'<w:gridCol w:w="{width // cols}"/>' * cols
    borders = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    indented = f'<w:tblInd w:w="{indent}" w:type="dxa"/>' if indent is not None else ""
    out = (f'<w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/>{indented}<w:tblBorders>{borders}'
           f'</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>')
    for k, row in enumerate(rows):
        properties = "<w:trPr><w:tblHeader/></w:trPr>" if header and k == 0 else ""
        cells = "".join(f'<w:tc><w:tcPr><w:tcW w:w="{width // cols}" w:type="dxa"/></w:tcPr>{p(text)}</w:tc>'
                        for text in row)
        out += f'<w:tr w14:paraId="{next(_ids):08X}" w14:textId="77777777">{properties}{cells}</w:tr>'
    return out + "</w:tbl>"


def docx(body: str) -> bytes:
    """Document.new()'s package, its body ``body`` (the section kept)."""
    source = zipfile.ZipFile(io.BytesIO(Document.new().to_bytes()))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in source.namelist():
            data = source.read(name)
            if name == "word/document.xml":
                text = data.decode()
                start = text.index("<w:body>") + len("<w:body>")
                text = text[:start] + body + text[text.index("<w:sectPr", start):]
                data = text.encode()
            archive.writestr(name, data)
    return out.getvalue()


def png(width: int = 40, height: int = 20) -> bytes:
    def chunk(kind: bytes, body: bytes) -> bytes:
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body))

    rows = b"".join(b"\x00" + b"\x80\x80\x80" * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


FILLER = "Filler text for the made-up page, long enough to wrap onto a second line of the column."


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


@pytest.fixture
def session(box):
    return box.session()


def ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, result.error and result.error.to_json())
    return result


def facts(box, session, doc: str, **arguments) -> dict:
    return ok(box, session, "check", doc=doc, include=["layout"], **arguments).data["layout"]


def call(box, session, doc: str, fix: dict):
    """Make the call a fact's ``fix`` names (each of its calls)."""
    for arguments in fix["calls"]:
        ok(box, session, fix["tool"], doc=doc, **arguments)


# -- text sizes, direct formatting, styles ----------------------------------------------------


def test_small_text_direct_formatting_and_styles_with_their_fixes(box, session):
    small, titled, emphasis = p("Small print", '<w:sz w:val="12"/><w:szCs w:val="12"/>'), \
        p("A title", '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>', '<w:pStyle w:val="Heading1"/><w:jc w:val="center"/>'), \
        p(runs='<w:r><w:t xml:space="preserve">Some </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>')
    d = session.open(docx(titled + small + emphasis + table([["a", "b"]]) + p("End")), "made.docx")
    layout = facts(box, session, d, min_size=9, styles=["Normal", "Heading 1", "No such style"])
    ids = [paragraph.id for paragraph in session.entry(d).document.paragraphs()]
    title, small_id, emphasis_id = ids[:3]

    assert layout["smallest_text"]["body"] == 6.0 and "table" in layout["smallest_text"]
    [entry] = layout["small_text"]
    assert entry["para"] == small_id and entry["kind"] == "body" and entry["size"] == 6.0
    assert entry["pages"] == [1] and entry["style_size"] > 9
    assert entry["fix"] == {"tool": "word_format", "calls": [{"targets": [small_id], "clear": ["size"]}]}

    found = {e["para"]: e for e in layout["direct_formatting"]}
    assert found[title]["style"] == "heading 1"
    assert found[title]["props"]["alignment"] == "center" and found[title]["style_has"]["alignment"] == "left"
    assert found[title]["props"]["font"] == "Arial" and found[title]["style_has"]["font"] != "Arial"
    assert set(found[title]["fix"]["calls"][0]["clear"]) == {"alignment", "font"}
    assert found[small_id]["props"] == {"size": 6.0}
    # Bold on one word of a paragraph is named, not cleared.
    assert found[emphasis_id]["in_part"] == ["bold"] and "fix" not in found[emphasis_id]
    assert layout["direct_formatting"][0]["para"] == title          # the most properties first

    [unlisted] = layout["styles_not_allowed"]
    assert unlisted["style"] == "Normal Table" and unlisted["kind"] == "table" and unlisted["count"] == 1
    assert layout["styles_unknown"] == ["No such style"]

    call(box, session, d, found[title]["fix"])
    call(box, session, d, entry["fix"])
    after = facts(box, session, d, min_size=9)
    assert after["small_text"] == []
    assert [e["para"] for e in after["direct_formatting"]] == [emphasis_id]
    inspected = ok(box, session, "word_inspect", doc=d, target=title).data["blocks"][0]
    assert inspected["effective"]["alignment"] == "left"


def test_word_format_clear_takes_properties_back_to_the_style(box, session):
    target = p("Indented", '<w:sz w:val="40"/><w:b/>', '<w:ind w:left="720" w:hanging="360"/><w:spacing w:after="600"/>')
    d = session.open(docx(target + p("End")), "clear.docx")
    paragraph = session.entry(d).document.paragraphs()[0].id
    ok(box, session, "word_format", doc=d, targets=[paragraph], clear=["indent_first", "size", "space_after"])
    xml = session.entry(d).document.package.tree(session.entry(d).document.package.document_part())
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    element = xml.find(f"{w}body/{w}p")
    assert element.find(f"{w}pPr/{w}ind").get(w + "left") == "720"
    assert element.find(f"{w}pPr/{w}ind").get(w + "hanging") is None
    assert element.find(f"{w}pPr/{w}spacing") is None
    assert element.find(f"{w}r/{w}rPr/{w}sz") is None and element.find(f"{w}r/{w}rPr/{w}b") is not None
    refused = box.dispatch(session, "word_format", {"doc": d, "targets": [paragraph], "size": 12, "clear": ["size"]})
    assert not refused.ok and refused.error.field == "clear"


# -- spacers ----------------------------------------------------------------------------------


def test_empty_paragraphs_used_as_space_and_the_space_kept_without_them(box, session):
    body = (p(FILLER) + p() + p() + p() + p("After the gap") + table([["a"]]) + p() + table([["b"]])
            + p(FILLER) + p())
    d = session.open(docx(body), "gaps.docx")
    document = session.entry(d).document
    ids = [paragraph.id for paragraph in document.paragraphs()]
    after_gap = ids[4]
    before = document.layout().where(after_gap)[0].top
    layout = facts(box, session, d)
    [gap] = layout["spacers"]          # not the one between the tables, nor the last paragraph
    assert gap["paras"] == ids[1:4] and gap["page"] == 1 and gap["next"] == after_gap
    assert gap["height_pt"] > 20
    assert gap["fix"] == {"tool": "word_delete", "calls": [{"target": f"{ids[1]}..{ids[3]}"}]}
    call(box, session, d, gap["fix"])
    call(box, session, d, gap["keep_space"])
    assert facts(box, session, d)["spacers"] == []
    moved = session.entry(d).document.layout().where(after_gap)[0].top
    assert abs(moved - before) < 1.0, (before, moved)


# -- keep with next ---------------------------------------------------------------------------


def test_a_keep_with_next_paragraph_left_at_its_pages_foot(box, session):
    # Word keeps the heading on page 1: the paragraph it keeps with starts a page of its own.
    heading = p("Heading kept with the next", ppr="<w:keepNext/>")
    forced = p("On a page of its own", ppr="<w:pageBreakBefore/>")
    d = session.open(docx(p(FILLER) + heading + forced + p(FILLER)), "kept.docx")
    ids = [paragraph.id for paragraph in session.entry(d).document.paragraphs()]
    [entry] = facts(box, session, d)["keep_with_next_at_foot"]
    assert entry == {"para": ids[1], "style": "Normal", "page": 1, "next": ids[2], "next_page": 2}
    # A plain paragraph before the break is no fact.
    plain = session.open(docx(p(FILLER) + p("Not kept") + p("On a page of its own", ppr="<w:pageBreakBefore/>")),
                         "plain.docx")
    assert facts(box, session, plain)["keep_with_next_at_foot"] == []


def test_a_chain_of_keep_with_next_paragraphs_longer_than_a_page(box, session):
    body = "".join(p(f"Kept line {k}", ppr="<w:keepNext/>") for k in range(90)) + p("The end")
    d = session.open(docx(body), "chain.docx")
    entries = facts(box, session, d)["keep_with_next_at_foot"]
    assert entries and all(e["next_page"] == e["page"] + 1 for e in entries)
    assert [e["page"] for e in entries] == list(range(1, len(entries) + 1))


# -- tables -----------------------------------------------------------------------------------


def test_a_split_row_and_header_rows_that_do_not_repeat(box, session):
    tall = table([["Head", "Head"]] + [[f"Row {r}", (FILLER + " ") * 40 if r == 3 else "short"] for r in range(8)],
                 header=True)
    plain = table([[f"Item {r}", "plain"] for r in range(80)])
    d = session.open(docx(p("Tables") + tall + p("Between") + plain + p("End")), "rows.docx")
    document = session.entry(d).document
    first, second = [t.id for t in document.tables()]
    layout = facts(box, session, d)
    [split] = [e for e in layout["split_rows"] if e["table"] == first]
    assert split["row"] == 4 and len(split["pages"]) >= 2          # the header row comes first
    # The header row is drawn again on every page the tall table goes on to: not a split.
    assert all(e["row"] != 0 for e in layout["split_rows"])
    [unrepeated] = layout["header_not_repeated"]
    assert unrepeated["table"] == second and unrepeated["header_rows"] == 0 and len(unrepeated["pages"]) >= 2
    call(box, session, d, unrepeated["fix"])
    assert facts(box, session, d)["header_not_repeated"] == []


# -- margins ----------------------------------------------------------------------------------


def test_a_picture_and_a_table_past_the_margins(box, session):
    d = session.open(docx(p("Before") + table([["wide", "table"]], width=12000) + p("Anchor") + p("End")), "wide.docx")
    document = session.entry(d).document
    anchor = document.paragraphs()[3].id
    image = session.add_blob(png(), "wide.png")
    made = ok(box, session, "word_drawings", doc=d, action="insert_picture", image=image, at=f"after:{anchor}",
              width=600, alt_text="A wide made-up picture")
    picture = made.data["id"]
    layout = facts(box, session, d)
    past = {e.get("drawing") or e.get("table"): e for e in layout["past_margins"]}
    table_id = document.tables()[0].id
    assert past[table_id]["past"]["right"] > 100 and "fix" not in past[table_id]
    assert past[picture]["past"]["right"] > 100
    call(box, session, d, past[picture]["fix"])
    assert picture not in {e.get("drawing") for e in facts(box, session, d)["past_margins"]}


# -- what the layout could not judge, pages, and the arguments ---------------------------------


def _vml_body() -> str:
    """A paragraph holding a VML shape that states no size: docx2svg draws a placeholder."""
    return (f'<w:p w14:paraId="{next(_ids):08X}" w14:textId="77777777"><w:r><w:pict>'
            '<v:rect style="position:absolute" fillcolor="#4472c4"/></w:pict></w:r>'
            '<w:r><w:t>A shape.</w:t></w:r></w:p>' + p("After it."))


def test_the_pages_the_layout_drew_placeholders_on(box, session):
    d = session.open(docx(_vml_body()), "shape.docx")
    layout = facts(box, session, d)
    [page] = layout["uncertain"]
    assert page["page"] == 1 and page["placeholders"][0]["reason"] == "drawing"
    assert page["placeholders"][0]["at"].startswith("p:")
    assert "unknown_from" not in layout


def test_pages_keep_the_facts_on_them_and_min_size_asks_for_the_facts(box, session):
    body = p("Small", '<w:sz w:val="12"/>') + p(FILLER, ppr="<w:pageBreakBefore/>") + p("Small too", '<w:sz w:val="12"/>')
    d = session.open(docx(body), "pages.docx")
    data = ok(box, session, "check", doc=d, min_size=9).data
    assert set(data) >= {"validate", "reflow", "fields", "layout"}
    assert sorted(e["pages"] for e in data["layout"]["small_text"]) == [[1], [2]]
    on_two = ok(box, session, "check", doc=d, pages=[2], min_size=9).data["layout"]
    assert [e["pages"] for e in on_two["small_text"]] == [[2]]


def test_layout_alone_is_no_reflow_baseline(box, session):
    d = session.open(docx(p("One") + p("Two")), "base.docx")
    ok(box, session, "check", doc=d, include=["reflow"])
    ok(box, session, "word_insert_text", doc=d, find="Two", text=" more")
    ok(box, session, "check", doc=d, include=["layout"])
    reflow = ok(box, session, "check", doc=d, include=["reflow"]).data["reflow"]
    assert reflow["since_last_check"]["changed"] == [1]


def test_a_document_refuses_deck_facts_and_names_layout():
    from ooxml_edit.tools import Toolbox as _Toolbox

    with _Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session = toolbox.session()
        d = session.open(docx(p("Text")), "deck-facts.docx")
        refused = toolbox.dispatch(session, "check", {"doc": d, "include": ["fit"]})
        assert not refused.ok and "layout" in refused.error.valid_options


def test_a_table_outdented_as_word_places_it_is_not_past_the_margin(box, session):
    """Word draws a table's first cell's text on the margin, the table outdented by that
    cell's left margin and half its outer border (below mode 15 by default; in mode 15 with
    the indent Word writes for it): that is no fact.  An indent that pushes it further out
    is.  Normal Table's cell margins are 108 twips, a w:sz 4 border's outer half 5."""
    aligned = table([["Text on the margin", "b"]], indent=-113)
    pushed = table([["Pushed out", "b"]], indent=-720)
    d = session.open(docx(p("Before") + aligned + p("Between") + pushed + p("End")), "outdent.docx")
    first, second = [t.id for t in session.entry(d).document.tables()]
    past = {e.get("table"): e for e in facts(box, session, d)["past_margins"]}
    assert first not in past
    assert 29 < past[second]["past"]["left"] < 32          # 36 pt, less the 5.65 Word allows


def test_markdown_column_alignment_through_styles_leaves_check_clean(box, session):
    """Round 22, Q3: a Markdown table's centred and right-aligned columns take the styles
    ``style_map`` names for ``table_cell_center`` and ``table_cell_right``, so check finds no
    direct formatting in them; with no style named, the cells keep their direct alignment,
    which check reports, and the result names the role that would avoid it."""
    made = Document.new()
    made.styles.add("Cell Centred", "paragraph", based_on="Normal", alignment="center")
    made.styles.add("Cell Right", "paragraph", based_on="Normal", alignment="right")
    markdown = "| Item | Count | Share |\n| :-- | :-: | --: |\n| One | 2 | 40% |\n| Two | 3 | 60% |\n"
    d = session.open(made.to_bytes(), "styled.docx")
    styled = ok(box, session, "word_insert_markdown", doc=d, at="end", markdown=markdown,
                style_map=[{"element": "table_cell_center", "style": "Cell Centred"},
                           {"element": "table_cell_right", "style": "Cell Right"}])
    assert not any("table_cell_" in w for w in styled.warnings)
    assert facts(box, session, d)["direct_formatting"] == []
    plain = session.open(made.to_bytes(), "plain.docx")
    direct = ok(box, session, "word_insert_markdown", doc=plain, at="end", markdown=markdown)
    assert sum("table_cell_center" in w or "table_cell_right" in w for w in direct.warnings) == 2
    found = facts(box, session, plain)["direct_formatting"]
    assert len(found) == 6 and all(set(e["props"]) == {"alignment"} for e in found)
