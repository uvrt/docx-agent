"""Round 36: a table parted from its lead-in or its line after (R29, check's ``table_parted``
layout fact), a span of blocks kept on one page (R30: ``word_format`` ``keep_together``,
``Document.keep_blocks_together``, and check's ``range`` saying whether each range is on one
page, kept, where it breaks and the fix), and a table's frame (R34: ``word_design_facts``
``table_frames``).

What the keeps take was measured in Word for Mac 16.106 (made-up probe, CHANGELOG 0.1.46):
the paragraph before a table kept with the next goes to the table's page with its first row;
a table on one page kept together and with the next goes whole to the line after's page,
the last row alone kept with the next takes only that row along; a span stays on one page
with every paragraph keeping its lines together and with the next (but the last) and every
row keeping with the next.  docx2svg draws these the same, which is what the layout facts
read.  Made-up documents only.
"""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402
from test_table_keeps import grid  # noqa: E402
from test_tools_layout_facts import FILLER, call, docx, facts, ok, p, with_style  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
KEEP = "<w:keepNext/>"
CAPTION_STYLE = ('<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/>'
                 '<w:basedOn w:val="Normal"/></w:style>')
#: A heading style that keeps with the next (its keep stays when a span is released).
HEAD_STYLE = ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="PartHead"><w:name w:val="Part head"/>'
              '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr></w:style>')


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


@pytest.fixture
def session(box):
    return box.session()


def _fill(count: int) -> str:
    return "".join(p(f"{FILLER} {k}") for k in range(count))


def _ids(session, d: str) -> list[str]:
    return [b["id"] for b in session.entry(d).document.blocks()]


def _pages(session, d: str, identifier: str) -> list[int]:
    return session.entry(d).document.layout().pages_of(identifier)


# -- R29: table_parted ---------------------------------------------------------------------------


def test_a_lead_in_left_on_the_page_before_its_table_and_its_fix(box, session):
    """Filler to the page's foot, a lead-in ending in a colon, an 8-row table on the next
    page; keeping the lead-in with the next takes it there (measured in Word)."""
    d = session.open(docx(_fill(20) + p("The figures below:") + grid(8) + p("After")), "lead.docx")
    ids = _ids(session, d)
    lead, table = ids[20], ids[21]
    [entry] = facts(box, session, d)["table_parted"]
    assert entry == {"table": table, "para": lead, "side": "lead_in", "kind": "colon",
                     "para_page": 1, "table_page": 2,
                     "fix": {"tool": "word_format", "calls": [{"targets": [lead], "keep_with_next": True}]}}
    call(box, session, d, entry["fix"])
    assert _pages(session, d, lead) == [2]
    after = facts(box, session, d)
    assert after["table_parted"] == [] and after["keep_with_next_at_foot"] == []


def test_a_long_paragraph_before_a_table_on_the_next_page_is_no_fact(box, session):
    lines = "".join(p(f"Line {k}") for k in range(25))
    d = session.open(docx(lines + p(f"{FILLER} {FILLER} {FILLER}") + grid(8) + p("After")), "long.docx")
    document = session.entry(d).document
    ids = _ids(session, d)
    assert max(document.layout().pages_of(ids[25])) < min(document.layout().pages_of(ids[26])), \
        "the made-up page must part them"
    assert facts(box, session, d)["table_parted"] == []


def test_a_caption_after_a_table_on_one_page_and_its_fix_moves_the_table(box, session):
    """The table ends at the page's foot and its caption starts the next: on one page, the
    fix keeps it together and with the next, and the whole table goes along."""
    body = (p("Intro") + _fill(15) + grid(8) + p("Source: made-up figures", ppr='<w:pStyle w:val="Caption"/>')
            + p("After"))
    d = session.open(with_style(docx(body), CAPTION_STYLE), "caption.docx")
    ids = _ids(session, d)
    table, caption = ids[16], ids[17]
    [entry] = facts(box, session, d)["table_parted"]
    assert entry["side"] == "line_after" and entry["kind"] == "caption"
    assert (entry["table"], entry["para"], entry["row"], entry["row_page"], entry["para_page"]) == \
        (table, caption, 7, 1, 2)
    assert entry["fix"] == {"tool": "word_format_table", "calls": [
        {"table": table, "scope": "table", "keep_with_next": True, "keep_together": True}]}
    call(box, session, d, entry["fix"])
    table = session.entry(d).document.tables()[0].id                 # its rows stamped
    assert _pages(session, d, table) == [2] and _pages(session, d, caption) == [2]
    assert facts(box, session, d)["table_parted"] == []


def test_a_line_after_a_table_over_two_pages_takes_its_last_row(box, session):
    """54 rows fill two pages to the foot; the line after starts the third.  Over two pages
    the fix keeps the last row alone with the next: that row goes along, the rest stays."""
    d = session.open(docx(grid(54) + p("Line after the rows") + p("After")), "rows.docx")
    ids = _ids(session, d)
    [entry] = facts(box, session, d)["table_parted"]
    lead = session.entry(d).document.tables()[0].cell(53, 0).paragraphs[0].id
    assert (entry["side"], entry["kind"], entry["row"], entry["row_page"], entry["para_page"]) == \
        ("line_after", "one_line", 53, 2, 3)
    assert entry["fix"] == {"tool": "word_format", "calls": [{"targets": [lead], "keep_with_next": True}]}
    call(box, session, d, entry["fix"])
    assert _pages(session, d, ids[1]) == [3] and _pages(session, d, ids[0]) == [1, 2, 3]
    assert facts(box, session, d)["table_parted"] == []


def test_a_last_row_kept_alone_moves_with_the_line_after(box, session):
    """The fix for a table over several pages: its last row's deciding paragraph keeps with
    the next, and only that row goes to the line's page (measured in Word)."""
    document = Document.open(docx(p("Intro") + _fill(15) + grid(8) + p("Line after the rows") + p("After")))
    table = document.tables()[0]
    lead = table.cell(7, 0).paragraphs[0].id
    document.format_paragraph(lead, keep_with_next=True)
    layout = document.layout()
    pages = layout.pages_of(document.tables()[0].id)
    assert pages == [1, 2]
    line = [b["id"] for b in document.blocks()][17]
    assert layout.pages_of(line) == [2]


def test_already_kept_or_a_page_break_between_is_no_table_parted(box, session):
    kept = docx(_fill(19) + p("The figures below:", ppr=KEEP) + grid(30) + p("After"))
    d = session.open(kept, "kept.docx")
    assert facts(box, session, d)["table_parted"] == []
    broken = docx(_fill(3) + p("The figures below:") + grid(4) + p("Line", ppr="<w:pageBreakBefore/>") + p("End"))
    d = session.open(broken, "break.docx")
    assert facts(box, session, d)["table_parted"] == []


def test_table_parted_is_scoped_by_range_and_pages(box, session):
    d = session.open(docx(_fill(20) + p("The figures below:") + grid(8) + p("After")), "scoped.docx")
    ids = _ids(session, d)
    assert len(facts(box, session, d, range=[ids[21]])["table_parted"]) == 1
    assert facts(box, session, d, range=[ids[0]])["table_parted"] == []
    assert facts(box, session, d, pages=[3])["table_parted"] == []


# -- R30: keep a span together -------------------------------------------------------------------


def _span_doc(fill: int = 17, head: str = "") -> bytes:
    words = " ".join(f"word{k:02d} lorem ipsum" for k in range(40))
    return docx(_fill(fill) + p("Head", ppr=head) + p(words) + grid(5) + p("Closing line") + p("After"))


def test_word_format_keep_together_keeps_a_span_on_one_page(box, session):
    d = session.open(_span_doc(), "span.docx")
    ids = _ids(session, d)
    span = f"{ids[17]}..{ids[20]}"
    before = ok(box, session, "check", doc=d, range=[span]).data["range"][0]
    assert before["one_page"] is False and before["kept"] is False
    assert before["breaks"] == [{"inside": ids[18], "pages": [1, 2]}]
    assert before["fix"] == {"tool": "word_format", "calls": [{"targets": [span], "keep_together": True}]}
    call(box, session, d, before["fix"])
    ids = _ids(session, d)
    for identifier in ids[17:21]:
        assert _pages(session, d, identifier) == [2]
    after = ok(box, session, "check", doc=d, range=[f"{ids[17]}..{ids[20]}"]).data["range"][0]
    assert after["one_page"] is True and after["kept"] is True and "breaks" not in after and "fix" not in after

    # What it wrote: lines kept and with the next on every paragraph but the last; the
    # table's rows all keep with the next; the last block keeps its lines only.
    document = session.entry(d).document
    paragraphs = [document.paragraph(i).effective for i in (ids[17], ids[18], ids[20])]
    assert [(e.keep_together, e.keep_with_next) for e in paragraphs] == [(True, True), (True, True), (True, False)]
    assert document.table_properties(ids[19])["keep_together"] is True
    assert document.table_properties(ids[19])["keep_with_next"] is True


def test_keep_together_again_is_a_no_op_and_false_releases_only_what_it_wrote(box, session):
    d = session.open(with_style(_span_doc(head='<w:pStyle w:val="PartHead"/>'), HEAD_STYLE), "release.docx")
    ids = _ids(session, d)
    span = f"{ids[17]}..{ids[20]}"
    ok(box, session, "word_format", doc=d, targets=[span], keep_together=True)
    again = ok(box, session, "word_format", doc=d, targets=[span], keep_together=True)
    assert again.changed == [] and any("already kept together" in w for w in again.warnings)
    ok(box, session, "word_format", doc=d, targets=[span], keep_together=False)
    document = session.entry(d).document
    head, text = document.paragraph(ids[17]).effective, document.paragraph(ids[18]).effective
    assert head.keep_with_next is True                      # its style's keep stays
    assert (text.keep_with_next, text.keep_together) == (False, False)
    assert document.table_properties(ids[19])["keep_together"] is False
    xml = document.package.tree(document.package.document_part())
    assert xml.find(f".//{_W}keepLines") is None
    released = ok(box, session, "word_format", doc=d, targets=[span], keep_together=False)
    assert released.changed == [] and any("not kept together" in w for w in released.warnings)


def test_keep_together_on_one_paragraph_keeps_its_lines(box, session):
    d = session.open(docx(p("A paragraph") + p("End")), "one.docx")
    target = _ids(session, d)[0]
    result = ok(box, session, "word_format", doc=d, targets=[target], keep_together=True)
    assert result.warnings == []
    effective = session.entry(d).document.paragraph(target).effective
    assert effective.keep_together is True and effective.keep_with_next is False


def test_a_span_taller_than_a_page_is_said_by_the_setter_and_the_check(box, session):
    d = session.open(docx(_fill(3) + p("Head") + grid(70) + p("Closing line") + p("After")), "tall.docx")
    ids = _ids(session, d)
    span = f"{ids[3]}..{ids[5]}"
    ok(box, session, "check", doc=d, include=["layout"])           # a layout to measure it in
    result = ok(box, session, "word_format", doc=d, targets=[span], keep_together=True)
    assert any("cannot be kept on one page" in w for w in result.warnings)
    ids = _ids(session, d)
    found = ok(box, session, "check", doc=d, range=[f"{ids[3]}..{ids[5]}"]).data["range"][0]
    assert found["one_page"] is False and found["kept"] is True and found["taller_than_page"] is True
    assert "fix" not in found


def test_check_range_of_a_cell_says_nothing_of_keeps(box, session):
    d = session.open(docx(p("Intro") + grid(3) + p("End")), "cell.docx")
    table = _ids(session, d)[1]
    found = ok(box, session, "check", doc=d, range=[f"{table}/c0,0"]).data["range"][0]
    assert "one_page" not in found and "kept" not in found


def test_keep_together_refuses_what_names_no_blocks_with_the_way_out(box, session):
    d = session.open(docx(p("Intro") + p("End")), "refuse.docx")
    refused = box.dispatch(session, "word_format", {"doc": d, "targets": ["s:99"], "keep_together": True})
    assert not refused.ok and "s:1" in refused.error.message


def test_a_section_is_kept_together_as_its_blocks(box, session):
    """Rule 1: a section (s:...) given to keep_together is its blocks, as check's range reads it."""
    d = session.open(_span_doc(), "section.docx")
    section = session.entry(d).document.sections()[0].id
    ok(box, session, "word_format", doc=d, targets=[section], keep_together=True)
    ids = _ids(session, d)
    effective = session.entry(d).document.paragraph(ids[0]).effective
    assert effective.keep_together and effective.keep_with_next


def test_a_range_lists_at_most_ten_breaks_and_counts_the_rest(box, session):
    body = p("Start") + "".join(p(f"Page {k}", ppr="<w:pageBreakBefore/>") for k in range(12)) + p("End")
    d = session.open(docx(body), "breaks.docx")
    ids = _ids(session, d)
    found = ok(box, session, "check", doc=d, range=[f"{ids[0]}..{ids[12]}"]).data["range"][0]
    assert len(found["breaks"]) == 10 and found["breaks_more"] == 2
    assert found["breaks"][0] == {"after": ids[0], "before": ids[1], "page": 2, "page_break": True}
    assert found["kept"] is False and "fix" not in found       # no keep undoes a page break


def test_keep_blocks_together_in_python_writes_word_flags_only_where_styles_lack_them():
    document = Document.open(with_style(_span_doc(head='<w:pStyle w:val="PartHead"/>'), HEAD_STYLE))
    ids = [b["id"] for b in document.blocks()]
    result = document.keep_blocks_together(f"{ids[17]}..{ids[20]}")
    assert result.changed
    xml = document.package.tree(document.package.document_part())
    head = next(e for e in xml.iter(_W + "p") if "".join(t.text or "" for t in e.iter(_W + "t")) == "Head")
    assert head.find(f"{_W}pPr/{_W}keepNext") is None              # its style keeps it
    assert head.find(f"{_W}pPr/{_W}keepLines") is not None
    assert document.layout().pages_of(ids[17]) == document.layout().pages_of(ids[20]) == [2]


# -- R34: table frames ---------------------------------------------------------------------------


def _table(width: int, columns: list[int], rows: list[tuple[str, str, str | None]], indent: int = 0) -> str:
    """A table of ``width`` twips (``w:tblW`` dxa), its grid ``columns``; each row: two cell
    texts and its ``w:trHeight`` (``"400"``, ``"400 exact"`` or ``None``)."""
    grid_ = "".join(f'<w:gridCol w:w="{c}"/>' for c in columns)
    indented = f'<w:tblInd w:w="{indent}" w:type="dxa"/>' if indent else ""
    out = (f'<w:tbl><w:tblPr><w:tblW w:w="{width}" w:type="dxa"/>{indented}</w:tblPr>'
           f'<w:tblGrid>{grid_}</w:tblGrid>')
    for first, second, height in rows:
        trpr = ""
        if height:
            value, _, rule = height.partition(" ")
            trpr = f'<w:trPr><w:trHeight w:val="{value}"{f" w:hRule={chr(34)}{rule}{chr(34)}" if rule else ""}/></w:trPr>'
        cells = "".join(f'<w:tc><w:tcPr><w:tcW w:w="{c}" w:type="dxa"/></w:tcPr>{p(t)}</w:tc>'
                        for c, t in zip(columns, (first, second)))
        out += f"<w:tr>{trpr}{cells}</w:tr>"
    return out + "</w:tbl>"


def test_table_frames_width_against_the_body_columns_and_empty_last_cells(box, session):
    document = Document.new()
    body = document.sections()[0].text_width            # twips
    full = _table(body, [body // 3, body - body // 3], [("Label", "Value", None), ("Ask", "", "400")])
    same = _table(body, [body // 2, body - body // 2], [("Label", "Value", None)])
    narrow = _table(body - 1440, [3000, body - 1440 - 3000],
                    [("Ask", "", "567 exact"), ("Ask", "", None), ("Label", "Value", None)])
    wide = _table(body, [body // 2, body - body // 2], [("Label", "Value", None)], indent=720)
    d = session.open(docx(p("Intro") + full + p("a") + same + p("b") + narrow + p("c") + wide + p("End")),
                     "frames.docx")
    data = ok(box, session, "word_design_facts", doc=d, include=["table_frames"]).data
    frames = data["table_frames"]
    assert data["usual_table_width"] == body / 20
    first, second, third, fourth = frames
    assert first["width"] == first["body"] == body / 20 and "over_body" not in first
    assert first["columns"] == [round(body // 3 / 20, 2), round((body - body // 3) / 20, 2)]
    assert first["empty_last_cell"] == [{"row": 1, "height": 20, "rule": "at_least"}]
    assert "empty_last_cell" not in second and "narrower_than_usual" not in second
    assert third["narrower_than_usual"] is True and third["width"] == (body - 1440) / 20
    assert third["empty_last_cell"] == [{"row": 0, "height": 28.35, "rule": "exact"}, {"row": 1, "height": None}]
    assert fourth["indent"] == 36 and fourth["over_body"] == 36
    # Not asked for: the default facts list them too, after the table styles.
    defaults = ok(box, session, "word_design_facts", doc=d).data
    assert list(defaults).index("table_frames") == list(defaults).index("table_styles") + 1


def test_a_nested_table_and_a_percent_width_in_table_frames():
    document = Document.new()
    body = document.sections()[0].text_width
    nested = _table(2000, [1000, 1000], [("in", "ner", None)])
    outer = (f'<w:tbl><w:tblPr><w:tblW w:w="2500" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="{body // 2}"/>'
             f'</w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="{body // 2}" w:type="dxa"/></w:tcPr>{nested}{p("x")}'
             '</w:tc></w:tr></w:tbl>')
    made = Document.open(docx(p("Intro") + outer + p("End")))
    [frame] = made.design_facts(["table_frames"])["table_frames"]
    assert frame["width"] == round(body / 20 / 2, 2) and "usual_table_width" not in made.design_facts(["table_frames"])
