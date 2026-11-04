"""The table style new tables get, and a style's tab stops (round 31, R3 and R4), as Word
writes them (measured in Word 16 for Mac through the oracle: ``feedback-round31/
measurements-C.md``):

* Table Design > Set as Default (this document) writes the settings' ``w:defaultTableStyle``
  (after ``w:hyphenationZone``, before ``w:characterSpacingControl``) and nothing else; Word's
  ``default table style`` reads that setting, not a table style's ``w:default``, and a table
  with no style of its own is drawn as before.
* Word's own tab stop on a style is ``w:tab w:val="right" w:leader="dot" w:pos=...``; its clear
  of an inherited stop ``w:tab w:val="clear" w:pos=...`` (no leader), which Word honours.
  A right stop past the right margin is drawn at the paragraph's right edge.

Made-up documents, ``Document.new()``'s package."""

from __future__ import annotations

from fractions import Fraction

import pytest
from lxml import etree
from ooxml_edit.tools import Toolbox

from docx_agent import Document, EditError
from docx_agent.layout import PX_PER_PT
from docx_agent.tools import FORMAT, GROUPS, TOOLS

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PLAN = ('<w:style xmlns:w="' + W_NS + '" w:type="table" w:customStyle="1" w:styleId="PlanTable">'
        '<w:name w:val="Plan Table"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="2E75B6"/></w:tblBorders></w:tblPr></w:style>')


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, result.error and result.error.to_json())
    return result


def with_plan_table() -> Document:
    """A document with a made-up table style of its own and one Table Grid table."""
    document = Document.new()
    document.insert_table(1, 2, after=document.paragraphs()[0].id, style="Table Grid")
    styles = document.package.tree(document.package.styles_part())
    styles.append(etree.fromstring(PLAN))
    document.package.mark_dirty(document.package.styles_part())
    document._invalidate()
    return Document.open(document.to_bytes())


def settings(document: Document):
    return document.package.tree(document.package.settings_part())


def part(document: Document, name: str) -> bytes:
    return etree.tostring(document.package.tree(name))


# -- R3: the default table style -----------------------------------------------------------------


def test_set_default_writes_the_settings_and_nothing_else():
    document = with_plan_table()
    styles_before = part(document, document.package.styles_part())
    body_before = part(document, document.package.document_part())
    assert document.styles.table_default() is None
    result = document.styles.set_default("Plan Table")
    assert result.changed and result.id == "style:PlanTable"
    [node] = settings(document).findall(_W + "defaultTableStyle")
    assert node.get(_W + "val") == "PlanTable"
    names = [child.tag.rpartition("}")[2] for child in settings(document)]
    assert names.index("defaultTableStyle") < names.index("compat")
    assert "defaultTabStop" not in names or names.index("defaultTabStop") < names.index("defaultTableStyle")
    assert part(document, document.package.styles_part()) == styles_before
    assert part(document, document.package.document_part()) == body_before
    assert document.styles.table_default().name == "Plan Table"
    again = document.styles.set_default("plan table")
    assert not again.changed and again.warnings == ["'Plan Table' was already the default table style"]
    document.styles.set_default(None)
    assert settings(document).find(_W + "defaultTableStyle") is None and document.styles.table_default() is None


def test_new_tables_and_markdown_tables_take_the_default_first():
    document = with_plan_table()
    first = document.paragraphs()[0].id
    assert document.table_properties(document.insert_table(1, 1, after=first).id)["style"] == "Table Grid"
    document.styles.set_default("Plan Table")
    assert document.table_properties(document.insert_table(1, 1, after=first).id)["style"] == "Plan Table"
    result = document.insert_markdown("| a | b |\n|---|---|\n| 1 | 2 |\n")
    table = next(b for b in result.blocks if b.startswith("t:"))
    assert document.table_properties(table)["style"] == "Plan Table"
    # A style map's own table style still wins.
    result = document.insert_markdown("| a |\n|---|\n| 1 |\n", style_map={"table": "Normal Table"})
    table = next(b for b in result.blocks if b.startswith("t:"))
    assert document.table_properties(table)["style"] == "Normal Table"


def test_a_built_in_default_is_written_and_a_missing_one_is_none():
    document = Document.new()
    assert document.styles.find("Table Grid") is None
    document.styles.set_default("Table Grid")
    assert document.styles.table_default().name == "Table Grid"
    with pytest.raises(EditError):
        document.styles.set_default("Normal")                   # a paragraph style
    with pytest.raises(EditError):
        document.styles.set_default("Plan Table")               # none by that name
    settings(document).find(_W + "defaultTableStyle").set(_W + "val", "Gone")
    assert document.styles.table_default() is None


def test_word_styles_sets_lists_and_describes_the_default(box):
    session = box.session()
    d = session.open(with_plan_table().to_bytes(), "plan.docx")
    result = ok(box, session, "word_styles", doc=d, action="set_default", name="Plan Table")
    assert result.summary == "New tables get Plan Table"
    listed = ok(box, session, "word_styles", doc=d, action="list", kind="table")
    assert listed.summary.endswith("; new tables get Plan Table")
    assert [row["name"] for row in listed.data if row.get("default")] == ["Plan Table"]
    assert ok(box, session, "word_styles", doc=d, action="describe", name="Plan Table").data["default"] is True
    refused = box.dispatch(session, "word_styles", {"doc": d, "action": "set_default", "name": "Normal"})
    assert refused.error.code == "not_found" and "Plan Table" in refused.error.valid_options
    refused = box.dispatch(session, "word_styles", {"doc": d, "action": "set_default", "name": "Plan Table",
                                                    "kind": "paragraph"})
    assert refused.error.field == "kind"
    again = ok(box, session, "word_styles", doc=d, action="set_default", name="Plan Table")
    assert again.summary == "No change: new tables get Plan Table"


# -- R4: tab stops ---------------------------------------------------------------------------------


def contents_styles() -> Document:
    """A contents style with a right dot stop at 346.95 pt and a style based on it."""
    document = Document.new()
    document.styles.add("Contents level", based_on="Normal", tabs=[{"position": 346.95, "alignment": "right",
                                                                    "leader": "dot"}])
    document.styles.add("Contents wide", based_on="Contents level")
    for name in ("Contents level", "Contents wide"):
        document.insert_paragraph("Entry\t12", after=document.paragraphs()[-1].id, style=name)
    return document


def own_tabs(document: Document, name: str) -> list[dict]:
    node = document.styles._node(document.styles.get(name).id)
    return [dict(tab.attrib) for tab in node.findall(f"{_W}pPr/{_W}tabs/{_W}tab")]


def test_tabs_are_written_as_word_writes_them():
    document = contents_styles()
    assert own_tabs(document, "Contents level") == [{f"{_W}val": "right", f"{_W}leader": "dot", f"{_W}pos": "6939"}]
    result = document.styles.modify("Contents wide", tabs=[{"position": 400, "alignment": "right", "leader": "dot"},
                                                           {"position": 346.95, "alignment": "clear"}])
    assert result.warnings == []
    assert own_tabs(document, "Contents wide") == [{f"{_W}val": "clear", f"{_W}pos": "6939"},
                                                   {f"{_W}val": "right", f"{_W}leader": "dot", f"{_W}pos": "8000"}]
    node = document.styles._node(document.styles.get("Contents wide").id)
    assert etree.tostring(node.find(f"{_W}pPr/{_W}tabs")).count(b"<w:tab ") == 2
    assert document.styles.inherited_tabs(document.styles.get("Contents wide").id) == {6939: ("right", "dot")}


def test_a_clear_word_honours_and_docx2svg_draws_as_word():
    document = contents_styles()
    document.styles.modify("Contents wide", tabs=[{"position": 346.95, "alignment": "clear"}])
    ends = {}
    for line in document.layout().pages[0].lines:
        texts = [s for s in line.spans if s.kind == "text" and "12" in "".join(s.chars)]
        if texts:
            ends[line.paragraph_id] = float(Fraction(texts[-1].end) / PX_PER_PT)
    level, wide = (ends[p.id[2:]] for p in document.paragraphs()[-2:])
    # The page's left margin plus the stop (Word's PDF: 417.26 with a 70.9 pt margin); cleared,
    # the 12 follows the next default stop (Word's: 117.43).
    left = document.sections()[-1].margins["left"]
    assert level == pytest.approx(left + 346.95, abs=1)
    assert wide < left + 60


def test_a_stop_at_the_text_columns_end_and_what_tabs_refuse():
    document = contents_styles()
    width = document.sections()[-1].text_width
    document.styles.modify("Contents wide", tabs=[{"alignment": "right", "leader": "dot"}])
    assert own_tabs(document, "Contents wide") == [{f"{_W}val": "right", f"{_W}leader": "dot", f"{_W}pos": str(width)}]
    warned = document.styles.modify("Contents wide", tabs=[{"position": 100, "alignment": "clear"}])
    assert warned.warnings == ["the style inherits no tab stop at 100 pt to clear (it inherits: 346.95)"]
    document.styles.modify("Contents wide", tabs=[])
    assert own_tabs(document, "Contents wide") == []
    for tabs in ([{"position": 10, "alignment": "middle"}], [{"position": 10, "leader": "stars"}],
                 [{"position": 2000}], [{"alignment": "clear"}], [{"position": 10}, {"position": 10}],
                 [{"position": 10, "alignment": "clear", "leader": "dot"}], [{"at": 10}]):
        with pytest.raises(EditError):
            document.styles.modify("Contents wide", tabs=tabs)
    document.styles.add("Marked", kind="character")
    with pytest.raises(EditError):
        document.styles.modify("Marked", tabs=[{"position": 10}])


def test_sections_of_different_widths_put_the_stop_at_the_widest_and_say_so():
    document = contents_styles()
    first = document.insert_section_break(after=document.paragraphs()[1].id).id
    document.set_section(first, margin_left=36, margin_right=36)
    widths = sorted({s.text_width for s in document.sections()})
    assert len(widths) == 2
    result = document.styles.modify("Contents wide", tabs=[{"alignment": "right"}])
    assert own_tabs(document, "Contents wide")[0][f"{_W}pos"] == str(widths[-1])
    assert result.warnings and result.warnings[0].startswith("the sections' text columns differ")


def test_word_styles_tabs(box):
    session = box.session()
    d = session.open(contents_styles().to_bytes(), "contents.docx")
    ok(box, session, "word_styles", doc=d, action="modify", name="Contents wide",
       tabs=[{"position": 346.95, "alignment": "clear"}, {"alignment": "right", "leader": "dot"}])
    described = ok(box, session, "word_styles", doc=d, action="describe", name="Contents wide").data
    width = session.entry(d).document.sections()[-1].text_width / 20
    assert described["declared"]["tabs"] == [{"position": 346.95, "alignment": "clear"},
                                             {"position": width, "alignment": "right", "leader": "dot"}]
    ok(box, session, "word_styles", doc=d, action="add", name="Contents third", based_on="Contents level",
       tabs=[{"position": 200, "alignment": "left"}])
    assert ok(box, session, "word_styles", doc=d, action="describe",
              name="Contents third").data["declared"]["tabs"] == [{"position": 200.0, "alignment": "left"}]
    refused = box.dispatch(session, "word_styles", {"doc": d, "action": "describe", "name": "Contents wide",
                                                    "tabs": [{"position": 10}]})
    assert refused.error.field == "tabs"
    ok(box, session, "word_styles", doc=d, action="modify", name="Contents wide", tabs=[])
    assert "tabs" not in ok(box, session, "word_styles", doc=d, action="describe",
                            name="Contents wide").data["declared"]
