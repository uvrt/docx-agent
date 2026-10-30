"""Round 25, W4: ``word_format`` ``keep_lines`` -- Word's "Keep lines together"
(``w:keepLines``; the Python API's ``keep_together``) -- set, turned off and cleared, and
reported by check's ``direct_formatting`` as ``keep_with_next`` is, its fix clearing it; the
styles' describe, word_compare_blocks, word_inspect and word_design_facts name it so too (a
table's ``keep_together`` is word_format_table's, another thing).  A paragraph that keeps its lines
together and does not fit at a page's foot moves whole: docx2svg's rule, which Word for Mac
16.106 drew the same in a made-up probe (CHANGELOG, round 25, W4).  Made-up documents only.
"""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402
from test_tools_layout_facts import FILLER, call, docx, facts, ok, p, with_style  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
#: A paragraph style that keeps its lines together.
WHOLE_STYLE = ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="Whole"><w:name w:val="Whole"/>'
               '<w:basedOn w:val="Normal"/><w:pPr><w:keepLines/></w:pPr></w:style>')


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


@pytest.fixture
def session(box):
    return box.session()


def _element(session, d: str, k: int = 0):
    document = session.entry(d).document
    return document.package.tree(document.package.document_part()).findall(f"{_W}body/{_W}p")[k]


def test_keep_lines_is_set_turned_off_and_cleared(box, session):
    d = session.open(docx(p("A paragraph") + p("End")), "keep.docx")
    target = session.entry(d).document.paragraphs()[0].id
    ok(box, session, "word_format", doc=d, targets=[target], keep_lines=True)
    node = _element(session, d).find(f"{_W}pPr/{_W}keepLines")
    assert node is not None and node.get(_W + "val") is None
    assert session.entry(d).document.paragraphs()[0].effective.keep_together
    inspected = ok(box, session, "word_inspect", doc=d, target=target).data["blocks"][0]
    assert inspected["effective"]["keep_lines"] is True and "keep_together" not in inspected["effective"]
    # The Python API keeps its name.
    state = session.entry(d).document.state(target)
    assert state["blocks"][0]["effective"]["keep_together"] is True

    ok(box, session, "word_format", doc=d, targets=[target], keep_lines=False)
    assert not session.entry(d).document.paragraphs()[0].effective.keep_together
    ok(box, session, "word_format", doc=d, targets=[target], clear=["keep_lines"])
    assert _element(session, d).find(f"{_W}pPr/{_W}keepLines") is None

    # The Python API's name is read as word_format's, and the call says so.
    result = ok(box, session, "word_format", doc=d, targets=[target], keep_together=True)
    assert _element(session, d).find(f"{_W}pPr/{_W}keepLines") is not None
    assert result.warnings[0] == "keep_together read as keep_lines"

    refused = box.dispatch(session, "word_format", {"doc": d, "targets": [target], "keep_lines": True,
                                                    "clear": ["keep_lines"]})
    assert not refused.ok and refused.error.field == "clear"


def test_check_reports_a_direct_keep_lines_and_its_fix_clears_it(box, session):
    whole = p("Kept whole", ppr="<w:keepLines/>")
    released = p("Released", ppr='<w:pStyle w:val="Whole"/><w:keepLines w:val="0"/>')
    as_styled = p("As styled", ppr='<w:pStyle w:val="Whole"/><w:keepLines/>')
    d = session.open(with_style(docx(whole + released + as_styled + p("End")), WHOLE_STYLE), "facts.docx")
    ids = [paragraph.id for paragraph in session.entry(d).document.paragraphs()]
    found = {e["para"]: e for e in facts(box, session, d)["direct_formatting"]}
    assert found[ids[0]]["props"] == {"keep_lines": True} and found[ids[0]]["style_has"] == {"keep_lines": False}
    assert found[ids[1]]["props"] == {"keep_lines": False} and found[ids[1]]["style_has"] == {"keep_lines": True}
    assert ids[2] not in found                       # the same as its style: not direct formatting
    assert found[ids[0]]["fix"] == {"tool": "word_format", "calls": [{"targets": [ids[0]], "clear": ["keep_lines"]}]}
    for k in (0, 1):
        call(box, session, d, found[ids[k]]["fix"])
    assert [e for e in facts(box, session, d)["direct_formatting"] if e["para"] in ids[:2]] == []
    document = session.entry(d).document
    assert [p_.effective.keep_together for p_ in document.paragraphs()[:3]] == [False, True, True]


def test_a_paragraph_kept_whole_moves_to_the_next_page(box, session):
    """Filler to near a page's foot, then a long paragraph: split across the pages, and
    on the next page whole once its lines are kept together."""
    long = " ".join([FILLER] * 5)
    d = session.open(docx("".join(p(f"{FILLER} {k}") for k in range(19)) + p(long) + p("End")), "move.docx")
    document = session.entry(d).document
    target = document.paragraphs()[19].id
    split = document.layout().where(target)
    assert len({box_.page for box_ in split}) == 2, "the made-up page must split the paragraph"
    ok(box, session, "word_format", doc=d, targets=[target], keep_lines=True)
    kept = session.entry(d).document.layout().where(target)
    assert {box_.page for box_ in kept} == {max(box_.page for box_ in split)}


def test_styles_describe_and_compare_blocks_name_keep_lines(box, session):
    d = session.open(with_whole(), "styles.docx")
    declared = ok(box, session, "word_styles", doc=d, action="describe", name="Whole").data["declared"]
    assert declared["keep_lines"] is True and "keep_together" not in declared
    facts_ = ok(box, session, "word_design_facts", doc=d, include=["declared_styles"]).data
    [whole] = [s for s in facts_["declared_styles"] if s["name"] == "Whole"]
    assert whole["look"]["keep_lines"] is True and "keep_together" not in whole["look"]
    python = session.entry(d).document.design_facts(["declared_styles"])["declared_styles"]
    assert [s for s in python if s["name"] == "Whole"][0]["look"]["keep_together"] is True

    document = Document.new()
    made = document.insert_markdown("Sample text.\n\nCopied text.", at="end")
    sample, copy = made.blocks[:2]
    document.format_paragraph(sample, keep_together=True)
    data = document.compare_blocks([copy], sample)
    [difference] = [x for x in data["differences"] if x["property"] == "keep_lines"]
    assert difference["value"] is None and difference["reference"] is True
    assert difference["fix"] == {"tool": "word_format", "calls": [{"targets": [copy], "keep_lines": True}]}


def with_whole() -> bytes:
    """A paragraph in the style that keeps its lines together."""
    return with_style(docx(p("Styled", ppr='<w:pStyle w:val="Whole"/>')), WHOLE_STYLE)
