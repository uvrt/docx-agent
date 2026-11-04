"""Styles inheriting a list level, a level's positions, and how a style numbers (round 30, W13
to W15).

Word's behaviour, measured on made-up documents (``feedback-round30/measurements-X.md``):

- A style based on a style linked to a level numbers at that level too: Word numbers its
  paragraphs, and its style description says the level.  Linking the base to a level (Word's
  "Link level to style") changes nothing in the derived style.
- "Numbering: None" on such a style is ``w:numPr`` with ``w:numId`` 0 and no ``w:ilvl`` in the
  style (Word's style redefined from a paragraph whose numbering was removed; the paragraph's
  own "None" is ``w:ilvl`` 0, ``w:numId`` 0).  Word keeps it on saving, draws no number and
  describes the style as "no bullets or numbering"; the paragraphs after it count on as if
  it were not there.
- "Define new multilevel list": the number's "Aligned at" and "Text indent at" are the
  level's ``w:ind`` -- ``left`` the text, ``hanging`` the difference, ``firstLine`` where the
  number is right of the text (``firstLine`` 0 where they are one); "Follow number with" a
  space or nothing is ``w:suff`` (a tab: none); "Add tab stop at" a ``num`` tab stop.  docx2svg
  draws every label and text where Word does (each case through Word's PDF).

The document: one multilevel list (numId 1) whose level 0 (``%1.``) is linked to "Chapter
title" and level 1 (``a.``) to "Section title"; "Front title" is based on "Chapter title",
"Figure title" on "Section title"; "Aside title" is in no list.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from docx_agent import Document, EditError
from docx_agent.edit.numbering import problems
from docx_agent.layout_facts import numbering_skips_level
from docx_agent.validate import check

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml."
LONG = " words to wrap" * 24


def _style(style_id: str, name: str, ppr: str = "", based: str = "Normal") -> str:
    return (f'<w:style w:type="paragraph" w:customStyle="1" w:styleId="{style_id}"><w:name w:val="{name}"/>'
            f'<w:basedOn w:val="{based}"/><w:qFormat/><w:pPr>{ppr}</w:pPr></w:style>')


def _lvl(level: int, fmt: str, text: str, left: int, hanging: int, style: str | None) -> str:
    return (f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            + (f'<w:pStyle w:val="{style}"/>' if style else "")
            + f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr></w:lvl>')


def _paragraph(text: str, style: str | None = None) -> str:
    return ("<w:p>" + (f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else "")
            + f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


#: Chapters and sections, a front title, a figure title, an aside, body text.
BODY = [("One", "ChapterTitle"), ("One first", "SectionTitle"), ("Front", "FrontTitle"),
        ("A figure", "FigureTitle"), ("One second", "SectionTitle"), ("Two", "ChapterTitle"),
        ("Aside", "AsideTitle"), ("Body text", None)]


def make(body=BODY, linked: bool = True, section: str = "%2.") -> bytes:
    chapter = '<w:numPr><w:numId w:val="1"/></w:numPr>' if linked else ""
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>'
              '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              + _style("ChapterTitle", "Chapter title", chapter)
              + _style("SectionTitle", "Section title", '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>')
              + _style("FrontTitle", "Front title", based="ChapterTitle")
              + _style("FigureTitle", "Figure title", based="SectionTitle")
              + _style("AsideTitle", "Aside title", '<w:ind w:left="1100" w:hanging="200"/>')
              + "</w:styles>")
    numbering = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS}>'
                 '<w:abstractNum w:abstractNumId="0"><w:nsid w:val="6A7B8C9D"/>'
                 '<w:multiLevelType w:val="multilevel"/>'
                 + _lvl(0, "decimal", "%1.", 432, 432, "ChapterTitle" if linked else None)
                 + _lvl(1, "lowerLetter", section, 576, 576, "SectionTitle")
                 + _lvl(2, "lowerLetter", "%3)", 1080, 360, None)
                 + '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
                + "".join(_paragraph(text, style) for text, style in body)
                + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" '
                'w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>'
                "</w:body></w:document>")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" ContentType="{_CT}document.main+xml"/>'
            f'<Override PartName="/word/styles.xml" ContentType="{_CT}styles+xml"/>'
            f'<Override PartName="/word/numbering.xml" ContentType="{_CT}numbering+xml"/></Types>'))
        archive.writestr("_rels/.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}officeDocument" Target="word/document.xml"/></Relationships>'))
        archive.writestr("word/_rels/document.xml.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}styles" Target="styles.xml"/>'
            f'<Relationship Id="rId2" Type="{_REL}numbering" Target="numbering.xml"/></Relationships>'))
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", styles)
        archive.writestr("word/numbering.xml", numbering)
    return out.getvalue()


def style_numpr(document: Document, style_id: str) -> str | None:
    root = document.package.tree(document.package.styles_part())
    node = next(s for s in root.iter(W + "style") if s.get(W + "styleId") == style_id)
    found = node.find(f"{W}pPr/{W}numPr")
    return None if found is None else re.sub(r' xmlns:\w+="[^"]+"', "", _xml(found))


def level_xml(document: Document, level: int) -> str:
    root = document.package.tree(document.package.numbering_part())
    lvl = next(n for n in root.iter(W + "lvl") if n.get(W + "ilvl") == str(level))
    return re.sub(r' xmlns:\w+="[^"]+"', "", _xml(lvl))


def _xml(node) -> str:
    from lxml import etree

    return etree.tostring(node).decode()


def lines(document: Document) -> list[tuple[str, float | None, float | None]]:
    """docx2svg's lines: (label, its x, the text's x), in points."""
    import docx2svg

    out = []
    scale = 2480 / 595.3
    for svg in docx2svg.convert_docx_to_svg(document.to_bytes()):
        for line in re.findall(r'<g data-docx-line="\d+"[^>]*>(.*?)</g>', svg, re.S):
            label, lx, tx = "", None, None
            for attrs, body in re.findall(r"<text ([^>]*)>([^<]*)</text>", line):
                x = round(float(re.search(r'x="([\d.]+)', attrs).group(1)) / scale, 1)
                if 'data-docx-kind="label"' in attrs:
                    label, lx = body.strip(), x
                elif tx is None:
                    tx = x
            out.append((label, lx, tx))
    return out


def labels(document: Document) -> list[str]:
    return [label for label, _, _ in lines(document)]


def valid(document: Document) -> None:
    assert problems(document) == []
    reopened = Document.open(document.to_bytes())
    assert [str(p) for p in check(reopened.package) if p.code.startswith(("numbering", "style"))] == []


# -- W13: a style inheriting a linked level -------------------------------------------------------


def test_a_style_based_on_a_linked_style_numbers_at_its_level():
    # Word numbers them so (measured): 1. a. 2. a. b. 3.
    assert labels(Document.open(make())) == ["1.", "a.", "2.", "a.", "b.", "3.", "", ""]


def test_none_stops_a_link_the_style_inherits():
    data = make()
    document = Document.open(data)
    result = document.styles.modify("Front title", numbering=None)
    assert result.changed and not result.warnings
    assert style_numpr(document, "FrontTitle") == '<w:numPr><w:numId w:val="0"/></w:numPr>'    # as Word
    assert style_numpr(document, "ChapterTitle") is not None                  # the base keeps its link
    document.styles.modify("Figure title", numbering=None)
    # Word's numbers (measured): the two unnumbered, the paragraphs after counting on.
    assert labels(document) == ["1.", "a.", "", "", "b.", "2.", "", ""]
    front = document.paragraphs()[2].id
    assert document.list_of(front) is None
    assert numbering_skips_level(document) == []
    valid(document)
    assert document.undo() and document.undo()
    assert document.to_bytes() == Document.open(data).to_bytes()


def test_none_on_a_style_with_its_own_link_and_an_inherited_one_stops_both():
    document = Document.open(make())
    document.styles.modify("Front title", numbering={"list": 1, "level": 2})
    assert style_numpr(document, "FrontTitle") == '<w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr>'
    document.styles.modify("Front title", numbering=None)
    assert style_numpr(document, "FrontTitle") == '<w:numPr><w:numId w:val="0"/></w:numPr>'
    assert labels(document)[2] == ""
    valid(document)


@pytest.mark.parametrize("name", ["Aside title", "Front title"])
def test_none_on_a_style_that_numbers_nothing_says_so(name):
    document = Document.open(make())
    if name == "Front title":
        document.styles.modify(name, numbering=None)              # numbering none already
    def parts():
        return [_xml(document.package.tree(part)) for part in (document.package.styles_part(),
                                                               document.package.numbering_part())]

    before = parts()
    result = document.styles.modify(name, numbering=None)
    assert not result.changed
    assert result.warnings == [f"{name!r} numbers nothing, neither its own nor through based_on: left as it was"]
    assert parts() == before


def test_linking_a_level_names_the_styles_that_inherit_it():
    document = Document.open(make(linked=False))
    assert labels(document)[:3] == ["", "a.", ""]
    result = document.styles.modify("Chapter title", numbering={"list": 1, "level": 0})
    assert result.warnings == ["'Front title' is based on 'Chapter title': it numbers at level 0 of list 1 "
                               "too, its 1 paragraph(s) (numbering none on it stops that)"]
    assert style_numpr(document, "FrontTitle") is None            # Word leaves it as it was
    assert labels(document)[:3] == ["1.", "a.", "2."]
    # A style based on it that stops numbering is not named.
    document.styles.modify("Front title", numbering=None)
    result = document.styles.modify("Chapter title", numbering={"list": 1, "level": 0})
    assert result.warnings == []


def test_a_list_number_of_zero_is_no_numbering_everywhere():
    document = Document.open(make())
    document.styles.modify("Front title", numbering=None)
    document.styles.modify("Figure title", numbering=None)
    with pytest.raises(EditError, match="'Front title' is in no list"):
        document.styles.modify("Aside title", numbering={"list": "Front title", "level": 1})


# -- W14: a level's positions and what follows its number ----------------------------------------


W14_BODY = [("First" + LONG, "ChapterTitle"), ("Second" + LONG, "SectionTitle"), ("Third", "SectionTitle"),
            ("Body text", None)]


@pytest.mark.parametrize("arguments, ind, label_x, text_x", [
    # Word's "Aligned at" 18, "Text indent at" 54 (measured: label 88.8, text 124.8 pt).
    ({"number_at": 18, "text_at": 54}, '<w:ind w:left="1080" w:hanging="720"/>', 88.8, 124.8),
    ({"number_at": 36, "text_at": 36}, '<w:ind w:left="720" w:firstLine="0"/>', 106.8, None),
    ({"number_at": 54, "text_at": 18}, '<w:ind w:left="360" w:firstLine="720"/>', 124.8, None),
    ({"text_at": 54}, '<w:ind w:left="1080" w:hanging="1080"/>', 70.8, 124.8),     # the number stays at 0
])
def test_number_and_text_positions_are_the_levels_indent(arguments, ind, label_x, text_x):
    data = make(W14_BODY)
    document = Document.open(data)
    second = document.paragraphs()[1].id
    document.set_list_format(second, **arguments)
    assert ind in level_xml(document, 1)
    label, lx, tx = lines(document)[next(k for k, line in enumerate(lines(document)) if line[0] == "a.")]
    assert lx == label_x and (text_x is None or tx == text_x)
    valid(document)
    assert document.undo() and document.to_bytes() == Document.open(data).to_bytes()


def test_follow_and_the_tab_stop_are_written_as_word_writes_them():
    document = Document.open(make(W14_BODY))
    second = document.paragraphs()[1].id
    document.set_list_format(second, follow="space")
    assert '<w:suff w:val="space"/><w:lvlText' in level_xml(document, 1)
    document.set_list_format(second, follow="nothing")
    assert '<w:suff w:val="nothing"/>' in level_xml(document, 1)
    document.set_list_format(second, follow="tab", number_at=18, text_at=54, tab_at=90)
    xml = level_xml(document, 1)
    assert "w:suff" not in xml                                    # a tab is the default: Word writes none
    assert '<w:pPr><w:tabs><w:tab w:val="num" w:pos="1800"/></w:tabs><w:ind w:left="1080" w:hanging="720"/>' in xml
    document.set_list_format(second, tab_at=36)                   # one num tab stop, moved
    xml = level_xml(document, 1)
    assert xml.count("<w:tab ") == 1 and '<w:tab w:val="num" w:pos="720"/>' in xml
    # A label then text in the paragraph's line (docx2svg as Word: the tab at 54 pt's text indent).
    assert [line for line in lines(document) if line[0] == "a."][0][1:] == (88.8, 106.8)
    valid(document)


@pytest.mark.parametrize("arguments, says", [
    ({"follow": "dash"}, "follow is one of tab, space, nothing"),
    ({"number_at": 2000}, "number_at is in points"),
    ({"tab_at": "1"}, "tab_at is in points"),
])
def test_positions_out_of_range_are_refused(arguments, says):
    document = Document.open(make(W14_BODY))
    with pytest.raises(EditError, match=says):
        document.set_list_format(document.paragraphs()[1].id, **arguments)


# -- the tools: word_lists format, word_styles describe, check -----------------------------------


@pytest.fixture(scope="module")
def box():
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def dispatch(box, session, tool, /, **arguments):
    return box.dispatch(session, tool, arguments)


def test_word_lists_sets_a_levels_positions(box):
    session = box.session()
    session.open(make(W14_BODY), "positions.docx")
    second = session.entry("d1").document.paragraphs()[1].id
    result = dispatch(box, session, "word_lists", doc="d1", targets=[second], action="format",
                      number_at=18, text_at=54, follow="space")
    assert result.ok, result.to_json()
    xml = level_xml(session.entry("d1").document, 1)
    assert '<w:suff w:val="space"/>' in xml and '<w:ind w:left="1080" w:hanging="720"/>' in xml
    result = dispatch(box, session, "word_lists", doc="d1", targets=[second], action="format", follow="dash")
    assert not result.ok and result.error.field == "follow"


def test_word_styles_describe_says_the_list_level(box):
    session = box.session()
    session.open(make(), "levels.docx")

    def described(name):
        result = dispatch(box, session, "word_styles", doc="d1", action="describe", name=name)
        assert result.ok, result.to_json()
        return result.data.get("numbering")

    assert described("Chapter title") == {"list": 1, "level": 0, "own": True}
    assert described("Figure title") == {"list": 1, "level": 1, "own": False, "from": "Section title"}
    assert described("Aside title") is None
    result = dispatch(box, session, "word_styles", doc="d1", action="modify", name="Front title",
                      numbering={"list": "none"})
    assert result.ok and result.summary == "Modified style Front title: unlinked from its list level"
    assert described("Front title") == {"list": "none"}
    result = dispatch(box, session, "word_styles", doc="d1", action="modify", name="Aside title",
                      numbering={"list": "none"})
    assert result.ok and result.summary == "Modified style Aside title: numbers nothing, left as it was"
    assert result.warnings and not result.changed


def test_numbering_skips_level_names_its_paragraph_para(box):
    lone = [("Lone first", "SectionTitle"), ("Lone second", "SectionTitle")]
    document = Document.open(make(lone, section="%1.%2"))
    [entry] = numbering_skips_level(document)
    first = document.paragraphs()[0].id
    assert entry["para"] == first and list(entry)[0] == "para" and "first" not in entry
    session = box.session()
    session.open(make(lone, section="%1.%2"), "lone.docx")
    result = dispatch(box, session, "check", doc="d1", include=["layout"])
    [entry] = result.data["layout"]["numbering_skips_level"]
    assert entry["para"] == first and "first" not in entry
