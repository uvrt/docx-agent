#!/usr/bin/env python3
"""Write the fixtures E2's reader is tested on, ``tests/fixtures/generated/markdown/*.docx``.

The corpus fixtures lack most of what ``to_markdown`` has to say, so these three are written
here, by hand, with every construct on purpose.  They sit one directory below the shared
corpus (``tests/conftest.py`` globs ``*/*.docx``), so E0's and E1's suites and the Word
oracle keep their corpus; E2's tests read them beside it.

``constructs.docx`` (mode 15, English style names)
    a title; headings 1-3 and 7; direct bold, italic and strikethrough; the Strong,
    Emphasis and HTML Code character styles; text full of Markdown's special characters,
    a line starting with ``#`` and one with ``1.``, a leading tab, a line break, hidden
    text; a two-paragraph Quote; a two-line HTML Preformatted block; an empty paragraph
    with a bottom border; nested bullets, two adjacent bullet lists, a numbered list with
    a sub-level, one restarted at 5 by a ``w:startOverride``, a bullet from List Bullet's
    numbering; external and internal hyperlinks, a bookmark, two footnotes (one of two
    paragraphs) and an endnote; an inline and a floating picture and a text box; PAGE and
    REF fields; a GFM-shaped table (header row, a right-aligned column, a ``|`` in a cell),
    and one with a nested table and a cell of two paragraphs; inline text and checkbox
    content controls and a block one with a tag; a continuous section break, a page break;
    a header and a footer.

``review.docx`` (mode 15)
    revisions by two authors: insertion, deletion, a move with its range markers, a
    paragraph mark deleted (two paragraphs joined in the final view), a paragraph inserted
    whole and one deleted whole, an empty paragraph with a deleted mark before a table, a
    deleted and an inserted table row, a run-formatting and a paragraph-style change;
    comments: one with a reply (``commentsExtended``), one resolved, ``commentsIds`` giving
    durable ids to two of the three.

``dutch-template.docx`` (mode 15)
    a style set localised to Dutch: ids as Dutch Word writes them, and headings named
    "Kop 1" and "Kop 2" (found by their outline level) beside one named "heading 3" with no
    outline level (found by name); Titel, Citaat (named Quote), Lijstopsomteken and
    Lijstnummering (List Bullet, List Number), Nadruk and Zwaar (Emphasis, Strong),
    HTML Preformatted.

``blank.docx`` (mode 15)
    a minimal document every reader is held to -- the document defaults, the four styles
    every document has (Normal, Default Paragraph Font, Normal Table, No List), no
    numbering, notes or theme, one empty paragraph and an A4 section; until E6 the blank
    document ``insert_markdown`` was held to (now ``Document.new()``).

Deterministic, like ``make_fixtures.py``, whose helpers it uses.

    python tools/make_markdown_fixtures.py
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import make_fixtures as base  # noqa: E402

from make_fixtures import (A, DECL, MC, NS, PIC, REL, W, WML, WP, WP14, content_types, png,  # noqa: E402
                           relationships)

base.OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "generated" / "markdown"

W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
ALICE = 'w:author="Alice" w:date="2026-09-30T10:00:00Z"'
BOB = 'w:author="Bob" w:date="2026-10-01T11:30:00Z"'


def para(pid: str, inner: str = "", *, style: str | None = None, props: str = "", mark: str = "") -> str:
    """A paragraph: ``props`` are pPr children after the style, ``mark`` the mark's rPr."""
    ppr = (f'<w:pStyle w:val="{style}"/>' if style else "") + props + (f"<w:rPr>{mark}</w:rPr>" if mark else "")
    return (f'<w:p w14:paraId="{pid}" w14:textId="77777777">' + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "")
            + inner + "</w:p>")


def run(text: str, props: str = "") -> str:
    escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return f'<w:r>{f"<w:rPr>{props}</w:rPr>" if props else ""}<w:t xml:space="preserve">{escaped}</w:t></w:r>'


def deleted(text: str, props: str = "") -> str:
    return f'<w:r>{f"<w:rPr>{props}</w:rPr>" if props else ""}<w:delText xml:space="preserve">{text}</w:delText></w:r>'


def listed(pid: str, text: str, num: int, level: int = 0) -> str:
    return para(pid, run(text), style="ListParagraph",
                props=f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num}"/></w:numPr>')


def picture(doc_pr: int, rid: str, name: str, descr: str, *, floating: bool = False) -> str:
    graphic = (f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{PIC}"><pic:pic xmlns:pic="{PIC}">'
               f'<pic:nvPicPr><pic:cNvPr id="{doc_pr}" name="{name}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
               f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
               '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>'
               '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')
    ids = f'wp14:anchorId="2B3C4D{doc_pr:02X}" wp14:editId="3C4D5E{doc_pr:02X}"'
    if not floating:
        return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="{WP}" '
                f'xmlns:wp14="{WP14}" {ids}><wp:extent cx="457200" cy="457200"/>'
                f'<wp:docPr id="{doc_pr}" name="{name}" descr="{descr}"/>{graphic}</wp:inline></w:drawing></w:r>')
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            f'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" '
            f'xmlns:wp="{WP}" xmlns:wp14="{WP14}" {ids}><wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="column"><wp:posOffset>4572000</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '<wp:extent cx="457200" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="{doc_pr}" name="{name}" descr="{descr}"/>'
            f'{graphic}</wp:anchor></w:drawing></w:r>')


def text_box(doc_pr: int, inner: str) -> str:
    ids = f'wp14:anchorId="2B3C4D{doc_pr:02X}" wp14:editId="3C4D5E{doc_pr:02X}"'
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            f'relativeHeight="251660288" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" '
            f'xmlns:wp="{WP}" xmlns:wp14="{WP14}" {ids}><wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="column"><wp:posOffset>3657600</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '<wp:extent cx="1828800" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="{doc_pr}" name="Text Box {doc_pr}"/>'
            f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{WPS}"><wps:wsp xmlns:wps="{WPS}">'
            '<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/>'
            '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>'
            f'<wps:txbx><w:txbxContent>{inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>'
            '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def field(instruction: str, result: str) -> str:
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run(result)
            + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def cell(inner: str, width: int = 3000, props: str = "") -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{props}</w:tcPr>{inner}</w:tc>'


def row(pid: str, cells: str, props: str = "") -> str:
    return f'<w:tr w14:paraId="{pid}" w14:textId="77777777">' + (f"<w:trPr>{props}</w:trPr>" if props else "") \
        + cells + "</w:tr>"


def table(rows: str, columns: int, width: int = 3000) -> str:
    grid = "".join(f'<w:gridCol w:w="{width}"/>' for _ in range(columns))
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" '
            f'w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>')


SECTION = ('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
           'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/>')


# -- styles ----------------------------------------------------------------------------------


def english_styles(extra: str = "", *, lists: bool = True) -> str:
    def style(kind: str, sid: str, name: str, body: str = "", based: str | None = "Normal") -> str:
        basis = f'<w:basedOn w:val="{based}"/>' if based else ""
        return f'<w:style w:type="{kind}" w:styleId="{sid}"><w:name w:val="{name}"/>{basis}<w:qFormat/>{body}</w:style>'

    heading = ('<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="{lvl}"/></w:pPr>'
               '<w:rPr><w:b/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>')
    return (
        f'{DECL}<w:styles xmlns:w="{W}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" '
        'w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
        '<w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>'
        '</w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
        '<w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
        + style("paragraph", "Title", "Title", '<w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>')
        + style("paragraph", "Heading1", "heading 1", heading.format(lvl=0, size=32))
        + style("paragraph", "Heading2", "heading 2", heading.format(lvl=1, size=28))
        + style("paragraph", "Heading3", "heading 3", heading.format(lvl=2, size=24))
        + style("paragraph", "Heading7", "heading 7", heading.format(lvl=6, size=22))
        + style("paragraph", "Quote", "Quote", '<w:pPr><w:ind w:left="864" w:right="864"/></w:pPr>'
                '<w:rPr><w:i/><w:iCs/></w:rPr>')
        + style("paragraph", "HTMLPreformatted", "HTML Preformatted",
                '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
                '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>')
        + style("paragraph", "ListParagraph", "List Paragraph", '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>')
        + (style("paragraph", "ListBullet", "List Bullet",
                 '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:contextualSpacing/></w:pPr>') if lists else "")
        + style("paragraph", "Header", "header")
        + style("paragraph", "Footer", "footer")
        + style("paragraph", "FootnoteText", "footnote text",
                '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>')
        + style("paragraph", "EndnoteText", "endnote text",
                '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
                '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>')
        + style("character", "FootnoteReference", "footnote reference",
                '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>', based="DefaultParagraphFont")
        + style("character", "EndnoteReference", "endnote reference",
                '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>', based="DefaultParagraphFont")
        + style("character", "Strong", "Strong", "<w:rPr><w:b/><w:bCs/></w:rPr>", based="DefaultParagraphFont")
        + style("character", "Emphasis", "Emphasis", "<w:rPr><w:i/><w:iCs/></w:rPr>", based="DefaultParagraphFont")
        + style("character", "HTMLCode", "HTML Code",
                '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
                '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>', based="DefaultParagraphFont")
        + style("character", "Hyperlink", "Hyperlink",
                '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>', based="DefaultParagraphFont")
        + '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
        '<w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>'
        + extra + "</w:styles>"
    )


def numbering(abstracts: list[tuple[int, str, list[tuple[str, str]]]], nums: list[tuple[int, int, int | None]]) -> str:
    """``abstracts``: (id, nsid, [(numFmt, lvlText) per level]); ``nums``: (numId,
    abstractNumId, startOverride of level 0 or None)."""
    out = [f'{DECL}<w:numbering xmlns:w="{W}">']
    for abstract, nsid, levels in abstracts:
        out.append(f'<w:abstractNum w:abstractNumId="{abstract}"><w:nsid w:val="{nsid}"/>'
                   '<w:multiLevelType w:val="hybridMultilevel"/>')
        for level, (fmt, text) in enumerate(levels):
            font = '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>' \
                if fmt == "bullet" and text == "" else ""
            out.append(f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                       f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                       f'<w:pPr><w:ind w:left="{720 * (level + 1)}" w:hanging="360"/></w:pPr>{font}</w:lvl>')
        out.append("</w:abstractNum>")
    for num, abstract, start in nums:
        override = (f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="{start}"/></w:lvlOverride>'
                    if start is not None else "")
        out.append(f'<w:num w:numId="{num}"><w:abstractNumId w:val="{abstract}"/>{override}</w:num>')
    out.append("</w:numbering>")
    return "".join(out)


BULLETS = [("bullet", ""), ("bullet", "o"), ("bullet", "▪")]
NUMBERS = [("decimal", "%1."), ("lowerLetter", "%2."), ("lowerRoman", "%3.")]


def settings(notes: bool) -> str:
    separators = ('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
                  '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>') if notes else ""
    return (f'{DECL}<w:settings xmlns:w="{W}"><w:defaultTabStop w:val="708"/>{separators}'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def notes_part(kind: str, notes: list[str]) -> str:
    root = "footnotes" if kind == "footnote" else "endnotes"
    out = [f"{DECL}<w:{root} {NS}>",
           f'<w:{kind} w:type="separator" w:id="-1">' + para(f"{'1D' if kind == 'footnote' else '2E'}0000F1",
                                                              '<w:r><w:separator/></w:r>') + f"</w:{kind}>",
           f'<w:{kind} w:type="continuationSeparator" w:id="0">'
           + para(f"{'1D' if kind == 'footnote' else '2E'}0000F2", '<w:r><w:continuationSeparator/></w:r>')
           + f"</w:{kind}>"]
    for number, body in enumerate(notes, start=1):
        out.append(f'<w:{kind} w:id="{number}">{body}</w:{kind}>')
    out.append(f"</w:{root}>")
    return "".join(out)


def note_paragraph(pid: str, kind: str, text: str, *, first: bool = True) -> str:
    style = "FootnoteText" if kind == "footnote" else "EndnoteText"
    reference = "FootnoteReference" if kind == "footnote" else "EndnoteReference"
    mark = (f'<w:r><w:rPr><w:rStyle w:val="{reference}"/></w:rPr><w:{kind}Ref/></w:r>' + run(" ")) if first else ""
    return para(pid, mark + run(text), style=style)


# -- constructs.docx -------------------------------------------------------------------------


def constructs() -> dict:
    ids = iter(f"6B{n:06X}" for n in range(1, 400))

    def pid() -> str:
        return next(ids)

    reference = '<w:r><w:rPr><w:rStyle w:val="{style}"/></w:rPr><w:{kind}Reference w:id="{n}"/></w:r>'
    footnote = reference.format(style="FootnoteReference", kind="footnote", n="{n}")
    endnote = reference.format(style="EndnoteReference", kind="endnote", n="{n}")
    body = [
        para(pid(), run("Markdown constructs"), style="Title"),
        para(pid(), run("Headings and text"), style="Heading1"),
        para(pid(), run("Plain, ") + run("bold", "<w:b/>") + run(", ") + run("italic", "<w:i/>") + run(", ")
             + run("struck", "<w:strike/>") + run(", ") + run("strong", '<w:rStyle w:val="Strong"/>') + run(", ")
             + run("emphasised", '<w:rStyle w:val="Emphasis"/>') + run(" and ")
             + run("inline code", '<w:rStyle w:val="HTMLCode"/>') + run(".")),
        para(pid(), run("A second level"), style="Heading2"),
        para(pid(), run("A third level"), style="Heading3"),
        para(pid(), run("A seventh level"), style="Heading7"),
        para(pid(), run("Stars * and _underscores_, a [bracket], <tag>, `ticks`, & an entity &copy;, "
                        "a {++critic++} brace, a back\\slash and a pipe |.")),
        para(pid(), run("# not a heading")),
        para(pid(), run("1. not a list")),
        para(pid(), '<w:r><w:tab/><w:t>indented by a tab</w:t></w:r>'),
        para(pid(), '<w:r><w:t>Line one</w:t><w:br/><w:t>line two</w:t></w:r>' + run(" hidden", "<w:vanish/>")
             + run(".")),
        para(pid(), run("A quotation."), style="Quote"),
        para(pid(), run("Its second paragraph."), style="Quote"),
        para(pid(), run("def answer():"), style="HTMLPreformatted"),
        para(pid(), run("    return 42"), style="HTMLPreformatted"),
        para(pid(), props='<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>'),
        para(pid(), run("Lists"), style="Heading1"),
        listed(pid(), "Bullet one", 2), listed(pid(), "Nested bullet", 2, 1), listed(pid(), "Deeper bullet", 2, 2),
        listed(pid(), "Bullet two", 2),
        listed(pid(), "Another list", 5), listed(pid(), "Its second item", 5),
        para(pid(), run("Numbered:")),
        listed(pid(), "Step one", 3), listed(pid(), "Step two", 3), listed(pid(), "Sub-step", 3, 1),
        listed(pid(), "Step three", 3),
        para(pid(), run("Restarted at five:")),
        listed(pid(), "Five", 4), listed(pid(), "Six", 4),
        para(pid(), run("From the style"), style="ListBullet"),
        para(pid(), run("Links, notes and pictures"), style="Heading1"),
        para(pid(), run("An ") + '<w:hyperlink r:id="rIdSite" w:history="1">'
             + run("external link", '<w:rStyle w:val="Hyperlink"/>') + "</w:hyperlink>" + run(", an ")
             + '<w:hyperlink w:anchor="notes" w:history="1">' + run("internal one", '<w:rStyle w:val="Hyperlink"/>')
             + "</w:hyperlink>" + run(" and a note.") + footnote.format(n=1) + run(" Another.")
             + endnote.format(n=1)),
        para(pid(), '<w:bookmarkStart w:id="10" w:name="notes"/>' + run("Notes are written at the end.")
             + '<w:bookmarkEnd w:id="10"/>' + footnote.format(n=2)),
        para(pid(), run("A picture ") + picture(1, "rIdBlue", "Picture 1", "A blue square") + run(" inline.")),
        para(pid(), picture(2, "rIdRed", "Floating", "A red square", floating=True)
             + run("Text beside a floating picture.")),
        para(pid(), text_box(3, para("6B0000F1", run("Inside a text box."))) + run("Text beside a text box.")),
        para(pid(), run("Page ") + field("PAGE", "1") + run(", see ")
             + field("REF notes \\h", "Notes are written at the end.") + run(".")),
        para(pid(), run("Tables"), style="Heading1"),
        table(row(pid(), cell(para(pid(), run("Item"))) + cell(para(pid(), run("Count"))), "<w:tblHeader/>")
              + row(pid(), cell(para(pid(), run("Pens"))) + cell(para(pid(), run("12"), props='<w:jc w:val="right"/>')))
              + row(pid(), cell(para(pid(), run("a | b"))) + cell(para(pid(), run("500"), props='<w:jc w:val="right"/>'))),
              2),
        para(pid(), run("A table GFM cannot hold:")),
        table(row(pid(), cell(para(pid(), run("Outer")))
                  + cell(table(row(pid(), cell(para(pid(), run("Inner")), 1400) + cell(para(pid(), run("cells")), 1400)),
                               2, 1400) + para(pid()))) + row(pid(), cell(para(pid(), run("Two"))
                                                                         + para(pid(), run("paragraphs")))
                                                                    + cell(para(pid(), run("Plain")))), 2),
        para(pid(), run("Content controls:")),
        para(pid(), run("Name: ") + '<w:sdt><w:sdtPr><w:id w:val="301"/><w:text/></w:sdtPr><w:sdtContent>'
             + run("Grace Hopper") + "</w:sdtContent></w:sdt>" + run(", agreed: ")
             + '<w:sdt><w:sdtPr><w:id w:val="302"/><w14:checkbox><w14:checked w14:val="1"/>'
             '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" '
             'w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent>'
             + run("☒", '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/>')
             + "</w:sdtContent></w:sdt>"),
        '<w:sdt><w:sdtPr><w:alias w:val="Clause"/><w:tag w:val="clause"/><w:id w:val="303"/><w:richText/></w:sdtPr>'
        '<w:sdtContent>' + para(pid(), run("A clause in a control.")) + para(pid(), run("Its second paragraph."))
        + "</w:sdtContent></w:sdt>",
        para(pid(), run("Before a section break."),
             props=f'<w:sectPr><w:type w:val="continuous"/>{SECTION}</w:sectPr>'),
        para(pid(), '<w:r><w:br w:type="page"/></w:r>' + run("After a page break.")),
        para(pid(), run("The end.")),
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>'
        f'<w:footerReference w:type="default" r:id="rIdFooter"/>{SECTION}</w:sectPr>',
    ]
    document = f"{DECL}<w:document {NS}><w:body>{''.join(body)}</w:body></w:document>"
    header = f"{DECL}<w:hdr {NS}>" + para("7A0000A1", run("Constructs, the header"), style="Header") + "</w:hdr>"
    footer = (f"{DECL}<w:ftr {NS}>" + para("7A0000A2", run("Page ") + '<w:fldSimple w:instr=" PAGE ">' + run("1")
                                           + "</w:fldSimple>", style="Footer") + "</w:ftr>")
    footnotes = notes_part("footnote", [
        note_paragraph("1D000001", "footnote", "A first footnote."),
        note_paragraph("1D000002", "footnote", "A second footnote,")
        + note_paragraph("1D000003", "footnote", "in two paragraphs.", first=False)])
    endnotes = notes_part("endnote", [note_paragraph("2E000001", "endnote", "An endnote.")])
    numbering_xml = numbering(
        [(0, "4A5B6C01", BULLETS), (1, "4A5B6C02", NUMBERS)],
        [(1, 0, None), (2, 0, None), (3, 1, None), (4, 1, 5), (5, 0, None)])
    return {
        "[Content_Types].xml": content_types({
            "word/document.xml": WML + ".document.main+xml",
            "word/styles.xml": WML + ".styles+xml",
            "word/settings.xml": WML + ".settings+xml",
            "word/numbering.xml": WML + ".numbering+xml",
            "word/header1.xml": WML + ".header+xml",
            "word/footer1.xml": WML + ".footer+xml",
            "word/footnotes.xml": WML + ".footnotes+xml",
            "word/endnotes.xml": WML + ".endnotes+xml",
        }).replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/>'
                   '<Default Extension="xml"'),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([
            ("rIdStyles", REL + "styles", "styles.xml", False),
            ("rIdSettings", REL + "settings", "settings.xml", False),
            ("rIdNumbering", REL + "numbering", "numbering.xml", False),
            ("rIdHeader", REL + "header", "header1.xml", False),
            ("rIdFooter", REL + "footer", "footer1.xml", False),
            ("rIdFootnotes", REL + "footnotes", "footnotes.xml", False),
            ("rIdEndnotes", REL + "endnotes", "endnotes.xml", False),
            ("rIdBlue", REL + "image", "media/image1.png", False),
            ("rIdRed", REL + "image", "media/image2.png", False),
            ("rIdSite", REL + "hyperlink", "https://example.com/markdown", True),
        ]),
        "word/document.xml": document,
        "word/styles.xml": english_styles(),
        "word/settings.xml": settings(True),
        "word/numbering.xml": numbering_xml,
        "word/header1.xml": header,
        "word/footer1.xml": footer,
        "word/footnotes.xml": footnotes,
        "word/endnotes.xml": endnotes,
        "word/media/image1.png": png(),
        "word/media/image2.png": png(color=(200, 40, 40)),
    }


# -- review.docx -----------------------------------------------------------------------------


def review() -> dict:
    ids = iter(f"7C{n:06X}" for n in range(1, 400))

    def pid() -> str:
        return next(ids)

    body = [
        para(pid(), run("Review"), style="Heading1"),
        para(pid(), run("The fee is ") + f'<w:del w:id="101" {ALICE}>' + deleted("ten") + "</w:del>"
             + f'<w:ins w:id="102" {ALICE}>' + run("twelve") + "</w:ins>" + run(" euros.")),
        para(pid(), run("Bob ") + f'<w:ins w:id="103" {BOB}>' + run("carefully ") + "</w:ins>" + run("read it.")),
        para(pid(), run("Moved from here: ") + f'<w:moveFromRangeStart w:id="104" w:name="move1" {ALICE}/>'
             + f'<w:moveFrom w:id="105" {ALICE}>' + run("the moved words") + "</w:moveFrom>"
             + '<w:moveFromRangeEnd w:id="104"/>' + run(".")),
        para(pid(), run("Moved to here: ") + f'<w:moveToRangeStart w:id="106" w:name="move1" {ALICE}/>'
             + f'<w:moveTo w:id="107" {ALICE}>' + run("the moved words") + "</w:moveTo>"
             + '<w:moveToRangeEnd w:id="106"/>' + run(".")),
        para(pid(), run("This paragraph's mark is deleted, so it joins "), mark=f'<w:del w:id="108" {BOB}/>'),
        para(pid(), run("the next one in the final view.")),
        para(pid(), f'<w:ins w:id="109" {ALICE}>' + run("A whole inserted paragraph.") + "</w:ins>",
             mark=f'<w:ins w:id="110" {ALICE}/>'),
        para(pid(), f'<w:del w:id="111" {BOB}>' + deleted("A whole deleted paragraph.") + "</w:del>",
             mark=f'<w:del w:id="112" {BOB}/>'),
        para(pid(), run("Text after the deleted paragraph.")),
        para(pid(), run("Now ") + '<w:r><w:rPr><w:b/>'
             f'<w:rPrChange w:id="113" {BOB}><w:rPr/></w:rPrChange></w:rPr><w:t>bold</w:t></w:r>'
             + run(" by a revision.")),
        para(pid(), run("Made a heading"), style="Heading2",
             props=f'<w:pPrChange w:id="114" {ALICE}><w:pPr/></w:pPrChange>'),
        para(pid(), run("A ") + '<w:commentRangeStart w:id="1"/>' + run("disputed") + '<w:commentRangeEnd w:id="1"/>'
             + '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>'
             + '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="2"/></w:r>'
             + run(" clause.")),
        para(pid(), run("A ") + '<w:commentRangeStart w:id="3"/>' + run("settled") + '<w:commentRangeEnd w:id="3"/>'
             + '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="3"/></w:r>'
             + run(" point.")),
        para(pid(), mark=f'<w:del w:id="115" {BOB}/>'),
        table(row(pid(), cell(para(pid(), run("Name"))) + cell(para(pid(), run("Score"))), "<w:tblHeader/>")
              + row(pid(), cell(para(pid(), f'<w:del w:id="116" {BOB}>' + deleted("Ann") + "</w:del>"))
                    + cell(para(pid(), f'<w:del w:id="117" {BOB}>' + deleted("3") + "</w:del>")),
                    f'<w:del w:id="118" {BOB}/>')
              + row(pid(), cell(para(pid(), f'<w:ins w:id="119" {ALICE}>' + run("Ben") + "</w:ins>"))
                    + cell(para(pid(), f'<w:ins w:id="120" {ALICE}>' + run("4") + "</w:ins>")),
                    f'<w:ins w:id="121" {ALICE}/>'), 2),
        para(pid(), run("The last paragraph.")),
        f"<w:sectPr>{SECTION}</w:sectPr>",
    ]
    comment = ('<w:comment w:id="{id}" {who} w:initials="{initials}">'
               '<w:p w14:paraId="{para}" w14:textId="77777777"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>'
               '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>'
               + '{text}</w:p></w:comment>')
    comments = (f"{DECL}<w:comments {NS}>"
                + comment.format(id=1, who=ALICE, initials="A", para="5D000001", text=run("Is this right?"))
                + comment.format(id=2, who=BOB, initials="B", para="5D000002", text=run("Yes, it is."))
                + comment.format(id=3, who=BOB, initials="B", para="5D000003", text=run("Done."))
                + "</w:comments>")
    extended = (f'{DECL}<w15:commentsEx xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15">'
                '<w15:commentEx w15:paraId="5D000001" w15:done="0"/>'
                '<w15:commentEx w15:paraId="5D000002" w15:paraIdParent="5D000001" w15:done="0"/>'
                '<w15:commentEx w15:paraId="5D000003" w15:done="1"/></w15:commentsEx>')
    durable = (f'{DECL}<w16cid:commentsIds xmlns:w16cid="{W16CID}" xmlns:mc="{MC}" mc:Ignorable="w16cid">'
               '<w16cid:commentId w16cid:paraId="5D000001" w16cid:durableId="1A2B3C4D"/>'
               '<w16cid:commentId w16cid:paraId="5D000002" w16cid:durableId="2B3C4D5E"/></w16cid:commentsIds>')
    extra = ('<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/>'
             '<w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
             '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/>'
             '<w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>')
    return {
        "[Content_Types].xml": content_types({
            "word/document.xml": WML + ".document.main+xml",
            "word/styles.xml": WML + ".styles+xml",
            "word/settings.xml": WML + ".settings+xml",
            "word/comments.xml": WML + ".comments+xml",
            "word/commentsExtended.xml": WML + ".commentsExtended+xml",
            "word/commentsIds.xml": WML + ".commentsIds+xml",
        }),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([
            ("rIdStyles", REL + "styles", "styles.xml", False),
            ("rIdSettings", REL + "settings", "settings.xml", False),
            ("rIdComments", REL + "comments", "comments.xml", False),
            ("rIdCommentsEx", "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
             "commentsExtended.xml", False),
            ("rIdCommentsIds", "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
             "commentsIds.xml", False),
        ]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{''.join(body)}</w:body></w:document>",
        "word/styles.xml": english_styles(extra, lists=False),
        "word/settings.xml": settings(False),
        "word/comments.xml": comments,
        "word/commentsExtended.xml": extended,
        "word/commentsIds.xml": durable,
    }


# -- dutch-template.docx ---------------------------------------------------------------------


def dutch() -> dict:
    ids = iter(f"5E{n:06X}" for n in range(1, 100))

    def pid() -> str:
        return next(ids)

    def style(kind: str, sid: str, name: str, body: str = "", based: str | None = "Standaard") -> str:
        basis = f'<w:basedOn w:val="{based}"/>' if based else ""
        return f'<w:style w:type="{kind}" w:styleId="{sid}"><w:name w:val="{name}"/>{basis}<w:qFormat/>{body}</w:style>'

    heading = ('<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/>{outline}</w:pPr>'
               '<w:rPr><w:b/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>')
    styles = (
        f'{DECL}<w:styles xmlns:w="{W}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" '
        'w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
        '<w:lang w:val="nl-NL" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>'
        '</w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Standaard"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="character" w:default="1" w:styleId="Standaardalinea-lettertype">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
        '<w:style w:type="table" w:default="1" w:styleId="Standaardtabel"><w:name w:val="Normal Table"/>'
        '<w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
        + style("paragraph", "Titel", "Titel", '<w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>')
        + style("paragraph", "Kop1", "Kop 1", heading.format(outline='<w:outlineLvl w:val="0"/>', size=32))
        + style("paragraph", "Kop2", "Kop 2", heading.format(outline='<w:outlineLvl w:val="1"/>', size=28))
        + style("paragraph", "Kop3", "heading 3", heading.format(outline="", size=24))
        + style("paragraph", "Citaat", "Quote", '<w:rPr><w:i/><w:iCs/></w:rPr>')
        + style("paragraph", "HTML-voorafopgemaakt", "HTML Preformatted",
                '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/></w:rPr>')
        + style("paragraph", "Lijstopsomteken", "List Bullet",
                '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:contextualSpacing/></w:pPr>')
        + style("paragraph", "Lijstnummering", "List Number",
                '<w:pPr><w:numPr><w:numId w:val="2"/></w:numPr><w:contextualSpacing/></w:pPr>')
        + style("character", "Zwaar", "Strong", "<w:rPr><w:b/><w:bCs/></w:rPr>", based="Standaardalinea-lettertype")
        + style("character", "Nadruk", "Emphasis", "<w:rPr><w:i/><w:iCs/></w:rPr>", based="Standaardalinea-lettertype")
        + "</w:styles>")
    body = [
        para(pid(), run("Een Nederlands sjabloon"), style="Titel"),
        para(pid(), run("Inleiding"), style="Kop1"),
        para(pid(), run("Tekst met ") + run("nadruk", '<w:rStyle w:val="Nadruk"/>') + run(" en ")
             + run("zwaar", '<w:rStyle w:val="Zwaar"/>') + run(".")),
        para(pid(), run("Achtergrond"), style="Kop2"),
        para(pid(), run("Details"), style="Kop3"),
        para(pid(), run("Een citaat."), style="Citaat"),
        para(pid(), run("Eerste punt"), style="Lijstopsomteken"),
        para(pid(), run("Tweede punt"), style="Lijstopsomteken"),
        para(pid(), run("Stap een"), style="Lijstnummering"),
        para(pid(), run("Stap twee"), style="Lijstnummering"),
        para(pid(), run("x = 1"), style="HTML-voorafopgemaakt"),
        f"<w:sectPr>{SECTION}</w:sectPr>",
    ]
    numbering_xml = numbering([(0, "6C7D8E01", BULLETS[:1]), (1, "6C7D8E02", NUMBERS[:1])], [(1, 0, None), (2, 1, None)])
    return {
        "[Content_Types].xml": content_types({
            "word/document.xml": WML + ".document.main+xml",
            "word/styles.xml": WML + ".styles+xml",
            "word/settings.xml": WML + ".settings+xml",
            "word/numbering.xml": WML + ".numbering+xml",
        }),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([
            ("rIdStyles", REL + "styles", "styles.xml", False),
            ("rIdSettings", REL + "settings", "settings.xml", False),
            ("rIdNumbering", REL + "numbering", "numbering.xml", False),
        ]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{''.join(body)}</w:body></w:document>",
        "word/styles.xml": styles,
        "word/settings.xml": settings(False),
        "word/numbering.xml": numbering_xml,
    }


def blank() -> dict[str, str]:
    styles = (
        f'{DECL}<w:styles xmlns:w="{W}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" '
        'w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
        '<w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>'
        '</w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
        '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
        "</w:styles>")
    body = para("0B1A0001") + f"<w:sectPr>{SECTION}</w:sectPr>"
    return {
        "[Content_Types].xml": content_types({
            "word/document.xml": WML + ".document.main+xml",
            "word/styles.xml": WML + ".styles+xml",
            "word/settings.xml": WML + ".settings+xml",
        }),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([
            ("rIdStyles", REL + "styles", "styles.xml", False),
            ("rIdSettings", REL + "settings", "settings.xml", False),
        ]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": styles,
        "word/settings.xml": settings(False),
    }


def main() -> None:
    for name, parts in (("constructs.docx", constructs()), ("review.docx", review()),
                        ("dutch-template.docx", dutch()), ("blank.docx", blank())):
        print(base.write(name, parts))


if __name__ == "__main__":
    main()
