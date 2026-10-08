#!/usr/bin/env python3
"""Measure what Word does for E6's authoring: a new document, a document from a template,
content copied between documents whose styles conflict, and Convert.

Through ``tests/oracle.py``'s machine-wide lock, Word, by AppleScript:

* ``new``, ``new-typed``: makes a new blank document (``make new document``: File > New,
  on this Word's Normal template) and saves it as it is, and again with a line typed;
* ``template-dotx``, ``template-docx``: makes a document from a template written here --
  a ``.dotx``, and the same content as a ``.docx`` -- (``create new document attached
  template``: File > New from a template) and saves it;
* ``as-template``: saves a document written here as a ``.dotx``;
* ``copy-formatted``: copies a source document written here into a destination written
  here, whose styles, list, note, comment, bookmark, picture, content control and paraIds
  collide with the source's, through ``formatted text`` (Word's own copy of formatted
  content between ranges, which does not use the clipboard); ``copy-formatted-tracked``
  the same with tracking on; ``copy-insert-file``: ``insert file`` of the source;
* ``paste-keep``, ``paste-destination``, ``paste-merge``, ``paste-default``: the clipboard's
  paste with each of Word's paste options.  **Only when the clipboard is empty** before the
  run, and it is emptied again after: the clipboard is the user's, and its contents could
  not otherwise be put back as they were;
* ``convert-*``: ``upgrade`` (Word's File > Info > Convert) on documents in modes 11, 12
  and 14 -- two written here with legacy compatibility options, and three corpus fixtures;
  ``resave-*``: two of them only opened and saved, to tell what Convert changes from what any
  save does.

What it wrote is read back into ``tests/observations/e6-word.json`` as facts: settings and
compatibility options, which parts and relationships a document has, the body's blocks
(normalised: rsids dropped, dates and authors replaced), style definitions by name for the
documents written here -- **never** the new document's own styles, theme or font table,
which are Word's Normal template: those are compared, part by part, with what
``Document.new()`` writes, and only the comparison is recorded.

    python tools/e6_probe.py                # every probe, the observations written
    python tools/e6_probe.py new paste-keep # some probes, printed

Nothing Word writes is committed: only these facts.
"""

from __future__ import annotations

import hashlib
import io
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "src"))

import oracle  # noqa: E402
from make_fixtures import DECL, NS, REL, WML, content_types, png, relationships  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "e6-word.json"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
FIXTURES = ROOT / "tests" / "fixtures"
DRAWING_NS = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
              'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
              'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
CT_APP = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
TYPED = ("Word typed this line into its own new document, long enough to wrap across the "
         "width of the page at least once, so that where the line breaks says something.")


# -- documents written here -------------------------------------------------------------------


def zipped(parts: dict[str, str | bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=(2026, 10, 4, 12, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data.encode("utf-8") if isinstance(data, str) else data)
    return buffer.getvalue()


def para(pid: str | None, inner: str = "", *, style: str | None = None, props: str = "") -> str:
    ids = f' w14:paraId="{pid}" w14:textId="77777777"' if pid else ""
    ppr = (f'<w:pStyle w:val="{style}"/>' if style else "") + props
    return f"<w:p{ids}>" + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + inner + "</w:p>"


def run(text: str, props: str = "") -> str:
    return f'<w:r>{f"<w:rPr>{props}</w:rPr>" if props else ""}<w:t xml:space="preserve">{text}</w:t></w:r>'


def style(kind: str, sid: str, name: str, *, based: str | None = None, nxt: str | None = None,
          ppr: str = "", rpr: str = "", custom: bool = False, default: bool = False) -> str:
    flags = (' w:customStyle="1"' if custom else "") + (' w:default="1"' if default else "")
    return (f'<w:style w:type="{kind}"{flags} w:styleId="{sid}"><w:name w:val="{name}"/>'
            + (f'<w:basedOn w:val="{based}"/>' if based else "") + (f'<w:next w:val="{nxt}"/>' if nxt else "")
            + "<w:qFormat/>" + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + (f"<w:rPr>{rpr}</w:rPr>" if rpr else "")
            + "</w:style>")


def styles_xml(face: str, size: int, extra: str) -> str:
    return (f'{DECL}<w:styles xmlns:w="{W}">'
            f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:eastAsia="{face}" '
            f'w:cs="{face}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/><w:lang w:val="en-GB" w:eastAsia="en-US" '
            'w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" '
            'w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
            '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name '
            'w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
            '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
            '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" '
            'w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
            '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
            '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
            + extra + "</w:styles>")


def numbering_xml(nsid: str, fmt: str = "decimal") -> str:
    text = "%1." if fmt == "decimal" else "•"
    return (f'{DECL}<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="{nsid}"/>'
            '<w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/>'
            f'<w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" '
            'w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/>'
            "</w:num></w:numbering>")


def settings_xml(mode: int | None, *, compat: str = "", extra: str = "", notes: bool = False) -> str:
    separators = ('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
                  '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>') if notes else ""
    mode_xml = (f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" '
                f'w:val="{mode}"/>') if mode else ""
    compat_xml = f"<w:compat>{compat}{mode_xml}</w:compat>" if (compat or mode_xml) else ""
    return f'{DECL}<w:settings xmlns:w="{W}"><w:defaultTabStop w:val="720"/>{extra}{separators}{compat_xml}</w:settings>'


def notes_xml(kind: str, text: str, pid: str) -> str:
    tag = kind[:-1]
    sep = (f'<w:{tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
           f'w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:{tag}>'
           f'<w:{tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
           f'w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:{tag}>')
    return (f'{DECL}<w:{kind} {NS}>{sep}<w:{tag} w:id="1"><w:p w14:paraId="{pid}" w14:textId="77777777"><w:r>'
            f'<w:{tag}Ref/></w:r><w:r><w:t xml:space="preserve"> {text}</w:t></w:r></w:p></w:{tag}></w:{kind}>')


def comments_xml(text: str, pid: str) -> str:
    return (f'{DECL}<w:comments {NS}><w:comment w:id="0" w:author="Probe" w:date="2026-10-01T09:00:00Z" '
            f'w:initials="P"><w:p w14:paraId="{pid}" w14:textId="77777777"><w:r><w:annotationRef/></w:r>'
            f'<w:r><w:t>{text}</w:t></w:r></w:p></w:comment></w:comments>')


def picture(rid: str, doc_pr: int, name: str) -> str:
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="457200" '
            f'cy="457200"/><wp:docPr id="{doc_pr}" name="{name}"/><a:graphic><a:graphicData '
            'uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
            f'<pic:cNvPr id="{doc_pr}" name="{name}.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
            f'<a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm>'
            '<a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
            "</a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>")


SECTION = ('<w:sectPr>{refs}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" '
           'w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/>'
           '<w:docGrid w:linePitch="360"/></w:sectPr>')


def package(body: str, *, main_type: str = WML + ".document.main+xml", styles: str, settings: str,
            parts: dict[str, str | bytes] = (), types: dict[str, str] = (), rels: list = (),
            defaults: dict[str, str] = (), core: str | None = None, app: str | None = None,
            namespaces: str = "") -> bytes:
    types = {"word/document.xml": main_type, "word/styles.xml": WML + ".styles+xml",
             "word/settings.xml": WML + ".settings+xml", **dict(types)}
    root_rels = [("rId1", REL + "officeDocument", "word/document.xml", False)]
    out: dict[str, str | bytes] = {}
    if core is not None:
        types["docProps/core.xml"] = CT_CORE
        root_rels.append(("rId2", REL_CORE, "docProps/core.xml", False))
        out["docProps/core.xml"] = core
    if app is not None:
        types["docProps/app.xml"] = CT_APP
        root_rels.append(("rId3", REL + "extended-properties", "docProps/app.xml", False))
        out["docProps/app.xml"] = app
    content = content_types(types)
    for extension, kind in dict(defaults).items():
        content = content.replace("<Default ", f'<Default Extension="{extension}" ContentType="{kind}"/><Default ', 1)
    out.update({
        "[Content_Types].xml": content,
        "_rels/.rels": relationships(root_rels),
        "word/_rels/document.xml.rels": relationships([("rIdStyles", REL + "styles", "styles.xml", False),
                                                       ("rIdSettings", REL + "settings", "settings.xml", False),
                                                       *rels]),
        "word/document.xml": f"{DECL}<w:document {NS} {namespaces}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": styles,
        "word/settings.xml": settings,
        **dict(parts),
    })
    return zipped(out)


def core_xml(title: str, creator: str) -> str:
    return (f'{DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/'
            'core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f"<dc:title>{title}</dc:title><dc:subject>Template subject</dc:subject><dc:creator>{creator}</dc:creator>"
            "<cp:keywords>template keywords</cp:keywords><cp:lastModifiedBy>Template Editor</cp:lastModifiedBy>"
            "<cp:revision>7</cp:revision>"
            '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-02T03:04:00Z</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2021-05-06T07:08:00Z</dcterms:modified>'
            "</cp:coreProperties>")


def app_xml() -> str:
    return (f'{DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm'
            "</Template><TotalTime>12</TotalTime><Pages>1</Pages><Words>9</Words><Application>Microsoft Office Word"
            "</Application><Company>Template Company</Company></Properties>")


#: The template's styles: a redefined heading 1, a style of its own of each kind.
TEMPLATE_STYLES = (
    style("paragraph", "Heading1", "heading 1", based="Normal", nxt="Normal",
          ppr='<w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/>',
          rpr='<w:b/><w:color w:val="C00000"/><w:sz w:val="36"/><w:szCs w:val="36"/>')
    + style("paragraph", "BrandBody", "Brand Body", based="Normal", custom=True,
            rpr='<w:color w:val="1F6F3F"/>')
    + style("character", "BrandMark", "Brand Mark", based="DefaultParagraphFont", custom=True,
            rpr='<w:b/><w:color w:val="7030A0"/>')
    + style("paragraph", "Header", "header", based="Normal")
    + style("paragraph", "Footer", "footer", based="Normal"))


def template(kind: str) -> bytes:
    main = WML + (".template.main+xml" if kind == "dotx" else ".document.main+xml")
    body = (para("6A000001", run("Template heading"), style="Heading1")
            + para("6A000002", run("Template body, ") + run("marked", '<w:rStyle w:val="BrandMark"/>'),
                   style="BrandBody")
            + para("6A000003", run("Template list item"),
                   props='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
            + para("6A000004", run("Template last paragraph."))
            + SECTION.format(refs='<w:headerReference w:type="default" r:id="rIdHeader"/>'
                                  '<w:footerReference w:type="default" r:id="rIdFooter"/>')
            .replace('w:w="11906" w:h="16838"', 'w:w="12240" w:h="15840"').replace('w:top="1440"', 'w:top="1000"'))
    header = (f'{DECL}<w:hdr {NS}>' + para("6A000010", run("Template header"), style="Header") + "</w:hdr>")
    footer = (f'{DECL}<w:ftr {NS}>' + para("6A000011", run("Template footer"), style="Footer") + "</w:ftr>")
    return package(body, main_type=main, styles=styles_xml("Calibri", 22, TEMPLATE_STYLES),
                   settings=settings_xml(15, extra='<w:evenAndOddHeaders/>'),
                   parts={"word/numbering.xml": numbering_xml("1A2B3C4D", "bullet"), "word/header1.xml": header,
                          "word/footer1.xml": footer},
                   types={"word/numbering.xml": WML + ".numbering+xml", "word/header1.xml": WML + ".header+xml",
                          "word/footer1.xml": WML + ".footer+xml"},
                   rels=[("rIdNumbering", REL + "numbering", "numbering.xml", False),
                         ("rIdHeader", REL + "header", "header1.xml", False),
                         ("rIdFooter", REL + "footer", "footer1.xml", False)],
                   core=core_xml("Template Title", "Template Author"), app=app_xml())


# -- copying between documents ------------------------------------------------------------------

#: Source and destination define the same names differently, and each has one of its own.
SOURCE_STYLES = (
    style("paragraph", "Heading1", "heading 1", based="Normal", nxt="Normal",
          ppr='<w:keepNext/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/>',
          rpr='<w:b/><w:color w:val="C00000"/><w:sz w:val="36"/><w:szCs w:val="36"/>')
    + style("paragraph", "Brand", "Brand", based="Normal", custom=True, rpr='<w:b/><w:color w:val="00B050"/>')
    + style("paragraph", "SourceOnly", "Source Only", based="Normal", custom=True,
            rpr='<w:color w:val="ED7D31"/>')
    + style("character", "Accent", "Accent", based="DefaultParagraphFont", custom=True,
            rpr='<w:color w:val="FF0000"/><w:u w:val="single"/>')
    + style("paragraph", "FootnoteText", "footnote text", based="Normal",
            ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', rpr='<w:sz w:val="20"/>')
    + style("character", "FootnoteReference", "footnote reference", based="DefaultParagraphFont",
            rpr='<w:vertAlign w:val="superscript"/>')
    + style("character", "Hyperlink", "Hyperlink", based="DefaultParagraphFont",
            rpr='<w:color w:val="0563C1"/><w:u w:val="single"/>'))
DEST_STYLES = (
    style("paragraph", "Heading1", "heading 1", based="Normal", nxt="Normal",
          ppr='<w:keepNext/><w:spacing w:before="480" w:after="0"/><w:outlineLvl w:val="0"/>',
          rpr='<w:i/><w:color w:val="1F4E79"/><w:sz w:val="28"/><w:szCs w:val="28"/>')
    + style("paragraph", "Brand", "Brand", based="Normal", custom=True, rpr='<w:i/><w:color w:val="7030A0"/>')
    + style("paragraph", "DestOnly", "Destination Only", based="Normal", custom=True,
            rpr='<w:color w:val="2E75B6"/>')
    + style("character", "Accent", "Accent", based="DefaultParagraphFont", custom=True,
            rpr='<w:b/><w:color w:val="0000FF"/>')
    + style("paragraph", "FootnoteText", "footnote text", based="Normal",
            ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', rpr='<w:sz w:val="20"/>')
    + style("character", "FootnoteReference", "footnote reference", based="DefaultParagraphFont",
            rpr='<w:vertAlign w:val="superscript"/>'))

PURPLE = png(4, 4, (112, 48, 160))


def copy_document(role: str) -> bytes:
    """The source (``role`` "source") or the destination of a copy.  They collide on: the
    styles named above, list instance 1 (both decimal lists over abstract 0), footnote 1,
    comment 0, the bookmark "Shared", picture ``docPr`` 1 and the same picture's bytes,
    content control id 5, and paraIds 7B000002-7B000004."""
    source = role == "source"
    word = "Source" if source else "Destination"
    listed = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
    body = (para("7B000001", run(f"{word} heading"), style="Heading1")
            + para("7B000002", run(f"{word} brand paragraph with ") + run("accented", '<w:rStyle w:val="Accent"/>')
                   + run(" words."), style="Brand")
            + para("7B000003", run(f"{word} list item one"), props=listed)
            + para("7B000004", run(f"{word} list item two"), props=listed)
            + para("7B000005", run(f"{word} note") + '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                   '<w:footnoteReference w:id="1"/></w:r>' + run(" and ") + '<w:commentRangeStart w:id="0"/>'
                   + run("comment") + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>'
                   + run(" and ") + '<w:bookmarkStart w:id="0" w:name="Shared"/>' + run("bookmark")
                   + '<w:bookmarkEnd w:id="0"/>' + run("."))
            + para("7B000006", run(f"{word} picture ") + picture("rIdImg", 1, "Picture 1"))
            + para("7B000007", run(f"{word} control: ") + '<w:sdt><w:sdtPr><w:id w:val="5"/><w:text/></w:sdtPr>'
                   f'<w:sdtContent>{run(word + " value")}</w:sdtContent></w:sdt>'))
    if source:
        body += (para("7B000008", run("Source only style."), style="SourceOnly")
                 + para("7B000009", run("Source link: ") + '<w:hyperlink r:id="rIdLink" w:history="1">'
                        + run("example", '<w:rStyle w:val="Hyperlink"/>') + "</w:hyperlink>")
                 + '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>'
                 '<w:gridCol w:w="3000"/></w:tblGrid><w:tr w14:paraId="7B00000A" w14:textId="77777777">'
                 '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' + para("7B00000B", run("Cell one"))
                 + '</w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' + para("7B00000C", run("Cell two"))
                 + "</w:tc></w:tr></w:tbl>" + para("7B00000D", run("Source last paragraph.")))
    else:
        # The copy goes into the empty last paragraph, as a paste at a document's end does.
        body += (para("7B000008", run("Destination only style."), style="DestOnly")
                 + para("7B000009", run("Destination end.")) + para("7B00000F"))
    body += SECTION.format(refs="")
    rels = [("rIdNumbering", REL + "numbering", "numbering.xml", False),
            ("rIdFootnotes", REL + "footnotes", "footnotes.xml", False),
            ("rIdEndnotes", REL + "endnotes", "endnotes.xml", False),
            ("rIdComments", REL + "comments", "comments.xml", False),
            ("rIdImg", REL + "image", "media/image1.png", False)]
    if source:
        rels.append(("rIdLink", REL + "hyperlink", "https://example.com/source", True))
    face, size = ("Georgia", 22) if source else ("Arial", 24)
    return package(
        body, styles=styles_xml(face, size, SOURCE_STYLES if source else DEST_STYLES),
        settings=settings_xml(15, notes=True),
        parts={"word/numbering.xml": numbering_xml("5A5A0001" if source else "6B6B0002"),
               "word/footnotes.xml": notes_xml("footnotes", f"{word} footnote.", "7B0000F1"),
               "word/endnotes.xml": notes_xml("endnotes", f"{word} endnote.", "7B0000F2"),
               "word/comments.xml": comments_xml(f"{word} comment", "7B0000F3"),
               "word/media/image1.png": PURPLE},
        types={"word/numbering.xml": WML + ".numbering+xml", "word/footnotes.xml": WML + ".footnotes+xml",
               "word/endnotes.xml": WML + ".endnotes+xml", "word/comments.xml": WML + ".comments+xml"},
        defaults={"png": "image/png"}, rels=rels, namespaces=DRAWING_NS)


# -- Convert ------------------------------------------------------------------------------------

#: Legacy options a Word 2003 (mode 11) document carries; which of them Convert keeps.
LEGACY_COMPAT = ("<w:doNotExpandShiftReturn/><w:useFELayout/><w:doNotUseHTMLParagraphAutoSpacing/>"
                 "<w:balanceSingleByteDoubleByteWidth/><w:ulTrailSpace/><w:doNotUseIndentAsNumberingTabStop/>"
                 "<w:adjustLineHeightInTable/><w:doNotBreakWrappedTables/><w:footnoteLayoutLikeWW8/>"
                 "<w:selectFldWithFirstOrLastChar/><w:useWord2002TableStyleRules/><w:growAutofit/>")


def legacy_document(mode: int = 11, compat: str = LEGACY_COMPAT) -> bytes:
    body = (para("7C000001", run(f"A document in compatibility mode {mode}."))
            + para("7C000002", run("Its compatibility options are the ones such a file carries."))
            + SECTION.format(refs=""))
    return package(body, styles=styles_xml("Times New Roman", 24, ""), settings=settings_xml(mode, compat=compat))


#: Every option of ``w:compat`` (ECMA-376's ``CT_Compat``), in schema order: which of them
#: Convert keeps, and which any save keeps.
ALL_COMPAT = (
    "useSingleBorderforContiguousCells", "wpJustification", "noTabHangInd", "noLeading", "spaceForUL",
    "noColumnBalance", "balanceSingleByteDoubleByteWidth", "noExtraLineSpacing", "doNotLeaveBackslashAlone",
    "ulTrailSpace", "doNotExpandShiftReturn", "spacingInWholePoints", "lineWrapLikeWord6",
    "printBodyTextBeforeHeader", "printColBlack", "wpSpaceWidth", "showBreaksInFrames", "subFontBySize",
    "suppressBottomSpacing", "suppressTopSpacing", "suppressSpacingAtTopOfPage", "suppressTopSpacingWP",
    "suppressSpBfAfterPgBrk", "swapBordersFacingPages", "convMailMergeEsc", "truncateFontHeightsLikeWP6",
    "mwSmallCaps", "usePrinterMetrics", "doNotSuppressParagraphBorders", "wrapTrailSpaces",
    "footnoteLayoutLikeWW8", "shapeLayoutLikeWW8", "alignTablesRowByRow", "forgetLastTabAlignment",
    "adjustLineHeightInTable", "autoSpaceLikeWord95", "noSpaceRaiseLower", "doNotUseHTMLParagraphAutoSpacing",
    "layoutRawTableWidth", "layoutTableRowsApart", "useWord97LineBreakRules", "doNotBreakWrappedTables",
    "doNotSnapToGridInCell", "selectFldWithFirstOrLastChar", "applyBreakingRules", "doNotWrapTextWithPunct",
    "doNotUseEastAsianBreakRules", "useWord2002TableStyleRules", "growAutofit", "useFELayout",
    "useNormalStyleForList", "doNotUseIndentAsNumberingTabStop", "useAltKinsokuLineBreakRules",
    "allowSpaceOfSameStyleInTable", "doNotSuppressIndentation", "doNotAutofitConstrainedTables",
    "autofitToFirstFixedWidthCell", "underlineTabInNumList", "displayHangulFixedWidth", "splitPgBreakAndParaMark",
    "doNotVertAlignCellWithSp", "doNotBreakConstrainedForcedTable", "doNotVertAlignInTxbx", "useAnsiKerningPairs",
    "cachedColBalance",
)
EVERY_COMPAT = "".join(f"<w:{name}/>" for name in ALL_COMPAT)


#: A Word 2010 document's options, East Asian layout among them, without an East Asian
#: theme language.
MODE14_COMPAT = "<w:useFELayout/>"


# -- scripts -------------------------------------------------------------------------------------

_HEAD = '''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
'''
_TAIL = '''
      close every document saving no
      quit saving no
    end tell
  end timeout
end run
'''


def _script(body: str, *, save: str = "save as d file name outputPath file format format document") -> str:
    return _HEAD + body + "\n      " + save + _TAIL


NEW = _script("set d to make new document")
NEW_TYPED = _script(f'set d to make new document\n      insert text "{TYPED}" at end of text object of d')
FROM_TEMPLATE = _script("set d to create new document attached template (item 3 of argv)")
AS_TEMPLATE = _script("open (POSIX file inputPath)\n      set d to active document",
                      save="save as d file name outputPath file format format template")
CONVERT = _script("open (POSIX file inputPath)\n      set d to active document\n      upgrade d")
RESAVE = _script("open (POSIX file inputPath)\n      set d to active document")

#: Opens the destination, then the source -- each held by its name: ``active document`` is a
#: reference that follows the active window -- and ``r`` is the destination's end (before
#: its last paragraph mark: the start of its empty last paragraph).  ``{copy}`` puts the source's content there.
_COPY = '''open (POSIX file inputPath)
      set d to document (name of active document)
      open (POSIX file (item 3 of argv))
      set s to document (name of active document)
      {track}
      set e to (end of content of text object of d) - 1
      set r to create range d start e end e
      {copy}'''


def copy_script(copy: str, track: bool = False) -> str:
    return _script(_COPY.replace("{track}", "set track revisions of d to true" if track else "")
                   .replace("{copy}", copy))


COPIES = {
    "copy-formatted": copy_script("set formatted text of r to formatted text of (text object of s)"),
    "copy-formatted-tracked": copy_script("set formatted text of r to formatted text of (text object of s)",
                                          track=True),
    "copy-insert-file": copy_script("insert file at r file name (item 3 of argv)"),
}
PASTES = {
    "paste-default": copy_script("copy object (text object of s)\n      paste object r"),
    "paste-keep": copy_script("copy object (text object of s)\n      paste and format r type "
                              "format original formatting"),
    "paste-destination": copy_script("copy object (text object of s)\n      paste and format r type "
                                     "use destination styles recovery"),
    "paste-merge": copy_script("copy object (text object of s)\n      paste and format r type "
                               "format surrounding formatting with emphasis"),
}


# -- the clipboard: only an empty one is used, and it is left empty --------------------------


def clipboard_empty() -> bool:
    info = subprocess.run(["osascript", "-e", "clipboard info"], capture_output=True, text=True)
    return info.returncode == 0 and not info.stdout.strip()


def empty_clipboard() -> None:
    subprocess.run(["osascript", "-l", "JavaScript", "-e",
                    'ObjC.import("AppKit"); $.NSPasteboard.generalPasteboard.clearContents'],
                   capture_output=True, check=False)


# -- running -------------------------------------------------------------------------------------


def stage(name: str, data: bytes, suffix: str) -> Path:
    """A second input (the source, the template) beside the oracle's staging, by content."""
    path = oracle.ORACLE_DIR / f"e6-{name}-{hashlib.sha256(data).hexdigest()[:12]}{suffix}"
    if not path.exists():
        path.write_bytes(data)
    return path


def corpus(name: str) -> bytes:
    return (FIXTURES / name).read_bytes()


def runs() -> dict[str, tuple]:
    """name -> (script, input document or None, extra-argument maker)."""
    source = lambda: [str(stage("source", copy_document("source"), ".docx"))]  # noqa: E731
    out = {
        "new": (NEW, None, None),
        "new-typed": (NEW_TYPED, None, None),
        "template-dotx": (FROM_TEMPLATE, None, lambda: [str(stage("template", template("dotx"), ".dotx"))]),
        "template-docx": (FROM_TEMPLATE, None, lambda: [str(stage("template", template("docx"), ".docx"))]),
        "as-template": (AS_TEMPLATE, template("docx"), None),
        "convert-11": (CONVERT, legacy_document(), None),
        "resave-11": (RESAVE, legacy_document(), None),
        "convert-all-11": (CONVERT, legacy_document(11, EVERY_COMPAT), None),
        "resave-all-11": (RESAVE, legacy_document(11, EVERY_COMPAT), None),
        "convert-all-12": (CONVERT, legacy_document(12, EVERY_COMPAT), None),
        "convert-all-14": (CONVERT, legacy_document(14, EVERY_COMPAT), None),
        "resave-all-14": (RESAVE, legacy_document(14, EVERY_COMPAT), None),
        "convert-14-fe": (CONVERT, legacy_document(14, MODE14_COMPAT), None),
        "resave-14-sample": (RESAVE, corpus("samplelib/sample-simple.docx"), None),
        "convert-12": (CONVERT, corpus("docx2svg/layout-sweep.docx"), None),
        "convert-14": (CONVERT, corpus("generated/mode14.docx"), None),
        "convert-14-sample": (CONVERT, corpus("samplelib/sample-simple.docx"), None),
    }
    for name, script in COPIES.items():
        out[name] = (script, copy_document("destination"), source)
    for name, script in PASTES.items():
        out[name] = (script, copy_document("destination"), source)
    return out


def measure(names: list[str]) -> dict[str, bytes]:
    plan = runs()
    saved: dict[str, bytes] = {}
    with oracle.session() as session:
        for name in names:
            script, document, extra = plan[name]
            clipboard = name.startswith("paste-")
            if clipboard and not clipboard_empty():
                print(f"{name}: skipped, the clipboard is not empty")
                continue
            try:
                outcome = session.run_script(script, document, name=f"e6-{name}", tag="saved",
                                             args=extra() if extra else [], timeout=200)
            finally:
                if clipboard:
                    empty_clipboard()
            if not outcome:
                print(f"{name}: Word did not save: {outcome.outcome} {outcome.detail}")
                continue
            if outcome.detail:
                print(f"{name}: {outcome.detail}")
            saved[name] = outcome.path.read_bytes()
    return saved


# -- reading -------------------------------------------------------------------------------------

_DROP = re.compile(r' w:rsid\w*="[^"]*"')


def normalise(xml: str) -> str:
    xml = re.sub(r' xmlns:\w+="[^"]*"', "", xml)
    xml = _DROP.sub("", xml)
    xml = re.sub(r'w:author="[^"]*"', 'w:author="{author}"', xml)
    xml = re.sub(r'w:date="[^"]*"', 'w:date="{date}"', xml)
    xml = re.sub(r'w:initials="[^"]*"', 'w:initials="{initials}"', xml)
    xml = re.sub(r"<w:rsid w:val=\"[^\"]*\"/>", "", xml)
    return xml


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def blocks(files: dict[str, bytes], part: str = "word/document.xml") -> list[str]:
    root = etree.fromstring(files[part])
    container = root.find(_W + "body")
    if container is None:
        container = root
    return [normalise(etree.tostring(child, encoding="unicode")) for child in container]


def settings_facts(files: dict[str, bytes]) -> str:
    text = normalise(files["word/settings.xml"].decode("utf-8"))
    text = re.sub(r"<w:rsids>.*?</w:rsids>|<w15:docId [^>]*/>|<w14:docId [^>]*/>", "", text)
    return re.sub(r"^<\?xml[^>]*>\s*", "", text)


def style_facts(files: dict[str, bytes]) -> dict[str, str]:
    root = etree.fromstring(files["word/styles.xml"])
    out = {}
    for node in root.findall(_W + "style"):
        name = node.find(_W + "name").get(_W + "val")
        out[name] = normalise(etree.tostring(node, encoding="unicode"))
    defaults = root.find(_W + "docDefaults")
    if defaults is not None:
        out["(docDefaults)"] = normalise(etree.tostring(defaults, encoding="unicode"))
    return out


def package_facts(files: dict[str, bytes]) -> dict:
    types = files["[Content_Types].xml"].decode("utf-8")
    rels = files.get("word/_rels/document.xml.rels", b"").decode("utf-8")
    return {
        "parts": sorted(files),
        "content_types": sorted(re.findall(r'PartName="([^"]+)" ContentType="([^"]+)"', types)),
        "defaults": sorted(re.findall(r'Extension="([^"]+)" ContentType="([^"]+)"', types)),
        "relationships": sorted(re.findall(r'Type="[^"]*/([^/"]+)" Target="([^"]+)"', rels)),
    }


def properties(files: dict[str, bytes]) -> dict:
    out = {}
    for part in ("docProps/core.xml", "docProps/app.xml"):
        if part in files:
            text = files[part].decode("utf-8")
            text = re.sub(r"<dcterms:(created|modified)([^>]*)>[^<]*<", r"<dcterms:\1\2>{date}<", text)
            # Word's user, as Word fills in a new document's author.
            text = re.sub(r"<(dc:creator|cp:lastModifiedBy)>(?!Template)[^<]*<", r"<\1>{user}<", text)
            out[part] = re.sub(r"^<\?xml[^>]*>\s*", "", re.sub(r' xmlns(:\w+)?="[^"]*"', "", text))
    return out


def observe(name: str, data: bytes) -> dict:
    files = parts(data)
    entry: dict = {"package": package_facts(files), "settings": settings_facts(files)}
    if name.startswith(("convert", "resave")):
        # What Convert changes is the settings (compared with a save's); nothing else.
        return entry
    if name.startswith("new"):
        # Word's Normal template: only what is not its style set, theme or font table.
        entry["body"] = blocks(files)
        entry["properties"] = properties(files)
        return entry
    entry["body"] = blocks(files)
    entry["styles"] = style_facts(files)
    entry["properties"] = properties(files)
    for part in sorted(files):
        if re.match(r"word/(numbering|footnotes|endnotes|comments|header\d+|footer\d+)\.xml$", part):
            entry[part] = normalise(re.sub(r"^<\?xml[^>]*>\s*", "", files[part].decode("utf-8")))
    media = {part: hashlib.sha256(files[part]).hexdigest()[:12] for part in files if part.startswith("word/media/")}
    if media:
        entry["media"] = media
    return entry


# -- the new document's tables -------------------------------------------------------------------

TABLES = ROOT / "src" / "docx_agent" / "edit" / "word_new.py"
_PREFIXES = {W: "w", "http://schemas.microsoft.com/office/word/2010/wordml": "w14",
             "http://schemas.openxmlformats.org/officeDocument/2006/math": "m"}


def _name(qname: str) -> str:
    namespace, _, local = qname[1:].partition("}") if qname.startswith("{") else ("", "", qname)
    return f"{_PREFIXES[namespace]}:{local}" if namespace else local


def node(element) -> tuple:
    """An element as ``(tag, ((attribute, value), ...), (children...))``: values, in Word's
    order, from which ``edit/blank.py`` writes the XML."""
    attributes = tuple((_name(key), value) for key, value in element.attrib.items()
                       if not key.endswith("}rsid") and "rsid" not in key)
    return (_name(element.tag), attributes, tuple(node(child) for child in element))


def english_name(name: str) -> str:
    """English Word's display name of a built-in style ("heading 1" -> "Heading 1")."""
    return " ".join(word[:1].upper() + word[1:] for word in name.split())


def new_tables(data: bytes) -> dict:
    """What a new document is made of, as values: read from the one Word made."""
    files = parts(data)
    styles = etree.fromstring(files["word/styles.xml"])
    by_id = {s.get(_W + "styleId"): s.find(_W + "name").get(_W + "val") for s in styles.findall(_W + "style")}

    def name_of(style) -> str:
        name = style.find(_W + "name").get(_W + "val")
        link = style.find(_W + "link")
        if style.get(_W + "type") == "character" and link is not None:
            # A linked style's name is the interface language's ("Kop 1 Char"); English
            # Word's is its paragraph style's English name and " Char".
            return english_name(by_id[link.get(_W + "val")]) + " Char"
        return name

    names = {sid: name_of(s) for sid, s in ((s.get(_W + "styleId"), s) for s in styles.findall(_W + "style"))}
    out_styles = []
    for style in styles.findall(_W + "style"):
        entry: dict = {"type": style.get(_W + "type"), "name": names[style.get(_W + "styleId")]}
        if style.get(_W + "default") == "1":
            entry["default"] = True
        if style.get(_W + "customStyle") == "1":
            entry["custom"] = True
        children = []
        for child in style:
            tag = _name(child.tag)
            if tag in ("w:name", "w:rsid"):
                continue
            if tag in ("w:basedOn", "w:next", "w:link"):
                children.append((tag, names[child.get(_W + "val")]))
            elif tag == "w:uiPriority":
                children.append((tag, int(child.get(_W + "val"))))
            elif tag in ("w:pPr", "w:rPr", "w:tblPr", "w:trPr", "w:tcPr"):
                children.append(node(child))
            else:
                children.append((tag,))
        entry["children"] = tuple(children)
        out_styles.append(entry)
    latent = styles.find(_W + "latentStyles")
    latent_rows = []
    for exception in latent:
        flags = "".join(flag for flag, key in (("s", "semiHidden"), ("u", "unhideWhenUsed"), ("q", "qFormat"),
                                                ("l", "locked")) if exception.get(_W + key) == "1")
        priority = exception.get(_W + "uiPriority")
        latent_rows.append((exception.get(_W + "name"), None if priority is None else int(priority), flags))
    defaults = styles.find(_W + "docDefaults")
    settings = etree.fromstring(files["word/settings.xml"])
    fonts = etree.fromstring(files["word/fontTable.xml"])
    web = etree.fromstring(files["word/webSettings.xml"])
    return {
        "DOC_DEFAULTS": node(defaults),
        "LATENT_DEFAULTS": tuple((_name(k), v) for k, v in latent.attrib.items()),
        "LATENT": tuple(latent_rows),
        "STYLES": tuple(out_styles),
        "FONTS": tuple(node(font) for font in fonts),
        "COMPAT": tuple((s.get(_W + "name"), s.get(_W + "val")) for s in settings.iter(_W + "compatSetting")),
        "MATH": node(settings.find("{http://schemas.openxmlformats.org/officeDocument/2006/math}mathPr")),
        "COLOR_MAPPING": tuple((_name(k), v) for k, v in settings.find(_W + "clrSchemeMapping").attrib.items()),
        "WEB_SETTINGS": tuple(_name(child.tag) for child in web),
    }


def write_tables(data: bytes) -> None:
    import pprint

    values = new_tables(data)
    lines = ['"""What Word 16.106 makes a new blank document of, as values: the document defaults, the',
             "latent styles, the styles a new document has, the font table, the compatibility settings,",
             "the math settings and the web settings.",
             "",
             "Generated by ``tools/e6_probe.py`` from the new document Word made (``make new document``)",
             "and saved; do not edit by hand.  Nothing here is Word's XML: :mod:`docx_agent.edit.blank`",
             "writes every part from these values.  An element is ``(tag, ((attribute, value), ...),",
             "(children...))``; a style is its type, its English name (the interface's own names of",
             "linked character styles, \"Kop 1 Char\", are given as English Word names them, \"Heading 1",
             'Char"), and its children in order, where ``basedOn``, ``next`` and ``link`` name a style',
             "and ``uiPriority`` is a number.  A latent style is ``(name, uiPriority, flags)``: ``s``",
             "semi-hidden, ``u`` unhidden when used, ``q`` quick style, ``l`` locked.",
             '"""', "", "# fmt: off"]
    for key, value in values.items():
        lines.append(f"{key} = " + pprint.pformat(value, width=110, sort_dicts=False))
        lines.append("")
    TABLES.write_text("\n".join(lines), encoding="utf-8")


# -- Document.new() against Word's own new document ---------------------------------------------


def _canonical_new(files: dict[str, bytes], part: str, *, word: bool) -> str:
    """A part of a new document in a form both sides can be compared in: rsids, the
    document id and the window's view dropped; style ids and the linked styles' localised
    names as English Word writes them; the theme's localised name; dates and authors as
    placeholders."""
    data = files[part]
    if part.endswith(".rels") or part == "[Content_Types].xml":
        root = etree.fromstring(data)
        children = sorted(root, key=lambda n: (n.get("PartName") or n.get("Extension") or n.get("Target") or ""))
        for child in children:
            root.remove(child)
            child.attrib.pop("Id", None)
            root.append(child)
        return etree.tostring(root, method="c14n").decode()
    root = etree.fromstring(data)
    for node in list(root.iter()):
        if not isinstance(node.tag, str):
            continue
        for key in list(node.attrib):
            if "rsid" in key:
                del node.attrib[key]
        name = etree.QName(node).localname
        if name in ("rsids", "rsid", "docId", "view") and node.getparent() is not None:
            node.getparent().remove(node)
    if part == "word/styles.xml":
        tables = new_tables(_zip(files)) if word else None
        names = {}
        for style in root.findall(_W + "style"):
            names[style.get(_W + "styleId")] = style.find(_W + "name").get(_W + "val")
        if word:
            english = {entry_id: entry for entry_id, entry in names.items()}
            from docx_agent.edit.blank import style_id as english_id

            renamed = {}
            for style, entry in zip(root.findall(_W + "style"), tables["STYLES"]):
                renamed[style.get(_W + "styleId")] = english_id(entry["name"])
                style.find(_W + "name").set(_W + "val", entry["name"])
            for node in root.iter(_W + "basedOn", _W + "next", _W + "link"):
                node.set(_W + "val", renamed[node.get(_W + "val")])
            for style in root.findall(_W + "style"):
                style.set(_W + "styleId", renamed[style.get(_W + "styleId")])
            del english
    if part == "word/theme/theme1.xml":
        root.set("name", "{theme}")
    if part == "docProps/core.xml":
        for node in root:
            if etree.QName(node).localname in ("created", "modified"):
                node.text = "{date}"
            if etree.QName(node).localname in ("creator", "lastModifiedBy"):
                node.text = "{author}"
    if part == "docProps/app.xml":
        for node in root:
            if etree.QName(node).localname in ("Application", "AppVersion"):
                root.remove(node)
    if part == "word/document.xml":
        for node in root.iter(_W + "p"):
            for key in list(node.attrib):
                del node.attrib[key]
    return etree.tostring(root, method="c14n").decode()


def _zip(files: dict[str, bytes]) -> bytes:
    return zipped(dict(files))


def compare_new(word: bytes) -> dict:
    """``Document.new()``, built with this Word's language and locale, against the new
    document Word made: per part, ``same`` or where they differ."""
    from datetime import datetime, timezone

    from docx_agent.edit import blank

    ours = parts(blank.build(language="nl-BE", locale=blank.METRIC, author="Word's user",
                             created=datetime(2026, 10, 4, tzinfo=timezone.utc)))
    theirs = parts(word)
    out: dict = {}
    for part in sorted(set(ours) | set(theirs)):
        if part not in ours or part not in theirs:
            out[part] = "only in " + ("Word's" if part in theirs else "ours")
            continue
        a, b = _canonical_new(ours, part, word=False), _canonical_new(theirs, part, word=True)
        if a == b:
            out[part] = "same"
        else:
            at = next((k for k in range(min(len(a), len(b))) if a[k] != b[k]), min(len(a), len(b)))
            out[part] = {"ours": a[max(0, at - 80):at + 160], "Word's": b[max(0, at - 80):at + 160]}
    return out


def main(names: list[str]) -> None:
    saved = measure(names or list(runs()))
    if names:
        for name, data in saved.items():
            print(name, json.dumps(observe(name, data), indent=1, ensure_ascii=False)[:60000])
        return
    previous = json.loads(OBSERVATIONS.read_text(encoding="utf-8")) if OBSERVATIONS.exists() else {}
    observations = {
        "_about": ("What Word 16.106 for Mac (Dutch interface, metric) wrote for tools/e6_probe.py: per probe "
                   "the package's parts, content types and relationships, the settings, the body's blocks, the "
                   "styles by name and the notes, comments, numbering, headers and footers of the documents "
                   "written here, and the properties. rsids dropped; authors and dates replaced by "
                   "placeholders. A new document's styles, theme and font table are Word's Normal template and "
                   "are not recorded: 'new-compared' holds how Document.new() compares with them. "
                   "Regenerate with python tools/e6_probe.py."),
        **{name: observe(name, data) for name, data in saved.items()},
    }
    if "new" in saved:
        write_tables(saved["new"])
        print(f"wrote {TABLES.relative_to(ROOT)}")
        observations["new-compared"] = compare_new(saved["new"])
    elif "new-compared" in previous:
        observations["new-compared"] = previous["new-compared"]
    OBSERVATIONS.write_text(json.dumps(observations, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main(sys.argv[1:])
