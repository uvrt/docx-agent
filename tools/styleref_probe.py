"""Made-up ``STYLEREF`` probe documents (production round 21, T2), and the Word run that
measured what Word stores for them: every field of the body and of every header updated
through AppleScript (Word in English, ``-AppleLocale en_US``), the document saved as .docx.

:func:`documents` -> ``[(name, .docx bytes)]``.  Every probe field is labelled (``C=[...]``)
and cached as ``?``, so a value Word did not compute shows as ``?``.  Word's values, read
from its PDF after it updated every field (English interface), are
docx2svg's ``tests/fixtures/styleref-observations.json`` (the same documents; its
``tools/make_styleref_probe.py``); what Word stored after updating them is
``tests/observations/styleref-word.json``, which ``tests/test_styleref_update.py`` holds
:meth:`FieldOps.update_fields` to.  Run (holds the machine-wide Word lock,
``tests/oracle.py``)::

    PYTHONPATH=src python tools/styleref_probe.py base:update names:update ...  Nine pages separated by page breaks: page 1 a title, a chapter title and a
heading; page 2 none; page 3 two chapter titles after body text; page 4 one at its foot;
page 5 one in a table cell; page 6 one in a text box only; page 7 none; a chapter title split
across pages 8 and 9 by a page break in it.  Standard library only.
"""
from __future__ import annotations

import io
import zipfile
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSDECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
)


def t(text: str, style: str | None = None) -> str:
    rpr = f'<w:rPr><w:rStyle w:val="{style}"/></w:rPr>' if style else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def fld(instruction: str, cached: str = "?") -> str:
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {escape(instruction)} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:t xml:space="preserve">{escape(cached)}</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def labelled(label: str, instruction: str) -> str:
    return t(f"{label}=[") + fld(instruction) + t("]")


def p(content: str = "", style: str | None = None, extra_ppr: str = "") -> str:
    ppr = ""
    if style or extra_ppr:
        ppr = "<w:pPr>" + (f'<w:pStyle w:val="{style}"/>' if style else "") + extra_ppr + "</w:pPr>"
    return f"<w:p>{ppr}{content}</w:p>"


def para(text: str, style: str | None = None) -> str:
    return p(t(text), style)


PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def table(cells: list[list[str]]) -> str:
    rows = "".join("<w:tr>" + "".join(f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{c}</w:tc>'
                                      for c in row) + "</w:tr>" for row in cells)
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>'
            + '<w:gridCol w:w="4000"/>' * len(cells[0]) + f"</w:tblGrid>{rows}</w:tbl>")


def text_box(inner: str, ident: int = 1) -> str:
    """A floating text box (wps) anchored in a paragraph, holding ``inner`` paragraphs."""
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>3200000</wp:posOffset>'
        '</wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        '<wp:extent cx="2200000" cy="700000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>'
        f'<wp:docPr id="{ident}" name="Box {ident}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2200000" cy="700000"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="6350"><a:solidFill>'
        '<a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent>{inner}</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t">'
        '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def styles(lang: str = "en-US") -> str:
    def pstyle(sid, name, extra="", based="Normal", rpr=""):
        return (f'<w:style w:type="paragraph" w:customStyle="{0 if sid in BUILTIN else 1}" w:styleId="{sid}">'
                f'<w:name w:val="{name}"/><w:basedOn w:val="{based}"/><w:next w:val="Normal"/><w:qFormat/>'
                f'<w:pPr>{extra}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NSDECL}>'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        f'<w:sz w:val="20"/><w:lang w:val="{lang}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>'
        + pstyle("Title", "Title", rpr='<w:sz w:val="36"/>')
        + pstyle("Heading1", "heading 1", '<w:keepNext/><w:numPr><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/>',
                 rpr='<w:b/><w:sz w:val="28"/>')
        + pstyle("Heading2", "heading 2", '<w:keepNext/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'
                 '<w:outlineLvl w:val="1"/>', based="Heading1", rpr='<w:sz w:val="24"/>')
        + pstyle("Chaptertitle", "Chapter title", rpr='<w:b/><w:color w:val="1F4E79"/>')
        + pstyle("Unusedstyle", "Unused style")
        + pstyle("Latestyle", "Late style", rpr="<w:i/>")
        + pstyle("Header", "header")
        + '<w:style w:type="character" w:customStyle="1" w:styleId="Keyterm"><w:name w:val="Key term"/>'
          '<w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>'
        + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
          '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/>'
          '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
        + "</w:styles>")


BUILTIN = {"Title", "Heading1", "Heading2", "Header"}

NUMBERING = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NSDECL}>'
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/>'
    '<w:lvlText w:val="Part %1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>'
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:pStyle w:val="Heading2"/>'
    '<w:lvlText w:val="%1-%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>'
    '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')


def header_part(lines: list[str]) -> str:
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NSDECL}>' + "".join(
        p(line, "Header") for line in lines) + "</w:hdr>"


def body_fields(label: str, instructions: list[str]) -> str:
    return p("".join(t(" ") + labelled(f"{label}{k}", ins) for k, ins in enumerate(instructions)))


def base_body(body_probe: list[str], split_heading: bool = True) -> str:
    """Pages 1-9 (explicit page breaks); ``body_probe`` instructions are put in the body at
    labelled spots B0 (before any chapter title), B1 (after), B2 (page 2, none on it),
    B3 (page 3, after two), B9 (last page)."""
    bp = lambda label: body_fields(label, body_probe) if body_probe else ""  # noqa: E731
    out = [
        # page 1
        bp("B0_"),
        para("Report one", "Title"),
        para("Opening line on page one."),
        para("Alpha first", "Chaptertitle"),
        para("Part one heading", "Heading1"),
        p(t("Body with ") + t("Term one", "Keyterm") + t(" in it.")),
        bp("B1_"),
        PAGE_BREAK,
        # page 2: nothing in the styles
        para("Page two holds only body text."),
        bp("B2_"),
        PAGE_BREAK,
        # page 3: body first, then two chapter titles, a heading 2, the late style
        para("Page three starts with body text."),
        para("Beta third", "Chaptertitle"),
        para("Between the two."),
        para("Gamma third", "Chaptertitle"),
        bp("B3_"),
        para("Sub point heading", "Heading2"),
        para("Late one", "Latestyle"),
        p(t("More ") + t("Term two", "Keyterm") + t(" and ") + t("Term three", "Keyterm") + t(".")),
        PAGE_BREAK,
        # page 4: a chapter title at the bottom only
        para("Page four body."),
        para("Second part heading", "Heading1"),
        para("Delta fourth", "Chaptertitle"),
        PAGE_BREAK,
        # page 5: a chapter title in a table cell only
        para("Page five body before a table."),
        table([[para("Cell plain"), para("Epsilon cell", "Chaptertitle")]]),
        para("Page five after the table."),
        PAGE_BREAK,
        # page 6: a chapter title in a text box only
        p(text_box(para("Zeta box", "Chaptertitle")) + t("Page six body with a text box.")),
        para("Page six more body."),
        PAGE_BREAK,
        # page 7: nothing
        para("Page seven body only."),
        PAGE_BREAK,
        # page 8 -> 9: a chapter title split across the page boundary by a page break in it
        para("Page eight body."),
    ]
    if split_heading:
        out.append(p(t("Eta split start") + '<w:r><w:br w:type="page"/></w:r>' + t("Eta split end"), "Chaptertitle"))
    out += [para("Page nine body."), bp("B9_")]
    return "".join(out)


SECT = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>{more}'
        '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
        'w:header="567" w:footer="567" w:gutter="0"/>{titlepg}</w:sectPr>')


def build(header: list[str], *, body_probe: list[str] = (), first: list[str] | None = None,
          even: list[str] | None = None, lang: str = "en-US", body: str | None = None) -> bytes:
    headers = {"rIdH1": header_part(header)}
    more = ""
    if first is not None:
        headers["rIdH2"] = header_part(first)
        more += '<w:headerReference w:type="first" r:id="rIdH2"/>'
    if even is not None:
        headers["rIdH3"] = header_part(even)
        more += '<w:headerReference w:type="even" r:id="rIdH3"/>'
    sect = SECT.format(more=more, titlepg="<w:titlePg/>" if first is not None else "")
    content = body if body is not None else base_body(list(body_probe))
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NSDECL} mc:Ignorable="w14">'
                f"<w:body>{content}{sect}</w:body></w:document>")
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NSDECL}>'
                + ("<w:evenAndOddHeaders/>" if even is not None else "")
                + f'<w:defaultTabStop w:val="720"/><w:themeFontLang w:val="{lang}"/></w:settings>')
    rels = ['<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
            '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
            '<Relationship Id="rIdT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>']
    overrides = []
    for rid in headers:
        rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="h{rid}.xml"/>')
        overrides.append(f'<Override PartName="/word/h{rid}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          + "".join(overrides) + "</Types>")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(rels) + "</Relationships>")
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles(lang))
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/settings.xml", settings)
        for rid, xml in headers.items():
            z.writestr(f"word/h{rid}.xml", xml)
    return out.getvalue()


CH = 'STYLEREF "Chapter title"'
BASE_HEADER = [labelled("T", "STYLEREF Title"), labelled("C", CH), labelled("L", CH + " \\l"),
               labelled("H", 'STYLEREF "Heading 1"'), labelled("P", "PAGE")]

NAMES = {
    "n1": 'STYLEREF "heading 1"', "n2": 'STYLEREF "HEADING 1"', "n3": "STYLEREF Heading1",
    "n4": 'STYLEREF "Kop 1"', "n5": 'STYLEREF "chapter title"', "n6": "STYLEREF Chaptertitle",
    "n7": 'STYLEREF "Key term"', "n8": 'STYLEREF "Key term" \\l', "n9": 'STYLEREF "Unused style"',
    "na": 'STYLEREF "No such style"', "nb": 'STYLEREF "Late style"', "nc": 'STYLEREF "Late style" \\l',
    "nd": "STYLEREF 1", "ne": "STYLEREF 2", "nf": 'STYLEREF " Chapter title "',
}
SWITCHES = {
    "s0": 'STYLEREF "Heading 1"', "s1": 'STYLEREF "Heading 1" \\n', "s2": 'STYLEREF "Heading 1" \\r',
    "s3": 'STYLEREF "Heading 1" \\w', "s4": 'STYLEREF "Heading 1" \\p', "s5": 'STYLEREF "Heading 1" \\s',
    "s6": 'STYLEREF "Heading 1" \\t', "s7": 'STYLEREF "Heading 1" \\n \\t', "s8": 'STYLEREF "Heading 2" \\w',
    "s9": 'STYLEREF "Heading 2" \\r', "sa": 'STYLEREF "Heading 2" \\n', "sb": 'STYLEREF "Heading 2" \\w \\t',
    "sc": 'STYLEREF "Chapter title" \\p', "sd": 'STYLEREF "Heading 1" \\* Upper', "se": 'STYLEREF "Heading 1" \\l \\n',
    "sf": 'STYLEREF "Chapter title" \\* MERGEFORMAT',
}





def sections_case() -> bytes:
    """Two sections, each its own header: section 2 starts on page 3."""
    data = build([labelled("A", CH), labelled("AL", CH + " \\l")])
    body = base_body([CH])
    # end section 1 after page 2's content: the page break before page 3 becomes a section break
    marker = para("Page two holds only body text.") + body_fields("B2_", [CH]) + PAGE_BREAK
    assert marker in body
    sect1 = ('<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>'
             '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
             'w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:pPr></w:p>')
    body = body.replace(marker, para("Page two holds only body text.") + body_fields("B2_", [CH]) + sect1)
    out = io.BytesIO()
    src = zipfile.ZipFile(io.BytesIO(data))
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            raw = src.read(item.filename).decode()
            if item.filename == "word/document.xml":
                start = raw.index("<w:body>") + len("<w:body>")
                end = raw.index("<w:sectPr>")
                raw = raw[:start] + body + raw[end:].replace('r:id="rIdH1"', 'r:id="rIdH2"', 1)
            elif item.filename == "word/_rels/document.xml.rels":
                raw = raw.replace("</Relationships>", '<Relationship Id="rIdH2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="hrIdH2.xml"/></Relationships>')
            elif item.filename == "[Content_Types].xml":
                raw = raw.replace("</Types>", '<Override PartName="/word/hrIdH2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
            z.writestr(item, raw)
        z.writestr("word/hrIdH2.xml", header_part([labelled("S", CH), labelled("SL", CH + " \\l")]))
    return out.getvalue()


def firsteven2_case() -> bytes:
    """First/even/odd headers; page 2 starts with a chapter title, page 3 with another."""
    body = base_body([])
    body = body.replace(para("Page two holds only body text."), para("Two top", "Chaptertitle") + para("Page two holds only body text."))
    return build([labelled("O", CH)], first=[labelled("F", CH)], even=[labelled("E", CH)], body=body)


def footer_part(lines: list[str]) -> str:
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {NSDECL}>' + "".join(
        p(line, "Header") for line in lines) + "</w:ftr>"


def extra_case() -> bytes:
    """Section 1 = pages 1-2 (header + footer), section 2 from page 3 with its own header.
    Page 1 has two chapter titles; page 2 an empty chapter title paragraph; page 3 a body
    field before its first chapter title; page 4 a chapter title with a tab in it."""
    bf = lambda label: body_fields(label, [CH, CH + " \\l"])  # noqa: E731
    sect1 = ('<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>'
             '<w:footerReference w:type="default" r:id="rIdF1"/>'
             '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
             'w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:pPr></w:p>')
    body = "".join([
        para("One A", "Chaptertitle"), para("Body."), para("One B", "Chaptertitle"), para("Body."),
        PAGE_BREAK,
        para("Page two."), p("", "Chaptertitle"), para("After empty."), bf("E2_"),
        sect1,
        bf("E3_"), para("Three A", "Chaptertitle"), para("Body."),
        PAGE_BREAK,
        p(t("Four") + "<w:r><w:tab/></w:r>" + t("tabbed"), "Chaptertitle"), para("Body."),
    ])
    sect2 = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdH2"/>'
             '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
             'w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>')
    data = build([labelled("A", CH)])
    out = io.BytesIO()
    src = zipfile.ZipFile(io.BytesIO(data))
    rel = ('<Relationship Id="{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{1}" '
           'Target="{2}"/>')
    ct = '<Override PartName="/word/{0}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{1}+xml"/>'
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            raw = src.read(item.filename).decode()
            if item.filename == "word/document.xml":
                start = raw.index("<w:body>") + len("<w:body>")
                raw = raw[:start] + body + sect2 + "</w:body></w:document>"
            elif item.filename == "word/_rels/document.xml.rels":
                raw = raw.replace("</Relationships>", rel.format("rIdH2", "header", "hrIdH2.xml")
                                  + rel.format("rIdF1", "footer", "frIdF1.xml") + "</Relationships>")
            elif item.filename == "[Content_Types].xml":
                raw = raw.replace("</Types>", ct.format("hrIdH2.xml", "header") + ct.format("frIdF1.xml", "footer") + "</Types>")
            elif item.filename == "word/hrIdH1.xml":
                raw = header_part([labelled("A", CH), labelled("AL", CH + " \\l")])
            z.writestr(item, raw)
        z.writestr("word/hrIdH2.xml", header_part([labelled("S", CH), labelled("SL", CH + " \\l")]))
        z.writestr("word/frIdF1.xml", footer_part([labelled("FT", CH), labelled("FTL", CH + " \\l")]))
    return out.getvalue()


LONG = ("Long chapter title that runs on well past the width of one header line so that the header "
        "wraps onto a second line on this page only")


def height_case(cached: str = "?") -> bytes:
    """A one-line header whose STYLEREF value wraps to two lines on page 2 only; the
    header reaches past the top margin when it has two lines."""
    body = "".join([para("Short one", "Chaptertitle"), para("Body line page one."), PAGE_BREAK,
                    para(LONG, "Chaptertitle"), para("Body line page two."), PAGE_BREAK,
                    para("Short three", "Chaptertitle"), para("Body line page three."), PAGE_BREAK,
                    para("Body line page four.")])
    header = [t("C=[") + fld(CH, cached) + t("]")]
    data = build(header, body=body)
    out = io.BytesIO()
    src = zipfile.ZipFile(io.BytesIO(data))
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            raw = src.read(item.filename).decode()
            if item.filename == "word/document.xml":
                raw = raw.replace('w:top="1418"', 'w:top="720"').replace('w:header="567"', 'w:header="360"')
            z.writestr(item, raw)
    return out.getvalue()

def case_document(name: str) -> bytes:
    """The probe document ``name`` (one of :data:`CASES`)."""
    if name == "base":
        return build(BASE_HEADER, body_probe=[CH, CH + " \\l", "STYLEREF Title"])
    if name == "names":
        return build([labelled(k, v) for k, v in NAMES.items()],
                     body_probe=['STYLEREF "Late style"', 'STYLEREF "Key term"'])
    if name == "switches":
        return build([labelled(k, v) for k, v in SWITCHES.items()],
                     body_probe=['STYLEREF "Heading 1" \\p', 'STYLEREF "Heading 2" \\p'])
    if name == "nl":
        return build([labelled(k, NAMES[k]) for k in ("n4", "n9", "na", "nb")] + [labelled("C", CH)], lang="nl-NL",
                     body_probe=['STYLEREF "Unused style"'])
    if name == "firsteven":
        return build([labelled("O", CH), labelled("P", "PAGE")], first=[labelled("F", CH), labelled("P", "PAGE")],
                     even=[labelled("E", CH), labelled("P", "PAGE")])
    if name == "firsteven2":
        return firsteven2_case()
    if name == "sections":
        return sections_case()
    if name == "extra":
        return extra_case()
    if name == "latent":
        return latent_case()
    if name == "height":
        return height_case()
    if name == "heightlong":
        return height_case(LONG)
    raise KeyError(name)


def latent_case() -> bytes:
    """Built-in styles not in the styles part, a number, no name, the default styles,
    case switches and a switch Word does not know."""
    return build([labelled("l1", 'STYLEREF "Heading 3"'), labelled("l2", "STYLEREF 3"),
                  labelled("l3", 'STYLEREF "Quote"'), labelled("l4", 'STYLEREF "TOC 1"'),
                  labelled("l5", "STYLEREF"), labelled("l6", 'STYLEREF "Header"'),
                  labelled("l7", 'STYLEREF "Normal"'), labelled("l8", 'STYLEREF "Default Paragraph Font"'),
                  labelled("l9", "STYLEREF Title \\* Upper"), labelled("la", 'STYLEREF "Chapter title" \\* Caps'),
                  labelled("lb", 'STYLEREF "Chapter title" \\* MERGEFORMAT \\l'),
                  labelled("lc", 'STYLEREF "Chapter title" \\x')])


#: The documents Word measured (``height`` and ``heightlong``: the body's top under a
#: header that wraps on one page, measured from the PDF's glyphs, not by the test).
CASES = ["base", "names", "switches", "nl", "firsteven", "firsteven2", "sections", "extra", "latent", "height",
         "heightlong"]


def documents() -> list[tuple[str, bytes]]:
    return [(name, case_document(name)) for name in CASES]


# -- the Word run ------------------------------------------------------------------------------

import subprocess  # noqa: E402
import sys  # noqa: E402
import time  # noqa: E402
from pathlib import Path  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tests"))
import oracle  # noqa: E402

#: Where the saved documents and PDFs go (outside the repository).
OUT = Path.cwd()

SCRIPT = '''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set pdfPath to item 3 of argv
  set mode to item 4 of argv
  set report to ""
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file inputPath)
      repeat 60 times
        if (count of documents) > 0 then exit repeat
        delay 0.5
      end repeat
      repeat 40 times
        try
          set d to active document
          get name of d
          exit repeat
        end try
        delay 0.5
      end repeat
      try
      if mode is "update" then
        set n to count of fields of d
        repeat with i from 1 to n
          try
            update field (field i of d)
          on error m
            set report to report & "body " & i & ": " & m & linefeed
          end try
        end repeat
        repeat with s from 1 to count of sections of d
          repeat with k in {header footer primary, header footer first page, header footer even pages}
            try
              set hf to get header (section s of d) index k
              set m to count of fields of text object of hf
              repeat with i from 1 to m
                update field (field i of text object of hf)
              end repeat
            on error e
              set report to report & "header " & s & ": " & e & linefeed
            end try
          end repeat
        end repeat
      end if
      save as d file name outputPath file format format document
      save as d file name pdfPath file format format PDF
      on error e
        -- every document closed unsaved, on an error too
        close d saving no
        quit saving no
        error e
      end try
      close d saving no
      quit saving no
    end tell
  end timeout
  return report
end run
'''


def wait(cond, n=300):
    for _ in range(n):
        if cond():
            return True
        time.sleep(0.1)
    return False


def run(case: str, mode: str, english: bool = True) -> bool:
    data = case_document(case)
    tag = f"{case}-{mode}{'' if english else '-nlui'}"
    staged = oracle.ORACLE_DIR / f"r21c-{tag}-in.docx"
    out = oracle.ORACLE_DIR / f"r21c-{tag}-out.docx"
    pdf = oracle.ORACLE_DIR / f"r21c-{tag}-out.pdf"
    staged.write_bytes(data)
    out.unlink(missing_ok=True)
    pdf.unlink(missing_ok=True)
    wait(lambda: not oracle.word_running(), 150)
    if english:
        subprocess.run(["open", "-g", "-a", str(oracle.WORD_APP), "--args", "-AppleLocale", "en_US",
                        "-AppleLanguages", "(en-US)"], check=True)
        wait(oracle.word_running)
        time.sleep(10)
    start = time.monotonic()
    done = subprocess.run(["osascript", "-e", SCRIPT, str(staged), str(out), str(pdf), mode],
                          capture_output=True, text=True, timeout=200)
    staged.unlink(missing_ok=True)
    print(tag, out.exists(), pdf.exists(), (done.stdout + done.stderr).strip()[:400], round(time.monotonic() - start, 1),
          flush=True)
    if out.exists():
        (OUT / f"{tag}-word.docx").write_bytes(out.read_bytes())
        out.unlink()
    ok = pdf.exists()
    if pdf.exists():
        (OUT / f"{tag}-word.pdf").write_bytes(pdf.read_bytes())
        pdf.unlink()
    return ok


if __name__ == "__main__":
    english = "--nl" not in sys.argv
    jobs = [a for a in sys.argv[1:] if not a.startswith("--")]
    with oracle.session() as word:
        for job in jobs:
            case, mode = job.split(":")
            for attempt in range(3):
                try:
                    if run(case, mode, english):
                        break
                except Exception as e:  # noqa: BLE001
                    print(job, "failed", e, flush=True)
                word.recover()
            word.recover()
