"""Made-up documents for the round-32 TOC probes (R12 numbered headings, R14 toc N tab stops):
the inputs of tools/toc_numbers_probe.py and tests/test_toc_numbers.py."""
import io, zipfile

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
            '</Relationships>')
SETTINGS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}>'
            '<w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')

TEXTS = {0: "%1", 1: "%1.%2", 2: "%1.%2.%3"}


def numbering(suff="tab", linked=True, start=3, texts=TEXTS, lvl_extra="", lvl_rpr="", ind=(432, 576, 720)):
    lvls = []
    for k in range(9):
        text = texts.get(k, "%1." * 0 + ".".join(f"%{i + 1}" for i in range(k + 1)))
        pstyle = f'<w:pStyle w:val="Heading{k + 1}"/>' if linked and k < 3 else ""
        s = f'<w:suff w:val="{suff}"/>' if suff != "tab" else ""
        left = ind[k] if k < 3 else 720 + 144 * k
        rpr = f'<w:rPr>{lvl_rpr}</w:rPr>' if lvl_rpr else ""
        lvls.append(f'<w:lvl w:ilvl="{k}"><w:start w:val="{start if k == 0 else 1}"/><w:numFmt w:val="decimal"/>{pstyle}{s}'
                    f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                    f'<w:pPr>{lvl_extra}<w:ind w:left="{left}" w:hanging="{left}"/></w:pPr>{rpr}</w:lvl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS}>'
            f'<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>{"".join(lvls)}</w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')


def styles(toc_ppr=("", "", ""), linked=True, font="Arial", sz=20, toc_ind=(0, 200, 400)):
    out = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>',
           f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
           f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:rPrDefault>'
           '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
           '</w:docDefaults>',
           '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>',
           '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>'
           '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>']
    for k in (1, 2, 3):
        num = f'<w:numPr><w:ilvl w:val="{k - 1}"/><w:numId w:val="1"/></w:numPr>' if linked else ""
        out.append(f'<w:style w:type="paragraph" w:styleId="Heading{k}"><w:name w:val="heading {k}"/>'
                   f'<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
                   f'<w:pPr><w:keepNext/>{num}<w:outlineLvl w:val="{k - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="{30 - 4 * k}"/></w:rPr></w:style>')
    for k in (1, 2, 3):
        out.append(f'<w:style w:type="paragraph" w:styleId="TOC{k}"><w:name w:val="toc {k}"/>'
                   f'<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>'
                   f'<w:pPr>{toc_ppr[k - 1]}<w:spacing w:after="100"/><w:ind w:left="{toc_ind[k - 1]}"/></w:pPr></w:style>')
    out.append('<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/>'
               '<w:next w:val="Normal"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>')
    out.append('<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>'
               '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
               '<w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>')
    out.append('</w:styles>')
    return "".join(out)


def r(text, rpr=""):
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def fld(kind):
    return f'<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>'


def instr(text, rpr=""):
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:instrText xml:space="preserve">{text}</w:instrText></w:r>'


HEADINGS = [(1, "Opening part"), (2, "Second step"), (3, "Third detail"), (2, "Fourth step"), (1, "Closing part")]


def toc():
    """Title and a stale TOC field (one entry, no numbers): Word's update rebuilds it."""
    head = fld("begin") + instr(' TOC \\o "1-3" \\h \\z \\u ') + fld("separate")
    return ('<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>' + r("Contents") + '</w:p>'
            + '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' + head + r("stale entry") + '</w:p>'
            + '<w:p>' + fld("end") + '</w:p>')


def headings(direct=False):
    out = []
    for k, (level, text) in enumerate(HEADINGS):
        num = f'<w:numPr><w:ilvl w:val="{level - 1}"/><w:numId w:val="1"/></w:numPr>' if direct else ""
        out.append(f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/>{num}</w:pPr>{r(text)}</w:p>'
                   + f'<w:p>{r("Body text under " + text.lower() + ".")}</w:p>')
    return "".join(out)


def document(body, width=11906, margin=1418):
    sect = (f'<w:sectPr><w:pgSz w:w="{width}" w:h="16838"/><w:pgMar w:top="1418" w:right="{margin}" '
            f'w:bottom="1418" w:left="{margin}" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


def package(body, styles_xml, numbering_xml, settings=None, **kw):
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", document(body, **kw))
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/settings.xml", settings or SETTINGS)
        z.writestr("word/numbering.xml", numbering_xml)
    return out.getvalue()


def tabs(pos, val="right", leader="dot"):
    lead = f' w:leader="{leader}"' if leader else ""
    return f'<w:tabs><w:tab w:val="{val}"{lead} w:pos="{pos}"/></w:tabs>'


BODY = toc() + headings()
CASES = {
    # R12: style-linked multilevel list, number then a tab / space / nothing.
    "n-tab": package(BODY, styles(), numbering()),
    "n-space": package(BODY, styles(), numbering(suff="space")),
    "n-nothing": package(BODY, styles(), numbering(suff="nothing")),
    # Numbering on the paragraphs, styles not linked.
    "n-direct": package(toc() + headings(direct=True), styles(linked=False), numbering(linked=False)),
    # "%1." with the level's own tab stop, the number bold and red.
    "n-lvl": package(BODY, styles(), numbering(texts={0: "%1.", 1: "%1.%2.", 2: "%1.%2.%3."},
                                               lvl_extra='<w:tabs><w:tab w:val="num" w:pos="851"/></w:tabs>',
                                               lvl_rpr='<w:b/><w:color w:val="FF0000"/>')),
    # 11 pt Calibri-sized text, other toc indents.
    "n-11pt": package(BODY, styles(sz=22, toc_ind=(0, 220, 440)), numbering()),
    "n-wide": package(BODY, styles(toc_ind=(0, 567, 1134)), numbering(texts={0: "Part %1", 1: "%1.%2", 2: "%1.%2.%3"})),
    # R14: toc N styles' own tabs (text column 9070 twips here).
    "t-none": package(BODY, styles(), numbering()),
    "t-end": package(BODY, styles(toc_ppr=(tabs(9070),) * 3), numbering()),
    "t-end10": package(BODY, styles(toc_ppr=(tabs(9060),) * 3), numbering()),
    "t-near": package(BODY, styles(toc_ppr=(tabs(9000),) * 3), numbering()),
    "t-far": package(BODY, styles(toc_ppr=(tabs(6000),) * 3), numbering()),
    "t-past": package(BODY, styles(toc_ppr=(tabs(9500),) * 3), numbering()),
    "t-noleader": package(BODY, styles(toc_ppr=(tabs(9070, leader=None),) * 3), numbering()),
    "t-left": package(BODY, styles(toc_ppr=(tabs(9070, val="left"),) * 3), numbering()),
    "t-only1": package(BODY, styles(toc_ppr=(tabs(9070), "", "")), numbering()),
    "t-plain": package(toc() + headings(), styles(linked=False, toc_ppr=(tabs(9070),) * 3), numbering(linked=False)),
}


# -- sweep: one list per heading (direct numbering), so each entry's number differs ---------------
SWEEP = ["%1", "%1.", "(%1)", "%1.0", "%1.00", "%1.000", "%1.0000", "Part %1", "Section %1", "Chapter %1.",
         "WWWW %1", "WWWWWW %1"]


def sweep_numbering(texts=SWEEP, suff="tab"):
    out = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS}>']
    for k, text in enumerate(texts):
        s = f'<w:suff w:val="{suff}"/>' if suff != "tab" else ""
        lvls = "".join(f'<w:lvl w:ilvl="{i}"><w:start w:val="3"/><w:numFmt w:val="decimal"/>{s}'
                       f'<w:lvlText w:val="{text.replace("%1", f"%{i + 1}")}"/><w:lvlJc w:val="left"/>'
                       f'<w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl>' for i in range(3))
        out.append(f'<w:abstractNum w:abstractNumId="{k}"><w:multiLevelType w:val="multilevel"/>{lvls}</w:abstractNum>')
    for k in range(len(texts)):
        out.append(f'<w:num w:numId="{k + 1}"><w:abstractNumId w:val="{k}"/></w:num>')
    return "".join(out) + "</w:numbering>"


def sweep_headings(levels=(1, 2, 3), n=None):
    n = len(SWEEP) if n is None else n
    out = []
    for level in levels:
        for k in range(n):
            out.append(f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/><w:numPr><w:ilvl w:val="{level - 1}"/>'
                       f'<w:numId w:val="{k + 1}"/></w:numPr></w:pPr>{r(f"Item {level}{k:02d}")}</w:p>')
    return "".join(out)


SWEEP_BODY = toc() + sweep_headings()
SETTINGS_567 = SETTINGS.replace('w:val="720"', 'w:val="567"')
CASES2 = {
    "s-arial10": package(SWEEP_BODY, styles(linked=False), sweep_numbering()),
    "s-arial12": package(SWEEP_BODY, styles(linked=False, sz=24, toc_ind=(0, 240, 480)), sweep_numbering()),
    "s-times10": package(SWEEP_BODY, styles(linked=False, font="Times New Roman", toc_ind=(0, 0, 0)), sweep_numbering()),
    "s-tab567": package(SWEEP_BODY, styles(linked=False), sweep_numbering(), settings=SETTINGS_567),
    "s-ind300": package(SWEEP_BODY, styles(linked=False, toc_ind=(300, 300, 300)), sweep_numbering()),
}


# -- set 3: toc 4-9 Word's own (styles absent), non-monotonic indents, more R14 forms -------------
def styles9(toc_ind=None, sz=20, toc_ppr=None, extra=""):
    """Heading 1-9 linked to list 1's levels; toc N styles only for the levels in ``toc_ind``."""
    out = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}>',
           f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
           f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:rPrDefault>'
           '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
           '</w:docDefaults>',
           '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>',
           '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>'
           '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>', extra]
    for k in range(1, 10):
        out.append(f'<w:style w:type="paragraph" w:styleId="Heading{k}"><w:name w:val="heading {k}"/>'
                   f'<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
                   f'<w:pPr><w:keepNext/><w:numPr><w:ilvl w:val="{k - 1}"/><w:numId w:val="1"/></w:numPr>'
                   f'<w:outlineLvl w:val="{k - 1}"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>')
    for k, ind in (toc_ind or {}).items():
        ppr = (toc_ppr or {}).get(k, "")
        out.append(f'<w:style w:type="paragraph" w:styleId="TOC{k}"><w:name w:val="toc {k}"/>'
                   f'<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>'
                   f'<w:pPr>{ppr}<w:spacing w:after="100"/><w:ind w:left="{ind}"/></w:pPr></w:style>')
    out.append('<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>'
               '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
               '<w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>')
    out.append('</w:styles>')
    return "".join(out)


def numbering9(texts=None):
    texts = texts or {}
    lvls = []
    for k in range(9):
        pstyle = f'<w:pStyle w:val="Heading{k + 1}"/>'
        text = texts.get(k, ".".join(f"%{i + 1}" for i in range(k + 1)))
        lvls.append(f'<w:lvl w:ilvl="{k}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>{pstyle}'
                    f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                    f'<w:pPr><w:ind w:left="{432 + 144 * k}" w:hanging="{432 + 144 * k}"/></w:pPr></w:lvl>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS}>'
            f'<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>{"".join(lvls)}</w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')


def toc9():
    head = fld("begin") + instr(' TOC \\o "1-9" \\h \\z \\u ') + fld("separate")
    return ('<w:p>' + r("Contents") + '</w:p><w:p>' + head + r("stale entry") + '</w:p><w:p>' + fld("end") + '</w:p>')


def headings9():
    return "".join(f'<w:p><w:pPr><w:pStyle w:val="Heading{k}"/></w:pPr>{r(f"Level {k} title")}</w:p>' for k in range(1, 10))


BODY9 = toc9() + headings9()
BASE_TAB = ('<w:style w:type="paragraph" w:styleId="TocBase"><w:name w:val="Toc Base"/><w:basedOn w:val="Normal"/>'
            '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9070"/></w:tabs></w:pPr></w:style>')
CASES3 = {
    "x-none10": package(BODY9, styles9(sz=20), numbering9()),
    "x-none11": package(BODY9, styles9(sz=22), numbering9()),
    "x-none12": package(BODY9, styles9(sz=24), numbering9()),
    "x-only3": package(BODY9, styles9(toc_ind={1: 0, 2: 200, 3: 400}), numbering9()),
    "x-odd": package(BODY9, styles9(toc_ind={1: 0, 2: 1000, 3: 500, 4: 1500, 5: 300, 6: 2500, 7: 2000, 8: 100, 9: 3000}),
                     numbering9()),
    "x-flat": package(BODY9, styles9(toc_ind={k: 0 for k in range(1, 10)}), numbering9()),
    "x-hanging": package(BODY9, styles9(toc_ind={1: 0, 2: 200, 3: 400}, toc_ppr={
        1: '', 2: '', 3: ''}).replace('<w:ind w:left="200"/>', '<w:ind w:left="600" w:hanging="400"/>'), numbering9()),
    # R14 more: a left tab only; a tab inherited from the style's base; a cleared inherited tab.
    "t-smallleft": package(BODY, styles(toc_ppr=(tabs(1000, val="left", leader=None),) * 3), numbering()),
    "t-inherit": package(BODY, styles().replace('<w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>',
                                                '<w:name w:val="toc 1"/><w:basedOn w:val="TocBase"/>')
                         .replace('</w:styles>', BASE_TAB + '</w:styles>'), numbering()),
    "t-clear": package(BODY, styles(toc_ppr=('<w:tabs><w:tab w:val="clear" w:pos="9070"/></w:tabs>', "", ""))
                       .replace('<w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>',
                                '<w:name w:val="toc 1"/><w:basedOn w:val="TocBase"/>')
                       .replace('</w:styles>', BASE_TAB + '</w:styles>'), numbering()),
}
