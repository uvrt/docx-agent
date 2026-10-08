"""Charts, embedded workbooks and SmartArt written as parts of a Word document.

What ``tools/make_chart_fixtures.py`` builds its documents from and ``tools/charts_probe.py``
its probes: a chart part as Word 365 writes one (``c14:style``, the ``c16r2`` declaration,
a ``c16:uniqueId`` on every series, ``c16r3`` display options, an embedded workbook named by
``c:externalData`` with ``c:autoUpdate``), the workbook itself (one sheet, shared strings, a
table over the data, laid out as Word's Insert Chart lays it out: categories down column A,
a series to a column), and a SmartArt data model with no presentation points and no cached
drawing, which Word lays out and writes on saving.

Every byte is written here.  The SmartArt layout, quick-style and colour definitions are the
ones pptx-agent's ``tests/fixtures/powerpoint-smartart.pptx`` carries (PowerPoint wrote them
there), read from the sibling checkout when a document is built: a definition must be whole
for Word to lay out anything but Basic Block List (docx2svg's ROADMAP, F.21).
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parent.parent
#: Where the SmartArt definitions are read from (``definitions``).
SMARTART_SOURCE = ROOT.parent / "pptx-agent" / "tests" / "fixtures" / "powerpoint-smartart.pptx"

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
C14 = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
C16 = "http://schemas.microsoft.com/office/drawing/2014/chart"
C16R2 = "http://schemas.microsoft.com/office/drawing/2015/06/chart"
C16R3 = "http://schemas.microsoft.com/office/drawing/2017/03/chart"
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = R + "/"
PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_CHART = REL + "chart"
REL_PACKAGE = REL + "package"
REL_DGM = {"data": REL + "diagramData", "layout": REL + "diagramLayout", "style": REL + "diagramQuickStyle",
           "colors": REL + "diagramColors"}
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_DML = "application/vnd.openxmlformats-officedocument.drawingml."
CT_DGM = {"data": _DML + "diagramData+xml", "layout": _DML + "diagramLayout+xml",
          "style": _DML + "diagramStyle+xml", "colors": _DML + "diagramColors+xml"}
_SML = "application/vnd.openxmlformats-officedocument.spreadsheetml."
STAMP = (2026, 10, 4, 12, 0, 0)

#: The URIs of the extensions Word writes into a chart.
UNIQUE_ID_URI = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}"
DISPLAY_URI = "{56B9EC1D-385E-4148-901F-78D8002777C0}"


def column(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def number(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return repr(value) if isinstance(value, float) else str(value)


def zip_parts(parts: dict[str, bytes | str]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data.encode() if isinstance(data, str) else data)
    return buffer.getvalue()


def relationships(rels: list[tuple[str, str, str]]) -> str:
    return (DECL + f'<Relationships xmlns="{PKG_RELS}">'
            + "".join(f'<Relationship Id="{i}" Type="{t}" Target="{target}"/>' for i, t, target in rels)
            + "</Relationships>")


# -- the chart's data ------------------------------------------------------------------------


@dataclass
class ChartData:
    """What a chart shows: categories (text, or numbers for a scatter's x values) down
    column A from row 2, each series' name in row 1 and its values below, a column each."""

    categories: list
    series: dict[str, list]
    #: The header of column A (Word's Insert Chart leaves it empty; a scatter's says
    #: ``X-Values``).
    corner: str | None = None
    sheet: str = "Sheet1"

    @property
    def last_row(self) -> int:
        return len(self.categories) + 1

    def ref(self, col: int, first: int, last: int | None = None) -> str:
        start = f"${column(col)}${first}"
        return f"{self.sheet}!{start}" + (f":${column(col)}${last}" if last is not None else "")


def workbook(data: ChartData, *, values: dict[str, list] | None = None) -> bytes:
    """The embedded workbook behind ``data`` -- or, with ``values``, one whose cells say
    something else than the chart's cache (a probe)."""
    shown = values or data.series
    strings: list[str] = []
    uses = 0

    def text_cell(ref: str, value: str) -> str:
        nonlocal uses
        if value not in strings:
            strings.append(value)
        uses += 1
        return f'<c r="{ref}" t="s"><v>{strings.index(value)}</v></c>'

    rows: list[str] = []
    width = len(shown) + 1
    header = [text_cell("A1", data.corner)] if data.corner else []
    for k, name in enumerate(shown, start=2):
        header.append(text_cell(f"{column(k)}1", name))
    rows.append(f'<row r="1" spans="1:{width}">' + "".join(header) + "</row>")
    for i, category in enumerate(data.categories):
        r = i + 2
        cells = [text_cell(f"A{r}", category) if isinstance(category, str)
                 else f'<c r="A{r}"><v>{number(category)}</v></c>']
        for k, points in enumerate(shown.values(), start=2):
            value = points[i]
            if value is not None:
                cells.append(f'<c r="{column(k)}{r}"><v>{number(value)}</v></c>')
        rows.append(f'<row r="{r}" spans="1:{width}">' + "".join(cells) + "</row>")
    last = f"{column(width)}{data.last_row}"
    names = [data.corner or "Column1"] + list(shown)
    sheet = (DECL + f'<worksheet xmlns="{X}" xmlns:r="{R}"><dimension ref="A1:{last}"/>'
             '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
             '<sheetFormatPr baseColWidth="10" defaultRowHeight="16"/>'
             f"<sheetData>{''.join(rows)}</sheetData>"
             '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
             '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')
    table = (DECL + f'<table xmlns="{X}" id="1" name="Table1" displayName="Table1" ref="A1:{last}" '
             f'totalsRowShown="0"><tableColumns count="{len(names)}">'
             + "".join(f'<tableColumn id="{k}" name="{escape(n)}"/>' for k, n in enumerate(names, start=1))
             + '</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
             'showRowStripes="1" showColumnStripes="0"/></table>')
    shared = (DECL + f'<sst xmlns="{X}" count="{uses}" uniqueCount="{len(strings)}">'
              + "".join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in strings) + "</sst>")
    styles = (DECL + f'<styleSheet xmlns="{X}"><fonts count="1"><font><sz val="12"/><color theme="1"/>'
              '<name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font></fonts>'
              '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/>'
              '</fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
              '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
              '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" '
              'defaultPivotStyle="PivotStyleLight16"/></styleSheet>')
    book = (DECL + f'<workbook xmlns="{X}" xmlns:r="{R}"><bookViews><workbookView xWindow="0" yWindow="0" '
            'windowWidth="16000" windowHeight="10000"/></bookViews>'
            f'<sheets><sheet name="{escape(data.sheet)}" sheetId="1" r:id="rId1"/></sheets>'
            '<calcPr calcId="191029"/></workbook>')
    types = (DECL + f'<Types xmlns="{CT}"><Default Extension="rels" '
             'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="xml" ContentType="application/xml"/>'
             f'<Override PartName="/xl/workbook.xml" ContentType="{_SML}sheet.main+xml"/>'
             f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{_SML}worksheet+xml"/>'
             f'<Override PartName="/xl/styles.xml" ContentType="{_SML}styles+xml"/>'
             f'<Override PartName="/xl/sharedStrings.xml" ContentType="{_SML}sharedStrings+xml"/>'
             f'<Override PartName="/xl/tables/table1.xml" ContentType="{_SML}table+xml"/></Types>')
    return zip_parts({
        "[Content_Types].xml": types,
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "xl/workbook.xml")]),
        "xl/workbook.xml": book,
        "xl/_rels/workbook.xml.rels": relationships([("rId1", REL + "worksheet", "worksheets/sheet1.xml"),
                                                     ("rId2", REL + "styles", "styles.xml"),
                                                     ("rId3", REL + "sharedStrings", "sharedStrings.xml")]),
        "xl/worksheets/sheet1.xml": sheet,
        "xl/worksheets/_rels/sheet1.xml.rels": relationships([("rId1", REL + "table", "../tables/table1.xml")]),
        "xl/tables/table1.xml": table,
        "xl/sharedStrings.xml": shared,
        "xl/styles.xml": styles,
    })


# -- the chart part ----------------------------------------------------------------------------


@dataclass
class ChartSpec:
    """One chart: ``kind`` ``column``, ``bar``, ``line``, ``pie``, ``doughnut``, ``scatter``
    or ``combo`` (columns, and a line for the last series), its data, a title (``None``:
    none, the auto title deleted) and the legend's position (``None``: no legend)."""

    kind: str
    data: ChartData
    title: str | None = None
    legend: str | None = "b"
    #: ``c:title``'s whole ``c:tx`` and ``c:txPr``, written as given (a probe's variants);
    #: ``None`` writes Word's own form for ``title``.
    title_xml: str | None = None
    #: The values the cache holds, when they are to differ from the workbook's (a probe).
    cache: dict[str, list] | None = None
    #: ``c:autoUpdate``'s value, or ``None`` for none.
    auto_update: str | None = "0"
    #: Number format of the value caches.
    format_code: str = "General"
    #: A seed for the series' ``c16:uniqueId``s, so two charts' ids differ.
    seed: int = 0
    #: The workbook's own data, when it is to say something else than the chart (a probe).
    book: ChartData | None = None


def _str_cache(values: list) -> str:
    points = "".join(f'<c:pt idx="{k}"><c:v>{escape(str(v))}</c:v></c:pt>' for k, v in enumerate(values)
                     if v is not None)
    return f'<c:strCache><c:ptCount val="{len(values)}"/>{points}</c:strCache>'


def _num_cache(values: list, code: str) -> str:
    points = "".join(f'<c:pt idx="{k}"><c:v>{number(v)}</c:v></c:pt>' for k, v in enumerate(values)
                     if v is not None)
    return f'<c:numCache><c:formatCode>{escape(code)}</c:formatCode><c:ptCount val="{len(values)}"/>{points}</c:numCache>'


def _unique(spec: ChartSpec, index: int) -> str:
    return (f'<c:extLst><c:ext uri="{UNIQUE_ID_URI}" xmlns:c16="{C16}">'
            f'<c16:uniqueId val="{{{spec.seed:04X}{index:04X}-0000-4000-8000-0000000000{index:02X}}}"/>'
            "</c:ext></c:extLst>")


def _series(spec: ChartSpec, index: int, name: str, plot: str) -> str:
    data = spec.data
    col = index + 2
    cache = (spec.cache or data.series)[name]
    tx = f'<c:tx><c:strRef><c:f>{data.ref(col, 1)}</c:f>{_str_cache([name])}</c:strRef></c:tx>'
    numeric_categories = all(not isinstance(c, str) for c in data.categories)
    if numeric_categories:
        cat = (f"<c:numRef><c:f>{data.ref(1, 2, data.last_row)}</c:f>"
               f"{_num_cache(data.categories, 'General')}</c:numRef>")
    else:
        cat = f"<c:strRef><c:f>{data.ref(1, 2, data.last_row)}</c:f>{_str_cache(data.categories)}</c:strRef>"
    val = f"<c:numRef><c:f>{data.ref(col, 2, data.last_row)}</c:f>{_num_cache(cache, spec.format_code)}</c:numRef>"
    head = f'<c:ser><c:idx val="{index}"/><c:order val="{index}"/>{tx}'
    if plot == "bar":
        body = f'<c:invertIfNegative val="0"/><c:cat>{cat}</c:cat><c:val>{val}</c:val>'
    elif plot == "line":
        body = f'<c:marker><c:symbol val="none"/></c:marker><c:cat>{cat}</c:cat><c:val>{val}</c:val><c:smooth val="0"/>'
    elif plot in ("pie", "doughnut"):
        body = f"<c:cat>{cat}</c:cat><c:val>{val}</c:val>"
    elif plot == "scatter":
        body = (f'<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker><c:xVal>{cat}</c:xVal>'
                f'<c:yVal>{val}</c:yVal><c:smooth val="0"/>')
    else:
        raise ValueError(plot)
    return head + body + _unique(spec, index) + "</c:ser>"


_AXES = (('<c:catAx><c:axId val="{c}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
          '<c:axPos val="{cpos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
          '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="{v}"/><c:crosses val="autoZero"/>'
          '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'),
         ('<c:valAx><c:axId val="{v}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
          '<c:axPos val="{vpos}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/>'
          '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
          '<c:crossAx val="{c}"/><c:crosses val="autoZero"/><c:crossBetween val="{between}"/></c:valAx>'))

_SCATTER_AXES = ('<c:valAx><c:axId val="{c}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
                 '<c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
                 '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="{v}"/>'
                 '<c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
                 + _AXES[1].replace("{between}", "midCat"))


def _title(spec: ChartSpec) -> str:
    if spec.title_xml is not None:
        return f'<c:title>{spec.title_xml}<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
    if spec.title is None:
        return '<c:autoTitleDeleted val="1"/>'
    return ('<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" '
            'wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" i="0" '
            'u="none" strike="noStrike" kern="1200" spc="0" baseline="0"><a:solidFill><a:schemeClr val="tx1">'
            '<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/>'
            '<a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="en-US"/><a:t>{escape(spec.title)}</a:t></a:r></a:p></c:rich></c:tx>'
            '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>')


def chart_xml(spec: ChartSpec, package_rel: str | None = "rId1") -> str:
    data = spec.data
    names = list(data.series)
    kind = spec.kind
    if kind in ("column", "bar"):
        plots = [("bar", names, f'<c:barDir val="{"col" if kind == "column" else "bar"}"/><c:grouping val="clustered"/>'
                                '<c:varyColors val="0"/>', '<c:gapWidth val="219"/><c:overlap val="-27"/>')]
    elif kind == "combo":
        plots = [("bar", names[:-1], '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>',
                  '<c:gapWidth val="219"/><c:overlap val="-27"/>'),
                 ("line", names[-1:], '<c:grouping val="standard"/><c:varyColors val="0"/>', '<c:marker val="1"/>')]
    elif kind == "line":
        plots = [("line", names, '<c:grouping val="standard"/><c:varyColors val="0"/>', '<c:marker val="1"/>')]
    elif kind == "pie":
        plots = [("pie", names, '<c:varyColors val="1"/>', '<c:firstSliceAng val="0"/>')]
    elif kind == "doughnut":
        plots = [("doughnut", names, '<c:varyColors val="1"/>', '<c:firstSliceAng val="0"/><c:holeSize val="75"/>')]
    elif kind == "scatter":
        plots = [("scatter", names, '<c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>', "")]
    else:
        raise ValueError(kind)
    axes = kind not in ("pie", "doughnut")
    body = []
    for plot, members, head, tail in plots:
        tag = {"bar": "barChart", "line": "lineChart", "pie": "pieChart", "doughnut": "doughnutChart",
               "scatter": "scatterChart"}[plot]
        series = "".join(_series(spec, names.index(name), name, plot) for name in members)
        ax = '<c:axId val="510001"/><c:axId val="510002"/>' if axes else ""
        body.append(f"<c:{tag}>{head}{series}{tail}{ax}</c:{tag}>")
    if kind == "scatter":
        body.append(_SCATTER_AXES.format(c="510001", v="510002", vpos="l"))
    elif axes:
        horizontal = kind == "bar"
        body.append(_AXES[0].format(c="510001", v="510002", cpos="l" if horizontal else "b"))
        body.append(_AXES[1].format(c="510001", v="510002", vpos="b" if horizontal else "l", between="between"))
    legend = (f'<c:legend><c:legendPos val="{spec.legend}"/><c:overlay val="0"/></c:legend>'
              if spec.legend else "")
    external = ""
    if package_rel is not None:
        update = f'<c:autoUpdate val="{spec.auto_update}"/>' if spec.auto_update is not None else ""
        external = f'<c:externalData r:id="{package_rel}">{update}</c:externalData>'
    return (DECL + f'<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{R}" xmlns:c16r2="{C16R2}">'
            '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>'
            f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="c14" xmlns:c14="{C14}">'
            '<c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>'
            f"<c:chart>{_title(spec)}<c:plotArea><c:layout/>{''.join(body)}</c:plotArea>{legend}"
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            f'<c:extLst><c:ext uri="{DISPLAY_URI}" xmlns:c16r3="{C16R3}"><c16r3:dataDisplayOptions16>'
            '<c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chart>'
            '<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" '
            'algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/>'
            '</a:schemeClr></a:solidFill><a:round/></a:ln></c:spPr>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p>'
            f"</c:txPr>{external}</c:chartSpace>")


# -- SmartArt ----------------------------------------------------------------------------------


@dataclass
class DiagramSpec:
    """A SmartArt data model: ``layout`` a layout's ``uniqueId`` tail (``default`` is Basic
    Block List, ``vList2`` Vertical Bullet List), ``nodes`` ``(text, level)`` depth first."""

    layout: str
    nodes: list[tuple[str, int]]
    seed: int = 0
    definitions: dict[str, bytes] = field(default_factory=dict)


def _model_id(seed: int, k: int) -> str:
    return f"{{{0x5B000000 + seed * 0x1000:08X}-0000-4000-8000-{k:012X}}}"


def data_model(spec: DiagramSpec) -> str:
    """The data model alone: the document point, each node with its parent and sibling
    transition points, and their parent-of connections -- no presentation points, no
    drawing: Word writes those."""
    ids = iter(range(1, 10_000))
    doc_id = _model_id(spec.seed, next(ids))
    empty = '<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t>'
    points = [f'<dgm:pt modelId="{doc_id}" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/'
              f'layout/{spec.layout}" loCatId="list" qsTypeId="urn:microsoft.com/office/officeart/2005/8/quickstyle/'
              'simple1" qsCatId="simple" csTypeId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2" '
              f'csCatId="accent1"/><dgm:spPr/>{empty}</dgm:pt>']
    connections = []
    parents = [doc_id]
    order: dict[str, int] = {}
    for text, level in spec.nodes:
        node, par, sib, cxn = (_model_id(spec.seed, next(ids)) for _ in range(4))
        del parents[level + 1:]
        parent = parents[level]
        position = order.get(parent, 0)
        order[parent] = position + 1
        points.append(f'<dgm:pt modelId="{node}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/><dgm:t><a:bodyPr/>'
                      f'<a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{escape(text)}</a:t></a:r></a:p></dgm:t></dgm:pt>')
        points.append(f'<dgm:pt modelId="{par}" type="parTrans" cxnId="{cxn}"><dgm:prSet/><dgm:spPr/>{empty}</dgm:pt>')
        points.append(f'<dgm:pt modelId="{sib}" type="sibTrans" cxnId="{cxn}"><dgm:prSet/><dgm:spPr/>{empty}</dgm:pt>')
        connections.append(f'<dgm:cxn modelId="{cxn}" srcId="{parent}" destId="{node}" srcOrd="{position}" '
                           f'destOrd="0" parTransId="{par}" sibTransId="{sib}"/>')
        parents.append(node)
    return (DECL + f'<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}"><dgm:ptLst>{"".join(points)}</dgm:ptLst>'
            f'<dgm:cxnLst>{"".join(connections)}</dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>')


def definitions(layout: str) -> dict[str, bytes]:
    """The layout, quick-style and colour definitions for ``layout``, from pptx-agent's
    fixture (module docstring)."""
    with zipfile.ZipFile(SMARTART_SOURCE) as archive:
        names = archive.namelist()
        for n in (1, 2):
            data = archive.read(f"ppt/diagrams/layout{n}.xml")
            if f'layout/{layout}"'.encode() in data:
                return {"layout": data, "style": archive.read(f"ppt/diagrams/quickStyle{n}.xml"),
                        "colors": archive.read(f"ppt/diagrams/colors{n}.xml")}
    raise KeyError(f"no definition for {layout} among {names}")


# -- drawings in WordprocessingML --------------------------------------------------------------

WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


def chart_graphic(rel_id: str) -> str:
    return (f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{CHART_URI}">'
            f'<c:chart xmlns:c="{C}" xmlns:r="{R}" r:id="{rel_id}"/></a:graphicData></a:graphic>')


def diagram_graphic(rels: dict[str, str]) -> str:
    return (f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{DIAGRAM_URI}">'
            f'<dgm:relIds xmlns:dgm="{DGM}" xmlns:r="{R}" r:dm="{rels["data"]}" r:lo="{rels["layout"]}" '
            f'r:qs="{rels["style"]}" r:cs="{rels["colors"]}"/></a:graphicData></a:graphic>')


def inline(doc_pr: int, name: str, graphic: str, cx: int, cy: int) -> str:
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{doc_pr}" name="{name}"/>'
            f"<wp:cNvGraphicFramePr/>{graphic}</wp:inline></w:drawing></w:r>")


def anchored(doc_pr: int, name: str, graphic: str, cx: int, cy: int, *, x: int = 0, y: int = 0,
             height: int = 251659264) -> str:
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            f'relativeHeight="{height}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>'
            f'{x}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>{y}'
            f'</wp:posOffset></wp:positionV><wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapTopAndBottom/><wp:docPr id="{doc_pr}" name="{name}"/><wp:cNvGraphicFramePr/>{graphic}'
            "</wp:anchor></w:drawing></w:r>")


def text_box_graphic(content: str, cx: int, cy: int) -> str:
    """A text box (``wps:wsp`` with ``w:txbxContent``) whose content is ``content``'s blocks."""
    return (f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{WPS}"><wps:wsp xmlns:wps="{WPS}"><wps:cNvSpPr txBox="1"/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect">'
            '<a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="6350"><a:solidFill>'
            '<a:prstClr val="black"/></a:solidFill></a:ln></wps:spPr>'
            f"<wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>"
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" '
            'anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>')


def group_graphic(chart_rel: str, member_id: int, shape_id: int, cx: int, cy: int) -> str:
    """A group of a rectangle and a chart (``wpg:graphicFrame``), side by side in child
    space equal to its extent."""
    half = cx // 3
    return (f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{WPG}"><wpg:wgp xmlns:wpg="{WPG}" xmlns:wps="{WPS}">'
            f'<wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/>'
            f'<a:chOff x="0" y="0"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></wpg:grpSpPr>'
            f'<wps:wsp><wps:cNvPr id="{shape_id}" name="Rectangle {shape_id}"/><wps:cNvSpPr/><wps:spPr><a:xfrm>'
            f'<a:off x="0" y="0"/><a:ext cx="{half}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
            '</a:prstGeom><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>'
            f'<wpg:graphicFrame><wpg:cNvPr id="{member_id}" name="Chart {member_id}"/><wpg:cNvFrPr/>'
            f'<wpg:xfrm><a:off x="{half}" y="0"/><a:ext cx="{cx - half}" cy="{cy}"/></wpg:xfrm>'
            f"{chart_graphic(chart_rel)}</wpg:graphicFrame></wpg:wgp></a:graphicData></a:graphic>")


# -- a document holding them -------------------------------------------------------------------


class Builder:
    """Charts and diagrams added to a document's parts, with their relationships and
    content types: ``chart(part, spec)`` returns the relationship id from ``part``."""

    def __init__(self, parts: dict[str, bytes]) -> None:
        self.parts = parts
        self.rels: dict[str, list[tuple[str, str, str]]] = {}
        self.overrides: dict[str, str] = {}
        self.charts = 0
        self.diagrams = 0
        self.xlsx = False

    def _rel(self, part: str, rel_type: str, target: str) -> str:
        existing = self.rels.setdefault(part, [])
        rel_id = f"rIdC{len(existing) + 1}"
        existing.append((rel_id, rel_type, target))
        return rel_id

    def chart(self, part: str, spec: ChartSpec, *, embed: bool = True) -> str:
        self.charts += 1
        n = self.charts
        chart = f"word/charts/chart{n}.xml"
        self.overrides[f"/{chart}"] = CT_CHART
        if embed:
            name = "Microsoft_Excel_Worksheet" + ("" if n == 1 else str(n - 1)) + ".xlsx"
            self.parts[f"word/embeddings/{name}"] = workbook(spec.book or spec.data)
            self.parts[f"word/charts/_rels/chart{n}.xml.rels"] = relationships(
                [("rId1", REL_PACKAGE, f"../embeddings/{name}")]).encode()
            self.xlsx = True
        self.parts[chart] = chart_xml(spec, "rId1" if embed else None).encode()
        return self._rel(part, REL_CHART, "charts/" + chart.rsplit("/", 1)[1])

    def diagram(self, part: str, spec: DiagramSpec) -> dict[str, str]:
        self.diagrams += 1
        n = self.diagrams
        files = {"data": data_model(spec).encode(), **(spec.definitions or definitions(spec.layout))}
        out = {}
        stems = {"data": "data", "layout": "layout", "style": "quickStyle", "colors": "colors"}
        for key, data in files.items():
            name = f"word/diagrams/{stems[key]}{n}.xml"
            self.parts[name] = data
            self.overrides[f"/{name}"] = CT_DGM[key]
            out[key] = self._rel(part, REL_DGM[key], f"diagrams/{stems[key]}{n}.xml")
        return out

    def finish(self) -> dict[str, bytes]:
        parts = self.parts
        for part, rels in self.rels.items():
            directory, name = part.rsplit("/", 1)
            path = f"{directory}/_rels/{name}.rels"
            text = parts[path].decode() if path in parts else relationships([])
            text = text.replace("</Relationships>", "".join(
                f'<Relationship Id="{i}" Type="{t}" Target="{target}"/>' for i, t, target in rels) + "</Relationships>")
            parts[path] = text.encode()
        types = parts["[Content_Types].xml"].decode()
        extra = "".join(f'<Override PartName="{name}" ContentType="{kind}"/>' for name, kind in self.overrides.items())
        if self.xlsx and 'Extension="xlsx"' not in types:
            extra = f'<Default Extension="xlsx" ContentType="{CT_XLSX}"/>' + extra
        parts["[Content_Types].xml"] = types.replace("</Types>", extra + "</Types>").encode()
        return parts


def read_parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}
