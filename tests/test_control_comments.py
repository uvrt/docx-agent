"""A comment on a content control (round 48): ``word_comments`` add with a ``cc:`` target.

The comment covers the whole control as Word anchors one when the control is selected
(measured with Word 16.106 by selecting each control's whole range -- its start and end marks
included -- and adding a comment; ``test_word_anchors_a_selected_control_so``):

* a control in the line: the start right before ``w:sdt`` (before the paragraph when the
  control opens it), the end right after it, then the reference run;
* a block, row or cell control: the start before ``w:sdt`` with
  ``w:displacedByCustomXml="next"``, the end and reference at the end of the text of its last
  paragraph -- one character sooner when that paragraph is empty.

Word re-saves the result unchanged (``test_word_keeps_our_control_comments``).

Tool design rules: 1 (a control is a target, as it is for ``word_delete``), 2 (an unknown
control is refused naming the document's controls), 6 (Word's own form).  Made-up text.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.validate import check
from ooxml_edit.tools import Toolbox


def _p(text: str, pid: int) -> str:
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p w14:paraId="{pid:08X}" w14:textId="77777777">{run}</w:p>'


def _run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def _sdt(tag: str, number: int, inner: str, kind: str = "<w:richText/>") -> str:
    return (f'<w:sdt><w:sdtPr><w:tag w:val="{tag}"/><w:id w:val="{number}"/>{kind}</w:sdtPr><w:sdtEndPr/>'
            f'<w:sdtContent>{inner}</w:sdtContent></w:sdt>')


_GRID = ('<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/>'
         '</w:tblPr><w:tblGrid><w:gridCol w:w="4536"/><w:gridCol w:w="4536"/></w:tblGrid>')


def _tc(inner: str) -> str:
    return f'<w:tc><w:tcPr><w:tcW w:w="4536" w:type="dxa"/></w:tcPr>{inner}</w:tc>'


def _levels_body() -> str:
    """Controls of every level: a table in a block control, a cell and a row control, controls
    in the line at a paragraph's start, end and middle, a block control ending in an empty
    paragraph, a block control holding one in the line."""
    n = iter(range(0x10000001, 0x10000100))
    return "".join([
        _p("Alpha before.", next(n)),
        _sdt("tbl", 101, f'<w:tbl>{_GRID}<w:tr>{_tc(_p("A", next(n)))}{_tc(_p("B", next(n)))}</w:tr>'
                         f'<w:tr>{_tc(_p("one", next(n)))}{_tc(_p("two", next(n)))}</w:tr></w:tbl>'),
        _p("Between.", next(n)),
        f'<w:tbl>{_GRID}<w:tr>{_sdt("cell", 102, _tc(_p("C", next(n))))}{_tc(_p("D", next(n)))}</w:tr>'
        + _sdt("row", 103, f'<w:tr>{_tc(_p("three", next(n)))}{_tc(_p("four", next(n)))}</w:tr>')
        + f'<w:tr>{_tc(_p("five", next(n)))}{_tc(_p("six", next(n)))}</w:tr></w:tbl>',
        _p("After table.", next(n)),
        f'<w:p w14:paraId="{next(n):08X}" w14:textId="77777777">{_sdt("startline", 104, _run("start value"), "<w:text/>")}'
        f'{_run(" rest of line.")}</w:p>',
        f'<w:p w14:paraId="{next(n):08X}" w14:textId="77777777">{_run("Line ends with ")}'
        f'{_sdt("endline", 105, _run("end value"), "<w:text/>")}</w:p>',
        _sdt("emptylast", 106, _p("Filled.", next(n)) + _p("", next(n))),
        _sdt("outer", 107, f'<w:p w14:paraId="{next(n):08X}" w14:textId="77777777">{_run("Outer text ")}'
                           f'{_sdt("inner", 108, _run("inner value"), "<w:text/>")}{_run(" more.")}</w:p>'),
        f'<w:p w14:paraId="{next(n):08X}" w14:textId="77777777">{_run("Rich ")}{_sdt("inlinerich", 109, _run("rich value"))}'
        f'{_run(" after.")}</w:p>',
        _p("Omega after.", next(n))])


def _edges_body() -> str:
    """The edges: a control first in the body, one opening a centred paragraph, an empty one,
    one ending in two empty paragraphs, controls first in a cell's paragraph and as a cell's
    block, one opening the paragraph of a block control, one showing its placeholder, the last."""
    n = iter(range(0x20000001, 0x20000100))

    def paragraph(inner: str, properties: str = "") -> str:
        return f'<w:p w14:paraId="{next(n):08X}" w14:textId="77777777">{properties}{inner}</w:p>'

    return "".join([
        _sdt("first", 201, _p("First block.", next(n))),
        paragraph(_sdt("startppr", 202, _run("centred value"), "<w:text/>") + _run(" tail."),
                  '<w:pPr><w:jc w:val="center"/></w:pPr>'),
        _sdt("oneempty", 203, _p("", next(n))),
        _p("Spacer one.", next(n)),
        _sdt("twoempty", 204, _p("Text here.", next(n)) + _p("", next(n)) + _p("", next(n))),
        f'<w:tbl>{_GRID}<w:tr>'
        + _tc(paragraph(_sdt("cellstart", 205, _run("cell value"), "<w:text/>") + _run(" in cell.")))
        + _tc(_sdt("cellblock", 206, _p("Block in cell.", next(n)))) + '</w:tr></w:tbl>',
        _sdt("nestouter", 207, paragraph(_sdt("neststart", 208, _run("nested value"), "<w:text/>")
                                         + _run(" after nested."))),
        paragraph(_run("Place ") + '<w:sdt><w:sdtPr><w:tag w:val="plc"/><w:id w:val="209"/><w:showingPlcHdr/>'
                  '<w:text/></w:sdtPr><w:sdtEndPr/><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/>'
                  '</w:rPr><w:t>Click here.</w:t></w:r></w:sdtContent></w:sdt>' + _run(" end.")),
        _p("Spacer two.", next(n)),
        _sdt("last", 210, _p("Last block.", next(n)))])


def _document(body: str) -> bytes:
    data = Document.new().to_bytes()
    source = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for item in source.infolist():
            raw = source.read(item.filename)
            if item.filename == "word/document.xml":
                xml = raw.decode()
                start = xml.index("<w:body>") + len("<w:body>")
                raw = (xml[:start] + body + xml[xml.index("<w:sectPr"):]).encode()
            archive.writestr(item, raw)
    return out.getvalue()


_SIMPLE = [(re.compile(r' w(?:14)?:(?:rsid\w*|paraId|textId)="[^"]*"| xml:space="preserve"'), ""),
           (re.compile(r"<w:(tblPr|tblGrid|tcPr|sdtPr|sectPr)>.*?</w:\1>"), ""),
           (re.compile(r'<w:rPr><w:rStyle w:val="[^"]*"/><w:sz w:val="(\d+)"/><w:szCs w:val="\d+"/></w:rPr>'),
            r"<REF \1>"),
           (re.compile(r"<w:rPr>.*?</w:rPr>"), "")]


def simple(data: bytes) -> str:
    """The body's blocks, rows, cells, text and comment marks; control properties dropped (Word
    writes a rich text control's without ``w:richText``); comment ids in order of appearance."""
    xml = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml").decode()
    body = xml[xml.index("<w:body>"):xml.index("</w:body>")]
    for pattern, replacement in _SIMPLE:
        body = pattern.sub(replacement, body)
    numbers: dict[str, str] = {}
    return re.sub(r'(commentRange(?:Start|End)|commentReference) w:id="(\d+)"',
                  lambda m: f'{m.group(1)} w:id="{numbers.setdefault(m.group(2), str(len(numbers)))}"', body)


#: Word 16.106's bodies after commenting on each named control selected whole (the oracle
#: test below measures the first again), simplified as :func:`simple` does.
WORD_LEVELS = (
    '<w:body><w:p><w:r><w:t>Alpha before.</w:t></w:r></w:p>'
    '<w:commentRangeStart w:id="0" w:displacedByCustomXml="next"/><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc>'
    '<w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc>'
    '<w:p><w:r><w:t>one</w:t></w:r></w:p></w:tc><w:tc>'
    '<w:p><w:r><w:t>two</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><REF 24><w:commentReference w:id="0"/></w:r></w:p></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Between.</w:t></w:r></w:p><w:tbl><w:tr>'
    '<w:commentRangeStart w:id="1" w:displacedByCustomXml="next"/>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:tc>'
    '<w:p><w:r><w:t>C</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><REF 24><w:commentReference w:id="1"/></w:r></w:p></w:tc></w:sdtContent></w:sdt><w:tc>'
    '<w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr>'
    '<w:commentRangeStart w:id="2" w:displacedByCustomXml="next"/>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>three</w:t></w:r></w:p></w:tc><w:tc>'
    '<w:p><w:r><w:t>four</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><REF 24><w:commentReference w:id="2"/></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt><w:tr><w:tc>'
    '<w:p><w:r><w:t>five</w:t></w:r></w:p></w:tc><w:tc>'
    '<w:p><w:r><w:t>six</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    '<w:p><w:r><w:t>After table.</w:t></w:r></w:p><w:commentRangeStart w:id="3"/><w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>start value</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="3"/><w:r><REF 24><w:commentReference w:id="3"/></w:r><w:r><w:t> rest of line.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Line ends with </w:t></w:r><w:commentRangeStart w:id="4"/>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>end value</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="4"/><w:r><REF 24><w:commentReference w:id="4"/></w:r></w:p>'
    '<w:commentRangeStart w:id="5" w:displacedByCustomXml="next"/><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Filled.</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><REF 24><w:commentReference w:id="5"/></w:r></w:p>'
    '<w:p/></w:sdtContent></w:sdt><w:commentRangeStart w:id="6" w:displacedByCustomXml="next"/>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:p><w:r><w:t>Outer text </w:t></w:r>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>inner value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> more.</w:t></w:r><w:commentRangeEnd w:id="6"/><w:r><REF 24><w:commentReference w:id="6"/></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Rich </w:t></w:r><w:commentRangeStart w:id="7"/>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>rich value</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="7"/><w:r><REF 24><w:commentReference w:id="7"/></w:r><w:r><w:t> after.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Omega after.</w:t></w:r></w:p>')

WORD_EDGES = (
    '<w:body><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>First block.</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>centred value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> tail.</w:t></w:r></w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:p/></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Spacer one.</w:t></w:r></w:p>'
    '<w:commentRangeStart w:id="0" w:displacedByCustomXml="next"/><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Text here.</w:t></w:r></w:p><w:commentRangeEnd w:id="0"/>'
    '<w:p><w:r><REF 24><w:commentReference w:id="0"/></w:r></w:p><w:p/></w:sdtContent></w:sdt>'
    '<w:tbl><w:tr><w:tc><w:commentRangeStart w:id="1"/><w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>cell value</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="1"/><w:r><REF 24><w:commentReference w:id="1"/></w:r><w:r><w:t> in cell.</w:t></w:r></w:p></w:tc><w:tc>'
    '<w:commentRangeStart w:id="2" w:displacedByCustomXml="next"/><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Block in cell.</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><REF 24><w:commentReference w:id="2"/></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:commentRangeStart w:id="3" w:displacedByCustomXml="prev"/><w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>nested value</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="3"/><w:r><REF 24><w:commentReference w:id="3"/></w:r><w:r><w:t> after nested.</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Place </w:t></w:r><w:commentRangeStart w:id="4"/>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>Click here.</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="4"/><w:r><REF 24><w:commentReference w:id="4"/></w:r><w:r><w:t> end.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Spacer two.</w:t></w:r></w:p>'
    '<w:commentRangeStart w:id="5" w:displacedByCustomXml="next"/><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Last block.</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><REF 24><w:commentReference w:id="5"/></w:r></w:p></w:sdtContent></w:sdt>')

WORD_FIRST = (
    '<w:body><w:commentRangeStart w:id="0" w:displacedByCustomXml="next"/>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>First block.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><REF 24><w:commentReference w:id="0"/></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:commentRangeStart w:id="1"/><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>centred value</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="1"/><w:r><REF 24><w:commentReference w:id="1"/></w:r><w:r><w:t> tail.</w:t></w:r></w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:p/></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Spacer one.</w:t></w:r></w:p><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Text here.</w:t></w:r></w:p><w:p/><w:p/></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc>'
    '<w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>cell value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> in cell.</w:t></w:r></w:p></w:tc><w:tc>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Block in cell.</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>nested value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> after nested.</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Place </w:t></w:r>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>Click here.</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> end.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Spacer two.</w:t></w:r></w:p><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Last block.</w:t></w:r></w:p></w:sdtContent></w:sdt>')

WORD_EMPTY = (
    '<w:body><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>First block.</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>centred value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> tail.</w:t></w:r></w:p>'
    '<w:commentRangeStart w:id="0" w:displacedByCustomXml="next"/>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:commentRangeEnd w:id="0" w:displacedByCustomXml="prev"/>'
    '<w:p><w:r><REF 24><w:commentReference w:id="0"/></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Spacer one.</w:t></w:r></w:p><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Text here.</w:t></w:r></w:p><w:p/><w:p/></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc>'
    '<w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>cell value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> in cell.</w:t></w:r></w:p></w:tc><w:tc>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Block in cell.</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>nested value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> after nested.</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Place </w:t></w:r>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>Click here.</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> end.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Spacer two.</w:t></w:r></w:p><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Last block.</w:t></w:r></w:p></w:sdtContent></w:sdt>')

WORD_NESTED = (
    '<w:body><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>First block.</w:t></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>centred value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> tail.</w:t></w:r></w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:p/></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Spacer one.</w:t></w:r></w:p><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Text here.</w:t></w:r></w:p><w:p/><w:p/></w:sdtContent></w:sdt><w:tbl><w:tr><w:tc>'
    '<w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>cell value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> in cell.</w:t></w:r></w:p></w:tc><w:tc>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Block in cell.</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>'
    '<w:commentRangeStart w:id="0" w:displacedByCustomXml="next"/><w:sdt><w:sdtEndPr/><w:sdtContent><w:p>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>nested value</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> after nested.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><REF 24><w:commentReference w:id="0"/></w:r></w:p></w:sdtContent></w:sdt>'
    '<w:p><w:r><w:t>Place </w:t></w:r>'
    '<w:sdt><w:sdtEndPr/><w:sdtContent><w:r><w:t>Click here.</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> end.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Spacer two.</w:t></w:r></w:p><w:sdt><w:sdtEndPr/><w:sdtContent>'
    '<w:p><w:r><w:t>Last block.</w:t></w:r></w:p></w:sdtContent></w:sdt>')


#: The controls each Word body above commented on, in order, and the body it starts from.
CASES = {"levels": (_levels_body, ["tbl", "cell", "row", "startline", "endline", "emptylast", "outer", "inlinerich"],
                    WORD_LEVELS),
         "edges": (_edges_body, ["twoempty", "cellstart", "cellblock", "neststart", "plc", "last"], WORD_EDGES),
         "first": (_edges_body, ["first", "startppr"], WORD_FIRST),
         "empty": (_edges_body, ["oneempty"], WORD_EMPTY),
         "nested": (_edges_body, ["nestouter"], WORD_NESTED)}


def commented(body, tags: list[str]) -> bytes:
    """The controls ``tags`` name commented on through the tool, in order."""
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session()
        session.open(_document(body()), "controls.docx")
        found = session.entry("d1").document
        by_tag = {c.tag: c.id for c in found.content_controls()}
        result = box.dispatch(session, "word_comments", {"doc": "d1", "action": "add", "items": [
            {"target": by_tag[tag], "text": f"On {tag}"} for tag in tags]})
        assert result.ok, result.error
        assert len(result.data["comments"]) == len(tags)
        return found.to_bytes()


@pytest.mark.parametrize("case", sorted(CASES))
def test_a_comment_on_a_control_is_anchored_as_word_anchors_one(case):
    body, tags, word = CASES[case]
    data = commented(body, tags)
    assert simple(data) == word
    before = check(Document.open(_document(body())).package)   # the made-up styles' names
    assert check(Document.open(data).package) == before


def test_the_comment_covers_the_controls_text():
    done = Document.open(commented(_levels_body, CASES["levels"][1]))
    assert [(c.text, c.anchor.text) for c in done.comments()] == [
        ("On tbl", "A\nB\none\ntwo"), ("On cell", "C"), ("On row", "three\nfour"),
        ("On startline", "start value"), ("On endline", "end value"),
        ("On emptylast", "Filled."), ("On outer", "Outer text inner value more."),
        ("On inlinerich", "rich value")]


def test_a_comment_on_a_control_can_be_replied_to_and_deleted():
    before = check(Document.open(_document(_edges_body())).package)
    document = Document.open(commented(_edges_body, ["first", "oneempty", "startppr"]))
    for comment in document.comments(replies=False):
        document.reply_to_comment(comment.id, "Seen.")
    assert check(document.package) == before
    for comment in document.comments(replies=False):
        document.delete_comment(comment.id)
    assert document.comments() == []
    assert "comment" not in simple(document.to_bytes())
    assert check(document.package) == before


def test_an_unknown_control_is_refused_naming_the_documents_controls():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session = box.session()
        session.open(_document(_edges_body()), "controls.docx")
        result = box.dispatch(session, "word_comments", {"doc": "d1", "action": "add", "items": [
            {"target": "cc:4242", "text": "Where?"}]})
    assert not result.ok and result.error.code == "invalid_arguments"
    assert "cc:201" in result.error.message and "cc:210" in result.error.message
    assert result.error.valid_options[:3] == ["cc:201", "cc:202", "cc:203"]


def test_removing_a_commented_control_keeps_its_comment():
    before = check(Document.open(_document(_edges_body())).package)
    document = Document.open(commented(_edges_body, ["first", "last", "startppr", "oneempty"]))
    for identifier in ("cc:201", "cc:210", "cc:202", "cc:203"):
        document.remove_control(identifier)
    assert [(c.text, c.anchor.text) for c in document.comments()] == [
        ("On first", "First block."), ("On last", "Last block."), ("On startppr", "centred value"),
        ("On oneempty", "")]
    # The marks Word wrote beside the controls no longer say they are beside one.
    assert "displacedByCustomXml" not in simple(document.to_bytes())
    assert check(document.package) == before
    Document.open(document.to_bytes())


# -- Word itself -----------------------------------------------------------------------------


#: Each control's whole range in Word's story of :func:`_levels_body` (each control's start
#: and end take a character; measured by reading the story character by character).
LEVELS_RANGES = [(14, 28), (39, 42), (46, 58), (83, 96), (126, 137), (138, 148), (149, 181), (187, 199)]


def _oracle():
    oracle = pytest.importorskip("oracle")
    if not oracle.available():
        pytest.skip("Microsoft Word or docx2svg's oracle scripts are not available")
    return oracle


@pytest.mark.oracle
def test_word_anchors_a_selected_control_so():
    """Word itself: select each control whole and comment (the measurement of
    :data:`WORD_LEVELS`)."""
    oracle = _oracle()
    script = '''
on run argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file (item 1 of argv))
      set d to active document
      set rs to {}
      repeat with k from 3 to (count of argv) by 2
        set end of rs to (create range d start ((item k of argv) as integer) end ((item (k + 1) of argv) as integer))
      end repeat
      repeat with k from 1 to (count of rs)
        select (item k of rs)
        make new Word comment at d with properties {comment text:("C" & k)}
      end repeat
      save as d file name (item 2 of argv) file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
end run
'''
    args = [str(x) for pair in LEVELS_RANGES for x in pair]
    try:
        with oracle.session() as word:
            out = word.run_script(script, _document(_levels_body()), name="r48-control-comment", tag="word",
                                  args=args, timeout=170)
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))
    assert out.ok, out.detail
    assert simple(out.path.read_bytes()) == WORD_LEVELS


@pytest.mark.oracle
@pytest.mark.parametrize("case", sorted(CASES))
def test_word_keeps_our_control_comments(case):
    """Word opens ours without a prompt, and its re-save keeps every mark where we put it."""
    oracle = _oracle()
    body, tags, _word = CASES[case]
    data = commented(body, tags)
    try:
        with oracle.session() as word:
            out = word.resave(data, name=f"r48-ours-{case}")
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))
    assert out.ok, (out.outcome, out.detail)
    assert simple(out.path.read_bytes()) == simple(data)
