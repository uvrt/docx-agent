#!/usr/bin/env python3
"""Measure what Word writes when it tracks changes, and what its Accept All and Reject All
make of them -- the facts docx-agent's tracking mode (E3) mirrors.

Probe documents are written here (mode 15, every paragraph and row with its paraId, so
Word keeps them and a case can be found again), and Word -- through ``tests/oracle.py``'s
machine-wide lock -- opens each, turns tracking on (``track revisions``), makes the case's
edits by AppleScript, and saves.  Each tracked result is then opened again and Word's own
``accept all revisions`` and ``reject all revisions`` are saved too.  A second document
makes the same paragraph-mark deletion with tracking off, for the join rule.

The cases (one paragraph or table each, named by its first word):

* text -- insert into a paragraph, delete a word, replace one, insert and then delete part
  of one's own insertion, delete across a paragraph mark, delete a whole paragraph, insert
  a paragraph, delete a field (a complex ``PAGE``), insert a picture;
* formatting -- bold a word (``w:rPrChange``), a paragraph style (``w:pPrChange``), an
  alignment, a list (``apply number default``), a section's margin (``w:sectPrChange``);
* moves -- a paragraph relocated in outline view;
* tables -- a row inserted, a row deleted, a column inserted, a column deleted, two cells
  merged across and two down, a cell split, a row's height, a cell's shading, a table's
  alignment;
* comments -- one added on a word (its parts as Word writes them);
* joins -- two paragraphs with different properties whose first mark is deleted, tracked
  (then accepted) and untracked.

    python tools/e3_probe.py           # writes tests/observations/e3-word.json

Nothing Word writes is committed: only the facts read from it, authors and dates
replaced by ``{author}``/``{date}`` and rsids dropped.
"""

from __future__ import annotations

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

import oracle  # noqa: E402
from make_fixtures import DECL, NS, REL, WML, content_types, png, relationships, settings, styles  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "e3-word.json"
COMMENT_STYLES = ROOT / "src" / "docx_agent" / "edit" / "word_comment_styles.py"
#: The styles Word adds with a document's first comment, by ``w:name``.
COMMENT_STYLE_NAMES = ("annotation text", "Annotation Text Char", "annotation reference",
                       "annotation subject", "Annotation Subject Char")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W

#: Paragraphs: (key, text, extra pPr).  Every case's paragraph starts with its key.
PARAGRAPHS = [
    ("Insert", "Insert alpha beta.", ""),
    ("Delete", "Delete gamma delta.", ""),
    ("Replace", "Replace epsilon zeta.", ""),
    ("Own", "Own eta theta.", ""),
    ("JoinA", "JoinA iota kappa.", '<w:jc w:val="left"/><w:spacing w:before="0" w:after="0"/>'),
    ("JoinB", "JoinB lambda mu.", '<w:jc w:val="right"/><w:spacing w:before="240" w:after="240"/>'),
    ("Mark", "Mark nu xi.", '<w:ind w:left="720"/>'),
    ("Marked", "Marked omicron pi.", '<w:jc w:val="center"/>'),
    ("Bold", "Bold rho sigma.", ""),
    ("Style", "Style tau upsilon.", ""),
    ("Align", "Align phi chi.", ""),
    ("After", "After psi omega.", ""),
    ("Gone", "Gone one two.", ""),
    ("Moved", "Moved three four.", ""),
    ("Stay", "Stay five six.", ""),
    ("Number", "Number seven eight.", ""),
    ("Field", None, ""),
    ("Resize", "Resize ten.", ""),
    ("Link", "Link twelve thirteen.", ""),
    ("Picture", "Picture nine.", ""),
    ("Comment", "Comment ten eleven.", ""),
]


def _para_id(k: int, row: bool = False) -> str:
    return f"{0x20000000 + (0x1000 if row else 0) + k:08X}"


def _paragraph(k: int, key: str, text: str | None, extra: str) -> str:
    ids = f' w14:paraId="{_para_id(k)}" w14:textId="77777777"'
    props = f"<w:pPr>{extra}</w:pPr>" if extra else ""
    if text is None:
        body = ('<w:r><w:t xml:space="preserve">Field page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>'
                '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> end.</w:t></w:r>')
    else:
        body = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
    return f"<w:p{ids}>{props}{body}</w:p>"


#: Tables: key -> rows x columns.  Cell text is ``<key> r<i>c<j>``.
TABLES = ["RowIns", "RowDel", "ColIns", "ColDel", "MergeH", "MergeV", "Split", "Props"]


def _table(t: int, key: str) -> str:
    rows = []
    for i in range(3):
        cells = "".join(
            f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
            f'<w:p w14:paraId="{_para_id(100 + t * 20 + i * 3 + j)}" w14:textId="77777777">'
            f'<w:r><w:t xml:space="preserve">{key} r{i + 1}c{j + 1}</w:t></w:r></w:p></w:tc>'
            for j in range(3))
        rows.append(f'<w:tr w14:paraId="{_para_id(t * 4 + i, row=True)}" w14:textId="77777777">{cells}</w:tr>')
    grid = '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
    props = '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
    after = f'<w:p w14:paraId="{_para_id(400 + t)}" w14:textId="77777777"/>'
    return f"<w:tbl>{props}{grid}{''.join(rows)}</w:tbl>{after}"


def _package(body: str, *, with_picture_rel: bool = False) -> bytes:
    section = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" '
               'w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    parts = {
        "[Content_Types].xml": content_types({"word/document.xml": WML + ".document.main+xml",
                                              "word/styles.xml": WML + ".styles+xml",
                                              "word/settings.xml": WML + ".settings+xml"}),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False)]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}{section}</w:body></w:document>",
        "word/styles.xml": styles(),
        "word/settings.xml": settings(15),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 10, 3, 12, 0, 0)), data)
    return buffer.getvalue()


def text_document() -> bytes:
    body = "".join(_paragraph(k, key, text, extra) for k, (key, text, extra) in enumerate(PARAGRAPHS))
    return _package(body)


def table_document() -> bytes:
    body = _paragraph(0, "Tables", "Tables follow.", "") + "".join(_table(t, key) for t, key in enumerate(TABLES))
    return _package(body)


#: The cell revision forms Word does not write itself, but may accept and reject.
CELL_CASES = ["CellIns", "CellDel", "CellMerge", "MergeAcross", "HMerge", "JoinX", "InsertMid", "InsertBeforeTable", "DeleteLast",
              "DeleteBeforeTable", "InsertLast", "InsertEnd"]
_AGENT = 'w:author="Agent" w:date="2026-10-03T12:00:00Z"'


def cells_document() -> bytes:
    """A column inserted with ``w:cellIns``, one deleted with ``w:cellDel``, two cells merged
    down with ``w:cellMerge``, and two merged across as docx-agent records it: the first
    cell's span in a ``w:tcPrChange``, the second cell ``w:cellDel``, its paragraph moved
    into the first as a deletion and an insertion."""
    counter = [500]

    def nid() -> int:
        counter[0] += 1
        return counter[0]

    def para(text: str, mark: str = "", deleted: bool = False, inserted: bool = False) -> str:
        properties = f"<w:pPr><w:rPr>{mark}</w:rPr></w:pPr>" if mark else ""
        run = ""
        if text:
            run = (f'<w:r><w:{"delText" if deleted else "t"} xml:space="preserve">{text}'
                   f'</w:{"delText" if deleted else "t"}></w:r>')
            if deleted:
                run = f'<w:del w:id="{nid()}" {_AGENT}>{run}</w:del>'
            if inserted:
                run = f'<w:ins w:id="{nid()}" {_AGENT}>{run}</w:ins>'
        return f'<w:p w14:paraId="{_para_id(nid())}" w14:textId="77777777">{properties}{run}</w:p>'

    def cell(content: str, extra: str = "", width: int = 2000) -> str:
        return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{extra}</w:tcPr>{content}</w:tc>'

    def table(rows: list[list[str]], grid: list[int], change: list[int] | None = None) -> str:
        cols = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid)
        if change is not None:
            cols += (f'<w:tblGridChange w:id="{nid()}"><w:tblGrid>'
                     + "".join(f'<w:gridCol w:w="{w}"/>' for w in change) + "</w:tblGrid></w:tblGridChange>")
        body = "".join(f'<w:tr w14:paraId="{_para_id(nid(), row=True)}" w14:textId="77777777">{"".join(r)}</w:tr>'
                       for r in rows)
        return (f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
                f"<w:tblGrid>{cols}</w:tblGrid>{body}</w:tbl>") + para("")

    out = []
    out.append(para("CellIns: column 3 inserted."))
    out.append(table([[cell(para(f"CellIns r{i}c1")), cell(para(f"CellIns r{i}c2")),
                       cell(para("", f'<w:ins w:id="{nid()}" {_AGENT}/>'), f'<w:cellIns w:id="{nid()}" {_AGENT}/>'),
                       cell(para(f"CellIns r{i}c3"))] for i in (1, 2, 3)], [2000] * 4, [2000] * 3))
    out.append(para("CellDel: column 2 deleted."))
    out.append(table([[cell(para(f"CellDel r{i}c1")),
                       cell(para(f"CellDel r{i}c2", f'<w:del w:id="{nid()}" {_AGENT}/>', deleted=True),
                            f'<w:cellDel w:id="{nid()}" {_AGENT}/>'),
                       cell(para(f"CellDel r{i}c3"))] for i in (1, 2, 3)], [2000] * 3))
    out.append(para("CellMerge: r1c1 and r2c1 merged down."))
    rows = []
    for i in (1, 2, 3):
        row = []
        for j in (1, 2, 3):
            extra = ""
            if j == 1 and i in (1, 2):
                extra = f'<w:cellMerge w:id="{nid()}" w:vMerge="{"rest" if i == 1 else "cont"}" {_AGENT}/>'
            row.append(cell(para(f"CellMerge r{i}c{j}"), extra))
        rows.append(row)
    out.append(table(rows, [2000] * 3))
    out.append(para("MergeAcross: r1c1 and r1c2 merged across."))
    first = cell(para("MergeAcross r1c1", f'<w:ins w:id="{nid()}" {_AGENT}/>')
                 + para("MergeAcross r1c2", inserted=True),
                 f'<w:gridSpan w:val="2"/><w:tcPrChange w:id="{nid()}" {_AGENT}><w:tcPr><w:tcW w:w="2000" '
                 'w:type="dxa"/></w:tcPr></w:tcPrChange>', width=4000)
    second = cell(para("MergeAcross r1c2", f'<w:del w:id="{nid()}" {_AGENT}/>', deleted=True),
                  f'<w:cellDel w:id="{nid()}" {_AGENT}/>')
    rows = [[first, second, cell(para("MergeAcross r1c3"))]]
    rows += [[cell(para(f"MergeAcross r{i}c{j}")) for j in (1, 2, 3)] for i in (2, 3)]
    out.append(table(rows, [2000] * 3))
    out.append(para("HMerge: r1c1 and r1c2 merged across with the legacy w:hMerge."))
    first = cell(para("HMerge r1c1", f'<w:ins w:id="{nid()}" {_AGENT}/>') + para("HMerge r1c2", inserted=True),
                 f'<w:hMerge w:val="restart"/><w:tcPrChange w:id="{nid()}" {_AGENT}><w:tcPr><w:tcW w:w="2000" '
                 'w:type="dxa"/></w:tcPr></w:tcPrChange>')
    second = cell(para("HMerge r1c2", deleted=True),
                  f'<w:hMerge w:val="continue"/><w:tcPrChange w:id="{nid()}" {_AGENT}><w:tcPr><w:tcW w:w="2000" '
                  'w:type="dxa"/></w:tcPr></w:tcPrChange>')
    rows = [[first, second, cell(para("HMerge r1c3"))]]
    rows += [[cell(para(f"HMerge r{i}c{j}")) for j in (1, 2, 3)] for i in (2, 3)]
    out.append(table(rows, [2000] * 3))
    # A deleted mark with no paragraph-properties change: whose properties does accepting keep?
    out.append('<w:p w14:paraId="2000F001" w14:textId="77777777"><w:pPr><w:ind w:left="1440"/><w:rPr>'
               f'<w:del w:id="{nid()}" {_AGENT}/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">JoinX first. </w:t>'
               '</w:r></w:p><w:p w14:paraId="2000F002" w14:textId="77777777"><w:pPr><w:jc w:val="center"/></w:pPr>'
               '<w:r><w:t>JoinY second.</w:t></w:r></w:p>')
    # A paragraph inserted after another as a paragraph of its own (its own mark inserted).
    out.append('<w:p w14:paraId="2000F003" w14:textId="77777777"><w:pPr><w:jc w:val="right"/></w:pPr><w:r>'
               '<w:t>InsertMid before.</w:t></w:r></w:p>'
               '<w:p w14:paraId="2000F004" w14:textId="77777777"><w:pPr><w:jc w:val="right"/><w:rPr>'
               f'<w:ins w:id="{nid()}" {_AGENT}/></w:rPr></w:pPr><w:ins w:id="{nid()}" {_AGENT}><w:r>'
               '<w:t>InsertMid new.</w:t></w:r></w:ins></w:p>'
               '<w:p w14:paraId="2000F005" w14:textId="77777777"><w:r><w:t>InsertMid after.</w:t></w:r></w:p>')
    # The same before a table, and at the end of the body (below).
    out.append('<w:p w14:paraId="2000F006" w14:textId="77777777"><w:r><w:t>InsertBeforeTable before.</w:t>'
               '</w:r></w:p><w:p w14:paraId="2000F007" w14:textId="77777777"><w:pPr><w:rPr>'
               f'<w:ins w:id="{nid()}" {_AGENT}/></w:rPr></w:pPr><w:ins w:id="{nid()}" {_AGENT}><w:r>'
               '<w:t>InsertBeforeTable new.</w:t></w:r></w:ins></w:p>')
    out.append(table([[cell(para("InsertBeforeTable cell."))]], [2000]))
    # A cell's last paragraph deleted, content and mark; and a paragraph before a table.
    out.append(table([[cell(para("DeleteLast keep.") + para("DeleteLast gone.", f'<w:del w:id="{nid()}" {_AGENT}/>',
                                                             deleted=True)),
                       cell(para("DeleteLast other."))]], [2000] * 2))
    out.append(para("DeleteBeforeTable gone.", f'<w:del w:id="{nid()}" {_AGENT}/>', deleted=True))
    out.append(table([[cell(para("DeleteBeforeTable cell."))]], [2000]))
    # The same at the end of a cell: no paragraph after it in its container.
    out.append(table([[cell(para("InsertLast cell.") + para("InsertLast new.", f'<w:ins w:id="{nid()}" {_AGENT}/>',
                                                             inserted=True)),
                       cell(para("InsertLast other."))]], [2000] * 2))
    out.append(para("InsertEnd before."))
    out.append(para("InsertEnd new.", f'<w:ins w:id="{nid()}" {_AGENT}/>', inserted=True))
    return _package("".join(out))


def join_document() -> bytes:
    """Two pairs of paragraphs with different properties, for an untracked join."""
    body = "".join(_paragraph(k, key, text, extra) for k, (key, text, extra) in enumerate(PARAGRAPHS[4:8]))
    return _package(body)


# -- the scripts ---------------------------------------------------------------------------------

def _locate(key: str) -> str:
    """AppleScript: ``k`` is the index of the paragraph whose text starts with ``key``."""
    return (f'set k to 0\n'
            f'      repeat with i from 1 to count of paragraphs of d\n'
            f'        if content of text object of paragraph i of d starts with "{key} " then\n'
            f'          set k to i\n'
            f'          exit repeat\n'
            f'        end if\n'
            f'      end repeat\n'
            f'      set s to start of content of text object of paragraph k of d')


def _at(offset_from: str, a: int, b: int | None = None) -> str:
    b = a if b is None else b
    return f"(create range d start ({offset_from} + {a}) end ({offset_from} + {b}))"


def _select(a: int, b: int | None = None) -> str:
    b = a if b is None else b
    return f"select (create range d start (s + {a}) end (s + {b}))"


#: Typing as a person types: over the selection, as Word's ``replace selection`` does.
BACKSPACE = "type backspace selection"


def _type(text: str) -> str:
    return f'type text selection text "{text}"'


def text_script(picture: bool) -> str:
    lines = []
    if picture:
        # A picture that is there before tracking starts, for the resize case.
        lines += [_locate("Resize"), f"make new inline picture at {_at('s', 7)} with properties "
                  f"{{file name:(item 3 of argv), link to file:false, save with document:true}}"]
    lines.append("set track revisions of d to true")

    def case(key: str, *steps: str) -> None:
        lines.append(_locate(key))
        lines.extend(steps)

    # Bottom up, so a case's edits never move the paragraphs of the cases above it.
    case("Comment", _select(8, 11), 'make new Word comment at d with properties {comment text:"Probe comment"}')
    if picture:
        case("Picture", f"make new inline picture at {_at('s', 8)} with properties "
                        f"{{file name:(item 3 of argv), link to file:false, save with document:true}}")
    case("Link", f'make new hyperlink object at d with properties {{text object:{_at("s", 5, 11)}, '
                 'hyperlink address:"https://example.com/"}')
    if picture:
        case("Resize", "set width of inline picture 1 of d to 72")
    case("Field", _select(6, 21), BACKSPACE)
    case("Number", "apply number default (list format of text object of paragraph k of d)")
    case("Moved", "set view type of view of active window to outline view",
         "relocate (text object of paragraph k of d) direction relocate down",
         "set view type of view of active window to print view")
    case("Gone", "select text object of paragraph k of d", BACKSPACE)
    case("After", "set e to end of content of text object of paragraph k of d",
         "select (create range d start (e - 1) end (e - 1))", "type paragraph selection",
         _type("Inserted paragraph."))
    case("Align", "set alignment of paragraph format of paragraph k of d to align paragraph center")
    case("Style", "set style of paragraph k of d to style heading2")
    case("Bold", f"set bold of font object of {_at('s', 5, 8)} to true")
    case("Mark", "set e to end of content of text object of paragraph k of d",
         "select (create range d start (e - 1) end e)", BACKSPACE)
    case("JoinA", "set e to end of content of text object of paragraph k of d",
         "select (create range d start (e - 7) end (e + 6))", BACKSPACE)
    case("Own", _select(4), _type("NEWTEXT "), _select(6, 9), BACKSPACE)
    case("Replace", _select(8, 15), _type("EPSILON"))
    case("Delete", _select(7, 13), BACKSPACE)
    case("Insert", _select(7), _type("INSERTED "))
    lines.append("set left margin of page setup of d to 90")
    return _script(lines)


def table_script() -> str:
    lines = ["set track revisions of d to true"]
    t = {key: k + 1 for k, key in enumerate(TABLES)}

    def cell(key: str, r: int, c: int) -> str:
        return f"(get cell from table (table {t[key]} of d) row {r} column {c})"

    lines += [
        f"set alignment of rows of table {t['Props']} of d to align row center",
        f"set texture of shading of {cell('Props', 2, 2)} to texture solid",
        f"set table item height (row 1 of table {t['Props']} of d) row height 40 height rule row height at least",
        f"split cell {cell('Split', 2, 2)} number of rows 1 number of columns 2",
        f"merge cell {cell('MergeV', 1, 1)} with {cell('MergeV', 2, 1)}",
        f"merge cell {cell('MergeH', 1, 1)} with {cell('MergeH', 1, 2)}",
        # Backspace on a selected column opens a dialog; cutting deletes it (through the
        # clipboard, which main() empties again).
        f"select text object of {cell('ColDel', 1, 2)}",
        "select column selection",
        "cut object selection",
        f"select text object of {cell('ColIns', 1, 2)}",
        "insert columns selection position insert on the right",
        f"select text object of row 2 of table {t['RowDel']} of d",
        BACKSPACE,
        f"select text object of {cell('RowIns', 1, 1)}",
        "insert rows selection position below",
    ]
    return _script(lines)


def join_script() -> str:
    """The JoinA/JoinB and Mark/Marked deletions, tracking off."""
    lines = []
    lines.append(_locate("Mark"))
    lines += ["set e to end of content of text object of paragraph k of d",
              "select (create range d start (e - 1) end e)", BACKSPACE]
    lines.append(_locate("JoinA"))
    lines += ["set e to end of content of text object of paragraph k of d",
              "select (create range d start (e - 7) end (e + 6))", BACKSPACE]
    return _script(lines)


def review_script(action: str) -> str:
    return _script([f"{action} all revisions d"])


def _script(lines: list[str]) -> str:
    body = "\n      ".join(line.replace("\n", "\n      ") for line in lines)
    return f'''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file inputPath)
      set d to active document
      {body}
      save as d file name outputPath file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
end run
'''


# -- reading ---------------------------------------------------------------------------------

_DROP = re.compile(r' w:rsid\w*="[^"]*"')


def normalise(xml: str) -> str:
    xml = re.sub(r' xmlns:\w+="[^"]*"', "", xml)
    xml = _DROP.sub("", xml)
    xml = re.sub(r'w:author="[^"]*"', 'w:author="{author}"', xml)
    xml = re.sub(r'w:date="[^"]*"', 'w:date="{date}"', xml)
    xml = re.sub(r'w:initials="[^"]*"', 'w:initials="{initials}"', xml)
    xml = re.sub(r'w15:userId="[^"]*"', 'w15:userId="{user}"', xml)
    xml = re.sub(r'w15:author="[^"]*"', 'w15:author="{author}"', xml)
    xml = re.sub(r'(w16du|w16cex):dateUtc="[^"]*"', r'\1:dateUtc="{date}"', xml)
    return xml


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def blocks(data: bytes) -> list[str]:
    """The body's blocks, normalised, each as one string."""
    root = etree.fromstring(parts(data)["word/document.xml"])
    body = root.find(_W + "body")
    return [normalise(etree.tostring(child, encoding="unicode")) for child in body]


def write_comment_styles(tracked: bytes) -> None:
    """The comment styles Word wrote into the tracked probe (it adds them with a document's
    first comment), as ``word_styles.py``'s entries are written (``e1_probe.read_styles``)."""
    import pprint

    from e1_probe import read_styles

    files = parts(tracked)
    measured = read_styles(files["word/styles.xml"], files.get("word/numbering.xml", b""))
    table = {name: measured[name] for name in COMMENT_STYLE_NAMES if name in measured}
    COMMENT_STYLES.write_text(
        '"""The styles Word 16.106 adds with a document\'s first comment, as it writes them.\n\n'
        "Generated by ``tools/e3_probe.py`` from the tracked probe Word saved (the observations are\n"
        "in ``tests/observations/e3-word.json``); do not edit by hand.  The same form as\n"
        "``word_styles.py``: keyed by ``w:name``, ``{id}`` and the references filled in when written.\n"
        '"""\n\n'
        f"STYLES: dict[str, dict] = {pprint.pformat(table, width=110)}\n")


def main() -> None:
    oracle.ORACLE_DIR.mkdir(parents=True, exist_ok=True)
    picture = oracle.ORACLE_DIR / "e3-probe.png"
    out: dict = {}
    with oracle.session() as session:
        picture.write_bytes(png(48, 24, (30, 120, 30)))
        try:
            runs = {
                "text": session.run_script(text_script(True), text_document(), name="e3-text", tag="tracked",
                                           args=[str(picture)], timeout=200),
                "tables": session.run_script(table_script(), table_document(), name="e3-tables", tag="tracked",
                                             timeout=200),
                "join-untracked": session.run_script(join_script(), join_document(), name="e3-join",
                                                     tag="untracked", timeout=200),
            }
        finally:
            picture.unlink(missing_ok=True)
        subprocess.run(["osascript", "-e", 'set the clipboard to ""'], check=False)
        runs["cells-pdf"] = session.export_pdf(cells_document(), name="e3-cells")
        for name in ("text", "tables", "cells"):
            if name != "cells" and not runs[name]:
                raise SystemExit(f"Word did not save {name}: {runs[name].outcome} {runs[name].detail}")
            tracked = runs[name].path.read_bytes() if name != "cells" else cells_document()
            for action in ("accept", "reject"):
                runs[f"{name}-{action}"] = session.run_script(review_script(action), tracked, name=f"e3-{name}",
                                                              tag=action, timeout=200)
    write_comment_styles(runs["text"].path.read_bytes())
    for name, outcome in runs.items():
        if not outcome:
            raise SystemExit(f"Word did not save {name}: {outcome.outcome} {outcome.detail}")
        if name.endswith("-pdf"):
            out[name] = {"opened": True, "text": oracle.pdf_pages(outcome.path)}
            continue
        data = outcome.path.read_bytes()
        files = parts(data)
        entry = {"blocks": blocks(data)}
        for part in sorted(files):
            if re.match(r"word/(comments\w*|people|settings)\.xml$", part):
                entry[part] = normalise(files[part].decode("utf-8"))
        rels = files.get("word/_rels/document.xml.rels", b"").decode("utf-8")
        entry["relationships"] = sorted(re.findall(r'Type="([^"]+)"', rels))
        entry["content_types"] = sorted(re.findall(r'PartName="([^"]+)"', files["[Content_Types].xml"].decode()))
        out[name] = entry
    OBSERVATIONS.write_text(json.dumps({
        "_about": ("What Word 16.106 for Mac wrote for tools/e3_probe.py: each probe's body blocks after "
                   "tracked edits, after its own Accept All and Reject All, and the untracked joins; the "
                   "comment, people and settings parts. Authors, dates, initials and user ids replaced by "
                   "placeholders, rsids dropped. Regenerate with python tools/e3_probe.py."),
        **out}, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
