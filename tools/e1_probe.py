#!/usr/bin/env python3
"""Measure what Word writes for the E1 operations docx-agent must write as Word does:
built-in styles on first use, its default bullet and number lists, and an inserted
picture's size.

A probe document is written here (mode 15, the handful of styles ``make_fixtures.py``
writes, no latent styles), and Word -- through ``tests/oracle.py``'s machine-wide lock --
opens it and, by AppleScript:

* applies each built-in paragraph style to a paragraph of its own and each built-in
  character style to one, by Word's language-independent constants (``style heading2``),
  since this Word's interface is Dutch and style *names* are localised there;
* applies its default bullets (``apply bullet default``) to two paragraphs and its default
  numbering (``apply number default``) to two more;
* inserts two pictures written here: one stating no density, one stating 300 dpi;

and saves.  What it wrote is read back:

* every style it defined, keyed by ``w:name`` (the English canonical name, whatever the
  interface language), with ``w:basedOn``, ``w:next`` and ``w:link`` turned from style ids
  -- which this Word localises (``Kop2``) -- into the names they point at, and ``w:rsid``
  dropped (docx-agent writes no rsids: ROADMAP.md, decision 6);
* the list definitions and which style the listed paragraphs got;
* each picture's extent.

    python tools/e1_probe.py        # writes tests/observations/e1-word.json and
                                    # src/docx_agent/edit/word_styles.py

Nothing Word writes is committed: only these facts.
"""

from __future__ import annotations

import io
import json
import pprint
import re
import struct
import sys
import zipfile
import zlib
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))

import oracle  # noqa: E402
from make_fixtures import DECL, NS, REL, WML, content_types, p, relationships, settings  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "e1-word.json"
TABLE = ROOT / "src" / "docx_agent" / "edit" / "word_styles.py"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W

#: Word's AppleScript constants for the built-in styles measured.
PARAGRAPH_STYLES = [
    "heading1", "heading2", "heading3", "heading4", "heading5", "heading6", "heading7", "heading8",
    "heading9", "title", "subtitle", "quote", "intense quote", "list paragraph", "list bullet",
    "list bullet2", "list bullet3", "list bullet4", "list bullet5", "list number", "list number2",
    "list number3", "list number4", "list number5", "caption", "html pre", "block quotation",
    "body text", "footnote text", "endnote text", "header", "footer", "toc heading", "plain text",
    "bibliography",
]
CHARACTER_STYLES = [
    "strong", "emphasis", "subtle emphasis", "intense emphasis", "book title", "subtle reference",
    "intense reference", "hyperlink", "hyperlink followed", "html code", "footnote reference",
    "endnote reference",
]
EXTRA = 6  # bullets x2, numbers x2, two pictures


def png(width: int, height: int, color: tuple[int, int, int], dpi: int | None = None) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(color) * width for _ in range(height))
    out = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    if dpi:
        per_metre = int(round(dpi / 0.0254))
        out += chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    return out + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def minimal_styles() -> str:
    """Only what every document has: the defaults, Normal, Default Paragraph Font, Normal
    Table and No List -- so Word defines every built-in style the probe applies itself."""
    return (
        f'{DECL}<w:styles xmlns:w="{W}">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" '
        'w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
        '<w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>'
        '</w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
        '<w:qFormat/></w:style>'
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/>'
        '<w:unhideWhenUsed/></w:style>'
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr></w:style>'
        '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
        "</w:styles>"
    )


def display_name(name: str) -> str:
    """How English Word shows a built-in style's name: each word capitalised."""
    return " ".join(word[:1].upper() + word[1:] for word in name.split())


def probe_document() -> bytes:
    count = len(PARAGRAPH_STYLES) + len(CHARACTER_STYLES) + EXTRA
    body = "".join(p(f"Probe paragraph {k}", f"1{k:07X}") for k in range(1, count + 1))
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" '
             'w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    parts = {
        "[Content_Types].xml": content_types({"word/document.xml": WML + ".document.main+xml",
                                              "word/styles.xml": WML + ".styles+xml",
                                              "word/settings.xml": WML + ".settings+xml"}),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False)]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": minimal_styles(),
        "word/settings.xml": settings(15),
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 10, 3, 12, 0, 0)), data)
    return buffer.getvalue()


def script() -> str:
    lines = []
    k = 1
    for constant in PARAGRAPH_STYLES:
        lines.append(f"set style of paragraph {k} of d to style {constant}")
        k += 1
    for constant in CHARACTER_STYLES:
        lines.append(f"set style of (create range d start (start of content of text object of paragraph {k} of d) "
                     f"end ((end of content of text object of paragraph {k} of d) - 1)) to style {constant}")
        k += 1
    lines.append(f"apply bullet default (list format of (create range d start (start of content of text object of "
                 f"paragraph {k} of d) end (end of content of text object of paragraph {k + 1} of d)))")
    lines.append(f"apply number default (list format of (create range d start (start of content of text object of "
                 f"paragraph {k + 2} of d) end (end of content of text object of paragraph {k + 3} of d)))")
    for offset, argument in ((4, "item 3 of argv"), (5, "item 4 of argv")):
        lines.append(f"set s to start of content of text object of paragraph {k + offset} of d")
        lines.append(f"make new inline picture at (create range d start s end s) with properties "
                     f"{{file name:({argument}), link to file:false, save with document:true}}")
    body = "\n      ".join(lines)
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


def english_id(name: str) -> str:
    """The id English Word gives a built-in style: its name, each word capitalised, without
    spaces -- except Normal Table's and No List's, which Word spells otherwise."""
    special = {"Normal Table": "TableNormal", "normal table": "TableNormal", "No List": "NoList"}
    if name in special:
        return special[name]
    return "".join(word[:1].upper() + word[1:] for word in name.split())


def _strip(node) -> str:
    text = etree.tostring(node, encoding="unicode")
    return re.sub(r' xmlns:\w+="[^"]*"', "", text).replace(f"<w:{etree.QName(node).localname} ",
                                                              f'<w:{etree.QName(node).localname} xmlns:w="{W}" ', 1)


def read_styles(xml: bytes, numbering: bytes) -> dict[str, dict]:
    root = etree.fromstring(xml)
    by_id = {s.get(_W + "styleId"): s for s in root.findall(_W + "style")}
    names = {sid: (s.find(_W + "name").get(_W + "val") if s.find(_W + "name") is not None else sid)
             for sid, s in by_id.items()}
    # A linked character style's name is localised ("Kop 2 Char"): it is its paragraph
    # style's English display name and " Char", as English Word names it.
    for sid, node in by_id.items():
        link = node.find(_W + "link")
        if node.get(_W + "type") == "character" and link is not None and link.get(_W + "val") in by_id:
            names[sid] = display_name(names[link.get(_W + "val")]) + " Char"
    instances = {}
    abstracts = {}
    if numbering:
        tree = etree.fromstring(numbering)
        abstracts = {n.get(_W + "abstractNumId"): n for n in tree.findall(_W + "abstractNum")}
        instances = {n.get(_W + "numId"): n.find(_W + "abstractNumId").get(_W + "val") for n in tree.findall(_W + "num")}
    out = {}
    for sid, node in by_id.items():
        node = etree.fromstring(etree.tostring(node))
        refs = {}
        for tag in ("basedOn", "next", "link"):
            child = node.find(_W + tag)
            if child is not None:
                refs[tag] = names.get(child.get(_W + "val"), child.get(_W + "val"))
                child.set(_W + "val", "{%s}" % tag)
        for child in node.findall(_W + "rsid"):
            node.remove(child)
        node.set(_W + "styleId", "{id}")
        node.find(_W + "name").set(_W + "val", names[sid])
        entry = {"type": node.get(_W + "type"), "refs": refs}
        num = node.find(f"{_W}pPr/{_W}numPr/{_W}numId")
        if num is not None and instances.get(num.get(_W + "val")) in abstracts:
            # A list style's own list: an abstract definition whose level names the style.
            abstract = etree.fromstring(etree.tostring(abstracts[instances[num.get(_W + "val")]]))
            abstract.set(_W + "abstractNumId", "{abstract}")
            for tag, placeholder in (("nsid", "{nsid}"), ("tmpl", "{tmpl}")):
                child = abstract.find(_W + tag)
                if child is not None:
                    child.set(_W + "val", placeholder)
            for pstyle in abstract.iter(_W + "pStyle"):
                pstyle.set(_W + "val", "{id}")
            for name in list(abstract.attrib):
                if not name.startswith(_W):
                    del abstract.attrib[name]  # w15:restartNumberingAfterBreak: Word's own bookkeeping
            num.set(_W + "val", "{num}")
            entry["numbering"] = _strip(abstract)
        entry["xml"] = _strip(node).replace("<w:style ", "<w:style ", 1)
        out[names[sid]] = entry
    return out


def main() -> None:
    oracle.ORACLE_DIR.mkdir(parents=True, exist_ok=True)
    pictures = [oracle.ORACLE_DIR / "e1-probe-nodpi.png", oracle.ORACLE_DIR / "e1-probe-300dpi.png"]
    data = probe_document()
    with oracle.session() as session:
        pictures[0].write_bytes(png(96, 48, (200, 30, 30)))
        pictures[1].write_bytes(png(300, 150, (30, 30, 200), dpi=300))
        try:
            outcome = session.run_script(script(), data, name="e1-probe", tag="builtins",
                                         args=[str(pictures[0]), str(pictures[1])], timeout=200)
        finally:
            for picture in pictures:
                picture.unlink(missing_ok=True)
    if not outcome:
        raise SystemExit(f"Word did not save the probe: {outcome.outcome} {outcome.detail}")
    with zipfile.ZipFile(outcome.path) as archive:
        saved_styles = archive.read("word/styles.xml")
        document = etree.fromstring(archive.read("word/document.xml"))
        numbering = archive.read("word/numbering.xml") if "word/numbering.xml" in archive.namelist() else b""
    outcome.path.unlink(missing_ok=True)

    measured = read_styles(saved_styles, numbering)
    ours = {"Normal", "Default Paragraph Font", "Normal Table", "No List"}
    paragraphs = document.find(_W + "body").findall(_W + "p")
    style_names = {sid: name for sid, name in
                   ((s.get(_W + "styleId"), s.find(_W + "name").get(_W + "val"))
                    for s in etree.fromstring(saved_styles).findall(_W + "style"))}
    applied = []
    for k, paragraph in enumerate(paragraphs[:len(PARAGRAPH_STYLES)]):
        node = paragraph.find(f"{_W}pPr/{_W}pStyle")
        applied.append((PARAGRAPH_STYLES[k], style_names.get(node.get(_W + "val")) if node is not None else None))
    k = len(PARAGRAPH_STYLES)
    for constant in CHARACTER_STYLES:
        node = paragraphs[k].find(f".//{_W}rStyle")
        applied.append((constant, style_names.get(node.get(_W + "val")) if node is not None else None))
        k += 1
    lists = []
    for paragraph in paragraphs[k:k + 4]:
        style = paragraph.find(f"{_W}pPr/{_W}pStyle")
        num = paragraph.find(f"{_W}pPr/{_W}numPr/{_W}numId")
        lists.append({"style": style_names.get(style.get(_W + "val")) if style is not None else None,
                      "numId": num.get(_W + "val") if num is not None else None})
    extents = []
    for paragraph in paragraphs[k + 4:k + 6]:
        extent = paragraph.find(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
        extents.append({"cx": int(extent.get("cx")), "cy": int(extent.get("cy"))} if extent is not None else None)
    abstracts = []
    if numbering:
        root = etree.fromstring(numbering)
        for node in root.findall(_W + "abstractNum"):
            levels = [{"ilvl": lvl.get(_W + "ilvl"),
                       "numFmt": lvl.find(_W + "numFmt").get(_W + "val"),
                       "lvlText": lvl.find(_W + "lvlText").get(_W + "val"),
                       "ind": {etree.QName(k).localname: v for k, v in lvl.find(f"{_W}pPr/{_W}ind").attrib.items()}
                       if lvl.find(f"{_W}pPr/{_W}ind") is not None else None,
                       "pStyle": style_names.get(lvl.find(_W + "pStyle").get(_W + "val"))
                       if lvl.find(_W + "pStyle") is not None else None}
                      for lvl in node.findall(_W + "lvl")]
            abstracts.append({"id": node.get(_W + "abstractNumId"),
                              "multiLevelType": node.find(_W + "multiLevelType").get(_W + "val"),
                              "levels": levels})

    observations = {
        "word": "16.106 (macOS), Dutch interface",
        "applied": applied,
        "style_ids_localised": {name: sid for sid, name in style_names.items()},
        "styles": {name: {"type": v["type"], "refs": v["refs"]} for name, v in sorted(measured.items())},
        "default_lists": {"paragraphs": lists, "abstracts": abstracts},
        "pictures": {"no-density 96x48 px": extents[0], "300 dpi 300x150 px": extents[1]},
    }
    OBSERVATIONS.write_text(json.dumps(observations, indent=1, ensure_ascii=False, sort_keys=True) + "\n")
    table = {name: v for name, v in sorted(measured.items()) if name not in ours}
    TABLE.write_text(
        '"""Word\'s built-in style definitions, as Word 16.106 writes them on a style\'s first use.\n\n'
        "Generated by ``tools/e1_probe.py`` from what Word wrote into a probe document (the\n"
        "observations are in ``tests/observations/e1-word.json``); do not edit by hand.  Keyed by\n"
        "``w:name``; each entry's ``xml`` has ``{id}``, ``{basedOn}``, ``{next}`` and ``{link}`` where\n"
        "the document's own ids go, and ``refs`` names the styles those point at.\n"
        '"""\n\n'
        f"STYLES: dict[str, dict] = {pprint.pformat(table, width=110)}\n")
    print(f"{len(table)} definitions; applied {applied}; lists {lists}; pictures {extents}")


if __name__ == "__main__":
    main()
