#!/usr/bin/env python3
"""Write the chart and SmartArt documents, ``tests/fixtures/generated/charts/``.

Each is built here -- ``Document.new()``'s parts with charts, workbooks and SmartArt data
models added by ``tools/chart_parts.py`` -- and then **opened and saved by Word** (the
machine-wide lock, ``tests/oracle.py``'s ``resave``), so what is committed is what Word
wrote for it: pptx-agent's SmartArt-fixture method.  Word's author and dates in
``docProps/core.xml`` are then normalised (``normalise_properties``); nothing else Word
wrote is changed.

``charts.docx``
    six inline charts in the body, each with an embedded workbook: clustered columns
    (titled), a line, a pie, a doughnut of two rings, a scatter, and a combination of
    columns and a line over the same categories.
``chart-places.docx``
    a chart in each place a Word chart can stand that is not the body's line: the default
    header, a floating text box, a group beside a rectangle (``wpg:graphicFrame``) and a
    footnote; and one floating in the body.
``smartart.docx``
    Basic Block List and Vertical Bullet List (two items with two bulleted items each),
    written as data models alone -- Word lays them out and writes the presentation points,
    the drawing and the definitions it uses.

They sit two directories below the corpus, so no other phase's suite picks them up; the
chart tests (``tests/test_charts.py``) hold them to every gate.

    python tools/make_chart_fixtures.py            # inputs to scratch, then Word saves them
    python tools/make_chart_fixtures.py --inputs   # the inputs only, to the given directory
"""

from __future__ import annotations

import argparse
import hashlib
import io
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "src"))

import chart_parts as cp  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "generated" / "charts"
CREATED = datetime(2026, 10, 4, tzinfo=timezone.utc)
#: What ``docProps/core.xml`` says after normalising.
AUTHOR = "docx-agent"
DATE = "2026-10-04T12:00:00Z"

CHART_CX, CHART_CY = 5486400, 3017520
SMALL_CX, SMALL_CY = 2743200, 1600200


def base_parts() -> dict[str, bytes]:
    from docx_agent import Document

    return cp.read_parts(Document.new(created=CREATED).to_bytes())


def p(text: str = "", inner: str = "", style: str | None = None) -> str:
    props = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{props}{run}{inner}</w:p>"


def _body(parts: dict[str, bytes], body: str, section_extra: str = "") -> None:
    document = parts["word/document.xml"].decode()
    start = document.index("<w:body>") + len("<w:body>")
    sect = document.index("<w:sectPr>")
    document = document[:start] + body + document[sect:]
    if section_extra:
        document = document.replace("<w:sectPr>", "<w:sectPr>" + section_extra, 1)
    parts["word/document.xml"] = document.encode()


# -- the documents -----------------------------------------------------------------------------


def chart_specs() -> list[cp.ChartSpec]:
    quarters = ["Q1", "Q2", "Q3", "Q4"]
    return [
        cp.ChartSpec("column", cp.ChartData(quarters, {"North": [4.3, 2.5, 3.5, 4.5], "South": [2.4, 4.4, 1.8, 2.8],
                                                       "West": [2, 2, 3, 5]}), title="Sales by region", seed=1),
        cp.ChartSpec("line", cp.ChartData(["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
                                          {"Visitors": [120, 135, 160, 158, 190, 210],
                                           "Signups": [12, 18, 15, 22, 30, 41]}), seed=2),
        cp.ChartSpec("pie", cp.ChartData(["Apples", "Pears", "Plums", "Cherries"], {"Share": [8.2, 3.2, 1.4, 1.2]}),
                     title="Orchard", legend="r", seed=3),
        cp.ChartSpec("doughnut", cp.ChartData(["Design", "Build", "Test"], {"Plan": [30, 50, 20], "Actual": [25, 60, 15]}),
                     legend="r", seed=4),
        cp.ChartSpec("scatter", cp.ChartData([0.7, 1.8, 2.6, 3.1, 4.4], {"Measured": [2.7, 3.2, 0.8, 4.1, 3.6]},
                                             corner="X-Values"), seed=5),
        cp.ChartSpec("combo", cp.ChartData(quarters, {"Revenue": [40, 52, 61, 70], "Cost": [31, 37, 42, 45],
                                                      "Margin": [9, 15, 19, 25]}), seed=6),
    ]


def charts_document() -> bytes:
    parts = base_parts()
    builder = cp.Builder(parts)
    body = [p("Six charts, each with its workbook.")]
    for k, spec in enumerate(chart_specs(), start=1):
        rel = builder.chart("word/document.xml", spec)
        body.append(p(f"Chart {k}: {spec.kind}."))
        body.append(p(inner=cp.inline(k, f"Chart {k}", cp.chart_graphic(rel), CHART_CX, CHART_CY)))
    body.append(p("The end."))
    _body(parts, "".join(body))
    return cp.zip_parts(builder.finish())


def places_specs() -> dict[str, cp.ChartSpec]:
    small = lambda seed, *names: cp.ChartSpec(  # noqa: E731
        "column", cp.ChartData(["A", "B", "C"], {n: [k + 1, k + 3, k + 2] for k, n in enumerate(names)}), seed=seed)
    return {"header": small(11, "Header"), "box": small(12, "In box"), "group": small(13, "Grouped", "Paired"),
            "note": small(14, "Noted"), "float": small(15, "Floating")}


def places_document() -> bytes:
    parts = base_parts()
    builder = cp.Builder(parts)
    specs = places_specs()
    # The header.
    header = "word/header1.xml"
    document = parts["word/document.xml"].decode()
    namespaces = document[document.index("<w:document ") + len("<w:document "):document.index(">", document.index(
        "<w:document "))]
    header_rel = builder.chart(header, specs["header"])
    parts[header] = (cp.DECL + f"<w:hdr {namespaces}>"
                     + p("A chart in the header:") + p(inner=cp.inline(1, "Chart 1", cp.chart_graphic(header_rel),
                                                                        SMALL_CX, SMALL_CY // 2))
                     + "</w:hdr>").encode()
    builder.overrides["/word/header1.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
    header_ref = builder._rel("word/document.xml", cp.REL + "header", "header1.xml")
    # A footnote.
    note_rel = builder.chart("word/footnotes.xml", specs["note"])
    separator = ('<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
                 'w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
                 '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" '
                 'w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')
    note = ('<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> A chart in a note:'
            '</w:t></w:r></w:p>' + p(inner=cp.inline(2, "Chart 2", cp.chart_graphic(note_rel), SMALL_CX, SMALL_CY // 2))
            + "</w:footnote>")
    parts["word/footnotes.xml"] = (cp.DECL + f"<w:footnotes {namespaces}>{separator}{note}</w:footnotes>").encode()
    builder.overrides["/word/footnotes.xml"] = \
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
    builder._rel("word/document.xml", cp.REL + "footnotes", "footnotes.xml")
    settings = parts["word/settings.xml"].decode()
    settings = settings.replace("<w:compat>", '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/>'
                                '</w:footnotePr><w:compat>', 1)
    parts["word/settings.xml"] = settings.encode()
    # The body: a text box holding a chart, a group with one, a floating one.
    box_rel = builder.chart("word/document.xml", specs["box"])
    group_rel = builder.chart("word/document.xml", specs["group"])
    float_rel = builder.chart("word/document.xml", specs["float"])
    box_content = p("A chart in a text box:") + p(inner=cp.inline(4, "Chart 4", cp.chart_graphic(box_rel),
                                                                     SMALL_CX, SMALL_CY))
    body = [
        p("Charts in other places.", inner='<w:r><w:footnoteReference w:id="1"/></w:r>'),
        p("A text box follows.", inner=cp.anchored(3, "Text Box 3", cp.text_box_graphic(box_content, SMALL_CX + 182880,
                                                                                         SMALL_CY + 640080),
                                                   SMALL_CX + 182880, SMALL_CY + 640080)),
        p("A group follows.", inner=cp.anchored(5, "Group 5", cp.group_graphic(group_rel, 7, 6, CHART_CX, SMALL_CY),
                                                CHART_CX, SMALL_CY, height=251660288)),
        p("A floating chart follows.", inner=cp.anchored(8, "Chart 8", cp.chart_graphic(float_rel), SMALL_CX, SMALL_CY,
                                                         height=251661312)),
        p("The end."),
    ]
    _body(parts, "".join(body), f'<w:headerReference w:type="default" r:id="{header_ref}"/>')
    return cp.zip_parts(builder.finish())


def diagram_specs() -> list[cp.DiagramSpec]:
    return [
        cp.DiagramSpec("default", [("Plan", 0), ("Write", 0), ("Measure", 0), ("Review", 0), ("Ship", 0)], seed=1),
        cp.DiagramSpec("vList2", [("Goals", 0), ("Faster edits", 1), ("Fewer prompts", 1), ("Risks", 0),
                                  ("Stale caches", 1), ("Lost formulas", 1)], seed=2),
    ]


def smartart_document() -> bytes:
    parts = base_parts()
    builder = cp.Builder(parts)
    body = [p("Two SmartArt diagrams.")]
    for k, spec in enumerate(diagram_specs(), start=1):
        rels = builder.diagram("word/document.xml", spec)
        body.append(p(f"Diagram {k}: {spec.layout}."))
        body.append(p(inner=cp.inline(k, f"Diagram {k}", cp.diagram_graphic(rels), CHART_CX, 3200400)))
    body.append(p("The end."))
    _body(parts, "".join(body))
    return cp.zip_parts(builder.finish())


DOCUMENTS = {"charts.docx": charts_document, "chart-places.docx": places_document, "smartart.docx": smartart_document}


# -- after Word ----------------------------------------------------------------------------------


def normalise_properties(data: bytes) -> bytes:
    """Word's author and dates in ``docProps/core.xml`` replaced by :data:`AUTHOR` and
    :data:`DATE`; every other entry as Word wrote it (order, compression and bytes)."""
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for info in source.infolist():
            content = source.read(info)
            if info.filename == "docProps/core.xml":
                text = content.decode("utf-8")
                text = re.sub(r"(<dc:creator>)[^<]*(</dc:creator>)", rf"\g<1>{AUTHOR}\g<2>", text)
                text = re.sub(r"(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)", rf"\g<1>{AUTHOR}\g<2>", text)
                text = re.sub(r"(<dcterms:(created|modified)[^>]*>)[^<]*(</dcterms:\2>)", rf"\g<1>{DATE}\g<3>", text)
                text = re.sub(r"<cp:lastPrinted>[^<]*</cp:lastPrinted>", "", text)
                content = text.encode("utf-8")
            clone = zipfile.ZipInfo(info.filename, date_time=(1980, 1, 1, 0, 0, 0))
            clone.compress_type = info.compress_type
            clone.external_attr = info.external_attr
            archive.writestr(clone, content)
    return buffer.getvalue()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--inputs", type=Path, help="write the inputs only, to this directory")
    args = parser.parse_args(argv)
    inputs = {name: build() for name, build in DOCUMENTS.items()}
    if args.inputs:
        args.inputs.mkdir(parents=True, exist_ok=True)
        for name, data in inputs.items():
            (args.inputs / name).write_bytes(data)
            print(f"wrote {args.inputs / name}")
        return 0
    import oracle

    OUT.mkdir(parents=True, exist_ok=True)
    with oracle.session() as word:
        for name, data in inputs.items():
            outcome = word.resave(data, name=f"fixture-{Path(name).stem}", timeout=240)
            if not outcome:
                print(f"{name}: Word did not save it ({outcome.outcome}: {outcome.detail})")
                return 1
            saved = normalise_properties(outcome.path.read_bytes())
            (OUT / name).write_bytes(saved)
            print(f"{name}: in {hashlib.sha256(data).hexdigest()}, saved {hashlib.sha256(saved).hexdigest()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
