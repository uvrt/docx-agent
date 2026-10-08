#!/usr/bin/env python3
"""Measure what Word does with charts and SmartArt that docx-agent wrote or edited.

Probe documents are written here (``tools/chart_parts.py``) or edited from the committed
fixtures (``tests/fixtures/generated/charts``) through docx-agent's own chart and diagram
API, and Word -- through ``tests/oracle.py``'s machine-wide lock, no clipboard -- exports
each to PDF and saves it again.  What it drew is read off the PDF (text, and the face,
size, weight and colour of a title's characters), and what it wrote off the saved file:

* ``cache`` -- a chart whose caches disagree with its embedded workbook (the cache says
  ``Q1``... ``Cached``, 10 to 40; the workbook ``W1``... ``Booked``, 900s), with no
  ``c:autoUpdate``, with ``0`` and with ``1``: which Word draws, and whether a save
  refreshes the cache from the workbook;
* ``extensions`` -- ``charts.docx`` with every chart edited (a value, a category and a
  series added, a title): the ``c14``, ``c16r2``, ``c16r3`` and ``c16:uniqueId``
  extensions after docx-agent's edit and after Word's save, and the workbooks;
* ``titles`` -- one chart per title form: a paragraph with no ``a:defRPr``, an empty one
  (ooxml-edit's own template), Word's chart-style form (docx-agent's template), a title
  with no text (``c:title`` alone, what a title just added carries), and axis titles in
  ooxml-edit's and docx-agent's forms: how Word draws each, and what it writes back;
* ``tracking`` -- ``charts.docx`` and ``smartart.docx`` with ``w:trackRevisions`` on and
  docx-agent's edits made under ``doc.tracking()``: the revisions Word's own review
  counts, and what its PDF shows;
* ``places`` -- ``chart-places.docx`` with the chart in the header, the footnote, the
  text box, the group and the floating one each edited: Word's PDF, and the ids Word keeps
  (``docPr``, the group member's ``wpg:cNvPr``);
* ``smartart`` -- pptx-agent E4's variants on ``smartart.docx``: the data model's text
  changed with the drawing stale, the drawing's text alone changed, the drawing removed,
  a node added (the drawing dropped, or kept stale), a node removed (likewise), and a
  text edit the drawing follows exactly: what Word's PDF shows, and whether a save
  rewrites the drawing;
* ``dictionary`` -- Word's AppleScript dictionary (``sdef``) searched for chart data,
  Edit Data or SmartArt objects (no Word needed).

What it found is written to ``tests/observations/charts-word.json``: facts read off Word's
PDFs and files, never the files themselves.

    python tools/charts_probe.py              # every probe, the observations written
    python tools/charts_probe.py titles       # one probe, printed

Nothing Word writes is committed: only these facts.
"""

from __future__ import annotations

import io
import json
import re
import subprocess
import sys
import warnings
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "src"))

import oracle  # noqa: E402
import chart_parts as cp  # noqa: E402
import make_chart_fixtures as fixtures  # noqa: E402

from docx_agent import Document  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "charts-word.json"
FIXTURES = ROOT / "tests" / "fixtures" / "generated" / "charts"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
DGM = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
DSP = "{http://schemas.microsoft.com/office/drawing/2008/diagram}"
WPG = "{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def build(builder_body, *, extra=None) -> bytes:
    """A document from ``Document.new()``'s parts: ``builder_body(builder)`` returns the body."""
    base = fixtures.base_parts()
    builder = cp.Builder(base)
    body = builder_body(builder)
    fixtures._body(base, body)
    if extra:
        extra(base)
    return cp.zip_parts(builder.finish())


def charts_of(data: bytes) -> list[dict]:
    """Every chart in a saved document, in drawing order: its id, caches and extensions."""
    document = Document.open(data)
    out = []
    for chart in document.charts():
        root = document.package.tree(chart.part)
        names = [entry["name"] for entry in chart.data["series"]]
        series = []
        for name, ser in zip(names, root.iter(C + "ser")):
            series.append({
                "name": name,
                "unique_id": next((n.get("val") for n in ser.iter() if isinstance(n.tag, str)
                                   and n.tag.endswith("}uniqueId")), None),
                "ext_uris": [e.get("uri") for e in ser.findall(f"{C}extLst/{C}ext")],
            })
        workbook = chart.workbook_part
        out.append({
            "id": chart.id, "part": chart.part, "data": chart.data, "series": series,
            "c14_style": root.find(".//{http://schemas.microsoft.com/office/drawing/2007/8/2/chart}style") is not None,
            "c16r2_declared": "http://schemas.microsoft.com/office/drawing/2015/06/chart" in root.nsmap.values(),
            "c16r3": root.find(".//{http://schemas.microsoft.com/office/drawing/2017/03/chart}dataDisplayOptions16")
            is not None,
            "lang": (root.find(C + "lang").get("val") if root.find(C + "lang") is not None else None),
            "auto_update": (root.find(f"{C}externalData/{C}autoUpdate").get("val")
                            if root.find(f"{C}externalData/{C}autoUpdate") is not None else None),
            "workbook": workbook,
            "workbook_sha": __import__("hashlib").sha256(document.package.read(workbook)).hexdigest()[:16]
            if workbook else None,
            "formulas": [f.text for f in root.iter(C + "f")],
        })
    return out


def pdf_text(path: Path) -> str:
    return "\n".join(oracle.pdf_pages(path))


def chars_of(path: Path, text: str) -> dict | None:
    """How Word drew ``text`` (its first occurrence): face, size, weight, colour."""
    found = oracle.find_in_pdf(oracle.pdf_chars(path), text)
    if not found:
        return None
    first = found[0]
    return {"font": first.font.split("+", 1)[-1], "size": first.size, "weight": first.weight,
            "color": "%02X%02X%02X" % first.color, "stroked": first.stroked}


# -- cache against workbook ------------------------------------------------------------------------


def cache_document() -> bytes:
    def body(builder: cp.Builder) -> str:
        out = [fixtures.p("Cache against workbook.")]
        for k, update in enumerate((None, "0", "1"), start=1):
            spec = cp.ChartSpec("column", cp.ChartData(["Q1", "Q2", "Q3", "Q4"], {f"Cached{k}": [10, 20, 30, 40]}),
                                auto_update=update, seed=20 + k,
                                book=cp.ChartData(["W1", "W2", "W3", "W4"], {f"Booked{k}": [900, 910, 920, 930]}))
            rel = builder.chart("word/document.xml", spec)
            out.append(fixtures.p(f"autoUpdate {update}."))
            out.append(fixtures.p(inner=cp.inline(k, f"Chart {k}", cp.chart_graphic(rel), fixtures.SMALL_CX * 2,
                                                  fixtures.SMALL_CY)))
        return "".join(out)
    return build(body)


def probe_cache(word: oracle.Session) -> dict:
    data = cache_document()
    pdf = word.export_pdf(data, name="charts-cache")
    saved = word.resave(data, name="charts-cache")
    text = pdf_text(pdf.path) if pdf else ""
    out = {"exported": bool(pdf), "saved": bool(saved)}
    for k in (1, 2, 3):
        out[f"chart{k}"] = {
            "pdf_shows_cache_name": f"Cached{k}" in text, "pdf_shows_workbook_name": f"Booked{k}" in text,
        }
    out["pdf_shows_cache_categories"] = all(q in text for q in ("Q1", "Q4"))
    out["pdf_shows_workbook_categories"] = any(w in text for w in ("W1", "W4"))
    out["pdf_value_labels"] = sorted(set(re.findall(r"\b\d{2,4}\b", text)) - {"0"})
    if saved:
        out["saved_charts"] = [{"data": c["data"], "auto_update": c["auto_update"], "workbook_sha": c["workbook_sha"]}
                               for c in charts_of(saved.path.read_bytes())]
        out["input_workbooks"] = [c["workbook_sha"] for c in charts_of(data)]
    return out


# -- extensions ------------------------------------------------------------------------------------


def edited_charts(tracking: bool = False) -> bytes:
    document = Document.open(FIXTURES / "charts.docx")
    if tracking:
        document.set_word_tracks_changes(True)

    def edits() -> None:
        for chart in document.charts():
            numeric = chart.chart_types == ["scatter"]
            chart.series[0].set_value(0, 777)
            chart.add_category(9.5 if numeric else "Added", None)
            chart.add_series("New series", [5] * chart.point_count)
            chart.set_category(1, 2.2 if numeric else "Renamed")
            chart.series[0].name = chart.series[0].name + " edited"
            chart.set_title(f"Edited {chart.id}")

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        if tracking:
            with document.tracking(author="Probe"):
                edits()
        else:
            edits()
    document._probe_warnings = [str(w.message) for w in caught]
    return document.to_bytes()


def probe_extensions(word: oracle.Session) -> dict:
    data = edited_charts()
    pdf = word.export_pdf(data, name="charts-extensions")
    saved = word.resave(data, name="charts-extensions")
    out = {"exported": bool(pdf), "ours": charts_of(data)}
    if saved:
        out["saved"] = charts_of(saved.path.read_bytes())
    if pdf:
        text = pdf_text(pdf.path)
        out["pdf_has"] = {label: label in text for label in ("Added", "Renamed", "New series", "777", "Edited d:1")}
    return out


# -- titles ----------------------------------------------------------------------------------------


def _rich(text: str, ppr: str, *, vertical: bool = False) -> str:
    rot = ' rot="-5400000" vert="horz"' if vertical else ""
    return (f"<c:tx><c:rich><a:bodyPr{rot}/><a:lstStyle/><a:p>{ppr}<a:r><a:rPr lang=\"en-US\"/><a:t>{text}</a:t>"
            "</a:r></a:p></c:rich></c:tx>")


def title_cases() -> dict[str, str | None]:
    from docx_agent.edit.charts import title_text
    from ooxml_edit.charts.dmltext import replace_body_text

    holder = etree.Element(C + "chartSpace", nsmap={"c": C[1:-1], "a": A[1:-1]})
    holder.append(title_text(False, lang="en-US"))
    replace_body_text(holder[0][0], "Word style")  # as set_title writes it
    word_form = re.sub(r' xmlns:\w+="[^"]+"', "", etree.tostring(holder[0]).decode())
    return {
        "none": _rich("No defaults", ""),
        "empty": _rich("Empty defaults", "<a:pPr><a:defRPr/></a:pPr>"),
        "word": word_form,
        "auto": "",
    }


def titles_document() -> bytes:
    cases = title_cases()

    def body(builder: cp.Builder) -> str:
        out = [fixtures.p("Titles.")]
        for k, (key, tx) in enumerate(cases.items(), start=1):
            spec = cp.ChartSpec("column", cp.ChartData(["A", "B"], {f"Series {key}": [1, 2]}), title_xml=tx,
                                legend=None, seed=30 + k)
            rel = builder.chart("word/document.xml", spec)
            out.append(fixtures.p(f"Title {key}."))
            out.append(fixtures.p(inner=cp.inline(k, f"Chart {k}", cp.chart_graphic(rel), fixtures.SMALL_CX * 2,
                                                  fixtures.SMALL_CY)))
        return "".join(out)
    return build(body)


def axis_titles_document() -> bytes:
    """Two charts whose value axis gets a title through the API: ooxml-edit's template and
    docx-agent's."""
    document = Document.open(build(lambda builder: "".join(
        [fixtures.p("Axis titles.")] + [
            fixtures.p(inner=cp.inline(k, f"Chart {k}", cp.chart_graphic(builder.chart(
                "word/document.xml", cp.ChartSpec("column", cp.ChartData(["A", "B"], {f"S{k}": [1, 2]}),
                                                  legend=None, seed=40 + k))), fixtures.SMALL_CX * 2,
                fixtures.SMALL_CY)) for k in (1, 2)])))
    import dataclasses

    from ooxml_edit.charts import chart as raw

    first, second = document.charts()
    # ooxml-edit's own template: the host without docx-agent's.
    plain = raw.Chart(lambda: dataclasses.replace(first._resolve(), title_template=None))
    plain.set_axis_title("value", "Plain axis")
    plain.set_title("Plain title")
    second.set_axis_title("value", "Styled axis")
    second.set_title("Styled title")
    return document.to_bytes()


def _saved_titles(data: bytes) -> list[dict]:
    def serialize(node):
        return etree.tostring(node)

    document = Document.open(data)
    out = []
    for chart in document.charts():
        root = document.package.tree(chart.part)
        title = root.find(f"{C}chart/{C}title")
        axis = root.find(f".//{C}valAx/{C}title")
        out.append({"id": chart.id, "title": chart.title,
                    "title_xml": None if title is None else re.sub(r' xmlns:\w+="[^"]+"', "",
                                                                    serialize(title).decode()),
                    "axis_xml": None if axis is None else re.sub(r' xmlns:\w+="[^"]+"', "", serialize(axis).decode())})
    return out


def docx2svg_title(data: bytes, text: str) -> dict | None:
    """How docx2svg draws ``text``: the ``<text>`` element's face, size and weight."""
    import docx2svg

    for svg in docx2svg.convert_docx_to_svg(data, docx2svg.ConvertOptions()):
        root = etree.fromstring(svg.encode())
        for node in list(root.iter("{http://www.w3.org/2000/svg}tspan")) + list(root.iter("{http://www.w3.org/2000/svg}text")):
            if " ".join("".join(node.itertext()).split()) == text:
                found = {}
                current = node
                while current is not None:
                    for key in ("font-family", "font-size", "font-weight", "fill"):
                        if key not in found and current.get(key) is not None:
                            found[key] = current.get(key)
                    current = current.getparent()
                return found
    return None


def probe_titles(word: oracle.Session) -> dict:
    out = {}
    for name, data, texts in (("titles", titles_document(), ["No defaults", "Empty defaults", "Word style",
                                                                "Series auto", "Chart Title", "Grafiektitel"]),
                              ("axis", axis_titles_document(), ["Plain axis", "Plain title", "Styled axis",
                                                                "Styled title"])):
        pdf = word.export_pdf(data, name=f"charts-{name}")
        saved = word.resave(data, name=f"charts-{name}")
        entry = {"exported": bool(pdf), "ours": _saved_titles(data)}
        if pdf:
            entry["word_draws"] = {t: chars_of(pdf.path, t) for t in texts}
            entry["pdf_text"] = [line for line in pdf_text(pdf.path).splitlines() if line.strip()][:40]
        entry["docx2svg_draws"] = {t: docx2svg_title(data, t) for t in texts}
        if saved:
            entry["saved"] = _saved_titles(saved.path.read_bytes())
        out[name] = entry
    return out


# -- tracking --------------------------------------------------------------------------------------

_COUNT_SCRIPT = '''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  with timeout of 170 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file inputPath)
      set d to active document
      set n to count of revisions of d
      set t to track revisions of d
      save as d file name outputPath file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
  return "revisions=" & n & " tracking=" & t
end run
'''


def edited_smartart(tracking: bool = False, policy: str | None = None) -> bytes:
    document = Document.open(FIXTURES / "smartart.docx")
    if tracking:
        document.set_word_tracks_changes(True)

    def edits() -> None:
        first, second = document.diagrams()
        first.set_text(1, "Drafted")
        second.set_text(second.node(1), "Much faster edits")
        second.node(0).add_child("Smaller files")

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        if tracking:
            with document.tracking(author="Probe"):
                edits()
        else:
            edits()
    return document.to_bytes()


def probe_tracking(word: oracle.Session) -> dict:
    out = {}
    for name, data in (("charts", edited_charts(tracking=True)), ("smartart", edited_smartart(tracking=True))):
        # The count is the script's answer, which the cache does not keep: always ask Word.
        for stale in oracle.ORACLE_DIR.glob(f"charts-tracking-{name}-*-counted.docx"):
            stale.unlink()
        counted = word.run_script(_COUNT_SCRIPT, data, name=f"charts-tracking-{name}", tag="counted")
        pdf = word.export_pdf(data, name=f"charts-tracking-{name}")
        entry = {"word_says": counted.detail if counted else f"failed: {counted.detail}",
                 "ours_has_markup": oracle._zip_has_markup(data)}
        if pdf:
            text = pdf_text(pdf.path)
            entry["pdf_has"] = {label: label in text for label in ("Added", "Renamed", "New series", "777", "Drafted",
                                                                   "Much faster edits", "Smaller files")}
        if counted:
            entry["saved_has_markup"] = oracle._zip_has_markup(counted.path.read_bytes())
        out[name] = entry
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        document = Document.open(FIXTURES / "charts.docx")
        with document.tracking(author="Probe"):
            document.charts()[0].series[0].set_value(0, 1)
    out["docx_agent_warns"] = [f"{type(w.message).__name__}: {w.message}" for w in caught]
    return out


# -- places ----------------------------------------------------------------------------------------


def edited_places() -> bytes:
    document = Document.open(FIXTURES / "chart-places.docx")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        for chart in document.charts():
            tag = chart.id.replace("d:", "").replace("/", "m")
            chart.series[0].set_value(1, 4321)
            chart.set_category(0, f"Cat{tag}")
            chart.series[0].name = f"Name{tag}"
    return document.to_bytes()


def _ids(data: bytes) -> dict:
    document = Document.open(data)
    root = document.package.tree(document.package.document_part())
    return {"docPr": sorted(int(n.get("id")) for n in root.iter(WP + "docPr")),
            "group_members": [(n.get("id"), n.get("name")) for n in root.iter(WPG + "cNvPr")],
            "charts": [c.id for c in document.charts()]}


def probe_places(word: oracle.Session) -> dict:
    data = edited_places()
    pdf = word.export_pdf(data, name="charts-places")
    saved = word.resave(data, name="charts-places")
    out = {"exported": bool(pdf), "ours": _ids(data)}
    if pdf:
        text = pdf_text(pdf.path)
        out["pdf_has"] = {f"{kind}{tag}": f"{kind}{tag}" in text for tag in ("1", "2", "4", "5m7", "8")
                          for kind in ("Cat", "Name")}
        out["pdf_has_4321"] = "4321" in text or "4.321" in text or "4 321" in text or "4500" in text
        out["pdf_numbers"] = sorted(set(re.findall(r"\b\d[\d.]{2,5}\b", text)))
    if saved:
        saved_data = saved.path.read_bytes()
        out["saved"] = _ids(saved_data)
        out["saved_charts"] = [{"id": c["id"], "data": c["data"]} for c in charts_of(saved_data)]
        names = parts(saved_data)
        out["saved_media"] = sorted(n for n in names if n.startswith("word/media/"))
        out["input_media"] = sorted(n for n in parts(data) if n.startswith("word/media/"))
        out["fallback_image_changed"] = {n: parts(data).get(n) != names.get(n) for n in out["saved_media"]}
    return out


# -- SmartArt --------------------------------------------------------------------------------------


def _smartart_variant(name: str) -> bytes:
    document = Document.open(FIXTURES / "smartart.docx")
    package = document.package
    second = document.diagrams()[1]
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        if name == "model-stale":
            # The data model's text changed; the drawing left as it was.
            root = package.tree(second.part)
            point = next(p for p in root.iter(DGM + "pt") if "".join(t.text or "" for t in p.iter(A + "t")) == "Goals")
            next(point.iter(A + "t")).text = "Objectives"
            package.mark_dirty(second.part)
        elif name == "drawing-only":
            part = second.drawing_part
            root = package.tree(part)
            node = next(t for t in root.iter(A + "t") if t.text == "Goals")
            node.text = "Drawn only"
            package.mark_dirty(part)
        elif name == "drawing-missing":
            document.diagram(second.id, on_inexact_drawing="drop")._drop_drawing(
                __import__("ooxml_edit.charts.diagram", fromlist=["_Model"])._Model(
                    package.tree(second.part), second._resolve()))
            package.mark_dirty(second.part)
        elif name == "added-dropped":
            document.diagram(second.id, on_inexact_drawing="drop").node(0).add_child("Added child")
        elif name == "added-stale":
            document.diagram(second.id, on_inexact_drawing="keep").node(0).add_child("Added child")
        elif name == "removed-dropped":
            diagram = document.diagram(second.id, on_inexact_drawing="drop")
            diagram.remove_node(diagram.node(3))  # "Risks" and its children
        elif name == "removed-stale":
            diagram = document.diagram(second.id, on_inexact_drawing="keep")
            diagram.remove_node(diagram.node(3))
        elif name == "text-exact":
            document.diagram(second.id).set_text(1, "Exactly edited")
        elif name == "top-added-dropped":
            first = document.diagrams()[0]
            document.diagram(first.id, on_inexact_drawing="drop").add_node("Sixth")
        else:
            raise KeyError(name)
    return document.to_bytes()


SMARTART_VARIANTS = ("model-stale", "drawing-only", "drawing-missing", "added-dropped", "added-stale",
                     "removed-dropped", "removed-stale", "text-exact", "top-added-dropped")
SMARTART_WORDS = ("Goals", "Objectives", "Drawn only", "Added child", "Risks", "Stale caches", "Exactly edited",
                  "Faster edits", "Sixth")


def _drawing_texts(data: bytes) -> list:
    document = Document.open(data)
    out = []
    for diagram in document.diagrams():
        part = diagram.drawing_part
        texts = None
        if part is not None:
            root = document.package.tree(part)
            texts = ["".join(t.text or "" for t in p.iter(A + "t")) for p in root.iter(A + "p")]
            texts = [t for t in texts if t]
        presentations = sum(1 for p in document.package.tree(diagram.part).iter(DGM + "pt")
                            if p.get("type") == "pres")
        out.append({"id": diagram.id, "texts": diagram.texts, "drawing": part, "drawing_texts": texts,
                    "presentation_points": presentations})
    return out


def probe_smartart(word: oracle.Session) -> dict:
    out = {}
    for variant in SMARTART_VARIANTS:
        data = _smartart_variant(variant)
        pdf = word.export_pdf(data, name=f"smartart-{variant}")
        saved = word.resave(data, name=f"smartart-{variant}")
        entry = {"exported": bool(pdf), "ours": _drawing_texts(data)}
        if pdf:
            text = pdf_text(pdf.path)
            entry["pdf_has"] = {w: w in text for w in SMARTART_WORDS}
        if saved:
            entry["saved"] = _drawing_texts(saved.path.read_bytes())
        out[variant] = entry
    return out


# -- the dictionary ----------------------------------------------------------------------------------


def probe_dictionary(_word=None) -> dict:
    sdef = subprocess.run(["sdef", str(oracle.WORD_APP)], capture_output=True, text=True).stdout
    classes = re.findall(r'<class name="([^"]+)"', sdef)
    commands = re.findall(r'<command name="([^"]+)"', sdef)
    words = re.compile(r"chart|smart ?art|workbook|excel|edit data", re.I)
    return {
        "classes": sorted(c for c in classes if words.search(c)),
        "commands": sorted(c for c in commands if words.search(c)),
        "properties": sorted(set(re.findall(r'<property name="([^"]*(?:chart|smart)[^"]*)"', sdef, re.I))),
        "enumerators": sorted(set(re.findall(r'<enumerator name="((?:inline shape|shape type|link type)[^"]*'
                                             r'(?:chart|smart)[^"]*)"', sdef, re.I))),
        "has_run_vb_macro": "run VB macro" in commands,
    }


PROBES = {"cache": probe_cache, "extensions": probe_extensions, "titles": probe_titles, "tracking": probe_tracking,
          "places": probe_places, "smartart": probe_smartart, "dictionary": probe_dictionary}


def main(argv: list[str]) -> int:
    chosen = argv or list(PROBES)
    results = {}
    if chosen == ["dictionary"]:
        print(json.dumps(probe_dictionary(), indent=1))
        return 0
    with oracle.session() as word:
        for name in chosen:
            results[name] = PROBES[name](word)
            print(f"{name}: done", file=sys.stderr)
    if argv:
        print(json.dumps(results, indent=1, ensure_ascii=False, default=str))
        return 0
    OBSERVATIONS.write_text(json.dumps({"word": "16.106 for Mac", **results}, indent=1, ensure_ascii=False,
                                       default=str) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
