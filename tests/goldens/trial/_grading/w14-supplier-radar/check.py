"""W14: a radar chart from a CSV inserted after a paragraph, with a legend and a caption."""
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm, xlsx_grid  # noqa: E402

from docx_agent import Document  # noqa: E402

INPUT = INPUTS / "w14-supplier-radar/input/supplier-assessment.docx"
CSV = INPUTS / "w14-supplier-radar/input/supplier-scores.csv"
ANNOUNCE = "The panel scored each supplier from 1 to 5 on five criteria"
NEXT = "Brightfold scores highest on quality"
CAPTION = "Figure 1: Supplier scores by criterion, 1 to 5"


def body(c, out: Path):
    rows = [line.split(",") for line in CSV.read_text(encoding="utf-8").split("\n") if line]
    criteria = [r[0] for r in rows[1:]]
    want = {name: [float(r[k]) for r in rows[1:]] for k, name in enumerate(rows[0]) if k}
    doc = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not doc.validate(), doc.validate())
    charts = doc.charts()
    c.check("one chart", len(charts) == 1, len(charts))
    if not charts:
        return
    ch = charts[0]
    c.check("a radar chart", ch.chart_types == ["radar"], ch.chart_types)
    c.check("categories are the criteria in order", list(ch.categories) == criteria, ch.categories)
    got = {s.name: [float(v) for v in s.values] for s in ch.series}
    c.check("one series per supplier, scores exact", got == want, got)
    c.check("a legend", ch.has_legend, ch.has_legend)
    book = ch.workbook_values()
    c.check("Edit Data holds the drawn values",
            book["categories"]["values"] == criteria
            and {s["name"]["value"]: s["values"]["values"] for s in book["series"]} == want, book)
    with zipfile.ZipFile(out) as z:
        grid = xlsx_grid(z.read(ch.workbook_part))
    c.check("workbook read independently agrees",
            [grid.get((i, 1)) for i in range(2, 2 + len(criteria))] == criteria, grid)
    texts = [norm(p.text) for p in doc.paragraphs()]
    where = next((i for i, t in enumerate(texts) if t.startswith(ANNOUNCE)), None)
    c.check("announcing paragraph kept", where is not None)
    if where is None:
        return
    paragraphs = doc.paragraphs()
    holder = paragraphs[where + 1]
    c.check("chart directly after the announcing paragraph",
            holder._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart")
            is not None and not texts[where + 1], texts[where + 1][:60])
    c.check("caption directly below the chart", texts[where + 2] == CAPTION, texts[where + 2])
    caption = paragraphs[where + 2]
    c.check("caption uses the Caption style", (caption.style_name or "").lower() == "caption",
            caption.style_name)
    fields = [f for f in doc.fields() if "SEQ" in f.instruction]
    c.check("caption numbered by a SEQ Figure field",
            len(fields) == 1 and "Figure" in fields[0].instruction, [f.instruction for f in fields])
    c.check("then the paragraph that followed", texts[where + 3].startswith(NEXT), texts[where + 3][:60])
    drawing = doc.drawing(ch.id)
    section = doc.sections()[0]
    room = section.text_width / 20  # twips
    c.check("chart within the text width", drawing.size[0] <= room + 0.5, (drawing.size, room))
    others = [t for t in texts if t and t != CAPTION]
    c.check("nothing else changed", others == [norm(p.text) for p in src.paragraphs()], others)


if __name__ == "__main__":
    main("w14-supplier-radar", body)
