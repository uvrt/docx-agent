"""W3: table + chart (cache and embedded workbook) + text updated from a CSV."""
import io
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

INPUT = INPUTS / "w3-figures-update/input/regional-sales-update.docx"
NEW = {"North": [5.2, 5.4, 5.5, 5.9], "South": [3.9, 3.8, 4.3, 4.6], "West": [4.0, 4.5, 4.7, 5.3]}
TABLE = [
    ["Region", "Q1", "Q2", "Q3", "Q4", "Total"],
    ["North", "5.2", "5.4", "5.5", "5.9", "22.0"],
    ["South", "3.9", "3.8", "4.3", "4.6", "16.6"],
    ["West", "4.0", "4.5", "4.7", "5.3", "18.5"],
    ["Total", "13.1", "13.7", "14.5", "15.8", "57.1"],
]


def body(c, out: Path):
    doc = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not doc.validate(), doc.validate())
    c.check("no tracked changes added", not doc.revisions(), len(doc.revisions()))
    t = doc.tables()[0]
    got = [[norm(t.cell(r, k).text) for k in range(t.column_count)] for r in range(len(t.rows))]
    c.check("table matches answer key", got == TABLE, got)
    c.check("table still 5 rows", len(t.rows) == 5, len(t.rows))
    text = norm(doc.paragraphs()[2].text) if len(doc.paragraphs()) > 2 else ""
    full = " ".join(norm(p.text) for p in doc.paragraphs())
    c.check("text: year-to-date EUR 57.1 million", "EUR 57.1 million" in full and "41.6" not in full)
    c.check("text: North largest at EUR 22.0 million", "North remains the largest region at EUR 22.0 million" in full)
    c.check("text: Q4 strongest at EUR 15.8 million",
            "Q4 was the strongest quarter, at EUR 15.8 million" in full, full[:400])
    charts = doc.charts()
    c.check("one chart", len(charts) == 1, len(charts))
    ch = charts[0]
    c.check("chart categories Q1-Q4", list(ch.categories) == ["Q1", "Q2", "Q3", "Q4"], ch.categories)
    vals = {s.name: [round(float(v), 4) if v is not None else None for v in s.values] for s in ch.series}
    c.check("chart cache values match CSV", vals == NEW, vals)
    c.check("chart title kept", ch.title == "Sales by region (EUR million)", ch.title)
    # embedded workbook: read it straight from the package (grader only)
    try:
        from common import xlsx_grid
        with zipfile.ZipFile(out) as z:
            g = xlsx_grid(z.read(ch.workbook_part))
        names = [g.get((1, j)) for j in (2, 3, 4)]
        cats = [g.get((i, 1)) for i in (2, 3, 4, 5)]
        vals_wb = {n: [round(float(g.get((i, j))), 4) if g.get((i, j)) is not None else None
                       for i in (2, 3, 4, 5)] for j, n in zip((2, 3, 4), names)}
        c.check("embedded workbook matches CSV", cats == ["Q1", "Q2", "Q3", "Q4"] and vals_wb == NEW,
                {"cats": cats, "vals": vals_wb})
    except Exception as e:  # noqa: BLE001
        c.check("embedded workbook matches CSV", False, f"{type(e).__name__}: {e}")
    # nothing else changed
    keep = ["To: Management team. From: Sales Operations.", "All figures in EUR million.",
            "The chart below shows each region's sales by quarter.",
            "Regional managers will present their plans for next year at the November meeting."]
    c.check("other paragraphs unchanged", all(k in full for k in keep))


if __name__ == "__main__":
    main("w3-figures-update", body)
