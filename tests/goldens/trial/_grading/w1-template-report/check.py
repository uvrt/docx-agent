"""W1: Markdown draft -> report on the company template, with a TOC."""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

DRAFT = (INPUTS / "w1-template-report/input/draft.md").read_text(encoding="utf-8")


def strip_md(s: str) -> str:
    s = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", s)
    return norm(s.replace("**", "").replace("*", ""))


def parse():
    h0, h1, h2, paras, items, rows = [], [], [], [], [], []
    for block in DRAFT.split("\n\n"):
        block = block.strip()
        if not block:
            continue
        if block.startswith("### "):
            h2.append(strip_md(block[4:]))
        elif block.startswith("## "):
            h1.append(strip_md(block[3:]))
        elif block.startswith("# "):
            h0.append(strip_md(block[2:]))
        elif block.startswith("|"):
            for line in block.splitlines():
                if "---" not in line:
                    rows.append([norm(c) for c in line.strip("|").split("|")])
        elif re.match(r"^(- |\d+\. )", block):
            for line in block.splitlines():
                items.append(strip_md(re.sub(r"^(- |\d+\. )", "", line)))
        else:
            paras.append(strip_md(block))
    return h0, h1, h2, paras, items, rows


def body(c, out: Path):
    doc = Document.open(out)
    h0, h1, h2, paras, items, rows = parse()
    c.check("validate() clean", not doc.validate(), doc.validate())
    ps = doc.paragraphs()
    by_text = {}
    for p in ps:
        by_text.setdefault(norm(p.text), []).append(p)

    def style_of(text):
        found = by_text.get(text)
        return found[0].style_name if found else None

    c.check("title in Report Title", (style_of(h0[0]) or "").lower() == "report title", style_of(h0[0]))
    bad = [(t, style_of(t)) for t in h1 if (style_of(t) or "").lower() != "heading 1"]
    c.check("## headings in Heading 1", not bad, bad)
    bad = [(t, style_of(t)) for t in h2 if (style_of(t) or "").lower() != "heading 2"]
    c.check("### headings in Heading 2", not bad, bad)
    missing = [t for t in paras if t not in by_text]
    c.check("every body paragraph present (exact text)", not missing, missing)
    bad = [(t, style_of(t)) for t in paras if t in by_text and (style_of(t) or "").lower() != "report body"]
    c.check("body paragraphs in Report Body", not bad, bad)
    missing = [t for t in items if t not in by_text]
    c.check("list items present", not missing, missing)
    notlist = [t for t in items if t in by_text and by_text[t][0].list is None]
    c.check("list items are list paragraphs", not notlist, notlist)
    numbered = [by_text[t][0].list.format for t in items[3:] if t in by_text and by_text[t][0].list]
    c.check("incident list numbered", numbered and all(f == "decimal" for f in numbered), numbered)
    # table
    tables = doc.tables()
    ok, detail = False, f"{len(tables)} tables"
    for t in tables:
        got = [[norm(t.cell(r, k).text) for k in range(t.column_count)] for r in range(len(t.rows))]
        if got == rows:
            ok = True
        else:
            detail = got
    c.check("results table matches draft", ok, detail)
    links = [h.target if hasattr(h, "target") else str(h) for h in doc.hyperlinks()]
    c.check("hyperlink kept", any("example.org/dw-guidance" in str(l) for l in links), links)
    guidance = [p.text for p in ps if "Template guidance" in p.text]
    c.check("template guidance text removed", not guidance, guidance)
    # header / footer
    stories = [s.name for s in doc.stories] if hasattr(doc.stories[0], "name") else [str(s) for s in doc.stories]
    head_text = " ".join(p.text for s in doc.stories if "header" in str(s) for p in doc.paragraphs(str(s).split()[1]))
    c.check("template header kept", "Northwind Water Authority | Internal" in head_text, head_text or stories)
    fields = doc.fields()
    page = [f for f in fields if f.keyword == "PAGE" and str(f.paragraph_id).startswith("footer")]
    c.check("footer PAGE field kept", page, [(f.keyword, f.paragraph_id) for f in fields][:10])
    # TOC
    tocs = [f for f in fields if f.keyword == "TOC"]
    c.check("one TOC field", len(tocs) == 1, len(tocs))
    if tocs:
        toc = tocs[0]
        res = toc.result or ""
        lines = [norm(l) for l in res.splitlines() if l.strip()]
        want = [t for t in h1 + h2]
        missing = [t for t in want if not any(l.startswith(t) for l in lines)]
        c.check("TOC lists every Heading 1/2", not missing, {"missing": missing, "toc": lines[:20]})
        nums = [re.search(r"(\d+)$", l) for l in lines]
        c.check("TOC entries carry page numbers", lines and all(nums), lines[:20])
        order = [p.id for p in ps]
        ids = [p.id for p in ps]
        try:
            toc_at = ids.index(toc.paragraph_id)
            prep = ids.index(by_text[paras[0]][0].id)
            first_h = ids.index(by_text[h1[0]][0].id)
            c.check("TOC after 'Prepared by' and before first heading", prep < toc_at < first_h,
                    (prep, toc_at, first_h))
        except Exception as e:  # noqa: BLE001
            c.check("TOC after 'Prepared by' and before first heading", False, e)
    props = doc.properties
    c.check("title property", props.get("title") == "Annual Water Quality Report 2025", props.get("title"))
    c.check("author property", props.get("author") == "Northwind Water Authority", props.get("author"))
    for name in ("Report Body", "Report Title"):
        c.check(f"style {name} defined", doc.styles.find(name) is not None)
    order_expected = h0 + h1
    pos = [ps.index(by_text[t][0]) if t in by_text else -1 for t in order_expected]
    c.check("headings in draft order", pos == sorted(pos) and -1 not in pos, pos)
    empties = sum(1 for p in ps if not p.text.strip() and p.list is None and not doc.to_markdown(p.id).count("d:"))
    c.check("no more than 2 empty paragraphs", empties <= 2, empties)


if __name__ == "__main__":
    main("w1-template-report", body)
