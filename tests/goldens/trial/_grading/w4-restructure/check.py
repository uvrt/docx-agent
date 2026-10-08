"""W4: move a section, renumber headings, restart a list, fix cross-references and the TOC."""
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

INPUT = INPUTS / "w4-restructure/input/security-policy.docx"
HEADINGS = [("heading 1", "1 Purpose"), ("heading 1", "2 Scope"), ("heading 1", "3 Access control"),
            ("heading 1", "4 Data retention"), ("heading 2", "4.1 Retention periods"),
            ("heading 2", "4.2 Disposal"), ("heading 1", "5 Acceptable use"),
            ("heading 1", "6 Incident response"), ("heading 1", "7 Review")]


def body(c, out: Path):
    doc = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not doc.validate(), doc.validate())
    c.check("no tracked changes", not doc.revisions(), len(doc.revisions()))
    ps = doc.paragraphs()
    heads = [((p.style_name or "").lower(), norm(p.text)) for p in ps
             if (p.style_name or "").lower().startswith("heading")]
    c.check("headings renumbered, in the new order", heads == HEADINGS, heads)
    texts = [norm(p.text) for p in ps]

    def at(prefix):
        for i, t in enumerate(texts):
            if t.startswith(prefix) and not (ps[i].style_name or "").lower().startswith("toc"):
                return i
        return -1

    order = [at("Exceptions to these rules"), at("4 Data retention"), at("Information is kept only"),
             at("Customer contracts"), at("4.2 Disposal"), at("Paper records are shredded"),
             at("5 Acceptable use")]
    c.check("whole section moved after Access control (with table)", -1 not in order and order == sorted(order),
            order)
    # lists
    def label(prefix):
        i = at(prefix)
        if i < 0:
            return None
        md = doc.to_markdown(ps[i].id, ids=False).strip()
        m = re.match(r"^(\d+)\.", md)
        return int(m.group(1)) if m else md[:30]

    inc = [label("Contain the incident"), label("Assess the impact"), label("Recover affected")]
    acc = [label("Managers request"), label("The system owner"), label("Access is reviewed")]
    c.check("incident steps numbered 1-3", inc == [1, 2, 3], inc)
    c.check("access steps still numbered 1-3", acc == [1, 2, 3], acc)
    # fields
    fields = doc.fields()
    layout = doc.layout()

    def page_of(prefix):
        i = at(prefix)
        if i < 0:
            return None
        where = layout.where(ps[i].id)
        return where[0].page if where else None

    refs = {f.instruction.split()[1]: f for f in fields if f.keyword == "REF"}
    pagerefs = {f.instruction.split()[1]: f for f in fields if f.keyword == "PAGEREF"}
    inc_p = norm(texts[at("Exceptions to these rules")]) if at("Exceptions to these rules") >= 0 else ""
    rev_p = norm(texts[at("This policy is reviewed")]) if at("This policy is reviewed") >= 0 else ""
    c.check("incident cross-reference shows '6 Incident response'",
            "described in section 6 Incident response (page" in inc_p, inc_p)
    c.check("retention cross-reference shows '4 Data retention'",
            "are in section 4 Data retention (page" in rev_p, rev_p)
    ref_fields_ok = sorted(norm(f.result or "") for f in refs.values())
    c.check("REF fields still fields, results updated", ref_fields_ok == ["4 Data retention", "6 Incident response"],
            ref_fields_ok)
    pin = re.search(r"\(page (\d+)\)", inc_p)
    prt = re.search(r"\(page (\d+)\)", rev_p)
    c.check("incident page reference correct", pin and int(pin.group(1)) == page_of("6 Incident response"),
            (pin and pin.group(1), page_of("6 Incident response")))
    c.check("retention page reference correct", prt and int(prt.group(1)) == page_of("4 Data retention"),
            (prt and prt.group(1), page_of("4 Data retention")))
    tocs = [f for f in fields if f.keyword == "TOC"]
    c.check("one TOC", len(tocs) == 1, len(tocs))
    if tocs:
        lines = [norm(l) for l in (tocs[0].result or "").splitlines() if l.strip()]
        names = [re.sub(r"\s*\d+$", "", l) for l in lines]
        c.check("TOC entries match new headings and order", names == [h for _, h in HEADINGS], lines)
        pages_ok = all(re.search(r"(\d+)$", l) and int(re.search(r"(\d+)$", l).group(1)) == page_of(h)
                       for l, (_, h) in zip(lines, HEADINGS))
        c.check("TOC page numbers match the layout", pages_ok and len(lines) == len(HEADINGS),
                [(l, page_of(h)) for l, (_, h) in zip(lines, HEADINGS)])
    # nothing else changed: the multiset of non-heading, non-TOC body texts
    def content(d):
        out = []
        for p in d.paragraphs():
            s = (p.style_name or "").lower()
            if s.startswith("heading") or s.startswith("toc"):
                continue
            t = norm(p.text)
            t = re.sub(r"section \d+ [A-Za-z ]+ \(page \d+\)", "section X", t)
            out.append(t)
        return Counter(out)

    diff = (content(src) - content(doc)) + (content(doc) - content(src))
    c.check("no other text changed", not diff, dict(diff))
    c.check("table unchanged", [[norm(t.cell(r, k).text) for k in range(t.column_count)] for t in doc.tables()
                                for r in range(len(t.rows))] ==
            [[norm(t.cell(r, k).text) for k in range(t.column_count)] for t in src.tables() for r in range(len(t.rows))])


if __name__ == "__main__":
    main("w4-restructure", body)
