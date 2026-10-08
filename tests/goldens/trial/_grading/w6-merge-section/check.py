"""W6: merge B's 'Working at height' section into A, in A's styles, with picture and footnote."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

A = INPUTS / "w6-merge-section/input/field-handbook.docx"
SECTION = [
    "Working at height",
    "Falls from ladders caused two lost-time injuries this year. From 1 September, the following rules apply "
    "to any work more than 2 metres above ground.",
    "Use a ladder only for short tasks of under 30 minutes.",
    "Keep three points of contact while climbing.",
    "Never stand on the top three rungs.",
    "Before you climb, check:",
    "The ladder has a current inspection tag.",
    "The ground is firm and level.",
    "A second person is present to foot the ladder.",
    "",  # the picture's paragraph
    "Figure: the safe ladder angle is 75 degrees, one unit out for every four units up.",
]
LOOKS = ("font", "sz", "color", "b", "i")


def eff(doc, pid):
    e = doc.state(pid)["blocks"][0]["effective"]
    return {k: e.get(k) for k in LOOKS}


def body(c, out: Path):
    doc = Document.open(out)
    src = Document.open(A)
    c.check("validate() clean", not doc.validate(), doc.validate())
    c.check("no tracked changes", not doc.revisions(), len(doc.revisions()))
    ps = doc.paragraphs()
    texts = [norm(p.text) for p in ps]
    try:
        i = texts.index("Working at height")
    except ValueError:
        c.check("section heading present", False, texts)
        return
    got = texts[i:i + len(SECTION)]
    c.check("section content complete and in order", got == SECTION, got)
    c.check("section placed after Site access, before Vehicle checks",
            texts.index("Site access") < i < texts.index("Vehicle checks")
            and texts[i - 1] == "Park only in marked bays." and texts[i + len(SECTION)] == "Vehicle checks",
            texts[max(0, i - 2):i + len(SECTION) + 2])
    for unwanted in ("Safety bulletin 2026-07", "Issued by the HSE team for all field staff.", "Hot weather",
                     "Drink water regularly"):
        c.check(f"not brought: {unwanted[:30]}", not any(t.startswith(unwanted) for t in texts))
    names = {s.name for s in doc.styles}
    c.check("no bulletin styles in the handbook", not ({"Bulletin Body", "Bulletin Note"} & names),
            {"Bulletin Body", "Bulletin Note"} & names)
    h = ps[i]
    c.check("heading uses Heading 1", (h.style_name or "").lower() == "heading 1", h.style_name)
    ref_h = ps[texts.index("Site access")]
    c.check("heading looks like the handbook's headings", eff(doc, h.id) == eff(src, ref_h.id),
            (eff(doc, h.id), eff(src, ref_h.id)))
    ref_body = eff(src, src.paragraphs()[1].id)
    body_idx = [i + 1, i + 5, i + 10]
    bad = [(texts[k][:30], ps[k].style_name, eff(doc, ps[k].id)) for k in body_idx
           if (ps[k].style_name or "") != "Handbook Body" or eff(doc, ps[k].id) != ref_body]
    c.check("body paragraphs and figure line in Handbook Body, looking like the handbook", not bad, bad)
    runs_bad = []
    for k in range(i, i + len(SECTION)):
        for run in doc.state(ps[k].id)["blocks"][0].get("runs", []):
            extra = {key for key in run if key in ("font", "sz", "color", "rFonts", "highlight")}
            if extra:
                runs_bad.append((texts[k][:25], run))
    c.check("no direct font/size/colour formatting carried over", not runs_bad, runs_bad[:4])
    bul = [ps[k].list for k in range(i + 2, i + 5)]
    num = [ps[k].list for k in range(i + 6, i + 9)]
    c.check("bullets are a bulleted list", all(x and x.format == "bullet" for x in bul), bul)
    c.check("checklist is a numbered list", all(x and x.format == "decimal" for x in num), num)
    labels = [doc.to_markdown(ps[k].id, ids=False).strip()[:2] for k in range(i + 6, i + 9)]
    c.check("checklist numbered 1-3", labels == ["1.", "2.", "3."], labels)
    pics = doc.pictures()
    c.check("one picture added", len(pics) == len(src.pictures()) + 1, len(pics))
    if pics:
        c.check("picture alt text kept", any("Ladder at a 75 degree angle" in (p.alt_text or "") for p in pics),
                [p.alt_text for p in pics])
        md = doc.to_markdown(ps[i + 9].id)
        c.check("picture sits between checklist and figure line", "![" in md or "d:" in md, md[:120])
    notes = doc.notes()
    c.check("footnotes: handbook's two plus one", len(notes) == 3, [(n.id, n.text) for n in notes])
    c.check("bulletin footnote text present", any("stepladders on uneven ground" in n.text for n in notes))
    md = doc.to_markdown(ps[i + 1].id, ids=False)
    c.check("footnote reference in the section's first paragraph", "[^" in md, md[:200])
    # handbook otherwise unchanged
    rest = [t for k, t in enumerate(texts) if not (i <= k < i + len(SECTION))]
    c.check("rest of the handbook unchanged", rest == [norm(p.text) for p in src.paragraphs()],
            [x for x in zip(rest, [norm(p.text) for p in src.paragraphs()]) if x[0] != x[1]][:3])


if __name__ == "__main__":
    main("w6-merge-section", body)
