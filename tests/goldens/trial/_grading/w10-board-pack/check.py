"""W10: five reports' Highlights sections copied under their departments' headings, nothing
else from the reports, and the title and author set."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

FOLDER = INPUTS / "w10-board-pack/input"
ORDER = [("Finance", "finance"), ("Operations", "operations"), ("People", "people"),
         ("Sales", "sales"), ("IT", "it")]


def highlights(path: Path) -> list[str]:
    texts = [norm(p.text) for p in Document.open(path).paragraphs()]
    start = texts.index("Highlights")
    return texts[start + 1:texts.index("Detail")]


def body(c, out: Path):
    doc = Document.open(out)
    c.check("validate() clean", not doc.validate(), doc.validate())
    texts = [norm(p.text) for p in doc.paragraphs()]
    c.check("the introduction kept", texts[:2] == [norm(p.text) for p in Document.open(
        FOLDER / "board-pack-q3.docx").paragraphs()][:2], texts[:2])
    heads = [texts.index(name) if name in texts else -1 for name, _ in ORDER]
    c.check("the five headings, in order", heads == sorted(heads) and -1 not in heads, heads)
    if -1 in heads:
        return
    for (name, key), start, end in zip(ORDER, heads, heads[1:] + [len(texts)]):
        section = texts[start + 1:end]
        want = ["Highlights"] + highlights(FOLDER / f"{key}-report.docx")
        c.check(f"{name}: its report's Highlights, nothing else", section == want, section)
    c.check("no Detail or Next quarter section came along",
            "Detail" not in texts and "Next quarter" not in texts)
    props = doc.properties
    c.check("title set", props.get("title") == "Board pack, Q3 2026", props.get("title"))
    c.check("author set", props.get("author") == "Company Secretary", props.get("author"))


if __name__ == "__main__":
    main("w10-board-pack", body)
