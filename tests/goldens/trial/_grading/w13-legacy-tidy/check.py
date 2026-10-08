"""W13: compatibility mode 15, and the pasted paragraph's direct formatting gone so it looks
like the other body text; the words unchanged."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

INPUT = INPUTS / "w13-legacy-tidy/input/travel-policy.docx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def run_formatting(paragraph) -> list[dict]:
    out = []
    for run in paragraph._element.iter(f"{W}r"):
        props = run.find(f"{W}rPr")
        out.append({} if props is None else {child.tag.split("}")[1] for child in props
                                             if child.tag.split("}")[1] not in ("lang",)})
    return out


def body(c, out: Path):
    doc, src = Document.open(out), Document.open(INPUT)
    c.check("validate() clean", not doc.validate(), doc.validate())
    c.check("compatibility mode 15", doc.compatibility_mode == 15, doc.compatibility_mode)
    texts = [norm(p.text) for p in doc.paragraphs()]
    c.check("the words unchanged", texts == [norm(p.text) for p in src.paragraphs()], texts)
    paragraphs = {norm(p.text)[:12]: p for p in doc.paragraphs()}
    pasted = paragraphs.get("Economy clas")
    train = paragraphs.get("Train is the")
    c.check("the pasted paragraph has no direct run formatting left",
            pasted is not None and all(not f for f in run_formatting(pasted)),
            pasted is not None and run_formatting(pasted))
    c.check("it has the body text's style", pasted is not None and train is not None
            and pasted.style_name == train.style_name,
            (pasted and pasted.style_name, train and train.style_name))
    others = [run_formatting(p) for p in doc.paragraphs() if not norm(p.text).startswith("Economy")]
    before = [run_formatting(p) for p in src.paragraphs() if not norm(p.text).startswith("Economy")]
    c.check("the other paragraphs unchanged", others == before)
    c.check("no tracked changes", not doc.changes())


if __name__ == "__main__":
    main("w13-legacy-tidy", body)
