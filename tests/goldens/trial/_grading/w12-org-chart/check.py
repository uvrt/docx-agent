"""W12: the picture inline after its announcing paragraph, 12 cm wide in proportion, centred,
with alternative text; the rest unchanged."""
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import INPUTS, main, norm  # noqa: E402

from docx_agent import Document  # noqa: E402

INPUT = INPUTS / "w12-org-chart/input/annual-report-draft.docx"
ANNOUNCE = "The chart below shows the new structure."
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
EMU_PER_CM = 360000


def body(c, out: Path):
    doc, src = Document.open(out), Document.open(INPUT)
    c.check("validate() clean", not doc.validate(), doc.validate())
    paragraphs = doc.paragraphs()
    texts = [norm(p.text) for p in paragraphs]
    where = texts.index(ANNOUNCE) if ANNOUNCE in texts else None
    c.check("announcing paragraph kept", where is not None)
    if where is None:
        return
    holder = paragraphs[where + 1]
    inline = holder._element.find(f".//{WP}inline")
    c.check("an inline picture in the next paragraph, on its own", inline is not None
            and not texts[where + 1], texts[where + 1][:40])
    if inline is None:
        return
    extent = inline.find(f"{WP}extent")
    cx, cy = int(extent.get("cx")), int(extent.get("cy"))
    c.check("12 cm wide", abs(cx - 12 * EMU_PER_CM) <= 12700, cx / EMU_PER_CM)
    c.check("in proportion (the image is 3:2)", abs(cx / cy - 1.5) < 0.01, cx / cy)
    jc = holder._element.find(f"{W}pPr/{W}jc")
    c.check("centred", jc is not None and jc.get(f"{W}val") == "center",
            None if jc is None else jc.get(f"{W}val"))
    alt = inline.find(f"{WP}docPr").get("descr") or ""
    c.check("alternative text says what it shows",
            alt == "Organisation chart: the managing board over three divisions", alt)
    with zipfile.ZipFile(out) as archive:
        media = [n for n in archive.namelist() if n.startswith("word/media/")]
    c.check("one image in the package", len(media) == 1, media)
    rest = [t for i, t in enumerate(texts) if i != where + 1]
    c.check("the text unchanged", rest == [norm(p.text) for p in src.paragraphs()], rest)


if __name__ == "__main__":
    main("w12-org-chart", body)
