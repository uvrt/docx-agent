#!/usr/bin/env python3
"""Measure the colour Word draws for a theme colour with a tint or shade (production round
21, T7): ``w:color``'s ``w:themeColor`` / ``w:themeTint`` / ``w:themeShade`` on text and
``w:shd``'s ``w:themeFill`` / ``w:themeFillTint`` / ``w:themeFillShade`` on a paragraph.

A made-up document: one paragraph style per swatch for text and one for shading, their
``w:val`` / ``w:fill`` a decoy (``FF00FF`` / ``00FF00``) so the colour drawn shows whether
Word follows the theme attributes.  Word -- through ``tests/oracle.py``'s machine-wide lock
-- exports it to PDF (in batches: one document of all 27 swatches did not open in time);
text colours are read from the PDF's characters, fills from a 200 dpi raster beside the
text, into ``tests/observations/theme-tints-word.json``.

    python tools/theme_tint_probe.py

Nothing Word writes is committed: only these facts.
"""
import io
import json
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "src"))

import oracle  # noqa: E402
from docx_agent import Document  # noqa: E402

OBSERVED = ROOT / "tests" / "observations" / "theme-tints-word.json"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

# (theme colour, tint, shade)
SWATCHES = [
    ("accent1", None, "BF"), ("accent1", "99", None), ("accent1", "66", None), ("accent1", "33", None),
    ("accent2", None, "80"), ("accent2", "BF", None), ("text1", "BF", None), ("text1", "80", None),
    ("background2", None, "BF"), ("background2", None, "E6"), ("background1", None, "A6"),
    ("background2", "BF", None), ("text2", "99", None), ("accent4", None, None), ("accent5", None, "40"),
    ("accent6", "1A", None), ("dark2", None, None), ("light2", None, "80"), ("accent1", "99", "BF"),
    ("accent3", "F2", None), ("accent3", None, "0D"), ("hyperlink", None, None), ("accent2", "00", None),
    ("accent2", None, "FF"), ("accent2", "FF", None), ("accent4", "7F", None), ("accent4", None, "7F"),
]


def make() -> bytes:
    doc = Document.new()
    first = doc.paragraphs()[0]
    first.text = "Colour probe"
    last = first.id
    for k, _ in enumerate(SWATCHES):
        doc.styles.add(f"Text sample {k}", based_on="Normal")
        doc.styles.add(f"Fill sample {k}", based_on="Normal")
    for k, _ in enumerate(SWATCHES):
        for kind in ("Text", "Fill"):
            made = doc.paragraph(last).insert_after(f"{kind} sample {k} lorem ipsum dolor")
            doc.paragraph(made.id).set_style(f"{kind} sample {k}")
            last = made.id
    data = doc.to_bytes()
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            raw = src.read(item.filename)
            if item.filename == "word/styles.xml":
                root = etree.fromstring(raw)
                by_name = {s.find("w:name", NS).get(f"{{{W}}}val"): s for s in root.findall("w:style", NS)}
                for k, (theme, tint, shade) in enumerate(SWATCHES):
                    text = by_name[f"Text sample {k}"]
                    rpr = text.find("w:rPr", NS)
                    if rpr is None:
                        rpr = etree.SubElement(text, f"{{{W}}}rPr")
                    color = etree.SubElement(rpr, f"{{{W}}}color")
                    color.set(f"{{{W}}}val", "FF00FF")
                    color.set(f"{{{W}}}themeColor", theme)
                    if tint:
                        color.set(f"{{{W}}}themeTint", tint)
                    if shade:
                        color.set(f"{{{W}}}themeShade", shade)
                    fill = by_name[f"Fill sample {k}"]
                    ppr = fill.find("w:pPr", NS)
                    if ppr is None:
                        ppr = etree.Element(f"{{{W}}}pPr")
                        fill.find("w:qFormat", NS).addnext(ppr)
                    shd = etree.SubElement(ppr, f"{{{W}}}shd")
                    shd.set(f"{{{W}}}val", "clear")
                    shd.set(f"{{{W}}}color", "auto")
                    shd.set(f"{{{W}}}fill", "00FF00")
                    shd.set(f"{{{W}}}themeFill", theme)
                    if tint:
                        shd.set(f"{{{W}}}themeFillTint", tint)
                    if shade:
                        shd.set(f"{{{W}}}themeFillShade", shade)
                raw = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            dst.writestr(item, raw)
    return out.getvalue()


def read(pdf: Path) -> dict:
    import pypdfium2

    chars = oracle.pdf_chars(pdf)
    document = pypdfium2.PdfDocument(str(pdf))
    results = {}
    for page_index, page in enumerate(document):
        textpage = page.get_textpage()
        bitmap = page.render(scale=200 / 72).to_pil().convert("RGB")
        height = page.get_height()
        full = "".join(c.char for c in chars[page_index])
        for k in range(len(SWATCHES)):
            for kind in ("Text", "Fill"):
                label = f"{kind} sample {k} "
                at = full.find(label)
                if at < 0:
                    continue
                if kind == "Text":
                    results.setdefault(str(k), {})["text"] = "%02X%02X%02X" % chars[page_index][at].color
                else:
                    # A pixel right of the text, mid-line.
                    left, bottom, right, top = textpage.get_charbox(at + len(label) + 5)
                    x = int((right + 120) * 200 / 72)
                    y = int((height - (bottom + top) / 2) * 200 / 72)
                    results.setdefault(str(k), {})["fill"] = "%02X%02X%02X" % bitmap.getpixel((x, y))
    document.close()
    return results


def main() -> None:
    swatches = list(SWATCHES)
    rows = []
    with oracle.session() as word:
        for start in range(0, len(swatches), 12):
            SWATCHES[:] = swatches[start:start + 12]
            got = word.export_pdf(make(), name=f"theme-tints-{start}")
            if not got.ok:
                print(start, got.outcome, got.detail[:200])
                continue
            found = read(got.path)
            got.path.unlink()
            for k, (theme, tint, shade) in enumerate(SWATCHES):
                drawn = found.get(str(k), {})
                rows.append({"theme": theme, "tint": tint, "shade": shade, "word": drawn.get("text"),
                             "fill": drawn.get("fill")})
    SWATCHES[:] = swatches
    OBSERVED.write_text(json.dumps({"swatches": rows}, indent=1) + "\n")


if __name__ == "__main__":
    main()
