#!/usr/bin/env python3
"""Write the templates E6 is tested on, ``tests/fixtures/generated/templates/``.

``brand.dotx`` (compatibility mode 15)
    a template of the kind File > New from a template starts from: its own styles (heading
    1 redefined, a paragraph and a character style of its own, header, footer), a bullet
    list, a header and a footer, Letter paper with a 1000-twip top margin, even and odd
    headers, a body of four paragraphs, and properties (title, subject, author, keywords,
    company) -- the probe's template (``tools/e6_probe.py``, ``template``), which Word made
    a document from.
``brand-macros.dotm``
    the same as a macro-enabled template: a VBA project and its ``vbaData.xml``, both
    stand-ins written here (the project's bytes are not a VBA project; nothing runs them),
    which a ``.docx`` made from it must not carry.
``brand-mode14.dotx``
    the same in compatibility mode 14, which a new document made from it leaves.

They sit two directories below the corpus, so no suite picks them up as documents.
Deterministic: running this again writes the same bytes.

    python tools/make_template_fixtures.py
"""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import e6_probe  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "generated" / "templates"
DATE = (2026, 10, 4, 12, 0, 0)
_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
_VBA_DATA = "http://schemas.microsoft.com/office/2006/relationships/wordVbaData"


def _rewrite(data: bytes, change) -> bytes:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts = change(parts)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, part in parts.items():
            info = zipfile.ZipInfo(name, date_time=DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, part)
    return buffer.getvalue()


def macros(parts: dict[str, bytes]) -> dict[str, bytes]:
    types = parts["[Content_Types].xml"].decode()
    types = types.replace(e6_probe.WML + ".template.main+xml",
                          "application/vnd.ms-word.template.macroEnabledTemplate.main+xml")
    types = types.replace(
        "</Types>",
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
        '<Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/></Types>')
    parts["[Content_Types].xml"] = types.encode()
    rels = parts["word/_rels/document.xml.rels"].decode()
    parts["word/_rels/document.xml.rels"] = rels.replace(
        "</Relationships>", f'<Relationship Id="rIdVba" Type="{_VBA}" Target="vbaProject.bin"/></Relationships>'
    ).encode()
    parts["word/vbaProject.bin"] = b"docx-agent test stand-in, not a VBA project\n"
    parts["word/_rels/vbaProject.bin.rels"] = (
        f'{e6_probe.DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{_VBA_DATA}" Target="vbaData.xml"/></Relationships>').encode()
    parts["word/vbaData.xml"] = (
        f'{e6_probe.DECL}<wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">'
        "<wne:mcds/></wne:vbaSuppData>").encode()
    return parts


def mode14(parts: dict[str, bytes]) -> dict[str, bytes]:
    settings = parts["word/settings.xml"].decode().replace('w:val="15"', 'w:val="14"')
    parts["word/settings.xml"] = settings.encode()
    return parts


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    base = e6_probe.template("dotx")
    outputs = {"brand.dotx": base, "brand-macros.dotm": _rewrite(base, macros),
               "brand-mode14.dotx": _rewrite(base, mode14)}
    for name, data in outputs.items():
        (OUT / name).write_bytes(data)
        print(f"wrote {(OUT / name).relative_to(ROOT)}")


if __name__ == "__main__":
    main()
