"""Every E5 operation on every fixture (ROADMAP.md, Phase E5, "Done when": the standard
gates): edit, save, reopen and read back what the edit wrote; no validity problem added --
grid consistency, unique drawing ids, content-control and custom XML integrity among them;
the compatibility mode kept; undo to the original bytes, redo to the edited ones; and the
edited document laid out by docx2svg, every paragraph the edit made placed by ``where()``
(or said to be past where the layout stops).
"""

from __future__ import annotations

import pytest

from docx_agent import Document
from docx_agent.layout import Unknown
from docx_agent.validate import check

from e5_edits import OPERATIONS
from test_roundtrip import entries


@pytest.mark.parametrize("operation", OPERATIONS, ids=lambda f: f.__name__[3:])
def test_operation(docx_path, operation):
    data = docx_path.read_bytes()
    document = Document.open(data)
    mode = document.compatibility_mode
    before = set(check(document.package))
    read_back = operation(document)
    edited = document.to_bytes()

    saved = Document.open(edited)
    assert read_back(saved), f"{operation.__name__} did not read back from the saved document"
    assert saved.compatibility_mode == mode
    assert set(check(saved.package)) - before == set()

    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(data)
    while document.redo():
        pass
    assert entries(document.to_bytes()) == entries(edited)


@pytest.mark.parametrize("operation", OPERATIONS, ids=lambda f: f.__name__[3:])
def test_operation_renders_and_places_what_it_made(docx_path, operation):
    document = Document.open(docx_path.read_bytes())
    known = {p.id for s in document.stories for p in s.paragraphs}
    operation(document)
    layout = document.layout()
    assert layout.pages
    made = [p for s in document.stories for p in s.paragraphs if p.id not in known and s.name == "body"]
    for paragraph in made:
        placed = layout.where(paragraph.id)
        if isinstance(placed, Unknown):
            # Past a stop, in a text box or a cell docx2svg does not place by id, or empty.
            assert placed.reason != "not-drawn" or not paragraph.text or paragraph._element.xpath(
                "ancestor::w:txbxContent or ancestor::w:tc",
                namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}), (paragraph.id, placed)
        else:
            assert all(p.page >= 1 for p in placed)
    document.render_svg(pages=[1])
