"""The losslessness gate.

Everything else rests on one property: reading a document and writing it back changes
nothing.  It is held over the whole corpus, every part compared byte for byte, and reading
-- every paragraph's text, every id, every table's grid -- must not change a byte either:
ids are derived, never stamped, until something is edited.
"""

from __future__ import annotations

import io
import zipfile

from docx_agent import Document
from docx_agent.oxml.package import WordPackage, normalize_part_path


def entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {normalize_part_path(info.filename): archive.read(info)
                for info in archive.infolist() if not info.is_dir()}


def order(data: bytes) -> list[str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return [info.filename for info in archive.infolist() if not info.is_dir()]


def test_open_and_save_is_byte_identical(docx_path):
    original = entries(docx_path.read_bytes())
    written = entries(WordPackage.open(str(docx_path)).to_bytes())
    assert set(written) == set(original)
    assert [name for name, data in original.items() if written[name] != data] == []


def test_entry_order_is_preserved(docx_path):
    data = docx_path.read_bytes()
    assert order(Document.open(data).to_bytes()) == order(data)


def test_reading_everything_changes_nothing(docx_path):
    """Parse every part, read every story, paragraph, run, table and section."""
    data = docx_path.read_bytes()
    document = Document.open(data)
    for name in document.package.part_names:
        if name.endswith((".xml", ".rels")):
            document.package.tree(name)
    for story in document.stories:
        for paragraph in story.paragraphs:
            paragraph.text, paragraph.id, paragraph.style
            for view in ("current", "original", "markup"):
                paragraph.text_in(view)
            for run in paragraph.runs:
                run.text, run.id
        for table in story.tables:
            table.grid()
            [row.id for row in table.rows]
    document.sections()
    assert document.package.dirty_parts == frozenset()
    assert entries(document.to_bytes()) == entries(data)


def test_save_writes_the_same_bytes_as_to_bytes(docx_path, tmp_path):
    document = Document.open(str(docx_path))
    target = tmp_path / "out.docx"
    document.save(target)
    assert entries(target.read_bytes()) == entries(document.to_bytes())
