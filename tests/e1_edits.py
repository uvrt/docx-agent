"""A representative set of E1's edits, for any document: what the render tests and the Word
oracle apply to every fixture, and what each should then show.

On the document's first paragraph of plain text (written first where there is none): a word
replaced and given the Strong style; after it, a Heading 2 paragraph (a built-in style the
document may lack), a numbered list of three -- the third restarted at 7 -- and a bullet, a
hyperlink to the web and one to a bookmark on the heading, a cross-reference to that
bookmark, a paragraph with 20 pt red bold text and an inline picture.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field

from docx_agent import Document
from docx_agent.edit import text as _text

HEADING = "E1 Heading Marker"
BIG = "E1 big text marker"


@dataclass
class Expected:
    #: Text each page's text, joined and with whitespace collapsed, must contain.
    texts: list[str] = field(default_factory=list)
    #: The ids the edits created or touched.
    ids: dict[str, str] = field(default_factory=dict)


def png(width: int, height: int, color=(30, 140, 60)) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(color) * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _first_plain(document: Document):
    for paragraph in document.paragraphs():
        atoms = _text.atoms(paragraph._element)
        words = [w for w in paragraph.text.split() if len(w.strip(".,;:()")) >= 4]
        if (len(atoms) >= 12 and words and all(a.field is None and a.char != _text.OBJECT for a in atoms)
                and paragraph.list is None and not paragraph.ends_section):
            return paragraph
    if document.paragraphs():
        paragraph = document.paragraphs()[0]
        paragraph.set_text("A paragraph written for the edits to work on.")
        return document.paragraph(paragraph.id)
    return document.paragraph(document.append_paragraph("A paragraph written for the edits to work on.").id)


def e1_edit_set(document: Document) -> Expected:
    expected = Expected()
    first = _first_plain(document)
    word = next(w.strip(".,;:()") for w in first.text.split()
                if len(w.strip(".,;:()")) >= 4 and first.text.count(w.strip(".,;:()")) == 1) \
        if any(first.text.count(w.strip(".,;:()")) == 1 and len(w.strip(".,;:()")) >= 4 for w in first.text.split()) \
        else first.text.split()[0]
    swapped = document.anchor(word, within=first.id).replace("E1swap")
    first_id = swapped.id
    document.anchor("E1swap", within=first_id).set_style("Strong")
    expected.texts.append("E1swap")

    heading = document.insert_paragraph(HEADING, after=first_id, style="Heading 2")
    document.paragraph(heading.id).range().add_bookmark("E1Heading")
    expected.texts.append(HEADING)

    previous = heading.id
    items = []
    for text in ("E1 alpha item", "E1 beta item", "E1 gamma item", "E1 bullet item"):
        created = document.insert_paragraph(text, after=previous)
        document.paragraph(created.id).set_style(None)
        document.paragraph(created.id).clear_direct_formatting()
        items.append(created.id)
        previous = created.id
    for identifier in items[:3]:
        document.paragraph(identifier).add_to_list("number")
    document.paragraph(items[2]).restart_numbering(at=7)
    document.paragraph(items[3]).add_to_list("bullet")
    expected.texts += ["1. E1 alpha item", "2. E1 beta item", "7. E1 gamma item"]

    document.anchor("alpha", within=items[0]).add_hyperlink("https://example.com/e1")
    document.anchor("beta", within=items[1]).add_hyperlink(anchor="E1Heading")
    bullet = document.paragraph(items[3])
    document.insert_cross_reference(f"{bullet.id}@{len(bullet.text)}", "E1Heading")
    expected.texts.append(f"E1 bullet item{HEADING}")

    big = document.insert_paragraph(BIG, after=items[3])
    document.paragraph(big.id).set_style(None)
    document.paragraph(big.id).clear_direct_formatting()
    if document.paragraph(big.id).list is not None:     # it continued the bullet item before it
        document.paragraph(big.id).remove_from_list()
    document.paragraph(big.id).range().format(size=20, bold=True, color="C00000")
    picture = document.insert_picture(f"{big.id}@{len(BIG)}", png(96, 48), alt_text="E1 picture")
    expected.texts.append(BIG)
    expected.ids = {"first": first_id, "heading": heading.id, "items": ",".join(items), "big": big.id,
                    "picture": picture.id}
    return expected
