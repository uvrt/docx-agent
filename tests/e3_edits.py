"""A representative set of tracked edits and comments, for any document: what the render
tests and the Word oracle apply to every fixture (ROADMAP.md, Phase E3).

Inside ``Document.tracking``: a word replaced and another deleted, text inserted, a word
made bold, a paragraph's style changed and one made a list item, a paragraph inserted and
one deleted, two paragraphs joined, a paragraph moved, a picture and a hyperlink inserted;
in the first plain table, a row inserted and one deleted, a column inserted.  Then a
comment with a reply, and a second comment resolved.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field

from docx_agent import Document, EditError
from docx_agent.edit import text as _text

AUTHOR = "E3 Agent"
DATE = "2026-10-03T12:00:00Z"
INSERTED = "E3 inserted words"
NEW_PARAGRAPH = "E3 tracked paragraph marker"
COMMENT = "E3 comment marker"
REPLY = "E3 reply marker"


@dataclass
class Expected:
    #: Text the final view (and Word's final view) must show.
    final: list[str] = field(default_factory=list)
    #: Text it must not show.
    gone: list[str] = field(default_factory=list)
    comments: list[str] = field(default_factory=list)


def png(width: int, height: int, color=(150, 40, 160)) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(color) * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _plain_paragraphs(document: Document, minimum: int = 12) -> list:
    out = []
    for paragraph in document.paragraphs():
        atoms = _text.atoms(paragraph._element)
        words = [w for w in paragraph.text.split() if len(w.strip(".,;:()")) >= 4]
        if (len(atoms) >= minimum and words and all(a.field is None and a.char != _text.OBJECT for a in atoms)
                and not paragraph.ends_section and paragraph._element.getparent().tag.endswith("body")):
            out.append(paragraph)
    return out


def _word(paragraph) -> str:
    words = [w.strip(".,;:()") for w in paragraph.text.split()]
    return next((w for w in words if len(w) >= 4 and paragraph.text.count(w) == 1), words[0])


def prepare(document: Document) -> None:
    """Paragraphs for the edits to work on, written untracked where a document has too few."""
    if len(_plain_paragraphs(document)) < 6:
        anchor = document.paragraphs()[-1].id if document.paragraphs() else None
        for k in range(6):
            text = f"A paragraph written for the tracked edits, number {k}."
            anchor = (document.insert_paragraph(text, after=anchor).id if anchor
                      else document.append_paragraph(text).id)


def e3_edit_set(document: Document, *, comments: bool = True, tracked: bool = True) -> Expected:
    expected = Expected()
    prepare(document)
    with document.tracking(author=AUTHOR, date=DATE) if tracked else _untracked():
        paragraphs = _plain_paragraphs(document)
        first, second, third, fourth, fifth = paragraphs[:5]
        word = _word(first)
        document.anchor(word, within=first.id).replace("E3swap")
        expected.final.append("E3swap")
        expected.gone.append(word)
        document.insert_text(f"{second.id}@0", INSERTED + " ")
        expected.final.append(INSERTED)
        third_word = _word(third)
        document.anchor(third_word, within=third.id).format(bold=True)
        document.paragraph(fourth.id).style = "Heading 2"
        created = document.insert_paragraph(NEW_PARAGRAPH, after=fourth.id)
        expected.final.append(NEW_PARAGRAPH)
        document.paragraph(created.id).add_to_list("number")
        gone = document.paragraph(fifth.id).text
        document.delete_block(fifth.id)
        expected.gone.append(gone.split()[0] + " " + gone.split()[1] if len(gone.split()) > 1 else gone)
        later = _plain_paragraphs(document)
        if len(later) >= 3:
            mover = later[-1]
            try:
                document.move_block(mover.id, before=document.paragraph(first.id).id)
            except EditError:
                pass
        picture_at = document.paragraph(second.id)
        document.insert_picture(f"{picture_at.id}@{len(picture_at.text)}", png(30, 20), alt_text="E3 picture")
        document.find(INSERTED, within=picture_at.id)[0].add_hyperlink("https://example.com/e3")
        table = _plain_table(document)
        if table is not None:
            document.insert_row(table.id, 0)
            document.delete_row(table.id, len(document.table(table.id).rows) - 1)
            document.insert_column(table.id, 0)
    if not comments:
        return expected
    target = document.paragraph(third.id)
    comment = document.add_comment(target.range(0, min(8, len(target.text))), COMMENT, author="E3 Reviewer",
                                   date=DATE)
    document.reply_to_comment(comment.id, REPLY, author=AUTHOR, date=DATE)
    resolved = document.add_comment(document.paragraph(second.id).range(0, 4), "E3 resolved comment",
                                    author="E3 Reviewer", date=DATE)
    document.resolve_comment(resolved.id)
    expected.comments = [COMMENT, REPLY, "E3 resolved comment"]
    return expected


def _untracked():
    from contextlib import nullcontext

    return nullcontext()


def _plain_table(document: Document):
    for table in document.tables():
        grid = table.grid()
        if len(grid) >= 2 and grid and all(len(line) == len(grid[0]) and all(c is not None for c in line)
                                           for line in grid) and len(grid[0]) >= 2:
            if len({c.id for line in grid for c in line}) == len(grid) * len(grid[0]):
                return table
    return None
