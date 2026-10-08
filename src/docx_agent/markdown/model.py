"""The neutral Markdown model both directions meet in: blocks and inline items.

The reader (:mod:`.read`) projects a document into it, the renderer (:mod:`.render`)
writes it as text, and the parser (:mod:`.parse`) turns Markdown back into the same
shapes -- which is how the reverse-parse check compares what was meant with what
markdown-it reads, and what ``insert_markdown`` (E2's write half) will consume.

Everything here is plain data.  Comments that carry ids are explicit: a :class:`Comment`
block before the block it labels, an :class:`Html` item inside a line.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Union

#: Inline marks, in the order they nest outermost first when they start together.
MARKS = ("strike", "strong", "em", "code")


# -- inline ----------------------------------------------------------------------------------


@dataclass(frozen=True)
class Text:
    """Literal text (escaped on writing) with marks and, inside a link, its target."""

    text: str
    marks: frozenset = frozenset()
    href: str | None = None


@dataclass(frozen=True)
class Html:
    """Raw inline HTML, written as is: in practice a comment, ``<!-- bm:target -->``."""

    text: str


@dataclass(frozen=True)
class Critic:
    """A CriticMarkup delimiter (``{++``, ``++}``, ``{>>``...), written as is: Markdown
    reads it as text."""

    text: str


@dataclass(frozen=True)
class Break:
    """A hard line break."""


@dataclass(frozen=True)
class Image:
    alt: str
    src: str
    title: str | None = None
    href: str | None = None


@dataclass(frozen=True)
class NoteRef:
    """A footnote reference, ``[^label]``."""

    label: str


Inline = Union[Text, Html, Critic, Break, Image, NoteRef]


# -- blocks ----------------------------------------------------------------------------------


@dataclass
class Comment:
    """An HTML comment on a line of its own: ids, and what Markdown cannot say."""

    text: str


@dataclass
class Paragraph:
    inline: list


@dataclass
class Heading:
    level: int
    inline: list


@dataclass
class CodeBlock:
    text: str


@dataclass
class Quote:
    blocks: list


@dataclass
class ListBlock:
    ordered: bool
    start: int = 1
    #: Each item is a list of blocks.
    items: list = field(default_factory=list)
    #: Write the alternative marker (``*`` or ``)``), so that a list right after another of
    #: the same kind stays a list of its own.
    alternate: bool = False


@dataclass
class Table:
    """A GFM table: the first row is the header row."""

    header: list
    rows: list
    #: Per column: ``None``, ``left``, ``center`` or ``right``.
    aligns: list


@dataclass
class HtmlBlock:
    """Raw HTML on lines of its own (a table GFM cannot express)."""

    text: str


@dataclass
class Rule:
    """A thematic break."""


@dataclass
class FootnoteDef:
    label: str
    blocks: list


Block = Union[Comment, Paragraph, Heading, CodeBlock, Quote, ListBlock, Table, HtmlBlock, Rule, FootnoteDef]


def plain_text(inline: list) -> str:
    """An inline sequence's characters, as a reader would see them: text, alt text and
    nothing of comments, delimiters or references."""
    out = []
    for item in inline:
        if isinstance(item, Text):
            out.append(item.text)
        elif isinstance(item, Break):
            out.append("\n")
    return "".join(out)


def has_content(inline: list) -> bool:
    """Whether ``inline`` writes anything a Markdown reader keeps: text that is not only
    whitespace, a break between such text, an image, a reference, a comment."""
    return any((isinstance(i, Text) and i.text.strip()) or isinstance(i, (Html, Critic, Image, NoteRef))
               for i in inline)
