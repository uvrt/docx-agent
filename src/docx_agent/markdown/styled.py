"""Text with inline emphasis, as ``word_set_text`` reads it: one paragraph's text in which
``**...**`` is the strong role's character style, ``*...*`` the emphasis role's (the styles
``insert_markdown``'s style map names, Strong and Emphasis by default), and
``[text]{Style name}`` a character style the document has, named.

The rules are CommonMark's for ``*`` (markdown-it-py's own emphasis rule, so ``5 * 3`` and an
unclosed ``**`` stay as they are), with these differences, chosen for prose:

* ``_`` is never emphasis: ``file_name_here`` and ``_draft_`` keep their underscores.
* Nothing else is Markdown: no links, code spans, HTML, entities or line-start blocks (``#``,
  ``-``, ``1.``) -- ``a & b``, ``<b>`` and ``# 3`` are their characters.
* A backslash before ASCII punctuation writes it literally (``\\*`` is ``*``, ``\\[`` is
  ``[``, ``\\\\`` a backslash), as CommonMark escapes; a backslash before anything else is a
  backslash (``C:\\Users``).
* A span is ``[`` text ``]`` right followed by ``{`` a character style's name ``}`` (an
  ``{style="Name"}`` or Pandoc's ``{custom-style="Name"}`` read as the name); its text may
  hold ``**``/``*``, written as direct bold/italic beside the span's style, as
  ``insert_markdown`` writes a run with several marks.  Spans do not nest.

:func:`parse` gives the plain text and the styled pieces.  Text with none of these
(``marked`` false and the same characters) is plain text: the caller writes it as before.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from functools import lru_cache

from markdown_it import MarkdownIt

#: Private-use stand-ins while markdown-it reads the text: a span's start and end, and an
#: underscore (never emphasis here).
_OPEN, _CLOSE, _UNDERSCORE = "\ue000", "\ue001", "\ue002"
_ESCAPABLE = set("!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~")
#: ``{Name}``, ``{style=Name}``, ``{style="Name"}``, ``{custom-style="Name"}``.
_NAMED = re.compile(r'\s*(?:(?:custom-)?style\s*=\s*)?(?:"([^"]+)"|\'([^\']+)\'|([^"\'{}]+?))\s*')


@dataclass(frozen=True)
class Piece:
    """``text`` at ``[start, end)`` of the plain text, with its marks (``strong``, ``em``)
    and the character style a span names (``None``: none)."""

    start: int
    end: int
    text: str
    marks: frozenset = frozenset()
    style: str | None = None


@dataclass
class Styled:
    #: The characters to write.
    plain: str
    #: The plain text cut into pieces, in order, covering all of it.
    pieces: list[Piece] = field(default_factory=list)

    @property
    def marked(self) -> bool:
        """Whether any piece is strong, emphasis or in a named style."""
        return any(piece.marks or piece.style for piece in self.pieces)

    def styles_named(self) -> list[str]:
        return list(dict.fromkeys(p.style for p in self.pieces if p.style))


@lru_cache(maxsize=1)
def _parser() -> MarkdownIt:
    """Inline only, with ``*`` emphasis and backslash escapes and nothing else (compiled:
    safe to share between threads)."""
    from .parse import compiled

    return compiled(MarkdownIt("zero").enable(["emphasis", "escape"]))


def _spans(text: str) -> tuple[str, list[str]]:
    """``text`` with each span's ``[``, ``]{name}`` replaced by stand-ins, the names in
    order; escaped brackets and underscores kept for markdown-it; underscores as stand-ins."""
    out: list[str] = []
    names: list[str] = []
    k, n = 0, len(text)
    open_at: int | None = None        # index in out of an unclosed "["
    while k < n:
        ch = text[k]
        if ch == "\\" and k + 1 < n and text[k + 1] in _ESCAPABLE:
            out.append(text[k:k + 2])
            k += 2
            continue
        if ch == "_":
            out.append(_UNDERSCORE)
        elif ch == "[":
            open_at = len(out)
            out.append(ch)
        elif ch == "]" and open_at is not None and k + 1 < n and text[k + 1] == "{":
            close = text.find("}", k + 2)
            body = text[k + 2:close] if close != -1 else ""
            found = _NAMED.fullmatch(body) if close != -1 and "{" not in body and "\n" not in body else None
            name = next((g for g in found.groups() if g), "").strip() if found else ""
            if name:
                out[open_at] = _OPEN
                out.append(_CLOSE)
                names.append(name)
                open_at = None
                k = close + 1
                continue
            out.append(ch)
        else:
            out.append(ch)
        k += 1
    return "".join(out), names


def parse(text: str) -> Styled:
    """The plain text and pieces of one paragraph's ``text`` (no ``\\n``; ``\\v`` and every
    other character pass through)."""
    prepared, names = _spans(text)
    if not names and "*" not in prepared and "\\" not in prepared:
        return Styled(text, [Piece(0, len(text), text)] if text else [])
    tokens = _parser().parseInline(prepared)[0].children or []
    pieces: list[Piece] = []
    marks: list[str] = []
    style: str | None = None
    position = 0
    spans = iter(names)

    def add(chunk: str) -> None:
        nonlocal position
        chunk = chunk.replace(_UNDERSCORE, "_")
        if not chunk:
            return
        current = frozenset(marks)
        last = pieces[-1] if pieces else None
        if last is not None and last.marks == current and last.style == style:
            pieces[-1] = Piece(last.start, position + len(chunk), last.text + chunk, current, style)
        else:
            pieces.append(Piece(position, position + len(chunk), chunk, current, style))
        position += len(chunk)

    for token in tokens:
        kind = token.type
        if kind in ("strong_open", "em_open"):
            marks.append("strong" if kind == "strong_open" else "em")
        elif kind in ("strong_close", "em_close"):
            mark = "strong" if kind == "strong_close" else "em"
            if mark in marks:
                marks.reverse()
                marks.remove(mark)
                marks.reverse()
        else:
            content = token.content
            # Stand-ins split the text: a span opens or closes between the pieces.
            for part in re.split(f"([{_OPEN}{_CLOSE}])", content):
                if part == _OPEN:
                    style = next(spans, None)
                elif part == _CLOSE:
                    style = None
                else:
                    add(part)
    plain = "".join(piece.text for piece in pieces)
    return Styled(plain, pieces)
