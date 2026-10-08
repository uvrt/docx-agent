"""The model written as Markdown text: CommonMark plus GFM tables and strikethrough, plus
footnotes.

Text is escaped so it reads back as text: backslash, backtick, ``*``, ``_``, ``[``, ``]``,
``<`` and ``~`` everywhere; ``|`` in table cells; ``&`` where it would start an entity;
``{`` where it would open CriticMarkup; and at the start of a line what would start a
block (``#``, ``>``, ``-``, ``+``, ``=``, ``1.``/``1)``).  Emphasis delimiters never touch
whitespace (it is moved outside them).

**Self-checked inline.**  CommonMark's emphasis rules (left- and right-flanking runs, the
rule of three) cannot always express a mark that starts or ends next to punctuation inside
a word.  Each paragraph's inline Markdown is therefore parsed back with markdown-it before
it is used; if it does not read back as meant, strikethrough and then emphasis are dropped
from that block (the text is never changed), and the model is updated to say so -- so
what the reverse-parse check compares is what was written.
"""

from __future__ import annotations

import re

from . import model as m
from .parse import model_inline, normalise_inline, parser, _inline as _parsed_inline

_ALWAYS = frozenset("\\`*_[]<~")
_ENTITY = re.compile(r"&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});")
_LINE_START = re.compile(r"([#>+=-])|(\d{1,9})([.)])")
_CRITIC_OPEN = ("++", "--", ">>", "==", "~~")
_DELIMITERS = {"strong": "**", "em": "*", "strike": "~~"}


def render(blocks: list) -> str:
    """The blocks as Markdown, ending in one newline."""
    lines = _blocks(blocks)
    while lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines) + "\n"


# -- blocks ----------------------------------------------------------------------------------


def _blocks(blocks: list) -> list[str]:
    """Lines; a :class:`~.model.Comment` sits directly above the block after it, other
    blocks are separated by a blank line."""
    out: list[str] = []
    for k, block in enumerate(blocks):
        lines = _block(block)
        if not lines:
            continue
        out.extend(lines)
        attached = isinstance(block, m.Comment) and k + 1 < len(blocks) and not isinstance(
            blocks[k + 1], (m.Comment, m.FootnoteDef))
        if not attached:
            out.append("")
    return out


def _block(block) -> list[str]:
    if isinstance(block, m.Comment):
        return [f"<!-- {block.text} -->"]
    if isinstance(block, m.Paragraph):
        _settle(block)
        text = _inline(block.inline)
        return text.split("\n") if text else []
    if isinstance(block, m.Heading):
        _settle(block, heading=True)
        text = _inline(block.inline, line_start=False)
        if text.endswith("#"):
            text = text[:-1] + "\\#"
        return [("#" * block.level + " " + text).rstrip()]
    if isinstance(block, m.CodeBlock):
        fence = "`" * max(3, _longest_run(block.text, "`") + 1)
        body = block.text if block.text.endswith("\n") or not block.text else block.text + "\n"
        return [fence] + body.split("\n")[:-1] + [fence]
    if isinstance(block, m.Quote):
        inner = _blocks(block.blocks)
        while inner and inner[-1] == "":
            inner.pop()
        return [("> " + line).rstrip() for line in inner]
    if isinstance(block, m.ListBlock):
        return _list(block)
    if isinstance(block, m.Table):
        return _table(block)
    if isinstance(block, m.HtmlBlock):
        return block.text.rstrip("\n").split("\n")
    if isinstance(block, m.Rule):
        return ["---"]
    if isinstance(block, m.FootnoteDef):
        inner = _blocks(block.blocks)
        while inner and inner[-1] == "":
            inner.pop()
        if not inner:
            return [f"[^{block.label}]:"]
        return [f"[^{block.label}]: {inner[0]}"] + [("    " + line).rstrip() for line in inner[1:]]
    raise TypeError(f"not a block: {block!r}")


def _list(block: m.ListBlock) -> list[str]:
    out: list[str] = []
    number = block.start
    for k, item in enumerate(block.items):
        if block.ordered:
            marker = f"{number}{')' if block.alternate else '.'} "
            number += 1
        else:
            marker = "* " if block.alternate else "- "
        inner = _item(item)
        if not inner:
            out.append(marker.rstrip())
            continue
        pad = " " * len(marker)
        out.append(marker + inner[0])
        out.extend((pad + line).rstrip() if line else "" for line in inner[1:])
        if k + 1 < len(block.items) and any(line == "" for line in inner):
            out.append("")  # a loose item stays readable; the AST forgets tightness
    return out


def _item(blocks: list) -> list[str]:
    """A list item's lines: a nested list follows its paragraph without a blank line."""
    lines: list[str] = []
    for k, block in enumerate(blocks):
        if lines and not isinstance(block, m.ListBlock) and not isinstance(blocks[k - 1], m.Comment):
            lines.append("")
        lines.extend(_block(block))
    return lines


def _table(block: m.Table) -> list[str]:
    for cells in [block.header] + block.rows:
        for cell in cells:
            _settle_inline(cell, cell=True)

    def row(cells: list) -> str:
        return "| " + " | ".join(_inline(cell, line_start=False, table=True) for cell in cells) + " |"

    delimiter = {None: "---", "left": ":---", "center": ":---:", "right": "---:"}
    lines = [row(block.header), "| " + " | ".join(delimiter[a] for a in block.aligns) + " |"]
    lines += [row(cells) for cells in block.rows]
    return lines


def _longest_run(text: str, char: str) -> int:
    runs = re.findall(re.escape(char) + "+", text)
    return max((len(r) for r in runs), default=0)


# -- inline ----------------------------------------------------------------------------------


def _inline(inline: list, *, line_start: bool = True, table: bool = False) -> str:
    """One block's inline items as Markdown (lines joined by ``\\n``).  A mark that
    continues across a link's start and lasts through all of its text stays open around
    it (``*foo [bar](/url)*``); any other closes before it, as the link's own marks close
    at its end."""
    out: list[str] = []
    open_marks: list[str] = []
    href: str | None = None
    #: How many marks were open outside the link being written.
    outer = 0
    at_start = line_start
    after_note = False

    def close_to(keep: frozenset) -> None:
        # Close from the innermost until what is open is all still wanted.
        while len(open_marks) > outer and not set(open_marks) <= keep:
            out.append(_DELIMITERS[open_marks.pop()])

    def close_all() -> None:
        nonlocal outer
        outer = 0
        close_to(frozenset())

    def close_link() -> None:
        nonlocal href, outer
        if href is not None:
            close_to(frozenset())
            out.append(f"]({_destination(href)})")
            href = None
            outer = 0

    items = _spaced(inline)
    for k, item in enumerate(items):
        item_href = item.href if isinstance(item, (m.Text, m.Image)) else None
        if item_href != href:
            close_link()
            if item_href is not None:
                # The marks every item of the link has may stay open around it.
                lasting = None
                n = k
                while n < len(items) and isinstance(items[n], (m.Text, m.Image)) and items[n].href == item_href:
                    marks = frozenset(items[n].marks) - {"code"} if isinstance(items[n], m.Text) else frozenset()
                    lasting = marks if lasting is None else lasting & marks
                    n += 1
                lasting = lasting or frozenset()
                while open_marks and not set(open_marks) <= lasting:
                    out.append(_DELIMITERS[open_marks.pop()])
                outer = len(open_marks)
                if out and out[-1].endswith("!") and not out[-1].endswith("\\!"):
                    out[-1] = out[-1][:-1] + "\\!"  # else ![ opens an image
                out.append("[")
                href = item_href
        if isinstance(item, m.Text):
            marks = frozenset(item.marks) - {"code"}
            close_to(marks)
            for mark in sorted(marks - set(open_marks), key=_order(items, k)):
                out.append(_DELIMITERS[mark])
                open_marks.append(mark)
            if "code" in item.marks:
                out.append(_code(item.text))
            else:
                text = _escape(item.text, table=table, after_note=after_note)
                if at_start:
                    text = _escape_line_start(text)
                out.append(text)
            at_start = at_start and not item.text.strip()
            after_note = False
            continue
        if href is None:
            close_all()
        else:
            close_to(frozenset())
        if isinstance(item, m.Html):
            out.append(item.text)
            at_start = False
        elif isinstance(item, m.Critic):
            out.append(item.text)
            at_start = False
        elif isinstance(item, m.Break):
            close_link()
            close_all()
            if _ends_block(items, k):
                # CommonMark has no hard break at the end of a block (``Title\\`` reads
                # back as a backslash): a trailing break -- a cover page's Shift+Enter --
                # is ``<br>``, which the reader takes back as a break.
                out.append("<br>")
                at_start = False
            else:
                out.append("\\\n")
                at_start = True
        elif isinstance(item, m.Image):
            title = f' "{_title(item.title)}"' if item.title else ""
            out.append(f"![{_escape(item.alt, table=table)}]({_destination(item.src)}{title})")
            at_start = False
        elif isinstance(item, m.NoteRef):
            out.append(f"[^{item.label}]")
            at_start = False
            after_note = True
            continue
        after_note = False
    close_link()
    close_all()
    return _edge_spaces("".join(out))


def _ends_block(items: list, k: int) -> bool:
    """Whether nothing but breaks, whitespace and comments follows ``items[k]``."""
    return all(isinstance(i, m.Break) or (isinstance(i, m.Text) and not i.text.strip())
               or (isinstance(i, m.Html) and i.text.startswith("<!--")) for i in items[k + 1:])


def _order(items: list, k: int):
    """Marks opened together nest by how long they last: the longest outermost."""
    def lasts(mark: str) -> int:
        n = k
        while n < len(items) and isinstance(items[n], m.Text) and mark in items[n].marks:
            n += 1
        return -n
    return lasts


def _spaced(inline: list) -> list:
    """Whitespace moved out of a mark where the mark opens or closes, so no delimiter
    touches it; between two pieces of text that share a mark it stays inside (``*foo
    **bar** baz*``: the spaces are emphasis)."""
    out: list = []
    items = list(inline)

    def neighbour(k: int) -> frozenset:
        if 0 <= k < len(items) and isinstance(items[k], m.Text) and items[k].text:
            return frozenset(items[k].marks) - {"code"}
        return frozenset()

    for current, item in enumerate(items):
        if not (isinstance(item, m.Text) and item.marks and item.text) or "code" in item.marks:
            out.append(item)
            continue
        marks = frozenset(item.marks)
        before, after = neighbour(current - 1), neighbour(current + 1)
        core = item.text.strip(" \t")
        if not core:
            out.append(m.Text(item.text, marks & before & after, item.href))
            continue
        lead = item.text[: len(item.text) - len(item.text.lstrip(" \t"))]
        trail = item.text[len(item.text.rstrip(" \t")):]
        if lead:
            out.append(m.Text(lead, marks & before, item.href))
        out.append(m.Text(core, marks, item.href))
        if trail:
            out.append(m.Text(trail, marks & after, item.href))
    return out


def _code(text: str) -> str:
    fence = "`" * (_longest_run(text, "`") + 1)
    pad = text.startswith("`") or text.endswith("`") or (
        text.startswith(" ") and text.endswith(" ") and text.strip(" ") != "")
    return f"{fence} {text} {fence}" if pad else f"{fence}{text}{fence}"


def _escape(text: str, *, table: bool = False, after_note: bool = False) -> str:
    out = []
    for k, char in enumerate(text):
        if char in _ALWAYS or (table and char == "|"):
            out.append("\\" + char)
        elif char == "&" and _ENTITY.match(text, k):
            out.append("\\&")
        elif char == "{" and text[k + 1:k + 3] in _CRITIC_OPEN:
            out.append("\\{")
        elif char == ":" and k == 0 and after_note:
            out.append("\\:")
        elif char == "\n":
            out.append(" ")
        else:
            out.append(char)
    return "".join(out)


def _escape_line_start(text: str) -> str:
    stripped = text.lstrip(" \t")
    match = _LINE_START.match(stripped)
    if match is None:
        return stripped
    if match.group(1):
        return "\\" + stripped
    return match.group(2) + "\\" + match.group(3) + stripped[match.end():]


def _destination(url: str) -> str:
    if re.search(r"[\s()<>]", url):
        return "<" + url.replace("<", "%3C").replace(">", "%3E") + ">"
    return url


def _title(title: str) -> str:
    return title.replace("\\", "\\\\").replace('"', '\\"')


#: Whitespace a Markdown parser trims from a line's ends (JavaScript's ``trim``) that is not
#: a space or a tab: written as a character reference there, so it stays.
_TRIMMED = frozenset("\xa0\u1680\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a"
                     "\u2028\u2029\u202f\u205f\u3000\ufeff")


def _edge_spaces(text: str) -> str:
    lines = text.split("\n")
    for k, line in enumerate(lines):
        start = 0
        while start < len(line) and (line[start] in _TRIMMED or line[start] in " \t"):
            start += 1
        end = len(line)
        while end > start and (line[end - 1] in _TRIMMED or line[end - 1] in " \t"):
            end -= 1
        if any(c in _TRIMMED for c in line[:start] + line[end:]):
            lead = "".join(f"&#x{ord(c):X};" if c in _TRIMMED else c for c in line[:start])
            trail = "".join(f"&#x{ord(c):X};" if c in _TRIMMED else c for c in line[end:])
            lines[k] = lead + line[start:end] + trail
    return "\n".join(lines)


# -- the self-check --------------------------------------------------------------------------


def _settle(block, *, heading: bool = False) -> None:
    _settle_inline(block.inline, heading=heading)


def _settle_inline(inline: list, *, heading: bool = False, cell: bool = False) -> None:
    """Make ``inline`` (in place) something Markdown reads back as meant: whitespace
    moved out of a mark where it opens or closes (:func:`_spaced`), then strikethrough
    and then emphasis dropped from the block if its marks would not survive."""
    inline[:] = _spaced(inline)
    for dropped in ((), ("strike",), ("strike", "strong", "em")):
        if dropped:
            inline[:] = [m.Text(i.text, frozenset(i.marks) - set(dropped), i.href) if isinstance(i, m.Text) else i
                         for i in inline]
        if _reads_back(inline, heading=heading, cell=cell):
            return


def _reads_back(inline: list, *, heading: bool, cell: bool) -> bool:
    if not any(isinstance(i, m.Text) and i.marks - {"code"} for i in inline):
        return True
    text = _inline(inline, line_start=not (heading or cell), table=cell)
    # The references' definitions are block-level: say they exist, as the whole text will.
    env: dict = {"footnotes": {"refs": {":" + i.label: -1 for i in inline if isinstance(i, m.NoteRef)},
                               "list": {}}}
    tokens = parser().parseInline(text, env)
    children = tokens[0].children if tokens else []
    return _parsed_inline(children or [], env) == model_inline(inline, heading=heading, cell=cell)


__all__ = ["render", "normalise_inline"]
