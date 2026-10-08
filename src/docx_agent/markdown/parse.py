"""Markdown in: markdown-it-py's tokens turned into one normalised shape, and the same shape
computed from the model -- the reverse-parse check.

The dialect is ROADMAP.md's: CommonMark plus GFM tables and strikethrough, plus footnotes
(``mdit-py-plugins``).  markdown-it-py is used behind this module only, so it can be
swapped (ROADMAP.md, "Dialect and parser").

**The normalised AST** is nested tuples, equal when two documents mean the same thing:

* blocks -- ``("comment", text)`` for an HTML block that is only a comment,
  ``("html", text)``, ``("p", inline)``, ``("h", level, inline)``, ``("code", text)``,
  ``("quote", blocks)``, ``("list", ordered, start, items)`` (an item is a tuple of blocks),
  ``("table", header, rows, aligns)``, ``("hr",)``, ``("footnote", label, blocks)``;
* inline -- ``("t", text, marks, href)`` with adjacent alike texts merged,
  ``("html", text)``, ``("br",)``, ``("img", alt, src, title, href)``, ``("fn", label)``.

What normalisation forgets, because Markdown does not keep it: list markers and
delimiters, emphasis delimiters, whitespace at the ends of a block or a line, a paragraph
that is empty, and the order footnote definitions are written in (markdown-it lists them
by first reference, as the reader writes them).
"""

from __future__ import annotations

import re
from functools import lru_cache

from markdown_it import MarkdownIt
from mdit_py_plugins.footnote import footnote_plugin

from . import model as m

_COMMENT = re.compile(r"<!--(.*?)-->", re.DOTALL)
_ALIGN = re.compile(r"text-align:\s*(left|center|right)")


@lru_cache(maxsize=1)
def parser() -> MarkdownIt:
    """The dialect: CommonMark, GFM tables and strikethrough, footnotes."""
    return MarkdownIt("commonmark").enable(["table", "strikethrough"]).use(footnote_plugin)


# -- Markdown -> AST -------------------------------------------------------------------------


def parse_ast(text: str) -> tuple:
    """``text``'s normalised AST."""
    env: dict = {}
    tokens = parser().parse(text, env)
    blocks, _ = _blocks(tokens, 0, env, None)
    return tuple(blocks)


def _blocks(tokens, i: int, env: dict, until: str | None) -> tuple[list, int]:
    out: list = []
    while i < len(tokens):
        token = tokens[i]
        kind = token.type
        if until is not None and kind == until:
            return out, i + 1
        if kind == "html_block":
            out.extend(_html_block(token.content))
            i += 1
        elif kind == "paragraph_open":
            inline = _inline(tokens[i + 1].children or [], env)
            if inline:
                out.append(("p", inline))
            i += 3
        elif kind == "heading_open":
            out.append(("h", int(token.tag[1]), _inline(tokens[i + 1].children or [], env)))
            i += 3
        elif kind in ("fence", "code_block"):
            out.append(("code", token.content))
            i += 1
        elif kind == "blockquote_open":
            inner, i = _blocks(tokens, i + 1, env, "blockquote_close")
            out.append(("quote", tuple(inner)))
        elif kind in ("bullet_list_open", "ordered_list_open"):
            ordered = kind == "ordered_list_open"
            start = int(token.attrGet("start") or 1) if ordered else 1
            close = kind.replace("open", "close")
            items = []
            i += 1
            while tokens[i].type != close:
                inner, i = _blocks(tokens, i + 1, env, "list_item_close")
                items.append(tuple(inner))
            out.append(("list", ordered, start, tuple(items)))
            i += 1
        elif kind == "table_open":
            table, i = _table(tokens, i + 1, env)
            out.append(table)
        elif kind == "hr":
            out.append(("hr",))
            i += 1
        elif kind == "footnote_block_open":
            i += 1
            while tokens[i].type != "footnote_block_close":
                label = env["footnotes"]["list"][tokens[i].meta["id"]]["label"]
                inner, i = _blocks(tokens, i + 1, env, "footnote_close")
                out.append(("footnote", label, tuple(inner)))
            i += 1
        else:
            i += 1  # footnote_anchor and closers of what was skipped
    return out, i


def _html_block(content: str) -> list:
    stripped = content.strip()
    comments = _COMMENT.findall(stripped)
    if comments and _COMMENT.sub("", stripped).strip() == "":
        return [("comment", c.strip()) for c in comments]
    return [("html", content.rstrip("\n"))]


def _table(tokens, i: int, env: dict) -> tuple[tuple, int]:
    header: list = []
    rows: list = []
    aligns: list = []
    current: list | None = None
    while tokens[i].type != "table_close":
        token = tokens[i]
        if token.type == "tr_open":
            current = []
        elif token.type == "tr_close":
            (rows if header else header).append(tuple(current))
            current = None
        elif token.type in ("th_open", "td_open"):
            if token.type == "th_open":
                match = _ALIGN.search(token.attrGet("style") or "")
                aligns.append(match.group(1) if match else None)
            current.append(_inline(tokens[i + 1].children or [], env))
            i += 2
        i += 1
    return ("table", header[0] if header else (), tuple(rows), tuple(aligns)), i + 1


def _inline(children, env: dict) -> tuple:
    items: list = []
    marks: list[str] = []
    hrefs: list[str] = []
    for token in children:
        kind = token.type
        href = hrefs[-1] if hrefs else None
        if kind == "text":
            items.append(("t", token.content, frozenset(marks), href))
        elif kind == "code_inline":
            items.append(("t", token.content, frozenset(marks + ["code"]), href))
        elif kind in ("softbreak",):
            items.append(("t", "\n", frozenset(marks), href))
        elif kind == "hardbreak":
            items.append(("br",))
        elif kind == "html_inline":
            items.append(("html", token.content))
        elif kind == "strong_open":
            marks.append("strong")
        elif kind == "em_open":
            marks.append("em")
        elif kind == "s_open":
            marks.append("strike")
        elif kind in ("strong_close", "em_close", "s_close"):
            marks.pop()
        elif kind == "link_open":
            hrefs.append(token.attrGet("href"))
        elif kind == "link_close":
            hrefs.pop()
        elif kind == "image":
            alt = "".join(c.content for c in token.children or [] if c.type in ("text", "text_special", "code_inline"))
            items.append(("img", alt, token.attrGet("src"), token.attrGet("title"), href))
        elif kind == "footnote_ref":
            items.append(("fn", env["footnotes"]["list"][token.meta["id"]]["label"]))
    return normalise_inline(items)


# -- model -> AST ----------------------------------------------------------------------------


def model_ast(blocks: list) -> tuple:
    """The normalised AST the model means: what :func:`parse_ast` must read back from
    :func:`docx_agent.markdown.render.render` of it.  Footnote definitions come last, in the
    order of their first reference, and one nothing references (outside raw HTML) is not
    there -- markdown-it lists them so."""
    body = [b for b in blocks if not isinstance(b, m.FootnoteDef)]
    notes = {b.label: b for b in blocks if isinstance(b, m.FootnoteDef)}
    out = [node for block in body for node in _model_block(block)]
    order: list[str] = []
    _references(out, order)
    done: list = []
    seen: set[str] = set()
    k = 0
    while k < len(order):
        label = order[k]
        k += 1
        if label in seen or label not in notes:
            continue
        seen.add(label)
        inner = model_ast(notes[label].blocks)
        _references(inner, order)
        done.append(("footnote", label, inner))
    return tuple(out + done)


def _references(node, order: list[str]) -> None:
    """Footnote labels referenced in ``node`` (an AST), in order, appended to ``order``."""
    if isinstance(node, tuple):
        if len(node) == 2 and node[0] == "fn" and isinstance(node[1], str):
            order.append(node[1])
            return
        for child in node:
            _references(child, order)
    elif isinstance(node, list):
        for child in node:
            _references(child, order)


def _model_block(block) -> list:
    if isinstance(block, m.Comment):
        return [("comment", block.text.strip())]
    if isinstance(block, m.Paragraph):
        inline = model_inline(block.inline)
        return [("p", inline)] if inline else []
    if isinstance(block, m.Heading):
        return [("h", block.level, model_inline(block.inline, heading=True))]
    if isinstance(block, m.CodeBlock):
        return [("code", block.text if block.text.endswith("\n") or not block.text else block.text + "\n")]
    if isinstance(block, m.Quote):
        return [("quote", model_ast(block.blocks))]
    if isinstance(block, m.ListBlock):
        return [("list", block.ordered, block.start if block.ordered else 1,
                 tuple(model_ast(item) for item in block.items))]
    if isinstance(block, m.Table):
        return [("table", tuple(model_inline(c, cell=True) for c in block.header),
                 tuple(tuple(model_inline(c, cell=True) for c in row) for row in block.rows),
                 tuple(block.aligns))]
    if isinstance(block, m.HtmlBlock):
        return [("html", block.text.rstrip("\n"))]
    if isinstance(block, m.Rule):
        return [("hr",)]
    raise TypeError(f"not a block: {block!r}")


def model_inline(inline: list, *, heading: bool = False, cell: bool = False) -> tuple:
    items: list = []
    for item in inline:
        if isinstance(item, m.Text):
            text = item.text
            if heading or cell:
                text = text.replace("\n", " ")
            items.append(("t", text, frozenset(item.marks), item.href))
        elif isinstance(item, m.Critic):
            items.append(("t", item.text, frozenset(), None))
        elif isinstance(item, m.Html):
            items.append(("html", item.text))
        elif isinstance(item, m.Break):
            items.append(("br",))
        elif isinstance(item, m.Image):
            items.append(("img", item.alt, item.src, item.title, item.href))
        elif isinstance(item, m.NoteRef):
            items.append(("fn", item.label))
    return normalise_inline(items)


def normalise_inline(items: list) -> tuple:
    """Merge alike texts, drop empty ones, and forget whitespace at the ends of the block
    and around hard breaks (Markdown keeps none of it).  Tabs count as spaces."""
    merged: list = []
    for item in items:
        if item[0] == "t":
            text = item[1].replace("\t", " ")
            if merged and merged[-1][0] == "t" and merged[-1][2:] == item[2:]:
                merged[-1] = ("t", merged[-1][1] + text) + item[2:]
                continue
            merged.append(("t", text) + item[2:])
        else:
            merged.append(item)
    # Whitespace at the ends of lines.
    for k, item in enumerate(merged):
        if item[0] != "t":
            continue
        text = item[1]
        if k == 0 or merged[k - 1][0] == "br":
            text = text.lstrip(" \n")
        if k == len(merged) - 1 or merged[k + 1][0] == "br":
            text = text.rstrip(" \n")
        merged[k] = ("t", text) + item[2:]
    out = [item for item in merged if item[0] != "t" or item[1]]
    # Dropping an empty text can make two texts adjacent: merge again.
    final: list = []
    for item in out:
        if final and item[0] == "t" and final[-1][0] == "t" and final[-1][2:] == item[2:]:
            final[-1] = ("t", final[-1][1] + item[1]) + item[2:]
        else:
            final.append(item)
    while final and final[-1][0] == "br":
        final.pop()
    return tuple(final)


def differences(expected: tuple, actual: tuple, path: str = "") -> list[str]:
    """Where two ASTs differ, as readable paths (empty when they agree)."""
    if expected == actual:
        return []
    if isinstance(expected, tuple) and isinstance(actual, tuple) and len(expected) == len(actual):
        out = []
        for k, (a, b) in enumerate(zip(expected, actual)):
            out += differences(a, b, f"{path}/{k}")
        return out or [f"{path}: {expected!r} != {actual!r}"]
    return [f"{path}: expected {expected!r}, read {actual!r}"]


# -- Markdown -> model (what insert_markdown writes) -----------------------------------------


def parse_model(text: str) -> list:
    """``text`` as model blocks (:mod:`.model`): what ``insert_markdown`` writes.  Footnote
    definitions come last, as :class:`~.model.FootnoteDef` in the order of their first
    reference (markdown-it keeps only those referenced); a soft line break is a space in
    its :class:`~.model.Text`, as Markdown means it; an HTML block that is only comments is
    a :class:`~.model.Comment` per comment."""
    env: dict = {}
    tokens = parser().parse(text, env)
    blocks, _ = _model_blocks(tokens, 0, env, None)
    return blocks


def _model_blocks(tokens, i: int, env: dict, until: str | None) -> tuple[list, int]:
    out: list = []
    while i < len(tokens):
        token = tokens[i]
        kind = token.type
        if until is not None and kind == until:
            return out, i + 1
        if kind == "html_block":
            stripped = token.content.strip()
            comments = _COMMENT.findall(stripped)
            if comments and _COMMENT.sub("", stripped).strip() == "":
                out += [m.Comment(c.strip()) for c in comments]
            else:
                out.append(m.HtmlBlock(token.content))
            i += 1
        elif kind == "paragraph_open":
            out.append(m.Paragraph(_model_inline(tokens[i + 1].children or [], env)))
            i += 3
        elif kind == "heading_open":
            out.append(m.Heading(int(token.tag[1]), _model_inline(tokens[i + 1].children or [], env)))
            i += 3
        elif kind in ("fence", "code_block"):
            out.append(m.CodeBlock(token.content))
            i += 1
        elif kind == "blockquote_open":
            inner, i = _model_blocks(tokens, i + 1, env, "blockquote_close")
            out.append(m.Quote(inner))
        elif kind in ("bullet_list_open", "ordered_list_open"):
            ordered = kind == "ordered_list_open"
            start = int(token.attrGet("start") or 1) if ordered else 1
            close = kind.replace("open", "close")
            items = []
            i += 1
            while tokens[i].type != close:
                inner, i = _model_blocks(tokens, i + 1, env, "list_item_close")
                items.append(inner)
            out.append(m.ListBlock(ordered, start, items, alternate=token.markup in ("*", "+", ")")))
            i += 1
        elif kind == "table_open":
            out.append(_model_table(tokens, i + 1, env))
            while tokens[i].type != "table_close":
                i += 1
            i += 1
        elif kind == "hr":
            out.append(m.Rule())
            i += 1
        elif kind == "footnote_block_open":
            i += 1
            while tokens[i].type != "footnote_block_close":
                label = env["footnotes"]["list"][tokens[i].meta["id"]]["label"]
                inner, i = _model_blocks(tokens, i + 1, env, "footnote_close")
                out.append(m.FootnoteDef(label, inner))
            i += 1
        else:
            i += 1
    return out, i


def _model_table(tokens, i: int, env: dict) -> m.Table:
    rows: list = []
    aligns: list = []
    current: list | None = None
    header = True
    while tokens[i].type != "table_close":
        token = tokens[i]
        if token.type == "tr_open":
            current = []
        elif token.type == "tr_close":
            rows.append(current)
            current = None
        elif token.type == "tbody_open":
            header = False
        elif token.type in ("th_open", "td_open"):
            if token.type == "th_open" and header:
                match = _ALIGN.search(token.attrGet("style") or "")
                aligns.append(match.group(1) if match else None)
            current.append(_model_inline(tokens[i + 1].children or [], env))
            i += 2
        i += 1
    return m.Table(rows[0] if rows else [], rows[1:], aligns)


def _model_inline(children, env: dict) -> list:
    items: list = []
    marks: list[str] = []
    hrefs: list[str] = []
    for token in children:
        kind = token.type
        href = hrefs[-1] if hrefs else None
        if kind == "text":
            items.append(m.Text(token.content, frozenset(marks), href))
        elif kind == "code_inline":
            items.append(m.Text(token.content, frozenset(marks + ["code"]), href))
        elif kind == "softbreak":
            items.append(m.Text(" ", frozenset(marks), href))
        elif kind == "hardbreak":
            items.append(m.Break())
        elif kind == "html_inline":
            items.append(m.Html(token.content))
        elif kind == "strong_open":
            marks.append("strong")
        elif kind == "em_open":
            marks.append("em")
        elif kind == "s_open":
            marks.append("strike")
        elif kind in ("strong_close", "em_close", "s_close"):
            marks.pop()
        elif kind == "link_open":
            hrefs.append(token.attrGet("href") or "")
        elif kind == "link_close":
            hrefs.pop()
        elif kind == "image":
            alt = "".join(c.content for c in token.children or [] if c.type in ("text", "text_special", "code_inline"))
            items.append(m.Image(alt, token.attrGet("src") or "", token.attrGet("title"), href))
        elif kind == "footnote_ref":
            items.append(m.NoteRef(env["footnotes"]["list"][token.meta["id"]]["label"]))
    return items


# -- the round trip's normalisation ------------------------------------------------------------


def roundtrip_ast(text: str) -> tuple:
    """``text``'s AST as the Markdown round trip compares it (ROADMAP.md, "Guarantees"):
    :func:`parse_ast`, and further forgetting what a Word document cannot keep or names
    otherwise -- HTML comments (``insert_markdown`` drops them), a soft line break is a space (Word has none), footnote labels are their
    order of first reference (``to_markdown`` writes ``fn:<id>``), a picture is its alt text
    (its address becomes the drawing's id, its title the drawing's name), and a left-aligned
    table column is an unaligned one (Word's left is its default)."""
    labels: dict[str, str] = {}
    tree = _without_comments(parse_ast(text))
    _labels(tree, labels)
    return _normalise(tree, labels)


def _without_comments(node):
    """The AST without HTML comments, which ``insert_markdown`` drops (``to_markdown``'s ids
    are comments)."""
    if not isinstance(node, tuple):
        return node
    kept = tuple(_without_comments(child) for child in node
                 if not (isinstance(child, tuple) and child and (child[0] == "comment" or (
                     child[0] == "html" and len(child) == 2 and isinstance(child[1], str)
                     and child[1].startswith("<!--")))))
    if kept and all(isinstance(c, tuple) and c and c[0] in ("t", "html", "br", "img", "fn") for c in kept):
        return normalise_inline(list(kept))
    return kept


def _labels(node, labels: dict[str, str]) -> None:
    if isinstance(node, tuple):
        if len(node) == 2 and node[0] == "fn" and isinstance(node[1], str):
            labels.setdefault(node[1], f"n{len(labels) + 1}")
            return
        for child in node:
            _labels(child, labels)


def _normalise(node, labels: dict[str, str]):
    if not isinstance(node, tuple) or not node:
        return node
    head = node[0]
    if head == "t" and len(node) == 4 and isinstance(node[1], str):
        return ("t", node[1].replace("\n", " "), node[2], node[3])
    if head == "fn" and len(node) == 2 and isinstance(node[1], str):
        return ("fn", labels.get(node[1], node[1]))
    if head == "img" and len(node) == 5:
        return ("img", node[1], node[4])
    if head == "footnote" and len(node) == 3 and isinstance(node[1], str):
        return ("footnote", labels.get(node[1], node[1]), _normalise(node[2], labels))
    if head == "table" and len(node) == 4:
        aligns = tuple(None if a == "left" else a for a in node[3])
        return ("table", _normalise(node[1], labels), _normalise(node[2], labels), aligns)
    out = tuple(_normalise(child, labels) for child in node)
    # Re-merge texts a replaced soft break made alike.
    if out and all(isinstance(c, tuple) for c in out) and any(c and c[0] == "t" for c in out):
        return normalise_inline(list(out)) if all(c and isinstance(c[0], str) and c[0] in
                                                  ("t", "html", "br", "img", "fn") for c in out) else out
    return out
