"""Raw HTML for what GFM cannot say: a table with merged cells, nested tables or more than
one paragraph in a cell (ROADMAP.md, "The Markdown layer": the fallback).

The HTML stays one Markdown HTML block: no blank line inside it, every newline in text
written as ``&#10;``.  Ids ride along -- ``data-id`` on the table and each cell, the
paragraphs' comments inside the cells -- so the agent can still address everything.
"""

from __future__ import annotations

from html import escape
from typing import TYPE_CHECKING

from . import model as m

if TYPE_CHECKING:  # pragma: no cover
    from .read import Options, Reader, TableRecord

_TAGS = {"strong": "strong", "em": "em", "code": "code", "strike": "del"}


def table_html(reader: "Reader", table: "TableRecord", options: "Options") -> str:
    from .read import build

    lines = [f'<table data-id="{escape(table.id)}">' if options.ids else "<table>"]
    for k, line in enumerate(table.rows):
        cells = []
        tag = "th" if k < table.header_rows else "td"
        for cell in line:
            attributes = f' data-id="{escape(cell.id)}"' if options.ids else ""
            if cell.row_span > 1:
                attributes += f' rowspan="{cell.row_span}"'
            if cell.column_span > 1:
                attributes += f' colspan="{cell.column_span}"'
            cells.append(f"<{tag}{attributes}>{blocks_html(build(reader, cell.records, options))}</{tag}>")
        lines.append("<tr>" + "".join(cells) + "</tr>")
    lines.append("</table>")
    return "\n".join(lines)


def blocks_html(blocks: list) -> str:
    out = []
    for block in blocks:
        if isinstance(block, m.Comment):
            out.append(f"<!-- {block.text} -->")
        elif isinstance(block, m.Paragraph):
            out.append(f"<p>{inline_html(block.inline)}</p>")
        elif isinstance(block, m.Heading):
            out.append(f"<h{block.level}>{inline_html(block.inline)}</h{block.level}>")
        elif isinstance(block, m.CodeBlock):
            out.append(f"<pre><code>{_text(block.text)}</code></pre>")
        elif isinstance(block, m.Quote):
            out.append(f"<blockquote>{blocks_html(block.blocks)}</blockquote>")
        elif isinstance(block, m.ListBlock):
            tag = "ol" if block.ordered else "ul"
            start = f' start="{block.start}"' if block.ordered and block.start != 1 else ""
            items = "".join(f"<li>{blocks_html(item)}</li>" for item in block.items)
            out.append(f"<{tag}{start}>{items}</{tag}>")
        elif isinstance(block, m.Table):
            rows = [block.header] + block.rows
            body = "".join("<tr>" + "".join(f"<td>{inline_html(c)}</td>" for c in row) + "</tr>" for row in rows)
            out.append(f"<table>{body}</table>")
        elif isinstance(block, m.HtmlBlock):
            out.append(block.text.replace("\n", ""))
        elif isinstance(block, m.Rule):
            out.append("<hr>")
    return "".join(out)


def inline_html(inline: list) -> str:
    out = []
    for item in inline:
        if isinstance(item, m.Text):
            text = _text(item.text)
            for mark in ("code", "em", "strong", "strike"):
                if mark in item.marks:
                    text = f"<{_TAGS[mark]}>{text}</{_TAGS[mark]}>"
            if item.href is not None:
                text = f'<a href="{escape(item.href)}">{text}</a>'
            out.append(text)
        elif isinstance(item, m.Html):
            out.append(item.text)
        elif isinstance(item, m.Critic):
            out.append(_text(item.text))
        elif isinstance(item, m.Break):
            out.append("<br>")
        elif isinstance(item, m.Image):
            title = f' title="{escape(item.title)}"' if item.title else ""
            image = f'<img src="{escape(item.src)}" alt="{escape(item.alt)}"{title}>'
            out.append(f'<a href="{escape(item.href)}">{image}</a>' if item.href else image)
        elif isinstance(item, m.NoteRef):
            out.append(f"<sup>[^{escape(item.label)}]</sup>")
    return "".join(out)


def _text(text: str) -> str:
    return escape(text, quote=False).replace("\n", "&#10;").replace("\v", "<br>")
