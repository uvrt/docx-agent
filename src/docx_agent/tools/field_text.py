"""A dataset's text written by reference (round 45, R65, R66): where each field went.

**One field as Markdown.**  ``word_insert_markdown``'s ``from`` names a table of rows
(``d:notes``, ``d:notes/#3:#9``) -- or one field of one row (``d:notes/#12/body``,
``d:notes/Doors/body``, ``t:checks/Alpha/note``): that value's text, written as the
``markdown`` argument is (its paragraphs, its lists as lists in the template's styles,
nested items at their level, its emphasis), verbatim -- nothing retyped.  ``from`` is its
one name (round 47: no alias).  A value that is not text -- a number, a date, a logical
value -- is refused with the way out (a figure goes into text as ``{{d:...}}``) and the
text fields of its row; an unknown dataset, row or field is refused as a reference to it
is, naming what there is.

**Several fields** go in one call as references in ``markdown``, each standing alone
between blank lines, with whatever is wanted around them (``## Doors\\n\\n{{d:notes/#12/body}}
\\n\\n## Roof\\n\\n{{d:notes/#13/body}}``): the dispatcher writes each value in its place
before the Markdown is read (:mod:`ooxml_edit.tools.textfacts`), so it is the same Markdown.

**Where each went.**  The result's ``data.facts`` names, for a value that made whole blocks
(``from``, or a reference standing alone), the blocks it made -- ``{"from":
"d:notes/#12/body", "blocks": ["p:...", ...], "chars": 1840, "sources": [...]}`` -- and no
longer echoes a long or several-line text (rule 4: the text is the caller's own; ``chars``
says how much was written).  The blocks are found by reading each value's Markdown alone for the blocks it
writes (:func:`shape`) and lining that up with the blocks the call made; when they do not
line up (a table written beside another table gets a separator paragraph), the facts say
so with a note rather than guess.  A short value written inside a sentence is listed as
before, with its text.  Each paragraph a field made keeps the field as its source in the
file (:meth:`docx_agent.Document.keep_text_facts`, as a figure written by reference keeps
its own), so a check of numbers reads the numbers in it as the data's.

**word_set_text** writes a field as a paragraph's text the same way (``text:
"{{d:notes/#2/summary}}"``); its facts name the paragraphs each value went into
(``blocks``).  **word_edit_table**'s ``set_cells`` takes one value per cell -- ``from:
"d:scores/#12/rating:0.0"`` (R66), with a number format or a date pattern as a reference
takes -- and references in a cell's ``text``; each fact names its cell (``cell``).
"""

from __future__ import annotations

import re
from typing import Any, Iterable, Mapping, Sequence

from ooxml_edit.tools import ToolError
from ooxml_edit.tools import textfacts
from ooxml_edit.tools.datasets import DATASET_PREFIX, row_range

from ..markdown import model as m

#: A value written by reference longer than this (or over several lines) is not echoed in
#: the result: ``chars`` says how long it was, ``blocks`` where it went.
ECHO_CHARS = 200
#: Column types whose values are not text.
_NOT_TEXT = {"number", "percent", "date", "datetime", "time", "bool"}
#: A source address after its document handle: ``d2/t:x`` -> ``t:x``.
_HANDLE = re.compile(r"^d[0-9]+/")


# -- one field of one row --------------------------------------------------------------------


def names_one_value(address: Any) -> bool:
    """Whether a ``from`` names one field of one row rather than rows: a dataset address
    past its id that is no run of rows (``d:notes/#12/body``, ``d:notes/Doors/body``)."""
    if not isinstance(address, str) or not address.startswith(DATASET_PREFIX):
        return False
    body = textfacts.strip_format(address)
    rest = body[len(DATASET_PREFIX):]
    return "/" in rest and row_range(body) is None


def might_name_one_value(address: Any) -> bool:
    """Whether a ``from`` that is no dataset (``t:checks/Alpha/note``, ``d2/t:...``) has the
    parts of one value -- a table, a row and a column -- so that a refusal as rows is worth
    reading it as one value."""
    if not isinstance(address, str) or address.startswith(DATASET_PREFIX):
        return False
    _source, segments, _fmt = textfacts.split(address)
    return len(segments) >= 2


def _rows_and_cell(call: Any, address: str, field: str) -> tuple[Any, int, int]:
    """The rows a one-value address reads and the value's ``(row, column)``, as the
    dispatcher's reference resolution picks them (``Call.fact``)."""
    source, segments, _fmt = textfacts.split(address)
    selected = False
    if source.startswith(DATASET_PREFIX):
        rows = call.source(source, field=field)
    elif len(segments) >= 2:
        segment, quoted = segments[0]
        numbered = re.fullmatch(r"#[1-9][0-9]*", segment)
        raw = textfacts.quote(segment) if quoted or not numbered else segment
        rows = call.source(f"{source}/{raw}", field=field)
        segments, selected = segments[1:], True
    else:
        rows = call.source(source, field=field)
    row, column = textfacts.pick(rows, segments, address, field=field, locale=call.locale,
                                 row_selected=selected)
    return rows, row, column


def is_text(value: Any, kind: str | None) -> bool:
    return isinstance(value, str) and kind not in _NOT_TEXT


def text_field(call: Any, address: str, field: str) -> tuple[str, dict[str, Any]]:
    """The text of the one text field ``address`` names, and the record of it
    (``Call.fact``'s).  A value that is not text is refused, naming the row's text fields
    and the way to write a figure; a blank, an unknown dataset, row or field as a
    reference's are."""
    try:
        rows, row, column = _rows_and_cell(call, address, field)
    except ToolError as error:
        with_rows(call, address, field, error)
        raise
    value = rows.dataset.rows[row][column]
    kind = rows.type(column)
    shown = _HANDLE.sub("", textfacts.strip_format(address))
    if value is not None and value != "" and not is_text(value, kind):
        columns = list(rows.dataset.columns)
        texts = [c for k, c in enumerate(columns)
                 if is_text(rows.dataset.rows[row][k], rows.type(k))]
        what = kind or _kind(value)
        prefix = shown.rsplit("/", 1)[0]
        raise ToolError("invalid_arguments", f"{shown} is a {what} value ({value!r}), not text: Markdown comes "
                        f"from a text field" + (f" ({', '.join(texts[:6])})" if texts else "") +
                        f"; a figure goes into the markdown as {{{{{shown}}}}}", field=field,
                        valid_options=[f"{prefix}/{textfacts.quote(c)}" for c in texts[:10]]
                        + [f"markdown: '{{{{{shown}}}}}'"])
    text, record = call.fact(address, field=field)
    return value_markdown(rows, row, column) or text, record


def value_markdown(rows: Any, row: int, column: int) -> str | None:
    """Round 48: the Markdown a source holds for one text value -- a workbook cell's bold
    and italic runs and its list lines (ooxml-edit's ``Dataset.markdown``) -- ``None`` when
    the value is its text."""
    found = getattr(rows.dataset, "markdown_at", None)
    return found(row, column) if found is not None else None


def reference_markdown(call: Any, fact: Any) -> str | None:
    """Round 48: the Markdown of the value a reference wrote (``fact``, the dispatcher's),
    read again from its source; ``None`` when it is its text or cannot be read again."""
    address = fact.address
    doc = fact.record.get("doc")
    if not address.startswith(DATASET_PREFIX) and isinstance(doc, str) and _HANDLE.match(doc + "/"):
        address = f"{doc}/{address}"
    try:
        rows, row, column = _rows_and_cell(call, address, fact.field)
    except ToolError:
        return None
    if not is_text(rows.dataset.rows[row][column], rows.type(column)):
        return None
    return value_markdown(rows, row, column)


def with_markdown(call: Any, markdown: str, facts: Sequence[Any]) -> tuple[str, list[Any]]:
    """Round 48: ``markdown`` (the text as the handler got it, every reference written as its
    value's text) with each value that has Markdown (:func:`reference_markdown`) written as
    it: a reference standing alone between blank lines as its Markdown -- its lists and
    emphasis, as ``from`` writes one field -- and one inside a sentence as its text on one
    line, escaped, so that its lines and characters stay text there.  The facts, their
    places and texts in the new Markdown.  A value without Markdown is left as it is."""
    from ooxml_edit.markdown import Run, inline_markdown

    out: list[str] = []
    placed: list[Any] = []
    at = shift = 0
    for fact in sorted(facts, key=lambda f: f.start):
        found = reference_markdown(call, fact)
        if found is None:
            placed.append(textfacts.Fact(fact.field, fact.start + shift, fact.end + shift, fact.text,
                                         fact.address, fact.record))
            continue
        if not _alone(markdown, fact.start, fact.end):
            line_start = not markdown[:fact.start].rpartition("\n")[2].strip()
            one_line = " ".join(line.strip() for line in fact.text.splitlines() if line.strip())
            found = inline_markdown([Run(one_line)], line_start=line_start).replace("|", "\\|")
        out.append(markdown[at:fact.start])
        out.append(found)
        at = fact.end
        start = fact.start + shift
        placed.append(textfacts.Fact(fact.field, start, start + len(found), found, fact.address,
                                     fact.record))
        shift += len(found) - (fact.end - fact.start)
    if not out:
        return markdown, list(facts)
    out.append(markdown[at:])
    return "".join(out), placed


def _kind(value: Any) -> str:
    if isinstance(value, bool):
        return "logical"
    if isinstance(value, (int, float)):
        return "number"
    return "date" if hasattr(value, "isoformat") else type(value).__name__


def with_rows(call: Any, address: str, field: str, error: ToolError) -> None:
    """A refusal of a row or column a dataset lacks given the way out (rule 2): whole
    addresses -- the first and last rows with the same field, the row's fields -- in place
    of none or of bare column names."""
    if error.code != "not_found" or not address.startswith(DATASET_PREFIX):
        return
    source, segments, _fmt = textfacts.split(address)
    if len(segments) != 2:
        return
    try:
        dataset = call.source(source, field=field).dataset
    except ToolError:
        return
    row, column = textfacts.quote(segments[0][0]) if segments[0][1] else segments[0][0], segments[1][0]
    count = len(dataset.rows)
    if not error.valid_options and re.fullmatch(r"#[1-9][0-9]*", row) and count:
        error.valid_options = list(dict.fromkeys([f"{source}/#1/{textfacts.quote(column)}",
                                                  f"{source}/#{count}/{textfacts.quote(column)}"]))
    elif error.valid_options and set(error.valid_options) <= set(dataset.columns):
        error.valid_options = [f"{source}/{row}/{textfacts.quote(c)}" for c in error.valid_options]


# -- the blocks Markdown writes --------------------------------------------------------------


def shape(blocks: Sequence[Any]) -> list[str]:
    """The top-level blocks the Markdown writer (:class:`docx_agent.markdown.write.Writer`)
    makes of model ``blocks``, as ``p`` (a paragraph) or ``t`` (a table), in order -- a
    separator paragraph between two tables included."""
    out: list[str] = []
    for block in blocks:
        made = _block(block)
        if made and out and made[0] == "t" and out[-1] == "t":
            out.append("p")
        out += made
    return out


def _block(block: Any) -> list[str]:
    if isinstance(block, (m.Paragraph,)):
        return ["p"] if block.inline else []
    if isinstance(block, (m.Heading, m.Rule)):
        return ["p"]
    if isinstance(block, m.CodeBlock):
        body = block.text[:-1] if block.text.endswith("\n") else block.text
        return ["p"] * len(body.split("\n"))
    if isinstance(block, m.Quote):
        return shape(block.blocks)
    if isinstance(block, m.ListBlock):
        out: list[str] = []
        for item in block.items:
            first = item[0] if item else None
            rest = item[1:] if isinstance(first, m.Paragraph) else item
            out.append("p")
            for inner in rest:
                out += _block(inner)
        return out
    if isinstance(block, m.Table):
        return ["t"]
    if isinstance(block, m.HtmlBlock):
        return ["p"] * len(block.text.rstrip("\n").split("\n"))
    return []


def markdown_shape(text: str) -> list[str]:
    """:func:`shape` of Markdown text, as ``insert_markdown`` parses it."""
    from ..markdown.parse import parse_model

    return shape([b for b in parse_model(text) if not isinstance(b, (m.FootnoteDef, m.Comment))])


def _alone(text: str, start: int, end: int) -> bool:
    """Whether ``text[start:end]`` stands alone: blank lines (or the text's ends) on both
    sides, spaces aside."""
    before = text[:start].rstrip(" \t")
    after = text[end:].lstrip(" \t")
    return (not before.strip() or before.endswith("\n\n") or re.search(r"\n[ \t]*\n$", before) is not None) \
        and (not after.strip() or after.startswith("\n\n") or re.match(r"^\n[ \t]*\n", after) is not None)


def placed_blocks(markdown: str, facts: Sequence[Any], blocks: Sequence[str]) -> tuple[dict[int, list[str]], bool]:
    """``{index in facts: the blocks it made}`` for the facts that stand alone in
    ``markdown`` (the text as the handler got it), and whether the blocks lined up.  The
    Markdown between them is read alone too; a value's blocks are its share of ``blocks``
    (a leading separator, after a table already there, and a trailing paragraph, after a
    table at a container's end, aside)."""
    alone = [(k, fact) for k, fact in enumerate(facts) if _alone(markdown, fact.start, fact.end)]
    if not alone:
        return {}, True
    pieces: list[tuple[int | None, str]] = []
    at = 0
    for k, fact in sorted(alone, key=lambda item: item[1].start):
        pieces.append((None, markdown[at:fact.start]))
        pieces.append((k, markdown[fact.start:fact.end]))
        at = fact.end
    pieces.append((None, markdown[at:]))
    expected: list[str] = []
    owners: list[int | None] = []
    for owner, piece in pieces:
        made = markdown_shape(piece) if piece.strip() else []
        if made and expected and made[0] == "t" and expected[-1] == "t":
            expected.append("p")
            owners.append(None)
        expected += made
        owners += [owner] * len(made)
    kinds = [b.rpartition("/")[2][:1] for b in blocks]
    for offset in (0, 1):
        window = kinds[offset:offset + len(expected)]
        extra = len(kinds) - offset - len(expected)
        if window == expected and 0 <= extra <= 1:
            out: dict[int, list[str]] = {}
            for owner, block in zip(owners, blocks[offset:]):
                if owner is not None:
                    out.setdefault(owner, []).append(block)
            return out, True
    return {}, False


# -- the result -------------------------------------------------------------------------------


def brief(fact: Any, blocks: Sequence[str] | None = None, *, cell: str | None = None) -> dict[str, Any]:
    """What a result says of one value written by reference: the reference, its text (a
    long or several-line one only counted, ``chars``), where it went, its sources."""
    out = fact.brief()
    text = out.get("text") or ""
    if len(text) > ECHO_CHARS or "\n" in text:
        out.pop("text", None)
        out["chars"] = len(text)
    if blocks is not None:
        out["blocks"] = list(blocks)
    if cell is not None:
        out["cell"] = cell
    return {k: out[k] for k in ("from", "text", "chars", "blocks", "cell", "sources", "stale") if k in out}


def keep_in_paragraphs(document: Any, record: Mapping[str, Any], blocks: Iterable[str]) -> int:
    """Keep ``record`` (a value's, ``Call.fact``'s) as the source of each paragraph of
    ``blocks`` with text, its whole text the part written; how many were kept."""
    kept = 0
    for block in blocks:
        if not block.startswith("p:"):
            continue
        try:
            text = document.paragraph(block).text
        except (KeyError, ValueError):
            continue
        if text and document.keep_text_facts(block, [{**record, "text": text, "p": 0, "at": 0}]):
            kept += 1
    return kept


def as_fact(address: str, field: str, text: str, record: Mapping[str, Any]) -> Any:
    """A :class:`ooxml_edit.tools.textfacts.Fact` for a value a handler resolved itself
    (``from``), as the dispatcher makes one for a reference."""
    shown = _HANDLE.sub("", textfacts.strip_format(address))
    return textfacts.Fact(field, 0, len(text), text, shown, dict(record))
