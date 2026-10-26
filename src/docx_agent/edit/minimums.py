"""The application's minimum text sizes, by kind of text (round 17, P6b).

One setting, read by everything that judges or writes text size: ``check``'s layout facts
(``small_text``, measured on the layout) and the builders' results (``word_insert_markdown``,
``word_edit_table``: the text they wrote that its styles make smaller, measured on the
style cascade, as Word draws a paragraph -- no autofit).  The agent tool layer sets it from
its format (``docx_format(min_text_size=...)``, :attr:`Document.min_text_size`); ``None`` --
the default -- is no minimum, and nothing reports it.

A minimum is one number of points for every kind, or a mapping by kind (:data:`TEXT_KINDS`);
a kind it leaves out is not checked.  ``caption`` is a paragraph in the Caption style,
wherever it is (its minimum is used instead of its place's: body or table).
"""

from __future__ import annotations

from collections.abc import Mapping
from typing import TYPE_CHECKING, Any, Iterable

from . import effective as _effective
from . import text as _text
from .errors import EditError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: The kinds of text a minimum is set for.
TEXT_KINDS = ("body", "table", "footnote", "endnote", "header", "footer", "caption")


class MinimumError(EditError):
    """A minimum text size setting that is not one."""


def text_minimums(value: Any) -> dict[str, float]:
    """``{kind: points}`` for a setting: ``None`` none, a number every kind, a mapping the
    kinds it names (``None`` values left out); :class:`MinimumError` for anything else."""
    if value is None:
        return {}

    def points(kind: str, given: Any) -> float:
        if isinstance(given, bool) or not isinstance(given, (int, float)) or not 1 <= given <= 400:
            raise MinimumError(f"min_text_size{f'[{kind!r}]' if kind else ''} {given!r}: points between "
                               f"1 and 400, or a mapping of {', '.join(TEXT_KINDS)} to points")
        return float(given)

    if isinstance(value, Mapping):
        unknown = [k for k in value if k not in TEXT_KINDS]
        if unknown:
            raise MinimumError(f"min_text_size has no kind {unknown[0]!r}; the kinds are {', '.join(TEXT_KINDS)}")
        found = {kind: points(kind, given) for kind, given in value.items() if given is not None}
        if not found:
            raise MinimumError(f"min_text_size is an empty mapping; name {', '.join(TEXT_KINDS)}, or give "
                               "one number")
        return found
    one = points("", value)
    return {kind: one for kind in TEXT_KINDS}


def is_caption(document: "Document", paragraph) -> bool:
    """Whether a paragraph is in the Caption style (by its id or its name)."""
    properties = paragraph.find(_W + "pPr")
    node = properties.find(_W + "pStyle") if properties is not None else None
    if node is None:
        return False
    style_id = node.get(_W + "val") or ""
    if style_id.lower() == "caption":
        return True
    found = next((s for s in document.styles if s.id == style_id), None)
    return bool(found is not None and (found.name or "").lower() == "caption")


def minimum_for(minimums: Mapping[str, float], kind: str, caption: bool) -> tuple[str, float | None]:
    """``(kind, minimum)`` of text of ``kind`` (a caption's own when it has one)."""
    if caption and "caption" in minimums:
        return "caption", minimums["caption"]
    return kind, minimums.get(kind)


def place_kind(document: "Document", part: str, paragraph) -> str:
    """``body``, ``table``, ``footnote``, ``endnote``, ``header`` or ``footer``: where a
    paragraph is."""
    story = document._story_of(part)
    if story == "body":
        node = paragraph.getparent()
        while node is not None:
            if node.tag == _W + "tc":
                return "table"
            node = node.getparent()
        return "body"
    if story.startswith("header"):
        return "header"
    if story.startswith("footer"):
        return "footer"
    return {"footnotes": "footnote", "endnotes": "endnote"}.get(story, story)


def styled_below(document: "Document", paragraphs: Iterable[str], minimums: Mapping[str, float]) -> list[dict]:
    """Of ``paragraphs`` (ids), the text whose effective size -- its styles and direct
    formatting, superscript and subscript left out -- is below its kind's minimum: one
    ``{"para", "kind", "size", "min", "style_size"}`` each, its smallest size, smallest
    first; ``fix`` clears a direct size where the styles give at least the minimum."""
    out = []
    for identifier in paragraphs:
        try:
            part, entry = document._resolve(identifier)
        except (KeyError, EditError):
            continue
        element = entry.element
        if element.tag != _W + "p":
            continue
        kind, minimum = minimum_for(minimums, place_kind(document, part, element), is_caption(document, element))
        if minimum is None:
            continue
        smallest = None
        walked = _text.walk(element)
        seen: set[int] = set()
        for atom in walked.atoms:
            if atom.run is None or atom.kind != "text" or id(atom.run) in seen or not atom.char.strip():
                continue
            seen.add(id(atom.run))
            run = _effective.run(document, element, atom.run, atom.char)
            if run.vertical_align in ("superscript", "subscript"):
                continue
            smallest = run.size if smallest is None else min(smallest, run.size)
        if smallest is None or smallest >= minimum:
            continue
        found: dict[str, Any] = {"para": entry.id, "kind": kind, "size": smallest, "min": minimum}
        style_size = _effective.run(document, element, None, direct=False).size
        found["style_size"] = style_size
        if style_size >= minimum:
            found["fix"] = {"tool": "word_format", "calls": [{"targets": [entry.id], "clear": ["size"]}]}
        out.append(found)
    out.sort(key=lambda e: e["size"])
    return out


__all__ = ["MinimumError", "TEXT_KINDS", "is_caption", "minimum_for", "place_kind", "styled_below",
           "text_minimums"]
