"""A copy's list items and page settings compared with its reference's, as Word applies them
(round 42, R49).

:mod:`.compare` reads a paragraph's direct formatting as its ``w:pPr`` states it, so two
list items that Word draws differently -- one in a list of dashes, one of round bullets; one
a level deeper; one that lost the "keep with next" its style gave it -- can compare the same,
and two that Word draws alike (a keep stated directly and by the style) can differ.  For the
copies :func:`docx_agent.edit.copies.compare_copy` compares, this module reads each pair of
paragraphs (:func:`.compare.paragraph_pairs`: a paragraph block and its reference, a table's
cell paragraphs) as Word applies them -- docDefaults, the table style, the paragraph style,
the list level, direct formatting (:mod:`.effective`, docx2svg's measured cascade) -- and
reports where they differ:

* ``list_level`` -- the item's list level (0-8), or ``None`` for a paragraph in no list (or
  one whose level draws no label: Word shows none);
* ``list_format`` -- the label its level draws: ``bullet`` and the character (a symbol
  face's character read as the Unicode character it shows: Symbol's U+F0B7 is "•"), or the
  number format and the level's text (``decimal %1.``);
* ``list_numbering`` -- for a numbered item: ``continues`` (one more than the item before it
  at its level, in the same list, since the last item at a shallower level -- in the copy,
  or in the reference) or ``starts at N`` (any other number: a first item, a restart, a list
  that continues another before the copy).  Counted as Word counts (docx2svg's
  :class:`~docx2svg.linebreak.ListCounters`, over the whole body), so a copy whose list
  goes on from the chapter before it is told;
* ``indent_left`` and ``indent_first`` (points; ``indent_first`` negative for a hanging
  indent): the paragraph's own, its style's or its list level's;
* ``keep_with_next``, ``keep_lines``, ``page_break_before``, ``widow_control``: true or false.

**Two items that look the same do not differ**: values, not the XML that states them (a
keep stated directly and by the style, an indent stated as the list gives it, a restarted
list of bullets, another list that draws the same label at the same place).  One
difference per cause: a pair whose paragraph styles differ is not compared here (the style
difference and its fix come first; a check after it compares the rest); a pair whose list
level differs (one in no list among them) is told by ``list_level`` alone, and one whose
label differs by ``list_format`` alone -- the page settings apart, which neither changes.
These differences replace :mod:`.compare`'s ``indent_left``, ``indent_first``,
``keep_with_next``, ``keep_lines`` and ``page_break_before`` for the paragraphs compared here.

**Fixes** (:func:`settings_fix`, the one place they are written): ``word_format`` with the
reference's value, or ``clear`` where the paragraph's style (and list) give it the
reference's value without its own; ``word_lists`` ``level``, ``add`` (a paragraph in no list:
a list of Word's own of the reference's kind, its label compared by the next check),
``remove``, ``format`` (the level's ``number_format`` and ``label`` -- a symbol face's
bullet as its Unicode character -- and, where the item's indents differ, ``text_at`` and
``number_at`` at the reference's: the copy's list only, Word's copy-on-write), ``restart``
(``start``) and ``continue`` (where an earlier list over the same definition is there to
continue).  ``widow_control`` has no ``word_format`` field:
``word_follow_sample`` gives it (the check's ``fix``).  None where no call makes it.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import effective as _effective

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_P, _TBL = _W + "p", _W + "tbl"

#: Every setting compared, in the order reported.
SETTINGS = ("list_level", "list_format", "list_numbering", "indent_left", "indent_first",
            "keep_with_next", "keep_lines", "page_break_before", "widow_control")
#: The list's settings: ``word_lists`` makes them; ``word_follow_sample`` keeps a copy's list.
LIST_SETTINGS = frozenset(SETTINGS[:3])
#: :mod:`.compare`'s properties these replace for the paragraphs compared here.
REPLACED = frozenset({"indent_left", "indent_first", "keep_with_next", "keep_lines", "page_break_before"})
#: The settings ``word_format`` sets by the same name.
_FORMAT_FIELDS = frozenset({"indent_left", "indent_first", "keep_with_next", "keep_lines", "page_break_before"})
_PAGE = ("keep_with_next", "keep_lines", "page_break_before", "widow_control")


# -- list items counted as Word counts them -----------------------------------------------------


def list_items(document: "Document") -> dict[int, tuple]:
    """Every body paragraph that draws a list label, by the ``id()`` of docx2svg's paragraph
    (:func:`.effective.model`): ``(list key, level, format, text, number)`` -- the list key
    the instances counting together share, the level, its ``w:numFmt`` and ``w:lvlText``
    (a symbol face's characters as Unicode) and the number it shows.  Counted in document
    order over the body, table cells included, as the layout counts (docx2svg's
    ``ListCounters``); kept for the document's state."""
    kept = getattr(document, "_tool_list_items", None)
    if kept is not None and kept[0] == document._state:
        return kept[1]
    from docx2svg.linebreak import ListCounters
    from docx2svg.model import Table
    from docx2svg.resolve import resolve_paragraph

    model = _effective.model(document).model
    counters = ListCounters()
    out: dict[int, tuple] = {}

    def walk(item: Any) -> None:
        if isinstance(item, Table):
            for row in item.rows:
                for cell in row:
                    for inner in cell:
                        walk(inner)
            return
        resolved = resolve_paragraph(model, item)
        num_id = resolved.get("numPr.numId")
        if not num_id:
            return
        level = resolved.get("numPr.ilvl", 0) or 0
        if counters.label(model, item, resolved) is None:
            return
        key = model.numbering_lists.get(num_id, ("num", num_id))
        found = model.numbering.get(num_id, {}).get(level)
        number = counters.counts.get(key, {}).get(level)
        out[id(item)] = (key, level, found.format if found is not None else None,
                         _shown(found), number)

    for block in model.body:
        walk(block)
    document._tool_list_items = (document._state, out)
    return out


def _shown(level: Any) -> str | None:
    """A level's text as it shows: a symbol face's characters as the Unicode they draw."""
    if level is None or level.text is None:
        return None
    from ooxml_common.text.symbol_fonts import symbol_face, to_unicode

    font = level.run.get("rFonts.ascii") if level.run else None
    face = symbol_face(getattr(font, "face", None))
    if face is None:
        return level.text
    return "".join(to_unicode(face, char) or char for char in level.text)


# -- one region's settings -----------------------------------------------------------------------


def region_settings(document: "Document", paragraphs: list) -> dict[int, dict[str, Any]]:
    """Each of ``paragraphs`` (elements, in document order: a copy's or a reference's) by
    ``id()``: its settings (the module's names), ``list_numbering`` read along the region."""
    items = list_items(document)
    state = _effective.model(document)
    out: dict[int, dict[str, Any]] = {}
    seen: dict[int, tuple] = {}             # level -> (list key, number) of the item last there
    for element in paragraphs:
        applied = _effective.paragraph(document, element)
        values: dict[str, Any] = {
            "indent_left": round(applied.indent_left, 2),
            "indent_first": round(applied.first_line - applied.hanging, 2),
            "keep_with_next": applied.keep_with_next, "keep_lines": applied.keep_together,
            "page_break_before": applied.page_break_before, "widow_control": applied.widow_control,
            "list_level": None, "list_format": None, "list_numbering": None}
        item = items.get(id(state.paragraph(element)))
        if item is not None:
            key, level, number_format, text, number = item
            values["list_level"] = level
            values["list_format"] = f"{number_format or 'decimal'} {text}"
            if number_format not in ("bullet", "none") and text and "%" in text:
                before = seen.get(level)
                values["list_numbering"] = ("continues" if before is not None and before[0] == key
                                            and number == before[1] + 1 else f"starts at {number}")
            seen[level] = (key, number)
            for deeper in [k for k in seen if k > level]:
                del seen[deeper]
        out[id(element)] = values
    return out


def compare_settings(document: "Document", reference_document: "Document", pairs: list[tuple],
                     ids: dict[int, tuple[str, str]], mine: dict[int, dict], theirs: dict[int, dict],
                     reference_ids: dict[int, str]) -> list[dict]:
    """The differences of ``pairs`` (``(paragraph, reference)`` elements, styles alike), each
    ``{"block", "property", "paragraphs", "value", "reference", "fix"}`` grouped per block,
    property and values; ``ids`` gives each paragraph's ``(block, paragraph id)``, ``mine``
    and ``theirs`` the regions' settings (:func:`region_settings`)."""
    groups: dict[tuple, list[str]] = {}
    samples: dict[tuple, str] = {}
    for paragraph, reference in pairs:
        where = ids.get(id(paragraph))
        if where is None:
            continue
        block, identifier = where
        a, b = mine.get(id(paragraph)), theirs.get(id(reference))
        if a is None or b is None:
            continue
        names = list(SETTINGS)
        if a["list_level"] != b["list_level"]:
            names = ["list_level", *_PAGE]   # the level's label, numbers and indents follow it
        elif a["list_format"] != b["list_format"]:
            names = ["list_format", *_PAGE]  # its fix places the label as the reference's
        for name in names:
            if a[name] != b[name]:
                key = (block, name, a[name], b[name])
                groups.setdefault(key, []).append(identifier)
                samples.setdefault(key, (reference_ids.get(id(reference), ""), b, a))
    out = []
    for (block, name, value, reference), paragraphs in groups.items():
        out.append({"block": block, "property": name, "paragraphs": paragraphs, "value": value,
                    "reference": reference,
                    "fix": settings_fix(document, name, value, reference, paragraphs, *samples[(block, name, value,
                                                                                                reference)])})
    order = {name: k for k, name in enumerate(SETTINGS)}
    out.sort(key=lambda d: order[d["property"]])
    return out


# -- fixes -----------------------------------------------------------------------------------------


def settings_fix(document: "Document", name: str, value: Any, reference: Any, paragraphs: list[str],
                 sample: str, sample_settings: dict, own_settings: dict) -> dict | None:
    """The calls that give ``paragraphs`` the reference's ``name`` (the module's "Fixes"):
    every fix of a list item or a page setting is written here, and only here.  ``sample``
    is the reference paragraph (for a call that makes a paragraph look like another, once
    one exists: that call would replace these); ``sample_settings`` its settings and
    ``own_settings`` the first paragraph's."""
    if name in _FORMAT_FIELDS:
        if _style_gives(document, name, reference, paragraphs):
            return {"tool": "word_format", "calls": [{"targets": paragraphs, "clear": [name]}]}
        return {"tool": "word_format", "calls": [{"targets": paragraphs, name: reference}]}
    if name == "list_level":
        if reference is None:
            return {"tool": "word_lists", "calls": [{"targets": paragraphs, "action": "remove"}]}
        if value is None:
            # A new list of Word's own at the reference's level; a check after it compares
            # its label with the reference's.
            kind = "bullet" if (sample_settings.get("list_format") or "").startswith("bullet ") else "number"
            return {"tool": "word_lists", "calls": [{"targets": paragraphs, "action": "add", "kind": kind,
                                                      "level": reference}]}
        return {"tool": "word_lists", "calls": [{"targets": paragraphs, "action": "level", "level": reference}]}
    if name == "list_format":
        number_format, _, label = reference.partition(" ")
        place: dict[str, float] = {}
        if (own_settings["indent_left"], own_settings["indent_first"]) != \
                (sample_settings["indent_left"], sample_settings["indent_first"]):
            # The label and text where the reference has them ("Define new multilevel list").
            place = {"text_at": sample_settings["indent_left"],
                     "number_at": round(sample_settings["indent_left"] + sample_settings["indent_first"], 2)}
        calls = [{"targets": [p], "action": "format", "level": level, "number_format": number_format,
                  "label": label, **place} for p, level in _first_per_list(document, paragraphs)]
        return {"tool": "word_lists", "calls": calls} if calls else None
    if name == "list_numbering":
        if reference.startswith("starts at "):
            return {"tool": "word_lists", "calls": [{"targets": [p], "action": "restart",
                                                      "start": int(reference.split()[-1])} for p in paragraphs]}
        calls = [{"targets": [p], "action": "continue"} for p in paragraphs if _can_continue(document, p)]
        return {"tool": "word_lists", "calls": calls} if len(calls) == len(paragraphs) else None
    return None


def _style_gives(document: "Document", name: str, reference: Any, paragraphs: list[str]) -> bool:
    """Whether every paragraph has ``name`` of its own and its style (and list) alone gives
    the reference's value: clearing it is the fix, the template's value kept (rule 6)."""
    from .compare import _PARAGRAPH

    tag, attributes = _PARAGRAPH[name]
    for identifier in paragraphs:
        element = document._resolve(identifier)[1].element
        properties = element.find(_W + "pPr")
        node = properties.find(_W + tag) if properties is not None else None
        if node is None or (attributes and not any(node.get(_W + a) is not None for a in attributes)):
            return False
        applied = _effective.paragraph(document, element, direct=False)
        own = (round(applied.indent_left, 2) if name == "indent_left" else
               round(applied.first_line - applied.hanging, 2) if name == "indent_first" else
               getattr(applied, "keep_together" if name == "keep_lines" else name))
        if own != reference:
            return False
    return True


def _first_per_list(document: "Document", paragraphs: list[str]) -> list[tuple[str, int]]:
    """The first of ``paragraphs`` in each list and level: one format call per list."""
    from .numbering import membership

    seen: set = set()
    out = []
    for identifier in paragraphs:
        found = membership(document, document._resolve(identifier)[1].element)
        if found is None or (found.num_id, found.level) in seen:
            continue
        seen.add((found.num_id, found.level))
        out.append((identifier, found.level))
    return out


def _can_continue(document: "Document", identifier: str) -> bool:
    """Whether ``word_lists`` ``continue`` finds a list to continue for the paragraph: an
    earlier paragraph in another list over the same definition."""
    from .numbering import Numbering, membership

    part, entry = document._resolve(identifier)
    found = membership(document, entry.element)
    if found is None:
        return False
    numbering = Numbering(document)
    mine = numbering.canonical(found.abstract_id)
    for other in reversed(document._index(part).paragraphs[:entry.index]):
        other_list = membership(document, other.element)
        if other_list is not None and other_list.num_id != found.num_id and \
                numbering.canonical(other_list.abstract_id) == mine:
            return True
    return False


__all__ = ["LIST_SETTINGS", "REPLACED", "SETTINGS", "compare_settings", "list_items", "region_settings",
           "settings_fix"]
