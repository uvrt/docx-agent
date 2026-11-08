"""What a replacement keeps of the paragraphs it replaces (round 42): ``insert_markdown`` at
``replace:p:A..p:B`` writes new paragraphs that take the replaced ones' paragraph settings,
as rewriting their text in place would have kept them.

**Matching.**  Each replaced paragraph is a *donor* of one kind: a heading of its outline
level, a list item of its list's kind (bullet or numbered) at its level, or a paragraph.
Each new body paragraph is of the kind Markdown says (``#``, a list item at its nesting,
a paragraph), and takes the donor of its kind at its position among them -- the first new
item the first old item's settings, and so on; past the last donor of its kind, that last
one's (but its page break before: one page break, not one per paragraph).  A list's nesting counts from the replaced items' top level: a flat list replacing
items at level 1 stays at level 1.  Empty donors are passed over while a non-empty one of
their kind is there.  A new paragraph with no donor of its kind is written as Markdown says.

**What is kept.**  The donor's paragraph properties, as they are: its style, list
membership (the same ``w:num``, so items join its list and count on in it), indents,
spacing, alignment, keeps, page break, widow control, tabs, borders, shading and the
paragraph mark's formatting -- never its section break or its revision records.

**What Markdown says wins.**  The kind (a heading's level, bullet or numbered) is
Markdown's.  A numbered list that starts at another number than 1, or a construct whose
style the style map names, keeps Markdown's style and numbering: only the donor's keeps,
page break and widow control are taken.  A list item nested deeper than any donor of its
kind takes the last one with a list of its own (``w:numPr``) at the new level, its indent
then the level's.

The instances (``w:num``) the writer made that no paragraph names any more are removed
(:func:`docx_agent.edit.numbering.reap`).
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..edit import numbering as _numbering
from ..edit import text as _text
from ..oxml.xml import Element, append_in_order, make, remove
from .stylemap import DEFAULT

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_PPR = _W + "pPr"
LISTS = ("bullet", "number")
#: What is kept beside a style or numbering Markdown names.
PAGE_SETTINGS = ("keepNext", "keepLines", "pageBreakBefore", "widowControl")
#: Never kept: the section a paragraph ends, and revision records.
_DROPPED = ("sectPr", "pPrChange")
_MARK_REVISIONS = ("ins", "del", "moveFrom", "moveTo", "rPrChange")
#: The names a result says (word_format's where it has one).
NAMES = {"keepNext": "keep_with_next", "keepLines": "keep_lines", "pageBreakBefore": "page_break_before",
         "widowControl": "widow_control", "numPr": "list", "ind": "indent", "spacing": "spacing",
         "jc": "alignment", "tabs": "tabs", "pBdr": "borders", "shd": "shading", "rPr": "mark_formatting",
         "outlineLvl": "outline_level", "contextualSpacing": "contextual_spacing",
         "suppressAutoHyphens": "no_hyphenation", "framePr": "frame"}
#: Markdown's construct (and level) for a kind, as the style map names it.
_CONSTRUCTS = {"paragraph": "paragraph", "heading": "heading", "bullet": "bullet_list", "number": "ordered_list"}


@dataclass
class Donor:
    id: str
    key: tuple[str, int]
    properties: Element | None
    #: Its list comes from its own ``w:numPr`` (not its style's).
    direct_list: bool
    empty: bool


def _donors(document: "Document", replaced: list[str]) -> list[Donor]:
    """The replaced paragraphs (tables and w:sdt blocks pass), read before anything changes."""
    from ..edit.document import _OutlineLevels
    from ..edit.ids import ParagraphEntry

    out: list[Donor] = []
    outline = _OutlineLevels(document)
    for identifier in replaced:
        try:
            _, entry = document._resolve(identifier)
        except Exception:                       # a control's id: its paragraphs are not donors
            continue
        if not isinstance(entry, ParagraphEntry):
            continue
        element = entry.element
        level = outline.effective(element)
        member = _numbering.membership(document, element)
        if level is not None:
            key = ("heading", min(level, 6))
        elif member is not None:
            key = (member.kind, member.level)
        else:
            key = ("paragraph", 0)
        properties = element.find(W_PPR)
        out.append(Donor(entry.id, key, copy.deepcopy(properties) if properties is not None else None,
                         member is not None and not member.from_style,
                         not _text.paragraph_text(element).strip()))
    full = {d.key for d in out if not d.empty}
    return [d for d in out if not d.empty or d.key not in full]


def keep_replaced(document: "Document", writer, donors: list[Donor]) -> tuple[list[tuple[Element, dict]], set[int]]:
    """Give the writer's tagged paragraphs their donors' settings; returns, per paragraph
    that kept any, ``(element, {"from", "settings"})``, and the instances it stopped naming."""
    groups: dict[tuple[str, int], list[Donor]] = {}
    for donor in donors:
        groups.setdefault(donor.key, []).append(donor)
    levels = [d.key[1] for d in donors if d.key[0] in LISTS]
    base = min(levels) if levels else 0
    seen: dict[tuple[str, int], int] = {}
    out: list[tuple[Element, dict]] = []
    dropped: set[int] = set()
    used: set[str] = set()
    for element, (kind, level), restarts in writer.tagged:
        mapped = (kind, base + level) if kind in LISTS else (kind, level)
        pool = groups.get(mapped, [])
        index = seen.get(mapped, 0)
        seen[mapped] = index + 1
        donor = pool[min(index, len(pool) - 1)] if pool else None
        relevel = False
        if donor is None and kind in LISTS:
            same = [d for d in donors if d.key[0] == kind and d.direct_list]
            if same:
                donor, relevel = same[-1], True
        if donor is None:
            continue
        named = _named(writer.map, kind, level)
        before = _num_ids(element)
        # A donor taken again (more new paragraphs than old) gives no second page break.
        settings = _apply(document, element, donor, mapped[1], page_only=named or restarts, relevel=relevel,
                          again=donor.id in used)
        used.add(donor.id)
        dropped |= before - _num_ids(element)
        if settings:
            out.append((element, {"from": donor.id, "settings": settings}))
    return out, dropped


def _named(style_map, kind: str, level: int) -> bool:
    """Whether the style map names its own style for this kind (Markdown then says it)."""
    construct = _CONSTRUCTS[kind]
    at = level + 1 if kind in LISTS else (level if kind == "heading" else 0)
    return style_map.style_for(construct, at) != DEFAULT.style_for(construct, at)


def _num_ids(element: Element) -> set[int]:
    node = element.find(f"{W_PPR}/{_W}numPr/{_W}numId")
    value = node.get(_W + "val") if node is not None else None
    return {int(value)} if value and value.isdigit() else set()


def _apply(document: "Document", element: Element, donor: Donor, level: int, *, page_only: bool,
           relevel: bool, again: bool) -> list[str]:
    source = donor.properties
    if again and source is not None and source.find(_W + "pageBreakBefore") is not None:
        source = copy.deepcopy(source)
        remove(source.find(_W + "pageBreakBefore"))
    target = element.find(W_PPR)
    if page_only:
        names = []
        for tag in PAGE_SETTINGS:
            node = source.find(_W + tag) if source is not None else None
            if node is None:
                continue
            if target is None:
                target = make("w:pPr")
                element.insert(0, target)
            old = target.find(_W + tag)
            if old is not None:
                remove(old)
            append_in_order(target, copy.deepcopy(node))
            names.append(NAMES[tag])
        return names
    fresh = copy.deepcopy(source) if source is not None else make("w:pPr")
    for tag in _DROPPED:
        for node in fresh.findall(_W + tag):
            remove(node)
    mark = fresh.find(_W + "rPr")
    if mark is not None:
        for tag in _MARK_REVISIONS:
            for node in mark.findall(_W + tag):
                remove(node)
        if not len(mark):
            remove(mark)
    if relevel:
        numbering = fresh.find(_W + "numPr")
        ilvl = numbering.find(_W + "ilvl") if numbering is not None else None
        if ilvl is None and numbering is not None:
            ilvl = make("w:ilvl")
            numbering.insert(0, ilvl)
        if ilvl is not None:
            ilvl.set(_W + "val", str(level))
        for node in fresh.findall(_W + "ind"):
            remove(node)
    if target is not None:
        element.replace(target, fresh)
    else:
        element.insert(0, fresh)
    if not len(fresh):
        remove(fresh)
        return []
    names = []
    for child in fresh:
        if not isinstance(child.tag, str):
            continue
        tag = child.tag[len(_W):] if child.tag.startswith(_W) else child.tag
        if tag == "pStyle":
            names.append(f"style {_style_name(document, child.get(_W + 'val'))}")
        elif tag == "numPr":
            ilvl = child.find(_W + "ilvl")
            names.append(f"list level {int(ilvl.get(_W + 'val') or 0) if ilvl is not None else 0}")
        else:
            names.append(NAMES.get(tag, tag))
    if donor.key[0] in LISTS and not any(n.startswith("list") for n in names):
        names.append(f"list level {donor.key[1]}")         # its style's numbering
    return names


def _style_name(document: "Document", style_id: str | None) -> str:
    found = next((s for s in document.styles if s.id == style_id), None)
    return found.name if found is not None and found.name else str(style_id)
