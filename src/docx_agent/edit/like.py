"""Paragraphs made to look like a given one (round 42, R47): :meth:`Document.format_like
<docx_agent.Document.format_like>` and ``word_insert_markdown``'s ``like``.

What makes a paragraph look as it does, besides its text and runs, is what its own
``w:pPr`` says beside its style.  A paragraph *like* a sample takes, from the sample's own
properties as its XML states them:

* its **paragraph style** (``w:pStyle``; none is the default paragraph style);
* its **list** (``w:numPr``): the same instance at the same level -- so it is an item of
  the sample's list, numbered on from the items before it, with the sample's marker -- as
  the sample states it: its own ``numId`` and ``ilvl`` where it has them, nothing where its
  style numbers it (the style's list then numbers both), ``numId`` 0 where it is taken out
  of the list its style would give it;
* its **direct indents** (``w:ind``);
* its **page settings**: keep with next, keep lines together, page break before and widow
  control (``w:keepNext``, ``w:keepLines``, ``w:pageBreakBefore``, ``w:widowControl``).

Each is taken whole: where the sample has none of its own, the paragraph's own is removed,
so that with the same style both resolve to the same.  Nothing else is taken: run
formatting (the text's emphasis and the paragraph mark's), spacing, alignment, tabs,
borders and shading stay the paragraph's own, as do its text, section break and change
records.  A paragraph that already looks like the sample is left as it is (no change).

Tracked, the change is a ``w:pPrChange`` as every paragraph-property change is.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from ..oxml.xml import Element, append_in_order, make, remove

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: What a paragraph takes from its sample: the part's name (as results say it) -> the
#: ``w:pPr`` children it is.
PARTS: dict[str, tuple[str, ...]] = {
    "style": ("pStyle",),
    "list": ("numPr",),
    "indent": ("ind",),
    "keep_with_next": ("keepNext",),
    "keep_lines": ("keepLines",),
    "page_break_before": ("pageBreakBefore",),
    "widow_control": ("widowControl",),
}

#: Not taken, said in the result: the rest of the sample's look a paragraph keeps its own.
NOT_TAKEN = "run formatting, spacing, alignment and tabs stay each paragraph's own"


@dataclass
class Look:
    """A sample paragraph's own properties that make its look (copies, not the sample's
    elements: a look outlives a sample deleted after it is taken)."""

    sample: str
    #: Part name -> its ``w:pPr`` child (a copy), or ``None`` where the sample has none.
    parts: dict[str, Element | None] = field(default_factory=dict)

    def deeper(self, document: "Document", levels: int) -> "Look | None":
        """The look of an item ``levels`` deeper in the sample's list, as Word demotes an
        item (Tab): the same style, list and page settings, the level ``levels`` lower and
        the level's own indent (no direct one).  ``None`` where the sample is in no list,
        or the list has no such level."""
        from . import numbering as _numbering

        found = _numbering.membership(document, self._element(document))
        if found is None or found.level + levels >= _numbering.LEVELS:
            return None
        if _numbering.Numbering(document).level(found.num_id, found.level + levels) is None:
            return None
        numbering = make("w:numPr")
        numbering.append(make("w:ilvl", **{"w:val": str(found.level + levels)}))
        numbering.append(make("w:numId", **{"w:val": str(found.num_id)}))
        parts = dict(self.parts)
        parts["list"] = numbering
        parts["indent"] = None
        return Look(self.sample, parts)

    def _element(self, document: "Document") -> Element:
        """A stand-in paragraph with this look's properties (for reading its membership)."""
        paragraph = make("w:p")
        take(paragraph, self, set())
        return paragraph


def look_of(document: "Document", identifier: str) -> Look:
    """The look of the paragraph ``identifier``: its own properties of :data:`PARTS`."""
    from .document import EditError

    try:
        part, entry = document._paragraph_entry(identifier)
    except (KeyError, ValueError) as error:
        raise EditError(f"no paragraph {identifier}: {error}") from None
    properties = entry.element.find(_W + "pPr")
    parts: dict[str, Element | None] = {}
    for name, tags in PARTS.items():
        node = properties.find(_W + tags[0]) if properties is not None else None
        parts[name] = _clean(node) if node is not None else None
    return Look(entry.id, parts)


def _clean(node: Element) -> Element:
    """A copy of a property without its change records (a list's ``w:ins``, a numbering
    change: the sample's history, not its look)."""
    out = copy.deepcopy(node)
    for child in list(out):
        if isinstance(child.tag, str) and child.tag in (_W + "ins", _W + "del", _W + "numberingChange"):
            remove(child)
    return out


def _key(node: Element | None) -> tuple | None:
    if node is None:
        return None
    return (node.tag, tuple(sorted(node.attrib.items())),
            tuple(_key(child) for child in node if isinstance(child.tag, str)))


def differs(paragraph: Element, look: Look, skip: set[str]) -> list[str]:
    """The parts of :data:`PARTS` (less ``skip``) in which ``paragraph`` differs from ``look``."""
    properties = paragraph.find(_W + "pPr")
    out = []
    for name, tags in PARTS.items():
        if name in skip:
            continue
        own = properties.find(_W + tags[0]) if properties is not None else None
        if _key(_clean(own) if own is not None else None) != _key(look.parts.get(name)):
            out.append(name)
    return out


def take(paragraph: Element, look: Look, skip: set[str]) -> list[str]:
    """Give ``paragraph`` the parts of ``look`` it differs in (less ``skip``); returns them."""
    changed = differs(paragraph, look, skip)
    if not changed:
        return []
    properties = paragraph.find(_W + "pPr")
    if properties is None:
        properties = make("w:pPr")
        paragraph.insert(0, properties)
    for name in changed:
        tag = PARTS[name][0]
        own = properties.find(_W + tag)
        if own is not None:
            remove(own)
        wanted = look.parts.get(name)
        if wanted is not None:
            append_in_order(properties, copy.deepcopy(wanted))
    if not len(properties):
        remove(properties)
    return changed


def format_like(document: "Document", identifier: str, look: Look, *,
                skip: set[str] | frozenset = frozenset()) -> tuple["EditResult", list[str]]:
    """Give the paragraph ``identifier`` ``look`` (one undo step, tracked when tracking is
    on): ``(result, parts set)``; nothing set is no change."""
    from . import numbering as _numbering
    from .document import EditResult

    part, entry = document._paragraph_entry(identifier)
    before = _numbering.membership(document, entry.element)
    if not differs(entry.element, look, set(skip)):
        return EditResult(entry.id, changed=False), []
    with document._edit():
        renames = document._prepare(part, [entry])
        with document._track_properties([part]):
            changed = take(entry.element, look, set(skip))
        removed = _numbering.reap(document, {before.num_id}) if before is not None and "list" in changed else []
        document.package.mark_dirty(part)
        if "list" in changed and _numbering.Numbering(document).part is not None:
            document.package.mark_dirty(_numbering.Numbering(document).part)
    result = EditResult(renames.get(entry.id, entry.id), renamed=renames)
    result.removed += removed
    return result, changed


# -- Markdown list items like given paragraphs (word_insert_markdown's like) -------------------

#: The list roles a level may be given for alone (a style map's keys): role -> (kind, depth).
ROLES = {"bullet": ("bullet", 1), "bullet2": ("bullet", 2), "bullet3": ("bullet", 3), "bullet4": ("bullet", 4),
         "bullet5": ("bullet", 5), "number": ("number", 1), "number2": ("number", 2), "number3": ("number", 3),
         "number4": ("number", 4), "number5": ("number", 5)}


class LevelLooks:
    """The looks Markdown list items take, by level (1 the top): ``{1: "p:3", 2: "p:4"}``
    -- every list's items at that level, bullets and numbers -- or by role (``bullet2``:
    a bullet list's second level only), or one paragraph for the top level.  A level given
    no paragraph, below one that is, follows the nearest given level's list a level deeper
    per level, as Word demotes an item (:meth:`Look.deeper`); a level above every given
    one is written as without ``like``."""

    def __init__(self, document: "Document", like: "str | dict") -> None:
        from .document import EditError

        given = {1: like} if isinstance(like, str) else dict(like or {})
        self.document = document
        self._looks: dict[tuple[str | None, int], Look] = {}
        for key, sample in given.items():
            if isinstance(key, str) and key in ROLES:
                where: tuple[str | None, int] = ROLES[key]
            elif isinstance(key, int) and not isinstance(key, bool) or (isinstance(key, str) and key.isdigit()):
                where = (None, int(key))
            else:
                raise EditError(f"like: {key!r} is a level (1 the top) or one of {', '.join(ROLES)}")
            if not 1 <= where[1] <= 9:
                raise EditError(f"like: level {where[1]} is 1 to 9")
            if sample:
                self._looks[where] = look_of(document, sample)
        self.warnings: list[str] = []
        self._told: set[tuple] = set()

    def __bool__(self) -> bool:
        return bool(self._looks)

    def at(self, kind: str, depth: int) -> Look | None:
        """The look of ``kind`` (``bullet``, ``number``) items at Markdown ``depth`` (1 the top)."""
        for level in range(depth, 0, -1):
            look = self._looks.get((kind, level)) or self._looks.get((None, level))
            if look is None:
                continue
            if level == depth:
                self._check(look, kind, depth)
                return look
            deeper = look.deeper(self.document, depth - level)
            if deeper is None:
                self._tell(("deep", look.sample, depth), f"like: level {depth} has no paragraph given, and "
                           f"{look.sample}'s list has no level {depth - level} below {look.sample}'s: written as without like")
            return deeper
        return None

    def _check(self, look: Look, kind: str, depth: int) -> None:
        from . import numbering as _numbering

        found = _numbering.membership(self.document, look._element(self.document))
        if found is None:
            self._tell(("none", look.sample), f"like: {look.sample} is in no list: the level {depth} items "
                       "are written like it, with no marker")
        elif found.kind != kind:
            self._tell(("kind", look.sample), f"like: the Markdown's level {depth} items are {kind}s and "
                       f"{look.sample} is a {found.kind} item: they follow {look.sample}")

    def _tell(self, key: tuple, text: str) -> None:
        if key not in self._told:
            self._told.add(key)
            self.warnings.append(text)
