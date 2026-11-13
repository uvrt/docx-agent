"""Lists and numbering: making a paragraph a list item, its level, restarting and continuing
numbering, leaving a list -- with ``numbering.xml`` kept whole.

Word's model: a paragraph is in a list when its ``w:numPr`` (direct, or from its style)
names a ``w:num``; a ``w:num`` is an *instance* over a ``w:abstractNum`` (the levels'
formats and indents) and may override levels (``w:lvlOverride``, ``w:startOverride``).
Every instance over one abstract definition counts on from the others unless it restarts,
which is why "restart numbering" is a new ``w:num`` over the same abstract definition with
a ``w:startOverride``, as Word does, and why a new, separate list gets an abstract
definition of its own (with a ``w:nsid`` of its own): sharing one would continue the other
list's numbers -- a list definition shared by accident.

**Integrity** is kept by every operation and checked by :func:`problems`: every ``numId``
a paragraph or style names resolves (``0`` means "no list"), every ``w:num`` names an
abstract definition that exists, levels are 0-8, and an instance or abstract definition an
operation stops using is removed (one the document had unused before is left as it was).
A level's format is changed copy-on-write: when another instance shares the abstract
definition, the instance changed gets a copy of its own.

New lists are written as Word's own Bullets and Numbering buttons write them: a
``hybridMultilevel`` definition of nine levels, 0.5 inch apart with a 0.25 inch hanging
indent; bullets ``Symbol`` U+F0B7, ``Courier New`` "o", ``Wingdings`` U+F0A7 in turn,
numbers ``1.``, ``a.``, ``i.`` in turn (ROADMAP.md, Phase E1, records how this was checked).
"""

from __future__ import annotations

import copy
import hashlib
from contextlib import contextmanager
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterator

from lxml import etree

from ..oxml.xml import Element, append_in_order, make, qn, remove
from .ids import ParagraphEntry

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
NUMBERING_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
LEVELS = 9

#: Word's own bullet list: per level, (character, font, template code).
_BULLETS = [("", "Symbol", "04090001"), ("o", "Courier New", "04090003"), ("", "Wingdings", "04090005")]
#: Word's own numbered list: per level, (format, justification, template code).
_NUMBERS = [("decimal", "left", "0409000F"), ("lowerLetter", "left", "04090019"), ("lowerRoman", "right", "0409001B")]


@dataclass(frozen=True)
class ListMembership:
    """Where a paragraph is in a list."""

    #: The ``w:num`` instance.
    num_id: int
    #: 0-based level (``w:ilvl``).
    level: int
    #: The ``w:abstractNum`` it is over.
    abstract_id: int | None
    #: The level's ``w:numFmt`` (``bullet``, ``decimal``, ``lowerLetter``...).
    format: str | None
    #: The level's ``w:lvlText`` (``%1.``, a bullet character).
    text: str | None
    #: Whether the numbering comes from the paragraph's style rather than its own ``w:numPr``.
    from_style: bool

    @property
    def kind(self) -> str:
        """``"bullet"`` or ``"number"``."""
        return "bullet" if self.format == "bullet" else "number"


@contextmanager
def reading(document: "Document") -> Iterator[None]:
    """Round 45: while nothing changes the numbering part (a read of it), its lists and
    definitions listed once (:meth:`Numbering.nums`, :meth:`Numbering.abstracts`) rather
    than on every question."""
    if document.__dict__.get("_numbering_held") is not None:
        yield
        return
    document.__dict__["_numbering_held"] = {}
    try:
        yield
    finally:
        document.__dict__["_numbering_held"] = None


class Numbering:
    """``numbering.xml``, read live."""

    def __init__(self, document: "Document") -> None:
        self._document = document

    @property
    def part(self) -> str | None:
        return self._document.package.numbering_part()

    @property
    def root(self) -> Element | None:
        part = self.part
        return self._document.package.tree(part) if part is not None else None

    def nums(self) -> dict[int, Element]:
        return self._listed("num", "numId")

    def abstracts(self) -> dict[int, Element]:
        return self._listed("abstractNum", "abstractNumId")

    def _listed(self, tag: str, attribute: str) -> dict[int, Element]:
        """The part's ``tag`` elements by their number; inside :func:`reading` listed once
        (round 45: a table of contents asked for every heading's list several times over)."""
        held = self._document.__dict__.get("_numbering_held")
        if held is not None:
            if "root" not in held:
                held["root"] = self.root
            root = held["root"]
            kept = held.get(tag)
            if kept is not None:
                return kept[1]
        else:
            root = self.root
        out: dict[int, Element] = {}
        for node in (root.findall(_W + tag) if root is not None else []):
            number = _int(node.get(_W + attribute))
            if number is not None:
                out[number] = node
        if held is not None:
            held[tag] = (root, out)
        return out

    def abstract_of(self, num_id: int) -> int | None:
        node = self.nums().get(num_id)
        ref = node.find(_W + "abstractNumId") if node is not None else None
        return _int(ref.get(_W + "val")) if ref is not None else None

    def canonical(self, abstract_id: int | None) -> int | None:
        """The definition whose levels and counts ``abstract_id`` uses: the first in the
        part with its ``w:nsid`` -- definitions sharing one are one list to Word, the first
        one's levels drawn (docx2svg's ROADMAP, "Numbering and lists -- measured", N.1)."""
        abstracts = self.abstracts()
        node = abstracts.get(abstract_id)
        nsid = node.find(_W + "nsid") if node is not None else None
        value = (nsid.get(_W + "val") or "").upper() if nsid is not None else ""
        if not value:
            return abstract_id
        for identifier, other in abstracts.items():
            found = other.find(_W + "nsid")
            if found is not None and (found.get(_W + "val") or "").upper() == value:
                return identifier
        return abstract_id

    def level(self, num_id: int, level: int) -> Element | None:
        """The ``w:lvl`` that applies: the instance's override, else its definition's (the
        first definition with the same ``w:nsid``, as Word reads it)."""
        num = self.nums().get(num_id)
        if num is None:
            return None
        for override in num.findall(_W + "lvlOverride"):
            if _int(override.get(_W + "ilvl")) == level and override.find(_W + "lvl") is not None:
                return override.find(_W + "lvl")
        abstract = self.abstracts().get(self.canonical(self.abstract_of(num_id)))
        if abstract is None:
            return None
        link = abstract.find(_W + "numStyleLink")
        if link is not None:
            linked = self._style_num(link.get(_W + "val"))
            if linked is not None and linked != num_id:
                return self.level(linked, level)
        return next((lvl for lvl in abstract.findall(_W + "lvl") if _int(lvl.get(_W + "ilvl")) == level), None)

    def _style_num(self, style_id: str | None) -> int | None:
        styles = self._document.package.styles_part()
        root = self._document.package.tree(styles) if styles else None
        if root is None or style_id is None:
            return None
        for style in root.findall(_W + "style"):
            if style.get(_W + "styleId") == style_id:
                ref = style.find(f"{_W}pPr/{_W}numPr/{_W}numId")
                return _int(ref.get(_W + "val")) if ref is not None else None
        return None


def _int(value: str | None) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


# -- membership ------------------------------------------------------------------------------


def _style_numbering(document: "Document", style_id: str | None) -> tuple[int | None, int | None]:
    """``(numId, ilvl)`` a paragraph style gives through its ``w:basedOn`` chain; inside
    :func:`reading` found once per style (round 57, R124: the styles were listed again for
    every paragraph asked about)."""
    held = document.__dict__.get("_numbering_held")
    if held is None:
        return _style_numbering_now(document, style_id)
    kept = held.setdefault("style numbering", {})
    if style_id not in kept:
        kept[style_id] = _style_numbering_now(document, style_id)
    return kept[style_id]


def _style_numbering_now(document: "Document", style_id: str | None) -> tuple[int | None, int | None]:
    part = document.package.styles_part()
    root = document.package.tree(part) if part else None
    if root is None:
        return None, None
    styles = {s.get(_W + "styleId"): s for s in root.findall(_W + "style")}
    if style_id is None:
        style_id = next((s.get(_W + "styleId") for s in styles.values()
                         if s.get(_W + "type") == "paragraph" and s.get(_W + "default") in ("1", "true")), None)
    seen = set()
    num_id = level = None
    while style_id and style_id not in seen and style_id in styles:
        seen.add(style_id)
        node = styles[style_id]
        numbering = node.find(f"{_W}pPr/{_W}numPr")
        if numbering is not None:
            if num_id is None and numbering.find(_W + "numId") is not None:
                num_id = _int(numbering.find(_W + "numId").get(_W + "val"))
            if level is None and numbering.find(_W + "ilvl") is not None:
                level = _int(numbering.find(_W + "ilvl").get(_W + "val"))
        based = node.find(_W + "basedOn")
        style_id = based.get(_W + "val") if based is not None else None
    return num_id, level


def style_list(document: "Document", style_id: str | None) -> tuple[int, int] | None:
    """``(numId, ilvl)`` a paragraph style numbers its paragraphs with, as Word numbers them, or
    ``None`` when it does not.  The style's ``w:numPr`` is its own or the first its
    ``basedOn`` chain gives (the default paragraph style's for ``None``); ``numId`` 0, an
    instance the document lacks and one naming no definition number nothing.  A level tied
    to another style (its ``w:pStyle`` names a style other than the one holding the
    ``w:numPr``) numbers nothing either: Word draws no label for such a paragraph, while a
    style inheriting the tied style's ``w:numPr`` through ``basedOn`` is numbered (measured,
    round 23: ``feedback-round23/measurements-F.md``)."""
    part = document.package.styles_part()
    root = document.package.tree(part) if part else None
    if root is None:
        return None
    styles = {s.get(_W + "styleId"): s for s in root.findall(_W + "style")}
    if style_id is None:
        style_id = next((s.get(_W + "styleId") for s in styles.values()
                         if s.get(_W + "type") == "paragraph" and s.get(_W + "default") in ("1", "true")), None)
    seen: set[str] = set()
    num_id = level = holder = None
    current = style_id
    while current and current not in seen and current in styles:
        seen.add(current)
        node = styles[current]
        numbering = node.find(f"{_W}pPr/{_W}numPr")
        if numbering is not None:
            if num_id is None and numbering.find(_W + "numId") is not None:
                num_id = _int(numbering.find(_W + "numId").get(_W + "val"))
                holder = current
            if level is None and numbering.find(_W + "ilvl") is not None:
                level = _int(numbering.find(_W + "ilvl").get(_W + "val"))
        based = node.find(_W + "basedOn")
        current = based.get(_W + "val") if based is not None else None
    level = level if level is not None else 0
    numbering_ = Numbering(document)
    if not num_id or num_id not in numbering_.nums() or numbering_.abstract_of(num_id) not in numbering_.abstracts():
        return None
    lvl = numbering_.level(num_id, level)
    tied = lvl.find(_W + "pStyle") if lvl is not None else None
    if tied is not None and tied.get(_W + "val") and tied.get(_W + "val") != holder:
        return None
    return num_id, level


def paragraph_style_id(paragraph: Element) -> str | None:
    """The style a paragraph names (``None``: the default paragraph style)."""
    found = paragraph.find(f"{_W}pPr/{_W}pStyle")
    return found.get(_W + "val") if found is not None else None


def indent_note(document: "Document", paragraph: Element | None, num_id: int, level: int, *,
                style_ind: Element | None = None) -> str | None:
    """A note for a paragraph given direct numbering over a style that numbers it with
    another indent: direct numbering takes the list level's ``w:ind``, numbering from the
    style the style's own (measured, round 23).  ``style_ind``: the style's indent, given
    instead of the paragraph."""
    numbering = Numbering(document)
    lvl = numbering.level(num_id, level)
    level_ind = lvl.find(f"{_W}pPr/{_W}ind") if lvl is not None else None
    if style_ind is None and paragraph is not None:
        style_ind = _style_indent(document, paragraph_style_id(paragraph))
    if level_ind is None or style_ind is None:
        return None

    def twips(node: Element, *names: str) -> int:
        for name in names:
            value = _int(node.get(_W + name))
            if value is not None:
                return value
        return 0

    def form(node: Element) -> tuple[int, int]:
        return twips(node, "left", "start"), twips(node, "hanging") - twips(node, "firstLine")

    if form(level_ind) == form(style_ind):
        return None
    left, hanging = form(level_ind)
    return (f"direct numbering: the paragraph takes its list level's indent (left {left / 20:g} pt, "
            f"hanging {hanging / 20:g} pt), not its style's")


def _style_indent(document: "Document", style_id: str | None) -> Element | None:
    part = document.package.styles_part()
    root = document.package.tree(part) if part else None
    if root is None:
        return None
    styles = {s.get(_W + "styleId"): s for s in root.findall(_W + "style")}
    seen: set[str] = set()
    while style_id and style_id not in seen and style_id in styles:
        seen.add(style_id)
        found = styles[style_id].find(f"{_W}pPr/{_W}ind")
        if found is not None:
            return found
        based = styles[style_id].find(_W + "basedOn")
        style_id = based.get(_W + "val") if based is not None else None
    return None


def membership(document: "Document", paragraph: Element) -> ListMembership | None:
    """The list a paragraph is in, or ``None``."""
    properties = paragraph.find(_W + "pPr")
    direct = properties.find(_W + "numPr") if properties is not None else None
    style = properties.find(_W + "pStyle") if properties is not None else None
    style_num, style_level = _style_numbering(document, style.get(_W + "val") if style is not None else None)
    num_id = level = None
    from_style = True
    if direct is not None:
        ref = direct.find(_W + "numId")
        if ref is not None:
            num_id = _int(ref.get(_W + "val"))
            from_style = False
        lvl = direct.find(_W + "ilvl")
        if lvl is not None:
            level = _int(lvl.get(_W + "val"))
    if num_id is None:
        num_id = style_num
    if level is None:
        level = style_level if style_level is not None else 0
    if not num_id:
        return None
    numbering = Numbering(document)
    if num_id not in numbering.nums():
        return None
    lvl = numbering.level(num_id, level)
    fmt = lvl.find(_W + "numFmt") if lvl is not None else None
    text = lvl.find(_W + "lvlText") if lvl is not None else None
    return ListMembership(num_id, level, numbering.abstract_of(num_id),
                          fmt.get(_W + "val") if fmt is not None else None,
                          text.get(_W + "val") if text is not None else None, from_style)


# -- references and integrity -----------------------------------------------------------------


def referenced_nums(document: "Document") -> set[int]:
    """Every ``numId`` a paragraph (any story) or a style names."""
    out: set[int] = set()
    parts = document._parts() + ([document.package.styles_part()] if document.package.styles_part() else [])
    for part in parts:
        root = document.package.tree(part)
        if root is None:
            continue
        for node in root.iter(_W + "numId"):
            if node.getparent() is not None and node.getparent().tag == _W + "numPr":
                value = _int(node.get(_W + "val"))
                if value:
                    out.add(value)
    return out


def problems(document: "Document") -> list[tuple[str, str]]:
    """Numbering integrity: ``(code, detail)`` for every numId that does not resolve, every
    instance over a missing abstract definition, and every level out of range."""
    numbering = Numbering(document)
    nums = numbering.nums()
    abstracts = numbering.abstracts()
    out: list[tuple[str, str]] = []
    for num_id in sorted(referenced_nums(document)):
        if num_id not in nums:
            out.append(("numbering-num-missing", str(num_id)))
    for num_id, node in sorted(nums.items()):
        abstract = numbering.abstract_of(num_id)
        if abstract is None or abstract not in abstracts:
            out.append(("numbering-abstract-missing", f"{num_id} -> {abstract}"))
    parts = document._parts() + ([document.package.styles_part()] if document.package.styles_part() else [])
    for part in parts:
        root = document.package.tree(part)
        for node in (root.iter(_W + "ilvl") if root is not None else []):
            value = _int(node.get(_W + "val"))
            if node.getparent().tag == _W + "numPr" and (value is None or not 0 <= value < LEVELS):
                out.append(("numbering-level-out-of-range", f"{part}: {node.get(_W + 'val')}"))
    return out


# -- writing ---------------------------------------------------------------------------------


def _ensure_part(document: "Document") -> Element:
    """The numbering part's root, created (as Word declares it) when the document has none."""
    numbering = Numbering(document)
    if numbering.root is not None:
        return numbering.root
    package = document.package
    part = "word/numbering.xml" if not package.has_part("word/numbering.xml") else \
        package.unused_part_name("word/numbering{n}.xml")
    data = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    package.add_part(part, data, NUMBERING_TYPE, override=True)
    package.add_relationship(package.document_part(), REL_NUMBERING, part)
    return package.tree(part)


def _next_id(nodes: dict[int, Element]) -> int:
    return max(nodes, default=0) + 1


def _nsid(document: "Document", seed: str) -> str:
    used = set()
    root = Numbering(document).root
    if root is not None:
        used = {(n.get(_W + "val") or "").upper() for n in root.iter(_W + "nsid")}
    counter = 0
    while True:
        value = hashlib.sha256(f"{seed}\0{counter}".encode()).hexdigest()[:8].upper()
        counter += 1
        if value not in used and value != "00000000":
            return value


def builtin_abstract(kind: str, abstract_id: int, nsid: str, template: str) -> Element:
    """Word's own Bullets or Numbering definition, as a ``w:abstractNum``."""
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    levels = []
    for level in range(LEVELS):
        indent = 720 * (level + 1)
        if kind == "bullet":
            char, font, code = _BULLETS[level % 3]
            levels.append(
                f'<w:lvl w:ilvl="{level}" w:tplc="{code}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
                f'<w:lvlText w:val="{char}"/><w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{indent}" w:hanging="360"/></w:pPr>'
                f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"'
                + (' w:cs="Courier New"' if font == "Courier New" else "")
                + ' w:hint="default"/></w:rPr></w:lvl>')
        else:
            fmt, justification, code = _NUMBERS[level % 3]
            hanging = 180 if justification == "right" else 360
            levels.append(
                f'<w:lvl w:ilvl="{level}" w:tplc="{code}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="%{level + 1}."/><w:lvlJc w:val="{justification}"/>'
                f'<w:pPr><w:ind w:left="{indent}" w:hanging="{hanging}"/></w:pPr></w:lvl>')
    xml = (f'<w:abstractNum {w} w:abstractNumId="{abstract_id}"><w:nsid w:val="{nsid}"/>'
           f'<w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="{template}"/>'
           + "".join(levels) + "</w:abstractNum>")
    return etree.fromstring(xml)


def add_abstract(document: "Document", node: Element) -> None:
    """Insert an abstract definition after the last one (they precede every instance)."""
    root = _ensure_part(document)
    existing = root.findall(_W + "abstractNum")
    if existing:
        existing[-1].addnext(node)
    else:
        first_num = root.find(_W + "num")
        if first_num is not None:
            first_num.addprevious(node)
        else:
            later = root.find(_W + "numIdMacAtCleanup")
            if later is not None:
                later.addprevious(node)
            else:
                root.append(node)


def add_num(document: "Document", abstract_id: int, overrides: dict[int, int] | None = None) -> int:
    """A new instance over ``abstract_id``, after the last one; ``overrides`` are
    ``level -> start``.  Returns its numId."""
    root = _ensure_part(document)
    nums = Numbering(document).nums()
    num_id = _next_id(nums)
    node = make("w:num")
    node.set(qn("w:numId"), str(num_id))
    reference = make("w:abstractNumId")
    reference.set(qn("w:val"), str(abstract_id))
    node.append(reference)
    for level, start in sorted((overrides or {}).items()):
        override = make("w:lvlOverride")
        override.set(qn("w:ilvl"), str(level))
        start_node = make("w:startOverride")
        start_node.set(qn("w:val"), str(start))
        override.append(start_node)
        node.append(override)
    existing = root.findall(_W + "num")
    if existing:
        existing[-1].addnext(node)
    else:
        later = root.find(_W + "numIdMacAtCleanup")
        if later is not None:
            later.addprevious(node)
        else:
            root.append(node)
    return num_id


def add_restart(document: "Document", abstract_id: int, num_id: int, level: int, start: int,
                style_id: str | None) -> int:
    """A new instance over ``abstract_id`` (the definition of ``num_id``, the style's
    instance) that restarts ``level`` at ``start``, for
    an item its style (``style_id``) numbers: Word's ``w:startOverride``; and where the
    style's indent differs from the level's, the whole level in the override with the
    style's indent in its ``w:pPr`` -- direct numbering takes the level's indent, so this is
    what keeps the item at its style's (round 24, W7: Word draws it there, keeps the form
    on saving, and the items after it count on from it; a level override without the
    ``w:startOverride`` does not restart).  Returns the new numId."""
    numbering = Numbering(document)
    new_id = add_num(document, abstract_id, {level: start})
    style_ind = _style_indent(document, style_id)
    if style_ind is None or indent_note(document, None, num_id, level, style_ind=style_ind) is None:
        return new_id
    abstract = numbering.abstracts().get(numbering.canonical(abstract_id))
    source = next((lvl for lvl in abstract.findall(_W + "lvl") if _int(lvl.get(_W + "ilvl")) == level),
                  None) if abstract is not None else None
    if source is None:
        return new_id
    copied = copy.deepcopy(source)
    # The level's own start too: with a whole level in the override, Word counts from the
    # level's w:start, not the w:startOverride beside it (measured).
    first = copied.find(_W + "start")
    if first is None:
        first = make("w:start")
        copied.insert(0, first)            # a level's first child (ECMA-376 17.9.6)
    first.set(qn("w:val"), str(start))
    properties = copied.find(_W + "pPr")
    if properties is None:
        properties = append_in_order(copied, make("w:pPr"))
    old = properties.find(_W + "ind")
    if old is not None:
        remove(old)
    append_in_order(properties, copy.deepcopy(style_ind))
    node = Numbering(document).nums()[new_id]
    override = next(o for o in node.findall(_W + "lvlOverride") if _int(o.get(_W + "ilvl")) == level)
    override.append(copied)
    return new_id


def new_list(document: "Document", kind: str) -> int:
    """A new list of Word's own ``bullet`` or ``number`` kind: an abstract definition of its
    own and an instance over it.  Returns the numId."""
    abstracts = Numbering(document).abstracts()
    abstract_id = max(abstracts, default=-1) + 1
    seed = f"{kind}\0{abstract_id}\0{len(Numbering(document).nums())}"
    node = builtin_abstract(kind, abstract_id, _nsid(document, seed + "nsid"), _nsid(document, seed + "tmpl"))
    add_abstract(document, node)
    return add_num(document, abstract_id)


def set_numbering(paragraph: Element, num_id: int, level: int) -> None:
    """Give a paragraph its own ``w:numPr``."""
    properties = paragraph.find(_W + "pPr")
    if properties is None:
        properties = append_in_order(paragraph, make("w:pPr"))
    numbering = properties.find(_W + "numPr")
    if numbering is None:
        numbering = append_in_order(properties, make("w:numPr"))
    for child in list(numbering):
        if child.tag in (_W + "ilvl", _W + "numId"):
            remove(child)
    ilvl = make("w:ilvl")
    ilvl.set(qn("w:val"), str(level))
    numbering.insert(0, ilvl)
    ref = make("w:numId")
    ref.set(qn("w:val"), str(num_id))
    ilvl.addnext(ref)


def clear_numbering(paragraph: Element) -> None:
    properties = paragraph.find(_W + "pPr")
    numbering = properties.find(_W + "numPr") if properties is not None else None
    if numbering is not None:
        remove(numbering)
        if not len(properties):
            remove(properties)


def reap(document: "Document", candidates: set[int]) -> list[str]:
    """Remove the instances among ``candidates`` nothing names any more, then the abstract
    definitions only they used.  Returns what was removed (``num:3``, ``abstractNum:2``)."""
    numbering = Numbering(document)
    if numbering.root is None:
        return []
    used = referenced_nums(document)
    nums = numbering.nums()
    removed: list[str] = []
    abstracts_touched: set[int] = set()
    for num_id in sorted(candidates):
        if num_id in nums and num_id not in used:
            abstract = numbering.abstract_of(num_id)
            if abstract is not None:
                abstracts_touched.add(abstract)
            remove(nums[num_id])
            removed.append(f"num:{num_id}")
    if not removed:
        return []
    still = {numbering.abstract_of(n) for n in numbering.nums()}
    abstracts = numbering.abstracts()
    for abstract in sorted(abstracts_touched):
        node = abstracts.get(abstract)
        if node is not None and abstract not in still and node.find(_W + "styleLink") is None:
            remove(node)
            removed.append(f"abstractNum:{abstract}")
    document.package.mark_dirty(numbering.part)
    return removed


def copy_abstract(document: "Document", abstract_id: int) -> int:
    """A copy of an abstract definition with an id and a ``w:nsid`` of its own."""
    numbering = Numbering(document)
    original = numbering.abstracts()[abstract_id]
    new_id = max(numbering.abstracts()) + 1
    node = etree.fromstring(etree.tostring(original))
    node.set(_W + "abstractNumId", str(new_id))
    nsid = node.find(_W + "nsid")
    value = _nsid(document, f"copy\0{abstract_id}\0{new_id}")
    if nsid is None:
        nsid = make("w:nsid")
        node.insert(0, nsid)
    nsid.set(_W + "val", value)
    add_abstract(document, node)
    return new_id


# -- a paragraph style linked to a list level (round 29, W12) ---------------------------------


def _styles_by_id(document: "Document") -> dict[str, Element]:
    part = document.package.styles_part()
    root = document.package.tree(part) if part else None
    return {} if root is None else {s.get(_W + "styleId"): s for s in root.findall(_W + "style")}


def definition_of(document: "Document", num_id: int) -> Element | None:
    """The ``w:abstractNum`` whose levels ``num_id`` numbers with: its definition (the first
    with its ``w:nsid``), or, for a definition that only names a list style
    (``w:numStyleLink``), the definition that style's list has."""
    numbering = Numbering(document)
    seen: set[int] = set()
    while num_id not in seen:
        seen.add(num_id)
        abstract = numbering.abstracts().get(numbering.canonical(numbering.abstract_of(num_id)))
        if abstract is None:
            return None
        link = abstract.find(_W + "numStyleLink")
        linked = numbering._style_num(link.get(_W + "val")) if link is not None else None
        if not linked:
            return abstract
        num_id = linked
    return None


def _definition_levels(document: "Document", definition: Element) -> list[tuple[int | None, Element]]:
    """``(ilvl, w:lvl)`` of a definition, and of every level override of an instance over it."""
    out = [(_int(lvl.get(_W + "ilvl")), lvl) for lvl in definition.findall(_W + "lvl")]
    for num_id, num in Numbering(document).nums().items():
        if definition_of(document, num_id) is not definition:
            continue
        for override in num.findall(_W + "lvlOverride"):
            lvl = override.find(_W + "lvl")
            if lvl is not None:
                out.append((_int(override.get(_W + "ilvl")), lvl))
    return out


def defined_levels(document: "Document", num_id: int) -> list[int]:
    """The levels (0-8) the definition ``num_id`` numbers with defines; none without one."""
    definition = definition_of(document, num_id)
    return [] if definition is None else sorted(
        {i for i in (_int(lvl.get(_W + "ilvl")) for lvl in definition.findall(_W + "lvl")) if i is not None})


def list_style(document: "Document", style_id: str) -> int | None:
    """The ``numId`` a paragraph style's own ``w:numPr`` (or its ``basedOn`` chain's) names:
    the list the style is in, linked to a level or not."""
    num_id, _ = _style_numbering(document, style_id)
    return num_id or None


def _drop_style_numbering(node: Element, indent: bool) -> bool:
    """Take a style's own ``w:numPr`` out of its ``w:pPr`` -- and its ``w:ind`` with ``indent``,
    as Word's "Link level to style" does to the style that held a level (measured, round 29:
    ``feedback-round29/measurements-W.md``).  Whether anything changed."""
    properties = node.find(_W + "pPr")
    if properties is None:
        return False
    changed = False
    for tag in ("numPr",) + (("ind",) if indent else ()):
        child = properties.find(_W + tag)
        if child is not None:
            remove(child)
            changed = True
    if not len(properties):
        remove(properties)
    return changed


def link_style_level(document: "Document", style_id: str, num_id: int, level: int) -> list[str]:
    """Link a paragraph style to a level of a list, as Word's "Link level to style" does
    (measured, round 29): the level (the definition's, and an instance's override of it)
    names the style (``w:pStyle``); the style's ``w:pPr`` gets ``w:numPr`` (``w:ilvl``,
    ``w:numId``) and loses its own ``w:ind`` (the level's indent applies); a style that held
    the level before loses its ``w:numPr`` and its ``w:ind`` (its paragraphs are no longer
    numbered); a level the style held before stops naming it.  Inside an edit; ``ValueError``
    when the list has no such level.  Returns the ids of the styles unlinked."""
    definition = definition_of(document, num_id)
    if definition is None or level not in defined_levels(document, num_id):
        raise ValueError(f"list {num_id} has no level {level}")
    levels = _definition_levels(document, definition)
    styles = _styles_by_id(document)
    unlinked: list[str] = []
    for ilvl, lvl in levels:
        tied = lvl.find(_W + "pStyle")
        holder = tied.get(_W + "val") if tied is not None else None
        if ilvl == level and holder != style_id:
            if holder and holder not in unlinked:
                unlinked.append(holder)
            if tied is None:
                tied = append_in_order(lvl, make("w:pStyle"))
            tied.set(qn("w:val"), style_id)
        elif ilvl != level and holder == style_id:
            remove(tied)                         # a style numbers at one level
    for holder in unlinked:
        node = styles.get(holder)
        own = node.find(f"{_W}pPr/{_W}numPr/{_W}numId") if node is not None else None
        theirs = _int(own.get(_W + "val")) if own is not None else None
        # Only a numPr in this list: a style numbered in another keeps its own.
        if node is not None and (own is None or theirs is not None and definition_of(document, theirs) is definition):
            _drop_style_numbering(node, indent=True)
    node = styles[style_id]
    properties = node.find(_W + "pPr")
    if properties is None:
        properties = append_in_order(node, make("w:pPr"))
    old = properties.find(_W + "numPr")
    if old is not None:
        remove(old)
    numbering = append_in_order(properties, make("w:numPr"))
    if level:
        numbering.append(make("w:ilvl", **{"w:val": str(level)}))
    numbering.append(make("w:numId", **{"w:val": str(num_id)}))
    indent = properties.find(_W + "ind")
    if indent is not None:
        remove(indent)
    document.package.mark_dirty(Numbering(document).part)
    document.package.mark_dirty(document.package.styles_part())
    return unlinked


def unlink_style(document: "Document", style_id: str) -> bool:
    """Unlink a paragraph style from the list level it is linked to, as Word's "Link level
    to style" set to none does (measured, round 29): no level names it any more, and the
    style loses its ``w:numPr`` -- and, where a level was linked to it, its ``w:ind``.  A
    style still numbered after that through its ``basedOn`` chain gets ``w:numPr`` with
    ``w:numId`` 0 and no ``w:ilvl``, as Word writes a derived style's "Numbering: None"
    (measured, round 30: ``feedback-round30/measurements-X.md``), so it numbers nothing.
    Inside an edit.  Whether anything changed."""
    numbering = Numbering(document)
    linked = False
    for definition in numbering.abstracts().values():
        for _, lvl in _definition_levels(document, definition):
            tied = lvl.find(_W + "pStyle")
            if tied is not None and tied.get(_W + "val") == style_id:
                remove(tied)
                linked = True
    node = _styles_by_id(document).get(style_id)
    own = node.find(f"{_W}pPr/{_W}numPr/{_W}numId") if node is not None else None
    if not linked and own is not None and own.get(_W + "val") == "0":
        return False                             # numbering none already
    changed = _drop_style_numbering(node, indent=linked) if node is not None else False
    if node is not None and style_list(document, style_id) is not None:
        properties = node.find(_W + "pPr")
        if properties is None:
            properties = append_in_order(node, make("w:pPr"))
        stop = append_in_order(properties, make("w:numPr"))
        stop.append(make("w:numId", **{"w:val": "0"}))
        changed = True
    if linked:
        document.package.mark_dirty(numbering.part)
    if changed:
        document.package.mark_dirty(document.package.styles_part())
    return linked or changed


def style_link(document: "Document", style_id: str) -> dict | None:
    """How a paragraph style numbers its paragraphs, for ``word_styles`` ``describe`` (round
    30, W15): ``{"list": numId, "level": n, "own": bool, "from": style id}`` -- ``own``
    whether the ``w:numId`` is the style's own, ``from`` the style in its ``basedOn`` chain
    that holds it when not; ``{"list": "none"}`` for a style whose own ``w:numId`` 0 stops
    numbering; ``None`` when it numbers nothing (as :func:`style_list`)."""
    styles = _styles_by_id(document)
    node = styles.get(style_id)
    own = node.find(f"{_W}pPr/{_W}numPr/{_W}numId") if node is not None else None
    if own is not None and _int(own.get(_W + "val")) == 0:
        return {"list": "none"}
    found = style_list(document, style_id)
    if found is None:
        return None
    holder, seen = style_id, set()
    while holder and holder not in seen and holder in styles:
        seen.add(holder)
        if styles[holder].find(f"{_W}pPr/{_W}numPr/{_W}numId") is not None:
            break
        based = styles[holder].find(_W + "basedOn")
        holder = based.get(_W + "val") if based is not None else None
    out: dict = {"list": found[0], "level": found[1], "own": holder == style_id}
    if holder != style_id:
        out["from"] = holder
    return out


def inheriting_styles(document: "Document", style_id: str) -> list[str]:
    """The paragraph styles numbered with ``style_id``'s ``w:numPr`` through their
    ``basedOn`` chain (none of their own on the way), as Word numbers them: a style based
    on a style linked to a level numbers at that level too (measured, round 30)."""
    styles = _styles_by_id(document)
    out = []
    for other, node in styles.items():
        if other == style_id or node.get(_W + "type") != "paragraph":
            continue
        seen: set[str] = set()
        current = other
        while current and current not in seen and current in styles:
            seen.add(current)
            if styles[current].find(f"{_W}pPr/{_W}numPr/{_W}numId") is not None:
                break
            based = styles[current].find(_W + "basedOn")
            current = based.get(_W + "val") if based is not None else None
        if current == style_id and style_list(document, other) is not None:
            out.append(other)
    return out


def following_members(document: "Document", part: str, entry: ParagraphEntry, num_id: int) -> list[ParagraphEntry]:
    """The paragraph and those after it in its story that are in the same instance."""
    out = []
    for other in document._index(part).paragraphs[entry.index:]:
        found = membership(document, other.element)
        if found is not None and found.num_id == num_id:
            out.append(other)
    return out


# -- the operations --------------------------------------------------------------------------

_KINDS = ("bullet", "number")


class ListOps:
    """The list half of :class:`docx_agent.Document`."""

    @property
    def numbering(self: "Document") -> Numbering:
        """The document's list definitions and instances (``numbering.xml``)."""
        return Numbering(self)

    def list_of(self: "Document", identifier: str) -> ListMembership | None:
        """The list a paragraph is in (its instance, level, format), or ``None``."""
        return membership(self, self._paragraph_entry(identifier)[1].element)

    def _list_edit(self: "Document", identifier: str, change) -> "EditResult":
        """Run ``change(part, entry, before)`` as one undo step, then remove the instances and
        abstract definitions it left unused."""
        from .document import EditResult

        part, entry = self._paragraph_entry(identifier)
        before = membership(self, entry.element)
        with self._edit():
            renames = self._prepare(part, [entry])
            with self._track_properties([part]):
                candidates = change(part, entry, before) or set()
            if before is not None:
                candidates.add(before.num_id)
            removed = reap(self, candidates)
            self.package.mark_dirty(part)
            if Numbering(self).part is not None and Numbering(self).root is not None:
                self.package.mark_dirty(Numbering(self).part)
        result = EditResult(renames.get(entry.id, entry.id), renamed=renames)
        result.removed += removed
        return result

    def add_to_list(self: "Document", identifier: str, kind: "str | int" = "bullet", *, level: int = 0,
                    continue_previous: bool = True) -> "EditResult":
        """Make a paragraph a list item.  ``kind`` is ``bullet`` or ``number`` (Word's own
        lists), or a numId to join that list instance.  With ``continue_previous`` (the
        default) a paragraph right after an item of a list of the same kind joins that list,
        as Word's buttons do; otherwise, a new list of its own begins."""
        from .document import EditError

        if not isinstance(level, int) or not 0 <= level < LEVELS:
            raise EditError(f"level is 0 to {LEVELS - 1}")
        if isinstance(kind, int) and not isinstance(kind, bool):
            if kind not in Numbering(self).nums():
                raise EditError(f"the document has no list instance {kind}")
        elif kind not in _KINDS:
            raise EditError("kind is 'bullet', 'number' or an existing numId")

        def change(part, entry, before):
            target = kind if isinstance(kind, int) else None
            if target is None and continue_previous:
                previous = _previous_paragraph(self, part, entry)
                found = membership(self, previous.element) if previous is not None else None
                if found is not None and found.kind == kind:
                    target = found.num_id
            if target is None:
                target = new_list(self, kind)
            set_numbering(entry.element, target, level)
            _list_paragraph(self, entry.element, on=True)
            return set()

        return self._list_edit(identifier, change)

    def set_list_level(self: "Document", identifier: str, level: int) -> "EditResult":
        """Move a list item to ``level`` (0-8; ``w:ilvl``)."""
        from .document import EditError

        if not isinstance(level, int) or not 0 <= level < LEVELS:
            raise EditError(f"level is 0 to {LEVELS - 1}")
        part, entry = self._paragraph_entry(identifier)
        found = membership(self, entry.element)
        if found is None:
            raise EditError(f"{entry.id} is not in a list")
        return self._list_edit(identifier, lambda part, entry, before: set_numbering(entry.element, found.num_id, level))

    def restart_numbering(self: "Document", identifier: str, at: int = 1) -> "EditResult":
        """Restart the list's numbering at this paragraph (from ``at``): a new instance over
        the same abstract definition with a ``w:startOverride``, as Word does, given to
        this paragraph and the list's items after it.  An item numbered by its style gets
        it alone, as Word writes a restart there (measured, round 23): the items after it,
        numbered by the style's instance over the same definition, count on from it."""
        from .document import EditError

        if not isinstance(at, int) or not 0 <= at <= 32767:
            raise EditError("at is 0 to 32767")
        part, entry = self._paragraph_entry(identifier)
        found = membership(self, entry.element)
        if found is None:
            raise EditError(f"{entry.id} is not in a list")
        if found.abstract_id is None:
            raise EditError(f"list instance {found.num_id} names no abstract definition")

        def change(part, entry, before):
            members = [entry] if found.from_style else following_members(self, part, entry, found.num_id)
            num_id = add_restart(self, found.abstract_id, found.num_id, found.level, at,
                                 paragraph_style_id(entry.element)) if found.from_style \
                else add_num(self, found.abstract_id, {found.level: at})
            for member in members:
                level = membership(self, member.element).level
                set_numbering(member.element, num_id, level)
            return set()

        result = self._list_edit(identifier, change)
        if found.from_style:
            moved = self.list_of(result.id)
            note = indent_note(self, self._paragraph_entry(result.id)[1].element, moved.num_id, moved.level) \
                if moved is not None else None
            if note is not None:
                result.warnings.append(f"{result.id}: {note}")
        return result

    def continue_numbering(self: "Document", identifier: str, from_id: str | None = None) -> "EditResult":
        """Continue the numbering of an earlier list (``from_id``'s; by default the nearest
        earlier list over the same abstract definition): this paragraph and the items of its
        list after it join that list's instance.  An item with numbering of its own whose
        style numbers it too also continues the nearest earlier list over the style's
        definition; joining a list over the style's definition, an item whose style numbers
        it at its level loses its own ``w:numPr`` and takes the style's numbering, indent
        included (instances over one definition count on from each other)."""
        from .document import EditError

        part, entry = self._paragraph_entry(identifier)
        found = membership(self, entry.element)
        if found is None:
            raise EditError(f"{entry.id} is not in a list")
        numbering = Numbering(self)
        own_style = None if found.from_style else style_list(self, paragraph_style_id(entry.element))
        style_abstract = numbering.canonical(numbering.abstract_of(own_style[0])) if own_style else None
        if from_id is not None:
            source = membership(self, self._paragraph_entry(from_id)[1].element)
            if source is None:
                raise EditError(f"{from_id} is not in a list")
            target = source.num_id
        else:
            target = None
            for other in reversed(self._index(part).paragraphs[:entry.index]):
                other_list = membership(self, other.element)
                if other_list is None or other_list.num_id == found.num_id:
                    continue
                if other_list.abstract_id == found.abstract_id or (
                        style_abstract is not None and numbering.canonical(other_list.abstract_id) == style_abstract):
                    target = other_list.num_id
                    break
            if target is None and own_style is not None:
                style_id = paragraph_style_id(entry.element)
                style = next((s for s in self.styles if s.id == style_id), None) if style_id else \
                    self.styles.default("paragraph")
                name = (style.name or style.id) if style is not None else style_id
                raise EditError(f"{entry.id} has direct numbering (list {found.num_id}) while its style {name!r} "
                                "numbers itself, and no earlier paragraph is in a list over either definition: "
                                "nothing to continue")
            if target is None:
                raise EditError(f"no earlier list over the same definition as {entry.id}'s")
        target_abstract = numbering.canonical(numbering.abstract_of(target))

        def change(part, entry, before):
            for member in following_members(self, part, entry, found.num_id):
                level = membership(self, member.element).level
                by_style = style_list(self, paragraph_style_id(member.element))
                if by_style is not None and by_style[1] == level and \
                        Numbering(self).canonical(Numbering(self).abstract_of(by_style[0])) == target_abstract:
                    clear_numbering(member.element)
                else:
                    set_numbering(member.element, target, level)
            return set()

        return self._list_edit(identifier, change)

    def remove_from_list(self: "Document", identifier: str) -> "EditResult":
        """Take a paragraph out of its list: its own ``w:numPr`` removed, or -- where its
        style puts it in a list -- ``numId`` 0, as Word writes it.  A paragraph in no list
        is as asked: no change, said in ``warnings`` (round 27)."""
        from .document import EditResult

        part, entry = self._paragraph_entry(identifier)
        found = membership(self, entry.element)
        if found is None:
            return EditResult(entry.id, changed=False, warnings=[f"{entry.id} was not in a list"])

        def change(part, entry, before):
            clear_numbering(entry.element)
            _list_paragraph(self, entry.element, on=False)
            if membership(self, entry.element) is not None:
                set_numbering(entry.element, 0, 0)
            return set()

        return self._list_edit(identifier, change)

    def set_list_format(self: "Document", identifier: str, *, level: int | None = None,
                        format: str | None = None, text: str | None = None, start: int | None = None,
                        font: str | None = None, number_at: float | None = None,
                        text_at: float | None = None, follow: str | None = None,
                        tab_at: float | None = None) -> "EditResult":
        """Change a level of a paragraph's list (its own level by default): ``format``
        (``w:numFmt``: ``decimal``, ``lowerLetter``, ``upperRoman``, ``bullet``...),
        ``text`` (``w:lvlText``: ``%1)``, a bullet character), ``start``, the label's
        ``font``.  Copy-on-write: if another instance shares the abstract definition, this
        list gets a copy of its own first, so no other list changes.

        Round 30, W14, as Word's "Define new multilevel list" writes them (measured:
        ``feedback-round30/measurements-X.md``), in points: ``number_at`` (the number's
        "Aligned at") and ``text_at`` ("Text indent at") -- the level's ``w:ind``,
        ``left`` the text's position and ``hanging`` the difference (``firstLine`` where
        the number is right of the text, ``firstLine`` 0 where they are equal); the one not
        given keeps its place.  ``follow`` (``tab``, ``space``, ``nothing``: ``w:suff``,
        absent for a tab), ``tab_at`` ("Add tab stop at": the level's ``num`` tab stop).
        A style linked to the level takes this indent: W12 gives it none of its own."""
        from .document import EditError

        for name, value in (("number_at", number_at), ("text_at", text_at), ("tab_at", tab_at)):
            if value is not None and (isinstance(value, bool) or not isinstance(value, (int, float))
                                      or not -1584 <= value <= 1584):
                raise EditError(f"{name} is in points, -1584 to 1584")
        if follow is not None and follow not in _FOLLOW:
            raise EditError(f"follow is one of {', '.join(_FOLLOW)}")

        part, entry = self._paragraph_entry(identifier)
        found = membership(self, entry.element)
        if found is None:
            raise EditError(f"{entry.id} is not in a list")
        level = found.level if level is None else level
        if not 0 <= level < LEVELS:
            raise EditError(f"level is 0 to {LEVELS - 1}")
        if format is not None and format not in NUMBER_FORMATS:
            raise EditError(f"format is one of {', '.join(sorted(NUMBER_FORMATS))}")
        if start is not None and not (isinstance(start, int) and 0 <= start <= 32767):
            raise EditError("start is 0 to 32767")
        if found.abstract_id is None:
            raise EditError(f"list instance {found.num_id} names no abstract definition")

        def change(part, entry, before):
            numbering = Numbering(self)
            sharing = [n for n in numbering.nums() if n != found.num_id and numbering.abstract_of(n) == found.abstract_id]
            abstract_id = found.abstract_id
            if sharing:
                abstract_id = copy_abstract(self, found.abstract_id)
                numbering.nums()[found.num_id].find(_W + "abstractNumId").set(_W + "val", str(abstract_id))
            abstract = Numbering(self).abstracts()[abstract_id]
            lvl = next((n for n in abstract.findall(_W + "lvl") if _int(n.get(_W + "ilvl")) == level), None)
            if lvl is None:
                raise EditError(f"the list's definition has no level {level}")
            for tag, value in (("w:start", start), ("w:numFmt", format), ("w:lvlText", text)):
                if value is None:
                    continue
                node = lvl.find(qn(tag))
                if node is None:
                    node = append_in_order(lvl, make(tag))
                node.set(qn("w:val"), str(value))
            if font is not None:
                properties = lvl.find(_W + "rPr")
                if properties is None:
                    properties = append_in_order(lvl, make("w:rPr"))
                fonts = properties.find(_W + "rFonts")
                if fonts is None:
                    fonts = append_in_order(properties, make("w:rFonts"))
                for attribute in ("ascii", "hAnsi"):
                    fonts.set(_W + attribute, font)
                fonts.set(_W + "hint", "default")
            if follow is not None:
                suffix = lvl.find(_W + "suff")
                if follow == "tab":
                    if suffix is not None:
                        remove(suffix)
                else:
                    if suffix is None:
                        suffix = append_in_order(lvl, make("w:suff"))
                    suffix.set(_W + "val", follow)
            if number_at is not None or text_at is not None or tab_at is not None:
                _level_positions(lvl, number_at, text_at, tab_at)
            return set()

        return self._list_edit(identifier, change)


#: What follows a level's number (``w:suff``; a tab when absent).
_FOLLOW = ("tab", "space", "nothing")


def _level_positions(lvl: Element, number_at: float | None, text_at: float | None,
                     tab_at: float | None) -> None:
    """A level's number and text positions (``w:ind``) and its tab stop, in points, as
    Word's "Define new multilevel list" writes them (measured, round 30): ``left`` the
    text's, ``hanging`` the number's distance left of it, ``firstLine`` its distance right
    of it, ``firstLine`` 0 where they are one; the tab a ``num`` tab stop in ``w:tabs``."""
    properties = lvl.find(_W + "pPr")
    if properties is None:
        properties = append_in_order(lvl, make("w:pPr"))
    if number_at is not None or text_at is not None:
        indent = properties.find(_W + "ind")
        if indent is None:
            indent = append_in_order(properties, make("w:ind"))

        def twips(*names: str) -> int:
            for name in names:
                value = _int(indent.get(_W + name))
                if value is not None:
                    return value
            return 0

        left = twips("left", "start")
        number = left - twips("hanging") + twips("firstLine")
        left = round(text_at * 20) if text_at is not None else left
        number = round(number_at * 20) if number_at is not None else number
        for name in ("start", "hanging", "firstLine", "hangingChars", "firstLineChars", "startChars", "leftChars"):
            indent.attrib.pop(_W + name, None)
        indent.set(_W + "left", str(left))
        if number < left:
            indent.set(_W + "hanging", str(left - number))
        else:
            indent.set(_W + "firstLine", str(number - left))
    if tab_at is not None:
        tabs = properties.find(_W + "tabs")
        if tabs is None:
            tabs = append_in_order(properties, make("w:tabs"))
        for tab in tabs.findall(_W + "tab"):
            if tab.get(_W + "val") == "num":
                remove(tab)
        tab = make("w:tab", **{"w:val": "num", "w:pos": str(round(tab_at * 20))})
        after = next((t for t in tabs.findall(_W + "tab") if (_int(t.get(_W + "pos")) or 0) > round(tab_at * 20)),
                     None)
        if after is not None:
            after.addprevious(tab)
        else:
            tabs.append(tab)


#: ``ST_NumberFormat`` values Word offers for a list level.
NUMBER_FORMATS = frozenset({
    "decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter", "ordinal", "cardinalText",
    "ordinalText", "decimalZero", "bullet", "none", "decimalEnclosedCircle", "decimalEnclosedParen",
    "decimalEnclosedFullstop",
})


def _list_paragraph(document: "Document", paragraph: Element, *, on: bool) -> None:
    """Word gives a paragraph of the default style that it makes a list item the List
    Paragraph style (measured: ``tools/e1_probe.py``); taking it out of the list here
    returns a List Paragraph paragraph to the default style, undoing that."""
    from . import formatting as _formatting

    styles = document.styles
    properties = paragraph.find(_W + "pPr")
    node = properties.find(_W + "pStyle") if properties is not None else None
    current = node.get(_W + "val") if node is not None else None
    default = styles.default("paragraph")
    if on:
        if current is None or (default is not None and current == default.id):
            style_id = styles._ensure("List Paragraph", "paragraph")
            properties = _formatting.properties_of(paragraph, "w:pPr", create=True)
            _formatting._set_child(properties, "w:pStyle", {"w:val": style_id})
    elif current is not None:
        found = styles.find("List Paragraph", "paragraph")
        if found is not None and found.id == current:
            remove(node)
            if not len(properties):
                remove(properties)


def _previous_paragraph(document: "Document", part: str, entry: ParagraphEntry) -> ParagraphEntry | None:
    """The paragraph right before ``entry`` in the same container, if it is a paragraph."""
    node = entry.element.getprevious()
    while node is not None and not (isinstance(node.tag, str) and node.tag in (_W + "p", _W + "tbl", _W + "sdt")):
        node = node.getprevious()
    if node is None or node.tag != _W + "p":
        return None
    return document._index(part).entry_for(node)
