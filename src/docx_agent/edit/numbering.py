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

import hashlib
from dataclasses import dataclass
from typing import TYPE_CHECKING

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
        root = self.root
        out: dict[int, Element] = {}
        for node in (root.findall(_W + "num") if root is not None else []):
            number = _int(node.get(_W + "numId"))
            if number is not None:
                out[number] = node
        return out

    def abstracts(self) -> dict[int, Element]:
        root = self.root
        out: dict[int, Element] = {}
        for node in (root.findall(_W + "abstractNum") if root is not None else []):
            number = _int(node.get(_W + "abstractNumId"))
            if number is not None:
                out[number] = node
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
    """``(numId, ilvl)`` a paragraph style gives through its ``w:basedOn`` chain."""
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
        this paragraph and the list's items after it."""
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
            members = following_members(self, part, entry, found.num_id)
            num_id = add_num(self, found.abstract_id, {found.level: at})
            for member in members:
                level = membership(self, member.element).level
                set_numbering(member.element, num_id, level)
            return set()

        return self._list_edit(identifier, change)

    def continue_numbering(self: "Document", identifier: str, from_id: str | None = None) -> "EditResult":
        """Continue the numbering of an earlier list (``from_id``'s; by default the nearest
        earlier list over the same abstract definition): this paragraph and the items of its
        list after it join that list's instance."""
        from .document import EditError

        part, entry = self._paragraph_entry(identifier)
        found = membership(self, entry.element)
        if found is None:
            raise EditError(f"{entry.id} is not in a list")
        if from_id is not None:
            source = membership(self, self._paragraph_entry(from_id)[1].element)
            if source is None:
                raise EditError(f"{from_id} is not in a list")
            target = source.num_id
        else:
            target = None
            for other in reversed(self._index(part).paragraphs[:entry.index]):
                other_list = membership(self, other.element)
                if (other_list is not None and other_list.num_id != found.num_id
                        and other_list.abstract_id == found.abstract_id):
                    target = other_list.num_id
                    break
            if target is None:
                raise EditError(f"no earlier list over the same definition as {entry.id}'s")

        def change(part, entry, before):
            for member in following_members(self, part, entry, found.num_id):
                set_numbering(member.element, target, membership(self, member.element).level)
            return set()

        return self._list_edit(identifier, change)

    def remove_from_list(self: "Document", identifier: str) -> "EditResult":
        """Take a paragraph out of its list: its own ``w:numPr`` removed, or -- where its
        style puts it in a list -- ``numId`` 0, as Word writes it."""
        from .document import EditError

        part, entry = self._paragraph_entry(identifier)
        found = membership(self, entry.element)
        if found is None:
            raise EditError(f"{entry.id} is not in a list")

        def change(part, entry, before):
            clear_numbering(entry.element)
            _list_paragraph(self, entry.element, on=False)
            if membership(self, entry.element) is not None:
                set_numbering(entry.element, 0, 0)
            return set()

        return self._list_edit(identifier, change)

    def set_list_format(self: "Document", identifier: str, *, level: int | None = None,
                        format: str | None = None, text: str | None = None, start: int | None = None,
                        font: str | None = None) -> "EditResult":
        """Change a level of a paragraph's list (its own level by default): ``format``
        (``w:numFmt``: ``decimal``, ``lowerLetter``, ``upperRoman``, ``bullet``...),
        ``text`` (``w:lvlText``: ``%1)``, a bullet character), ``start``, the label's
        ``font``.  Copy-on-write: if another instance shares the abstract definition, this
        list gets a copy of its own first, so no other list changes."""
        from .document import EditError

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
            return set()

        return self._list_edit(identifier, change)


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
