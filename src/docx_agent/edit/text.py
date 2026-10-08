"""A paragraph's text: read across runs, fields, hyperlinks, content controls and revisions,
and written back keeping per-character formatting.

**The text** is defined once (ROADMAP.md, "Text ranges and anchors"): ``w:t`` characters;
``w:tab`` and ``w:ptab`` -> ``\\t``; a line break (``w:br``, ``w:cr``) -> ``\\v``; a page or
column break -> ``\\f``; ``w:noBreakHyphen`` -> U+2011; ``w:softHyphen`` -> U+00AD; ``w:sym``
-> its character; a field -> its *result* (instructions are never text); a drawing, a note
reference, an equation or any other object -> U+FFFC, one character.

**The view** decides which revisions count: ``current`` (insertions in, deletions out --
Word's "No Markup"), ``original`` (the reverse) or ``markup`` (both).  Containers are walked,
never flattened: ``w:hyperlink``, ``w:smartTag``, ``w:customXml``, ``w:dir``, ``w:bdo``,
``w:fldSimple``, an inline ``w:sdt``'s ``w:sdtContent``, and the revision containers.

**Writing** (:func:`set_text`) is an in-place diff, after pptx-agent's E1 rule: each new
character takes the formatting of the character it replaces, or of the one before it when it
is inserted.  Characters that survive stay in the ``w:t`` they were in, so a run, its
properties, a hyperlink or a field around unchanged text are not rebuilt; only the ``w:t``
nodes whose text changes are rewritten, and only runs the edit empties are removed.
"""

from __future__ import annotations

import copy
import difflib
from dataclasses import dataclass, field

from ..oxml.xml import REVISION_PROPERTY_TAGS, XML_SPACE, Element, qn, remove

VIEWS = ("current", "original", "markup")

LINE_BREAK = "\v"
PAGE_BREAK = "\f"
OBJECT = "￼"
NON_BREAKING_HYPHEN = "‑"
SOFT_HYPHEN = "­"

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

W_P = _W + "p"
W_R = _W + "r"
W_T = _W + "t"
W_PPR = _W + "pPr"
W_RPR = _W + "rPr"

#: Containers whose content is the paragraph's in every view.
_TRANSPARENT = frozenset(_W + name for name in ("hyperlink", "smartTag", "customXml", "dir", "bdo"))
#: Revision containers, and the views that include their content.
_SHOWN_IN = {
    _W + "ins": ("current", "markup"),
    _W + "moveTo": ("current", "markup"),
    _W + "del": ("original", "markup"),
    _W + "moveFrom": ("original", "markup"),
}
#: Run content that is one character of text.
_CHARACTERS = {
    _W + "tab": "\t",
    _W + "ptab": "\t",
    _W + "cr": LINE_BREAK,
    _W + "noBreakHyphen": NON_BREAKING_HYPHEN,
    _W + "softHyphen": SOFT_HYPHEN,
}
#: Run content that is an object: one U+FFFC.
_OBJECTS = frozenset({
    _W + "drawing", _W + "pict", _W + "object", _W + "footnoteReference",
    _W + "endnoteReference", _W + "footnoteRef", _W + "endnoteRef", _W + "contentPart",
    _W + "ruby", _MC + "AlternateContent",
})
#: Paragraph content that is an object.
_INLINE_OBJECTS = frozenset({_M + "oMath", _M + "oMathPara"})


@dataclass
class Atom:
    """One character of a paragraph's text and where it comes from."""

    char: str
    #: ``text`` (a character of a ``w:t``), ``deleted`` (of a ``w:delText``), or ``element``
    #: (the whole of ``node`` is the character: a tab, a break, an object).
    kind: str
    node: Element
    #: Index in ``node.text`` for ``text`` and ``deleted`` atoms.
    offset: int = 0
    #: The ``w:r`` the character is in (``None`` for an equation).
    run: Element | None = None
    #: In a field's result, where Word may recompute it.
    in_field: bool = False
    #: The innermost field the character is in the result of: its ``w:fldChar`` begin, or
    #: its ``w:fldSimple``; ``None`` outside fields.
    field: Element | None = None


@dataclass
class _Walk:
    view: str
    atoms: list[Atom] = field(default_factory=list)
    runs: list[Element] = field(default_factory=list)
    #: The open complex fields, innermost last: ``instruction`` or ``result``.
    fields: list[str] = field(default_factory=list)
    simple_field: int = 0
    #: The open fields' elements (a begin ``w:fldChar`` or a ``w:fldSimple``), innermost last.
    owners: list[Element] = field(default_factory=list)

    def visible(self) -> bool:
        return all(mode == "result" for mode in self.fields)

    def in_field(self) -> bool:
        return bool(self.fields) or self.simple_field > 0

    def owner(self) -> Element | None:
        return self.owners[-1] if self.owners else None


def walk(paragraph: Element, view: str = "current") -> _Walk:
    """The paragraph's characters and runs, in document order, in ``view``."""
    if view not in VIEWS:
        raise ValueError(f"no text view {view!r}; expected one of {', '.join(VIEWS)}")
    state = _Walk(view)
    _container(paragraph, state)
    return state


def atoms(paragraph: Element, view: str = "current") -> list[Atom]:
    return walk(paragraph, view).atoms


def paragraph_text(paragraph: Element, view: str = "current") -> str:
    return "".join(atom.char for atom in walk(paragraph, view).atoms)


def runs(paragraph: Element, view: str = "current") -> list[Element]:
    """Every ``w:r`` of the paragraph in ``view``, in document order -- inside hyperlinks,
    insertions, content controls and fields too.  A text box's paragraphs are their own."""
    return walk(paragraph, view).runs


def run_text(paragraph: Element, run: Element, view: str = "current") -> str:
    return "".join(atom.char for atom in walk(paragraph, view).atoms if atom.run is run)


def _container(node: Element, state: _Walk) -> None:
    for child in node:
        tag = child.tag
        if not isinstance(tag, str):
            continue
        if tag == W_R:
            _run(child, state)
        elif tag in _TRANSPARENT:
            _container(child, state)
        elif tag == _W + "fldSimple":
            state.simple_field += 1
            state.owners.append(child)
            _container(child, state)
            state.owners.pop()
            state.simple_field -= 1
        elif tag == _W + "sdt":
            content = child.find(_W + "sdtContent")
            if content is not None:
                _container(content, state)
        elif tag in _SHOWN_IN:
            if state.view in _SHOWN_IN[tag]:
                _container(child, state)
        elif tag in _INLINE_OBJECTS:
            if state.visible():
                state.atoms.append(Atom(OBJECT, "element", child, in_field=state.in_field(),
                                        field=state.owner()))


def _run(run: Element, state: _Walk) -> None:
    state.runs.append(run)
    for child in run:
        tag = child.tag
        if not isinstance(tag, str) or tag == W_RPR:
            continue
        if tag == _W + "fldChar":
            kind = child.get(_W + "fldCharType")
            if kind == "begin":
                state.fields.append("instruction")
                state.owners.append(child)
            elif kind == "separate" and state.fields:
                state.fields[-1] = "result"
            elif kind == "end" and state.fields:
                state.fields.pop()
                state.owners.pop()
            continue
        if not state.visible():
            continue
        in_field = state.in_field()
        owner = state.owner()
        if tag == W_T:
            for index, char in enumerate(child.text or ""):
                state.atoms.append(Atom(char, "text", child, index, run, in_field, owner))
        elif tag == _W + "delText":
            if state.view != "current":
                for index, char in enumerate(child.text or ""):
                    state.atoms.append(Atom(char, "deleted", child, index, run, in_field, owner))
        elif tag == _W + "br":
            char = PAGE_BREAK if child.get(_W + "type") in ("page", "column") else LINE_BREAK
            state.atoms.append(Atom(char, "element", child, 0, run, in_field, owner))
        elif tag in _CHARACTERS:
            state.atoms.append(Atom(_CHARACTERS[tag], "element", child, 0, run, in_field, owner))
        elif tag == _W + "sym":
            state.atoms.append(Atom(_symbol(child), "element", child, 0, run, in_field, owner))
        elif tag in _OBJECTS:
            state.atoms.append(Atom(OBJECT, "element", child, 0, run, in_field, owner))


def _symbol(node: Element) -> str:
    try:
        return chr(int(node.get(_W + "char") or "", 16))
    except (ValueError, OverflowError):
        return OBJECT


# -- writing ---------------------------------------------------------------------------------

#: Characters below this length that happen to match inside a changed stretch are folded
#: into the change, so a rewrite does not keep stray letters in their old runs (pptx-agent's
#: rule and threshold).
_MIN_KEPT = 3

#: Characters a written text spells as an element rather than in a ``w:t``.
_SPECIAL = {
    "\t": ("w:tab", {}),
    LINE_BREAK: ("w:br", {}),
    PAGE_BREAK: ("w:br", {"w:type": "page"}),
    NON_BREAKING_HYPHEN: ("w:noBreakHyphen", {}),
    SOFT_HYPHEN: ("w:softHyphen", {}),
}


def check_text(text: str) -> None:
    if "\n" in text or "\r" in text:
        raise ValueError("a paragraph's text cannot contain '\\n'; insert a paragraph, or use "
                         "'\\v' for a line break")


def set_text(paragraph: Element, new_text: str) -> bool:
    """Give ``paragraph`` the text ``new_text`` in the current view, keeping per-character
    formatting.  Returns whether anything changed.

    An object (U+FFFC) can be kept or deleted but not created: the new text may hold one
    only where the old text had one, unchanged.
    """
    check_text(new_text)
    items = atoms(paragraph, "current")
    old_text = "".join(atom.char for atom in items)
    if old_text == new_text:
        return False

    deleted: set[int] = set()
    #: (atom index, "before" | "after", text), or (atom index | None, "run", text)
    inserts: list[tuple[int | None, str, str]] = []
    for tag, i1, i2, j1, j2 in coarse_opcodes(old_text, new_text):
        if tag == "equal":
            continue
        deleted.update(range(i1, i2))
        text = new_text[j1:j2]
        if not text:
            continue
        if OBJECT in text:
            raise ValueError("new text cannot create an object (U+FFFC); only keep one")
        inserts.append(_anchor(items, i1, i2) + (text,))

    _apply(paragraph, items, deleted, inserts)
    return True


def _plain(atom: Atom) -> bool:
    return atom.kind == "text" and not atom.in_field


def _anchor(items: list[Atom], i1: int, i2: int) -> tuple[int | None, str]:
    """Where new text for the old stretch ``[i1, i2)`` goes: the first character it replaces
    when that is text, else after the character before the stretch, else before the one
    after it -- preferring text that is not a field's result -- else a run of its own."""
    for k in range(i1, i2):
        if _plain(items[k]):
            return k, "before"
    for k in range(i1, i2):
        if items[k].kind == "text":
            return k, "before"
    before = i1 - 1 if i1 > 0 else None
    after = i2 if i2 < len(items) else None
    for test in (_plain, lambda atom: atom.kind == "text"):
        if before is not None and test(items[before]):
            return before, "after"
        if after is not None and test(items[after]):
            return after, "before"
    # No text to join: a new run, after the character before or before the one after.
    if before is not None:
        return before, "run-after"
    if after is not None:
        return after, "run-before"
    return None, "run-after"


def _apply(paragraph: Element, items: list[Atom], deleted: set[int],
           inserts: list[tuple[int | None, str, str]]) -> None:
    before: dict[int, list[str]] = {}
    after: dict[int, list[str]] = {}
    new_runs: list[tuple[int | None, str, str]] = []
    for index, side, text in inserts:
        if side == "before":
            before.setdefault(index, []).append(text)  # type: ignore[arg-type]
        elif side == "after":
            after.setdefault(index, []).append(text)  # type: ignore[arg-type]
        else:
            new_runs.append((index, side, text))

    touched_runs: list[Element] = []

    # Rewrite each w:t whose text changes, all at once from its atoms.
    by_node: dict[int, list[int]] = {}
    for index, atom in enumerate(items):
        if atom.kind == "text":
            by_node.setdefault(id(atom.node), []).append(index)
    for indices in by_node.values():
        changed = any(k in deleted or k in before or k in after for k in indices)
        if not changed:
            continue
        node = items[indices[0]].node
        pieces: list[str] = []
        for k in indices:
            pieces.extend(before.get(k, ()))
            if k not in deleted:
                pieces.append(items[k].char)
            pieces.extend(after.get(k, ()))
        # Characters of the node that are not atoms (none, in a well-formed run) stay.
        _replace_text_node(node, "".join(pieces))
        touched_runs.append(items[indices[0]].run)  # type: ignore[arg-type]

    # Remove deleted tabs, breaks and objects.
    for k in sorted(deleted):
        atom = items[k]
        if atom.kind == "element":
            remove(atom.node)
            if atom.run is not None:
                touched_runs.append(atom.run)

    # Text with no text beside it gets a run of its own.
    for index, side, text in new_runs:
        reference = items[index] if index is not None else None
        _insert_run(paragraph, reference, side, text)

    seen: set[int] = set()
    for run in touched_runs:
        if run is None or id(run) in seen:
            continue
        seen.add(id(run))
        if all(not isinstance(child.tag, str) or child.tag == W_RPR for child in run):
            remove(run)


def _replace_text_node(node: Element, text: str) -> None:
    """Replace a ``w:t`` with the elements that spell ``text`` -- ``w:t`` for plain stretches,
    ``w:tab``/``w:br``/... for the special characters -- in the same run."""
    elements = _spell(text, template=node)
    for element in elements:
        node.addprevious(element)
    remove(node)


def _spell(text: str, template: Element | None = None) -> list[Element]:
    out: list[Element] = []
    plain: list[str] = []

    def flush() -> None:
        if plain:
            out.append(_t("".join(plain), template))
            plain.clear()

    for char in text:
        if char in _SPECIAL:
            flush()
            tag, attributes = _SPECIAL[char]
            element = _make(tag)
            for name, value in attributes.items():
                element.set(qn(name), value)
            out.append(element)
        else:
            plain.append(char)
    flush()
    return out


def _make(tag: str) -> Element:
    from lxml import etree

    return etree.Element(qn(tag))


def _t(text: str, template: Element | None) -> Element:
    node = _make("w:t")
    if template is not None:
        for name, value in template.attrib.items():
            node.set(name, value)
    node.text = text
    if text != text.strip() or "  " in text:
        node.set(XML_SPACE, "preserve")
    return node


def run_properties_for(paragraph: Element, reference: Element | None) -> Element | None:
    """A copy of the run properties new text near ``reference`` (a run) takes: the run's own,
    else the paragraph mark's without its revision marks."""
    if reference is not None:
        properties = reference.find(qn("w:rPr"))
        return copy.deepcopy(properties) if properties is not None else None
    mark = paragraph.find(W_PPR)
    mark = mark.find(W_RPR) if mark is not None else None
    if mark is None:
        return None
    properties = copy.deepcopy(mark)
    for child in list(properties):
        if child.tag in REVISION_PROPERTY_TAGS:
            remove(child)
    return properties if len(properties) else None


def make_run(text: str, properties: Element | None) -> Element:
    run = _make("w:r")
    if properties is not None:
        run.append(properties)
    for element in _spell(text):
        run.append(element)
    return run


def _insert_run(paragraph: Element, reference: Atom | None, side: str, text: str) -> None:
    source = reference.run if reference is not None else None
    run = make_run(text, run_properties_for(paragraph, source))
    if reference is None:
        paragraph.append(run)  # an empty paragraph: after the properties and any markers
        return
    anchor = reference.run if reference.run is not None else reference.node
    if side == "run-after":
        anchor.addnext(run)
    else:
        anchor.addprevious(run)


def coarse_opcodes(old: str, new: str) -> list[tuple[str, int, int, int, int]]:
    """Character opcodes with coincidental short matches folded into the edits around them
    (pptx-agent's ``_coarse_opcodes``)."""
    raw = difflib.SequenceMatcher(None, old, new, autojunk=False).get_opcodes()
    result: list[list] = []
    for index, (tag, i1, i2, j1, j2) in enumerate(raw):
        interior = 0 < index < len(raw) - 1
        if tag == "equal" and interior and (i2 - i1) < _MIN_KEPT:
            tag = "replace"
        if tag != "equal" and result and result[-1][0] != "equal":
            result[-1][2] = i2
            result[-1][4] = j2
            continue
        result.append([("equal" if tag == "equal" else "replace"), i1, i2, j1, j2])
    return [tuple(op) for op in result]  # type: ignore[misc]
