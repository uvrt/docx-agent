"""Styles: found by name, applied first; direct formatting second.

A style is named by what users say -- its ``w:name`` ("heading 1", "Quote", "Intense
Emphasis"), compared without case, or one of its ``w:aliases`` -- and only then by its
``w:styleId``, because ids are localised: a Dutch template's "heading 1" has the id
``Kop1``, and asking for ``"Heading 1"`` must find it.  A built-in style the document does
not define yet (it is only a latent style) is written as Word writes it on first use
(:mod:`docx_agent.edit.builtin_styles`, measured); a name that is neither is refused.

``doc.styles`` (a property: ``for style in doc.styles``) lists and finds styles;
``Styles.add`` creates one (based on another, with run and paragraph formatting);
``Styles.modify`` changes one's formatting; ``Styles.remove`` removes one (refused while it
is used, unless given a replacement) and ``Styles.purge_unused`` every unused one.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, append_in_order, make, qn, remove
from . import formatting as _formatting
from .errors import EditError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
KINDS = ("paragraph", "character", "table", "numbering")


class StyleError(EditError):
    """No such style, or one of the wrong kind."""


@dataclass(frozen=True)
class Style:
    """One ``w:style`` as the document defines it."""

    id: str
    name: str | None
    kind: str
    based_on: str | None
    default: bool
    #: ``w:aliases``: other names the style answers to.
    aliases: tuple[str, ...] = ()
    #: ``w:link``: the paired character (or paragraph) style.
    link: str | None = None


def _style(node: Element) -> Style:
    def value(tag: str) -> str | None:
        child = node.find(_W + tag)
        return child.get(_W + "val") if child is not None else None

    aliases = tuple(a.strip() for a in (value("aliases") or "").split(",") if a.strip())
    default = (node.get(_W + "default") or "0").lower() in ("1", "true", "on")
    return Style(node.get(_W + "styleId") or "", value("name"), node.get(_W + "type") or "paragraph",
                 value("basedOn"), default, aliases, value("link"))


class Styles:
    """The document's styles (``styles.xml``), found by name or id."""

    def __init__(self, document: "Document") -> None:
        self._document = document

    def _root(self) -> Element | None:
        part = self._document.package.styles_part()
        return self._document.package.tree(part) if part is not None else None

    def _nodes(self) -> list[Element]:
        root = self._root()
        return [] if root is None else root.findall(_W + "style")

    def __iter__(self):
        return iter([_style(node) for node in self._nodes()])

    def __len__(self) -> int:
        return len(self._nodes())

    def __call__(self) -> list[Style]:
        """``doc.styles()``: the styles as a list (``doc.styles`` is a property, iterable)."""
        return list(self)

    def __getitem__(self, name: str) -> Style:
        """``doc.styles["Heading 1"]``: :meth:`get`."""
        return self.get(name)

    def _node(self, style_id: str) -> Element | None:
        return next((node for node in self._nodes() if node.get(_W + "styleId") == style_id), None)

    def find(self, name: str, kind: str | None = None) -> Style | None:
        """The style ``name`` names -- by ``w:name`` (any case), then an alias, then its id --
        among those of ``kind``; ``None`` if the document defines none."""
        styles = [s for s in self if kind is None or s.kind == kind]
        wanted = name.casefold()
        for test in (lambda s: (s.name or "").casefold() == wanted,
                     lambda s: any(a.casefold() == wanted for a in s.aliases),
                     lambda s: s.id == name,
                     lambda s: s.id.casefold() == wanted):
            for style in styles:
                if test(style):
                    return style
        return None

    def get(self, name: str, kind: str | None = None) -> Style:
        found = self.find(name, kind)
        if found is None:
            raise StyleError(f"the document has no {kind + ' ' if kind else ''}style {name!r}")
        return found

    def default(self, kind: str = "paragraph") -> Style | None:
        return next((s for s in self if s.kind == kind and s.default), None)

    def latent(self, name: str) -> bool:
        """Whether ``name`` is one of the document's latent (built-in, undefined) styles."""
        root = self._root()
        latent = root.find(_W + "latentStyles") if root is not None else None
        if latent is None:
            return False
        wanted = name.casefold()
        return any((node.get(_W + "name") or "").casefold() == wanted for node in latent.findall(_W + "lsdException"))

    # -- resolving for an edit ---------------------------------------------------------------

    def resolve(self, name: str, kind: str) -> tuple[str, list[str]]:
        """The id of the ``kind`` style ``name`` names, and the ids of the styles adding it
        would write first (a built-in one the document does not define, and what it points
        at).  Nothing changes; :meth:`_ensure` writes them."""
        found = self.find(name, kind)
        if found is not None:
            return found.id, []
        other = self.find(name)
        if other is not None:
            raise StyleError(f"{name!r} is a {other.kind} style, not a {kind} style")
        from .builtin_styles import definition_chain

        chain = definition_chain(name, kind, self)
        if chain is None:
            raise StyleError(f"the document has no {kind} style {name!r}, and Word has no built-in one by that name")
        return chain[0][0], [style_id for style_id, _ in chain]

    def _ensure(self, name: str, kind: str) -> str:
        """The style's id, writing a built-in definition the document lacks (inside an edit)."""
        found = self.find(name, kind)
        if found is not None:
            return found.id
        from .builtin_styles import write

        self.resolve(name, kind)
        return write(name, kind, self)

    def _append(self, node: Element) -> None:
        root = self._root()
        if root is None:
            root = self._document._create_styles_part()
        root.append(node)
        self._document.package.mark_dirty(self._document.package.styles_part())

    # -- editing -----------------------------------------------------------------------------

    def add(self, name: str, kind: str = "paragraph", *, based_on: str | None = None,
            next_style: str | None = None, **formatting) -> "EditResult":
        """A new style named ``name`` (refused if the name is taken), based on ``based_on``
        (a name), with run formatting (and, for a paragraph style, paragraph formatting)
        as :mod:`docx_agent.edit.formatting` takes it.  Its id is the name without spaces
        or punctuation, numbered if taken."""
        from .document import EditError, EditResult

        if kind not in ("paragraph", "character"):
            raise EditError("add() makes paragraph and character styles")
        if not name.strip() or len(name) > 253:
            raise EditError("a style name is 1 to 253 characters")
        if self.find(name) is not None:
            raise EditError(f"the document already has a style {name!r}")
        if based_on is not None:
            self.resolve(based_on, kind)
        if next_style is not None:
            self.resolve(next_style, "paragraph")
        run_values = {k: v for k, v in formatting.items() if k in _formatting.RUN_PROPERTIES}
        paragraph_values = {k: v for k, v in formatting.items() if k not in _formatting.RUN_PROPERTIES}
        try:
            _formatting.check_run(run_values)
            if paragraph_values and kind != "paragraph":
                raise _formatting.FormattingError("a character style has no paragraph formatting")
            _formatting.check_paragraph(paragraph_values)
        except _formatting.FormattingError as error:
            raise EditError(str(error)) from None
        document = self._document
        with document._edit():
            base = self._ensure(based_on, kind) if based_on is not None else None
            following = self._ensure(next_style, "paragraph") if next_style is not None else None
            style_id = self._new_id(name)
            node = make("w:style")
            node.set(qn("w:type"), kind)
            node.set(qn("w:customStyle"), "1")
            node.set(qn("w:styleId"), style_id)
            append_in_order(node, make("w:name", **{"w:val": name}))
            if base is not None:
                append_in_order(node, make("w:basedOn", **{"w:val": base}))
            if following is not None:
                append_in_order(node, make("w:next", **{"w:val": following}))
            append_in_order(node, make("w:qFormat"))
            _formatting.write_paragraph(node, paragraph_values)
            _formatting.write_run(node, run_values, document._theme_color)
            self._append(node)
        return EditResult(f"style:{style_id}", created=[f"style:{style_id}"])

    def modify(self, name: str, kind: str | None = None, **formatting) -> "EditResult":
        """Change a style's run and paragraph formatting (``None`` removes a value)."""
        from .document import EditError, EditResult

        style = self.find(name, kind)
        if style is None:
            raise EditError(f"the document has no style {name!r}")
        run_values = {k: v for k, v in formatting.items() if k in _formatting.RUN_PROPERTIES}
        paragraph_values = {k: v for k, v in formatting.items() if k not in _formatting.RUN_PROPERTIES}
        try:
            _formatting.check_run(run_values)
            _formatting.check_paragraph(paragraph_values)
        except _formatting.FormattingError as error:
            raise EditError(str(error)) from None
        if paragraph_values and style.kind not in ("paragraph", "table"):
            raise EditError(f"{name!r} is a {style.kind} style: it has no paragraph formatting")
        document = self._document
        with document._edit():
            node = self._node(style.id)
            _formatting.write_paragraph(node, paragraph_values)
            _formatting.write_run(node, run_values, document._theme_color)
            for tag in ("w:pPr", "w:rPr"):
                child = node.find(qn(tag))
                if child is not None and not len(child):
                    remove(child)
            document.package.mark_dirty(document.package.styles_part())
        return EditResult(f"style:{style.id}")

    # -- removing ---------------------------------------------------------------------------

    def usage(self, name: str, kind: str | None = None) -> dict[str, int]:
        """Where a style is used: ``{"content": n, "styles": n}`` -- references from the
        document's paragraphs, runs, tables and list levels (every story, numbering too), and
        from other styles (``w:basedOn``, ``w:next``, ``w:link``)."""
        style = self.get(name, kind)
        return {"content": len(self._content_references(style.id)),
                "styles": len(self._style_references(style.id))}

    def _content_references(self, style_id: str) -> list[tuple[str, Element]]:
        package = self._document.package
        styles_part = package.styles_part()
        out: list[tuple[str, Element]] = []
        for part in package.part_names:
            if part == styles_part or not part.startswith("word/") or not part.endswith(".xml"):
                continue
            root = package.tree(part)
            if root is None:
                continue
            for node in root.iter(_W + "pStyle", _W + "rStyle", _W + "tblStyle", _W + "numStyleLink",
                                  _W + "styleLink"):
                if node.get(_W + "val") == style_id:
                    out.append((part, node))
        return out

    def _style_references(self, style_id: str) -> list[Element]:
        return [node for style in self._nodes() if style.get(_W + "styleId") != style_id
                for node in style.iter(_W + "basedOn", _W + "next", _W + "link", _W + "pStyle", _W + "rStyle",
                                       _W + "tblStyle", _W + "numStyleLink", _W + "styleLink")
                if node.get(_W + "val") == style_id]

    def _used(self, node: Element) -> bool:
        """Whether a style is used by content, or by another style otherwise than a link."""
        style_id = node.get(_W + "styleId") or ""
        return bool(self._content_references(style_id)) or any(
            ref.tag != _W + "link" for ref in self._style_references(style_id)) or \
            (node.get(_W + "default") or "0") in ("1", "true", "on")

    def remove(self, name: str, kind: str | None = None, *, replacement: str | None = None) -> "EditResult":
        """Remove a style's definition (found as :meth:`find` finds it), one undo step.

        Refused while the document uses it -- a paragraph, run, table or list level names
        it -- unless ``replacement`` (a style of the same kind) is given: every use then
        names the replacement.  Other styles based on it are based on what it was based on;
        one that names it as its next style or its link loses that.  The default style of
        its kind cannot be removed::

            doc.styles.remove("Bulletin Body", replacement="Handbook Body")"""
        from .document import EditError, EditResult

        style = self.find(name, kind)
        if style is None:
            raise StyleError(f"the document has no {kind + ' ' if kind else ''}style {name!r}")
        if style.default:
            raise EditError(f"{style.name!r} is the document's default {style.kind} style: it cannot be removed")
        target = None
        if replacement is not None:
            found = self.find(replacement, style.kind)
            if found is None and self.find(replacement) is None and not self.latent(replacement):
                from .builtin_styles import definition_chain

                if definition_chain(replacement, style.kind, self) is None:
                    raise StyleError(f"the document has no {style.kind} style {replacement!r}")
            if found is not None and found.id == style.id:
                raise EditError("a style cannot replace itself")
        uses = self._content_references(style.id)
        if uses and replacement is None:
            raise EditError(f"{style.name!r} is used {len(uses)} time(s): give replacement= (a {style.kind} "
                            "style) to point those at")
        document = self._document
        with document._edit():
            if replacement is not None:
                target = self._ensure(replacement, style.kind)
            node = self._node(style.id)
            based = node.find(_W + "basedOn")
            based_on = based.get(_W + "val") if based is not None else None
            for part, use in self._content_references(style.id):
                use.set(_W + "val", target)
                document.package.mark_dirty(part)
            for ref in self._style_references(style.id):
                if ref.tag == _W + "basedOn" and based_on is not None:
                    ref.set(_W + "val", based_on)
                elif ref.tag in (_W + "pStyle", _W + "rStyle", _W + "tblStyle") and target is not None:
                    ref.set(_W + "val", target)
                else:
                    remove(ref)
            remove(node)
            document.package.mark_dirty(document.package.styles_part())
        return EditResult(None, removed=[f"style:{style.id}"],
                          warnings=[f"{len(uses)} use(s) now name {replacement!r}"] if uses else [])

    def purge_unused(self, *, builtin: bool = False) -> "EditResult":
        """Remove every style the document does not use -- no paragraph, run, table or list
        level names it, and no style kept is based on it, follows it or links to it --
        except the defaults; only the document's own (custom) styles unless ``builtin=True``
        (Word's built-in ones are harmless and Word lists them anyway).  One undo step;
        ``removed`` names them (``style:<id>``)."""
        from .document import EditResult

        document = self._document
        removed: list[str] = []

        def purgeable(node: Element) -> bool:
            return (node.get(_W + "default") or "0") not in ("1", "true", "on") and (
                builtin or (node.get(_W + "customStyle") or "0") in ("1", "true", "on"))

        with document._edit():
            changed = True
            while changed:
                changed = False
                for node in list(self._nodes()):
                    style = _style(node)
                    if not purgeable(node):
                        continue
                    if self._content_references(style.id):
                        continue
                    references = self._style_references(style.id)
                    if any(ref.tag != _W + "link" for ref in references):
                        continue
                    # A linked pair (a paragraph style and its "... Char") goes together:
                    # kept while the other one is used.
                    if any(self._used(ref.getparent()) or not purgeable(ref.getparent()) for ref in references):
                        continue
                    for ref in self._style_references(style.id):
                        remove(ref)
                    remove(node)
                    removed.append(f"style:{style.id}")
                    changed = True
            if removed:
                document.package.mark_dirty(document.package.styles_part())
        return EditResult(None, removed=removed, changed=bool(removed))

    def _new_id(self, name: str) -> str:
        base = re.sub(r"[^0-9A-Za-z]", "", name) or "Style"
        taken = {s.id.casefold() for s in self}
        candidate, number = base, 1
        while candidate.casefold() in taken:
            number += 1
            candidate = f"{base}{number}"
        return candidate

    def __repr__(self) -> str:
        return f"<Styles {len(self)}>"
