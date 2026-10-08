"""Formatting operations on the document: styles first, then direct formatting.

Each is one undo step, checks its values before it changes anything, stamps per the
document's policy, and returns an :class:`EditResult`.  Formatting a range splits runs at
its edges (:func:`docx_agent.edit.inline.isolate`), writes the runs inside, and merges
what the split left once the runs on either side came out alike, so no edit leaves empty or
redundant runs.  A range across paragraphs formats each paragraph's part of it and the
marks of every paragraph but the last, as Word does for a selection across paragraph ends.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..oxml.xml import Element, append_in_order, make, qn, remove
from . import formatting as _formatting
from . import inline as _inline
from . import text as _text
from .ids import ParagraphEntry
from .ranges import TextRange
from .styles import StyleError, Styles

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

#: Word's theme colour names -> the colour scheme's slots (the default mapping; a document's
#: ``w:clrSchemeMapping`` may say otherwise for the text and background names).
_SCHEME = {
    "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2", "accent1": "accent1",
    "accent2": "accent2", "accent3": "accent3", "accent4": "accent4", "accent5": "accent5",
    "accent6": "accent6", "hyperlink": "hlink", "followedHyperlink": "folHlink",
    "text1": "dk1", "background1": "lt1", "text2": "dk2", "background2": "lt2",
}
_MAPPING = {"text1": "t1", "background1": "bg1", "text2": "t2", "background2": "bg2"}
_MAPPED = {"dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2"}


class FormatOps:
    """The formatting half of :class:`docx_agent.Document`."""

    @property
    def styles(self: "Document") -> Styles:
        """The document's styles, by name: ``doc.styles["Heading 1"]``."""
        return Styles(self)

    def _theme_color(self: "Document", name: str) -> str | None:
        """The theme's value for a theme colour name, as Word writes it in ``w:val``."""
        main = self.package.document_part()
        parts = self.package.related_parts_of_type(main, _THEME)
        root = self.package.tree(parts[0]) if parts else None
        scheme = root.find(f"{_A}themeElements/{_A}clrScheme") if root is not None else None
        if scheme is None:
            return None
        slot = _SCHEME[name]
        if name in _MAPPING:
            settings = self.package.settings_part()
            mapping = self.package.tree(settings).find(_W + "clrSchemeMapping") if settings else None
            if mapping is not None and mapping.get(_W + _MAPPING[name]):
                slot = _MAPPED.get(mapping.get(_W + _MAPPING[name]), slot)
        node = scheme.find(_A + slot)
        if node is None:
            return None
        for child in node:
            if child.tag == _A + "srgbClr":
                return (child.get("val") or "").upper() or None
            if child.tag == _A + "sysClr":
                return (child.get("lastClr") or "").upper() or None
        return None

    def _create_styles_part(self: "Document") -> Element:
        from ..oxml.package import REL_STYLES

        part = "word/styles.xml"
        data = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        self.package.add_part(part, data, "application/vnd.openxmlformats-officedocument."
                              "wordprocessingml.styles+xml", override=True)
        self.package.add_relationship(self.package.document_part(), REL_STYLES, part)
        return self.package.tree(part)

    # -- styles ------------------------------------------------------------------------------

    def set_paragraph_style(self: "Document", identifier: str, name: str | None) -> "EditResult":
        """Apply the paragraph style ``name`` (found by name, alias or id; a built-in one the
        document lacks is added as Word writes it) to a paragraph; ``None`` or the default
        style removes ``w:pStyle``, as Word writes a Normal paragraph."""
        from .document import EditError, EditResult

        part, entry = self._paragraph_entry(identifier)
        styles = self.styles
        try:
            if name is not None:
                styles.resolve(name, "paragraph")
        except StyleError as error:
            raise EditError(str(error)) from None
        with self._edit():
            renames = self._prepare(part, [entry])
            with self._track_properties([part]):
                style_id = styles._ensure(name, "paragraph") if name is not None else None
                default = styles.default("paragraph")
                properties = entry.element.find(_W + "pPr")
                if style_id is None or (default is not None and style_id == default.id):
                    node = properties.find(_W + "pStyle") if properties is not None else None
                    if node is not None:
                        remove(node)
                        if not len(properties):
                            remove(properties)
                else:
                    properties = _formatting.properties_of(entry.element, "w:pPr", create=True)
                    _formatting._set_child(properties, "w:pStyle", {"w:val": style_id})
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames)

    def set_character_style(self: "Document", where: TextRange, name: str | None) -> "EditResult":
        """Apply the character style ``name`` to a range's runs (``None`` removes it)."""
        from .document import EditError

        styles = self.styles
        try:
            if name is not None:
                styles.resolve(name, "character")
        except StyleError as error:
            raise EditError(str(error)) from None

        def apply(owner: Element) -> None:
            style_id = styles._ensure(name, "character") if name is not None else None
            default = styles.default("character")
            properties = _formatting.properties_of(owner, "w:rPr", create=style_id is not None)
            if properties is None:
                return
            if style_id is None or (default is not None and style_id == default.id):
                _formatting._set_child(properties, "w:rStyle", None)
            else:
                _formatting._set_child(properties, "w:rStyle", {"w:val": style_id})
            if owner.tag == _W + "r" and not len(properties):
                remove(properties)

        return self._format_range(where, apply, marks=False)

    # -- direct formatting -------------------------------------------------------------------

    def format_range(self: "Document", where: TextRange, **values) -> "EditResult":
        """Direct run formatting (:mod:`docx_agent.edit.formatting`) on a range's runs."""
        from .document import EditError

        try:
            _formatting.check_run(values)
        except _formatting.FormattingError as error:
            raise EditError(str(error)) from None
        return self._format_range(where, lambda owner: _formatting.write_run(owner, values, self._theme_color))

    def format_paragraph(self: "Document", identifier: str, **values) -> "EditResult":
        """Direct paragraph formatting; run formatting given too applies to every run of
        the paragraph and to its mark."""
        from .document import EditError, EditResult

        run_values = {k: v for k, v in values.items() if k in _formatting.RUN_PROPERTIES}
        paragraph_values = {k: v for k, v in values.items() if k not in _formatting.RUN_PROPERTIES}
        try:
            _formatting.check_paragraph(paragraph_values)
            _formatting.check_run(run_values)
        except _formatting.FormattingError as error:
            raise EditError(str(error)) from None
        part, entry = self._paragraph_entry(identifier)
        with self._edit():
            renames = self._prepare(part, [entry])
            with self._track_properties([part]):
                _formatting.write_paragraph(entry.element, paragraph_values)
                if run_values:
                    runs = _text.runs(entry.element)
                    for run in runs:
                        _formatting.write_run(run, run_values, self._theme_color)
                    owner = _formatting.properties_of(entry.element, "w:pPr", create=True)
                    _formatting.write_run(owner, run_values, self._theme_color)
                    if not len(owner):
                        remove(owner)
                    _inline.merge_runs(runs)
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames)

    def format_run(self: "Document", identifier: str, index: int, **values) -> "EditResult":
        """Direct formatting on one run, as it is (no split)."""
        from .document import EditError, EditResult

        try:
            _formatting.check_run(values)
        except _formatting.FormattingError as error:
            raise EditError(str(error)) from None
        part, entry = self._paragraph_entry(identifier)
        runs = _text.runs(entry.element)
        if not 0 <= index < len(runs):
            raise EditError(f"{entry.id} has no run r{index}")
        with self._edit():
            renames = self._prepare(part, [entry])
            with self._track_properties([part]):
                _formatting.write_run(_text.runs(entry.element)[index], values, self._theme_color)
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames)

    def clear_formatting(self: "Document", target: str | TextRange, *, paragraph: bool = True) -> "EditResult":
        """Remove direct formatting -- a range's run formatting, or a paragraph's run,
        mark and (``paragraph=True``) paragraph formatting -- keeping styles, list
        membership, language and revision records: how a mess becomes a styled document
        again."""
        from .document import EditResult

        if isinstance(target, TextRange):
            return self._format_range(target, _formatting.clear_run, marks=False)
        part, entry = self._paragraph_entry(target)
        with self._edit():
            renames = self._prepare(part, [entry])
            with self._track_properties([part]):
                for run in _text.runs(entry.element):
                    _formatting.clear_run(run)
                if paragraph:
                    _formatting.clear_paragraph(entry.element)
                else:
                    properties = entry.element.find(_W + "pPr")
                    mark = properties.find(_W + "rPr") if properties is not None else None
                    if mark is not None:
                        _formatting.clear_run(mark)
                _inline.merge_runs(_text.runs(entry.element))
            self.package.mark_dirty(part)
        return EditResult(renames.get(entry.id, entry.id), renamed=renames)

    def _format_range(self: "Document", where: TextRange, apply, *, marks: bool = True) -> "EditResult":
        from .document import EditError, EditResult

        if where.view != "current":
            raise EditError(f"{where.id} was taken in the {where.view} view; edits take the current view")
        part, entries = where._entries()
        spans = []
        for k, entry in enumerate(entries):
            length = len(_text.atoms(entry.element))
            start = where.start if k == 0 else 0
            end = where.end if k == len(entries) - 1 else length
            if not 0 <= start <= end <= length:
                raise EditError(f"{where.id} is outside its paragraphs' text")
            try:
                _inline.check_span(_text.atoms(entry.element), start, end, deleting=False)
            except _inline.SpanError as error:
                raise EditError(f"{where.id}: {error}") from None
            spans.append((entry, start, end))
        with self._edit(), self._track_properties([part]):
            renames = self._prepare(part, entries)
            for k, (entry, start, end) in enumerate(spans):
                runs = _inline.isolate(entry.element, start, end)
                for run in runs:
                    apply(run)
                if marks and k < len(spans) - 1:
                    mark_owner = _formatting.properties_of(entry.element, "w:pPr", create=True)
                    apply(mark_owner)
                    if not len(mark_owner):
                        remove(mark_owner)
                _inline.merge_around(runs)
                _inline.prune(entry.element)
            self.package.mark_dirty(part)
        first = entries[0]
        return EditResult(renames.get(first.id, first.id), renamed=renames)

    def _paragraph_entry(self: "Document", identifier: str) -> tuple[str, ParagraphEntry]:
        from .document import EditError

        part, entry = self._resolve(identifier)
        if not isinstance(entry, ParagraphEntry):
            raise EditError(f"{identifier} is not a paragraph")
        return part, entry
