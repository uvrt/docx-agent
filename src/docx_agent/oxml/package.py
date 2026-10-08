"""The ``.docx`` package: ooxml-edit's OPC container plus WordprocessingML's entry points.

Reading, editing and losslessly writing parts is :mod:`ooxml_edit.opc`'s; this module only
knows where a Word document keeps its main part, its stories (headers, footers, notes,
comments), its styles, numbering and settings, and which content types make a package a
document, a macro-enabled document or a template.
"""

from __future__ import annotations

import posixpath

from . import xml as _xml  # noqa: F401  registers the w: namespaces before qn() is used
from ooxml_edit.opc import (  # noqa: F401  (re-exported)
    CONTENT_TYPES_PART,
    REL_OFFICE_DOCUMENT,
    OpcPackage,
    Relationship,
    normalize_part_path,
    rels_path_for,
)
from .xml import find, qn
from ooxml_common.kinds import MAIN_CONTENT_TYPES as _ALL_MAIN_CONTENT_TYPES

_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
REL_STYLES = _REL + "styles"
REL_NUMBERING = _REL + "numbering"
REL_SETTINGS = _REL + "settings"
REL_HEADER = _REL + "header"
REL_FOOTER = _REL + "footer"
REL_FOOTNOTES = _REL + "footnotes"
REL_ENDNOTES = _REL + "endnotes"
REL_COMMENTS = _REL + "comments"

#: The main part's content type for each kind of Word package
#: (:data:`ooxml_common.kinds.MAIN_CONTENT_TYPES`, Word's share).
MAIN_CONTENT_TYPES: dict[str, str] = {kind: _ALL_MAIN_CONTENT_TYPES[kind] for kind in ("docx", "dotx", "docm", "dotm")}

#: Relationship types of the parts that hold stories, in the order they are listed.
STORY_RELATIONSHIPS = (REL_HEADER, REL_FOOTER, REL_FOOTNOTES, REL_ENDNOTES, REL_COMMENTS)

#: The compatibility mode a document has when its settings name none: Word 2007's.
DEFAULT_COMPATIBILITY_MODE = 12


class WordPackage(OpcPackage):
    """An open Word package (``.docx``, ``.docm``, ``.dotx``, ``.dotm``)."""

    def document_part(self) -> str:
        """The main document part (``word/document.xml`` in every file Word writes)."""
        main = self.main_document_part()
        if main is not None:
            return main
        if self.has_part("word/document.xml"):
            return "word/document.xml"
        raise ValueError("not a Word package: no main document part")

    @property
    def kind(self) -> str | None:
        """``docx``, ``docm``, ``dotx`` or ``dotm``, from the main part's content type."""
        content_type = self.content_type(self.document_part())
        for kind, known in MAIN_CONTENT_TYPES.items():
            if content_type == known:
                return kind
        return None

    def story_parts(self) -> list[str]:
        """Every part related to the main part that holds a story, in relationship order:
        headers, footers, footnotes, endnotes, comments."""
        main = self.document_part()
        out: list[str] = []
        for rel_type in STORY_RELATIONSHIPS:
            for part in self.related_parts_of_type(main, rel_type):
                if part not in out and self.has_part(part):
                    out.append(part)
        return out

    def _related(self, rel_type: str) -> str | None:
        parts = self.related_parts_of_type(self.document_part(), rel_type)
        return parts[0] if parts and self.has_part(parts[0]) else None

    def styles_part(self) -> str | None:
        return self._related(REL_STYLES)

    def numbering_part(self) -> str | None:
        return self._related(REL_NUMBERING)

    def settings_part(self) -> str | None:
        return self._related(REL_SETTINGS)

    def compatibility_mode(self) -> int:
        """``w:compatSetting w:name="compatibilityMode"`` from the settings, or 12 (Word
        2007's mode) when the document names none."""
        part = self.settings_part()
        root = self.tree(part) if part is not None else None
        compat = find(root, "w:compat")
        if compat is not None:
            for setting in compat.findall(qn("w:compatSetting")):
                if setting.get(qn("w:name")) == "compatibilityMode":
                    try:
                        return int(setting.get(qn("w:val")) or "")
                    except ValueError:
                        break
        return DEFAULT_COMPATIBILITY_MODE


def story_name(package: WordPackage, part: str) -> str:
    """The name a part's story goes by in ids: ``body`` for the main part, else the part's
    stem (``header2``, ``footnotes``, ``comments``)."""
    if part == package.document_part():
        return "body"
    return posixpath.splitext(posixpath.basename(part))[0]
