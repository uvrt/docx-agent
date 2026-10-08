"""Making documents: a new one (blank or from a template), saving one as a template, and
bringing an older one into compatibility mode 15.

Each is measured against Word 16.106 for Mac (``tools/e6_probe.py``,
``tests/observations/e6-word.json``):

**From nothing** the parts are :mod:`.blank`'s: what Word makes for File > New > Blank
Document.

**From a template** -- a ``.dotx`` or a ``.docx``, as a path or bytes -- as Word's File >
New from a template (``create new document attached template``) makes one.  Word keeps
the template's whole package: its styles, numbering, theme, settings, sections, headers
and footers, **and its body**; so does :meth:`Document.new` (``keep_content=False`` drops
the body, keeping the last section).  The main part becomes a document
(``wordprocessingml.document.main``).  From a ``.dotx`` path Word attaches the template
(``w:attachedTemplate`` and an external relationship to ``file:///`` and its path) and
names it in ``app.xml``'s ``Template``; from a ``.docx`` it attaches nothing and the
template stays ``Normal.dotm``.  The core properties start again as Word starts them --
the title, subject, keywords and description kept, the creator and last modifier the
author, revision 1, created and modified now -- and ``app.xml`` keeps the template's company.
**docx-agent attaches no template unless asked** (``attach=True``): measured by the
oracle, Word for Mac asks for access to an attached template outside its sandbox every
time it opens the document (its export blocked on the question), and an attached
macro-enabled template would bring back the macros the ``.docx`` must not carry.
``app.xml`` then names ``Normal.dotm``, as for a document attached to nothing.
A **macro-enabled** template (``.dotm``) or document (``.docm``) gives a ``.docx``: its VBA
project and what belongs to it (``vbaData.xml``, the key-map customisations) are dropped,
with a :class:`MacrosDropped` warning, since a ``.docx`` that carries them does not open.
A template below compatibility mode 15 is converted (:func:`upgrade`), because every
document docx-agent creates is a modern one (ROADMAP.md, "Decisions"), with a
:class:`ModeUpgraded` warning.

**Saving as a template** (:meth:`Document.save_as_template`): Word changes the main part's
content type and nothing else (measured: ``save as ... file format format template``).

**Convert** (:meth:`Document.upgrade_to_modern`, Word's File > Info > Convert, its
``upgrade`` command) changes only ``w:compat`` in the settings (measured on modes 11, 12
and 14, with every one of the 65 legacy options set):

* ``compatibilityMode`` becomes 15, and the settings Word writes beside it are set:
  ``overrideTableStyleFontSizeAndJustification``, ``enableOpenTypeFeatures``,
  ``doNotFlipMirrorIndents``, ``differentiateMultirowTableHeaders`` (1) and
  ``useWord2013TrackBottomHyphenation`` (0);
* of the legacy options, the ten Word 2013's layout still honours stay (:data:`KEPT`) and
  the other 55 go.  ``w:useFELayout``, which Word drops on *any* save of a document whose
  languages are not East Asian, is a save's doing, not Convert's: it stays.
"""

from __future__ import annotations

import copy
import os
import urllib.parse
import warnings
from datetime import datetime
from typing import TYPE_CHECKING

from lxml import etree
from ooxml_common.kinds import KINDS, MAIN_CONTENT_TYPES as _MAIN_CONTENT_TYPES, REL_VBA_PROJECT
from ooxml_common.kinds import kind_for as _kind_for

from ..oxml.xml import make, qn, remove
from . import blank
from .properties import APP_NS, core_part, now, write_core

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
REL_ATTACHED_TEMPLATE = _REL + "attachedTemplate"
REL_VBA = REL_VBA_PROJECT
REL_KEYMAP = "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations"
REL_VBA_DATA = "http://schemas.microsoft.com/office/2006/relationships/wordVbaData"
_MACRO_RELATIONSHIPS = (REL_VBA, REL_KEYMAP, REL_VBA_DATA)
CT_MACRO_DOCUMENT = _MAIN_CONTENT_TYPES["docm"]
CT_MACRO_TEMPLATE = _MAIN_CONTENT_TYPES["dotm"]
_COMPAT_URI = "http://schemas.microsoft.com/office/word"


class AuthoringWarning(UserWarning):
    """Something about a new document the caller should know."""


class MacrosDropped(AuthoringWarning):
    """A macro-enabled template's VBA project was not carried into the ``.docx``."""


class ModeUpgraded(AuthoringWarning):
    """A template below compatibility mode 15 gave a document converted to mode 15."""


class TemplateOpened(AuthoringWarning):
    """:meth:`Document.open` opened a template (``.dotx``, ``.dotm``): edits change the
    template itself.  ``Document.new(template=...)`` makes a document from it."""


#: The kind of Word package each file extension holds: Word refuses a file whose main
#: part's content type is another kind's (measured: ``tests/test_oracle_trial.py``).  Word's
#: share of :data:`ooxml_common.kinds.EXTENSION_KINDS`.
EXTENSION_KINDS = {kind.extension: name for name, kind in KINDS.items() if kind.application == "word"}


def kind_for(target) -> str | None:
    """The package kind a path's extension names (``docx``, ``docm``, ``dotx``, ``dotm``),
    or ``None`` for another extension, a file object or bytes
    (:func:`ooxml_common.kinds.kind_for`, Word's kinds only)."""
    return _kind_for(target, "word")


def has_macros(package) -> bool:
    """Whether the main part has a VBA project (or what belongs to one)."""
    main = package.document_part()
    return any(rel.type in _MACRO_RELATIONSHIPS for rel in package.relationships(main).values())


def content_types_as(package, content_type: str) -> bytes:
    """``[Content_Types].xml`` with the main part's ``Override`` naming ``content_type``,
    serialised; the package itself is not changed
    (:meth:`ooxml_edit.opc.OpcPackage.content_types_with`)."""
    return package.content_types_with(package.document_part(), content_type)


def write_as(document: "Document", target, kind: str, app: bytes | None) -> None:
    """Write ``document`` to ``target`` as a package of ``kind``: the main part's content
    type set to that kind's, and nothing else changed -- but for a VBA project, which a
    ``.docx`` or ``.dotx`` cannot carry (Word does not open one that does): dropped from
    the written copy, with a :class:`MacrosDropped` warning, as Word's Save As drops it."""
    from ..oxml.package import MAIN_CONTENT_TYPES, WordPackage
    from .properties import REL_APP, _part

    package = document.package
    wanted = MAIN_CONTENT_TYPES[kind]
    if not KINDS[kind].macro_enabled and has_macros(package):
        copy_ = type(document)(WordPackage.open(package.to_bytes()), stamping=document.stamping)
        dropped = _drop_macros(copy_)
        warnings.warn(MacrosDropped(f"a .{kind} cannot carry macros: dropped "
                                    f"{', '.join(dropped) or 'the VBA project'} from the file written"),
                      stacklevel=3)
        set_main_content_type(copy_.package, wanted)
        copy_.package.save(target, {_part(package, REL_APP): app} if app is not None else None)
        return
    replacements = {"[Content_Types].xml": content_types_as(package, wanted)}
    if app is not None:
        replacements[_part(package, REL_APP)] = app
    package.save(target, replacements)


# -- Convert ------------------------------------------------------------------------------------

#: Every child of ``w:compat`` but ``w:compatSetting``, in schema order (ECMA-376
#: ``CT_Compat``).
COMPAT_OPTIONS = (
    "useSingleBorderforContiguousCells", "wpJustification", "noTabHangInd", "noLeading", "spaceForUL",
    "noColumnBalance", "balanceSingleByteDoubleByteWidth", "noExtraLineSpacing", "doNotLeaveBackslashAlone",
    "ulTrailSpace", "doNotExpandShiftReturn", "spacingInWholePoints", "lineWrapLikeWord6",
    "printBodyTextBeforeHeader", "printColBlack", "wpSpaceWidth", "showBreaksInFrames", "subFontBySize",
    "suppressBottomSpacing", "suppressTopSpacing", "suppressSpacingAtTopOfPage", "suppressTopSpacingWP",
    "suppressSpBfAfterPgBrk", "swapBordersFacingPages", "convMailMergeEsc", "truncateFontHeightsLikeWP6",
    "mwSmallCaps", "usePrinterMetrics", "doNotSuppressParagraphBorders", "wrapTrailSpaces",
    "footnoteLayoutLikeWW8", "shapeLayoutLikeWW8", "alignTablesRowByRow", "forgetLastTabAlignment",
    "adjustLineHeightInTable", "autoSpaceLikeWord95", "noSpaceRaiseLower", "doNotUseHTMLParagraphAutoSpacing",
    "layoutRawTableWidth", "layoutTableRowsApart", "useWord97LineBreakRules", "doNotBreakWrappedTables",
    "doNotSnapToGridInCell", "selectFldWithFirstOrLastChar", "applyBreakingRules", "doNotWrapTextWithPunct",
    "doNotUseEastAsianBreakRules", "useWord2002TableStyleRules", "growAutofit", "useFELayout",
    "useNormalStyleForList", "doNotUseIndentAsNumberingTabStop", "useAltKinsokuLineBreakRules",
    "allowSpaceOfSameStyleInTable", "doNotSuppressIndentation", "doNotAutofitConstrainedTables",
    "autofitToFirstFixedWidthCell", "underlineTabInNumList", "displayHangulFixedWidth", "splitPgBreakAndParaMark",
    "doNotVertAlignCellWithSp", "doNotBreakConstrainedForcedTable", "doNotVertAlignInTxbx", "useAnsiKerningPairs",
    "cachedColBalance",
)
#: The legacy options Convert keeps (measured: these ten of the 65, in every mode).  And
#: ``useFELayout``, which a save drops, not Convert (module docstring).
KEPT = frozenset({"spaceForUL", "balanceSingleByteDoubleByteWidth", "noExtraLineSpacing",
                  "doNotLeaveBackslashAlone", "ulTrailSpace", "doNotExpandShiftReturn", "suppressBottomSpacing",
                  "adjustLineHeightInTable", "doNotUseHTMLParagraphAutoSpacing", "applyBreakingRules",
                  "useFELayout"})
#: The ``w:compatSetting``\\ s Convert writes, in Word's order, and their values.
MODERN_SETTINGS = (("compatibilityMode", "15"), ("overrideTableStyleFontSizeAndJustification", "1"),
                   ("enableOpenTypeFeatures", "1"), ("doNotFlipMirrorIndents", "1"),
                   ("differentiateMultirowTableHeaders", "1"), ("useWord2013TrackBottomHyphenation", "0"))


def upgrade(document: "Document") -> list[str]:
    """Convert's change to the settings, in place (inside an edit): the legacy options it
    removes, by name."""
    from ..oxml.xml import insert_in_order

    part = document.package.settings_part() or document._create_settings_part()
    root = document.package.tree(part)
    compat = root.find(_W + "compat")
    if compat is None:
        compat = make("w:compat")
        insert_in_order(root, compat)
    removed: list[str] = []
    options: dict[str, etree._Element] = {}
    settings: list[etree._Element] = []
    others: list[etree._Element] = []
    for child in list(compat):
        if not isinstance(child.tag, str):
            others.append(child)
            continue
        name = etree.QName(child).localname
        if child.tag == _W + "compatSetting":
            settings.append(child)
        elif child.tag.startswith(_W) and name in COMPAT_OPTIONS:
            if name in KEPT:
                options[name] = child
            else:
                removed.append(name)
        else:
            others.append(child)
        compat.remove(child)
    by_name = {s.get(_W + "name"): s for s in settings if s.get(_W + "uri") in (None, _COMPAT_URI)}
    for name in COMPAT_OPTIONS:
        if name in options:
            compat.append(options[name])
    placed: list[etree._Element] = []
    for name, value in MODERN_SETTINGS:
        node = by_name.get(name)
        if node is None:
            node = make("w:compatSetting")
            node.set(_W + "name", name)
            node.set(_W + "uri", _COMPAT_URI)
        node.set(_W + "val", value)
        compat.append(node)
        placed.append(node)
    for node in settings:
        if all(node is not other for other in placed):
            compat.append(node)
    for node in others:
        compat.append(node)
    document.package.mark_dirty(part)
    return removed


# -- templates ----------------------------------------------------------------------------------


def _read(source) -> tuple[bytes, str | None]:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source), None
    if isinstance(source, (str, os.PathLike)):
        path = os.path.abspath(os.fspath(source))
        with open(path, "rb") as handle:
            return handle.read(), path
    if hasattr(source, "read"):
        return source.read(), None
    raise TypeError(f"a template is a path or bytes, not {type(source).__name__}")


def set_main_content_type(package, content_type: str) -> None:
    """Point the main part's ``Override`` at ``content_type`` (document <-> template)."""
    from ..oxml.package import CONTENT_TYPES_PART

    main = package.document_part()
    root = package.tree(CONTENT_TYPES_PART)
    found = False
    for node in root:
        if (node.get("PartName") or "").lstrip("/") == main:
            node.set("ContentType", content_type)
            found = True
    if found:
        package.mark_dirty(CONTENT_TYPES_PART)
    else:
        # ooxml-edit's declare_content_type appends an Override beside an existing one.
        package.declare_content_type(main, content_type, override=True)


def _drop_macros(document: "Document") -> list[str]:
    """Remove the VBA project and what belongs to it; the parts removed."""
    package = document.package
    main = package.document_part()
    dropped: list[str] = []
    for rel in list(package.relationships(main).values()):
        if rel.type in _MACRO_RELATIONSHIPS:
            dropped += package.release(main, [rel.id])
    # vbaData.xml hangs off the VBA project, which release reaped with it.  The VBA
    # project's Default content type goes when nothing else is a .bin of that type.
    from ..oxml.package import CONTENT_TYPES_PART

    types = package.tree(CONTENT_TYPES_PART)
    for node in list(types):
        if (node.get("ContentType") or "") == "application/vnd.ms-office.vbaProject" and node.get("Extension"):
            extension = "." + node.get("Extension").lower()
            if not any(name.lower().endswith(extension) and package.content_type(name) == node.get("ContentType")
                       for name in package.part_names if "Override" not in name):
                types.remove(node)
                package.mark_dirty(CONTENT_TYPES_PART)
    return dropped


def _drop_body(document: "Document") -> None:
    """The template's body goes, but for one empty paragraph in the default style and the
    last section; notes and comments it referenced go with it."""
    part = document.package.document_part()
    root = document.package.tree(part)
    body = root.find(_W + "body")
    section = body.find(_W + "sectPr")
    for block in [child for child in body if child is not section]:
        body.remove(block)
    paragraph = make("w:p")
    from .ids import ensure_w14, generate, stamp

    stamp(paragraph, *generate(part, document._used(), 2))
    if section is not None:
        section.addprevious(paragraph)
    else:
        body.append(paragraph)
    ensure_w14(root)
    document.package.mark_dirty(part)
    # Notes and comments nothing references any more.
    for story_part in document.package.story_parts():
        tree = document.package.tree(story_part)
        if tree is None:
            continue
        changed = False
        for note in list(tree):
            if note.tag in (_W + "footnote", _W + "endnote") and note.get(_W + "type") is None:
                tree.remove(note)
                changed = True
            elif note.tag == _W + "comment":
                tree.remove(note)
                changed = True
        if changed:
            document.package.mark_dirty(story_part)
    if document._comment_part("comments") is not None:
        for key in ("extended", "ids", "extensible", "comments"):
            document._drop_comment_part(key)
        document._tidy_people(document._authors_in_use() | {
            person.get("{http://schemas.microsoft.com/office/word/2012/wordml}author")
            for person in (document._comment_tree("people") if document._comment_part("people") else [])})
    # Headers and footers only the dropped sections used.
    main = document.package.document_part()
    used = document.package.referenced_values(main)
    stale = [rel.id for rel in document.package.relationships(main).values()
             if rel.type in (_REL + "header", _REL + "footer") and rel.id not in used]
    if stale:
        document.package.release(main, stale)


def from_template(cls, source, *, keep_content: bool = True, title: str | None = None,
                  author: str | None = None, created: datetime | None = None, attach: bool | None = None,
                  stamping: str = "document") -> "Document":
    data, path = _read(source)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", TemplateOpened)
        document = cls.open(data, stamping=stamping)
    package = document.package
    kind = package.kind
    if kind is None:
        raise ValueError("not a Word document or template")
    created = (created or now()).replace(microsecond=0)
    if kind in ("docm", "dotm"):
        dropped = _drop_macros(document)
        warnings.warn(MacrosDropped(f"the template's macros are not carried into a .docx: dropped "
                                    f"{', '.join(dropped) or 'its VBA project'}"), stacklevel=3)
    set_main_content_type(package, blank.CT_DOCUMENT)
    if not keep_content:
        _drop_body(document)
    mode = package.compatibility_mode()
    if mode < 15:
        upgrade(document)
        warnings.warn(ModeUpgraded(f"the template is in compatibility mode {mode}: the new document is "
                                   "converted to mode 15"), stacklevel=3)
    template_name = None
    if attach is None:
        # Word attaches a .dotx; docx-agent does not unless asked (module docstring): an
        # attached template outside Word's sandbox makes Word ask for access to it on every
        # opening, and a macro-enabled one would bring its macros back.
        attach = False
    if attach:
        if path is None:
            raise ValueError("a template given as bytes has no path to attach")
        _attach(document, path)
        template_name = os.path.basename(path)
    # The core properties start again (Word's): title, subject, keywords, description kept.
    write_core(package, {"author": author, "last_modified_by": author, "revision": 1, "created": created,
                         "modified": created, **({"title": title} if title is not None else {})})
    _app_template(document, template_name or "Normal.dotm")
    return cls.open(document.to_bytes(), stamping=stamping)


def _attach(document: "Document", path: str) -> None:
    from ..oxml.xml import insert_in_order

    package = document.package
    part = package.settings_part() or document._create_settings_part()
    root = package.tree(part)
    for node in root.findall(_W + "attachedTemplate"):
        rid = node.get(qn("r:id"))
        if rid:
            package.remove_relationship(part, rid)
        remove(node)
    target = "file:///" + urllib.parse.quote(path)
    rid = package.add_external_relationship(part, REL_ATTACHED_TEMPLATE, target)
    node = make("w:attachedTemplate")
    node.set(qn("r:id"), rid)
    insert_in_order(root, node)
    package.mark_dirty(part)


def _app_template(document: "Document", name: str) -> None:
    from .properties import REL_APP, _part, app_xml

    package = document.package
    part = _part(package, REL_APP)
    if part is None:
        package.add_part("docProps/app.xml", app_xml(), blank.CT_APP, override=True)
        package.add_relationship("", REL_APP, "docProps/app.xml")
        part = "docProps/app.xml"
    root = etree.fromstring(package.read(part))
    node = root.find("{%s}Template" % APP_NS)
    if node is None:
        node = etree.Element("{%s}Template" % APP_NS)
        root.insert(0, node)
    if node.text != name:
        node.text = name
        package.replace_part(part, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                             + etree.tostring(root, encoding="UTF-8"))


def new(cls, *, page=None, orientation: str | None = None, language: str | None = None,
        locale: blank.Locale | None = None, title: str | None = None, author: str | None = None,
        created: datetime | None = None, template=None, keep_content: bool = True, attach: bool | None = None,
        stamping: str = "document") -> "Document":
    from .properties import as_datetime

    created = as_datetime(created)
    if template is None:
        data = blank.build(page="A4" if page is None else page, orientation=orientation or "portrait",
                           language=language or "en-US", locale=locale, title=title, author=author,
                           created=created)
        return cls.open(data, stamping=stamping)
    document = from_template(cls, template, keep_content=keep_content, title=title, author=author,
                             created=created, attach=attach, stamping=stamping)
    if page is not None or orientation is not None or language is not None:
        _restyle(document, page, orientation, language)
        document = cls.open(document.to_bytes(), stamping=stamping)
    return document


def _restyle(document: "Document", page, orientation: str | None, language: str | None) -> None:
    """A template's page size, orientation and language replaced where they are stated."""
    package = document.package
    if page is not None or orientation is not None:
        part = package.document_part()
        for section in package.tree(part).iter(_W + "sectPr"):
            size = section.find(_W + "pgSz")
            if size is None:
                continue
            current = "landscape" if size.get(_W + "orient") == "landscape" else "portrait"
            if page is not None:
                width, height = blank.page_size(page, orientation or current)
            else:
                width, height = (int(size.get(_W + "w", "11906")), int(size.get(_W + "h", "16838")))
                if (orientation == "landscape") != (width > height):
                    width, height = height, width
            size.set(_W + "w", str(width))
            size.set(_W + "h", str(height))
            if (orientation or current) == "landscape":
                size.set(_W + "orient", "landscape")
            elif size.get(_W + "orient") is not None:
                del size.attrib[_W + "orient"]
        package.mark_dirty(part)
    if language is not None:
        styles = package.styles_part()
        root = package.tree(styles) if styles is not None else None
        defaults = None if root is None else root.find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr")
        if defaults is not None:
            lang = defaults.find(_W + "lang")
            if lang is None:
                from ..oxml.xml import insert_in_order

                lang = make("w:lang")
                insert_in_order(defaults, lang)
            lang.set(_W + "val", language)
            package.mark_dirty(styles)
        settings = package.settings_part()
        node = package.tree(settings).find(_W + "themeFontLang") if settings else None
        if node is not None:
            node.set(_W + "val", language)
            package.mark_dirty(settings)


# -- the operations on a document -----------------------------------------------------------------


class AuthoringOps:
    """:meth:`Document.save_as_template` and :meth:`Document.upgrade_to_modern`."""

    def save_as_template(self: "Document", target) -> None:
        """Write the document as a template (``.dotx``; a macro-enabled document as
        ``.dotm``) to ``target``, without changing the document: Word's Save As Template
        changes the main part's content type and nothing else (measured).  A ``.dotx`` or
        ``.dotm`` target decides which (a ``.dotx`` drops a VBA project, with
        :class:`MacrosDropped`); a ``.docx`` or ``.docm`` one is refused -- that is
        :meth:`save`."""
        from .errors import EditError
        from .properties import refreshed_app

        kind = kind_for(target)
        if kind in ("docx", "docm"):
            raise EditError(f"save_as_template writes a .dotx or .dotm, not a .{kind}: save() writes a document")
        if kind is None:
            kind = "dotx" if self.package.kind in ("docx", "dotx") else "dotm"
        write_as(self, target, kind, refreshed_app(self))

    def upgrade_to_modern(self: "Document") -> "EditResult":
        """Bring the document into compatibility mode 15 as Word's Convert does (module
        docstring): one undo step, opt-in only -- no other edit changes the mode.  The
        result's ``reflow`` is docx2svg's comparison of the pages before and after
        (:meth:`DocumentLayout.compare`); ``warnings`` name the legacy options removed.
        A document already in mode 15 is left as it is."""
        from .document import EditResult

        if self.package.compatibility_mode() >= 15:
            return EditResult(None, changed=False)
        before = self.layout()
        with self._edit():
            removed = upgrade(self)
        after = self.layout()
        return EditResult(None, warnings=[f"removed compatibility option {name}" for name in removed],
                          reflow=after.compare(before))


__all__ = ["AuthoringOps", "AuthoringWarning", "COMPAT_OPTIONS", "EXTENSION_KINDS", "KEPT", "MacrosDropped",
           "ModeUpgraded", "TemplateOpened", "content_types_as", "from_template", "has_macros", "kind_for", "new",
           "upgrade", "write_as"]
