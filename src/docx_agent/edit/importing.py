"""Copying blocks from another document: :meth:`Document.copy_blocks`.

The blocks -- paragraphs, tables, block content controls, or a range of them -- are copied
with everything they depend on, as Word copies content between documents.  Measured
(``tools/e6_probe.py``: a source and a destination whose styles, list, footnote, comment,
bookmark, picture, content control and paraIds all collide; Word's ``formatted text``
copy, ``insert file``, and the clipboard's paste with each paste option):

**Styles, matched by name** (``styles=``):

* ``"use_destination"`` (the default; what Word's default paste, Use Destination Styles,
  ``insert file`` and ``formatted text`` all did): a style the destination has by the same
  name is the destination's -- the copy takes its look; a style it lacks is imported,
  with what it is based on, its next and linked styles, and, for a paragraph style, the
  source's document defaults where they differ from the destination's (Word gave the
  imported paragraph style the source's default face and size; not a character style);
* ``"keep_source"`` (Keep Source Formatting): no style is imported -- every copied
  paragraph is the destination's default paragraph style, and the source's effective
  paragraph and run formatting, where it differs from the destination's default, is
  written as direct formatting on the paragraph, its mark and its runs (character styles
  too: a footnote reference became superscript);
* ``"merge"`` (Merge Formatting): the default paragraph style, and of the formatting only
  bold, italic and underline, direct -- but a heading keeps its level and a list item its
  list: a paragraph whose style is a heading (by name or outline level) or a list style
  takes the destination's style of that name (a built-in one it lacks added as Word
  writes it), its formatting left to that style.  (Word's own Merge Formatting makes them
  Normal paragraphs; the end-to-end trial needed them kept, so this departs from Word.)

**Mapping styles** (``style_map=``, ``unmapped=``): ``style_map`` names, by source style name,
the destination style a source style becomes (``{"Bulletin Body": "Handbook Body"}``; a
built-in name the destination lacks is added as Word writes it); it applies before the policy
looks a style up by name, in every policy.  ``unmapped="body"`` makes a source style the map
does not name and the destination lacks *by name* -- Word's built-in ones (List Number,
annotation text...) added as Word writes them -- the destination's body style -- the
paragraph style most of its body's plain paragraphs (not headings, not list items) use, else
its default -- instead of importing it; a character style becomes the paragraph's own font
(no ``w:rStyle``), a table style the default table style.  ``unmapped="import"`` (the
default) imports it, as Word does.

**Lists** (``lists=``): each source list becomes a list of its own in the destination, its
abstract definition copied (its ``w:nsid`` kept unless the destination uses it) and a new
instance over it, so it does not continue the destination's lists (Word: a new ``numId``
2, abstract 0 copied) -- or, with ``"continue"``, joins the list of the destination
paragraph before the insertion when its first level has the same number format.

**Pictures and other parts**: every relationship the copies use is made again in the
destination -- media de-duplicated by content (Word gave both pictures one
``image1.png``), other parts (charts and what they embed) copied with their own
relationships, hyperlinks as external relationships.

**Notes and comments** referenced by the copies are copied with them and renumbered (Word:
footnote 2, comment 2), their content imported as the blocks' is; a comment's resolved
state and reply link come along, with a new durable id.

**Ids**: a bookmark whose name the destination has is **renamed** (``Name_1``...), and the
copies' links and fields that point at it follow -- Word *drops* such a bookmark from the
copy (``bookmarks="drop"`` does as Word); bookmark ids are renumbered.  A paraId, a
drawing's ``docPr`` id, ``wp14`` ids and a content control's id the destination (or an
earlier copy) already uses are re-issued (Word re-issued the ``docPr`` id and the control's
id; on paraIds it rewrote *every* paraId in the document, which docx-agent does not: only
the colliding copies get new ones, so every other id lasts).  The result's ``copied`` maps
each source id to its copy's (``p:``, ``t:``, ``fn:``/``en:``, ``comment#``, ``sdt#``).

**Tracked** (``track=`` or ``doc.tracking(...)``), the copy is an insertion, as Word's
tracked paste is: every paragraph's mark and runs, every row, in one revision group.

**Approximations**: the copies take the source's current view (insertions kept, deletions
dropped, formatting changes' records dropped); a paragraph's section break is not copied
(the copy takes the destination's sections); a data-bound control's binding is dropped
(its text stays); in ``keep_source`` the table style's conditional formatting is not
flattened into its cells' paragraphs, and theme fonts and colours are written as the
source's faces and values only where the two themes differ.
"""

from __future__ import annotations

import copy
import posixpath
import re
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, insert_in_order, make, remove
from . import ids as _ids
from . import numbering as _numbering
from .errors import EditError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_O = "{urn:schemas-microsoft-com:office:office}"
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
W_P, W_TBL, W_SDT = _W + "p", _W + "tbl", _W + "sdt"
_BLOCKS = (W_P, W_TBL, W_SDT, _W + "customXml")

STYLE_POLICIES = ("use_destination", "keep_source", "merge")
UNMAPPED_POLICIES = ("import", "body")
#: Source style names a heading's or list's paragraph keeps under ``merge``.
_LIST_STYLE = re.compile(r"(list (bullet|number|continue)( [2-5])?|list paragraph)", re.IGNORECASE)
LIST_POLICIES = ("separate", "continue")
BOOKMARK_POLICIES = ("rename", "drop")

#: Run properties that toggle along a style hierarchy (ECMA-376 17.7.3).
TOGGLES = frozenset(_W + name for name in ("b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
                                           "outline", "shadow", "emboss", "imprint", "vanish"))
#: Properties whose attributes merge level by level, rather than the element replacing.
_MERGED = frozenset(_W + name for name in ("rFonts", "lang", "spacing", "ind"))
#: Paragraph properties that are not formatting.
_NOT_FORMATTING = frozenset(_W + name for name in ("pStyle", "rPr", "sectPr", "pPrChange", "numPr", "rStyle"))
_THEME_FONT_ATTRIBUTES = {"asciiTheme": "ascii", "hAnsiTheme": "hAnsi", "eastAsiaTheme": "eastAsia",
                          "cstheme": "cs"}
#: What Merge Formatting keeps.
_MERGE_KEEPS = (_W + "b", _W + "i", _W + "u")


def _on(node: Element | None) -> bool | None:
    if node is None:
        return None
    return (node.get(_W + "val") or "true").lower() not in ("0", "false", "off", "none")


def _canonical(node: Element) -> bytes:
    return etree.tostring(node, method="c14n")


class _Styles:
    """One document's styles, by id and by name."""

    def __init__(self, document: "Document") -> None:
        part = document.package.styles_part()
        self.root = document.package.tree(part) if part is not None else None
        self.by_id: dict[str, Element] = {}
        self.by_name: dict[tuple[str, str], Element] = {}
        if self.root is not None:
            for node in self.root.findall(_W + "style"):
                self.add(node)

    def add(self, node: Element) -> None:
        self.by_id[node.get(_W + "styleId")] = node
        name = node.find(_W + "name")
        if name is not None:
            self.by_name.setdefault((node.get(_W + "type"), (name.get(_W + "val") or "").casefold()), node)

    def name(self, style_id: str) -> str | None:
        node = self.by_id.get(style_id)
        found = None if node is None else node.find(_W + "name")
        return None if found is None else found.get(_W + "val")

    def default(self, kind: str) -> Element | None:
        return next((n for n in self.by_id.values() if n.get(_W + "type") == kind
                     and n.get(_W + "default") in ("1", "true", "on")), None)

    def chain(self, style_id: str | None, kind: str) -> list[Element]:
        """The style and what it is based on, root first (the default style when none)."""
        node = self.by_id.get(style_id) if style_id else self.default(kind)
        out: list[Element] = []
        seen: set[int] = set()
        while node is not None and id(node) not in seen:
            seen.add(id(node))
            out.append(node)
            based = node.find(_W + "basedOn")
            node = self.by_id.get(based.get(_W + "val")) if based is not None else None
        return out[::-1]

    def defaults(self, tag: str) -> Element | None:
        if self.root is None:
            return None
        return self.root.find(f"{_W}docDefaults/{_W}{tag}Default/{_W}{tag}")


def _merge_into(accumulated: dict[str, Element], node: Element | None, *, toggles: dict | None = None) -> None:
    """Fold one level's properties into ``accumulated`` (tag -> element): later wins, the
    attributes of fonts, languages, spacing and indents merge one by one, and toggles
    (``toggles`` given) flip."""
    if node is None:
        return
    for child in node:
        if not isinstance(child.tag, str) or child.tag in _NOT_FORMATTING or child.tag.endswith("Change"):
            continue
        if toggles is not None and child.tag in TOGGLES:
            toggles[child.tag] = toggles.get(child.tag, False) ^ bool(_on(child))
            continue
        if child.tag in _MERGED and child.tag in accumulated:
            merged = copy.deepcopy(accumulated[child.tag])
            for key, value in child.attrib.items():
                local = etree.QName(key).localname
                if child.tag == _W + "rFonts":
                    # A face named outright replaces the theme's, and the other way round.
                    for theme, plain in _THEME_FONT_ATTRIBUTES.items():
                        if local == plain:
                            merged.attrib.pop(_W + theme, None)
                        elif local == theme:
                            merged.attrib.pop(_W + plain, None)
                if child.tag == _W + "ind" and local in ("hanging", "firstLine"):
                    merged.attrib.pop(_W + "hanging", None)
                    merged.attrib.pop(_W + "firstLine", None)
                merged.set(key, value)
            accumulated[child.tag] = merged
        else:
            accumulated[child.tag] = copy.deepcopy(child)


class _Effective:
    """Effective paragraph and run formatting in one document, at the XML level: document
    defaults, the paragraph style's chain, the character style's chain, direct formatting."""

    def __init__(self, styles: _Styles) -> None:
        self.styles = styles

    def paragraph(self, style_id: str | None, direct: Element | None) -> dict[str, Element]:
        out: dict[str, Element] = {}
        _merge_into(out, self.styles.defaults("pPr"))
        for node in self.styles.chain(style_id, "paragraph"):
            _merge_into(out, node.find(_W + "pPr"))
        _merge_into(out, direct)
        return out

    def run(self, paragraph_style: str | None, run_style: str | None, direct: Element | None) -> dict[str, Element]:
        out: dict[str, Element] = {}
        _merge_into(out, self.styles.defaults("rPr"))
        base_toggles = {tag: bool(_on(out.pop(tag))) for tag in list(out) if tag in TOGGLES}
        levels: list[dict[str, bool]] = []
        for chain in (self.styles.chain(paragraph_style, "paragraph"),
                      self.styles.chain(run_style, "character") if run_style else []):
            flips: dict[str, bool] = {}
            for node in chain:
                level: dict[str, bool] = {}
                _merge_into(out, node.find(_W + "rPr"), toggles=level)
                flips.update(level)
            levels.append(flips)
        for tag in TOGGLES:
            if any(tag in level for level in levels):
                value = False
                for level in levels:
                    value ^= level.get(tag, False)
            elif tag in base_toggles:
                value = base_toggles[tag]
            else:
                continue
            node = make("w:" + etree.QName(tag).localname)
            if not value:
                node.set(_W + "val", "0")
            out[tag] = node
        if direct is not None:
            for child in direct:
                if isinstance(child.tag, str) and child.tag in TOGGLES:
                    out[child.tag] = copy.deepcopy(child)
            _merge_into(out, direct)
        return out


def _theme_root(document: "Document") -> Element | None:
    main = document.package.document_part()
    parts = document.package.related_parts_of_type(
        main, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
    return document.package.tree(parts[0]) if parts and document.package.has_part(parts[0]) else None


def _theme_faces(document: "Document") -> dict[str, str]:
    """``minorHAnsi`` -> the face the theme names, and so on."""
    root = _theme_root(document)
    out: dict[str, str] = {}
    if root is None:
        return out
    for which in ("major", "minor"):
        font = root.find(f"{_A}themeElements/{_A}fontScheme/{_A}{which}Font")
        if font is None:
            continue
        for script, key in (("latin", "HAnsi"), ("ea", "EastAsia"), ("cs", "Bidi")):
            face = font.find(_A + script)
            if face is not None and face.get("typeface"):
                out[f"{which}{key}"] = face.get("typeface")
    return out


def _theme_colors(document: "Document") -> dict[str, str]:
    root = _theme_root(document)
    scheme = None if root is None else root.find(f"{_A}themeElements/{_A}clrScheme")
    out: dict[str, str] = {}
    for slot in (scheme if scheme is not None else []):
        value, system = slot.find(_A + "srgbClr"), slot.find(_A + "sysClr")
        out[etree.QName(slot).localname] = (value.get("val") if value is not None else
                                           system.get("lastClr") if system is not None else "")
    return out


# -- the importer -------------------------------------------------------------------------------


class Importer:
    """What one :meth:`Document.copy_blocks` brings over, and the maps it keeps."""

    def __init__(self, destination: "Document", source: "Document", *, styles: str, lists: str,
                 bookmarks: str, at_part: str, anchor: Element | None, style_map: dict | None = None,
                 unmapped: str = "import") -> None:
        self.dest = destination
        #: Source style name (casefolded) -> destination style name.
        self.name_map = {str(k).casefold(): v for k, v in (style_map or {}).items()}
        self.unmapped = unmapped
        self._body_style: str | None | bool = False
        self.source = source
        self.style_policy = styles
        self.list_policy = lists
        self.bookmark_policy = bookmarks
        self.part = at_part
        self.anchor = anchor
        self.src_styles = _Styles(source)
        self.dest_styles = _Styles(destination)
        self.style_map: dict[str, str] = {}
        self.num_map: dict[int, int] = {}
        self.abstract_map: dict[int | None, int] = {}
        self.part_map: dict[str, str] = {}
        self.copied: dict[str, str] = {}
        self.warnings: list[str] = []
        #: (kind, number, notes part, note element) for every note copied.
        self.notes_created: list[tuple[str, int, str, Element]] = []
        self._doc_prs: set[int] | None = None
        self._controls: set[int] | None = None
        self._annotation = destination._next_annotation_id()
        self._bookmarks = set(self._dest_bookmark_names())
        self._renamed_bookmarks: dict[str, str] = {}
        self._source_faces = _theme_faces(source)
        self._themes_differ = self._source_faces != _theme_faces(destination)
        self._colors_differ = _theme_colors(source) != _theme_colors(destination)

    def bring(self, elements: list[Element], source_part: str, dest_part: str) -> None:
        """Make ``elements`` (copies, not yet placed) fit the destination: current view,
        styles, lists, relationships, notes, comments, bookmarks, ids."""
        for element in elements:
            _current_view(element)
            for section in list(element.iter(_W + "sectPr")):
                if section.getparent() is not None and section.getparent().tag == _W + "pPr":
                    remove(section)
            for binding in list(element.iter(_W + "dataBinding")):
                remove(binding)
        self._formatting(elements)
        self._lists(elements)
        self._relationships(elements, source_part, dest_part)
        self._notes(elements)
        self._comments(elements)
        self._bookmarks_in(elements)
        self._drawing_ids(elements)
        self._control_ids(elements)

    # -- styles -----------------------------------------------------------------------------

    def _formatting(self, elements: list[Element]) -> None:
        kinds = {_W + "pStyle": "paragraph", _W + "rStyle": "character", _W + "tblStyle": "table"}
        if self.style_policy == "use_destination":
            for element in elements:
                for node in list(element.iter(*kinds)):
                    _set_or_drop(node, self.style(node.get(_W + "val"), kinds[node.tag]))
            return
        effective = _Effective(self.src_styles)
        destination = _Effective(self.dest_styles)
        base_paragraph = destination.paragraph(None, None)
        base_run = destination.run(None, None, None)
        for element in elements:
            for node in list(element.iter(_W + "tblStyle")):
                _set_or_drop(node, self.style(node.get(_W + "val"), "table"))
            for paragraph in list(element.iter(W_P)):
                if self.style_policy == "merge" and self._keeps_structure(paragraph):
                    continue
                self._flatten(paragraph, effective, base_paragraph, base_run)

    def _keeps_structure(self, paragraph: Element) -> bool:
        """Merge Formatting here: a heading or a list item keeps a style of its kind -- the
        destination's of the same name (or the one ``style_map`` names) -- and its list."""
        properties = paragraph.find(_W + "pPr")
        node = None if properties is None else properties.find(_W + "pStyle")
        style_id = None if node is None else node.get(_W + "val")
        if not style_id:
            return False
        name = self.src_styles.name(style_id) or style_id
        outline = None
        for level in self.src_styles.chain(style_id, "paragraph"):
            found = level.find(f"{_W}pPr/{_W}outlineLvl")
            if found is not None:
                outline = found.get(_W + "val")
                break
        heading = re.fullmatch(r"heading [1-9]", name, re.IGNORECASE) or (outline is not None and outline != "9")
        listed = _LIST_STYLE.fullmatch(name) or any(
            level.find(f"{_W}pPr/{_W}numPr") is not None for level in self.src_styles.chain(style_id, "paragraph"))
        if not (heading or listed):
            return False
        wanted = self.name_map.get(name.casefold())
        if wanted is None and not re.fullmatch(r"heading [1-9]", name, re.IGNORECASE) and outline is not None \
                and outline.isdigit() and int(outline) <= 8 and heading:
            wanted = f"heading {int(outline) + 1}"
        try:
            target = self._destination_style(wanted or name, "paragraph")
        except EditError:
            return False
        if target is None:
            return False
        numbering = None if properties is None else properties.find(_W + "numPr")
        if numbering is None and listed:
            # The list a style gave: written on the paragraph, so it stays a list whatever
            # the destination style says.
            for level in reversed(self.src_styles.chain(style_id, "paragraph")):
                found = level.find(f"{_W}pPr/{_W}numPr")
                if found is not None:
                    numbering = copy.deepcopy(found)
                    if numbering.find(_W + "ilvl") is None:
                        numbering.insert(0, make("w:ilvl", **{"w:val": "0"}))
                    insert_in_order(properties, numbering)
                    break
        node.set(_W + "val", target)
        for run in paragraph.iter(_W + "rStyle"):
            _set_or_drop(run, self.style(run.get(_W + "val"), "character"))
        return True

    def _destination_style(self, name: str, kind: str) -> str | None:
        """The id of the destination's ``kind`` style ``name`` (a built-in one it lacks added
        as Word writes it), or ``None`` if it has none and Word has no such built-in style."""
        from .styles import StyleError

        styles = self.dest.styles
        found = styles.find(name, kind)
        if found is not None:
            return found.id
        try:
            style_id = styles._ensure(name, kind)
        except StyleError:
            return None
        self.dest_styles = _Styles(self.dest)
        self.warnings.append(f"style {name!r} added as Word writes it")
        return style_id

    def body_style(self) -> str | None:
        """The destination's body style: the paragraph style most of its body's plain
        paragraphs use (no heading, no list), else its default; ``None`` is the default."""
        if self._body_style is not False:
            return self._body_style  # type: ignore[return-value]
        counts: dict[str | None, int] = {}
        root = self.dest.package.tree(self.dest.package.document_part())
        body = root.find(_W + "body")
        for paragraph in (body.findall(W_P) if body is not None else []):
            properties = paragraph.find(_W + "pPr")
            if properties is not None and properties.find(_W + "numPr") is not None:
                continue
            node = None if properties is None else properties.find(_W + "pStyle")
            style_id = None if node is None else node.get(_W + "val")
            if style_id is not None:
                name = self.dest_styles.name(style_id) or ""
                chain = self.dest_styles.chain(style_id, "paragraph")
                if re.fullmatch(r"heading [1-9]|title|subtitle", name, re.IGNORECASE) or _LIST_STYLE.fullmatch(name) \
                        or any(level.find(f"{_W}pPr/{_W}outlineLvl") is not None
                               or level.find(f"{_W}pPr/{_W}numPr") is not None for level in chain):
                    continue
            if not "".join(t.text or "" for t in paragraph.iter(_W + "t")).strip():
                continue
            counts[style_id] = counts.get(style_id, 0) + 1
        best = max(counts.items(), key=lambda kv: kv[1])[0] if counts else None
        default = self.dest_styles.default("paragraph")
        if best is not None and default is not None and best == default.get(_W + "styleId"):
            best = None
        self._body_style = best
        return best

    def _flatten(self, paragraph: Element, effective: _Effective, base_paragraph: dict, base_run: dict) -> None:
        """Keep Source Formatting or Merge Formatting on one paragraph: the default style,
        and the source's look (all of it, or only bold, italic and underline) direct."""
        properties = paragraph.find(_W + "pPr")
        style_node = None if properties is None else properties.find(_W + "pStyle")
        style_id = None if style_node is None else style_node.get(_W + "val")
        merge = self.style_policy == "merge"
        numbering = None if properties is None else properties.find(_W + "numPr")
        if numbering is not None:
            numbering = copy.deepcopy(numbering)
        else:
            for node in reversed(self.src_styles.chain(style_id, "paragraph")):
                found = node.find(f"{_W}pPr/{_W}numPr")
                if found is not None:
                    numbering = copy.deepcopy(found)
                    if numbering.find(_W + "ilvl") is None:
                        numbering.insert(0, make("w:ilvl", **{"w:val": "0"}))
                    break
        new_properties = make("w:pPr")
        if numbering is not None:
            insert_in_order(new_properties, numbering)
        if not merge:
            for tag, node in effective.paragraph(style_id, properties).items():
                if tag in base_paragraph and _canonical(base_paragraph[tag]) == _canonical(node):
                    continue
                insert_in_order(new_properties, node)
        mark = None if properties is None else properties.find(_W + "rPr")
        mark_values = self._run_values(effective, style_id, None, mark, base_run, merge)
        if mark is not None:
            for record in mark:
                if record.tag in (_W + "ins", _W + "del"):
                    insert_in_order(mark_values, copy.deepcopy(record))
        if len(mark_values):
            insert_in_order(new_properties, mark_values)
        if properties is not None:
            remove(properties)
        if len(new_properties):
            paragraph.insert(0, new_properties)
        for run in list(paragraph.iter(_W + "r")):
            direct = run.find(_W + "rPr")
            run_style = None
            if direct is not None and direct.find(_W + "rStyle") is not None:
                run_style = direct.find(_W + "rStyle").get(_W + "val")
            values = self._run_values(effective, style_id, run_style, direct, base_run, merge)
            if direct is not None:
                remove(direct)
            if len(values):
                run.insert(0, values)

    def _run_values(self, effective: _Effective, paragraph_style, run_style, direct, base_run, merge) -> Element:
        out = make("w:rPr")
        for tag, node in effective.run(paragraph_style, run_style, direct).items():
            if merge:
                if tag not in _MERGE_KEEPS:
                    continue
                if tag == _W + "u" and (node.get(_W + "val") or "single") == "none":
                    continue
                if tag != _W + "u" and not _on(node):
                    continue
            if tag in base_run and _canonical(base_run[tag]) == _canonical(node):
                continue
            if tag in TOGGLES and not _on(node) and tag not in base_run:
                continue
            insert_in_order(out, self._concrete(node))
        return out

    def _concrete(self, node: Element) -> Element:
        """Theme fonts and colours as the source's faces and values, where the destination's
        theme names others."""
        if node.tag == _W + "rFonts" and self._themes_differ:
            node = copy.deepcopy(node)
            for theme, plain in _THEME_FONT_ATTRIBUTES.items():
                value = node.attrib.pop(_W + theme, None)
                if value is not None and value in self._source_faces:
                    node.set(_W + plain, self._source_faces[value])
        elif node.tag == _W + "color" and self._colors_differ:
            node = copy.deepcopy(node)
            for key in ("themeColor", "themeShade", "themeTint"):
                node.attrib.pop(_W + key, None)
        return node

    def style(self, style_id: str | None, kind: str) -> str | None:
        """The destination's id for a source style: the destination's own by name, else the
        source's definition imported."""
        if not style_id:
            return style_id
        if style_id in self.style_map:
            return self.style_map[style_id]
        node = self.src_styles.by_id.get(style_id)
        if node is None:
            return style_id
        name = self.src_styles.name(style_id) or style_id
        kind = node.get(_W + "type") or kind
        wanted = self.name_map.get(name.casefold())
        if wanted is not None:
            target = self._destination_style(wanted, kind)
            if target is None:
                raise EditError(f"style_map: the destination has no {kind} style {wanted!r} "
                                "(and Word no built-in one by that name)")
            self.style_map[style_id] = target
            return target
        existing = self.dest_styles.by_name.get((kind, name.casefold()))
        if existing is None and node.get(_W + "default") in ("1", "true", "on"):
            existing = self.dest_styles.default(kind)
        if existing is not None:
            self.style_map[style_id] = existing.get(_W + "styleId")
            return self.style_map[style_id]
        if self.unmapped == "body":
            builtin = self._destination_style(name, kind) if kind in ("paragraph", "character", "table") else None
            if builtin is not None:
                # A built-in style (List Number, annotation text...): Word's own definition,
                # not the source's.
                self.style_map[style_id] = builtin
                return builtin
            target = self.body_style() if kind == "paragraph" else None
            self.style_map[style_id] = target  # type: ignore[assignment]
            self.warnings.append(f"style {name!r} not imported: "
                                 + (f"the destination's body style" if kind == "paragraph" else "dropped"))
            return target
        imported = copy.deepcopy(node)
        new_id = self._free_style_id(style_id, name)
        self.style_map[style_id] = new_id
        imported.set(_W + "styleId", new_id)
        imported.attrib.pop(_W + "default", None)
        for child in list(imported):
            if child.tag == _W + "rsid":
                remove(child)
        for tag in ("basedOn", "next", "link"):
            ref = imported.find(_W + tag)
            if ref is None:
                continue
            target_node = self.src_styles.by_id.get(ref.get(_W + "val"))
            if target_node is None:
                remove(ref)
                continue
            ref.set(_W + "val", self.style(ref.get(_W + "val"), target_node.get(_W + "type")))
        for numbering in imported.iter(_W + "numId"):
            numbering.set(_W + "val", str(self.num(int(numbering.get(_W + "val") or 0))))
        if kind == "paragraph" and self._based_in_destination(imported):
            self._fold_defaults(imported)
        self._append_style(imported)
        self.warnings.append(f"style {name!r} imported from the source")
        return new_id

    def _based_in_destination(self, imported: Element) -> bool:
        based = imported.find(_W + "basedOn")
        return based is None or based.get(_W + "val") in self.dest_styles.by_id

    def _fold_defaults(self, imported: Element) -> None:
        """The source's document defaults where the destination's differ, into a paragraph
        style imported (measured: Word gave it the source's default face and size)."""
        for tag in ("rPr", "pPr"):
            source = self.src_styles.defaults(tag)
            dest = self.dest_styles.defaults(tag)
            if source is None:
                continue
            ours = {child.tag: child for child in (dest if dest is not None else [])}
            container = imported.find(_W + tag)
            for child in source:
                if not isinstance(child.tag, str):
                    continue
                if child.tag in ours and _canonical(ours[child.tag]) == _canonical(child):
                    continue
                if container is not None and container.find(child.tag) is not None:
                    continue
                if container is None:
                    container = make("w:" + tag)
                    insert_in_order(imported, container)
                insert_in_order(container, self._concrete(copy.deepcopy(child)))

    def _free_style_id(self, style_id: str, name: str) -> str:
        taken = {key.casefold() for key in self.dest_styles.by_id}
        candidate, number = style_id, 1
        while candidate.casefold() in taken:
            number += 1
            candidate = f"{re.sub(r'[^0-9A-Za-z]', '', name) or 'Style'}{number}"
        return candidate

    def _append_style(self, node: Element) -> None:
        if self.dest_styles.root is None:
            self.dest._create_styles_part()
            self.dest_styles = _Styles(self.dest)
        self.dest_styles.root.append(node)
        self.dest_styles.add(node)
        self.dest.package.mark_dirty(self.dest.package.styles_part())

    # -- lists ------------------------------------------------------------------------------

    def _lists(self, elements: list[Element]) -> None:
        for element in elements:
            for node in element.iter(_W + "numId"):
                if node.getparent() is not None and node.getparent().tag == _W + "numPr":
                    node.set(_W + "val", str(self.num(int(node.get(_W + "val") or 0))))

    def num(self, num_id: int) -> int:
        """The destination instance a source list's paragraphs go into."""
        if num_id == 0:
            return 0
        if num_id in self.num_map:
            return self.num_map[num_id]
        source = _numbering.Numbering(self.source)
        nums = source.nums()
        if num_id not in nums:
            return 0
        if self.list_policy == "continue":
            joined = self._continued(num_id)
            if joined is not None:
                self.num_map[num_id] = joined
                return joined
        new_abstract = self.abstract(source.abstract_of(num_id))
        new_num = _numbering.add_num(self.dest, new_abstract)
        destination = _numbering.Numbering(self.dest)
        node = destination.nums()[new_num]
        for override in nums[num_id].findall(_W + "lvlOverride"):
            node.append(copy.deepcopy(override))
        self.num_map[num_id] = new_num
        self.dest.package.mark_dirty(destination.part)
        return new_num

    def abstract(self, abstract_id: int | None) -> int:
        if abstract_id in self.abstract_map:
            return self.abstract_map[abstract_id]
        original = _numbering.Numbering(self.source).abstracts().get(abstract_id)
        destination = _numbering.Numbering(self.dest)
        new_id = max(destination.abstracts(), default=-1) + 1
        if original is None:
            node = _numbering.builtin_abstract("number", new_id, "00000001", "00000001")
        else:
            node = copy.deepcopy(original)
        node.set(_W + "abstractNumId", str(new_id))
        nsid = node.find(_W + "nsid")
        used = set()
        if destination.root is not None:
            used = {(n.get(_W + "val") or "").upper() for n in destination.root.iter(_W + "nsid")}
        if nsid is None or (nsid.get(_W + "val") or "").upper() in used:
            if nsid is None:
                nsid = make("w:nsid")
                node.insert(0, nsid)
            nsid.set(_W + "val", _numbering._nsid(self.dest, f"import\0{abstract_id}\0{new_id}"))
        for picture in list(node.iter(_W + "lvlPicBulletId")):
            remove(picture)
        self.abstract_map[abstract_id] = new_id
        for link in list(node.iter(_W + "styleLink", _W + "numStyleLink", _W + "pStyle")):
            kind = "paragraph" if link.tag == _W + "pStyle" else "numbering"
            _set_or_drop(link, self.style(link.get(_W + "val"), kind))
        _numbering.add_abstract(self.dest, node)
        return new_id

    def _continued(self, num_id: int) -> int | None:
        """The destination list before the insertion, when its first level counts as the
        copied one's does."""
        if self.anchor is None:
            return None
        root = self.dest.package.tree(self.part)
        paragraphs = list(root.iter(W_P))
        last = self.anchor if self.anchor.tag == W_P else next(
            (p for p in reversed(list(self.anchor.iter(W_P)))), None)
        if last is None or all(p is not last for p in paragraphs):
            return None
        index = next(k for k, p in enumerate(paragraphs) if p is last)
        previous = None
        for paragraph in reversed(paragraphs[:index + 1]):
            found = paragraph.find(f"{_W}pPr/{_W}numPr/{_W}numId")
            if found is not None and found.get(_W + "val") not in (None, "0"):
                previous = int(found.get(_W + "val"))
                break
        if previous is None:
            return None

        def number_format(level: Element | None) -> str | None:
            fmt = None if level is None else level.find(_W + "numFmt")
            return None if fmt is None else fmt.get(_W + "val")

        source_format = number_format(_numbering.Numbering(self.source).level(num_id, 0))
        dest_format = number_format(_numbering.Numbering(self.dest).level(previous, 0))
        return previous if source_format is not None and source_format == dest_format else None

    # -- relationships and parts ------------------------------------------------------------

    def _relationships(self, elements: list[Element], source_part: str, dest_part: str) -> None:
        cache: dict[str, str] = {}
        for element in elements:
            for node in element.iter():
                if not isinstance(node.tag, str):
                    continue
                for key, value in list(node.attrib.items()):
                    if key.startswith(_R) or key == _O + "relid":
                        if value not in cache:
                            cache[value] = self.relationship(source_part, value, dest_part)
                        node.set(key, cache[value])

    def relationship(self, source_part: str, rid: str, dest_part: str) -> str:
        """A relationship from ``dest_part`` like ``source_part``'s ``rid``."""
        rel = self.source.package.relationships(source_part).get(rid)
        if rel is None:
            return rid
        package = self.dest.package
        if rel.is_external:
            return package.add_external_relationship(dest_part, rel.type, rel.target)
        target = rel.target_part
        if target is None or not self.source.package.has_part(target):
            return rid
        return package.add_relationship(dest_part, rel.type, self.copy_part(target))

    def copy_part(self, source_part: str) -> str:
        """The destination's copy of a source part: media by content, anything else copied
        once, with its own relationships made again."""
        if source_part in self.part_map:
            return self.part_map[source_part]
        package = self.dest.package
        data = self.source.package.read(source_part)
        content_type = self.source.package.content_type(source_part)
        directory = posixpath.dirname(source_part)
        if posixpath.basename(directory) == "media":
            existing = package.find_part_with_bytes(data, directory)
            if existing is not None:
                self.part_map[source_part] = existing
                return existing
        base, extension = posixpath.splitext(posixpath.basename(source_part))
        target = source_part
        if package.has_part(target):
            stem = re.sub(r"\d+$", "", base) or base
            target = package.unused_part_name(posixpath.join(directory, f"{stem}{{n}}{extension}"))
        package.add_part(target, data, content_type)
        self.part_map[source_part] = target
        for rel_id, rel in self.source.package.relationships(source_part).items():
            if rel.is_external:
                new_id = package.add_external_relationship(target, rel.type, rel.target)
            elif rel.target_part is not None and self.source.package.has_part(rel.target_part):
                new_id = package.add_relationship(target, rel.type, self.copy_part(rel.target_part))
            else:
                continue
            if new_id != rel_id:
                self._rename_relationship(target, rel_id, new_id)
        return target

    def _rename_relationship(self, part: str, old: str, new: str) -> None:
        """A copied part's relationship got another id: its references follow."""
        try:
            root = self.dest.package.tree(part)
        except etree.XMLSyntaxError:
            return
        if root is None:
            return
        for node in root.iter():
            if not isinstance(node.tag, str):
                continue
            for key, value in list(node.attrib.items()):
                if (key.startswith(_R) or key == _O + "relid") and value == old:
                    node.set(key, new)
        self.dest.package.mark_dirty(part)

    # -- notes ------------------------------------------------------------------------------

    def _notes(self, elements: list[Element]) -> None:
        for kind in ("footnote", "endnote"):
            references = [node for element in elements for node in element.iter(_W + f"{kind}Reference")]
            if not references:
                continue
            source_part = self.source._notes_part(kind)
            source_root = None if source_part is None else self.source.package.tree(source_part)
            if source_root is None:
                for node in references:
                    remove(node)
                continue
            dest_part = self.dest._notes_part(kind, create=True)
            dest_root = self.dest.package.tree(dest_part)
            used = [int(n.get(_W + "id")) for n in dest_root.findall(_W + kind)
                    if (n.get(_W + "id") or "").lstrip("-").isdigit()]
            next_id = max(used + [0]) + 1
            by_id = {n.get(_W + "id"): n for n in source_root.findall(_W + kind)}
            prefix = "fn" if kind == "footnote" else "en"
            for reference in references:
                old = reference.get(_W + "id")
                note = by_id.get(old)
                if note is None:
                    continue
                copied = copy.deepcopy(note)
                copied.set(_W + "id", str(next_id))
                reference.set(_W + "id", str(next_id))
                content = [child for child in copied if isinstance(child.tag, str)]
                for child in content:
                    _current_view(child)
                self._formatting(content)
                self._lists(content)
                self._relationships(content, source_part, dest_part)
                self._bookmarks_in(content)
                self._drawing_ids(content)
                self._control_ids(content)
                dest_root.append(copied)
                self.notes_created.append((kind, next_id, dest_part, copied))
                self.copied[f"{prefix}:{old}"] = f"{prefix}:{next_id}"
                next_id += 1
            self.dest.package.mark_dirty(dest_part)

    # -- comments ---------------------------------------------------------------------------

    def _comments(self, elements: list[Element]) -> None:
        markers = [node for element in elements
                   for node in element.iter(_W + "commentRangeStart", _W + "commentRangeEnd", _W + "commentReference")]
        if not markers:
            return
        source_root = self.source._comment_tree("comments")
        by_id = {} if source_root is None else {n.get(_W + "id"): n for n in source_root.findall(_W + "comment")}
        mapping: dict[str, str] = {}
        for node in markers:
            old = node.get(_W + "id")
            if old not in by_id:
                parent = node.getparent()
                if node.tag == _W + "commentReference" and parent is not None and parent.tag == _W + "r" \
                        and len(parent) <= 2:
                    remove(parent)
                else:
                    remove(node)
                continue
            if old not in mapping:
                mapping[old] = str(self._annotation)
                self._annotation += 1
            node.set(_W + "id", mapping[old])
        if not mapping:
            return
        source_part = self.source._comment_part("comments")
        dest_root = self.dest._comment_tree("comments", create=True)
        dest_part = self.dest._comment_part("comments")
        extended = self.source._comment_tree("extended")
        source_ex = {} if extended is None else {(n.get(_W15 + "paraId") or "").upper(): n
                                                 for n in extended.findall(_W15 + "commentEx")}
        para_map: dict[str, str] = {}
        new_comments = []
        used = self.dest._used()
        for old, new in mapping.items():
            comment = copy.deepcopy(by_id[old])
            comment.set(_W + "id", new)
            old_last = _last_para(comment)
            content = [child for child in comment if isinstance(child.tag, str)]
            self._formatting(content)
            self._lists(content)
            self._relationships(content, source_part, dest_part)
            for paragraph in comment.iter(W_P):
                raw = paragraph.get(_ids.PARA_ID)
                para_id, text_id = _ids.generate(dest_part + "\0import", used, 2)
                _ids.stamp(paragraph, para_id, text_id)
                if raw:
                    para_map[raw.upper()] = para_id
            dest_root.append(comment)
            new_comments.append((comment, old_last))
            if comment.get(_W + "author"):
                self.dest._ensure_person(comment.get(_W + "author"))
            self.copied[f"comment#{old}"] = f"comment#{new}"
        extended_dest = self.dest._comment_tree("extended", create=True)
        for comment, old_last in new_comments:
            entry = source_ex.get((old_last or "").upper())
            item = etree.SubElement(extended_dest, _W15 + "commentEx")
            item.set(_W15 + "paraId", _last_para(comment))
            parent = None if entry is None else entry.get(_W15 + "paraIdParent")
            if parent and parent.upper() in para_map:
                item.set(_W15 + "paraIdParent", para_map[parent.upper()])
            item.set(_W15 + "done", (entry.get(_W15 + "done") if entry is not None else None) or "0")
        self.dest._complete_comment_parts()
        _ids.ensure_w14(dest_root)
        for key in ("comments", "extended", "ids", "extensible"):
            part = self.dest._comment_part(key)
            if part is not None:
                self.dest.package.mark_dirty(part)

    # -- bookmarks, drawings, controls -----------------------------------------------------

    def _dest_bookmark_names(self) -> list[str]:
        return [node.get(_W + "name") or "" for part in self.dest._parts()
                for node in self.dest.package.tree(part).iter(_W + "bookmarkStart")]

    def _bookmarks_in(self, elements: list[Element]) -> None:
        starts = [n for element in elements for n in element.iter(_W + "bookmarkStart")]
        ends = [n for element in elements for n in element.iter(_W + "bookmarkEnd")]
        ids: dict[str, str] = {}
        dropped: set[str] = set()
        for start in starts:
            name = start.get(_W + "name") or ""
            old_id = start.get(_W + "id")
            if name in self._bookmarks:
                if self.bookmark_policy == "drop" or name == "_GoBack":
                    dropped.add(old_id)
                    remove(start)
                    continue
                new_name = self._free_bookmark(name)
                self._renamed_bookmarks[name] = new_name
                start.set(_W + "name", new_name)
                self.warnings.append(f"bookmark {name!r} renamed {new_name!r}: the destination has one")
                name = new_name
            self._bookmarks.add(name)
            ids[old_id] = str(self._annotation)
            self._annotation += 1
            start.set(_W + "id", ids[old_id])
        for end in ends:
            old_id = end.get(_W + "id")
            if old_id in ids and old_id not in dropped:
                end.set(_W + "id", ids[old_id])
            else:
                remove(end)  # dropped, or its start was not copied
        if not self._renamed_bookmarks:
            return
        from .links import _rename_in_instruction

        for element in elements:
            for link in element.iter(_W + "hyperlink"):
                anchor = link.get(_W + "anchor")
                if anchor in self._renamed_bookmarks:
                    link.set(_W + "anchor", self._renamed_bookmarks[anchor])
            for node in element.iter(_W + "instrText"):
                for old, new in self._renamed_bookmarks.items():
                    node.text = _rename_in_instruction(node.text or "", old, new)
            for node in element.iter(_W + "fldSimple"):
                for old, new in self._renamed_bookmarks.items():
                    node.set(_W + "instr", _rename_in_instruction(node.get(_W + "instr") or "", old, new))

    def _free_bookmark(self, name: str) -> str:
        number = 1
        while True:
            suffix = f"_{number}"
            candidate = name[:40 - len(suffix)] + suffix
            if candidate not in self._bookmarks:
                return candidate
            number += 1

    def _drawing_ids(self, elements: list[Element]) -> None:
        if self._doc_prs is None:
            self._doc_prs = set()
            for part in self.dest._parts():
                for node in self.dest.package.tree(part).iter(_WP + "docPr"):
                    try:
                        self._doc_prs.add(int(node.get("id")))
                    except (TypeError, ValueError):
                        pass
        for element in elements:
            for node in element.iter(_WP + "docPr"):
                try:
                    value = int(node.get("id"))
                except (TypeError, ValueError):
                    value = None
                if value is None or value in self._doc_prs:
                    value = max(self._doc_prs, default=0) + 1
                    node.set("id", str(value))
                self._doc_prs.add(value)

    def _control_ids(self, elements: list[Element]) -> None:
        if self._controls is None:
            self._controls = set()
            for _, _, sdt in self.dest._content_controls():
                node = sdt.find(f"{_W}sdtPr/{_W}id")
                try:
                    self._controls.add(abs(int(node.get(_W + "val"))))
                except (AttributeError, TypeError, ValueError):
                    pass
        for element in elements:
            for node in element.iter(_W + "id"):
                parent = node.getparent()
                if parent is None or parent.tag != _W + "sdtPr":
                    continue
                try:
                    value = int(node.get(_W + "val"))
                except (TypeError, ValueError):
                    continue
                if abs(value) in self._controls:
                    old = value
                    value = int(_ids.generate(f"control\0{old}\0{len(self._controls)}", set(), 1)[0], 16)
                    value = value % 2_000_000_000 + 1
                    while value in self._controls:
                        value = (value * 7 + 13) % 2_000_000_000 + 1
                    node.set(_W + "val", str(value))
                    self.copied[f"sdt#{old}"] = f"sdt#{value}"
                self._controls.add(abs(value))


def _set_or_drop(node: Element, value: str | None) -> None:
    """Point a style reference at ``value``, or remove it (``None``: the default style)."""
    if value is None:
        remove(node)
    else:
        node.set(_W + "val", value)


def _last_para(comment: Element) -> str:
    paragraphs = list(comment.iter(W_P))
    return paragraphs[-1].get(_ids.PARA_ID, "") if paragraphs else ""


def _current_view(element: Element) -> None:
    """The source's current view of a copied block: insertions kept, deletions dropped,
    records of changes dropped."""
    for node in list(element.iter(_W + "del", _W + "moveFrom")):
        parent = node.getparent()
        if parent is None:
            continue
        if parent.tag == _W + "trPr" and node.tag == _W + "del":
            row = parent.getparent()
            if row is not None and row.getparent() is not None:
                remove(row)
                continue
        remove(node)
    for node in list(element.iter(_W + "ins", _W + "moveTo")):
        parent = node.getparent()
        if parent is None:
            continue
        if parent.tag in (_W + "rPr", _W + "trPr", _W + "tcPr"):
            remove(node)
            continue
        for child in list(node):
            node.addprevious(child)
        remove(node)
    for tag in ("rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange",
                "tblGridChange", "numberingChange", "cellIns", "cellDel", "cellMerge",
                "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"):
        for node in list(element.iter(_W + tag)):
            if node.getparent() is not None:
                remove(node)


# -- the operation -------------------------------------------------------------------------------


def _source_blocks(source: "Document", ids) -> tuple[str, list[Element]]:
    """The source blocks ``ids`` names: one id, ``first..last`` (blocks of one container),
    or a list of either; all in one story."""
    items = [ids] if isinstance(ids, str) else list(ids)
    if not items:
        raise EditError("copy_blocks needs at least one block id")
    part: str | None = None
    elements: list[Element] = []
    for item in items:
        if not isinstance(item, str):
            raise EditError("a block to copy is named by its id, or 'first..last'")
        first, sep, last = item.partition("..")
        where, entry = source._resolve(first)
        if part is None:
            part = where
        elif where != part:
            raise EditError("the blocks copied are in one story")
        chosen = [entry.element]
        if sep:
            other_part, other = source._resolve(last)
            if other_part != part or other.element.getparent() is not entry.element.getparent():
                raise EditError(f"{item!r}: a range copied is blocks of one container")
            siblings = [c for c in entry.element.getparent() if isinstance(c.tag, str) and c.tag in _BLOCKS]
            a, b = siblings.index(entry.element), siblings.index(other.element)
            if b < a:
                a, b = b, a
            chosen = siblings[a:b + 1]
        for element in chosen:
            if all(element is not e for e in elements):
                elements.append(element)
    return part, elements


def _block_ids(document: "Document", part: str, elements: list[Element]) -> list[list[str]]:
    """Per element, the ids of what it holds that ``copied`` maps: the block itself
    (``p:`` or ``t:``), and every paragraph in it."""
    index = document._index(part)
    by_element = {id(e.element): e.id for e in index.paragraphs}
    by_element.update({id(e.element): e.id for e in index.tables})
    out = []
    for element in elements:
        ids = [by_element.get(id(element), "")]
        ids += [by_element.get(id(p), "") for p in element.iter(W_P) if p is not element]
        out.append(ids)
    return out


class ImportOps:
    """:meth:`Document.copy_blocks`."""

    def copy_blocks(self: "Document", source: "Document", ids, *, at: str = "end",
                    styles: str = "use_destination", lists: str = "separate",
                    bookmarks: str = "rename", style_map: dict | None = None,
                    unmapped: str = "import") -> "EditResult":
        """Copy blocks of ``source`` (another :class:`Document`) here, with what they depend
        on (:mod:`docx_agent.edit.importing` says how, measured against Word).

        ``ids``: a paragraph or table id, ``"first..last"`` (blocks of one container), or a
        list of them.  ``at``: as :meth:`insert_markdown` takes it -- ``"end"``,
        ``"end:<story>"``, ``"after:<id>"``, ``"before:<id>"``, ``"replace:<id>[..<id>]"``.
        ``styles``: ``"use_destination"`` (Word's default paste), ``"keep_source"`` or
        ``"merge"``; ``lists``: ``"separate"`` or ``"continue"``; ``bookmarks``:
        ``"rename"`` or ``"drop"`` (Word's).  ``style_map``: source style name ->
        destination style name (``{"Bulletin Body": "Handbook Body"}``); ``unmapped``:
        ``"import"`` (Word's) or ``"body"`` -- a source style neither mapped nor in the
        destination by name becomes the destination's body style instead of being imported
        (a built-in one is added as Word writes it)::

            handbook.copy_blocks(bulletin, bulletin.section_blocks("p:3B212964"), at="before:p:69E36C9F",
                                 unmapped="body")

        ``"merge"`` keeps a heading's level and a list item's list.  One undo step; tracked
        where the document tracks.  ``created`` lists every new paragraph and table id, ``blocks`` the
        top-level ones, ``copied`` maps each source id to its copy's, ``warnings`` name
        the styles imported and the bookmarks renamed."""
        from ..markdown.write import _is_last, _locate, _next_block, _place
        from .document import Document, EditResult

        if not isinstance(source, Document):
            raise EditError("copy_blocks copies from another Document")
        if source is self:
            raise EditError("copy_blocks copies between documents: within one, use move_block")
        if styles not in STYLE_POLICIES:
            raise EditError(f"styles is one of {', '.join(STYLE_POLICIES)}")
        if lists not in LIST_POLICIES:
            raise EditError(f"lists is one of {', '.join(LIST_POLICIES)}")
        if bookmarks not in BOOKMARK_POLICIES:
            raise EditError(f"bookmarks is one of {', '.join(BOOKMARK_POLICIES)}")
        if unmapped not in UNMAPPED_POLICIES:
            raise EditError(f"unmapped is one of {', '.join(UNMAPPED_POLICIES)}")
        if style_map is not None and not (isinstance(style_map, dict)
                                          and all(isinstance(v, str) for v in style_map.values())):
            raise EditError("style_map maps source style names to destination style names")
        source_part, originals = _source_blocks(source, ids)
        source_ids = _block_ids(source, source_part, originals)
        place = _locate(self, at)
        if self.package.tree(place.part).tag in (_W + "footnotes", _W + "endnotes", _W + "comments"):
            raise EditError("blocks are copied into the body, a header or a footer, or a cell")
        tracking = self._active_tracking()
        removed: list[str] = []
        with self._edit():
            renames = self._prepare(place.part, [], place.position)
            if place.replaced and tracking is not None:
                last = self._resolve(place.replaced[-1])[1].element
                following = _next_block(last)
                for identifier in place.replaced:
                    removed += self.delete_block(identifier).removed
                place.before = following
            if place.before is not None:
                anchor = place.before.getprevious()
            else:
                blocks = [c for c in place.container if isinstance(c.tag, str) and c.tag in _BLOCKS]
                anchor = blocks[-1] if blocks else None
            importer = Importer(self, source, styles=styles, lists=lists, bookmarks=bookmarks,
                                at_part=place.part, anchor=anchor, style_map=style_map, unmapped=unmapped)
            elements = [copy.deepcopy(element) for element in originals]
            importer.bring(elements, source_part, place.part)
            _place(place, elements)
            if place.replaced and tracking is None:
                for identifier in place.replaced:
                    removed += self.delete_block(identifier).removed
            trailing = None
            if _is_last(elements[-1]) and elements[-1].tag == W_TBL:
                trailing = make("w:p")
                elements[-1].addnext(trailing)
                elements.append(trailing)
            created, per_element = _stamp(self, place.part, elements)
            for _, _, notes_part, note in importer.notes_created:
                _stamp_notes(self, notes_part, note, created)
            if tracking is not None:
                _track(self, place.part, elements, importer, tracking)
            self.package.mark_dirty(place.part)
        copied = dict(importer.copied)
        for old_ids, new_ids in zip(source_ids, per_element):
            for old, new in zip(old_ids, new_ids):
                if old and new:
                    copied[old] = new
        top = [ids_[0] for ids_ in per_element if ids_ and ids_[0]]
        return EditResult(top[0] if top else None, created=created, renamed=renames, removed=removed,
                          warnings=importer.warnings, blocks=top, copied=copied)


def _stamp(document: "Document", part: str, elements: list[Element]) -> tuple[list[str], list[list[str]]]:
    """Keep a copy's paraIds, textIds and drawing ids where the destination does not use
    them; re-issue the rest.  Returns every id created in document order, and per element
    its own id followed by its paragraphs' (as :func:`_block_ids` lists the source's)."""
    root = document.package.tree(part)
    # The proxies are kept alive while their ids are compared: lxml makes a new one (with
    # another id) for a node whose proxy has gone.
    nodes = [node for element in elements for node in element.iter()]
    placed = {id(node) for node in nodes}
    used: set[int] = set()
    for node in root.iter(W_P, _W + "tr", _WP + "inline", _WP + "anchor"):
        if id(node) in placed:
            continue
        for name in (_ids.PARA_ID, _ids.TEXT_ID, _ids.ANCHOR_ID, _ids.EDIT_ID):
            raw = node.get(name)
            if raw is not None and _ids.valid_long_hex(raw):
                used.add(int(raw, 16))
    for other in document._parts():
        if other != part:
            used |= _ids.used_long_hex([document.package.tree(other)])
    story = document._story_of(part)
    prefix = "" if story == "body" else f"{story}/"
    drawings = False

    def keep_or_issue(node: Element, names: tuple[str, str], seed: str) -> str:
        values: list[str | None] = []
        for name in names:
            raw = node.get(name)
            if raw is not None and _ids.valid_long_hex(raw) and int(raw, 16) not in used:
                used.add(int(raw, 16))
                values.append(raw.upper())
            else:
                values.append(None)
        if None in values:
            fresh = iter(_ids.generate(seed, used, values.count(None)))
            values = [value if value is not None else next(fresh) for value in values]
        for name, value in zip(names, values):
            node.set(name, value)
        return values[0]

    created: list[str] = []
    per_element: list[list[str]] = []
    for element in elements:
        first_row = None
        paragraphs: list[str] = []
        own = None
        for node in element.iter(W_P, _W + "tr", _WP + "inline", _WP + "anchor"):
            if node.tag == W_P:
                para_id = keep_or_issue(node, (_ids.PARA_ID, _ids.TEXT_ID), part)
                if node is element:
                    own = f"{prefix}p:{para_id}"
                else:
                    paragraphs.append(f"{prefix}p:{para_id}")
            elif node.tag == _W + "tr":
                row_id = keep_or_issue(node, (_ids.PARA_ID, _ids.TEXT_ID), part + "\0row")
                first_row = first_row or row_id
            else:
                keep_or_issue(node, (_ids.ANCHOR_ID, _ids.EDIT_ID), part + "\0drawing")
                drawings = True
        if element.tag == W_TBL and first_row is not None:
            own = f"{prefix}t:{first_row}"
        elif own is None and paragraphs:
            own = paragraphs[0]  # a block control: its first paragraph
        ids = ([own] if own else [""]) + paragraphs
        per_element.append(ids)
        created += [i for i in ids if i and i not in created]
    _ids.ensure_w14(root, drawings=drawings)
    return created, per_element


def _stamp_notes(document: "Document", part: str, note: Element, created: list[str]) -> None:
    used = document._used()
    story = document._story_of(part)
    for node in note.iter(W_P, _W + "tr"):
        para_id, text_id = _ids.generate(part + "\0import", used, 2)
        _ids.stamp(node, para_id, text_id)
        if node.tag == W_P:
            created.append(f"{story}/p:{para_id}")
    _ids.ensure_w14(document.package.tree(part),
                    drawings=any(True for _ in note.iter(_WP + "inline", _WP + "anchor")))


def _track(document: "Document", part: str, elements: list[Element], importer: Importer, tracking) -> None:
    """The copy as one insertion (Word's tracked paste: every mark, run and row)."""
    from ..markdown.write import _insert_content, _track_table
    from ..revisions import track as _track_mod
    from ..revisions.stamp import Stamp

    stamp = Stamp(document, tracking)
    for element in elements:
        if element.tag == W_P:
            _insert_content(element, stamp, part)
            _track_mod.mark(element, "ins", stamp, part)
        elif element.tag == W_TBL:
            _track_table(element, stamp, part)
        else:
            for paragraph in element.iter(W_P):
                _insert_content(paragraph, stamp, part)
                _track_mod.mark(paragraph, "ins", stamp, part)
    last = elements[-1]
    if last.tag == W_P and _track_mod.is_last_in_container(last):
        previous = elements[0].getprevious()
        while previous is not None and not (isinstance(previous.tag, str) and previous.tag in _BLOCKS):
            previous = previous.getprevious()
        if previous is not None and previous.tag == W_P and _track_mod.mark_record(previous) is None \
                and previous.find(f"{_W}pPr/{_W}sectPr") is None:
            # E3's typing form at a container's end: the mark before is the inserted one.
            record = _track_mod.mark_record(last)
            run_properties = record.getparent()
            remove(record)
            if not len(run_properties):
                remove(run_properties)
            _track_mod.mark(previous, "ins", stamp, part)
            change = previous.find(f"{_W}pPr/{_W}pPrChange/{_W}pPr")
            old = _track_mod.clean_properties(change if change is not None else previous.find(_W + "pPr"), "w:pPr")
            _track_mod.record_paragraph_change(last, old, stamp, part)
    for _, _, notes_part, note in importer.notes_created:
        paragraphs = [c for c in note if isinstance(c.tag, str)]
        for k, element in enumerate(paragraphs):
            if element.tag == W_P:
                _insert_content(element, stamp, notes_part)
                if k + 1 < len(paragraphs):
                    _track_mod.mark(element, "ins", stamp, notes_part)
            else:
                _track_table(element, stamp, notes_part)
    stamp.finish()


def copy_blocks(source: "Document", ids, *, to: "Document", at: str = "end", **options) -> "EditResult":
    """:meth:`Document.copy_blocks` with the source first: copy ``ids`` of ``source`` into
    ``to``."""
    return to.copy_blocks(source, ids, at=at, **options)


__all__ = ["BOOKMARK_POLICIES", "ImportOps", "Importer", "LIST_POLICIES", "STYLE_POLICIES", "copy_blocks"]
