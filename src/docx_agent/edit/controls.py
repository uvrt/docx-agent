"""Content controls (``w:sdt``): read, filled, inserted and removed -- and a data-bound one's
custom XML part kept in step with what it shows.

What Word 16.106 does (``tools/e5_probe.py``; Word's AppleScript has no content-control
object, so the controls were written here as Word 365 writes them and Word typed into them):

* **Filling by typing** replaces the content's text in the content's first run; **tracked**,
  the old text ``w:del`` and the new ``w:ins`` *inside* ``w:sdtContent``.  A date control's
  ``w:fullDate`` follows a date typed into it (untracked); a drop-down list refuses typing
  (its content stays one of its items).
* **A data-bound control** shows what its custom XML node says: Word replaced a stale
  cached content with the node's text on opening.  So a fill writes the node as well (as
  docx4j does) -- otherwise Word puts the old value back.  Word's own typing into a bound
  control left the node unchanged in the saved file (measured): the cache and the node
  disagreed, and the next opening shows the node's.
* **Re-saving** keeps every control, adds an empty ``w:sdtEndPr`` after ``w:sdtPr`` and
  writes ``w:tag`` before ``w:id`` (the schema's order); docx-agent writes them so.

docx-agent fills plain and rich text, drop-down lists (an item's display text, the node its
value), combo boxes (an item, or any text), dates (the date in the control's own picture and
``w:fullDate``) and check boxes (``w14:checked`` and its state's glyph); respects ``w:lock``
(content locked: no fill; control locked: no removal); and, after a review accepts or
rejects a fill, writes the bound node from what the control then shows.  Inserting and
removing a control is not a revision in Word (it has none for the wrapper): tracked, the
content's text is inserted or deleted, the wrapper is not, and the result says so.
"""

from __future__ import annotations

import copy
import datetime as _dt
import re
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, insert_in_order, make, remove
from . import ids as _ids
from . import text as _text

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_W14 = "{%s}" % W14
DS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
REL_CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
REL_CUSTOM_XML_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
KINDS = ("text", "rich-text", "drop-down", "combo-box", "date", "checkbox")
#: Word's check box glyphs (MS Gothic) as it writes a new one.
CHECKED, UNCHECKED, CHECK_FONT = "2612", "2610", "MS Gothic"


def _sdt_child(sdt: Element, tag: str) -> Element | None:
    properties = sdt.find(_W + "sdtPr")
    return properties.find(tag) if properties is not None else None


def items(sdt: Element) -> list[tuple[str, str]]:
    """A list control's items: ``(display text, value)``."""
    for tag in (_W + "dropDownList", _W + "comboBox"):
        node = _sdt_child(sdt, tag)
        if node is not None:
            return [(item.get(_W + "displayText") or item.get(_W + "value") or "", item.get(_W + "value") or "")
                    for item in node.findall(_W + "listItem")]
    return []


def binding(sdt: Element) -> dict | None:
    node = _sdt_child(sdt, _W + "dataBinding")
    if node is None:
        return None
    return {"xpath": node.get(_W + "xpath"), "store": node.get(_W + "storeItemID"),
            "prefixes": node.get(_W + "prefixMappings") or ""}


def lock(sdt: Element) -> str | None:
    node = _sdt_child(sdt, _W + "lock")
    return node.get(_W + "val") if node is not None else None


def checked(sdt: Element) -> bool | None:
    box = _sdt_child(sdt, _W14 + "checkbox")
    if box is None:
        return None
    state = box.find(_W14 + "checked")
    return state is not None and (state.get(_W14 + "val") or "0") in ("1", "true")


def _prefixes(mappings: str) -> dict[str, str]:
    return dict(re.findall(r"xmlns:(\w+)\s*=\s*['\"]([^'\"]*)['\"]", mappings))


def _date_text(sdt: Element, moment: _dt.date) -> str:
    from .fields import format_date

    picture = _sdt_child(sdt, _W + "date")
    fmt = picture.find(_W + "dateFormat") if picture is not None else None
    return format_date(fmt.get(_W + "val") if fmt is not None else "d-M-yyyy", moment)


def _parse_date(value) -> _dt.date:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    return _dt.date.fromisoformat(str(value)[:10])


class ControlOps:
    """Content controls' values, on :class:`docx_agent.Document`."""

    def _control(self: "Document", identifier: str) -> tuple[str, Element]:
        from .document import EditError

        for part, found, element in self._content_controls():
            if found == identifier:
                return part, element
        raise EditError(f"no content control {identifier!r}")

    # -- reading ----------------------------------------------------------------------------

    def control_value(self: "Document", identifier: str):
        """What a control holds: a check box's ``True``/``False``, a date control's date
        (``w:fullDate``, or ``None``), a list's chosen item's value where its text is an
        item's, else the text."""
        from .annotations import ContentControl, control_kind

        part, sdt = self._control(identifier)
        kind = control_kind(sdt)
        if kind == "checkbox":
            return checked(sdt)
        text = ContentControl(self, identifier).text
        if kind == "date":
            node = _sdt_child(sdt, _W + "date")
            full = node.get(_W + "fullDate") if node is not None else None
            return _dt.date.fromisoformat(full[:10]) if full else None
        if kind in ("drop-down", "combo-box"):
            for display, value in items(sdt):
                if display == text:
                    return value
        return text

    # -- filling ----------------------------------------------------------------------------

    def fill_control(self: "Document", identifier: str, value) -> "EditResult":
        """Fill a control as typing into it would: plain or rich text (a line per paragraph
        in a block-level control), a drop-down list's item (its display text or value), a
        combo box's item or any text, a date (a date, a datetime or ISO 8601: shown in the
        control's own picture, ``w:fullDate`` set), a check box (``True``/``False``).  The
        content keeps its first run's formatting; a placeholder shown goes.  A data-bound
        control's custom XML node is written too.  Tracked, the old text is deleted and the
        new inserted inside the control, as Word records typing there."""
        from .annotations import control_kind, control_level
        from .document import EditError, EditResult

        part, sdt = self._control(identifier)
        kind = control_kind(sdt)
        if kind not in KINDS:
            raise EditError(f"{identifier} is a {kind} control: its content is not filled as text")
        if lock(sdt) in ("contentLocked", "sdtContentLocked"):
            raise EditError(f"{identifier}'s content is locked")
        node_value = None
        if kind == "checkbox":
            if not isinstance(value, bool):
                raise EditError("a check box is filled with True or False")
            box = _sdt_child(sdt, _W14 + "checkbox")
            state = box.find(_W14 + ("checkedState" if value else "uncheckedState"))
            text = chr(int(state.get(_W14 + "val") if state is not None else (CHECKED if value else UNCHECKED), 16))
            node_value = "true" if value else "false"
        elif kind == "date":
            try:
                moment = _parse_date(value)
            except (TypeError, ValueError):
                raise EditError(f"{value!r} is not a date") from None
            text = _date_text(sdt, moment)
            node_value = moment.isoformat() + "T00:00:00"
        elif kind in ("drop-down", "combo-box"):
            text = str(value)
            choice = next(((d, v) for d, v in items(sdt) if text in (d, v)), None)
            if choice is None and kind == "drop-down":
                raise EditError(f"{value!r} is not one of {identifier}'s items {[d for d, _ in items(sdt)]}")
            if choice is not None:
                text, node_value = choice
            else:
                node_value = text
        else:
            text = str(value)
            node_value = text
        lines = text.split("\n")
        for line in lines:
            _text.check_text(line)
        if len(lines) > 1 and control_level(sdt) != "block":
            raise EditError(f"{identifier} sits in a line: one line of text")
        bound = binding(sdt)
        if bound is not None:
            self._bound_node(bound)  # refuse before anything changes
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare_many([])
            part, sdt = self._control(identifier)
            if control_level(sdt) == "block":
                self._fill_block(part, sdt, lines, tracking)
            else:
                self._fill_inline(part, sdt, text, tracking)
            properties = sdt.find(_W + "sdtPr")
            for node in properties.findall(_W + "showingPlcHdr"):
                remove(node)
            if kind == "checkbox":
                box = _sdt_child(sdt, _W14 + "checkbox")
                state = box.find(_W14 + "checked")
                if state is None:
                    state = etree.SubElement(box, _W14 + "checked")
                    box.insert(0, state)
                state.set(_W14 + "val", "1" if value else "0")
            if kind == "date":
                _sdt_child(sdt, _W + "date").set(_W + "fullDate", node_value.replace("T00:00:00", "T00:00:00Z"))
            if bound is not None:
                self._write_bound(bound, node_value)
            self.package.mark_dirty(part)
        return EditResult(identifier, renamed=renames)

    def _fill_inline(self: "Document", part: str, sdt: Element, text: str, tracking) -> None:
        content = sdt.find(_W + "sdtContent")
        if content is None:
            content = make("w:sdtContent")
            sdt.append(content)
        runs = [run for run in content.iter(_W + "r")]
        template = next((r for r in runs if r.find(_W + "t") is not None or r.find(_W + "delText") is not None),
                        runs[0] if runs else None)
        properties = copy.deepcopy(template.find(_W + "rPr")) if template is not None and \
            template.find(_W + "rPr") is not None else None
        if properties is None:
            sdt_rpr = _sdt_child(sdt, _W + "rPr")
            properties = copy.deepcopy(sdt_rpr) if sdt_rpr is not None else None
        if properties is not None:
            for node in list(properties):
                if node.tag in (_W + "rStyle",) and node.get(_W + "val") in ("PlaceholderText", "Tekstvantijdelijkeaanduiding"):
                    remove(node)
                elif node.tag == _W + "rPrChange":
                    remove(node)
            if not len(properties):
                properties = None
        new_run = _text.make_run(text, properties) if text else None
        if tracking is None:
            for child in list(content):
                remove(child)
            if new_run is not None:
                content.append(new_run)
            return
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp

        stamp = Stamp(self, tracking)
        shown = [r for r in content.iter(_W + "r") if _track.revision_owner(r) is None
                 or _track.revision_owner(r).tag in _track.INSERTED]
        if shown:
            self._release(part, _track.delete_runs(shown, stamp, part))
        if new_run is not None:
            _track.insert_nodes(content, len(content), [new_run], stamp, part)
        stamp.finish()

    def _fill_block(self: "Document", part: str, sdt: Element, lines: list[str], tracking) -> None:
        content = sdt.find(_W + "sdtContent")
        paragraphs = [p for p in content if isinstance(p.tag, str) and p.tag == _W + "p"]
        index = self._index(part)
        if not paragraphs:
            paragraph = make("w:p")
            content.append(paragraph)
            para_id, text_id = self._fresh_ids(part)
            _ids.stamp(paragraph, para_id, text_id)
            self._invalidate()
            index = self._index(part)
            paragraphs = [paragraph]
        ids = [index.entry_for(p).id for p in paragraphs]
        # Each paragraph takes a line (keeping its properties); the lines left over follow
        # the last; the paragraphs left over go from the end.
        for identifier, line in zip(ids, lines):
            self._set_text(identifier, line)
        last = ids[min(len(ids), len(lines)) - 1]
        for line in lines[len(ids):]:
            last = self.insert_paragraph(line, after=last).id
        for identifier in reversed(ids[len(lines):]):
            self.delete_block(identifier)

    # -- the custom XML part ----------------------------------------------------------------

    def _store_parts(self: "Document") -> dict[str, str]:
        """``{store item id (upper case): custom XML part}``."""
        main = self.package.document_part()
        out = {}
        for part in self.package.related_parts_of_type(main, REL_CUSTOM_XML):
            if not self.package.has_part(part):
                continue
            for props in self.package.related_parts_of_type(part, REL_CUSTOM_XML_PROPS):
                root = self.package.tree(props) if self.package.has_part(props) else None
                if root is not None and root.get("{%s}itemID" % DS):
                    out[root.get("{%s}itemID" % DS).upper()] = part
        return out

    def _bound_node(self: "Document", bound: dict) -> tuple[str, Element]:
        from .document import EditError

        stores = self._store_parts()
        part = stores.get((bound["store"] or "").upper())
        if part is None:
            if len(stores) == 1 and not bound["store"]:
                part = next(iter(stores.values()))
            else:
                raise EditError(f"the custom XML part {bound['store']} the control is bound to is missing")
        root = self.package.tree(part)
        try:
            found = etree.ElementTree(root).xpath(bound["xpath"], namespaces=_prefixes(bound["prefixes"]))
        except etree.XPathError as error:
            raise EditError(f"the control's binding {bound['xpath']!r} does not evaluate: {error}") from None
        if not found or not isinstance(found[0], etree._Element):
            raise EditError(f"the bound node {bound['xpath']!r} is not in the custom XML part "
                            "(docx-agent writes a binding's node, it does not make one)")
        return part, found[0]

    def _write_bound(self: "Document", bound: dict, value: str) -> None:
        part, node = self._bound_node(bound)
        if node.text != value:
            for child in list(node):
                node.remove(child)
            node.text = value
            self.package.mark_dirty(part)

    def _sync_bound_controls(self: "Document") -> None:
        """Every data-bound control's node set to what the control shows (after a review
        changed what it shows): Word fills a bound control from its node on opening."""
        from .annotations import ContentControl, control_kind

        for part, identifier, sdt in self._content_controls():
            bound = binding(sdt)
            if bound is None:
                continue
            try:
                _, node = self._bound_node(bound)
            except Exception:
                continue
            kind = control_kind(sdt)
            if kind == "checkbox":
                value = "true" if checked(sdt) else "false"
            elif kind == "date":
                date = _sdt_child(sdt, _W + "date")
                full = date.get(_W + "fullDate") if date is not None else None
                value = full.replace("Z", "") if full else ContentControl(self, identifier).text
            elif kind in ("drop-down", "combo-box"):
                text = ContentControl(self, identifier).text
                value = next((v for d, v in items(sdt) if d == text), text)
            else:
                value = ContentControl(self, identifier).text
            if node.text != value and not len(node):
                self._write_bound(bound, value)

    # -- inserting and removing -------------------------------------------------------------

    def insert_control(self: "Document", at, kind: str = "text", *, text: str = "", tag: str | None = None,
                       alias: str | None = None, items: list | None = None, date=None, checked: bool = False,
                       date_format: str = "d-M-yyyy", after: str | None = None,
                       before: str | None = None) -> "EditResult":
        """A content control as Word 365 writes one: at a position (``at``, a range or its
        id; in the line), or a block-level one after or before a block (``at=None``,
        ``after=``/``before=``; rich or plain text, a line per paragraph).  ``kind`` is
        ``text``, ``rich-text``, ``drop-down``, ``combo-box`` (``items``: display texts, or
        ``(display, value)`` pairs), ``date`` (``date``, ``date_format``) or ``checkbox``
        (``checked``).  The result's id is the control's (``cc:<w:id>``).  Tracked, its
        content is inserted; the control itself is no revision in Word."""
        from .document import EditError, EditResult
        from .ranges import TextRange

        if kind not in KINDS:
            raise EditError(f"kind is one of {', '.join(KINDS)}")
        pairs = [(i, i) if isinstance(i, str) else (str(i[0]), str(i[1])) for i in (items or [])]
        if kind in ("drop-down", "combo-box") and not pairs:
            raise EditError("a list control needs items")
        if kind == "date" and date is not None:
            try:
                date = _parse_date(date)
            except (TypeError, ValueError):
                raise EditError(f"{date!r} is not a date") from None
        content_text = text
        if kind == "drop-down" and not text:
            content_text = pairs[0][0]
        if kind == "date" and date is not None and not text:
            from .fields import format_date

            content_text = format_date(date_format, date)
        if kind == "checkbox":
            content_text = chr(int(CHECKED if checked else UNCHECKED, 16))
        for line in content_text.split("\n"):
            _text.check_text(line)
        block = at is None
        if block and (after is None) == (before is None):
            raise EditError("give a position (at=), or after= or before= a block for a block-level control")
        if not block and "\n" in content_text:
            raise EditError("a control in the line holds one line")
        tracking = self._active_tracking()
        if not block:
            where = at if isinstance(at, TextRange) else self.range(at)
            if where.view != "current":
                raise EditError("edits take the current view")
            part, entries = where._entries()
            entry = entries[0]
        else:
            part, entry = self._resolve(after if after is not None else before)
        with self._edit():
            renames = self._prepare(part, [entry] if not block else [])
            number = self._next_control_id()
            sdt = make("w:sdt")
            properties = make("w:sdtPr")
            sdt.append(properties)
            gothic = None
            if kind == "checkbox":
                gothic = make("w:rPr")
                gothic.append(make("w:rFonts", **{"w:ascii": CHECK_FONT, "w:eastAsia": CHECK_FONT,
                                                  "w:hAnsi": CHECK_FONT, "w:hint": "eastAsia"}))
                properties.append(copy.deepcopy(gothic))
            if alias:
                insert_in_order(properties, make("w:alias", **{"w:val": alias}))
            if tag:
                insert_in_order(properties, make("w:tag", **{"w:val": tag}))
            insert_in_order(properties, make("w:id", **{"w:val": str(number)}))
            insert_in_order(properties, _kind_element(kind, pairs, date, date_format, checked))
            sdt.append(make("w:sdtEndPr"))
            content = make("w:sdtContent")
            sdt.append(content)
            new_paragraphs = []
            if block:
                anchor = entry.element
                used = self._used()
                for line in content_text.split("\n"):
                    paragraph = make("w:p")
                    if line:
                        paragraph.append(_text.make_run(line, None))
                    para_id, text_id = _ids.generate(part + "\0control", used, 2)
                    _ids.stamp(paragraph, para_id, text_id)
                    content.append(paragraph)
                    new_paragraphs.append(paragraph)
                (anchor.addnext if after is not None else anchor.addprevious)(sdt)
            else:
                from . import inline as _inline

                run = _inline.new_run(entry.element, where.start, [])
                _inline.hoist(run)
                properties_run = run.find(_W + "rPr")
                if gothic is not None:
                    if properties_run is not None:
                        remove(properties_run)
                    properties_run = copy.deepcopy(gothic)
                if content_text:
                    content.append(_text.make_run(content_text, copy.deepcopy(properties_run)
                                                  if properties_run is not None else None))
                run.addprevious(sdt)
                remove(run)
                self._renew_text_id(part, entry)
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                if block:
                    for paragraph in new_paragraphs:
                        _track.insert_paragraph_content(paragraph, stamp, part)
                        _track.mark(paragraph, "ins", stamp, part)
                else:
                    runs = [r for r in content if isinstance(r.tag, str) and r.tag == _W + "r"]
                    if runs:
                        for r in runs:
                            remove(r)
                        _track.insert_nodes(content, 0, runs, stamp, part)
                stamp.finish()
            _ids.ensure_w14(self.package.tree(part))
            self.package.mark_dirty(part)
            self._invalidate()
            made = [f"cc:{number}"]
            if new_paragraphs:
                index = self._index(part)
                made += [index.entry_for(p).id for p in new_paragraphs]
        warnings = ["Word has no revision for a control itself: its content is tracked, the control is not"] \
            if tracking is not None else []
        return EditResult(f"cc:{number}", created=made, renamed=renames, warnings=warnings)

    def _next_control_id(self: "Document") -> int:
        used = set()
        for _, _, sdt in self._content_controls():
            node = _sdt_child(sdt, _W + "id")
            try:
                used.add(int(node.get(_W + "val")))
            except (AttributeError, TypeError, ValueError):
                pass
        value = _ids.generate("content-control", set(), 1)[0]
        number = int(value, 16) % 2_000_000_000 + 1
        while number in used or -number in used:
            number = (number * 7 + 13) % 2_000_000_000 + 1
        return number

    def remove_control(self: "Document", identifier: str, *, keep_content: bool = True) -> "EditResult":
        """Remove a content control: its content stays where it was (``keep_content``), or
        goes with it.  A control locked against deletion (``sdtLocked``) is refused.  Word
        has no revision for a control: tracked, the removal is applied untracked (content
        going with it is deleted as tracked text) and the result says so."""
        from .annotations import control_level
        from .document import EditError, EditResult, _rehome_markers

        part, sdt = self._control(identifier)
        if lock(sdt) in ("sdtLocked", "sdtContentLocked"):
            raise EditError(f"{identifier} is locked against deletion")
        level = control_level(sdt)
        if level in ("row", "cell") and not keep_content:
            raise EditError(f"{identifier} holds a table's {level}s: remove them with the table operations")
        if not keep_content and level == "block":
            content = sdt.find(_W + "sdtContent")
            if content is not None and any(True for _ in content.iter(_W + "footnoteReference",
                                                                      _W + "endnoteReference",
                                                                      _W + "commentReference")):
                raise EditError(f"{identifier} holds a note or comment reference")
        tracking = self._active_tracking()
        warnings = []
        with self._edit():
            renames = self._prepare_many([])
            part, sdt = self._control(identifier)
            content = sdt.find(_W + "sdtContent")
            children = [c for c in content] if content is not None else []
            if keep_content or level == "inline":
                if not keep_content:
                    if tracking is not None:
                        from ..revisions import track as _track
                        from ..revisions.stamp import Stamp

                        stamp = Stamp(self, tracking)
                        runs = [r for r in content.iter(_W + "r")]
                        self._release(part, _track.delete_runs(runs, stamp, part))
                        stamp.finish()
                        children = [c for c in content]
                    else:
                        from . import inline as _inline

                        self._release(part, _inline.relationship_ids(content))
                        children = [c for c in children if isinstance(c.tag, str) and c.tag in (
                            _W + "bookmarkStart", _W + "bookmarkEnd", _W + "commentRangeStart",
                            _W + "commentRangeEnd")]
                for child in children:
                    sdt.addprevious(child)
                paragraph = sdt.getparent()
                remove(sdt)
                if paragraph is not None and paragraph.tag == _W + "p":
                    entry = self._index(part).entry_for(paragraph)
                    if entry is not None:
                        self._renew_text_id(part, entry)
            else:
                if tracking is not None:
                    from ..revisions import track as _track
                    from ..revisions.stamp import Stamp

                    stamp = Stamp(self, tracking)
                    for paragraph in [p for p in children if isinstance(p.tag, str) and p.tag == _W + "p"]:
                        self._release(part, _track.delete_content(paragraph, stamp, part))
                        if _track.mark_record(paragraph) is None:
                            _track.mark(paragraph, "del", stamp, part)
                    stamp.finish()
                    for child in children:
                        sdt.addprevious(child)
                    remove(sdt)
                else:
                    from . import inline as _inline

                    _rehome_markers(sdt)
                    self._release(part, _inline.relationship_ids(sdt))
                    remove(sdt)
            if tracking is not None:
                warnings.append("Word has no revision for a control itself: it is removed untracked")
            self.package.mark_dirty(part)
        return EditResult(None, renamed=renames, removed=[identifier], warnings=warnings)


def _kind_element(kind: str, pairs: list[tuple[str, str]], date, date_format: str, is_checked: bool) -> Element:
    if kind == "text":
        return make("w:text")
    if kind == "rich-text":
        return make("w:richText")
    if kind in ("drop-down", "combo-box"):
        node = make("w:dropDownList" if kind == "drop-down" else "w:comboBox")
        for display, value in pairs:
            node.append(make("w:listItem", **{"w:displayText": display, "w:value": value}))
        return node
    if kind == "date":
        node = make("w:date")
        if date is not None:
            node.set(_W + "fullDate", date.isoformat() + "T00:00:00Z")
        node.append(make("w:dateFormat", **{"w:val": date_format}))
        node.append(make("w:lid", **{"w:val": "en-GB"}))
        node.append(make("w:storeMappedDataAs", **{"w:val": "dateTime"}))
        node.append(make("w:calendar", **{"w:val": "gregorian"}))
        return node
    box = etree.Element(_W14 + "checkbox", nsmap={"w14": W14})
    for tag, attributes in (("checked", {"val": "1" if is_checked else "0"}),
                            ("checkedState", {"val": CHECKED, "font": CHECK_FONT}),
                            ("uncheckedState", {"val": UNCHECKED, "font": CHECK_FONT})):
        child = etree.SubElement(box, _W14 + tag)
        for name, value in attributes.items():
            child.set(_W14 + name, value)
    return box
