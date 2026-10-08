"""Footnotes and endnotes: insert, edit, delete and move them; their numbering (E4).

As Word writes them (``tools/e4_probe.py``, ``notes`` and ``notes2``; E2's ``e2_probe.py``
for the first note):

* **A note** is a ``w:footnote`` (``w:endnote``) whose first paragraph, in Footnote Text
  (Endnote Text), opens with the ``w:footnoteRef`` run in Footnote Reference (Endnote
  Reference) and a space; its reference is a run in Footnote Reference holding the
  ``w:footnoteReference``.  A note's id is one above the largest in use (Word renumbers them
  on saving: measured in E0, so ``fn:<id>`` holds for the session).
* **The parts**: a document's first note brings ``footnotes.xml`` *and* ``endnotes.xml``,
  each with its separator (``w:id`` -1) and continuation separator (0), and the settings'
  ``w:footnotePr``/``w:endnotePr`` naming them (E2's measurement).
* **Deleting** a note removes its reference and the note; a notes part left with only its
  separators stays, as Word leaves it (``notes2``).  Tracked, the reference is deleted as a
  revision and the note stays until that is accepted (which removes it with its reference:
  ``review.py``).
* **Moving** a note moves its reference; tracked, the reference is deleted where it was
  and a copy of the note is inserted where it goes -- a reference has one note -- so
  accepting leaves the copy and rejecting the original.
* **Numbering** is the section's (``w:sectPr/w:footnotePr``: ``set_section``'s
  ``footnote_format``...): Word writes the settings' format, start and restart nowhere and
  docx2svg measured that it ignores them there (its 4.12, 4.13).  The endnotes' position
  (``sectEnd``, ``docEnd``) is the document's, in the settings, where docx2svg measured Word
  honouring it; the footnotes' (``beneathText``) is the section's, where Word writes it
  (``notes``).
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from ..oxml.xml import Element, append_in_order, make, remove
from . import ids as _ids
from . import inline as _inline
from . import text as _text
from .errors import EditError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult
    from .ranges import TextRange

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_P = _W + "p"
W_PPR = _W + "pPr"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
_STYLES = {"footnote": ("footnote text", "footnote reference"), "endnote": ("endnote text", "endnote reference")}
POSITIONS = {"footnote": ("pageBottom", "beneathText"), "endnote": ("sectEnd", "docEnd")}


def new_notes_part(document: "Document", kind: str) -> str:
    """A footnotes (or endnotes) part with its two separators, as Word writes it with a
    document's first footnote, its relationship, content type and settings entry
    (``tools/e2_probe.py``)."""
    package = document.package
    plural = kind + "s"
    path = f"word/{plural}.xml"
    if package.has_part(path):
        path = package.unused_part_name(f"word/{plural}{{n}}.xml")
    paragraph = '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:{0}/></w:r></w:p>'
    data = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<w:{plural} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<w:{kind} w:type="separator" w:id="-1">{paragraph.format("separator")}</w:{kind}>'
            f'<w:{kind} w:type="continuationSeparator" w:id="0">{paragraph.format("continuationSeparator")}'
            f'</w:{kind}></w:{plural}>').encode()
    package.add_part(path, data, f"{_WML}.{plural}+xml", override=True)
    package.add_relationship(package.document_part(), REL + plural, path)
    root = package.tree(path)
    used = document._used()
    for paragraph_node in root.iter(W_P):
        para_id, text_id = _ids.generate(path, used, 2)
        _ids.stamp(paragraph_node, para_id, text_id)
    _ids.ensure_w14(root)
    package.mark_dirty(path)
    settings = package.settings_part()
    if settings is not None:
        settings_root = package.tree(settings)
        node = settings_root.find(_W + f"{kind}Pr")
        if node is None:
            node = append_in_order(settings_root, make(f"w:{kind}Pr"))
        if node.find(_W + kind) is None:
            for value in ("-1", "0"):
                append_in_order(node, make(f"w:{kind}", **{"w:id": value}))
        package.mark_dirty(settings)
    document._invalidate()
    return path


class NoteOps:
    """Footnotes and endnotes, on :class:`docx_agent.Document`."""

    def _notes_part(self: "Document", kind: str, *, create: bool = False) -> str | None:
        main = self.package.document_part()
        found = self.package.related_parts_of_type(main, REL + kind + "s")
        if found and self.package.has_part(found[0]):
            return found[0]
        if not create:
            return None
        part = new_notes_part(self, kind)
        other = "endnote" if kind == "footnote" else "footnote"
        if not self.package.related_parts_of_type(main, REL + other + "s"):
            new_notes_part(self, other)
        return part

    def _note_element(self: "Document", identifier: str) -> tuple[str, Element]:
        for part, found, element in self._notes():
            if found == identifier:
                return part, element
        raise KeyError(f"no note {identifier!r}")

    def _note_reference(self: "Document", identifier: str) -> tuple[str, Element] | None:
        """The part and ``w:footnoteReference`` (``w:endnoteReference``) of a note."""
        kind = "footnote" if identifier.startswith("fn:") else "endnote"
        number = identifier.split(":", 1)[1]
        for part in self._parts():
            root = self.package.tree(part)
            if root is None or root.tag in (_W + "footnotes", _W + "endnotes"):
                continue
            for node in root.iter(_W + f"{kind}Reference"):
                if node.get(_W + "id") == number:
                    return part, node
        return None

    def _place(self: "Document", at: "str | TextRange"):
        """``(part, entry, offset)`` a note's reference goes at: a position, or a range's end."""
        from .ranges import TextRange

        where = at if isinstance(at, TextRange) else self.range(at)
        if where.view != "current":
            raise EditError("edits take the current view")
        part, entries = where._entries()
        entry = entries[-1]
        offset = where.end
        if part != self.package.document_part():
            raise EditError("a note's reference goes in the body (Word has no note in a header, a footer, "
                            "a note or a comment)")
        items = _text.atoms(entry.element)
        if not 0 <= offset <= len(items):
            raise EditError(f"{where.id} is outside its paragraph's text")
        if 0 < offset < len(items) and items[offset].field is not None and items[offset - 1].field is items[offset].field:
            raise EditError(f"{where.id} is inside a field's result")
        return part, entry, offset

    def insert_footnote(self: "Document", at: "str | TextRange", text: str = "") -> "EditResult":
        """A footnote whose reference stands at ``at`` (a position, or a range's end) and
        whose text is ``text`` (a line per paragraph).  The result's ``id`` is ``fn:<id>``;
        ``created`` has it and the note's paragraph ids."""
        return self._insert_note("footnote", at, text)

    def insert_endnote(self: "Document", at: "str | TextRange", text: str = "") -> "EditResult":
        """An endnote whose reference stands at ``at`` (a position, or a range's end), as
        :meth:`insert_footnote` writes a footnote: ``doc.insert_endnote(doc.anchor("2027"), "Source.")``."""
        return self._insert_note("endnote", at, text)

    def _insert_note(self: "Document", kind: str, at, text: str) -> "EditResult":
        from .document import EditResult
        from . import formatting as _formatting

        _text.check_text(text.replace("\n", ""))
        part, entry, offset = self._place(at)
        text_style, reference_style = _STYLES[kind]
        self.styles.resolve(text_style, "paragraph")
        self.styles.resolve(reference_style, "character")
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [entry])
            entry = self._index(part).entry_for(entry.element) or entry
            text_id = self.styles._ensure(text_style, "paragraph")
            reference_id = self.styles._ensure(reference_style, "character")
            notes = self._notes_part(kind, create=True)
            root = self.package.tree(notes)
            number = max([int(n.get(_W + "id")) for n in root.findall(_W + kind)
                          if (n.get(_W + "id") or "").lstrip("-").isdigit()] + [0]) + 1
            note = make(f"w:{kind}", **{"w:id": str(number)})
            paragraphs = _note_paragraphs(kind, text, text_id, reference_id)
            for paragraph in paragraphs:
                note.append(paragraph)
            root.append(note)
            used = self._used()
            for paragraph in paragraphs:
                para_id, text_id_ = _ids.generate(notes, used, 2)
                _ids.stamp(paragraph, para_id, text_id_)
            _ids.ensure_w14(root)
            reference = make(f"w:{kind}Reference", **{"w:id": str(number)})
            run = _inline.new_run(entry.element, offset, [reference])
            properties = _formatting.properties_of(run, "w:rPr", create=True)
            _formatting._set_child(properties, "w:rStyle", {"w:val": reference_id})
            if tracking is not None:
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                self._track_inserted(part, [run])
                stamp = Stamp(self, tracking)
                for k, paragraph in enumerate(paragraphs):
                    _track.insert_paragraph_content(paragraph, stamp, notes)
                    if k + 1 < len(paragraphs):
                        _track.mark(paragraph, "ins", stamp, notes)
                stamp.finish()
            self._renew_text_id(part, entry)
            self.package.mark_dirty(notes)
            self.package.mark_dirty(part)
            story = self._story_of(notes)
            ids = [f"{story}/p:{p.get(_ids.PARA_ID)}" for p in paragraphs]
        prefix = "fn" if kind == "footnote" else "en"
        return EditResult(f"{prefix}:{number}", created=[f"{prefix}:{number}", *ids], renamed=renames)

    def edit_note(self: "Document", identifier: str, text: str) -> "EditResult":
        """Replace a note's text (a line per paragraph), keeping its number mark and the
        first paragraph's formatting; tracked like any text edit."""
        from .document import EditResult

        _text.check_text(text.replace("\n", ""))
        part, note = self._note_element(identifier)
        index = self._index(part)
        held = list(note.iter(W_P))  # held: stable id()s
        inside = {id(p) for p in held}
        entries = [e for e in index.paragraphs if id(e.element) in inside]
        if not entries:
            raise EditError(f"{identifier} has no paragraph")
        lines = text.split("\n")
        first = entries[0]
        items = _text.atoms(first.element)
        start = 0
        while start < len(items) and items[start].char == _text.OBJECT:
            start += 1
        if start < len(items) and items[start].char == " ":
            start += 1
        renames: dict[str, str] = {}
        created: list[str] = []
        removed: list[str] = []
        with self._edit():
            for entry in entries[1:]:
                result = self.delete_block(entry.id)
                removed += result.removed
                renames.update(result.renamed)
            first_id = renames.get(first.id, first.id)
            length = len(self.paragraph(first_id).text)
            result = self.range(f"{first_id}@{start}:{length}").replace(lines[0])
            renames.update(result.renamed)
            previous = renames.get(first_id, first_id)
            for line in lines[1:]:
                result = self.insert_paragraph(line, after=previous)
                created += result.created
                renames.update(result.renamed)
                previous = result.id
        return EditResult(identifier, created=created, renamed=renames, removed=removed)

    def delete_note(self: "Document", identifier: str) -> "EditResult":
        """Delete a note: its reference and the note.  Tracked, the reference is deleted as
        a revision (accepting it removes the note)."""
        from .document import EditResult

        notes_part, note = self._note_element(identifier)
        found = self._note_reference(identifier)
        tracking = self._active_tracking()
        if tracking is not None and found is not None:
            part, reference = found
            run = reference.getparent()
            with self._edit():
                renames = self._prepare(part, [])
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                stamp = Stamp(self, tracking)
                run = _isolate_reference(reference)
                released = _track.delete_runs([run], stamp, part)
                stamp.finish()
                self._release(part, released)
                self.package.mark_dirty(part)
                gone = not _attached(reference)
                if gone:
                    # One's own inserted reference, deleted, goes outright (E3): its note too.
                    notes_part, note = self._note_element(identifier)
                    remove(note)
                    self.package.mark_dirty(notes_part)
            if gone:
                return EditResult(None, renamed=renames, removed=[identifier])
            return EditResult(identifier, renamed=renames)
        removed = [identifier] + [e.id for e in self._index(notes_part).paragraphs
                                  if any(e.element is p for p in note.iter(W_P))]
        with self._edit():
            renames = self._prepare_many([])
            notes_part, note = self._note_element(identifier)
            found = self._note_reference(identifier)
            if found is not None:
                part, reference = found
                _remove_reference(reference)
                self.package.mark_dirty(part)
            remove(note)
            self.package.mark_dirty(notes_part)
        return EditResult(None, renamed={k: v for k, v in renames.items() if k not in removed}, removed=removed)

    def move_note(self: "Document", identifier: str, to: "str | TextRange") -> "EditResult":
        """Move a note's reference to ``to``; the note keeps its id.  Tracked, the reference
        is deleted where it was and a copy of the note inserted where it goes (a reference
        has one note), whose id is the result's."""
        from .document import EditResult

        found = self._note_reference(identifier)
        if found is None:
            raise EditError(f"{identifier} has no reference to move")
        kind = "footnote" if identifier.startswith("fn:") else "endnote"
        part, entry, offset = self._place(to)
        tracking = self._active_tracking()
        if tracking is not None:
            notes_part, note = self._note_element(identifier)
            with self._edit():
                renames = self._prepare(part, [entry])
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp

                entry = self._index(part).entry_for(entry.element) or entry
                old_part, reference = self._note_reference(identifier)
                old_run = reference.getparent()
                properties = copy.deepcopy(old_run.find(_W + "rPr"))
                root = self.package.tree(notes_part)
                number = max([int(n.get(_W + "id")) for n in root.findall(_W + kind)
                              if (n.get(_W + "id") or "").lstrip("-").isdigit()] + [0]) + 1
                twin = copy.deepcopy(note)
                twin.set(_W + "id", str(number))
                used = self._used()
                for paragraph in twin.iter(W_P):
                    para_id, text_id = _ids.generate(notes_part, used, 2)
                    _ids.stamp(paragraph, para_id, text_id)
                note.addnext(twin)
                run = _inline.new_run(entry.element, offset, [make(f"w:{kind}Reference", **{"w:id": str(number)})])
                if properties is not None:
                    old = run.find(_W + "rPr")
                    if old is not None:
                        remove(old)
                    run.insert(0, properties)
                stamp = Stamp(self, tracking)
                moved = _isolate_reference(reference)
                self._release(old_part, _track.delete_runs([moved], stamp, old_part))
                stamp.finish()
                if not _attached(reference):
                    remove(note)  # one's own inserted reference went outright: its note too
                self._track_inserted(part, [run])
                stamp = Stamp(self, tracking)
                paragraphs = list(twin.iter(W_P))
                for k, paragraph in enumerate(paragraphs):
                    _track.insert_paragraph_content(paragraph, stamp, notes_part)
                    if k + 1 < len(paragraphs):
                        _track.mark(paragraph, "ins", stamp, notes_part)
                stamp.finish()
                self._renew_text_id(part, entry)
                for touched in {part, old_part, notes_part}:
                    self.package.mark_dirty(touched)
            prefix = "fn" if kind == "footnote" else "en"
            return EditResult(f"{prefix}:{number}", created=[f"{prefix}:{number}"], renamed=renames)
        with self._edit():
            renames = self._prepare(part, [entry])
            entry = self._index(part).entry_for(entry.element) or entry
            old_part, reference = self._note_reference(identifier)
            old_run = reference.getparent()
            properties = copy.deepcopy(old_run.find(_W + "rPr"))
            run = _inline.new_run(entry.element, offset, [copy.deepcopy(reference)])
            if properties is not None:
                old = run.find(_W + "rPr")
                if old is not None:
                    remove(old)
                run.insert(0, properties)
            _remove_reference(reference)
            self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
            self.package.mark_dirty(old_part)
        return EditResult(identifier, renamed=renames)

    def set_note_settings(self: "Document", kind: str = "endnote", *, position: str | None = None) -> "EditResult":
        """The document's note settings: the endnotes' ``position`` (``sectEnd``: at each
        section's end; ``docEnd``, the default: at the document's).  A footnote's position,
        numbering and restart are a section's (:meth:`set_section`), where Word writes and
        reads them."""
        from .document import EditResult

        if kind != "endnote":
            raise EditError("footnotes are positioned and numbered per section: set_section(footnote_position=...)")
        if position not in POSITIONS["endnote"]:
            raise EditError(f"an endnote position is one of {', '.join(POSITIONS['endnote'])}")
        part = self.package.settings_part()
        root = self.package.tree(part) if part is not None else None
        node = root.find(_W + "endnotePr") if root is not None else None
        current = node.find(_W + "pos") if node is not None else None
        value = None if position == "docEnd" else position
        if (current.get(_W + "val") if current is not None else None) == value:
            return EditResult(None, changed=False)
        with self._edit():
            part = part or self._create_settings_part()
            root = self.package.tree(part)
            node = root.find(_W + "endnotePr")
            if node is None:
                node = append_in_order(root, make("w:endnotePr"))
            for old in node.findall(_W + "pos"):
                remove(old)
            if value is not None:
                append_in_order(node, make("w:pos", **{"w:val": value}))
            if not len(node):
                remove(node)
            self.package.mark_dirty(part)
        return EditResult(None)


def _note_paragraphs(kind: str, text: str, text_style: str | None, reference_style: str | None) -> list[Element]:
    """A note's paragraphs as Word writes them: each in the text style, the first opening
    with the number mark in the reference style and a space."""
    out = []
    for k, line in enumerate(text.split("\n")):
        paragraph = make("w:p")
        if text_style is not None:
            properties = make("w:pPr")
            properties.append(make("w:pStyle", **{"w:val": text_style}))
            paragraph.append(properties)
        if k == 0:
            mark = make("w:r")
            if reference_style is not None:
                run_properties = make("w:rPr")
                run_properties.append(make("w:rStyle", **{"w:val": reference_style}))
                mark.append(run_properties)
            mark.append(make(f"w:{kind}Ref"))
            paragraph.append(mark)
            paragraph.append(_text.make_run(" ", None))
        if line:
            paragraph.append(_text.make_run(line, None))
        out.append(paragraph)
    return out


def _isolate_reference(reference: Element) -> Element:
    """The run holding only ``reference`` (split out of a run that holds more)."""
    run = reference.getparent()
    children = [c for c in run if isinstance(c.tag, str) and c.tag != _W + "rPr"]
    if children == [reference]:
        return run
    properties = run.find(_W + "rPr")
    alone = make("w:r")
    if properties is not None:
        alone.append(copy.deepcopy(properties))
    before = [c for c in children[:children.index(reference)]]
    after = [c for c in children[children.index(reference) + 1:]]
    if after:
        tail = make("w:r")
        if properties is not None:
            tail.append(copy.deepcopy(properties))
        for child in after:
            tail.append(child)
        run.addnext(tail)
    alone.append(reference)
    run.addnext(alone)
    if not before:
        remove(run)
    return alone


def _attached(node: Element) -> bool:
    root = node
    while root.getparent() is not None:
        root = root.getparent()
    return root.tag == _W + "document"


def _remove_reference(reference: Element) -> None:
    """Take a reference out with its run, when the run holds nothing else."""
    run = reference.getparent()
    remove(reference)
    if run is not None and not any(isinstance(c.tag, str) and c.tag != _W + "rPr" for c in run):
        parent = run.getparent()
        remove(run)
        while parent is not None and parent.tag in (_W + "ins", _W + "del") and not len(parent):
            grand = parent.getparent()
            remove(parent)
            parent = grand


__all__ = ["NoteOps", "new_notes_part"]
