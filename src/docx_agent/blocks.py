"""Edits of different blocks of one Word document at the same time (round 45, lane B): the Word
format's :class:`ooxml_edit.tools.blocks.BlockEdits` (``DocumentFormat.blocks``).

ooxml-edit runs a changing call this module admits on a **draft** -- an editable copy of the
document's last committed version (:func:`docx_agent.layout.read_copy`'s ``make.draft``) --
with no lock of the document held, then applies what it changed to the document under the
lock, briefly, when what it read is as it was (ooxml-edit's ``tools/blocks.py`` says the
protocol).  What runs on a draft: ``word_set_text``, ``word_insert_markdown`` (``end``,
``before:``, ``after:``; no pictures or footnotes) and ``word_edit_table``'s ``set_cells`` and
``sync`` (round 49); and
(round 48, R83) the edits of known blocks -- ``replace_text`` in a range, ``word_format`` of
blocks, runs, spans or text found within them, ``word_format_table``, ``word_delete`` of a
block, span or text found within one, ``word_links`` adding a link or a bookmark to a range --
and batches of any of them; text written by reference only from datasets (``{{d:...}}``,
``from: d:...``).  Whatever reads or changes the whole document keeps the document's lock for
its whole run, as before: text found anywhere, a story other than the body, a section, a
heading's range or a bookmark as an address (named, or a ref to one: found on the draft); a
style given by name that has to be made, a style or list definition changed, the settings.

**What a draft records** (:class:`DraftOps`, :class:`DraftPool`): every block an id named
(``Document._resolve``; ``_direct``: whether an id names anything), every id drawn from the
document's sequences (:func:`docx_agent.edit.ids.generate`), the facts about the whole document
the call used (:func:`observe`: the table style most tables use, the list definition the
document's lists use most, whether the body is one empty paragraph), and what a patch cannot
carry -- then the call runs as before on the document: ids stamped, a tracked change, an
annotation id (comments, bookmarks, revisions), revision ids repeated, ids kept from a copy,
drawing numbers, media, a layout, renames.

**The patch** (:meth:`Blocks.prepared`): the body's top-level blocks before and after, from
``word/document.xml``'s bytes (each block's own bytes, ids included), compared in order: each
run of blocks that differ is a *region*.  A *window* is what the change depends on (round 47,
R78), its bytes (gaps between blocks included) as they were: blocks changed under their own
ids are their own window; blocks put between two are placed by the one the call named --
before it (``before:X``: X) or after it (``after:P``: P) -- unless they hold a table, which
reads both neighbours (a separator beside a table, a paragraph at a container's end), or
neither neighbour was named (both, :data:`CONTEXT`); blocks removed or replaced, with their
neighbours; each block the call named, alone.  Every other part the call changed, before and
after.  The draws, and whether the ids drawn can be renamed (:func:`_renamable`).

**The check** (:meth:`Blocks.check`, under the document's lock): no part but the body changed
since the draft's version (ooxml-edit counts the parts each change touched) -- the provenance
item's entries merged instead, entry by entry (:func:`_merge_sources`); the body's head
(the root's namespaces) and tail (the last section's properties) as they were; each window in
the document now, found by its first block's id, byte for byte as it was; each id that named
nothing still naming nothing, each id that named something still naming something; every
fact the call used the same now; the renames, never given before; the document fully stamped
and its renames as they were.  Else a conflict: the call is prepared again.

**Apply and settle**: each window's bytes replaced by the draft's, ids renamed where the
document's sequence moved on (another call drew the same values first); the other parts
replaced; the id pool as the draws leave it; then the ledger read at the new version, the
spacing count, the checks' layout engine, and the reflow's "before" (the layout of the
document as it was just before) -- all as the call's own run on the document leaves them.

**Shared state** (round 45's inventory; ``tests/test_block_edits.py`` tests each):

=========================  =================================================================
item                       rule
=========================  =================================================================
body blocks                windows (what the change depends on) compared byte for byte;
                           regions spliced in
w14:paraId / textId        drawn again on the document (renamed): never reused, unique
id ledger, retired ids     read at the new version after the patch (``_note_ids``)
renames (aliases)          must be as they were; a call that renames: whole lock
stamping                   the document fully stamped (``_fast_ids``); a call that stamps:
                           whole lock
tracked changes            tracking on: whole lock; revision ids repeated: whole lock
annotation ids             a bookmark's drawn again when the patch applies (one above the
                           document's largest), renamed in its marks; comments', revisions':
                           whole lock
footnotes, endnotes        Markdown footnotes: not admitted; the parts: unchanged or conflict
comments part              unchanged since the draft's version, or conflict
numbering.xml              unchanged since, or conflict; a list restarted: the part replaced;
                           a definition changed by the call: whole lock (round 48)
styles.xml                 styles others only added: merged, the call's after theirs (one
                           there by its id: theirs); else unchanged since, or conflict; a
                           style changed, or made for a style given by name: whole lock
settings.xml, theme        unchanged since, or conflict; changed by the call: whole lock
bookmarks by name          the names a link or a new bookmark looked for (in any case), the
                           same; a bookmark as an address: whole lock
document.xml.rels          relationships others only added: merged (a target there already:
                           its id; a new one the next rIdN free, r:id renamed in the windows)
relationships, parts,      unchanged since, or conflict; added or replaced as the draft did,
[Content_Types]            in its order; ids drawn again renamed in an added XML part as in a
                           changed one (any other part added with ids renamed: conflict)
provenance item            merged entry by entry: the call's own entries as it left them, the
                           document's others as they are (an entry both changed: conflict);
                           added by the call while another added one: conflict
sectPr, section breaks     the body's tail as it was, every section break as it was (round
                           48: a table's width reads its section); a region holding a section
                           break changed: whole lock
spans, text ranges         every block between their ends (round 48): one window, in a row
fields and TOC             in the body's blocks (windows); field updates: whole lock
docProps                   unchanged since, or conflict (no admitted call writes them)
list numbering             the list style's numbering counts in Word; no admitted call reads
                           a list's place (its windows hold the neighbours it reads)
datasets, refs, keys       ooxml-edit checks them (the same datasets, not stale; the refs as
                           they were when used; a retry key not used meanwhile)
version, undo, owners      ooxml-edit: the patch is one undo step, the call's, its version
layouts, checks            per version, after the patch (the copy of the version it made)
=========================  =================================================================

**What a read covered** (round 47, :func:`touches`): a ranged ``check`` or ``render`` says the
blocks it read (``docx_agent.scope.ReadScope``); ooxml-edit warns it that the document
"changed while this call read it" only when a change made meanwhile touched those blocks --
changed or removed one, or put blocks between two of them -- or a part but the body other
than the registrations, the document's properties, the provenance item and lists only added
to the numbering.  Both depend on the layout, so for them any change at or before the range's
last block, or to the last section's properties, touches it too (blocks moved to other
pages); a change after the range does not.
"""

from __future__ import annotations

import contextlib
import difflib
import json
import re
import threading
import weakref
from collections import Counter
from contextlib import ExitStack
from dataclasses import dataclass, field
from typing import Any, Iterator, Mapping

from ooxml_edit.tools.blocks import BlockConflict, renamed, token_renamer
from lxml import etree
from ooxml_edit.xml import parse_xml, serialize

from .edit import ids as _ids

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_BODY = _W + "body"

#: Blocks of context either side of a place where a call added or removed blocks: the
#: neighbours it read there, which must be as they were -- and still each other's neighbours
#: -- for the change to apply (round 47, R78: what an edit depends on).  An insert reads the
#: blocks it goes between and nothing further (a table beside a table takes a separator; a
#: container's end, a paragraph); a block changed under its own id, and a block a call
#: named, is its own window -- ``word_set_text`` with lines also reads the paragraphs after
#: its paragraph, one per further line (:func:`_window_after`).
CONTEXT = 1

#: The calls a draft takes (:func:`admits`).
TOOLS = ("word_set_text", "word_insert_markdown", "word_edit_table",
         # Round 48 (R83): edits of known blocks.
         "replace_text", "word_format", "word_format_table", "word_delete", "word_links")

#: Drafts kept per document to reuse (each a parsed copy of it).
KEEP_DRAFTS = 2

_REFERENCE = re.compile(r"\{\{\s*([^}]*?)\s*\}\}")
_HEX_TOKEN = re.compile(rb"(?<![0-9A-Fa-f])([0-9A-Fa-f]{8})(?![0-9A-Fa-f])")
#: Where an id drawn may stand in what a call writes and says: a paraId or textId, a block
#: id (``p:``, ``t:``, ``tr:``), a provenance entry's key.
_ID_CONTEXTS = (b'w14:paraId="', b'w14:textId="', b"p:", b"t:", b"tr:")


# -- what a draft takes --------------------------------------------------------------------------


def admits(tool: str, arguments: Mapping[str, Any]) -> bool:
    """Whether a validated changing call may run on a draft: one of :data:`TOOLS`, naming
    blocks by id (or ``$ref``), text by reference only from datasets."""
    if tool == "word_set_text":
        items = arguments.get("items") or []
        targets = [arguments.get("target")] + [item.get("target") for item in items if isinstance(item, Mapping)]
        texts = [arguments.get("text")] + [item.get("text") for item in items if isinstance(item, Mapping)]
        return (all(_plain(target) for target in targets if target is not None)
                and all(_from_datasets(text) for text in texts))
    if tool == "word_insert_markdown":
        at = arguments.get("at")
        markdown = arguments.get("markdown")
        if not isinstance(at, str) or not _place(at):
            return False
        if isinstance(markdown, str) and ("![" in markdown or "[^" in markdown):
            return False
        likes = arguments.get("like") or []
        return (_from_datasets(markdown) and _source(arguments.get("from"))
                and all(_plain(sample) for sample in ([likes] if isinstance(likes, str) else likes) if sample))
    if tool == "word_edit_table":
        cells = arguments.get("cells") or []
        # Round 49 (R90): a sync is confined to its table (rows added and deleted in it).
        return (arguments.get("action") in ("set_cells", "sync") and _plain(arguments.get("table"))
                and _source(arguments.get("from"))
                and all(isinstance(cell, Mapping) and _source(cell.get("from")) and _from_datasets(cell.get("text"))
                        for cell in cells))
    # Round 48 (R83): edits whose effect is confined to the blocks they name.  Text found
    # anywhere (no range, no within), a story other than the body, a section, a heading's
    # range or a bookmark named as an address read the whole document: they keep the lock.
    if tool == "replace_text":
        return (arguments.get("stories") in (None, "body") and not arguments.get("slides")
                and _blocky(arguments.get("range")) and _from_datasets(arguments.get("replace")))
    if tool == "word_format":
        targets = _listed(arguments.get("targets"))
        finds = _listed(arguments.get("find"))
        return (bool(targets) != bool(finds)
                and all(_blocky(target) for target in targets)
                and all(isinstance(item, Mapping) and _blocky(item.get("within")) for item in finds)
                and (arguments.get("like") is None or _plain(arguments.get("like"))))
    if tool == "word_format_table":
        return _plain(arguments.get("table"))
    if tool == "word_delete":
        target, find = arguments.get("target"), arguments.get("find")
        return (_blocky(target) and find is None) or (target is None and find is not None
                                                      and _blocky(arguments.get("within")))
    if tool == "word_links":
        return (arguments.get("action") in ("add_hyperlink", "add_bookmark") and arguments.get("find") is None
                and _blocky(arguments.get("target")))
    return False


#: One end of a block address a draft takes (round 48): a paragraph, table, row or control
#: id, a cell or row of a table, a place or span of text in a paragraph, a picture or note --
#: in the body or a story (``header1/p:...``); a ref (what it names is resolved on the draft:
#: a heading's range, a section or a bookmark found there keeps the lock).  Not a section
#: (``s:``), a heading's range (``h:``) or a bookmark named as such.
_BLOCK_END = re.compile(r"(?:[A-Za-z]+\d*/)?(?:p|t|tr|cc|d|fn|en):[^\s.$]+|\$[\w-]+")


def _blocky(address: Any) -> bool:
    """A block address (:data:`_BLOCK_END`), or a span of two (``p:A..t:B``, open at
    neither end): what it covers is the blocks between them, which the draft records."""
    if not isinstance(address, str) or not address or "h:" in address:
        return False
    ends = address.split("..")
    return len(ends) in (1, 2) and all(_BLOCK_END.fullmatch(end.strip()) for end in ends)


def _listed(value: Any) -> list:
    if value is None:
        return []
    return list(value) if isinstance(value, (list, tuple)) else [value]


def alone(tool: str, arguments: Mapping[str, Any]) -> bool:
    """Whether an admitted call is worth a draft with no other edit in flight: a fill from
    a dataset (a ``from: "d:..."`` of ``word_edit_table`` or ``word_insert_markdown``, or of
    a cell), whose rows may be many, so an edit that comes while it runs need not wait."""
    if tool not in ("word_edit_table", "word_insert_markdown"):
        return False
    cells = arguments.get("cells") or [] if tool == "word_edit_table" else []
    return any(isinstance(value, str) and value.startswith("d:")
               for value in [arguments.get("from")] + [cell.get("from") for cell in cells
                                                        if isinstance(cell, Mapping)])


def _plain(address: Any) -> bool:
    """An id or a ref, not a span, a heading's range or a place in text."""
    return isinstance(address, str) and bool(address) and ".." not in address and "h:" not in address \
        and "@" not in address


def _place(at: str) -> bool:
    if at in ("end", "end:body"):
        return True
    kind, _, rest = at.partition(":")
    if kind in ("before", "after"):
        return _plain(rest)
    return kind not in ("replace", "end") and _plain(at)


def _source(value: Any) -> bool:
    """No source, or a dataset (``d:``): rows of a table of a document are read now."""
    return value is None or value == "" or (isinstance(value, str) and value.startswith("d:"))


def _from_datasets(text: Any) -> bool:
    """Text whose references (``{{...}}``) all name datasets."""
    if not isinstance(text, str) or "{{" not in text:
        return True
    return all(found.startswith("d:") for found in _REFERENCE.findall(text))


# -- what a draft records ------------------------------------------------------------------------


class _Recording:
    """What a call did on a draft that its patch, or its check, needs."""

    def __init__(self, body: str) -> None:
        self.body = body
        #: What a patch cannot carry, in words: the call runs on the document instead.
        self.cannot: list[str] = []
        #: ``(seed, values)``: every draw from the document's id sequences, in order.
        self.draws: list[tuple[str, list[str]]] = []
        #: The body's top-level blocks an id named, by id.
        self.blocks: set[str] = set()
        #: Ids named that named nothing; ids looked up (``_direct``) and whether they did.
        self.missing: set[str] = set()
        self.exists: dict[str, bool] = {}
        #: Facts of the whole document the call used: ``(name, args) -> value``.
        self.facts: dict[tuple, Any] = {}
        #: Annotation ids drawn (round 48: a bookmark's), in order.
        self.annotations: list[int] = []

    def read(self, document: Any, part: str, element: Any) -> None:
        """A block an id named, ``element`` in ``part``: its top-level block, if the body's."""
        if part != self.body:
            return
        node = element
        while node is not None and node.getparent() is not None and node.getparent().tag != W_BODY:
            node = node.getparent()
        if node is None or node.getparent() is None:
            return
        try:
            self.blocks.add(document._block_id(part, node))
        except Exception:  # noqa: BLE001 -- a block without an id: no window can find it
            self.cannot.append("a block without an id")


class DraftPool(_ids.IdPool):
    """A draft's id pool (:attr:`docx_agent.edit.ids.Ledger.pool`): records every draw."""

    __slots__ = ("recording",)

    def __init__(self, values: Any = ()) -> None:
        super().__init__(values)
        self.recording: _Recording | None = None

    def drew(self, seed: str, values: list[str]) -> None:
        if self.recording is not None:
            self.recording.draws.append((seed, list(values)))


def observe(document: Any, name: str, value: Any, *args: Any) -> Any:
    """``value``, the fact ``name`` (with ``args``) of the whole document a call used: on a
    draft, recorded, to be found the same when its change is applied (:func:`_fact`)."""
    recording = document.__dict__.get("_recording") if hasattr(document, "__dict__") else None
    if recording is not None:
        recording.facts.setdefault((name, args), value)
    return value


def _fact(document: Any, name: str, args: tuple) -> Any:
    from .markdown import write as _write

    if name == "table_style_in_use":
        return _write.table_style_in_use(document)
    if name == "list_definition":
        return _write._document_definition(document, *args)
    if name == "lone_empty_paragraph":
        return _write._lone_empty_paragraph(document)
    if name == "bookmarks":
        return _bookmark_names(document)
    if name == "bookmark_named":
        return _named(document._bookmark_markers(), *args)
    raise KeyError(name)


def _named(markers: Mapping[str, Any], folded: str) -> tuple:
    return tuple(sorted(name for name in markers if name.casefold() == folded))


def _bookmark_names(document: Any) -> tuple:
    """The document's bookmark names as a call looking one up reads them (round 48)."""
    with document._unrecorded() if isinstance(document, DraftOps) else contextlib.nullcontext():
        return tuple(sorted(document._bookmark_markers()))


class DraftOps:
    """What a draft's class adds to :class:`docx_agent.Document` (``layout._draft_class``):
    the recording.  Off a call (no recording), the document's own behaviour."""

    def _draft_cannot(self, why: str) -> None:
        recording = self.__dict__.get("_recording")
        if recording is not None:
            recording.cannot.append(why)

    def _resolve(self, identifier: str) -> Any:
        recording = self.__dict__.get("_recording")
        try:
            part, entry = super()._resolve(identifier)  # type: ignore[misc]
        except KeyError:
            if recording is not None:
                recording.missing.add(identifier)
            raise
        if recording is not None:
            recording.read(self, part, entry.element)
        return part, entry

    def _resolve_block(self, identifier: str) -> Any:
        found = super()._resolve_block(identifier)  # type: ignore[misc]
        recording = self.__dict__.get("_recording")
        if recording is not None and isinstance(identifier, str) and identifier.startswith(("cc:", "cc@")):
            recording.read(self, found[0], found[1].element)
        return found

    def _span_elements(self, first: Any, last: Any, movable: Any) -> Any:
        # Round 48: a span's blocks, every one of them -- one put between them meanwhile is
        # one the call would have changed too (its window: all of them, in a row).
        elements, part = super()._span_elements(first, last, movable)  # type: ignore[misc]
        recording = self.__dict__.get("_recording")
        if recording is not None:
            for element in elements:
                recording.read(self, part, element)
        return elements, part

    def _scope_paragraphs(self, address: str) -> Any:
        # Round 48: the paragraphs an address holds (a span, a range, a cell...): their blocks.
        part, entries = super()._scope_paragraphs(address)  # type: ignore[misc]
        recording = self.__dict__.get("_recording")
        if recording is not None:
            for entry in entries:
                recording.read(self, part, entry.element)
        return part, entries

    def range(self, identifier: str, *, view: str = "current") -> Any:
        # Round 48: a text range's paragraphs, from its first to its last.
        found = super().range(identifier, view=view)  # type: ignore[misc]
        recording = self.__dict__.get("_recording")
        if recording is not None:
            try:
                part, entries = found._entries()
            except (KeyError, ValueError):
                return found
            for entry in entries:
                recording.read(self, part, entry.element)
        return found

    def _expand_heading(self, address: Any) -> Any:
        # A heading's range ends where the outline says: what a ref may name too.
        found = super()._expand_heading(address)  # type: ignore[misc]
        if found != address:
            self._draft_cannot("a heading's range")
        return found

    def _section_paragraphs(self, identifier: str) -> Any:
        self._draft_cannot("a section's paragraphs")
        return super()._section_paragraphs(identifier)  # type: ignore[misc]

    def bookmark(self, name: str) -> Any:
        # A bookmark named as an address (or a ref to one): where it is, the whole
        # document's to say.  (A link to one only reads the names: _bookmark_markers.)
        found = super().bookmark(name)  # type: ignore[misc]
        if self.__dict__.get("_linking") is None:
            self._draft_cannot("a bookmark's range")
        return found

    def add_hyperlink(self, where: Any, url: Any = None, **options: Any) -> Any:
        # A link to a bookmark reads whether one of that name is there.
        anchor = options.get("anchor")
        self.__dict__["_linking"] = anchor
        try:
            return super().add_hyperlink(where, url, **options)  # type: ignore[misc]
        finally:
            self.__dict__.pop("_linking", None)

    def add_bookmark(self, where: Any, name: str) -> Any:
        # A bookmark added reads whether its name is taken (in any case).
        self.__dict__["_linking"] = name
        try:
            return super().add_bookmark(where, name)  # type: ignore[misc]
        finally:
            self.__dict__.pop("_linking", None)

    def _bookmark_markers(self) -> Any:
        # Round 48: the bookmarks by name, a fact of the whole document -- for a link to one
        # or a bookmark added, only the names that are that one's (in any case).
        found = super()._bookmark_markers()  # type: ignore[misc]
        recording = self.__dict__.get("_recording")
        if recording is not None:
            name = self.__dict__.get("_linking")
            if isinstance(name, str):
                recording.facts.setdefault(("bookmark_named", (name.casefold(),)),
                                           _named(found, name.casefold()))
            else:
                recording.facts.setdefault(("bookmarks", ()), tuple(sorted(found)))
        return found

    def _direct(self, identifier: str) -> Any:
        found = super()._direct(identifier)  # type: ignore[misc]
        recording = self.__dict__.get("_recording")
        if recording is not None:
            recording.exists.setdefault(identifier, found is not None)
        return found

    def _follow_text_facts(self) -> Any:
        # Every edit's bookkeeping of values written by reference: it reads each paragraph
        # an entry of the provenance item names, wherever it is.  Not the call's reads: a
        # change to such a paragraph's text, or its going, changes the item too (the change
        # that made it ran this), and the item must be as it was for the patch to apply.
        with self._unrecorded():
            return super()._follow_text_facts()  # type: ignore[misc]

    def _prune_provenance(self) -> Any:
        # The same for the records whose table or paragraph is gone (the item again).
        with self._unrecorded():
            return super()._prune_provenance()  # type: ignore[misc]

    @contextlib.contextmanager
    def _unrecorded(self) -> Iterator[None]:
        recording = self.__dict__.pop("_recording", None)
        try:
            yield
        finally:
            if recording is not None:
                self.__dict__["_recording"] = recording

    def _stamp_document(self) -> Any:
        renames = super()._stamp_document()  # type: ignore[misc]
        if renames:
            self._draft_cannot("ids stamped")
        return renames

    def _prepare(self, part: str, entries: Any, position: Any = None, *, including: bool = False) -> Any:
        if self.stamping != "document":  # type: ignore[attr-defined]
            self._draft_cannot("stamping by paragraph")
        return super()._prepare(part, entries, position, including=including)  # type: ignore[misc]

    def _prepare_many(self, targets: Any) -> Any:
        if self.stamping != "document":  # type: ignore[attr-defined]
            self._draft_cannot("stamping by paragraph")
        return super()._prepare_many(targets)  # type: ignore[misc]

    def _unique_revision_ids(self, before: Any) -> Any:
        if before is not None and self._revision_id_counts() != before:  # type: ignore[attr-defined]
            self._draft_cannot("revision ids")
        return super()._unique_revision_ids(before)  # type: ignore[misc]

    def _next_annotation_id(self) -> Any:
        # Round 48: drawn on the draft, drawn again on the document when the patch applies
        # (a bookmark's: renamed in what the call wrote); any other use keeps the lock.
        value = super()._next_annotation_id()  # type: ignore[misc]
        recording = self.__dict__.get("_recording")
        if recording is not None:
            recording.annotations.append(value)
        return value

    def _remember_ids(self, values: Any) -> Any:
        self._draft_cannot("ids kept")
        return super()._remember_ids(values)  # type: ignore[misc]

    def _next_doc_pr(self) -> Any:
        self._draft_cannot("drawing numbers")
        return super()._next_doc_pr()  # type: ignore[misc]

    def _media(self, part: str, data: bytes) -> Any:
        self._draft_cannot("media")
        return super()._media(part, data)  # type: ignore[misc]

    def _active_tracking(self) -> Any:
        found = super()._active_tracking()  # type: ignore[misc]
        if found is not None:
            self._draft_cannot("tracked")
        return found

    def layout(self, **options: Any) -> Any:
        self._draft_cannot("laid out")
        return super().layout(**options)  # type: ignore[misc]

    def _rename(self, renames: Any) -> Any:
        if renames:
            self._draft_cannot("ids renamed")
        return super()._rename(renames)  # type: ignore[misc]

    def _note_replaced(self, old: Any, new: Any) -> Any:
        self._draft_cannot("a block replaced")
        return super()._note_replaced(old, new)  # type: ignore[misc]


# -- the body in bytes ---------------------------------------------------------------------------


@dataclass
class _Body:
    """``word/document.xml``'s bytes, read for blocks: where the body's content begins and
    ends, each top-level block's span and id."""

    data: bytes
    start: int
    end: int
    spans: list[tuple[int, int]]
    ids: list[str]
    #: Every paragraph's and row's paraId in the part (upper case), each once: the part is
    #: fully stamped (``touched._stamped``), counted so another state's is read from the
    #: blocks that differ alone (:func:`_reread_body`).
    stamps: "Counter[bytes]" = field(default_factory=Counter, repr=False)
    _where: dict[str, int] | None = None

    @property
    def count(self) -> int:
        return len(self.spans)

    def block(self, k: int) -> bytes:
        begin, end = self.spans[k]
        return self.data[begin:end]

    def blocks(self) -> list[bytes]:
        return [self.data[begin:end] for begin, end in self.spans]

    def at(self, k: int) -> int:
        """Where a window from block ``k`` begins: after the block before it."""
        return self.spans[k - 1][1] if k > 0 else self.start

    def to(self, k: int) -> int:
        """Where a window up to block ``k`` (not included) ends: where it begins."""
        return self.spans[k][0] if k < len(self.spans) else self.end

    def head(self) -> bytes:
        return self.data[:self.start]

    def tail(self) -> bytes:
        return self.data[self.spans[-1][1] if self.spans else self.start:]

    def where(self, identifier: str) -> int | None:
        if self._where is None:
            self._where = {identifier: k for k, identifier in enumerate(self.ids)}
        return self._where.get(identifier)


_BODY_OPEN = re.compile(rb"<w:body(?=[\s>])")


#: Bodies read lately, by their bytes (the same state is read by the draft, the check and a
#: batch's ops): :func:`read_body`'s answers, the latest last.
_BODIES: "dict[bytes, _Body | None]" = {}
_BODIES_LOCK = threading.Lock()
_BODIES_KEPT = 8


def read_body(data: bytes | None) -> _Body | None:
    """The body's blocks from the part's bytes (``touched``'s scan), each named by its id;
    ``None`` when they cannot be read so -- the part does not write ``w:``, its tags do
    not nest, or a paragraph, row or content control lacks an id of its own (the document is
    not fully stamped).  The same bytes are read once (:data:`_BODIES`)."""
    if data is None:
        return None
    with _BODIES_LOCK:
        if data in _BODIES:
            found = _BODIES.pop(data)
            _BODIES[data] = found
            return found
        base = next((body for body in reversed(_BODIES.values()) if body is not None), None)
    found = (base is not None and _reread_body(data, base)) or _read_body(data)
    with _BODIES_LOCK:
        _BODIES[data] = found
        while len(_BODIES) > _BODIES_KEPT:
            _BODIES.pop(next(iter(_BODIES)))
    return found


def _read_body(data: bytes) -> _Body | None:
    from .tools import touched as _touched

    root_end = data.find(b">", data.find(b"<w:document"))
    if root_end < 0 or _touched._W_DECLARED not in data[:root_end]:
        return None
    opening = _BODY_OPEN.search(data, root_end)
    end = data.rfind(b"</w:body>")
    if opening is None or end < 0:
        return None
    start = data.index(b">", opening.end()) + 1
    spans = _touched._blocks(data, start, end)
    if spans is None or _touched._W14_DECLARED not in data[:root_end]:
        return None
    stamps = _stamps(data)
    if stamps is None or any(count != 1 for count in stamps.values()):
        return None
    ids = _touched._span_ids(data, spans)
    if ids is None or len(set(ids)) != len(ids):
        return None
    return _Body(data, start, end, spans, ids, stamps)


def _stamps(data: bytes) -> "Counter[bytes] | None":
    """The paraIds in ``data`` (upper case) when every paragraph and row there has one
    (``touched._stamped``'s reading), else ``None``."""
    from .tools import touched as _touched

    live = _touched._FALLBACK.sub(b"", data) if b"<mc:Fallback" in data else data
    stamped = _touched._ALL_PARA_IDS.findall(live)
    if len(stamped) != len(_touched._PARA_OR_ROW.findall(live)):
        return None
    return Counter(map(bytes.upper, stamped))


def _reread_body(data: bytes, base: _Body) -> _Body | None:
    """:func:`_read_body` from ``base``, another state of a body: the blocks before the
    first byte that differs and after the last are ``base``'s, only those between are read
    again (as ``touched`` reads a change).  ``None`` when the bytes differ outside the
    body's blocks, or what is between does not read so."""
    from .tools import touched as _touched

    old = base.data
    prefix = _touched._common_prefix(old, data)
    suffix = _touched._common_suffix(old, data, min(len(old), len(data)) - prefix)
    if prefix < base.start or len(old) - suffix > base.end:
        return None
    delta = len(data) - len(old)
    spans = base.spans
    before = 0
    while before < len(spans) and spans[before][1] <= prefix:
        before += 1
    after = 0
    while after < len(spans) - before and spans[len(spans) - 1 - after][0] >= len(old) - suffix:
        after += 1
    begin = spans[before - 1][1] if before else base.start
    stop = spans[len(spans) - after][0] + delta if after else base.end + delta
    middle = _touched._blocks(data, begin, stop)
    if middle is None:
        return None
    # Stamped as the whole part is read (every paraId once): the base's, less those of its
    # blocks between, plus those of the blocks between now.
    added, removed = _stamps(data[begin:stop]), _stamps(old[begin:stop - delta])
    if added is None or removed is None or not base.stamps:
        return None
    stamps = base.stamps.copy()
    stamps.subtract(removed)
    stamps.update(added)
    if any(stamps[value] != 1 for value in added):
        return None
    stamps = +stamps
    found = _touched._span_ids(data, middle)
    if found is None:
        return None
    ids = base.ids[:before] + found + base.ids[len(spans) - after:]
    if len(set(ids)) != len(ids):
        return None
    shifted = [(start + delta, end + delta) for start, end in spans[len(spans) - after:]]
    return _Body(data, base.start, base.end + delta, spans[:before] + middle + shifted, ids, stamps)


def _regions(before: _Body, after: _Body) -> list[tuple[int, int, int, int]]:
    """Each run of blocks that differ: ``(i1, i2, j1, j2)``, blocks ``i1..i2`` before became
    ``j1..j2`` after."""
    old, new = before.blocks(), after.blocks()
    n, m = len(old), len(new)
    head = 0
    while head < n and head < m and old[head] == new[head]:
        head += 1
    tail = 0
    while tail < n - head and tail < m - head and old[n - 1 - tail] == new[m - 1 - tail]:
        tail += 1
    if head == n and head == m:
        return []
    middle_old, middle_new = old[head:n - tail], new[head:m - tail]
    if not middle_old or not middle_new:
        return [(head, n - tail, head, m - tail)]
    matcher = difflib.SequenceMatcher(None, middle_old, middle_new, autojunk=False)
    return [(head + i1, head + i2, head + j1, head + j2)
            for tag, i1, i2, j1, j2 in matcher.get_opcodes() if tag != "equal"]


@dataclass
class _Window:
    """Blocks ``first..last`` (not included) of the body before the call (``first_after
    ..last_after`` after it), and their bytes then and after."""

    first: int
    last: int
    first_after: int
    last_after: int
    before: bytes
    after: bytes
    ids: list[str]
    #: The ids of the blocks after the call (``first_after..last_after``), and each block's
    #: bytes before and after.
    after_ids: list[str] = field(default_factory=list)
    blocks_before: list[bytes] = field(default_factory=list)
    blocks_after: list[bytes] = field(default_factory=list)


def _windows(before: _Body, after: _Body, regions: list[tuple[int, int, int, int]],
             named: list[int], wide: int) -> list[_Window]:
    """The windows: each region -- with :data:`CONTEXT` blocks either side where it adds or
    removes blocks, none where it changes blocks under their own ids; ``wide`` more after it
    for a call that reads further -- and each block named (``named``), merged where they
    meet."""
    n = before.count
    spans: list[tuple[int, int, list[tuple[int, int, int, int]]]] = []
    anchors = {before.ids[k] for k in named}
    for region in regions:
        i1, i2, j1, j2 = region
        if i2 > i1 and before.ids[i1:i2] == after.ids[j1:j2]:
            # Blocks changed under their own ids: those blocks (and what the call reads after).
            spans.append((i1, min(n, i2 + wide), [region]))
            continue
        ahead, behind = CONTEXT, CONTEXT
        if i1 == i2 and not any(b"<w:tbl" in after.block(j) for j in range(j1, j2)):
            # Blocks put between two: placed by the one the call named -- before it, or after
            # it -- and reading nothing else there (a table reads both: beside a table it
            # takes a separator, at a container's end a paragraph after it).
            named_before = i1 > 0 and before.ids[i1 - 1] in anchors
            named_after = i1 < n and before.ids[i1] in anchors
            if named_after and not named_before:
                ahead = 0
            elif named_before and not named_after:
                behind = 0
        spans.append((max(0, i1 - ahead), min(n, i2 + max(behind, wide)), [region]))
    for k in named:
        spans.append((k, k + 1, []))
    spans.sort(key=lambda item: (item[0], item[1]))
    merged: list[list[Any]] = []
    for first, last, held in spans:
        if merged and first <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], last)
            merged[-1][2].extend(held)
        else:
            merged.append([first, last, list(held)])
    out: list[_Window] = []
    shift = 0
    for first, last, held in merged:
        grown = sum((j2 - j1) - (i2 - i1) for i1, i2, j1, j2 in held)
        first_after, last_after = first + shift, last + shift + grown
        out.append(_Window(first, last, first_after, last_after,
                           before.data[before.at(first):before.to(last)],
                           after.data[after.at(first_after):after.to(last_after)],
                           before.ids[first:last], after.ids[first_after:last_after],
                           [before.block(k) for k in range(first, last)],
                           [after.block(k) for k in range(first_after, last_after)]))
        shift += grown
    return out


# -- the patch -----------------------------------------------------------------------------------


@dataclass
class Patch:
    """What a call changed on its draft (:meth:`Blocks.prepared`)."""

    body: str
    head: bytes
    tail: bytes
    #: How many top-level blocks the body had before the call.
    count: int
    windows: list[_Window]
    #: Every other part changed: ``(path, before, after)``, ``None`` for absent; in the
    #: draft's order of its parts.
    parts: list[tuple[str, bytes | None, bytes | None]]
    draws: list[tuple[str, list[str]]]
    #: Whether the ids drawn can be renamed (:func:`_renamable`).
    renamable: bool
    missing: set[str]
    exists: dict[str, bool]
    facts: dict[tuple, Any]
    aliases: dict[str, str]
    spaced: int
    engine: Any
    after_body: bytes | None
    #: The version the draft was of; the draft's id pool before and after the call.
    version: int = -1
    pool_before: tuple[int, dict] = (0, {})
    pool_after: Any = None
    before_body: bytes | None = None
    #: The provenance item (``edit/provenance.py``) as the draft's version had it, by path;
    #: ``None``: it had none.  Its entries are merged with the document's (:func:`_merge_sources`).
    provenance: str | None = None
    #: The committed version's bytes (the parts the call did not change, as it read them).
    committed_data: bytes | None = None
    #: Annotation ids the call drew (bookmarks'), in order.
    annotations: list[int] = field(default_factory=list)


@dataclass
class Plan:
    """A patch checked against the document now (:meth:`Blocks.check`): its new body and
    parts, the renames and the id pool they leave."""

    patch: Patch
    data: bytes | None
    renames: dict[str, str]
    pool: Any
    parts: list[tuple[str, bytes | None, bytes | None]] = field(default_factory=list)
    #: The body's bytes before the patch, and the renames that are not the same name.
    before: bytes | None = None
    moved: dict[str, str] = field(default_factory=dict)
    placed: list = field(default_factory=list)
    #: Round 48: the parts merged with others' changes (path -> ``"sources"``, ``"rels"``,
    #: ``"styles"``), the call's relationships and bookmark ids as the document gives them.
    merged: dict[str, str] = field(default_factory=dict)
    relinked: dict[str, str] = field(default_factory=dict)
    remarked: dict[str, str] = field(default_factory=dict)

    def rewritten(self, window: "_Window") -> "_Window":
        """An op's window as the call's apply writes it (relationship and bookmark ids)."""
        if self.relinked:
            window = _relinked_window(window, self.relinked)
        if self.remarked:
            window = _remarked_window(window, self.remarked)
        return window


def _renamable(draws: list[tuple[str, list[str]]], arguments: Mapping[str, Any], outputs: list[bytes]) -> bool:
    """Whether the ids a call drew can be renamed in what it wrote and said by their text
    alone: none of them is in its arguments, and each place one is in its outputs is an id's
    (:data:`_ID_CONTEXTS`) -- text, a number format, a hash are never renamed."""
    drawn = {value.encode() for _seed, values in draws for value in values}
    if not drawn:
        return True
    given = json.dumps(arguments, ensure_ascii=False, default=str).encode("utf-8").upper()
    if any(token in drawn for token in _HEX_TOKEN.findall(given)):
        return False
    for data in outputs:
        for match in _HEX_TOKEN.finditer(data):
            token = match.group(1)
            if token in drawn and not data[max(0, match.start() - 12):match.start()].endswith(_ID_CONTEXTS):
                return False
    return True


def _renamed_bytes(data: bytes, renames: Mapping[str, str]) -> bytes:
    if not renames:
        return data
    rename = token_renamer(renames)
    return rename(data.decode("utf-8")).encode("utf-8")


# -- the drafts ----------------------------------------------------------------------------------


class Draft:
    """A call's draft: the copy, and what the call does on it (:class:`_Recording`)."""

    def __init__(self, document: Any, committed: Any) -> None:
        self.document = document
        self.committed = committed
        self.recording = _Recording(document.package.document_part())
        self.why: str | None = None
        document.__dict__["_recording"] = self.recording
        pool = document._ledger.pool
        if isinstance(pool, DraftPool):
            pool.recording = self.recording
        self.aliases = dict(document._aliases)
        self.replaced = dict(document._ledger.replaced)
        #: The id pool's size and cursors when the call began (:func:`_redraw`'s shortcut).
        self.pool = (len(pool), dict(pool.cursors)) if pool is not None else (0, {})
        self.spaced = getattr(document, "typography_spaced", 0)
        self._stack = ExitStack()
        self.change = self._stack.enter_context(document.history.watch())

    def stop(self) -> None:
        """The call has run: stop recording, close the watch."""
        self._stack.close()
        self.document.__dict__.pop("_recording", None)
        pool = self.document._ledger.pool
        if isinstance(pool, DraftPool):
            pool.recording = None


class Blocks:
    """The Word format's :class:`ooxml_edit.tools.blocks.BlockEdits`."""

    def __init__(self) -> None:
        self._kept: "weakref.WeakKeyDictionary[Any, list[tuple[Any, Any]]]" = weakref.WeakKeyDictionary()
        self._lock = threading.Lock()

    # -- ooxml-edit's protocol --------------------------------------------------------------------

    def admits(self, tool: str, arguments: Mapping[str, Any]) -> bool:
        return admits(tool, arguments)

    def alone(self, tool: str, arguments: Mapping[str, Any]) -> bool:
        return alone(tool, arguments)

    def draft(self, document: Any, committed: Any) -> Draft | None:
        make = getattr(getattr(committed, "make", None), "draft", None)
        if make is None:
            return None
        copy = None
        with self._lock:
            kept = self._kept.get(document) or []
            for k, (version, found) in enumerate(kept):
                if version is committed:
                    copy = found
                    del kept[k]
                    break
        if copy is None:
            copy = make(committed.data)
        else:
            copy.typography_spaced = getattr(document, "typography_spaced", 0)
        return Draft(copy, committed)

    def prepared(self, draft: Draft, result: Any, arguments: Mapping[str, Any]) -> Patch | None:
        draft.stop()
        recording, document = draft.recording, draft.document
        why = list(dict.fromkeys(recording.cannot))
        if not isinstance(document._ledger.pool, DraftPool):
            why.append("no id pool")
        if document._aliases != draft.aliases or document._ledger.replaced != draft.replaced:
            why.append("ids renamed")
        change = draft.change
        body = document.package.document_part()
        before_bytes, after_bytes = change.before(body), change.after(body)
        before, after = read_body(before_bytes), read_body(after_bytes)
        if before is None or after is None:
            why.append("a body whose blocks do not all have ids")
        if why:
            draft.why = "; ".join(why)
            return None
        assert before is not None and after is not None
        if before.head() != after.head() or before.tail() != after.tail():
            draft.why = "the body's namespaces or last section changed"
            return None
        regions = _regions(before, after)
        wide = _window_after(arguments)
        named = sorted(k for k in (before.where(identifier) for identifier in recording.blocks) if k is not None)
        windows = _windows(before, after, regions, named, wide)
        for window in windows:
            if b"<w:sectPr" in window.before or b"<w:sectPr" in window.after:
                if _section_breaks(window.before) != _section_breaks(window.after):
                    draft.why = "a section break changed"
                    return None
        parts: list[tuple[str, bytes | None, bytes | None]] = []
        order = {path: k for k, path in enumerate(_entry_names(document.package))}
        for path in sorted(change.parts - {body}, key=lambda p: (order.get(p, len(order)), p)):
            parts.append((path, change.before(path), change.after(path)))
        styled = _names_styles(arguments)
        whole = next((path for path, old, new in parts if _whole_document_part(path, old, new, styled)), None)
        if whole is not None:
            # A style made or changed, a list definition changed, the settings: what every
            # block reads (round 48: such a call keeps the lock).
            draft.why = f"{whole} changed"
            return None
        outputs = [window.after for window in windows] + [after for _p, _b, after in parts if after]
        outputs.append(json.dumps(_said(result), ensure_ascii=False, default=str).encode("utf-8"))
        if recording.annotations:
            # Ids drawn for anything but bookmarks (a comment's, a revision's) cannot be
            # drawn again by their text alone: the lock.
            drawn = {str(value).encode() for value in recording.annotations}
            written = [window.after for window in windows] + [after for _p, _b, after in parts if after]
            every = sum(1 for data in written for found in re.finditer(rb'\bw:id="(\d+)"', data)
                        if found.group(1) in drawn)
            marks = sum(1 for data in written for found in _BOOKMARK_ID.finditer(data) if found.group(2) in drawn)
            if every != marks:
                draft.why = "annotation ids"
                return None
        item = document._provenance_item()
        if item is not None and change.before(item) is None:
            item = None             # added by the call: the document's version had none
        if item is None:
            item = next((path for path, old, _new in parts if old is not None and _is_sources(old)), None)
        return Patch(body=body, head=before.head(), tail=before.tail(), count=before.count, windows=windows,
                     parts=parts,
                     draws=list(recording.draws), renamable=_renamable(recording.draws, arguments, outputs),
                     missing=set(recording.missing), exists=dict(recording.exists), facts=dict(recording.facts),
                     aliases=draft.aliases,
                     spaced=getattr(document, "typography_spaced", 0) - draft.spaced,
                     engine=document.__dict__.get("_checks_engine"), after_body=after_bytes,
                     version=draft.committed.version, pool_before=draft.pool,
                     pool_after=document._ledger.pool, before_body=before_bytes, provenance=item,
                     committed_data=getattr(draft.committed, "data", None),
                     annotations=list(recording.annotations))

    def check(self, document: Any, patch: Patch, changed: frozenset[str] | None) -> Plan:
        if changed is None:
            raise BlockConflict("the document changed outside the tools")
        others = sorted(changed - {patch.body})
        merge = None
        if patch.provenance is not None and patch.provenance in others \
                and document._provenance_item() == patch.provenance:
            # Values' sources, kept per table and paragraph: merged entry by entry with the
            # document's (each kept for the call's own blocks), not a conflict.
            others.remove(patch.provenance)
            merge = patch.provenance
        # Round 48: the document part's relationships (a hyperlink's) and the styles (a style
        # made on first use) others only added to: merged -- the call's additions after
        # theirs, as the call run now adds them -- or, where the call did not change them,
        # nothing the call read.
        merges: dict[str, Any] = {}
        drafted = {path: (before, after) for path, before, after in patch.parts}
        for path in list(others):
            kind = _additive_kind(path)
            if kind is None or not document.package.has_part(path):
                continue
            live_bytes = document.package.read(path)
            mine = drafted.get(path)
            base = mine[0] if mine is not None else _part_before(patch, path)
            if base is None or not _only_added(kind, base, live_bytes):
                continue
            others.remove(path)
            if mine is not None:
                merges[path] = kind
        if others:
            raise BlockConflict(f"{', '.join(others[:3])} changed")
        if document._aliases != patch.aliases:
            raise BlockConflict("ids were renamed")
        data = _body_now(document, patch.body)
        live = read_body(data)
        if live is None:
            raise BlockConflict("the document is not fully stamped")
        if live.head() != patch.head or live.tail() != patch.tail:
            raise BlockConflict("the body's namespaces or last section changed")
        if patch.before_body is not None and (b"<w:sectPr" in data or b"<w:sectPr" in patch.before_body) \
                and _section_breaks(data) != _section_breaks(patch.before_body):
            # A section's page and margins: what a table's width, a layout of any block reads.
            raise BlockConflict("a section changed")
        if (b"<wp:inline" in data or b"<wp:anchor" in data) and document._stamping_due():
            raise BlockConflict("a drawing needs ids")
        placed = _placed(live, data, patch.windows, patch.count)
        for identifier in patch.missing:
            try:
                document._resolve(identifier)
            except KeyError:
                continue
            raise BlockConflict(f"{identifier} names a block now")
        seen = {identifier for _begin, _end, window in placed for identifier in window.ids}
        made = {value for _seed, values in patch.draws for value in values}
        for identifier, existed in patch.exists.items():
            if identifier in seen:
                continue            # in a window found the same
            if existed and identifier.rpartition(":")[2].upper() in made:
                continue            # one the call made itself (its patch makes it)
            if existed and live.where(identifier) is not None:
                continue            # a block of the body, there now
            if (document._direct(identifier) is not None) != existed:
                raise BlockConflict(f"{identifier} {'is gone' if existed else 'names a block now'}")
        for (name, args), value in patch.facts.items():
            if _fact(document, name, args) != value:
                raise BlockConflict(f"the document's {name.replace('_', ' ')} changed")
        renames, pool = _redraw(document, patch)
        moved = {old: new for old, new in renames.items() if old != new}
        # Ids drawn again are renamed in what the call wrote -- the body's windows and every
        # part it changed or added (the provenance item a first fill adds) -- by their text.
        if moved and (not patch.renamable or any(after is not None and not path.endswith((".xml", ".rels"))
                                                 for path, _b, after in patch.parts)):
            raise BlockConflict("ids drawn meanwhile, and these cannot be renamed")
        parts = [(path, before, _renamed_bytes(after, moved) if after is not None else None)
                 for path, before, after in patch.parts]
        if merge is not None:
            parts = [(path, before, after) if path != merge
                     else (path, *_merge_sources(document.package.read(path), before, after))
                     for path, before, after in parts]
        relinked: dict[str, str] = {}
        if merges:
            merged_parts = []
            for path, before, after in parts:
                kind = merges.get(path)
                if kind is not None:
                    live_bytes = document.package.read(path)
                    after, ids = _merge_added(kind, live_bytes, before, after)
                    before = live_bytes
                    relinked.update(ids)
                merged_parts.append((path, before, after))
            parts = merged_parts
            if relinked:
                # The call's new relationships under the ids the document gives them now.
                placed = [(begin, end, _relinked_window(window, relinked)) for begin, end, window in placed]
        ids: dict[str, str] = {}
        if patch.annotations:
            # Its bookmarks' ids, drawn again: one above the largest the document has now.
            offset = document._next_annotation_id() - patch.annotations[0]
            if offset:
                ids = {str(value): str(value + offset) for value in patch.annotations}
                placed = [(begin, end, _remarked_window(window, ids)) for begin, end, window in placed]
                parts = [(path, before, _remarked(after, ids) if after is not None and path.endswith(".xml") else after)
                         for path, before, after in parts]
        if moved and any(before is None for _p, before, _a in parts):
            parts = _named_again(patch.parts, parts, moved)
        merged = {**{path: kind for path, kind in merges.items()}, **({merge: "sources"} if merge else {})}
        return Plan(patch, _spliced(data, live, placed, moved), renames, pool, parts, before=data, moved=moved,
                    placed=placed, merged=merged, relinked=relinked, remarked=ids)

    def replay_changes(self, document: Any, plan: Plan, op_changes: list[tuple[int, str, Any]]
                       ) -> list[tuple[int, str, Any]] | None:
        """A batch's changes op by op (each ``(op, doc, change)`` on the draft) as the
        document's: each op's blocks spliced in turn into the body as it is now, for what
        each op changed (``Result.touched_by_op``) read from the document, as a batch run on
        it reads it; ``None`` when the body is as the draft's was and no id was renamed: the
        draft's are the document's.  :class:`BlockConflict` when they cannot be (never,
        where the patch applies: each op's blocks lie in its windows)."""
        body = plan.patch.body
        current = plan.before
        if current == plan.patch.before_body and not plan.moved and not plan.merged \
                and not plan.relinked and not plan.remarked:
            return None             # the body as the draft's: the draft's changes are the document's
        # Parts merged with others' changes: each op's additions merged in turn into the
        # document's (an op that adds what the document has already changed nothing there).
        states = {path: before for path, before, _after in plan.parts if path in plan.merged}
        out: list[tuple[int, str, Any]] = []
        for op, doc, change in sorted(op_changes, key=lambda item: item[0]):
            parts = set(change.parts)
            befores: dict[str, bytes | None] = {path: change.before(path) for path in parts - {body}}
            afters: dict[str, bytes | None] = {path: change.after(path) for path in parts - {body}}
            for path in list(parts - {body}):
                kind = plan.merged.get(path)
                if kind is None:
                    continue
                state = states[path]
                if kind == "sources":
                    made_part = _merge_sources(state, befores[path], afters[path])[1]
                else:
                    made_part = _merge_added(kind, state, befores[path], afters[path])[0]
                if made_part == state:
                    parts.discard(path)
                    continue
                befores[path], afters[path], states[path] = state, made_part, made_part
            if body in parts:
                old, new = read_body(change.before(body)), read_body(change.after(body))
                live = read_body(current)
                if old is None or new is None or live is None:
                    raise BlockConflict("an op's change cannot be read by blocks")
                windows = [plan.rewritten(window) for window in _windows(old, new, _regions(old, new), [], CONTEXT)]
                made = _spliced(current, live, _placed(live, current, windows, old.count), plan.moved)
                befores[body], afters[body] = current, made
                current = made
            out.append((op, doc, _Replayed(frozenset(parts), befores, afters)))
        if current != plan.data:
            raise BlockConflict("the ops' changes do not add up to the call's")
        return out

    def apply(self, document: Any, plan: Plan) -> Mapping[str, str]:
        package = document.package
        patch = plan.patch
        if plan.data is not None and plan.data != plan.before:
            _splice_tree(package, patch.body, plan)
        for path, before, after in plan.parts:
            if after is None:
                if package.has_part(path):      # (a part's relationships go with it)
                    package.remove_part(path)
            elif before is None or not package.has_part(path):
                package.replace_part(path, after)
            elif path.endswith((".xml", ".rels")):
                package.set_tree(path, parse_xml(after))
            else:
                package.replace_part(path, after)
        document._ledger.pool = plan.pool
        document._invalidate()
        return plan.renames

    def settle(self, document: Any, plan: Plan, *, touched: list[Any], before: bytes) -> None:
        from .tools._base import Before, Engine, seen_before

        document._note_ids()
        # The body's bytes at this version, as the patch was checked to make them: the next
        # patch's check reads them, while the document is at this version, without writing
        # the part again.
        document.__dict__["_block_body"] = (document.history.version,
                                            plan.data if plan.data is not None else plan.before)
        document.typography_spaced = getattr(document, "typography_spaced", 0) + plan.patch.spaced
        engine = plan.patch.engine
        if isinstance(engine, Engine):
            document._checks_engine = Engine(engine._run, engine.timeout, keyed=engine.keyed, document=document)
        for k, scope in enumerate(touched):
            if isinstance(scope, Before):
                # What the reflow is measured from: the document as it was just before.
                touched[k] = Before(seen_before(document, before))

    def rename(self, value: Any, renames: Mapping[str, str]) -> Any:
        return renamed(value, token_renamer(renames))

    def touches(self, change: Any, scope: Any) -> bool:
        return touches(change, scope)

    def drew(self, draft: Draft) -> bool:
        return any(values for _seed, values in draft.recording.draws)

    def release(self, draft: Draft, *, current: Any = None) -> None:
        if draft.recording is not None and "_recording" in draft.document.__dict__:
            draft.stop()
        if current is None:
            return
        document = draft.document
        document.history.clear()
        document.converter = None
        document.__dict__.pop("_checks_engine", None)
        live = document.__dict__.get("_live")
        if live is None:
            return
        with self._lock:
            kept = self._kept.setdefault(live, [])
            kept.insert(0, (current, document))
            del kept[KEEP_DRAFTS:]


def _body_now(document: Any, part: str) -> bytes | None:
    """The body part's bytes now: those the last patch applied made, while the document is
    still at the version it left (equal versions are equal content); else the part
    written."""
    kept = document.__dict__.get("_block_body")
    if kept is not None and kept[0] == document.history.version:
        return kept[1]
    return document.package.read(part)


#: Parts a change may touch without changing what a read of some blocks says: registrations
#: of parts ([Content_Types], relationships) and the document's properties.  (The
#: provenance item is told by its content, :func:`_is_sources`; ``numbering.xml`` touches a
#: read only where a list definition or instance there changed, :func:`_numbering_added`.)
_UNREAD_PARTS = re.compile(r"(^|/)\[Content_Types\]\.xml$|\.rels$|^docProps/")


def touches(change: Any, scope: Any) -> bool:
    """ooxml-edit's ``BlockEdits.touches`` (round 47): whether ``change`` touched what a read
    of some ranges covered (``scope``, :class:`docx_agent.scope.ReadScope`) -- a block of it
    changed, removed or put between two of its blocks; or any other part changed but those
    that say nothing a read of blocks shows (:data:`_UNREAD_PARTS`, the provenance item, a
    list only added to the numbering).  Anything that cannot be told: touched."""
    from .scope import ReadScope

    if not isinstance(scope, ReadScope):
        return True
    for part in change.parts:
        before, after = change.before(part), change.after(part)
        if part == scope.part:
            if _body_touches(before, after, scope.blocks, layout=scope.layout):
                return True
        elif _UNREAD_PARTS.search(part):
            continue
        elif any(data is not None and _is_sources(data) for data in (before, after)):
            continue
        elif part.endswith("numbering.xml") and _numbering_added(before, after):
            continue
        else:
            return True
    return False


def _body_touches(before: bytes | None, after: bytes | None, blocks: frozenset, *,
                  layout: bool = False) -> bool:
    """Whether the body's change touched ``blocks``; with ``layout`` (a read of their pages)
    also any change before the last of them, which can move them to other pages."""
    old, new = read_body(before), read_body(after)
    if old is None or new is None:
        return True
    last = max((k for k, identifier in enumerate(old.ids) if identifier in blocks), default=None)
    if layout and (last is None or old.tail() != new.tail()):
        return True                 # (the last section's properties: every page's)
    for i1, i2, j1, j2 in _regions(old, new):
        if layout and last is not None and i1 <= last:
            return True             # at or before the range's last block: its pages may move
        if any(identifier in blocks for identifier in old.ids[i1:i2]) \
                or any(identifier in blocks for identifier in new.ids[j1:j2]):
            return True
        if i1 == i2 and i1 > 0 and old.ids[i1 - 1] in blocks \
                and (i1 == old.count or old.ids[i1] in blocks):
            return True             # put between two of its blocks (or after its last, at the end)
    return False


def _numbering_added(before: bytes | None, after: bytes | None) -> bool:
    """Whether ``numbering.xml`` (or ``styles.xml``) only gained definitions and instances
    (lists, styles added): every element it had is there as it was."""
    if before is None or after is None:
        return False
    try:
        old, new = parse_xml(before), parse_xml(after)
    except Exception:  # noqa: BLE001 -- cannot be told
        return False
    kept = {etree.tostring(node) for node in new}
    return all(etree.tostring(node) in kept for node in old)


def _named_again(drafted: list, parts: list, moved: Mapping[str, str]) -> list:
    """The parts a call added, its ids drawn again (``moved``): a provenance item's
    properties name it by a hash of the key it was made for (``edit/provenance.py``:
    ``uuid5(NS#key)``), which renaming its text cannot reach -- named again from the key
    renamed."""
    import uuid

    from .edit import provenance as _provenance

    out = []
    keys = [node.get("key") or "" for _path, before, after in drafted
            if before is None and after is not None and _is_sources(after) for node in parse_xml(after)]
    for path, before, after in parts:
        if before is None and after is not None and b"datastoreItem" in after[:600]:
            root = parse_xml(after)
            attribute = "{%s}itemID" % _provenance.DS
            made = {"{%s}" % str(uuid.uuid5(uuid.NAMESPACE_URL, f"{_provenance.NS}#{key}")).upper(): key
                    for key in keys}
            key = made.get(root.get(attribute) or "")
            if key is not None:
                renamed_key = token_renamer(moved)(key)
                root.set(attribute, "{%s}" % str(uuid.uuid5(uuid.NAMESPACE_URL,
                                                            f"{_provenance.NS}#{renamed_key}")).upper())
                after = serialize(root)
        out.append((path, before, after))
    return out


#: Parts every block reads (round 48): a call that changes one keeps the document's lock.
_WHOLE_DOCUMENT_PARTS = re.compile(r"(^|/)(stylesWithEffects|settings|webSettings|fontTable)\.xml$"
                                   r"|(^|/)theme/")

#: The arguments that name a style to give blocks (``word_format``, ``word_format_table``).
_STYLE_KEYS = ("paragraph_style", "character_style", "style")


def _names_styles(arguments: Mapping[str, Any]) -> bool:
    ops = [op.get("arguments") for op in arguments.get("ops") or [] if isinstance(op, Mapping)]
    return any(isinstance(found, Mapping) and any(found.get(key) for key in _STYLE_KEYS)
               for found in [arguments, *ops])


def _whole_document_part(path: str, before: bytes | None, after: bytes | None, styled: bool = False) -> bool:
    """Whether a draft's change of ``path`` is one of the whole document: settings, theme,
    fonts; a style changed, or made for a call that gives blocks a style by name (a
    Markdown heading's style added on its first use stays a patch, as before); the
    numbering but for lists only added to it."""
    if _WHOLE_DOCUMENT_PARTS.search(path):
        return True
    if path.endswith("styles.xml") and before is not None:
        return styled or not _numbering_added(before, after)
    if path.endswith("numbering.xml") and before is not None:
        return not _numbering_added(before, after)
    return False


# -- parts others only added to (round 48) ---------------------------------------------------

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _additive_kind(path: str) -> str | None:
    """How a part others only added to merges with a call's change: ``"rels"`` -- the main
    document part's relationships, by ``Id`` -- or ``"styles"`` (by ``w:styleId``)."""
    if path.endswith("_rels/document.xml.rels"):
        return "rels"
    if path.endswith("/styles.xml") or path == "styles.xml":
        return "styles"
    return None


def _part_before(patch: "Patch", path: str) -> bytes | None:
    """``path`` as the draft's version had it, for a part the call did not change: the
    committed version's bytes (a part unchanged by the call is the version's)."""
    data = getattr(patch, "committed_data", None)
    if data is None:
        return None
    import io
    import zipfile

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            return archive.read(path)
    except KeyError:
        return None


def _keyed(kind: str, data: bytes) -> tuple[Any, list[tuple[Any, Any]]]:
    root = parse_xml(data)
    if kind == "rels":
        return root, [(node.get("Id"), node) for node in root]
    return root, [(node.get(_W + "styleId") if node.tag == _W + "style" else (node.tag, k), node)
                  for k, node in enumerate(root)]


def _same(a: Any, b: Any) -> bool:
    return etree.tostring(a, with_tail=False) == etree.tostring(b, with_tail=False)


def _only_added(kind: str, before: bytes, after: bytes) -> bool:
    """Whether ``after`` is ``before`` with elements added (each it had there as it was)."""
    _root, old = _keyed(kind, before)
    _made, new = _keyed(kind, after)
    found = dict(new)
    return all(key in found and _same(node, found[key]) for key, node in old)


def _merge_added(kind: str, live: bytes, before: bytes | None, after: bytes | None) -> tuple[bytes, dict]:
    """The call's additions to a part (``before`` -> ``after``, additions only, else a
    :class:`BlockConflict`) made to the document's ``live`` bytes, as the call run on them
    would: a relationship to a target the document has already is that one (its ``Id``), a
    new one the next ``rIdN`` the document has free, after the others; a style the document
    has already is its, a new one after the others.  The merged bytes and the ids renamed."""
    from ooxml_edit.opc import _match_tail

    if before is None or after is None or not _only_added(kind, before, after):
        raise BlockConflict(f"a {kind} part changed in place by the call and meanwhile")
    root, now = _keyed(kind, live)
    _old_root, old = _keyed(kind, before)
    _new_root, new = _keyed(kind, after)
    had = {key for key, _node in old}
    present = dict(now)
    renamed: dict[str, str] = {}
    for key, node in new:
        if key in had:
            continue
        if kind == "rels":
            same = next((other for _k, other in now
                         if other.get("Type") == node.get("Type") and other.get("Target") == node.get("Target")
                         and (other.get("TargetMode") == "External") == (node.get("TargetMode") == "External")),
                        None)
            if same is not None:
                renamed[key] = same.get("Id")
                continue
            used = {other.get("Id") for other in root}
            index = len(used) + 1
            while f"rId{index}" in used:
                index += 1
            identifier = f"rId{index}"
            copy = etree.SubElement(root, node.tag)
            for attribute, value in node.attrib.items():
                copy.set(attribute, identifier if attribute == "Id" else value)
            _match_tail(copy)
            now.append((identifier, copy))
            if identifier != key:
                renamed[key] = identifier
        else:
            if key in present:
                continue            # the document has it now: the call would have used it
            copy = etree.fromstring(etree.tostring(node, with_tail=False))
            root.append(copy)
            _match_tail(copy)
            present[key] = copy
    return serialize(root), renamed


_REL_REFERENCE = re.compile(rb'(r:(?:id|embed|link|pict)=")(rId\d+)(")')


def _relinked_window(window: "_Window", renamed: Mapping[str, str]) -> "_Window":
    """``window`` with the call's new relationship ids as the document gives them."""
    def relink(data: bytes) -> bytes:
        return _REL_REFERENCE.sub(lambda m: m.group(1) + renamed.get(m.group(2).decode(), m.group(2).decode()).encode()
                                  + m.group(3), data)

    import dataclasses

    return dataclasses.replace(window, after=relink(window.after),
                               blocks_after=[relink(block) for block in window.blocks_after])


_BOOKMARK_ID = re.compile(rb'(<w:bookmark(?:Start|End)\b[^>]*?\bw:id=")(\d+)(")')


def _remarked(data: bytes, ids: Mapping[str, str]) -> bytes:
    """``data`` with the bookmarks of the ids drawn (``ids``) under the ones drawn again."""
    return _BOOKMARK_ID.sub(lambda m: m.group(1) + ids.get(m.group(2).decode(), m.group(2).decode()).encode()
                            + m.group(3), data)


def _remarked_window(window: "_Window", ids: Mapping[str, str]) -> "_Window":
    import dataclasses

    return dataclasses.replace(window, after=_remarked(window.after, ids),
                               blocks_after=[_remarked(block, ids) for block in window.blocks_after])


def _is_sources(data: bytes) -> bool:
    from .edit import provenance as _provenance

    head = data[:data.find(b">", data.find(b"sources")) + 1] if b"sources" in data[:400] else b""
    return _provenance.NS.encode() in head


def _source_entries(data: bytes) -> tuple[Any, dict[tuple[str, str], Any]]:
    from .edit import provenance as _provenance

    root = parse_xml(data)
    if root.tag != _provenance._P + "sources":
        raise BlockConflict("the provenance item is not one")
    entries: dict[tuple[str, str], Any] = {}
    for node in root:
        if not isinstance(node.tag, str) or node.tag != _provenance._P + "source":
            raise BlockConflict("the provenance item holds what is not an entry")
        name = (node.get("kind") or "", node.get("key") or "")
        if name in entries:
            raise BlockConflict("the provenance item repeats an entry")
        entries[name] = node
    return root, entries


def _merge_sources(live: bytes | None, before: bytes | None, after: bytes | None) -> tuple[bytes, bytes]:
    """The provenance item (``edit/provenance.py``) another call changed since the draft's
    version, with the call's own change of it: ``(the document's bytes, the merged bytes)``.
    Each entry (a table's, a paragraph's: ``kind`` and ``key``) the call added, changed or
    dropped, as it did -- where the document's entry is still the one the call saw (else a
    :class:`BlockConflict`: both changed it); every other entry as the document has it, in
    its order, the call's new ones after them, as the call run on the document now adds
    them.  Never an item left empty (the call would have removed it): a conflict."""
    if live is None or before is None or after is None:
        raise BlockConflict("the provenance item came or went meanwhile")
    root, now = _source_entries(live)
    _root, old = _source_entries(before)
    _made, new = _source_entries(after)

    def same(a: Any, b: Any) -> bool:
        return (a is not None and b is not None and a.tag == b.tag and dict(a.attrib) == dict(b.attrib)
                and a.text == b.text and not len(a) and not len(b))

    for name in set(old) | set(new):
        mine_before, mine_after = old.get(name), new.get(name)
        if mine_before is not None and mine_after is not None and same(mine_before, mine_after):
            continue                                        # not the call's change
        theirs = now.get(name)
        if (theirs is None) != (mine_before is None) or (theirs is not None and not same(theirs, mine_before)):
            raise BlockConflict(f"the provenance of {name[1]} changed meanwhile")
        if mine_after is None:
            root.remove(theirs)                             # dropped by the call
        elif theirs is not None:
            theirs.text = mine_after.text                   # changed by the call, in place
            for attribute, value in mine_after.attrib.items():
                theirs.set(attribute, value)
    for name, node in new.items():
        if name not in old:
            copy = etree.SubElement(root, node.tag)         # added by the call: after the others
            for attribute, value in node.attrib.items():
                copy.set(attribute, value)
            copy.text = node.text
    if not len(root):
        raise BlockConflict("the provenance item would be left empty")
    return live, serialize(root)


def _placed(live: _Body, data: bytes, windows: list[_Window], count: int) -> list[tuple[int, int, _Window]]:
    """Each window in the body now (``live``, read from ``data``): found by its first block's
    id, the same blocks after it, byte for byte as they were, reaching the body's end where
    it did (``count``: the blocks the body had); in order, apart.  :class:`BlockConflict`
    naming what moved or changed otherwise."""
    placed: list[tuple[int, int, _Window]] = []
    for window in windows:
        size = window.last - window.first
        if window.first == 0:
            begin = 0
        else:
            found = live.where(window.ids[0])
            if found is None:
                raise BlockConflict(f"{window.ids[0]} is gone")
            begin = found
        end = begin + size
        if end > live.count or (window.last == count and end != live.count):
            raise BlockConflict("the blocks around a change moved")
        if live.ids[begin:end] != window.ids or data[live.at(begin):live.to(end)] != window.before:
            raise BlockConflict(f"a block near {window.ids[0] if window.ids else 'the end'} changed")
        placed.append((begin, end, window))
    placed.sort(key=lambda item: item[0])
    for (_b1, e1, _w1), (b2, _e2, _w2) in zip(placed, placed[1:]):
        if b2 < e1:
            raise BlockConflict("changes overlap")
    return placed


def _spliced(data: bytes, live: _Body, placed: list[tuple[int, int, _Window]], moved: Mapping[str, str]) -> bytes:
    """``data`` with each placed window's bytes replaced by what the call made of them, its
    ids renamed (``moved``)."""
    out = bytearray(data)
    for begin, end, window in sorted(placed, key=lambda item: item[0], reverse=True):
        out[live.at(begin):live.to(end)] = _renamed_bytes(window.after, moved)
    return bytes(out)


_BLOCK_TAGS = frozenset(_W + tag for tag in ("p", "tbl", "sdt", "customXml"))


def _splice_tree(package: Any, part: str, plan: "Plan") -> None:
    """The body's windows changed in the document's tree itself, as an edit in place changes
    it: a block that did not change is left as it is, one changed under its own id is
    rewritten in place (the same element), new blocks are inserted and removed ones go -- so
    a reader holding the tree or one of its elements finds it as after an edit -- then the
    part's bytes are checked against the plan's; anything else (text between blocks, a
    window not read so), and any difference, installs the planned bytes as the part's
    tree."""
    root = package.tree(part)
    body = root.find(W_BODY) if root is not None else None
    live = read_body(plan.before)
    done = False
    if body is not None and live is not None and plan.before is not None:
        blocks = [node for node in body if isinstance(node.tag, str) and node.tag in _BLOCK_TAGS]
        if len(blocks) == live.count and all(
                _in_place(body, blocks, begin, end, window, plan, live.start)
                for begin, end, window in sorted(plan.placed, key=lambda item: item[0], reverse=True)):
            package.mark_dirty(part)
            done = package.read(part) == plan.data
    if not done:
        package.set_tree(part, parse_xml(plan.data))


def _in_place(body: Any, blocks: list, begin: int, end: int, window: "_Window", plan: "Plan",
              start: int) -> bool:
    """One window of :func:`_splice_tree`; ``False`` when it cannot be done so (nothing changed)."""
    after = _renamed_bytes(window.after, plan.moved)
    head = plan.before[:start] if plan.before is not None else b""
    fragment = parse_xml(head + after + b"</w:body></w:document>").find(W_BODY)
    made = list(fragment)
    old = blocks[begin:end]
    if end == len(blocks):
        # A window to the body's end holds the last section's properties after its blocks:
        # as they were (the check compared the body's tail), left where they are.
        while made and (not isinstance(made[-1].tag, str) or made[-1].tag not in _BLOCK_TAGS):
            made.pop()
    if (any(not isinstance(node.tag, str) or node.tag not in _BLOCK_TAGS or (node.tail or "").strip()
            for node in made) or (fragment.text or "").strip() or len(made) != len(window.after_ids)
            or any(node.getnext() is not None and node.getnext() not in old and node is not old[-1]
                   for node in old)):
        return False
    rename = token_renamer(plan.moved) if plan.moved else (lambda text: text)
    ids = [rename(identifier) for identifier in window.after_ids]
    # Each block between the window's ends is a block of it (no bookmark or other node between).
    first = blocks[begin - 1] if begin > 0 else None
    span = []
    node = first.getnext() if first is not None else (body[0] if len(body) else None)
    if end < len(blocks):
        stop = blocks[end]
    else:
        stop = old[-1].getnext() if old else node
    while node is not None and node is not stop:
        span.append(node)
        node = node.getnext()
    if span != old:
        return False
    mine = dict(zip(window.ids, old))
    before_bytes = dict(zip(window.ids, window.blocks_before))
    after_bytes = [_renamed_bytes(data, plan.moved) for data in window.blocks_after]
    anchor = first
    used = set()
    for new, identifier, data in zip(made, ids, after_bytes):
        element = mine.get(identifier)
        if element is not None and identifier not in used:
            used.add(identifier)
            if before_bytes.get(identifier) != data:
                element.attrib.clear()
                element.attrib.update(new.attrib)
                element.text = new.text
                for child in list(element):
                    element.remove(child)
                for child in list(new):
                    element.append(child)
            placed = element
            if (anchor.getnext() if anchor is not None else body[0]) is not placed:
                tail = placed.tail
                body.remove(placed)
                placed.tail = tail
                _after(body, anchor, placed)
        else:
            new.tail = None
            _after(body, anchor, new)
            placed = new
        anchor = placed
    for identifier, element in mine.items():
        if identifier not in used:
            body.remove(element)
    return True


def _after(body: Any, anchor: Any, node: Any) -> None:
    if anchor is None:
        body.insert(0, node)
    else:
        anchor.addnext(node)


class _Replayed:
    """An op's change read from bytes, with :class:`ooxml_edit.history.Change`'s interface
    (what ``touched`` reads): the parts it changed, each one's bytes before and after."""

    def __init__(self, parts: frozenset[str], before: dict[str, bytes | None],
                 after: dict[str, bytes | None]) -> None:
        self.parts = parts
        self.changed = bool(parts)
        self._before, self._after = before, after

    def before(self, path: str) -> bytes | None:
        return self._before.get(path)

    def after(self, path: str) -> bytes | None:
        return self._after.get(path)


def _redraw(document: Any, patch: Patch) -> tuple[dict[str, str], Any]:
    """The draws made again on the document's sequences now -- what the call would draw
    there -- on a copy of its pool: ``old -> new`` for each value, and the pool after.

    Nothing committed since the draft's version, and the document's pool as the draft's was
    (pools only grow: the same size is the same set), the draws are the draft's, and its
    pool after the call is the document's: taken as they are."""
    draws = patch.draws
    pool = document._id_pool()
    if (document.history.version == patch.version and (len(pool), pool.cursors) == patch.pool_before
            and patch.pool_after is not None):
        same = _ids.IdPool(patch.pool_after)
        same.cursors.update(patch.pool_after.cursors)
        return {value: value for _seed, values in draws for value in values}, same
    scratch = _ids.IdPool(pool)
    scratch.cursors.update(pool.cursors)
    if draws:
        scratch |= _ids.used_long_hex(document._roots())
    renames: dict[str, str] = {}
    for seed, values in draws:
        for old, new in zip(values, _ids.generate(seed, scratch, len(values))):
            renames[old] = new
    return renames, scratch


def _window_after(arguments: Mapping[str, Any]) -> int:
    """How many blocks after a change a call reads: ``word_set_text`` with lines reads the
    paragraphs after its paragraph, one per further line (whether they already hold them);
    others none."""
    texts = [arguments.get("text")] + [item.get("text") for item in arguments.get("items") or []
                                       if isinstance(item, Mapping)]
    for op in arguments.get("ops") or []:
        if isinstance(op, Mapping) and isinstance(op.get("arguments"), Mapping):
            inner = op["arguments"]
            texts += [inner.get("text")] + [item.get("text") for item in inner.get("items") or []
                                            if isinstance(item, Mapping)]
    lines = max((text.count("\n") + 1 for text in texts if isinstance(text, str)), default=1)
    return lines - 1


def _section_breaks(data: bytes) -> list[bytes]:
    return re.findall(rb"<w:sectPr(?=[\s>]).*?</w:sectPr>|<w:sectPr(?=[\s>])[^>]*/>", data, re.S)


def _entry_names(package: Any) -> list[str]:
    from ooxml_common.opc import normalize_part_path

    return [normalize_part_path(entry.filename) for entry in package._entries]


def _said(result: Any) -> Any:
    """What a result says, for :func:`_renamable`: its JSON, and the application's parts."""
    return {"json": result.to_json(), "touched_by_op": result.touched_by_op}


BLOCKS = Blocks()

__all__ = ["BLOCKS", "Blocks", "CONTEXT", "DraftOps", "DraftPool", "Patch", "TOOLS", "admits", "alone", "observe",
           "read_body"]
