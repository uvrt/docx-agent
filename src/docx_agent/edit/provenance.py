"""Source addresses kept in the saved document: where a table's or a chart's values came from.

A table placed from a source (``word_insert_markdown``'s ``from``, ``word_edit_table``'s
``set_cells`` ``from``) and a chart (``edit_chart``'s ``from``) are reported in their
result's ``from``; the *record* of it (:mod:`ooxml_edit.provenance`) is also kept in the
document, so it can be read back after a save, in another session, or after Word has
opened and saved the file.

**Where.**  One custom XML item related from the document part -- the shape Word for Mac
was measured to keep through open, edit and save (2026-10-14, the oracle's W1):
``customXml/itemN.xml`` with its ``customXml/itemPropsN.xml`` (a ``ds:datastoreItem``
whose ``ds:schemaRef`` names :data:`NS`), a ``customXmlProps`` relationship from the item
and a ``customXml`` relationship from ``word/document.xml``::

    <prov:sources xmlns:prov="urn:uvrt:docx-agent:provenance:v1">
      <prov:source kind="table" key="t:04C2BB95">{"v":1,"source":"d:scores",...}</prov:source>
      <prov:source kind="chart" key="d:1">{"v":1,"source":"t:sales",...}</prov:source>
    </prov:sources>

The item is found by its root element's namespace, never by its file name (Word renumbers
items).  Other custom XML items -- a bibliography, cover-page properties, a content
control's data store -- are left as they are.

**Keys.**  A table's key is its docx-agent id, ``t:<paraId>`` -- its first row's
``w14:paraId``, which Word keeps through a save when every paragraph and row has one, as
docx-agent stamps them (:mod:`docx_agent.edit.ids`) -- with its story before it outside
the body (``header1/t:...``).  A chart's is its drawing's id, ``d:<docPr id>``.  So a key
is the address the tools take now, and it round-trips.  A volatile id (``t@body/2``, a
repeated ``d:3#1``) is no key: nothing is kept for it.

**Reading.**  :meth:`ProvenanceOps.provenance` maps each key whose table or chart is in the
document *with its tracked changes accepted* to its record: an entry whose object is gone
(deleted, or deleted as a tracked change), or whose text is not a record, is ignored and
counted (:attr:`Provenance.ignored`).  An edit that takes a table or chart out of the
file drops its entry in the same undo step (every outermost edit and batch ends by pruning
what is gone), so undoing it brings the object back with its record; a save writes the
document as it is and prunes nothing.  A table deleted as a tracked change is still in
the file: its entry stays, ignored on reading, and rejecting the deletion brings it back.

**Text written by reference** (round 36).  A value written into a paragraph by reference
(``{{d:by_topic/door/mean}}``, :mod:`ooxml_edit.tools.textfacts`) is kept in the same item,
one entry per paragraph, keyed by the paragraph's id (``p:<paraId>``, ``header1/p:...``):
``kind="text"`` and the record ``{"v":1,"paragraph":<its text>,"facts":[...]}`` -- each
fact an :mod:`ooxml_edit.provenance` fact (``source``, ``format``, ``text``, ``document``,
``version``, ``sources``, ``p`` 0 and ``at``, the character in the paragraph's text as
``find_text`` reads it: insertions in, deletions out), and ``paragraph`` the text the
places count in.  So a fact is true while its characters are its text:

* every edit here ends by following them (:meth:`ProvenanceOps._follow_text_facts`, in
  its undo step): a paragraph whose text changed keeps the facts the change did not touch,
  moved with the text before them (:func:`ooxml_edit.provenance.facts_after_edit`), and
  loses the ones it touched or borders; a paragraph renamed keeps its entry under its new
  id; a paragraph gone, or left with no fact, takes its entry away;
* reading (:meth:`ProvenanceOps.text_facts`) applies the same rule between the text kept
  and the text now, for a change made elsewhere -- Word keeps a paragraph's ``w14:paraId``
  through an edit of it and a save (measured, ROADMAP.md "Paragraphs"); one edit in
  between is one changed span, so a value with changes on both sides of it in one
  session elsewhere is dropped, never kept wrongly;
* a paragraph split keeps the facts of the part that keeps its id; text copied or moved
  to new paragraphs (a copy, a paste) takes no fact along; two paragraphs joined keep the
  first one's.

Everything is written in the caller's undo step (:meth:`~docx_agent.Document._edit`): the
item, its properties and the relationship come and go together, and the last record
removed takes all three away.
"""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, Any, Mapping

from lxml import etree
from ooxml_edit import provenance as _records
from ooxml_edit.xml import serialize

from . import ids as _ids
from . import text as _text
from .controls import DS, REL_CUSTOM_XML, REL_CUSTOM_XML_PROPS

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

#: The item's namespace, and the schema its properties name.
NS = "urn:uvrt:docx-agent:provenance:v1"
_P = "{%s}" % NS
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
#: A custom XML item's content type (Word's ``Default`` for ``.xml``) and its properties'.
ITEM_TYPE = "application/xml"
PROPS_TYPE = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"



class Provenance(dict):
    """``{address: record}`` for the tables and charts the document keeps a record of;
    ``ignored`` counts the entries left out -- their table or chart is gone (with tracked
    changes accepted), or their text is not a record."""

    ignored: int = 0


def kind_of(key: str) -> str | None:
    """``table`` for a durable table id, ``chart`` for a durable drawing id, ``text`` for a
    durable paragraph id (its text facts), else ``None``."""
    local = key.rpartition("/")[2] if not key.startswith("d:") else key
    if local.startswith("t:") and "#" not in local and _ids.valid_long_hex(local[2:]):
        return "table"
    if key.startswith("d:") and "#" not in key:
        return "chart"
    if local.startswith("p:") and _ids.valid_long_hex(local[2:]):
        return TEXT
    return None


#: An entry's kind for a paragraph's text facts.
TEXT = "text"


class ProvenanceOps:
    """Records of where placed values came from, on :class:`docx_agent.Document`."""

    def provenance(self: "Document") -> Provenance:
        """Where the values of the document's tables and charts came from, as kept in the
        file: ``{address: record}`` -- ``t:<paraId>`` for a table, ``d:<id>`` for a chart,
        each an :mod:`ooxml_edit.provenance` record (a dict, ``"v": 1``).  Only tables and
        charts the document has with its tracked changes accepted; the entries left out
        are counted in the result's ``ignored``::

            for address, record in doc.provenance().items():
                print(address, record["source"], record.get("document"))"""
        out = Provenance()
        item = self._provenance_item()
        if item is None:
            return out
        for node in _entries(self.package.tree(item)):
            key = node.get("key") or ""
            if node.get("kind") == TEXT:
                continue                       # a paragraph's text facts: text_facts()
            try:
                found = _records.loads(node.text or "")
            except _records.RecordError:
                out.ignored += 1
                continue
            if kind_of(key) is None or not self._provenance_target(key, accepted=True) or key in out:
                out.ignored += 1
                continue
            out[key] = found
        return out

    def store_provenance(self: "Document", address: str, record: Mapping[str, Any] | None) -> bool:
        """Keep ``record`` (an :mod:`ooxml_edit.provenance` record) for the table or chart
        ``address``, replacing the one it had; ``None`` drops it.  One undo step (or part of
        the caller's).  Entries whose table or chart is no longer in the file are pruned
        on the way.  ``False`` when nothing could be kept: ``address`` is no durable table
        or chart id here (a dropped record is still dropped)."""
        if record is not None:
            _records.dumps(record)  # refuse what is not a record before anything changes
        key = self._provenance_key(address)
        item = self._provenance_item()
        if item is None and (record is None or key is None):
            return key is not None
        with self._edit():
            root = self.package.tree(item) if item is not None else None
            changed = False
            for node in _entries(root):
                if node.get("key") in (address, key) and (record is None or node.get("key") != key):
                    root.remove(node)
                    changed = True
            if record is not None and key is not None:
                if root is None:
                    item, root = self._add_provenance_item(key)
                text = _records.dumps(record)
                node = next((n for n in _entries(root) if n.get("key") == key), None)
                if node is None:
                    node = etree.SubElement(root, _P + "source")
                    node.set("kind", kind_of(key) or "")
                    node.set("key", key)
                if node.text != text:
                    node.text = text
                    changed = True
            if changed:
                self.package.mark_dirty(item)
            self._prune_provenance()
        return key is not None

    # -- text written by reference (round 36) -----------------------------------------------

    def text_facts(self: "Document") -> dict[str, list[dict[str, Any]]]:
        """The values written into paragraphs by reference that are still their text, as
        kept in the file: ``{paragraph id: facts}``, each fact an :mod:`ooxml_edit.provenance`
        fact (``source``, ``text``, ``p`` 0, ``at`` its character in the paragraph's text,
        ``sources``...), in order.  A paragraph whose text changed since -- edited elsewhere
        -- keeps the facts the change left alone (see the module's notes)::

            for paragraph, facts in doc.text_facts().items():
                for fact in facts:
                    print(paragraph, fact["at"], fact["text"], fact["source"])"""
        out: dict[str, list[dict[str, Any]]] = {}
        item = self._provenance_item()
        if item is None:
            return out
        for node in _entries(self.package.tree(item)):
            if node.get("kind") != TEXT:
                continue
            key = node.get("key") or ""
            found = self._text_entry(node)
            element = self._text_paragraph(key)
            if found is None or element is None:
                continue
            kept, facts = found
            now = _text.paragraph_text(element)
            facts = _records.facts_after_edit(facts, kept, now)
            if facts:
                out.setdefault(self._text_key(key), []).extend(facts)
        return out

    def keep_text_facts(self: "Document", paragraph: str, facts: list[Mapping[str, Any]]) -> bool:
        """Keep ``facts`` -- :mod:`ooxml_edit.provenance` facts whose ``p`` is 0 and ``at``
        their place in the paragraph's text now -- for the paragraph ``paragraph``, with the
        facts it kept that they do not overlap; one undo step (or part of the caller's).  A
        fact whose text is not where it says is left out.  ``False`` when nothing could be
        kept: the paragraph has no durable id (a positional ``p@...`` one)."""
        try:
            _part, entry = self._resolve(paragraph)
        except KeyError:
            return False
        if not isinstance(entry, _ids.ParagraphEntry) or kind_of(entry.id) != TEXT:
            return False
        now = _text.paragraph_text(entry.element)
        new = _records.facts_found([{**dict(f), "p": 0} for f in facts], now)
        with self._edit():
            item = self._provenance_item()
            root = self.package.tree(item) if item is not None else None
            node = next((n for n in _entries(root) if n.get("key") == entry.id
                         and n.get("kind") == TEXT), None)
            old: list[dict[str, Any]] = []
            if node is not None:
                found = self._text_entry(node)
                if found is not None:
                    old = _records.facts_after_edit(found[1], found[0], now)
            spans = [(f["at"], f["at"] + len(f["text"])) for f in new]
            kept = [f for f in old if not any(f["at"] < end and start < f["at"] + len(f["text"])
                                              for start, end in spans)]
            merged = sorted(kept + new, key=lambda f: f["at"])
            if not merged:
                if node is not None and root is not None and item is not None:
                    root.remove(node)
                    self.package.mark_dirty(item)
                return True
            if root is None:
                item, root = self._add_provenance_item(entry.id)
            if node is None:
                node = etree.SubElement(root, _P + "source")
                node.set("kind", TEXT)
                node.set("key", entry.id)
            text = _records.dumps({"v": _records.VERSION, "paragraph": now, "facts": merged})
            if node.text != text and item is not None:
                node.text = text
                self.package.mark_dirty(item)
        return True

    def _text_entry(self: "Document", node: Any) -> tuple[str, list[dict[str, Any]]] | None:
        """A text entry's paragraph text and facts, ``None`` when it holds no record."""
        try:
            found = _records.loads(node.text or "")
        except _records.RecordError:
            return None
        kept, facts = found.get("paragraph"), found.get("facts")
        if not isinstance(kept, str) or not isinstance(facts, list):
            return None
        return kept, facts

    def _text_paragraph(self: "Document", key: str) -> Any:
        """The paragraph element a text entry's key names now (a renamed one followed), or
        ``None``."""
        try:
            _part, entry = self._resolve(key)
        except KeyError:
            return None
        return entry.element if isinstance(entry, _ids.ParagraphEntry) else None

    def _text_key(self: "Document", key: str) -> str:
        """The id a text entry's paragraph has now (its key, or the id it was renamed to)."""
        try:
            return self._resolve(key)[1].id
        except KeyError:
            return key

    def _follow_text_facts(self: "Document") -> None:
        """Inside an edit: every paragraph's text facts after it (see the module's notes) --
        moved with their text, dropped where the edit touched them, the entry under the id
        its paragraph has now, or gone with its paragraph or its last fact."""
        item = self._provenance_item()
        if item is None:
            return
        root = self.package.tree(item)
        nodes = [node for node in _entries(root) if node.get("kind") == TEXT]
        if root is None or not nodes:
            return
        self._invalidate()  # the edit may have left the indexes stale
        changed = False
        for node in nodes:
            key = node.get("key") or ""
            found = self._text_entry(node)
            element = self._text_paragraph(key)
            if found is None or element is None:
                continue                       # gone: pruned with the other records
            kept, facts = found
            now = _text.paragraph_text(element)
            name = self._text_key(key)
            if now == kept and name == key:
                continue
            facts = _records.facts_after_edit(facts, kept, now)
            changed = True
            if not facts or kind_of(name) != TEXT:
                root.remove(node)
                continue
            node.set("key", name)
            node.text = _records.dumps({"v": _records.VERSION, "paragraph": now, "facts": facts})
        if changed:
            self.package.mark_dirty(item)

    def _prune_provenance(self: "Document") -> None:
        """Inside an edit: drop the entries whose table or chart is no longer in the file
        (deleted, or never a durable id), and the item, its properties and relationship
        when no entry is left.  Every outermost edit ends with it, so the edit that deletes
        a table drops its record in the same undo step."""
        item = self._provenance_item()
        if item is None:
            return
        self._invalidate()  # the edit may have left the indexes stale
        root = self.package.tree(item)
        gone = [node for node in _entries(root)
                if not self._provenance_target(node.get("key") or "", accepted=False)]
        for node in gone:
            root.remove(node)
        if not _entries(root):
            self._remove_provenance_item(item)
        elif gone:
            self.package.mark_dirty(item)

    def _provenance_keys(self: "Document") -> list[str]:
        """Every key the item holds, whatever its object's state."""
        item = self._provenance_item()
        return [node.get("key") or "" for node in _entries(self.package.tree(item))] if item else []

    # -- the item ---------------------------------------------------------------------------

    def _provenance_item(self: "Document") -> str | None:
        """The custom XML item holding the records, found by its root's namespace."""
        main = self.package.document_part()
        for part in self.package.related_parts_of_type(main, REL_CUSTOM_XML):
            if not self.package.has_part(part):
                continue
            try:
                root = self.package.tree(part)
            except etree.XMLSyntaxError:
                continue
            if root is not None and root.tag == _P + "sources":
                return part
        return None

    def _add_provenance_item(self: "Document", first: str) -> tuple[str, Any]:
        """A new, empty item, its properties and the relationships to both."""
        package = self.package
        number = 1
        while package.has_part(f"customXml/item{number}.xml") or \
                package.has_part(f"customXml/itemProps{number}.xml"):
            number += 1
        item, props = f"customXml/item{number}.xml", f"customXml/itemProps{number}.xml"
        # Deterministic, as every id docx-agent writes: the same edit writes the same bytes.
        item_id = "{%s}" % str(uuid.uuid5(uuid.NAMESPACE_URL, f"{NS}#{first}")).upper()
        package.add_part(item, serialize(etree.Element(_P + "sources", nsmap={"prov": NS})), ITEM_TYPE)
        store = etree.Element("{%s}datastoreItem" % DS, nsmap={"ds": DS})
        store.set("{%s}itemID" % DS, item_id)
        refs = etree.SubElement(store, "{%s}schemaRefs" % DS)
        etree.SubElement(refs, "{%s}schemaRef" % DS).set("{%s}uri" % DS, NS)
        package.add_part(props, serialize(store), PROPS_TYPE, override=True)
        package.add_relationship(item, REL_CUSTOM_XML_PROPS, props)
        package.add_relationship(package.document_part(), REL_CUSTOM_XML, item)
        return item, package.tree(item)

    def _remove_provenance_item(self: "Document", item: str) -> None:
        """The item, its properties and the relationship to it, gone together."""
        package = self.package
        main = package.document_part()
        props = [p for p in package.related_parts_of_type(item, REL_CUSTOM_XML_PROPS) if package.has_part(p)]
        for rel in list(package.relationships(main).values()):
            if rel.type == REL_CUSTOM_XML and not rel.is_external and rel.target_part == item:
                package.remove_relationship(main, rel.id)
        package.remove_part(item)
        for part in props:
            if not package.relationship_sources(part):
                package.remove_part(part)

    # -- keys -------------------------------------------------------------------------------

    def _provenance_key(self: "Document", address: str) -> str | None:
        """The durable key ``address`` names now: the table's or chart's own id (an alias
        followed), or ``None`` when it has none."""
        if not isinstance(address, str):
            return None
        if address.startswith("d:"):
            return address if kind_of(address) and self._provenance_target(address, accepted=False) else None
        try:
            _part, entry = self._resolve(address)
        except KeyError:
            return None
        if not isinstance(entry, _ids.TableEntry):
            return None
        return entry.id if kind_of(entry.id) == "table" else None

    def _provenance_target(self: "Document", key: str, *, accepted: bool) -> bool:
        """Whether the table, chart or paragraph ``key`` is in the document -- with its
        tracked changes accepted (``accepted``), or in the file at all."""
        kind = kind_of(key)
        if kind == TEXT:
            return self._text_paragraph(key) is not None
        if kind == "table":
            found = self._direct(key)
            if found is None or not isinstance(found[1], _ids.TableEntry):
                return False
            return not (accepted and all(_row_deleted(row) for row, _ in found[1].rows))
        if kind == "chart":
            for _part, identifier, frame, what in self._graphics():
                if identifier == key:
                    return what == "chart" and not (accepted and _inside_deletion(frame))
        return False


def _entries(root: Any) -> list[Any]:
    return [node for node in root if node.tag == _P + "source"] if root is not None else []


def _row_deleted(row: Any) -> bool:
    """A row deleted as a tracked change (``w:trPr/w:del``)."""
    properties = row.find(_W + "trPr")
    return properties is not None and properties.find(_W + "del") is not None


def _inside_deletion(node: Any) -> bool:
    """A drawing inside a tracked deletion: gone once the changes are accepted."""
    parent = node.getparent()
    while parent is not None:
        if parent.tag in (_W + "del", _W + "moveFrom"):
            return True
        parent = parent.getparent()
    return False


def accepted_rows(table: Any) -> list[int]:
    """The table's rows (as the tools count them) that stay when its tracked changes are
    accepted: every row but those deleted as a tracked change."""
    return [k for k, (row, _) in enumerate(table.rows) if not _row_deleted(row)]


def pending_rows(table: Any) -> bool:
    """Whether a row of the table is inserted or deleted as a tracked change."""
    for row, _ in table.rows:
        properties = row.find(_W + "trPr")
        if properties is not None and (properties.find(_W + "del") is not None
                                       or properties.find(_W + "ins") is not None):
            return True
    return False


def count(rows: str) -> int:
    """How many rows a record's ``rows`` (``1-3,5``) names."""
    total = 0
    for run in filter(None, (rows or "").split(",")):
        first, _, last = run.partition("-")
        total += (int(last) - int(first) + 1) if last else 1
    return total


__all__ = ["ITEM_TYPE", "NS", "PROPS_TYPE", "Provenance", "ProvenanceOps", "TEXT", "accepted_rows",
           "count", "kind_of", "pending_rows"]
