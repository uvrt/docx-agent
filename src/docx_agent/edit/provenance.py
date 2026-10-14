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
    """``table`` for a durable table id, ``chart`` for a durable drawing id, else ``None``."""
    local = key.rpartition("/")[2] if not key.startswith("d:") else key
    if local.startswith("t:") and "#" not in local and _ids.valid_long_hex(local[2:]):
        return "table"
    if key.startswith("d:") and "#" not in key:
        return "chart"
    return None


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
        """Whether the table or chart ``key`` is in the document -- with its tracked
        changes accepted (``accepted``), or in the file at all."""
        kind = kind_of(key)
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


__all__ = ["ITEM_TYPE", "NS", "PROPS_TYPE", "Provenance", "ProvenanceOps", "accepted_rows", "count",
           "kind_of", "pending_rows"]
