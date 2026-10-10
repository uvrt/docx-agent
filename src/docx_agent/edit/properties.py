"""The document's metadata: ``docProps/core.xml`` and ``docProps/app.xml``.

**Core properties** -- title, subject, author (``dc:creator``), keywords, description,
last modified by, revision, created, modified, category, status and language
(``dc:language``) -- are ordinary, undoable edits of ``core.xml``
(:meth:`Document.set_properties`); a document without the part gets one, related from the
package root, as Word writes it.  ``dc:language`` is the metadata's language: the
language the text is proofed in is the document defaults' ``w:lang``, which
:meth:`Document.new` sets and :meth:`Document.set_properties` leaves alone.

**Extended properties** (``app.xml``) restate the document's statistics for indexers and
the Finder: pages, words, characters (with and without spaces), lines and paragraphs.  Word
rewrites them on every save from a count it keeps in the background, and does not check
them on opening -- measured: Word saved a four-paragraph document it had just made from a
template with ``Paragraphs`` 1 and ``Lines`` 1, and a one-line document whose line it had
just typed with ``Words`` 1 -- so a stale count breaks nothing, but a document should not
claim one page when it has five.  :func:`refreshed_app` brings them up to date as the
document is *saved* (:meth:`Document.save`), and only when the body changed since it was
opened (what another application counted otherwise stays): words, characters and
paragraphs are counted from the body's text in the final view (Word's rules: words are
runs of non-space characters, characters exclude spaces and paragraph marks, paragraphs
are those with text); pages and lines are docx2svg's, taken from the layout of the state
being saved when one has been made (:meth:`Document.layout`, ``render_svg``), and
otherwise left as they were -- a save never lays the document out on its own.  The edits
themselves never touch ``app.xml``, so undo and redo stay exact.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult

CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CORE_PART = "docProps/core.xml"
APP_PART = "docProps/app.xml"
REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
REL_APP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
APPLICATION = "docx-agent"
_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

#: Property name -> its element, in the order Word writes them (the rest after).
CORE_PROPERTIES: dict[str, str] = {
    "title": "dc:title",
    "subject": "dc:subject",
    "author": "dc:creator",
    "keywords": "cp:keywords",
    "description": "dc:description",
    "last_modified_by": "cp:lastModifiedBy",
    "revision": "cp:revision",
    "created": "dcterms:created",
    "modified": "dcterms:modified",
    "category": "cp:category",
    "status": "cp:contentStatus",
    "language": "dc:language",
}
_DATED = {"dcterms:created", "dcterms:modified"}


def _c(tag: str) -> str:
    prefix, _, name = tag.partition(":")
    return "{%s}%s" % (CORE_NS[prefix], name)


def w3cdtf(moment: datetime) -> str:
    """``2026-10-04T09:24:00Z``: UTC, whole seconds, as Word writes it."""
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=timezone.utc)
    return moment.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_w3cdtf(text: str | None) -> datetime | None:
    if not text:
        return None
    value = text.strip()
    if value.endswith("Z"):
        value = value[:-1] + "+00:00"
    try:
        moment = datetime.fromisoformat(value)
    except ValueError:
        return None
    return moment if moment.tzinfo else moment.replace(tzinfo=timezone.utc)


def as_datetime(value) -> datetime | None:
    """``created=``'s value: ``None``, a datetime (naive is UTC) or an ISO 8601 string, as
    ``tracking(date=...)`` takes (``"2026-10-05"``, ``"2026-10-05T09:00:00Z"``)."""
    if value is None or isinstance(value, datetime):
        return value if value is None or value.tzinfo else value.replace(tzinfo=timezone.utc)
    if isinstance(value, str):
        text = value.strip()
        try:
            moment = datetime.fromisoformat(text[:-1] + "+00:00" if text.endswith("Z") else text)
        except ValueError:
            raise ValueError(f"{value!r} is not an ISO 8601 date") from None
        return moment if moment.tzinfo else moment.replace(tzinfo=timezone.utc)
    raise TypeError(f"a date is a datetime or an ISO 8601 string, not {type(value).__name__}")


def now() -> datetime:
    return datetime.now(timezone.utc).replace(microsecond=0)


def _serialize(root: Element) -> bytes:
    return _DECL + etree.tostring(root, encoding="UTF-8")


def core_xml(*, title: str | None, author: str | None, created: datetime) -> bytes:
    """A new ``core.xml``, as Word writes one for a new document (Word: its user's name as
    creator and last modifier; here the ``author`` given, or none)."""
    root = etree.Element(_c("cp:coreProperties"), nsmap=CORE_NS)
    stamp = w3cdtf(created)
    for tag, text in (("dc:title", title or ""), ("dc:subject", ""), ("dc:creator", author),
                      ("cp:keywords", ""), ("dc:description", ""), ("cp:lastModifiedBy", author),
                      ("cp:revision", "1"), ("dcterms:created", stamp), ("dcterms:modified", stamp)):
        if text is None:
            continue
        node = etree.SubElement(root, _c(tag))
        node.text = text
        if tag in _DATED:
            node.set(_c("xsi:type"), "dcterms:W3CDTF")
    return _serialize(root)


#: ``app.xml``'s statistics, by the names :func:`statistics` gives them.
STATISTICS = {"pages": "Pages", "words": "Words", "characters": "Characters", "lines": "Lines",
              "paragraphs": "Paragraphs", "characters_with_spaces": "CharactersWithSpaces"}


def app_xml(stats: dict | None = None) -> bytes:
    """A new ``app.xml`` in Word's element order: a new document's statistics (one page,
    nothing counted) unless ``stats`` are given."""
    stats = {"pages": 1, "words": 0, "characters": 0, "lines": 0, "paragraphs": 0, "characters_with_spaces": 0,
             **(stats or {})}
    root = etree.Element("{%s}Properties" % APP_NS, nsmap={None: APP_NS, "vt": VT_NS})

    def add(tag: str, text: str) -> None:
        etree.SubElement(root, "{%s}%s" % (APP_NS, tag)).text = text

    add("Template", "Normal.dotm")
    add("TotalTime", "0")
    add("Pages", str(stats["pages"]))
    add("Words", str(stats["words"]))
    add("Characters", str(stats["characters"]))
    add("Application", APPLICATION)
    add("DocSecurity", "0")
    add("Lines", str(stats["lines"]))
    add("Paragraphs", str(stats["paragraphs"]))
    add("ScaleCrop", "false")
    add("Company", "")
    add("LinksUpToDate", "false")
    add("CharactersWithSpaces", str(stats["characters_with_spaces"]))
    add("SharedDoc", "false")
    add("HyperlinksChanged", "false")
    return _serialize(root)


# -- core properties ---------------------------------------------------------------------------


def _part(package, relationship: str) -> str | None:
    for rel in package.relationships("").values():
        if rel.type == relationship and rel.target_part and package.has_part(rel.target_part):
            return rel.target_part
    return None


def core_part(package, *, create: bool = False) -> str | None:
    part = _part(package, REL_CORE)
    if part is None and create:
        part = package.unused_part_name(CORE_PART) if package.has_part(CORE_PART) else CORE_PART
        root = etree.Element(_c("cp:coreProperties"), nsmap=CORE_NS)
        package.add_part(part, _serialize(root), CT_CORE, override=True)
        package.add_relationship("", REL_CORE, part)
    return part


def read_core(package) -> dict:
    """Every core property the document sets: name -> text (dates as ``datetime``)."""
    part = core_part(package)
    root = None if part is None else package.tree(part)
    out: dict = {}
    if root is None:
        return out
    for name, tag in CORE_PROPERTIES.items():
        node = root.find(_c(tag))
        if node is None:
            continue
        text = node.text or ""
        if tag in _DATED:
            out[name] = parse_w3cdtf(text)
        elif name == "revision":
            out[name] = int(text) if text.strip().isdigit() else text
        else:
            out[name] = text
    return out


def write_core(package, values: dict) -> bool:
    """Set (a value) or remove (``None``) core properties in place; ``True`` when anything
    changed.  New elements go where Word writes them."""
    part = core_part(package, create=True)
    root = package.tree(part)
    order = list(CORE_PROPERTIES.values())
    changed = False
    for name, value in values.items():
        tag = CORE_PROPERTIES[name]
        node = root.find(_c(tag))
        if value is None:
            if node is not None:
                root.remove(node)
                changed = True
            continue
        if isinstance(value, datetime):
            text = w3cdtf(value)
        else:
            text = str(value)
        if node is None:
            node = etree.Element(_c(tag))
            rank = order.index(tag)
            after = None
            for child in root:
                child_tag = etree.QName(child).namespace, etree.QName(child).localname
                known = next((t for t in order if (CORE_NS[t.split(":")[0]], t.split(":")[1]) == child_tag), None)
                if known is not None and order.index(known) < rank:
                    after = child
            if after is None:
                root.insert(0, node)
            else:
                after.addnext(node)
        if tag in _DATED and node.get(_c("xsi:type")) != "dcterms:W3CDTF":
            node.set(_c("xsi:type"), "dcterms:W3CDTF")
            changed = True
        if node.text != text:
            node.text = text
            changed = True
    if changed:
        package.mark_dirty(part)
    return changed


def check_values(values: dict) -> None:
    unknown = sorted(set(values) - set(CORE_PROPERTIES))
    if unknown:
        raise TypeError(f"unknown document propert{'ies' if len(unknown) > 1 else 'y'}: {', '.join(unknown)} "
                        f"(one of {', '.join(CORE_PROPERTIES)})")
    for name in ("created", "modified"):
        value = values.get(name)
        if value is not None and not isinstance(value, datetime):
            if parse_w3cdtf(str(value)) is None:
                raise ValueError(f"{name} is a datetime or a W3CDTF date, not {value!r}")
            values[name] = parse_w3cdtf(str(value))
    revision = values.get("revision")
    if revision is not None and not (str(revision).isdigit()):
        raise ValueError("revision is a whole number")


class PropertyOps:
    """Document properties: :attr:`Document.properties` and :meth:`Document.set_properties`."""

    @property
    def properties(self: "Document") -> dict:
        """The core properties the document sets (``title``, ``author``, ``created``...),
        and its ``statistics`` as ``app.xml`` states them."""
        out = read_core(self.package)
        stats = read_app(self.package)
        if stats:
            out["statistics"] = stats
        return out

    def set_properties(self: "Document", **values) -> "EditResult":
        """Set core properties (:data:`CORE_PROPERTIES`'s names; ``None`` removes one): one
        undo step.  Dates are ``datetime``\\ s or W3CDTF strings, written in UTC."""
        from .document import EditResult

        check_values(values)
        current = read_core(self.package)
        if all((current.get(name) == value) or (value is None and name not in current)
               for name, value in values.items()):
            return EditResult("properties", changed=False)
        with self._edit():
            write_core(self.package, values)
        return EditResult("properties")


# -- app.xml -----------------------------------------------------------------------------------


def read_app(package) -> dict:
    part = _part(package, REL_APP)
    if part is None:
        return {}
    try:
        root = etree.fromstring(package.read(part), etree.XMLParser(resolve_entities=False))
    except etree.XMLSyntaxError:
        return {}
    out = {}
    for name, tag in STATISTICS.items():
        node = root.find("{%s}%s" % (APP_NS, tag))
        if node is not None and (node.text or "").strip().isdigit():
            out[name] = int(node.text)
    return out


def text_statistics(document: "Document") -> dict:
    """Words, characters with and without spaces, and paragraphs with text: the body's,
    in the final view, by Word's rules."""
    words = characters = spaces = paragraphs = 0
    for paragraph in document.paragraphs("body"):
        text = paragraph.text_in("current")
        if not text.strip():
            continue
        paragraphs += 1
        words += len(text.split())
        visible = sum(1 for ch in text if not ch.isspace())
        characters += visible
        spaces += sum(1 for ch in text if ch in "  \t")
    return {"words": words, "characters": characters, "characters_with_spaces": characters + spaces,
            "paragraphs": paragraphs}


def layout_statistics(document: "Document") -> dict:
    """Pages and lines from docx2svg's layout of the state being saved, when it has been
    made (and laid the whole document out); nothing otherwise."""
    from ..layout import cache_key, layout_options

    conversion = document._layouts.get(cache_key(document.to_bytes(), layout_options(document, {})))
    if conversion is None:
        return {}
    layout = conversion.layout
    if layout.pages_known is None:
        return {}
    from ..layout import _in_body

    lines = sum(1 for page in layout.pages for line in page.text_lines() if _in_body(line))
    return {"pages": layout.pages_known, "lines": lines}


def refreshed_app(document: "Document") -> bytes | None:
    """``app.xml`` with the statistics brought up to date, or ``None`` when there is no
    ``app.xml``, the body has not changed since the document was opened, or the counts are
    already right."""
    package = document.package
    part = _part(package, REL_APP)
    if part is None or package.document_part() not in package.changed_parts():
        return None
    try:
        root = etree.fromstring(package.read(part), etree.XMLParser(resolve_entities=False))
    except etree.XMLSyntaxError:
        return None
    stats = text_statistics(document)
    stats.update(layout_statistics(document))
    changed = False
    for name, value in stats.items():
        node = root.find("{%s}%s" % (APP_NS, STATISTICS[name]))
        if node is not None and node.text != str(value):
            node.text = str(value)
            changed = True
    return _serialize(root) if changed else None


__all__ = ["CORE_PROPERTIES", "PropertyOps", "app_xml", "core_xml", "read_app", "read_core", "refreshed_app",
           "text_statistics", "w3cdtf", "write_core"]
