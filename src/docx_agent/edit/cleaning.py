"""A saved file with what Word keeps but never draws taken out (round 16, F11):
``save_document``'s ``clean``.

Documents that have been through many hands carry weight that changes nothing on the page:

* **orphaned parts** -- a stray zip entry no content type is declared for
  (``[trash]/0000.dat``, left by an old Office), or a part no chain of relationships
  reaches (``OpcPackage.drop_orphans``, ooxml-edit's, every format's);
* **an external ``attachedTemplate``** -- ``settings.xml``'s link to the template the
  document was made from on someone's disk (``file:///C:/Users/...``), which Word tries to
  reach on opening;
* **revision save ids** -- every ``w:rsid*`` attribute (``w:rsidR``, ``w:rsidRPr``,
  ``w:rsidRDefault``, ``w:rsidP``, ``w:rsidDel``, ``w:rsidSect``, ``w:rsidTr``) and
  ``settings.xml``'s ``w:rsids`` list, by which Word compares versions of a document; none
  is drawn, and Word starts a new list on its next save;
* **pasted web divisions** -- ``webSettings.xml``'s ``w:divs``, the HTML ``div`` tree text
  pasted from a web page brings, with each paragraph's and row's ``w:divId`` pointing into
  it.  Only when no division sets a margin or a border: those are drawn, and then the
  tree is kept whole.

Measured in Word 16.106 (round 16, ``feedback-round16/probe/f11_probe.py``, made-up
documents): with rsids on every paragraph, run and section, 6,000 listed, an external
template link, 400 zero-margin divisions and an unreferenced picture part, the cleaned file
(64,769 -> 13,066 bytes) opens without a prompt and its PDF is the original's pixel for
pixel on all 9 pages (100 dpi), the text the same.  With a ``[trash]/0000.dat`` entry too
-- a content type declared for it or not -- Word does not open the original without a
dialog (its export times out) and opens the cleaned copy (130,433 -> 13,066 bytes) at once.
A division with a 1440-twip left margin and a border is drawn (without it 8 of 9 pages
differ): kept, and the cleaned copy draws as the original.  The document in the session is
untouched: only the bytes saved are cleaned.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.opc import OpcPackage

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
REL_ATTACHED_TEMPLATE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"
REL_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
REL_WEB_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
#: A division's properties that are drawn (ECMA-376 17.15.2.8 ``div``): its margins, when
#: not 0, and its borders.
_DRAWN = ("marLeft", "marRight", "marTop", "marBottom")


def clean(data: bytes) -> tuple[bytes, dict[str, Any]]:
    """``data`` (a .docx, .dotx, .docm or .dotm) with the weight above taken out, and what
    was taken: ``{"orphaned_parts": {"count", "bytes", "parts"}, "rsids": {"attributes",
    "listed", "bytes"}, "web_divs": {"count", "bytes"}, "attached_template": {"target",
    "bytes"}, "bytes": {"before", "after"}}`` -- a kind only when something of it was
    there, ``bytes`` the XML (or part) bytes it held, uncompressed; ``bytes.before`` and
    ``after`` the files' sizes."""
    package = OpcPackage.open(data)
    report: dict[str, Any] = {}
    main = package.main_document_part()
    settings = _related(package, main, REL_SETTINGS)
    template = _attached_template(package, settings)
    if template:
        report["attached_template"] = template
    rsids = _rsids(package, settings)
    if rsids["attributes"] or rsids["listed"]:
        report["rsids"] = rsids
    divs = _web_divs(package, main)
    if divs:
        report["web_divs"] = divs
    removed = package.drop_orphans()
    if removed:
        report["orphaned_parts"] = {"count": len(removed), "bytes": sum(size for _, size in removed),
                                    "parts": [part for part, _ in removed][:20]}
    out = package.to_bytes() if report else data
    report["bytes"] = {"before": len(data), "after": len(out)}
    return out, report


def _related(package: OpcPackage, part: str | None, rel_type: str) -> str | None:
    if part is None:
        return None
    found = package.related_parts_of_type(part, rel_type)
    return found[0] if found else None


def _size(package: OpcPackage, part: str) -> int:
    return len(package.read(part) or b"")


def _attached_template(package: OpcPackage, settings: str | None) -> dict[str, Any] | None:
    """Remove settings' external ``attachedTemplate`` and its relationship."""
    if settings is None:
        return None
    root = package.tree(settings)
    for rel_id, rel in list(package.relationships(settings).items()):
        if rel.type != REL_ATTACHED_TEMPLATE or not rel.is_external:
            continue
        before = _size(package, settings)
        for node in root.iter(_W + "attachedTemplate"):
            if node.get(_R_ID) == rel_id:
                node.getparent().remove(node)
                break
        package.mark_dirty(settings)
        package.remove_relationship(settings, rel_id)
        return {"target": rel.target, "bytes": before - _size(package, settings)}
    return None


def _word_parts(package: OpcPackage) -> list[str]:
    return [part for part in package.part_names
            if part.endswith(".xml") and "wordprocessingml" in (package.content_type(part) or "")]


def _rsids(package: OpcPackage, settings: str | None) -> dict[str, int]:
    """Remove every ``w:rsid*`` attribute of the WordprocessingML parts and settings'
    ``w:rsids`` list: how many attributes, how many ids the list held, the bytes."""
    attributes = listed = saved = 0
    for part in _word_parts(package):
        data = package.read(part) or b""
        if b"rsid" not in data:
            continue
        root = package.tree(part)
        changed = False
        for node in root.iter():
            if not isinstance(node.tag, str):
                continue
            names = [name for name in node.attrib if name.startswith(_W + "rsid")]
            for name in names:
                del node.attrib[name]
            attributes += len(names)
            changed = changed or bool(names)
        if part == settings:
            for node in [node for node in root.iter(_W + "rsids")]:
                listed += sum(1 for child in node if isinstance(child.tag, str))
                node.getparent().remove(node)
                changed = True
        if changed:
            package.mark_dirty(part)
            saved += len(data) - _size(package, part)
    return {"attributes": attributes, "listed": listed, "bytes": saved}


def _web_divs(package: OpcPackage, main: str | None) -> dict[str, int] | None:
    """Remove ``webSettings``' ``w:divs`` and every ``w:divId`` -- unless a division sets
    a margin or a border, which Word draws."""
    web = _related(package, main, REL_WEB_SETTINGS)
    if web is None:
        return None
    root = package.tree(web)
    tree = [node for node in root.iter(_W + "divs") if node.getparent() is root]
    if not tree:
        return None
    for division in root.iter(_W + "div"):
        if division.find(_W + "divBdr") is not None and any(
                isinstance(side.tag, str) and side.get(_W + "val") not in (None, "nil", "none")
                for side in division.find(_W + "divBdr")):
            return None
        for name in _DRAWN:
            node = division.find(_W + name)
            if node is not None and (node.get(_W + "val") or "0") not in ("0", "auto"):
                return None
    before = _size(package, web)
    count = sum(1 for _ in root.iter(_W + "div"))
    for node in tree:
        root.remove(node)
    package.mark_dirty(web)
    saved = before - _size(package, web)
    for part in _word_parts(package):
        data = package.read(part) or b""
        if b"divId" not in data:
            continue
        found = list(package.tree(part).iter(_W + "divId"))
        for node in found:
            node.getparent().remove(node)
        if found:
            package.mark_dirty(part)
            saved += len(data) - _size(package, part)
    return {"count": count, "bytes": saved}
