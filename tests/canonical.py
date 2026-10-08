"""A canonical form of a package, for comparing what two routes to one document made
(ROADMAP.md, "Accept and reject": the invariant is held in canonical form).

What it forgets, and why each is not a difference in the document:

* ``w14:paraId``/``w14:textId`` (optionally): which element survives a join or a move is
  measured, not chosen, and a tracked edit's paragraphs are new elements;
* ``w:rsid*``, ``wp14:anchorId``/``editId``, namespace declarations and ``mc:Ignorable``;
* run boundaries: adjacent runs with the same properties and only text-like content are one
  run, adjacent ``w:t`` one ``w:t`` (a tracked insertion is a run of its own; accepted, it
  sits beside the text it joined);
* relationship ids: an ``r:id`` stands for its relationship's type and target, a media
  part for its content's hash, and a relationships part is the set of what it holds;
* revision ids and dates where asked (``revisions=False`` drops every revision record's
  ``w:id``, ``w:date`` and ``w16du:dateUtc``).
"""

from __future__ import annotations

import hashlib
import io
import posixpath
import re
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
WP14 = "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}"
MC_IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"
W16DU = "{http://schemas.microsoft.com/office/word/2023/wordml/word16du}dateUtc"
_TEXTUAL = {_W + "rPr", _W + "t", _W + "tab", _W + "br"}
_IGNORED_ATTRIBUTES = {MC_IGNORABLE, WP14 + "anchorId", WP14 + "editId", W14 + "textId"}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _merge_runs(root) -> None:
    for parent in list(root.iter()):
        children = [c for c in parent if isinstance(c.tag, str)]
        previous = None
        for child in children:
            if (previous is not None and child.tag == _W + "r" and previous.tag == _W + "r"
                    and previous.getnext() is child
                    and all(c.tag in _TEXTUAL for c in previous if isinstance(c.tag, str))
                    and all(c.tag in _TEXTUAL for c in child if isinstance(c.tag, str))
                    and _props(previous) == _props(child)):
                for grandchild in list(child):
                    if grandchild.tag != _W + "rPr":
                        previous.append(grandchild)
                parent.remove(child)
                continue
            previous = child
    for run in root.iter(_W + "r"):
        previous = None
        for child in list(run):
            if child.tag == _W + "t" and previous is not None and previous.tag == _W + "t":
                previous.text = (previous.text or "") + (child.text or "")
                run.remove(child)
                continue
            previous = child


def _props(run) -> bytes:
    properties = run.find(_W + "rPr")
    return etree.tostring(properties, method="c14n") if properties is not None and len(properties) else b""


def _annotation_ids(root, comments: dict[str, str]) -> None:
    """Annotation ids stand for what they name: a bookmark's for its name, a comment's for
    its place among the comments -- their numbers depend on what else was allocated."""
    names = {node.get(_W + "id"): node.get(_W + "name") for node in root.iter(_W + "bookmarkStart")}
    for node in root.iter(_W + "bookmarkStart", _W + "bookmarkEnd"):
        node.set(_W + "id", "bm:" + str(names.get(node.get(_W + "id"))))
    for node in root.iter(_W + "commentRangeStart", _W + "commentRangeEnd", _W + "commentReference",
                          _W + "comment"):
        node.set(_W + "id", comments.get(node.get(_W + "id"), "c?"))


def canonical_xml(data: bytes, targets: dict[str, str], *, ids: bool = True, revisions: bool = True,
                  bookmarks: bool = True, comments: dict[str, str] | None = None) -> str:
    root = etree.fromstring(data)
    _annotation_ids(root, comments or {})
    if not bookmarks:
        for node in list(root.iter(_W + "bookmarkStart", _W + "bookmarkEnd")):
            node.getparent().remove(node)
    for node in list(root.iter(_W + "rPr", _W + "pPr", _W + "tcPr", _W + "trPr")):
        if not len(node) and not node.attrib and node.getparent() is not None \
                and node.getparent().tag in (_W + "r", _W + "p", _W + "tc", _W + "tr"):
            node.getparent().remove(node)
    _merge_runs(root)
    lines: list[str] = []

    def walk(node, depth: int) -> None:
        if not isinstance(node.tag, str):
            return
        attributes = []
        for key, value in sorted(node.attrib.items()):
            if key in _IGNORED_ATTRIBUTES or key.startswith(_W + "rsid") or key == XML_SPACE:
                continue
            if not ids and key == W14 + "paraId":
                continue
            if not revisions and key in (_W + "id", _W + "date", W16DU) and node.tag != _W + "comment" \
                    and "bookmark" not in node.tag and "comment" not in node.tag:
                continue
            if key.startswith(R):
                value = targets.get(value, value)
            attributes.append(f"{key}={value!r}")
        text = (node.text or "") if node.tag in (_W + "t", _W + "delText", _W + "instrText", _W + "delInstrText") \
            else (node.text or "").strip()
        lines.append("  " * depth + node.tag + " " + " ".join(attributes) + (f" |{text}|" if text else ""))
        for child in node:
            walk(child, depth + 1)

    walk(root, 0)
    return "\n".join(lines)


def _relationships(files: dict[str, bytes], part: str) -> dict[str, str]:
    folder, name = posixpath.split(part)
    rels = posixpath.join(folder, "_rels", name + ".rels")
    out: dict[str, str] = {}
    if rels not in files:
        return out
    for rel in etree.fromstring(files[rels]):
        target = rel.get("Target")
        if rel.get("TargetMode") != "External":
            resolved = posixpath.normpath(posixpath.join(folder, target))
            if resolved in files and not resolved.endswith(".xml"):
                target = "sha:" + hashlib.sha256(files[resolved]).hexdigest()[:16]
            else:
                target = resolved
        out[rel.get("Id")] = f"{rel.get('Type').rpartition('/')[2]}->{target}"
    return out


def canonical(data: bytes, *, ids: bool = True, revisions: bool = True, bookmarks: bool = True,
              skip: tuple[str, ...] = ()) -> dict[str, str]:
    """Every part of a package in canonical form, by name (media by content)."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    out: dict[str, str] = {}
    comments: dict[str, str] = {}
    if "word/comments.xml" in files:
        for k, node in enumerate(etree.fromstring(files["word/comments.xml"]).iter(_W + "comment")):
            comments[node.get(_W + "id")] = f"c{k}"
    for name, content in files.items():
        if name in skip:
            continue
        if name.endswith(".rels"):
            folder = posixpath.dirname(posixpath.dirname(name))
            owner = posixpath.join(folder, posixpath.basename(name)[:-5])
            owner = owner.lstrip("/")
            out[name] = "\n".join(sorted(_relationships(files, owner).values()))
        elif name == "[Content_Types].xml":
            root = etree.fromstring(content)
            used = {n.rpartition(".")[2].lower() for n in files}
            # A default for an extension no part has is not a difference (ooxml-edit leaves
            # one behind when it reaps the last picture of its kind).
            out[name] = "\n".join(sorted(
                f"{c.get('PartName') or c.get('Extension')}={c.get('ContentType')}" for c in root
                if not (c.get("PartName") or "").startswith("/word/media/")
                and (c.get("Extension") is None or c.get("Extension").lower() in used)))
        elif name.endswith(".xml"):
            out[name] = canonical_xml(content, _relationships(files, name), ids=ids, revisions=revisions,
                                      bookmarks=bookmarks, comments=comments)
        elif name.startswith("word/media/"):
            out["media:" + hashlib.sha256(content).hexdigest()[:16]] = "media"
        else:
            out[name] = hashlib.sha256(content).hexdigest()
    # An empty numbering part (no list left in it) is no list at all.
    numbering = out.get("word/numbering.xml")
    if numbering is not None and len(numbering.splitlines()) == 1:
        del out["word/numbering.xml"]
        for key in ("word/_rels/document.xml.rels", "[Content_Types].xml"):
            if key in out:
                out[key] = "\n".join(line for line in out[key].splitlines() if "numbering" not in line)
    return out


def difference(left: dict[str, str], right: dict[str, str], limit: int = 40) -> str:
    """A readable account of how two canonical forms differ ('' when they do not)."""
    import difflib

    lines: list[str] = []
    for name in sorted(set(left) | set(right)):
        if left.get(name) == right.get(name):
            continue
        if name not in left or name not in right:
            lines.append(f"{name}: only in {'right' if name not in left else 'left'}")
            continue
        diff = list(difflib.unified_diff(left[name].splitlines(), right[name].splitlines(), lineterm="", n=2))
        lines.append(f"{name}:")
        lines += diff[2:2 + limit]
    return "\n".join(lines)


def without_added_styles(form: dict[str, str], original: dict[str, str]) -> dict[str, str]:
    """``form`` without the style definitions ``original`` lacks: Word keeps a style an
    edit added when the edit is rejected (style definitions are not revisions)."""
    name = "word/styles.xml"
    if name not in form or name not in original:
        return form
    keep = set(re.findall(r"styleId='([^']*)'", original[name]))
    blocks = re.split(r"\n(?=  \S)", form[name])
    kept = [b for b in blocks if not (m := re.search(r"}style .*?styleId='([^']*)'", b.split("\n")[0]))
            or m.group(1) in keep]
    return {**form, name: "\n".join(kept)}


def _blocks_of(text: str) -> tuple[str, list[str]]:
    """A canonical part split into its root line and its top-level children's blocks."""
    lines = text.split("\n")
    blocks: list[str] = []
    for line in lines[1:]:
        if line.startswith("  ") and not line.startswith("    "):
            blocks.append(line)
        elif blocks:
            blocks[-1] += "\n" + line
    return lines[0], blocks


def without_added_definitions(form: dict[str, str], original: dict[str, str]) -> dict[str, str]:
    """``form`` without what an insertion adds that is not a revision, as Word keeps it when
    the insertion is rejected: style definitions (:func:`without_added_styles`), the list
    definitions those styles carry (a level naming a style ``original`` lacks, and the
    instances over it), and footnote and endnote parts holding nothing but their
    separators, with their relationships, content types and settings entries."""
    form = dict(without_added_styles(form, original))
    keep = set(re.findall(r"styleId='([^']*)'", original.get("word/styles.xml", "")))
    name = "word/numbering.xml"
    if name in form:
        root, blocks = _blocks_of(form[name])
        gone = set()
        kept = []
        for block in blocks:
            styles = re.findall(r"}pStyle .*?val='([^']*)'", block)
            if "}abstractNum " in block.split("\n")[0] and styles and not set(styles) <= keep:
                gone |= set(re.findall(r"abstractNumId='([^']*)'", block.split("\n")[0]))
                continue
            kept.append(block)
        kept = [b for b in kept if not ("}num " in b.split("\n")[0]
                                        and set(re.findall(r"}abstractNumId .*?val='([^']*)'", b)) & gone)]
        if name not in original and not kept:
            del form[name]
            for key in ("word/_rels/document.xml.rels", "[Content_Types].xml"):
                if key in form:
                    form[key] = "\n".join(line for line in form[key].splitlines() if "numbering" not in line)
        else:
            form[name] = "\n".join([root] + kept)
    for kind in ("footnote", "endnote"):
        name = f"word/{kind}s.xml"
        if name in form and name not in original:
            _, blocks = _blocks_of(form[name])
            if all("type='separator'" in b.split("\n")[0] or "type='continuationSeparator'" in b.split("\n")[0]
                   for b in blocks):
                del form[name]
                form.pop(f"word/_rels/{kind}s.xml.rels", None)
                for key in ("word/_rels/document.xml.rels", "[Content_Types].xml"):
                    if key in form:
                        form[key] = "\n".join(line for line in form[key].splitlines() if f"{kind}s" not in line)
                settings = form.get("word/settings.xml")
                if settings is not None and f"}}{kind}Pr" not in original.get("word/settings.xml", ""):
                    root, blocks = _blocks_of(settings)
                    form["word/settings.xml"] = "\n".join(
                        [root] + [b for b in blocks if f"}}{kind}Pr " not in b.split("\n")[0] + " "])
    for name in [n for n in form if n.endswith(".rels") and n not in original and not form[n].strip()]:
        del form[name]  # a relationships part left empty (ooxml-edit's release)
    return form
