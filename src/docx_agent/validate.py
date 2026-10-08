"""Structural validity checks: what Word repairs, or refuses, when it is wrong.

Validation as a tool, not a hope (ROADMAP.md, "Testing strategy", layer 2).  Each check
returns :class:`Problem`\\ s with a stable ``code`` and a ``detail`` that does not depend on
positions, so a test can compare the problems of a document before and after an edit: an
edit may leave a problem the document already had (a repeated paraId it did not touch), but
must never add one.

E0's checks: parts well-formed; children in schema order (for every element whose sequence
:mod:`docx_agent.oxml.xml` registers); relationships resolve, every part has a content type,
no ``Override`` names a missing part and no part is orphaned; paraIds unique per part and in
range, each textId with a paraId; bookmarks paired and their names unique; comment ranges
and references paired with comments; every note reference has its note; complex fields
balanced; no ``w:t`` inside ``w:del`` and no ``w:delText`` outside one; every ``w:tc`` ends
in a ``w:p``; the body's ``w:sectPr`` last.

E3's: every revision record has an author and an id no other record in any story has;
move ranges are paired (start and end by id) and every move name has a source and a
destination; every comment has its reference; the modern comment parts agree with
``comments.xml`` -- each comment's last paragraph has its ``w15:commentEx`` and
``w16cid:commentId``, durable ids are unique and ``commentsExtensible`` names only them,
a reply's ``w15:paraIdParent`` is a comment's.

E1's: every relationship id a part's XML spells (``r:id``, ``r:embed``...) is one of that
part's relationships; **numbering integrity** (every ``numId`` a paragraph or style names
resolves, every instance names an abstract definition that exists, levels 0-8); every
style a paragraph, run or style names (``w:pStyle``, ``w:rStyle``, ``w:basedOn``, ``w:next``,
``w:link``) exists and ``w:basedOn`` chains end; ``wp:docPr`` ids unique; bookmark names
valid (a letter or ``_``, then letters, digits and underscores, at most 40 characters).

E2's: every prefix a part's ``mc:Ignorable`` lists is declared on its root (Word does not
open a part that names one it cannot resolve).

Given the file name it is written to (``target``): the main part's content type is the one
its extension names (``content-type-extension``) -- Word refuses a ``.docx`` whose main part
is a template's, and a ``.dotx`` whose main part is a document's (measured,
``tests/test_oracle_trial.py``) -- and a ``.docx`` or ``.dotx`` carries no VBA project.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

from lxml import etree

from .oxml.package import CONTENT_TYPES_PART, WordPackage
from .oxml.xml import CHILD_ORDER, Element, prefixed_name
from .edit.ids import PARA_ID, TEXT_ID, live_elements, valid_long_hex

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_BOOKMARK_NAME = re.compile(r"[^\W\d][\w]{0,39}")


@dataclass(frozen=True, order=True)
class Problem:
    code: str
    part: str
    detail: str = ""

    def __str__(self) -> str:
        return f"[{self.code}] {self.part}: {self.detail}"


def check(package: WordPackage, *, target=None) -> list[Problem]:
    """Every problem found in the package, sorted; with ``target`` (the file name it is
    written to), also whether its kind fits the extension."""
    problems: list[Problem] = []
    if target is not None:
        problems += extension_problems(package, target)
    roots: dict[str, Element] = {}
    for name in package.part_names:
        if not name.endswith((".xml", ".rels")):
            continue
        data = package.read(name)
        try:
            etree.fromstring(data, etree.XMLParser(resolve_entities=False))
        except etree.XMLSyntaxError as error:
            problems.append(Problem("not-well-formed", name, str(error)))
            continue
        roots[name] = package.tree(name)

    problems += _package_problems(package)
    word_parts = [name for name in roots if name.startswith("word/") and name.endswith(".xml")]
    for name in word_parts:
        root = roots[name]
        problems += _order_problems(name, root)
        problems += _paraid_problems(name, root)
        problems += _revision_problems(name, root)
        problems += _field_problems(name, root)
        problems += _table_problems(name, root)
        problems += _ignorable_problems(name, root)
    problems += _range_problems(package, roots)
    problems += _reference_problems(package, roots)
    problems += _numbering_problems(package, roots)
    problems += _style_problems(package, roots)
    problems += _drawing_problems(package, roots)
    problems += _tracking_problems(package, roots)
    problems += _comment_part_problems(package, roots)
    body = roots.get(package.document_part())
    if body is not None:
        problems += _body_problems(package.document_part(), body)
    return sorted(problems)


def extension_problems(package: WordPackage, target) -> list[Problem]:
    """The main part's content type against ``target``'s extension (``.docx``, ``.docm``,
    ``.dotx``, ``.dotm``; another extension is not checked)."""
    from ooxml_common.kinds import KINDS, kind_mismatch

    from .edit.authoring import has_macros, kind_for

    wanted = kind_for(target)
    if wanted is None:
        return []
    out: list[Problem] = []
    kind = package.kind
    mismatch = kind_mismatch(package.content_type(package.document_part()), target)
    if mismatch is not None:
        actual = mismatch.content_type
        out.append(Problem("content-type-extension", CONTENT_TYPES_PART,
                           f"a .{wanted} file whose main part is {'a .' + kind + chr(39) + 's' if kind else 'of no Word kind'} "
                           f"({actual}): Word refuses it; save() writes {mismatch.expected_type}"))
    if not KINDS[wanted].macro_enabled and has_macros(package):
        out.append(Problem("macros-in-macro-free-file", package.document_part(),
                           f"a .{wanted} cannot carry a VBA project"))
    return out


_MC_IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"


def _ignorable_problems(name: str, root: Element) -> list[Problem]:
    """Every prefix ``mc:Ignorable`` lists is declared on the root: Word refuses a part
    that names an ignorable prefix it cannot resolve (measured: it does not open it)."""
    listed = (root.get(_MC_IGNORABLE) or "").split()
    return [Problem("ignorable-prefix-undeclared", name, prefix) for prefix in listed if prefix not in root.nsmap]


def _package_problems(package: WordPackage) -> list[Problem]:
    out: list[Problem] = []
    parts = set(package.part_names)
    for name in parts:
        if name == CONTENT_TYPES_PART:
            continue
        if package.content_type(name) is None:
            out.append(Problem("no-content-type", name))
    types = package.tree(CONTENT_TYPES_PART)
    if types is not None:
        for override in types:
            if not isinstance(override.tag, str) or not override.tag.endswith("}Override"):
                continue
            target = (override.get("PartName") or "").lstrip("/")
            if target and target not in parts:
                out.append(Problem("override-without-part", CONTENT_TYPES_PART, target))
    for owner in [""] + sorted(parts):
        if owner and owner.endswith(".rels"):
            continue
        for relationship in package.relationships(owner).values():
            if not relationship.is_external and relationship.target_part not in parts:
                out.append(Problem("dangling-relationship", owner or "/",
                                   f"{relationship.type.rpartition('/')[2]} -> {relationship.target_part}"))
    for orphan in sorted(package.unreachable_parts()):
        out.append(Problem("orphaned-part", orphan))
    return out


def _order_problems(name: str, root: Element) -> list[Problem]:
    out: list[Problem] = []
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        parent = prefixed_name(element)
        if parent not in CHILD_ORDER:
            continue
        ranks = _ranks(CHILD_ORDER[parent])
        unknown = len(ranks) + 1
        last = -1
        for child in element:
            if not isinstance(child.tag, str):
                continue
            rank = ranks.get(prefixed_name(child), unknown)
            if rank < last:
                out.append(Problem("child-order", name, f"{prefixed_name(child)} in {parent}"))
                break
            last = rank
    return out


def _ranks(order) -> dict[str, int]:
    ranks: dict[str, int] = {}
    for rank, step in enumerate(order):
        for name in (step,) if isinstance(step, str) else step:
            ranks[name] = rank
    return ranks


def _paraid_problems(name: str, root: Element) -> list[Problem]:
    out: list[Problem] = []
    counts: Counter[str] = Counter()
    for node in live_elements(root, frozenset({_W + "p", _W + "tr"})):
        raw = node.get(PARA_ID)
        text = node.get(TEXT_ID)
        if text is not None and raw is None:
            out.append(Problem("textid-without-paraid", name, text))
        if text is not None and not valid_long_hex(text):
            out.append(Problem("textid-out-of-range", name, text))
        if raw is None:
            continue
        if not valid_long_hex(raw):
            out.append(Problem("paraid-out-of-range", name, raw))
        counts[raw.upper()] += 1
    out += [Problem("paraid-repeated", name, raw) for raw, n in counts.items() if n > 1]
    return out


def _revision_problems(name: str, root: Element) -> list[Problem]:
    out: list[Problem] = []
    for node in root.iter(_W + "t"):
        # Moved-from text stays in w:t (as Word writes it); deleted text is w:delText.
        if any(a.tag == _W + "del" and _owns_runs(a) for a in node.iterancestors()):
            out.append(Problem("text-in-deletion", name, (node.text or "")[:30]))
    for node in root.iter(_W + "delText"):
        if not any(a.tag == _W + "del" for a in node.iterancestors()):
            out.append(Problem("deleted-text-outside-deletion", name, (node.text or "")[:30]))
    for node in root.iter(_W + "hyperlink", _W + "fldSimple"):
        # A revision container holds runs; Word drops a hyperlink inside one (measured).
        if any(a.tag in (_W + "ins", _W + "del", _W + "moveFrom", _W + "moveTo") and _owns_runs(a)
               for a in node.iterancestors()):
            out.append(Problem("hyperlink-in-revision", name, node.tag.rpartition("}")[2]))
    return out


def _owns_runs(revision: Element) -> bool:
    """A ``w:del`` that holds runs (not the paragraph-mark or row ``w:del`` in properties)."""
    parent = revision.getparent()
    return parent is None or parent.tag not in (_W + "rPr", _W + "trPr")


def _field_problems(name: str, root: Element) -> list[Problem]:
    depth = 0
    for node in root.iter(_W + "fldChar"):
        if any(a.tag in (_W + "del", _W + "moveFrom") for a in node.iterancestors()):
            continue
        kind = node.get(_W + "fldCharType")
        if kind == "begin":
            depth += 1
        elif kind == "end":
            depth -= 1
            if depth < 0:
                return [Problem("field-unbalanced", name, "end without begin")]
    return [Problem("field-unbalanced", name, "begin without end")] if depth else []


def _table_problems(name: str, root: Element) -> list[Problem]:
    out: list[Problem] = []
    for cell in root.iter(_W + "tc"):
        blocks = [child for child in cell if child.tag in (_W + "p", _W + "tbl", _W + "sdt", _W + "customXml")]
        if not blocks or blocks[-1].tag == _W + "tbl":
            out.append(Problem("cell-not-ending-in-paragraph", name))
    for table in live_elements(root, frozenset({_W + "tbl"})):
        out += [Problem(code, name, detail) for code, detail in grid_problems(table)]
    return out


#: Cell and grid revisions: a table holding one is consistent once reviewed, not before.
_STRUCTURAL = frozenset(_W + name for name in ("cellIns", "cellDel", "cellMerge", "tcPrChange", "tblGridChange",
                                                "trPrChange"))


def grid_problems(table: Element) -> list[tuple[str, str]]:
    """E5's grid consistency: every row covers the grid exactly (``w:gridBefore``, its
    cells' spans, ``w:gridAfter``), and a vertical merge goes on only under a cell of the
    same columns.  A table holding cell or grid revisions is held to it once reviewed."""
    rows = []
    for row in table.iter(_W + "tr"):
        owner = row.getparent()
        while owner is not None and owner.tag != _W + "tbl":
            owner = owner.getparent()
        if owner is table:
            rows.append(row)
    if any(node.tag in _STRUCTURAL for row in rows for node in row.iter()) \
            or table.find(f"{_W}tblGrid/{_W}tblGridChange") is not None:
        return []
    grid = table.find(_W + "tblGrid")
    count = len(grid.findall(_W + "gridCol")) if grid is not None else 0
    out = []
    previous: dict[int, int] = {}
    first_cell = table.find(f"{_W}tr/{_W}tc")
    label = "".join(first_cell.itertext())[:30] if first_cell is not None else ""

    def number(node, default=0):
        try:
            return int(node.get(_W + "val")) if node is not None else default
        except (TypeError, ValueError):
            return default

    for k, row in enumerate(rows):
        properties = row.find(_W + "trPr")
        column = number(properties.find(_W + "gridBefore")) if properties is not None else 0
        after = number(properties.find(_W + "gridAfter")) if properties is not None else 0
        current: dict[int, int] = {}
        for cell in row.iter(_W + "tc"):
            owner = cell.getparent()
            while owner is not None and owner.tag != _W + "tr":
                owner = owner.getparent()
            if owner is not row:
                continue
            span = number(cell.find(f"{_W}tcPr/{_W}gridSpan"), 1)
            if span < 1:
                out.append(("grid-span-invalid", f"{label!r} row {k}"))
                span = 1
            merge = cell.find(f"{_W}tcPr/{_W}vMerge")
            if merge is not None and (merge.get(_W + "val") or "continue") == "continue" \
                    and previous.get(column) != span:
                out.append(("vmerge-without-start", f"{label!r} row {k} column {column}"))
            current[column] = span
            column += span
        if column + after != count:
            out.append(("grid-row-mismatch", f"{label!r} row {k}: {column + after} of {count} columns"))
        previous = current
    return out


def _body_problems(name: str, root: Element) -> list[Problem]:
    body = root.find(_W + "body")
    if body is None:
        return [Problem("no-body", name)]
    elements = [child for child in body if isinstance(child.tag, str)]
    sections = [child for child in elements if child.tag == _W + "sectPr"]
    if sections and elements[-1] is not sections[-1]:
        return [Problem("body-section-not-last", name)]
    if len(sections) > 1:
        return [Problem("body-has-two-sections", name)]
    return []


def _range_problems(package: WordPackage, roots: dict[str, Element]) -> list[Problem]:
    out: list[Problem] = []
    story_parts = [package.document_part()] + package.story_parts()
    bookmarks_start: Counter[str] = Counter()
    bookmarks_end: Counter[str] = Counter()
    names: Counter[str] = Counter()
    comment_starts, comment_ends, comment_refs = set(), set(), set()
    note_refs = {"footnote": set(), "endnote": set()}
    for part in story_parts:
        root = roots.get(part)
        if root is None:
            continue
        for node in root.iter(_W + "bookmarkStart"):
            bookmarks_start[node.get(_W + "id") or ""] += 1
            names[node.get(_W + "name") or ""] += 1
        for node in root.iter(_W + "bookmarkEnd"):
            bookmarks_end[node.get(_W + "id") or ""] += 1
        comment_starts |= {n.get(_W + "id") for n in root.iter(_W + "commentRangeStart")}
        comment_ends |= {n.get(_W + "id") for n in root.iter(_W + "commentRangeEnd")}
        comment_refs |= {n.get(_W + "id") for n in root.iter(_W + "commentReference")}
        note_refs["footnote"] |= {n.get(_W + "id") for n in root.iter(_W + "footnoteReference")}
        note_refs["endnote"] |= {n.get(_W + "id") for n in root.iter(_W + "endnoteReference")}
    main = package.document_part()
    for mark in sorted(set(bookmarks_start) | set(bookmarks_end)):
        if bookmarks_start[mark] != bookmarks_end[mark]:
            out.append(Problem("bookmark-unpaired", main, mark))
    folded: Counter[str] = Counter()
    for name, count in names.items():
        folded[name.casefold()] += count
        if not _BOOKMARK_NAME.fullmatch(name):
            out.append(Problem("bookmark-name-invalid", main, name))
    for name, count in folded.items():
        if count > 1:
            out.append(Problem("bookmark-name-repeated", main, name))
    for mark in sorted(comment_starts ^ comment_ends):
        out.append(Problem("comment-range-unpaired", main, str(mark)))
    comments = set()
    for part in package.story_parts():
        root = roots.get(part)
        if root is not None and root.tag == _W + "comments":
            comments |= {n.get(_W + "id") for n in root.iter(_W + "comment")}
    for mark in sorted(comment_refs - comments):
        out.append(Problem("comment-reference-without-comment", main, str(mark)))
    for kind in ("footnote", "endnote"):
        notes = set()
        for part in package.story_parts():
            root = roots.get(part)
            if root is not None and root.tag == _W + f"{kind}s":
                notes |= {n.get(_W + "id") for n in root.iter(_W + kind)}
        for mark in sorted(note_refs[kind] - notes):
            out.append(Problem(f"{kind}-reference-without-note", main, str(mark)))
    return out


_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"


def _reference_problems(package: WordPackage, roots: dict[str, Element]) -> list[Problem]:
    """Every ``r:`` attribute names a relationship of its own part."""
    out: list[Problem] = []
    for name, root in roots.items():
        if not name.endswith(".xml") or name.endswith(".rels") or name == CONTENT_TYPES_PART:
            continue
        relationships = None
        for node in root.iter():
            if not isinstance(node.tag, str):
                continue
            for attribute, value in node.attrib.items():
                if attribute.startswith(_R):
                    if relationships is None:
                        relationships = package.relationships(name)
                    if value and value not in relationships:
                        out.append(Problem("relationship-id-unknown", name, value))
    return out


def _story_roots(package: WordPackage, roots: dict[str, Element]) -> list[tuple[str, Element]]:
    parts = [package.document_part()] + package.story_parts()
    return [(part, roots[part]) for part in parts if part in roots]


def _numbering_problems(package: WordPackage, roots: dict[str, Element]) -> list[Problem]:
    numbering_part = package.numbering_part()
    numbering = roots.get(numbering_part) if numbering_part else None
    nums: dict[str, str | None] = {}
    abstracts: set[str] = set()
    if numbering is not None:
        abstracts = {n.get(_W + "abstractNumId") for n in numbering.findall(_W + "abstractNum")}
        for num in numbering.findall(_W + "num"):
            ref = num.find(_W + "abstractNumId")
            nums[num.get(_W + "numId")] = ref.get(_W + "val") if ref is not None else None
    out: list[Problem] = []
    styles_part = package.styles_part()
    owners = _story_roots(package, roots) + ([(styles_part, roots[styles_part])] if styles_part in roots else [])
    for part, root in owners:
        for node in root.iter(_W + "numPr"):
            num = node.find(_W + "numId")
            level = node.find(_W + "ilvl")
            if num is not None and num.get(_W + "val") not in ("0", None) and num.get(_W + "val") not in nums:
                out.append(Problem("numbering-num-missing", part, num.get(_W + "val")))
            if level is not None:
                try:
                    ok = 0 <= int(level.get(_W + "val")) <= 8
                except (TypeError, ValueError):
                    ok = False
                if not ok:
                    out.append(Problem("numbering-level-out-of-range", part, str(level.get(_W + "val"))))
    for num_id, abstract in nums.items():
        if abstract not in abstracts:
            out.append(Problem("numbering-abstract-missing", numbering_part, f"{num_id} -> {abstract}"))
    if numbering is not None:
        nsids = Counter((n.get(_W + "val") or "").upper() for n in numbering.iter(_W + "nsid"))
        out += [Problem("numbering-nsid-repeated", numbering_part, nsid) for nsid, n in nsids.items() if n > 1]
    return out


def _style_problems(package: WordPackage, roots: dict[str, Element]) -> list[Problem]:
    styles_part = package.styles_part()
    styles = roots.get(styles_part) if styles_part else None
    if styles is None:
        return []
    by_id = {s.get(_W + "styleId"): s for s in styles.findall(_W + "style")}
    out: list[Problem] = []
    for part, root in _story_roots(package, roots):
        for tag in ("pStyle", "rStyle", "tblStyle"):
            for node in root.iter(_W + tag):
                if node.get(_W + "val") not in by_id:
                    out.append(Problem("style-missing", part, f"{tag} {node.get(_W + 'val')}"))
    for style_id, node in by_id.items():
        for tag in ("basedOn", "next", "link"):
            ref = node.find(_W + tag)
            if ref is not None and ref.get(_W + "val") not in by_id:
                out.append(Problem("style-missing", styles_part, f"{style_id} {tag} {ref.get(_W + 'val')}"))
        seen = set()
        current = style_id
        while current in by_id and current not in seen:
            seen.add(current)
            ref = by_id[current].find(_W + "basedOn")
            current = ref.get(_W + "val") if ref is not None else None
        if current in seen:
            out.append(Problem("style-based-on-cycle", styles_part, style_id))
    return out


def _drawing_problems(package: WordPackage, roots: dict[str, Element]) -> list[Problem]:
    counts: Counter[str] = Counter()
    for part, root in _story_roots(package, roots):
        for node in live_elements(root, frozenset({_WP + "docPr"})):
            counts[node.get("id") or ""] += 1
    return [Problem("docpr-id-repeated", package.document_part(), raw) for raw, n in counts.items() if n > 1]


_RECORDS = frozenset(_W + name for name in (
    "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange", "tblPrChange",
    "trPrChange", "tcPrChange", "numberingChange", "cellIns", "cellDel", "cellMerge"))


def _tracking_problems(package: WordPackage, roots: dict[str, Element]) -> list[Problem]:
    """Revision records: an author each, ids unique across stories; moves paired."""
    out: list[Problem] = []
    ids: Counter[str] = Counter()
    main = package.document_part()
    for part, root in _story_roots(package, roots):
        starts: dict[str, dict[str, str]] = {"from": {}, "to": {}}
        ends: dict[str, set[str]] = {"from": set(), "to": set()}
        for node in live_elements(root, _RECORDS | frozenset(_W + f"move{k}Range{e}" for k in ("From", "To")
                                                             for e in ("Start", "End"))):
            tag = node.tag
            if tag in _RECORDS:
                if node.get(_W + "author") is None:
                    out.append(Problem("revision-without-author", part, tag.rpartition("}")[2]))
                if node.get(_W + "id") is not None:
                    ids[node.get(_W + "id")] += 1
                continue
            side = "from" if "moveFrom" in tag else "to"
            if tag.endswith("Start"):
                starts[side][node.get(_W + "id") or ""] = node.get(_W + "name") or ""
            else:
                ends[side].add(node.get(_W + "id") or "")
        for side in ("from", "to"):
            for mark in sorted(set(starts[side]) ^ ends[side]):
                out.append(Problem("move-range-unpaired", part, f"{side} {mark}"))
        names_from, names_to = set(starts["from"].values()), set(starts["to"].values())
        for name in sorted(names_from ^ names_to):
            out.append(Problem("move-unpaired", part, name))
    out += [Problem("revision-id-repeated", main, raw) for raw, n in ids.items() if n > 1]
    return out


_W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
_W16CID = "{http://schemas.microsoft.com/office/word/2016/wordml/cid}"
_W16CEX = "{http://schemas.microsoft.com/office/word/2018/wordml/cex}"


def _comment_part_problems(package: WordPackage, roots: dict[str, Element]) -> list[Problem]:
    """``comments.xml`` and the modern parts beside it agree; every comment is referenced."""
    out: list[Problem] = []
    main = package.document_part()

    def related(kind: str) -> Element | None:
        found = package.related_parts_of_type(main, kind)
        return roots.get(found[0]) if found else None

    comments = related("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
    if comments is None:
        return out
    references: set[str] = set()
    for _, root in _story_roots(package, roots):
        references |= {n.get(_W + "id") for n in root.iter(_W + "commentReference")}
    last_ids = []
    for comment in comments.findall(_W + "comment"):
        if comment.get(_W + "id") not in references:
            out.append(Problem("comment-without-reference", main, comment.get(_W + "id") or ""))
        paragraphs = comment.findall(".//" + _W + "p")
        last_ids.append((paragraphs[-1].get(PARA_ID) or "").upper() if paragraphs else "")
    extended = related("http://schemas.microsoft.com/office/2011/relationships/commentsExtended")
    if extended is not None:
        entries = {(n.get(_W15 + "paraId") or "").upper(): n for n in extended.findall(_W15 + "commentEx")}
        for para in last_ids:
            if para and para not in entries:
                out.append(Problem("comment-without-extended-entry", main, para))
        for node in entries.values():
            parent = (node.get(_W15 + "paraIdParent") or "").upper()
            if parent and parent not in entries:
                out.append(Problem("comment-parent-unknown", main, parent))
    durable_ids: Counter[str] = Counter()
    ids_root = related("http://schemas.microsoft.com/office/2016/09/relationships/commentsIds")
    if ids_root is not None:
        known = {(n.get(_W16CID + "paraId") or "").upper() for n in ids_root.findall(_W16CID + "commentId")}
        for node in ids_root.findall(_W16CID + "commentId"):
            durable_ids[(node.get(_W16CID + "durableId") or "").upper()] += 1
        for para in last_ids:
            if para and para not in known:
                out.append(Problem("comment-without-durable-id", main, para))
    out += [Problem("comment-durable-id-repeated", main, raw) for raw, n in durable_ids.items() if n > 1]
    extensible = related("http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible")
    if extensible is not None and ids_root is not None:
        for node in extensible.findall(_W16CEX + "commentExtensible"):
            raw = (node.get(_W16CEX + "durableId") or "").upper()
            if raw not in durable_ids:
                out.append(Problem("comment-extensible-unknown", main, raw))
    return out
