"""Copies compared with the reference they were made from (round 35, R19).

A chapter made from a sample -- a heading, a few paragraphs, a table, copied from a sample
in a collection of samples and then filled -- is a *copy*; the sample's blocks it was made
from are its *reference*, in this document or another open one.  :func:`compare_copy` says,
for one copy, whether it still follows its reference and what differs, and where the
reference's example text is still in it.  It measures; it changes nothing
(``word_follow_sample`` makes a copy follow, ``word_compare_blocks`` pages one block's
differences).

**Blocks.**  A copy and a reference are each a block (``p:...``, ``t:...``) or a span of
the body's blocks (``p:A..t:B``; a section ``s:2`` and the other addresses ``word_read``
takes read as the top-level blocks they touch).

**Corresponding blocks.**  The copy's blocks are paired with the reference's in order, as
:mod:`.compare` pairs a cell's paragraphs: each block is read as its kind (paragraph or
table), its style (by name) and its text, normalised (case and runs of white space
ignored; a table's text is its first row's).  Blocks pair only with blocks of their kind;
the alignment keeps the order (no crossing pairs), pairs as many blocks matching by text
(:data:`.compare.TEXT_SIMILARITY` alike) as it can, then as many more matching by style.  A
copy block left over after a paired one, of its kind and in its reference block's style,
repeats it (one sample paragraph is the pattern for the copy's several).  A reference block
left unpaired is ``missing_block``, a copy block ``extra_block``.

**Differences** are :mod:`.compare`'s for each pair, each with its fix calls, and a
``kind``: ``own`` -- the copy's own content, kept by design (its row count, a cell's line
count, a line or block the sample lacks or has: text is the copy's); ``structure`` --
another kind of block, other columns or merges, a missing table; ``changed_by_copy`` -- a
style the copy has where no tool can give it the reference's (no such style here, and Word
has none built in by its name); ``format`` -- formatting ``word_follow_sample`` gives the
copy: its widths, table, row, cell and paragraph properties, styles, run formatting -- and
(round 42, R49) each paired paragraph's list item and page settings as Word applies them
(:mod:`.copy_settings`: ``list_level``, ``list_format``, ``list_numbering``, the indents and
the keeps, ``word_lists`` or ``word_format`` the fix).  A copy **follows** its reference when
every difference is its own.

**Example text** is the reference's own text the copy was meant to replace.  Precisely: the
text of every paragraph in the reference's blocks (in table cells too; insertions in,
deletions out, as Word shows the current text), but those in a table's header rows (rows
that repeat on each page, ``w:tblHeader``: column labels a copy keeps), split into
sentences (after ``.``, ``!``, ``?`` or ``…`` and a space, and at a line break or tab),
each normalised as above.  A sentence of at least :data:`EXAMPLE_MIN_WORDS` words or
:data:`EXAMPLE_MIN_CHARS` characters is an *example unit*; a shorter one -- a label such as
"Owner:" -- is not.  A copy paragraph holds example text when its normalised text contains
a unit as whole words (not inside a longer word).  Each such paragraph is reported once,
with the start of each unit it holds and the reference's paragraph; ``first_row`` marks a unit from
a table's first row that is not a marked header row (a label row, perhaps: the caller
decides).
"""

from __future__ import annotations

import hashlib
import re
from typing import TYPE_CHECKING, Any

from lxml import etree

from . import copy_settings as _settings
from . import text as _text
from .compare import _by_text, _paragraph_style, _style_name, compare_blocks, styles_read
from .compare import paragraph_pairs as _paragraph_pairs
from .follow import KIND as _FOLLOW_KIND

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_P, _TBL, _TR, _SDT = _W + "p", _W + "tbl", _W + "tr", _W + "sdt"
_TXBX = _W + "txbxContent"

#: A sentence of at least this many words is example text (the module's rule) ...
EXAMPLE_MIN_WORDS = 3
#: ... or of at least this many characters (one long word, a code, a URL).
EXAMPLE_MIN_CHARS = 20
#: The characters of a unit an entry shows.
EXAMPLE_SHOWN = 60
#: Where a reference paragraph's text breaks into sentences.
_SENTENCE = re.compile(r"(?<=[.!?…])\s+|[\n\t\v\f\r]+")

#: Each difference's ``kind`` (the module's "Differences").
KINDS = ("own", "structure", "changed_by_copy", "format")
_OWN = frozenset({"rows", "paragraphs", "missing_paragraph", "extra_paragraph", "extra_block", "list"})


def blocks_of(document: "Document", address: str) -> list[tuple[str, Any]]:
    """``(id, element)`` of the blocks ``address`` names, in order: a block or a span as the
    block tools read them (a block-level content control: the blocks in it); any other
    address ``word_read`` takes (a section, a bookmark, a cell) as the body's top-level
    blocks its paragraphs are in.  ``KeyError`` for one that names nothing or is not in
    the body."""
    from .addresses import span_ends
    from .document import _top_block
    from .errors import EditError

    main = document.package.document_part()
    elements = None
    address = document._expand_heading(address)
    if span_ends(address) is not None:
        elements, part = document._blocks_of(address)
    else:
        try:
            part, entry = document._resolve_block(address)
            if entry.element.tag in (_P, _TBL, _SDT):
                elements = [entry.element]
        except (KeyError, EditError):
            elements = None
        if elements is None:
            part, entries = document._scope_paragraphs(address)
            elements = []
            for found in entries:
                top = _top_block(found.element)
                if top is not None and not any(top is e for e in elements):
                    elements.append(top)
    if part != main:
        raise KeyError(f"{address} is not in the body: a copy and its reference are the body's blocks")
    index = document._index(main)
    out = []
    for element in elements:
        if element.tag == _SDT:
            content = element.find(_W + "sdtContent")
            element_list = [c for c in content if c.tag in (_P, _TBL)] if content is not None else []
        else:
            element_list = [element]
        for node in element_list:
            found = index.entry_for(node)
            if found is not None:
                out.append((found.id, node))
    return out


def digest(document: "Document", blocks: list[tuple[str, Any]]) -> str:
    """The blocks' content, as one digest: their ids and their XML, byte for byte."""
    hasher = hashlib.blake2b(digest_size=16)
    for identifier, element in blocks:
        hasher.update(identifier.encode())
        hasher.update(b"\0")
        hasher.update(etree.tostring(element))
        hasher.update(b"\1")
    return hasher.hexdigest()


def styles_digest(document: "Document") -> str:
    """What the comparison reads besides the blocks: the document's styles, theme and
    numbering (style names, the values a style, a theme font or a list level gives a
    property; what the numbers before the blocks show, :func:`list_context`, apart)."""
    package = document.package
    hasher = hashlib.blake2b(digest_size=16)
    main = package.document_part()
    parts = [package.styles_part(), package.numbering_part()]
    parts += package.related_parts_of_type(
        main, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
    for part in parts:
        if part is None:
            continue
        tree = package.tree(part)
        hasher.update(part.encode())
        hasher.update(etree.tostring(tree) if tree is not None else b"")
    return hasher.hexdigest()


# -- one copy ----------------------------------------------------------------------------------


def compare_copy(document: "Document", copy: list[tuple[str, Any]], reference_document: "Document",
                 reference: list[tuple[str, Any]]) -> dict[str, Any]:
    """One copy (its blocks, :func:`blocks_of`) compared with its reference (blocks of
    ``reference_document``, which may be ``document``): ``{"follows", "pairs",
    "differences", "example_text"}`` -- the pairs ``[copy block, reference block]``, every
    difference (``block``, ``property``, ..., ``fix``, ``kind``; the module says which) and
    every place the reference's example text is still in the copy."""
    with styles_read():
        return _compare_copy(document, copy, reference_document, reference)


def _compare_copy(document: "Document", copy: list[tuple[str, Any]], reference_document: "Document",
                  reference: list[tuple[str, Any]]) -> dict[str, Any]:
    other = reference_document if reference_document is not document else None
    pairs, missing, extra = align(document, copy, reference_document, reference)
    differences: list[dict] = []
    settings, compared_here = _settings_differences(document, copy, reference_document, reference, pairs)
    for (identifier, _), (reference_id, _) in pairs:
        found = compare_blocks(document, [identifier], reference_id, other)
        for difference in found["differences"]:
            if difference["property"] in _settings.REPLACED and difference.get("paragraphs"):
                # Read as Word applies them here (R49): the paragraphs compared so leave it.
                left = [p for p in difference["paragraphs"] if p not in compared_here]
                if not left:
                    continue
                difference["paragraphs"] = left
            difference["kind"] = kind_of(difference)
            differences.append(difference)
        for difference in settings.get(identifier, ()):
            difference["kind"] = kind_of(difference)
            differences.append(difference)
    for k in missing:
        reference_id, element = reference[k]
        differences.append({"block": None, "property": "missing_block",
                            "reference": {"block": reference_id, **_described(reference_document, element)},
                            "fix": None, "kind": "structure" if element.tag == _TBL else "own"})
    for k in extra:
        identifier, element = copy[k]
        differences.append({"block": identifier, "property": "extra_block",
                            "value": _described(document, element), "fix": None, "kind": "own"})
    return {"follows": all(d["kind"] == "own" for d in differences),
            "pairs": [[a, b] for (a, _), (b, _) in pairs],
            "differences": differences,
            "example_text": example_text(document, copy, reference_document, reference)}


def _region(blocks: list[tuple[str, Any]]) -> list[Any]:
    """The paragraphs of ``blocks`` in document order -- in table cells too, not in text
    boxes (the body's own, as its lists count)."""
    out = []
    for _, element in blocks:
        for paragraph in ([element] if element.tag == _P else element.iter(_P)):
            if not any(a.tag == _TXBX for a in paragraph.iterancestors()):
                out.append(paragraph)
    return out


def _settings_differences(document: "Document", copy: list[tuple[str, Any]], reference_document: "Document",
                          reference: list[tuple[str, Any]], pairs: list[tuple]) -> tuple[dict, set]:
    """``(differences by copy block, the ids of the paragraphs compared)``: each pair's
    paragraphs of the same style compared as Word applies them (:mod:`.copy_settings`)."""
    index = document._index(document.package.document_part())
    reference_index = reference_document._index(reference_document.package.document_part())
    paragraph_pairs: list[tuple] = []
    ids: dict[int, tuple[str, str]] = {}
    reference_ids: dict[int, str] = {}
    for (identifier, element), (_, reference_element) in pairs:
        for paragraph, other in _paragraph_pairs(document, element, reference_document, reference_element):
            if _paragraph_style(document, paragraph) != _paragraph_style(reference_document, other):
                continue                      # the style's difference first (the module's rule)
            entry, other_entry = index.entry_for(paragraph), reference_index.entry_for(other)
            if entry is None:
                continue
            ids[id(paragraph)] = (identifier, entry.id)
            if other_entry is not None:
                reference_ids[id(other)] = other_entry.id
            paragraph_pairs.append((paragraph, other))
    if not paragraph_pairs:
        return {}, set()
    mine = _settings.region_settings(document, _region(copy))
    theirs = _settings.region_settings(reference_document, _region(reference))
    by_block: dict[str, list[dict]] = {}
    for difference in _settings.compare_settings(document, reference_document, paragraph_pairs, ids, mine,
                                                 theirs, reference_ids):
        by_block.setdefault(difference["block"], []).append(difference)
    return by_block, {i for _, i in ids.values()}


def list_context(document: "Document", blocks: list[tuple[str, Any]]) -> str:
    """What decides the numbers the blocks' list items show besides their own XML and the
    styles and numbering (:func:`styles_digest`): the list items before them and in them,
    as a digest of each body paragraph that may be one -- its style, its own ``w:numPr``,
    whether its mark is deleted -- up to the blocks' last paragraph (``""`` where none
    is).  Part of what a kept comparison is kept by (:mod:`docx_agent.tools.scoped`): read
    from the XML alone, nothing laid out or parsed."""
    last = None
    for _, element in reversed(blocks):
        last = element if element.tag == _P else next(reversed(list(element.iter(_P))), None)
        if last is not None:
            break
    if last is None:
        return ""
    _, upto, digests = _skeleton(document)
    return digests[upto.get(id(last), 0)]


def _skeleton(document: "Document") -> tuple[list, dict[int, int], list[str]]:
    """The body's list skeleton (:func:`.copy_settings.skeleton`)."""
    found = _settings.skeleton(document)
    return found.elements, found.upto, found.digests


def kind_of(difference: dict) -> str:
    """A :mod:`.compare` difference's ``kind`` (the module's "Differences")."""
    name = difference["property"]
    if name in _OWN:
        return "own"
    if name in ("paragraph_style", "table_style"):
        return "format" if difference.get("fix") else "changed_by_copy"
    if name == "widths":
        return "format" if difference.get("fix") else "structure"
    if _FOLLOW_KIND.get(name) == "structure":
        return "structure"
    return "format"


def _described(document: "Document", element: Any) -> dict[str, Any]:
    kind, style, text = _signature(document, element)
    shown = text[:40].rstrip() + ("…" if len(text) > 40 else "")
    return {"kind": kind, "style": style, "text": shown}


def _signature(document: "Document", element: Any) -> tuple[str, str | None, str]:
    """A block's kind, style name and text, normalised (a table's: its first row's)."""
    if element.tag == _TBL:
        node = element.find(f"{_W}tblPr/{_W}tblStyle")
        style = _style_name(document, node.get(_W + "val")) if node is not None else None
        row = element.find(_TR)
        text = " | ".join(_text.paragraph_text(p) for p in row.iter(_P)) if row is not None else ""
        return "table", style, " ".join(text.split()).casefold()
    if element.tag == _P:
        return ("paragraph", _paragraph_style(document, element),
                " ".join(_text.paragraph_text(element).split()).casefold())
    return "block", None, ""


def align(document: "Document", copy: list[tuple[str, Any]], reference_document: "Document",
          reference: list[tuple[str, Any]]) -> tuple[list[tuple], list[int], list[int]]:
    """The copy's blocks paired with the reference's (the module's "Corresponding blocks"):
    ``(pairs, missing, extra)`` -- ``((copy id, element), (reference id, element))`` pairs,
    the indexes of the reference's blocks left unpaired and of the copy's."""
    m, n = len(copy), len(reference)
    mine = [_signature(document, element) for _, element in copy]
    theirs = [_signature(reference_document, element) for _, element in reference]
    weight = m + n + 1

    def score(i: int, j: int) -> int:
        if mine[i][0] != theirs[j][0]:
            return 0
        return weight * _by_text(mine[i][2], theirs[j][2]) + (mine[i][1] == theirs[j][1])

    scores = [[score(i, j) for j in range(n)] for i in range(m)]
    best = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(m - 1, -1, -1):
        for j in range(n - 1, -1, -1):
            best[i][j] = max(best[i + 1][j], best[i][j + 1],
                             scores[i][j] + best[i + 1][j + 1] if scores[i][j] else 0)
    aligned: dict[int, int] = {}
    i = j = 0
    while i < m and j < n:
        if scores[i][j] and best[i][j] == scores[i][j] + best[i + 1][j + 1]:
            aligned[i] = j
            i, j = i + 1, j + 1
        elif best[i][j] == best[i][j + 1]:
            j += 1
        else:
            i += 1
    paired = dict(aligned)
    previous = None
    for i in range(m):
        if i in aligned:
            previous = aligned[i]
        elif previous is not None and mine[i][:2] == theirs[previous][:2]:
            paired[i] = previous                      # a repeat of the pattern
    pairs = [(copy[i], reference[j]) for i, j in sorted(paired.items())]
    missing = sorted(set(range(n)) - set(aligned.values()))
    extra = [i for i in range(m) if i not in paired]
    return pairs, missing, extra


# -- example text ------------------------------------------------------------------------------


def example_units(document: "Document", blocks: list[tuple[str, Any]]) -> list[tuple[str, str, bool]]:
    """The reference's example units (the module's rule): ``(unit, paragraph id, first
    row)``, each unit once, in order."""
    index = document._index(document.package.document_part())
    seen: set[str] = set()
    out = []
    for _, element in blocks:
        for paragraph in ([element] if element.tag == _P else list(element.iter(_P))):
            row = next(paragraph.iterancestors(_TR), None)
            first_row = False
            if row is not None:
                header = row.find(f"{_W}trPr/{_W}tblHeader")
                if header is not None and (header.get(_W + "val") or "1").lower() not in ("0", "false", "off"):
                    continue
                table = row.getparent()
                first_row = table is not None and table.find(_TR) is row
            text = _text.paragraph_text(paragraph)
            for sentence in _SENTENCE.split(text):
                unit = " ".join(sentence.split()).casefold()
                if not unit or unit in seen:
                    continue
                if len(unit.split()) < EXAMPLE_MIN_WORDS and len(unit) < EXAMPLE_MIN_CHARS:
                    continue
                seen.add(unit)
                entry = index.entry_for(paragraph)
                out.append((unit, entry.id if entry is not None else "", first_row))
    return out


def example_text(document: "Document", copy: list[tuple[str, Any]], reference_document: "Document",
                 reference: list[tuple[str, Any]]) -> list[dict[str, Any]]:
    """Every copy paragraph holding example units of the reference: ``{"para", "text",
    "reference"}`` -- ``text`` the start of each unit it holds, ``reference`` the reference
    paragraph of the first; ``first_row`` where one is from a table's unmarked first row --
    in the copy's order."""
    units = [(re.compile(r"(?<!\w)" + re.escape(unit) + r"(?!\w)"), unit, source, first)
             for unit, source, first in example_units(reference_document, reference)]
    if not units:
        return []
    index = document._index(document.package.document_part())
    out = []
    for _, element in copy:
        for paragraph in ([element] if element.tag == _P else list(element.iter(_P))):
            text = " ".join(_text.paragraph_text(paragraph).split()).casefold()
            if not text:
                continue
            entry = index.entry_for(paragraph)
            found: dict[str, Any] | None = None
            for pattern, unit, source, first in units:
                if len(unit) <= len(text) and pattern.search(text):
                    shown = unit[:EXAMPLE_SHOWN].rstrip() + ("\u2026" if len(unit) > EXAMPLE_SHOWN else "")
                    if found is None:
                        found = {"para": entry.id if entry is not None else None, "text": [],
                                 "reference": source}
                        out.append(found)
                    found["text"].append(shown)
                    if first:
                        found["first_row"] = True
    return out
