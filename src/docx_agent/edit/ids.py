"""Stable addressing: what every paragraph, table, row, cell and section is called.

pptx-agent's rule carries over: **ids name things, positions name places, and the id says
which it is.**  ROADMAP.md, "Addressing and stable ids", is the specification; in short:

* **A paragraph** is ``p:<paraId>`` -- Word's own ``w14:paraId``, verbatim, so it is the id
  docx2svg puts in ``data-docx-id`` -- prefixed with its story outside the body
  (``header2/p:1A2B3C4D``, ``footnotes/p:...``).
* **Missing:** a paragraph with no paraId is ``p@<story>/<n>`` (its 0-based place among
  the story's paragraphs in document order), honest about being volatile, until the first
  edit **stamps** it with a paraId and a textId.  Reading never stamps.
* **Duplicated** (a paraId is unique in its part by specification, but tools and copying
  repeat them): the first occurrence keeps ``p:<paraId>``, later ones are
  ``p:<paraId>#<k>`` (the k-th repeat).  The first edit re-issues a repeat's paraId, and
  one outside the valid range (0 < id < 0x80000000).
* **Every rename is kept as an alias**, so an id an agent holds still resolves after the
  paragraph is stamped or re-issued -- and, because an alias is consulted only while its
  target exists, after the edit is undone too.

**What the first edit stamps** is the whole document by default (``Document(stamping=
"document")``): every paragraph and table row, in every story, that lacks a valid and
unique paraId, and every drawing that lacks a ``wp14:anchorId`` and ``wp14:editId``.  Measured in Word (ROADMAP.md, "Durable across a Word save -- measured";
``tools/paraid_probe.py``): Word keeps paraIds it did not write -- through saves, edits
elsewhere and edits of the paragraph itself (whose textId it renews) -- only when every
paragraph and row of every story has a valid, unique one and every drawing its two ids;
one paragraph, row or drawing without them, in any story, or one paraId out of range, and
it writes none at all (or, in a document with
comments, new ones throughout).  A repeat it replaces, keeping the first.  ``stamping=
"paragraph"`` stamps only what an edit touches, for callers who want fewer bytes changed and
do not need ids to outlive a Word save.

Generated paraIds are deterministic: the first values of a sequence seeded by the part's
name that no paraId or textId anywhere in the document uses, so the same edit on the same
document writes the same bytes, and undoing and redoing it writes them again.
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from lxml import etree

from ..oxml.xml import MC, W14, Element, qn

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_P = _W + "p"
W_TBL = _W + "tbl"
W_TR = _W + "tr"
W_TC = _W + "tc"
PARA_ID = "{%s}paraId" % W14
TEXT_ID = "{%s}textId" % W14
MC_FALLBACK = "{%s}Fallback" % MC
MC_IGNORABLE = "{%s}Ignorable" % MC
WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
WP_INLINE = _WP + "inline"
WP_ANCHOR = _WP + "anchor"
ANCHOR_ID = "{%s}anchorId" % WP14
EDIT_ID = "{%s}editId" % WP14
MC_PROBE = "{%s}ProcessContent" % MC

#: ``ST_LongHexNumber`` values Word accepts as a paraId or textId: above 0, below 0x80000000.
PARA_ID_LIMIT = 0x80000000

_HEX = re.compile(r"[0-9A-Fa-f]{1,8}")


def valid_long_hex(raw: str | None) -> bool:
    return raw is not None and _HEX.fullmatch(raw) is not None and 0 < int(raw, 16) < PARA_ID_LIMIT


def live_elements(root: Element, tags: frozenset[str]) -> list[Element]:
    """Elements with one of ``tags`` in document order, skipping ``mc:Fallback`` subtrees: a
    fallback repeats its choice's content (paraIds included, which the specification allows)
    and is not what Word shows."""
    out: list[Element] = []
    stack = [iter(root)]
    while stack:
        for node in stack[-1]:
            tag = node.tag
            if not isinstance(tag, str) or tag == MC_FALLBACK:
                continue
            if tag in tags:
                out.append(node)
            if len(node):
                stack.append(iter(node))
                break
        else:
            stack.pop()
    return out


@dataclass
class ParagraphEntry:
    element: Element
    id: str
    #: The ``w14:paraId`` as written, or ``None``.
    raw: str | None
    #: ``ok``, ``missing`` (positional id), ``repeat`` (``#k``) or ``invalid`` (out of range).
    state: str
    #: 0-based place among the story's paragraphs.
    index: int

    @property
    def volatile(self) -> bool:
        """Whether the id can change without the paragraph being edited."""
        return self.state in ("missing", "repeat")

    @property
    def needs_stamp(self) -> bool:
        return self.state != "ok"


@dataclass
class TableEntry:
    element: Element
    id: str
    rows: list[tuple[Element, str]] = field(default_factory=list)


class PartIndex:
    """Every paragraph and table of one story part, by id and by element.

    Built from the tree on demand and thrown away on any change: rebuilding is one pass,
    and an index that outlived an edit is how ids silently drift.
    """

    def __init__(self, story: str, root: Element) -> None:
        self.story = story
        self.root = root
        prefix = "" if story == "body" else f"{story}/"
        self.paragraphs: list[ParagraphEntry] = []
        self.by_id: dict[str, ParagraphEntry | TableEntry] = {}
        self._by_element: dict[int, ParagraphEntry | TableEntry] = {}

        everything = live_elements(root, frozenset({W_P, W_TR, W_TBL}))
        seen: dict[str, int] = {}
        repeat_of: dict[int, int] = {}
        for node in everything:
            if node.tag == W_TBL:
                continue
            raw = node.get(PARA_ID)
            if raw is None:
                continue
            key = raw.upper()
            repeat_of[id(node)] = seen.get(key, 0)
            seen[key] = seen.get(key, 0) + 1

        index = 0
        for node in everything:
            if node.tag != W_P:
                continue
            raw = node.get(PARA_ID)
            if raw is None:
                entry = ParagraphEntry(node, f"p@{story}/{index}", None, "missing", index)
            else:
                repeat = repeat_of[id(node)]
                if repeat:
                    entry = ParagraphEntry(node, f"{prefix}p:{raw}#{repeat}", raw, "repeat", index)
                else:
                    state = "ok" if valid_long_hex(raw) else "invalid"
                    entry = ParagraphEntry(node, f"{prefix}p:{raw}", raw, state, index)
            self.paragraphs.append(entry)
            self.by_id[entry.id] = entry
            self._by_element[id(node)] = entry
            index += 1

        self.tables: list[TableEntry] = []
        for number, node in enumerate(n for n in everything if n.tag == W_TBL):
            rows = [row for row in live_elements(node, frozenset({W_TR})) if _owner_table(row) is node]
            first = rows[0].get(PARA_ID) if rows else None
            unique = first is not None and valid_long_hex(first) and repeat_of.get(id(rows[0])) == 0
            table_id = f"{prefix}t:{first}" if unique else f"t@{story}/{number}"
            entry_t = TableEntry(node, table_id)
            for k, row in enumerate(rows):
                raw = row.get(PARA_ID)
                ok = raw is not None and valid_long_hex(raw) and repeat_of.get(id(row)) == 0
                entry_t.rows.append((row, f"{prefix}tr:{raw}" if ok else f"{table_id}/r{k}"))
            self.tables.append(entry_t)
            self.by_id[table_id] = entry_t
            self._by_element[id(node)] = entry_t

    def entry_for(self, element: Element) -> ParagraphEntry | TableEntry | None:
        return self._by_element.get(id(element))

    def paragraph_ids(self) -> set[str]:
        return {entry.raw.upper() for entry in self.paragraphs if entry.raw}


def _owner_table(row: Element) -> Element | None:
    node = row.getparent()
    while node is not None and node.tag != W_TBL:
        node = node.getparent()
    return node


# -- generating and stamping -----------------------------------------------------------------


def used_long_hex(roots: list[Element]) -> set[int]:
    """Every paraId and textId, and every drawing's anchorId and editId, written anywhere in
    ``roots``."""
    used: set[int] = set()
    for root in roots:
        for node in root.iter(W_P, W_TR, WP_INLINE, WP_ANCHOR):
            for name in (PARA_ID, TEXT_ID, ANCHOR_ID, EDIT_ID):
                raw = node.get(name)
                if raw is not None and _HEX.fullmatch(raw):
                    used.add(int(raw, 16))
    return used


def drawings_needing_ids(root: Element) -> list[Element]:
    """The drawings (``wp:inline``, ``wp:anchor``) without a valid ``wp14:anchorId`` and
    ``wp14:editId``: Word keeps no paraId in a document that has one (measured)."""
    return [node for node in live_elements(root, frozenset({WP_INLINE, WP_ANCHOR}))
            if not (valid_long_hex(node.get(ANCHOR_ID)) and valid_long_hex(node.get(EDIT_ID)))]


def generate(seed: str, used: set[int], count: int = 1) -> list[str]:
    """``count`` new ``ST_LongHexNumber`` values, deterministic in ``seed`` and ``used``:
    the first members of the seed's sequence that ``used`` does not hold, in ``1 ..
    0x7FFFFFFF``, upper-case and eight digits wide as Word writes them.  ``used`` gains them."""
    out: list[str] = []
    counter = 0
    while len(out) < count:
        digest = hashlib.sha256(f"{seed}\0{counter}".encode()).digest()
        counter += 1
        value = int.from_bytes(digest[:4], "big") & 0x7FFFFFFF
        if value == 0 or value in used:
            continue
        used.add(value)
        out.append(f"{value:08X}")
    return out


def ensure_w14(root: Element, *, drawings: bool = False, extra: dict[str, str] | None = None) -> None:
    """Declare ``w14`` (and ``wp14``, with ``drawings``) on the part's root and list them in
    ``mc:Ignorable``, once attributes in them have been written somewhere in the tree.

    lxml declares a namespace where it is first used, under a made-up prefix; Word wants
    them declared on the root, and a consumer that predates Word 2010 must be told it may
    ignore them.  A root that already declares and ignores them is left alone.
    Declarations the document has, used or not, are kept: ``mc:Ignorable`` names prefixes.
    ``extra`` names further ignorable namespaces by prefix (``w16du``, for a revision's
    ``w16du:dateUtc``).
    """
    wanted = {"w14": W14, **({"wp14": WP14} if drawings else {}), **(extra or {})}
    found = {name: next((p for p, uri in root.nsmap.items() if uri == want and p), None)
             for name, want in wanted.items()}
    mc = next((p for p, uri in root.nsmap.items() if uri == MC and p), None)
    if mc is None or None in found.values():
        # Keep every declaration the document made; not the ones lxml just made up for the
        # attributes written before this call, which the top declarations replace.
        ours = set(wanted.values()) | {MC}
        prefixes: set[str] = set()
        for node in root.iter():
            if isinstance(node.tag, str):
                prefixes.update(p for p, uri in node.nsmap.items()
                                if p and not (uri in ours and p not in (*wanted, "mc")))
        top = {name: wanted[name] for name, prefix in found.items() if prefix is None}
        # cleanup_namespaces declares at the top only what is used below it, so mc is used
        # for a moment on a child; the attribute goes, the declaration stays.
        probe = next((child for child in root if isinstance(child.tag, str)), None)
        if mc is None and probe is not None:
            top["mc"] = MC
            probe.set(MC_PROBE, "")
        etree.cleanup_namespaces(root, top_nsmap=top, keep_ns_prefixes=sorted(prefixes))
        if probe is not None:
            probe.attrib.pop(MC_PROBE, None)
        found = {name: next((p for p, uri in root.nsmap.items() if uri == want and p), name)
                 for name, want in wanted.items()}
    ignorable = (root.get(MC_IGNORABLE) or "").split()
    missing = [prefix for prefix in found.values() if prefix not in ignorable]
    if missing:
        root.set(MC_IGNORABLE, " ".join(ignorable + missing))


def stamp(element: Element, para_id: str, text_id: str) -> None:
    element.set(PARA_ID, para_id)
    element.set(TEXT_ID, text_id)
