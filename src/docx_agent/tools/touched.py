"""What a call changed in a Word document, for the application (``Result.touched``; round 16,
F12): the heading sections of the body that hold the blocks it changed.

ooxml-edit reports one ``{"part"}`` entry per part whose bytes changed (round 13).  For the
body that says only "the document"; an application that builds a report section by section
wants to know *which* section, so a body change is reported per **enclosing heading
section** instead::

    {"doc": "d1", "part": "word/document.xml", "heading": "p:301", "level": 2,
     "blocks": ["p:301", "t:7"]}

``heading`` is the innermost heading before the changed block (the block itself when it is
a heading) and ``level`` its level (1-9); ``blocks`` the first and last block of the
section it owns, as ``word_move`` and ``word_read`` take a range (``p:301..t:7``): the
heading and everything after it up to the next heading at its level or above (a lower
number) -- ``Document.section_blocks``' section, the outline level read as ``describe`` and
``word_read`` read it.  Blocks before the first heading are ``"heading": None, "level": 0``.
A block is a body paragraph, table or block-level content control, by the id it has after
the call.  One entry per section, in document order.  A change to the body that is in no
block (the last section's properties, the part's namespaces) is the plain ``{"part"}``
entry, as before; so is every other part (headers, notes, styles, numbering).

How: the body's top-level blocks before and after, each by a digest of its XML less its
``w14:paraId``/``w14:textId`` and a drawing's ``wp14:anchorId``/``wp14:editId`` (so the
first edit's stamping of ids throughout, which changes every block's bytes and nothing Word
draws, reports no section), compared in order (``difflib``); a removed block is reported in
the section that now holds the place it was in.  The blocks are found in the part's bytes
without parsing it (the tags of paragraphs, tables and content control blocks, in a part
that writes WordprocessingML as ``w:``, which Word and this library do), else from a
parse; each state's summary is kept, and the state after a call is read from the one
before it, only the stretch between the first and last bytes that differ read again.  On
a made-up 41-page document (444 blocks, 272 KB) a call spends 4.5 ms here (median; 11 ms
the first).

**Pages are not reported.**  Which pages a change moved is known only after laying the
document out again (a second or more for a long document, against milliseconds for this),
and the layout before is kept only by ``render``, of the last version drawn.  An application
that redraws only what moved renders after the change and compares the pages' images by
content: ``render`` keys each page by its SVG, so an unchanged page is the same key (the
toolbox's ``ContentStore``).
"""

from __future__ import annotations

import difflib
import hashlib
import re
from typing import Any

from lxml import etree

from ..edit import ids as _ids
from ..markdown.stylemap import DEFAULT as _STYLE_MAP

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % _W_NS
_BLOCKS = frozenset({_W + "p", _W + "tbl", _W + "sdt", _W + "customXml"})
#: The ids stamping writes, which change no block's content.
_STAMPS = re.compile(rb' (?:w14:paraId|w14:textId|wp14:anchorId|wp14:editId)="[^"]*"')
#: The tags that open and close blocks, in a part that writes the main namespace as ``w:``.
_TAG = re.compile(rb"<(/?)w:(p|tbl|sdt|customXml|body)(?=[\s/>])")
_W_DECLARED = b'xmlns:w="' + _W_NS.encode() + b'"'
_W14_DECLARED = b'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
_PARA_IDS = re.compile(rb'<w:(p|tr)(?=[\s/>])([^>]*)>')
_PARA_ID = re.compile(rb' w14:paraId="([^"]*)"')
_PARA_OR_ROW = re.compile(rb"<w:(?:p|tr)(?=[\s/>])")
_ALL_PARA_IDS = re.compile(rb'<w:(?:p|tr)(?=[\s/>])[^>]*? w14:paraId="([^"]*)"')
_FALLBACK = re.compile(rb"<mc:Fallback(?=[\s>]).*?</mc:Fallback>", re.S)
_STYLE = re.compile(rb'<w:pStyle w:val="([^"]*)"')
_OUTLINE = re.compile(rb'<w:outlineLvl w:val="([^"]*)"')
_CONTROL_ID = re.compile(rb'<w:sdtPr(?=[\s>]).*?(?:<w:id w:val="(-?\d+)"|</w:sdtPr>)', re.S)
#: Summaries kept per document: the latest few states' (a batch's ops, an undo and redo).
_KEPT = 4


def touched(document: Any, change: Any) -> list[dict[str, Any]]:
    """``DocumentFormat.touched`` for a Word document (see the module docstring)."""
    body = document.package.document_part()
    out: list[dict[str, Any]] = []
    for part in sorted(change.parts):
        if part == body:
            found = body_sections(document, change.before(part), change.after(part),
                                  styles_changed=document.package.styles_part() in change.parts)
            if found is not None:
                out += [{"part": part, **item} for item in found]
                continue
        out.append({"part": part})
    return out


def body_sections(document: Any, before: bytes | None, after: bytes | None, *,
                  styles_changed: bool = True) -> list[dict[str, Any]] | None:
    """The heading sections whose blocks differ between the body's bytes ``before`` and
    ``after``, plus a plain ``{}`` when something outside the blocks changed too; ``None``
    when either state is missing (the part added or removed)."""
    if before is None or after is None:
        return None
    old = _summary(document, before)
    new = _summary(document, after, old)
    changed: set[int] = set()
    matcher = difflib.SequenceMatcher(None, old.digests, new.digests, autojunk=False)
    for tag, _i1, _i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if j2 > j1:
            changed.update(range(j1, j2))
        elif new.digests:
            changed.add(max(0, min(j1 - 1, len(new.digests) - 1)))
    out: list[dict[str, Any]] = []
    if changed:
        levels, ids = new.levels(document, styles_changed), new.ids()
        owner, current = [], 0
        for index, level in enumerate(levels):
            if level is not None:
                current = index
            owner.append(current)
        for start in sorted({owner[j] for j in changed}):
            level = levels[start]
            end = next((k for k in range(start + 1, len(levels)) if levels[k] is not None
                        and (level is None or levels[k] <= level)), len(levels))
            out.append({"heading": ids[start] if level is not None else None, "level": level or 0,
                        "blocks": [ids[start], ids[end - 1]]})
    new.kept = None     # its levels are its own now: no chain of states held
    if old.rest != new.rest or not out:
        out.append({})
    return out


class _Summary:
    """One state of the body: its blocks' spans and digests, what is not a block as one
    digest; the blocks' ids and heading levels when asked (the state after a call only)."""

    def __init__(self, data: bytes, spans: list[tuple[int, int]] | None, digests: list[bytes], rest: bytes,
                 ids: list[str] | None = None, levels: list[int | None] | None = None,
                 body: tuple[int, int] | None = None) -> None:
        self.data, self.spans, self.digests, self.rest, self.body = data, spans, digests, rest, body
        self._ids, self._levels = ids, levels
        #: Where this state's blocks came from another's (:func:`_rescan`): the blocks kept
        #: before and after the new ones, and that state, whose levels they share.
        self.kept: tuple[int, int, "_Summary"] | None = None

    def ids(self) -> list[str]:
        if self._ids is None:
            self._ids = _fast_ids(self.data, self.spans) or _parsed(self.data)[0]
        return self._ids

    def levels(self, document: Any, styles_changed: bool = True) -> list[int | None]:
        if self._levels is None:
            levels = _Levels(document)
            kept = self.kept
            if kept is not None and not styles_changed and kept[2]._levels is not None:
                before, after, base = kept
                middle = self.spans[before:len(self.spans) - after]
                self._levels = (base._levels[:before] + [levels.of_bytes(self.data, s, e) for s, e in middle]
                                + base._levels[len(base._levels) - after:])
            else:
                self._levels = [levels.of_bytes(self.data, start, end) for start, end in self.spans]
        return self._levels


def _summary(document: Any, data: bytes, base: _Summary | None = None) -> _Summary:
    """``data``'s summary, kept on the document by the bytes' digest (the latest
    :data:`_KEPT`); from ``base``'s where only part of the body changed."""
    key = hashlib.blake2b(data, digest_size=16).digest()
    cache: dict[bytes, _Summary] = document.__dict__.setdefault("_touched_summaries", {})
    found = cache.get(key)
    if found is not None:
        cache[key] = cache.pop(key)      # the latest last
        return found
    found = (base is not None and _rescan(data, base)) or _scan(data) or _parse(document, data)
    cache[key] = found
    while len(cache) > _KEPT:
        cache.pop(next(iter(cache)))
    return found


def _digest(data: bytes) -> bytes:
    return hashlib.blake2b(_STAMPS.sub(b"", data), digest_size=16).digest()


def _blocks(data: bytes, begin: int, stop: int) -> list[tuple[int, int]] | None:
    """The top-level blocks between ``begin`` and ``stop`` (both outside any block), by
    their tags; ``None`` when the tags do not nest there."""
    depth, start = 0, begin
    spans: list[tuple[int, int]] = []
    for match in _TAG.finditer(data, begin, stop):
        if match.group(2) == b"body":
            return None
        if match.group(1):
            depth -= 1
            if depth < 0:
                return None
            if depth == 0:
                spans.append((start, data.index(b">", match.end()) + 1))
            continue
        close = data.index(b">", match.end())
        if data[close - 1] == 0x2F:          # "/": empty
            if depth == 0:
                spans.append((match.start(), close + 1))
            continue
        if depth == 0:
            start = match.start()
        depth += 1
    return spans if depth == 0 else None


def _rest(data: bytes, spans: list[tuple[int, int]], body: tuple[int, int]) -> bytes:
    """A digest of what is in the body but in no block (the last section's properties);
    the root's namespaces, which the first edit adds to, are not the body's."""
    rest = hashlib.blake2b(digest_size=16)
    previous = body[0]
    for begin, end in spans:
        rest.update(data[previous:begin])
        previous = end
    rest.update(data[previous:body[1]])
    return rest.digest()


def _scan(data: bytes) -> _Summary | None:
    """The body's blocks from its bytes, by their tags; ``None`` when the part does not
    write the main namespace as ``w:``, or the tags do not nest."""
    root_end = data.find(b">", data.find(b"<w:document"))
    if root_end < 0 or _W_DECLARED not in data[:root_end]:
        return None
    opening = re.compile(rb"<w:body(?=[\s>])").search(data, root_end)
    body_end = data.rfind(b"</w:body>")
    if opening is None or body_end < 0:
        return None
    body = (data.index(b">", opening.end()) + 1, body_end)
    spans = _blocks(data, *body)
    if spans is None:
        return None
    return _Summary(data, spans, [_digest(data[begin:end]) for begin, end in spans], _rest(data, spans, body),
                    body=body)


def _rescan(data: bytes, base: _Summary) -> _Summary | None:
    """``data``'s summary from ``base``'s, another state of the same body: the blocks
    before the first byte that differs and after the last are ``base``'s, only those
    between are read again -- a call's cost is its change's, not the document's.  ``None``
    when the change reaches outside the body, or ``base`` was not read from its bytes."""
    if base.spans is None or base.body is None:
        return None
    old = base.data
    prefix = _common_prefix(old, data)
    suffix = _common_suffix(old, data, min(len(old), len(data)) - prefix)
    delta = len(data) - len(old)
    body_at, body_end = base.body
    if prefix < body_at or len(old) - suffix > body_end:
        return None
    spans = base.spans
    before = 0
    while before < len(spans) and spans[before][1] <= prefix:
        before += 1
    after = 0
    while after < len(spans) - before and spans[len(spans) - 1 - after][0] >= len(old) - suffix:
        after += 1
    begin = spans[before - 1][1] if before else body_at
    stop = spans[len(spans) - after][0] + delta if after else body_end + delta
    middle = _blocks(data, begin, stop)
    if middle is None:
        return None
    shifted = [(s + delta, e + delta) for s, e in spans[len(spans) - after:]]
    new_spans = spans[:before] + middle + shifted
    body = (body_at, body_end + delta)
    digests = (base.digests[:before] + [_digest(data[s:e]) for s, e in middle]
               + base.digests[len(spans) - after:])
    found = _Summary(data, new_spans, digests, _rest(data, new_spans, body), body=body)
    found.kept = (before, after, base)
    return found


def _common_prefix(a: bytes, b: bytes) -> int:
    """How many leading bytes ``a`` and ``b`` share (compared a block at a time)."""
    limit, at, step = min(len(a), len(b)), 0, 4096
    while at < limit and a[at:at + step] == b[at:at + step]:
        at += step
    if at >= limit:
        return limit
    low, high = at, min(at + step, limit)
    while low < high:
        middle = (low + high + 1) // 2
        if a[at:middle] == b[at:middle]:
            low = middle
        else:
            high = middle - 1
    return low


def _common_suffix(a: bytes, b: bytes, limit: int) -> int:
    """How many trailing bytes ``a`` and ``b`` share, at most ``limit``."""
    at, step = 0, 4096
    while at < limit and a[len(a) - min(at + step, limit):len(a) - at] == b[len(b) - min(at + step, limit):len(b) - at]:
        at = min(at + step, limit)
    if at >= limit:
        return limit
    low, high = at, min(at + step, limit)
    while low < high:
        middle = (low + high + 1) // 2
        if a[len(a) - middle:len(a) - at] == b[len(b) - middle:len(b) - at]:
            low = middle
        else:
            high = middle - 1
    return low


def _fast_ids(data: bytes, spans: list[tuple[int, int]] | None) -> list[str] | None:
    """The blocks' ids from the bytes when every paragraph and row has a paraId, none
    repeated (as after this library's first edit stamps them all), and every content
    control block an id of its own; ``None`` otherwise (:func:`_parsed` names them as
    ``Document`` does)."""
    if spans is None or _W14_DECLARED not in data[:data.find(b">", data.find(b"<w:document"))]:
        return None
    live = _FALLBACK.sub(b"", data) if b"<mc:Fallback" in data else data
    stamped = _ALL_PARA_IDS.findall(live)
    if len(stamped) != len(_PARA_OR_ROW.findall(live)) or len(set(map(bytes.upper, stamped))) != len(stamped):
        return None
    ids: list[str] = []
    sdt_ids: set[bytes] = set()
    for begin, end in spans:
        head = data[begin:data.index(b">", begin) + 1]
        if head.startswith(b"<w:p") and head[4:5] in (b" ", b">", b"/"):
            ids.append("p:" + _PARA_ID.search(head).group(1).decode())
        elif head.startswith(b"<w:tbl"):
            row = _PARA_IDS.search(data, begin, end)
            while row is not None and row.group(1) != b"tr":
                row = _PARA_IDS.search(data, row.end(), end)
            raw = _PARA_ID.search(row.group(2)).group(1).decode() if row is not None else None
            if not _ids.valid_long_hex(raw):
                return None
            ids.append("t:" + raw)
        elif head.startswith(b"<w:sdt"):
            found = _CONTROL_ID.search(data, begin, end)
            if found is None or found.group(1) is None or found.group(1) in sdt_ids:
                return None
            sdt_ids.add(found.group(1))
            ids.append("cc:" + found.group(1).decode())
        else:
            return None
    return ids


def _parse(document: Any, data: bytes) -> _Summary:
    ids, levels, digests, rest = _parsed(data, document)
    return _Summary(data, None, digests, rest, ids, levels)


def _parsed(data: bytes, document: Any = None) -> tuple[list[str], list[int | None], list[bytes], bytes]:
    """Ids, heading levels (with ``document``), digests and the rest's digest from a parse
    of the part: ``Document``'s own naming (``PartIndex``)."""
    root = etree.fromstring(data, etree.XMLParser(huge_tree=True))
    body = root.find(_W + "body")
    rest = hashlib.blake2b(digest_size=16)
    if body is None:
        rest.update(data)
        return [], [], [], rest.digest()
    index = _ids.PartIndex("body", root)
    sdts = _control_ids(root)
    levels = _Levels(document) if document is not None else None
    ids: list[str] = []
    heading_levels: list[int | None] = []
    digests: list[bytes] = []
    for node in body:
        xml = _STAMPS.sub(b"", re.sub(rb' xmlns(?::[\w.-]+)?="[^"]*"', b"", etree.tostring(node)))
        if not isinstance(node.tag, str) or node.tag not in _BLOCKS:
            rest.update(xml)
            continue
        entry = index.entry_for(node)
        ids.append(entry.id if entry is not None else sdts.get(id(node), (None, f"b@body/{len(ids)}"))[1])
        heading_levels.append(levels.of(node) if levels is not None and node.tag == _W + "p" else None)
        digests.append(hashlib.blake2b(xml, digest_size=16).digest())
    return ids, heading_levels, digests, rest.digest()


def _control_ids(root: Any) -> dict[int, tuple[Any, str]]:
    """``id(sdt) -> (sdt, cc:<id>)`` for the body's content control blocks, as ``Document``
    names them; each element held, so its ``id()`` stays its own (lxml makes a proxy per
    access)."""
    out: dict[int, tuple[Any, str]] = {}
    seen: dict[str, int] = {}
    for number, node in enumerate(_ids.live_elements(root, frozenset({_W + "sdt"}))):
        value = node.find(f"{_W}sdtPr/{_W}id")
        raw = value.get(_W + "val") if value is not None else None
        if raw is None:
            out[id(node)] = (node, f"cc@body/{number}")
            continue
        count = seen.get(raw, 0)
        seen[raw] = count + 1
        out[id(node)] = (node, f"cc:{raw}" + (f"#{count}" if count else ""))
    return out


class _Levels:
    """A body paragraph's heading level (1-9) or ``None``, as ``to_markdown`` reads ``#``:
    the outline level it states, else its style's (by name, or the outline level the style
    declares through its ``basedOn`` chain) -- ``document._HeadingLevels``' rule, from the
    styles part alone."""

    def __init__(self, document: Any) -> None:
        package = document.package
        part = package.styles_part()
        root = package.tree(part) if part else None
        self.nodes = {n.get(_W + "styleId"): n for n in root.findall(_W + "style")} if root is not None else {}
        self.default = next((sid for sid, n in self.nodes.items()
                             if n.get(_W + "type") == "paragraph" and n.get(_W + "default") in ("1", "true", "on")),
                            None)
        self.memo: dict[str | None, int | None] = {}

    def of(self, paragraph: Any) -> int | None:
        properties = paragraph.find(_W + "pPr")
        direct = properties.find(_W + "outlineLvl") if properties is not None else None
        if direct is not None:
            return _direct(direct.get(_W + "val"))
        style = properties.find(_W + "pStyle") if properties is not None else None
        return self.style(style.get(_W + "val") if style is not None else None)

    def of_bytes(self, data: bytes, start: int, end: int) -> int | None:
        """:meth:`of` for the block at ``data[start:end]`` (``None`` unless a paragraph),
        read from its properties' bytes."""
        if not data.startswith(b"<w:p", start) or data[start + 4:start + 5] not in (b" ", b">", b"/"):
            return None
        at = data.index(b">", start) + 1
        if data[at - 2] == 0x2F:             # an empty paragraph: its style is the default
            return self.style(None)
        while at < end and data[at] in b" \t\r\n":
            at += 1
        style_id = None
        if data.startswith(b"<w:pPr", at) and data[at + 6:at + 7] in (b" ", b">", b"/"):
            stop = data.find(b"</w:pPr>", at, end)
            stop = data.index(b">", at) if stop < 0 else stop
            change = data.find(b"<w:pPrChange", at, stop)
            properties = data[at:stop if change < 0 else change]
            direct = _OUTLINE.search(properties)
            if direct is not None:
                return _direct(direct.group(1).decode())
            found = _STYLE.search(properties)
            style_id = found.group(1).decode() if found is not None else None
        return self.style(style_id)

    def style(self, style_id: str | None) -> int | None:
        if style_id not in self.memo:
            self.memo[style_id] = self._style_level(style_id)
        return self.memo[style_id]

    def _style_level(self, style_id: str | None) -> int | None:
        node = self.nodes.get(style_id) if style_id else None
        if node is None:
            style_id = self.default
            node = self.nodes.get(style_id) if style_id else None
        name_node = node.find(_W + "name") if node is not None else None
        name = name_node.get(_W + "val") if name_node is not None else None
        outline = None
        current, seen = style_id, set()
        while current and current not in seen and current in self.nodes:
            seen.add(current)
            level = self.nodes[current].find(f"{_W}pPr/{_W}outlineLvl")
            if level is not None:
                try:
                    outline = int(level.get(_W + "val"))
                except (TypeError, ValueError):
                    outline = None
                break
            based = self.nodes[current].find(_W + "basedOn")
            current = based.get(_W + "val") if based is not None else None
        return _STYLE_MAP.heading_level(name, outline)


def _direct(value: str | None) -> int | None:
    """A paragraph's own ``w:outlineLvl`` as a heading level: 0-8 are 1-9; 9 (body text)
    and anything else, none."""
    try:
        level = int(value)
    except (TypeError, ValueError):
        return None
    return level + 1 if 0 <= level <= 8 else None
