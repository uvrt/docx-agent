"""Layout facts: what a reviewer judges a document's pages by, measured (round 16, F8).

``check`` with ``include: ["layout"]`` (or ``min_size`` or ``styles``) reports these, read
from the one layout docx2svg made of the document's state -- the layout ``reflow`` reports
and ``render`` draws (:class:`docx_agent.layout.DocumentLayout`), never a second one -- and
from the style cascade docx2svg resolves (:mod:`docx_agent.edit.effective`).  Facts, not
verdicts: each entry says what is there and where (block ids as the Word tools take them,
1-based pages as ``render`` counts them), and, where a tool can make the change, the call
that makes it (``{"tool", "calls": [arguments...]}``, one call per entry: ``word_format``
takes its targets as one list with the values for all, so it is ``calls``, not ``items``).

``smallest_text``
    The smallest size drawn, pt, per kind: ``body``, ``table`` (text in a table cell),
    ``footnote``, ``endnote``, ``header``, ``footer``.  Superscript and subscript text (a
    note's reference mark) is left out: it is drawn small by design.
``small_text`` (with ``min_size``)
    ``{"para", "kind", "size", "chars", "pages", "style_size"}``: text drawn below the
    minimum, one entry per paragraph, kind and size, smallest first.  ``style_size`` is what
    the paragraph's styles give it; where that is at least the minimum the small size is
    direct formatting, and ``fix`` clears it (``word_format`` ``clear: ["size"]``).
``direct_formatting``
    ``{"para", "style", "props", "style_has", "pages", "fix"}``: properties set on the
    paragraph or on all its text that make it differ from what its styles (paragraph,
    table and character style) give it -- ``props`` the direct values, ``style_has`` the
    styles' -- in ``word_format``'s names (``size``, ``font``, ``bold``, ``alignment``,
    ``space_before``...); ``fix`` clears them (``word_format`` ``clear``), back to the
    style.  ``in_part``: run properties set on only some of the text (emphasis, perhaps),
    named, not cleared by ``fix`` -- a range target (``p:A@4:11``) clears those.
``styles_not_allowed``, ``styles_unknown`` (with ``styles``)
    ``{"style", "kind", "count", "blocks", "pages"}`` for each paragraph or table style in
    use that the list does not name (most used first); the listed names the document has no
    style for.
``spacers``
    ``{"paras", "height_pt", "page", "next", "fix", "keep_space"}``: top-level empty
    paragraphs, a run of them, used as vertical space: ``height_pt`` how far the next block
    would move up without them, ``fix`` deletes them (``word_delete``), ``keep_space``
    gives the next paragraph that space as ``space_before`` instead (``word_format``).
    ``page_end`` where the run reaches the foot of its page (the next block is on a later
    page).  One empty paragraph between two tables (it keeps them apart) or ending the
    document is not one.
``keep_with_next_at_foot``
    ``{"para", "style", "page", "next", "next_page"}``: a keep-with-next paragraph drawn
    last on its page, its next block on a later page.
``split_rows``
    ``{"table", "row", "pages"}``: a row (from 0) whose content is drawn on two pages.  No
    Word tool sets a row's "allow row to break across pages" yet: no fix.
``header_not_repeated``
    ``{"table", "pages", "header_rows"}``: a table on more than one page whose first row is
    not drawn again on each later page; ``fix`` marks one header row (``word_format_table``
    ``header_rows: 1``) where the table marks none, ``not_repeated_on`` the pages where a
    marked header is not drawn.
``past_margins``
    ``{"drawing" | "table", "page", "past": {"left"|"right"|"top"|"bottom": pt}}``: a
    picture or other drawing, or a table, drawn past the page's margins by more than half a
    point -- a table as drawn: its borders, shading and text (a borderless table whose text
    stays inside is not past them), beyond Word's own placement of it: a table is outdented
    by its first cell's left margin and half its outer border so that the cell's text
    lines up with the margin (and the same on the right), so only what reaches further
    counts -- an indent that pushes it out does.  ``fix`` for an inline picture that only runs past
    the right margin: resized to fit (``word_drawings`` ``resize``, keeping its aspect).
``uncertain``
    ``{"page", "placeholders", "approximations", "story_stops"}``: the pages the layout drew
    a placeholder on, laid out by a rule it has not measured, or could not draw a header,
    footer or note on -- a reviewer cannot judge those parts from the render.
    ``unknown_from``: the page the layout stopped at, when it was made to stop.

Each list holds at most :data:`CAP` entries, the most telling first; ``<fact>_left_out``
counts the rest.  ``pages`` (``check``'s) keeps the entries on those pages.
"""

from __future__ import annotations

import re
from fractions import Fraction
from typing import TYPE_CHECKING, Any, Iterable

from docx2svg.paths import resolve_path

from .edit import effective as _effective
from .edit import text as _text
from .layout import PX_PER_PT, _block_id, _in_body, _story_kind

if TYPE_CHECKING:  # pragma: no cover
    from .edit.document import Document
    from .layout import DocumentLayout

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

#: Entries per fact, most telling first (contrast keeps 12 per slide; a document's facts
#: span many pages).
CAP = 20

#: Past a margin by at most this, pt, is on it: the layout's device-pixel rounding.
MARGIN_TOLERANCE = 0.5

#: The run properties compared, in ``word_format``'s names, with the elements that
#: declare them (a run that declares none of them is not looked at).
_RUN_DECLARES = {
    "bold": ("b", "bCs"), "italic": ("i", "iCs"), "underline": ("u",), "strike": ("strike",),
    "size": ("sz", "szCs"), "font": ("rFonts",), "color": ("color",), "highlight": ("highlight",),
}
_PARAGRAPH_DECLARES = ("jc", "spacing", "ind", "keepNext", "pageBreakBefore")

#: A table's row as a line's or rule's path names it: the outermost table in the body (in
#: block-level w:sdt too) and its 1-based row.
_ROW = re.compile(r"^(w:body(?:/w:sdt\[\d+\]/w:sdtContent)*/w:tbl\[\d+\])/w:tr\[(\d+)\]")

#: What makes an empty-looking paragraph not empty: something drawn or a break.
_CONTENT = frozenset(_W + tag for tag in (
    "drawing", "pict", "object", "br", "cr", "sym", "tab", "ptab", "fldChar", "fldSimple",
    "footnoteReference", "endnoteReference", "sectPr", "framePr"))


def _pt(px) -> float:
    return round(float(Fraction(px) / PX_PER_PT), 2)


def layout_facts(layout: "DocumentLayout", *, pages: Iterable[int] | None = None,
                 min_size: float | None = None, styles: list[str] | None = None) -> dict[str, Any]:
    """The facts the module lists, of ``layout`` (and the document state it is of)."""
    return _Facts(layout, pages, min_size, styles).run()


class _Facts:
    def __init__(self, layout: "DocumentLayout", pages, min_size, styles) -> None:
        self.layout = layout
        self.document: "Document" = layout._document
        self.wanted = None if pages is None else set(pages)
        self.min_size = None if min_size is None else float(min_size)
        self.styles = styles
        document = self.document
        self.main = document.package.document_part()
        self.style_names = {s.id: s.name or s.id for s in document.styles}
        default = document.styles.default("paragraph")
        self.default_paragraph = default.id if default is not None else None
        default = document.styles.default("table")
        self.default_table = default.id if default is not None else None
        #: canon id -> [(1-based page, line)], every story's.
        self.placed: dict[str, list] = {}
        for number, lines in enumerate(layout._lines, start=1):
            for key, line in lines:
                if key is not None:
                    self.placed.setdefault(layout._canon(key), []).append((number, line))
        self.top_level = self._top_level()

    # -- helpers ---------------------------------------------------------------------------

    def pages_of(self, identifier: str) -> list[int]:
        return sorted({page for page, _ in self.placed.get(self.layout._canon(identifier), ())})

    def shown(self, pages: Iterable[int]) -> bool:
        return self.wanted is None or any(p in self.wanted for p in pages)

    def style_of(self, element) -> str | None:
        properties = element.find(_W + "pPr")
        node = properties.find(_W + "pStyle") if properties is not None else None
        style_id = node.get(_W + "val") if node is not None else self.default_paragraph
        return self.style_names.get(style_id, style_id)

    def _top_level(self) -> list[tuple[str, Any, str]]:
        """The body's blocks in order, block-level w:sdt opened: (kind, element, id)."""
        index = self.document._index(self.main)
        body = self.document.package.tree(self.main).find(_W + "body")
        out = []

        def visit(container) -> None:
            for child in container:
                if child.tag == _W + "sdt":
                    content = child.find(_W + "sdtContent")
                    if content is not None:
                        visit(content)
                elif child.tag in (_W + "p", _W + "tbl"):
                    entry = index.entry_for(child)
                    if entry is not None:
                        out.append(("p" if child.tag == _W + "p" else "t", child, entry.id))

        if body is not None:
            visit(body)
        return out

    def first_line(self, kind: str, element, identifier: str):
        """``(page, top px)`` where a top-level block is first drawn, or ``None``."""
        if kind == "p":
            found = self.placed.get(self.layout._canon(identifier))
            if not found:
                return None
            page = found[0][0]
            return page, min(line.top for p, line in found if p == page)
        best = None
        index = self.document._index(self.main)
        for paragraph in element.iter(_W + "p"):
            entry = index.entry_for(paragraph)
            for page, line in self.placed.get(self.layout._canon(entry.id), ()) if entry else ():
                if best is None or (page, line.top) < best:
                    best = (page, line.top)
        return best

    @staticmethod
    def capped(entries: list, name: str, out: dict) -> None:
        out[name] = entries[:CAP]
        if len(entries) > CAP:
            out[f"{name}_left_out"] = len(entries) - CAP

    # -- the facts -------------------------------------------------------------------------

    def run(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        self.text_sizes(out)
        self.capped(self.direct_formatting(), "direct_formatting", out)
        if self.styles is not None:
            self.style_list(out)
        self.capped(self.spacers(), "spacers", out)
        self.capped(self.keep_with_next(), "keep_with_next_at_foot", out)
        split, unrepeated = self.tables()
        self.capped(split, "split_rows", out)
        self.capped(unrepeated, "header_not_repeated", out)
        self.capped(self.past_margins(), "past_margins", out)
        self.uncertain(out)
        return out

    def text_sizes(self, out: dict) -> None:
        smallest: dict[str, float] = {}
        small: dict[tuple, dict] = {}
        for number, lines in enumerate(self.layout._lines, start=1):
            if self.wanted is not None and number not in self.wanted:
                continue
            for key, line in lines:
                kind = _kind(line)
                for span in line.spans:
                    if span.kind not in ("text", "label") or span.vertical_align in ("superscript", "subscript"):
                        continue
                    chars = sum(1 for c in span.chars if not c.isspace())
                    if not chars:
                        continue
                    size = span.half_points / 2
                    if size < smallest.get(kind, 10_000):
                        smallest[kind] = size
                    if self.min_size is not None and size < self.min_size and key is not None:
                        entry = small.setdefault((self.layout._canon(key), kind, size),
                                                 {"chars": 0, "pages": set()})
                        entry["chars"] += chars
                        entry["pages"].add(number)
        out["smallest_text"] = {kind: smallest[kind] for kind in sorted(smallest)}
        if self.min_size is None:
            return
        entries = []
        baselines: dict[str, float | None] = {}
        for (identifier, kind, size), seen in sorted(small.items(), key=lambda kv: (kv[0][2], -kv[1]["chars"])):
            if identifier not in baselines:
                baselines[identifier] = self._style_size(identifier)
            entry = {"para": identifier, "kind": kind, "size": size, "chars": seen["chars"],
                     "pages": sorted(seen["pages"])}
            style_size = baselines[identifier]
            if style_size is not None:
                entry["style_size"] = style_size
                if style_size >= self.min_size:
                    entry["fix"] = {"tool": "word_format", "calls": [{"targets": [identifier], "clear": ["size"]}]}
            entries.append(entry)
        self.capped(entries, "small_text", out)

    def _style_size(self, identifier: str) -> float | None:
        try:
            _, entry = self.document._resolve(identifier)
            return _effective.run(self.document, entry.element, None, direct=False).size
        except Exception:  # noqa: BLE001 -- a paragraph the resolver cannot read: no size
            return None

    def direct_formatting(self) -> list[dict]:
        document = self.document
        entries = []
        order = 0
        for part in document._parts():
            index = document._index(part)
            if index.story == "comments":
                continue
            for entry in index.paragraphs:
                order += 1
                found = self._direct(entry.element, entry.id)
                if found is None:
                    continue
                pages = self.pages_of(entry.id)
                if not self.shown(pages):
                    continue
                found["pages"] = pages
                entries.append((-len(found["props"]), order, found))
        entries.sort(key=lambda e: (e[0] == 0, e[0], e[1]))
        return [e[2] for e in entries]

    def _direct(self, element, identifier: str) -> dict | None:
        document = self.document
        props: dict[str, Any] = {}
        style_has: dict[str, Any] = {}
        properties = element.find(_W + "pPr")
        if properties is not None and any(properties.find(_W + tag) is not None for tag in _PARAGRAPH_DECLARES):
            mine = _effective.paragraph(document, element)
            styled = _effective.paragraph(document, element, direct=False)
            for name, value, base in _paragraph_values(mine, styled):
                if value != base:
                    props[name], style_has[name] = value, base
        walked = _text.walk(element)
        counts: dict[int, tuple[Any, str]] = {}
        for atom in walked.atoms:
            if atom.run is not None and atom.kind == "text":
                run, text = counts.get(id(atom.run), (atom.run, ""))
                counts[id(atom.run)] = (run, text + atom.char)
        total = sum(len(text) for _, text in counts.values())
        overridden: dict[str, dict] = {}
        for run, text in counts.values():
            rpr = run.find(_W + "rPr")
            if rpr is None or not text or not any(rpr.find(_W + tag) is not None
                                                  for tags in _RUN_DECLARES.values() for tag in tags):
                continue
            mine = _effective.run(document, element, run, text)
            styled = _effective.run(document, element, run, text, direct=False)
            for name, tags in _RUN_DECLARES.items():
                if not any(rpr.find(_W + tag) is not None for tag in tags):
                    continue
                value, base = getattr(mine, name), getattr(styled, name)
                if value != base:
                    seen = overridden.setdefault(name, {"chars": 0, "values": {}, "base": base})
                    seen["chars"] += len(text)
                    seen["values"][value] = seen["values"].get(value, 0) + len(text)
        in_part = []
        for name, seen in overridden.items():
            if seen["chars"] >= total:
                props[name] = max(seen["values"], key=seen["values"].get)
                style_has[name] = seen["base"]
            else:
                in_part.append(name)
        if not props and not in_part:
            return None
        found: dict[str, Any] = {"para": identifier, "style": self.style_of(element), "props": props,
                                 "style_has": style_has}
        if in_part:
            found["in_part"] = in_part
        if props:
            found["fix"] = {"tool": "word_format", "calls": [{"targets": [identifier], "clear": list(props)}]}
        return found

    def style_list(self, out: dict) -> None:
        document = self.document
        allowed: set[str] = set()
        unknown = []
        for name in self.styles:
            style = document.styles.find(name)
            if style is None:
                unknown.append(name)
            else:
                allowed.add(style.id)
        groups: dict[tuple[str, str], dict] = {}
        for part in document._parts():
            index = document._index(part)
            if index.story == "comments":
                continue
            for entry in index.paragraphs:
                properties = entry.element.find(_W + "pPr")
                node = properties.find(_W + "pStyle") if properties is not None else None
                style_id = node.get(_W + "val") if node is not None else self.default_paragraph
                if style_id is not None and style_id not in allowed:
                    self._count(groups, "paragraph", style_id, entry.id)
            for table in index.tables:
                node = table.element.find(f"{_W}tblPr/{_W}tblStyle")
                style_id = node.get(_W + "val") if node is not None else self.default_table
                if style_id is not None and style_id not in allowed:
                    self._count(groups, "table", style_id, table.id)
        entries = []
        for (kind, style_id), seen in groups.items():
            pages = sorted({p for b in seen["blocks"] for p in self.pages_of(b)}
                           if kind == "paragraph" else {p for b in seen["blocks"] for p in self._table_pages(b)})
            if not self.shown(pages) and pages:
                continue
            entries.append({"style": self.style_names.get(style_id, style_id), "kind": kind,
                            "count": len(seen["blocks"]), "blocks": seen["blocks"][:10], "pages": pages[:20]})
        entries.sort(key=lambda e: -e["count"])
        self.capped(entries, "styles_not_allowed", out)
        out["styles_unknown"] = unknown

    @staticmethod
    def _count(groups: dict, kind: str, style_id: str, identifier: str) -> None:
        groups.setdefault((kind, style_id), {"blocks": []})["blocks"].append(identifier)

    def _table_pages(self, identifier: str) -> list[int]:
        try:
            _, entry = self.document._resolve(identifier)
        except KeyError:
            return []
        index = self.document._index(self.main)
        pages = set()
        for paragraph in entry.element.iter(_W + "p"):
            found = index.entry_for(paragraph)
            if found is not None:
                pages.update(self.pages_of(found.id))
        return sorted(pages)

    def spacers(self) -> list[dict]:
        blocks = self.top_level
        entries = []
        k = 0
        while k < len(blocks):
            if not self._empty(blocks[k]):
                k += 1
                continue
            start = k
            while k < len(blocks) and self._empty(blocks[k]):
                k += 1
            run = blocks[start:k]
            before = blocks[start - 1] if start else None
            after = blocks[k] if k < len(blocks) else None
            if after is None:
                run = run[:-1]                    # the paragraph the document ends with
            elif len(run) == 1 and before is not None and before[0] == "t" and after[0] == "t":
                continue                          # keeps two tables apart
            if not run:
                continue
            entry = self._spacer(run, before, after)
            if entry is not None and self.shown([entry["page"]]):
                entries.append(entry)
        entries.sort(key=lambda e: -e["height_pt"])
        return entries

    def _empty(self, block) -> bool:
        kind, element, identifier = block
        if kind != "p":
            return False
        if any(node.tag in _CONTENT for node in element.iter() if isinstance(node.tag, str)):
            return False
        found = self.placed.get(self.layout._canon(identifier))
        if not found:
            return False                          # not drawn (hidden, or past a stop)
        return not any(not c.isspace() for _, line in found for span in line.spans for c in span.chars)

    def _spacer(self, run: list, before, after) -> dict | None:
        """One run of spacers.  Between two paragraphs the space before and the space after
        collapse to the larger (docx2svg's measured rule, :func:`docx2svg.vertical.paragraph_gap_px`),
        so the gap the run makes is measured from the previous block's text to the next
        block's: without the run, the next paragraph keeps its place with a space before of
        that gap."""
        first = self.first_line(*run[0])
        if first is None:
            return None
        page, top = first
        ids = [identifier for _, _, identifier in run]
        entry: dict[str, Any] = {"paras": ids, "page": page}
        space = self._space
        # The gap above the first spacer: its space before against the space after of the
        # paragraph before it (none after a table or at a page's top).
        previous_after = space(before, "space_after") if before is not None else 0.0
        above = max(previous_after, space(run[0], "space_before"))
        following = self.first_line(*after) if after is not None else None
        target = ids[0] if len(ids) == 1 else f"{ids[0]}..{ids[-1]}"
        entry["fix"] = {"tool": "word_delete", "calls": [{"target": target}]}
        if following is not None and following[0] == page:
            gap = round(above + _pt(following[1] - top), 1)
            entry["height_pt"] = round(gap - max(previous_after, space(after, "space_before")), 1)
            entry["next"] = after[2]
            if after[0] == "p" and gap <= 1584:
                entry["keep_space"] = {"tool": "word_format",
                                       "calls": [{"targets": [after[2]], "space_before": gap}]}
        else:
            bottom = max(line.top + line.pitch for _, _, i in run
                         for p, line in self.placed.get(self.layout._canon(i), ()) if p == page)
            entry["height_pt"] = round(above + _pt(bottom - top), 1)
            entry["page_end"] = True
            if after is not None:
                entry["next"] = after[2]
        return entry

    def _space(self, block, side: str) -> float:
        """A top-level paragraph's effective space before or after, pt (0 for a table)."""
        if block[0] != "p":
            return 0.0
        return getattr(_effective.paragraph(self.document, block[1]), side)

    def keep_with_next(self) -> list[dict]:
        position = {self.layout._canon(identifier): k for k, (_, _, identifier) in enumerate(self.top_level)}
        entries = []
        for number, lines in enumerate(self.layout._lines, start=1):
            if not self.shown([number]):
                continue
            body = [(line.top, key) for key, line in lines if key is not None and _in_body(line)]
            if not body:
                continue
            last = self.layout._canon(max(body, key=lambda b: b[0])[1])
            k = position.get(last)
            if k is None or k + 1 >= len(self.top_level) or self.top_level[k][0] != "p":
                continue
            if max(self.pages_of(last)) != number:
                continue                          # it goes on over the page
            following = self.first_line(*self.top_level[k + 1])
            if following is None or following[0] <= number:
                continue
            element = self.top_level[k][1]
            if not _effective.paragraph(self.document, element).keep_with_next:
                continue
            entries.append({"para": self.top_level[k][2], "style": self.style_of(element), "page": number,
                            "next": self.top_level[k + 1][2], "next_page": following[0]})
        return entries

    def tables(self) -> tuple[list[dict], list[dict]]:
        rows: dict[str, dict[int, dict[int, set]]] = {}
        table_pages: dict[str, set] = {}
        for number, page in enumerate(self.layout.pages, start=1):
            for line in page.lines:
                match = _ROW.match(line.path) if line.story is None else None
                if match:
                    table, row = match.group(1), int(match.group(2)) - 1
                    rows.setdefault(table, {}).setdefault(row, {}).setdefault(number, set()).add(
                        (line.path, line.number))
                    table_pages.setdefault(table, set()).add(number)
            for rule in page.rules:
                match = _ROW.match(rule.path or "")
                if match:
                    table_pages.setdefault(match.group(1), set()).add(number)
        split, unrepeated = [], []
        tree = self.document.package.tree(self.main)
        for table, pages in table_pages.items():
            if len(pages) < 2:
                continue
            identifier = _block_id(self.document, self.main, table) or table
            element = resolve_path(tree, table)
            header_rows = _header_rows(element) if element is not None else 0
            ordered = sorted(pages)
            for row, drawn in sorted(rows.get(table, {}).items()):
                if len(drawn) < 2:
                    continue
                repeated = row < header_rows and len({frozenset(v) for v in drawn.values()}) == 1
                if not repeated and self.shown(drawn):
                    split.append({"table": identifier, "row": row, "pages": sorted(drawn)})
            if not self.shown(ordered):
                continue
            head = rows.get(table, {}).get(0, {})
            missing = [p for p in ordered[1:] if p not in head]
            if header_rows == 0:
                unrepeated.append({"table": identifier, "pages": ordered, "header_rows": 0,
                                   "fix": {"tool": "word_format_table",
                                           "calls": [{"table": identifier, "scope": "table", "header_rows": 1}]}})
            elif missing:
                unrepeated.append({"table": identifier, "pages": ordered, "header_rows": header_rows,
                                   "not_repeated_on": missing})
        return split, unrepeated

    def past_margins(self) -> list[dict]:
        document = self.document
        tree = document.package.tree(self.main)
        frames = [(identifier, frame) for part, identifier, frame in document._drawings() if part == self.main]
        by_frame = {id(frame): identifier for identifier, frame in frames}
        entries = []
        for number, page in enumerate(self.layout.pages, start=1):
            if not self.shown([number]):
                continue
            left, right, top, bottom = page.column
            boxes = []
            for picture in page.pictures:
                if picture.part is None and picture.path.startswith("w:body"):
                    boxes.append((picture, True, False))
            for placed in page.floats:
                if placed.story is None and placed.path.startswith("w:body"):
                    boxes.append((placed, False, True))
            for box, inline, vertical in boxes:
                past = _past(box.x, box.x + box.width, box.y, box.y + box.height, page.column, vertical)
                if not past:
                    continue
                element = resolve_path(tree, box.path)
                identifier = _drawing_id(element, by_frame) or box.path
                entry = {"drawing": identifier, "page": number, "past": past}
                if inline and list(past) == ["right"] and identifier.startswith("d:") and box.x < right:
                    width = _pt(right - box.x)
                    entry["fix"] = {"tool": "word_drawings",
                                    "calls": [{"action": "resize", "target": identifier, "width": width}]}
                entries.append(entry)
            extents: dict[str, list] = {}
            for rule in page.rules:
                match = _ROW.match(rule.path or "")
                if match:
                    extent = extents.setdefault(match.group(1), [rule.x, rule.x + rule.width])
                    extent[0], extent[1] = min(extent[0], rule.x), max(extent[1], rule.x + rule.width)
            for line in page.lines:
                match = _ROW.match(line.path) if line.story is None else None
                if match and line.spans:
                    low = min(span.xs[0] for span in line.spans if span.xs)
                    high = max(span.end for span in line.spans)
                    extent = extents.setdefault(match.group(1), [low, high])
                    extent[0], extent[1] = min(extent[0], low), max(extent[1], high)
            for table, (low, high) in extents.items():
                # Word's own placement draws a table's first cell's text on the margin, the
                # table outdented by that cell's left margin and its outer border's outer
                # half (and the last cell's text ending on the right one): only past that.
                left_allow, right_allow = self._table_allowance(table)
                past = _past(low + left_allow, high - right_allow, None, None, page.column, False)
                if past:
                    entries.append({"table": _block_id(document, self.main, table) or table, "page": number,
                                    "past": past})
        entries.sort(key=lambda e: -max(e["past"].values()))
        return entries

    def _table_allowance(self, path: str) -> tuple:
        """``(left, right)`` device px a table may reach past the margins as Word places
        it by default: the first row's first cell's left margin plus its left outer
        border's outer half, and the last cell's right margin plus its right one's
        (docx2svg's resolved table, :func:`docx2svg.table.resolve_table`)."""
        allowances = getattr(self, "_allowances", None)
        if allowances is None:
            allowances = self._allowances = {}
            try:
                from docx2svg import model as _model
                from docx2svg.table import outer_part, resolve_table

                state = _effective.model(self.document).model
                for block in state.body or ():
                    if isinstance(block, _model.Table) and block.path:
                        resolved = resolve_table(state, block)
                        cells = resolved.rows[0].cells if resolved.rows and resolved.rows[0].cells else None
                        if not cells:
                            continue
                        first, last = cells[0], cells[-1]
                        left = first.margins.get("left", 0) + outer_part(first.left_edge[0])
                        right = last.margins.get("right", 0) + outer_part(last.right_edge[0])
                        allowances[block.path] = (Fraction(left, 20) * PX_PER_PT, Fraction(right, 20) * PX_PER_PT)
            except Exception:  # noqa: BLE001 -- a table the resolver cannot read: no allowance
                pass
        return allowances.get(path, (0, 0))

    def uncertain(self, out: dict) -> None:
        coverage = self.layout.coverage
        pages: dict[int, dict] = {}

        def add(page, kind: str, value: dict) -> None:
            if page is None or not self.shown([page]):
                return
            pages.setdefault(page, {"page": page}).setdefault(kind, []).append(value)

        if coverage is not None:
            ids = self.layout._placeholder_ids
            for entry in coverage.placeholders or ():
                add(entry.page, "placeholders", {"reason": entry.reason, "at": ids.get(entry.path) or entry.path})
            for entry in coverage.approximations or ():
                at = (_block_id(self.document, self.main, entry.path) if entry.path and entry.path.startswith("w:body")
                      else None) or entry.path
                add(entry.page, "approximations", {"reason": entry.reason, "at": at})
            for entry in coverage.story_stops or ():
                add(entry.page, "story_stops", {"reason": entry.reason})
        entries = [pages[p] for p in sorted(pages)]
        self.capped(entries, "uncertain", out)
        if self.layout.stopped is not None:
            out["unknown_from"] = self.layout.stopped.page


def _kind(line) -> str:
    story = line.story or _story_kind(line.path)
    if story == "body":
        return "table" if "/w:tc[" in line.path else "body"
    return {"footnotes": "footnote", "endnotes": "endnote"}.get(story, story)


def _paragraph_values(mine, styled):
    """``(word_format's name, direct, styles')`` for each paragraph property compared."""
    def first(effective):
        return round(effective.first_line - effective.hanging, 2)

    def line(effective):
        if effective.line_spacing is None:
            return None
        return (round(effective.line_spacing, 2) if effective.line_rule == "auto"
                else [effective.line_rule, round(effective.line_spacing, 2)])

    yield "alignment", mine.alignment, styled.alignment
    yield "space_before", mine.space_before, styled.space_before
    yield "space_after", mine.space_after, styled.space_after
    yield "line_spacing", line(mine), line(styled)
    yield "indent_left", mine.indent_left, styled.indent_left
    yield "indent_right", mine.indent_right, styled.indent_right
    yield "indent_first", first(mine), first(styled)
    yield "keep_with_next", mine.keep_with_next, styled.keep_with_next
    yield "page_break_before", mine.page_break_before, styled.page_break_before


def _header_rows(table) -> int:
    count = 0
    for row in table.findall(_W + "tr"):
        node = row.find(f"{_W}trPr/{_W}tblHeader")
        if node is None or (node.get(_W + "val") or "1").lower() in ("0", "false", "off"):
            break
        count += 1
    return count


def _past(x0, x1, y0, y1, column, vertical: bool) -> dict[str, float]:
    left, right, top, bottom = column
    past = {}
    for side, amount in (("left", left - x0), ("right", x1 - right)):
        if _pt(amount) > MARGIN_TOLERANCE:
            past[side] = _pt(amount)
    if vertical:
        for side, amount in (("top", top - y0), ("bottom", y1 - bottom)):
            if _pt(amount) > MARGIN_TOLERANCE:
                past[side] = _pt(amount)
    return past


def _drawing_id(element, by_frame: dict) -> str | None:
    if element is None:
        return None
    for node in element.iter(_WP + "inline", _WP + "anchor"):
        found = by_frame.get(id(node))
        if found is not None:
            return found
    return None


__all__ = ["CAP", "MARGIN_TOLERANCE", "layout_facts"]
