"""A copy's list items and page settings compared with its reference's, as Word applies them
(round 42, R49).

:mod:`.compare` reads a paragraph's direct formatting as its ``w:pPr`` states it, so two
list items that Word draws differently -- one in a list of dashes, one of round bullets; one
a level deeper; one that lost the "keep with next" its style gave it -- can compare the same,
and two that Word draws alike (a keep stated directly and by the style) can differ.  For the
copies :func:`docx_agent.edit.copies.compare_copy` compares, this module reads each pair of
paragraphs (:func:`.compare.paragraph_pairs`: a paragraph block and its reference, a table's
cell paragraphs) as Word applies them -- docDefaults, the table style, the paragraph style,
the list level, direct formatting (:mod:`.effective`, docx2svg's measured cascade) -- and
reports where they differ:

* ``list_level`` -- the item's list level (0-8), or ``None`` for a paragraph in no list (or
  one whose level draws no label: Word shows none);
* ``list_format`` -- the label its level draws: ``bullet`` and the character (a symbol
  face's character read as the Unicode character it shows: Symbol's U+F0B7 is "•"), or the
  number format and the level's text (``decimal %1.``);
* ``list_numbering`` -- for a numbered item: ``continues`` (one more than the item before it
  at its level, in the same list, since the last item at a shallower level -- in the copy,
  or in the reference) or ``starts at N`` (any other number: a first item, a restart, a list
  that continues another before the copy).  Counted as Word counts (docx2svg's
  :class:`~docx2svg.linebreak.ListCounters`, over the whole body), so a copy whose list
  goes on from the chapter before it is told.  Compared (round 43) only where the sample's
  own count begins inside the sample, or goes on there from an item of the sample: then
  each copy's begins and goes on alike.  Where the sample's item is numbered by the
  document's outline (:func:`_by_outline`: a heading whose number is its place among the
  headings) or its count runs on from before the sample (by design), its level and label
  are compared and its number never -- nor where the copy's item is numbered by the
  outline: no restart is offered there.  Decided from how the sample's count behaves
  (:data:`COUNT`), never from a style's name.  Where the caller states that the copies'
  lists run on (round 43, :class:`Run`: never inferred), a later copy's item where the
  sample's count begins goes on from the copy before it instead (:data:`RUNS_ON`);
* ``indent_left`` and ``indent_first`` (points; ``indent_first`` negative for a hanging
  indent): the paragraph's own, its style's or its list level's;
* ``keep_with_next``, ``keep_lines``, ``page_break_before``, ``widow_control``: true or false.

**Two items that look the same do not differ**: values, not the XML that states them (a
keep stated directly and by the style, an indent stated as the list gives it, a restarted
list of bullets, another list that draws the same label at the same place).  One
difference per cause: a pair whose paragraph styles differ is not compared here (the style
difference and its fix come first; a check after it compares the rest); a pair whose list
level differs (one in no list among them) is told by ``list_level`` alone, and one whose
label differs by ``list_format`` alone -- the page settings apart, which neither changes.
These differences replace :mod:`.compare`'s ``indent_left``, ``indent_first``,
``keep_with_next``, ``keep_lines`` and ``page_break_before`` for the paragraphs compared here.

**Fixes** (:func:`settings_fix`, the one place they are written).  A reference in the
same document: ``word_format`` ``like`` the reference's paragraph, one call per reference
paragraph, wherever taking its look changes nothing else compared here (:func:`_like_takes`:
not for where a numbered item's numbering starts, nor to move a numbered item into the
reference's list, nor a page setting of a paragraph whose list differs).  Otherwise -- and
always for a reference in another document -- the explicit calls: ``word_format`` with the
reference's value, or ``clear`` where the paragraph's style (and list) give it the
reference's value without its own; ``word_lists`` ``level``, ``add`` (a paragraph in no list:
a list of Word's own of the reference's kind, its label compared by the next check),
``remove``, ``format`` (the level's ``number_format`` and ``label`` -- a symbol face's
bullet as its Unicode character -- and, where the item's indents differ, ``text_at`` and
``number_at`` at the reference's: the copy's list only, Word's copy-on-write), ``restart``
(``start``) and ``continue`` (where an earlier list over the same definition is there to
continue; in a stated run, with ``join``: the previous copy's list, and no restart inside
a running list).  ``widow_control`` has no ``word_format`` field: ``like`` gives it, else
``word_follow_sample`` (the check's ``fix``).  None where no call makes it.
"""

from __future__ import annotations

from collections import OrderedDict
from typing import TYPE_CHECKING, Any

from lxml import etree

from . import effective as _effective

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_P, _TBL = _W + "p", _W + "tbl"

#: Every setting compared, in the order reported.
SETTINGS = ("list_level", "list_format", "list_numbering", "indent_left", "indent_first",
            "keep_with_next", "keep_lines", "page_break_before", "widow_control")
#: The list's settings: ``word_lists`` makes them; ``word_follow_sample`` keeps a copy's list.
LIST_SETTINGS = frozenset(SETTINGS[:3])
#: Round 43: per list item of a region (:func:`region_settings`), how its number comes --
#: ``outline`` (numbered by the outline: :func:`_by_outline`), ``own`` (its count begins
#: inside the region, or it continues an item of the region) or ``runs_on`` (its count
#: goes on from an item before the region): what :func:`_counted` decides by.  Not a
#: setting, never reported.
COUNT = "list_count"
#: Round 43 (run on): per list item, its list key (the instances counting together) and
#: whether it leads its count in the region (no item of its list at its level or a
#: shallower one before it there).  Not settings, never reported.
LIST_KEY, LEAD = "list_key", "list_lead"
#: Round 43: ``list_numbering``'s value where the caller states the copies' lists run on
#: (:class:`Run`) and a copy's item goes on from the copy before it, in the same list.
RUNS_ON = "runs on"
#: :mod:`.compare`'s properties these replace for the paragraphs compared here.
REPLACED = frozenset({"indent_left", "indent_first", "keep_with_next", "keep_lines", "page_break_before"})
#: The settings ``word_format`` sets by the same name.
_FORMAT_FIELDS = frozenset({"indent_left", "indent_first", "keep_with_next", "keep_lines", "page_break_before"})
_PAGE = ("keep_with_next", "keep_lines", "page_break_before", "widow_control")


# -- list items counted as Word counts them -----------------------------------------------------


def list_items(document: "Document") -> dict[int, tuple]:
    """Every body paragraph that draws a list label, by the ``id()`` of its element (held
    by the reader, :func:`reader`): ``(list key, level, format, text, number)`` -- the list
    key the instances counting together share, the level, its ``w:numFmt`` and
    ``w:lvlText`` (a symbol face's characters as Unicode) and the number it shows.  Counted
    in document order over the body, table cells included, as the layout counts (docx2svg's
    ``ListCounters``); kept for the document's state."""
    return reader(document).items


def reader(document: "Document") -> "_Reader":
    """What this module reads of the document's paragraphs, for its state: the light
    reader (:class:`_Light`) where it reads the document as docx2svg would, else docx2svg's
    whole parse (:class:`_Whole`)."""
    kept = getattr(document, "_tool_settings_reader", None)
    if kept is not None and kept[0] == document._state:
        return kept[1]
    found = _Light.of(document) or _Whole(document)
    document._tool_settings_reader = (document._state, found)
    return found


class _Reader:
    """The list items (:attr:`items`, by element ``id()``) and each paragraph's formatting
    as Word applies it (:meth:`applied`)."""

    items: dict[int, tuple]

    def applied(self, element: Any) -> "_effective.EffectiveParagraph":
        raise NotImplementedError


def _count(model: Any, paragraphs: "list[tuple[Any, Any]]") -> dict[int, tuple]:
    """``items`` of ``(element, docx2svg paragraph)`` in document order, counted."""
    from docx2svg.linebreak import ListCounters
    from docx2svg.resolve import resolve_paragraph

    counters = ListCounters()
    out: dict[int, tuple] = {}
    for element, paragraph in paragraphs:
        resolved = resolve_paragraph(model, paragraph)
        num_id = resolved.get("numPr.numId")
        if not num_id:
            continue
        level = resolved.get("numPr.ilvl", 0) or 0
        key = model.numbering_lists.get(num_id, ("num", num_id))
        # Round 43: whether the item begins its count -- no item before it at its level
        # since the count last began (a shallower item, the document's start), or a restart
        # (``w:startOverride``) its instance makes here -- read before it is counted.
        fresh = (level not in counters.counts.get(key, {})
                 or ((num_id, level) not in counters.used and (num_id, level) in model.numbering_restarts))
        if counters.label(model, paragraph, resolved) is None:
            continue
        found = model.numbering.get(num_id, {}).get(level)
        number = counters.counts.get(key, {}).get(level)
        out[id(element)] = (key, level, found.format if found is not None else None, _shown(found), number,
                            fresh, _by_outline(resolved, paragraph, found))
    return out


def _by_outline(resolved: Any, paragraph: Any, level: Any) -> bool:
    """Whether an item is numbered by the document's outline: the paragraph stands at an
    outline level (``w:outlineLvl`` 0-8, its own or its style's) and its numbering is its
    style's (no ``w:numId`` of its own) or a level tied to a paragraph style (``w:pStyle``
    in the level: Word's multilevel list linked to the heading styles).  Read from the
    structure, never the style's name: its number is its place in the outline."""
    outline = resolved.get("outlineLvl")
    if not isinstance(outline, int) or not 0 <= outline <= 8:
        return False
    own = getattr(getattr(paragraph, "properties", None), "numbering_id", None)
    return not own or getattr(level, "style", None) is not None


class _Whole(_Reader):
    """From docx2svg's parse of the whole package (:func:`.effective.model`): its
    paragraphs mapped to our elements (held by the model's map)."""

    def __init__(self, document: "Document") -> None:
        from docx2svg.model import Table

        self.document = document
        state = _effective.model(document)
        elements = state.elements
        ordered: list[tuple[Any, Any]] = []

        def walk(item: Any) -> None:
            if isinstance(item, Table):
                for row in item.rows:
                    for cell in row:
                        for inner in cell:
                            walk(inner)
                return
            element = elements.get(id(item))
            ordered.append((element if element is not None else item, item))

        for block in state.model.body:
            walk(block)
        self.items = _count(state.model, ordered)

    def applied(self, element: Any) -> "_effective.EffectiveParagraph":
        return _effective.paragraph(self.document, element)


class _Light(_Reader):
    """The paragraphs read from the document's own XML, as docx2svg's parse reads them --
    each paragraph's properties (docx2svg's own reader), its table's style and the cell's
    conditional formats (docx2svg's ``cell_conditions``) -- resolved against the styles,
    numbering and settings alone (:func:`_frame`, kept while those parts are the same):
    no parse of the body, which a check of one copy's blocks paid for every version.
    The same values as docx2svg's whole parse (``tests/test_copy_settings.py`` compares
    them, paragraph for paragraph); only where the final view is the XML's own -- no
    paragraph mark or table row deleted or moved away (docx2svg joins those), else
    :class:`_Whole`."""

    @classmethod
    def of(cls, document: "Document") -> "_Light | None":
        try:
            from docx2svg.parse import document as parse
            from docx2svg.resolve import resolve_paragraph  # noqa: F401
            properties_of, conditions_of, look_of = (parse._paragraph_properties, parse.cell_conditions,
                                                     parse._table_look)
        except (ImportError, AttributeError):     # another docx2svg: the whole parse
            return None
        body = document.package.tree(document.package.document_part()).find(_W + "body")
        if body is None or skeleton(document).removed:
            return None
        frame = _frame(document)
        if frame is None:
            return None
        return cls(document, body, frame, properties_of, conditions_of, look_of)

    def __init__(self, document: "Document", body: Any, frame: Any, properties_of: Any, conditions_of: Any,
                 look_of: Any) -> None:
        from docx2svg.paths import row_level

        self.document, self.frame = document, frame
        self._properties_of = properties_of
        self._context: dict[int, tuple] = {}     # id(element) -> (table style, conditions)
        self._held: list = []                    # our elements, so their id()s stay theirs
        order: list = []

        def blocks(parent: Any, style: "str | None", conditions: tuple) -> None:
            for node in parent:
                tag = node.tag
                if tag == _P:
                    self._held.append(node)
                    self._context[id(node)] = (style, conditions)
                    order.append(node)
                elif tag == _TBL:
                    table(node)
                elif tag == _W + "sdt":
                    content = node.find(_W + "sdtContent")
                    if content is not None:
                        blocks(content, style, conditions)
                elif tag == _W + "customXml":
                    blocks(node, style, conditions)

        def table(element: Any) -> None:
            properties = element.find(_W + "tblPr")
            named = properties.find(_W + "tblStyle") if properties is not None else None
            style = (named.get(_W + "val") if named is not None else None) or ""
            # The look read as the cells' conditions are asked for (:meth:`_paragraph`).
            look: list = []
            tblLook = properties.find(_W + "tblLook") if properties is not None else None
            rows = list(row_level(element, "tr"))
            for k, row in enumerate(rows):
                cells = list(row_level(row, "tc"))
                for j, cell in enumerate(cells):
                    blocks(cell, style, (look, tblLook, k, j, len(rows), len(cells)))

        self._conditions_of, self._look_of = conditions_of, look_of
        blocks(body, None, ())
        listed = skeleton(document).listed
        self._models = _model_cache(document)
        self.items = _count(frame, [(e, self._paragraph(e)) for e in order if id(e) in listed])

    def _paragraph(self, element: Any) -> Any:
        """docx2svg's paragraph of ``element``, as its parse builds it -- kept by the bytes
        of its ``w:pPr``, its table's style and its cell's conditions (:func:`_model_cache`)."""
        from docx2svg import model

        style, cell = self._context.get(id(element), (_effective._table_style(element), ()))
        conditions: tuple = ()
        if cell:
            look, tblLook, k, j, rows, cells = cell
            if not look:
                look.append(self._look_of(tblLook))
            conditions = self._conditions_of(look[0], k, j, rows, cells)
        properties = element.find(_W + "pPr")
        key = (etree.tostring(properties) if properties is not None else b"", style, conditions)
        found = self._models.get(key)
        if found is None:
            found = self._models[key] = model.Paragraph(properties=self._properties_of(properties),
                                                        table_style_id=style, table_conditions=conditions)
        return found

    def applied(self, element: Any) -> "_effective.EffectiveParagraph":
        if id(element) not in self._context:
            return _effective.paragraph(self.document, element)   # not the body's own
        return _effective._paragraph(_FrameState(self.frame), self._paragraph(element))


class _FrameState:
    """What :func:`.effective._paragraph` reads of its state: the model."""

    __slots__ = ("model",)

    def __init__(self, model: Any) -> None:
        self.model = model


#: docx2svg paragraphs kept per document (:meth:`_Light._paragraph`), past which the oldest go.
MODELS_KEPT = 4096


def _model_cache(document: "Document") -> "OrderedDict":
    """The document's kept docx2svg paragraphs, bounded (:data:`MODELS_KEPT`): a paragraph
    is a value of its properties (frozen), so one built is the same for any state."""
    found = getattr(document, "_tool_paragraph_models", None)
    if found is None:
        found = document._tool_paragraph_models = OrderedDict()
    while len(found) > MODELS_KEPT:
        found.popitem(last=False)
    return found


class _Skeleton:
    """The body's paragraphs in document order (held, so their ``id()`` stays theirs), per
    paragraph (``upto``) how many of those up to it may be list items, the digest of the
    first ``n`` of those per ``n`` (``digests``), and whether a paragraph mark or a row is
    deleted or moved away (``removed``: docx2svg's final view joins them)."""

    __slots__ = ("elements", "upto", "digests", "removed", "listed")

    def __init__(self, elements: list, upto: dict, digests: list, removed: bool, listed: set) -> None:
        self.elements, self.upto, self.digests, self.removed = elements, upto, digests, removed
        #: The ``id()`` of each paragraph that may be a list item (:func:`_maybe_listed`).
        self.listed = listed


def skeleton(document: "Document") -> _Skeleton:
    """The body's list skeleton for the document's state (what
    :func:`docx_agent.edit.copies.list_context` keys a comparison by): for each paragraph
    that may be a list item (:func:`_maybe_listed`) its style, its own ``w:numPr`` and
    whether its mark is deleted -- from the XML alone."""
    import hashlib

    kept = getattr(document, "_tool_list_skeleton", None)
    if kept is not None and kept[0] == document._state:
        return kept[1]
    body = document.package.tree(document.package.document_part()).find(_W + "body")
    maybe = _maybe_listed(document)
    hasher = hashlib.blake2b(digest_size=16)
    elements = list(body.iter(_P)) if body is not None else []
    upto: dict[int, int] = {}
    digests = [hasher.hexdigest()]
    removed = body is not None and body.find(f".//{_W}trPr/{_W}del") is not None
    listed: set = set()
    for element in elements:
        mark = element.find(f"{_W}pPr/{_W}rPr")
        deleted = mark is not None and mark.find(_W + "del") is not None
        if mark is not None and (deleted or mark.find(_W + "moveFrom") is not None):
            removed = True
        if maybe(element):
            listed.add(id(element))
            own = element.find(f"{_W}pPr/{_W}numPr")
            named = element.find(f"{_W}pPr/{_W}pStyle")
            hasher.update(((named.get(_W + "val") or "") if named is not None else "").encode() + b"\0")
            hasher.update(etree.tostring(own) if own is not None else b"-")
            hasher.update(b"d" if deleted else b"k")
            digests.append(hasher.hexdigest())
        upto[id(element)] = len(digests) - 1
    found = _Skeleton(elements, upto, digests, removed, listed)
    document._tool_list_skeleton = (document._state, found)
    return found


def _maybe_listed(document: "Document") -> Any:
    """``element -> bool``: whether a paragraph may be in a list -- numbering of its own, a
    paragraph style whose chain states numbering, or (where docDefaults or a table style
    states any) every paragraph.  The others resolve to none, so are not resolved."""
    styles = document.package.tree(document.package.styles_part()) if document.package.styles_part() else None
    if styles is None:
        return lambda element: element.find(f"{_W}pPr/{_W}numPr") is not None
    numbered: dict = {}
    by_id = {node.get(_W + "styleId"): node for node in styles.iter(_W + "style")}
    if styles.find(f"{_W}docDefaults/{_W}pPrDefault/{_W}pPr/{_W}numPr") is not None or any(
            node.get(_W + "type") == "table" and node.find(f".//{_W}numPr") is not None for node in by_id.values()):
        return lambda element: True
    default = next((i for i, node in by_id.items() if node.get(_W + "type") == "paragraph"
                    and node.get(_W + "default") in ("1", "true", "on")), None)

    def chain(style: "str | None") -> bool:
        if style not in numbered:
            seen, current, found = set(), style, False
            while current and current not in seen and current in by_id:
                seen.add(current)
                node = by_id[current]
                if node.find(f"{_W}pPr/{_W}numPr") is not None:
                    found = True
                    break
                based = node.find(_W + "basedOn")
                current = based.get(_W + "val") if based is not None else None
            numbered[style] = found
        return numbered[style]

    def maybe(element: Any) -> bool:
        if element.find(f"{_W}pPr/{_W}numPr") is not None:
            return True
        named = element.find(f"{_W}pPr/{_W}pStyle")
        return chain(named.get(_W + "val") if named is not None else default)

    return maybe


#: A main part with an empty body: :func:`_frame` parses the styles with it.
_EMPTY = (b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          b"<w:body/></w:document>")


def _frame(document: "Document") -> Any:
    """docx2svg's document of the styles, numbering and settings alone (as
    :func:`docx2svg.parse.document.parse_package` reads them), kept on the document while
    those parts' bytes are the same; ``None`` where docx2svg cannot be asked so."""
    import hashlib

    from lxml import etree

    package = document.package
    main = package.document_part()
    named = {"styles": package.styles_part(), "numbering": package.numbering_part()}
    settings = package.related_parts_of_type(
        main, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings")
    named["settings"] = settings[0] if settings else None
    data = {}
    hasher = hashlib.blake2b(digest_size=16)
    for name, part in named.items():
        tree = package.tree(part) if part is not None else None
        data[name] = etree.tostring(tree) if tree is not None else None
        hasher.update(name.encode() + b"\0" + (data[name] or b"-") + b"\1")
    key = hasher.digest()
    kept = getattr(document, "_tool_settings_frame", None)
    if kept is not None and kept[0] == key:
        return kept[1]
    try:
        from docx2svg.parse.document import parse_document, settings_read
        from docx2svg.parse.styles import parse_numbering
    except ImportError:
        return None
    frame = parse_document(_EMPTY, data["styles"])
    if data["numbering"] is not None:
        numbering = parse_numbering(data["numbering"], frame.styles)
        frame.numbering, frame.numbering_lists = numbering.levels, numbering.lists
        frame.numbering_restarts = numbering.restarts
    if data["settings"] is not None:
        for name, value in settings_read(data["settings"]).items():
            setattr(frame, name, value)
    document._tool_settings_frame = (key, frame)
    return frame


def _shown(level: Any) -> str | None:
    """A level's text as it shows: a symbol face's characters as the Unicode they draw."""
    if level is None or level.text is None:
        return None
    from ooxml_common.text.symbol_fonts import symbol_face, to_unicode

    font = level.run.get("rFonts.ascii") if level.run else None
    face = symbol_face(getattr(font, "face", None))
    if face is None:
        return level.text
    return "".join(to_unicode(face, char) or char for char in level.text)


# -- one region's settings -----------------------------------------------------------------------


def region_settings(document: "Document", paragraphs: list) -> dict[int, dict[str, Any]]:
    """Each of ``paragraphs`` (elements, in document order: a copy's or a reference's) by
    ``id()``: its settings (the module's names), ``list_numbering`` read along the region."""
    read = reader(document)
    items = read.items
    out: dict[int, dict[str, Any]] = {}
    seen: dict[int, tuple] = {}             # level -> (list key, number) of the item last there
    shallowest: dict[Any, int] = {}         # list key -> the shallowest level of its items so far
    for element in paragraphs:
        applied = read.applied(element)
        values: dict[str, Any] = {
            "indent_left": round(applied.indent_left, 2),
            "indent_first": round(applied.first_line - applied.hanging, 2),
            "keep_with_next": applied.keep_with_next, "keep_lines": applied.keep_together,
            "page_break_before": applied.page_break_before, "widow_control": applied.widow_control,
            "list_level": None, "list_format": None, "list_numbering": None, COUNT: None,
            LIST_KEY: None, LEAD: False}
        item = items.get(id(element))
        if item is not None:
            key, level, number_format, text, number, fresh, outline = item
            values["list_level"] = level
            values["list_format"] = f"{number_format or 'decimal'} {text}"
            values[LIST_KEY] = key
            values[LEAD] = shallowest.get(key, 9) > level
            shallowest[key] = min(shallowest.get(key, 9), level)
            if number_format not in ("bullet", "none") and text and "%" in text:
                before = seen.get(level)
                goes_on = before is not None and before[0] == key and number == before[1] + 1
                values["list_numbering"] = "continues" if goes_on else f"starts at {number}"
                values[COUNT] = ("outline" if outline else "own" if goes_on or fresh else "runs_on")
            seen[level] = (key, number)
            for deeper in [k for k in seen if k > level]:
                del seen[deeper]
        out[id(element)] = values
    return out


def compare_settings(document: "Document", reference_document: "Document", pairs: list[tuple],
                     ids: dict[int, tuple[str, str]], mine: dict[int, dict], theirs: dict[int, dict],
                     reference_ids: dict[int, str], run: "Run | None" = None) -> list[dict]:
    """The differences of ``pairs`` (``(paragraph, reference)`` elements, styles alike), each
    ``{"block", "property", "paragraphs", "value", "reference", "fix"}`` grouped per block,
    property and values; ``ids`` gives each paragraph's ``(block, paragraph id)``, ``mine``
    and ``theirs`` the regions' settings (:func:`region_settings`); ``run`` (round 43) where
    the caller states the copy's lists run on (:class:`Run`)."""
    groups: dict[tuple, list[str]] = {}
    samples: dict[tuple, tuple] = {}
    likes: dict[tuple, dict[str, str | None]] = {}
    joins: dict[str, str] = {}
    held: set[tuple] = set()                 # groups whose restart would break a stated run
    same = document is reference_document
    for paragraph, reference in pairs:
        where = ids.get(id(paragraph))
        if where is None:
            continue
        block, identifier = where
        a, b = mine.get(id(paragraph)), theirs.get(id(reference))
        if a is None or b is None:
            continue
        names = list(SETTINGS)
        if a["list_level"] != b["list_level"]:
            names = ["list_level", *_PAGE]   # the level's label, numbers and indents follow it
        elif a["list_format"] != b["list_format"]:
            names = ["list_format", *_PAGE]  # its fix places the label as the reference's
        for name in names:
            value, expected = a[name], b[name]
            if name == "list_numbering":
                if not _counted(a, b):
                    continue
                if run is not None:
                    value, expected = run.compared(a, b, identifier, joins)
            if value != expected:
                key = (block, name, value, expected)
                groups.setdefault(key, []).append(identifier)
                if run is not None and name == "list_numbering" and not b[LEAD]:
                    held.add(key)            # a restart inside a running list moves the run
                sample = reference_ids.get(id(reference), "")
                samples.setdefault(key, (sample, b, a))
                # Round 42: the same document's sample paragraph to look like, where taking
                # its look changes nothing else this comparison reads (:func:`_like_takes`).
                likes.setdefault(key, {})[identifier] = (
                    sample if same and sample and _like_takes(name, paragraph, reference, a, b) else None)
    out = []
    for key, paragraphs in groups.items():
        block, name, value, reference = key
        out.append({"block": block, "property": name, "paragraphs": paragraphs, "value": value,
                    "reference": reference,
                    "fix": settings_fix(document, name, value, reference, paragraphs, *samples[key],
                                        likes=likes[key], joins=joins, restart=key not in held)})
    order = {name: k for k, name in enumerate(SETTINGS)}
    out.sort(key=lambda d: order[d["property"]])
    return out


# -- fixes -----------------------------------------------------------------------------------------


def settings_fix(document: "Document", name: str, value: Any, reference: Any, paragraphs: list[str],
                 sample: str, sample_settings: dict, own_settings: dict,
                 likes: "dict[str, str | None] | None" = None, joins: "dict[str, str] | None" = None,
                 restart: bool = True) -> dict | None:
    """The calls that give ``paragraphs`` the reference's ``name`` (the module's "Fixes"):
    every fix of a list item or a page setting is written here, and only here.  ``sample``
    is the reference paragraph, ``sample_settings`` its settings and ``own_settings`` the
    first paragraph's.  ``likes`` (round 42): per paragraph, the sample paragraph of the
    same document it may look like (``None``: not so) -- where every paragraph has one,
    ``word_format`` ``like``, one call per sample paragraph.  Round 43, lists stated to run
    on (:class:`Run`): ``joins`` gives each paragraph the previous copy's item whose list it
    goes on from (``word_lists`` ``continue`` ``join``), and ``restart`` false holds back a
    restart that would break the run."""
    if likes and all(likes.get(p) for p in paragraphs):
        by_sample: dict[str, list[str]] = {}
        for identifier in paragraphs:
            by_sample.setdefault(likes[identifier], []).append(identifier)
        return {"tool": "word_format", "calls": [{"targets": targets, "like": like}
                                                 for like, targets in by_sample.items()]}
    if name in _FORMAT_FIELDS:
        if _style_gives(document, name, reference, paragraphs):
            return {"tool": "word_format", "calls": [{"targets": paragraphs, "clear": [name]}]}
        return {"tool": "word_format", "calls": [{"targets": paragraphs, name: reference}]}
    if name == "list_level":
        if reference is None:
            return {"tool": "word_lists", "calls": [{"targets": paragraphs, "action": "remove"}]}
        if value is None:
            # A new list of Word's own at the reference's level; a check after it compares
            # its label with the reference's.
            kind = "bullet" if (sample_settings.get("list_format") or "").startswith("bullet ") else "number"
            return {"tool": "word_lists", "calls": [{"targets": paragraphs, "action": "add", "kind": kind,
                                                      "level": reference}]}
        return {"tool": "word_lists", "calls": [{"targets": paragraphs, "action": "level", "level": reference}]}
    if name == "list_format":
        number_format, _, label = reference.partition(" ")
        place: dict[str, float] = {}
        if (own_settings["indent_left"], own_settings["indent_first"]) != \
                (sample_settings["indent_left"], sample_settings["indent_first"]):
            # The label and text where the reference has them ("Define new multilevel list").
            place = {"text_at": sample_settings["indent_left"],
                     "number_at": round(sample_settings["indent_left"] + sample_settings["indent_first"], 2)}
        calls = [{"targets": [p], "action": "format", "level": level, "number_format": number_format,
                  "label": label, **place} for p, level in _first_per_list(document, paragraphs)]
        return {"tool": "word_lists", "calls": calls} if calls else None
    if name == "list_numbering":
        if reference == RUNS_ON:
            calls = [{"targets": [p], "action": "continue", "join": joins[p]} for p in paragraphs
                     if joins and joins.get(p)]
            return {"tool": "word_lists", "calls": calls} if calls and len(calls) == len(paragraphs) else None
        if reference.startswith("starts at "):
            if not restart:
                return None
            return {"tool": "word_lists", "calls": [{"targets": [p], "action": "restart",
                                                      "start": int(reference.split()[-1])} for p in paragraphs]}
        calls = [{"targets": [p], "action": "continue"} for p in paragraphs if _can_continue(document, p)]
        return {"tool": "word_lists", "calls": calls} if len(calls) == len(paragraphs) else None
    return None


class Run:
    """Round 43: the caller states that a copy's numbered lists run on through the copies
    of one sample (one list counting 1-2 in the first copy, 3-4 in the next, ...) --
    never inferred: in Word's XML a copy that counts on by design and one that should have
    begun again are alike.  ``previous`` is ``None`` for the run's first copy (it numbers as
    its sample shows) and otherwise the copy before it's numbered items, in order: ``(level,
    list key, label, paragraph id)``.

    Where the sample's own count begins at the start of its region (:data:`LEAD`, ``starts
    at N``), a later copy's item is to go on from the copy before it: in the list of that
    copy's last item at its level and label (else at its level), its count carried on, no
    restart between -- :data:`RUNS_ON`; any other number (a restart, a separate list) is
    the difference, its fix ``word_lists`` ``continue`` joining that list (``join``), never
    a restart.  Where the copy before has no such item, the item compares as its sample.
    The level, the label and every other item compare as before; a restart is offered
    only on an item that leads its count (in a run's first copy, or with nothing to run on
    from), never inside a running list (:func:`compare_settings`)."""

    def __init__(self, previous: "list[tuple] | None" = None):
        self.previous = previous

    @classmethod
    def after(cls, document: "Document", blocks: "list[Any] | None") -> "Run":
        """The run as it stands after the copy whose region's paragraphs are ``blocks``
        (``None``: no copy before)."""
        if blocks is None:
            return cls(None)
        index = document._index(document.package.document_part())
        found = region_settings(document, blocks)
        items = []
        for element in blocks:
            values = found[id(element)]
            entry = index.entry_for(element)
            if values["list_numbering"] is not None and entry is not None:
                items.append((values["list_level"], values[LIST_KEY], values["list_format"], entry.id))
        return cls(items)

    def target(self, own: dict) -> "tuple | None":
        """The previous copy's item ``own`` goes on from: its last at ``own``'s level and
        label, else its last at the level."""
        level = [item for item in self.previous or () if item[0] == own["list_level"]]
        alike = [item for item in level if item[2] == own["list_format"]]
        return (alike or level or [None])[-1]

    def compared(self, own: dict, sample: dict, identifier: str, joins: dict[str, str]) -> tuple:
        """``(value, expected)`` of ``list_numbering`` for a copy's item and its sample's."""
        if self.previous is None or not (sample[LEAD] and sample[COUNT] == "own"
                                         and str(sample["list_numbering"]).startswith("starts at ")):
            return own["list_numbering"], sample["list_numbering"]
        found = self.target(own)
        if found is None:
            return own["list_numbering"], sample["list_numbering"]
        joins[identifier] = found[3]
        goes_on = own[COUNT] == "runs_on" and own[LIST_KEY] == found[1]
        return (RUNS_ON if goes_on else own["list_numbering"]), RUNS_ON


def _counted(own: dict, sample: dict) -> bool:
    """Whether where a copy's item stands in its count is compared with its sample's
    (round 43): only where the sample's own count begins (or begins again) inside the
    sample's region and goes on there (:data:`COUNT` ``own``) -- a numbered list of the
    sample's own, which each copy then numbers alike.  Not where the sample's item is
    numbered by the outline (its number is its place among the document's headings) or its
    count runs on from before the sample's region (it runs on by design): there the level
    and the label compare, never the number.  Nor where the copy's item is numbered by the
    outline: a restart there would break the outline's numbers."""
    return sample.get(COUNT) == "own" and own.get(COUNT) != "outline"


def _like_takes(name: str, paragraph: Any, reference: Any, own: dict, sample: dict) -> bool:
    """Whether ``word_format`` ``like`` the reference gives ``paragraph`` the reference's
    ``name`` and changes nothing else this comparison reads.  ``like`` takes the sample's
    style, list (``w:numPr``), direct indents and page settings whole
    (:mod:`docx_agent.edit.like`); the pair's styles are alike here, so:

    * a page setting or an indent: where the two state the same list (or none) -- else
      the paragraph would join the sample's list too;
    * the list level or label: where the reference draws bullets, or one of the two is in
      no list -- an item of a numbered list joining the sample's would count on from it
      (``word_lists``' explicit calls keep its numbers);
    * where its numbering starts: never (``like`` continues the sample's list)."""
    if name in LIST_SETTINGS:
        if name == "list_numbering":
            return False
        return (own["list_level"] is None or sample["list_level"] is None
                or (sample["list_format"] or "").startswith("bullet "))
    return _numbering_stated(paragraph) == _numbering_stated(reference)


def _numbering_stated(paragraph: Any) -> tuple:
    """A paragraph's own ``w:numPr`` as ``like`` compares it: its level and instance."""
    found = paragraph.find(f"{_W}pPr/{_W}numPr")
    if found is None:
        return ()
    level, number = found.find(_W + "ilvl"), found.find(_W + "numId")
    return (level.get(_W + "val") if level is not None else None,
            number.get(_W + "val") if number is not None else None)


def _style_gives(document: "Document", name: str, reference: Any, paragraphs: list[str]) -> bool:
    """Whether every paragraph has ``name`` of its own and its style (and list) alone gives
    the reference's value: clearing it is the fix, the template's value kept (rule 6)."""
    from .compare import _PARAGRAPH

    tag, attributes = _PARAGRAPH[name]
    for identifier in paragraphs:
        element = document._resolve(identifier)[1].element
        properties = element.find(_W + "pPr")
        node = properties.find(_W + tag) if properties is not None else None
        if node is None or (attributes and not any(node.get(_W + a) is not None for a in attributes)):
            return False
        applied = _effective.paragraph(document, element, direct=False)
        own = (round(applied.indent_left, 2) if name == "indent_left" else
               round(applied.first_line - applied.hanging, 2) if name == "indent_first" else
               getattr(applied, "keep_together" if name == "keep_lines" else name))
        if own != reference:
            return False
    return True


def _first_per_list(document: "Document", paragraphs: list[str]) -> list[tuple[str, int]]:
    """The first of ``paragraphs`` in each list and level: one format call per list."""
    from .numbering import membership

    seen: set = set()
    out = []
    for identifier in paragraphs:
        found = membership(document, document._resolve(identifier)[1].element)
        if found is None or (found.num_id, found.level) in seen:
            continue
        seen.add((found.num_id, found.level))
        out.append((identifier, found.level))
    return out


def _can_continue(document: "Document", identifier: str) -> bool:
    """Whether ``word_lists`` ``continue`` finds a list to continue for the paragraph: an
    earlier paragraph in another list over the same definition."""
    from .numbering import Numbering, membership

    part, entry = document._resolve(identifier)
    found = membership(document, entry.element)
    if found is None:
        return False
    numbering = Numbering(document)
    mine = numbering.canonical(found.abstract_id)
    for other in reversed(document._index(part).paragraphs[:entry.index]):
        other_list = membership(document, other.element)
        if other_list is not None and other_list.num_id != found.num_id and \
                numbering.canonical(other_list.abstract_id) == mine:
            return True
    return False


__all__ = ["LIST_SETTINGS", "REPLACED", "RUNS_ON", "Run", "SETTINGS", "compare_settings", "list_items", "region_settings",
           "settings_fix"]
