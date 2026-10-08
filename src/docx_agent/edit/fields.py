"""Fields: reading them, inserting them as Word does, and computing their results (E4).

A field is complex (``w:fldChar`` begin, ``w:instrText``, separate, its result, end, across
runs and paragraphs) or simple (``w:fldSimple``).  What Word writes was measured
(``tools/e4_probe.py``; ``tests/observations/e4-word.json``):

* ``PAGE``: `` PAGE  \\* MERGEFORMAT `` complex, its result run ``w:noProof``; ``NUMPAGES``
  and ``SECTIONPAGES`` (in a footer) the same instruction as a ``w:fldSimple`` (``headers2``).
* ``DATE``: `` DATE \\@ "d MMMM yyyy" `` complex, the date cached (``fields``).
* ``SEQ`` (a caption): a paragraph in Caption, ``Figure `` then `` SEQ Figure \\* ARABIC `` as
  a ``w:fldSimple``, then the caption's text (``fields``).
* ``REF`` and ``PAGEREF`` with ``\\h`` (E1's, the page now the number Word shows: the
  page's number in its section's format).
* ``HYPERLINK``: Word saves the field as a ``w:hyperlink`` in the Hyperlink style
  (``fields``), so that is what a hyperlink field is written as.
* **A table of contents** (``tocfield``, ``tocstyles``): `` TOC \\o "1-3" \\h \\z \\u `` (the
  switches of Word's Table of Contents; the probe's Insert Field form added ``\\*
  MERGEFORMAT``); its begin, instruction and separate open the first entry's paragraph and
  its end is the paragraph after the entries (an empty paragraph, as when Word puts it in
  one).  Each entry is a paragraph in TOC *n* with a right tab stop with a dot leader at the
  text width less 10 twips (9016 of 9026, 9376 of 9386: measured on two widths) and a
  ``w:noProof`` mark; with ``\\h`` a ``w:hyperlink`` to the heading's bookmark holding the
  heading's text (Hyperlink, ``w:noProof``), a tab and a `` PAGEREF _Toc… \\h `` field (every
  run ``w:noProof``, ``w:webHidden`` with ``\\z``, and the empty run Word writes between the
  instruction and the separate).  Each heading gets a hidden bookmark ``_Toc<9 digits>``
  around its text (Word reuses one that is there).  The TOC 1-9 styles are Word's
  (``word_toc_styles.py``).

**Results are computed here**, never by asking Word to (``w:updateFields`` is not set:
Word would prompt on every open): ``PAGE``, ``NUMPAGES``, ``SECTIONPAGES`` and ``PAGEREF``
from docx2svg's layout -- the page's number as Word shows it (``w:pgNumType``'s start and
format, docx2svg's ``fields.field_text`` for the switches); ``REF`` from the bookmark;
``SEQ`` by counting; ``DATE`` from the date given; a table of contents rebuilt from the
headings with their pages.  A table of contents changes the layout it reports, so
:meth:`FieldOps.update_fields` lays the document out again until the numbers stop moving
(at most ``passes`` times).  **Where docx2svg's layout cannot say** (a target past where it
stops, a header no page shows), the result is left as it was and the field is listed in the
result's ``unknown`` -- never marked dirty: a ``w:dirty`` field makes Word ask the user on
opening, as ``w:updateFields`` would (measured: Word's export blocks on the question).
"""

from __future__ import annotations

import copy
import datetime as _dt
import re
from dataclasses import dataclass, field as _field
from typing import TYPE_CHECKING

from ..oxml.xml import XML_SPACE, Element, make, remove
from . import ids as _ids
from . import inline as _inline
from . import text as _text
from .errors import EditError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, EditResult
    from .ranges import TextRange

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_P = _W + "p"
W_R = _W + "r"
W_PPR = _W + "pPr"
W_RPR = _W + "rPr"
_DELETED = (_W + "del", _W + "moveFrom")

#: The table of contents' instruction, as Word's Table of Contents writes it.
TOC_INSTRUCTION = ' TOC \\o "{first}-{last}"{h}{z}{u} '
#: What Word writes in a table of contents that finds no headings.
EMPTY_TOC = "No table of contents entries found."
#: Where a TOC entry's right tab stop is: the text width less this (measured).
TOC_TAB_INSET = 10
#: The first number of a ``_Toc`` bookmark made here (Word's are 9 digits).
TOC_BOOKMARK_BASE = 100000000
#: English month and day names (``DATE``'s pictures; the probe's en-GB document).
_MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
           "November", "December")
_DAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")


# -- reading ---------------------------------------------------------------------------------


@dataclass
class _Found:
    """A field as the scanner finds it."""

    part: str
    number: int
    begin: Element  # the begin w:fldChar, or the w:fldSimple
    simple: bool
    instruction: str = ""
    separate: Element | None = None
    end: Element | None = None
    nested: list = _field(default_factory=list)

    @property
    def keyword(self) -> str:
        words = self.instruction.split()
        return words[0].upper() if words else ""


def _deleted(node: Element) -> bool:
    return any(ancestor.tag in _DELETED for ancestor in node.iterancestors())


def scan(part: str, root: Element) -> list[_Found]:
    """Every field of a part in document order (nested ones too), deleted ones aside."""
    out: list[_Found] = []
    stack: list[_Found] = []
    for node in root.iter(_W + "fldChar", _W + "instrText", _W + "fldSimple"):
        if _deleted(node):
            continue
        if node.tag == _W + "fldSimple":
            found = _Found(part, len(out), node, True, node.get(_W + "instr") or "")
            out.append(found)
            continue
        if node.tag == _W + "instrText":
            if stack and stack[-1].separate is None:
                stack[-1].instruction += node.text or ""
            continue
        kind = node.get(_W + "fldCharType")
        if kind == "begin":
            found = _Found(part, len(out), node, False)
            if stack:
                stack[-1].nested.append(found)
            stack.append(found)
            out.append(found)
        elif kind == "separate" and stack:
            stack[-1].separate = node
        elif kind == "end" and stack:
            stack.pop().end = node
    return out


@dataclass(frozen=True)
class Field:
    """A field (``fld:<story>/<n>``, the n-th in its story: positional, as runs are)."""

    document: "Document"
    id: str

    def _found(self) -> _Found:
        return self.document._field(self.id)

    @property
    def instruction(self) -> str:
        """The field's instruction (``" PAGE "``, ``" TOC \\o "1-3" "``...)."""
        return self._found().instruction

    @property
    def keyword(self) -> str:
        """The field's kind: ``PAGE``, ``TOC``, ``REF``..."""
        return self._found().keyword

    @property
    def result(self) -> str:
        """The cached result's text (a table of contents': its entries, a line each)."""
        return self.document._result_text(self._found())

    @property
    def dirty(self) -> bool:
        """Whether Word is asked to update the field on opening (``w:dirty``)."""
        return (self._found().begin.get(_W + "dirty") or "false").lower() in ("1", "true", "on")

    @property
    def paragraph_id(self) -> str | None:
        """The id of the paragraph the field begins in."""
        found = self._found()
        paragraph = _paragraph_of(found.begin)
        entry = self.document._index(found.part).entry_for(paragraph) if paragraph is not None else None
        return entry.id if entry is not None else None

    def update(self, **options) -> "EditResult":
        """Compute this field's cached result: :meth:`Document.update_fields`."""
        return self.document.update_fields([self.id], **options)

    def __repr__(self) -> str:
        return f"<Field {self.id} {self.instruction.strip()!r}>"


def _paragraph_of(node: Element) -> Element | None:
    while node is not None and node.tag != W_P:
        node = node.getparent()
    return node


def _run_of(node: Element) -> Element:
    return node.getparent() if node.getparent() is not None and node.getparent().tag == W_R else node


def switches(instruction: str) -> dict[str, str | bool]:
    """A field's switches: ``\\o "1-3"`` -> ``{"o": "1-3"}``, ``\\h`` -> ``{"h": True}``,
    ``\\* ARABIC`` -> ``{"*": "ARABIC"}``.  A switch's argument is a quoted string, any word
    after ``\\@``, ``\\#``, ``\\*``, or a number after a letter (``SEQ``'s ``\\r 3``)."""
    out: dict[str, str | bool] = {}
    tokens = re.findall(r'"[^"]*"|\\.|[^\s\\"]+', instruction)
    k = 0
    while k < len(tokens):
        token = tokens[k]
        k += 1
        if not token.startswith("\\") or len(token) != 2:
            continue
        flag = token[1]
        key = flag if flag in "@#*!" else flag.lower()
        argument: str | bool = True
        if k < len(tokens) and not tokens[k].startswith("\\"):
            following = tokens[k]
            if following.startswith('"') or flag in "@#*" or following.isdigit():
                argument = following.strip('"')
                k += 1
        out.setdefault(key, argument)
    return out


def _argument(instruction: str) -> str | None:
    """A field's first argument (``REF Name \\h`` -> ``Name``)."""
    words = re.findall(r'"[^"]*"|\S+', instruction)
    if len(words) < 2 or words[1].startswith("\\"):
        return None
    return words[1].strip('"')


def format_date(picture: str, moment: _dt.date) -> str:
    """A ``DATE``'s ``\\@`` picture (``d MMMM yyyy``) applied, English names."""
    out = []
    for token in re.findall(r"'[^']*'|d{1,4}|M{1,4}|y{2,4}|H{1,2}|h{1,2}|m{1,2}|s{1,2}|AM/PM|am/pm|.", picture):
        if token.startswith("'"):
            out.append(token.strip("'"))
        elif token[0] == "d":
            out.append({1: str(moment.day), 2: f"{moment.day:02d}", 3: _DAYS[moment.weekday()][:3],
                        4: _DAYS[moment.weekday()]}[len(token)])
        elif token[0] == "M":
            out.append({1: str(moment.month), 2: f"{moment.month:02d}", 3: _MONTHS[moment.month - 1][:3],
                        4: _MONTHS[moment.month - 1]}[len(token)])
        elif token[0] == "y":
            out.append(f"{moment.year % 100:02d}" if len(token) == 2 else str(moment.year))
        elif token[0] in "Hhms" or token.lower() == "am/pm":
            hour = getattr(moment, "hour", 0)
            minute = getattr(moment, "minute", 0)
            second = getattr(moment, "second", 0)
            if token.lower() == "am/pm":
                value = "AM" if hour < 12 else "PM"
                out.append(value if token.isupper() else value.lower())
            elif token[0] == "H":
                out.append(f"{hour:02d}" if len(token) == 2 else str(hour))
            elif token[0] == "h":
                twelve = hour % 12 or 12
                out.append(f"{twelve:02d}" if len(token) == 2 else str(twelve))
            elif token[0] == "m":
                out.append(f"{minute:02d}" if len(token) == 2 else str(minute))
            else:
                out.append(f"{second:02d}" if len(token) == 2 else str(second))
        else:
            out.append(token)
    return "".join(out)


def _number(value: int, fmt: str | None) -> str:
    from docx2svg.fields import field_text

    return field_text(f" PAGE \\* {fmt} " if fmt else " PAGE ", "PAGE", page=value, pages=None,
                      section_pages=None, section_format=None)


# -- the Document half -----------------------------------------------------------------------


@dataclass
class _Context:
    """What one round of computing results knows: the layout, sections, bookmarks."""

    layout: object
    sections: list
    date: _dt.date | _dt.datetime
    unknown: list[str] = _field(default_factory=list)


class FieldOps:
    """Reading, inserting and updating fields, on :class:`docx_agent.Document`."""

    # -- reading -----------------------------------------------------------------------------

    def _all_fields(self: "Document") -> list[_Found]:
        out = []
        for part in self._parts():
            root = self.package.tree(part)
            if root is not None:
                out += scan(part, root)
        return out

    def fields(self: "Document", story: str | None = None) -> list[Field]:
        """Every field, in story and document order (nested ones after their outer)."""
        out = []
        for part in self._parts() if story is None else [self._part_for_story(story)]:
            name = self._story_of(part)
            out += [Field(self, f"fld:{name}/{found.number}") for found in scan(part, self.package.tree(part))]
        return out

    def field(self: "Document", identifier: str) -> Field:
        """A field by id (``fld:body/0``: its story and its place there); ``KeyError`` if none."""
        found = Field(self, identifier)
        found._found()
        return found

    def _field(self: "Document", identifier: str) -> _Found:
        if not identifier.startswith("fld:"):
            raise KeyError(f"{identifier!r} is not a field id")
        story, _, number = identifier[4:].rpartition("/")
        part = self._part_for_story(story)
        if part is None:
            raise KeyError(f"no story {story!r}")
        found = scan(part, self.package.tree(part))
        try:
            return found[int(number)]
        except (ValueError, IndexError):
            raise KeyError(f"no field {identifier!r}") from None

    def _field_id(self: "Document", found: _Found) -> str:
        return f"fld:{self._story_of(found.part)}/{found.number}"

    def _result_text(self: "Document", found: _Found) -> str:
        if found.simple:
            return "".join(t.text or "" for t in found.begin.iter(_W + "t"))
        if found.separate is None or found.end is None:
            return ""
        nodes = _between(found.separate, found.end)
        text = []
        for node in nodes:
            if node.tag == _W + "t" and not _deleted(node):
                text.append(node.text or "")
            elif node.tag == _W + "tab" and node.getparent().tag == W_R:
                text.append("\t")
            elif node.tag == W_P:
                text.append("\n")
        return "".join(text).strip("\n")

    # -- inserting ---------------------------------------------------------------------------

    def _field_place(self: "Document", at: "str | TextRange"):
        from .ranges import TextRange

        where = at if isinstance(at, TextRange) else self.range(at)
        if where.view != "current":
            raise EditError("edits take the current view")
        part, entries = where._entries()
        entry = entries[-1]
        offset = where.end
        items = _text.atoms(entry.element)
        if not 0 <= offset <= len(items):
            raise EditError(f"{where.id} is outside its paragraph's text")
        if 0 < offset < len(items) and items[offset].field is not None and items[offset - 1].field is items[offset].field:
            raise EditError(f"{where.id} is inside a field's result")
        return part, entry, offset

    def insert_field(self: "Document", at: "str | TextRange", instruction: str, *, result: str | None = None,
                     simple: bool = False) -> "EditResult":
        """A field at a position: complex (begin, instruction, separate, result, end), or a
        ``w:fldSimple`` with ``simple``.  ``result`` is the cached result; ``None``
        computes it where :meth:`update_fields` can (``PAGE``, ``NUMPAGES``,
        ``SECTIONPAGES``, ``PAGEREF``, ``REF``, ``SEQ``, ``DATE``), else leaves it empty
        and lists it in ``unknown``.  A ``HYPERLINK`` field is written as Word saves one: a
        ``w:hyperlink`` (:meth:`insert_hyperlink`).  The result's ``id`` is the field's."""
        keyword = instruction.split()[0].upper() if instruction.split() else ""
        if keyword == "TOC":
            raise EditError("a table of contents is inserted with insert_toc")
        if keyword == "HYPERLINK":
            target = _argument(instruction)
            anchor = switches(instruction).get("l")
            text = result or target or ""
            if isinstance(anchor, str):
                return self.insert_hyperlink(at, text, anchor=anchor)
            return self.insert_hyperlink(at, text, target)
        if not instruction.startswith(" "):
            instruction = " " + instruction
        if not instruction.endswith(" "):
            instruction += " "
        return self._insert_field(at, instruction, result, simple=simple)

    def _insert_field(self: "Document", at, instruction: str, result: str | None, *, simple: bool = False,
                      proof: bool = True) -> "EditResult":
        from .document import EditResult

        part, entry, offset = self._field_place(at)
        tracking = self._active_tracking()
        with self._edit():
            renames = self._prepare(part, [entry])
            entry = self._index(part).entry_for(entry.element) or entry
            runs = _field_runs(entry.element, offset, instruction, result or "", simple=simple, proof=proof)
            if simple:
                # A w:fldSimple stands in no revision container (as a hyperlink: E3).
                from ..revisions import track as _track
                from ..revisions.stamp import Stamp, Tracking

                _track.place_outside_revisions(runs[0], Stamp(self, tracking or Tracking.make()))
            # A w:fldSimple cannot stand in a w:ins: its result run is the insertion.
            self._track_inserted(part, [r for r in runs[0] if r.tag == W_R] if simple else runs)
            self._renew_text_id(part, entry)
            self.package.mark_dirty(part)
            self._invalidate()
            begin = runs[0] if simple else runs[0].find(_W + "fldChar")
            found = next(f for f in scan(part, self.package.tree(part)) if f.begin is begin)
            identifier = self._field_id(found)
            unknown = self._compute([found]) if result is None else []
            self._invalidate()
        return EditResult(identifier, created=[identifier], renamed=renames, unknown=unknown)

    def insert_page_number(self: "Document", at: "str | TextRange", kind: str = "PAGE") -> "EditResult":
        """``PAGE``, ``NUMPAGES`` or ``SECTIONPAGES`` as Word inserts them (a complex
        ``PAGE``, simple ``NUMPAGES`` and ``SECTIONPAGES``: measured), the result computed."""
        kind = kind.upper()
        if kind not in ("PAGE", "NUMPAGES", "SECTIONPAGES"):
            raise EditError("kind is PAGE, NUMPAGES or SECTIONPAGES")
        return self._insert_field(at, f" {kind}  \\* MERGEFORMAT ", None, simple=kind != "PAGE")

    def insert_date(self: "Document", at: "str | TextRange", picture: str = "d MMMM yyyy", *,
                    date: "_dt.date | None" = None) -> "EditResult":
        """A ``DATE`` field with its picture, the date cached (today, UTC, unless given)."""
        moment = date or _dt.datetime.now(_dt.timezone.utc).date()
        return self._insert_field(at, f' DATE \\@ "{picture}" ', format_date(picture, moment))

    def insert_sequence(self: "Document", at: "str | TextRange", label: str = "Figure") -> "EditResult":
        """A ``SEQ <label> \\* ARABIC`` field (simple, as Word's caption writes it), its
        number counted."""
        if not re.fullmatch(r"[^\W\d]\w*", label):
            raise EditError(f"{label!r} is not a sequence name")
        return self._insert_field(at, f" SEQ {label} \\* ARABIC ", None, simple=True)

    def insert_caption(self: "Document", *, after: str | None = None, before: str | None = None,
                       label: str = "Figure", text: str = "") -> "EditResult":
        """A caption paragraph (Caption style) after or before a block: ``label``, a space,
        its ``SEQ`` number, then ``text`` (``": Results"``), as Word's Insert Caption writes
        it.  The result's ``id`` is the paragraph's."""
        from .document import EditResult

        if not re.fullmatch(r"[^\W\d]\w*", label):
            raise EditError(f"{label!r} is not a caption label")
        self.styles.resolve("caption", "paragraph")
        tracking = self._active_tracking()
        with self._edit():
            inserted = self.insert_paragraph(f"{label} ", after=after, before=before, style="caption")
            paragraph = self.paragraph(inserted.id)
            field = self._insert_field(f"{paragraph.id}@{len(paragraph.text)}", f" SEQ {label} \\* ARABIC ", None,
                                       simple=True)
            if text:
                # The caption's text in a run of its own, plain (Word's: ``fields``).
                element = self.paragraph(paragraph.id)._element
                run = _text.make_run(text, None)
                element.append(run)
                part, _ = self._resolve(paragraph.id)
                if tracking is not None:
                    self._track_inserted(part, [run])
                self.package.mark_dirty(part)
        return EditResult(inserted.id, created=inserted.created + field.created,
                          renamed={**inserted.renamed, **field.renamed})

    def insert_hyperlink(self: "Document", at: "str | TextRange", text: str, url: str | None = None, *,
                         anchor: str | None = None) -> "EditResult":
        """``text`` inserted at a position and made a hyperlink -- a ``HYPERLINK`` field as
        Word saves one (a ``w:hyperlink``: measured)."""
        from .document import EditResult

        if not text:
            raise EditError("a hyperlink needs text")
        part, entry, offset = self._field_place(at)
        with self._edit():
            inserted = self.insert_text(f"{entry.id}@{offset}", text)
            paragraph = self.paragraph(inserted.id)
            link = self.add_hyperlink(paragraph.range(offset, offset + len(text)), url, anchor=anchor)
        return EditResult(link.id, created=link.created, renamed={**inserted.renamed, **link.renamed})

    # -- the table of contents ---------------------------------------------------------------

    def insert_toc(self: "Document", *, before: str | None = None, after: str | None = None,
                   levels: tuple[int, int] = (1, 3), hyperlinks: bool = True, hide_in_web: bool = True,
                   outline_levels: bool = True) -> "EditResult":
        """A table of contents of the headings at ``levels`` (outline levels 1-9 their
        styles give: ``\\o``; with ``outline_levels``, paragraphs' own too: ``\\u``), with
        hyperlinks (``\\h``) and its tab leaders and page numbers hidden on the Web
        (``\\z``), as Word writes one (module docstring); its page numbers computed from
        docx2svg's layout until they hold.  The result's ``id`` is the field's, ``created``
        the entries' paragraph ids (and the paragraph holding the field's end);
        ``unknown`` the entries whose page the layout cannot say."""
        from .document import EditResult

        if (after is None) == (before is None):
            raise EditError("give exactly one of after= or before=")
        first, last = levels
        if not 1 <= first <= last <= 9:
            raise EditError("levels are (first, last), 1 to 9")
        instruction = TOC_INSTRUCTION.format(first=first, last=last, h=" \\h" if hyperlinks else "",
                                             z=" \\z" if hide_in_web else "", u=" \\u" if outline_levels else "")
        anchor_id = after if after is not None else before
        part, entry = self._resolve(anchor_id)  # type: ignore[arg-type]
        if part != self.package.document_part():
            raise EditError("a table of contents goes in the body")
        tracking = self._active_tracking()
        # Before a paragraph, Word's form (``tocfield``): the field's end opens that
        # paragraph; elsewhere an empty paragraph holds it (as one Word puts a TOC in).
        own_holder = not (before is not None and entry.element.tag == W_P)
        with self._edit():
            if own_holder:
                holder = self.insert_paragraph("", after=after, before=before, track=False)
                renames = dict(holder.renamed)
                paragraph = self._index(part).by_id[holder.id].element
                properties = paragraph.find(W_PPR)
                if properties is not None:
                    remove(properties)
            else:
                renames = self._prepare(part, [entry], entry.index, including=True)
                paragraph = entry.element
            begin = make("w:fldChar", **{"w:fldCharType": "begin"})
            end = make("w:fldChar", **{"w:fldCharType": "end"})
            at = 1 if paragraph.find(W_PPR) is not None else 0
            for children in ([begin], [_instruction(instruction)], [make("w:fldChar", **{"w:fldCharType": "separate"})],
                             [end]):
                run = make("w:r")
                for child in children:
                    run.append(child)
                paragraph.insert(at, run)
                at += 1
            self._invalidate()
            found = next(f for f in scan(part, self.package.tree(part)) if f.begin is begin)
            unknown, _ = self._update_loop([found], passes=4)
            if tracking is not None:
                self._track_toc(part, end, tracking, own_holder)
            if not own_holder:
                self._renew_text_id(part, self._index(part).entry_for(paragraph))
            self._invalidate()
            found = next(f for f in scan(part, self.package.tree(part)) if f.end is end)
            identifier = self._field_id(found)
            made = _toc_paragraphs(found) if own_holder else _toc_paragraphs(found)[:-1]
            ids = [e.id for e in self._index(part).paragraphs if any(e.element is p for p in made)]
        return EditResult(identifier, created=[identifier, *ids], renamed=renames, unknown=unknown,
                          warnings=self._empty_tocs([identifier]))

    def _empty_tocs(self: "Document", identifiers: list[str] | None = None) -> list[str]:
        """A warning for each table of contents (of ``identifiers``, else every one) that
        found no entries -- Word's "No table of contents entries found." -- since a TOC put in
        before the headings it lists stays empty until the fields are updated (trial 2, N8)."""
        empty = []
        for item in self.fields():
            if item.keyword != "TOC" or (identifiers is not None and item.id not in identifiers):
                continue
            if item.result.strip() in ("", EMPTY_TOC):
                empty.append(f"{item.id}: the table of contents has no entries: no heading at its levels "
                             "yet; add the headings, then update the fields")
        return empty

    def _track_toc(self: "Document", part: str, end: Element, tracking, own_holder: bool) -> None:
        """A new table of contents as one insertion: its entries' marks and runs inserted
        (inside their hyperlinks, never around them: E3), and the paragraph holding its
        end -- its mark too when it is new, else only the end's run (Word: ``tracked``)."""
        from ..revisions import track as _track
        from ..revisions.stamp import Stamp

        found = next(f for f in scan(part, self.package.tree(part)) if f.end is end)
        stamp = Stamp(self, tracking)
        paragraphs = _toc_paragraphs(found)
        entries = paragraphs if own_holder else paragraphs[:-1]
        for paragraph in entries:
            _insert_content(paragraph, stamp, part)
        if not own_holder:
            end_run = _run_of(end)
            container = stamp.make("w:ins", part)
            end_run.addprevious(container)
            container.append(end_run)
        for paragraph in entries:
            if _track.is_last_in_container(paragraph):
                previous = entries[0].getprevious()
                while previous is not None and not (isinstance(previous.tag, str) and previous.tag in (W_P, _W + "tbl")):
                    previous = previous.getprevious()
                if previous is not None and previous.tag == W_P and _track.mark_record(previous) is None \
                        and previous.find(f"{W_PPR}/{_W}sectPr") is None:
                    _track.mark(previous, "ins", stamp, part)
                    continue
            _track.mark(paragraph, "ins", stamp, part)
        stamp.finish()

    # -- updating ----------------------------------------------------------------------------

    def update_fields(self: "Document", identifiers: list[str] | None = None, *, date=None,
                      passes: int = 4) -> "EditResult":
        """Compute the cached results of the document's fields (or those named): page
        numbers from docx2svg's layout, laid out again until they hold (at most
        ``passes``), tables of contents rebuilt from the headings, ``REF`` from bookmarks,
        ``SEQ`` counted, ``DATE`` from ``date`` (today, UTC).  Results are caches: not
        tracked.  What the layout cannot say is left as it was and listed in ``unknown``
        (never marked dirty: Word would ask on opening); ``count`` is how many results
        changed."""
        from .document import EditResult

        before = {self._field_id(f): self._result_text(f) for f in self._all_fields()}
        with self._edit():
            with self._untracked():
                found = self._all_fields()
                if identifiers is not None:
                    wanted = set(identifiers)
                    found = [f for f in found if self._field_id(f) in wanted]
                    missing = wanted - {self._field_id(f) for f in found}
                    if missing:
                        raise EditError(f"no field {', '.join(sorted(missing))}")
                unknown, _ = self._update_loop(found, passes=passes, date=date)
        after = {self._field_id(f): self._result_text(f) for f in self._all_fields()}
        count = sum(1 for key, value in after.items() if before.get(key) != value)
        return EditResult(None, changed=count > 0 or bool(unknown), count=count, unknown=unknown,
                          warnings=[f"{u}: the layout cannot say" for u in unknown]
                          + self._empty_tocs(identifiers))

    def _untracked(self: "Document"):
        from contextlib import contextmanager

        @contextmanager
        def scope():
            self._track_stack.append(False)
            try:
                yield
            finally:
                self._track_stack.pop()

        return scope()

    def _update_loop(self: "Document", found: list[_Found], *, passes: int, date=None) -> tuple[list[str], list[str]]:
        """Compute ``found``'s results, laying out again until they hold.  A field is known
        across passes by its end (a table of contents' begin is rebuilt, its end kept) or,
        for a simple field, its element."""
        keys = [f.begin if f.simple or f.end is None else f.end for f in found]

        def current() -> list[_Found]:
            return [f for f in self._all_fields()
                    if any((f.begin if f.simple or f.end is None else f.end) is key for key in keys)]

        unknown: list[str] = []
        previous = None
        settled = False
        for _ in range(max(1, passes)):
            unknown = self._compute(current(), date=date)
            self._invalidate()
            state = [(f.instruction, self._result_text(f)) for f in self._all_fields()]
            if state == previous:
                settled = True
                break
            previous = state
        if not settled and passes > 1:
            unknown = unknown + ["(the page numbers did not settle)"]
        return unknown, []

    def _compute(self: "Document", found: list[_Found], date=None) -> list[str]:
        """Write each field's result from one layout; the ids of those it cannot compute."""
        needs_layout = any(f.keyword in ("PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "TOC") for f in found)
        context = _Context(self.layout() if needs_layout else None, self.sections(),
                           date or _dt.datetime.now(_dt.timezone.utc).date())
        sequences: dict[str, int] = {}
        wanted = [f for f in found if f.keyword == "SEQ"]
        # SEQ numbers count every SEQ field of the body in order, updated or not.
        for item in self._all_fields():
            if item.keyword == "SEQ":
                value = self._sequence_value(item, sequences)
                if any(item.begin is f.begin for f in wanted):
                    self._write_result(item, value, context)
        tocs = [f for f in found if f.keyword == "TOC"]
        for item in found:
            keyword = item.keyword
            if keyword == "SEQ" or not _attached(item.begin):
                continue
            if keyword != "TOC" and any(item in toc.nested for toc in tocs):
                continue  # rebuilt with its table of contents
            if keyword == "TOC":
                self._rebuild_toc(item, context)
                continue
            try:
                value = self._value(item, context)
            except _Unknown:
                # Left as it is, never marked dirty: Word asks about a dirty field on
                # opening (measured: the export blocks on it), as it would for updateFields.
                context.unknown.append(self._field_id(item))
                continue
            if value is not None:
                self._write_result(item, value, context)
                if not item.simple and item.begin.get(_W + "dirty") is not None:
                    del item.begin.attrib[_W + "dirty"]
        return context.unknown

    def _sequence_value(self: "Document", item: _Found, sequences: dict[str, int]) -> str | None:
        name = (_argument(item.instruction) or "").upper()
        options = switches(item.instruction)
        if "r" in options and isinstance(options["r"], str) and options["r"].isdigit():
            sequences[name] = int(options["r"])
        elif "c" not in options:
            sequences[name] = sequences.get(name, 0) + 1
        value = sequences.get(name, 0)
        if "h" in options:
            return ""
        fmt = options.get("*")
        return _number(value, fmt if isinstance(fmt, str) and fmt.upper() != "MERGEFORMAT" else None)

    def _value(self: "Document", item: _Found, context: _Context) -> str | None:
        keyword = item.keyword
        if keyword in ("PAGE", "NUMPAGES", "SECTIONPAGES"):
            return self._page_field(item, context)
        if keyword == "PAGEREF":
            name = _argument(item.instruction)
            page = self._bookmark_page(name, context) if name else None
            if page is None:
                raise _Unknown()
            return _page_text(item.instruction, page, context)
        if keyword == "REF":
            name = _argument(item.instruction)
            try:
                return self.bookmark(name).text.split("\n")[0].replace(_text.OBJECT, "") if name else None
            except KeyError:
                raise _Unknown() from None
        if keyword == "DATE":
            picture = switches(item.instruction).get("@")
            return format_date(picture if isinstance(picture, str) else "d-M-yyyy", context.date)
        return None

    def _page_field(self: "Document", item: _Found, context: _Context) -> str:
        from docx2svg.fields import Uncomputable, field_text

        layout = context.layout
        paragraph = _paragraph_of(item.begin)
        entry = self._index(item.part).entry_for(paragraph) if paragraph is not None else None
        if entry is None or layout is None:
            raise _Unknown()
        placed = layout.where(entry.id)
        if not placed:
            raise _Unknown()
        page = layout.pages[placed[0].page - 1]
        info = page.info
        pages = None if layout.stopped else len(layout.pages)
        section_pages = None
        if info is not None and not layout.stopped:
            section_pages = sum(1 for p in layout.pages if p.info is not None and p.info.section == info.section
                                and not p.info.blank)
        fmt = self._section_format(info.section if info is not None else 0, context)
        try:
            return field_text(item.instruction, item.keyword, page=info.number if info else placed[0].page,
                              pages=pages, section_pages=section_pages, section_format=fmt)
        except Uncomputable:
            raise _Unknown() from None

    def _section_format(self: "Document", index: int, context: _Context) -> str | None:
        sections = context.sections
        if 0 <= index < len(sections):
            fmt = sections[index].page_numbering["format"]
            return None if fmt == "decimal" else fmt
        return None

    def _bookmark_page(self: "Document", name: str, context: _Context):
        """The page (docx2svg's) a bookmark starts on, or ``None`` where the layout cannot
        say."""
        layout = context.layout
        try:
            where = self.bookmark(name).range()
        except KeyError:
            return None
        placed = layout.where(where.start_id) if layout is not None else None
        if not placed:
            return None
        return layout.pages[placed[0].page - 1]

    def page_label(self: "Document", identifier: str) -> str | None:
        """The page number Word shows for the page a block starts on (its section's start
        and format), from docx2svg's layout; ``None`` where the layout cannot say."""
        context = _Context(self.layout(), self.sections(), _dt.date.today())
        placed = context.layout.where(identifier)
        if not placed:
            return None
        try:
            return _page_text(" PAGEREF x ", context.layout.pages[placed[0].page - 1], context)
        except _Unknown:
            return None

    def _write_result(self: "Document", item: _Found, value: str, context: _Context) -> None:
        """Replace a field's cached result with ``value``: the old result's first run takes
        the text (its formatting, and any revision holding it, kept) and the others go; with
        no result yet, a run in the begin run's formatting (``w:noProof`` for page numbers,
        as Word writes them)."""
        if item.simple:
            if "".join(t.text or "" for t in item.begin.iter(_W + "t")) == value:
                return
            old = [r for r in item.begin.iter(W_R)]
            if old:
                _retext(old, value)
            else:
                properties = make("w:rPr")
                properties.append(make("w:noProof"))
                item.begin.append(_result_run(value, properties))
            self.package.mark_dirty(item.part)
            return
        if item.separate is not None and self._result_text(item) == value:
            return
        begin_run = _run_of(item.begin)
        end_run = _run_of(item.end) if item.end is not None else None
        if end_run is None:
            return
        if _paragraph_of(item.begin) is not _paragraph_of(item.end):
            return  # a result across paragraphs: only a table of contents is rebuilt
        if item.separate is None:
            separate = make("w:r")
            if begin_run.find(W_RPR) is not None:
                separate.append(copy.deepcopy(begin_run.find(W_RPR)))
            separate.append(make("w:fldChar", **{"w:fldCharType": "separate"}))
            end_run.addprevious(separate)
            item.separate = separate[-1]
        separate_run = _run_of(item.separate)
        old = [node for node in _between(item.separate, item.end)
               if node.tag == W_R and not any(a.tag == W_R for a in node.iterancestors())
               and not any(c.tag in (_W + "fldChar", _W + "instrText") for c in node)]
        if old:
            _retext(old, value)
        else:
            properties = copy.deepcopy(begin_run.find(W_RPR)) if begin_run.find(W_RPR) is not None else None
            if item.keyword in ("PAGE", "PAGEREF", "NUMPAGES", "SECTIONPAGES"):
                from . import formatting as _formatting

                properties = properties if properties is not None else make("w:rPr")
                _formatting._set_child(properties, "w:noProof", {})
            separate_run.addnext(_result_run(value, properties))
        self.package.mark_dirty(item.part)

    def _rebuild_toc(self: "Document", item: _Found, context: _Context) -> None:
        """Rebuild a table of contents' entries from the headings, with their pages."""
        part = item.part
        options = switches(item.instruction)
        levels = (1, 9)
        if isinstance(options.get("o"), str) and re.fullmatch(r"\d-\d", options["o"]):
            levels = tuple(int(v) for v in options["o"].split("-"))  # type: ignore[assignment]
        hyperlinks, hidden, outline = "h" in options, "z" in options, "u" in options
        omit = options.get("n")
        omit_levels = (1, 9) if omit is True else tuple(int(v) for v in omit.split("-")) if isinstance(omit, str) \
            and re.fullmatch(r"\d-\d", omit) else None
        entries_place = _toc_paragraphs(item)
        inside = {id(p) for p in entries_place[:-1]}  # the end's paragraph may be a heading
        headings = [h for h in self._headings(levels, outline) if id(h[0].element) not in inside]
        style_ids = {}
        for _, level in headings:
            style_ids[level] = style_ids.get(level) or self.styles._ensure(f"toc {level}", "paragraph")
        link_style = self.styles._ensure("Hyperlink", "character") if hyperlinks and headings else None
        tab = self._toc_tab(entries_place[-1])
        bookmarks = [self._toc_bookmark(entry.element) for entry, _ in headings]
        pages = []
        for (entry, _), name in zip(headings, bookmarks):
            page = None
            if context.layout is not None:
                placed = context.layout.where(entry.id)
                if placed:
                    try:
                        page = _page_text(" PAGEREF x \\h ", context.layout.pages[placed[0].page - 1], context)
                    except _Unknown:
                        page = None
            if page is None:
                context.unknown.append(f"{self._field_id(item)} ({entry.id})")
            pages.append(page)
        # The new entries.
        begin_run = _run_of(item.begin)
        head = _field_head(item)
        new_paragraphs = []
        for k, ((entry, level), name, page) in enumerate(zip(headings, bookmarks, pages)):
            paragraph = make("w:p")
            properties = make("w:pPr")
            properties.append(make("w:pStyle", **{"w:val": style_ids[level]}))
            tabs = make("w:tabs")
            tabs.append(make("w:tab", **{"w:val": "right", "w:leader": "dot", "w:pos": str(tab)}))
            properties.append(tabs)
            mark = make("w:rPr")
            mark.append(make("w:noProof"))
            properties.append(mark)
            paragraph.append(properties)
            if k == 0:
                for run in head:
                    paragraph.append(run)
            container = paragraph
            if hyperlinks:
                container = make("w:hyperlink", **{"w:anchor": name, "w:history": "1"})
                paragraph.append(container)
            container.append(_toc_text_run(_heading_text(entry.element), link_style))
            number = omit_levels is None or not (omit_levels[0] <= level <= omit_levels[1])
            if number:
                for run in _toc_page_runs(name, page or "", hidden):
                    container.append(run)
            new_paragraphs.append(paragraph)
        if not headings:
            paragraph = make("w:p")
            for run in head:
                paragraph.append(run)
            paragraph.append(_text.make_run(EMPTY_TOC, None))
            new_paragraphs.append(paragraph)
        # Take the old entries out: everything from the begin to the end's paragraph.
        end_paragraph = _paragraph_of(item.end)
        begin_paragraph = _paragraph_of(item.begin)
        end_run = _run_of(item.end)
        holding = list(end_run.iterancestors())
        for node in [n for n in end_paragraph if isinstance(n.tag, str) and n is not end_run and n.tag != W_PPR
                     and not any(n is a for a in holding) and _before(n, end_run)]:
            remove(node)
        if begin_paragraph is not end_paragraph:
            for paragraph in entries_place[1:-1]:
                remove(paragraph)
            kept_before = [n for n in begin_paragraph if isinstance(n.tag, str) and n.tag != W_PPR
                           and n is not begin_run and _before(n, begin_run)]
            if kept_before:
                for node in [n for n in begin_paragraph if n is begin_run or _before(begin_run, n)]:
                    remove(node)
                anchor = begin_paragraph
                for paragraph in new_paragraphs:
                    anchor.addnext(paragraph)
                    anchor = paragraph
            else:
                for paragraph in new_paragraphs:
                    begin_paragraph.addprevious(paragraph)
                remove(begin_paragraph)
        else:
            for paragraph in new_paragraphs:
                end_paragraph.addprevious(paragraph)
        used = self._used()
        for paragraph in new_paragraphs:
            para_id, text_id = _ids.generate(part + "\0toc", used, 2)
            _ids.stamp(paragraph, para_id, text_id)
        _ids.ensure_w14(self.package.tree(part))
        self.package.mark_dirty(part)
        self._invalidate()
        item.begin = head[0].find(_W + "fldChar")

    def _headings(self: "Document", levels: tuple[int, int], outline: bool) -> list:
        """``(entry, level)`` of the body's paragraphs a TOC lists: outline levels from
        their styles (``\\o``), and their own with ``outline`` (``\\u``); empty ones aside."""
        part = self.package.document_part()
        out = []
        for entry in self._index(part).paragraphs:
            level = self._outline_level(entry.element, outline)
            if level is None or not levels[0] <= level + 1 <= levels[1]:
                continue
            if not _heading_text(entry.element).strip():
                continue
            out.append((entry, level + 1))
        return out

    def _outline_level(self: "Document", paragraph: Element, own: bool) -> int | None:
        properties = paragraph.find(W_PPR)
        if own and properties is not None:
            node = properties.find(_W + "outlineLvl")
            if node is not None and (node.get(_W + "val") or "").isdigit():
                value = int(node.get(_W + "val"))
                return value if value < 9 else None
        style = properties.find(_W + "pStyle") if properties is not None else None
        style_id = style.get(_W + "val") if style is not None else None
        if style_id is None:
            default = self.styles.default("paragraph")
            style_id = default.id if default is not None else None
        seen = set()
        while style_id and style_id not in seen:
            seen.add(style_id)
            node = self.styles._node(style_id)
            if node is None:
                return None
            level = node.find(f"{W_PPR}/{_W}outlineLvl")
            if level is not None and (level.get(_W + "val") or "").isdigit():
                value = int(level.get(_W + "val"))
                return value if value < 9 else None
            based = node.find(_W + "basedOn")
            style_id = based.get(_W + "val") if based is not None else None
        return None

    def _toc_tab(self: "Document", paragraph: Element) -> int:
        """A TOC entry's right tab stop: its section's text width less 10 twips (measured)."""
        try:
            section = self.section_of(self._index(self.package.document_part()).entry_for(paragraph).id)
            width = section.text_width
        except (KeyError, AttributeError):
            width = None
        return (width or 9026) - TOC_TAB_INSET

    def _toc_bookmark(self: "Document", paragraph: Element) -> str:
        """The heading's ``_Toc`` bookmark: one starting at its text's start (Word reuses
        it), else a new one around its text."""
        for start in paragraph.iter(_W + "bookmarkStart"):
            name = start.get(_W + "name") or ""
            if re.fullmatch(r"_Toc\d+", name):
                return name
        numbers = [int(name[4:]) for name in self._bookmark_markers() if re.fullmatch(r"_Toc\d+", name)]
        name = f"_Toc{max(numbers + [TOC_BOOKMARK_BASE]) + 1}"
        mark_id = str(self._next_annotation_id())
        start = make("w:bookmarkStart", **{"w:id": mark_id, "w:name": name})
        end = make("w:bookmarkEnd", **{"w:id": mark_id})
        # Around the heading's text: after its properties (Word's), and after a table of
        # contents' end should the heading hold it (put before the heading).
        children = [c for c in paragraph if isinstance(c.tag, str)]
        at = 0
        for k, child in enumerate(children):
            if child.tag == W_PPR or any(True for _ in child.iter(_W + "fldChar", _W + "instrText")) \
                    and not any(True for _ in child.iter(_W + "t")):
                at = k + 1
            elif child.tag != _W + "proofErr":
                break
        if at < len(children):
            children[at].addprevious(start)
        else:
            paragraph.append(start)
        paragraph.append(end)
        part = self.package.document_part()
        self.package.mark_dirty(part)
        return name


class _Unknown(Exception):
    """A result the layout cannot say."""


def _page_text(instruction: str, page, context: _Context) -> str:
    """What a ``PAGEREF`` to a page shows: its number in its section's format, unless a
    switch formats it (docx2svg's ``field_text``)."""
    from docx2svg.fields import Uncomputable, field_text

    info = page.info
    number = info.number if info is not None else page.number + 1
    sections = context.sections
    fmt = None
    if info is not None and 0 <= info.section < len(sections):
        fmt = sections[info.section].page_numbering["format"]
        fmt = None if fmt == "decimal" else fmt
    try:
        return field_text(instruction.replace("PAGEREF", "PAGE", 1), "PAGE", page=number, pages=None,
                          section_pages=None, section_format=fmt)
    except Uncomputable:
        raise _Unknown() from None


def _between(start: Element, end: Element) -> list[Element]:
    """The elements after ``start``'s run and before ``end``'s, in document order."""
    root = start
    while root.getparent() is not None:
        root = root.getparent()
    first, last = _run_of(start), _run_of(end)
    out = []
    inside = False
    for node in root.iter():
        if node is first:
            inside = True
            continue
        if node is last:
            break
        if inside and isinstance(node.tag, str) and not any(a is first for a in node.iterancestors()):
            out.append(node)
    return out


def _attached(node: Element) -> bool:
    """Whether an element is still in a part's tree (a rebuilt table of contents' old
    fields are not)."""
    root = node
    while root.getparent() is not None:
        root = root.getparent()
    return root.tag in {_W + name for name in ("document", "hdr", "ftr", "footnotes", "endnotes", "comments")}


def _before(left: Element, right: Element) -> bool:
    """Whether ``left`` comes before ``right`` in document order."""
    root = left
    while root.getparent() is not None:
        root = root.getparent()
    for node in root.iter():
        if node is left:
            return True
        if node is right:
            return False
    return False


def _retext(runs: list[Element], value: str) -> None:
    """``value`` into the first of a result's runs (its properties and the revision holding
    it kept); the other runs go, and a revision container left empty with them."""
    first = runs[0]
    for child in [c for c in first if isinstance(c.tag, str) and c.tag != W_RPR]:
        remove(child)
    for child in _text._spell(value or ""):
        first.append(child)
    for run in runs[1:]:
        parent = run.getparent()
        remove(run)
        while parent is not None and parent.tag in (_W + "ins", _W + "del") and not len(parent):
            grand = parent.getparent()
            remove(parent)
            parent = grand


def _instruction(text: str) -> Element:
    node = make("w:instrText")
    node.set(XML_SPACE, "preserve")
    node.text = text
    return node


def _result_run(value: str, properties: Element | None) -> Element:
    run = make("w:r")
    if properties is not None:
        run.append(copy.deepcopy(properties))
    for child in _text._spell(value or ""):
        run.append(child)
    return run


def _field_runs(paragraph: Element, offset: int, instruction: str, result: str, *, simple: bool,
                proof: bool) -> list[Element]:
    """A field's runs put at ``offset``, the result run ``w:noProof`` (as Word writes the
    fields it computes)."""
    if simple:
        node = make("w:fldSimple", **{"w:instr": instruction})
        properties = make("w:rPr")
        properties.append(make("w:noProof"))
        node.append(_result_run(result, properties))
        parent, index = _inline.position(paragraph, offset)
        parent.insert(index, node)
        return [node]
    begin = make("w:fldChar", **{"w:fldCharType": "begin"})
    first = _inline.new_run(paragraph, offset, [begin])
    properties = first.find(W_RPR)
    runs = [first]
    previous = first
    result_properties = copy.deepcopy(properties) if properties is not None else make("w:rPr")
    if proof:
        from . import formatting as _formatting

        _formatting._set_child(result_properties, "w:noProof", {})
    for children, props in (([_instruction(instruction)], properties),
                            ([make("w:fldChar", **{"w:fldCharType": "separate"})], properties),
                            (_text._spell(result), result_properties),
                            ([make("w:fldChar", **{"w:fldCharType": "end"})], properties)):
        run = make("w:r")
        if props is not None and len(props):
            run.append(copy.deepcopy(props))
        for child in children:
            run.append(child)
        previous.addnext(run)
        previous = run
        runs.append(run)
    return runs


def _field_head(item: _Found) -> list[Element]:
    """A TOC field's begin, instruction and separate runs, new, as the first entry opens with."""
    out = []
    for child in (make("w:fldChar", **{"w:fldCharType": "begin"}), _instruction(item.instruction),
                  make("w:fldChar", **{"w:fldCharType": "separate"})):
        run = make("w:r")
        run.append(child)
        out.append(run)
    return out


def _toc_paragraphs(item: _Found) -> list[Element]:
    """The paragraphs from a TOC field's begin to its end, both included."""
    first, last = _paragraph_of(item.begin), _paragraph_of(item.end) if item.end is not None else None
    out = [first]
    node = first
    while last is not None and node is not last:
        node = node.getnext()
        if node is None:
            break
        if isinstance(node.tag, str) and node.tag == W_P:
            out.append(node)
    return out


def _heading_text(paragraph: Element) -> str:
    text = _text.paragraph_text(paragraph).replace(_text.OBJECT, "")
    return text.replace(_text.LINE_BREAK, " ").replace("\f", " ").strip()


def _toc_text_run(text: str, link_style: str | None) -> Element:
    properties = make("w:rPr")
    if link_style is not None:
        properties.append(make("w:rStyle", **{"w:val": link_style}))
    properties.append(make("w:noProof"))
    return _result_run(text, properties)


def _toc_page_runs(bookmark: str, page: str, hidden: bool) -> list[Element]:
    """An entry's tab and its ``PAGEREF`` field's runs, as Word writes them (``w:noProof``,
    ``w:webHidden`` with ``\\z``; the empty run between the instruction and the separate)."""
    def properties() -> Element:
        node = make("w:rPr")
        node.append(make("w:noProof"))
        if hidden:
            node.append(make("w:webHidden"))
        return node

    out = []
    for children in ([make("w:tab")], [make("w:fldChar", **{"w:fldCharType": "begin"})],
                     [_instruction(f" PAGEREF {bookmark} \\h ")], [],
                     [make("w:fldChar", **{"w:fldCharType": "separate"})], None,
                     [make("w:fldChar", **{"w:fldCharType": "end"})]):
        if children is None:
            out.append(_result_run(page, properties()))
            continue
        run = make("w:r")
        run.append(properties())
        for child in children:
            run.append(child)
        out.append(run)
    return out


def _insert_content(paragraph: Element, stamp, part: str) -> None:
    """A new paragraph's runs inserted, those in a hyperlink inside it (E3: never a
    hyperlink in a ``w:ins``)."""
    from ..revisions import track as _track

    groups: list[list[Element]] = []
    for child in list(paragraph):
        if not isinstance(child.tag, str) or child.tag == W_PPR:
            continue
        if child.tag == _W + "hyperlink":
            inner = [c for c in child if isinstance(c.tag, str)]
            if inner:
                container = stamp.make("w:ins", part)
                inner[0].addprevious(container)
                for node in inner:
                    container.append(node)
            groups.append([])
            continue
        if child.tag in (_W + "bookmarkStart", _W + "bookmarkEnd"):
            groups.append([])
            continue
        if not groups or not groups[-1]:
            groups.append([])
        groups[-1].append(child)
    for group in groups:
        if not group:
            continue
        container = stamp.make("w:ins", part)
        group[0].addprevious(container)
        for node in group:
            container.append(node)
    del _track


__all__ = ["Field", "FieldOps", "TOC_INSTRUCTION", "format_date", "scan", "switches"]
