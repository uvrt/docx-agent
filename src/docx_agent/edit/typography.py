"""Spacing around punctuation as Word's proofing of the text's language expects (round 35,
R21): written with the text a call writes, and reported by ``check`` where it is missing.

**Measured** in Word 16 for Mac (``tools/typography_probe.py``; the facts in
``tests/observations/typography-word.json``):

* *Proofing.*  Word's grammar check of French (``fr-FR``, ``fr-CH`` and a bare ``fr``) flags
  ``:`` ``;`` ``!`` ``?`` with no space before them (``Bonjour:``, ``Fin?``, ``Quoi?!``,
  ``« oui »!``) and guillemets with no space inside (``«bonjour»``, ``« déjà»``); an
  ordinary space, U+00A0 and U+202F all pass.  Canadian French (``fr-CA``) flags the
  guillemets only.  For Belgian, Luxembourg and Monaco French this Word flags nothing at
  all (its French grammar check does not run for them): they take French's rule here, as
  ``fr`` alone does.  Word also flags a mark between two words (``12:30``, ``x?y``,
  ``v2:beta``), whose fix is not a space; those are left alone.  English (US, GB), German,
  Dutch, Spanish, Italian and Portuguese flag the opposite: any space -- ordinary, U+00A0 or
  U+202F -- before ``:`` ``;`` ``!`` ``?``; Spanish, Italian and Portuguese also inside
  guillemets.
* *Line breaks.*  With an ordinary space before ``!`` ``?`` ``:`` ``;`` or ``»`` Word breaks
  a line before the mark, which then starts the next line; with U+00A0 or U+202F it never
  does.
* *What Word writes.*  Word's AutoFormat turns straight quotes in French text into ``«``
  U+00A0 ... U+00A0 ``»`` (fr-FR, fr-CA, fr-CH, fr-BE alike): U+00A0, the space written
  here.  It leaves the spaces before ``:`` ``;`` ``!`` ``?`` as they are; AutoFormat as you
  type cannot be driven by a script (typed text is not autoformatted).
* *Language.*  A run's language is its own ``w:lang``, else its character style's, its
  paragraph style's (through ``w:basedOn``), ``w:docDefaults``'; ``w:themeFontLang`` is not
  it (a document with only that is in Word's own editing language).  A table style's
  language counts between the defaults and the paragraph style, as the cascade has it.

**Writing** (:func:`apply_written`): in text a call wrote -- recorded as it is spelled into
``w:t`` nodes (:class:`Recorder`) -- whose language has a rule here (:data:`WRITE_RULES`),
an ordinary space the call wrote before a mark or inside guillemets becomes U+00A0, and
where there is none, U+00A0 is put in next to what the call wrote.  An ordinary space that
was there before stays (``check`` reports it).  Nothing is removed: the languages that want
no space get no change on writing -- text tagged English that holds French spacing is more
often French with the wrong language than English with stray spaces -- and ``check``
reports them as that (``language_mismatch``).  Text in a field's result, in a run Word does not proof
(``w:noProof``), or written as code (Markdown code, the code styles) is not looked at.
``Document.typography`` (``docx_format(typography=False)``) turns writing it off.

**Checking** (:func:`problems`): every place in the document where the rule of the text's
language is not met, per paragraph, with ``replace_text`` calls that fix them; spaces a
language that wants none has are a ``language_mismatch`` whose fix sets the language
(``Document.set_language``, round 35, R24), never a deletion.
"""

from __future__ import annotations

import contextvars
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Callable, Iterable

from ..oxml.xml import Element
from . import text as _text

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

NBSP = "\u00a0"
NNBSP = "\u202f"
#: The marks French puts a space before.
MARKS = ":;!?"
OPEN, CLOSE = "«", "»"
#: The spaces that keep a mark on its word's line (both pass Word's proofing; U+00A0 is
#: what Word writes).
NO_BREAK = (NBSP, NNBSP)
_SPACES = (" ", NBSP, NNBSP)

#: What is written, by language (lower case; a tag not listed takes its primary
#: subtag's): ``full`` -- before ``: ; ! ?`` and inside guillemets; ``guillemets`` -- inside
#: guillemets only.
WRITE_RULES = {"fr": "full", "fr-ca": "guillemets"}
#: Languages whose proofing flags any space before ``: ; ! ?`` (``check`` only).
NO_SPACE_BEFORE = frozenset({"en", "de", "nl", "es", "it", "pt"})
#: ... and inside guillemets.
NO_SPACE_INSIDE = frozenset({"es", "it", "pt"})

#: What may end the word a mark follows (a letter or digit too).
_WORD_END = "»)]}”’\"'%.…°€$£"
#: What may follow a mark (whitespace and the end too): not a letter, a digit, ``/``...
_AFTER_MARK = "»”’\"'.,…"
#: Characters a fix's literal find stops at (an object, a break).
_STOPS = frozenset({"\ufffc", "\v", "\f", "\n"})  # an object (text.OBJECT), breaks

#: The names of the code styles (``docx_agent.markdown.stylemap``'s code and code block).
CODE_STYLES = frozenset({"html code", "verbatim char", "html preformatted", "source code"})


def rule_of(language: str | None, rules: dict[str, str] = WRITE_RULES) -> str | None:
    """The rule ``language`` (a ``w:lang`` value) writes, or ``None``."""
    if not language:
        return None
    tag = language.lower()
    return rules.get(tag) or rules.get(tag.split("-")[0])


def _primary(language: str | None) -> str:
    return (language or "").lower().split("-")[0]


# -- what a call wrote -------------------------------------------------------------------------


class Recorder:
    """The characters a call spelled into ``w:t`` nodes, by node: their offsets in its text.
    Splits and joins of a recorded node carry its offsets along (``inline``)."""

    def __init__(self) -> None:
        self._nodes: dict[int, tuple[Element, set[int]]] = {}
        #: Round 45: how many times what it holds changed -- spacing applied at one count, in
        #: an unchanged document, is applied (``Document._space_written``).
        self.version = 0

    def __bool__(self) -> bool:
        return bool(self._nodes)

    def add(self, node: Element, offsets: Iterable[int] | None = None) -> None:
        self.version += 1
        length = len(node.text or "")
        found = set(range(length)) if offsets is None else {k for k in offsets if 0 <= k < length}
        if found:
            self._nodes[id(node)] = (node, found)
        else:
            self._nodes.pop(id(node), None)

    def offsets(self, node: Element) -> set[int]:
        held = self._nodes.get(id(node))
        return held[1] if held is not None and held[0] is node else set()

    def written(self, node: Element, offset: int) -> bool:
        return offset in self.offsets(node)

    def split(self, node: Element, left: Element, at: int) -> None:
        """``node``'s text ``[:at]`` moved to ``left`` (before it)."""
        held = self.offsets(node)
        if not held:
            return
        self.version += 1
        self._nodes[id(left)] = (left, {k for k in held if k < at})
        self._nodes[id(node)] = (node, {k - at for k in held if k >= at})

    def join(self, kept: Element, gone: Element, shift: int) -> None:
        """``gone``'s text appended to ``kept``'s, which was ``shift`` long."""
        self.version += 1
        moved = {k + shift for k in self.offsets(gone)}
        self._nodes.pop(id(gone), None)
        if moved:
            self._nodes[id(kept)] = (kept, self.offsets(kept) | moved)

    def paragraphs(self) -> list[Element]:
        """The paragraphs holding recorded nodes still in a tree, in no particular order."""
        out: dict[int, Element] = {}
        for node, _ in self._nodes.values():
            paragraph = node.getparent()
            while paragraph is not None and paragraph.tag != _text.W_P:
                paragraph = paragraph.getparent()
            if paragraph is not None:
                out[id(paragraph)] = paragraph
        return list(out.values())


_RECORDER: contextvars.ContextVar[Recorder | None] = contextvars.ContextVar("docx_agent_typography",
                                                                             default=None)


def recorder() -> Recorder | None:
    """The recorder of the edit in progress on this thread, if it records."""
    return _RECORDER.get()


def record_all(element: Element) -> None:
    """Every ``w:t`` under ``element`` written by the call (text a builder serialised and
    parsed again, which drops what was recorded)."""
    record = _RECORDER.get()
    if record is not None:
        for node in element.iter(_W + "t"):
            record.add(node)


def start() -> contextvars.Token:
    return _RECORDER.set(Recorder())


def stop(token: contextvars.Token) -> None:
    _RECORDER.reset(token)


# -- the language and proofing of a run --------------------------------------------------------


class Cascade:
    """A run's language, whether Word proofs it and whether it is code, through the styles
    (the module's *Language*)."""

    def __init__(self, document: "Document") -> None:
        self.styles: dict[str, Element] = {}
        self.default_paragraph: str | None = None
        self.default_table: str | None = None
        self.defaults: Element | None = None
        package = document.package
        part = package.styles_part()
        root = package.tree(part) if part else None
        if root is not None:
            self.defaults = root.find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr")
            for style in root.iter(_W + "style"):
                identifier = style.get(_W + "styleId")
                if identifier is None:
                    continue
                self.styles[identifier] = style
                if style.get(_W + "default") in ("1", "true", "on"):
                    kind = style.get(_W + "type")
                    if kind == "paragraph":
                        self.default_paragraph = identifier
                    elif kind == "table":
                        self.default_table = identifier
        self._code = {identifier for identifier, style in self.styles.items()
                      if (_name(style) or "").lower() in CODE_STYLES}
        self._cache: dict[tuple, tuple[str | None, bool, bool, str | None]] = {}

    def _chain(self, identifier: str | None) -> list[Element]:
        out: list[Element] = []
        seen: set[str] = set()
        while identifier and identifier not in seen and identifier in self.styles:
            seen.add(identifier)
            style = self.styles[identifier]
            out.append(style)
            based = style.find(_W + "basedOn")
            identifier = based.get(_W + "val") if based is not None else None
        return out

    def _chain_ids(self, identifier: str | None) -> set[str]:
        return {style.get(_W + "styleId") for style in self._chain(identifier)}

    def of(self, paragraph: Element, run: Element | None) -> tuple[str | None, bool, bool]:
        """``(language, proofed, code)`` of ``run`` in ``paragraph``."""
        return self.of_with_source(paragraph, run)[:3]  # type: ignore[return-value]

    def of_with_source(self, paragraph: Element,
                       run: Element | None) -> tuple[str | None, bool, bool, str | None]:
        """``(language, proofed, code, source)``: :meth:`of` and where the language comes
        from -- ``"run"`` (the run's own ``w:lang``), ``"style:<name>"`` or ``"default"``
        (``w:docDefaults``'); ``None`` with no language."""
        properties = paragraph.find(_W + "pPr")
        node = properties.find(_W + "pStyle") if properties is not None else None
        paragraph_style = node.get(_W + "val") if node is not None else self.default_paragraph
        table_style = None
        holder = paragraph.getparent()
        while holder is not None and holder.tag != _W + "tbl":
            holder = holder.getparent()
        if holder is not None:
            node = holder.find(f"{_W}tblPr/{_W}tblStyle")
            table_style = node.get(_W + "val") if node is not None else self.default_table
        direct = run.find(_W + "rPr") if run is not None else None
        node = direct.find(_W + "rStyle") if direct is not None else None
        character_style = node.get(_W + "val") if node is not None else None
        key = (paragraph_style, table_style, character_style)
        base = self._cache.get(key)
        if base is None:
            levels: list[tuple[Element | None, str]] = [(self.defaults, "default")]
            for style_id in (table_style, paragraph_style, character_style):
                # Each level's chain, its own declaration first.
                levels.extend(reversed([(s.find(_W + "rPr"), "style:" + (_name(s) or s.get(_W + "styleId") or ""))
                                        for s in self._chain(style_id)]))
            language = source = None
            proofed = True
            for level, where in levels:
                if level is None:
                    continue
                found = level.find(_W + "lang")
                if found is not None and found.get(_W + "val"):
                    language, source = found.get(_W + "val"), where
                found = level.find(_W + "noProof")
                if found is not None:
                    proofed = not _on(found)
            code = bool(self._code & (self._chain_ids(paragraph_style) | self._chain_ids(character_style)))
            base = (language, proofed, code, source)
            self._cache[key] = base
        language, proofed, code, source = base
        if direct is not None:
            found = direct.find(_W + "lang")
            if found is not None and found.get(_W + "val"):
                language, source = found.get(_W + "val"), "run"
            found = direct.find(_W + "noProof")
            if found is not None:
                proofed = not _on(found)
        return language, proofed, code, source


def _name(style: Element) -> str | None:
    node = style.find(_W + "name")
    return node.get(_W + "val") if node is not None else None


def _on(node: Element) -> bool:
    return node.get(_W + "val") not in ("0", "false", "off")


# -- spotting ----------------------------------------------------------------------------------


@dataclass(frozen=True)
class Spot:
    """One place the rule is not met.  ``kind``: ``breakable`` (an ordinary space where the
    rule wants U+00A0), ``missing`` (no space), ``space_before`` / ``space_inside`` (a space
    the language does not want).  ``action``: ``replace`` the character at ``at`` with
    U+00A0, ``insert`` U+00A0 before position ``at``, or ``delete`` the character at ``at``.
    ``mark``: the punctuation's position; ``language``: its run's."""

    kind: str
    action: str
    at: int
    mark: int
    language: str


def _wordish(char: str) -> bool:
    return char.isalnum() or char in _WORD_END


def _ends_mark(text: str, j: int) -> bool:
    return j >= len(text) or text[j].isspace() or text[j] in _AFTER_MARK


def find(text: str, language_at: Callable[[int], str | None], looked_at: Callable[[int], bool],
         *, check: bool = False) -> list[Spot]:
    """The places in ``text`` where the rule of the language at each mark is not met;
    ``looked_at(i)`` says whether character ``i`` is text the rule applies to.  With
    ``check`` also the spaces the languages that want none have."""
    spots: list[Spot] = []
    n = len(text)
    opens = [i for i, c in enumerate(text) if c == OPEN]
    closes = [i for i, c in enumerate(text) if c == CLOSE]
    for i, char in enumerate(text):
        if char not in MARKS and char not in (OPEN, CLOSE):
            continue
        if not looked_at(i):
            continue
        language = language_at(i)
        rule = rule_of(language)
        primary = _primary(language)
        if char in MARKS:
            if i == 0 or text[i - 1] in MARKS:
                continue  # the first of ?! or !!! only
            j = i
            while j < n and text[j] in MARKS:
                j += 1
            if not _ends_mark(text, j):
                continue  # 12:30, x?y, http://
            before = text[i - 1]
            if rule == "full":
                if before == " " and i >= 2 and _wordish(text[i - 2]) and looked_at(i - 1):
                    spots.append(Spot("breakable", "replace", i - 1, i, language))
                elif _wordish(before) and before not in MARKS:
                    spots.append(Spot("missing", "insert", i, i, language))
            elif check and primary in NO_SPACE_BEFORE and rule is None:
                if before in _SPACES and i >= 2 and _wordish(text[i - 2]) and looked_at(i - 1):
                    spots.append(Spot("space_before", "delete", i - 1, i, language))
            continue
        # Guillemets, paired in the paragraph: an opening one with a closing one after it.
        if char == OPEN:
            if not any(k > i for k in closes) or i + 1 >= n:
                continue
            after = text[i + 1]
            if rule is not None:
                if after == " " and i + 2 < n and not text[i + 2].isspace() and text[i + 2] != CLOSE \
                        and looked_at(i + 1):
                    spots.append(Spot("breakable", "replace", i + 1, i, language))
                elif not after.isspace() and after != CLOSE:
                    spots.append(Spot("missing", "insert", i + 1, i, language))
            elif check and primary in NO_SPACE_INSIDE and after in _SPACES and looked_at(i + 1):
                spots.append(Spot("space_inside", "delete", i + 1, i, language))
        else:
            if not any(k < i for k in opens) or i == 0:
                continue
            before = text[i - 1]
            if rule is not None:
                if before == " " and i >= 2 and not text[i - 2].isspace() and text[i - 2] != OPEN \
                        and looked_at(i - 1):
                    spots.append(Spot("breakable", "replace", i - 1, i, language))
                elif not before.isspace() and before != OPEN:
                    spots.append(Spot("missing", "insert", i, i, language))
            elif check and primary in NO_SPACE_INSIDE and before in _SPACES and looked_at(i - 1):
                spots.append(Spot("space_inside", "delete", i - 1, i, language))
    return spots


class _Paragraph:
    """A paragraph's current text, each character's atom, and the rule's view of them."""

    def __init__(self, cascade: Cascade, paragraph: Element) -> None:
        self.element = paragraph
        self.atoms = _text.atoms(paragraph)
        self.text = "".join(atom.char for atom in self.atoms)
        self.cascade = cascade
        self._runs: dict[int, tuple[str | None, bool, bool, str | None]] = {}

    def _of(self, i: int) -> tuple[str | None, bool, bool, str | None]:
        run = self.atoms[i].run
        key = id(run)
        found = self._runs.get(key)
        if found is None:
            found = self._runs[key] = self.cascade.of_with_source(self.element, run)
        return found

    def source_at(self, i: int) -> str | None:
        return self._of(i)[3]

    def language_at(self, i: int) -> str | None:
        return self._of(i)[0]

    def looked_at(self, i: int) -> bool:
        atom = self.atoms[i]
        if atom.kind != "text" or atom.in_field:
            return False
        _, proofed, code, _ = self._of(i)
        return proofed and not code


# -- writing ----------------------------------------------------------------------------------


def apply_written(document: "Document", record: Recorder,
                  paragraphs: list[Element]) -> tuple[int, list[Element]]:
    """Put the module's spacing in the text ``record`` holds in ``paragraphs`` (the
    document's): the number of spaces written or made U+00A0, and the paragraphs changed
    (their parts are the caller's to mark)."""
    if not paragraphs:
        return 0, []
    cascade = Cascade(document)
    made = 0
    changed: list[Element] = []
    for element in paragraphs:
        paragraph = _Paragraph(cascade, element)
        atoms = paragraph.atoms

        def written(i: int) -> bool:
            atom = atoms[i]
            return atom.kind == "text" and record.written(atom.node, atom.offset)

        edits: dict[int, tuple[Element, list[tuple[int, str]]]] = {}

        def edit(atom_index: int, offset: int, action: str) -> None:
            node = atoms[atom_index].node
            edits.setdefault(id(node), (node, []))[1].append((offset, action))

        for spot in find(paragraph.text, paragraph.language_at, paragraph.looked_at):
            if spot.action == "replace":
                if written(spot.at):
                    edit(spot.at, atoms[spot.at].offset, "replace")
                continue
            # An insertion at ``spot.at``: between characters ``at - 1`` and ``at``, into the
            # punctuation's node where the call wrote it, else into its neighbour's.
            left, right = spot.at - 1, spot.at
            mark_side = right if spot.mark == right else left
            other = left if mark_side == right else right
            for side in (mark_side, other):
                if 0 <= side < len(atoms) and written(side):
                    offset = atoms[side].offset + (1 if side == left else 0)
                    edit(side, offset, "insert")
                    break
        for node, changes in edits.values():
            value = node.text or ""
            for offset, action in sorted(changes, reverse=True):
                if action == "replace":
                    value = value[:offset] + NBSP + value[offset + 1:]
                else:
                    value = value[:offset] + NBSP + value[offset:]
                made += 1
            node.text = value
        if edits:
            changed.append(element)
    return made, changed


# -- checking ---------------------------------------------------------------------------------


#: Marks listed per paragraph at most (``marks_left_out`` counts the rest).
MARKS_PER_PARAGRAPH = 10


#: The kinds of :class:`Spot` a language that wants no space reports.
_MISMATCH_KINDS = ("space_before", "space_inside")


def problems(document: "Document") -> list[dict[str, Any]]:
    """``check``'s ``punctuation_spacing``: per paragraph of every story, the places the rule
    of the text's language is not met (the module's), in document order.

    ``{"para", "language", "marks": [{"at", "kind"}], "fix"}``: ``at`` the 0-based offset of
    the space (or of the mark a space is missing before; after ``«``, of the character after
    it), ``kind`` the module's (:class:`Spot`), and the mark's ``language`` where it is not
    the paragraph's first; ``fix``: ``replace_text`` calls in the paragraph, each find the
    shortest text around one or more marks that occurs there once.  Making them (or
    rewriting the text: written text gets the spacing) clears them.

    Spaces a language that wants none has (``space_before``, ``space_inside``: French
    spacing in text tagged English, Dutch...) are most often French text under the wrong
    language tag, not stray spaces; deleting them would undo the author's French.  They are
    an entry of their own, ``"kind": "language_mismatch"``, with ``source`` -- where the tag
    comes from (``default``, ``style:<name>``, ``run``) -- and no ``replace_text``: its
    ``fix`` sets the language (``word_styles`` ``set_language``: the document's -- its default
    and the styles in the template's own language -- or, for a style in another language,
    that style's) to a French one (:func:`french_guess`), and a run's own tag has none."""
    cascade = Cascade(document)
    guess: list[str] = []
    own: list[list[tuple[str, str]]] = []
    out: list[dict[str, Any]] = []
    for part in document._parts():
        for entry in document._index(part).paragraphs:
            raw = "".join(node.text or "" for node in entry.element.iter(_W + "t"))
            if not any(c in raw for c in MARKS + OPEN + CLOSE):
                continue  # nothing to look at (cheaper than the paragraph's atoms)
            paragraph = _Paragraph(cascade, entry.element)
            spots = find(paragraph.text, paragraph.language_at, paragraph.looked_at, check=True)
            if not spots:
                continue
            spacing = [spot for spot in spots if spot.kind not in _MISMATCH_KINDS]
            mismatched = [spot for spot in spots if spot.kind in _MISMATCH_KINDS]
            if spacing:
                found = _entry(entry.id, spacing)
                found["fix"] = {"tool": "replace_text",
                                "calls": _fix_calls(entry.id, paragraph.text, spacing[:MARKS_PER_PARAGRAPH])}
                out.append(found)
            if mismatched:
                if not guess:
                    guess.append(french_guess(document))
                source = paragraph.source_at(mismatched[0].mark) or "default"
                found = _entry(entry.id, mismatched)
                found = {"para": found.pop("para"), "kind": "language_mismatch", **found, "source": source}
                if source == "run":
                    found["note"] = (f"French spacing in runs tagged {mismatched[0].language} themselves: likely "
                                     "French text under the wrong tag; not a spacing fault if so")
                else:
                    call: dict[str, Any] = {"action": "set_language", "language": guess[0]}
                    if source.startswith("style:") and not _template_own(document, own, mismatched[0].language):
                        # A style in a language the template does not otherwise use: that
                        # style's alone.  In the template's own language (Normal, the headings
                        # of an English template): the document's, which retags them all.
                        call["name"] = source[len("style:"):]
                    found["note"] = (f"French spacing in text tagged {mismatched[0].language}: likely French "
                                     "under the wrong language; the fix sets it (the text's own region if not "
                                     f"{guess[0]}: fr-BE, fr-CH, fr-CA...)")
                    found["fix"] = {"tool": "word_styles", "calls": [call]}
                out.append(found)
    return out


def _template_own(document: "Document", own: list, language: str | None) -> bool:
    """Whether ``language`` is the template's own (``docx_agent.edit.language``'s rule);
    ``own`` caches the rule's answer across one :func:`problems`."""
    from .language import own_languages

    if not own:
        part = document.package.styles_part()
        own.append(own_languages(document.package.tree(part) if part else None, "val"))
    return any(_same_tag(language, found) for found, _ in own[0])


def _same_tag(a: str | None, b: str | None) -> bool:
    return a is not None and b is not None and a.replace("_", "-").lower() == b.replace("_", "-").lower()


def _entry(identifier: str, spots: list[Spot]) -> dict[str, Any]:
    shown = spots[:MARKS_PER_PARAGRAPH]
    found: dict[str, Any] = {
        "para": identifier, "language": spots[0].language,
        "marks": [{"at": spot.at, "kind": spot.kind, **({"language": spot.language}
                                                        if spot.language != spots[0].language else {})}
                  for spot in shown]}
    if len(spots) > len(shown):
        found["marks_left_out"] = len(spots) - len(shown)
    return found


def french_guess(document: "Document") -> str:
    """The French tag a ``language_mismatch`` fix proposes: the document locale's when it is
    French, else a French tag the document's defaults or styles state, else ``fr-FR``."""
    locale = getattr(getattr(document, "locale", None), "language", None)
    if _primary(locale) == "fr":
        return str(locale)
    package = document.package
    part = package.styles_part()
    root = package.tree(part) if part else None
    if root is not None:
        for lang in root.iter(_W + "lang"):
            value = lang.get(_W + "val")
            if _primary(value) == "fr" and "-" in (value or ""):
                return str(value)
    return "fr-FR"


def _window(text: str, spot: Spot) -> tuple[int, int]:
    """The stretch a fix of ``spot`` rewrites: from the start of the word before to the mark."""
    start = min(spot.at, spot.mark)
    end = max(spot.at + (0 if spot.action == "insert" else 1), spot.mark + 1)
    while start > 0 and text[start - 1] not in _STOPS and not text[start - 1].isspace():
        start -= 1
    return start, end


def _fix_calls(identifier: str, text: str, spots: list[Spot]) -> list[dict[str, Any]]:
    """``replace_text`` calls that fix ``spots``: each the shortest stretch around them,
    widened to whole words, that occurs once in the paragraph; spots whose stretches touch
    share one call, so no call's find is changed by another's."""
    windows: list[list[Any]] = []
    for spot in sorted(spots, key=lambda spot: spot.at):
        start, end = _window(text, spot)
        if windows and start <= windows[-1][1]:
            windows[-1][1] = max(windows[-1][1], end)
            windows[-1][2].append(spot)
        else:
            windows.append([start, end, [spot]])

    def unique(a: int, b: int) -> bool:
        return text.find(text[a:b]) == a and text.find(text[a:b], a + 1) < 0

    calls = []
    for start, end, held in windows:
        while not unique(start, end):
            grew = False
            if start > 0 and text[start - 1] not in _STOPS:
                start -= 1
                grew = True
            if end < len(text) and text[end] not in _STOPS:
                end += 1
                grew = True
            if not grew:
                break
        piece = text[start:end]
        for spot in sorted(held, key=lambda spot: spot.at, reverse=True):
            at = spot.at - start
            if spot.action == "replace":
                piece = piece[:at] + NBSP + piece[at + 1:]
            elif spot.action == "insert":
                piece = piece[:at] + NBSP + piece[at:]
            else:
                piece = piece[:at] + piece[at + 1:]
        calls.append({"range": identifier, "find": text[start:end], "replace": piece, "expect": "one"})
    return calls
