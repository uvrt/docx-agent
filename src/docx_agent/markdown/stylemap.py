"""The mapping between Markdown and Word's styles: one declarative table, read in both
directions (mammoth's idea; ROADMAP.md, "The Markdown layer").

Each :class:`Rule` says which Markdown construct a Word style (by ``w:name``, never by
id, so a localised template works) stands for.  ``style`` is the name the write half
(``insert_markdown``, :mod:`.write`) applies; ``reads`` are further names the read half
also takes for the construct -- Pandoc's names, and the styles a story's own paragraphs
carry (``footnote text``, ``header``) -- which the write half never writes.  A document
or caller may override any rule (:meth:`StyleMap.with_rules`).

Reading, in this order:

* **Headings** are found by name (``heading 1``-``heading 9``, any case) and then by the
  outline level a paragraph style declares, through its ``w:basedOn`` chain
  (``w:outlineLvl`` 0-8), so a template whose heading styles are renamed into another
  language ("Kop 1") still has headings.  Levels 7-9 are written as level 6, with their
  style named on the block.
* **Lists** come from list membership (``w:numPr``, direct or through the style), not
  from the style's name: a level's ``w:numFmt`` ``bullet`` is a bullet list, any other a
  numbered one; ``none`` is not a list.
* **Paragraph styles** by the rule naming them; a style no rule names is shown as
  ``style: <name>`` on its block.
* **Character styles** add a mark (code); emphasis and strong are read from the
  *effective* formatting instead (see :mod:`docx_agent.markdown.read`), which covers the
  Emphasis and Strong styles, other styles that make text bold or italic, and direct
  ``w:b``/``w:i`` alike.

**Writing**, ``paragraph`` is the body's paragraphs only: a table cell's paragraphs are
``table_cell`` and a footnote's are ``footnote_text``, each the document's default (Normal,
and footnote text) unless a rule names another -- so mapping ``paragraph`` to a body style
leaves cells and notes as they were.

A caller passes ``insert_markdown(style_map=...)`` a :class:`StyleMap` or a plain dict
(:meth:`StyleMap.from_dict`), whose keys are :data:`DICT_KEYS`::

    doc.insert_markdown(md, style_map={"h1": "Report Title", "h2": "Heading 1",
                                       "paragraph": "Report Body"})
"""

from __future__ import annotations

import re
from dataclasses import dataclass

#: Markdown constructs a rule can name.
CONSTRUCTS = (
    "paragraph", "heading", "bullet_list", "ordered_list", "blockquote", "code_block",
    "code", "emphasis", "strong", "strikethrough", "link", "table", "table_cell", "footnote_text",
    "footnote_reference", "thematic_break",
)

#: Which kind of style each construct is written with.
STYLE_KINDS = {
    "paragraph": "paragraph", "heading": "paragraph", "bullet_list": "paragraph",
    "ordered_list": "paragraph", "blockquote": "paragraph", "code_block": "paragraph", "table_cell": "paragraph",
    "footnote_text": "paragraph", "code": "character", "emphasis": "character",
    "strong": "character", "link": "character", "footnote_reference": "character",
    "table": "table", "strikethrough": "direct", "thematic_break": "direct",
}


@dataclass(frozen=True)
class Rule:
    """One row: Markdown ``construct`` (at ``level`` for headings and lists, 1-based) is
    Word's style ``style`` (a ``w:name``), or -- for ``direct`` constructs -- the direct
    formatting ``direct`` names."""

    construct: str
    style: str | None
    level: int = 0
    reads: tuple[str, ...] = ()
    #: For constructs Word has no style for: the direct formatting (``w:strike``, a
    #: bottom border on an empty paragraph).
    direct: str | None = None

    @property
    def kind(self) -> str:
        return STYLE_KINDS[self.construct]

    def names(self) -> tuple[str, ...]:
        return ((self.style,) if self.style else ()) + self.reads


#: The default table (decided: ROADMAP.md, "Decisions" 3).
DEFAULT_RULES: tuple[Rule, ...] = (
    Rule("paragraph", "Normal", reads=("Body Text", "First Paragraph", "Compact", "List Paragraph",
                                       "footnote text", "endnote text", "annotation text", "header", "footer")),
    *(Rule("heading", f"heading {n}", n) for n in range(1, 7)),
    Rule("bullet_list", "List Bullet", 1),
    *(Rule("bullet_list", f"List Bullet {n}", n) for n in range(2, 6)),
    Rule("ordered_list", "List Number", 1),
    *(Rule("ordered_list", f"List Number {n}", n) for n in range(2, 6)),
    Rule("blockquote", "Quote", reads=("Block Text",)),
    Rule("code_block", "HTML Preformatted", reads=("Source Code",)),
    Rule("code", "HTML Code", reads=("Verbatim Char",)),
    Rule("emphasis", "Emphasis"),
    Rule("strong", "Strong"),
    Rule("strikethrough", None, direct="w:strike"),
    Rule("link", "Hyperlink"),
    Rule("table", "Table Grid"),
    Rule("table_cell", "Normal"),
    Rule("footnote_text", "footnote text"),
    Rule("footnote_reference", "footnote reference", reads=("endnote reference",)),
    Rule("thematic_break", None, direct="w:pBdr/w:bottom"),
)

_HEADING_NAME = re.compile(r"heading ([1-9])", re.IGNORECASE)

#: The keys a plain dict ``style_map`` takes: each names a construct (and level), and its
#: value is a style name (``w:name``, an alias or an id, as ``Styles.find`` finds it).
DICT_KEYS: dict[str, tuple[str, int]] = {
    "paragraph": ("paragraph", 0),
    **{f"h{n}": ("heading", n) for n in range(1, 7)},
    "quote": ("blockquote", 0),
    "code": ("code_block", 0),
    "inline_code": ("code", 0),
    "bullet": ("bullet_list", 1),
    **{f"bullet{n}": ("bullet_list", n) for n in range(2, 6)},
    "number": ("ordered_list", 1),
    **{f"number{n}": ("ordered_list", n) for n in range(2, 6)},
    "table": ("table", 0),
    "table_cell": ("table_cell", 0),
    "footnote": ("footnote_text", 0),
    "footnote_reference": ("footnote_reference", 0),
    "emphasis": ("emphasis", 0),
    "strong": ("strong", 0),
    "link": ("link", 0),
}
#: Other spellings the dict takes for the same keys.
_DICT_ALIASES = {
    **{f"heading{n}": f"h{n}" for n in range(1, 7)},
    **{f"heading {n}": f"h{n}" for n in range(1, 7)},
    **{f"heading_{n}": f"h{n}" for n in range(1, 7)},
    "blockquote": "quote", "code_block": "code", "body": "paragraph", "cell": "table_cell",
    "footnote_text": "footnote", "bullet1": "bullet", "number1": "number",
    "bullet_list": "bullet", "ordered_list": "number",
}


class StyleMap:
    """The mapping table, with lookups for each direction."""

    def __init__(self, rules: tuple[Rule, ...] = DEFAULT_RULES) -> None:
        for rule in rules:
            if rule.construct not in CONSTRUCTS:
                raise ValueError(f"unknown Markdown construct {rule.construct!r}")
        self.rules = tuple(rules)
        self._by_name: dict[tuple[str, str], Rule] = {}
        for rule in self.rules:
            for name in rule.names():
                self._by_name.setdefault((rule.kind, name.casefold()), rule)

    #: The default table (:data:`DEFAULT`), set below the class.
    DEFAULT: "StyleMap"

    def with_rules(self, *rules: Rule) -> "StyleMap":
        """A copy in which ``rules`` replace the rows for the same construct and level."""
        keys = {(r.construct, r.level) for r in rules}
        return StyleMap(tuple(rules) + tuple(r for r in self.rules if (r.construct, r.level) not in keys))

    @classmethod
    def from_dict(cls, mapping: dict, base: "StyleMap | None" = None) -> "StyleMap":
        """``base`` (the default table) with the rows ``mapping`` names replaced::

            StyleMap.from_dict({"h1": "Report Title", "h2": "Heading 1", "paragraph": "Report Body"})

        Keys (:data:`DICT_KEYS`): ``paragraph`` (body paragraphs only), ``h1``-``h6``,
        ``quote``, ``code`` (a code block), ``inline_code``, ``bullet``, ``bullet2``-
        ``bullet5``, ``number``, ``number2``-``number5``, ``table`` (a table style),
        ``table_cell`` (its cells' paragraphs), ``footnote`` (a footnote's text),
        ``footnote_reference``, ``emphasis``, ``strong``, ``link``; ``heading1`` and
        ``blockquote`` are taken too.  A value is a style's name; a reading rule's other
        names (``reads``) are kept."""
        base = base or DEFAULT
        if isinstance(mapping, StyleMap):
            return mapping
        rules: list[Rule] = []
        for key, value in dict(mapping).items():
            if not isinstance(key, str):
                raise TypeError(f"a style_map key is a string, not {type(key).__name__}")
            name = key.strip().lower()
            name = _DICT_ALIASES.get(name, name)
            if name not in DICT_KEYS:
                raise ValueError(f"unknown style_map key {key!r}: the keys are {', '.join(DICT_KEYS)}")
            if value is not None and not isinstance(value, str):
                raise TypeError(f"style_map[{key!r}] is a style name, not {type(value).__name__}")
            construct, level = DICT_KEYS[name]
            old = next((r for r in base.rules if r.construct == construct and r.level == level), None)
            rules.append(Rule(construct, value, level, reads=old.reads if old is not None else ()))
        return base.with_rules(*rules)

    def to_dict(self) -> dict[str, str | None]:
        """The table as :meth:`from_dict` takes it (the rows a dict can name)."""
        out: dict[str, str | None] = {}
        for key, (construct, level) in DICT_KEYS.items():
            rule = next((r for r in self.rules if r.construct == construct and r.level == level), None)
            if rule is not None:
                out[key] = rule.style
        return out

    def __repr__(self) -> str:
        return f"StyleMap({self.to_dict()!r})"

    # -- reading -------------------------------------------------------------------------------

    def paragraph_rule(self, name: str | None) -> Rule | None:
        """The rule a paragraph style's name maps to, or ``None``."""
        return self._by_name.get(("paragraph", (name or "").casefold())) if name else None

    def character_rule(self, name: str | None) -> Rule | None:
        return self._by_name.get(("character", (name or "").casefold())) if name else None

    def heading_level(self, name: str | None, outline_level: int | None) -> int | None:
        """A paragraph style's heading level (1-9): by its name, else by its outline level."""
        rule = self.paragraph_rule(name)
        if rule is not None and rule.construct == "heading":
            return rule.level
        match = _HEADING_NAME.fullmatch(name or "")
        if match:
            return int(match.group(1))
        if outline_level is not None and 0 <= outline_level <= 8:
            return outline_level + 1
        return None

    # -- writing (what insert_markdown applies) -------------------------------------------------

    def style_for(self, construct: str, level: int = 0) -> Rule | None:
        """The rule the write half applies for ``construct`` at ``level``: the exact level's
        row, else the deepest row below it."""
        rows = [r for r in self.rules if r.construct == construct]
        exact = [r for r in rows if r.level == level]
        if exact:
            return exact[0]
        below = sorted((r for r in rows if r.level <= level), key=lambda r: r.level)
        return below[-1] if below else (rows[0] if rows else None)


DEFAULT = StyleMap()
StyleMap.DEFAULT = DEFAULT


def coerce(style_map) -> StyleMap:
    """``insert_markdown``'s ``style_map``: ``None`` (the default table), a
    :class:`StyleMap`, or a dict (:meth:`StyleMap.from_dict`)."""
    if style_map is None:
        return DEFAULT
    if isinstance(style_map, StyleMap):
        return style_map
    if isinstance(style_map, dict):
        return StyleMap.from_dict(style_map)
    raise TypeError(f"style_map is a dict such as {{'h1': 'Title', 'paragraph': 'Body Text'}} or a StyleMap, "
                    f"not {type(style_map).__name__}")


__all__ = ["CONSTRUCTS", "DEFAULT", "DEFAULT_RULES", "DICT_KEYS", "Rule", "StyleMap", "coerce"]
