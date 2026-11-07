"""One way to name a place across the tools that put content somewhere (round 40).

**The canonical form is ``at``**: one string holding the side and the anchor together --
``end``, ``after:<id>``, ``before:<id>``, ``replace:<id>[..<id>]``, a position
``p:A@5``, or a bare id, read as after it.  Six of the nine placing tools took it already
(``word_insert_markdown``, ``word_copy_from``, ``word_drawings``, ``word_fields``,
``word_notes``, ``word_controls``); it is one field, so a model cannot give a side without
an anchor (``where`` alone) or the reverse, and it costs no optional parameter in the
strict set; ``$ref`` works inside it (``after:$intro``).  ``word_insert_text`` takes it
now too, beside its ``target`` and ``where``.

The other tools' names for a place are read as each tool's own (``Tool.aliases``, a note
in the warnings), where renaming says the same thing:

* ``after: X`` is ``at: X`` -- a bare id is after it (blocks, text), at its end (a mark);
* ``target: X`` without ``where`` is ``at: X``;
* ``position`` and ``location`` are ``at``;
* ``at``, ``target`` given to a tool that places by ``after``/``before`` (``word_move``,
  ``word_sections``) are its ``after``, whose value may carry its side (``before:X``).

The rest -- ``before: X``, ``where`` with a target, ``find`` where a tool takes no text --
would change meaning if renamed (the side is in the name), so the call is refused with the
arguments that work in ``details.call`` (rule 2) and the message says how to write them.
"""

from __future__ import annotations

from typing import Any, Callable, Mapping

from ooxml_edit.tools import ToolError

#: The names the placing tools give a place, which another placing tool may lack.
NAMES = ("where", "target", "after", "before", "position", "location", "at", "find")

#: What a tool's ``at`` (or ``after``) is a place of: ``blocks`` (new blocks:
#: ``word_insert_markdown``, ``word_copy_from``), ``text`` (``word_insert_text``), ``mark``
#: (a note, field or control: a position or the end of a range), ``picture``
#: (``word_drawings``: a position, or ``after:``/``before:`` a block for a paragraph of its
#: own), ``sides`` (``word_move``, ``word_sections``: ``after`` and ``before`` fields).
KINDS = ("blocks", "text", "mark", "picture", "sides")

#: Values whose prefix names the side: written as they are.
SIDES = ("after:", "before:", "replace:", "end")


def is_place(value: Any) -> bool:
    """Whether ``value`` reads as an address or a place, not as a ref name: a ref name is
    letters, digits, ``_`` and ``-`` only, so a ``:``, ``@``, ``..`` or ``$`` says it is not."""
    return isinstance(value, str) and (any(c in value for c in ":@$") or ".." in value)


def aliases(tool: str, kind: str, own: Mapping[str, Any], *, prefix: str = "",
            find: bool = False) -> dict[str, Callable[[Mapping[str, Any]], str | None]]:
    """The ``Tool.aliases`` of a placing tool ``tool`` of ``kind`` whose parameters (at
    ``prefix``: ``items[].``) are ``own``: one for each name of :data:`NAMES` it lacks
    (``find``: whether it takes ``find`` itself -- then it is no alias)."""
    assert kind in KINDS, kind
    out: dict[str, Callable[[Mapping[str, Any]], str | None]] = {}
    for name in NAMES:
        if name in own or (name == "find" and find):
            continue
        out[prefix + name] = _alias(tool, kind, name, prefix)
    return out


def _alias(tool: str, kind: str, name: str, prefix: str) -> Callable[[Mapping[str, Any]], str | None]:
    field = "after" if kind == "sides" else "at"

    def meant(node: Mapping[str, Any]) -> str | None:
        value = node.get(name)
        if not isinstance(value, str) or not value:
            return None                       # nothing said, or a type the validator names
        if name == "find":
            _refuse(tool, kind, node, prefix, name, f"{tool} places by {field}, an id, not by text: "
                    f"find_text gives the id of the paragraph holding {short(value)!r}; then {field}: "
                    "'after:<that id>'")
        side, anchor, used = _read(name, value, node)
        if anchor is None:
            return None                       # where alone: the validator says it is unknown
        if kind == "sides":
            if side is None:                  # the value says its side, or is after its anchor
                if anchor.startswith("replace:"):
                    _refuse(tool, kind, node, prefix, name, f"{tool} places beside a block: give after "
                            "or before, not replace")
                return field
            _refuse(tool, kind, node, prefix, name, None, {side: anchor}, used)
        written = _write(kind, side, anchor)
        if written is None:
            _refuse(tool, kind, node, prefix, name, f"{tool} places by at: a position is "
                    "p:<id>@<offset>; before a range is its start's offset")
        if "where" not in used and written == value:
            return field                      # renamed, it says the same
        _refuse(tool, kind, node, prefix, name, None, {field: written}, used)
        return None

    return meant


def _read(name: str, value: str, node: Mapping[str, Any]) -> tuple[str | None, str | None, tuple[str, ...]]:
    """``(side, anchor, names used)`` of an alias given: ``after: X`` is (after, X); ``where``
    reads ``target`` (or a ``ref`` holding an address) as its anchor; ``None`` for a side the
    value says itself, or none (after)."""
    if name in ("after", "before"):
        return name, value, (name,)
    if name == "where":
        if value not in ("after", "before"):
            return None, None, ("where",)      # no side: the validator says it is unknown
        side = value
        if node.get("target") is not None:
            return side, node["target"], ("where", "target")
        if is_place(node.get("ref")):
            return side, node["ref"], ("where", "ref")
        return side, None, ("where",)
    if name == "target":
        where = node.get("where")
        if where is not None:
            return (where if where in ("after", "before") else None), value, ("target", "where")
        return None, value, ("target",)
    return None, value, (name,)                # at, position, location: the value as it is


def _write(kind: str, side: str | None, anchor: str) -> str | None:
    """The tool's own ``at`` for ``side`` of ``anchor``; ``None`` when it has no such place."""
    if side is None or anchor.startswith(SIDES):
        return anchor
    if kind in ("blocks", "text"):
        return anchor if side == "after" else f"before:{anchor}"      # a bare id is after it
    if kind == "picture":
        return anchor if "@" in anchor else f"{side}:{anchor}"
    # a mark (a note, field or control): at a position, or a range's end; before: its start
    if side == "after":
        return anchor
    head, at, span = anchor.partition("@")
    if not at:
        return f"{anchor}@0" if anchor.startswith("p:") and ".." not in anchor else None
    return f"{head}@{span.split(':')[0].split('..')[0]}"


def _refuse(tool: str, kind: str, node: Mapping[str, Any], prefix: str, name: str, message: str | None,
            meant: dict[str, Any] | None = None, drop: tuple[str, ...] = ()) -> None:
    """Refuse a place named another tool's way, with the arguments that say it this tool's
    way (``details``: ``call``, or ``item`` for an item of ``items``)."""
    where = prefix + name
    if meant is not None:
        fields = ", ".join(f"{k}: {v!r}" for k, v in meant.items())
        message = (f"{name} is another tool's way to name a place; {tool} takes "
                   f"{' and '.join(meant)}: write {fields}")
    details = None
    if meant is not None:
        rewritten = {k: v for k, v in node.items() if k not in drop and k != name}
        rewritten.update(meant)
        details = {"item": rewritten} if prefix else {"call": {"tool": tool, "arguments": rewritten}}
    raise ToolError("invalid_arguments", message or "", field=where,
                    valid_options=list(meant) if meant else (["after", "before"] if kind == "sides" else ["at"]),
                    details=details)


def short(text: str, limit: int = 40) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[:limit - 1] + "…"


def side_of(value: str | None) -> tuple[str, str] | None:
    """An ``after``/``before`` field's value read with its side: ``before:X`` given as
    ``after`` (an ``at`` read as ``after``) is before ``X``; ``("end", "")`` for ``end``;
    ``None`` for none."""
    if value is None:
        return None
    for side in ("after", "before"):
        if value.startswith(side + ":"):
            return side, value[len(side) + 1:]
    return ("end", "") if value == "end" else ("after", value)
