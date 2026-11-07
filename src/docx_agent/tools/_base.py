"""What every Word tool shares: the decorator, tracking mode, the layout engine, checks.

* :func:`word_tool` makes a tool of a handler, as :func:`ooxml_edit.tools.tool` does, and
  wraps it: the document's layouts run in the toolbox's worker pool with a deadline
  (:class:`Engine`), and a changing call is written as tracked changes when the session set
  tracking on for the document (:func:`tracked`).
* :func:`checks` is the docx format's ``checks`` hook: after a changing call -- or once, at
  the end of a ``batch`` -- the ``validate()`` delta against the document as opened and the
  pages the change reflowed.
* :func:`outcome` turns the library's :class:`~docx_agent.EditResult`\\ s into the result
  envelope: what changed, was created, removed or renamed, and the warnings.
* Addresses go to the library verbatim (``p:3B212964``, ``p:3B212964@4:11``,
  ``p:A..t:B``, ``t:``, ``c:``, ``rev:``, ``fn:``, ``d:``, ``s:``, story names).  The
  dispatcher replaces a ``$name`` ref in a declared target (``Tool.refs``) before the
  handler runs; :func:`resolve` does it for the rest (``at: after:$intro``).

Lengths are points at the boundary; the library takes points already.  Dates written come
from the session clock.
"""

from __future__ import annotations

import functools
import time
from contextlib import contextmanager
from typing import Any, Callable, Iterable, Iterator, Mapping, Sequence

from ooxml_edit.tools import Result, ToolError, page_list, tool as _tool

from ooxml_edit.tools.worker import takes_affinity

from ..layout import (adopt, affinity, canon_of, known_summary, layout_key, summary_for, cache_key, convert_bytes, held_state, lay_out_bytes,
                      lay_out_unlocked, layout_options, twin, twin_of)

#: A Word document of this many pages or fewer gets its reflow in every edit's checks.
REFLOW_PAGES = 20

STALE = "stale: call check"


# -- the layout engine: docx2svg in the worker pool, with a deadline ---------------------------


class Engine:
    """``document.converter``: docx-agent lays a document out through it
    (:func:`docx_agent.layout.lay_out_bytes` in a worker process), so a layout, a reflow,
    ``update_fields``' passes and a TOC's page numbers share one deadline
    (``Limits.layout_timeout``, 30 s), and a worker still busy at it is killed and the call
    reports ``timeout``.

    Round 38: a document's layouts stay in the worker that made them, and its layout tasks
    go there (ooxml-edit 0.90.0's affinity, keyed by the document's lineage), waiting for
    that worker while busy at most as long as laying the document out whole took
    (:func:`docx_agent.layout.affinity`) -- ``keyed``: the pool takes affinity -- of
    ``document``'s layouts."""

    def __init__(self, run: Callable[..., Any], timeout: float, *, keyed: bool = False,
                 document: Any = None) -> None:
        self._run = run
        self.timeout = timeout
        self.deadline = time.monotonic() + timeout
        self.keyed = keyed
        self.document = document

    def arm(self, timeout: float | None = None) -> "Engine":
        if timeout is not None:
            self.timeout = timeout
        self.deadline = time.monotonic() + self.timeout
        return self

    def __call__(self, data: bytes, options: dict) -> tuple:
        return self._run(convert_bytes, data, options, timeout=self._remaining())

    def lay_out(self, data: bytes, options: dict, state: Any = None) -> tuple:
        """:func:`docx_agent.layout.lay_out_bytes` in a worker process, under the deadline:
        the layout kept for the next version (round 33), in the worker holding it (round 38)."""
        return self._run(lay_out_bytes, data, options, state, timeout=self._remaining(),
                         **self.keys(state))

    def check(self, data: bytes, options: dict, state: Any, request: dict) -> dict:
        """``check``'s layout facts made in the worker holding the document's layout (round
        38, :func:`docx_agent.tools.check_worker.check_bytes`), under the deadline."""
        from .check_worker import check_bytes

        return self._run(check_bytes, data, options, state, request, timeout=self._remaining(),
                         **self.keys(state))

    def keys(self, state: Any) -> dict:
        """The pool's ``affinity`` and ``affinity_wait`` for a task handed ``state``: the
        document's lineage, and the wait :func:`docx_agent.layout.affinity` says."""
        if not self.keyed or not held_state(state):
            return {}
        wait = affinity(self.document)[1] if self.document is not None else None
        return {"affinity": ("docx-agent layout", state[3]), "affinity_wait": wait}

    def _remaining(self) -> float:
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise ToolError("timeout", f"laying the document out did not finish within "
                            f"{self.timeout:g} s", details={"timeout": self.timeout})
        return remaining


def engage(call: Any) -> None:
    """Point the call's Word documents' layouts at the toolbox's worker pool."""
    timeout = call.limits.layout_timeout
    pool = call.toolbox.pool
    for entry in call.entries.values():
        if entry.kind == "docx":
            entry.document.converter = Engine(pool.run, timeout, keyed=takes_affinity(pool),
                                              document=entry.document)


def release(entries: Iterable[Any]) -> None:
    for entry in entries:
        if getattr(entry, "kind", None) == "docx":
            entry.document.converter = None


#: How many times a changing call lays its document out ahead of its change (round 33,
#: R16c) when another call changed it meanwhile: each time from the layout before, so a few
#: seconds; past them the change lays out what is left under the lock (from the last one).
AHEAD_TRIES = 2


def lay_out_ahead(call: Any) -> bool:
    """Lay the call's document out with its lock let go, ahead of a change that reads the
    layout first (``word_fields`` update: page numbers; a page cross-reference; the upgrade's
    reflow) -- ooxml-edit's ``Tool.aheads`` (0.79.0), run before the call's undo step.

    The layout is of the version the lock was let go at, kept as every layout is
    (:func:`docx_agent.layout.convert`); when another call changed the document meanwhile,
    it is laid out again from that one, at most :data:`AHEAD_TRIES` times in all.  The
    handler, run after under the lock, finds the layout of the state it holds cached, or
    lays out what changed since from the last one.  Whether the layout ahead was of the
    state the call then holds.  A layout that fails is left to the handler, which meets it
    again and says so: the work ahead never changes what the call does."""
    if call.entry is None or not call.may_unlock:
        return False
    entry, document = call.entry, call.document
    for _ in range(AHEAD_TRIES):
        version = entry.version
        # The engine again, under the lock: a call that ran while the lock was let go
        # pointed the document's converter at its own engine, or at none, as it ended.
        engage(call)
        try:
            lay_out_unlocked(document, call.unlocked, lambda: entry.version == version)
        except Exception:  # noqa: BLE001 -- the handler lays out again and reports it
            return False
        if entry.version == version:
            return True
    return False


def update_fields_ahead(call: Any, ids: list[str] | None) -> bool:
    """``word_fields`` update's work ahead (round 33, R16c): the whole update -- every pass
    and its layout -- on a copy of the document opened from its bytes, with the lock let go;
    then the copy's layouts are the document's (:func:`docx_agent.layout.adopt`).  The
    update under the lock, writing what the copy's wrote, finds each pass's layout of those
    very bytes cached, and lays nothing out; where another call changed the document
    meanwhile, it is done again from there (at most :data:`AHEAD_TRIES` times in all), and
    past that the update lays out what changed since under the lock.  Whether the copy's
    update was of the state the call then holds.  Anything that fails is left to the update
    under the lock, which meets it again and says so."""
    if call.entry is None or not call.may_unlock:
        return False
    entry, document = call.entry, call.document
    for _ in range(AHEAD_TRIES):
        engage(call)
        version = entry.version
        try:
            wanted = resolve(call, ids) if ids else None
        except Exception:  # noqa: BLE001 -- the update refuses it, under the lock
            return False
        data, like, date = entry.snapshot(), twin_of(document), call.now().date()
        copy = None
        with call.unlocked():
            try:
                copy = twin(data, like)
                copy.update_fields(wanted, date=date)
            except Exception:  # noqa: BLE001 -- the update lays out again and reports it
                copy = None
        if copy is None:
            return False
        adopt(document, copy)
        if entry.version == version:
            return True
    return False


def seen_before(document: Any) -> Any:
    """What a changing call's reflow is measured from: the summary of the layout of the
    document's current state when it was laid out already -- here, or (round 38) in the
    worker that made a check's facts -- else ``None``."""
    options = layout_options(document, {})
    data = document.to_bytes()
    key = layout_key(data, options)
    found = known_summary(document, key)
    if found is not None:
        return found
    conversion = document._layouts.get(cache_key(data, options))
    return summary_for(document, key, conversion.layout) if conversion is not None else None


def cached_layout(document: Any, *, fetch: bool = True) -> Any:
    """The layout of the document's current state if it was laid out already, else ``None``.
    Round 38: laid out in the worker that made a check's facts, it is fetched from there
    (nothing laid out again) -- ``fetch`` false: only whether it was (``True``)."""
    options = layout_options(document, {})
    data = document.to_bytes()
    conversion = document._layouts.get(cache_key(data, options))
    if conversion is not None:
        return conversion.layout
    if known_summary(document, layout_key(data, options)) is None:
        return None
    return document.layout() if fetch else True


# -- tracking mode -----------------------------------------------------------------------------


@contextmanager
def tracked(call: Any, entry: Any) -> Iterator[None]:
    """Every edit inside written as tracked changes, when the session tracks this document
    (``word_set_tracking``), by its author and dated by the session clock."""
    mode = entry.tracking if entry is not None else None
    if not mode or not mode.get("on"):
        yield
        return
    with entry.document.tracking(author=mode["author"], date=call.now(), initials=mode.get("initials")):
        yield


def author_of(call: Any, given: str | None) -> str | None:
    """The author a comment is written by: the one given, else the tracking author."""
    if given:
        return given
    mode = call.entry.tracking if call.entry is not None else None
    return mode.get("author") if mode else None


# -- refs --------------------------------------------------------------------------------------


def resolve(call: Any, value: Any) -> Any:
    """An address as the library takes it: a ``$name`` ref replaced (the dispatcher already
    did so for a tool's declared targets; this covers the rest)."""
    if isinstance(value, list):
        return [resolve(call, item) for item in value]
    if isinstance(value, str) and ".." in value and "$" in value:
        # A span's ends, either a ref: $intro..$outro, p:A..$end, $start.. (the dispatcher
        # resolved a leading one).
        first, _, last = value.partition("..")
        return "..".join(call.resolve_ref(end.strip()) if end.strip().startswith("$") else end
                         for end in (first, last))
    if isinstance(value, str) and value.startswith("$"):
        value = call.resolve_ref(value)
    return _heading_range(call, value)


def _heading_range(call: Any, value: Any) -> Any:
    """A heading's range (``h:p:A``; round 36, R32) as the span of blocks it is, so every
    parameter taking a span takes it; a refusal names the body's headings."""
    if not (isinstance(value, str) and (value.startswith("h:") or "..h:" in value)):
        return value
    try:
        document = call.document
    except Exception:  # noqa: BLE001 -- no document: the handler refuses the address itself
        return value
    if document is None or not hasattr(document, "_expand_heading"):
        return value
    from ..edit.addresses import AddressError

    try:
        return document._expand_heading(value)
    except AddressError as error:
        raise ToolError("invalid_arguments", str(error), valid_options=list(error.options)[:50]) from None


def block_of(where: Any, *args: str) -> str | None:
    """:meth:`Document.block_of <docx_agent.Document.block_of>` with a ``$name`` ref resolved
    first (refs live in the session, not the document): the top-level body block holding an
    id.  In a handler, ``block_of(call, identifier)`` -- the call's document, its refs as the
    tools resolve them; in the application, outside a call, ``block_of(session, doc,
    identifier)`` -- the open document ``doc`` (``"d1"``), under its lock, its refs as
    :meth:`Session.resolve_ref <ooxml_edit.tools.Session.resolve_ref>` resolves them::

        block_of(session, "d1", "$intro")     # -> "t:5B1F07A2"
    """
    with _addressed(where, args, 1, "block_of") as (document, (identifier,)):
        return document.block_of(identifier)


def blocks_between(where: Any, *args: str) -> list[str]:
    """:meth:`Document.blocks_between <docx_agent.Document.blocks_between>`, either end a
    ``$name`` ref: ``blocks_between(call, first, last)`` in a handler,
    ``blocks_between(session, doc, first, last)`` outside a call (see :func:`block_of`)."""
    with _addressed(where, args, 2, "blocks_between") as (document, (first, last)):
        return document.blocks_between(first, last)


@contextmanager
def _addressed(where: Any, args: Sequence[str], count: int, name: str) -> Iterator[tuple[Any, list[str]]]:
    """The document and the ``count`` ids of a call's ``(call, *ids)`` or the application's
    ``(session, doc, *ids)``, each ``$name`` resolved: :func:`block_of`'s and
    :func:`blocks_between`'s one way in."""
    in_call = hasattr(where, "document")      # a session has documents, not one
    if in_call and len(args) == count:        # a handler's call: it holds the document's lock
        yield where.document, [resolve(where, value) for value in args]
        return
    if in_call or len(args) != count + 1:
        ends = "identifier" if count == 1 else "first, last"
        raise TypeError(f"{name} takes (call, {ends}) in a handler or (session, doc, {ends}) outside one")
    from .format import KIND                  # format imports this module

    doc, *values = args
    entry = where.entry(doc)                  # not_found names the open documents
    if entry.kind != KIND:
        raise ToolError("invalid_arguments", f"{doc} is a {entry.kind} document; {name} reads a Word "
                        "document's blocks", field="doc")
    with entry.lock:
        yield entry.document, [where.resolve_ref(doc, value) for value in values]


def remember(call: Any, ref: str | None, identifier: str | None) -> None:
    """Name ``identifier`` ``ref`` for later calls (``"$ref"``)."""
    if ref and identifier is not None:
        call.define_ref(ref, identifier)


# -- checks ------------------------------------------------------------------------------------


class Before:
    """What a changing call touches first: the layout of the state before it (or ``None``),
    for the reflow the checks report.  Equal to every other, so a batch keeps the first."""

    def __init__(self, layout: Any) -> None:
        self.layout = layout

    def __eq__(self, other: object) -> bool:
        return isinstance(other, Before)

    def __hash__(self) -> int:
        return 0

    def __repr__(self) -> str:
        return "Before()"


def checks(entry: Any, touched: Sequence[Any]) -> dict[str, Any]:
    """The docx format's ``checks``: the ``validate()`` delta against the document's problems
    when opened, and the reflow -- the pages the change touched, from the layout before (when
    it was laid out) and after, for a document of :data:`REFLOW_PAGES` pages or fewer
    (``"stale: call check"`` beyond)."""
    document = entry.document
    before = next((t.layout for t in touched if isinstance(t, Before)), None)
    out: dict[str, Any] = {"validate": validate_delta(entry), "reflow": None}
    try:
        if before is not None:
            if before.count > REFLOW_PAGES:
                out["reflow"] = STALE
            else:
                engine = document.converter
                if isinstance(engine, Engine):
                    engine.arm()
                try:
                    # By summaries (round 38), ids followed through the document's aliases now.
                    after = document.layout()
                    now = summary_for(document, after.key if after.key is not None else layout_key(
                        document.to_bytes(), layout_options(document, {})), after)
                    out["reflow"] = reflow_json(now.compare(before, canon_of(document)))
                except ToolError as error:
                    out["reflow"] = f"{error.code}: {error.message}"
    finally:
        release([entry])
    return out


def validate_delta(entry: Any, problems: Sequence[Any] | None = None) -> dict[str, Any]:
    """The document's problems (``problems``, else ``validate()``'s) against those it had
    when opened."""
    key = problem_key
    baseline = {key(problem) for problem in entry.baseline_problems}
    problems = list(entry.document.validate()) if problems is None else list(problems)
    now = {key(problem) for problem in problems}
    return {"new": [str(p) for p in problems if key(p) not in baseline],
            "fixed": [str(p) for p in entry.baseline_problems if key(p) not in now],
            "baseline": len(entry.baseline_problems)}


def problem_key(problem: Any) -> Any:
    return (getattr(problem, "code", None), getattr(problem, "part", None), str(problem))


def reflow_json(reflow: Any) -> dict[str, Any]:
    data: dict[str, Any] = {"changed": list(reflow.changed), "pages": list(reflow.page_count)}
    if reflow.moved:
        data["moved"] = {k: list(v) for k, v in list(reflow.moved.items())[:20]}
    if reflow.unknown:
        data["unknown"] = list(reflow.unknown)
    return data


# -- small text a builder wrote (round 17, P6b) -------------------------------------------------


def note_small_text(call: Any, result: Result, blocks: Sequence[str]) -> Result:
    """``result`` with ``small_text`` -- the text of ``blocks`` (paragraph, table or cell ids)
    its styles make smaller than the application's minimum for its kind
    (:func:`docx_agent.edit.minimums.styled_below`) -- and a warning saying so; ``result``
    unchanged without a minimum or with nothing below it."""
    from ..edit.minimums import styled_below, text_minimums

    document = call.document
    minimums = text_minimums(getattr(document, "min_text_size", None))
    if not minimums:
        return result
    paragraphs: list[str] = []
    for identifier in blocks:
        try:
            found = document.get(identifier)
        except (KeyError, ValueError):
            continue
        if hasattr(found, "grid"):           # a table: every cell's paragraphs
            seen = set()
            for line in found.grid():
                for cell in line:
                    if cell is not None and cell.id not in seen:
                        seen.add(cell.id)
                        paragraphs += [p.id for p in cell.paragraphs]
        elif hasattr(found, "paragraphs") and not hasattr(found, "text_boxes"):
            paragraphs += [p.id for p in found.paragraphs]
        else:
            paragraphs.append(identifier)
    below = styled_below(document, list(dict.fromkeys(paragraphs)), minimums)
    if below:
        if isinstance(result.data, dict):
            result.data["small_text"] = below[:20]
        smallest = below[0]
        result.warnings.append(f"small_text: {len(below)} paragraph(s) below the minimum, the smallest "
                               f"{smallest['para']} at {smallest['size']:g} pt ({smallest['kind']}, minimum "
                               f"{smallest['min']:g} pt)")
    return result


# -- results -----------------------------------------------------------------------------------


def outcome(results: Iterable[Any], summary: str, *, data: Any = None,
            changed: Sequence[str] = ()) -> Result:
    """The envelope for one or more :class:`~docx_agent.EditResult`\\ s."""
    result = Result(summary=summary, data=data)
    for edit in results:
        if edit is None:
            continue
        for key, values in (("created", edit.created), ("removed", edit.removed)):
            target = getattr(result, key)
            target.extend(v for v in values if v not in target)
        if edit.id is not None and edit.changed and edit.id not in result.created \
                and edit.id not in result.changed:
            result.changed.append(edit.id)
        result.renamed.update(edit.renamed)
        result.warnings.extend(w for w in edit.warnings if w not in result.warnings)
        for unknown in edit.unknown:
            note = f"{unknown}: the layout cannot say"
            if note not in result.warnings:
                result.warnings.append(note)
    for identifier in changed:
        if identifier not in result.changed:
            result.changed.append(identifier)
    return result


# -- the decorator -----------------------------------------------------------------------------


def word_tool(name: str, description: str, params: Mapping[str, Any], *, group: str,
              mutates: bool = False, exactly_one: Sequence[Sequence[str]] = (),
              documents: Any = ("doc",), priority: int | None = None, track: bool = True,
              refs: Sequence[str] = (), batchable: bool = True,
              reads: Callable[[Mapping[str, Any]], bool] | None = None,
              sources: Sequence[str] = (), facts: Sequence[str] = (),
              aliases: Mapping[str, Any] | None = None,
              action_roles: Mapping[str, str] | None = None,
              call_role: Callable[[Mapping[str, Any]], str | None] | None = None,
              review_core: bool = False,
              action_fields: Mapping[str, Sequence[str]] | None = None,
              ahead: Callable[..., Any] | None = None
              ) -> Callable[[Callable[..., Any]], Any]:
    """A Word tool (kind ``docx``): ``handler(call, **arguments) -> Result | data``.

    Around the handler: the layout engine; for a changing call, the layout before it (for
    the checks' reflow) and tracking mode (``track=False`` for one that is never tracked,
    like a comment).  ``refs`` names the arguments that may be ``$name`` refs; ``reads``
    the calls of a changing tool that only read (``word_changes`` list), which run as
    reading calls (``ooxml_edit.tools.Tool.reads``); ``sources`` the arguments that name
    rows to take, a dataset or another document's table (``Tool.sources``); ``facts`` the
    text that may hold references to one value, ``{{d:x/key/column}}``, written as the value
    before the handler runs (``Tool.facts``); ``aliases`` the argument names read as its own
    (``Tool.aliases``); ``action_roles`` and ``call_role`` the roles of a changing tool's
    actions that are not ``edit``'s (``Tool.action_roles``: a ``list`` action is
    ``review``'s); ``review_core`` loads it in the core of a role set without ``edit``
    (``Tool.review_core``); ``action_fields`` the parameters only some actions take, not sent
    to a role set offered none of them (``Tool.action_fields``, ooxml-edit 0.69.0);
    ``ahead`` what a changing call does before its change, its lock let go
    (``Tool.aheads``, ooxml-edit 0.79.0: :func:`lay_out_ahead`), with the layout engine."""

    def make(handler: Callable[..., Any]) -> Any:
        return _tool(name, description, params, group=group, mutates=mutates, kind="docx",
                     exactly_one=exactly_one, documents=documents, priority=priority, refs=refs,
                     batchable=batchable, reads=reads, sources=sources, facts=facts,
                     aliases=aliases, action_roles=action_roles,
                     call_role=call_role, review_core=review_core,
                     action_fields=action_fields,
                     ahead=_engaged(ahead) if ahead is not None else None)(
                         wrap(handler, mutates=mutates, track=track))

    return make


def _engaged(work: Callable[..., Any]) -> Callable[..., Any]:
    """``work`` with the layout engine, as a handler has it (:func:`wrap`)."""
    @functools.wraps(work)
    def run(call: Any, **arguments: Any) -> Any:
        engage(call)
        try:
            return work(call, **arguments)
        finally:
            release(call.entries.values())

    return run


#: What a call's result says when the text it wrote got its language's spacing (round 35, R21).
SPACED = ("{} space(s) before : ; ! ? or inside « » written as no-break spaces (U+00A0), as Word's "
          "proofing of the text's language expects")


def _space_written(call: Any) -> int:
    """The spacing of the text the call wrote, now, inside its undo step: how many spaces."""
    space = getattr(call.document, "space_written", None)
    return space() if space is not None else 0


def wrap(handler: Callable[..., Any], *, mutates: bool, track: bool = True) -> Callable[..., Any]:
    @functools.wraps(handler)
    def run(call: Any, **arguments: Any) -> Any:
        engage(call)
        keep = False
        try:
            if mutates and call.entry is not None and call.changing:
                if not any(isinstance(t, Before) for t in call.context.touched.get(call.entry.doc_id, [])):
                    call.touch(Before(seen_before(call.document)))
                if track:
                    with tracked(call, call.entry):
                        returned = handler(call, **arguments)
                        spaced = _space_written(call)
                else:
                    returned = handler(call, **arguments)
                    spaced = _space_written(call)
                if getattr(call, "facts", None):
                    # Round 36: what was written by reference is kept on its paragraph.
                    from .textfacts import keep_written

                    keep_written(call, returned, arguments)
                if spaced and isinstance(returned, Result):
                    returned.warnings.append(SPACED.format(spaced))
                keep = True  # the checks hook lays out after the call, then releases
            else:
                returned = handler(call, **arguments)
            return returned
        finally:
            if not keep:
                release(call.entries.values())

    return run


# -- arguments ---------------------------------------------------------------------------------


#: A paragraph's properties the Python API names otherwise than the tools: ``keep_together``
#: (``Paragraph.keep_together``, ``w:keepLines``) is ``keep_lines`` to a model, as
#: ``word_format`` names it -- a table's ``keep_together`` (``word_format_table``, check's
#: fact) is another thing (round 25, W4).
TOOL_NAMES = {"keep_together": "keep_lines"}


def tool_names(data: Any, within: str) -> Any:
    """``data`` with the Python API's paragraph names (:data:`TOOL_NAMES`) as the tools say
    them, in every object under a key ``within`` (``effective``, ``look``) at any depth."""
    if isinstance(data, list):
        return [tool_names(item, within) for item in data]
    if not isinstance(data, dict):
        return data
    out = {}
    for key, value in data.items():
        if key == within and isinstance(value, dict):
            value = {TOOL_NAMES.get(k, k): v for k, v in value.items()}
        out[key] = tool_names(value, within)
    return out


def need(value: Any, name: str, why: str = "") -> Any:
    if value is None:
        raise ToolError("invalid_arguments", f"{name} is required{why}", field=name)
    return value


def one_of(arguments: Mapping[str, Any], *names: str) -> str:
    """The one of ``names`` given; ``invalid_arguments`` unless exactly one is."""
    given = [n for n in names if arguments.get(n) is not None]
    if len(given) != 1:
        raise ToolError("invalid_arguments", f"give exactly one of {', '.join(names)}"
                        + (f"; got {', '.join(given)}" if given else ""),
                        field=(given[1] if len(given) > 1 else names[0]), valid_options=list(names))
    return given[0]


def page(items: Sequence[Any], cursor: str | None, limit: int) -> tuple[list[Any], int, str | None]:
    return page_list(items, cursor=cursor, limit=limit)


def short(text: str, limit: int = 120) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= limit else text[:limit - 1] + "…"
