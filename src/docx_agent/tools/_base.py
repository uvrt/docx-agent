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

from ..layout import cache_key, convert_bytes, layout_options

#: A Word document of this many pages or fewer gets its reflow in every edit's checks.
REFLOW_PAGES = 20

STALE = "stale: call check"


# -- the layout engine: docx2svg in the worker pool, with a deadline ---------------------------


class Engine:
    """``document.converter``: docx-agent lays a document out through it
    (:func:`docx_agent.layout.convert_bytes` in a worker process), so a layout, a reflow,
    ``update_fields``' passes and a TOC's page numbers share one deadline
    (``Limits.layout_timeout``, 30 s), and a worker still busy at it is killed and the call
    reports ``timeout``."""

    def __init__(self, run: Callable[..., Any], timeout: float) -> None:
        self._run = run
        self.timeout = timeout
        self.deadline = time.monotonic() + timeout

    def arm(self, timeout: float | None = None) -> "Engine":
        if timeout is not None:
            self.timeout = timeout
        self.deadline = time.monotonic() + self.timeout
        return self

    def __call__(self, data: bytes, options: dict) -> tuple:
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise ToolError("timeout", f"laying the document out did not finish within "
                            f"{self.timeout:g} s", details={"timeout": self.timeout})
        return self._run(convert_bytes, data, options, timeout=remaining)


def engage(call: Any) -> None:
    """Point the call's Word documents' layouts at the toolbox's worker pool."""
    timeout = call.limits.layout_timeout
    for entry in call.entries.values():
        if entry.kind == "docx":
            entry.document.converter = Engine(call.toolbox.pool.run, timeout)


def release(entries: Iterable[Any]) -> None:
    for entry in entries:
        if getattr(entry, "kind", None) == "docx":
            entry.document.converter = None


def cached_layout(document: Any) -> Any:
    """The layout of the document's current state if it was laid out already, else ``None``."""
    conversion = document._layouts.get(cache_key(document.to_bytes(), layout_options(document, {})))
    return conversion.layout if conversion is not None else None


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
    if isinstance(value, str) and value.startswith("$"):
        return call.resolve_ref(value)
    return value


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
            if before.page_count > REFLOW_PAGES:
                out["reflow"] = STALE
            else:
                engine = document.converter
                if isinstance(engine, Engine):
                    engine.arm()
                try:
                    out["reflow"] = reflow_json(document.layout().compare(before))
                except ToolError as error:
                    out["reflow"] = f"{error.code}: {error.message}"
    finally:
        release([entry])
    return out


def validate_delta(entry: Any) -> dict[str, Any]:
    key = problem_key
    baseline = {key(problem) for problem in entry.baseline_problems}
    problems = list(entry.document.validate())
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
              reads: Callable[[Mapping[str, Any]], bool] | None = None) -> Callable[[Callable[..., Any]], Any]:
    """A Word tool (kind ``docx``): ``handler(call, **arguments) -> Result | data``.

    Around the handler: the layout engine; for a changing call, the layout before it (for
    the checks' reflow) and tracking mode (``track=False`` for one that is never tracked,
    like a comment).  ``refs`` names the arguments that may be ``$name`` refs; ``reads``
    the calls of a changing tool that only read (``word_changes`` list), which run as
    reading calls (``ooxml_edit.tools.Tool.reads``)."""

    def make(handler: Callable[..., Any]) -> Any:
        return _tool(name, description, params, group=group, mutates=mutates, kind="docx",
                     exactly_one=exactly_one, documents=documents, priority=priority, refs=refs,
                     batchable=batchable, reads=reads)(wrap(handler, mutates=mutates, track=track))

    return make


def wrap(handler: Callable[..., Any], *, mutates: bool, track: bool = True) -> Callable[..., Any]:
    @functools.wraps(handler)
    def run(call: Any, **arguments: Any) -> Any:
        engage(call)
        keep = False
        try:
            if mutates and call.entry is not None and call.changing:
                if not any(isinstance(t, Before) for t in call.context.touched.get(call.entry.doc_id, [])):
                    call.touch(Before(cached_layout(call.document)))
                if track:
                    with tracked(call, call.entry):
                        returned = handler(call, **arguments)
                else:
                    returned = handler(call, **arguments)
                keep = True  # the checks hook lays out after the call, then releases
            else:
                returned = handler(call, **arguments)
            return returned
        finally:
            if not keep:
                release(call.entries.values())

    return run


# -- arguments ---------------------------------------------------------------------------------


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
