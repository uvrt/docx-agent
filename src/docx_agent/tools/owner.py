"""One layout owner per document: a check waits for the document's next layout, laid out
once for every check waiting, and reads the version it is of (lane O, step 2).

With many writers on one long document every check laid out its own version: six writers
committing between their checks made six layouts of six versions, one after the other in
the one worker holding the document's layout, and a check waited for the five before it.
Now a check that needs the layout (``reflow``, ``layout``, a ``range``, ``numbers``) asks
the document's :class:`~docx_agent.layout.Owner` before it reads anything
(ooxml-edit's ``DocumentFormat.read_version``, :func:`read_version`): when the newest
layout is of the version the document has, or a newer one, it reads that; else it waits for
the layout in flight, and then for the next one, which lays out the newest committed version
for every check waiting then.  The check then reads *that* version -- its validation,
fields, numbers and layout facts are all of it -- and its result says so when it is newer
than the version at the call (ooxml-edit's warning).  What it reports of a version is what a
check of that version alone reports.

A relayout on commit (:func:`commit_relayout`): a change committed while the document's
checks are asking (one in the last :data:`ASKED_SECONDS`) starts the layout of the new
version at once, when none is in flight, so the next check finds it laid out or in flight.
"""

from __future__ import annotations

import threading
import time
from typing import Any

from ..layout import _keeper, own_layout, reading

#: The tools a document's owner serves (their reads wait for its next layout).
SERVED = frozenset({"check"})
def _asked_seconds() -> float:
    """``DOCX_AGENT_COMMIT_RELAYOUT`` (seconds; ``0`` turns it off) where it is set to a
    number, else 30."""
    import os

    try:
        return max(0.0, float(os.environ.get("DOCX_AGENT_COMMIT_RELAYOUT", "30")))
    except ValueError:
        return 30.0


#: A change committed this many seconds after a check of its document asked starts the
#: layout of its version at once (:func:`commit_relayout`); ``0`` turns it off.
ASKED_SECONDS = _asked_seconds()


def _needs_layout(arguments: Any) -> bool:
    """Whether a check with these arguments lays the document out (its handler or its
    numbers' pages)."""
    if not isinstance(arguments, dict):
        return False
    include = arguments.get("include")
    wanted = list(include or ["validate", "reflow", "fields"])
    return bool({"reflow", "layout", "numbers"} & set(wanted)) or any(
        arguments.get(name) is not None for name in ("range", "min_size", "styles"))


def _engine(toolbox: Any, session: Any, document: Any) -> Any:
    from ooxml_edit.tools.worker import takes_affinity

    from ._base import Engine

    pool = toolbox.pool
    if not takes_affinity(pool):
        return None       # a pool without affinity: each layout where it lands, as before
    return Engine(pool.run, session.limits.layout_timeout, keyed=True, document=document)


def _committed(entry: Any, fmt: Any) -> Any:
    """The document's committed version now, as a read takes it (ooxml-edit's
    ``Session.read_lock``): kept first, under its lock, where the version now is not yet."""
    committed = entry.committed
    if committed is None or committed.version != entry.version:
        with entry.lock:
            entry.commit(fmt)
            committed = entry.committed
    return committed


def read_version(toolbox: Any, session: Any, entry: Any, tool: Any, arguments: Any) -> Any:
    """ooxml-edit's ``DocumentFormat.read_version`` for a Word document: for a check that
    needs the layout, ``(committed, at_call)`` -- the version the owner laid out for it (the
    module's notes) and the version the document had when the check came; ``None`` for any
    other read, or where the owner could not lay the document out (the check then reads
    its own version, as before)."""
    _WAITED.pop(threading.get_ident(), None)        # none left by an earlier call
    if getattr(tool, "name", None) not in SERVED or not _needs_layout(arguments):
        return None
    if entry.pinned is not None or getattr(entry, "kind", None) != "docx":
        return None
    document = entry.document
    engine = _engine(toolbox, session, document)
    if engine is None:
        return None
    fmt = session.formats.get(entry.kind)
    committed = _committed(entry, fmt)
    if committed is None:
        return None
    at_call = committed.version
    _keeper(document)
    owner = document._layout_owner

    made: list[dict] = []

    def lay(target: Any) -> bool:
        found = own_layout(document, target.data, engine.arm())
        if found is not None:
            made.append(found)
        return found is not None

    began = time.perf_counter()
    served = owner.serve(at_call, lambda: entry.committed, lay)
    waited = time.perf_counter() - began
    if made:
        # This check laid the version out (for every check waiting): the worker's seconds
        # laying it out and reading its facts, and the rest of the wait.
        _WAITED[threading.get_ident()] = {
            "layout": made[0]["layout"], "facts": made[0]["facts"],
            "worker_wait": max(0.0, waited - made[0]["layout"] - made[0]["facts"])}
    elif waited >= 0.001:
        _WAITED[threading.get_ident()] = {"worker_wait": waited}   # for another's layout
    if served is None:
        return None
    reading(owner, served.version)      # until its facts are made (check_facts) or it ends
    return served, at_call


#: Per thread, where the last check on it spent its wait for its document's layout
#: (:func:`read_version`; the hook and the check run on one thread, one after the other).
_WAITED: dict[int, dict[str, float]] = {}


def layout_waited() -> dict[str, float]:
    """Where this thread's check spent its wait for the owner's layout, taken: ``layout``
    and ``facts`` (the worker's, where it laid the version out) and ``worker_wait`` (the
    rest: the worker's queue, the checks reading the version before, another's layout) --
    the parts ``check``'s ``timing`` says (round 47); empty where it did not wait."""
    return _WAITED.pop(threading.get_ident(), None) or {}


def layout_first(call: Any) -> None:
    """Lane O: a change that lays its document out ahead of its work (``word_fields``
    update) first waits -- the document's lock let go -- for the owner's layout of the
    version the document has: the one its checks read, laid out once and held where the
    document's layouts are, rather than laid out again beside them (and, its holder busy
    with them, laid out whole on another worker)."""
    entry = call.entry
    if entry is None or not call.may_unlock or getattr(entry, "kind", None) != "docx":
        return
    from ._base import Engine, engage

    engage(call)
    engine = call.document.converter
    committed = entry.committed
    if not isinstance(engine, Engine) or not engine.keyed or committed is None \
            or committed.version != entry.version:
        return
    document = entry.document
    _keeper(document)
    owner = document._layout_owner
    with call.unlocked():
        try:
            owner.serve(committed.version, lambda: entry.committed,
                        lambda target: own_layout(document, target.data, engine.arm()) is not None, ask=False, wait=True)
        except Exception:  # noqa: BLE001 -- the work ahead lays the document out itself
            pass


def commit_relayout(entry: Any, engine: Any) -> None:
    """Lane O: a change committed while its document's checks ask lays the new version out
    at once (:data:`ASKED_SECONDS`), on a thread of its own, when nothing is laid out for
    the document then -- the next check finds the layout made or in flight.  Once per check
    asking (:meth:`~docx_agent.layout.Owner.take_arm`): of a run of edits between two
    checks only the first is laid out at once (the next check's layout starts from it)."""
    if not ASKED_SECONDS or engine is None:
        return
    live = getattr(entry, "live", None) or entry
    document = live.document
    _keeper(document)
    owner = document._layout_owner
    if not owner.take_arm(ASKED_SECONDS):
        return
    committed = live.committed
    if committed is None:
        return

    def run() -> None:
        try:
            owner.serve(committed.version, lambda: live.committed,
                        lambda target: own_layout(document, target.data, engine.arm()) is not None, ask=False)
        except Exception:  # noqa: BLE001 -- the next check lays the version out itself
            pass

    threading.Thread(target=run, name="docx-agent relayout on commit", daemon=True).start()
