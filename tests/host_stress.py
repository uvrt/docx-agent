"""The application's own edits in the serial-replay stress (round 48): ``block_stress.run`` with
some of each thread's calls replaced by an edit the application makes through the library while
holding ``session.lock`` -- a paragraph added after a block, a block's text set, both in one
``batch()``, or the lock held with nothing changed.  Letting the lock go commits the change, so
the calls on drafts after it draft at once; every result, the bytes and the id state must equal
the calls run one at a time in the order their changes were committed, as in ``block_stress``.
Not collected itself: ``test_host_edits_serial.py`` runs it small, and ``python
tests/host_stress.py --seeds 0-19 --threads 8 --calls 16`` runs it long.
"""

from __future__ import annotations

import argparse
import json
import random
import sys
import time
from typing import Any

from ooxml_edit.tools import Result, Toolbox

import block_stress

#: The name a planned application edit goes by (no tool has it).
HOST = "host_edit"


def with_host_edits(plans: list[list[tuple[str, dict]]], seed: int) -> list[list[tuple[str, dict]]]:
    """``plans`` with about a third of each thread's calls that name a block replaced by an
    application edit of that block (of the first block a batch's first op names)."""
    out = []
    for k, calls in enumerate(plans):
        rng = random.Random(seed * 104729 + k)
        mine = []
        for c, (name, arguments) in enumerate(calls):
            given = arguments if name != "batch" else (arguments.get("ops") or [{}])[0].get("arguments", {})
            target = given.get("target")
            if isinstance(target, str) and target.startswith(("p:", "t:")) and rng.random() < 0.35:
                kind = rng.choice(["insert_after", "set_text", "both", "nothing"])
                mine.append((HOST, {"doc": given["doc"], "kind": kind, "target": target.split("..")[0],
                                    "text": f"Added by the application {k}.{c}."}))
            else:
                mine.append((name, arguments))
        out.append(mine)
    return out


def host_edit(session: Any, doc: str, kind: str, target: str, text: str) -> Result:
    """The application's edit: under ``session.lock``, through the library."""
    with session.lock([doc]) as [entry]:
        document = entry.document
        try:
            if kind == "nothing":
                document.paragraphs()
                said = "held the lock"
            elif kind == "both":
                with document.batch():
                    document.paragraph(target).insert_after(text)
                    document.paragraph(target).set_text(text.replace("Added", "Set"))
                said = "inserted and set, one step"
            elif kind == "insert_after":
                document.paragraph(target).insert_after(text)
                said = "inserted"
            else:
                document.paragraph(target).set_text(text.replace("Added", "Set"))
                said = "set"
        except Exception as exc:  # noqa: BLE001 -- the application's own failure, said as it is
            said = f"failed: {type(exc).__name__}"
        result = Result(summary=said)
    result.doc, result.version, result.tool = doc, entry.version, HOST
    return result


class HostToolbox(Toolbox):
    """A toolbox whose ``dispatch`` runs :data:`HOST` as :func:`host_edit`."""

    def dispatch(self, session: Any, name: str, arguments: Any = None, **options: Any) -> Result:
        if name == HOST:
            return host_edit(session, **arguments)
        return super().dispatch(session, name, arguments, **options)


def run(seed: int, threads: int, calls: int, **options: Any) -> dict:
    """:func:`block_stress.run` with :func:`with_host_edits`'s calls; also how many application
    edits there were (``host``), and how many drafts conflicted for a change made outside the
    tools (``outside``: none, now that letting the lock go commits it)."""
    original_plan, original_box, original_normal = block_stress.plan, block_stress._box, block_stress.normal
    counted = {"host": 0, "outside": 0}

    def planned(*args: Any, **kwargs: Any) -> list[list[tuple[str, dict]]]:
        return with_host_edits(original_plan(*args, **kwargs), seed)

    def boxed(mode: str) -> Toolbox:
        made = original_box(mode)
        made.__class__ = HostToolbox
        return made

    def counting(result: Any) -> dict:
        found = original_normal(result)
        counted["host"] += result.tool == HOST
        counted["outside"] += sum(reason == "the document changed outside the tools"
                                  for reason in (result.block or {}).get("conflicts") or ())
        return found

    block_stress.plan, block_stress._box, block_stress.normal = planned, boxed, counting
    try:
        found = block_stress.run(seed, threads, calls, **options)
    finally:
        block_stress.plan, block_stress._box, block_stress.normal = original_plan, original_box, original_normal
    return {**found, **counted}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--seeds", default="0-4")
    parser.add_argument("--threads", type=int, default=6)
    parser.add_argument("--calls", type=int, default=12)
    parser.add_argument("--chapters", type=int, default=4)
    parser.add_argument("--mode", default="always")
    args = parser.parse_args()
    failed = 0
    for seed in block_stress._seeds(args.seeds):
        started = time.perf_counter()
        found = run(seed, args.threads, args.calls, chapters=args.chapters, mode=args.mode,
                    first_fill=seed % 2 == 1)
        print(json.dumps({"seed": seed, "s": round(time.perf_counter() - started, 1),
                          **{k: v for k, v in found.items() if k != "differ"}}), flush=True)
        if not found["same"]:
            failed += 1
            print(json.dumps(found["differ"], indent=1), flush=True)
    sys.exit(1 if failed else 0)
