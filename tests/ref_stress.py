"""Refs by several callers in the serial-replay stress (round 47): ``block_stress.run`` with each
thread's calls naming what they insert and using refs -- its own, a name every thread gives
(the first to define it owns it; the others are refused and may use it), and names nobody
gave yet.  Every result, the bytes and the id state must equal the calls run one at a time in
the order their changes were committed, as in ``block_stress``.  Not collected itself:
``test_ref_owners_serial.py`` runs it small, and ``python tests/ref_stress.py --seeds 0-19
--threads 8 --calls 16`` runs it long.
"""

from __future__ import annotations

import argparse
import json
import random
import sys
import time
from typing import Any

import block_stress

#: Names every thread gives: owned by whichever defines them first.
SHARED = ("note", "spot")


def with_refs(plans: list[list[tuple[str, dict]]], seed: int) -> list[list[tuple[str, dict]]]:
    """``plans`` with refs: some inserts name what they make (``ref``), some places and targets
    say ``$name`` -- the thread's own name, a shared one, or another thread's."""
    out = []
    threads = len(plans)
    for k, calls in enumerate(plans):
        rng = random.Random(seed * 7919 + k)
        mine = []
        for name, arguments in calls:
            arguments = json.loads(json.dumps(arguments))
            ops = ([op for op in arguments.get("ops") or []] if name == "batch"
                   else [{"tool": name, "arguments": arguments}])
            for op in ops:
                tool, given = op["tool"], op["arguments"]
                roll = rng.random()
                if tool == "word_insert_markdown":
                    if roll < 0.45:
                        given["ref"] = rng.choice([*SHARED, f"own{k}", f"own{k}"])
                    elif roll < 0.65:
                        given["at"] = f"after:${rng.choice([*SHARED, f'own{k}', f'own{rng.randrange(threads)}'])}"
                elif tool == "word_set_text" and roll < 0.3:
                    given["target"] = f"${rng.choice([*SHARED, f'own{k}', f'own{rng.randrange(threads)}'])}"
            mine.append((name, arguments))
        out.append(mine)
    return out


def run(seed: int, threads: int, calls: int, **options: Any) -> dict:
    """:func:`block_stress.run` with :func:`with_refs`'s calls; also how many results defined
    refs (``defined``), were refused for another caller's ref (``taken``) or named a ref nobody
    had defined then (``missing``), the concurrent run's and the replay's counted together."""
    original = block_stress.plan
    counted = {"defined": 0, "taken": 0, "missing": 0}

    def planned(*args: Any, **kwargs: Any) -> list[list[tuple[str, dict]]]:
        return with_refs(original(*args, **kwargs), seed)

    normal = block_stress.normal

    def counting(result: Any) -> dict:
        found = normal(result)
        text = json.dumps(found)
        counted["taken"] += text.count("only its definer points it elsewhere")
        counted["missing"] += text.count("no ref $")
        counted["defined"] += bool(result.refs)
        return found

    block_stress.plan = planned
    block_stress.normal = counting
    try:
        found = block_stress.run(seed, threads, calls, **options)
    finally:
        block_stress.plan = original
        block_stress.normal = normal
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
