"""The layouts a render worker keeps (``docx_agent.layout._RECENT``) used from
several threads at once, as an in-process runner's threads use them.

With the GIL off (Python 3.15t) the unlocked dict's eviction raised from two threads at
once, and two tasks handed the same pickle got the same layout and the same
``InstalledFonts``, which laying out and drawing write to.  Eight threads here take and
put back layouts of more pickles than are kept, for about a second, and no two may hold
the same objects at once.  Under a GIL build it passes either way, and quickly.
"""

from __future__ import annotations

import pickle
import threading
import time

from docx_agent import layout as render  # where the kept layouts live (round 33)

THREADS = 8
SECONDS = 1.0
CALLS = 20_000


def test_no_two_tasks_hold_the_same_layout(monkeypatch):
    monkeypatch.setattr(render, "_RECENT", {})
    blobs = [pickle.dumps((f"layout {k}", {"fonts": k})) for k in range(render._RECENT_SIZE + 2)]
    guard = threading.Lock()
    held: set[int] = set()
    errors: list[str] = []
    start = threading.Barrier(THREADS)

    def run(seed: int) -> None:
        start.wait()
        deadline = time.monotonic() + SECONDS
        for k in range(CALLS):
            if k % 4 == 0 and time.monotonic() > deadline:
                return
            blob = blobs[(seed + k) % len(blobs)]
            try:
                laid = render._taken(blob)
                assert laid == pickle.loads(blob)
                with guard:
                    shared = id(laid) in held or id(laid[1]) in held
                    held.update((id(laid), id(laid[1])))
                if shared:
                    errors.append("two tasks hold one layout")
                    return
                with guard:
                    held.difference_update((id(laid), id(laid[1])))
                render._remember(blob, laid)
            except Exception as error:  # noqa: BLE001 -- any error is the failure
                errors.append(repr(error))
                return

    threads = [threading.Thread(target=run, args=(seed,)) for seed in range(THREADS)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    assert errors == []
    assert len(render._RECENT) <= render._RECENT_SIZE
