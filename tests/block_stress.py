"""Serializability of edits on drafts (round 45, lane B): a randomized stress run and its serial
replay.  Not collected itself: ``test_block_edits_serial.py`` runs it small, and
``python tests/block_stress.py --seeds 0-49 --threads 8 --calls 20`` runs it long.

N threads, each a caller (a session view with its own context), make M calls each on one
made-up report: edits of their own chapter and of others' (the same paragraphs, tables and
places, so changes overlap), batches with reads after their last change, calls that keep the
document's lock (formatting, fields), undo of one's own change, reads and checks -- with
``block_edits="always"``, so every admitted edit is prepared on a draft.  The order changes
were committed in is recorded (``DocumentEntry.note_change``, under the document's lock).

Then the same calls run one at a time on a fresh session, in that order: each changing call in
its turn, and each call that changed nothing (a read, a refusal, a no-op) right after the commit
that made the version it saw.  Every call's result (less what is about time: ``duration``,
``data.timing``, a check's reflow, which is there when the layout before was) and the final
bytes and id state must be the same.
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import functools
import hashlib
import json
import random
import sys
import threading
import time
from collections import Counter
from typing import Any

from ooxml_edit.tools import Dataset, Result, Toolbox, string, tool
from ooxml_edit.tools.limits import Limits
from ooxml_edit.tools import DocumentEntry

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

CLOCK = lambda: dt.datetime(2026, 11, 9, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
WORDS = ("lantern meadow river stone harbour quiet orchard copper valley signal morning ledger "
         "window garden timber bridge market evening canvas thread pebble summit").split()
COLUMNS = ("code", "name", "state")


def sentence(rng: random.Random, n: int = 8) -> str:
    return " ".join(rng.choice(WORDS) for _ in range(n)).capitalize() + "."


@tool("ds_rows", "Test only: keep a made-up dataset.", {"id": string("Its id."), "rows": string("count")},
      documents=())
def ds_rows(call: Any, id: str, rows: str) -> Result:
    rng = random.Random(int(rows))
    values = [(f"K-{k + 1:03d}", sentence(rng, 2), rng.choice(["open", "closed"])) for k in range(int(rows))]
    call.define_dataset(id, Dataset(COLUMNS, values), source=None)
    return Result(summary=f"kept d:{id}")


@functools.lru_cache(maxsize=4)
def made(chapters: int = 4) -> bytes:
    """A made-up report: per chapter a heading, sections, paragraphs, a bullet and a numbered
    list, a small table, an end paragraph; a table to fill; a paragraph with a value written
    by reference; a bookmark and a cross-reference; a comment over two paragraphs."""
    rng = random.Random(45)
    with Toolbox([*TOOLS, ds_rows], formats=[FORMAT], groups=GROUPS, workers=0, block_edits="off") as box:
        session = box.session(clock=CLOCK)
        document = Document.new()
        document.paragraphs()[0].set_text("Report of made-up things")
        doc = session.open(document.to_bytes(), "report.docx")

        def call(name: str, **arguments: Any) -> dict:
            found = box.dispatch(session, name, {"doc": doc, **arguments} if name != "ds_rows" else arguments)
            if not found.ok:
                raise AssertionError(json.dumps(found.to_json())[:2000])
            return found.to_json()

        call("ds_rows", id="rows", rows="6")
        for n in range(chapters):
            parts = [f"# Chapter {n + 1} {sentence(rng, 2)[:-1]}"]
            for s in range(2):
                parts.append(f"## Section {n + 1}.{s + 1}")
                parts += [" ".join(sentence(rng, 12) for _ in range(2)) for _ in range(2)]
            parts.append(f"- {sentence(rng, 4)}\n- {sentence(rng, 5)}\n  - {sentence(rng, 3)}")
            parts.append(f"1. {sentence(rng, 4)}\n2. {sentence(rng, 3)}")
            parts.append("| Item | Value |\n|---|---|\n" + "\n".join(f"| {sentence(rng, 2)[:-1]} | {k} |"
                                                                    for k in range(3)))
            parts.append(f"End of chapter {n + 1}.")
            call("word_insert_markdown", at="end", markdown="\n\n".join(parts))
        call("word_insert_markdown", at="end", markdown="Rows in full\n\n| code | name | state |\n|---|---|---|\n| x | y | z |")
        entry = session.entry(doc)
        paragraphs = entry.document.paragraphs()
        sections = [p.id for p in paragraphs if p.text.startswith("Section ")]
        call("word_links", action="add_bookmark", target=sections[0], bookmark="first_section")
        ends = [p for p in paragraphs if p.text.startswith("End of chapter ")]
        call("word_fields", action="insert_cross_reference", at=f"{ends[0].id}@{len(ends[0].text)}",
             bookmark="first_section", kind="text")
        call("word_set_text", target=[p.id for p in paragraphs if len(p.text) > 60][1],
             text="A value from the rows: {{d:rows/#2/name}} and more text after it.")
        long = [p.id for p in paragraphs if len(p.text) > 60]
        call("word_comments", action="add", items=[{"target": f"{long[3]}@3..{long[4]}@9", "text": "A remark on these."}])
        return entry.document.to_bytes()


class Layout:
    """The made-up report's ids, read once: per chapter its heading, its paragraphs, its end
    paragraph and its table; the table to fill."""

    def __init__(self, document: Document) -> None:
        paragraphs = document.paragraphs()
        ids = [p.id for p in paragraphs]
        self.heads = [p.id for p in paragraphs if p.text.startswith("Chapter ")]
        self.ends = [p.id for p in paragraphs if p.text.startswith("End of chapter ")]
        self.bodies = []
        for head, end in zip(self.heads, self.ends):
            span = ids[ids.index(head):ids.index(end)]
            self.bodies.append([i for i in span if len(document.paragraph(i).text) > 40 and "/" not in i])
        tables = document.tables()
        self.fill = next(t.id for t in tables if t.cell(0, 0).text == "code")
        self.tables = [t.id for t in tables if t.cell(0, 0).text == "Item"]


def plan(seed: int, threads: int, calls: int, layout: Layout, doc: str) -> list[list[tuple[str, dict]]]:
    """Each thread's calls, from ``seed``: mostly its own chapter's, often another's."""
    out = []
    chapters = len(layout.heads)
    for k in range(threads):
        rng = random.Random(seed * 1000 + k)
        own = k % chapters
        mine: list[tuple[str, dict]] = []
        for c in range(calls):
            chapter = own if rng.random() < 0.7 else rng.randrange(chapters)
            body = layout.bodies[chapter]
            target = rng.choice(body)
            hot = layout.bodies[0][0]
            roll = rng.random()
            md = (f"### Note {k}.{c}\n\n{sentence(rng, 10)}" if rng.random() < 0.5
                  else f"- {sentence(rng, 4)}\n- {sentence(rng, 3)}" if rng.random() < 0.5
                  else f"| A | B |\n|---|---|\n| {sentence(rng, 1)[:-1]} | {c} |")
            place = rng.choice([f"before:{layout.ends[chapter]}", f"after:{target}", f"before:{target}",
                                f"after:{layout.tables[chapter]}", "end"])
            if roll < 0.30:
                ops = [{"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": place, "markdown": md}},
                       {"tool": "word_set_text", "arguments": {"doc": doc, "target": rng.choice(body),
                                                               "text": sentence(rng, 14)}}]
                if rng.random() < 0.3:
                    ops.append({"tool": "word_read", "arguments": {"doc": doc, "range": f"{layout.heads[chapter]}.."
                                                                                       f"{layout.ends[chapter]}"}})
                mine.append(("batch", {"ops": ops}))
            elif roll < 0.45:
                text = sentence(rng, 12) if rng.random() < 0.8 else f"{sentence(rng, 5)}\n{sentence(rng, 4)}"
                mine.append(("word_set_text", {"doc": doc, "target": hot if rng.random() < 0.2 else target,
                                               "text": text}))
            elif roll < 0.55:
                mine.append(("word_insert_markdown", {"doc": doc, "at": place, "markdown": md}))
            elif roll < 0.60:
                mine.append(("word_edit_table", {"doc": doc, "table": layout.tables[chapter], "action": "set_cells",
                                                 "cells": [{"row": 1, "col": 1, "text": f"{k}.{c}"}]}))
            elif roll < 0.64:
                mine.append(("word_edit_table", {"doc": doc, "table": layout.fill, "action": "set_cells",
                                                 "grow": True, "from": "d:rows"}))
            elif roll < 0.68:
                mine.append(("word_set_text", {"doc": doc, "target": target,
                                               "text": f"From the rows: {{{{d:rows/#{1 + c % 5}/name}}}}."}))
            elif roll < 0.76:
                mine.append(("word_read", {"doc": doc, "range": f"{layout.heads[chapter]}..{layout.ends[chapter]}"}))
            elif roll < 0.80:
                mine.append(("check", {"doc": doc, "range": [f"{layout.heads[chapter]}..{layout.ends[chapter]}"]}))
            elif roll < 0.86:
                mine.append(("word_format", {"doc": doc, "targets": [target], "bold": rng.random() < 0.5}))
            elif roll < 0.89:
                mine.append(("undo", {"doc": doc}))
            elif roll < 0.92:
                mine.append(("word_set_text", {"doc": doc, "target": "p:7FFFFFF0", "text": "gone"}))
            else:
                mine.append(("describe", {"doc": doc}))
        out.append(mine)
    return out


def normal(result: Result) -> dict:
    """What a result says that is not about time."""
    out = copy.deepcopy(result.to_json())
    data = out.get("data")
    if isinstance(data, dict):
        data.pop("timing", None)
        for op in data.get("ops") or []:
            if isinstance(op.get("data"), dict):
                op["data"].pop("timing", None)
    checks = out.get("checks")
    if isinstance(checks, dict):
        checks.pop("reflow", None)
    out["warnings"] = [w for w in out.get("warnings") or [] if "changed while this call read it" not in w]
    out["touched"] = result.touched
    out["touched_by_op"] = {str(op): found for op, found in result.touched_by_op.items()}
    return out


class Recorder:
    """The commits, in order, by the call that made each (``note_change`` runs under the lock)."""

    def __init__(self) -> None:
        self.local = threading.local()
        self.commits: list[tuple[int, int]] = []     # (call, version after)
        self.lock = threading.Lock()
        self._original = DocumentEntry.note_change

        recorder = self

        def note_change(entry: Any, parts: Any, before: Any) -> None:
            recorder._original(entry, parts, before)
            call = getattr(recorder.local, "call", None)
            if call is not None:
                with recorder.lock:
                    recorder.commits.append((call, entry.version))

        self._patched = note_change

    def __enter__(self) -> "Recorder":
        DocumentEntry.note_change = self._patched  # type: ignore[method-assign]
        return self

    def __exit__(self, *exc: Any) -> None:
        DocumentEntry.note_change = self._original  # type: ignore[method-assign]


def _box(mode: str) -> Toolbox:
    return Toolbox([*TOOLS, ds_rows], formats=[FORMAT], groups=GROUPS, workers=0, block_edits=mode,
                   limits=Limits(render_timeout=600, layout_timeout=600, batch_timeout=600))


def run(seed: int, threads: int, calls: int, chapters: int = 4, mode: str = "always") -> dict:
    """One stress run and its replay; ``{"same": bool, ...}`` with what differed."""
    data = made(chapters)
    layout = Layout(Document.open(data))
    with _box(mode) as box, Recorder() as recorder:
        session = box.session(clock=CLOCK)
        doc = session.open(data, "report.docx")
        box.dispatch(session, "ds_rows", {"id": "rows", "rows": "6"})
        plans = plan(seed, threads, calls, layout, doc)
        made_calls: list[dict] = []
        lock = threading.Lock()
        errors: list[str] = []

        def worker(k: int) -> None:
            view = session.view(context={"agent": f"b{k}"})
            for c, (name, arguments) in enumerate(plans[k]):
                number = k * 10_000 + c
                with lock:
                    start = len(recorder.commits)
                recorder.local.call = number
                try:
                    result = box.dispatch(view, name, arguments)
                except Exception as exc:  # noqa: BLE001
                    errors.append(f"{name}: {exc!r}")
                    continue
                finally:
                    recorder.local.call = None
                with lock:
                    end = len(recorder.commits)
                made_calls.append({"call": number, "thread": k, "name": name, "arguments": arguments,
                                   "result": normal(result), "version": result.version, "start": start,
                                   "end": end, "block": result.block})
                time.sleep(0)

        workers = [threading.Thread(target=worker, args=(k,)) for k in range(threads)]
        for thread in workers:
            thread.start()
        for thread in workers:
            thread.join()
        final = session.entry(doc).document.to_bytes()
        state = session.entry(doc).document.id_state()
        commits = list(recorder.commits)
    order = _order(made_calls, commits)
    replayed = replay(data, order, threads)
    differ = []
    for item, again in zip(order, replayed["results"]):
        if item["result"] != again:
            differ.append({"call": item["call"], "name": item["name"], "concurrent": item["result"],
                           "serial": again})
    same = not differ and not errors and final == replayed["final"] and state == replayed["state"]
    drafted = [c for c in made_calls if c["block"]]
    return {"same": same, "differ": differ[:3], "errors": errors[:5], "bytes_same": final == replayed["final"],
            "state_same": state == replayed["state"], "calls": len(made_calls), "commits": len(commits),
            "drafted": len(drafted),
            "conflicts": sum(len(c["block"].get("conflicts") or []) for c in drafted),
            "locked": sum(1 for c in drafted if c["block"].get("lock")),
            "why": dict(Counter([reason.split(" near ")[0] for c in drafted for reason in c["block"].get("conflicts") or []]
                                + ["lock: " + c["block"]["lock"] for c in drafted if c["block"].get("lock")])),
            "sha": hashlib.sha256(final).hexdigest()[:16]}


def _order(calls: list[dict], commits: list[tuple[int, int]]) -> list[dict]:
    """The serial order: each commit's call in turn; each call that committed nothing right
    after the last commit before it ended that made the version it saw (before every commit
    when it saw the version the run began with)."""
    by_call = {item["call"]: item for item in calls}
    committing = {call for call, _ in commits}
    after: dict[int, list[dict]] = {}
    for item in sorted(calls, key=lambda c: c["call"]):
        if item["call"] in committing:
            continue
        at = -1
        for k in range(min(item["end"], len(commits)) - 1, -1, -1):
            if commits[k][1] == item["version"]:
                at = k
                break
        after.setdefault(at, []).append(item)
    out = list(after.get(-1, []))
    seen: set[int] = set()
    for k, (call, _version) in enumerate(commits):
        if call not in seen:
            seen.add(call)
            out.append(by_call[call])
        out += after.get(k, [])
    return out


def replay(data: bytes, order: list[dict], threads: int) -> dict:
    """The calls one at a time, in ``order``, each by its own caller, on a fresh session."""
    with _box("off") as box:
        session = box.session(clock=CLOCK)
        doc = session.open(data, "report.docx")
        box.dispatch(session, "ds_rows", {"id": "rows", "rows": "6"})
        views = {k: session.view(context={"agent": f"b{k}"}) for k in range(threads)}
        results = [normal(box.dispatch(views[item["thread"]], item["name"], item["arguments"])) for item in order]
        entry = session.entry(doc)
        return {"results": results, "final": entry.document.to_bytes(), "state": entry.document.id_state()}


def _seeds(text: str) -> list[int]:
    first, _, last = text.partition("-")
    return list(range(int(first), int(last or first) + 1))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--seeds", default="0-4")
    parser.add_argument("--threads", type=int, default=6)
    parser.add_argument("--calls", type=int, default=12)
    parser.add_argument("--chapters", type=int, default=4)
    parser.add_argument("--mode", default="always")
    args = parser.parse_args()
    failed = 0
    for seed in _seeds(args.seeds):
        started = time.perf_counter()
        found = run(seed, args.threads, args.calls, args.chapters, args.mode)
        print(json.dumps({"seed": seed, "s": round(time.perf_counter() - started, 1),
                          **{k: v for k, v in found.items() if k not in ("differ",)}}), flush=True)
        if not found["same"]:
            failed += 1
            print(json.dumps(found["differ"], indent=1)[:6000], flush=True)
    sys.exit(1 if failed else 0)
