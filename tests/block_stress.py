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

The version a call that changed nothing saw is the one its result describes (``version``).
docx-agent 0.1.112 serves a check from its document's next layout, which can be of a version
committed while the check waited; its result then describes that version, and says so
(``of version V: N commit(s) newer than at the call (version v), ...``).  Such a check is
placed after the commit that made V, like any read of V, and compared there whole -- but for
that note, which a call run alone never has: the note is checked against the commits recorded
instead (:func:`read_problems`): some commit made v, no earlier than the version the call
began at, and V is the N-th commit after it.  Every call that changed nothing must have seen
the version it began at or one committed while it ran (never one already replaced then: an
undo brings an earlier version number back), and the serial run's results carry no such note.
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import functools
import hashlib
import io
import json
import random
import re
import sys
import threading
import time
import zipfile
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
def made(chapters: int = 4, *, first_fill: bool = False) -> bytes:
    """A made-up report: per chapter a heading, sections, paragraphs, a bullet and a numbered
    list, a small table, an end paragraph; a table to fill; a paragraph with a value written
    by reference; a bookmark and a cross-reference; a comment over two paragraphs.
    ``first_fill``: the paragraph's text written plainly, so the document keeps no provenance
    yet -- its first fill or value by reference adds the part."""
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
             text=("A value from the rows: " + ("a name" if first_fill else "{{d:rows/#2/name}}")
                   + " and more text after it."))
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


def made_short(parts: int = 10, *, first_fill: bool = False) -> bytes:
    """A report as callers start one (round 47): ``parts`` headings with no paragraph or one
    short one between them, a table to fill at the end, and -- unless ``first_fill`` -- a
    value written by reference (so provenance is kept)."""
    rng = random.Random(47)
    with Toolbox([*TOOLS, ds_rows], formats=[FORMAT], groups=GROUPS, workers=0, block_edits="off") as box:
        session = box.session(clock=CLOCK)
        document = Document.new()
        document.paragraphs()[0].set_text("Report of made-up things")
        doc = session.open(document.to_bytes(), "report.docx")

        def call(name: str, **arguments: Any) -> None:
            found = box.dispatch(session, name, {"doc": doc, **arguments} if name != "ds_rows" else arguments)
            if not found.ok:
                raise AssertionError(json.dumps(found.to_json())[:2000])

        call("ds_rows", id="rows", rows="6")
        lines = [f"- {sentence(rng, 3)}\n- {sentence(rng, 3)}"]      # a list: the numbering part there
        for n in range(parts):
            lines.append(f"# Part {n + 1} {sentence(rng, 2)[:-1]}")
            if n % 2 == 0:
                lines.append(sentence(rng, 6))
        lines.append("Rows in full\n\n| code | name | state |\n|---|---|---|\n| x | y | z |")
        call("word_insert_markdown", at="end", markdown="\n\n".join(lines))
        if not first_fill:
            first = next(p for p in session.entry(doc).document.paragraphs() if p.text.startswith("Part 1 "))
            call("word_insert_markdown", at=f"after:{first.id}", markdown="A value: {{d:rows/#2/name}}.")
        return session.entry(doc).document.to_bytes()


class ShortLayout:
    """:func:`made_short`'s ids: each part's heading and the paragraphs after it."""

    def __init__(self, document: Document) -> None:
        paragraphs = document.paragraphs()
        ids = [p.id for p in paragraphs]
        self.heads = [p.id for p in paragraphs if p.text.startswith("Part ")]
        rows = next(p.id for p in paragraphs if p.text == "Rows in full")
        bounds = self.heads[1:] + [rows]
        self.bodies = [ids[ids.index(head) + 1:ids.index(bound)] for head, bound in zip(self.heads, bounds)]
        self.bounds = bounds
        self.fill = next(t.id for t in document.tables() if t.cell(0, 0).text == "code")


def plan_short(seed: int, threads: int, calls: int, layout: ShortLayout, doc: str) -> list[list[tuple[str, dict]]]:
    """The team's shape (round 47): each thread writes into its own part of a report whose
    parts are still all but empty -- at its heading, before the next one, after what it
    wrote -- with checks and reads of its part between, now and then a fill or an edit of
    the paragraph a part began with."""
    out = []
    parts = len(layout.heads)
    for k in range(threads):
        rng = random.Random(seed * 1000 + k)
        own = k % parts
        head, bound, body = layout.heads[own], layout.bounds[own], layout.bodies[own]
        mine: list[tuple[str, dict]] = []
        for c in range(calls):
            md = (f"### Note {k}.{c}\n\n{sentence(rng, 10)}" if rng.random() < 0.5
                  else f"- {sentence(rng, 4)}\n- {sentence(rng, 3)}" if rng.random() < 0.5
                  else f"| A | B |\n|---|---|\n| {sentence(rng, 1)[:-1]} | {c} |")
            place = rng.choice([f"after:{head}", f"before:{bound}"] + ([f"after:{body[-1]}"] if body else []))
            roll = rng.random()
            if roll < 0.25:
                mine.append(("word_insert_markdown", {"doc": doc, "at": place, "markdown": md}))
            elif roll < 0.45:
                # A table of a dataset's rows, its source kept (the provenance item).
                mine.append(("word_insert_markdown", {"doc": doc, "at": place, "from": "d:rows"}))
            elif roll < 0.60:
                ops = [{"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": place, "markdown": md}},
                       {"tool": "word_read", "arguments": {"doc": doc, "range": f"{head}..{bound}"}}]
                mine.append(("batch", {"ops": ops}))
            elif roll < 0.66:
                # The round 48 mix: edits of the part's blocks in one batch.
                if body:
                    mine.append(("batch", {"ops": [block_edit(rng, doc, f"{head}..{bound}", body[0], None)
                                                   for _ in range(rng.choice([1, 2, 3]))]}))
                else:
                    mine.append(("word_format", {"doc": doc, "targets": [head], "italic": True}))
            elif roll < 0.72:
                mine.append(("check", {"doc": doc, "range": [f"{head}..{bound}"]}))
            elif roll < 0.80:
                mine.append(("word_read", {"doc": doc, "range": f"{head}..{bound}"}))
            elif roll < 0.88 and body:
                mine.append(("word_set_text", {"doc": doc, "target": body[0], "text": sentence(rng, 12)}))
            elif roll < 0.92:
                mine.append(("word_edit_table", {"doc": doc, "table": layout.fill, "action": "set_cells",
                                                 "grow": True, "from": "d:rows"}))
            elif roll < 0.95:
                # A late failure: blocks made (ids drawn), then an op that fails, all undone.
                mine.append(("batch", {"atomic": True, "ops": [
                    {"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": place, "markdown": md}},
                    {"tool": "word_set_text", "arguments": {"doc": doc, "target": "p:7FFFFFF0", "text": "x"}}]}))
            else:
                mine.append(("word_insert_markdown", {"doc": doc, "at": f"before:{bound}",
                                                      "markdown": f"{sentence(rng, 6)} {{{{d:rows/#{1 + c % 5}/name}}}}"}))
        out.append(mine)
    return out


def block_edit(rng: random.Random, doc: str, span: str, target: str, table: str | None) -> dict:
    """One edit of known blocks (round 48, R83) as a batch op: text replaced in a span, a
    paragraph, part of one, a span or found text formatted, a table formatted, text or a
    paragraph deleted, a link or a bookmark added -- the mix callers send in one batch."""
    word = rng.choice(WORDS)
    kind = rng.random()
    if kind < 0.2:
        return {"tool": "replace_text", "arguments": {"doc": doc, "find": word, "replace": word.upper(),
                                                      "expect": "all", "range": span}}
    if kind < 0.45:
        how = rng.random()
        where = ({"targets": [target]} if how < 0.4 else {"targets": [f"{target}@0:5"]} if how < 0.6
                 else {"targets": [span]} if how < 0.8 else {"find": [{"text": word, "within": span, "expect": "all"}]})
        return {"tool": "word_format", "arguments": {"doc": doc, **where, rng.choice(["bold", "italic", "underline"]):
                                                     rng.random() < 0.7}}
    if kind < 0.6 and table is not None:
        return {"tool": "word_format_table", "arguments": {"doc": doc, "table": table, "scope": "table",
                                                           **rng.choice([{"width_of": "body"}, {"shading": "DDEEFF"},
                                                                         {"borders": "single"}, {"header_rows": 1}])}}
    if kind < 0.75:
        return {"tool": "word_delete", "arguments": {"doc": doc, "find": word, "within": target}
                if rng.random() < 0.7 else {"doc": doc, "target": f"{target}@0:3"}}
    if kind < 0.9:
        return {"tool": "word_links", "arguments": {"doc": doc, "action": "add_hyperlink", "target": f"{target}@0:4",
                                                    **({"url": "https://example.com/a"} if rng.random() < 0.6
                                                       else {"bookmark": "first_section"})}}
    return {"tool": "word_links", "arguments": {"doc": doc, "action": "add_bookmark", "target": f"{target}@1:6",
                                                "bookmark": f"mark_{rng.randrange(10_000)}"}}


def plan(seed: int, threads: int, calls: int, layout: Layout, doc: str, sync: bool = False
         ) -> list[list[tuple[str, dict]]]:
    """Each thread's calls, from ``seed``: mostly its own chapter's, often another's.
    ``sync`` (round 49): a quarter of the calls keep the table to fill in step by key instead
    -- from one of two datasets, its rows gone listed or removed."""
    out = []
    chapters = len(layout.heads)
    for k in range(threads):
        rng = random.Random(seed * 1000 + k)
        # Round 48's edits of known blocks drawn from their own sequence: the calls as before.
        extra = random.Random((seed * 1000 + k) * 31 + 7)
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
            if sync and extra.random() < 0.25:
                mine.append(("word_edit_table", {"doc": doc, "table": layout.fill, "action": "sync",
                                                 "from": extra.choice(["d:rows", "d:other"]),
                                                 "remove": extra.random() < 0.5}))
                continue
            place = rng.choice([f"before:{layout.ends[chapter]}", f"after:{target}", f"before:{target}",
                                f"after:{layout.tables[chapter]}", "end"])
            if roll < 0.30:
                ops = [{"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": place, "markdown": md}},
                       {"tool": "word_set_text", "arguments": {"doc": doc, "target": rng.choice(body),
                                                               "text": sentence(rng, 14)}}]
                span = f"{layout.heads[chapter]}..{layout.ends[chapter]}"
                for _ in range(extra.choice([0, 1, 2, 3])):
                    ops.append(block_edit(extra, doc, span, extra.choice(body), layout.tables[chapter]))
                if rng.random() < 0.3:
                    ops.append({"tool": "word_read", "arguments": {"doc": doc, "range": span}})
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
                if rng.random() < 0.5:
                    mine.append(("word_set_text", {"doc": doc, "target": "p:7FFFFFF0", "text": "gone"}))
                else:
                    # A late failure: blocks made (ids drawn), then an op that fails, all undone.
                    mine.append(("batch", {"atomic": True, "ops": [
                        {"tool": "word_insert_markdown", "arguments": {"doc": doc, "at": place, "markdown": md}},
                        {"tool": "word_set_text", "arguments": {"doc": doc, "target": "p:7FFFFFF0", "text": "x"}}]}))
            else:
                mine.append(("describe", {"doc": doc}))
        out.append(mine)
    return out


#: The note of a check served a version committed after it was called (docx-agent 0.1.112).
NEWER = re.compile(r"^of version (\d+): (\d+) commit\(s\) newer than at the call \(version (\d+)\), ")


def newer_note(result: Result) -> tuple[int, int, int] | None:
    """``(V, N, v)`` of ``result``'s :data:`NEWER` note, or ``None``."""
    for warning in result.warnings:
        found = NEWER.match(warning)
        if found:
            return int(found.group(1)), int(found.group(2)), int(found.group(3))
    return None


def normal(result: Result) -> dict:
    """What a result says that is not about time (a version note, :data:`NEWER`, is about
    when the call came: checked apart, :func:`read_problems`)."""
    out = copy.deepcopy(result.to_json())
    data = out.get("data")
    if isinstance(data, dict):
        data.pop("timing", None)
        if isinstance(data.get("reflow"), dict):
            # What a check says changed since the last check -- whichever ran last, in time.
            data["reflow"].pop("since_last_check", None)
        for op in data.get("ops") or []:
            if isinstance(op.get("data"), dict):
                op["data"].pop("timing", None)
    checks = out.get("checks")
    if isinstance(checks, dict):
        checks.pop("reflow", None)
    out["warnings"] = [w for w in out.get("warnings") or [] if "changed while this call read it" not in w
                       and not NEWER.match(w)]
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

        def note_change(entry: Any, parts: Any, before: Any, *more: Any) -> None:
            recorder._original(entry, parts, before, *more)
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


def run(seed: int, threads: int, calls: int, chapters: int = 4, mode: str = "always",
        first_fill: bool = False, shape: str = "report", sync: bool = False, tracked: bool = False) -> dict:
    """One stress run and its replay; ``{"same": bool, ...}`` with what differed.
    ``first_fill``: on a report with no provenance yet (:func:`made`).  ``shape``:
    ``"report"`` (:func:`made`, :func:`plan`) or ``"short"`` (:func:`made_short`, one part per
    thread, :func:`plan_short`).  ``sync``: :func:`plan`'s syncs (round 49), from ``d:rows``
    and ``d:other``.  ``tracked`` (round 50): every edit a tracked change (:data:`AUTHOR`)."""
    if shape == "short":
        data = made_short(threads, first_fill=first_fill)
        layout: Any = ShortLayout(Document.open(data))
    else:
        data = made(chapters, first_fill=first_fill)
        layout = Layout(Document.open(data))
    with _box(mode) as box, Recorder() as recorder:
        session = box.session(clock=CLOCK)
        doc = session.open(data, "report.docx")
        box.dispatch(session, "ds_rows", {"id": "rows", "rows": "6"})
        if sync:
            box.dispatch(session, "ds_rows", {"id": "other", "rows": "8"})
        if tracked:
            track(box, session, doc)
        plans = (plan_short(seed, threads, calls, layout, doc) if shape == "short"
                 else plan(seed, threads, calls, layout, doc, sync=sync))
        initial = session.entry(doc).version
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
                                   "end": end, "block": result.block, "note": newer_note(result)})
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
    errors += read_problems(made_calls, commits, initial)
    replayed = replay(data, order, threads, sync=sync, tracked=tracked)
    errors += [f"call {item['call']} ({item['name']}) run alone says it was served a later version"
               for item, note in zip(order, replayed["notes"]) if note is not None]
    differ = []
    for item, again in zip(order, replayed["results"]):
        if item["result"] != again:
            differ.append({"call": item["call"], "name": item["name"], "concurrent": item["result"],
                           "serial": again})
    repeated = repeated_ids(final)
    same = (not differ and not errors and final == replayed["final"] and state == replayed["state"]
            and not repeated)
    drafted = [c for c in made_calls if c["block"]]
    return {"same": same, "differ": differ[:3], "errors": errors[:5], "bytes_same": final == replayed["final"],
            "state_same": state == replayed["state"], "repeated_ids": repeated[:5],
            "revisions": revision_count(final), "calls": len(made_calls), "commits": len(commits),
            "drafted": len(drafted),
            "conflicts": sum(len(c["block"].get("conflicts") or []) for c in drafted),
            "locked": sum(1 for c in drafted if c["block"].get("lock")),
            "why": dict(Counter([reason.split(" near ")[0] for c in drafted for reason in c["block"].get("conflicts") or []]
                                + ["lock: " + c["block"]["lock"] for c in drafted if c["block"].get("lock")])),
            "sha": hashlib.sha256(final).hexdigest()[:16]}


_ANNOTATION_ID = re.compile(rb'<w:(bookmarkStart|ins|del|moveFrom|moveTo|rPrChange|pPrChange|tblPrChange|trPrChange'
                            rb'|tcPrChange|tblGridChange|sectPrChange|numberingChange|cellIns|cellDel)\b[^>]*?\bw:id="(\d+)"')


def _annotation_ids(data: bytes) -> list[tuple[bytes, bytes]]:
    """``(tag, id)`` of every annotation (a bookmark's start, a revision) in ``data``'s parts."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return [found for name in archive.namelist() if name.endswith(".xml")
                for found in _ANNOTATION_ID.findall(archive.read(name))]


def repeated_ids(data: bytes) -> list[str]:
    """Annotation ids (revisions', bookmarks') written more than once in ``data``'s parts."""
    counts = Counter(value for _tag, value in _annotation_ids(data))
    return sorted(value.decode() for value, count in counts.items() if count > 1)


def revision_count(data: bytes) -> int:
    """The revision records in ``data``'s parts."""
    return sum(1 for tag, _value in _annotation_ids(data) if tag != b"bookmarkStart")


def _placed(item: dict, commits: list[tuple[int, int]]) -> int:
    """The index of the commit a call that committed nothing is placed after: the last one
    before it ended that made the version it saw (``-1``: none, the version the run began
    with)."""
    for k in range(min(item["end"], len(commits)) - 1, -1, -1):
        if commits[k][1] == item["version"]:
            return k
    return -1


def read_problems(calls: list[dict], commits: list[tuple[int, int]], initial: int) -> list[str]:
    """What is wrong with the versions the calls that committed nothing saw (the module's
    notes): each saw the version it began at or one committed while it ran -- its placing
    commit (:func:`_placed`) is no earlier than the last one before it began -- and a
    :data:`NEWER` note is true of the commits recorded: V is the version the result describes,
    and some commit no earlier than the version the call began at made v, N commits before V's."""
    committing = {call for call, _ in commits}
    out = []
    for item in calls:
        if item["call"] in committing or item["version"] is None:
            continue                    # (a call refused before it saw a version: none to check)
        at = _placed(item, commits)
        name = f"call {item['call']} ({item['name']})"
        if at < item["start"] - 1 or (at == -1 and item["version"] != initial):
            out.append(f"{name} saw version {item['version']}, not the one it began at "
                       f"(commit {item['start']}) nor one committed while it ran")
        note = item.get("note")
        if note is None:
            continue
        described, count, at_call = note
        if described != item["version"]:
            out.append(f"{name}: its note names version {described}, its result {item['version']}")
            continue
        made = [k for k in range(max(item["start"] - 1, -1), at)
                if (commits[k][1] if k >= 0 else initial) == at_call]
        if not any(at - k == count for k in made):
            out.append(f"{name}: of version {described}, {count} commit(s) newer than version {at_call} at the "
                       f"call -- the commits recorded say otherwise ({[commits[k][1] for k in range(max(at - count - 1, 0), at + 1)]})")
    return out


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


#: The author of every edit in a tracked run (round 50).
AUTHOR = "Stress Writer"


def track(box: Toolbox, session: Any, doc: str) -> None:
    """Tracking on for ``doc``: every edit after it a tracked change by :data:`AUTHOR`."""
    assert box.dispatch(session, "word_set_tracking", {"doc": doc, "on": True, "author": AUTHOR}).ok


def replay(data: bytes, order: list[dict], threads: int, sync: bool = False, tracked: bool = False) -> dict:
    """The calls one at a time, in ``order``, each by its own caller, on a fresh session."""
    with _box("off") as box:
        session = box.session(clock=CLOCK)
        doc = session.open(data, "report.docx")
        box.dispatch(session, "ds_rows", {"id": "rows", "rows": "6"})
        if sync:
            box.dispatch(session, "ds_rows", {"id": "other", "rows": "8"})
        if tracked:
            track(box, session, doc)
        views = {k: session.view(context={"agent": f"b{k}"}) for k in range(threads)}
        found = [box.dispatch(views[item["thread"]], item["name"], item["arguments"]) for item in order]
        entry = session.entry(doc)
        return {"results": [normal(result) for result in found], "notes": [newer_note(result) for result in found],
                "final": entry.document.to_bytes(), "state": entry.document.id_state()}


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
    parser.add_argument("--shape", default="report", choices=("report", "short"),
                        help="report: chapters of text; short: one all but empty part per thread")
    parser.add_argument("--sync", action="store_true", help="syncs of the table to fill (round 49)")
    parser.add_argument("--tracked", action="store_true", help="every edit a tracked change (round 50)")
    parser.add_argument("--first-fill", default="odd", choices=("none", "odd", "all"),
                        help="which seeds run on a report with no provenance yet (default: odd ones)")
    args = parser.parse_args()
    failed = 0
    for seed in _seeds(args.seeds):
        started = time.perf_counter()
        first = args.first_fill == "all" or (args.first_fill == "odd" and seed % 2 == 1)
        found = run(seed, args.threads, args.calls, args.chapters, args.mode, first_fill=first, shape=args.shape,
                    sync=args.sync, tracked=args.tracked)
        print(json.dumps({"seed": seed, "first_fill": first, "s": round(time.perf_counter() - started, 1),
                          **{k: v for k, v in found.items() if k not in ("differ",)}}), flush=True)
        if not found["same"]:
            failed += 1
            print(json.dumps(found["differ"], indent=1), flush=True)
    sys.exit(1 if failed else 0)
