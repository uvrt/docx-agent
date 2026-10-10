"""Golden transcripts: record a task's tool calls, replay them, grade the output.

A transcript is JSON::

    {"task": "w2-review-comments", "clock": "2026-10-07T09:00:00+00:00",
     "open": ["hosting-agreement-summary.docx"],      # opened by the app: d1, d2...
     "blobs": ["draft.md"],                           # registered by the app: b1, b2...
     "calls": [{"tool": ..., "arguments": {...}, "expect": {...}}],
     "output": {"name": ..., "sha256": ...}}

``expect`` is the normalised result: ok, summary, changed, created, removed, renamed, refs,
warnings, the validate delta's new problems, and three things about data -- ``data_sha``, a
digest of all of it; ``data_keys``, its schema (:func:`data_keys`); and ``data_sha_core``, a
digest of what the layout does not measure (:func:`core_data`).  Replaying gives the same
results and a byte-identical output (the session clock is fixed) where Office's faces are
installed; elsewhere it gives the same results less what the layout measures
(:func:`core_mismatches`).  ``tools/refresh_goldens.py`` records the fields.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
from pathlib import Path
from typing import Any, Callable

FIELDS = ("ok", "summary", "changed", "created", "removed", "renamed", "refs", "warnings")


def _png(width: int, height: int, rgb: tuple[int, int, int]) -> bytes:
    """A plain PNG, made here: no raster is kept in the repository."""
    import struct
    import zlib

    row = b"\x00" + bytes(rgb) * width

    def chunk(kind: bytes, body: bytes) -> bytes:
        return (struct.pack(">I", len(body)) + kind + body
                + struct.pack(">I", zlib.crc32(kind + body) & 0xFFFFFFFF))

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header)
            + chunk(b"IDAT", zlib.compress(row * height, 9)) + chunk(b"IEND", b""))


#: Inputs made at replay time instead of kept as files: name -> bytes.
GENERATED = {"org-chart.png": lambda: _png(300, 200, (21, 96, 130))}


def blob_bytes(inputs: Path, name: str) -> bytes:
    path = inputs / name
    if not path.exists() and name in GENERATED:
        return GENERATED[name]()
    return path.read_bytes()


def normalise(result: Any, tool: str) -> dict[str, Any]:
    body = result.to_json()
    out = {key: body.get(key) for key in FIELDS if body.get(key) not in (None, [], {})}
    out["ok"] = body["ok"]
    if not body["ok"]:
        out["error"] = {"code": body["error"]["code"], "message": body["error"]["message"]}
    checks = body.get("checks") or {}
    validate = checks.get("validate") if isinstance(checks, dict) else None
    if validate and validate.get("new"):
        out["new_problems"] = validate["new"]
    data = body.get("data")
    if data not in (None, {}, []):
        out["data_sha"] = digest(data)
        keys = data_keys(data)
        if keys:
            out["data_keys"] = keys
        out["data_sha_core"] = digest(core_data(tool, data))
    return out


def digest(data: Any) -> str:
    return hashlib.sha256(json.dumps(data, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:16]


# -- What the layout measures ---------------------------------------------------------------
#
# The one place the expectations are split.  A transcript is recorded on a Mac with
# Office's faces; a runner without them (CI's: the open substitutes at best) lays the same
# documents out differently, or stops.  What depends on that layout is
#
# * the data fields in LAYOUT_FIELDS -- reflow and coverage facts, page counts, the pages
#   rendered, and the size of a saved document (its app.xml carries the page count): left
#   out of ``data_sha_core``, though their keys stay in ``data_keys``;
# * the keys in CONDITIONAL_KEYS, which docx-agent adds to those fields only when the
#   layout stopped, substituted or approximated (layout.coverage_facts and
#   tools/shared._layout_facts): left out of ``data_keys``, so a runner where the layout
#   stops has the schema a Mac has;
# * every result of a call layout_call() names (a field update, a TOC: the page numbers
#   are measured, and written into the document) -- and, for a call that writes them, every
#   later result in the session, which reads that document back: of those only ``ok``, the
#   error and ``data_keys`` hold everywhere.
#
# Everything else -- ok, summary, ids, warnings, new problems, the schema and
# ``data_sha_core`` -- must replay identically on every runner.  If a new transcript fails
# the core replay on a runner without the faces, a result depends on the layout in a way
# this split does not name: add it here, with why.

#: Data fields the layout measures, by tool ("*": every tool).
LAYOUT_FIELDS: dict[str, frozenset[str]] = {
    "*": frozenset({"reflow", "coverage", "pages", "pages_estimated"}),
    "save_document": frozenset({"size"}),
}
#: Keys a layout field has only when the layout stopped, substituted or approximated.
CONDITIONAL_KEYS: dict[str, frozenset[str]] = {
    "coverage": frozenset({"stop", "story_stops", "substituted_fonts", "missing_fonts", "pages_estimated",
                           "approximations", "approximations_total"}),
    "reflow": frozenset({"stopped", "layout_warnings", "error"}),
}


def layout_call(tool: str, arguments: dict[str, Any]) -> str | None:
    """``"document"`` when the call writes what the layout measured into the document (its
    result and every later one in the session depend on the layout), ``"result"`` when only
    its own result does, else ``None``."""
    if tool == "word_fields":
        action = arguments.get("action")
        if action in ("update", "insert_toc", "insert_field") or (
                action == "insert_cross_reference" and arguments.get("kind") == "page"):
            return "document"
    if tool == "word_inspect" and arguments.get("layout"):
        return "result"
    return None


def data_keys(data: Any) -> list[str]:
    """The schema of a result's data: its keys and, for a dict-valued one, that dict's keys
    (``coverage.status``); for a list of dicts, the keys its items have (``[].id``).  Less
    :data:`CONDITIONAL_KEYS`."""
    keys: set[str] = set()
    if isinstance(data, dict):
        for key, value in data.items():
            keys.add(key)
            if isinstance(value, dict):
                skip = CONDITIONAL_KEYS.get(key, frozenset())
                keys.update(f"{key}.{sub}" for sub in value if sub not in skip)
    elif isinstance(data, list):
        keys.update(f"[].{key}" for item in data if isinstance(item, dict) for key in item)
    return sorted(keys)


def core_data(tool: str, data: Any) -> Any:
    """``data`` less the fields the layout measures (:data:`LAYOUT_FIELDS`)."""
    if not isinstance(data, dict):
        return data
    measured = LAYOUT_FIELDS["*"] | LAYOUT_FIELDS.get(tool, frozenset())
    return {key: value for key, value in data.items() if key not in measured}


def writes_layout(transcript: dict[str, Any]) -> bool:
    """Whether a call writes what the layout measured into the document (:func:`layout_call`):
    then the saved document, and a check of its page numbers, depend on the layout too."""
    return any(layout_call(call["tool"], call["arguments"]) == "document" for call in transcript["calls"])


def core_mismatches(transcript: dict[str, Any], results: list[dict[str, Any]]) -> list:
    """Where ``results`` (normalised, one per call) differ from the recorded expectations in
    what does not depend on the layout.  ``data_sha`` is never compared here."""
    mismatches, written = [], False
    for index, (call, got) in enumerate(zip(transcript["calls"], results)):
        expect = call["expect"]
        kind = layout_call(call["tool"], call["arguments"])
        written = written or kind == "document"
        if written or kind:
            fields = ("ok", "error", "data_keys")
        else:
            fields = tuple(sorted((set(expect) | set(got)) - {"data_sha"}))
        want = {key: expect.get(key) for key in fields}
        have = {key: got.get(key) for key in fields}
        if want != have:
            mismatches.append((index, call["tool"], want, have))
    if len(results) != len(transcript["calls"]):
        mismatches.append((len(results), "calls", len(transcript["calls"]), len(results)))
    return mismatches


def replay(transcript: dict[str, Any], inputs: Path, *,
           on_result: Callable[[int, Any], None] | None = None) -> tuple[list[dict[str, Any]], list]:
    """Run a transcript's calls on fresh copies of its inputs: every call's normalised
    result, and the session's outputs.  ``on_result(index, result)`` sees each raw result."""
    from ooxml_edit.tools import Toolbox
    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    clock_at = dt.datetime.fromisoformat(transcript["clock"])
    results = []
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session = box.session(clock=lambda: clock_at)
        for name in transcript.get("open", []):
            session.open((inputs / name).read_bytes(), name)
        for name in transcript.get("blobs", []):
            session.add_blob(blob_bytes(inputs, name), name)
        for index, call in enumerate(transcript["calls"]):
            result = box.dispatch(session, call["tool"], call["arguments"])
            if on_result is not None:
                on_result(index, result)
            results.append(normalise(result, call["tool"]))
        outputs = session.take_outputs()
    return results, outputs


def run(transcript: dict[str, Any], inputs: Path, *, record: bool = False) -> tuple[dict[str, Any], list]:
    """Run a transcript's calls on fresh copies of its inputs.  Returns the transcript (with
    ``expect`` and ``output`` filled in when recording, and ``_mismatches``: every call whose
    result is not the one recorded) and the session's outputs."""
    results, outputs = replay(transcript, inputs)
    mismatches = []
    for index, (call, got) in enumerate(zip(transcript["calls"], results)):
        if record:
            call["expect"] = got
        elif got != call.get("expect"):
            mismatches.append((index, call["tool"], call.get("expect"), got))
    if record and outputs:
        last = outputs[-1]
        transcript["output"] = {"name": last.name, "sha256": hashlib.sha256(last.data).hexdigest()}
    transcript["_mismatches"] = mismatches
    return transcript, outputs
