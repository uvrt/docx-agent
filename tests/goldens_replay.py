"""Golden transcripts: record a task's tool calls, replay them, grade the output.

A transcript is JSON::

    {"task": "w2-review-comments", "clock": "2026-10-07T09:00:00+00:00",
     "open": ["hosting-agreement-summary.docx"],      # opened by the app: d1, d2...
     "blobs": ["draft.md"],                           # registered by the app: b1, b2...
     "calls": [{"tool": ..., "arguments": {...}, "expect": {...}}],
     "output": {"name": ..., "sha256": ...}}

``expect`` is the normalised result: ok, summary, changed, created, removed, renamed, refs,
warnings, the validate delta's new problems, and a digest of data.  Replaying gives the same
results and a byte-identical output (the session clock is fixed).
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
from pathlib import Path
from typing import Any

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


def normalise(result: Any) -> dict[str, Any]:
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
        out["data_sha"] = hashlib.sha256(json.dumps(data, sort_keys=True, ensure_ascii=False)
                                         .encode()).hexdigest()[:16]
    return out


def run(transcript: dict[str, Any], inputs: Path, *, record: bool = False) -> tuple[dict[str, Any], list]:
    """Run a transcript's calls on fresh copies of its inputs.  Returns the transcript (with
    ``expect`` and ``output`` filled in when recording) and the session's outputs."""
    from ooxml_edit.tools import Toolbox
    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    clock_at = dt.datetime.fromisoformat(transcript["clock"])
    mismatches = []
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        session = box.session(clock=lambda: clock_at)
        for name in transcript.get("open", []):
            session.open((inputs / name).read_bytes(), name)
        for name in transcript.get("blobs", []):
            session.add_blob(blob_bytes(inputs, name), name)
        for index, call in enumerate(transcript["calls"]):
            result = box.dispatch(session, call["tool"], call["arguments"])
            got = normalise(result)
            if record:
                call["expect"] = got
            elif got != call.get("expect"):
                mismatches.append((index, call["tool"], call.get("expect"), got))
        outputs = session.take_outputs()
    if record and outputs:
        last = outputs[-1]
        transcript["output"] = {"name": last.name, "sha256": hashlib.sha256(last.data).hexdigest()}
    transcript["_mismatches"] = mismatches
    return transcript, outputs
