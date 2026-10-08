"""Shared helpers for the trial's check scripts: record named checks, print JSON."""
from __future__ import annotations

import json
import re
import sys
import traceback
from pathlib import Path

TRIAL = Path(__file__).resolve().parents[1]
INPUTS = TRIAL / "_inputs"


class Checks:
    def __init__(self, task: str, output: Path):
        self.task = task
        self.output = Path(output)
        self.results: list[dict] = []

    def check(self, name: str, ok, detail: str = "") -> bool:
        self.results.append({"check": name, "ok": bool(ok), "detail": str(detail)[:600]})
        return bool(ok)

    def run(self, name: str, fn) -> None:
        """Run fn(); an exception fails the check with its message."""
        try:
            res = fn()
            if isinstance(res, tuple):
                self.check(name, res[0], res[1])
            else:
                self.check(name, res)
        except Exception as error:  # noqa: BLE001
            self.check(name, False, f"{type(error).__name__}: {error}")

    def finish(self) -> None:
        passed = sum(r["ok"] for r in self.results)
        print(json.dumps({"task": self.task, "output": str(self.output), "passed": passed,
                          "total": len(self.results), "results": self.results}, ensure_ascii=False, indent=1))


def norm(text: str) -> str:
    text = text.replace(" ", " ").replace("￼", "")
    text = text.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    return re.sub(r"\s+", " ", text).strip()


def main(task: str, body) -> None:
    output = Path(sys.argv[1])
    checks = Checks(task, output)
    if not output.exists():
        checks.check("output exists", False, f"{output} missing")
        checks.finish()
        return
    checks.check("output exists", True)
    try:
        body(checks, output)
    except Exception:  # noqa: BLE001
        checks.check("checker ran", False, traceback.format_exc()[-600:])
    checks.finish()


def xlsx_grid(data: bytes) -> dict:
    """{(row, col): value} of the first worksheet of an xlsx (grader-side reader)."""
    import io
    import zipfile
    from lxml import etree

    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    z = zipfile.ZipFile(io.BytesIO(data))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = etree.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", ns):
            shared.append("".join(si.itertext()))
    sheet = sorted(n for n in z.namelist() if n.startswith("xl/worksheets/sheet"))[0]
    root = etree.fromstring(z.read(sheet))
    grid = {}
    for c in root.iter("{%s}c" % ns["m"]):
        ref = c.get("r")
        col_letters = "".join(ch for ch in ref if ch.isalpha())
        row = int("".join(ch for ch in ref if ch.isdigit()))
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - 64)
        v = c.find("m:v", ns)
        isv = c.find("m:is", ns)
        if c.get("t") == "s" and v is not None:
            val = shared[int(v.text)]
        elif c.get("t") == "inlineStr" and isv is not None:
            val = "".join(isv.itertext())
        elif v is not None:
            try:
                val = float(v.text)
            except ValueError:
                val = v.text
        else:
            val = None
        grid[(row, col)] = val
    return grid


# ---------------------------------------------------------------- PowerPoint helpers
def notes_of(deck, number: int) -> str:
    """Speaker notes of 1-based slide ``number``, read from to_outline (Markdown-unescaped)."""
    md = deck.to_outline(slides=[number])
    if "\nNotes:\n" not in md:
        return ""
    text = md.split("\nNotes:\n", 1)[1]
    text = re.sub(r"\\([\\`*_{}\[\]()#+\-.!|>~])", r"\1", text)
    return norm(text)


def title_of(slide) -> str:
    for sh in slide.shapes:
        ph = sh.placeholder
        if ph and ph[0] in ("title", "ctrTitle"):
            return norm(sh.text)
    return ""


def all_shapes(slide):
    out = []
    stack = list(slide.shapes)
    while stack:
        sh = stack.pop(0)
        out.append(sh)
        try:
            stack[0:0] = list(sh.children or [])
        except Exception:  # noqa: BLE001
            pass
    return out
