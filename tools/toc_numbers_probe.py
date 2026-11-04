#!/usr/bin/env python3
"""Measure what Word's own table-of-contents update writes for numbered headings and next to
``toc N`` styles' own tab stops (production round 32, R12 and R14).

Made-up documents (``tests/toc_number_cases.py``): a stale ``TOC \\o "1-3" \\h \\z \\u`` and
headings numbered through a list -- linked to the heading styles or on the paragraphs, a
number then a tab, a space or nothing, numbers of many widths ("3" to "WWWWWW 3"), several
fonts, sizes and ``toc N`` indents -- and ``toc N`` styles with right, left, inherited and
cleared stops.  Word -- through ``tests/oracle.py``'s machine-wide lock, Normal.dotm backed up
and restored -- runs ``update (table of contents 1)`` on each and saves it; each entry's level,
tab stops and the texts of its runs before the page number are read back into
``tests/observations/toc-numbers-word.json``.

    python tools/toc_numbers_probe.py [case ...]

Nothing Word writes is committed: only these facts.
"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

OBSERVED = ROOT / "tests" / "observations" / "toc-numbers-word.json"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NORMAL = (Path.home() / "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized"
          / "Normal.dotm")
UPDATE = '''
on run argv
    with timeout of 120 seconds
        tell application "Microsoft Word"
            activate
            open (POSIX file (item 1 of argv))
            set d to active document
            update (table of contents 1 of d)
            save as d file name (item 2 of argv) file format format document
            close d saving no
            quit saving no
        end tell
    end timeout
end run
'''


def entries(data: bytes) -> list[dict]:
    """Each TOC entry Word wrote: level (Word renames the toc styles in its own language),
    tab stops, and its runs' texts before the page number's tab (a tab run as ``\\t``)."""
    root = etree.fromstring(zipfile.ZipFile(__import__("io").BytesIO(data)).read("word/document.xml"))
    out = []
    for p in root.iter(W + "p"):
        style = p.find(f"{W}pPr/{W}pStyle")
        found = re.match(r"\D+(\d)$", style.get(W + "val") if style is not None else "")
        if not found or not style.get(W + "val").lower().startswith(("toc", "inhopg")):
            continue
        tabs = [[t.get(W + "val"), int(t.get(W + "pos"))] + ([t.get(W + "leader")] if t.get(W + "leader") else [])
                for t in p.iterfind(f"{W}pPr/{W}tabs/{W}tab")]
        runs = []
        for r in (p.find(W + "hyperlink") if p.find(W + "hyperlink") is not None else p).iter(W + "r"):
            if r.find(W + "fldChar") is not None or r.find(W + "instrText") is not None:
                break
            t = r.find(W + "t")
            runs.append(t.text if t is not None else ("\t" if r.find(W + "tab") is not None else ""))
        out.append({"level": int(found.group(1)), "tabs": tabs, "runs": runs[:-1] if runs[-1:] == ["\t"] else runs})
    return out


def main(only: list[str]) -> None:
    import oracle
    from toc_number_cases import CASES, CASES2, CASES3

    cases = {k: v for k, v in {**CASES, **CASES2, **CASES3}.items() if not k.startswith("x-none")}
    observed = json.loads(OBSERVED.read_text()) if OBSERVED.exists() else {"cases": {}}
    before = hashlib.sha256(NORMAL.read_bytes()).hexdigest()
    backup = Path(tempfile.mkdtemp()) / "Normal.dotm"
    shutil.copy2(NORMAL, backup)
    try:
        with oracle.session() as word:
            for name, data in cases.items():
                if only and name not in only:
                    continue
                out = word.run_script(UPDATE, data, name="tocnum", tag=name)
                print(name, out.outcome, out.detail[:200], flush=True)
                if out.ok:
                    observed["cases"][name] = entries(out.path.read_bytes())
    finally:
        if hashlib.sha256(NORMAL.read_bytes()).hexdigest() != before:
            shutil.copy2(backup, NORMAL)
    OBSERVED.write_text(json.dumps(observed, indent=1, ensure_ascii=False) + "\n")


if __name__ == "__main__":
    main(sys.argv[1:])
