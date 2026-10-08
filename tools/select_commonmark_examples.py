#!/usr/bin/env python3
"""Select the CommonMark spec's examples ``insert_markdown`` is held to, into
``tests/corpus/commonmark/examples.json``.

The examples are the CommonMark Spec's (version 0.31.2, by John MacFarlane), licensed
CC BY-SA 4.0; the selection is committed with that attribution and under that licence
(``tests/corpus/commonmark/PROVENANCE.md``), so the tests need no network.  What is kept:
the Markdown of every example except

* the sections whose construct is raw HTML (*HTML blocks*, *Raw HTML*), and any other
  example holding raw HTML that is not a comment: ``insert_markdown`` refuses raw HTML by
  design (ROADMAP.md, "The mapping");
* any example matching a pattern given with ``--exclude`` (repeatable, case-insensitive):
  the repository's trace check, whose patterns are not written here -- none matched in
  0.31.2, counted all the same.

Each example keeps its number and section; the expected HTML is not kept (the round trip
compares Markdown with Markdown).

    python tools/select_commonmark_examples.py --exclude PATTERN [--exclude PATTERN ...]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from docx_agent.markdown.parse import parser  # noqa: E402

VERSION = "0.31.2"
SOURCE = f"https://spec.commonmark.org/{VERSION}/spec.json"
OUT = ROOT / "tests" / "corpus" / "commonmark" / "examples.json"
HTML_SECTIONS = {"HTML blocks", "Raw HTML"}
_COMMENT = re.compile(r"\s*(<!--.*?-->\s*)+", re.DOTALL)


def holds_html(markdown: str) -> bool:
    """Whether markdown-it reads raw HTML other than comments in ``markdown``."""
    for token in parser().parse(markdown, {}):
        if token.type == "html_block" and not _COMMENT.fullmatch(token.content):
            return True
        for child in token.children or []:
            if child.type == "html_inline" and not child.content.startswith("<!--"):
                return True
    return False


def main() -> None:
    arguments = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    arguments.add_argument("--exclude", action="append", default=[], help="a pattern to leave out")
    options = arguments.parse_args()
    trace = re.compile("|".join(f"(?:{p})" for p in options.exclude), re.IGNORECASE) if options.exclude else None
    with urllib.request.urlopen(SOURCE, timeout=60) as response:
        spec = json.load(response)
    kept, html, traced = [], 0, 0
    for example in spec:
        markdown = example["markdown"]
        if trace is not None and (trace.search(markdown) or trace.search(example["html"])):
            traced += 1
            continue
        if example["section"] in HTML_SECTIONS or holds_html(markdown):
            html += 1
            continue
        kept.append({"example": example["example"], "section": example["section"], "markdown": markdown})
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({
        "spec": f"CommonMark Spec {VERSION}",
        "source": SOURCE,
        "author": "John MacFarlane",
        "copyright": "Copyright (C) 2014-24 John MacFarlane",
        "licence": "CC BY-SA 4.0, https://creativecommons.org/licenses/by-sa/4.0/",
        "changes": (f"A selection: the Markdown of {len(kept)} of the {len(spec)} examples, without their "
                    f"HTML; {html} holding raw HTML and {traced} matching the trace check left out "
                    "(tools/select_commonmark_examples.py)."),
        "left_out": {"raw_html": html, "trace_check": traced},
        "examples": kept,
    }, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"kept {len(kept)} of {len(spec)}; raw HTML {html}; trace check {traced}")


if __name__ == "__main__":
    main()
