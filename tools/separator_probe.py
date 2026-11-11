#!/usr/bin/env python3
"""Measure what Word does, tracking, when a table is inserted where a paragraph must come
with it -- between two tables, in a story's last paragraph, at the end of a paragraph's text
-- and then deleted again (round 51).

The probe document is ``Start.``, table 1, an empty paragraph, table 2, ``Middle.``,
``Last paragraph.``, table 3 and an empty last paragraph.  Word -- through ``tests/oracle.py``'s
machine-wide lock -- opens it, turns tracking on and, from the bottom up, inserts a 2x2 table
(and types into its first cell) in the last paragraph (``end``), at the end of ``Last paragraph.``
(``text``) and in the paragraph between tables 1 and 2 (``between``).  ``made`` is the
result; ``deleted`` the same with each new table then deleted (selected, Backspace).

    python tools/separator_probe.py      # writes tests/observations/separator-word.json

Nothing Word writes is committed: only the facts read from it, authors and dates replaced
and rsids dropped.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))

import oracle  # noqa: E402
from e3_probe import BACKSPACE, _locate, _package, _script, _type  # noqa: E402
from rows_probe import _paragraph, _row, _table  # noqa: E402
from tracking_probe import read, user_name  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "separator-word.json"


def document() -> bytes:
    def table(key: str) -> str:
        return _table([_row(key, 1), _row(key, 2)])

    return _package("".join([_paragraph("Start."), table("One"), _paragraph(), table("Two"),
                             _paragraph("Middle."), _paragraph("Last paragraph."), table("Three"), _paragraph()]))


def _new_table(at: str, index: str, typed: str, delete: bool) -> list[str]:
    lines = [f"set x to {at}",
             "make new table at d with properties {text object:(create range d start x end x), "
             "number of rows:2, number of columns:2}",
             f"set n to table ({index}) of d",
             "set e to start of content of text object of (get cell from table n row 1 column 1)",
             "select (create range d start e end e)", _type(typed)]
    if delete:
        lines += ["select text object of n", BACKSPACE]
    return lines


def script(delete: bool) -> str:
    lines = ["set track revisions of d to true"]
    lines += _new_table("(end of content of text object of d) - 1", "count of tables of d", "End typed", delete)
    lines += [_locate("Last"), "set e to end of content of text object of paragraph k of d"]
    lines += _new_table("e - 1", "3", "Text typed", delete)
    lines += _new_table("end of content of text object of table 1 of d", "2", "Between typed", delete)
    return _script(lines)


def main() -> None:
    out: dict = {}
    with oracle.session() as session:
        me = user_name()
        session.recover()
        data = document()
        for name, delete in (("made", False), ("deleted", True)):
            outcome = session.run_script(script(delete), data, name=f"separator-{name}", tag="word", timeout=200)
            if not outcome:
                raise SystemExit(f"Word did not save {name}: {outcome.outcome} {outcome.detail}")
            out[name] = [item.replace(me, "{author}") for item in read(outcome.path.read_bytes())]
            outcome.path.unlink()
    OBSERVATIONS.write_text(json.dumps({
        "_about": ("What Word 16.106 for Mac wrote for tools/separator_probe.py: tables inserted, tracked, "
                   "between two tables, in a story's last paragraph and at the end of a paragraph's text "
                   "(made), and the same each deleted again (deleted). Authors ({author} the Word user) and "
                   "dates replaced, rsids dropped. Regenerate with python tools/separator_probe.py."),
        **{key: out[key] for key in sorted(out)}}, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
