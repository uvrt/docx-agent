#!/usr/bin/env python3
"""Measure what Word does, tracking, when a row or a table that is a pending insertion is
deleted -- by the author who inserted it and by another -- and when text typed into a cell
is cleared again (round 51).

One probe document (mode 15, every row and paragraph with its paraId) holds a table per
case, each followed by a paragraph ``After <Key>.``; cell text is ``<Key> r<i>c<j>``.  Word -- through
``tests/oracle.py``'s machine-wide lock -- opens it, turns tracking on, makes each case's
edits by AppleScript (from the last table up, so no case moves the tables of those above it)
and saves; Word's own Accept All and Reject All of what it saved are read too.

* ``Live``: Word inserts a row below row 2, types into its first cell, then deletes the row
  (the row selected, Backspace).  ``LiveEmpty``: the same with nothing typed.
* ``LiveClear``: a row inserted and typed into, then the typed text selected and deleted
  (the row stays).
* ``Mine`` / ``Other``: the file holds row 3 as a pending insertion (its mark, its cells'
  paragraph marks and text) by the Word user (by name) / by another author; Word deletes it.
* ``MineOtherText``: row 3 inserted by the Word user, its first cell's text by another
  author (``MineOtherMid``: the same as row 2 of 3); ``OtherMineText``: the reverse.  ``MinePlain``: row 3 inserted by the Word user,
  its cells' text and marks not revisions (as after a partial accept).  Each deleted.
* ``Two``: rows 2 (not a revision) and 3 (the Word user's insertion) deleted together.
* ``Typed``: text typed at the end of cell (2, 1) and then deleted; in cell (2, 2) typed and
  then the whole cell's text deleted.
* ``TableMine``: a table every row of which is the Word user's insertion, deleted whole;
  ``TableMixed``: the same with its second row another author's.
* ``TableLive``: a table Word inserts (at the paragraph ``Here``), types into, then deletes
  (Word writes the new table as rows inserted, each cell's mark inserted, the typed text an
  insertion: checked by hand once).

    python tools/rows_probe.py           # writes tests/observations/rows-word.json

Nothing Word writes is committed: only the facts read from it, authors and dates replaced
(``{author}`` the Word user, ``{other}`` the file's other author) and rsids dropped.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))

import oracle  # noqa: E402
from e3_probe import BACKSPACE, _package, _script, _type  # noqa: E402
from tracking_probe import OTHER, read, user_name  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "rows-word.json"
DATE = 'w:date="2026-11-01T10:00:00Z"'

#: Case key -> what row 3 (or every row) of its table holds: (row author, per-cell text
#: author, per-cell mark author); ``None`` is no revision.
TABLES = ["Live", "LiveEmpty", "LiveClear", "Mine", "Other", "MineOtherText", "MineOtherMid", "OtherMineText",
          "MinePlain",
          "Two", "Typed", "TableMine", "TableMixed", "TableLive"]

_ids = iter(range(1000, 100000))
_paras = iter(range(1, 100000))


def _pid() -> str:
    return f"{0x40000000 + next(_paras):08X}"


def _record(tag: str, author: str | None) -> str:
    return f'<w:{tag} w:id="{next(_ids)}" w:author="{author}" {DATE}/>' if author else ""


def _cell(text: str, text_author: str | None, mark_author: str | None) -> str:
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
    if text_author:
        run = f'<w:ins w:id="{next(_ids)}" w:author="{text_author}" {DATE}>{run}</w:ins>'
    mark = _record("ins", mark_author)
    props = f"<w:pPr><w:rPr>{mark}</w:rPr></w:pPr>" if mark else ""
    return (f'<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
            f'<w:p w14:paraId="{_pid()}" w14:textId="77777777">{props}{run}</w:p></w:tc>')


def _row(key: str, i: int, row_author=None, text_authors=(None, None), mark_authors=(None, None)) -> str:
    record = _record("ins", row_author)
    props = f"<w:trPr>{record}</w:trPr>" if record else ""
    cells = "".join(_cell(f"{key} r{i}c{j + 1}", text_authors[j], mark_authors[j]) for j in range(2))
    return f'<w:tr w14:paraId="{_pid()}" w14:textId="77777777">{props}{cells}</w:tr>'


def _table(rows: list[str]) -> str:
    grid = '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
    props = '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
    return f"<w:tbl>{props}{grid}{''.join(rows)}</w:tbl>"


def _paragraph(text: str = "") -> str:
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p w14:paraId="{_pid()}" w14:textId="77777777">{run}</w:p>'


def document(me: str) -> bytes:
    def inserted(key: str, i: int, row: str, texts: tuple, marks: tuple | None = None) -> str:
        return _row(key, i, row, texts, texts if marks is None else marks)

    plain = lambda key, i: _row(key, i)  # noqa: E731
    out = [_paragraph("Tables follow.")]
    for key in TABLES:
        if key == "MineOtherMid":
            rows = [plain(key, 1), inserted(key, 2, me, (OTHER, me), (me, me)), plain(key, 3)]
        elif key in ("Mine", "Other", "MineOtherText", "OtherMineText", "MinePlain", "Two"):
            author = OTHER if key.startswith("Other") else me
            texts = {"MineOtherText": (OTHER, me), "OtherMineText": (me, me), "MinePlain": (None, None)}.get(
                key, (author, author))
            marks = {"MineOtherText": (me, me), "OtherMineText": (OTHER, OTHER)}.get(key, texts)
            rows = [plain(key, 1), plain(key, 2), inserted(key, 3, author, texts, marks)]
        elif key == "TableMine":
            rows = [inserted(key, i, me, (me, me)) for i in (1, 2, 3)]
        elif key == "TableMixed":
            rows = [inserted(key, 1, me, (me, me)), inserted(key, 2, OTHER, (OTHER, OTHER)),
                    inserted(key, 3, me, (me, me))]
        elif key == "TableLive":
            out.append(_paragraph("Here"))
            continue
        else:
            rows = [plain(key, i) for i in (1, 2, 3)]
        out.append(_table(rows))
        out.append(_paragraph(f"After {key}."))
    out.append(_paragraph("Tail end."))
    return _package("".join(out))


def script() -> str:
    t = {key: k + 1 for k, key in enumerate(k for k in TABLES if k != "TableLive")}

    def cell(key: str, r: int, c: int) -> str:
        return f"(get cell from table (table {t[key]} of d) row {r} column {c})"

    def row(key: str, r: int) -> str:
        return f"(row {r} of table {t[key]} of d)"

    def delete_row(key: str, r: int) -> list[str]:
        return [f"select text object of {row(key, r)}", BACKSPACE]

    def end_of_cell(key: str, r: int, c: int) -> list[str]:
        return [f"set e to end of content of text object of {cell(key, r, c)}",
                "select (create range d start (e - 1) end (e - 1))"]

    lines = ["set track revisions of d to true"]
    # TableLive: a 2x2 table at the paragraph "Here", typed into, then deleted.
    lines += ['set k to 0',
              'repeat with i from 1 to count of paragraphs of d',
              '  if content of text object of paragraph i of d starts with "Here" then',
              '    set k to i',
              '    exit repeat',
              '  end if',
              'end repeat',
              "set h to text object of paragraph k of d",
              "set n to make new table at d with properties {text object:(create range d start (start of content of h) "
              "end (start of content of h)), number of rows:2, number of columns:2}",
              f"set n to table {len(t) + 1} of d",
              "set e to start of content of text object of (get cell from table n row 1 column 1)",
              "select (create range d start e end e)",
              _type("TableLive typed"),
              "select text object of n", BACKSPACE]
    lines += [f"select text object of table {t['TableMixed']} of d", BACKSPACE]
    lines += [f"select text object of table {t['TableMine']} of d", BACKSPACE]
    # Typed: typed then deleted; typed then the whole cell's text deleted.
    lines += end_of_cell("Typed", 2, 2) + [_type(" TYPEDTWO"),
                                            f"select text object of {cell('Typed', 2, 2)}",
                                            "set e to end of content of text object of selection",
                                            "select (create range d start (start of content of text object of selection) "
                                            "end (e - 1))", BACKSPACE]
    lines += end_of_cell("Typed", 2, 1) + [_type(" TYPEDONE"),
                                            "set e to end of content of text object of selection",
                                            "select (create range d start (e - 9) end e)", BACKSPACE]
    lines += [f"select (create range d start (start of content of text object of {row('Two', 2)}) "
              f"end (end of content of text object of {row('Two', 3)}))", BACKSPACE]
    for key in ("MinePlain", "OtherMineText"):
        lines += delete_row(key, 3)
    lines += delete_row("MineOtherMid", 2)
    for key in ("MineOtherText", "Other", "Mine"):
        lines += delete_row(key, 3)
    # LiveClear: a row inserted, typed into, the typed text deleted.
    lines += [f"select text object of {cell('LiveClear', 2, 1)}", "insert rows selection position below"]
    lines += end_of_cell("LiveClear", 3, 1) + [_type("CLEARED"),
                                                "select text object of " + cell("LiveClear", 3, 1)]
    lines += [f"set e to end of content of text object of {cell('LiveClear', 3, 1)}",
              f"select (create range d start (start of content of text object of {cell('LiveClear', 3, 1)}) "
              "end (e - 1))", BACKSPACE]
    lines += [f"select text object of {cell('LiveEmpty', 2, 1)}", "insert rows selection position below"]
    lines += delete_row("LiveEmpty", 3)
    lines += [f"select text object of {cell('Live', 2, 1)}", "insert rows selection position below"]
    lines += end_of_cell("Live", 3, 1) + [_type("LIVENEW")]
    lines += delete_row("Live", 3)
    return _script(lines)


ACCEPT = _script(["accept all revisions d"])
REJECT = _script(["reject all revisions d"])


def main() -> None:
    out: dict = {}
    with oracle.session() as session:
        me = user_name()
        session.recover()
        data = document(me)
        if "--keep" in sys.argv:
            Path("rows-in.docx").write_bytes(data)
        outcome = session.run_script(script(), data, name="rows-tracked", tag="word", timeout=200)
        if not outcome:
            raise SystemExit(f"Word did not save: {outcome.outcome} {outcome.detail}")
        saved = outcome.path.read_bytes()
        outcome.path.unlink()
        results = {"tracked": saved}
        for name, review in (("accept", ACCEPT), ("reject", REJECT)):
            done = session.run_script(review, saved, name=f"rows-{name}", tag="word", timeout=200)
            if not done:
                raise SystemExit(f"Word did not save {name}: {done.outcome} {done.detail}")
            results[name] = done.path.read_bytes()
            done.path.unlink()
    for name, data in results.items():
        out[name] = [item.replace(me, "{author}") for item in read(data)]
    OBSERVATIONS.write_text(json.dumps({
        "_about": ("What Word 16.106 for Mac wrote for tools/rows_probe.py: rows and tables that are pending "
                   "insertions deleted by their author and by another, text typed into cells then deleted, and "
                   "Word's Accept All and Reject All of the result. Authors ({author} the Word user, {other} "
                   "another) and dates replaced, rsids dropped. Regenerate with python tools/rows_probe.py."),
        **{key: out[key] for key in sorted(out)}}, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
