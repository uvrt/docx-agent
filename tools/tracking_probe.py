#!/usr/bin/env python3
"""Measure what Word records when, tracking, it formats what is itself a pending insertion,
when paragraphs are typed one after another at a document's end, and which paragraph a
deleted paragraph mark keeps (round 50).

Probe documents are written here (mode 15, every paragraph with its paraId), and Word --
through ``tests/oracle.py``'s machine-wide lock -- opens each, turns tracking on, makes the
edits by AppleScript and saves; the joins are accepted with Word's own ``accept all
revisions`` too, and exported in its final view (No Markup).

* ``format``: a paragraph typed after one (Enter at its end, then the text) and one typed
  before one (the text, then Enter), each then given a heading style, centred and a word
  made bold; a word typed into a paragraph, then made bold; a paragraph whose mark the
  typing made an insertion, then formatted; a paragraph and a run the file holds as
  another author's insertion, and as the Word user's own (by name), formatted.
* ``selection``: the same through the selection, as a command acts on what is selected,
  and bold and centring switched on before the text is typed.
* ``end-text-first``, ``end-style-first``: three paragraphs typed at the document's end,
  each given a heading style after (or before) its text is typed.
* ``join``: a mark deleted (``Jone``) and a whole paragraph deleted (``Kone``), tracked; the
  file also holds a deleted mark (``Lone``) and a deleted paragraph (``Mone``) with no
  ``w:pPrChange`` on the paragraph after, as another tool may write them.  ``join-accept``
  is Word's Accept All of what it saved, ``join-input-accept`` of the file as given, and
  ``join-final`` where the final view draws each paragraph (the left and right edges of its
  first line, in points; the first paragraphs are left-indented and centred, the second
  right-aligned).

    python tools/tracking_probe.py           # writes tests/observations/tracking-word.json

Nothing Word writes is committed: only the facts read from it, authors and dates replaced
(``{author}`` the Word user, ``{other}`` the file's other author) and rsids dropped.
"""

from __future__ import annotations

import io
import json
import subprocess
import zipfile
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))

import oracle  # noqa: E402
from e3_probe import BACKSPACE, _locate, _package, _script, _type, blocks  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "tracking-word.json"
OTHER = "Other Writer"
DATE = 'w:date="2026-11-01T10:00:00Z"'


def para(k: int, content: str, ppr: str = "", mark: str = "") -> str:
    rpr = f"<w:rPr>{mark}</w:rPr>" if mark else ""
    props = f"<w:pPr>{ppr}{rpr}</w:pPr>" if (ppr or rpr) else ""
    return f'<w:p w14:paraId="{0x30000000 + k:08X}" w14:textId="77777777">{props}{content}</w:p>'


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def inserted(author: str, identifier: int, content: str = "") -> str:
    if not content:
        return f'<w:ins w:id="{identifier}" w:author="{author}" {DATE}/>'
    return f'<w:ins w:id="{identifier}" w:author="{author}" {DATE}>{content}</w:ins>'


# -- documents ----------------------------------------------------------------------------------


def format_document(self_name: str) -> bytes:
    out: list[str] = []
    for key in ("Sone", "Stwo", "Sthree", "Sfour"):
        out.append(para(len(out), run(f"{key} existing alpha beta.")))
    for tag, author in (("O", OTHER), ("M", self_name)):
        out.append(para(len(out), run(f"{tag}one existing.")))
        k = len(out)
        out.append(para(k, inserted(author, 100 + k, run(f"{tag}onenew inserted paragraph.")),
                        mark=inserted(author, 200 + k)))
        k = len(out)
        out.append(para(k, run(f"{tag}two existing ") + inserted(author, 300 + k, run("XXRUNXX ")) + run("epsilon.")))
    out.append(para(len(out), run("Tail end.")))
    return _package("".join(out))


def selection_document() -> bytes:
    return _package("".join(para(k, run(f"{key} existing alpha beta."))
                            for k, key in enumerate(("Aone", "Atwo", "Athree", "Tail"))))


def end_document() -> bytes:
    return _package(para(0, run("Existing first.")) + para(1, run("Existing last.")))


def join_document() -> bytes:
    first = '<w:spacing w:before="240" w:after="0"/><w:ind w:left="720"/><w:jc w:val="center"/>'
    second = '<w:spacing w:before="0" w:after="480"/><w:jc w:val="right"/>'
    gone = f'<w:del w:id="900" w:author="{OTHER}" {DATE}/>'
    deleted_text = (f'<w:del w:id="901" w:author="{OTHER}" {DATE}><w:r><w:delText xml:space="preserve">'
                    f'Mone whole paragraph.</w:delText></w:r></w:del>')
    return _package("".join([
        para(0, run("Jone first paragraph."), first), para(1, run("Jtwo second paragraph."), second),
        para(2, run("Kone whole paragraph."), first), para(3, run("Ktwo next paragraph."), second),
        para(4, run("Lone first paragraph."), first, gone), para(5, run("Ltwo second paragraph."), second),
        para(6, deleted_text, first, f'<w:del w:id="902" w:author="{OTHER}" {DATE}/>'),
        para(7, run("Mtwo next paragraph."), second),
        para(8, run("Tail end."))]))


# -- scripts ------------------------------------------------------------------------------------


def _end_of(key: str) -> list[str]:
    return [_locate(key), "set e to end of content of text object of paragraph k of d",
            "select (create range d start (e - 1) end (e - 1))"]


def _formatted(key: str, bold_from: int, bold_to: int) -> list[str]:
    return [_locate(key), "set style of paragraph k of d to style heading2",
            "set alignment of paragraph format of paragraph k of d to align paragraph center",
            f"set bold of font object of (create range d start (s + {bold_from}) end (s + {bold_to})) to true"]


def format_script() -> str:
    lines = ["set track revisions of d to true"]
    lines += _end_of("Sone") + ["type paragraph selection", _type("Sonenew typed paragraph.")]
    lines += _formatted("Sonenew", 8, 13)
    lines += [_locate("Stwo"), "select (create range d start s end s)", _type("Stwonew typed before."),
              "type paragraph selection"]
    lines += _formatted("Stwonew", 8, 13)
    lines += [_locate("Sthree"), "select (create range d start (s + 7) end (s + 7))", _type("RUNINS ")]
    lines += [_locate("Sthree"), "set bold of font object of (create range d start (s + 7) end (s + 13)) to true"]
    lines += _end_of("Sfour") + ["type paragraph selection", _type("Sfournew typed.")]
    lines += [_locate("Sfour"), "set alignment of paragraph format of paragraph k of d to align paragraph right",
              "set bold of font object of (create range d start (s + 6) end (s + 14)) to true"]
    for tag in ("O", "M"):
        lines += _formatted(f"{tag}onenew", 8, 16)
        lines += [_locate(f"{tag}two"),
                  "set bold of font object of (create range d start (s + 14) end (s + 22)) to true"]
    return _script(lines)


def selection_script() -> str:
    lines = ["set track revisions of d to true"]
    lines += _end_of("Aone") + ["type paragraph selection", _type("Aonenew typed paragraph.")]
    lines += [_locate("Aonenew"), "select text object of paragraph k of d",
              "set alignment of paragraph format of selection to align paragraph center",
              "select (create range d start (s + 8) end (s + 13))", "set bold of font object of selection to true"]
    lines += _end_of("Atwo") + ["type paragraph selection",
                                "set alignment of paragraph format of selection to align paragraph center",
                                "set bold of font object of selection to true", _type("Atwonew typed bold.")]
    lines += [_locate("Athree"), "select (create range d start (s + 7) end (s + 7))", _type("FIRST "),
              "select (create range d start (s + 7) end (s + 12))", "set italic of font object of selection to true"]
    return _script(lines)


def end_script(style_first: bool) -> str:
    lines = ["set track revisions of d to true"]
    last = "paragraph (count of paragraphs of d) of d"
    for name in ("Pone", "Ptwo", "Pthree"):
        lines += ["set e to end of content of text object of d",
                  "select (create range d start (e - 1) end (e - 1))", "type paragraph selection"]
        typed, styled = _type(f"{name} typed."), f"set style of {last} to style heading2"
        lines += [styled, typed] if style_first else [typed, styled]
    return _script(lines)


def join_script() -> str:
    return _script(["set track revisions of d to true",
                    _locate("Jone"), "set e to end of content of text object of paragraph k of d",
                    "select (create range d start (e - 1) end e)", BACKSPACE,
                    _locate("Kone"), "select text object of paragraph k of d", BACKSPACE])


ACCEPT = _script(["accept all revisions d"])


# -- reading ------------------------------------------------------------------------------------


def read(data: bytes) -> list[str]:
    """Body blocks, normalised (``e3_probe.normalise``), the other author's records saying
    ``{other}``."""
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            content = source.read(item.filename)
            if item.filename == "word/document.xml":
                content = content.replace(f'w:author="{OTHER}"'.encode(), b'w:xauthor="{other}"')
            target.writestr(item, content)
    marked = buffer.getvalue()
    return [item.replace('w:xauthor="{other}"', 'w:author="{other}"') for item in blocks(marked)]


def edges(pdf: Path) -> dict[str, list[float]]:
    """Left and right edge (points) of the first line of each paragraph, by its key."""
    import pypdfium2

    document = pypdfium2.PdfDocument(str(pdf))
    page = document[0]
    text_page = page.get_textpage()
    text = text_page.get_text_range()
    out: dict[str, list[float]] = {}
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        start = text.find(line)
        first = text_page.get_charbox(start)
        last = text_page.get_charbox(start + len(line) - 1)
        out[line.split()[0]] = [round(first[0], 1), round(last[2], 1)]
    document.close()
    return out


def user_name() -> str:
    done = subprocess.run(["osascript", "-e", 'tell application "Microsoft Word" to get user name'],
                          capture_output=True, text=True, timeout=120)
    return done.stdout.strip()


def main() -> None:
    out: dict = {}
    with oracle.session() as session:
        me = user_name()
        session.recover()
        runs = {
            "format": (format_document(me), format_script()),
            "selection": (selection_document(), selection_script()),
            "end-text-first": (end_document(), end_script(False)),
            "end-style-first": (end_document(), end_script(True)),
            "join": (join_document(), join_script()),
        }
        saved: dict[str, bytes] = {}
        for name, (data, script) in runs.items():
            outcome = session.run_script(script, data, name=f"tracking-{name}", tag="word", timeout=200)
            if not outcome:
                raise SystemExit(f"Word did not save {name}: {outcome.outcome} {outcome.detail}")
            saved[name] = outcome.path.read_bytes()
            outcome.path.unlink()
        for name, data in (("join-accept", saved["join"]), ("join-input-accept", runs["join"][0])):
            outcome = session.run_script(ACCEPT, data, name=f"tracking-{name}", tag="word", timeout=200)
            if not outcome:
                raise SystemExit(f"Word did not save {name}: {outcome.outcome} {outcome.detail}")
            saved[name] = outcome.path.read_bytes()
            outcome.path.unlink()
        pdf = session.export_pdf(runs["join"][0], name="tracking-join-final")
        if not pdf:
            raise SystemExit(f"Word did not export the join: {pdf.outcome} {pdf.detail}")
        out["join-final"] = edges(pdf.path)
        pdf.path.unlink()
    for name, data in saved.items():
        out[name] = [item.replace(me, "{author}") for item in read(data)]
    OBSERVATIONS.write_text(json.dumps({
        "_about": ("What Word 16.106 for Mac wrote for tools/tracking_probe.py: formatting what is a "
                   "pending insertion, paragraphs typed at the end, deleted paragraph marks and Accept All "
                   "of them, and the final view's line edges. Authors ({author} the Word user, {other} "
                   "another) and dates replaced, rsids dropped. Regenerate with python tools/tracking_probe.py."),
        **{key: out[key] for key in sorted(out)}}, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
