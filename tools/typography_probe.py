#!/usr/bin/env python3
"""Measure what Word expects of the spacing around ``: ; ! ?`` and guillemets, per language
(round 35, R21: ``docx_agent.edit.typography``).

Probe documents are written here (one run per paragraph, the language on the run) and Word
-- through ``tests/oracle.py``'s machine-wide lock -- opens each and, by AppleScript:

* ``proofing``: waits until its grammar check settles (the count of its grammatical errors
  the same three times running, two seconds apart) and lists the paragraphs it flags, for
  each language in :data:`LANGUAGES`, over :data:`PROOFED` (ordinary space, none, U+00A0,
  U+202F before each mark and inside guillemets) and :data:`EDGES` (marks between words,
  URLs, runs of marks...).
* ``autoformat``: runs AutoFormat over straight quotes and marks in French text and reads
  back what it wrote (AutoFormat as you type cannot be driven: typed text is not
  autoformatted -- e2_probe.py saw the same).
* ``line_breaks``: exports a PDF of paragraphs whose line fills up just before `` !`` (and
  ``?``, ``:``, ``;``, ``»``), the space ordinary, U+00A0 or U+202F, and counts the lines that
  start with the mark.
* ``language``: the language Word reports for a run with ``w:lang`` on ``w:docDefaults``,
  the paragraph style, the run, and with only ``w:themeFontLang``.

    python tools/typography_probe.py      # writes tests/observations/typography-word.json

Nothing Word writes is committed: only these facts.  Back up Normal.dotm first (Word's
grammar settings live with it); the probe changes no setting.
"""

from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "tools"))

import oracle  # noqa: E402
from make_fixtures import DECL, NS, REL, WML, content_types, relationships  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "typography-word.json"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NBSP, NNBSP = "\u00a0", "\u202f"

#: Languages whose proofing is measured.
LANGUAGES = ["fr", "fr-FR", "fr-CH", "fr-CA", "fr-BE", "fr-LU", "fr-MC", "en-US", "en-GB", "de-DE", "nl-NL",
             "es-ES", "it-IT", "pt-PT"]

#: Each mark with an ordinary space, none, U+00A0 and U+202F before it; guillemets likewise.
PROOFED = [f"{word}{space}{mark} {tail}" for word, mark, tail in
           (("Bonjour", ":", "le monde."), ("Vraiment", "?", "Oui."), ("Attention", "!", "Oui."),
            ("Premier", ";", "second."))
           for space in (" ", "", NBSP, NNBSP)] + \
          [f"Il a dit «{space}bonjour{space}» hier." for space in (" ", "", NBSP, NNBSP)]

EDGES = ["Voir http://example.com/page pour plus.", "Rendez-vous à 12:30 demain.", "Le ratio est 1:2 ici.",
         "Le fichier C:\\dossier est vide.", "Quoi?! Non.", "Quoi ?! Non.", "Non!!! Jamais.", "Le point (?) reste.",
         "Il hésite...? Non.", "Fin?", "Titre:", "Il a dit « oui »!", "Bonjour :le monde.", "Sourire :-) oui.",
         "Mot\t: tab.", "Le mot « déjà» ici.", "Le mot «déjà » ici.", "L’API v2:beta tourne.", "x?y"]


def document(paragraphs: list[str], *, run_lang: str | None = None, default_lang: str | None = "en-US",
             normal_lang: str | None = None, theme_lang: str | None = None) -> bytes:
    def escape(text: str) -> str:
        return text.replace("&", "&amp;").replace("<", "&lt;")

    def run(text: str) -> str:
        if "\t" in text:
            head, _, tail = text.partition("\t")
            return run(head) + f"<w:r>{lang}<w:tab/></w:r>" + run(tail)
        return f'<w:r>{lang}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>' if text else ""

    lang = f'<w:rPr><w:lang w:val="{run_lang}"/></w:rPr>' if run_lang else ""
    body = "".join(f"<w:p>{run(text)}</w:p>" for text in paragraphs)
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
             'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    defaults = f'<w:lang w:val="{default_lang}"/>' if default_lang else ""
    normal = f'<w:rPr><w:lang w:val="{normal_lang}"/></w:rPr>' if normal_lang else ""
    styles = (f'{DECL}<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" '
              f'w:hAnsi="Calibri"/><w:sz w:val="22"/>{defaults}</w:rPr></w:rPrDefault></w:docDefaults>'
              f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>{normal}'
              '</w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
              '<w:name w:val="Default Paragraph Font"/></w:style></w:styles>')
    theme = f'<w:themeFontLang w:val="{theme_lang}"/>' if theme_lang else ""
    settings = (f'{DECL}<w:settings xmlns:w="{W}"><w:compat><w:compatSetting w:name="compatibilityMode" '
                f'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>{theme}</w:settings>')
    parts = {
        "[Content_Types].xml": content_types({"word/document.xml": WML + ".document.main+xml",
                                              "word/styles.xml": WML + ".styles+xml",
                                              "word/settings.xml": WML + ".settings+xml"}),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False)]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": styles,
        "word/settings.xml": settings,
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 11, 5, 12, 0, 0)), data)
    return buffer.getvalue()


def _script(body: str, *, result: str = '""') -> str:
    return f'''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set out to ""
  with timeout of 240 seconds
    tell application "Microsoft Word"
      activate
      open (POSIX file inputPath)
      set d to active document
      {body}
      save as d file name outputPath file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
  return out
end run
'''


GRAMMAR = _script('''
      set prev to -1
      set stable to 0
      repeat 25 times
        delay 2
        set c to -2
        try
          set c to count of grammatical errors of d
        end try
        if c = prev then
          set stable to stable + 1
          if stable >= 3 then exit repeat
        else
          set stable to 0
        end if
        set prev to c
      end repeat
      set g to grammatical errors of d
      repeat with e in g
        try
          set out to out & (start of content of e) & linefeed
        end try
      end repeat''')

LANGUAGE = _script('set out to ((language ID of text object of paragraph 1 of d) as string)')


def flagged(starts: str, paragraphs: list[str]) -> list[str]:
    """The paragraphs an error starts in (Word's ranges count one character per paragraph mark)."""
    offsets, at = [], 0
    for text in paragraphs:
        offsets.append((at, at + len(text)))
        at += len(text) + 1
    out = []
    for line in starts.split():
        start = int(line)
        for (a, b), text in zip(offsets, paragraphs):
            if a <= start <= b and text not in out:
                out.append(text)
    return out


def _paragraph_texts(path: Path) -> list[str]:
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    return ["".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p)) for p in re.findall(r"<w:p[ >].*?</w:p>", xml)]


def _shown(text: str) -> str:
    return text.replace(NBSP, "{U+00A0}").replace(NNBSP, "{U+202F}").replace("\t", "{tab}")


def _answer(outcome: "oracle.Outcome") -> str:
    """What a script returned; its saved document is removed, so that the next run asks Word
    again rather than find it cached (a cached run returns no answer)."""
    if outcome.path is not None:
        outcome.path.unlink(missing_ok=True)
    return outcome.detail


def main(languages: list[str] = LANGUAGES) -> None:
    observed: dict = {"proofing": {}, "edges_fr": None, "autoformat": {}, "line_breaks": {}, "language": {}}
    with oracle.session() as session:
        for lang in languages:
            outcome = session.run_script(GRAMMAR, document(PROOFED, run_lang=lang), name=f"typo-{lang}",
                                         tag="grammar", timeout=280)
            if not outcome:
                observed["proofing"][lang] = {"error": outcome.detail[:200]}
                continue
            observed["proofing"][lang] = [_shown(t) for t in flagged(_answer(outcome), PROOFED)]
        outcome = session.run_script(GRAMMAR, document(EDGES, run_lang="fr-FR"), name="typo-edges", tag="grammar",
                                     timeout=280)
        answer = _answer(outcome)
        observed["edges_fr"] = {"flagged": [_shown(t) for t in flagged(answer, EDGES)],
                                "passed": [_shown(t) for t in EDGES if t not in flagged(answer, EDGES)]}
        typed = ['Il a dit "bonjour" hier.', "Il a dit «bonjour» hier.", "Vraiment? Oui : non !"]
        for lang in ("fr-FR", "fr-CA", "fr-CH", "fr-BE"):
            outcome = session.run_script(_script("auto format d"), document(typed, run_lang=lang),
                                         name=f"typo-af-{lang}", tag="autoformat")
            observed["autoformat"][lang] = [_shown(t) for t in _paragraph_texts(outcome.path)]
        for name, space in (("ordinary", " "), ("U+00A0", NBSP), ("U+202F", NNBSP)):
            lines = [f"{'mot ' * 12}{'m' * k}{space}{mark} suite de la phrase ici." for mark in "!?:;»"
                     for k in range(24)]
            outcome = session.export_pdf(document(lines, run_lang="fr-FR"), name=f"typo-breaks-{name}")
            starts: dict[str, int] = {}
            for page in oracle.pdf_pages(outcome.path):
                for line in re.split(r"\r\n|\r|\n", page):
                    if line.strip()[:1] in "!?:;»":
                        starts[line.strip()[:1]] = starts.get(line.strip()[:1], 0) + 1
            observed["line_breaks"][name] = {"lines_starting_with_a_mark": starts}
        cases = {"docDefaults fr-FR": dict(default_lang="fr-FR"),
                 "Normal fr-FR": dict(normal_lang="fr-FR"),
                 "Normal fr-FR, run en-GB": dict(normal_lang="fr-FR", run_lang="en-GB"),
                 "docDefaults fr-FR, Normal de-DE": dict(default_lang="fr-FR", normal_lang="de-DE"),
                 "themeFontLang fr-FR only": dict(default_lang=None, theme_lang="fr-FR"),
                 "themeFontLang fr-FR, docDefaults en-US": dict(theme_lang="fr-FR"),
                 "none": dict(default_lang=None)}
        for name, options in cases.items():
            outcome = session.run_script(LANGUAGE, document(["Bonjour le monde."], **options),
                                         name="typo-lang", tag=re.sub(r"\W+", "-", name))
            observed["language"][name] = _answer(outcome)
    OBSERVATIONS.write_text(json.dumps(observed, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main(sys.argv[1:] or LANGUAGES)
