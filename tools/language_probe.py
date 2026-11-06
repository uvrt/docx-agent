#!/usr/bin/env python3
"""Measure what Word writes for a document's language, and which language tags it knows
(round 35, R24: :meth:`docx_agent.Document.set_language`, ``docx_agent.edit.language``).

Probe documents are written here and Word -- through ``tests/oracle.py``'s machine-wide
lock -- opens each and, by AppleScript:

* ``tags``: sets each paragraph of a document to one of the languages of Word's
  dictionary (``WdLanguageID``, every value but *none* and *no proofing*) and saves it; the
  ``w:lang`` Word wrote for each says the tag it writes for that language and which of
  ``w:val`` / ``w:eastAsia`` / ``w:bidi`` holds it.  Two passes, over two sets of document
  defaults, so that a language equal to a default (written as nothing) is seen in the other.
* ``reads``: the language Word reports for text whose only ``w:lang`` is
  ``w:docDefaults``' with a given value (case, ``_``, a script, unknown tags...).
* ``actions``: what Word writes into an open document -- styles with and without their own
  ``w:lang``, a run with its own -- when a script sets the language of the Normal style, of
  all the text, or of the Normal template (``language ID`` of ``normal template``: what a
  default for new documents is; it also writes Normal.dotm, which this probe puts back,
  byte for byte, from the copy it takes first).

    python tools/language_probe.py      # writes tests/observations/language-word.json

Word's Language dialog's *Default...* button is not driven here: System Events sees none of
Word's windows while the screen is locked, and a dialog left open blocks Word's scripting.
Nothing Word writes is committed: only these facts.  Back up Normal.dotm first anyway.
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

OBSERVATIONS = ROOT / "tests" / "observations" / "language-word.json"
NORMAL = oracle.GROUP_CONTAINER / "User Content.localized" / "Templates.localized" / "Normal.dotm"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: The two sets of document defaults the ``tags`` passes run over.
DEFAULTS = ('<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>',
            '<w:lang w:val="de-DE" w:eastAsia="zh-CN" w:bidi="he-IL"/>')

#: Tags read back from ``w:docDefaults`` (``reads``).
READS = ["fr-BE", "FR-be", "fr_BE", "fr-Latn-BE", "fr-be-x-test", "fr", "fr-XX", "nl-BE", "nl", "en-US", "en",
         "xx-YY", "zz", "x-none", "sr-Latn-RS", "de-CH-1996"]

#: Paragraphs of the ``actions`` document, and their styles: Normal with no language,
#: Heading 1 with its own, a custom paragraph style with its own, a character style with
#: its own, a run with its own.
ACTION_STYLES = (
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
    '<w:qFormat/><w:rPr><w:b/><w:lang w:val="en-GB"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Bodyde"><w:name w:val="Body de"/>'
    '<w:basedOn w:val="Normal"/><w:rPr><w:lang w:val="de-DE" w:eastAsia="ja-JP"/></w:rPr></w:style>'
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
    '<w:name w:val="Default Paragraph Font"/></w:style>'
    '<w:style w:type="character" w:customStyle="1" w:styleId="Tagit"><w:name w:val="Tag it"/>'
    '<w:rPr><w:i/><w:lang w:val="it-IT"/></w:rPr></w:style>')
ACTION_BODY = (
    '<w:p><w:r><w:t>Plain paragraph one.</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading text</w:t></w:r></w:p>'
    '<w:p><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Run tagged en-US.</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:pStyle w:val="Bodyde"/></w:pPr><w:r><w:t>Body de paragraph.</w:t></w:r></w:p>'
    '<w:p><w:r><w:rPr><w:rStyle w:val="Tagit"/></w:rPr><w:t>Tag it run.</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Last plain paragraph.</w:t></w:r></w:p>')
ACTIONS = {
    "nothing": "",
    "Normal style": "set language ID of (get Word style style normal of d) to belgian french",
    "all text": "set language ID of text object of d to belgian french",
    "normal template": "set language ID of normal template to belgian french",
}


def document(body: str, *, default: str | None = DEFAULTS[0], styles: str = "",
             theme: str | None = '<w:themeFontLang w:val="en-US"/>') -> bytes:
    styles_xml = (f'{DECL}<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" '
                  f'w:hAnsi="Calibri"/><w:sz w:val="22"/>{default or ""}</w:rPr></w:rPrDefault></w:docDefaults>'
                  + (styles or '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
                     '</w:style>') + '</w:styles>')
    settings = (f'{DECL}<w:settings xmlns:w="{W}"><w:compat><w:compatSetting w:name="compatibilityMode" '
                f'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>{theme or ""}</w:settings>')
    body += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
             'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    parts = {
        "[Content_Types].xml": content_types({"word/document.xml": WML + ".document.main+xml",
                                              "word/styles.xml": WML + ".styles+xml",
                                              "word/settings.xml": WML + ".settings+xml"}),
        "_rels/.rels": relationships([("rId1", REL + "officeDocument", "word/document.xml", False)]),
        "word/_rels/document.xml.rels": relationships([("rId1", REL + "styles", "styles.xml", False),
                                                       ("rId2", REL + "settings", "settings.xml", False)]),
        "word/document.xml": f"{DECL}<w:document {NS}><w:body>{body}</w:body></w:document>",
        "word/styles.xml": styles_xml,
        "word/settings.xml": settings,
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            archive.writestr(zipfile.ZipInfo(name, date_time=(2026, 11, 6, 12, 0, 0)), data)
    return buffer.getvalue()


def _script(body: str) -> str:
    return f'''
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set out to ""
  with timeout of 280 seconds
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


def _language_names() -> list[str]:
    """The values of Word's ``WdLanguageID``, from its scripting dictionary."""
    import subprocess

    sdef = subprocess.run(["sdef", str(oracle.WORD_APP)], capture_output=True, text=True, check=True).stdout
    block = re.search(r'<enumeration name="WdLanguageID".*?</enumeration>', sdef, re.S).group(0)
    return [n for n in re.findall(r'<enumerator name="([^"]+)"', block)
            if n not in ("language none", "language no proofing")]


def _xml(path: Path, part: str) -> str:
    return zipfile.ZipFile(path).read(part).decode("utf-8")


def _langs(xml: str) -> list[dict[str, str]]:
    return [dict(re.findall(r'w:(\w+)="([^"]*)"', attrs)) for attrs in re.findall(r"<w:lang ([^>]*?)/>", xml)]


def _answer(outcome: "oracle.Outcome") -> str:
    if outcome.path is not None:
        outcome.path.unlink(missing_ok=True)
    return outcome.detail


def tags(session: "oracle.Session", names: list[str]) -> dict[str, dict[str, str]]:
    """Per language name, the ``w:lang`` attribute Word wrote for it, over both defaults."""
    body = "".join(f"<w:p><w:r><w:t>Sample {i}.</w:t></w:r></w:p>" for i in range(len(names)))
    sets = "\n".join(f"      try\n        set language ID of text object of paragraph {i + 1} of d to {name}\n"
                     "      end try" for i, name in enumerate(names))
    found: dict[str, dict[str, str]] = {}
    for k, default in enumerate(DEFAULTS):
        outcome = session.run_script(_script(sets), document(body, default=default), name="lang-tags",
                                     tag=f"pass{k}", timeout=290)
        if not outcome:
            raise SystemExit(f"tags pass {k}: {outcome.detail[:300]}")
        xml = _xml(outcome.path, "word/document.xml")
        outcome.path.unlink(missing_ok=True)
        paragraphs = re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
        for name, paragraph in zip(names, paragraphs):
            for attrs in _langs(paragraph)[:1]:
                found.setdefault(name, {}).update(attrs)
    return found


READ = _script('set out to ((language ID of text object of paragraph 1 of d) as string)\n'
               '      try\n        set out to out & " / " & (name of (retrieve language '
               '(language ID of text object of paragraph 1 of d)))\n      end try')


def main() -> None:
    observed: dict = {"tags": {}, "reads": {}, "actions": {}}
    names = _language_names()
    template = NORMAL.read_bytes()
    try:
        with oracle.session() as session:
            written = tags(session, names)
            slots: dict[str, set[str]] = {"val": set(), "eastAsia": set(), "bidi": set()}
            for name in names:
                for slot, tag in written.get(name, {}).items():
                    slots[slot].add(tag)
            observed["tags"] = {slot: sorted(found, key=str.lower) for slot, found in slots.items()}
            observed["tags_not_written"] = sorted(n for n in names if n not in written)
            for tag in READS:
                outcome = session.run_script(READ, document('<w:p><w:r><w:t>Bonjour.</w:t></w:r></w:p>',
                                                            default=f'<w:lang w:val="{tag}"/>', theme=None),
                                             name="lang-read", tag=re.sub(r"\W+", "-", tag))
                observed["reads"][tag] = _answer(outcome)
            reads_back = "\n".join(f'      set out to out & ((language ID of text object of paragraph {i} of d) '
                                   'as string) & linefeed' for i in range(1, 7))
            for action, line in ACTIONS.items():
                outcome = session.run_script(_script(line + "\n" + reads_back),
                                             document(ACTION_BODY, styles=ACTION_STYLES),
                                             name="lang-act", tag=re.sub(r"\W+", "-", action))
                if not outcome:
                    observed["actions"][action] = {"error": outcome.detail[:200]}
                    continue
                styles = _xml(outcome.path, "word/styles.xml")
                defaults = re.search(r"<w:docDefaults>.*?</w:docDefaults>", styles, re.S).group(0)
                own = {}
                for match in re.finditer(r'<w:style [^>]*w:styleId="([^"]*)".*?</w:style>', styles, re.S):
                    found = _langs(match.group(0))
                    if found:
                        own[re.search(r'<w:name w:val="([^"]*)"', match.group(0)).group(1)] = found[0]
                observed["actions"][action] = {
                    "docDefaults": _langs(defaults),
                    "styles": own,
                    "themeFontLang": re.findall(r'<w:themeFontLang ([^>]*?)/>', _xml(outcome.path, "word/settings.xml")),
                    "runs": _langs(_xml(outcome.path, "word/document.xml")),
                    "proofed as": [line for line in outcome.detail.splitlines() if line.strip()],
                    "Normal.dotm changed": NORMAL.read_bytes() != template,
                }
                outcome.path.unlink(missing_ok=True)
                if NORMAL.read_bytes() != template:
                    # What Word wrote into the template: its defaults' and styles' languages.
                    saved = zipfile.ZipFile(io.BytesIO(NORMAL.read_bytes())).read("word/styles.xml").decode()
                    defaults = re.search(r"<w:docDefaults>.*?</w:docDefaults>", saved, re.S).group(0)
                    observed["actions"][action]["Normal.dotm"] = {
                        "docDefaults": _langs(defaults),
                        "styles": {m.group(1): _langs(m.group(0))[0] for m in
                                   re.finditer(r'<w:style [^>]*w:styleId="([^"]*)".*?</w:style>', saved, re.S)
                                   if _langs(m.group(0))}}
                    NORMAL.write_bytes(template)
    finally:
        if NORMAL.read_bytes() != template:
            NORMAL.write_bytes(template)
    OBSERVATIONS.write_text(json.dumps(observed, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
