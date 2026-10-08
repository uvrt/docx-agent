#!/usr/bin/env python3
"""Measure what Word does with a package whose kind (document or template, macro-enabled or
not) disagrees with its file's extension: the end-to-end trial's template finding
(ROADMAP.md, "Trial findings").

Through ``tests/oracle.py``'s machine-wide lock, Word, by AppleScript, opens and exports to
PDF each of these, staged under the extension named (a dialog -- Word's "cannot open" or
repair prompt -- blocks the export: ``rejected``):

* ``template-as-docx``: ``brand.dotx``'s package unchanged, as a ``.docx`` (what
  ``Document.open("x.dotx").save("y.docx")`` wrote before ``save`` set the kind);
* ``template-saved-docx``: ``Document.open(brand.dotx).save("y.docx")`` now;
* ``template-saved-dotx``: the same saved as ``.dotx`` (a template Word opens to edit);
* ``document-as-docm``: a document's package unchanged, as a ``.docm``;
* ``document-saved-docm``: ``save("y.docm")``: the macro-enabled document content type,
  no VBA project;
* ``macro-type-as-docx``: that package as a ``.docx``;

and ``word-saves-docm``: Word's own Save As macro-enabled document (``format
documentME``) of a document without macros -- the main part's content type it writes, and
whether it adds a VBA project.

    python tools/trial_probe.py            # every probe, written to tests/observations/trial-word.json
    python tools/trial_probe.py NAME...    # some, printed

Only outcomes and content types are recorded; nothing Word writes is committed.
"""

from __future__ import annotations

import hashlib
import io
import json
import sys
import tempfile
import warnings
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "src"))

import oracle  # noqa: E402
from docx_agent import Document  # noqa: E402

OBSERVATIONS = ROOT / "tests" / "observations" / "trial-word.json"
BRAND = ROOT / "tests" / "fixtures" / "generated" / "templates" / "brand.dotx"
#: Long enough for Word to open and export a one-page document on a cold start; a dialog
#: never returns.
TIMEOUT = 90

SAVE_DOCM = """
on run argv
    set inputPosixPath to item 1 of argv
    set outputPosixPath to item 2 of argv
    with timeout of 180 seconds
        tell application "Microsoft Word"
            activate
            open (POSIX file inputPosixPath)
            set d to active document
            save as d file name outputPosixPath file format format documentME
            close d saving no
            quit saving no
        end tell
    end timeout
end run
"""


def _saved(document: Document, suffix: str) -> bytes:
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / f"probe{suffix}"
        document.save(path)
        return path.read_bytes()


def _template() -> Document:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return Document.open(BRAND)


def _document() -> Document:
    return Document.new(template=BRAND, created=datetime(2026, 10, 5, tzinfo=timezone.utc))


def main_type(data: bytes) -> str:
    """The main part's content type in a package's ``[Content_Types].xml``."""
    from lxml import etree

    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        types = etree.fromstring(archive.read("[Content_Types].xml"))
    for node in types:
        if node.get("PartName") == "/word/document.xml":
            return node.get("ContentType")
    return ""


def cases() -> dict[str, tuple[bytes, str]]:
    """name -> (package bytes, the extension it is staged under)."""
    template = _template()
    document = _document()
    docm = _saved(document, ".docm")
    return {
        "template-as-docx": (template.to_bytes(), ".docx"),
        "template-saved-docx": (_saved(template, ".docx"), ".docx"),
        "template-saved-dotx": (_saved(template, ".dotx"), ".dotx"),
        "document-as-docm": (document.to_bytes(), ".docm"),
        "document-saved-docm": (docm, ".docm"),
        "macro-type-as-docx": (docm, ".docx"),
    }


def measure(names: list[str]) -> dict:
    out: dict = {}
    every = cases()
    with oracle.session() as word:
        for name, (data, suffix) in every.items():
            if names and name not in names:
                continue
            outcome = word.export_pdf(data, name=f"trial-{name}", timeout=TIMEOUT, suffix=suffix)
            out[name] = {"extension": suffix, "main_content_type": main_type(data), "outcome": outcome.outcome}
            print(name, outcome.outcome, f"{outcome.seconds:.1f}s", outcome.detail[:120])
        if not names or "word-saves-docm" in names:
            source = _document().to_bytes()
            digest = hashlib.sha256(source + SAVE_DOCM.encode()).hexdigest()[:16]
            target = oracle.ORACLE_DIR / f"trial-word-saves-docm-{digest}.docm"
            outcome = target.exists() and oracle.Outcome(True, "done", target) or word._run(
                ["-e", SAVE_DOCM], source, target, "trial-word-saves-docm", digest, 150)
            entry = {"outcome": outcome.outcome}
            if outcome:
                data = outcome.path.read_bytes()
                with zipfile.ZipFile(io.BytesIO(data)) as saved:
                    entry["main_content_type"] = main_type(data)
                    entry["vba_project"] = any("vbaProject" in name for name in saved.namelist())
            out["word-saves-docm"] = entry
            print("word-saves-docm", entry)
    return out


def main(names: list[str]) -> None:
    observed = measure(names)
    if names:
        print(json.dumps(observed, indent=1))
        return
    OBSERVATIONS.write_text(json.dumps({
        "_about": ("What Word 16.106 for Mac did with packages whose kind and extension agree or disagree "
                   "(tools/trial_probe.py): the outcome of opening and exporting each (done, or rejected: a "
                   "dialog blocked it), and the content type Word's own Save As macro-enabled document writes. "
                   "Regenerate with python tools/trial_probe.py."),
        **observed}, indent=1) + "\n", encoding="utf-8")
    print(f"wrote {OBSERVATIONS.relative_to(ROOT)}")


if __name__ == "__main__":
    main(sys.argv[1:])
