"""Word itself on the end-to-end trial's fixes, opt-in (``pytest -m oracle``; ROADMAP.md,
"Trial findings").

* **The template reproduction**: the trial's file (a ``.dotx`` opened and saved as a
  ``.docx`` before ``save`` set the kind) opened with ``Document.open`` and saved again as a
  ``.docx`` opens in Word unprompted, and saved as a ``.dotx`` is a template Word opens.
  What Word does with each kind under each extension is ``tools/trial_probe.py``'s.
* **A section merged** from the trial's bulletin into its handbook in the handbook's styles
  (``copy_blocks(unmapped="body")``) exports unprompted, its footnote in it.
* **A section moved, tracked** (``move_blocks(section_blocks(...))``) exports unprompted, in
  Word's final view the section at its new place, once.
"""

from __future__ import annotations

import warnings
from pathlib import Path

import pytest

import oracle
from docx_agent import Document
from docx_agent.edit.authoring import TemplateOpened

pytestmark = pytest.mark.oracle

ROOT = Path(__file__).resolve().parents[1]
TRIAL = ROOT / "tests" / "fixtures" / "generated" / "trial"


@pytest.fixture(scope="module")
def word():
    try:
        with oracle.session() as session:
            yield session
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))


def _text(outcome: oracle.Outcome) -> str:
    assert outcome, f"Word did not export it: {outcome.outcome} {outcome.detail}"
    return "\n".join(oracle.pdf_pages(outcome.path))


def _saved(document: Document, tmp_path: Path, name: str) -> bytes:
    target = tmp_path / name
    document.save(target)
    return target.read_bytes()


def test_the_reproduction_saved_as_a_docx_opens_in_word(word, tmp_path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", TemplateOpened)
        document = Document.open(TRIAL / "template-saved-as-docx.docx")
    data = _saved(document, tmp_path, "report.docx")
    assert Document.open(data).package.kind == "docx"
    assert "Template guidance" in _text(word.export_pdf(data, name="trial-repro-docx"))


def test_the_reproduction_saved_as_a_dotx_is_a_template_word_opens(word, tmp_path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", TemplateOpened)
        document = Document.open(TRIAL / "template-saved-as-docx.docx")
        data = _saved(document, tmp_path, "report.dotx")
        assert Document.open(data).package.kind == "dotx"
    assert "Template guidance" in _text(word.export_pdf(data, name="trial-repro-dotx", suffix=".dotx"))


def test_a_tracked_section_move_opens_in_word(word):
    import sys

    sys.path.insert(0, str(ROOT / "tests"))
    from test_trial_sections import heading, policy

    document = policy()
    with document.tracking(author="Claude", date="2026-10-05T09:00:00Z"):
        document.move_blocks(document.section_blocks(heading(document, "3 Data retention")),
                             after=document.section_blocks(heading(document, "1 Purpose"))[-1])
    text = _text(word.export_pdf(document.to_bytes(), name="trial-section-move"))
    # Word's final view: the section where it was moved to, once.
    assert text.index("3 Data retention") < text.index("2 Access control")
    assert text.count("Shred paper") == 1


def test_the_merged_section_opens_in_word(word):
    handbook = Document.open(TRIAL / "field-handbook.docx")
    bulletin = Document.open(TRIAL / "safety-bulletin.docx")
    handbook.copy_blocks(bulletin, bulletin.section_blocks("p:3B212964"), at="before:p:69E36C9F",
                         unmapped="body")
    text = _text(word.export_pdf(handbook.to_bytes(), name="trial-merged-section"))
    assert "Working at height" in text and "stepladders on uneven ground" in text
