"""A template opened and saved as a document (ROADMAP.md, "Trial findings", 1): ``save`` sets
the main part's content type from the extension, ``validate(target=...)`` reports a
mismatch, and ``Document.open`` on a template warns.  What Word does with each kind under
each extension is measured (``tools/trial_probe.py``, ``tests/observations/trial-word.json``)
and held to Word by ``test_oracle_trial.py``."""

from __future__ import annotations

import io
import json
import warnings
import zipfile
from pathlib import Path

import pytest

from docx_agent import Document
from docx_agent.edit.authoring import MacrosDropped, TemplateOpened
from docx_agent.oxml.package import MAIN_CONTENT_TYPES

ROOT = Path(__file__).parent
TRIAL = ROOT / "fixtures" / "generated" / "trial"
TEMPLATES = ROOT / "fixtures" / "generated" / "templates"
WORD = json.loads((ROOT / "observations" / "trial-word.json").read_text(encoding="utf-8"))


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def opened(path: Path) -> Document:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", TemplateOpened)
        return Document.open(path)


def test_the_reproduction_validates_clean_alone_and_not_as_a_docx():
    """The trial's file: a template's package in a .docx, which Word refuses -- validate()
    found nothing; with the target it says why."""
    document = opened(TRIAL / "template-saved-as-docx.docx")
    assert document.package.kind == "dotx"
    assert document.validate() == []
    problems = document.validate(target="report.docx")
    assert [p.code for p in problems] == ["content-type-extension"]
    assert "Word refuses" in problems[0].detail and MAIN_CONTENT_TYPES["docx"] in problems[0].detail
    assert document.validate(target="report.dotx") == []


@pytest.mark.parametrize("source", [TRIAL / "template-saved-as-docx.docx", TEMPLATES / "brand.dotx"],
                         ids=["reproduction", "brand"])
@pytest.mark.parametrize("extension", [".docx", ".dotx", ".docm", ".dotm"])
def test_save_writes_the_kind_its_extension_names_and_nothing_else(tmp_path, source, extension):
    document = opened(source)
    before = document.to_bytes()
    target = tmp_path / f"out{extension}"
    document.save(target)
    assert document.to_bytes() == before                       # the open document is unchanged
    written = parts(target.read_bytes())
    ours = parts(before)
    assert set(written) == set(ours)
    assert {name for name in written if written[name] != ours[name]} <= {"[Content_Types].xml"}
    reopened = opened(target)
    assert reopened.package.kind == extension[1:]
    assert reopened.validate(target=target) == []
    assert [p.text for p in reopened.paragraphs()] == [p.text for p in document.paragraphs()]


def test_the_kinds_follow_what_word_accepts():
    """Each kind under the extension save gives it opened in Word; the mismatches it refuses."""
    assert WORD["template-as-docx"]["outcome"] == "rejected"
    assert WORD["template-saved-docx"]["outcome"] == "done"
    assert WORD["template-saved-docx"]["main_content_type"] == MAIN_CONTENT_TYPES["docx"]
    assert WORD["template-saved-dotx"]["outcome"] == "done"
    assert WORD["template-saved-dotx"]["main_content_type"] == MAIN_CONTENT_TYPES["dotx"]
    # .docm: a document's content type is refused; the macro-enabled one, without a VBA
    # project, opens -- and is what Word's own Save As .docm writes.
    assert WORD["document-as-docm"]["outcome"] == "rejected"
    assert WORD["document-saved-docm"]["outcome"] == "done"
    assert WORD["document-saved-docm"]["main_content_type"] == MAIN_CONTENT_TYPES["docm"]
    assert WORD["word-saves-docm"]["main_content_type"] == MAIN_CONTENT_TYPES["docm"]
    assert WORD["word-saves-docm"]["vba_project"] is False
    assert WORD["macro-type-as-docx"]["outcome"] == "rejected"


def test_a_vba_project_is_dropped_from_a_docx_written_and_kept_in_a_docm(tmp_path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", TemplateOpened)
        template = Document.open(TEMPLATES / "brand-macros.dotm")
    assert template.package.kind == "dotm"
    assert any("vbaProject" in name for name in template.package.part_names)
    assert [p.code for p in template.validate(target="x.docx")] == ["content-type-extension",
                                                                      "macros-in-macro-free-file"]
    with pytest.warns(MacrosDropped):
        template.save(tmp_path / "out.docx")
    written = opened(tmp_path / "out.docx")
    assert written.package.kind == "docx" and not any("vbaProject" in n for n in written.package.part_names)
    assert written.validate(target=tmp_path / "out.docx") == []
    template.save(tmp_path / "out.docm")
    kept = opened(tmp_path / "out.docm")
    assert kept.package.kind == "docm" and any("vbaProject" in n for n in kept.package.part_names)
    assert any("vbaProject" in name for name in template.package.part_names)   # the open one keeps it


def test_another_extension_or_a_file_object_gets_the_package_as_it_is(tmp_path):
    document = opened(TEMPLATES / "brand.dotx")
    document.save(tmp_path / "out.zip")
    assert (tmp_path / "out.zip").read_bytes() == document.to_bytes()
    buffer = io.BytesIO()
    document.save(buffer)
    assert buffer.getvalue() == document.to_bytes()


def test_opening_a_template_warns_and_points_to_new():
    with pytest.warns(TemplateOpened, match=r"Document\.new\(template=\.\.\.\)"):
        Document.open(TEMPLATES / "brand.dotx")
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        Document.new(template=TEMPLATES / "brand.dotx")       # making a document does not warn
        Document.open(ROOT / "fixtures" / "generated" / "pilot" / "agreement-summary.docx")


def test_save_as_template_takes_the_template_extension(tmp_path):
    from docx_agent import EditError

    document = Document.new(template=TEMPLATES / "brand.dotx")
    document.save_as_template(tmp_path / "mine.dotm")
    assert opened(tmp_path / "mine.dotm").package.kind == "dotm"
    with pytest.raises(EditError, match="save\\(\\) writes a document"):
        document.save_as_template(tmp_path / "mine.docx")
