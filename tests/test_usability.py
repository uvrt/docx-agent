"""The public API the end-to-end pilot looked for (ROADMAP.md, "Usability (end-to-end
pilot)"): validation from a Document, render options checked with a useful error."""

from __future__ import annotations

import pytest

from docx_agent import Document, EditError
from docx_agent.validate import Problem, check
from conftest import FIXTURE_DIR

PILOT = FIXTURE_DIR / "generated" / "pilot" / "agreement-summary.docx"


def test_validate_is_the_checker_on_the_document():
    document = Document.open(PILOT)
    assert document.validate() == check(document.package) == []
    with document.tracking(author="Claude"):
        document.anchor("24 months").replace("36 months")
    assert document.validate() == []


def test_save_with_validate_refuses_a_document_with_problems(tmp_path, monkeypatch):
    document = Document.open(PILOT)
    document.save(tmp_path / "ok.docx", validate=True)
    assert (tmp_path / "ok.docx").exists()
    monkeypatch.setattr(Document, "validate", lambda self: [Problem("made-up", "word/document.xml", "x")])
    with pytest.raises(EditError, match=r"not saved: 1 validity problem.*made-up"):
        document.save(tmp_path / "bad.docx", validate=True)
    assert not (tmp_path / "bad.docx").exists()


@pytest.mark.parametrize("method", ["render_svg", "render_png", "layout"])
def test_an_unknown_render_option_names_the_accepted_ones(method):
    document = Document.open(PILOT)
    with pytest.raises(TypeError) as raised:
        getattr(document, method)(view="markup")
    message = str(raised.value)
    assert f"{method}() got unexpected keyword argument(s) 'view'" in message
    assert "width" in message and "font_dirs" in message
    assert 'to_markdown(view="markup")' in message
    assert ("pages" in message) == (method != "layout")


def test_accepted_render_options_still_pass_through():
    document = Document.open(PILOT)
    assert document.render_svg(pages=[1], glyph_size="device")[0].startswith("<svg")
