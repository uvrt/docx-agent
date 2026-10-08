"""Structural validity, after every edit set on every fixture.

An edit may leave a problem the document already had -- ids-and-markup.docx repeats a
paraId and has one out of range, on purpose -- but must never add one.
"""

from __future__ import annotations

from docx_agent import Document
from docx_agent.validate import Problem, check

from test_edit import edit_set

#: The problems a fixture is made with.
KNOWN = {
    "style-document": {
        Problem("style-missing", "word/styles.xml", "Heading2 link Heading2Char"),
    },
    "ids-and-markup": {
        Problem("paraid-out-of-range", "word/document.xml", "80000001"),
        Problem("paraid-repeated", "word/document.xml", "2B3C4D05"),
    },
}


def test_fixtures_are_valid_as_read(docx_path):
    assert set(check(Document.open(docx_path).package)) == KNOWN.get(docx_path.stem, set())


def test_edits_add_no_problem(docx_path):
    document = Document.open(docx_path)
    before = set(check(document.package))
    edit_set(document)
    after = set(check(Document.open(document.to_bytes()).package))
    assert after - before == set()


def test_editing_the_repeats_and_the_out_of_range_id_clears_their_problems(markup_doc):
    document = Document.open(markup_doc)
    document.paragraph("p:2B3C4D05#1").set_text("Second, edited.")
    document.paragraph("p:80000001").set_text("Now in range.")
    assert check(document.package) == []


def test_the_checks_see_what_they_check(markup_doc):
    """A deliberately broken document: each check fires."""
    document = Document.open(markup_doc)
    body = document.package.tree("word/document.xml").find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")
    paragraph = document.paragraph("p:1A2B3C02")._element
    properties = paragraph.makeelement(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    paragraph.append(properties)                       # pPr after the runs
    body.append(body[-1].__copy__())                    # a second body sectPr
    document.package.mark_dirty("word/document.xml")
    codes = {problem.code for problem in check(document.package)}
    assert {"child-order", "body-has-two-sections"} <= codes
