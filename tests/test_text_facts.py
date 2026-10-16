"""Text facts in a Word document (round 5, T4): ``{{d:scores/Beta/rating}}`` in a text tool's
argument is written as the value, in the session's locale, and the result lists it with its
sources -- as in pptx-agent.  A Word document does not keep text facts in the file yet (the
deck does): a paragraph's id changes when it is edited, so the key a table's record uses
does not carry over.  A deck's slide-number reference (``{{slide:s:262}}``) is refused
(round 6, R8).

The data is made up: the toy dataset and grid of :mod:`test_tables_from_data`.
"""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.textfacts", reason="needs ooxml-edit 0.33 (text facts)")

from ooxml_common.locales import Locale  # noqa: E402
from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

from test_tables_from_data import GRID, _session, call, ds_put, failed, ok  # noqa: E402

NL_BE = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


def first_paragraph(session):
    return session.entry("d1").document.paragraphs()[0]


def test_the_text_tools_declare_their_text(toolbox):
    assert toolbox.tools["word_set_text"].facts == ("text", "items[].text")
    assert toolbox.tools["word_insert_text"].facts == ("text", "items[].text")
    assert toolbox.tools["word_insert_markdown"].facts == ("markdown",)
    assert toolbox.tools["replace_text"].facts == ("replace",)


def test_set_text_writes_the_value_in_the_locale_and_reports_it(toolbox):
    session = _session(toolbox, locale=NL_BE)
    ok(call(toolbox, session, "ds_put", id="scores", typed="yes"))
    target = first_paragraph(session).id
    result = ok(call(toolbox, session, "word_set_text", doc="d1", target=target,
                     text="Gamma {{d:scores/Gamma/rating:0.0}}, due {{d:scores/Gamma/due}}"))
    assert first_paragraph(session).text == "Gamma 1,3, due 15/3/2026"
    assert result.data["facts"] == [
        {"from": "d:scores/Gamma/rating", "text": "1,3", "sources": ["G!B4"]},
        {"from": "d:scores/Gamma/due", "text": "15/3/2026", "sources": ["G!C4"]}]


def test_markdown_insert_and_replace_take_references_too(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
    ok(call(toolbox, session, "word_insert_markdown", doc="d1", at="end",
            markdown="Alpha rated **{{t:checks/Alpha/rating}}**."))
    texts = [p.text for p in session.entry("d1").document.paragraphs()]
    assert "Alpha rated 3.5." in texts
    result = ok(call(toolbox, session, "replace_text", doc="d1", find="3.5",
                     replace="{{d:scores/Beta/rating}}", expect="one"))
    assert result.data["facts"][0]["text"] == "4"
    assert "Alpha rated 4." in [p.text for p in session.entry("d1").document.paragraphs()]


def test_a_slide_reference_is_refused_and_an_escaped_one_written(toolbox):
    session = _session(toolbox)
    target = first_paragraph(session).id
    before = first_paragraph(session).text
    error = failed(call(toolbox, session, "word_set_text", doc="d1", target=target,
                        text="See slide {{slide:s:262}}"), "invalid_arguments")
    assert error.message.startswith("{{slide:s:262}}: slide references are for decks; a "
                                    "Word document has no slides.")
    assert error.field == "text"
    assert first_paragraph(session).text == before
    ok(call(toolbox, session, "word_set_text", doc="d1", target=target,
            text="See \\{{slide:s:262}}"))
    assert first_paragraph(session).text == "See {{slide:s:262}}"


def test_an_escape_writes_the_reference_and_a_bad_one_changes_nothing(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
    target = first_paragraph(session).id
    ok(call(toolbox, session, "word_set_text", doc="d1", target=target,
            text="Literally \\{{d:scores/Beta/rating}} and {{name}}"))
    assert first_paragraph(session).text == "Literally {{d:scores/Beta/rating}} and {{name}}"
    error = failed(call(toolbox, session, "word_set_text", doc="d1", target=target,
                        text="{{d:scores/Delta/rating}}"), "not_found")
    assert "d:scores/Beta" in error.valid_options
    assert first_paragraph(session).text.startswith("Literally")
