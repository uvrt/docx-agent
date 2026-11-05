"""Spacing around punctuation by the text's language (round 35, R21:
:mod:`docx_agent.edit.typography`, measured in ``tests/observations/typography-word.json``).

Text a call writes in French gets a no-break space (U+00A0) before ``: ; ! ?`` and inside
guillemets -- an ordinary space it wrote made U+00A0, a missing one put in -- and nothing
else changes; ``check``'s ``punctuation_spacing`` reports what is left, with fixes.  The
documents are made up here.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

import pytest

from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.edit import typography
from docx_agent.edit.typography import NBSP, Cascade, find, problems, rule_of
from docx_agent.tools import GROUPS, TOOLS, docx_format

ROOT = Path(__file__).resolve().parent.parent
OBSERVED = json.loads((ROOT / "tests" / "observations" / "typography-word.json").read_text(encoding="utf-8"))
NB = "{U+00A0}"


def french(*texts: str, language: str = "fr-FR", spaced: bool = True) -> Document:
    document = Document.new(language=language)
    document.typography = spaced
    first = document.paragraphs()[0]
    if texts:
        first.set_text(texts[0])
        last = first.id
        for text in texts[1:]:
            last = document.insert_paragraph(text, after=last).id
    return document


def texts(document: Document, story: str = "body") -> list[str]:
    return [p.text for p in document.paragraphs(story)]


# -- the measurements the rules stand on -------------------------------------------------------


def test_the_rules_are_the_measured_ones():
    proofing = OBSERVED["proofing"]
    # French (and fr alone, fr-CH): no space before a mark, or inside guillemets, is flagged;
    # an ordinary space, U+00A0 and U+202F all pass.
    for lang in ("fr", "fr-FR", "fr-CH"):
        flagged = proofing[lang]
        assert "Bonjour: le monde." in flagged and "Vraiment? Oui." in flagged
        assert "Il a dit «bonjour» hier." in flagged
        assert not any(t.startswith(("Bonjour :", f"Bonjour{NB}", "Bonjour{U+202F}")) for t in flagged)
        assert rule_of(lang) == "full"
    # Canadian French: the guillemets only.
    assert proofing["fr-CA"] == ["Il a dit «bonjour» hier."]
    assert rule_of("fr-CA") == "guillemets"
    # English, German, Dutch, Spanish, Italian, Portuguese flag any space before a mark.
    for lang in ("en-US", "en-GB", "de-DE", "nl-NL", "es-ES", "it-IT", "pt-PT"):
        flagged = proofing[lang]
        assert "Vraiment ? Oui." in flagged and f"Vraiment{NB}? Oui." in flagged
        assert "Vraiment? Oui." not in flagged
        assert lang.split("-")[0] in typography.NO_SPACE_BEFORE and rule_of(lang) is None
        assert (f"Il a dit «{NB}bonjour{NB}» hier." in flagged) == (lang.split("-")[0] in typography.NO_SPACE_INSIDE)
    # Line breaks: an ordinary space lets a mark start a line; a no-break one never does.
    assert OBSERVED["line_breaks"]["ordinary"]["lines_starting_with_a_mark"]
    assert not OBSERVED["line_breaks"]["U+00A0"]["lines_starting_with_a_mark"]
    assert not OBSERVED["line_breaks"]["U+202F"]["lines_starting_with_a_mark"]
    # What Word writes: U+00A0 inside guillemets.
    assert OBSERVED["autoformat"]["fr-FR"][0] == f"Il a dit «{NB}bonjour{NB}» hier."


def test_marks_between_words_are_left_alone_as_their_fix_is_not_a_space():
    """Word flags 12:30 and x?y in French too; URLs, paths and (?) it does not."""
    edges = OBSERVED["edges_fr"]
    assert "Rendez-vous à 12:30 demain." in edges["flagged"]
    assert "Voir http://example.com/page pour plus." in edges["passed"]
    document = french("Rendez-vous à 12:30, voir http://example.com/page ou C:\\dossier (?) x?y v2:beta.")
    assert NBSP not in texts(document)[0]


# -- writing -----------------------------------------------------------------------------------


def test_french_text_gets_no_break_spaces_where_written():
    document = french("Bonjour : le monde! Pourquoi? « oui » et «non». Quoi?! Fin :")
    assert texts(document) == [f"Bonjour{NBSP}: le monde{NBSP}! Pourquoi{NBSP}? «{NBSP}oui{NBSP}» et "
                               f"«{NBSP}non{NBSP}». Quoi{NBSP}?! Fin{NBSP}:"]
    assert document.typography_spaced == 9


def test_a_space_already_there_is_not_changed_by_a_later_edit():
    document = french("Avant : après", spaced=False)
    document.typography = True
    paragraph = document.paragraphs()[0]
    document.insert_text(f"{paragraph.id}@13", " ; suite")
    assert texts(document) == [f"Avant : après{NBSP}; suite"]   # the old ' :' stays


def test_other_languages_are_written_as_given():
    document = french("Note : this ! and « that »", language="en-US")
    assert texts(document) == ["Note : this ! and « that »"]
    assert document.typography_spaced == 0


def test_canadian_french_gets_the_guillemets_only():
    document = french("Il a dit «oui» : non!", language="fr-CA")
    assert texts(document) == [f"Il a dit «{NBSP}oui{NBSP}» : non!"]


def test_belgian_french_takes_frenchs_rule():
    document = french("Vraiment?", language="fr-BE")
    assert texts(document) == [f"Vraiment{NBSP}?"]


def test_the_language_is_the_runs_through_its_styles():
    document = french("Un.", "Deux.", language="en-US")
    styles = document.package.tree(document.package.styles_part())
    normal = next(s for s in styles.iter(f"{{{typography._W[1:-1]}}}style")
                  if s.get(typography._W + "styleId") == "Normal")
    from lxml import etree

    rpr = normal.find(typography._W + "rPr")
    if rpr is None:
        rpr = etree.SubElement(normal, typography._W + "rPr")
    etree.SubElement(rpr, typography._W + "lang").set(typography._W + "val", "fr-FR")
    cascade = Cascade(document)
    element = document.paragraphs()[0]._element
    run = element.find(typography._W + "r")
    assert cascade.of(element, run)[0] == "fr-FR"
    lang = etree.SubElement(run.find(typography._W + "rPr") if run.find(typography._W + "rPr") is not None
                            else etree.SubElement(run, typography._W + "rPr"), typography._W + "lang")
    lang.set(typography._W + "val", "en-GB")
    assert Cascade(document).of(element, run)[0] == "en-GB"


def test_text_without_proofing_or_in_a_field_is_not_looked_at():
    assert find("Bonjour : x", lambda i: "fr-FR", lambda i: False) == []
    spots = find("Bonjour : x", lambda i: "fr-FR", lambda i: True)
    assert [(s.kind, s.at) for s in spots] == [("breakable", 7)]


def test_undo_takes_the_text_and_its_spacing_together():
    document = french("Début.")
    paragraph = document.paragraphs()[0]
    paragraph.set_text("Début : fin")
    assert texts(document) == [f"Début{NBSP}: fin"]
    document.undo()
    assert texts(document) == ["Début."]


def test_off_writes_the_text_as_given():
    document = french("Bonjour : le monde!", spaced=False)
    assert texts(document) == ["Bonjour : le monde!"]


# -- the tools ---------------------------------------------------------------------------------


CLOCK = lambda: datetime.datetime(2026, 11, 5, 12, 0)  # noqa: E731


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[docx_format()], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box, document: Document | None = None):
    session = box.session(clock=CLOCK)
    session.open((document or french("Début.")).to_bytes(), "rapport.docx")
    return session


def _ok(box, session, name, **arguments):
    result = box.dispatch(session, name, arguments)
    assert result.ok, json.dumps(result.to_json(), default=str)[:800]
    return result


def _spaced_note(result) -> list[str]:
    return [w for w in result.warnings if "no-break" in w]


def test_every_writing_tool_spaces_what_it_writes(box):
    session = _session(box)
    first = session.entry("d1").document.paragraphs()[0].id
    result = _ok(box, session, "word_insert_markdown", doc="d1", at="end", markdown=(
        "# Titre: un\n\nBonjour : le **monde**!\n\n- point: un\n\n| Col: A | B |\n|---|---|\n| x ? | «y» |\n\n"
        "Code `a ? b : c` ici!\n\n```\nx = a ? b : c;\n```\n\nNote[^1].\n\n[^1]: Voir: ici!\n"))
    assert _spaced_note(result) and _spaced_note(result)[0].startswith("11 space(s)")
    _ok(box, session, "word_set_text", doc="d1", target=first, text="Début : fin")
    _ok(box, session, "word_insert_text", doc="d1", find="fin", where="after", text=" encore?")
    _ok(box, session, "replace_text", doc="d1", find="encore", replace="toujours!", expect="one")
    document = session.entry("d1").document
    _ok(box, session, "word_edit_table", doc="d1", table=document.tables()[0].id, action="set_cells",
        cells=[{"row": 1, "col": 1, "text": "Cellule: oui\nAutre?"}])
    _ok(box, session, "word_comments", doc="d1", action="add", author="Relecteur",
        items=[{"target": first, "text": "Pourquoi? Voir: ici"}])
    _ok(box, session, "word_headers_footers", doc="d1", section="s:body", which="header", action="set",
        text="En-tête: rapport")
    _ok(box, session, "word_notes", doc="d1", action="insert", kind="endnote", find="Début", text="Fin: note!")
    _ok(box, session, "word_drawings", doc="d1", action="insert_text_box", find="Début", text="Boîte: texte!",
        width=200, height=50)
    document = session.entry("d1").document
    body = texts(document)
    assert body[0] == f"Début\ufffc\ufffc{NBSP}: fin toujours{NBSP}!{NBSP}?"
    assert f"Titre{NBSP}: un" in body and f"Bonjour{NBSP}: le monde{NBSP}!" in body
    assert f"point{NBSP}: un" in body and f"Col{NBSP}: A" in body and f"x{NBSP}?" in body
    assert f"Cellule{NBSP}: oui" in body and f"Autre{NBSP}?" in body
    assert f"Code a ? b : c ici{NBSP}!" in body and "x = a ? b : c;" in body     # code as written
    assert f"Boîte{NBSP}: texte{NBSP}!" in body
    assert texts(document, "comments") == [f"Pourquoi{NBSP}? Voir{NBSP}: ici"]
    assert f"En-tête{NBSP}: rapport" in [p.text for s in document.stories if s.name.startswith("header")
                                          for p in document.paragraphs(s.name)]
    assert texts(document, "footnotes")[-1].endswith(f"Voir{NBSP}: ici{NBSP}!")
    assert texts(document, "endnotes")[-1].endswith(f"Fin{NBSP}: note{NBSP}!")
    assert document.validate() == []


def test_a_control_filled_is_spaced(box):
    session = _session(box)
    made = _ok(box, session, "word_controls", doc="d1", action="insert", find="Début", type="text", title="Titre")
    result = _ok(box, session, "word_controls", doc="d1", action="fill", control=made.data["id"],
                 value="Valeur: oui!")
    assert texts(session.entry("d1").document) == [f"DébutValeur{NBSP}: oui{NBSP}!."]
    assert _spaced_note(result)


def test_tracked_writing_is_spaced_inside_the_insertion(box):
    session = _session(box)
    _ok(box, session, "word_set_tracking", doc="d1", on=True, author="Relecteur")
    first = session.entry("d1").document.paragraphs()[0].id
    _ok(box, session, "word_set_text", doc="d1", target=first, text="Début : fin!")
    document = session.entry("d1").document
    assert texts(document) == [f"Début{NBSP}: fin{NBSP}!"]
    assert document.validate() == []
    document.reject_all()
    assert texts(document) == ["Début."]


def test_the_option_off_writes_as_given_and_says_nothing():
    with Toolbox(TOOLS, formats=[docx_format(typography=False)], groups=GROUPS, workers=0) as toolbox:
        session = _session(toolbox)
        first = session.entry("d1").document.paragraphs()[0].id
        result = _ok(toolbox, session, "word_set_text", doc="d1", target=first, text="Début : fin!")
        assert texts(session.entry("d1").document) == ["Début : fin!"]
        assert not _spaced_note(result)


def test_writing_what_is_there_already_is_a_no_op(box):
    session = _session(box, french("Début : fin!"))
    first = session.entry("d1").document.paragraphs()[0].id
    result = _ok(box, session, "word_set_text", doc="d1", target=first, text=f"Début{NBSP}: fin{NBSP}!")
    assert not result.changed and not _spaced_note(result)


# -- check -------------------------------------------------------------------------------------


def test_problems_list_what_is_left_with_fixes_that_clear_them(box):
    document = french("Bonjour : le monde! « oui » et «non».", "Rien à voir.", "Une : deux : trois", spaced=False)
    found = problems(document)
    assert [e["para"] for e in found] == [document.paragraphs()[0].id, document.paragraphs()[2].id]
    assert [m["kind"] for m in found[0]["marks"]] == ["breakable", "missing", "breakable", "breakable",
                                                      "missing", "missing"]
    assert found[0]["language"] == "fr-FR"
    session = _session(box, document)
    report = _ok(box, session, "check", doc="d1", include=["layout"]).data["layout"]["punctuation_spacing"]
    assert [e["para"] for e in report] == [e["para"] for e in found]
    for entry in report:
        assert entry["fix"]["tool"] == "replace_text"
        for call in entry["fix"]["calls"]:
            _ok(box, session, "replace_text", doc="d1", **call)
    assert _ok(box, session, "check", doc="d1", include=["layout"]).data["layout"]["punctuation_spacing"] == []
    assert texts(session.entry("d1").document)[0] == (f"Bonjour{NBSP}: le monde{NBSP}! «{NBSP}oui{NBSP}» et "
                                                      f"«{NBSP}non{NBSP}».")


def test_check_reports_the_spaces_english_does_not_want():
    document = french("Note : this! Really ?", "Fine: yes.", language="en-US")
    found = problems(document)
    assert len(found) == 1 and found[0]["language"] == "en-US"
    assert [m["kind"] for m in found[0]["marks"]] == ["space_before", "space_before"]
    assert found[0]["fix"]["calls"][0] == {"range": document.paragraphs()[0].id, "find": "Note :",
                                           "replace": "Note:", "expect": "one"}


def test_check_is_scoped_to_a_range(box):
    from test_scoped_check import scoped_tools

    document = french("Un : deux", "Trois : quatre", spaced=False)
    second = document.paragraphs()[1].id
    with Toolbox(scoped_tools(), formats=[docx_format()], groups=GROUPS, workers=0) as toolbox:
        session = _session(toolbox, document)
        report = _ok(toolbox, session, "check", doc="d1", range=[second]).data["layout"]
        assert [e["para"] for e in report["punctuation_spacing"]] == [second]


def test_a_paragraph_lists_ten_marks_and_counts_the_rest():
    document = french(" ".join(f"mot{k} :" for k in range(14)), spaced=False)
    entry = problems(document)[0]
    assert len(entry["marks"]) == 10 and entry["marks_left_out"] == 4
