"""The language a text is written in, next to its tag (round 49, R94:
:mod:`docx_agent.edit.text_language`, ``check``'s layout fact ``text_language``).

The detector's measured claims hold on the made-up samples (``tests/fixtures/
language-samples.json``, ``tools/language_accuracy.py``); the fact lists what differs, folds
a whole document under the wrong default into one entry with the fix, groups consecutive
paragraphs, counts what agrees, and stays capped.  The documents are made up here.
"""

from __future__ import annotations

import datetime
import json
import sys
from pathlib import Path

import pytest
from lxml import etree

from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.edit import text_language
from docx_agent.edit.text_language import (LANGUAGES, MIN_CONFIDENCE, MIN_WORDS, PROFILES, detect,
                                           words_of)
from docx_agent.tools import GROUPS, TOOLS, docx_format

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = json.loads((ROOT / "tests" / "fixtures" / "language-samples.json").read_text(encoding="utf-8"))
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CLOCK = lambda: datetime.datetime(2026, 11, 10, 12, 0)  # noqa: E731

FR = ["Les copies de sauvegarde tournent chaque nuit et sont gardées pendant trente jours au moins.",
      "Un fichier supprimé par erreur peut être remis en place depuis le portail par chaque personne.",
      "Pour tout ce qui a plus d'un mois, il faut passer par l'équipe technique qui le fera vite."]
EN = ["The new office opens on the second floor next month, and the teams near the stairs move first.",
      "Please pack your desk before Friday and label each box with your name and the new room number.",
      "Coffee and tea are free on every floor, and the kitchen is cleaned each evening after six."]
NL = "De fietsroute begint bij het oude station en volgt het kanaal over een afstand van twaalf kilometer."


def made_up(*texts: str, language: str = "en-GB") -> Document:
    document = Document.new(language=language)
    first = document.paragraphs()[0]
    first.set_text(texts[0])
    last = first.id
    for text in texts[1:]:
        last = document.insert_paragraph(text, after=last).id
    return Document.open(document.to_bytes())


def tag_runs(document: Document, starts: str, language: str, *, no_proof: bool = False) -> Document:
    """Each run of the paragraphs starting with ``starts`` tagged ``language`` (``w:lang``)."""
    part = document.package.document_part()
    for paragraph in document.package.tree(part).iter(_W + "p"):
        text = "".join(t.text or "" for t in paragraph.iter(_W + "t"))
        if not text.startswith(starts):
            continue
        for run in paragraph.iter(_W + "r"):
            properties = run.find(_W + "rPr")
            if properties is None:
                properties = etree.Element(_W + "rPr")
                run.insert(0, properties)
            if no_proof:
                etree.SubElement(properties, _W + "noProof")
            etree.SubElement(properties, _W + "lang").set(_W + "val", language)
    document.package.mark_dirty(part)
    return Document.open(document.to_bytes())


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[docx_format()], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _ok(box, session, tool, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, json.dumps(result.to_json(), default=str)[:800]
    return result


def fact(box, document: Document, **arguments) -> tuple[dict, object]:
    session = box.session(clock=CLOCK)
    session.open(document.to_bytes(), "made-up.docx")
    found = _ok(box, session, "check", doc="d1", include=["layout"], **arguments).data["layout"]
    return found.get("text_language"), session


# -- the detector, as measured ------------------------------------------------------------------


def _windows(text: str, n: int) -> list[str]:
    words = text.split()
    return [" ".join(words[i:i + n]) for i in range(0, len(words) - n + 1, 3)]


def test_the_measured_accuracy_holds_at_the_minimum_length():
    """At MIN_WORDS words, cut anywhere in the made-up text of every set: no window is listed
    as another language (wrong and MIN_CONFIDENCE sure), 97% of those judged are right."""
    judged = right = wrong_and_sure = 0
    for name in ("tuning", "held_out", "fresh"):
        for language in LANGUAGES:
            for paragraph in SAMPLES[name][language]:
                for window in _windows(paragraph, MIN_WORDS):
                    found = detect(window)
                    if found.language is None:
                        continue
                    judged += 1
                    right += found.language == language
                    wrong_and_sure += found.language != language and found.confidence >= MIN_CONFIDENCE
    assert wrong_and_sure == 0
    assert judged > 1200 and right / judged >= 0.97


def test_every_whole_sample_paragraph_is_its_language():
    """126 paragraphs of 35-55 words: all right; all but 5 Portuguese ones (close to Spanish
    and Italian: 0.39-0.46) sure enough to be listed were they tagged wrong."""
    unsure = []
    for name in ("tuning", "held_out", "fresh"):
        for language in LANGUAGES:
            for paragraph in SAMPLES[name][language]:
                found = detect(paragraph)
                assert found.language == language, (paragraph, found)
                if found.confidence < MIN_CONFIDENCE:
                    unsure.append(language)
    assert unsure == ["pt"] * 5


def test_too_few_words_or_function_words_are_not_judged():
    assert detect("Le rapport annuel").language is None                       # 3 words
    assert detect("Jan Peeters, Marie Dubois, Hans Weber, Ana Ruiz, Luca Neri, Emma Smith").language is None
    assert detect("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor "
                  "incididunt ut labore et dolore magna aliqua").language is None
    assert words_of("l’équipe qu'il dell'anno") == ["l'", "équipe", "qu'", "il", "dell'", "anno"]


def test_the_tool_reproduces_the_measurement():
    sys.path.insert(0, str(ROOT / "tools"))
    try:
        import language_accuracy
    finally:
        sys.path.pop(0)
    result = language_accuracy.measure(SAMPLES["fresh"], MIN_WORDS)
    assert result["total"]["false_positives"] == 0 and result["total"]["judged"] > 150


def test_the_profiles_are_function_words_of_each_language():
    assert set(PROFILES) == set(LANGUAGES) == {"en", "fr", "nl", "de", "es", "it", "pt"}
    assert all(40 <= len(words) <= 120 for words in PROFILES.values())


# -- the fact -----------------------------------------------------------------------------------


def test_a_document_in_another_language_than_its_default_is_one_entry_with_the_fix(box):
    """Rule 2 (the way out) and rule 3: the fix made, the check has nothing left to say."""
    found, session = fact(box, made_up(*FR, "Résumé", *FR))
    document = found["document"]
    assert {k: document[k] for k in ("tag", "source", "detected", "share", "paragraphs", "mismatched")} == {
        "tag": "en-GB", "source": "default", "detected": "fr", "share": 1.0, "paragraphs": 6, "mismatched": 6}
    assert document["fix"] == {"tool": "word_styles", "calls": [{"action": "set_language", "language": "fr-FR"}]}
    assert "mismatched" not in found and found["too_short"] == 1             # folded, not listed
    _ok(box, session, "word_styles", doc="d1", **document["fix"]["calls"][0])
    after = _ok(box, session, "check", doc="d1", include=["layout"]).data["layout"]["text_language"]
    assert after["document"]["tag"] == "fr-FR" and after["document"]["detected"] == "fr"
    assert "fix" not in after["document"] and "mismatched" not in after


def test_a_passage_in_another_language_is_listed_and_consecutive_paragraphs_are_one_range(box):
    document = made_up(*EN, *FR[:2], "Annexe", FR[2], *EN, NL)
    found, _ = fact(box, document)
    ids = [p.id for p in document.paragraphs()]
    assert found["document"] == {"tag": "en-GB", "source": "default", "detected": "en",
                                 "share": found["document"]["share"], "words": found["document"]["words"],
                                 "paragraphs": 10}
    first, second = found["mismatched"]
    assert first == {"range": f"{ids[3]}..{ids[6]}", "paragraphs": 4, "tag": "en-GB", "source": "default",
                     "detected": "fr", "confidence": first["confidence"], "words": first["words"]}
    assert first["confidence"] >= MIN_CONFIDENCE
    assert second["para"] == ids[10] and second["detected"] == "nl"
    assert found["too_short"] == 1                                         # "Annexe"


def test_text_tagged_its_own_language_agrees_and_a_wrong_run_tag_says_run(box):
    document = tag_runs(made_up(*EN, FR[0], EN[0] + " Again."), FR[0][:12], "fr-FR")
    found, _ = fact(box, document)
    assert "mismatched" not in found                                       # its runs say fr-FR
    document = tag_runs(made_up(*EN, FR[0]), EN[1][:12], "de-DE")
    found, _ = fact(box, document)
    run, french = found["mismatched"]
    assert (run["tag"], run["source"], run["detected"]) == ("de-DE", "run", "en")
    assert (french["tag"], french["source"]) == ("en-GB", "default")


def test_text_word_does_not_proof_is_not_judged(box):
    document = tag_runs(made_up(*EN, FR[0]), FR[0][:12], "en-GB", no_proof=True)
    found, _ = fact(box, document)
    assert "mismatched" not in found and found["document"]["paragraphs"] == 3


def test_the_list_is_capped_and_says_what_was_left_out(box):
    """Rule 4: at most CAP entries, the rest counted."""
    texts = []
    for i in range(25):
        texts += [EN[i % 3], FR[i % 3]]
    found, _ = fact(box, made_up(*texts))
    assert len(found["mismatched"]) == 20 and found["mismatched_left_out"] == 5


def test_a_range_says_what_its_words_are_and_proposes_no_document_fix(box):
    document = made_up(*EN, *FR, *EN)
    ids = [p.id for p in document.paragraphs()]
    found, _ = fact(box, document, range=f"{ids[3]}..{ids[5]}")
    assert "document" not in found and found["checked"]["detected"] == "fr"
    assert "fix" not in found["checked"] and "mismatched" not in found["checked"]
    assert found["mismatched"] == [{"range": f"{ids[3]}..{ids[5]}", "paragraphs": 3, "tag": "en-GB",
                                    "source": "default", "detected": "fr",
                                    "confidence": found["mismatched"][0]["confidence"],
                                    "words": found["mismatched"][0]["words"]}]


def test_a_document_too_short_to_judge_has_no_fact(box):
    found, _ = fact(box, made_up("Titre", "Résumé court"))
    assert found is None


def test_the_fix_takes_a_tag_the_document_already_names():
    document = made_up(*FR)
    styles = document.package.tree(document.package.styles_part())
    style = next(s for s in styles.iter(_W + "style") if s.get(_W + "type") == "paragraph")
    etree.SubElement(etree.SubElement(style, _W + "rPr"), _W + "lang").set(_W + "val", "fr-BE")
    assert text_language.tag_for(document, "fr") == "fr-BE"
    assert text_language.tag_for(document, "pt") == "pt-PT"
