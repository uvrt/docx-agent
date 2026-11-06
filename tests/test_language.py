"""An existing document's language (round 35, R24: :meth:`Document.set_language`,
:mod:`docx_agent.edit.language`; measured in ``tests/observations/language-word.json``).

Setting it writes what a new document in that language has -- ``w:docDefaults``' ``w:lang``
and ``w:themeFontLang`` -- and the styles that stated the old default; text written after
it gets the language's spacing; ``check`` no longer offers to delete French spaces in text
tagged another language.  The documents are made up here.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

import pytest
from lxml import etree

from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.edit.errors import EditError
from docx_agent.edit.language import WORD_TAGS, LanguageError, word_tag
from docx_agent.edit.typography import NBSP, problems
from docx_agent.tools import GROUPS, TOOLS, docx_format

ROOT = Path(__file__).resolve().parent.parent
OBSERVED = json.loads((ROOT / "tests" / "observations" / "language-word.json").read_text(encoding="utf-8"))
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W


def template(*texts: str, language: str = "en-US", normal: str | None = None,
             heading: str | None = "en-GB") -> Document:
    """A made-up report template: its default language, and styles that state their own."""
    document = Document.new(language=language)
    styles = document.package.tree(document.package.styles_part())
    for style in styles.iter(_W + "style"):
        if style.get(_W + "styleId") == "Normal" and normal:
            etree.SubElement(etree.SubElement(style, _W + "rPr"), _W + "lang").set(_W + "val", normal)
    if heading:
        document.styles.add("Chapter title", "paragraph", bold=True)
        chapter = next(s for s in styles.iter(_W + "style") if s.get(_W + "styleId") == document.styles.get(
            "Chapter title").id)
        rpr = chapter.find(_W + "rPr")
        etree.SubElement(rpr, _W + "lang").set(_W + "val", heading)
    first = document.paragraphs()[0]
    if texts:
        first.set_text(texts[0])
        last = first.id
        for text in texts[1:]:
            last = document.insert_paragraph(text, after=last).id
    document.package.mark_dirty(document.package.styles_part())
    return Document.open(document.to_bytes())


def defaults(document: Document) -> tuple[dict, dict]:
    found = document.languages()
    return found["default"], found["theme"]


# -- the measurements --------------------------------------------------------------------------


def test_the_tags_are_the_ones_word_writes():
    assert {slot: sorted(tags, key=str.lower) for slot, tags in WORD_TAGS.items()} == OBSERVED["tags"]
    assert {"fr-BE", "nl-BE", "fr-FR", "en-US"} <= WORD_TAGS["val"]
    assert "ja-JP" in WORD_TAGS["eastAsia"] and "ar-SA" in WORD_TAGS["bidi"]


def test_what_word_writes_when_it_sets_a_language():
    actions = OBSERVED["actions"]
    before = actions["nothing"]
    # None of Word's language commands rewrites the defaults, the theme or another style.
    for action in ("Normal style", "all text", "normal template"):
        assert actions[action]["docDefaults"] == before["docDefaults"]
        assert actions[action]["themeFontLang"] == before["themeFontLang"]
        for name, lang in before["styles"].items():
            assert actions[action]["styles"][name] == lang
    assert actions["Normal style"]["styles"]["Normal"] == {"val": "fr-BE"}
    assert {lang["val"] for lang in actions["all text"]["runs"]} == {"fr-BE"}
    assert actions["normal template"]["Normal.dotm changed"] and not actions["normal template"]["runs"]
    assert actions["normal template"]["proofed as"] == before["proofed as"]       # the open document as it was


def test_tags_are_read_as_word_reads_them():
    reads = OBSERVED["reads"]
    for given in ("fr-BE", "FR-be", "fr_BE", "fr-Latn-BE"):
        assert reads[given].startswith("belgian french")
        assert word_tag(given)[:2] == ("fr-BE", "val")
    assert reads["xx-YY"] == reads["zz"] == "belgian dutch"     # the measuring Word's own language
    assert word_tag("fr-BE")[2] is None and "fr-BE" in word_tag("fr_be")[2]
    assert word_tag("ja-JP")[1] == "eastAsia" and word_tag("he-IL")[1] == "bidi"


@pytest.mark.parametrize("given, wanted", [("fr", "fr-BE"), ("fr-XX", "fr-FR"), ("xx-YY", "en-US"),
                                           ("", "fr-FR"), ("fr-be-x-test", "fr-BE")])
def test_a_tag_word_does_not_know_is_refused_with_the_ones_that_would_do(given, wanted):
    with pytest.raises(LanguageError) as refused:
        word_tag(given)
    assert wanted in refused.value.options or wanted in str(refused.value)
    assert isinstance(refused.value, EditError)


def test_a_bare_language_lists_its_own_country_first():
    with pytest.raises(LanguageError) as refused:
        word_tag("nl")
    assert refused.value.options[:2] == ["nl-NL", "nl-BE"]


# -- set_language ------------------------------------------------------------------------------


def test_it_writes_what_a_new_document_in_the_language_has():
    document = template()
    result = document.set_language("fr-BE")
    assert result.changed
    assert defaults(document) == defaults(Document.new(language="fr-BE"))
    assert document.validate() == []


def test_styles_stating_the_old_default_follow_and_others_keep_theirs():
    document = template(normal="en-US", heading="en-GB")
    result = document.set_language("fr-BE")
    styles = document.languages()["styles"]
    assert styles["Normal"] == {"val": "fr-BE"} and styles["Chapter title"] == {"val": "en-GB"}
    assert any("Chapter title (en-GB)" in w for w in result.warnings)


@pytest.mark.parametrize("styles, normal, chapter", [("all", "fr-BE", "fr-BE"), ("none", "en-US", "en-GB"),
                                                     (["Chapter title"], "en-US", "fr-BE")])
def test_the_styles_option(styles, normal, chapter):
    document = template(normal="en-US", heading="en-GB")
    document.set_language("fr-BE", styles=styles)
    found = document.languages()
    assert found["default"]["val"] == "fr-BE"
    assert found["styles"]["Normal"]["val"] == normal and found["styles"]["Chapter title"]["val"] == chapter


def test_a_style_alone():
    document = template(heading="en-GB")
    document.set_language("fr-FR", styles=["Chapter title"], default=False)
    assert document.languages()["default"]["val"] == "en-US"
    assert document.languages()["styles"]["Chapter title"] == {"val": "fr-FR"}
    with pytest.raises(LanguageError):
        document.set_language("fr-FR", default=False)


def test_runs_keep_their_own_tags_unless_asked():
    document = template("Bonjour.", "Guten Tag.", "Encore.")
    paragraphs = document.paragraphs()
    for paragraph, tag in zip(paragraphs, ("en-US", "de-DE", "en-US")):
        run = paragraph._element.find(_W + "r")
        rpr = run.find(_W + "rPr")
        if rpr is None:
            rpr = etree.SubElement(run, _W + "rPr")
            run.insert(0, rpr)
        etree.SubElement(rpr, _W + "lang").set(_W + "val", tag)
    document.package.mark_dirty(document.package.document_part())
    document = Document.open(document.to_bytes())
    result = document.set_language("fr-BE")
    assert document.languages()["runs"] == {"en-US": 2, "de-DE": 1}
    assert any("runs=True also sets the 2 in en-US" in w for w in result.warnings)
    document.set_language("fr-BE", runs=True)
    assert document.languages()["runs"] == {"en-US": 2, "de-DE": 1}         # the old default is fr-BE now
    other = template("Bonjour.")
    run = other.paragraphs()[0]._element.find(_W + "r")
    rpr = etree.Element(_W + "rPr")
    run.insert(0, rpr)
    etree.SubElement(rpr, _W + "lang").set(_W + "val", "en-US")
    other.package.mark_dirty(other.package.document_part())
    other = Document.open(other.to_bytes())
    other.set_language("fr-BE", runs=True)
    assert other.languages()["runs"] == {"fr-BE": 1}


def test_east_asian_and_complex_script_languages_go_where_word_writes_them():
    document = template()
    document.set_language("ja-JP")
    assert document.languages()["default"] == {"val": "en-US", "eastAsia": "ja-JP", "bidi": "ar-SA"}
    document.set_language("he-IL")
    assert document.languages()["default"]["bidi"] == "he-IL"
    assert document.languages()["theme"] == {"val": "en-US", "eastAsia": "ja-JP", "bidi": "he-IL"}


def test_one_undo_step_and_a_no_op_is_not_an_error():
    document = template(normal="en-US")
    before = document.to_bytes()
    document.set_language("nl-BE")
    document.undo()
    assert document.to_bytes() == before
    document.set_language("nl-BE")
    again = document.set_language("nl_BE")
    assert not again.changed and any("nl-BE already" in w for w in again.warnings)


def test_a_document_without_a_theme_language_gets_one():
    document = template()
    settings = document.package.tree(document.package.settings_part())
    settings.remove(settings.find(_W + "themeFontLang"))
    document.package.mark_dirty(document.package.settings_part())
    document = Document.open(document.to_bytes())
    assert document.languages()["theme"] == {}
    document.set_language("fr-BE")
    assert document.languages()["theme"] == {"val": "fr-BE"} and document.validate() == []


# -- writing and check -------------------------------------------------------------------------


CLOCK = lambda: datetime.datetime(2026, 11, 6, 12, 0)  # noqa: E731


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[docx_format()], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _session(box, document: Document):
    session = box.session(clock=CLOCK)
    session.open(document.to_bytes(), "rapport.docx")
    return session


def _ok(box, session, tool, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, json.dumps(result.to_json(), default=str)[:800]
    return result


def _spacing(box, session) -> list[dict]:
    return _ok(box, session, "check", doc="d1", include=["layout"]).data["layout"]["punctuation_spacing"]


def test_an_english_template_set_to_belgian_french_is_written_with_french_spacing(box):
    session = _session(box, template("Titre"))
    result = _ok(box, session, "word_styles", doc="d1", action="set_language", language="fr-BE")
    assert result.data["default"]["val"] == "fr-BE" and "fr-BE" in result.summary
    first = session.entry("d1").document.paragraphs()[0].id
    _ok(box, session, "word_set_text", doc="d1", target=first, text="Résumé : les points clés!")
    _ok(box, session, "word_insert_markdown", doc="d1", at="end",
        markdown="Il a dit « oui » : pourquoi?\n\n| Mesure | Valeur |\n|---|---|\n| Délai : | court! |\n")
    document = session.entry("d1").document
    texts = [p.text for p in document.paragraphs()]
    assert texts[0] == f"Résumé{NBSP}: les points clés{NBSP}!"
    assert f"Il a dit «{NBSP}oui{NBSP}»{NBSP}: pourquoi{NBSP}?" in texts
    assert f"Délai{NBSP}:" in texts and f"court{NBSP}!" in texts
    assert _spacing(box, session) == []
    assert document.validate() == []


def test_belgian_dutch_gets_no_french_rules(box):
    session = _session(box, template("Titel"))
    _ok(box, session, "word_styles", doc="d1", action="set_language", language="nl-BE")
    first = session.entry("d1").document.paragraphs()[0].id
    _ok(box, session, "word_set_text", doc="d1", target=first, text="Besluit: goed! Waarom?")
    assert session.entry("d1").document.paragraphs()[0].text == "Besluit: goed! Waarom?"
    assert _spacing(box, session) == []


def test_french_spacing_in_english_text_is_a_language_mismatch_never_a_deletion(box):
    document = template("Bonjour : le monde ! Pourquoi ?", "Fine: yes.")
    found = problems(document)
    assert len(found) == 1
    entry = found[0]
    assert entry["kind"] == "language_mismatch" and entry["language"] == "en-US" and entry["source"] == "default"
    assert [m["kind"] for m in entry["marks"]] == ["space_before"] * 3
    assert entry["fix"] == {"tool": "word_styles", "calls": [{"action": "set_language", "language": "fr-FR"}]}
    assert "replace_text" not in json.dumps(found)
    # Following the fix: the text is French now, its ordinary spaces the ones to make no-break.
    session = _session(box, document)
    for call in entry["fix"]["calls"]:
        _ok(box, session, "word_styles", doc="d1", **call)
    after = _spacing(box, session)
    assert after and all(e.get("kind") != "language_mismatch" and e["fix"]["tool"] == "replace_text"
                         for e in after)
    for item in after:
        for call in item["fix"]["calls"]:
            assert NBSP in call["replace"] and call["replace"].replace(NBSP, "") == call["find"].replace(" ", "")
            _ok(box, session, "replace_text", doc="d1", **call)
    assert _spacing(box, session) == []


def test_the_fix_names_the_style_the_tag_comes_from_and_a_french_tag_the_document_has(box):
    document = template("Chapitre : un", language="fr-BE", heading="en-US")
    paragraph = document.paragraphs()[0]
    document.set_paragraph_style(paragraph.id, "Chapter title")
    entry = problems(document)[0]
    assert entry["source"] == "style:Chapter title"
    assert entry["fix"]["calls"] == [{"action": "set_language", "language": "fr-BE", "name": "Chapter title"}]
    session = _session(box, document)
    _ok(box, session, "word_styles", doc="d1", **entry["fix"]["calls"][0])
    assert _spacing(box, session) == []        # French now: the no-break space it was written with


def test_a_runs_own_tag_gets_no_fix():
    document = template("Bonjour : le monde")
    run = document.paragraphs()[0]._element.find(_W + "r")
    rpr = etree.Element(_W + "rPr")
    run.insert(0, rpr)
    etree.SubElement(rpr, _W + "lang").set(_W + "val", "en-GB")
    document.package.mark_dirty(document.package.document_part())
    entry = problems(Document.open(document.to_bytes()))[0]
    assert entry["source"] == "run" and "fix" not in entry and "not a spacing fault" in entry["note"]


def test_french_text_keeps_its_replace_text_fixes():
    document = Document.new(language="fr-FR")
    document.typography = False
    document.paragraphs()[0].set_text("Bonjour : le monde!")
    entry = problems(document)[0]
    assert "kind" not in entry and entry["fix"]["tool"] == "replace_text"


def test_the_tool_refuses_with_the_way_out(box):
    session = _session(box, template())
    result = box.dispatch(session, "word_styles", {"doc": "d1", "action": "set_language", "language": "fr"})
    assert not result.ok
    error = result.to_json()["error"]
    assert error["field"] == "language" and "fr-BE" in error["valid_options"]
    missing = box.dispatch(session, "word_styles", {"doc": "d1", "action": "set_language", "language": "fr-BE",
                                                    "name": "No such style"})
    assert not missing.ok and missing.to_json()["error"]["code"] == "not_found"
    no_op = _ok(box, session, "word_styles", doc="d1", action="set_language", language="en_us")
    assert no_op.summary.startswith("No change")
    wrong = box.dispatch(session, "word_styles", {"doc": "d1", "action": "remove", "name": "Normal",
                                                  "language": "fr-BE"})
    assert not wrong.ok and wrong.to_json()["error"]["field"] == "language"


def test_undo_through_the_session(box):
    session = _session(box, template())
    _ok(box, session, "word_styles", doc="d1", action="set_language", language="fr-BE")
    _ok(box, session, "undo", doc="d1")
    assert session.entry("d1").document.languages()["default"]["val"] == "en-US"


# -- Word ----------------------------------------------------------------------------------------


@pytest.mark.oracle
def test_word_proofs_the_text_in_the_language_set():
    import oracle

    document = template("Plain text.", "Chapter text.", normal="en-US", heading="en-GB")
    document.set_paragraph_style(document.paragraphs()[1].id, "Chapter title")
    document.set_language("fr-BE")
    script = '''
on run argv
  set out to ""
  with timeout of 120 seconds
    tell application "Microsoft Word"
      open (POSIX file (item 1 of argv))
      set d to active document
      repeat with i from 1 to 2
        set out to out & ((language ID of text object of paragraph i of d) as string) & linefeed
      end repeat
      save as d file name (item 2 of argv) file format format document
      close d saving no
      quit saving no
    end tell
  end timeout
  return out
end run
'''
    try:
        with oracle.session() as word:
            outcome = word.run_script(script, document.to_bytes(), name="set-language", tag="proofed")
    except oracle.WordBusy as busy:
        pytest.skip(str(busy))
    assert outcome, outcome.detail
    if outcome.path is not None:
        outcome.path.unlink(missing_ok=True)
    assert outcome.detail.split() == ["belgian", "french", "english", "uk"]


# -- the template's own language (round 35, R25) -------------------------------------------------


ENGLISH = ("Normal", "Heading1", "Heading2", "Heading1Char", "Heading2Char", "Title", "TitleChar",
           "PageFooter", "PageFooterChar", "ReportDate")


def _lang(style: etree._Element, **attributes: str) -> None:
    rpr = style.find(_W + "rPr")
    if rpr is None:
        rpr = etree.SubElement(style, _W + "rPr")
    lang = rpr.find(_W + "lang")
    if lang is None:
        lang = etree.SubElement(rpr, _W + "lang")
    for slot, value in attributes.items():
        lang.set(_W + slot, value)


def english_template(*texts: str, default: str | None = "nl-BE") -> Document:
    """A made-up report template made in an English Word: the default ``default`` (or none),
    Normal, the headings, their character styles, footer and date styles in en-US; one
    paragraph style in German (with a child stating nothing) and one character style in
    Italian, on purpose."""
    document = Document.new(language="en-US")
    document.styles.add("Page footer", "paragraph", based_on="Normal")
    document.styles.add("Page footer Char", "character")
    document.styles.add("Report date", "paragraph", based_on="Page footer")
    document.styles.add("Quotation German", "paragraph", based_on="Normal", italic=True)
    document.styles.add("Quotation German child", "paragraph", based_on="Quotation German")
    document.styles.add("Term Italian", "character", italic=True)
    root = document.package.tree(document.package.styles_part())
    by_id = {s.get(_W + "styleId"): s for s in root.iter(_W + "style")}
    by_name = {s.find(_W + "name").get(_W + "val"): s for s in root.iter(_W + "style")}
    for identifier in ENGLISH:
        style = by_id.get(identifier)
        if style is None:
            style = by_name[{"PageFooter": "Page footer", "PageFooterChar": "Page footer Char",
                             "ReportDate": "Report date"}[identifier]]
        _lang(style, val="en-US")
    _lang(by_name["Quotation German"], val="de-DE")
    _lang(by_name["Term Italian"], val="it-IT")
    lang = root.find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr/{_W}lang")
    if default is None:
        del lang.attrib[_W + "val"]
    else:
        lang.set(_W + "val", default)
    first = document.paragraphs()[0]
    if texts:
        first.set_text(texts[0])
        last = first.id
        for text in texts[1:]:
            last = document.insert_paragraph(text, after=last).id
    document.package.mark_dirty(document.package.styles_part())
    return Document.open(document.to_bytes())


FRENCH = f"Résumé{NBSP}: les points clés{NBSP}! Pourquoi{NBSP}?"


@pytest.mark.parametrize("default", ["nl-BE", None, "en-US"])
def test_the_template_s_own_language_follows_and_other_languages_stay_named(default):
    document = english_template(FRENCH, default=default)
    before = document.languages()["default"]
    result = document.set_language("fr-BE")
    styles = document.languages()["styles"]
    english = [name for name, lang in styles.items() if name not in ("Quotation German", "Term Italian")]
    assert len(english) == len(ENGLISH) and all(styles[name] == {"val": "fr-BE"} for name in english), styles
    assert styles["Quotation German"] == {"val": "de-DE"} and styles["Term Italian"] == {"val": "it-IT"}
    assert "Quotation German child" not in styles         # states nothing: German through its parent
    kept = next(w for w in result.warnings if w.startswith("styles keeping"))
    assert "Quotation German (de-DE)" in kept and "Term Italian (it-IT)" in kept and "en-US" not in kept
    own = [w for w in result.warnings if w.startswith("the template's own language")]
    if default == "en-US":
        assert own == []                                     # the default is all there was
    else:
        assert own and "en-US (Normal's)" in own[0] and (default is None or "nl-BE (the default)" in own[0])
    after = document.languages()["default"]
    assert after["val"] == "fr-BE" and {k: v for k, v in after.items() if k != "val"} == \
        {k: v for k, v in before.items() if k != "val"}      # eastAsia and bidi as they were
    assert document.validate() == []


def test_french_in_the_english_template_is_one_call_away_and_then_clean(box):
    document = english_template(FRENCH, f"Il a dit «{NBSP}oui{NBSP}».")
    entry = problems(document)[0]
    assert entry["kind"] == "language_mismatch" and entry["source"] == "style:Normal"
    # Normal's language is the template's own: the fix is the document's, not Normal's alone.
    assert entry["fix"]["calls"] == [{"action": "set_language", "language": "fr-FR"}]
    session = _session(box, document)
    result = _ok(box, session, "word_styles", doc="d1", action="set_language", language="fr-BE")
    assert "Quotation German (de-DE)" in json.dumps(result.to_json())
    assert _spacing(box, session) == []
    first = session.entry("d1").document.paragraphs()[0].id
    _ok(box, session, "word_set_text", doc="d1", target=first, text="Conclusion : tout va bien!")
    _ok(box, session, "word_insert_markdown", doc="d1", at="end", markdown="# Chapitre : deux\n\nPourquoi?\n")
    texts = [p.text for p in session.entry("d1").document.paragraphs()]
    assert f"Conclusion{NBSP}: tout va bien{NBSP}!" in texts
    assert f"Chapitre{NBSP}: deux" in texts and f"Pourquoi{NBSP}?" in texts
    assert _spacing(box, session) == []


def test_a_style_in_another_language_keeps_its_own_fix():
    document = english_template(FRENCH)
    document.set_paragraph_style(document.paragraphs()[0].id, "Quotation German")
    entry = problems(document)[0]
    assert entry["source"] == "style:Quotation German"
    assert entry["fix"]["calls"] == [{"action": "set_language", "language": "fr-FR", "name": "Quotation German"}]


def test_runs_in_the_template_s_own_language_follow_when_asked():
    document = english_template("One.", "Two.", "Three.")
    for paragraph, tag in zip(document.paragraphs(), ("en-US", "nl-BE", "de-DE")):
        run = paragraph._element.find(_W + "r")
        rpr = etree.Element(_W + "rPr")
        run.insert(0, rpr)
        etree.SubElement(rpr, _W + "lang").set(_W + "val", tag)
    document.package.mark_dirty(document.package.document_part())
    document = Document.open(document.to_bytes())
    plain = document.set_language("fr-BE")
    assert document.languages()["runs"] == {"en-US": 1, "nl-BE": 1, "de-DE": 1}
    assert any("runs=True also sets the 2 in nl-BE or en-US" in w for w in plain.warnings)
    document.undo()
    document.set_language("fr-BE", runs=True)
    assert document.languages()["runs"] == {"fr-BE": 2, "de-DE": 1}


def _bare(*languages: str | None) -> Document:
    """No default language, none on Normal; styles stating ``languages`` (None: nothing)."""
    document = Document.new(language="en-US")
    for i, _ in enumerate(languages):
        document.styles.add(f"Body {i}", "paragraph", based_on="Normal")
    root = document.package.tree(document.package.styles_part())
    root.find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr/{_W}lang").attrib.pop(_W + "val")
    by_name = {s.find(_W + "name").get(_W + "val"): s for s in root.iter(_W + "style")}
    for i, language in enumerate(languages):
        if language:
            _lang(by_name[f"Body {i}"], val=language)
    document.package.mark_dirty(document.package.styles_part())
    return Document.open(document.to_bytes())


@pytest.mark.parametrize("languages, own", [
    (("en-US", "en-US", "de-DE"), ["en-US"]),         # more than half, at least two
    (("en-US", "de-DE"), []),                         # a tie
    (("en-US", "en-US", "de-DE", "it-IT"), []),       # two of four: not more than half
    (("de-DE", None), []),                            # a single style is an exception
    (("en_us", "EN-US", "de-DE"), ["en_us"]),         # any spelling, the first one's kept
])
def test_without_a_default_or_normal_the_language_most_styles_state(languages, own):
    from docx_agent.edit.language import own_languages

    document = _bare(*languages)
    root = document.package.tree(document.package.styles_part())
    assert [tag for tag, _ in own_languages(root)] == own
    result = document.set_language("fr-BE")
    styles = document.languages()["styles"]
    for i, language in enumerate(languages):
        if language:
            assert styles[f"Body {i}"]["val"] == ("fr-BE" if own and _same_tag(language, own[0]) else language)
    if not own:
        assert any("states no language of its own" in w for w in result.warnings)


def _same_tag(a: str, b: str) -> bool:
    return a.replace("_", "-").lower() == b.replace("_", "-").lower()


def test_the_rule_reads_only_the_attribute_the_language_is_written_in():
    from docx_agent.edit.language import own_languages

    document = english_template()
    root = document.package.tree(document.package.styles_part())
    normal = next(s for s in root.iter(_W + "style") if s.get(_W + "styleId") == "Normal")
    _lang(normal, eastAsia="zh-CN")
    document.package.mark_dirty(document.package.styles_part())
    document = Document.open(document.to_bytes())
    root = document.package.tree(document.package.styles_part())
    assert own_languages(root, "val") == [("nl-BE", "default"), ("en-US", "Normal")]
    assert [why for _, why in own_languages(root, "eastAsia")][-1] == "Normal"
    document.set_language("fr-BE")
    styles = document.languages()["styles"]
    assert styles["Normal"] == {"val": "fr-BE", "eastAsia": "zh-CN"}        # eastAsia not touched
    document.set_language("ja-JP")
    assert document.languages()["styles"]["Normal"] == {"val": "fr-BE", "eastAsia": "ja-JP"}
    assert document.languages()["styles"]["heading 1"] == {"val": "fr-BE"}   # no eastAsia of its own
