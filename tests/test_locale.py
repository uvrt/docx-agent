"""The regional settings a document is drawn and edited in: ``Document.locale``, and through
the tool layer ``Toolbox(locale=...)`` -- to every layout and render (in the worker process
too), to new charts, to the language of new chart text where the document states none, and
to a new document's language and separators.  en-US unless given, never the machine's.

The one accepted change of output: a new blank document's ``w:decimalSymbol`` and
``w:listSeparator`` are the locale's, so without one they are en-US's ``.`` and ``,`` where
they were the measuring Word's ``,`` and ``;``.  A locale with a decimal comma writes what
was written before, part for part (the digests below are of main before the change).
"""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import re
import warnings
import zipfile

import pytest
from ooxml_common.locales import EN_US, Locale
from ooxml_edit.tools import Toolbox

import docx2svg
from docx_agent import Document
from docx_agent.edit import blank
from docx_agent.tools import FORMAT, GROUPS, TOOLS

DUTCH = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy", language="nl-NL",
               list_separator=";")
#: The separators of the Word the new document was measured on, with the language a new
#: document had by default: what ``Document.new()`` wrote before the locale.
COMMA = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy", language="en-US",
               list_separator=";")
#: Every document here is made at this time, so none depends on the day the test runs.
CREATED = dt.datetime(2026, 10, 7, 12, 0, tzinfo=dt.timezone.utc)
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"

#: Digests of the parts (:func:`_parts_digest`) of ``Document.new(created=CREATED, ...)`` on
#: main before the locale: separators ``,`` and ``;`` (``.`` and ``,`` for Letter).
BEFORE = {
    (): "9049b9401ab2c9c19f027c20bd6bd6a8a05dff6cd62b4f076b1fbbfd4cc234a7",
    (("page", "Letter"),): "29ad6fd8c49d83d9c3c8d7019a354a31db0dcc10ed5cfef2577c4238a9c14293",
    (("language", "nl-NL"),): "915bd13a4991e5eb5eaf4b41d57d919e8eaa86cf7d9211b6db1027e6d35a29a0",
}


def _parts_digest(data: bytes) -> str:
    """The package's parts, names and bytes, hashed: independent of the zip's compression."""
    digest = hashlib.sha256()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in sorted(archive.namelist()):
            digest.update(name.encode() + b"\0" + archive.read(name) + b"\0")
    return digest.hexdigest()


def _settings(document: Document) -> tuple[str, str]:
    root = document.package.tree("word/settings.xml")
    return root.find(W + "decimalSymbol").get(W + "val"), root.find(W + "listSeparator").get(W + "val")


# -- a new document ------------------------------------------------------------------------------


def test_without_a_locale_a_new_document_is_en_us():
    document = Document.new(created=CREATED)
    assert document.locale is None and _settings(document) == (".", ",")
    assert Document.new(locale=EN_US, created=CREATED).to_bytes() == document.to_bytes()
    letter = Document.new(page="Letter", created=CREATED)
    assert _settings(letter) == (".", ",")
    # Letter wrote en-US's separators before: unchanged.
    assert _parts_digest(letter.to_bytes()) == BEFORE[(("page", "Letter"),)]


@pytest.mark.parametrize("options", [(), (("language", "nl-NL"),)])
def test_a_comma_locale_writes_what_was_written_before(options):
    document = Document.new(created=CREATED, locale=COMMA, **dict(options))
    assert _settings(document) == (",", ";")
    assert _parts_digest(document.to_bytes()) == BEFORE[options]


def test_a_new_document_takes_the_locales_language_and_separators():
    document = Document.new(locale=DUTCH, created=CREATED)
    assert document.locale is DUTCH and _settings(document) == (",", ";")
    styles = document.package.tree("word/styles.xml")
    assert styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}lang").get(W + "val") == "nl-NL"
    assert document.package.tree("word/settings.xml").find(W + "themeFontLang").get(W + "val") == "nl-NL"
    german = Document.new(locale=DUTCH, language="de-DE", created=CREATED)
    styles = german.package.tree("word/styles.xml")
    assert styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}lang").get(W + "val") == "de-DE"


def test_from_a_template_the_locale_is_kept_and_the_settings_are_the_templates():
    template = Document.new(created=CREATED).to_bytes()
    document = Document.new(template=template, locale=DUTCH, created=CREATED)
    assert document.locale is DUTCH and _settings(document) == (".", ",")


def test_the_deprecated_blank_locale_writes_what_it_wrote():
    with pytest.warns(DeprecationWarning):
        legacy = blank.Locale(margin=1000)
    with pytest.warns(DeprecationWarning):
        document = Document.new(locale=legacy, created=CREATED)
    assert document.locale is None and _settings(document) == (",", ";")
    margins = document.package.tree("word/document.xml").find(f"{W}body/{W}sectPr/{W}pgMar")
    assert margins.get(W + "left") == "1000"
    with pytest.warns(DeprecationWarning):
        metric = Document.new(locale=blank.METRIC, created=CREATED)
    assert _parts_digest(metric.to_bytes()) == BEFORE[()]
    with pytest.warns(DeprecationWarning):
        us = Document.new(locale=blank.US, page="Letter", created=CREATED)
    assert _parts_digest(us.to_bytes()) == BEFORE[(("page", "Letter"),)]


def test_page_defaults_and_locale_are_separate():
    with warnings.catch_warnings():
        warnings.simplefilter("error", DeprecationWarning)
        document = Document.new(page_defaults=blank.US, locale=DUTCH, created=CREATED)
    root = document.package.tree("word/settings.xml")
    assert root.find(W + "defaultTabStop").get(W + "val") == "720" and _settings(document) == (",", ";")


# -- charts and renders --------------------------------------------------------------------------


def _chart_document(locale=None) -> Document:
    document = Document.new(locale=locale, created=CREATED)
    document.insert_chart(document.paragraphs()[0].id, "column", ["Q1", "Q2"],
                          [{"name": "North", "values": [0.4, 1.3]}])
    return document


def _numbers(svg: str) -> set[str]:
    return set(re.findall(r">([0-9][0-9.,]*)<", svg))


def _chart_lang(document: Document) -> str:
    (part,) = [p for p in document.package.part_names if p.startswith("word/charts/chart")]
    return document.package.tree(part).find(f"{C}lang").get("val")


def test_a_new_chart_takes_the_locales_language_where_the_document_states_none():
    document = Document.new(created=CREATED)
    styles = document.package.tree("word/styles.xml")
    node = styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}lang")
    node.getparent().remove(node)
    document.package.mark_dirty("word/styles.xml")
    document.locale = DUTCH
    document.insert_chart(document.paragraphs()[0].id, "pie", ["a", "b"], [{"name": "s", "values": [1, 2]}])
    assert _chart_lang(document) == "nl-NL"
    # A document that states its language keeps it.
    stated = Document.new(language="de-DE", created=CREATED)
    stated.locale = DUTCH
    stated.insert_chart(stated.paragraphs()[0].id, "pie", ["a", "b"], [{"name": "s", "values": [1, 2]}])
    assert _chart_lang(stated) == "de-DE"


def test_a_render_takes_the_locale():
    english = _numbers(_chart_document().render_svg()[0])
    dutch = _numbers(_chart_document(DUTCH).render_svg()[0])
    assert "0.2" in english and "0,2" in dutch and "0.2" not in dutch
    # an option given to the call wins
    assert "0.2" in _numbers(_chart_document(DUTCH).render_svg(locale=EN_US)[0])


def test_the_png_takes_the_locale():
    if "resvg" not in docx2svg.available_backends():
        pytest.skip("needs resvg-py")
    assert _chart_document(DUTCH).render_png(width=320) != _chart_document(DUTCH).render_png(width=320,
                                                                                             locale=EN_US)


# -- the tool layer ------------------------------------------------------------------------------


def _box(**options) -> Toolbox:
    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, **options)


def _render(box: Toolbox, data: bytes, **session_options) -> bytes:
    session = box.session(**session_options)
    session.open(data, "chart.docx")
    result = box.dispatch(session, "render", {"doc": "d1", "pages": [1], "width": 320})
    assert result.ok, result.error
    return result.images[0].data


def _direct(data: bytes, locale) -> bytes:
    document = Document.open(data)
    document.locale = locale
    return document.render_png(pages=[1], width=320)[0]


@pytest.fixture
def chart_bytes() -> bytes:
    if "resvg" not in docx2svg.available_backends():
        pytest.skip("needs resvg-py")
    return _chart_document().to_bytes()


def test_the_toolbox_locale_reaches_the_document_and_the_worker(chart_bytes):
    english, dutch = _direct(chart_bytes, None), _direct(chart_bytes, DUTCH)
    assert english != dutch
    with _box(workers=1) as plain:
        assert _render(plain, chart_bytes) == english
    with _box(workers=1, locale=DUTCH) as configured:
        session = configured.session()
        session.open(chart_bytes, "chart.docx")
        assert session.entry("d1").document.locale is DUTCH
        assert _render(configured, chart_bytes) == dutch               # in the worker, in Dutch
        assert _render(configured, chart_bytes, locale=EN_US) == english
    with _box(workers=0, locale=DUTCH) as in_process:
        assert _render(in_process, chart_bytes) == dutch


def test_the_render_cache_is_keyed_by_the_locale(chart_bytes):
    with _box(workers=0) as box:
        session = box.session()
        session.open(chart_bytes, "chart.docx")
        first = box.dispatch(session, "render", {"doc": "d1", "pages": [1], "width": 320})
        session.entry("d1").document.locale = DUTCH
        second = box.dispatch(session, "render", {"doc": "d1", "pages": [1], "width": 320})
    assert first.images[0].data != second.images[0].data


def test_a_subprocess_worker_gets_the_locale_from_its_arguments(chart_bytes):
    """What a daemonic process renders with: the locale crosses as an argument."""
    from ooxml_edit.tools import SubprocessPool

    with _box(runner=SubprocessPool(1), locale=DUTCH) as box:
        assert _render(box, chart_bytes) == _direct(chart_bytes, DUTCH)


def test_the_layout_in_the_worker_takes_the_locale():
    """``check``'s layout runs in the worker pool (``Engine``) with the document's options,
    the locale among them."""
    with _box(workers=1, locale=DUTCH) as box:
        session = box.session()
        session.open(_chart_document().to_bytes(), "chart.docx")
        result = box.dispatch(session, "check", {"doc": "d1", "include": ["reflow"]})
        assert result.ok, result.error
        document = session.entry("d1").document
        assert any(repr(DUTCH) in key for key in document._layouts)


def test_new_document_takes_the_session_locale():
    with _box(locale=DUTCH) as box:
        session = box.session()
        made = box.dispatch(session, "new_document", {"kind": "docx"})
        assert made.ok, made.error
        document = session.entry("d1").document
        assert document.locale is DUTCH and _settings(document) == (",", ";")
    with _box() as box:
        session = box.session()
        assert box.dispatch(session, "new_document", {"kind": "docx"}).ok
        assert _settings(session.entry("d1").document) == (".", ",")


def test_the_definitions_and_prompt_do_not_change():
    with _box() as plain, _box(locale=DUTCH) as configured:
        for provider in ("anthropic", "openai-responses"):
            assert json.dumps(plain.definitions(provider)) == json.dumps(configured.definitions(provider))
        assert plain.system_prompt() == configured.system_prompt()
