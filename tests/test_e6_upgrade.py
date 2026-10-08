"""E6: ``upgrade_to_modern()``, held to what Word's Convert wrote (``tools/e6_probe.py``:
``convert-*`` against ``resave-*`` in ``tests/observations/e6-word.json``) on documents in modes
11, 12 and 14 -- with every legacy compatibility option set, and on three corpus fixtures --
and to the gates; and the properties' edits.
"""

from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import pytest

from docx_agent import Document
from docx_agent.edit.authoring import COMPAT_OPTIONS, KEPT
from docx_agent.layout import Reflow
from docx_agent.validate import check

from conftest import fixture_id, fixture_paths
from test_roundtrip import entries

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "tools"))
import e6_probe  # noqa: E402

OBSERVATIONS = json.loads((HERE / "observations" / "e6-word.json").read_text(encoding="utf-8"))
FIXTURES = HERE / "fixtures"

#: probe -> the document Word converted.
CONVERTED = {
    "convert-11": lambda: e6_probe.legacy_document(),
    "convert-all-11": lambda: e6_probe.legacy_document(11, e6_probe.EVERY_COMPAT),
    "convert-all-12": lambda: e6_probe.legacy_document(12, e6_probe.EVERY_COMPAT),
    "convert-all-14": lambda: e6_probe.legacy_document(14, e6_probe.EVERY_COMPAT),
    "convert-14-fe": lambda: e6_probe.legacy_document(14, e6_probe.MODE14_COMPAT),
    "convert-12": lambda: (FIXTURES / "docx2svg" / "layout-sweep.docx").read_bytes(),
    "convert-14": lambda: (FIXTURES / "generated" / "mode14.docx").read_bytes(),
    "convert-14-sample": lambda: (FIXTURES / "samplelib" / "sample-simple.docx").read_bytes(),
}


def _compat(settings: str) -> list[str]:
    block = re.search(r"<w:compat>(.*?)</w:compat>", settings)
    return re.findall(r"<w:(\w+)/>|<w:compatSetting w:name=\"(\w+)\"[^>]*w:val=\"(\d+)\"", block.group(1)) \
        if block else []


def _ours(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return re.sub(r' xmlns:\w+="[^"]*"', "", archive.read("word/settings.xml").decode("utf-8"))


def test_word_keeps_ten_options_and_writes_five_settings():
    """What the probe found: of the 65 legacy options Convert kept ten, in every mode; a plain
    save kept all but ``useFELayout`` (a save's doing, not Convert's)."""
    for probe in ("convert-all-11", "convert-all-12", "convert-all-14"):
        kept = {flag for flag, _, _ in _compat(OBSERVATIONS[probe]["settings"]) if flag}
        assert kept == KEPT - {"useFELayout"}, probe
    for probe in ("resave-all-11", "resave-all-14"):
        kept = {flag for flag, _, _ in _compat(OBSERVATIONS[probe]["settings"]) if flag}
        assert kept == set(COMPAT_OPTIONS) - {"useFELayout"}, probe


@pytest.mark.parametrize("probe", list(CONVERTED))
def test_convert_as_word_converts(probe):
    document = Document.open(CONVERTED[probe]())
    result = document.upgrade_to_modern()
    assert result.changed and document.compatibility_mode == 15
    ours = _compat(_ours(document.to_bytes()))
    word = _compat(OBSERVATIONS[probe]["settings"])
    # useFELayout: Word drops it on any save of a document without East Asian languages.
    assert [c for c in ours if c[0] != "useFELayout"] == [c for c in word if c[0] != "useFELayout"]
    assert isinstance(result.reflow, Reflow)
    removed = {w.split()[-1] for w in result.warnings}
    assert removed <= set(COMPAT_OPTIONS) - KEPT


def test_convert_changes_only_the_settings():
    data = (FIXTURES / "samplelib" / "sample-simple.docx").read_bytes()
    document = Document.open(data)
    document.upgrade_to_modern()
    with zipfile.ZipFile(io.BytesIO(data)) as before, zipfile.ZipFile(io.BytesIO(document.to_bytes())) as after:
        changed = {n for n in before.namelist() if before.read(n) != after.read(n)}
    assert changed == {"word/settings.xml"}


def test_a_modern_document_is_left_alone():
    document = Document.new()
    before = document.to_bytes()
    result = document.upgrade_to_modern()
    assert not result.changed and document.to_bytes() == before


def test_no_other_edit_changes_the_mode():
    document = Document.open(FIXTURES / "generated" / "mode14.docx")
    document.insert_markdown("# Heading\n\nText.\n")
    document.insert_table(2, 2, after=document.paragraphs()[0].id)
    assert document.compatibility_mode == 14


def test_a_document_without_settings_gets_them():
    document = Document.open(FIXTURES / "docx2svg" / "layout-sweep.docx")
    assert document.package.settings_part() is None or document.compatibility_mode == 12
    document.upgrade_to_modern()
    assert document.compatibility_mode == 15 and check(document.package) == []


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_every_fixture_upgrades_and_reports_its_reflow(path, tmp_path):
    data = path.read_bytes()
    document = Document.open(data)
    before = set(check(document.package))
    result = document.upgrade_to_modern()
    if not result.changed:
        assert document.compatibility_mode == 15
        return
    assert set(check(document.package)) - before == set()
    reflow = result.reflow
    assert reflow.page_count[0] >= 1 and reflow.page_count[1] >= 1
    assert all(1 <= page <= max(reflow.page_count) for page in reflow.changed)
    upgraded = document.to_bytes()
    document.save(tmp_path / "upgraded.docx")
    assert Document.open(tmp_path / "upgraded.docx").compatibility_mode == 15
    assert document.undo() and entries(document.to_bytes()) == entries(data)
    assert document.redo() and entries(document.to_bytes()) == entries(upgraded)


# -- properties --------------------------------------------------------------------------------


def test_set_properties_is_one_undo_step(tmp_path):
    document = Document.open(FIXTURES / "samplelib" / "sample-simple.docx")
    original = document.to_bytes()
    when = datetime(2026, 10, 4, 12, 0, tzinfo=timezone.utc)
    result = document.set_properties(title="A title", author="Claude", language="en-GB", modified=when,
                                     keywords="one; two", category="Report", status="Draft")
    assert result.changed
    properties = document.properties
    assert properties["title"] == "A title" and properties["author"] == "Claude"
    assert properties["language"] == "en-GB" and properties["modified"] == when
    assert properties["category"] == "Report" and properties["status"] == "Draft"
    document.save(tmp_path / "props.docx")
    assert Document.open(tmp_path / "props.docx").properties["title"] == "A title"
    assert not document.set_properties(title="A title").changed
    assert document.set_properties(title=None).changed and "title" not in document.properties
    while document.undo():
        pass
    assert entries(document.to_bytes()) == entries(original)


def test_a_document_without_core_properties_gets_them():
    document = Document.open(FIXTURES / "generated" / "mode14.docx")
    assert "title" not in document.properties
    document.set_properties(title="Made")
    saved = Document.open(document.to_bytes())
    assert saved.properties["title"] == "Made"
    assert check(saved.package) == []


@pytest.mark.parametrize("values", [{"colour": "red"}, {"created": "yesterday"}, {"revision": "two"}])
def test_bad_properties_are_refused(values):
    with pytest.raises((TypeError, ValueError)):
        Document.new().set_properties(**values)
