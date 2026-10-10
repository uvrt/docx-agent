"""The application's own font folders through the tool layer: ``Toolbox(font_dirs=...)``,
a session's own, or ``OOXML_FONT_DIRS`` -- for the layout (reflow, coverage, saving's
facts) and the render, in the worker process too.

Production: with Aptos only in an application's folder, docx2svg given
``font_dirs=[folder]`` laid a document out completely, and the ``render`` tool reported it
incomplete, Aptos missing.  The face here is an open one -- Cousine from the
``pptx2svg-fonts`` bundle, or Liberation Mono where CI installs it -- relabelled in a
temporary folder as a family nothing else answers to, so no system copy can stand in.
"""

from __future__ import annotations

import io
import json
import multiprocessing
import sys
import warnings
import zipfile
from pathlib import Path

import pytest

from ooxml_common.fonts import bundle_dir
from ooxml_common.fonts.office import FONT_DIRS_ENV
from ooxml_common.fonts.sfnt import relabel
from ooxml_edit.tools import Toolbox

import docx2svg
from docx2svg import fonts
from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS

FAMILY = "Fontdirs Probe Mono"
_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _open_face() -> Path | None:
    """An open monospaced face to relabel: the bundle's Cousine, else Liberation Mono."""
    bundle = bundle_dir()
    if bundle is not None and (bundle / "Cousine-Regular.ttf").is_file():
        return bundle / "Cousine-Regular.ttf"
    found = fonts.installed_index(tuple(fonts.default_font_dirs())).get(("liberation mono", False, False))
    return Path(found[0]) if found and found[0].endswith(".ttf") and found[1] == 0 else None


@pytest.fixture(autouse=True)
def _no_environment(monkeypatch):
    monkeypatch.delenv(FONT_DIRS_ENV, raising=False)


@pytest.fixture
def folder(tmp_path) -> Path:
    source = _open_face()
    if source is None:
        pytest.skip("no open monospaced face here (pptx2svg-fonts, or fonts-liberation)")
    target = tmp_path / "app-fonts" / "probe"
    target.mkdir(parents=True)
    data = relabel(source.read_bytes(), FAMILY, bold=False, italic=False)
    (target / "FontdirsProbeMono-Regular.ttf").write_bytes(data)
    return tmp_path / "app-fonts"


def _document() -> bytes:
    """One page of text in the probe's family, A4."""
    run = f'<w:r><w:rPr><w:rFonts w:ascii="{FAMILY}" w:hAnsi="{FAMILY}"/></w:rPr><w:t>Licensed face</w:t></w:r>'
    styles = (f'<w:styles {_W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{FAMILY}" '
              f'w:hAnsi="{FAMILY}" w:cs="{FAMILY}"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '</w:styles>')
    document = (f'<w:document {_W}><w:body><w:p>{run}</w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" '
                'w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" '
            f'ContentType="application/xml"/><Override PartName="/word/document.xml" '
            f'ContentType="{_CT}.document.main+xml"/><Override PartName="/word/styles.xml" '
            f'ContentType="{_CT}.styles+xml"/></Types>'))
        archive.writestr("_rels/.rels", (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}/officeDocument" Target="word/document.xml"/></Relationships>'))
        archive.writestr("word/_rels/document.xml.rels", (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_REL}/styles" Target="styles.xml"/></Relationships>'))
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", styles)
    return buffer.getvalue()


def _complete(facts: dict) -> bool:
    return facts["complete"] and FAMILY not in facts.get("missing_fonts", [])


def test_the_document_lays_out_with_its_folder(folder, monkeypatch):
    doc = Document.open(_document())
    unnamed = doc.layout().coverage_facts()
    assert not unnamed["complete"] and unnamed["missing_fonts"] == [FAMILY]
    doc.font_dirs = (str(folder),)
    assert _complete(doc.layout().coverage_facts())
    doc.font_dirs = ()
    assert not doc.layout().coverage_facts()["complete"]
    monkeypatch.setenv(FONT_DIRS_ENV, str(folder))
    assert _complete(Document.open(_document()).layout().coverage_facts())


def _box(**options) -> Toolbox:
    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, **options)


def _facts(box: Toolbox, **session_options) -> tuple[dict, dict, dict]:
    """The coverage ``render``, ``check`` and ``save_document`` report."""
    session = box.session(**session_options)
    session.open(_document(), "report.docx")
    render = box.dispatch(session, "render", {"doc": "d1", "pages": [1], "width": 400})
    check = box.dispatch(session, "check", {"doc": "d1", "include": ["reflow"]})
    save = box.dispatch(session, "save_document", {"doc": "d1", "name": "out.docx", "format": "docx"})
    assert render.ok and check.ok and save.ok, (render.error, check.error, save.error)
    return render.data["coverage"], check.data["reflow"]["coverage"], save.data["coverage"]


def test_the_toolbox_reaches_the_layout_and_the_worker(folder, monkeypatch):
    with _box(workers=1) as plain:
        assert not any(_complete(facts) for facts in _facts(plain))

    with _box(workers=1, font_dirs=[folder]) as configured:
        assert all(_complete(facts) for facts in _facts(configured))
        # A session overrides the toolbox: [] is none (no environment variable either).
        assert not any(_complete(facts) for facts in _facts(configured, font_dirs=[]))

    # The environment variable, read where the toolbox runs, reaches the worker too.
    monkeypatch.setenv(FONT_DIRS_ENV, str(folder))
    with _box(workers=1) as from_env:
        assert all(_complete(facts) for facts in _facts(from_env))


def test_in_process_and_the_png_match_a_direct_render(folder):
    if "resvg" not in docx2svg.available_backends():
        pytest.skip("needs resvg-py")
    direct = Document.open(_document())
    direct.font_dirs = (str(folder),)
    expected = direct.render_png(pages=[1], width=400)[0]
    with _box(workers=0, font_dirs=[folder]) as box:
        session = box.session()
        session.open(_document(), "report.docx")
        render = box.dispatch(session, "render", {"doc": "d1", "pages": [1], "width": 400})
    assert render.ok and render.images[0].data == expected and _complete(render.data["coverage"])


def test_the_definitions_and_prompt_do_not_change(folder):
    with _box() as plain, _box(font_dirs=[folder]) as configured:
        for provider in ("anthropic", "openai-responses"):
            assert json.dumps(plain.definitions(provider)) == json.dumps(configured.definitions(provider))
        assert plain.system_prompt() == configured.system_prompt()


# Python 3.14 made forkserver Linux's default start method (fork before).  The toolbox's
# pool names spawn itself; under each method this platform has, the layout and the folders
# handed to it cross to the worker the same.  fork only on Linux: macOS' system libraries
# are not safe to fork with threads running.
_METHODS = [method for method in multiprocessing.get_all_start_methods()
            if method != "fork" or sys.platform.startswith("linux")]


@pytest.mark.parametrize("method", _METHODS)
def test_the_worker_sees_the_folder_under(method, folder):
    with warnings.catch_warnings():
        # fork() with threads running is a DeprecationWarning since 3.12; this test asks.
        warnings.simplefilter("ignore", DeprecationWarning)
        with _box(workers=1, start_method=method, font_dirs=[folder]) as box:
            assert all(_complete(facts) for facts in _facts(box))
            assert not box.pool.in_process
