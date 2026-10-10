import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

import pytest

FIXTURE_DIR = Path(__file__).parent / "fixtures"
#: E6's new document, written where the corpus is read from (``new/document``): every gate
#: and every phase's edit set holds for a document docx-agent made, as for one it opened.
NEW_DOCUMENT = Path(tempfile.gettempdir()) / "docx-agent-tests" / "new" / "document.docx"
#: Its creation date: fixed, so its bytes -- and the oracle's cache -- are the same each run.
NEW_CREATED = datetime(2026, 10, 4, tzinfo=timezone.utc)


def new_document() -> Path:
    """``Document.new()``'s document, written (atomically: xdist's workers each collect) when
    it is not there or differs."""
    from docx_agent import Document

    data = Document.new(created=NEW_CREATED).to_bytes()
    if not NEW_DOCUMENT.exists() or NEW_DOCUMENT.read_bytes() != data:
        NEW_DOCUMENT.parent.mkdir(parents=True, exist_ok=True)
        staged = NEW_DOCUMENT.with_name(f".{os.getpid()}.docx")
        staged.write_bytes(data)
        os.replace(staged, NEW_DOCUMENT)
    return NEW_DOCUMENT


def fixture_paths() -> list[Path]:
    return sorted(FIXTURE_DIR.glob("*/*.docx")) + [new_document()]


def fixture_id(path: Path) -> str:
    return f"{path.parent.name}/{path.stem}"


@pytest.fixture(params=fixture_paths(), ids=fixture_id)
def docx_path(request) -> Path:
    """Every fixture in turn -- the corpus a test must hold for, not one happy document."""
    return request.param


@pytest.fixture(scope="session")
def markup_doc() -> Path:
    """docx-agent's own fixture: every id edge case and the markup text reading walks."""
    return FIXTURE_DIR / "generated" / "ids-and-markup.docx"


@pytest.fixture(scope="session")
def long_doc() -> Path:
    """36 pages in docx2svg's layout, no paraIds: the multi-page reflow fixture."""
    return FIXTURE_DIR / "samplelib" / "sample-long.docx"


MARKDOWN_DIR = FIXTURE_DIR / "generated" / "markdown"
#: The chart and SmartArt documents Word saved (``tools/make_chart_fixtures.py``).
CHARTS_DIR = FIXTURE_DIR / "generated" / "charts"


def reading_paths() -> list[Path]:
    """The corpus, E2's own fixtures and the chart documents (``generated/markdown`` and
    ``generated/charts``, one level below the corpus so the other phases' suites and the Word
    oracle keep theirs): what every reader is held to."""
    return fixture_paths() + sorted(MARKDOWN_DIR.glob("*.docx")) + sorted(CHARTS_DIR.glob("*.docx"))


@pytest.fixture(params=reading_paths(), ids=fixture_id)
def reading_path(request) -> Path:
    return request.param


@pytest.fixture(scope="session")
def constructs_doc() -> Path:
    return MARKDOWN_DIR / "constructs.docx"


@pytest.fixture(scope="session")
def review_doc() -> Path:
    return MARKDOWN_DIR / "review.docx"


@pytest.fixture(scope="session")
def dutch_doc() -> Path:
    return MARKDOWN_DIR / "dutch-template.docx"


#: The fixtures the heaviest per-fixture suites hold by default -- Word-authored with tables
#: in mode 14, a document of every id edge case and revisions, a localised template with
#: lists, and E6's new document; ``--run-slow`` holds them on every fixture.
REPRESENTATIVE = {"wordto/sample-with-table", "generated/ids-and-markup", "generated/lists-and-styles",
                  "new/document"}
#: Those suites, by test function: each runs on :data:`REPRESENTATIVE` by default and on every
#: fixture with ``--run-slow``.
SWEPT = {
    "test_e1_operations.py::test_operation",
    "test_e3_invariant.py::test_accept_all_is_the_edit_and_reject_all_the_original",
    "test_e3_operations.py::test_tracked_operation",
    "test_e3_render.py::test_e3_edits_render_and_read",
    "test_e4_operations.py::test_operation",
    "test_e4_operations.py::test_operation_renders_and_places_what_it_made",
    "test_e4_tracked.py::test_accept_all_is_the_edit_and_reject_all_the_original",
    "test_e5_operations.py::test_operation",
    "test_e5_operations.py::test_operation_renders_and_places_what_it_made",
    "test_e5_tracked.py::test_accept_all_is_the_edit_and_reject_all_the_original",
    "test_markdown_write.py::test_the_corpus_round_trips_in_every_fixture",
    "test_markdown_write.py::test_tracked_insertion_accepts_to_the_edit_and_rejects_to_the_original",
    "test_e6_copy.py::test_every_fixture_copies_into_a_new_document_and_another",
    "test_e6_upgrade.py::test_every_fixture_upgrades_and_reports_its_reflow",
}
#: Randomised sweeps: the seeds held by default (``--run-slow``: every seed).
SEEDS = {"test_e5_table_sweep.py::test_sweep_keeps_the_grid": {"0", "1"},
         "test_e5_table_sweep.py::test_sweep_tracked_accepts_and_rejects": {"0"}}


def _sweep_kept(item) -> bool:
    """Whether the default run holds this item of a swept suite."""
    key = f"{Path(str(item.fspath)).name}::{item.originalname}"
    if key in SEEDS:
        return item.callspec.id in SEEDS[key] if hasattr(item, "callspec") else True
    if key not in SWEPT or not hasattr(item, "callspec"):
        return True
    for value in item.callspec.params.values():
        if isinstance(value, Path):
            return fixture_id(value) in REPRESENTATIVE
    return True


def pytest_addoption(parser):
    parser.addoption("--run-slow", action="store_true", default=False,
                     help="also run the exhaustive sweeps marked slow (ROADMAP.md, \"Testing strategy\")")


def pytest_collection_modifyitems(config, items):
    """The slow sweeps run with ``--run-slow``.  The Word oracle never runs under xdist's
    parallel workers (ROADMAP.md, "Word is one instance per machine"): there, its tests
    skip; run ``pytest -m oracle`` without ``-n``."""
    if not config.getoption("--run-slow"):
        later = pytest.mark.skip(reason="an exhaustive sweep: run with --run-slow")
        for item in items:
            if item.get_closest_marker("slow") is not None or not _sweep_kept(item):
                item.add_marker(later)
    if not os.environ.get("PYTEST_XDIST_WORKER"):
        return
    serial = pytest.mark.skip(reason="the Word oracle runs serially: pytest -m oracle, without -n")
    for item in items:
        if item.get_closest_marker("oracle") is not None:
            item.add_marker(serial)


# -- faces ----------------------------------------------------------------------------------
#
# The fixtures are laid out (by docx2svg) in the faces Word uses -- Calibri, Cambria,
# Georgia, Aptos -- which only a machine with Office has.  Without one, docx2svg stops the
# layout where a paragraph cannot be measured, and a test of a render, a reflow or a field's
# page number then says nothing about this code.  On such a machine, a test whose layout
# reports a face absent skips; where the faces are installed, nothing here applies.  CI's
# Linux and macOS runners install the open substitutes (Carlito, Liberation), which docx2svg
# lays Calibri, Arial, Times New Roman and Courier New out with, so only a document in a
# face without one (Cambria, Georgia, Aptos) skips there.

_OFFICE_FACES = ("Calibri", "Cambria", "Georgia", "Aptos")
_FACE_ABSENT = ("layout-stopped:unmeasurable", "layout-stopped:no face metrics", "line-numbers-not-drawn")


def _office_faces_absent() -> bool:
    from docx2svg.fonts import InstalledFonts, default_font_dirs

    fonts = InstalledFonts(dirs=tuple(default_font_dirs()))
    return any(fonts.face(family) is None for family in _OFFICE_FACES)


OFFICE_FACES_ABSENT = _office_faces_absent()


def _skip_if_a_face_is_absent(warnings, coverage=None) -> None:
    # docx2svg's coverage names the faces it found nowhere, whatever the stop is called (a
    # footnote in an absent face stops as "footnote not measurable").
    missing = getattr(coverage, "missing_fonts", None)
    if missing:
        pytest.skip(f"a face this document is laid out in is not installed here: {', '.join(missing)}")
    for warning in warnings:  # docx2svg's (code, message, page) tuples or its Warning objects
        code, message = (warning.code, warning.message) if hasattr(warning, "code") else warning[:2]
        if code in _FACE_ABSENT or (code.startswith("layout-stopped:") and "cannot be measured" in message):
            pytest.skip(f"a face this document is laid out in is not installed here: [{code}] {message}")


@pytest.fixture(autouse=True)
def _skip_where_faces_are_absent(request, monkeypatch):
    # ``any_faces``: a test that holds without Office's faces (it leaves what they measure out).
    if not OFFICE_FACES_ABSENT or request.node.get_closest_marker("any_faces") is not None:
        return
    import docx2svg

    from docx_agent import layout as agent_layout

    lay_out, convert = docx2svg._lay_out, agent_layout.convert

    def checked_lay_out(source, options):
        result = lay_out(source, options)
        _skip_if_a_face_is_absent(result[0].warnings, getattr(result[0], "coverage", None))
        return result

    def checked_convert(document, **options):
        # A tool layer's worker pool lays out in another process: read what came back.
        conversion = convert(document, **options)
        _skip_if_a_face_is_absent(conversion.layout.warnings, getattr(conversion.layout, "coverage", None))
        return conversion

    monkeypatch.setattr(docx2svg, "_lay_out", checked_lay_out)
    monkeypatch.setattr(agent_layout, "convert", checked_convert)
