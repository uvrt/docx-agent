"""Incremental, shared layouts (round 33, R16): a document is laid out from its last layout
(:func:`docx_agent.layout.lay_out_bytes`, docx2svg's ``previous``), concurrent callers asking
for one state share one layout, and ``layout_facts(pages=...)`` does only those pages' work.

The rule each test holds: what an incremental layout answers is what a layout of the bytes
alone answers -- every page's SVG byte for byte, the page signatures, the warnings, the
coverage and every layout fact -- after edits of each kind (text, table rows, a section
break, a style, fields, deletions), including those docx2svg lays out whole for
(:data:`docx2svg.relayout.FALLBACKS`).  Made-up documents only."""

from __future__ import annotations

import threading
import time

import pytest

from docx_agent import Document
from docx_agent import layout as layout_module
from docx_agent.layout_facts import layout_facts

# What each test compares holds with whatever faces the layout finds: run where Office's
# faces are absent too (CI), laid out with the stand-ins docx2svg picks.
pytestmark = pytest.mark.any_faces

WORDS = ("lantern meadow river stone harbour quiet orchard copper valley signal morning ledger window "
         "garden timber bridge market evening canvas thread pebble summit").split()


def _text(seed: int, sentences: int) -> str:
    out = []
    for s in range(sentences):
        words = [WORDS[(seed * 7 + s * 5 + k * 3) % len(WORDS)] for k in range(12)]
        out.append(" ".join(words).capitalize() + ".")
    return " ".join(out)


def report(chapters: int = 4) -> bytes:
    """A made-up report of some ten pages: contents, chapters starting on new pages, a table
    in each (one long enough to break across pages), a footer with the page number."""
    doc = Document.new()
    parts = ["# Report title", _text(0, 2)]
    for c in range(1, chapters + 1):
        parts.append(f"# Chapter {c}")
        for p in range(3):
            parts.append(_text(c * 10 + p, 6))
        rows = 45 if c == 2 else 5
        parts.append("| Item | Owner | Amount |\n|---|---|---|\n"
                     + "\n".join(f"| Item {c}.{k} | {WORDS[k % len(WORDS)]} | {k * 7} |" for k in range(rows)))
        parts.append(_text(c * 10 + 9, 3))
    doc.insert_markdown("\n\n".join(parts))
    for paragraph in doc.paragraphs():
        if paragraph.text.startswith("Chapter "):
            paragraph.format(page_break_before=True)
    doc.insert_toc(after=doc.paragraphs()[1].id, levels=(1, 1))
    section = doc.sections()[0]
    doc.add_footer(section.id, "default", "Page ")
    doc.insert_field(f"{section.footer().paragraphs[0].id}@5", "PAGE", result="1")
    return doc.to_bytes()


@pytest.fixture(scope="module")
def made() -> bytes:
    return report()


def _answers(document: Document) -> dict:
    """Everything a layout answers, comparable."""
    layout = document.layout()
    return {"pages": layout.page_count,
            "signatures": [layout.signature(n) for n in range(1, layout.page_count + 1)],
            "warnings": [str(w) for w in layout.warnings],
            "coverage": layout.coverage_facts(),
            "facts": layout_facts(layout),
            "svgs": document.render_svg()}


def _alone(data: bytes) -> dict:
    return _answers(Document.open(data))


def _body(document: Document) -> list:
    return [p for p in document.paragraphs() if len(p.text) > 200]


def _edits():
    """``(name, edit(document))``: one edit of each kind."""

    def text(document):
        _body(document)[4].set_text("A shorter paragraph now.")

    def longer(document):
        paragraph = _body(document)[2]
        paragraph.set_text(paragraph.text + " " + _text(99, 8))

    def rows(document):
        table = document.tables()[0]
        document.insert_row(table.id, 2)
        document.insert_row(table.id, 3)

    def delete_row(document):
        long = max(document.tables(), key=lambda t: len(t.rows))
        document.delete_row(long.id, 10)

    def section(document):
        document.insert_section_break(after=_body(document)[5].id)

    def style(document):
        document.styles.modify("Heading 1", size=23, space_before=30)

    def fields(document):
        _body(document)[1].set_text(_text(55, 30))
        document.update_fields()

    def delete(document):
        _body(document)[6].delete()

    def nothing(document):
        pass

    return [("text", text), ("longer", longer), ("rows", rows), ("delete_row", delete_row),
            ("section", section), ("style", style), ("fields", fields), ("delete", delete),
            ("nothing", nothing)]


@pytest.mark.parametrize("name,edit", _edits(), ids=[name for name, _ in _edits()])
def test_a_layout_from_the_last_one_answers_as_a_layout_of_the_bytes_alone(made, name, edit):
    document = Document.open(made)
    document.layout()                     # the layout the next one is made from
    before = document.to_bytes()
    edit(document)
    calls = []
    real = layout_module.lay_out_bytes

    def counted(data, options, state=None):
        calls.append(state)
        return real(data, options, state)

    layout_module.lay_out_bytes, saved = counted, real
    try:
        incremental = _answers(document)
    finally:
        layout_module.lay_out_bytes = saved
    assert incremental == _alone(document.to_bytes())
    if name == "nothing":
        assert calls == []                # the same state: its layout cached
    elif name != "fields":                # (update_fields laid it out itself)
        assert document.to_bytes() != before
        assert calls and (isinstance(calls[-1], bytes) or calls[-1][0] == "previous")


def test_a_run_of_edits_each_laid_out_from_the_one_before(made):
    document = Document.open(made)
    for _, edit in _edits():
        document.layout()
        edit(document)
    assert _answers(document) == _alone(document.to_bytes())


def test_the_kept_layout_is_the_last_and_one_per_document(made):
    document = Document.open(made)
    document.layout()
    first = document._layout_kept
    _body(document)[3].set_text("Short.")
    document.layout()
    assert document._layout_kept is not first and len(document._layout_kept) == 2
    assert layout_module.kept_layout(document, document.to_bytes(),
                                     layout_module.layout_options(document, {})) == document._layout_kept[1]


def test_a_render_and_a_layout_of_one_state_share_its_layout(made):
    document = Document.open(made)
    calls = []
    real = layout_module.lay_out_bytes
    layout_module.lay_out_bytes = lambda *a: calls.append(1) or real(*a)
    try:
        document.render_svg(pages=[2])
        document.layout()
        document._layouts.clear()         # the conversion dropped: the kept layout answers
        document.layout()
    finally:
        layout_module.lay_out_bytes = real
    assert len(calls) == 1


# -- shared: one layout per state in flight ---------------------------------------------------


def test_concurrent_callers_of_one_state_share_one_layout(made):
    document = Document.open(made)
    started = threading.Event()
    calls = []
    real = layout_module.lay_out_bytes

    def slow(data, options, state=None):
        calls.append(1)
        started.set()
        time.sleep(0.3)                   # the others ask while it runs
        return real(data, options, state)

    layout_module.lay_out_bytes = slow
    got: list = [None] * 8
    errors: list = []

    def ask(k: int) -> None:
        try:
            got[k] = document.layout()
        except Exception as error:  # noqa: BLE001 -- reported below
            errors.append(error)

    try:
        threads = [threading.Thread(target=ask, args=(k,)) for k in range(8)]
        threads[0].start()
        started.wait(10)
        for thread in threads[1:]:
            thread.start()
        for thread in threads:
            thread.join()
    finally:
        layout_module.lay_out_bytes = real
    assert errors == [] and len(calls) == 1
    assert all(layout is got[0] for layout in got)


def test_a_failed_layout_fails_every_caller_waiting_for_it_and_is_not_kept(made):
    document = Document.open(made)
    started = threading.Event()
    real = layout_module.lay_out_bytes

    def failing(data, options, state=None):
        started.set()
        time.sleep(0.2)
        raise RuntimeError("the layout failed")

    layout_module.lay_out_bytes = failing
    errors: list = []

    def ask() -> None:
        try:
            document.layout()
        except RuntimeError as error:
            errors.append(str(error))

    try:
        first = threading.Thread(target=ask)
        first.start()
        started.wait(10)
        second = threading.Thread(target=ask)
        second.start()
        first.join()
        second.join()
    finally:
        layout_module.lay_out_bytes = real
    assert errors == ["the layout failed"] * 2
    assert document._layout_flights == {} and document._layout_kept is None
    assert document.layout().page_count > 1          # and the next one is laid out


def test_waiting_callers_wait_unlocked(made):
    """A caller that waits for another's layout does so inside ``unlocked`` (the tool layer
    lets the document's lock go there), as the one laying out does."""
    document = Document.open(made)
    started = threading.Event()
    real = layout_module.lay_out_bytes
    inside: list = []

    def slow(data, options, state=None):
        started.set()
        time.sleep(0.3)
        return real(data, options, state)

    class Unlocked:
        def __init__(self, name):
            self.name = name

        def __enter__(self):
            inside.append(("in", self.name))

        def __exit__(self, *exc):
            inside.append(("out", self.name))

    layout_module.lay_out_bytes = slow
    try:
        first = threading.Thread(target=lambda: layout_module.convert(document, unlocked=lambda: Unlocked("a")))
        first.start()
        started.wait(10)
        layout_module.convert(document, unlocked=lambda: Unlocked("b"))
        first.join()
    finally:
        layout_module.lay_out_bytes = real
    assert ("in", "b") in inside and ("out", "b") in inside


# -- page-scoped facts --------------------------------------------------------------------------


def test_facts_for_some_pages_are_the_whole_documents_on_those_pages(made):
    document = Document.open(made)
    long = max(document.tables(), key=lambda t: len(t.rows))
    long.cell(1, 1).paragraphs[0].format(alignment="center")      # direct formatting in it
    _body(document)[5].format(space_after=30)
    layout = document.layout()
    whole = layout_facts(layout)
    for pages in ([1], [2, 3], [4], list(range(1, layout.page_count + 1))):
        some = layout_facts(layout, pages=pages)
        for name in ("direct_formatting", "column_fit", "numbering_skips_level", "small_text", "spacers",
                     "keep_with_next_at_foot", "split_rows", "past_margins"):
            expected = [e for e in whole.get(name, []) if _on(document, layout, e, pages)]
            assert some.get(name, []) == expected, (name, pages)


def _on(document, layout, entry: dict, pages: list[int]) -> bool:
    if "pages" in entry:
        return any(p in pages for p in entry["pages"])
    if "page" in entry:
        return entry["page"] in pages
    return any(p in pages for p in layout.pages_of(entry["para"]))


def test_facts_for_some_pages_do_only_those_pages_work(made, monkeypatch):
    from docx_agent import layout_facts as facts_module

    document = Document.open(made)
    layout = document.layout()
    fitted, direct = [], []
    real_fit, real_direct = facts_module._fit_columns, facts_module._Facts._direct
    monkeypatch.setattr(facts_module, "_fit_columns", lambda *a: fitted.append(1) or real_fit(*a))
    monkeypatch.setattr(facts_module._Facts, "_direct", lambda self, *a: direct.append(1) or real_direct(self, *a))
    layout_facts(layout)
    whole = (len(fitted), len(direct))
    fitted.clear()
    direct.clear()
    layout_facts(layout, pages=[1])
    assert len(fitted) < whole[0] and len(direct) < whole[1]


# -- through the tool layer ---------------------------------------------------------------------


def _toolbox(**options):
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, **options)


def test_concurrent_checks_of_one_version_lay_it_out_once_unlocked(made, monkeypatch):
    """Lone checks let the document's lock go while they lay out -- and while they wait for
    another's layout -- so a read meanwhile does not wait for them, and six of one version
    cost one layout."""
    from docx_agent.tools import _base

    laid = []
    started = threading.Event()
    ended: list = []
    run = _base.Engine.lay_out

    def slow(self, data, options, state=None):
        laid.append(1)
        started.set()
        time.sleep(1.0)
        try:
            return run(self, data, options, state)
        finally:
            ended.append(time.perf_counter())

    monkeypatch.setattr(_base.Engine, "lay_out", slow)
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        first = session.entry(doc).document.paragraphs()[0].id
        results: list = [None] * 6

        def check(k: int) -> None:
            results[k] = box.dispatch(session, "check", {"doc": doc, "include": ["reflow"]})

        threads = [threading.Thread(target=check, args=(k,)) for k in range(6)]
        threads[0].start()
        started.wait(30)
        for thread in threads[1:]:
            thread.start()
        read = box.dispatch(session, "word_read", {"doc": doc, "range": first})
        read_at = time.perf_counter()
        for thread in threads:
            thread.join()
    assert read.ok and read_at < ended[0]          # read while the layout went on
    assert all(r.ok for r in results) and len(laid) == 1
    assert len({r.data["reflow"]["pages"] for r in results}) == 1


def test_a_check_after_a_render_of_the_same_version_lays_nothing_out(made, monkeypatch):
    pytest.importorskip("resvg_py", reason="render needs the rasteriser")
    from docx_agent.tools import _base

    laid = []
    run = _base.Engine.lay_out
    monkeypatch.setattr(_base.Engine, "lay_out", lambda self, *a: laid.append(1) or run(self, *a))
    with _toolbox(workers=0) as box:
        session = box.session()
        doc = session.open(made, "report.docx")
        assert box.dispatch(session, "render", {"doc": doc, "pages": [2], "width": 300}).ok
        check = box.dispatch(session, "check", {"doc": doc, "include": ["reflow", "layout"]})
    assert check.ok and laid == []
