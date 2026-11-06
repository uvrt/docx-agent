"""Round 35, R27: an application's live view from the layout the tools share
(:mod:`docx_agent.liveview`).

What each test holds: ``view`` lists every page with its blocks and the pages whose drawing
changed since a token -- the pages docx2svg's own ``changed_pages`` names for the same two
versions; ``page_svg`` draws one page of the token's version, as ``render_svg`` draws it,
pictures by name or embedded, text as text or outlines; the version is laid out once for
the tools and the view together; and views beside checks and edits all finish (round 34's
rule: no wait for a layout in progress under the document's lock).  Calls in threads are
joined with a deadline: a deadlock fails the test, never hangs the run.  Made-up documents
only."""

from __future__ import annotations

import struct
import threading
import time
import zlib

import docx2svg
import pytest

from docx_agent import Document
from docx_agent import layout as layout_module
from docx_agent import liveview
from docx_agent.liveview import UnknownToken, embed_images, page_svg, view

pytestmark = pytest.mark.any_faces

#: How long the calls of one threaded test may take in all before it calls them deadlocked.
DEADLINE = 90.0
#: How long the slowed layout engine stalls before each layout.
STALL = 0.4

WORDS = ("lantern meadow river stone harbour quiet orchard copper valley signal morning ledger window "
         "garden timber bridge market evening canvas thread pebble summit").split()


def _text(seed: int, sentences: int) -> str:
    out = []
    for s in range(sentences):
        out.append(" ".join(WORDS[(seed * 7 + s * 5 + k * 3) % len(WORDS)] for k in range(12)).capitalize() + ".")
    return " ".join(out)


def _png(width: int, height: int, color=(40, 120, 200)) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(color) * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def report(chapters: int = 3) -> bytes:
    """A made-up report of some six pages: chapters on new pages, a table each, a picture,
    a footer with the page number."""
    doc = Document.new()
    parts = ["# Report title", _text(0, 2)]
    for c in range(1, chapters + 1):
        parts.append(f"# Chapter {c}")
        for p in range(3):
            parts.append(_text(c * 10 + p, 6))
        parts.append("| Item | Owner | Amount |\n|---|---|---|\n"
                     + "\n".join(f"| Item {c}.{k} | {WORDS[k % len(WORDS)]} | {k * 7} |" for k in range(8)))
    doc.insert_markdown("\n\n".join(parts))
    for paragraph in doc.paragraphs():
        if paragraph.text.startswith("Chapter "):
            paragraph.format(page_break_before=True)
    doc.insert_picture(f"{doc.paragraphs()[1].id}@0", _png(40, 30), width=60)
    section = doc.sections()[0]
    doc.add_footer(section.id, "default", "Page ")
    doc.insert_field(f"{section.footer().paragraphs[0].id}@5", "PAGE", result="1")
    return doc.to_bytes()


@pytest.fixture(scope="module")
def made() -> bytes:
    return report()


def _body(document: Document) -> list:
    return [p for p in document.paragraphs() if len(p.text) > 200]


def _edits():
    """``(name, edit(document))``: edits of each kind a builder makes."""

    def text(document):
        _body(document)[4].set_text("A shorter paragraph now.")

    def longer(document):
        paragraph = _body(document)[2]
        paragraph.set_text(paragraph.text + " " + _text(99, 12))

    def row(document):
        document.insert_row(document.tables()[1].id, 2)

    def delete(document):
        _body(document)[6].delete()

    def style(document):
        document.styles.modify("Heading 1", size=23)

    def nothing(document):
        pass

    return [(f.__name__, f) for f in (text, longer, row, delete, style, nothing)]


# -- view --------------------------------------------------------------------------------------


def test_the_first_view_lists_every_page_with_its_blocks(made):
    document = Document.open(made)
    shown = view(document)
    layout = document.layout()
    assert shown.page_count == layout.page_count > 3
    assert shown.changed == list(range(1, shown.page_count + 1))
    assert [list(p.blocks) for p in shown.pages] == [layout.blocks_on(n) for n in range(1, layout.page_count + 1)]
    assert all(p.width > 500 and p.height > 700 for p in shown.pages)
    again = view(document, since=shown.token)
    assert again.token == shown.token and again.changed == [] and not again.stale


@pytest.mark.parametrize("name, edit", _edits(), ids=[name for name, _ in _edits()])
def test_changed_pages_are_docx2svgs_for_the_same_versions(made, name, edit):
    document = Document.open(made)
    shown = view(document)
    edit(document)
    after = view(document, since=shown.token)
    # docx2svg's own answer for the same two versions: laid out from the one before.
    before, _, _ = docx2svg.lay_out_docx(made, docx2svg.ConvertOptions(**layout_module.layout_options(document, {})))
    new, _, _ = docx2svg.lay_out_docx(document.to_bytes(),
                                      docx2svg.ConvertOptions(**layout_module.layout_options(document, {})),
                                      previous=before)
    assert after.changed == new.changed_pages, name
    if name == "nothing":
        assert after.token == shown.token and after.changed == []


def test_changes_since_an_older_token_are_every_page_that_differs(made):
    """Two edits on two pages, the view asked since before both."""
    document = Document.open(made)
    first = view(document)
    _body(document)[0].set_text("Changed on the first chapter's page.")
    view(document, since=first.token)
    _body(document)[-1].set_text("Changed on the last chapter's page.")
    since_first = view(document, since=first.token)
    layout = document.layout()
    on = sorted(set(layout.pages_of(_body(document)[0].id)) | set(layout.pages_of(_body(document)[-1].id)))
    assert set(on) <= set(since_first.changed)


def test_a_stale_token_says_so_and_lists_every_page(made):
    document = Document.open(made)
    shown = view(document, since="0000-not-a-token")
    assert shown.stale and "every page" in shown.note
    assert shown.changed == list(range(1, shown.page_count + 1))
    other = view(Document.open(made))
    assert view(document, since=other.token).stale                     # another document's


def test_old_tokens_are_dropped_past_the_kept_number(made, monkeypatch):
    monkeypatch.setattr(liveview, "VIEW_TOKENS", 2)
    document = Document.open(made)
    tokens = [view(document).token]
    for k in range(2):
        _body(document)[k].set_text(f"Edit {k}.")
        tokens.append(view(document).token)
    assert view(document, since=tokens[0]).stale
    with pytest.raises(UnknownToken, match="call view"):
        page_svg(document, 1, tokens[0])


# -- page_svg ----------------------------------------------------------------------------------


def test_a_page_is_drawn_as_render_svg_draws_it(made):
    document = Document.open(made)
    shown = view(document)
    for number in (1, shown.page_count):
        embedded = page_svg(document, number, shown.token, images="embedded")
        assert embedded.svg == document.render_svg([number])[0]
        assert embedded.images == {}
        assert "data-docx-agent-id" in embedded.svg


def test_pictures_come_back_by_name(made):
    document = Document.open(made)
    shown = view(document)
    page = next(p.number for p in shown.pages if document.paragraphs()[1].id in p.blocks)
    drawn = page_svg(document, page, shown.token)
    [(name, data)] = drawn.images.items()
    assert name.startswith("img/") and name.endswith(".png") and len(name) == len("img/") + 32 + 4
    assert data.startswith(b"\x89PNG") and "base64," not in drawn.svg
    assert embed_images(drawn.svg, drawn.images) == page_svg(document, page, shown.token, images="embedded").svg
    with pytest.raises(KeyError, match="no such entry"):
        embed_images(drawn.svg, {})


def test_outlines_with_one_glyph_set(made):
    document = Document.open(made)
    shown = view(document)
    glyphs: dict = {}
    one = page_svg(document, 1, shown.token, text="outlines", glyphs=glyphs)
    count = len(glyphs)
    two = page_svg(document, 2, shown.token, text="outlines", glyphs=glyphs)
    assert count > 0 and len(glyphs) >= count
    for drawn in (one, two):
        assert "<use" in drawn.svg and "<text" not in drawn.svg and "data-docx-agent-id" in drawn.svg
    alone = page_svg(document, 1, shown.token, text="outlines")
    assert docx2svg.embed_glyphs(one.svg, glyphs) == alone.svg


def test_an_older_tokens_page_is_that_versions(made):
    """The document moved on: the page drawn is the token's version, with its ids."""
    document = Document.open(made)
    shown = view(document)
    target = _body(document)[0]
    page = document.layout().pages_of(target.id)[0]
    before = page_svg(document, page, shown.token, images="embedded").svg
    target.set_text("Rewritten after the view.")
    assert page_svg(document, page, shown.token, images="embedded").svg == before
    assert Document.open(made).render_svg([page])[0] == before


def test_refusals_say_the_way_out(made):
    document = Document.open(made)
    shown = view(document)
    with pytest.raises(IndexError, match=f"1 to {shown.page_count}"):
        page_svg(document, shown.page_count + 1, shown.token)
    with pytest.raises(ValueError, match="text, outlines"):
        page_svg(document, 1, shown.token, text="paths")
    with pytest.raises(ValueError, match="external, embedded"):
        page_svg(document, 1, shown.token, images="inline")
    with pytest.raises(UnknownToken, match="call view"):
        page_svg(document, 1, "nothing")
    with pytest.raises(TypeError, match=r"\(document, page, token\) or \(session, doc, page, token\)"):
        page_svg(document, 1)


# -- one layout per version, shared, and no wait under the lock -----------------------------------


def _toolbox():
    from ooxml_edit.tools import Toolbox

    from docx_agent.tools import FORMAT, GROUPS, TOOLS

    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0)


def _counted(monkeypatch, stall: float = 0.0) -> tuple[list, threading.Event]:
    """docx2svg's layouts counted (and each stalled by ``stall``); the event is set when the
    first began."""
    calls: list = []
    started = threading.Event()
    run = docx2svg.lay_out_docx

    def counted(*args, **kwargs):
        calls.append(1)
        started.set()
        time.sleep(stall)
        return run(*args, **kwargs)

    monkeypatch.setattr(docx2svg, "lay_out_docx", counted)
    return calls, started


def test_a_view_after_a_check_lays_nothing_out(made, monkeypatch):
    pytest.importorskip("ooxml_edit.tools.shared")
    calls, _ = _counted(monkeypatch)
    with _toolbox() as box:
        session = box.session()
        d1 = session.open(made, "report.docx")
        assert box.dispatch(session, "check", {"doc": d1, "include": ["reflow"]}).ok
        laid = len(calls)
        shown = view(session, d1)
        assert len(calls) == laid
        assert shown.version == session.entry(d1).version
        drawn = page_svg(session, d1, 1, shown.token)
        assert len(calls) == laid and drawn.svg.startswith("<svg")
        # An edit, then a check: the view of the new version finds that layout too.
        target = _body(session.entry(d1).document)[0].id
        assert box.dispatch(session, "word_set_text", {"doc": d1, "target": target, "text": "Shorter."}).ok
        assert box.dispatch(session, "check", {"doc": d1}).ok
        laid = len(calls)
        after = view(session, d1, since=shown.token)
        assert len(calls) == laid and after.changed


def _all_finish(jobs: list) -> list:
    """Every job run in its own daemon thread, all joined by :data:`DEADLINE`: a deadlock
    fails the test, never hangs the run.  Each job's result or error."""
    out: list = [None] * len(jobs)

    def run(k, job):
        try:
            out[k] = job()
        except BaseException as error:  # noqa: BLE001 -- reported below
            out[k] = error

    threads = [threading.Thread(target=run, args=(k, job), daemon=True) for k, job in enumerate(jobs)]
    for thread in threads:
        thread.start()
    ends = time.monotonic() + DEADLINE
    for thread in threads:
        thread.join(max(0.0, ends - time.monotonic()))
    hung = [t for t in threads if t.is_alive()]
    if hung:
        pytest.fail(f"deadlock: {len(hung)} of {len(threads)} calls still waiting after {DEADLINE:g} s")
    errors = [r for r in out if isinstance(r, BaseException)]
    assert not errors, errors
    return out


def test_a_view_waits_for_a_checks_layout_with_the_lock_let_go(made, monkeypatch):
    """A lone check lays the version out (stalled); views asked meanwhile wait for that one
    layout, the document's lock let go -- one layout in all."""
    calls, started = _counted(monkeypatch, STALL)
    with _toolbox() as box:
        session = box.session()
        d1 = session.open(made, "report.docx")

        def check():
            result = box.dispatch(session, "check", {"doc": d1, "include": ["reflow"]})
            assert result.ok, result.error
            return result.data["reflow"]["pages"]

        first = threading.Thread(target=check, daemon=True)
        first.start()
        assert started.wait(DEADLINE)
        views = _all_finish([lambda: view(session, d1)] * 3)
        first.join(DEADLINE)
        assert not first.is_alive()
    assert len(calls) == 1
    assert len({v.token for v in views}) == 1


def test_a_view_by_a_thread_holding_the_lock_lays_out_on_its_own(made, monkeypatch):
    """The application's thread holds the document's lock: its view never waits for the
    check's layout in progress (whose owner takes that lock back) -- it lays out alone."""
    calls, started = _counted(monkeypatch, STALL)
    with _toolbox() as box:
        session = box.session()
        d1 = session.open(made, "report.docx")
        entry = session.entry(d1)
        first = threading.Thread(target=lambda: box.dispatch(session, "check", {"doc": d1, "include": ["reflow"]}),
                                 daemon=True)
        first.start()
        assert started.wait(DEADLINE)

        def held():
            with entry.lock:
                return view(session, d1)

        [shown] = _all_finish([held])
        first.join(DEADLINE)
        assert not first.is_alive()
    assert shown.page_count > 3 and len(calls) == 2


def test_views_checks_and_edits_together_all_finish(made, monkeypatch):
    """Builders at once: lone checks, read_batches, edits and the application's views and
    page drawing, a few rounds; every view consistent with the version it names."""
    _counted(monkeypatch, 0.1)
    with _toolbox() as box:
        session = box.session()
        d1 = session.open(made, "report.docx")
        body = [p.id for p in _body(session.entry(d1).document)]
        tokens: list = []

        def looking():
            shown = view(session, d1, since=tokens[-1] if tokens else None)
            tokens.append(shown.token)
            for number in shown.changed[:2]:
                page_svg(session, d1, number, shown.token, text="outlines")
            return shown

        def call(name, arguments):
            def run():
                result = box.dispatch(session, name, arguments)
                assert result.ok, result.error and result.error.to_json()
                return result
            return run

        for k in range(2):
            out = _all_finish([
                call("check", {"doc": d1, "include": ["reflow"]}), looking,
                call("word_set_text", {"doc": d1, "target": body[1 + k], "text": f"Round {k}."}),
                looking, call("read_batch", {"ops": [{"tool": "check", "arguments": {"doc": d1}}]}),
                call("check", {"doc": d1, "range": [body[3]]}), looking])
            assert all(o is not None for o in out)
        final = view(session, d1)
        assert final.page_count == session.entry(d1).document.layout().page_count
