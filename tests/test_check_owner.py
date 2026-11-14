"""One layout owner per document, one trip per check (lane O, docx-agent 0.1.112).

* A check that needs the layout waits for the document's next layout -- laid out once for
  every check waiting, of the newest committed version -- and reads that version
  (:class:`docx_agent.layout.Owner`, ooxml-edit's ``read_version``): its result is the one a
  check of that version alone gives, and says when it is newer than at the call.
* Its numbers' pages go to the worker in the check's own trip (``check_plan``): one task
  per check, not two.
* The layout facts every check of a layout reads alike are made once per layout.

Made-up documents only."""

from __future__ import annotations

import copy
import dataclasses
import threading
import time

import pytest

pytest.importorskip("ooxml_edit.tools.numbers", reason="needs ooxml-edit (numbers)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.layout import Owner  # noqa: E402
from docx_agent.layout_facts import layout_facts  # noqa: E402
from docx_agent.scope import Scope, ranges_of  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS as WORD_TOOLS  # noqa: E402
from docx_agent.tools import _base as base_module  # noqa: E402
from docx_agent.tools import owner as owner_module  # noqa: E402
from docx_agent.tools.scoped import range_pages  # noqa: E402

from test_tables_from_data import ds_put  # noqa: E402

TOOLS = [*WORD_TOOLS, ds_put]
INCLUDE = ["validate", "layout", "numbers", "reflow", "fields"]

pytestmark = pytest.mark.any_faces

WORDS = "lantern meadow river stone harbour quiet orchard copper valley signal".split()


def _text(seed: int, words: int) -> str:
    return " ".join(WORDS[(seed + k) % len(WORDS)] for k in range(words)).capitalize() + "."


@pytest.fixture(scope="module")
def made() -> bytes:
    doc = Document.new()
    parts = ["# Made-up report"]
    for c in range(1, 5):
        parts += [f"# Part {c}", f"{_text(c, 60)} Rated {c}.5 of 5 in part {c}.",
                  "- first point\n- second point",
                  f"| Item | Amount |\n|---|---|\n| lantern | {c}2.5 |\n| meadow | {c}7.25 |",
                  _text(c + 3, 160)]
    doc.insert_markdown("\n\n".join(parts))
    return doc.to_bytes()


def _session(box, data):
    session = box.session()
    doc = session.open(data, "report.docx")
    assert box.dispatch(session, "ds_put", {"id": "values"}).ok
    return session, doc, session.entry(doc).document


def _heads(document):
    return [p.id for p in document.paragraphs() if p.text.startswith("Part ")]


def _same(result) -> dict:
    """A check's result without its timing and the version note."""
    data = copy.deepcopy(result.data)
    data.pop("timing", None)
    return {"data": data, "warnings": [w for w in result.warnings if "newer than at the call" not in w]}


# -- the owner itself --------------------------------------------------------------------------


@dataclasses.dataclass
class _Version:
    version: int


def test_owner_lays_out_once_for_every_check_waiting_and_then_the_newest():
    owner = Owner()
    newest = [_Version(1)]
    laid: list[int] = []
    gate = threading.Event()

    def lay(target):
        laid.append(target.version)
        if target.version == 1:
            assert gate.wait(5)
        return True

    served: dict[str, int] = {}

    entered: list[str] = []

    def check(name, version):
        entered.append(name)
        found = owner.serve(version, lambda: newest[0], lay)
        served[name] = found.version
        owner.release(found.version)            # its facts made

    first = threading.Thread(target=check, args=("first", 1))
    first.start()
    while not laid:
        time.sleep(0.01)
    # Commits and checks while version 1 is laid out: all served by one layout of the newest.
    newest[0] = _Version(2)
    waiting = [threading.Thread(target=check, args=(f"w{v}", v)) for v in (1, 2)]
    for thread in waiting:
        thread.start()
    newest[0] = _Version(4)
    late = threading.Thread(target=check, args=("w4", 4))
    late.start()
    while len(entered) < 4:
        time.sleep(0.01)
    time.sleep(0.2)
    gate.set()
    for thread in [first, *waiting, late]:
        thread.join(5)
    assert served["first"] == 1 and served["w1"] in (1, 4)     # version 1 includes it
    assert served["w2"] == 4 and served["w4"] == 4             # the next layout: the newest
    assert laid == [1, 4] and owner.flights == 2 and owner.readers == {}


def test_the_next_layout_waits_for_the_checks_reading_the_last_one(monkeypatch):
    """Their facts are made from the layout they read, before it is taken on to the next."""
    import docx_agent.layout as layout_module

    monkeypatch.setattr(layout_module, "READERS_WAIT", 5.0)
    owner = Owner()
    order: list[str] = []
    assert owner.serve(1, lambda: _Version(1), lambda target: True).version == 1      # a reader of 1

    def later():
        owner.serve(2, lambda: _Version(2), lambda target: order.append("layout 2") or True)

    thread = threading.Thread(target=later)
    thread.start()
    time.sleep(0.3)
    order.append("facts of 1")
    owner.release(1)
    thread.join(5)
    assert order == ["facts of 1", "layout 2"]
    # A reader that never says it is done is waited for READERS_WAIT s at most.
    monkeypatch.setattr(layout_module, "READERS_WAIT", 0.2)
    started = time.monotonic()
    assert owner.serve(3, lambda: _Version(3), lambda target: True).version == 3
    assert time.monotonic() - started < 2


def test_owner_does_not_serve_a_version_whose_layout_failed(monkeypatch):
    import docx_agent.layout as layout_module

    monkeypatch.setattr(layout_module, "READERS_WAIT", 0.1)
    owner = Owner()
    assert owner.serve(3, lambda: _Version(3), lambda target: False) is None
    assert owner.serve(2, lambda: _Version(3), lambda target: True) is None   # 3 includes 2: failed
    assert owner.serve(4, lambda: _Version(4), lambda target: True).version == 4


# -- checks served by the owner --------------------------------------------------------------


@pytest.mark.parametrize("workers", [0, 2])
def test_a_check_reports_what_a_check_of_its_version_alone_reports(made, workers):
    """With the owner and without it (no read_version), the same results, version by version."""
    without = dataclasses.replace(FORMAT, read_version=None)
    results = {}
    for name, fmt in (("owner", FORMAT), ("alone", without)):
        with Toolbox(TOOLS, formats=[fmt], groups=GROUPS, workers=workers) as box:
            session, doc, document = _session(box, made)
            heads = _heads(document)
            got = [box.dispatch(session, "check", {"doc": doc, "include": INCLUDE})]
            for k in range(2):
                box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{heads[k + 1]}",
                                                               "markdown": f"A new paragraph {k}: rated 4.{k}."})
                got.append(box.dispatch(session, "check", {"doc": doc, "include": INCLUDE,
                                                           "range": [f"{heads[k]}..{heads[k + 1]}"]}))
                got.append(box.dispatch(session, "check", {"doc": doc, "include": ["validate", "reflow"]}))
            assert all(r.ok for r in got), [r.to_json() for r in got if not r.ok]
            results[name] = [(r.version, _same(r)) for r in got]
    assert results["owner"] == results["alone"]


def test_a_check_with_numbers_goes_to_the_worker_once(made, monkeypatch):
    trips: list[dict] = []
    real = base_module.Engine.check

    def counted(self, data, options, state, request):
        trips.append(request)
        return real(self, data, options, state, request)

    monkeypatch.setattr(base_module.Engine, "check", counted)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=2) as box:
        session, doc, document = _session(box, made)
        heads = _heads(document)
        box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{heads[2]}",
                                                       "markdown": "One more paragraph: 3.25 of it."})
        trips.clear()
        result = box.dispatch(session, "check", {"doc": doc, "include": INCLUDE,
                                                 "range": [f"{heads[1]}..{heads[2]}"]})
        assert result.ok, result.to_json()
        assert result.data["numbers"]["texts"] and result.data["range"]
        # The owner's layout, then one trip with the range's facts and the numbers' pages.
        assert len(trips) == 2, trips
        assert trips[0]["ranges"] is None and not trips[0]["layout"] and "pages_of" not in trips[0]
        assert trips[1]["ranges"] and trips[1]["layout"] and trips[1]["pages_of"]
        assert {"layout", "numbers"} <= set(result.data["timing"])
        trips.clear()
        again = box.dispatch(session, "check", {"doc": doc, "include": INCLUDE,
                                                "range": [f"{heads[1]}..{heads[2]}"]})
        assert again.ok and trips == []                     # the same check of the same version


def test_a_check_waiting_for_a_layout_reads_the_newest_version_and_says_so(made, monkeypatch):
    """A check called while another version is laid out is served by the next layout, of the
    newest version: its result is of that version, and a warning says how many commits newer."""
    gate = threading.Event()
    started = threading.Event()
    real = owner_module.own_layout
    calls: list[int] = []

    def held(document, data, converter):
        calls.append(1)
        if len(calls) == 2:            # the layout after the first edit: held up
            started.set()
            assert gate.wait(30)
        return real(document, data, converter)

    monkeypatch.setattr(owner_module, "own_layout", held)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, document = _session(box, made)
        heads = _heads(document)
        assert box.dispatch(session, "check", {"doc": doc}).ok                       # layout 1
        edit = lambda k: box.dispatch(session, "word_insert_markdown", {  # noqa: E731
            "doc": doc, "at": f"before:{heads[1]}", "markdown": f"Paragraph {k}."})
        assert edit(1).ok
        out: dict = {}
        first = threading.Thread(target=lambda: out.setdefault("first", box.dispatch(session, "check", {"doc": doc})))
        first.start()
        assert started.wait(30)        # the layout of the first edit's version is in flight
        assert edit(2).ok
        at_call = session.entry(doc).version
        second = threading.Thread(target=lambda: out.setdefault("second", box.dispatch(session, "check", {"doc": doc})))
        second.start()
        time.sleep(0.5)
        assert edit(3).ok              # committed while the second check waits
        newest = session.entry(doc).version
        gate.set()
        first.join(60)
        second.join(60)
        assert out["first"].ok and out["first"].version == at_call - 1
        assert out["second"].ok and out["second"].version == newest == at_call + 1
        said = [w for w in out["second"].warnings if "newer than at the call" in w]
        assert said == [f"of version {newest}: 1 commit(s) newer than at the call (version {at_call}), "
                        "laid out for every check waiting meanwhile"]
        # What it reports is what a check of that version alone reports.
        alone = dataclasses.replace(FORMAT, read_version=None)
        data = session.entry(doc).document.to_bytes()
    with Toolbox(TOOLS, formats=[alone], groups=GROUPS, workers=0) as box:
        other = box.session()
        copy_doc = other.open(data, "report.docx")
        reference = box.dispatch(other, "check", {"doc": copy_doc})
        assert reference.ok
        mine, theirs = _same(out["second"]), _same(reference)
        mine["data"]["reflow"].pop("since_last_check", None)
        assert mine == theirs


def test_owner_places_versions_by_commit_order_not_number():
    """0.1.113: an undo brings an earlier version number back; the version committed last is
    the newest whatever its number."""

    @dataclasses.dataclass
    class _Committed:
        version: int
        seq: int

    owner = Owner()
    edited, undone = _Committed(version=2, seq=1), _Committed(version=1, seq=2)
    assert owner.serve(edited, lambda: edited, lambda target: True) is edited
    owner.release(edited)
    # A check called once the edit was undone: version 1, committed after version 2.
    laid: list = []
    served = owner.serve(undone, lambda: undone, lambda target: laid.append(target) or True)
    assert served is undone and laid == [undone]
    owner.release(undone)
    assert owner.serve(edited, lambda: undone, lambda target: True) is undone     # committed later


@pytest.mark.parametrize("workers", [0, 2])
def test_a_check_after_an_undo_reads_the_undone_document(made, workers):
    """A check called after an undo reads the document as the undo left it -- never the layout
    of the version before the undo, though that one's number is higher."""
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=workers) as box:
        session, doc, document = _session(box, made)
        heads = _heads(document)
        assert box.dispatch(session, "check", {"doc": doc}).ok
        inserted = box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{heads[1]}",
                                                                  "markdown": "A paragraph the undo takes out."})
        assert inserted.ok
        edited = box.dispatch(session, "check", {"doc": doc, "include": INCLUDE})
        assert edited.ok and edited.version == inserted.version
        assert box.dispatch(session, "undo", {"doc": doc}).ok
        now = session.entry(doc).version
        assert now < edited.version
        after = box.dispatch(session, "check", {"doc": doc, "include": INCLUDE})
        assert after.ok and after.version == now, (after.version, now)
        assert not [w for w in after.warnings if "newer than at the call" in w]
        data = session.entry(doc).document.to_bytes()
    alone = dataclasses.replace(FORMAT, read_version=None)
    with Toolbox(TOOLS, formats=[alone], groups=GROUPS, workers=0) as box:
        other = box.session()
        copy_doc = other.open(data, "report.docx")
        assert box.dispatch(other, "ds_put", {"id": "values"}).ok
        reference = box.dispatch(other, "check", {"doc": copy_doc, "include": INCLUDE})
    mine, theirs = _same(after), _same(reference)
    for result in (mine, theirs):
        result["data"]["reflow"].pop("since_last_check", None)
    assert mine == theirs


def test_relayout_on_commit_lays_out_the_new_version_while_checks_ask(made, monkeypatch):
    # As for a long document: an edit's own checks say "stale" and lay nothing out.
    monkeypatch.setattr(base_module, "REFLOW_PAGES", 0)
    monkeypatch.setattr(owner_module, "ASKED_SECONDS", 60.0)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, document = _session(box, made)
        heads = _heads(document)
        assert box.dispatch(session, "check", {"doc": doc}).ok
        owner = session.entry(doc).document._layout_owner
        flights = owner.flights
        assert box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{heads[1]}",
                                                              "markdown": "Laid out at once."}).ok
        deadline = time.monotonic() + 60
        while owner.flights == flights and time.monotonic() < deadline:
            time.sleep(0.05)
        assert owner.flights == flights + 1 and owner.done.version == session.entry(doc).version
        result = box.dispatch(session, "check", {"doc": doc})
        assert result.ok and result.version == owner.done.version
        assert owner.flights == flights + 1                   # nothing laid out again
        # Once per check asking: of two edits after it, the first is laid out at once.
        for k in range(2):
            assert box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{heads[2]}",
                                                                  "markdown": f"Edit {k}."}).ok
            deadline = time.monotonic() + 60
            while owner.flying is not None or (k == 0 and owner.flights == flights + 1
                                               and time.monotonic() < deadline):
                time.sleep(0.05)
        time.sleep(0.5)
        assert owner.flights == flights + 2 and owner.done.version == session.entry(doc).version - 1
    monkeypatch.setattr(owner_module, "ASKED_SECONDS", 0.0)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, document = _session(box, made)
        assert box.dispatch(session, "check", {"doc": doc}).ok
        owner = session.entry(doc).document._layout_owner
        flights = owner.flights
        assert box.dispatch(session, "word_insert_markdown", {"doc": doc, "at": f"before:{_heads(document)[1]}",
                                                              "markdown": "Not laid out yet."}).ok
        time.sleep(0.5)
        assert owner.flights == flights                     # off: the next check lays it out


# -- facts made once per layout ----------------------------------------------------------------


def test_layout_facts_made_once_per_layout_are_the_facts_made_afresh(made):
    document = Document.open(made)
    layout = document.layout()
    heads = _heads(document)
    for span in ([f"{heads[0]}..{heads[1]}"], [f"{heads[2]}..{heads[3]}"], None):
        scope = Scope.of_ids(document, Scope(document, ranges_of(span)).spec()) if span else None
        on = range_pages(scope, layout, extent=True)[1] if scope is not None else None
        kept = layout_facts(layout, pages=on, scope=scope)
        again = layout_facts(layout, pages=on, scope=scope)
        fresh_document = Document.open(made)
        fresh = fresh_document.layout()
        fresh_scope = (Scope.of_ids(fresh_document, Scope(fresh_document, ranges_of(span)).spec())
                       if span else None)
        assert kept == again == layout_facts(fresh, pages=on, scope=fresh_scope)
