"""What a process keeps of its layouts, by memory (round 39): ``docx_agent.layout.LAYOUT_MEMORY``.

Two layouts a worker kept a document's next check from laying it out whole only while no more
than two documents were built per worker; with six built at once on two workers, every check
laid its document out afresh.  The layouts held (``_HELD``), the pickled ones kept
(``_RECENT``) and the measurements by lineage (``_MEASURES``) are bounded by a byte budget
instead: a page's worth (``PAGE_COST``) per page held, the least recently used dropped first,
the newest kept whatever its size.
"""

from __future__ import annotations

from types import SimpleNamespace

from docx_agent import layout as L


def _laid(pages: int) -> tuple:
    return (SimpleNamespace(pages=[object()] * pages), None, [])


def _put(token: str, pages: int, lineage: str) -> None:
    with L._HELD_LOCK:
        L._HELD_TAKEN[lineage] = L._HELD_TAKEN.get(lineage, 0) + 1
    L._put_held((token,), L._Held(_laid(pages)), lineage)


def test_held_layouts_are_kept_within_the_budget_least_recently_used_out(monkeypatch):
    monkeypatch.setattr(L, "_HELD", type(L._HELD)())
    monkeypatch.setattr(L, "_HELD_TAKEN", {})
    monkeypatch.setattr(L, "LAYOUT_MEMORY", 30 * L.PAGE_COST)
    for k in range(3):                         # three ten-page documents fill the budget
        _put(f"doc {k} v1", 10, f"lineage {k}")
    assert [token[0] for token in L._HELD] == ["doc 0 v1", "doc 1 v1", "doc 2 v1"]
    # Document 0's next version: taken out and put back, it is the most recently used.
    found, how = L._take_held(("doc 0 v2",), ("doc 0 v1",), "lineage 0")
    assert how == "kept" and found is not None
    L._put_held(("doc 0 v2",), found, "lineage 0")
    _put("doc 3 v1", 10, "lineage 3")          # one more: document 1, used longest ago, goes
    assert [token[0] for token in L._HELD] == ["doc 2 v1", "doc 0 v2", "doc 3 v1"]
    _put("doc 4 v1", 25, "lineage 4")          # a long one: room made for it
    assert [token[0] for token in L._HELD] == ["doc 4 v1"]
    _put("doc 5 v1", 40, "lineage 5")          # one over the budget alone is still kept
    assert [token[0] for token in L._HELD] == ["doc 5 v1"]
    assert not L._HELD_TAKEN


def test_many_small_documents_are_held_where_two_were(monkeypatch):
    monkeypatch.setattr(L, "_HELD", type(L._HELD)())
    monkeypatch.setattr(L, "_HELD_TAKEN", {})
    monkeypatch.setattr(L, "LAYOUT_MEMORY", 256 * 2 ** 20)
    for k in range(30):
        _put(f"doc {k}", 10, f"lineage {k}")
    assert len(L._HELD) == 25                  # 25 ten-page layouts in 256 MB


def test_pickled_layouts_are_kept_within_a_quarter_of_the_budget(monkeypatch):
    import pickle

    monkeypatch.setattr(L, "_RECENT", {})
    monkeypatch.setattr(L, "LAYOUT_MEMORY", 40 * L.PAGE_COST)
    for k in range(4):
        L._remember(pickle.dumps(k), _laid(4))  # 4 pages each: two in a quarter of 40
    assert len(L._RECENT) == 2
    L._remember(pickle.dumps("long"), _laid(30))
    assert len(L._RECENT) == 1


def test_the_budget_comes_from_the_environment(monkeypatch):
    monkeypatch.setenv("DOCX_AGENT_LAYOUT_MEMORY", "512")
    assert L._memory_budget() == 512 * 2 ** 20
    for value in ("", "none", "0", "-3"):
        monkeypatch.setenv("DOCX_AGENT_LAYOUT_MEMORY", value)
        assert L._memory_budget() == 256 * 2 ** 20
    assert L._MEASURES.entries == L.LAYOUT_MEMORY // 2 // L.MEASURE_COST
