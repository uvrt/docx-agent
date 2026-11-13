"""Round 55 (R116, R119): ``word_changes`` list gives one change per edit as it was made, and
a sample document with the shape ``accept_loses_text`` names ships with the package.
Made-up documents only.

* R116: a tracked column deletion, a merge and a cell's new text, made by one author in one
  minute, were one listed "replacement" -- every record in a table joined every structure
  record there.  Now each edit is its own entry, with its ids and what it was
  (``structure``: column deletion, column insertion, merge, split, row deletion, row
  insertion), and accepting or rejecting one leaves the others; a row or column inserted or
  deleted whole, a merge, a split stay one entry each.
* Two records sharing one ``w:id`` are listed ``rev:95`` and ``rev:95#1``; accepting several
  ids in one call resolves them all first (accepting the first had renamed the second, which
  was then refused), and an id that is not there is refused with the ids that are.
* R119: ``docx_agent.samples`` holds ``accept_loses_text.docx``; check names its two cells,
  accept moves their text after the row, reject keeps it.

Tool-design rules: 2 (a refused id comes with the listed ids and the list call), 5 (several
ids, one call, one undo step), 7 (the sample shows check's fact).
"""

from __future__ import annotations

import datetime as dt
import itertools

import pytest
from lxml import etree
from ooxml_edit.tools import Toolbox

from docx_agent import Document, samples
from docx_agent.layout_facts import accept_loses_text
from docx_agent.tools import FORMAT, GROUPS, TOOLS

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % W
ME = "Editor A"
CLOCK = lambda: dt.datetime(2026, 11, 13, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
ROWS = [("Area", "Lead", "Phase", "Remark"), ("Harbour", "Ines", "draft", "awaits figures"),
        ("Orchard", "Tomas", "done", "signed off"), ("Quarry", "Aleid", "open", "two gaps")]
MARKDOWN = "\n".join(["| " + " | ".join(ROWS[0]) + " |", "|---|---|---|---|"]
                     + ["| " + " | ".join(r) + " |" for r in ROWS[1:]])
EDITS = {
    "delete_column": dict(action="delete_column", index=2),
    "set_cells": dict(action="set_cells", cells=[{"row": 3, "col": 0, "text": "Quarry North"}]),
    "merge": dict(action="merge", first={"row": 1, "col": 0}, last={"row": 1, "col": 1}),
    "merge_down": dict(action="merge", first={"row": 1, "col": 0}, last={"row": 2, "col": 0}),
    "split": dict(action="split", first={"row": 2, "col": 1}),
    "split_rows": dict(action="split", first={"row": 2, "col": 1}, rows=2, columns=1),
    "insert_column": dict(action="insert_column", index=1),
    "insert_row": dict(action="insert_row", index=1),
    "delete_row": dict(action="delete_row", index=2),
    "first_cell": dict(action="set_cells", cells=[{"row": 1, "col": 0, "text": "Harbour East"}]),
}


@pytest.fixture(scope="module")
def box():
    with Toolbox(list(TOOLS), formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def call(box, session, name, /, **arguments):
    return box.dispatch(session, name, arguments)


def ok(box, session, name, /, **arguments):
    result = call(box, session, name, **arguments)
    assert result.ok, (name, result.to_json())
    return result


def _edited(box, *edits: str):
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "report.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", doc="d1", at=f"after:{first}",
       markdown=f"Field report.\n\n{MARKDOWN}\n\nAfter the table.")
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    table = session.entry("d1").document.tables()[-1].id
    for edit in edits:
        ok(box, session, "word_edit_table", doc="d1", table=table, **EDITS[edit])
    return session


def _listed(box, session) -> list[dict]:
    return ok(box, session, "word_changes", doc="d1", action="list").data


def _shape(listed: list[dict]) -> list[tuple]:
    return sorted((c["kind"], c.get("structure"), c.get("old_text", ""), c.get("text", "")) for c in listed)


# -- R116: one change per edit ---------------------------------------------------------------------


def test_a_column_deletion_and_a_cells_text_are_two_changes(box):
    listed = _listed(box, _edited(box, "delete_column", "set_cells"))
    assert _shape(listed) == [("deletion", "column deletion", "Phase draft done open", ""),
                              ("insertion", None, "", "North")]


@pytest.mark.parametrize("order", [("merge", "set_cells", "delete_column"), ("set_cells", "merge", "delete_column"),
                                   ("merge", "delete_column", "set_cells")])  # (a merge after it: lane C)
def test_a_merge_a_column_deletion_and_a_cells_text_are_three_changes(box, order):
    listed = _listed(box, _edited(box, *order))
    assert _shape(listed) == [("deletion", "column deletion", "Phase draft done open", ""),
                              ("insertion", None, "", "North"),
                              ("table", "merge", "", "Ines")]
    assert len({r for c in listed for r in c.get("records", [c["id"]])}) == sum(
        len(c.get("records", [c["id"]])) for c in listed)                    # no record in two


@pytest.mark.parametrize("edit, kind, structure", [
    ("delete_column", "deletion", "column deletion"), ("insert_column", "insertion", "column insertion"),
    ("merge", "table", "merge"), ("merge_down", "table", "merge"), ("split", "table", "split"),
    ("split_rows", "table", "split"), ("insert_row", "insertion", "row insertion"),
    ("delete_row", "deletion", "row deletion")])
def test_each_structure_edit_is_one_change(box, edit, kind, structure):
    listed = _listed(box, _edited(box, edit))
    assert [(c["kind"], c.get("structure")) for c in listed] == [(kind, structure)]


@pytest.mark.parametrize("edit", ["merge", "merge_down", "split", "insert_column", "delete_row"])
def test_a_cells_new_text_beside_a_structure_edit_is_its_own_change(box, edit):
    # set_cells on row 1's first cell before the edit: a cell the merges take text into.
    listed = _listed(box, _edited(box, "first_cell", edit))
    assert len(listed) == 2 and len([c for c in listed if c.get("structure")]) == 1
    text = next(c for c in listed if not c.get("structure"))
    assert "East" in text.get("text", "") and "Ines" not in text.get("text", "")


def test_each_change_is_accepted_alone_and_in_any_order_gives_accept_all(box):
    edits = ("merge", "set_cells", "delete_column")
    all_at_once = _edited(box, *edits)
    ok(box, all_at_once, "word_changes", doc="d1", action="accept", all=True)
    expected = ok(box, all_at_once, "word_read", doc="d1").data
    for k in range(len(edits)):
        order = list(itertools.permutations(range(3)))[k * 2]
        session = _edited(box, *edits)
        ids = [c["id"] for c in _listed(box, session)]
        for step, position in enumerate(order):
            result = ok(box, session, "word_changes", doc="d1", action="accept", ids=[ids[position]])
            assert result.to_json()["checks"]["validate"]["new"] == []            # a valid file each time
            assert len(_listed(box, session)) == 2 - step                    # one gone, the rest kept
        assert ok(box, session, "word_read", doc="d1").data == expected


def test_rejecting_one_change_keeps_the_others(box):
    session = _edited(box, "merge", "set_cells", "delete_column")
    listed = _listed(box, session)
    column = next(c for c in listed if c.get("structure") == "column deletion")
    ok(box, session, "word_changes", doc="d1", action="reject", ids=[column["id"]])
    left = _listed(box, session)
    assert sorted(c.get("structure") or c["kind"] for c in left) == ["insertion", "merge"]
    read = ok(box, session, "word_read", doc="d1", view="original").data
    assert "Phase" in read and "| Area | Lead | Phase | Remark |" in read


# -- two records, one w:id -----------------------------------------------------------------------------


def _shared_ids(box):
    session = box.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "notes.docx")
    first = session.entry("d1").document.paragraphs()[0].id
    ok(box, session, "word_insert_markdown", doc="d1", at=f"after:{first}", markdown="Alpha one.\n\nBravo two.")
    ok(box, session, "word_set_tracking", doc="d1", on=True, author=ME)
    ok(box, session, "replace_text", doc="d1", find="one", replace="uno")
    ok(box, session, "replace_text", doc="d1", find="two", replace="dos")
    document = session.entry("d1").document
    part = document.package.document_part()
    for record in document.package.tree(part).iter(_W + "ins", _W + "del"):
        record.set(_W + "id", "95")
    document.package.mark_dirty(part)
    again = box.session(clock=CLOCK)
    again.open(document.to_bytes(), "notes.docx")
    return again


@pytest.mark.parametrize("ids", [["rev:95", "rev:95#2"], ["rev:95#2", "rev:95"], ["rev:95#1", "rev:95#3"]])
def test_ids_sharing_one_w_id_are_accepted_together(box, ids):
    session = _shared_ids(box)
    assert [c["id"] for c in _listed(box, session)] == ["rev:95", "rev:95#2"]
    result = ok(box, session, "word_changes", doc="d1", action="accept", ids=ids)
    assert result.data["records"] == (4 if "rev:95" in ids else 2)
    read = ok(box, session, "word_read", doc="d1", view="markup").data
    assert "uno" in read and "dos" in read and "{++" not in read


def test_an_id_not_there_is_refused_with_the_ids_that_are(box):
    session = _shared_ids(box)
    before = session.entry("d1").document.to_bytes()
    refused = call(box, session, "word_changes", doc="d1", action="accept", ids=["rev:95", "rev:96"])
    assert not refused.ok
    error = refused.to_json()["error"]
    assert error["code"] == "not_found" and "'rev:96'" in error["message"]
    assert error["valid_options"] == ["rev:95", "rev:95#2"]
    assert error["details"]["call"] == {"tool": "word_changes", "arguments": {"doc": "d1", "action": "list"}}
    assert session.entry("d1").document.to_bytes() == before            # nothing accepted


# -- R119: the sample ---------------------------------------------------------------------------------


def test_the_sample_is_installed_and_has_the_shape(box):
    assert samples.path("accept_loses_text").is_file()
    data = samples.read("accept_loses_text")
    with pytest.raises(KeyError, match="accept_loses_text"):
        samples.read("nothing")
    found = accept_loses_text(Document.open(data))
    assert [(e["text"], e["change"], e["author"]) for e in found] == [("Kept. Last.", "rev:901", "Editor B"),
                                                                      ("Kept. Last.", "rev:902", "Editor B")]
    assert [e["cell"].rsplit("/", 1)[1] for e in found] == ["c0,1", "c1,1"]
    # The table ids are stable (rows carry paraIds), so the facts name them across calls.
    assert all(e["cell"].startswith("t:") for e in found)


def test_check_accept_and_reject_on_the_sample(box):
    session = box.session(clock=CLOCK)
    session.open(samples.read("accept_loses_text"), "sample.docx")
    facts = ok(box, session, "check", doc="d1", include=["layout"]).data["layout"]["accept_loses_text"]
    assert [f["fix"]["calls"][0] for f in facts] == [{"action": "reject", "ids": ["rev:901"]},
                                                     {"action": "reject", "ids": ["rev:902"]}]
    accepted = ok(box, session, "word_changes", doc="d1", action="accept", all=True)
    assert len([w for w in accepted.warnings if "now after the row" in w]) == 2
    read = ok(box, session, "word_read", doc="d1").data
    assert read.count("Kept.") == 2 and read.count("Last.") == 2
    ok(box, session, "word_undo", doc="d1")
    ok(box, session, "word_changes", doc="d1", action="reject", all=True)
    assert accept_loses_text(session.entry("d1").document) == []
    tree = etree.fromstring(etree.tostring(session.entry("d1").document.package.tree(
        session.entry("d1").document.package.document_part())))
    cells = ["".join(t.text for t in tc.iter(_W + "t")) for tc in tree.iter(_W + "tc")]
    assert cells.count("Kept.Last.") == 2


def test_the_sample_is_what_the_recipe_writes():
    import importlib.util
    from pathlib import Path

    spec = importlib.util.spec_from_file_location(
        "make_samples", Path(__file__).resolve().parents[1] / "tools" / "make_samples.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    ours, shipped = Document.open(module.accept_loses_text()), Document.open(samples.read("accept_loses_text"))
    assert ours.to_markdown(view="markup") == shipped.to_markdown(view="markup")
