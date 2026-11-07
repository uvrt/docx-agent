"""Round 35: ``check`` and ``render`` scoped to some blocks (R18), and copies compared with the
reference they were made from (R19).

R18: ``check(range=[...])`` reports the facts about those blocks -- the layout's, on the pages
they fall on; ``validate``'s problems found in them -- and each range's pages;
``render(range=[...])`` draws those pages.  R19: ``check(copies=[...])`` says per copy whether
it follows its reference, what differs (with its fixes and kind), and where the reference's
example text is still in it; each comparison is kept by what both sides hold, so an edit
elsewhere finds it kept.

The shared definitions of ``check`` and ``render`` are ooxml-edit's (``range`` and
``copies`` since 0.81.0, the reference's document locked with the copy's).  Every document
here is made up.
"""

from __future__ import annotations

import threading
import time

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.edit import copies as copies_module  # noqa: E402
from docx_agent.scope import Scope, ranges_of  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS, scoped  # noqa: E402
from docx_agent.validate import Problem, check as validate_package, problems_in  # noqa: E402

pytestmark = pytest.mark.any_faces

@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, arguments, result.error and result.error.to_json())
    return result


# -- a report of chapters ------------------------------------------------------------------------

WORDS = "lantern meadow river stone harbour quiet orchard copper valley signal".split()


def _report() -> tuple[bytes, list[list[str]]]:
    """Three chapters, each a heading and paragraphs over more than a page; in each, one
    paragraph made bold by hand (a direct-formatting fact).  The blocks of each chapter."""
    doc = Document.new()
    parts = []
    for c in range(1, 4):
        parts.append(f"# Chapter {c}")
        for p in range(6 if c != 3 else 26):
            parts.append(" ".join(WORDS[(c + p + k) % len(WORDS)] for k in range(70)).capitalize() + ".")
    doc.insert_markdown("\n\n".join(parts))
    blocks = [b["id"] for b in doc.blocks()]
    heads = [k for k, b in enumerate(doc.blocks()) if b["level"] is not None]
    chapters = [blocks[a:b] for a, b in zip(heads, heads[1:] + [len(blocks)])]
    for chapter in chapters:
        doc.paragraph(chapter[2]).format(bold=True)
    return doc.to_bytes(), chapters


@pytest.fixture(scope="module")
def made():
    return _report()


def _span(blocks: list[str]) -> str:
    return f"{blocks[0]}..{blocks[-1]}"


# -- R18: check(range) -----------------------------------------------------------------------------


def test_a_range_reports_its_blocks_facts_and_pages(box, made):
    data, chapters = made
    session = box.session()
    d1 = session.open(data, "report.docx")
    whole = _ok(box, session, "check", doc=d1, include=["layout"]).data["layout"]
    flagged = {e["para"] for e in whole["direct_formatting"]}
    assert {chapter[2] for chapter in chapters} <= flagged          # one bold paragraph per chapter
    result = _ok(box, session, "check", doc=d1, range=[_span(chapters[1])])
    scoped_facts = result.data["layout"]
    assert [e["para"] for e in scoped_facts["direct_formatting"]] == [chapters[1][2]]
    layout = session.entry(d1).document.layout()
    pages = sorted({p for b in chapters[1] for p in layout.pages_of(b)})
    reported = dict(result.data["range"][0])
    reported.pop("extent"), reported.pop("height")          # R26's (tests/test_range_extent.py)
    for key in ("one_page", "kept", "breaks", "fix"):         # round 36, R30's (tests/test_keep_spans.py)
        reported.pop(key, None)
    for key in ("if_kept", "if_released", "effect_left_out"):  # round 40's (tests/test_keep_trials.py)
        reported.pop(key, None)
    assert [reported] == [{"range": _span(chapters[1]), "blocks": len(chapters[1]), "pages": pages}]
    assert "in 1 range(s)" in result.summary
    assert "layout" in result.data and "reflow" in result.data          # range adds layout to the default


def test_several_ranges_in_one_check(box, made):
    data, chapters = made
    session = box.session()
    d1 = session.open(data, "report.docx")
    result = _ok(box, session, "check", doc=d1, include=["layout"], range=[_span(chapters[0]), chapters[2][2]])
    assert [r["range"] for r in result.data["range"]] == [_span(chapters[0]), chapters[2][2]]
    assert result.data["range"][1]["blocks"] == 1
    paragraphs = [e["para"] for e in result.data["layout"]["direct_formatting"]]
    assert paragraphs == [chapters[0][2], chapters[2][2]]


def test_a_range_with_pages_keeps_the_facts_on_those_pages(box, made):
    data, chapters = made
    session = box.session()
    d1 = session.open(data, "report.docx")
    on = _ok(box, session, "check", doc=d1, include=["layout"], range=[_span(chapters[2])]).data["range"][0]["pages"]
    assert len(on) > 4                                                   # the long chapter
    elsewhere = [p for p in range(1, on[0])]
    result = box.dispatch(session, "check", {"doc": d1, "include": ["layout"], "range": [_span(chapters[2])],
                                             "pages": elsewhere})
    assert result.ok
    assert result.data["layout"]["direct_formatting"] == []
    assert any("none of the range's blocks" in w for w in result.warnings)


def test_validate_problems_are_found_in_their_blocks(box, made):
    data, chapters = made
    doc = Document.open(data)
    main = doc.package.document_part()
    element = doc._resolve(chapters[0][1])[1].element
    properties = element.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    if properties is None:
        from lxml import etree

        properties = etree.SubElement(element, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
        element.insert(0, properties)
    from lxml import etree

    style = etree.SubElement(properties, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
    style.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "NoSuchStyle")
    properties.insert(0, style)
    doc.package.mark_dirty(main)
    session = box.session()
    d1 = session.open(doc.to_bytes(), "report.docx")
    first = _ok(box, session, "check", doc=d1, include=["validate"], range=[_span(chapters[0])]).data["validate"]
    assert [p for p in first["problems"] if "style-missing" in p] == ["[style-missing] word/document.xml: pStyle NoSuchStyle"]
    assert first["elsewhere"] == 0
    second = _ok(box, session, "check", doc=d1, include=["validate"], range=[_span(chapters[1])]).data["validate"]
    assert second["problems"] == [] and second["elsewhere"] == 1 and second["baseline"] == 1


def test_problems_in_blocks_by_what_they_carry():
    """The locator: problems a block's elements carry, or that its XML alone shows again."""
    doc = Document.new()
    doc.insert_markdown("First paragraph.\n\nSecond paragraph.\n\n| A | B |\n|---|---|\n| 1 | 2 |\n")
    main = doc.package.document_part()
    tree = doc.package.tree(main)
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    paragraphs = tree.findall(f".//{w}body/{w}p")
    table = tree.find(f".//{w}tbl")
    from lxml import etree

    mark = etree.SubElement(paragraphs[0], w + "bookmarkStart")
    mark.set(w + "id", "7")
    mark.set(w + "name", "Lonely")
    cell = table.find(f".//{w}tc")
    for p in cell.findall(w + "p"):
        cell.remove(p)
    etree.SubElement(cell, w + "tbl")                     # a cell ending in a table
    doc.package.mark_dirty(main)
    problems = validate_package(doc.package)
    codes = {p.code for p in problems}
    assert {"bookmark-unpaired", "cell-not-ending-in-paragraph"} <= codes
    first = {p.code for p in problems_in(problems, main, [paragraphs[0]])}
    assert "bookmark-unpaired" in first and "cell-not-ending-in-paragraph" not in first
    second = {p.code for p in problems_in(problems, main, [paragraphs[1]])}
    assert not second & {"bookmark-unpaired", "cell-not-ending-in-paragraph"}
    in_table = {p.code for p in problems_in(problems, main, [table])}
    assert "cell-not-ending-in-paragraph" in in_table
    # A problem of another part, or of the package, is in no block.
    assert problems_in([Problem("orphaned-part", "word/media/x.png")], main, [paragraphs[0]]) == []


def test_a_range_that_names_nothing_is_refused_with_the_headings(box, made):
    """Rule 2: the way out -- the sections to name, as heading ranges (h:)."""
    data, chapters = made
    session = box.session()
    d1 = session.open(data, "report.docx")
    result = box.dispatch(session, "check", {"doc": d1, "range": ["p:DEADBEEF"]})
    assert not result.ok
    error = result.error.to_json()
    assert error["field"] == "range" and "range[0]" in error["message"]
    assert error["valid_options"][0].startswith("h:" + chapters[0][0])     # round 36: a heading's range
    assert any("Chapter 2" in option for option in error["valid_options"])


def test_one_range_said_as_a_string_and_placeholders():
    """Rule 1: a lone string is one range; blanks and repeats are dropped."""
    assert ranges_of("p:A..p:B") == ["p:A..p:B"]
    assert ranges_of(["", " p:A ", "p:A"]) == ["p:A"]
    assert ranges_of([]) == [] and ranges_of(None) == []


def test_an_empty_range_is_no_range(box, made):
    data, _ = made
    session = box.session()
    d1 = session.open(data, "report.docx")
    result = _ok(box, session, "check", doc=d1, include=["validate"], range=[])
    assert "range" not in result.data and "elsewhere" not in result.data["validate"]


def test_a_cell_range_is_about_the_cell_and_its_table(made):
    doc = Document.new()
    doc.insert_markdown("Before.\n\n| A | B |\n|---|---|\n| one | two |\n\nAfter.\n")
    table = doc.tables()[0]
    scope = Scope(doc, [f"{table.id}/c1,1"])
    assert table.id in scope.tables
    assert scope.about({"table": table.id, "row": 1})
    before = doc.paragraphs()[0].id
    assert not scope.about({"para": before})
    assert scope.about({"page": 1, "placeholders": []})                     # a page's, kept on its pages


# -- R18: render(range) ----------------------------------------------------------------------------


def test_render_draws_the_pages_a_range_falls_on(box, made):
    pytest.importorskip("resvg_py")
    data, chapters = made
    session = box.session()
    d1 = session.open(data, "report.docx")
    result = _ok(box, session, "render", doc=d1, range=[chapters[1][1]], width=300)
    layout = session.entry(d1).document.layout()
    assert result.data["pages"] == layout.pages_of(chapters[1][1])
    assert result.data["range"][0]["pages"] == result.data["pages"]
    # More pages than a call's images: one sheet, said.
    long = _ok(box, session, "render", doc=d1, range=[_span(chapters[2])], width=300)
    assert "sheet" in long.data and long.data["pages"] == long.data["range"][0]["pages"]
    assert any("drawn as one sheet" in w for w in long.warnings)
    refused = box.dispatch(session, "render", {"doc": d1, "range": [chapters[1][1]], "pages": [99]})
    assert not refused.ok and refused.error.to_json()["valid_options"] == result.data["pages"]


# -- R18: no wait for a layout under the lock ------------------------------------------------------

DEADLINE = 90.0


def test_scoped_checks_lone_and_in_a_read_batch_all_finish(made, monkeypatch):
    """Round 34's rule holds for a range: lone scoped checks lay out with the lock let go, a
    read_batch's scoped check on its own snapshot; none waits for the other."""
    from docx_agent.tools import _base

    data, chapters = made
    started = threading.Event()
    run = _base.Engine.lay_out
    facts = _base.Engine.check

    def slow(engine, data, options, state=None):
        started.set()
        time.sleep(0.4)
        return run(engine, data, options, state)

    def slow_facts(engine, data, options, state, request):     # a lone check's (round 38)
        started.set()
        time.sleep(0.4)
        return facts(engine, data, options, state, request)

    monkeypatch.setattr(_base.Engine, "lay_out", slow)
    monkeypatch.setattr(_base.Engine, "check", slow_facts)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session = toolbox.session()
        d1 = session.open(data, "report.docx")
        scoped_check = {"doc": d1, "range": [_span(chapters[1])]}
        results: list = []
        lock = threading.Lock()

        def call(name, arguments):
            found = toolbox.dispatch(session, name, arguments)
            with lock:
                results.append(found)

        threads = [threading.Thread(target=call, args=("check", scoped_check), daemon=True) for _ in range(2)]
        for thread in threads:
            thread.start()
        assert started.wait(DEADLINE)
        batch = {"ops": [{"tool": "check", "arguments": scoped_check},
                         {"tool": "word_inspect", "arguments": {"doc": d1, "target": chapters[1][1],
                                                                "layout": True}}]}
        later = [threading.Thread(target=call, args=("read_batch", batch), daemon=True)]
        for thread in later:
            thread.start()
        ends = time.monotonic() + DEADLINE
        for thread in threads + later:
            thread.join(max(0.0, ends - time.monotonic()))
        assert not any(t.is_alive() for t in threads + later), "deadlock"
    assert all(r.ok for r in results), [r.to_json() for r in results if not r.ok]


# -- R19: copies against their reference -----------------------------------------------------------

SAMPLE = ("# Part title\n\n"
          "Describe the purpose of this part here, in a few sentences. Then name who owns it.\n\n"
          "| Item | Owner | Status |\n| --- | --- | --- |\n| Write the first item here | Someone | Open |\n\n"
          "Close with a short summary of the part.\n")


def _samples() -> tuple[bytes, str]:
    sample = Document.new()
    sample.insert_markdown(SAMPLE, at="end")
    sample.set_table(sample.tables()[0].id, header_rows=1)
    blocks = [b["id"] for b in sample.blocks()]
    return sample.to_bytes(), _span(blocks)


def _copies(box, session, count: int = 6) -> tuple[str, str, str, list[str]]:
    """A samples document and a report holding ``count`` copies of its one sample."""
    data, reference = _samples()
    samples = session.open(data, "samples.docx")
    report = Document.new()
    report.insert_markdown("# Report\n\nAn opening paragraph.", at="end")
    d1 = session.open(report.to_bytes(), "report.docx")
    made = []
    for _ in range(count):
        blocks = _ok(box, session, "word_copy_from", doc=d1, source_doc=samples, range=reference,
                     at="end").data["blocks"]
        made.append(_span(blocks))
    return d1, samples, reference, made


def _pairs(made: list[str], samples: str, reference: str) -> list[dict]:
    return [{"blocks": c, "reference": {"doc": samples, "block": reference}} for c in made]


def _blocks(session, doc: str, span: str) -> list[str]:
    document = session.entry(doc).document
    return [identifier for identifier, _ in copies_module.blocks_of(document, span)]


def test_copies_follow_or_not_with_fixes_and_example_text(box):
    session = box.session()
    d1, samples, reference, made = _copies(box, session)
    # Copy 0 filled in full; copy 1 filled, its table's first body row shaded by hand; the
    # rest left with the sample's text.
    for k in (0, 1):
        blocks = _blocks(session, d1, made[k])
        _ok(box, session, "word_set_text", doc=d1, target=blocks[1], text=f"Part {k} keeps the books in order.")
        _ok(box, session, "word_set_text", doc=d1, target=blocks[3], text=f"Part {k} is done.")
        _ok(box, session, "word_set_text", doc=d1, target=f"{blocks[2]}/c1,0", text=f"Ledger {k}")
    second_table = _blocks(session, d1, made[1])[2]
    _ok(box, session, "word_format_table", doc=d1, table=second_table, scope="row", row=1, shading="F2F2F2")
    result = _ok(box, session, "check", doc=d1, include=[], copies=_pairs(made, samples, reference))
    found = result.data["copies"]
    items = found["items"]
    assert found["compared"] == 6 and found["follow"] == 5
    assert [item["follows"] for item in items] == [True, False, True, True, True, True]
    assert "example_text" not in items[0] and "example_text" not in items[1]
    assert found["with_example_text"] == 4
    left = items[2]["example_text"]
    assert [e["text"] for e in left][0] == ["describe the purpose of this part here, in a few sentences.",
                                            "then name who owns it."]
    assert all("part title" not in t for e in left for t in e["text"])     # a two-word heading: a label
    assert all("item" != t for e in left for t in e["text"])                # the header row: kept by design
    differences = items[1]["differences"]
    assert differences and all(d["kind"] == "format" for d in differences)
    assert {d["property"] for d in differences} == {"shading"}
    # Rule 5: one call makes every copy follow its reference.
    assert found["fix"] == {"tool": "word_follow_sample",
                            "calls": [{"sample": {"block": _blocks(session, samples, reference)[2],
                                                  "doc": samples}, "copies": [second_table]}]}
    for call in found["fix"]["calls"]:
        _ok(box, session, "word_follow_sample", doc=d1, **call)
    again = _ok(box, session, "check", doc=d1, include=[], copies=_pairs(made, samples, reference))
    assert again.data["copies"]["follow"] == 6
    assert "copies: 6 of 6 follow their reference, 4 with its example text" in again.summary


def test_a_comparison_is_kept_while_neither_side_changes(box):
    session = box.session()
    d1, samples, reference, made = _copies(box, session)
    pairs = _pairs(made, samples, reference)
    first = _ok(box, session, "check", doc=d1, include=[], copies=pairs).data["copies"]
    assert first["kept"] == 0
    # An edit elsewhere in the report: every comparison kept.
    opening = session.entry(d1).document.paragraphs()[1].id
    _ok(box, session, "word_set_text", doc=d1, target=opening, text="Another opening paragraph.")
    second = _ok(box, session, "check", doc=d1, include=[], copies=pairs).data["copies"]
    assert second["kept"] == 6 and all(item["cached"] for item in second["items"])
    assert [i["follows"] for i in second["items"]] == [i["follows"] for i in first["items"]]
    # An edit inside copy 3: that one compared again.
    _ok(box, session, "word_set_text", doc=d1, target=_blocks(session, d1, made[3])[1], text="Filled in.")
    third = _ok(box, session, "check", doc=d1, include=[], copies=pairs).data["copies"]
    assert [item["cached"] for item in third["items"]] == [True, True, True, False, True, True]
    assert "example_text" in third["items"][2] and len(third["items"][3]["example_text"]) == 2
    # The reference changed: every copy compared again.
    _ok(box, session, "word_format", doc=samples, targets=[_blocks(session, samples, reference)[1]],
        alignment="center")
    fourth = _ok(box, session, "check", doc=d1, include=[], copies=pairs).data["copies"]
    assert fourth["kept"] == 0 and fourth["follow"] == 0
    assert all(any(d["property"] == "alignment" for d in item["differences"]) for item in fourth["items"])


def test_the_kept_comparisons_are_bounded(box, monkeypatch):
    monkeypatch.setattr(scoped, "COPY_CACHE", 2)
    session = box.session()
    d1, samples, reference, made = _copies(box, session, count=4)
    for k in range(4):
        _ok(box, session, "word_set_text", doc=d1, target=_blocks(session, d1, made[k])[1], text=f"Copy {k}.")
    _ok(box, session, "check", doc=d1, include=[], copies=_pairs(made, samples, reference))
    assert len(session.entry(d1).document._tool_copy_cache) == 2


def test_the_reply_keeps_to_its_budget(box, monkeypatch):
    """Rule 4: past the budget, differences are counted and the reply says how to read them."""
    monkeypatch.setattr(scoped, "DIFFERENCES_PER_REPLY", 1)
    session = box.session()
    d1, samples, reference, made = _copies(box, session, count=2)
    for span in made:
        table = _blocks(session, d1, span)[2]
        _ok(box, session, "word_format_table", doc=d1, table=table, scope="row", row=1, shading="F2F2F2")
        _ok(box, session, "word_format", doc=d1, targets=[_blocks(session, d1, span)[1]], alignment="right")
    found = _ok(box, session, "check", doc=d1, include=[], copies=_pairs(made, samples, reference)).data["copies"]
    assert sum(len(item["differences"]) for item in found["items"]) == 1
    assert sum(item.get("more", 0) for item in found["items"]) >= 3
    assert "word_compare_blocks" in found["left_out"]


def test_copies_refused_with_the_way_out(box):
    """Rule 2: a reference document not open, a copy naming nothing, a malformed item."""
    session = box.session()
    d1, samples, reference, made = _copies(box, session, count=1)
    result = box.dispatch(session, "check", {"doc": d1, "copies": [{"blocks": made[0],
                                                                    "reference": {"doc": "d9", "block": reference}}]})
    assert not result.ok
    result = box.dispatch(session, "check", {"doc": d1, "include": [], "copies": [
        {"blocks": "p:DEADBEEF", "reference": {"doc": samples, "block": reference}}]})
    assert not result.ok and result.error.to_json()["valid_options"]
    assert "copies[0].blocks" in result.error.message


def test_example_text_rules():
    """Units: sentences of three words or twenty characters, not in a header row; matched as
    whole words; a unit from an unmarked first row is marked."""
    reference = Document.new()
    reference.insert_markdown("Owner:\n\nState the aim. Keep it short.\n\n| Area name | Value |\n|---|---|\n"
                              "| Put the area here | 0 |\n", at="end")
    units = [u for u, _, _ in copies_module.example_units(reference, copies_module.blocks_of(
        reference, _span([b["id"] for b in reference.blocks()])))]
    assert "owner:" not in units and "state the aim." in units and "keep it short." in units
    assert "put the area here" in units and "area name" not in units        # two words: a label
    copy = Document.new()
    copy.insert_markdown("Owner:\n\nState the aims. Keep it short.\n\n| Area name | Value |\n|---|---|\n"
                         "| Put the area here | 0 |\n", at="end")
    found = copies_module.example_text(copy, copies_module.blocks_of(copy, _span([b["id"] for b in copy.blocks()])),
                                       reference, copies_module.blocks_of(
                                           reference, _span([b["id"] for b in reference.blocks()])))
    texts = [t for e in found for t in e["text"]]
    assert "keep it short." in texts and "state the aim." not in texts     # "aims": another word
    assert "put the area here" in texts
    assert not any(e.get("first_row") for e in found)                       # the second row
