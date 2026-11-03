"""What a batch of Word ops costs does not grow with the document (round 28, W8), and its
results are the ones the ops gave before: the same with and without what makes it so.

What each op of a batch repeated over the whole document, and what it does now:

* the layout ``keep_together`` reads for its warning (a table taller than a page): laid
  out once, after the batch's last op (``Call.defer``), not after every op;
* ``word_edit_table``'s table as Markdown: the table read alone (``Reader.block``), the
  list items before it counted, not the whole body read for it;
* the effective formatting's paragraphs: found as asked (``_Model._found``), not every
  paragraph's path resolved from the root after every edit;
* the id index, the drawings without ids, the ids in use and the section breaks: found by
  lxml's walk (or XPath).

The golden-style test runs one mixed batch -- formatting, cells set, rows, keeps that warn,
text, Markdown, a failing op -- and each of its ops alone, with all of that switched off and
on, and compares every result and the document's bytes."""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Limits, Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.edit import effective as _effective  # noqa: E402
from docx_agent.edit import ids as _ids  # noqa: E402
from docx_agent.edit.numbering import membership  # noqa: E402
from docx_agent.edit.sections import W_BODY, W_PPR, W_SECTPR, in_body, section_id  # noqa: E402
from docx_agent.markdown.read import Reader  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

from conftest import fixture_id, fixture_paths  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 11, 3, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731


def report() -> bytes:
    """Three tables: the first with a list item in a cell that counts on from the body's
    list, followed by a paragraph and a table (a keep with the next warns); the third taller
    than a page."""
    doc = Document.new()
    rows = "\n".join(f"| Line {k} | {k * 3} |" for k in range(70))
    doc.insert_markdown("\n\n".join([
        "# Report title", "An opening paragraph.", "1. First step\n2. Second step",
        "| Name | Count |\n|---|---|\n| North | 4 |\n| South | 5 |",
        "A paragraph between two tables.",
        "| Kind | Share |\n|---|---|\n| Paper | 40 |\n| Glass | 60 |",
        "# Chapter title", "Some words before the long table.",
        "| Line | Value |\n|---|---|\n" + rows,
        "A closing paragraph.",
    ]))
    listed = next(p for p in doc.paragraphs() if membership(doc, doc._resolve(p.id)[1].element) is not None)
    number = membership(doc, doc._resolve(listed.id)[1].element).num_id
    doc.add_to_list(doc.tables()[0].cell(1, 0).paragraphs[0].id, number)
    return doc.to_bytes()


def mixed(doc: str, document: Document) -> list[dict]:
    first, second, long = [t.id for t in document.tables()]
    texts = [p.id for p in document.paragraphs() if p.text.startswith(("An opening", "A closing", "Some words"))]
    def op(tool, **arguments):
        return {"tool": tool, "arguments": {"doc": doc, **arguments}}
    return [
        op("word_format_table", table=first, scope="table", borders="single"),
        op("word_edit_table", table=first, action="set_cells", cells=[{"row": 2, "col": 1, "text": "6"}]),
        op("word_format_table", table=first, scope="table", keep_together=True, keep_with_next=True),
        op("word_format_table", table=long, scope="table", keep_together=True),
        op("word_set_text", target=texts[0], text="A new opening paragraph."),
        op("word_format", targets=[texts[1]], bold=True),
        op("word_edit_table", table=second, action="insert_row"),
        op("word_edit_table", table="t:7FFF0000", action="insert_row"),
        op("word_insert_markdown", at=f"after:{texts[2]}", markdown="New words.\n\n- one\n- two"),
        op("word_format_table", table=second, scope="row", row=0, shading="D9E2F3"),
        op("word_edit_table", table=long, action="set_cells",
           cells=[{"row_label": "Line 5", "col_label": "Value", "text": "fifteen"}]),
        op("word_format_table", table=long, scope="table", header_rows=1),
        op("word_edit_table", table=first, action="set_cells", cells=[{"row": 1, "col": 1, "text": "7"}]),
    ]


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _plain(result) -> dict:
    data = result.to_json()
    data.pop("duration", None)
    return {"result": data, "touched": result.touched,
            "touched_by_op": {str(k): v for k, v in result.touched_by_op.items()}}


def _run(box, data: bytes) -> dict:
    session = box.session(clock=CLOCK)
    doc = session.open(data, "report.docx")
    document = session.entry(doc).document
    document.layout()                  # laid out before: the checks' reflow compares with it
    ops = mixed(doc, document)
    out = {"batch": _plain(box.dispatch(session, "batch", {"ops": ops})),
           "bytes": document.to_bytes().hex()}
    alone = []
    for op in ops:
        single = box.session(clock=CLOCK)
        single.open(data, "report.docx")
        single.entry(doc).document.layout()
        alone.append(_plain(box.dispatch(single, op["tool"], op["arguments"])))
    out["alone"] = alone
    return out


def _unoptimised(monkeypatch, box) -> None:
    monkeypatch.setattr(box, "batch_defers", False)
    monkeypatch.setattr(Reader, "block", lambda self, part, identifier: None)

    def every(self, element):
        self._map_all()
        return self._paragraphs.get(id(element))

    monkeypatch.setattr(_effective._Model, "_found", every)


def test_a_mixed_batch_gives_the_results_it_gave_before(box, monkeypatch):
    data = report()
    read_alone = []
    block = Reader.block

    def counted(self, part, identifier):
        found = block(self, part, identifier)
        read_alone.append(found is not None)
        return found

    with monkeypatch.context() as patch:
        patch.setattr(Reader, "block", counted)
        fast = _run(box, data)
    with monkeypatch.context() as patch:
        _unoptimised(patch, box)
        slow = _run(box, data)
    assert any(read_alone)                         # the tables were read alone
    assert json.dumps(fast, sort_keys=True, default=str) == json.dumps(slow, sort_keys=True, default=str)
    batch = fast["batch"]["result"]
    warnings = batch["warnings"]
    taller = [w for w in warnings if "taller than a page" in w]
    loose = [w for w in warnings if "does not keep with the next" in w]
    assert taller and taller[0].startswith("ops[3]:") and loose and loose[0].startswith("ops[2]:")
    assert batch["data"]["failed"] == [7] and batch["checks"]["reflow"] is not None
    first_table = batch["data"]["ops"][1]["data"]["markdown"]
    assert '<ol start="3">' in first_table         # counted on from the body's list
    assert fast["alone"][3]["result"]["warnings"] == [taller[0].removeprefix("ops[3]: ")]


def test_keep_together_lays_the_document_out_once_per_batch(box, monkeypatch):
    from docx_agent.tools import _base

    data = report()
    session = box.session(clock=CLOCK)
    doc = session.open(data, "report.docx")
    tables = [t.id for t in session.entry(doc).document.tables()]
    ops = [{"tool": "word_format_table", "arguments": {"doc": doc, "table": t, "scope": "table",
                                                       "keep_together": True}} for t in tables * 2]
    laid = []
    run = _base.Engine.__call__

    def counting(self, data, options):
        laid.append(1)
        return run(self, data, options)

    monkeypatch.setattr(_base.Engine, "__call__", counting)
    result = box.dispatch(session, "batch", {"ops": ops})
    assert result.ok and sum("taller than a page" in w for w in result.warnings) == 2
    assert len(laid) == 1                          # after the last op; each op's own would be six


def test_a_batch_past_its_deadline_says_what_was_not_run(box):
    data = report()
    session = box.session(clock=CLOCK, limits=Limits(batch_timeout=0.0))
    doc = session.open(data, "report.docx")
    tables = [t.id for t in session.entry(doc).document.tables()]
    ops = [{"tool": "word_format_table", "arguments": {"doc": doc, "table": t, "scope": "table",
                                                       "borders": "single"}} for t in tables]
    result = box.dispatch(session, "batch", {"ops": ops})
    assert not result.ok and result.error.code == "timeout"
    assert result.error.details["not_run"] == {"count": 3, "first": 0, "last": 2, "ran": 0,
                                               "seconds": result.error.details["not_run"]["seconds"]}
    assert "Not run: 3 op(s), ops[0]..ops[2]" in result.error.message
    assert "resend them as smaller batches" in result.error.message
    assert session.entry(doc).document.to_bytes() == Document.open(data).to_bytes()


# -- the walks lxml does now find what the Python walks found -------------------------------------


def _walk(root, tags):
    """The Python walk ``live_elements`` was: ``tags`` in document order, no ``mc:Fallback``."""
    out, stack = [], [iter(root)]
    while stack:
        for node in stack[-1]:
            tag = node.tag
            if not isinstance(tag, str) or tag == _ids.MC_FALLBACK:
                continue
            if tag in tags:
                out.append(node)
            if len(node):
                stack.append(iter(node))
                break
        else:
            stack.pop()
    return out


def _records(document):
    """``_section_records`` as it was: over the index's paragraphs."""
    part = document.package.document_part()
    out = []
    for entry in document._index(part).paragraphs:
        properties = entry.element.find(W_PPR)
        sect = properties.find(W_SECTPR) if properties is not None else None
        if sect is not None and in_body(entry.element):
            out.append((section_id(entry.id), sect, entry.element))
    last = document.package.tree(part).find(W_BODY).find(W_SECTPR)
    return out + ([("s:body", last, None)] if last is not None else [])


def _section_of(document, identifier):
    """``section_of`` as it was: by every element's place in the tree."""
    part, entry = document._resolve(identifier)
    nodes = list(document.package.tree(part).iter())  # held: one proxy per element, stable id()s
    order = {id(node): k for k, node in enumerate(nodes)}
    records = _records(document)
    for found, _, paragraph in records:
        if paragraph is None or order[id(paragraph)] >= order[id(entry.element)]:
            return found
    return records[-1][0]


TAGS = [frozenset({_ids.W_P, _ids.W_TR, _ids.W_TBL}), frozenset({_ids.WP_INLINE, _ids.WP_ANCHOR}),
        frozenset({_ids._W + "sdt"}), frozenset({_ids._W + "ins", _ids._W + "del"})]


@pytest.mark.parametrize("path", fixture_paths(), ids=fixture_id)
def test_the_walks_find_what_they_found(path: Path):
    document = Document.open(path.read_bytes())
    for part in document._parts():
        root = document.package.tree(part)
        for node in [root, *list(root.iter(_ids.W_TBL, _ids.MC_FALLBACK))[:20]]:
            for tags in TAGS:
                assert _ids.live_elements(node, tags) == _walk(node, tags)
        used = set()
        for node in root.iter(_ids.W_P, _ids.W_TR, _ids.WP_INLINE, _ids.WP_ANCHOR):
            for name in (_ids.PARA_ID, _ids.TEXT_ID, _ids.ANCHOR_ID, _ids.EDIT_ID):
                raw = node.get(name)
                if raw is not None and _ids._HEX.fullmatch(raw):
                    used.add(int(raw, 16))
        assert _ids.used_long_hex([root]) == used
    assert document._section_records() == _records(document)
    body = document._index(document.package.document_part())
    for identifier in [e.id for e in body.paragraphs] + [t.id for t in body.tables]:
        assert document.section_of(identifier).id == _section_of(document, identifier)
    eager, lazy = _effective._Model(document), _effective._Model(document)
    eager._map_all()
    for part in document._parts():
        for element in document.package.tree(part).iter(_ids.W_P):
            want, got = eager._paragraphs.get(id(element)), lazy._found(element)
            assert (want and want.path) == (got and got.path)
            if lazy._mapped:
                lazy = _effective._Model(document)
