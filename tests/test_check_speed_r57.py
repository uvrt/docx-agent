"""A range-scoped check of a long report costs what the range asks (round 57, R124).

* ``include: ["numbers"]`` names each text's page from the worker holding the layout
  (:func:`docx_agent.layout.check_facts`, ``pages_of``): the layout itself never comes to
  the caller's process for it (on a 300-page report that was seconds of the interpreter
  every call shares).
* ``numbering_skips_level`` lists the numbering part and reads each style's numbering once
  per check, not once per numbered paragraph.
* ``column_fit`` finds each table's cell margins by the table's path, not by resolving every
  body table's.
* ``data.timing`` says each include's seconds: ``numbers`` and ``validation`` (the save
  gate's, ooxml-edit 0.115.0) besides the parts it said (round 47).

Tool design rules: 4 (the timing: a few numbers, the same names whatever the document's
size; nothing else in the result changes).  Made-up documents only."""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.numbers", reason="needs ooxml-edit (numbers)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent import layout as layout_module  # noqa: E402
from docx_agent.edit import numbering as numbering_module  # noqa: E402
from docx_agent.layout_facts import numbering_skips_level  # noqa: E402
from docx_agent.markdown import write as write_module  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS as WORD_TOOLS  # noqa: E402

from test_tables_from_data import ds_put  # noqa: E402

#: With a made-up dataset to compare numbers with (test_tables_from_data's).
TOOLS = [*WORD_TOOLS, ds_put]

pytestmark = pytest.mark.any_faces

WORDS = "lantern meadow river stone harbour quiet orchard copper valley signal".split()


def _text(seed: int, words: int) -> str:
    return " ".join(WORDS[(seed + k) % len(WORDS)] for k in range(words)).capitalize() + "."


@pytest.fixture(scope="module")
def made() -> bytes:
    doc = Document.new()
    parts = ["# Made-up report"]
    for c in range(1, 4):
        parts += [f"# Part {c}", f"{_text(c, 60)} Rated {c}.5 of 5 in part {c}.",
                  "- first point\n- second point\n  - a deeper point",
                  "1. one step\n2. another step",
                  f"| Item | Amount |\n|---|---|\n| lantern | {c}2.5 |\n| meadow | {c}7.25 |",
                  _text(c + 3, 120)]
    doc.insert_markdown("\n\n".join(parts))
    return doc.to_bytes()


def _session(box, data):
    session = box.session()
    doc = session.open(data, "report.docx")
    assert box.dispatch(session, "ds_put", {"id": "values"}).ok
    return session, doc, session.entry(doc).document


@pytest.mark.parametrize("workers", [0, 2])
def test_numbers_pages_come_from_where_the_layout_is_held(made, workers, monkeypatch):
    """Each text's page as before, the layout never laid out (nor sent) to this process."""
    pulled = []
    real = layout_module.lay_out_unlocked

    def counted(*args, **kwargs):
        pulled.append(1)
        return real(*args, **kwargs)

    monkeypatch.setattr(layout_module, "lay_out_unlocked", counted)
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=workers) as box:
        session, doc, document = _session(box, made)
        result = box.dispatch(session, "check", {"doc": doc, "include": ["numbers", "validate"]})
        assert result.ok, result.to_json()
        texts = result.data["numbers"]["texts"]
        # The pages the whole layout says, laid out here (as the check did before).
        laid = Document.open(made).layout()
        assert texts and all(t.get("page") for t in texts), texts
        paragraphs = [t for t in texts if t["at"].startswith("p:")]
        assert paragraphs and all(t["page"] == laid.pages_of(t["at"])[0] for t in paragraphs), texts
        assert len({t["page"] for t in texts}) > 1                 # more than one page named
        assert pulled == []
        # A range: only its texts, their pages named the same way.
        heads = [p.id for p in document.paragraphs() if p.text.startswith("Part ")]
        scoped = box.dispatch(session, "check", {"doc": doc, "include": ["numbers"],
                                                 "range": [f"{heads[1]}..{heads[2]}"]})
        assert scoped.ok, scoped.to_json()
        found = scoped.data["numbers"]["texts"]
        assert found and len(found) < len(texts) and all(t.get("page") for t in found), found
        assert pulled == []


def test_numbers_on_pages_keep_only_the_texts_drawn_there(made):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, _document = _session(box, made)
        on_one = box.dispatch(session, "check", {"doc": doc, "include": ["numbers"], "pages": [1]})
        assert on_one.ok and on_one.data["numbers"]["texts"]
        past = box.dispatch(session, "check", {"doc": doc, "include": ["numbers"], "pages": [40]})
        assert past.ok and past.data["numbers"]["texts"] == []


def test_timing_says_numbers_and_validation(made):
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        session, doc, document = _session(box, made)
        ids = [b["id"] for b in document.blocks()]
        result = box.dispatch(session, "check", {"doc": doc, "range": [f"{ids[1]}..{ids[4]}"],
                                                 "include": ["validate", "layout", "numbers", "reflow",
                                                             "fields"]})
        assert result.ok, result.to_json()
        timing = result.data["timing"]
        assert {"validate", "layout", "facts", "fields", "numbers", "validation"} <= set(timing), timing
        assert all(isinstance(v, float) and v >= 0 for v in timing.values())
        assert sum(timing.values()) <= result.duration + 0.01
        assert list(result.data)[-1] == "timing"
        # numbers alone: its result as before (no timing added to a result that had none).
        alone = box.dispatch(session, "check", {"doc": doc, "include": ["numbers"]})
        assert alone.ok and "timing" not in alone.data


def test_numbering_skips_level_reads_each_style_once(made, monkeypatch):
    document = Document.open(made)
    calls: list = []
    real = numbering_module._style_numbering_now

    def counted(doc, style_id):
        calls.append(style_id)
        return real(doc, style_id)

    monkeypatch.setattr(numbering_module, "_style_numbering_now", counted)
    before = numbering_skips_level(document)
    assert len(calls) == len(set(calls))                 # each style once
    assert document.__dict__.get("_numbering_held") is None   # nothing kept after
    monkeypatch.setattr(numbering_module, "_style_numbering_now", real)
    # The same facts as read without the shared read (each question on its own).
    assert numbering_skips_level(Document.open(made)) == before


def test_cell_margins_by_path_are_those_found_by_element(made):
    document = Document.open(made)
    main = document.package.document_part()
    tree = document.package.tree(main)
    tables = list(tree.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"))
    assert len(tables) == 3
    for n, table in enumerate(tables, start=1):
        found = write_module._cell_margins(document, table)
        assert found
        assert write_module._cell_margins(document, table, f"w:body/w:tbl[{n}]") == found
        # A wrong hint is only a hint: the table is still found.
        assert write_module._cell_margins(document, table, "w:body/w:tbl[99]") == found
