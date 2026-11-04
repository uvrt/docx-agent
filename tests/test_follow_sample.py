"""Round 31, R8: ``word_follow_sample`` -- copies of a sample table (or paragraph) made to
follow the revised sample in one call, each keeping its own text and values; what could not
follow reported per copy with why.  The check is the comparison's: after following,
``compare_blocks`` finds nothing left but what the result names (and the copies' own row
counts).  Made-up documents only."""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.edit.follow import NOT_FOLLOWED_CAP, WHY, _textless  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

AUTHOR, DATE = "A. Reviewer", "2026-11-04T09:00:00Z"
SAMPLE = ("Parts.\n\n"
          "| Part | Maker | Stock |\n| --- | --- | --- |\n| Sample part | Someone | 12 |\n\n"
          "A closing line.")


def _texts(document: Document, table: str) -> list[list[str]]:
    t = document.table(table)
    return [[cell.text if cell is not None else None for cell in line] for line in t.grid()]


def _document() -> tuple[Document, str, list[str]]:
    """A sample and two copies filled with their own values (one with an extra row), then the
    sample revised: header row repeated and shaded, centred, a wider first column, a body
    cell centred vertically, the body text bold at 9 pt."""
    document = Document.new()
    document.insert_markdown(SAMPLE, at="end")
    sample = document.tables()[0].id
    copies = [document.copy_blocks(document, sample, at="end").blocks[0] for _ in range(2)]
    document.table(copies[0]).cell(1, 0).paragraphs[0].set_text("Gear wheel")
    document.insert_row(copies[0], 1)
    document.table(copies[0]).cell(2, 0).paragraphs[0].set_text("Axle")
    document.table(copies[1]).cell(1, 2).paragraphs[0].set_text("40")
    document.set_table(sample, header_rows=1, alignment="center")
    for column in range(3):
        document.set_cell(sample, 0, column, shading="1F3864")
    document.set_cell(sample, 1, 1, vertical_alignment="center")
    document.set_column_width(sample, 0, 200)
    for column in range(3):
        document.table(sample).cell(1, column).paragraphs[0].format(bold=True, size=9)
    return document, sample, copies


def _left(document: Document, sample: str, copies: list[str]) -> list[str]:
    """What a comparison still finds, less the copies' own row counts and the run formatting
    of paragraphs with no text (their marks took it, which a comparison does not read)."""
    found = (_textless(document, d) for d in document.compare_blocks(copies, sample)["differences"])
    return [d["property"] for d in found if d is not None and d["property"] != "rows"]


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, arguments, result.error and result.error.to_json())
    return result


def test_copies_follow_the_revised_sample_and_keep_their_values():
    document, sample, copies = _document()
    texts = {c: _texts(document, c) for c in copies}
    assert {"widths", "table_alignment", "shading", "vertical_alignment", "bold", "size"} \
        <= set(_left(document, sample, copies))
    result, report = document.follow_sample(sample, copies)
    assert result.changed and result.count == 2
    assert _left(document, sample, copies) == []
    assert {c: _texts(document, c) for c in copies} == texts       # their own text and rows
    assert [len(document.table(c).rows) for c in copies] == [3, 2]
    first, second = report["copies"]
    assert first["copy"] == copies[0] and first["changed"] and first["table"] == ["alignment"]
    assert first["widths"] and first["rows"] == 3 and first["cells"] and first["runs"]
    assert first["not_followed"] == [] and second["not_followed"] == []
    # The extra row took the sample's last body row's look, and its empty cells its run
    # formatting on their paragraph marks (text typed there is bold at 9 pt).
    table = document.table(copies[0])
    assert document.cell_properties(copies[0], 2, 1)["vertical_alignment"] == "center"
    mark = table.cell(2, 1).paragraphs[0]._element.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr/"
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
    assert mark is not None and mark.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b") is not None
    assert document.validate() == []
    # One undo step takes every copy back.
    assert document.undo()
    assert {"widths", "bold"} <= set(_left(document, sample, copies))


def test_preview_reports_the_same_and_changes_nothing():
    document, sample, copies = _document()
    before = document.to_bytes()
    _, preview = document.follow_sample(sample, copies, preview=True)
    assert document.to_bytes() == before
    result, applied = document.follow_sample(sample, copies)
    assert preview["preview"] and not applied["preview"]
    assert preview["copies"] == applied["copies"]
    # Following again finds nothing to change: no step.
    again, report = document.follow_sample(sample, copies)
    assert not again.changed and not any(e["changed"] for e in report["copies"])


def test_a_copy_s_own_emphasis_stays():
    document, sample, copies = _document()
    paragraph = document.table(copies[1]).cell(1, 0).paragraphs[0]
    paragraph.set_text("Note: a part")
    paragraph.range(0, 5).format(italic=True)
    document.follow_sample(sample, copies[1])
    element = document.table(copies[1]).cell(1, 0).paragraphs[0]._element
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    italic = [r for r in element.iter(w + "r") if r.find(f"{w}rPr/{w}i") is not None]
    assert italic and "".join(t.text for r in italic for t in r.iter(w + "t")) == "Note:"
    assert all(r.find(f"{w}rPr/{w}b") is not None for r in element.iter(w + "r"))


def test_what_cannot_follow_is_reported_with_why():
    document, sample, copies = _document()
    # One copy gains a column, the other has two cells merged and is in a list.
    document.insert_column(copies[0], 2)
    document.merge_cells(copies[1], (1, 1), (1, 2))
    listed = document.table(copies[1]).cell(1, 0).paragraphs[0].id
    document.add_to_list(listed, "bullet")
    _, report = document.follow_sample(sample, copies)
    wider, merged = report["copies"]
    found = {d["property"]: d for d in wider["not_followed"]}
    assert set(found) >= {"columns", "widths"}
    assert found["columns"]["why"] == WHY["columns"] and found["columns"]["value"] == 4
    assert wider["changed"]                                  # the table's own properties followed
    found = {d["property"]: d for d in merged["not_followed"]}
    assert found["merges"]["why"] == WHY["merges"] and found["merges"]["fix"]["tool"] == "word_edit_table"
    assert found["list"]["paragraphs"] == [listed] and found["list"]["why"] == WHY["list"]
    # Every entry says why.
    for entry in report["copies"]:
        assert all(d["why"] for d in entry["not_followed"])
        assert len(entry["not_followed"]) <= NOT_FOLLOWED_CAP


def test_the_sample_itself_and_another_kind_of_block():
    document, sample, copies = _document()
    # Rule 3: nothing to do is no error.
    result, report = document.follow_sample(sample, [sample])
    assert not result.changed and report["copies"] == []
    assert any("the sample itself" in w for w in result.warnings)
    paragraph = document.paragraphs()[0].id
    _, report = document.follow_sample(sample, [paragraph])
    assert report["copies"][0]["changed"] is False
    assert report["copies"][0]["not_followed"][0]["property"] == "kind"


def test_a_paragraph_follows_a_sample_paragraph():
    document = Document.new()
    document.insert_markdown("Lead one.\n\nLead two, a copy.\n\nLead three, a copy.", at="end")
    sample, *copies = [p.id for p in document.paragraphs()[:3]]
    document.get(sample).format(alignment="center", space_after=18, italic=True, color="2F5496")
    result, report = document.follow_sample(sample, copies)
    assert result.count == 2
    assert [d for d in document.compare_blocks(copies, sample)["differences"]] == []
    assert [document.get(c).text for c in copies] == ["Lead two, a copy.", "Lead three, a copy."]


def test_tracked_the_copies_changes_are_revisions_and_rows_untracked():
    document, sample, copies = _document()
    with document.tracking(author=AUTHOR, date=DATE):
        result, _ = document.follow_sample(sample, copies)
    assert any("does not track a row" in w for w in result.warnings)
    xml = document.package.tree(document.package.document_part())
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for tag in ("tblPrChange", "tblGridChange", "tcPrChange", "rPrChange"):
        assert next(xml.iter(w + tag), None) is not None, tag
    assert document.validate() == []
    # Rejecting them takes the copies back (but the rows, untracked as Word does).
    document.reject_all()
    assert {"widths", "bold", "shading", "table_alignment"} <= set(_left(document, sample, copies))


def test_from_another_document_its_styles_by_name(box):
    session = box.session()
    samples = Document.new()
    samples.insert_markdown(SAMPLE, at="end")
    samples.styles.add("Part text", "paragraph", size=9)
    d1 = session.open(samples.to_bytes(), "samples.docx")
    report = Document.new()
    report.insert_markdown("A report.", at="end")
    d2 = session.open(report.to_bytes(), "report.docx")
    sample = session.entry(d1).document.tables()[0].id
    copies = [_ok(box, session, "word_copy_from", doc=d2, source_doc=d1, range=sample, at="end").data["blocks"][0]
              for _ in range(2)]
    # The sample is revised: Table Grid's built-in style, a narrower column, and its body text
    # in a style the report lacks (a custom one) and one Word has (built-in).
    _ok(box, session, "word_format_table", doc=d1, table=sample, scope="column", column=2, width=60)
    cells = session.entry(d1).document.table(sample)
    _ok(box, session, "word_format", doc=d1, targets=[cells.cell(1, 0).paragraphs[0].id], paragraph_style="Part text")
    _ok(box, session, "word_format", doc=d1, targets=[cells.cell(1, 1).paragraphs[0].id], paragraph_style="Caption")
    preview = _ok(box, session, "word_follow_sample", doc=d2, sample={"doc": d1, "block": sample}, copies=copies,
                  preview=True)
    assert preview.changed == [] and "would follow" in preview.summary
    result = _ok(box, session, "word_follow_sample", doc=d2, sample={"doc": d1, "block": sample}, copies=copies)
    assert result.changed == copies and result.data["copies"] == preview.data["copies"]
    for entry in result.data["copies"]:
        assert entry["widths"]
        styles = [d for d in entry["not_followed"] if d["property"] == "paragraph_style"]
        assert [(d["reference"], d["why"]) for d in styles] == [("Part text", WHY["paragraph_style"])]
    document = session.entry(d2).document
    assert document.styles.find("Caption", "paragraph") is not None       # added, a built-in
    assert document.styles.find("Part text") is None
    assert document.validate() == []
    _ok(box, session, "undo", doc=d2)
    assert document.styles.find("Caption", "paragraph") is None


def test_in_a_batch_and_by_reference(box):
    session = box.session()
    document, sample, copies = _document()
    doc = session.open(document.to_bytes(), "parts.docx")
    result = _ok(box, session, "batch", ops=[
        {"tool": "word_format_table", "arguments": {"doc": doc, "table": sample, "scope": "table", "alignment": "right"}},
        {"tool": "word_follow_sample", "arguments": {"doc": doc, "sample": {"block": sample}, "copies": copies}}])
    assert result.ok
    assert _left(session.entry(doc).document, sample, copies) == []


# -- the tool design rules (tool-design-rules.md) ------------------------------------------------


def test_rule_1_addresses_other_tools_take_and_placeholders():
    document, sample, copies = _document()
    cell = f"{copies[0]}/c1,1"
    paragraph = document.table(copies[1]).cell(1, 0).paragraphs[0].id
    sample_cell = document.table(sample).cell(0, 0).paragraphs[0].id
    result, report = document.follow_sample(sample_cell, [cell, "", paragraph, copies[0]])
    assert [e["copy"] for e in report["copies"]] == copies and report["sample"] == sample
    notes = " ".join(result.warnings)
    assert f"{cell} is in table {copies[0]}" in notes and "empty copy address was ignored" in notes
    assert f"is a paragraph in table {copies[1]}" in notes and "named twice" in notes
    assert _left(document, sample, copies) == []


def test_rule_2_a_refusal_names_the_tables_and_their_shapes():
    document, sample, copies = _document()
    with pytest.raises(Exception) as caught:
        document.follow_sample(sample, ["t:0BADF00D"])
    message = str(caught.value)
    assert "t:0BADF00D" in message and f"{copies[0]} (3 rows x 3 columns)" in message
    with pytest.raises(Exception, match="copies names at least one block"):
        document.follow_sample(sample, [])


def test_rule_3_a_copy_already_following_does_not_fail_a_batch(box):
    session = box.session()
    document, sample, copies = _document()
    doc = session.open(document.to_bytes(), "parts.docx")
    follow = {"tool": "word_follow_sample", "arguments": {"doc": doc, "sample": {"doc": "", "block": sample},
                                                          "copies": copies, "preview": False}}
    result = _ok(box, session, "batch", ops=[follow, follow])
    assert result.ok
    again = _ok(box, session, "word_follow_sample", doc=doc, sample={"block": sample}, copies=copies)
    assert again.changed == [] and "nothing left to follow" in again.summary
    assert all(e["changed"] is False for e in again.data["copies"])


def test_rule_4_the_reply_has_a_budget_and_names_what_it_left_out(monkeypatch):
    from docx_agent.edit import follow as _follow

    document, sample, copies = _document()
    for copy in copies:
        document.insert_column(copy, 2)
        document.merge_cells(copy, (1, 0), (1, 1))
    monkeypatch.setattr(_follow, "REPLY_CAP", 3)
    _, report = document.follow_sample(sample, copies)
    shown = sum(len(e["not_followed"]) for e in report["copies"])
    assert shown == 3 and sum(e.get("more", 0) for e in report["copies"]) > 0
    assert "word_compare_blocks" in report["left_out"]
