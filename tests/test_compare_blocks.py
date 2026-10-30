"""Round 17, P3: ``word_compare_blocks`` -- tables and paragraphs copied from a sample,
compared with the sample (here or in another open document), every difference with the
reference's value and the calls that give it; the fix calls run through the tools, and a
second comparison finds nothing left to fix."""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit 0.4 (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.edit.compare import WIDTH_TOLERANCE, compare_blocks  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

SAMPLE = ("Samples.\n\n"
          "| Item | Owner | Status |\n| --- | --- | --- |\n| Sample item | Someone | Open |\n\n"
          "A closing line.")


def _sample() -> tuple[Document, str]:
    """A made-up sample: a table whose header row repeats and is shaded, its body row's
    middle cell centred vertically, a heading-styled paragraph after it."""
    sample = Document.new()
    sample.insert_markdown(SAMPLE, at="end")
    table = sample.tables()[0].id
    sample.set_table(table, header_rows=1)
    for column in range(3):
        sample.set_cell(table, 0, column, shading="1F3864")
    sample.set_cell(table, 1, 1, vertical_alignment="center")
    return sample, table


def _differences(data: dict, block: str | None = None) -> dict[str, dict]:
    return {d["property"]: d for d in data["differences"] if block is None or d["block"] == block}


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


def _ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, arguments, result.error and result.error.to_json())
    return result


def _apply(box, session, doc: str, data: dict) -> int:
    """Every fix call of a comparison, in the order listed; how many."""
    count = 0
    for difference in data["differences"]:
        if difference["fix"]:
            for call in difference["fix"]["calls"]:
                _ok(box, session, difference["fix"]["tool"], doc=doc, **call)
                count += 1
    return count


def test_a_drifted_copy_from_another_document_is_fixed_toward_its_sample(box):
    sample, reference = _sample()
    session = box.session()
    d2 = session.open(sample.to_bytes(), "samples.docx")
    report = Document.new()
    report.insert_markdown("A report.", at="end")
    d1 = session.open(report.to_bytes(), "report.docx")
    first = _ok(box, session, "word_copy_from", doc=d1, source_doc=d2, range=reference, at="end").data["blocks"][0]
    second = _ok(box, session, "word_copy_from", doc=d1, source_doc=d2, range=reference, at="end").data["blocks"][0]
    document = session.entry(d1).document
    # Filled: another item row in each (more rows than the sample: a difference by design).
    for table in (first, second):
        _ok(box, session, "word_edit_table", doc=d1, table=table, action="insert_row", index=1, side="after")
    # The second drifts: no repeated header, a shaded item row, a column made narrower, a
    # cell's text bold and centred, two cells merged, a vertical alignment lost.
    _ok(box, session, "word_format_table", doc=d1, table=second, scope="table", header_rows=0)
    _ok(box, session, "word_format_table", doc=d1, table=second, scope="row", row=1, shading="F2F2F2")
    _ok(box, session, "word_format_table", doc=d1, table=second, scope="column", column=0, width=100)
    cell = document.table(second).cell(1, 2).paragraphs[0].id
    _ok(box, session, "word_format", doc=d1, targets=[cell], bold=True, alignment="center")
    _ok(box, session, "word_edit_table", doc=d1, table=second, action="merge", first={"row": 2, "col": 1},
        last={"row": 2, "col": 2})
    _ok(box, session, "word_format_table", doc=d1, table=second, scope="cell", row=1, column=1,
        vertical_alignment="bottom")

    result = _ok(box, session, "word_compare_blocks", doc=d1, blocks=[first, second],
                 reference={"doc": d2, "block": reference})
    data = result.data
    assert data["reference"] == reference and data["compared"] == [first, second] and data["same"] == []
    # The untouched copy differs only by its extra row, which has no fix.
    assert {d["property"]: (d["value"], d["reference"], d["fix"]) for d in data["differences"]
            if d["block"] == first} == {"rows": (2, 1, None)}
    found = _differences(data, second)
    assert {"header_rows", "rows", "merges", "widths", "shading", "vertical_alignment", "alignment", "bold"} \
        <= set(found)
    assert found["header_rows"]["value"] == 0 and found["header_rows"]["reference"] == 1
    assert found["shading"]["cells"] == [f"{second}/c1,*"]
    assert (found["shading"]["value"], found["shading"]["reference"]) == ("F2F2F2", None)
    assert found["shading"]["fix"] == {"tool": "word_format_table", "calls": [
        {"table": second, "scope": "row", "row": 1, "shading": "none"}]}
    assert found["widths"]["columns"] == [0] and found["widths"]["value"][0] == 100
    assert found["merges"]["value"] == ["c1..c2"] and found["merges"]["reference"] == []
    assert found["bold"]["value"] is True and found["bold"]["reference"] is None
    assert found["bold"]["fix"]["calls"] == [{"targets": [cell], "clear": ["bold"]}]
    # Every difference has the shape ppt_compare_slides's have, whatever it compares.
    for difference in data["differences"]:
        assert {"block", "property", "value", "reference", "fix"} <= set(difference)
    assert "Compared 2 block(s)" in result.summary

    assert _apply(box, session, d1, data) >= 7
    again = _ok(box, session, "word_compare_blocks", doc=d1, blocks=[first, second],
                reference={"doc": d2, "block": reference}).data
    assert [d["property"] for d in again["differences"]] == ["rows", "rows"], again["differences"]
    assert all(d["fix"] is None for d in again["differences"])
    assert session.entry(d1).document.validate() == []


def test_paragraphs_compare_their_style_and_direct_formatting():
    document = Document.new()
    made = document.insert_markdown("# A heading\n\nBody text.\n\n# Another heading\n\nMore text.", at="end")
    first, _, second, body = made.blocks[:4]
    document.format_paragraph(second, size=20, space_after=3)
    document.set_paragraph_style(body, "Heading 2")
    data = document.compare_blocks([second, body], first)
    found = _differences(data, second)
    assert set(found) == {"size", "space_after"}
    assert found["size"]["value"] == 20 and found["size"]["reference"] is None
    assert found["size"]["fix"] == {"tool": "word_format", "calls": [{"targets": [second], "clear": ["size"]}]}
    other = _differences(data, body)
    assert other["paragraph_style"]["value"] == "heading 2" and other["paragraph_style"]["reference"] == "heading 1"
    assert other["paragraph_style"]["fix"]["calls"] == [{"targets": [body], "paragraph_style": "heading 1"}]
    # A table against a paragraph: only the kind.
    document.insert_markdown("| a | b |\n| --- | --- |\n| 1 | 2 |", at="end")
    table = document.tables()[0].id
    assert compare_blocks(document, [table], first)["differences"] == [
        {"block": table, "property": "kind", "value": "table", "reference": "paragraph", "fix": None}]


def test_the_same_sample_in_the_same_document_and_the_width_tolerance(box):
    sample, reference = _sample()
    copy = sample.copy_blocks(sample, reference, at="end").blocks[0]
    assert sample.compare_blocks([copy], reference)["same"] == [copy]
    columns = sample.table(reference).properties["columns"]
    sample.set_column_width(copy, 1, columns[1] + WIDTH_TOLERANCE)
    assert sample.compare_blocks([copy], reference)["same"] == [copy]
    sample.set_column_width(copy, 1, columns[1] + 1)
    found = _differences(sample.compare_blocks([copy], reference))
    assert set(found) == {"widths"} and found["widths"]["columns"] == [1]
    # Through the tool, the reference in the same document: no doc in it.
    session = box.session()
    d = session.open(sample.to_bytes(), "samples.docx")
    data = _ok(box, session, "word_compare_blocks", doc=d, blocks=[copy], reference={"block": reference}).data
    assert _apply(box, session, d, data) == 1
    assert _ok(box, session, "word_compare_blocks", doc=d, blocks=[copy],
               reference={"block": reference}).data["same"] == [copy]
    error = box.dispatch(session, "word_compare_blocks", {"doc": d, "blocks": [copy],
                                                          "reference": {"doc": "d9", "block": reference}})
    assert not error.ok and error.error.code == "not_found"


def test_table_style_and_borders_are_fixed_where_the_tool_writes_them(box):
    sample, reference = _sample()
    session = box.session()
    d = session.open(sample.to_bytes(), "samples.docx")
    _ok(box, session, "word_format_table", doc=d, table=reference, scope="table", borders="single")
    # Borders alone leave the table's style as it was (they set it to "single" before).
    assert session.entry(d).document.table(reference).properties["style"] == "Table Grid"
    _ok(box, session, "word_format_table", doc=d, table=reference, scope="cell", row=1, column=0, borders="dashed")
    copy = _ok(box, session, "word_copy_from", doc=d, source_doc=d, range=reference, at="end").data["blocks"][0]
    assert _ok(box, session, "word_compare_blocks", doc=d, blocks=[copy],
               reference={"block": reference}).data["same"] == [copy]
    _ok(box, session, "word_format_table", doc=d, table=copy, scope="table", style="Normal Table", borders="dotted")
    _ok(box, session, "word_format_table", doc=d, table=copy, scope="cell", row=1, column=0, borders="double")
    data = _ok(box, session, "word_compare_blocks", doc=d, blocks=[copy], reference={"block": reference}).data
    found = _differences(data)
    assert set(found) == {"table_style", "table_borders", "borders"}, data["differences"]
    assert (found["table_style"]["value"], found["table_style"]["reference"]) == ("Normal Table", "Table Grid")
    assert found["table_borders"]["reference"]["top"] == "single 0.5pt auto"
    assert found["table_borders"]["fix"]["calls"] == [{"table": copy, "scope": "table", "borders": "single"}]
    # The cell's four sides, and the sides its neighbours share with it (setting a cell's
    # border sets the neighbour's side it meets too): only the cell's own has a fix, which
    # gives the neighbours theirs back.
    borders = {tuple(e["cells"]): e for e in data["differences"] if e["property"] == "borders"}
    assert set(borders) == {(f"{copy}/c1,0",), (f"{copy}/c0,0",), (f"{copy}/c1,1",)}
    assert borders[(f"{copy}/c1,0",)]["fix"]["calls"] == [{"table": copy, "scope": "cell", "row": 1, "column": 0,
                                                          "borders": "dashed"}]
    assert borders[(f"{copy}/c1,1",)]["value"] == {"left": "double 0.5pt auto"}
    assert borders[(f"{copy}/c1,1",)]["fix"] is None
    assert _apply(box, session, d, data) == 3
    assert _ok(box, session, "word_compare_blocks", doc=d, blocks=[copy],
               reference={"block": reference}).data["same"] == [copy]


# -- round 26, W6: a cell's paragraphs paired by an alignment -------------------------------------


def _two_paragraph_cell() -> tuple[Document, str, str]:
    """A made-up sample whose notes cell holds a heading-styled label line and a body line,
    and a copy of it in the same document: ``(document, sample, copy)``."""
    document = Document.new()
    document.insert_markdown("| Item | Notes |\n| --- | --- |\n| Item one | Body text here |", at="end")
    sample = document.tables()[0].id
    document.table(sample).cell(1, 1).paragraphs[0].insert_before("Section label", style="Heading 2")
    copy = document.copy_blocks(document, sample, at="end").blocks[0]
    assert document.compare_blocks([copy], sample)["same"] == [copy]
    return document, sample, copy


def _notes(document: Document, table: str) -> list:
    return document.table(table).cell(1, 1).paragraphs


def test_a_cell_missing_its_label_line_is_never_restyled_as_the_label(box):
    document, sample, copy = _two_paragraph_cell()
    label, body = _notes(document, copy)
    label.delete()
    body.set_text("Filled in by hand")
    document.format_paragraph(body.id, bold=True)
    data = document.compare_blocks([copy], sample)
    found = _differences(data)
    # By position the body line would pair with the label and be given its style.
    assert "paragraph_style" not in found, data["differences"]
    assert set(found) == {"paragraphs", "missing_paragraph", "bold"}
    assert (found["paragraphs"]["value"], found["paragraphs"]["reference"], found["paragraphs"]["fix"]) == (1, 2, None)
    assert found["paragraphs"]["cells"] == [f"{copy}/c1,1"]
    assert found["missing_paragraph"]["cells"] == [f"{copy}/c1,1"] and found["missing_paragraph"]["fix"] is None
    assert found["missing_paragraph"]["reference"] == {"paragraph_style": "heading 2", "text": "Section label",
                                                       "index": 0}
    # The body line pairs with the sample's body line (by style): its direct formatting is compared.
    assert found["bold"]["paragraphs"] == [body.id]
    assert found["bold"]["fix"] == {"tool": "word_format", "calls": [{"targets": [body.id], "clear": ["bold"]}]}
    # Through the tool: the fixes leave only what no tool fixes.
    session = box.session()
    d = session.open(document.to_bytes(), "cells.docx")
    data = _ok(box, session, "word_compare_blocks", doc=d, blocks=[copy], reference={"block": sample}).data
    assert _apply(box, session, d, data) == 1
    again = _ok(box, session, "word_compare_blocks", doc=d, blocks=[copy], reference={"block": sample}).data
    assert [e["property"] for e in again["differences"]] == ["paragraphs", "missing_paragraph"]


def test_a_cell_with_a_line_the_sample_has_not_reports_it_as_extra():
    document, copy_of_sample, sample = _two_paragraph_cell()   # the two-line cell is the copy now
    label, body = _notes(document, sample)
    label.delete()
    data = document.compare_blocks([copy_of_sample], sample)
    found = _differences(data)
    assert set(found) == {"paragraphs", "extra_paragraph"}, data["differences"]
    assert (found["paragraphs"]["value"], found["paragraphs"]["reference"]) == (2, 1)
    extra = found["extra_paragraph"]
    assert extra["paragraphs"] == [_notes(document, copy_of_sample)[0].id]
    assert extra["value"] == {"paragraph_style": "heading 2", "text": "Section label"}
    assert extra["reference"] is None and extra["fix"] is None


def test_reordered_lines_are_missing_and_extra_not_restyled():
    document, sample, copy = _two_paragraph_cell()
    label, body = _notes(document, copy)
    label.move(after=body.id)
    data = document.compare_blocks([copy], sample)
    found = _differences(data)
    # Equal counts, but the texts say the pairs are crossed: no style fixes, no count.
    assert set(found) == {"missing_paragraph", "extra_paragraph"}, data["differences"]
    missing = found["missing_paragraph"]["reference"]
    assert found["extra_paragraph"]["value"] == {k: v for k, v in missing.items() if k != "index"}
    assert all(d["fix"] is None for d in data["differences"])


def test_equal_counts_keep_pairing_by_position():
    document, sample, copy = _two_paragraph_cell()
    label, body = _notes(document, copy)
    # Filled with other text and the label restyled: still compared line by line.
    label.set_text("Another label")
    body.set_text("Other text entirely")
    document.set_paragraph_style(label.id, "Heading 3")
    found = _differences(document.compare_blocks([copy], sample))
    assert set(found) == {"paragraph_style"}
    assert found["paragraph_style"]["paragraphs"] == [label.id]
    assert found["paragraph_style"]["fix"]["calls"] == [{"targets": [label.id], "paragraph_style": "heading 2"}]


def test_a_label_matching_by_text_is_restyled_and_like_lines_repeat_their_pattern():
    document, sample, copy = _two_paragraph_cell()
    label, body = _notes(document, copy)
    # The label's text a little edited and restyled, the body line gone: the label still
    # pairs by its text, so its style is fixed; the body line is missing.
    body.delete()
    label.set_text("Section labels")
    document.set_paragraph_style(label.id, "Heading 3")
    found = _differences(document.compare_blocks([copy], sample))
    assert set(found) == {"paragraphs", "missing_paragraph", "paragraph_style"}
    assert found["missing_paragraph"]["reference"]["text"] == "Body text here"
    assert found["paragraph_style"]["fix"]["calls"] == [{"targets": [label.id], "paragraph_style": "heading 2"}]
    # A sample line is the pattern for more lines like it after it: a cell filled with
    # three body lines under its label is as the sample.
    document, sample, copy = _two_paragraph_cell()
    _, body = _notes(document, copy)
    body.insert_after("A second body line")
    _notes(document, copy)[2].insert_after("A third body line")
    assert len(_notes(document, copy)) == 4
    data = document.compare_blocks([copy], sample)
    assert data["same"] == [copy], data["differences"]
