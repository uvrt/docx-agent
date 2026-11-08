"""Round 42: ``word_set_text`` writes a paragraph's style and its inline emphasis in the
template's character styles in one call (``paragraph_style``; ``**strong**``, ``*emphasis*``,
``[text]{Character style}``), and text without markup is written as before.  One test (at
least) per tool design rule it touches.  Made-up documents, Document.new()'s package."""

from __future__ import annotations

import pytest
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.markdown.styled import parse
from docx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tables_from_data import GRID, _session, ds_put  # noqa: E402
from test_tables_from_data import call as dispatch  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

MD = """# Results

The mean was **4.0/5** over the year.

Second paragraph here.

| Name | Value |
|---|---|
| left | right |
"""


def made() -> bytes:
    document = Document.new()
    document.insert_markdown(MD, at="end")
    document.styles.add("Report Body", "paragraph", based_on="Normal", space_after=9)
    document.styles.add("Score Mark", "character", bold=True, color="1F4E79")
    return document.to_bytes()


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        yield toolbox


@pytest.fixture
def opened(box):
    session = box.session()
    doc = session.open(made(), "styled.docx")
    document = session.entry(doc).document
    return session, doc, document


def call(box, session, doc, **arguments):
    return box.dispatch(session, "word_set_text", {"doc": doc, **arguments})


def ok(box, session, doc, **arguments):
    result = call(box, session, doc, **arguments)
    assert result.ok, result.error and result.error.to_json()
    return result


def runs(document, identifier):
    """(text, style id, bold, italic) per text run."""
    element = document._resolve(identifier)[1].element
    out = []
    for run in element.iter(W + "r"):
        text = "".join(t.text or "" for t in run.iter(W + "t"))
        if not text:
            continue
        properties = run.find(W + "rPr")
        style = properties.find(W + "rStyle") if properties is not None else None
        out.append((text, style.get(W + "val") if style is not None else None,
                    properties is not None and properties.find(W + "b") is not None,
                    properties is not None and properties.find(W + "i") is not None))
    return out


def body(document):
    return [p for p in document.paragraphs() if p.text]


# -- the reading of the text ---------------------------------------------------------------------


def test_the_markup_read_is_strong_emphasis_and_spans_only():
    styled = parse("Mean **2.0/5**, below *3.0/5*; see [the scale]{Score Mark}.")
    assert styled.plain == "Mean 2.0/5, below 3.0/5; see the scale."
    assert [(p.text, sorted(p.marks), p.style) for p in styled.pieces if p.marks or p.style] == [
        ("2.0/5", ["strong"], None), ("3.0/5", ["em"], None), ("the scale", [], "Score Mark")]
    # Nothing else is Markdown: links, code, HTML, entities, line starts, underscores.
    for text in ("a & b &amp; <b> `c` [x](y) ~~s~~", "# 3 items", "- not a list", "file_name and _draft_",
                 "5 * 3 * 2", "**unclosed", "[a bracket] {not a span}", "[x]{}"):
        assert parse(text).plain == text and not parse(text).marked, text


def test_literal_asterisks_are_escaped_with_a_backslash():
    styled = parse(r"Rated \*\*\* (three) and **bold \* star**, path C:\Users")
    assert styled.plain == "Rated *** (three) and bold * star, path C:\\Users"
    assert [p.text for p in styled.pieces if p.marks] == ["bold * star"]
    assert parse(r"\[x]{Score Mark}").plain == "[x]{Score Mark}"


def test_span_forms_and_marks_inside_a_span():
    for text in ('[x]{Score Mark}', '[x]{style="Score Mark"}', "[x]{custom-style='Score Mark'}"):
        assert [(p.text, p.style) for p in parse(text).pieces] == [("x", "Score Mark")], text
    inner = parse("[**a** b]{Score Mark}").pieces
    assert [(p.text, sorted(p.marks), p.style) for p in inner] == [("a", ["strong"], "Score Mark"),
                                                                    (" b", [], "Score Mark")]


# -- writing ---------------------------------------------------------------------------------------


def test_style_and_emphasis_in_one_call_with_the_runs_said(box, opened):
    session, doc, document = opened
    target = body(document)[1].id
    result = ok(box, session, doc, items=[{"target": target, "paragraph_style": "Report Body",
                                           "text": "The mean was **2.0/5**, below *3.0/5* in [March]{Score Mark}."}])
    paragraph = document.paragraph(target)
    assert paragraph.style_name == "Report Body"
    assert paragraph.text == "The mean was 2.0/5, below 3.0/5 in March."
    styles = {s.id: s.name for s in document.styles}
    assert [(t, styles.get(s)) for t, s, _, _ in runs(document, target) if s] == [
        ("2.0/5", "Strong"), ("3.0/5", "Emphasis"), ("March", "Score Mark")]
    assert result.data["styled"] == [
        {"range": f"{target}@13:18", "text": "2.0/5", "style": "Strong"},
        {"range": f"{target}@26:31", "text": "3.0/5", "style": "Emphasis"},
        {"range": f"{target}@35:40", "text": "March", "style": "Score Mark"}]
    # The ranges said are the characters styled.
    for entry in result.data["styled"]:
        assert document.range(entry["range"]).text == entry["text"]
    assert "styled" in result.summary
    # One undo step for the style, the text and the emphasis.
    ok_undo = box.dispatch(session, "undo", {"doc": doc})
    assert ok_undo.ok and document.paragraph(target).text == "The mean was 4.0/5 over the year."
    assert document.paragraph(target).style_name == "Normal"


def test_markup_states_the_emphasis_old_emphasis_goes_other_formatting_stays(box, opened):
    session, doc, document = opened
    target = body(document)[1].id
    # The old text has Strong on 4.0/5; give the rest a colour (direct formatting not emphasis).
    document.range(f"{target}@0:12").format(color="C00000", bold=True)
    ok(box, session, doc, target=target, text="The mean was 4.0/5 over the *year*.")
    look = runs(document, target)
    # 4.0/5 is unmarked now: no Strong; "The mean was" keeps its colour but not its bold.
    assert all(style is None for text, style, _, _ in look if "4.0/5" in text)
    assert not any(bold for _, _, bold, _ in look)
    element = document._resolve(target)[1].element
    assert any(c.get(W + "val") == "C00000" for c in element.iter(W + "color"))
    assert [t for t, s, _, _ in look if s] == ["year"]


def test_plain_text_keeps_its_formatting_as_before(box, opened):
    session, doc, document = opened
    target = body(document)[1].id
    result = ok(box, session, doc, target=target, text="The mean was 4.0/5 over the decade.")
    assert result.data is None
    assert [(t, s) for t, s, _, _ in runs(document, target) if s] == [("4.0/5", "Strong")]
    # A lone asterisk or an underscore is text, written as before.
    ok(box, session, doc, target=target, text="The mean was 4.0/5 * 2 over file_name.")
    assert document.paragraph(target).text == "The mean was 4.0/5 * 2 over file_name."
    assert [(t, s) for t, s, _, _ in runs(document, target) if s] == [("4.0/5", "Strong")]


def test_rule_1_a_style_beside_items_is_each_items_and_an_id_names_a_style(box, opened):
    session, doc, document = opened
    first, second = body(document)[1].id, body(document)[2].id
    ok(box, session, doc, paragraph_style="Report Body",
       items=[{"target": first, "text": "One."}, {"target": second, "text": "Two.", "paragraph_style": ""}])
    assert {document.paragraph(first).style_name, document.paragraph(second).style_name} == {"Report Body"}
    style_id = document.styles.find("Report Body", "paragraph").id
    heading = body(document)[0].id
    ok(box, session, doc, target=heading, text="Results", paragraph_style=style_id)
    assert document.paragraph(heading).style_name == "Report Body"


def test_rule_2_an_unknown_style_is_refused_with_the_ones_there_are(box, opened):
    session, doc, document = opened
    target = body(document)[1].id
    before = document.paragraph(target).text
    refused = call(box, session, doc, items=[{"target": target, "text": "x", "paragraph_style": "Made Up Style"}])
    assert not refused.ok and refused.error.code == "not_found"
    assert refused.error.field == "items[0].paragraph_style" and "Report Body" in refused.error.valid_options
    refused = call(box, session, doc, items=[{"target": target, "text": "**x** in [y]{No Such Style}"}])
    assert not refused.ok and refused.error.field == "items[0].text"
    assert "Score Mark" in refused.error.valid_options and "\\[" in refused.error.message
    refused = call(box, session, doc, target=target, text="[y]{Report Body}")
    assert not refused.ok and "paragraph style" in refused.error.message
    # Refused whole, nothing written; an earlier good item included.
    refused = call(box, session, doc, items=[{"target": target, "text": "Good."},
                                             {"target": target, "text": "x", "paragraph_style": "Nope"}])
    assert not refused.ok and document.paragraph(target).text == before


def test_rule_3_the_same_again_changes_nothing(box, opened):
    session, doc, document = opened
    target = body(document)[1].id
    arguments = {"items": [{"target": target, "paragraph_style": "Report Body",
                            "text": "The mean was **2.0/5** in [March]{Score Mark}."}]}
    ok(box, session, doc, **arguments)
    again = ok(box, session, doc, **arguments)
    assert again.changed == [] and any("already in Report Body" in w for w in again.warnings)
    assert again.data["styled"]            # what the runs are, said again
    same_style = ok(box, session, doc, target=target, text=document.paragraph(target).text,
                    paragraph_style="Report Body")
    assert same_style.changed == []


def test_rule_4_the_runs_listed_are_capped_and_the_rest_counted(box, opened):
    session, doc, document = opened
    target = body(document)[1].id
    text = " ".join(f"**w{k}**" for k in range(45))
    result = ok(box, session, doc, target=target, text=text)
    assert len(result.data["styled"]) == 40 and result.data["left_out"] == 5
    assert any("5 more styled run(s)" in w for w in result.warnings)


def test_rule_6_the_template_owns_the_styles_one_it_lacks_added_as_word_writes_it(box):
    session = box.session()
    document = Document.new()
    document.insert_markdown("Plain one.", at="end")
    doc = session.open(document.to_bytes(), "bare.docx")
    document = session.entry(doc).document
    assert document.styles.find("Strong", "character") is None
    target = body(document)[0].id
    result = ok(box, session, doc, target=target, text="Plain **one**.")
    assert "style added: Strong, as Word writes it" in result.warnings
    strong = document.styles.find("Strong", "character")
    assert strong is not None and [(t, s) for t, s, _, _ in runs(document, target) if s] == [("one", strong.id)]
    # No warning when the style is there, or for a role the text does not use.
    again = ok(box, session, doc, target=target, text="Plain **two**.")
    assert not any("style added" in w for w in again.warnings)


def test_lines_and_cells_take_the_style_and_the_emphasis(box, opened):
    session, doc, document = opened
    target = body(document)[2].id
    result = ok(box, session, doc, target=target, paragraph_style="Report Body", text="One **a**\nTwo *b*")
    new = result.created[0]
    assert document.paragraph(new).style_name == "Report Body" == document.paragraph(target).style_name
    assert [r["range"] for r in result.data["styled"]] == [f"{target}@4:5", f"{new}@4:5"]
    table = document.tables()[0].id
    cell = f"{table}/c1,1"
    result = ok(box, session, doc, target=cell, text="**big** value\nsecond *line*")
    paragraphs = document.table(table).cell(1, 1).paragraphs
    assert [p.text for p in paragraphs] == ["big value", "second line"]
    assert [r["text"] for r in result.data["styled"]] == ["big", "line"]


def test_marks_beside_a_span_are_direct(box, opened):
    session, doc, document = opened
    target = body(document)[1].id
    result = ok(box, session, doc, target=target, text="A [**b**]{Score Mark} c.")
    assert result.data["styled"] == [{"range": f"{target}@2:3", "text": "b", "style": "Score Mark", "bold": True}]
    mark = document.styles.find("Score Mark", "character").id
    assert ("b", mark, True, False) in runs(document, target)


def test_tracked_emphasis_is_a_formatting_revision_and_stays_valid(box, opened):
    session, doc, document = opened
    assert box.dispatch(session, "word_set_tracking", {"doc": doc, "on": True, "author": "Writer"}).ok
    target = body(document)[1].id
    ok(box, session, doc, target=target, text="The mean was **4.0/5** over the *year*.")
    element = document._resolve(target)[1].element
    assert element.find(f".//{W}rPrChange") is not None
    assert document.validate() == []


def test_styled_text_lays_out_as_the_same_text_written_by_hand():
    """The layout check: set_text's result renders as a paragraph written with the same
    styles by hand (insert_markdown, the character style applied to a range)."""
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session()
        doc = session.open(made(), "a.docx")
        document = session.entry(doc).document
        target = body(document)[1].id
        assert toolbox.dispatch(session, "word_set_text", {"doc": doc, "target": target, "paragraph_style": "Report Body",
                                                           "text": "Mean **2.0/5**, below *3.0/5* in [May]{Score Mark}."}).ok
        written = document.render_svg()
    hand = Document.open(made())
    target = body(hand)[1].id
    hand.paragraph(target).set_style("Report Body")
    hand.paragraph(target).set_text("Mean 2.0/5, below 3.0/5 in May.")
    hand.range(f"{target}@0:{len('Mean 2.0/5, below 3.0/5 in May.')}").set_style(None)
    hand.range(f"{target}@5:10").set_style("Strong")
    hand.range(f"{target}@18:23").set_style("Emphasis")
    hand.range(f"{target}@27:30").set_style("Score Mark")
    assert written == hand.render_svg()


def test_a_value_written_by_reference_inside_markup_is_kept_on_its_paragraph():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as toolbox:
        session = _session(toolbox)
        assert dispatch(toolbox, session, "ds_put", id="scores").ok
        document = session.entry("d1").document
        target = document.paragraphs()[0].id
        for text in ("Gamma rated **{{d:scores/Gamma/rating}}** of [5]{Strong}.",
                     "[Gamma {{d:scores/Gamma/rating}}]{Strong} rated."):
            result = dispatch(toolbox, session, "word_set_text", doc="d1", target=target, text=text)
            assert result.ok and result.data["facts"][0]["text"] == "1.25" and result.data["styled"]
            kept = document.text_facts()[target]
            assert [(f["text"], document.paragraph(target).text[f["at"]:f["at"] + 4]) for f in kept] == \
                [("1.25", "1.25")]
