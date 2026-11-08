"""Paragraphs like a given one (round 42, R47): ``word_format``'s ``like`` and
``word_insert_markdown``'s ``like`` -- the targets, or new list items, take the sample's
paragraph style, its list at its level (numbered on from it), its direct indents and its page
settings; not its run formatting.  Laid out (docx2svg), a new item is drawn as the sample's
items are: the same marker at the same place, its text at the same place.  Made-up template
and text only."""

from __future__ import annotations

import datetime as dt
import warnings

import pytest
from lxml import etree
from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.edit import numbering as _numbering
from docx_agent.tools import FORMAT, GROUPS, TOOLS

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
CLOCK = lambda: dt.datetime(2026, 11, 8, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731


def _xml(text: str):
    return etree.fromstring(f'<w:root xmlns:w="{W_NS}">{text}</w:root>')[0]


def _level(ilvl: int, fmt: str, text: str, left: int) -> str:
    return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{left}" w:hanging="360"/></w:pPr></w:lvl>')


def template() -> Document:
    """A made-up template: List Bullet numbers itself with a three-level bullet definition
    (a dot, a dash, a ring), List Number with a two-level numbered one (1., a)), and a body
    style, Plan Body."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        document = Document.new()
    _numbering._ensure_part(document)
    _numbering.add_abstract(document, _xml(
        '<w:abstractNum w:abstractNumId="1"><w:nsid w:val="6B1A0001"/><w:multiLevelType w:val="multilevel"/>'
        + _level(0, "bullet", "•", 360) + _level(1, "bullet", "–", 720) + _level(2, "bullet", "◦", 1080)
        + "</w:abstractNum>"))
    _numbering.add_abstract(document, _xml(
        '<w:abstractNum w:abstractNumId="2"><w:nsid w:val="6B1A0002"/><w:multiLevelType w:val="multilevel"/>'
        + _level(0, "decimal", "%1.", 360) + _level(1, "lowerLetter", "%2)", 720) + "</w:abstractNum>"))
    bullet = _numbering.add_num(document, 1)
    number = _numbering.add_num(document, 2)
    styles = document.package.tree(document.package.styles_part())
    for style_id, name, ppr in (
            ("ListBullet", "List Bullet", f'<w:numPr><w:numId w:val="{bullet}"/></w:numPr>'),
            ("ListNumber", "List Number", f'<w:numPr><w:numId w:val="{number}"/></w:numPr>'),
            ("PlanBody", "Plan Body", '<w:spacing w:after="120"/>')):
        styles.append(_xml(f'<w:style w:type="paragraph" w:styleId="{style_id}"><w:name w:val="{name}"/>'
                           f'<w:basedOn w:val="Normal"/><w:pPr>{ppr}</w:pPr></w:style>'))
    document.package.mark_dirty(document.package.styles_part())
    document.package.mark_dirty(document.package.numbering_part())
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return Document.open(document.to_bytes())


def sample() -> tuple[Document, dict[str, str]]:
    """The template with a made-up section: a heading, a body paragraph, a nested bullet
    list (each point followed by its examples, a level down, as Word's Tab makes them: the
    style's list at level 1), a numbered list, a body paragraph after.  The second example
    keeps with the next paragraph and has an indent of its own; the body paragraph is Plan Body and
    keeps its lines together."""
    document = template()
    made = document.insert_markdown(
        "# River plan\n\nThe plan has three points.\n\n- First point\n- example one\n- example two\n"
        "- Second point\n\nSteps to take:\n\n1. Walk\n2. Ford\n\nThe plan ends here.\n")
    ids = dict(zip(["heading", "body", "point1", "example1", "example2", "point2", "intro", "step1", "step2",
                    "end"], made.blocks))
    for key in ("example1", "example2"):
        document.set_list_level(ids[key], 1)
    document.format_paragraph(ids["example2"], keep_with_next=True, indent_left=40)
    document.set_paragraph_style(ids["body"], "Plan Body")
    document.format_paragraph(ids["body"], keep_together=True)
    return document, ids


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


@pytest.fixture
def made(box):
    document, ids = sample()
    session = box.session(clock=CLOCK)
    session.open(document.to_bytes(), "plan.docx")
    return session, ids


def run(box, session, name, **arguments):
    return box.dispatch(session, name, arguments).to_json()


def doc_of(session) -> Document:
    return session.entry("d1").document


def own(document: Document, identifier: str, tag: str):
    found = document._resolve(identifier)[1].element.find(f"{W}pPr/{W}{tag}")
    return None if found is None else dict(found.attrib) or etree.tostring(found)


def numbering_of(document: Document, identifier: str) -> tuple | None:
    node = document._resolve(identifier)[1].element.find(f"{W}pPr/{W}numPr")
    if node is None:
        return None
    return tuple(child.get(W + "val") for child in node)


def labels(document: Document, ids: list[str]) -> list[str | None]:
    blocks = {}
    for identifier in ids:
        for block in document.state(identifier)["blocks"]:
            blocks[block["id"]] = (block.get("list") or {}).get("label")
    return [blocks.get(i) for i in ids]


def style_of(document: Document, identifier: str) -> str | None:
    """The style id a paragraph names."""
    return document.paragraph(identifier).style


def drawn(document: Document, identifier: str) -> tuple:
    """How docx2svg draws a paragraph's first line: each span's characters and pen x (the
    marker, then the text), less the text itself."""
    layout = document.layout()
    for lines in layout._lines:
        for key, line in lines:
            if key is not None and layout._canon(key) in layout._wanted(identifier):
                return tuple((span.chars[0] if span.chars else "", float(span.xs[0]) if span.xs else None)
                             for span in line.spans)
    raise AssertionError(f"{identifier} is not drawn")


# -- word_format like ----------------------------------------------------------------------------


def test_targets_join_the_samples_list_at_its_level_numbered_on(box, made):
    session, ids = made
    added = doc_of(session).insert_paragraph("Third", after=ids["end"]).id
    later = doc_of(session).insert_paragraph("Fourth", after=added).id
    result = run(box, session, "word_format", doc="d1", targets=[f"{added}..{later}"], like=ids["step2"])
    assert result["ok"], result
    document = doc_of(session)
    assert result["data"]["like"]["set"] == {added: ["style"], later: ["style"]}
    assert style_of(document, added) == style_of(document, ids["step1"]) == "ListNumber"
    assert numbering_of(document, added) is None             # the style numbers it, as it does the sample
    assert labels(document, [ids["step1"], ids["step2"], added, later]) == ["1.", "2.", "3.", "4."]


def test_a_nested_sample_gives_its_level_and_its_own_numbering_only_where_it_has_it(box, made):
    session, ids = made
    document = doc_of(session)
    # A separate list of its own, as a builder's new items came out: another definition.
    stray = document.insert_paragraph("example three", after=ids["example2"]).id
    document.add_to_list(stray, "bullet", level=1, continue_previous=False)
    stray_num = document.list_of(stray).num_id
    result = run(box, session, "word_format", doc="d1", targets=[stray], like=ids["example1"])
    assert result["ok"], result
    document = doc_of(session)
    assert numbering_of(document, stray) == numbering_of(document, ids["example1"]) == ("1", "1")
    assert document.list_of(stray).num_id == document.list_of(ids["example1"]).num_id
    assert stray_num not in document.numbering.nums()         # the list left unused is gone
    assert f"num:{stray_num}" in result["removed"]
    assert labels(document, [ids["example1"], stray]) == ["–", "–"]


def test_indents_and_page_settings_are_taken_whole_and_runs_stay(box, made):
    session, ids = made
    document = doc_of(session)
    target = document.insert_paragraph("example four", after=ids["point2"], style="Plan Body").id
    document.format_paragraph(target, keep_together=True, page_break_before=True, bold=True, alignment="center")
    result = run(box, session, "word_format", doc="d1", targets=target, like=ids["example2"])
    assert result["ok"], result
    document = doc_of(session)
    assert result["data"]["like"]["set"][target] == ["style", "list", "indent", "keep_with_next", "keep_lines",
                                                     "page_break_before"]
    assert "run formatting" in result["data"]["like"]["not_taken"]
    assert own(document, target, "ind") == own(document, ids["example2"], "ind")
    assert own(document, target, "keepNext") is not None
    assert own(document, target, "keepLines") is None and own(document, target, "pageBreakBefore") is None
    assert own(document, target, "jc") is not None                             # alignment not taken
    assert document.paragraph(target).runs[0].bold                             # nor the runs' formatting


def test_explicit_settings_in_the_same_call_win(box, made):
    session, ids = made
    target = doc_of(session).insert_paragraph("Aside", after=ids["end"]).id
    result = run(box, session, "word_format", doc="d1", targets=target, like=ids["example2"],
                 keep_with_next=False, indent_left=10)
    assert result["ok"], result
    document = doc_of(session)
    assert "keep_with_next" not in result["data"]["like"]["set"][target]
    assert own(document, target, "keepNext") == {f"{W}val": "0"} or own(document, target, "keepNext") is None
    assert own(document, target, "ind")[f"{W}left"] == "200"
    assert numbering_of(document, target) == ("1", "1")                         # the list is the sample's
    styled = run(box, session, "word_format", doc="d1", targets=target, like=ids["step1"],
                 paragraph_style="Plan Body")
    assert styled["ok"] and style_of(doc_of(session), target) == "PlanBody"


def test_already_alike_is_no_change_and_the_sample_among_its_targets_is_left(box, made):
    session, ids = made
    version = session.entry("d1").version
    result = run(box, session, "word_format", doc="d1", targets=[f"{ids['point1']}..{ids['point2']}"],
                 like=ids["point1"])
    assert result["ok"], result
    # The examples differ (their level, one's indent and keep); the points are alike.
    assert set(result["data"]["like"]["set"]) == {ids["example1"], ids["example2"]}
    assert result["data"]["like"]["alike"] == [ids["point1"], ids["point2"]]
    again = run(box, session, "word_format", doc="d1", targets=[ids["point2"], ids["example1"]], like=ids["point1"])
    assert again["ok"] and again["changed"] == []
    assert again["warnings"] == [f"every target already looks like {ids['point1']}: nothing changed"]
    assert session.entry("d1").version == version + 1                        # the no-op made no version


def test_a_heading_sample_makes_a_heading_and_says_the_outline_moved(box, made):
    session, ids = made
    result = run(box, session, "word_format", doc="d1", targets=ids["intro"], like=ids["heading"])
    assert result["ok"], result
    document = doc_of(session)
    assert style_of(document, ids["intro"]) == style_of(document, ids["heading"])
    assert any("heading" in w.lower() for w in result["warnings"]), result["warnings"]


def test_a_sample_in_no_list_takes_the_target_out_of_its_list(box, made):
    session, ids = made
    result = run(box, session, "word_format", doc="d1", targets=ids["step2"], like=ids["body"])
    assert result["ok"], result
    document = doc_of(session)
    assert document.list_of(ids["step2"]) is None and style_of(document, ids["step2"]) == "PlanBody"
    assert own(document, ids["step2"], "keepLines") is not None                 # the body's keep, taken


def test_a_table_cell_paragraph_as_the_sample(box, made):
    session, ids = made
    made_table = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{ids['end']}",
                     markdown="| Item | Note |\n|---|---|\n| a | b |\n")
    document = doc_of(session)
    table = made_table["data"]["blocks"][0]
    cell = document.get(table).cell(1, 0).paragraphs[0].id
    document.format_paragraph(cell, indent_left=6, keep_with_next=True)
    result = run(box, session, "word_format", doc="d1", targets=ids["intro"], like=f"{table}/c1,0")
    assert result["ok"], result
    document = doc_of(session)
    assert f"like {table}/c1,0 read as its paragraph {cell}" in result["warnings"]
    assert style_of(document, ids["intro"]) == style_of(document, cell)
    assert own(document, ids["intro"], "ind") == own(document, cell, "ind")


def test_a_sample_of_several_paragraphs_or_none_is_refused_with_the_way_out(box, made):
    session, ids = made
    several = run(box, session, "word_format", doc="d1", targets=ids["intro"],
                  like=f"{ids['point1']}..{ids['point2']}")
    assert not several["ok"] and several["error"]["field"] == "like"
    assert ids["point1"] in several["error"]["valid_options"]
    assert f"like: {ids['point1']!r}" in several["error"]["message"]
    missing = run(box, session, "word_format", doc="d1", targets=ids["intro"], like="p:0BADBEEF")
    assert not missing["ok"] and missing["error"]["field"] == "like"
    ranged = run(box, session, "word_format", doc="d1", targets=ids["intro"], like=f"{ids['point1']}@0:3")
    assert ranged["ok"] and f"like {ids['point1']}@0:3 read as its paragraph {ids['point1']}" in ranged["warnings"]


def test_tracked_like_is_a_paragraph_property_change_rejected_back(box, made):
    session, ids = made
    target = doc_of(session).insert_paragraph("Tracked", after=ids["end"]).id
    assert run(box, session, "word_set_tracking", doc="d1", on=True, author="Lane")["ok"]
    result = run(box, session, "word_format", doc="d1", targets=target, like=ids["example1"])
    assert result["ok"], result
    document = doc_of(session)
    element = document._resolve(target)[1].element
    assert element.find(f"{W}pPr/{W}pPrChange") is not None
    assert numbering_of(document, target) == ("1", "1")
    listed = run(box, session, "word_changes", doc="d1", action="reject", all=True)
    assert listed["ok"], listed
    assert doc_of(session).list_of(target) is None


def test_one_undo_step(box, made):
    session, ids = made
    before = doc_of(session).to_bytes()
    target = ids["intro"]
    assert run(box, session, "word_format", doc="d1", targets=target, like=ids["point1"])["ok"]
    assert run(box, session, "undo", doc="d1")["ok"]
    assert doc_of(session).to_bytes() == before


def test_drawn_as_the_samples_items_are(box, made):
    session, ids = made
    document = doc_of(session)
    point = document.insert_paragraph("Third point", after=ids["point2"]).id
    example = document.insert_paragraph("example five", after=point).id
    assert run(box, session, "word_format", doc="d1", targets=point, like=ids["point1"])["ok"]
    assert run(box, session, "word_format", doc="d1", targets=example, like=ids["example1"])["ok"]
    document = doc_of(session)
    assert drawn(document, point)[0] == drawn(document, ids["point1"])[0]          # the marker and its place
    assert drawn(document, point)[1][1] == drawn(document, ids["point1"])[1][1]    # the text's place
    assert drawn(document, example)[0] == drawn(document, ids["example1"])[0]
    assert drawn(document, example)[1][1] == drawn(document, ids["example1"])[1][1]


# -- the way out of removing an item's own list --------------------------------------------------


def test_removing_a_stray_lists_numbering_offers_joining_the_list_keeping_the_markers(box, made):
    session, ids = made
    document = doc_of(session)
    strays = []
    after = ids["point2"]
    for text, level in (("Third point", 0), ("example six", 1)):
        after = document.insert_paragraph(text, after=after, style="List Bullet").id
        document.add_to_list(after, "bullet", level=level, continue_previous=False)
        strays.append(after)
    refused = run(box, session, "word_lists", doc="d1", targets=strays, action="remove")
    assert not refused["ok"]
    message = refused["error"]["message"]
    assert "join an earlier list at their level" in message and f'"like": "{ids["example2"]}"' in message
    assert f"word_format like {ids['example2']}" in refused["error"]["valid_options"]
    # The call it names keeps the marker at the item's level.
    assert run(box, session, "word_format", doc="d1", targets=[strays[1]], like=ids["example2"])["ok"]
    assert labels(doc_of(session), [strays[1]]) == ["–"]


# -- word_insert_markdown like -------------------------------------------------------------------


def test_new_nested_items_follow_the_sample_at_each_level(box, made):
    session, ids = made
    result = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{ids['point2']}",
                 markdown="- Third point\n  - example seven\n  - example eight\n- Fourth point\n",
                 like=[ids["point1"], ids["example1"]])
    assert result["ok"], result
    document = doc_of(session)
    new = result["data"]["blocks"]
    assert [style_of(document, i) for i in new] == ["ListBullet"] * 4
    assert [numbering_of(document, i) for i in new] == [None, ("1", "1"), ("1", "1"), None]
    assert labels(document, new) == ["•", "–", "–", "•"]
    assert {document.list_of(i).num_id for i in new} == {document.list_of(ids["point1"]).num_id}
    assert drawn(document, new[1])[0] == drawn(document, ids["example1"])[0]
    assert drawn(document, new[1])[1][1] == drawn(document, ids["example1"])[1][1]


def test_a_level_below_the_given_ones_is_the_list_a_level_deeper(box, made):
    session, ids = made
    result = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{ids['point2']}",
                 markdown="- Third point\n  - example\n    - detail\n", like=ids["point1"])
    assert result["ok"], result
    document = doc_of(session)
    new = result["data"]["blocks"]
    assert [document.list_of(i).level for i in new] == [0, 1, 2]
    assert labels(document, new) == ["•", "–", "◦"]
    assert [style_of(document, i) for i in new] == ["ListBullet"] * 3        # Word's Tab: the style stays


def test_numbered_items_count_on_from_the_sample(box, made):
    session, ids = made
    result = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{ids['step2']}",
                 markdown="1. Cross\n2. Rest\n", like=[ids["step1"]])
    assert result["ok"], result
    document = doc_of(session)
    assert labels(document, [ids["step2"], *result["data"]["blocks"]]) == ["2.", "3.", "4."]


def test_a_style_map_naming_a_paragraph_is_read_as_like(box, made):
    session, ids = made
    result = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{ids['point2']}",
                 markdown="- Third point\n  - example nine\n",
                 style_map=[{"element": "bullet2", "style": ids["example1"]}])
    assert result["ok"], result
    assert any("read as like" in w for w in result["warnings"])
    document = doc_of(session)
    new = result["data"]["blocks"]
    assert numbering_of(document, new[1]) == ("1", "1") and labels(document, new) == ["•", "–"]


def test_an_empty_level_is_written_without_like(box, made):
    session, ids = made
    result = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{ids['end']}",
                 markdown="- top\n  - nested\n", like=["", ids["example1"]])
    assert result["ok"], result
    document = doc_of(session)
    new = result["data"]["blocks"]
    assert numbering_of(document, new[1]) == ("1", "1")
    assert style_of(document, new[0]) == "ListBullet" and numbering_of(document, new[0]) is None


def test_a_sample_in_no_list_and_another_kind_are_said(box, made):
    session, ids = made
    plain = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{ids['end']}",
                markdown="- one\n", like=ids["body"])
    assert plain["ok"] and any("is in no list" in w for w in plain["warnings"])
    assert doc_of(session).list_of(plain["data"]["blocks"][0]) is None
    other = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{ids['end']}",
                markdown="1. one\n", like=ids["point1"])
    assert other["ok"] and any("are numbers and" in w for w in other["warnings"])


def test_replacing_the_sample_itself_keeps_its_look(box, made):
    session, ids = made
    result = run(box, session, "word_insert_markdown", doc="d1",
                 at=f"replace:{ids['example1']}..{ids['example2']}",
                 markdown="- example ten\n- example eleven\n", like=ids["example2"])
    assert result["ok"], result
    document = doc_of(session)
    new = result["data"]["blocks"]
    assert [numbering_of(document, i) for i in new] == [("1", "1")] * 2
    assert own(document, new[0], "keepNext") is not None
    assert labels(document, [ids["point1"], *new, ids["point2"]]) == ["•", "–", "–", "•"]
    assert _numbering.problems(document) == []


def test_a_copied_sample_extended_follows_the_copy(box, made):
    session, ids = made
    copied = run(box, session, "word_copy_from", doc="d1", source_doc="d1",
                 range=f"{ids['point1']}..{ids['point2']}", at=f"after:{ids['end']}")
    assert copied["ok"], copied
    document = doc_of(session)
    copy = copied["data"]["blocks"]
    last_point, first_example = copy[-1], copy[1]
    result = run(box, session, "word_insert_markdown", doc="d1", at=f"after:{last_point}",
                 markdown="- Third point\n  - example twelve\n", like=[copy[0], first_example])
    assert result["ok"], result
    document = doc_of(session)
    new = result["data"]["blocks"]
    # Each level as the copy's item of that level is (the copy's own lists, as it was copied).
    assert [numbering_of(document, i) for i in new] == [numbering_of(document, copy[0]),
                                                         numbering_of(document, first_example)]
    assert [document.list_of(i).num_id for i in new] == [document.list_of(copy[0]).num_id,
                                                          document.list_of(first_example).num_id]
    assert [document.list_of(i).level for i in new] == [0, 1]
    assert drawn(document, new[0])[0] == drawn(document, copy[0])[0]
    assert drawn(document, new[1])[1][1] == drawn(document, first_example)[1][1]


def test_like_wins_over_the_replaced_items_kept_settings(box, made):
    session, ids = made
    # The replaced item is a nested example keeping with the next; like names a top-level point.
    result = run(box, session, "word_insert_markdown", doc="d1", at=f"replace:{ids['example2']}",
                 markdown="- Second example, now a point\n", like=ids["point1"])
    assert result["ok"], result
    document = doc_of(session)
    new = result["data"]["blocks"][0]
    assert numbering_of(document, new) is None and document.list_of(new).level == 0
    assert own(document, new, "keepNext") is None and own(document, new, "ind") is None
    assert "kept" not in result["data"]
