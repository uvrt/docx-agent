"""Round 42, R49: ``check(copies=[...])`` compares each copy's list items -- level, label,
numbering, indents -- and every paragraph's page settings with its reference's, as Word
applies them (:mod:`docx_agent.edit.copy_settings`): a dash among round bullets, an item a
level up, a lost keep are each told, with the call that fixes it; copies that look the same
as their sample are not, however their XML says it.  Every document here is made up."""

from __future__ import annotations

import pytest

pytest.importorskip("ooxml_edit.tools.shared", reason="needs ooxml-edit (the shared tools)")

from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent import Document  # noqa: E402
from docx_agent.edit import copies as copies_module  # noqa: E402
from docx_agent.edit import copy_settings  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS, scoped  # noqa: E402

pytestmark = pytest.mark.any_faces

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

SAMPLE = ("# Part title\n\nDescribe the part here, in a few sentences.\n\n"
          "- First point of the part\n    - A detail under it\n- Second point of the part\n\n"
          "1. Step one of the part\n2. Step two of the part\n\nClose with a short summary of the part.\n")


@pytest.fixture
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def _ok(box, session, tool: str, /, **arguments):
    result = box.dispatch(session, tool, arguments)
    assert result.ok, (tool, arguments, result.error and result.error.to_json())
    return result


def _sample() -> tuple[bytes, list[str]]:
    sample = Document.new()
    sample.insert_markdown(SAMPLE, at="end")
    blocks = [b["id"] for b in sample.blocks()]
    sample.paragraph(blocks[1]).format(keep_with_next=True)     # the description stays with its list
    return sample.to_bytes(), blocks


def _made(box, session, count: int, *, same_document: bool = False) -> tuple[str, str, str, list[list[str]]]:
    """A report of ``count`` copies of the sample: ``(report doc, sample doc, sample span, copies)``.
    ``same_document``: the sample is the report's own first chapter."""
    data, blocks = _sample()
    reference = f"{blocks[0]}..{blocks[-1]}"
    samples = session.open(data, "samples.docx")
    if same_document:
        d1 = samples
    else:
        report = Document.new()
        report.insert_markdown("# Report\n\nAn opening paragraph.", at="end")
        d1 = session.open(report.to_bytes(), "report.docx")
    made = [_ok(box, session, "word_copy_from", doc=d1, source_doc=samples, range=reference,
                at="end").data["blocks"] for _ in range(count)]
    return d1, samples, reference, made


def _pairs(made: list[list[str]], samples: str, reference: str) -> list[dict]:
    return [{"blocks": f"{m[0]}..{m[-1]}", "reference": {"doc": samples, "block": reference}} for m in made]


def _restart_numbers(box, session, doc: str, made: list[list[str]]) -> None:
    """Each copy's numbered list starts at 1 (copied by style, it would go on counting)."""
    for blocks in made[1:]:
        _ok(box, session, "word_lists", doc=doc, targets=[blocks[5]], action="restart", start=1)


def _check(box, session, doc: str, made, samples, reference) -> dict:
    return _ok(box, session, "check", doc=doc, include=[], copies=_pairs(made, samples, reference)).data["copies"]


def _settings(item: dict) -> dict[str, dict]:
    return {d["property"]: d for d in item["differences"] if d["property"] in copy_settings.SETTINGS}


def test_dashes_a_wrong_level_and_a_lost_keep_are_each_found_with_their_fix(box):
    session = box.session()
    d1, samples, reference, made = _made(box, session, 4)
    _restart_numbers(box, session, d1, made)
    assert _check(box, session, d1, made, samples, reference)["follow"] == 4
    # Copy 1: its first point made a list of its own, of dashes.
    first = made[1][2]
    _ok(box, session, "word_lists", doc=d1, targets=[first], action="add", kind="bullet")
    _ok(box, session, "word_lists", doc=d1, targets=[first], action="format", level=0, number_format="bullet",
        label="–")
    # Copy 2: the detail a level up.  Copy 3: the description's keep lost.
    _ok(box, session, "word_lists", doc=d1, targets=[made[2][3]], action="level", level=0)
    _ok(box, session, "word_format", doc=d1, targets=[made[3][1]], keep_with_next=False)
    found = _check(box, session, d1, made, samples, reference)
    items = found["items"]
    assert [item["follows"] for item in items] == [True, False, False, False]
    dashes = _settings(items[1])["list_format"]
    assert dashes["value"] == "bullet –" and dashes["reference"] == "bullet •"
    assert dashes["paragraphs"] == [first] and dashes["kind"] == "format"
    level = _settings(items[2])["list_level"]
    assert (level["value"], level["reference"], level["paragraphs"]) == (0, 1, [made[2][3]])
    assert level["fix"] == {"tool": "word_lists", "calls": [{"targets": [made[2][3]], "action": "level",
                                                             "level": 1}]}
    keep = _settings(items[3])["keep_with_next"]
    assert (keep["value"], keep["reference"]) == (False, True)
    assert keep["fix"] == {"tool": "word_format", "calls": [{"targets": [made[3][1]], "keep_with_next": True}]}
    # One difference per cause: neither the dashes' nor the level's indents are told apart.
    assert set(_settings(items[1])) == {"list_format"} and set(_settings(items[2])) == {"list_level"}
    # Rule 5: the list fixes of every copy in one list of calls; follow_sample for the rest.
    assert found["list_fix"]["tool"] == "word_lists"
    assert {c["action"] for c in found["list_fix"]["calls"]} == {"format", "level"}
    assert found["fix"]["calls"] == [{"sample": {"block": _blocks_of(session, samples, reference)[1],
                                                 "doc": samples}, "copies": [made[3][1]]}]
    # The fixes make every copy follow again.
    for call in found["list_fix"]["calls"]:
        _ok(box, session, "word_lists", doc=d1, **call)
    for call in keep["fix"]["calls"]:
        _ok(box, session, "word_format", doc=d1, **call)
    again = _check(box, session, d1, made, samples, reference)
    assert again["follow"] == 4, [i["differences"] for i in again["items"]]


def _blocks_of(session, doc: str, span: str) -> list[str]:
    return [i for i, _ in copies_module.blocks_of(session.entry(doc).document, span)]


def test_copies_that_look_the_same_do_not_differ(box):
    """The same look stated otherwise: a keep the heading's style gives, stated again; a
    list item's indent stated as its list gives it; the bullets restarted (a new instance,
    the same bullets); an item in another list that draws the same."""
    session = box.session()
    d1, samples, reference, made = _made(box, session, 3)
    _restart_numbers(box, session, d1, made)
    document = session.entry(d1).document
    heading, point, detail = made[0][0], made[0][2], made[0][3]
    _ok(box, session, "word_format", doc=d1, targets=[heading], keep_with_next=True)
    from docx_agent.edit import effective

    applied = effective.paragraph(document, document._resolve(point)[1].element)
    _ok(box, session, "word_format", doc=d1, targets=[point], indent_left=applied.indent_left,
        indent_first=applied.first_line - applied.hanging)
    _ok(box, session, "word_lists", doc=d1, targets=[made[1][2]], action="restart", start=1)
    # Copy 2's detail: a new list of Word's own bullets, the same "o" at the same place.
    document.add_to_list(made[2][3], "bullet", level=1, continue_previous=False)
    found = _check(box, session, d1, made, samples, reference)
    assert found["follow"] == 3, [item["differences"] for item in found["items"]]
    assert "list_fix" not in found and "fix" not in found
    # Their XML differs all the same.
    assert document.list_of(made[2][3]).num_id != document.list_of(detail).num_id
    assert document.list_of(made[1][2]).num_id != document.list_of(made[0][2]).num_id
    states = [document._resolve(i)[1].element.find(f"{W}pPr/{W}{tag}") is not None
              for i, tag in ((heading, "keepNext"), (made[1][0], "keepNext"), (point, "ind"), (made[1][2], "ind"))]
    assert states == [True, False, True, False]


def test_a_numbered_list_going_on_from_the_chapter_before_is_told(box):
    """Copies of a numbered list in its style count on from the copy before them, as Word
    numbers them: told as where each starts, fixed by a restart."""
    session = box.session()
    d1, samples, reference, made = _made(box, session, 3)
    found = _check(box, session, d1, made, samples, reference)
    numbering = [_settings(item).get("list_numbering") for item in found["items"]]
    assert numbering[0] is None
    assert [(n["value"], n["reference"]) for n in numbering[1:]] == [("starts at 3", "starts at 1"),
                                                                     ("starts at 5", "starts at 1")]
    assert numbering[1]["fix"] == {"tool": "word_lists", "calls": [
        {"targets": [made[1][5]], "action": "restart", "start": 1}]}
    for call in found["list_fix"]["calls"]:
        _ok(box, session, "word_lists", doc=d1, **call)
    assert _check(box, session, d1, made, samples, reference)["follow"] == 3


def test_a_kept_comparison_follows_the_numbers_before_it(box):
    """A comparison is kept by what both sides hold -- and by the numbers the copy's lists
    show, which the items before the copy decide: an item added before it compares it again."""
    session = box.session()
    d1, samples, reference, made = _made(box, session, 2)
    _restart_numbers(box, session, d1, made)
    pairs = _pairs(made[1:], samples, reference)
    first = _ok(box, session, "check", doc=d1, include=[], copies=pairs).data["copies"]
    assert first["follow"] == 1
    kept = _ok(box, session, "check", doc=d1, include=[], copies=pairs).data["copies"]
    assert kept["kept"] == 1
    # The second copy's numbers restarted from the first's: undo it, its list goes on again.
    _ok(box, session, "word_lists", doc=d1, targets=[made[1][5]], action="continue")
    again = _ok(box, session, "check", doc=d1, include=[], copies=pairs).data["copies"]
    assert again["kept"] == 0 and again["follow"] == 0


def test_the_sample_in_the_same_document(box):
    """The sample the report's own chapter: its copies' lists by style go on from it."""
    session = box.session()
    d1, _, reference, made = _made(box, session, 1, same_document=True)
    found = _ok(box, session, "check", doc=d1, include=[], copies=[
        {"blocks": f"{made[0][0]}..{made[0][-1]}", "reference": {"block": reference}}]).data["copies"]
    numbering = _settings(found["items"][0])["list_numbering"]
    assert (numbering["value"], numbering["reference"]) == ("starts at 3", "starts at 1")


def test_the_list_fix_keeps_to_its_budget(box, monkeypatch):
    """Rule 4: past the budget the list fix's calls are counted."""
    monkeypatch.setattr(scoped, "LIST_CALLS", 1)
    session = box.session()
    d1, samples, reference, made = _made(box, session, 3)
    found = _check(box, session, d1, made, samples, reference)
    assert len(found["list_fix"]["calls"]) == 1 and found["list_fix"]["more"] == 1


def test_clear_is_the_fix_where_the_style_gives_the_reference(box):
    """Rule 6: a keep the copy turned off directly, where its style keeps it: cleared."""
    session = box.session()
    d1, samples, reference, made = _made(box, session, 1)
    _ok(box, session, "word_format", doc=d1, targets=[made[0][0]], keep_with_next=False)   # the heading's
    found = _check(box, session, d1, made, samples, reference)
    keep = _settings(found["items"][0])["keep_with_next"]
    assert keep["fix"] == {"tool": "word_format", "calls": [{"targets": [made[0][0]], "clear": ["keep_with_next"]}]}
