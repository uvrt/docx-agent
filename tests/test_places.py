"""One way to name a place across the placing tools (round 40, :mod:`docx_agent.tools.positions`).

``at`` is the form: ``end``, ``after:<id>``, ``before:<id>``, ``replace:<id>``, a position
``p:A@5``, or a bare id (after it).  Another tool's name for a place is read as the tool's
own where renaming says the same thing (a warning says so); where it would not -- the side
is in the name (``before``, ``where``) -- the call is refused and ``details`` holds the
arguments that work, which these tests dispatch as given.  Made-up text only."""

from __future__ import annotations

import datetime as dt

import pytest

from ooxml_edit.tools import Toolbox

from docx_agent import Document
from docx_agent.tools import FORMAT, GROUPS, TOOLS
from docx_agent.tools import positions

CLOCK = lambda: dt.datetime(2026, 11, 7, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def made(toolbox):
    session = toolbox.session(clock=CLOCK)
    session.open(Document.new().to_bytes(), "plan.docx")
    result = toolbox.dispatch(session, "word_insert_markdown", {
        "doc": "d1", "at": "end", "markdown": "Opening line.\n\nMiddle line.\n\nClosing line."})
    assert result.ok, result.to_json()
    return session, result.data["blocks"]


def run(toolbox, session, name, **arguments):
    return toolbox.dispatch(session, name, arguments)


def ok(result):
    assert result.ok, result.to_json()
    return result


def texts(session) -> list[str]:
    document = session.entry("d1").document
    return [p.text for p in document.paragraphs()]


def redo(toolbox, session, error):
    """Dispatch the call a refusal gives (``details.call``)."""
    meant = error.details["call"]
    return ok(toolbox.dispatch(session, meant["tool"], meant["arguments"]))


# -- every placing tool reads the names it lacks -------------------------------------------------


PLACING = {"word_insert_text": "at", "word_insert_markdown": "at", "word_copy_from": "at",
           "word_drawings": "at", "word_controls": "at", "word_notes": "at", "word_fields": "at",
           "word_move": "after", "word_sections": "after"}


@pytest.mark.parametrize("name", sorted(PLACING))
def test_each_placing_tool_takes_its_own_form_and_reads_the_others(name):
    tool = next(t for t in TOOLS if t.name == name)
    params = set(tool.schema["properties"])
    assert PLACING[name] in params
    read = {alias for alias in tool.aliases if "." not in alias}
    lacking = {n for n in positions.NAMES if n not in params}
    if name in ("word_move", "word_sections"):
        lacking -= {"target", "where"}              # their target is what moves, or a section
    if "find" in params:
        lacking.discard("find")
    assert lacking <= read, lacking - read


def test_the_strict_set_has_room_for_insert_markdowns_ref_and_no_more():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        strict = [d for d in box.definitions("anthropic") if d.get("strict")]
    optional = sum(len(set(d["input_schema"]["properties"]) - set(d["input_schema"].get("required", ())))
                   for d in strict)
    assert len(strict) == 6 and optional <= 24
    assert "word_insert_markdown" in {d["name"] for d in strict}


# -- blocks: word_insert_markdown, word_copy_from ------------------------------------------------


def test_after_and_a_target_alone_are_read_as_at(toolbox, made):
    session, (opening, middle, closing) = made
    result = ok(run(toolbox, session, "word_insert_markdown", doc="d1", after=opening, markdown="One."))
    assert result.warnings[0] == "after read as at"
    result = ok(run(toolbox, session, "word_insert_markdown", doc="d1", target=middle, markdown="Two."))
    assert result.warnings[0] == "target read as at"
    result = ok(run(toolbox, session, "word_insert_markdown", doc="d1", position="end", markdown="Three."))
    assert texts(session) == ["Opening line.", "One.", "Middle line.", "Two.", "Closing line.", "Three."]


@pytest.mark.parametrize("arguments, at", [
    ({"before": "{middle}"}, "before:{middle}"),
    ({"target": "{middle}", "where": "before"}, "before:{middle}"),
    ({"target": "{middle}", "where": "after"}, "{middle}"),
])
def test_a_place_whose_side_is_in_its_name_is_refused_with_the_call_that_works(toolbox, made, arguments, at):
    session, (opening, middle, closing) = made
    arguments = {k: v.format(middle=middle) for k, v in arguments.items()}
    result = run(toolbox, session, "word_insert_markdown", doc="d1", markdown="New.", **arguments)
    error = result.error
    assert error.code == "invalid_arguments" and error.valid_options == ["at"]
    assert "word_insert_markdown takes at" in error.message
    assert error.details["call"]["arguments"] == {"doc": "d1", "markdown": "New.", "at": at.format(middle=middle)}
    assert texts(session) == ["Opening line.", "Middle line.", "Closing line."]
    redo(toolbox, session, error)
    new = texts(session).index("New.")
    assert texts(session)[new + (1 if at.startswith("before") else -1)] == "Middle line."


def test_find_where_a_tool_places_by_id_names_find_text(toolbox, made):
    session, _ = made
    error = run(toolbox, session, "word_insert_markdown", doc="d1", find="Middle", markdown="x").error
    assert error.code == "invalid_arguments" and "find_text" in error.message and error.field == "find"


def test_ref_names_the_blocks_made_and_a_span_is_placed_beside(toolbox, made):
    session, (opening, middle, closing) = made
    one = ok(run(toolbox, session, "word_insert_markdown", doc="d1", at=f"after:{opening}",
                 markdown="A.\n\nB.", ref="pair"))
    first, last = one.data["blocks"]
    assert one.refs == {"pair": f"{first}..{last}"}
    ok(run(toolbox, session, "word_insert_markdown", doc="d1", at="after:$pair", markdown="After pair."))
    ok(run(toolbox, session, "word_insert_markdown", doc="d1", at="before:$pair", markdown="Before pair."))
    assert texts(session)[0:5] == ["Opening line.", "Before pair.", "A.", "B.", "After pair."]
    ok(run(toolbox, session, "word_format", doc="d1", targets=["$pair"], bold=True))
    single = ok(run(toolbox, session, "word_insert_markdown", doc="d1", at="end", markdown="Last.", ref="tail"))
    assert single.refs == {"tail": single.data["blocks"][0]}


def test_an_address_given_as_ref_is_refused_naming_at(toolbox, made):
    session, (opening, middle, closing) = made
    error = run(toolbox, session, "word_insert_markdown", doc="d1", at="end", ref=middle, markdown="x").error
    assert error.field == "ref" and f"at: 'after:{middle}'" in error.message
    assert texts(session) == ["Opening line.", "Middle line.", "Closing line."]


def test_copy_from_reads_after_and_names_its_blocks(toolbox, made):
    session, (opening, middle, closing) = made
    result = ok(run(toolbox, session, "word_copy_from", doc="d1", source_doc="d1", range=closing, after=opening,
                    ref="copy"))
    assert result.warnings[0] == "after read as at" and result.refs == {"copy": result.data["blocks"][0]}
    assert texts(session)[0:2] == ["Opening line.", "Closing line."]
    error = run(toolbox, session, "word_copy_from", doc="d1", source_doc="d1", range=closing, before=opening).error
    redo(toolbox, session, error)
    assert texts(session)[0:2] == ["Closing line.", "Opening line."]


# -- text: word_insert_text ---------------------------------------------------------------------


def test_insert_text_takes_at_beside_target_and_where(toolbox, made):
    session, (opening, middle, closing) = made
    ok(run(toolbox, session, "word_insert_text", doc="d1", at=f"before:{opening}@0:7", text="The "))
    ok(run(toolbox, session, "word_insert_text", doc="d1", at=f"{middle}@6", text=" quiet"))
    ok(run(toolbox, session, "word_insert_text", doc="d1", target=f"{closing}@0:7", where="before", text="A "))
    result = ok(run(toolbox, session, "word_insert_text", doc="d1", items=[
        {"at": f"after:{closing}@15", "text": "!"}, {"after": f"{opening}@0:3", "text": ","}]))
    assert result.warnings == ["items[1].after read as at"]
    assert texts(session) == ["The, Opening line.", "Middle quiet line.", "A Closing line.!"]
    error = run(toolbox, session, "word_insert_text", doc="d1", at=f"after:{opening}@0", where="before",
                text="x").error
    assert error.field == "where"


def test_insert_text_before_is_refused_with_the_call_and_the_item(toolbox, made):
    session, (opening, middle, closing) = made
    error = run(toolbox, session, "word_insert_text", doc="d1", before=f"{middle}@0", text="Very ").error
    assert error.details["call"]["arguments"] == {"doc": "d1", "text": "Very ", "at": f"before:{middle}@0"}
    redo(toolbox, session, error)
    assert texts(session)[1] == "Very Middle line."
    error = run(toolbox, session, "word_insert_text", doc="d1", items=[{"before": f"{closing}@0", "text": "x"}]).error
    assert error.field == "items[].before" and error.details["item"] == {"text": "x", "at": f"before:{closing}@0"}


def test_insert_text_ref_holding_an_address_and_no_place(toolbox, made):
    session, (opening, middle, closing) = made
    error = run(toolbox, session, "word_insert_text", doc="d1", ref=f"{middle}@6", text="!").error
    assert error.field == "ref" and error.valid_options == ["at"]
    redo(toolbox, session, error)
    assert texts(session)[1] == "Middle! line."


# -- marks: a note, a field, a content control ------------------------------------------------


def test_marks_read_after_as_the_end_and_before_as_the_start(toolbox, made):
    session, (opening, middle, closing) = made
    result = ok(run(toolbox, session, "word_notes", doc="d1", action="insert", after=middle, text="A note."))
    assert result.warnings[0] == "after read as at"
    error = run(toolbox, session, "word_controls", doc="d1", action="insert", before=f"{closing}@8:12").error
    assert error.details["call"]["arguments"]["at"] == f"{closing}@8"
    redo(toolbox, session, error)
    ok(run(toolbox, session, "word_fields", doc="d1", action="insert_date", after=opening))
    document = session.entry("d1").document
    assert len(document.notes()) == 1 and len(document.content_controls()) == 1 and len(document.fields()) == 1


def test_a_toc_or_caption_takes_at_as_the_placing_tools_write_it(toolbox, made):
    session, (opening, middle, closing) = made
    ok(run(toolbox, session, "word_fields", doc="d1", action="insert_caption", at=f"before:{closing}",
           text="Made up"))
    assert texts(session)[2].startswith("Figure") and texts(session)[3] == "Closing line."


# -- sides: word_move, word_sections ------------------------------------------------------------


@pytest.mark.parametrize("at, order", [
    ("before:{opening}", ["Closing line.", "Opening line.", "Middle line."]),
    ("after:{opening}", ["Opening line.", "Closing line.", "Middle line."]),
    ("{opening}", ["Opening line.", "Closing line.", "Middle line."]),
])
def test_move_reads_at_with_its_side(toolbox, made, at, order):
    session, (opening, middle, closing) = made
    result = ok(run(toolbox, session, "word_move", doc="d1", blocks=closing, at=at.format(opening=opening)))
    assert result.warnings[0] == "at read as after"
    assert texts(session) == order


def test_move_to_the_end_and_a_replace_refused(toolbox, made):
    session, (opening, middle, closing) = made
    ok(run(toolbox, session, "word_move", doc="d1", blocks=opening, position="end"))
    assert texts(session)[-1] == "Opening line."
    error = run(toolbox, session, "word_move", doc="d1", blocks=opening, at=f"replace:{middle}").error
    assert error.code == "invalid_arguments" and error.valid_options == ["after", "before"]


def test_a_section_break_before_a_block(toolbox, made):
    session, (opening, middle, closing) = made
    ok(run(toolbox, session, "word_sections", doc="d1", action="insert_break", at=f"before:{closing}"))
    sections = ok(run(toolbox, session, "word_sections", doc="d1", action="list")).data
    assert len(sections) == 2


# -- batches ------------------------------------------------------------------------------------


def test_a_batch_reads_the_aliases_and_refuses_naming_the_op(toolbox, made):
    session, (opening, middle, closing) = made
    ops = [{"tool": "word_insert_markdown", "arguments": {"doc": "d1", "after": opening, "markdown": "One."}},
           {"tool": "word_insert_text", "arguments": {"doc": "d1", "after": f"{closing}@13", "text": "!"}}]
    result = ok(run(toolbox, session, "batch", ops=ops, atomic=True))
    assert "ops[0]: after read as at" in result.warnings
    ops.append({"tool": "word_insert_markdown", "arguments": {"doc": "d1", "before": closing, "markdown": "Two."}})
    result = run(toolbox, session, "batch", ops=ops, atomic=True)
    assert not result.ok and result.error.field == "ops[2].arguments.before"
    assert result.error.details["call"]["arguments"]["at"] == f"before:{closing}"
    assert texts(session).count("One.") == 1                      # the atomic batch changed nothing


def test_existing_callers_are_unchanged(toolbox, made):
    session, (opening, middle, closing) = made
    ok(run(toolbox, session, "word_insert_text", doc="d1", target=f"{middle}@0:6", where="after", text="!"))
    ok(run(toolbox, session, "word_insert_text", doc="d1", find="Closing", text="The ", where="before"))
    ok(run(toolbox, session, "word_insert_markdown", doc="d1", at=f"replace:{opening}", markdown="Replaced."))
    ok(run(toolbox, session, "word_move", doc="d1", blocks=closing, before=middle))
    assert texts(session) == ["Replaced.", "The Closing line.", "Middle! line."]
