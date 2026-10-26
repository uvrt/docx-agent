"""Which error code a tool call's failure gets (review, S9).

ooxml-edit 0.66.0 types its refusals (``ooxml_edit.errors``): its dispatcher gives an
``ArgumentError`` ``invalid_arguments`` and a ``UnitError`` (a value out of its range)
``unit``, in every format.  Before, every plain ``ValueError`` was ``unit`` -- removing a
diagram's last node among them.  The library's own errors keep the codes ``ERRORS`` gives
them.
"""

from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("ooxml_edit.errors", reason="needs ooxml-edit 0.66 (typed refusals)")

from ooxml_edit.errors import ArgumentError, UnitError  # noqa: E402
from ooxml_edit.tools import Toolbox  # noqa: E402

from docx_agent.edit.errors import EditError  # noqa: E402
from docx_agent.edit.formatting import FormattingError  # noqa: E402
from docx_agent.edit.inline import SpanError  # noqa: E402
from docx_agent.edit.pictures import PictureError  # noqa: E402
from docx_agent.edit.ranges import AmbiguousAnchor, AnchorNotFound  # noqa: E402
from docx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402
from docx_agent.tools.format import ERRORS  # noqa: E402

SMARTART = Path(__file__).parent / "fixtures" / "generated" / "charts" / "smartart.docx"


@pytest.fixture(scope="module")
def box():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        yield toolbox


def test_removing_the_last_node_is_invalid_arguments(box):
    session = box.session()
    doc = session.open(SMARTART.read_bytes(), SMARTART.name)
    for _ in range(10):
        result = box.dispatch(session, "edit_smartart", {
            "doc": doc, "target": "d:1", "action": "remove_node", "node": 0})
        if not result.ok:
            break
    assert result.error.code == "invalid_arguments"
    assert result.error.message == "d:1: a diagram keeps at least one node"


@pytest.mark.parametrize("error, code", [
    (ArgumentError("a diagram keeps at least one node"), "invalid_arguments"),
    (UnitError("text is longer than a cell holds"), "unit"),
    # The library's own, as ERRORS maps them (unchanged by S9).
    (FormattingError("size is in points, 1 to 1638"), "unit"),
    (SpanError("no span"), "refused"),
    (PictureError("no picture"), "refused"),
    (EditError("not this way"), "refused"),
    (AnchorNotFound("nowhere"), "not_found"),
    (AmbiguousAnchor("twice", []), "ambiguous"),
    # Python's own ValueError is still unit; its TypeError internal.
    (ValueError("invalid literal"), "unit"),
    (TypeError("unsupported operand"), "internal"),
])
def test_the_codes(box, error, code):
    assert box.map_exception(error).code == code


def test_every_mapped_error_is_tested():
    tested = {FormattingError, SpanError, PictureError, EditError, AnchorNotFound,
              AmbiguousAnchor}
    assert set(ERRORS) == tested
