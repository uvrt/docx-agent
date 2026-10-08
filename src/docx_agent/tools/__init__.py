"""Word tools for an agent tool layer (``ooxml_edit.tools``): definitions and handlers.

An application puts these in a :class:`ooxml_edit.tools.Toolbox` with the docx
:data:`FORMAT` (and, beside them, another library's tools for its documents)::

    from ooxml_edit.tools import Toolbox
    from docx_agent.tools import TOOLS, FORMAT, GROUPS

    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    session = toolbox.session(clock=my_clock)
    d1 = session.open(docx_bytes, name="agreement.docx")
    tools = toolbox.definitions("anthropic", groups="core")
    result = toolbox.dispatch(session, "describe", {"doc": d1})

Everything is in memory: documents and inputs are bytes under handles, saved files go to
the application.  Lengths are points.  Every mutating call is one undo step, written as
tracked changes when ``word_set_tracking`` turned tracking on, and returns ``checks``: the
``validate()`` delta against the document as opened, and the pages it reflowed.  Layout
runs in the toolbox's worker pool under ``Limits.layout_timeout`` (30 s).

The shared tools (``open_document``, ``save_document``, ``find_text``, ``render``...) are
one definition for every format (``ooxml_edit.tools.shared``); this package gives their
Word handlers.
"""

from __future__ import annotations

from ooxml_edit.tools import CORE, ToolGroup, shared as _shared

from .format import FORMAT, PROMPT
from . import read, text, review, structure, objects, style

GROUPS = [
    *_shared.GROUPS,
    ToolGroup("word_text", "Word: inspect formatting, insert and delete text, write Markdown."),
    ToolGroup("word_review", "Word: tracking, tracked changes (list, accept, reject), comments."),
    ToolGroup("word_structure", "Word: move and copy sections, page setup, headers and footers, "
              "fields and TOC, notes, links and bookmarks."),
    ToolGroup("word_objects", "Word: tables, pictures and other drawings, text boxes, content controls."),
    ToolGroup("word_style", "Word: formatting text and paragraphs, lists, the style sheet, templates."),
]

#: Every Word-specific tool, in the roadmap's order.
WORD_TOOLS = (read.TOOLS + text.TOOLS + review.TOOLS + structure.TOOLS + objects.TOOLS
              + style.TOOLS)

from . import shared  # noqa: E402  (the Word handlers of the shared definitions)

TOOLS = shared.TOOLS + WORD_TOOLS

CORE_TOOLS = [t.name for t in TOOLS if t.group == CORE]

__all__ = ["CORE_TOOLS", "FORMAT", "GROUPS", "PROMPT", "TOOLS", "WORD_TOOLS"]
