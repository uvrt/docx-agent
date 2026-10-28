"""Word tools for an agent tool layer (``ooxml_edit.tools``): definitions and handlers.

An application puts these in a :class:`ooxml_edit.tools.Toolbox` with the docx
:data:`FORMAT` (and, beside them, another library's tools for its documents)::

    from ooxml_edit.tools import Toolbox
    from docx_agent.tools import TOOLS, FORMAT, GROUPS

    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    session = toolbox.session(clock=my_clock)
    d1 = session.open(docx_bytes, name="agreement.docx")
    tools = toolbox.definitions("anthropic", groups="core")
    openai_tools = toolbox.definitions("openai-responses")   # or "openai-chat"
    result = toolbox.dispatch(session, "describe", {"doc": d1})

Everything is in memory: documents and inputs are bytes under handles, saved files go to
the application.  Lengths are points.  Every mutating call is one undo step, written as
tracked changes when ``word_set_tracking`` turned tracking on, and returns ``checks``: the
``validate()`` delta against the document as opened, and the pages it reflowed.  Layout
runs in the toolbox's worker pool under ``Limits.layout_timeout`` (30 s).

The shared tools (``open_document``, ``save_document``, ``find_text``, ``render``...) are
one definition for every format (``ooxml_edit.tools.shared``); this package gives their
Word handlers.

How to run them with each provider (Claude, OpenAI Responses, Chat Completions):
``README.md`` in this package.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Mapping
from typing import Any

from ooxml_edit.tools import CORE, DocumentFormat, ToolGroup, shared as _shared

from ._base import block_of, blocks_between
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


def docx_format(*, min_text_size: "float | dict[str, float] | None" = None, assets: Any = None) -> DocumentFormat:
    """:data:`FORMAT` with the application's own settings (round 17, P6) -- never tool
    arguments, so the definitions the model is sent stay the same unless a setting adds a
    tool; ``deck_format``'s counterpart:

    ``min_text_size`` (points): ``check`` reports ``small_text``, the text drawn smaller
    (its layout facts, as ``min_size`` does), and the builders -- ``word_insert_markdown``
    and ``word_edit_table`` -- the text they wrote that its styles make smaller.  One number
    is the minimum for all text; a mapping sets it per kind -- ``body``, ``table`` (text in a
    table cell), ``footnote``, ``endnote``, ``header``, ``footer``, ``caption`` (a
    paragraph in the Caption style) -- and a kind it leaves out is not checked
    (:mod:`docx_agent.edit.minimums`)::

        toolbox = Toolbox(TOOLS, formats=[docx_format(min_text_size=10)], groups=GROUPS)
        docx_format(min_text_size={"body": 10, "table": 9, "footnote": 8, "caption": 9})

    ``assets``: the application's asset catalogue -- brand icons, logos -- any object with
    ``search(query, limit)`` and ``get(id) -> (bytes, MIME type)``
    (:class:`ooxml_edit.tools.AssetCatalogue`), such as :class:`ooxml_edit.tools.FolderCatalogue`,
    the protocol pptx-agent's ``deck_format(assets=...)`` takes.  With one the toolbox offers
    ``word_find_asset`` (deferred, in ``word_objects``) and ``word_drawings``
    ``insert_picture`` takes ``asset``, an id, in place of ``image``; without one neither is
    sent::

        from ooxml_edit.tools import FolderCatalogue

        icons = FolderCatalogue("brand-icons")               # brand-icons/index.json
        toolbox = Toolbox(TOOLS, formats=[docx_format(assets=icons)], groups=GROUPS)
    """
    from ..edit.minimums import text_minimums

    text_minimums(min_text_size)                 # refused here, not at the first check
    if isinstance(min_text_size, Mapping):
        min_text_size = {kind: float(points) for kind, points in min_text_size.items() if points is not None}
    elif min_text_size is not None:
        min_text_size = float(min_text_size)
    if assets is not None and not (callable(getattr(assets, "search", None))
                                   and callable(getattr(assets, "get", None))):
        raise ValueError("assets is a catalogue: an object with search(query, limit) and get(id), such as "
                         "ooxml_edit.tools.FolderCatalogue")
    tools = () if assets is None else (objects.drawings_tool(assets), objects.find_asset_tool(assets))
    configure = FORMAT.configure

    def configured(document: Any, session: Any) -> None:
        if configure is not None:
            configure(document, session)
        document.min_text_size = dict(min_text_size) if isinstance(min_text_size, dict) else min_text_size

    return dataclasses.replace(FORMAT, configure=configured, tools=tools)


__all__ = ["CORE_TOOLS", "FORMAT", "GROUPS", "PROMPT", "TOOLS", "WORD_TOOLS", "block_of", "blocks_between",
           "docx_format"]
