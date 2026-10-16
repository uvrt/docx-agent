"""The docx :class:`~ooxml_edit.tools.DocumentFormat`: how the tool layer opens, recognises
and validates a Word document, which warnings it collects, how library errors map to error
codes, and the Word system-prompt fragment (mechanics only: no house style)."""

from __future__ import annotations

import io
import zipfile
from typing import Any

from ooxml_edit.charts import ChartDataWarning
from ooxml_edit.tools import DocumentFormat

from ..edit.authoring import AuthoringWarning
from ..edit.charts import DiagramDrawingDropped, UntrackedChartEdit
from ..edit.document import Document
from ..edit.errors import EditError
from ..edit.formatting import FormattingError
from ..edit.inline import SpanError
from ..edit.pictures import PictureError
from ..edit.ranges import AmbiguousAnchor, AnchorNotFound
from ._base import checks, problem_key
from .table_data import source_hint

KIND = "docx"
EXTENSIONS = (".docx", ".docm", ".dotx", ".dotm")

PROMPT = """\
Word documents
- describe first: headings with ids, sections, styles, comments, revisions, fields, \
tables, pages. word_read gives Markdown with an id comment before each block; pass a range \
to read part of it.
- Addresses: p:3B212964 a paragraph; p:3B212964@4:11 characters 4-11 of it; p:A@5 a \
position; p:A..t:B the blocks from A through B; t: table, c: comment, rev: revision, \
fn:/en: note, d: drawing, cc: content control, fld: field, s:body the last section, \
header1/footer1 stories. A paragraph's id can change when it is edited: the result's \
renamed maps old to new; the old id keeps working.
- Never paste word_read's Markdown into a text tool: it is escaped. Pass plain text.
- A figure from a dataset or table goes into text as {{d:<id>/<key>/<column>}} or \
{{t:<table>/<key>/<column>}}, :0.0 after it to round: the tool writes the value in the \
locale and reports its source.
- To change words inside a paragraph, prefer replace_text with expect one, or \
word_insert_text and word_delete with find; word_set_text rewrites a whole paragraph. \
word_format sets styles and run or paragraph formatting. A new table is a Markdown \
table through word_insert_markdown.
- Tracked changes: word_set_tracking on makes every later edit a revision by its author; \
word_changes lists, accepts and rejects them. \
Comments, replies and resolving are never revisions.
- Fields (TOC, page numbers, cross-references) are caches: after changing headings or \
moving text, word_fields update.
- Renders show the final view (revisions accepted, comments hidden); read markup with \
word_read view markup.
- Before saving: check with include validate, reflow and fields; update stale fields.
"""

#: Claude's strict mode fits only some tools (20 tools, 24 optional parameters and, measured,
#: about 32 free-text strings per request): the writing tools the reference transcripts
#: (w1-w6, batch-30) call most -- 23 of their calls, 24 optional parameters, 23 strings.
#: word_changes counts its list calls too.  edit_chart left the list in T4: with its add
#: action it has 14 optional parameters, more than half the request's.  X1's from,
#: from_columns and header took word_insert_markdown to seven, and edit_smartart, which
#: filled the last two places, out of the strict set (23 optional parameters now).
STRICT_FIRST = ("replace_text", "word_changes", "word_set_tracking", "word_insert_markdown",
                "word_template")


def open_docx(data: bytes) -> Document:
    return Document.open(data)


def configure(document: Document, session: Any) -> None:
    """A document joining a session lays out and renders with the session's font folders
    (``Toolbox(font_dirs=...)``; ``None``: ``OOXML_FONT_DIRS``) and is drawn and edited in
    its locale (``Toolbox(locale=...)``; en-US unless given)."""
    document.font_dirs = session.font_dirs
    document.locale = session.locale


def store_provenance(document: Document, address: str, record: Any) -> None:
    """Keep a placement's record in the document's own custom XML item
    (:mod:`docx_agent.edit.provenance`), keyed by the table's ``t:`` or the chart's ``d:``
    id; ``None`` drops it.  Inside the call's undo step."""
    document.store_provenance(address, record)


def read_provenance(document: Document) -> dict[str, Any]:
    """``{address: record}`` for the tables and charts the document has (tracked changes
    accepted); an entry whose object is gone is left out."""
    return dict(document.provenance())


def detect(data: bytes, name: str) -> bool:
    if not data.startswith(b"PK"):
        return False
    if name.lower().endswith(EXTENSIONS):
        return True
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            return "word/document.xml" in archive.namelist()
    except zipfile.BadZipFile:
        return False


def problems(document: Document) -> list[Any]:
    return list(document.validate())


def summary(document: Document) -> dict[str, Any]:
    """What ``open_document`` says of a Word document: cheap facts, no layout."""
    from .read import heading_tree

    title = document.properties.get("title")
    return {"title": title, "paragraphs": len(document.paragraphs()), "tables": len(document.tables()),
            "headings": len(heading_tree(document)), "comments": len(document.comments()),
            "revisions": len(document.revisions()), "template": document.package.kind in ("dotx", "dotm")}


def restored(document: Document) -> None:
    """Parts were put back behind the document's back -- a ``batch`` op that failed, rolled
    back inside the batch's open undo step: the trees were reparsed, so every index is
    stale (what the document's own ``undo`` does after the history's)."""
    document._invalidate()


def _candidates(error: AmbiguousAnchor) -> list[str]:
    out = []
    for candidate in error.candidates[:50]:
        try:
            out.append(f"{candidate.id} {candidate.context()!r}")
        except Exception:  # noqa: BLE001 -- a candidate that cannot say its context
            out.append(str(getattr(candidate, "id", candidate)))
    return out


ERRORS = {
    AmbiguousAnchor: ("ambiguous", _candidates),
    AnchorNotFound: "not_found",
    FormattingError: "unit",
    SpanError: "refused",
    PictureError: "refused",
    EditError: "refused",
}

#: The references a deck resolves in its own text that a Word document refuses in words of
#: its own (``DocumentFormat.own_facts``), with no deck format loaded too: pptx-agent's
#: ``{{slide:s:262}}``, a slide's number.
OWN_REFUSED = {"slide:": "slide references are for decks; a Word document has no slides"}

FORMAT = DocumentFormat(
    kind=KIND, open=open_docx, detect=detect, problems=problems, problem_key=problem_key,
    warnings=(AuthoringWarning, UntrackedChartEdit, DiagramDrawingDropped, ChartDataWarning),
    prompt=PROMPT, errors=ERRORS, summary=summary, checks=checks, strict_first=STRICT_FIRST,
    configure=configure, source_hint=source_hint, store_provenance=store_provenance,
    read_provenance=read_provenance, restored=restored, own_facts=dict(OWN_REFUSED))
