"""docx-agent: an AI-editable Word layer over lxml, ooxml-edit and docx2svg.

Typical use::

    from docx_agent import Document

    doc = Document.open("report.docx")
    first = doc.paragraphs()[0]
    result = first.set_text("Revenue grew 14% in the third quarter.")
    before = doc.layout()
    doc.insert_paragraph("A new paragraph.", after=result.id)
    print(doc.layout().compare(before).changed)    # the pages the edit changed
    doc.undo()

    with doc.tracking(author="Claude"):              # every edit a revision
        doc.anchor("draft").replace("final")
    doc.add_comment(doc.anchor("final"), "Changed at the client's request.")
    doc.revisions(author="Claude")[0].accept()

    section = doc.insert_section_break(after=first.id, kind="nextPage").id   # E4: structure
    doc.set_section(section, orientation="landscape", columns=2)
    doc.add_footer(section, "default", "Page ")
    doc.insert_toc(before=doc.paragraphs()[0].id)          # page numbers from docx2svg
    doc.update_fields()

    table = doc.insert_table(3, 3, after=first.id, data=[["Item", "Qty", "Price"]], header_rows=1).id
    doc.merge_cells(table, (1, 0), (2, 0))                       # E5: tables, drawings, controls
    doc.set_table(table, width="80%", alignment="center")
    doc.set_cell(table, 0, 1, shading="D9E2F3", vertical_alignment="center")
    picture = doc.insert_picture(f"{first.id}@0", "logo.png", width=72).id
    doc.float_drawing(picture, wrap="square", horizontal_from="margin", x_align="right")
    box = doc.insert_text_box(f"{first.id}@0", "A note beside the text", x=300).id
    doc.drawing(box).paragraphs[0].set_text("A note, edited through the paragraph API")
    control = doc.insert_control(f"{first.id}@0", "drop-down", items=["Draft", "Final"]).id
    doc.fill_control(control, "Final")
    doc.save("report-edited.docx")

    chart = doc.chart("d:3")                                     # charts and SmartArt, anywhere
    chart.series[0].set_value(2, 4285)                           # the cache and the workbook cell
    chart.add_category("Q4", [4400, 530])                        # cells, formulas, table follow
    doc.diagram("d:4").node(0).add_child("Smaller files")

    new = Document.new(title="Summary", author="Claude")         # E6: Word's own new document
    new.copy_blocks(doc, f"{first.id}..{first.id}", at="end")    # with its styles, lists, media
    new.save("summary.docx")

See ROADMAP.md for the design and what each phase adds.
"""

from .edit import (AmbiguousAnchor, AnchorNotFound, Bookmark, Cell, Document, EditError, EditResult,
                   Hyperlink, ListMembership, Paragraph, Picture, Row, Run, Section, Story, Table, TextRange)
from .edit.annotations import Comment, ContentControl, Drawing, Note
from .edit.charts import (Chart, ChartDataWarning, ChartEditError, Diagram, DiagramDrawingDropped, Member,
                          UntrackedChartEdit)
from .edit.fields import Field
from .edit.importing import copy_blocks
from .markdown.stylemap import DEFAULT as DEFAULT_STYLE_MAP, Rule, StyleMap
from .revisions.changes import Change
from .revisions.review import Revision
from .revisions.stamp import Tracking

__version__ = "0.0.1"

__all__ = ["AmbiguousAnchor", "AnchorNotFound", "Bookmark", "Cell", "Change", "Chart", "ChartDataWarning", "ChartEditError",
           "Comment", "ContentControl", "DEFAULT_STYLE_MAP", "Diagram", "DiagramDrawingDropped", "Document", "Drawing", "EditError",
           "EditResult", "Field", "Hyperlink", "ListMembership", "Member", "Note", "Paragraph", "Picture", "Revision", "Row",
           "Rule", "Run", "Section", "Story", "StyleMap", "Table", "TextRange", "Tracking", "UntrackedChartEdit",
           "__version__", "copy_blocks"]
