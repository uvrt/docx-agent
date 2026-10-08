# A tour of the API

## What is covered

**Status: phases E0 to E6 of [`ROADMAP.md`](../ROADMAP.md), and charts and SmartArt** -- the semantic API on E0's
foundation (a lossless package, live views with stable ids, undo, rendering and reflow
feedback): text ranges and anchors, find and replace across runs, styles and direct
formatting, lists, hyperlinks, bookmarks, cross-references and pictures; reading the
document as Markdown with ids, or as structured JSON, and writing Markdown into it in its
own styles; tracked changes and comments; sections, headers and footers, notes, fields and
tables of contents; tables (made, formatted, merged and split, columns across spans),
floating drawings, text boxes and shapes, and content controls (data-bound ones too);
and authoring -- a new document as Word makes one (blank, or from a template), blocks
copied between documents with their styles, lists, pictures, notes and comments, an older
document converted to compatibility mode 15, the document's properties; and charts and
SmartArt wherever Word holds them, their data and embedded workbooks kept in step -- each
written as Word writes it, tracked where Word tracks.

It builds on [`ooxml-edit`](https://github.com/uvrt/ooxml-edit) (the lxml-based OPC
package, ordered insertion and undo, and its charts subpackage, extracted from pptx-agent)
and renders and lays out through `docx2svg`.

The phases' features in brief. Unlike the recipes above, these snippets are illustrative:
their files and ids are placeholders.

```python
from docx_agent import Document

doc = Document.open("report.docx")
for paragraph in doc.paragraphs():             # body paragraphs, tables and controls too
    print(paragraph.id, paragraph.text)        # p:1A2B3C4D, or p@body/7 until first edited

before = doc.layout()                          # docx2svg's layout, in these ids
result = doc.paragraph("p@body/7").set_text("Revenue grew 14% in the third quarter.")
print(result.id, result.renamed)               # the stamped id, and every rename
doc.insert_paragraph("A new paragraph.", after=result.id)

after = doc.layout()
print(after.where(result.id))                  # [Placement(page=2, top=412.5, bottom=468.1, ...)]
print(after.compare(before).changed)           # the pages the edits changed
svgs = doc.render_svg(pages=after.compare(before).changed)   # data-docx-agent-id on paragraphs

doc.undo(); doc.undo()                         # the original bytes again
doc.save("report.docx")
```

The semantic API (E1):

```python
doc = Document.open("report.docx")

where = doc.anchor("net revenue")                # one place, or AmbiguousAnchor listing them all
where.replace("operating revenue")               # across runs; formatting around it kept
doc.replace(r"Q(\d)", r"quarter \1", regex=True) # every match, one undo step
doc.paragraph("p:3A1F09C2").set_text("Revenue grew 15%.")   # mixed formatting kept

heading = doc.paragraph("p:0C11D2E7")
heading.style = "Heading 2"                      # by name: a Dutch template's Kop2 too
doc.anchor("operating revenue").format(bold=True, color="accent1")   # direct formatting
heading.effective.space_before                   # what Word applies (docx2svg's resolver)

item = doc.paragraph("p:1B2C3D4E")
item.add_to_list("number")                       # Word's own list, a definition of its own
item.restart_numbering(at=1)                     # a new instance over the same definition

doc.anchor("operating revenue").add_bookmark("Revenue")
doc.anchor("see above").add_hyperlink(anchor="Revenue")
doc.anchor("our site").add_hyperlink("https://example.com/")
doc.insert_picture("p:1B2C3D4E@0", "chart.png", width=200, alt_text="Revenue by quarter")
```

Tables, drawings and content controls (E5), each as Word writes it, tracked where Word
tracks (merges and splits too, in the forms Word's own Accept All and Reject All review):

```python
table = doc.insert_table(4, 3, after="p:0C11D2E7", data=[["Item", "Qty", "Price"]], header_rows=1).id
doc.merge_cells(table, (1, 0), (2, 0))           # grows to whole cells; across spans too
doc.split_cell(table, 3, 1, rows=1, columns=2)
doc.insert_column(table, 0)                       # the grid kept consistent, always
doc.set_table(table, width="80%", alignment="center", look={"banded_rows": False})
doc.set_cell(table, 0, 1, shading="D9E2F3", borders={"bottom": "double"}, vertical_alignment="center")
doc.set_row(table, 0, height=24, height_rule="exact")

logo = doc.insert_picture("p:0C11D2E7@0", "logo.png", width=72).id
doc.float_drawing(logo, wrap="square", horizontal_from="margin", x_align="right")
doc.set_drawing(logo, z_order="front", lock_anchor=True)
box = doc.insert_text_box("p:0C11D2E7@0", "Beside the text", x=300).id
doc.drawing(box).paragraphs[0].set_text("Edited like any paragraph")

status = doc.insert_control("p:0C11D2E7@0", "drop-down", items=["Draft", "Final"]).id
doc.fill_control(status, "Final")                 # a bound control's custom XML node too
```

Charts and SmartArt, in the body, headers, footers, notes, text boxes and groups: the
chart's caches and its embedded workbook edited together, one undo step each (the editing
is ooxml-edit's `ooxml_edit.charts`, shared with pptx-agent):

```python
chart = doc.chart("d:3")                          # doc.charts(); a group's: doc.chart("d:5/7")
chart.series[0].set_value(2, 4285)                # the cache and the workbook cell
chart.add_category("Q4", [4400, 530])             # cells move, formulas and the table grow
chart.series[1].name = "Gross margin"
chart.set_title("Margins")
doc.state()                                       # each chart's and diagram's JSON; chart.apply(json)
diagram = doc.diagram("d:4")
diagram.set_text(1, "Much faster edits")          # and the cached drawing, where exact
diagram.node(0).add_child("Smaller files")        # Word lays the diagram out again
```

Authoring (E6): a new document as Word makes one, from a template as File > New makes
one, blocks copied in from another document as Word pastes them, and Convert:

```python
doc = Document.new(title="Quarterly report", author="Claude")      # A4, mode 15, Word's styles and theme
doc = Document.new(page="Letter", orientation="landscape", language="en-GB")
doc = Document.new(template="brand.dotx", keep_content=False)     # the template's styles, sections, headers

source = Document.open("appendix.docx")
result = doc.copy_blocks(source, "p:1A2B3C4D..t:5E6F7A8B", at="end",
                         styles="use_destination")                # or "keep_source", "merge"
result.copied                                     # each source id -> its copy's id
doc.save_as_template("mine.dotx")

old = Document.open("legacy.docx")                # mode 14 stays mode 14 through every edit...
old.upgrade_to_modern().reflow.changed            # ...until asked: the pages that reflowed
doc.set_properties(subject="Q3", keywords="revenue")
```

Writing (E2): Markdown in, as the document's own styles, one undo step (one revision group
when tracking).

```python
result = doc.insert_markdown("""## Next steps

1. Sign the contract.[^1]
2. Book the **kick-off**.

[^1]: Before the end of the month.
""", at="after:p:0C11D2E7")
result.blocks                                     # ['p:...', 'p:...', 'p:...']: the new blocks
doc.to_markdown(f"{result.blocks[0]}..{result.blocks[-1]}", ids=False)   # reads back the same
```

Reading (E2): Markdown with ids, in three views of the revisions, and the JSON state.

```python
print(doc.to_markdown())                         # every block with its id, nothing changed
# <!-- p:3A1F09C2 -->
# ## Results
#
# <!-- p:0C11D2E7 -->
# Revenue grew **14%** in the third quarter.[^fn:2]
#
# - First point <!-- p:1B2C3D4E -->

doc.to_markdown("s:body", view="markup")         # a section, revisions as {++...++}<!-- rev:12 Alice -->
doc.to_markdown("p:0C11D2E7..p:1B2C3D4E", ids=False)   # a range, as plain Markdown
state = doc.state("p:0C11D2E7", layout=True)     # styles, runs, lists, revisions, placements
state["blocks"][0]["effective"]["sz"]            # 11.0: what Word applies (docx2svg resolver)
doc.get("rev:12"), doc.get("fn:2"), doc.get("d:7")   # every id either names resolves
```

What docx-agent guarantees, and tests on every fixture:

- **Lossless.** Opening and saving writes every part back byte for byte; reading anything
  changes nothing; only edited parts are serialised again.
- **Stable ids.** Every paragraph has an id: Word's own `w14:paraId` (`p:<paraId>`), or a
  positional `p@<story>/<n>` until the first edit stamps the document. Ids resolve across
  inserts, deletes, undo and redo, and renames are kept as aliases. Stamped ids survive
  Word: it keeps them on saving and on edits, in every compatibility mode (measured; see
  ROADMAP.md, "Durable across a Word save -- measured").
- **Undo is exact.** Every edit is one step; undoing gives back the original bytes.
- **Formatting survives editing.** Replacing text keeps the formatting of the characters
  around it, and of every character that survives; formatting a range leaves no empty or
  redundant run. Held to an independent model over a randomised sweep of matches across
  run, hyperlink, revision and field boundaries.
- **Valid.** Schema order, relationships, content types, unique ids, paired ranges,
  numbering integrity, style references: no edit adds a problem (`doc.validate()`,
  `docx_agent.validate`).
- **Word opens it.** Every fixture with E1's edits exports from Word without a repair
  prompt, showing the text, styles and list numbers the edits wrote.
- **Reading is read-only.** `to_markdown` and `state` change no byte and stamp no id; the
  Markdown reads back through markdown-it-py as the model it was written from, and its
  final view's text is what docx2svg draws.
- **Markdown round trip.** `insert_markdown` then `to_markdown` gives the same CommonMark
  AST, in a blank document and every fixture, over the CommonMark spec's examples of the
  supported constructs (those outside the guarantee listed with their reasons) and
  hand-written documents; it names only the document's styles (by name: a localised
  template works), adding a built-in one it lacks as Word writes it.
- **Rendered honestly.** Where docx2svg stops laying out, a block past the stop is
  `Unknown`, never guessed.
