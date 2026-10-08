# Common tasks

Each `python` block below runs as written in `tests/test_readme.py`, in a folder holding
the files it names, so these recipes cannot rot.

Short recipes for what an agent does most. Every block runs as written:
`tests/test_readme.py` runs each one on its own, in a folder holding `agreement.docx` (the
supplier agreement summary in `tests/fixtures/generated/pilot`, with five reviewers'
comments), `charts.docx`, `brand.dotx` and `logo.png`. The ids (`p:3B212964`, `c:76FCC91F`)
are that document's; in yours, read them from `to_markdown()` first.

### Open, read, and edit by id

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
print(doc.to_markdown())                       # every block, its id in a comment: <!-- p:3B212964 -->
print(doc.to_markdown(view="markup"))          # tracked changes {++ins++}{--del--} and comments {>>...<<} in place
print(doc.to_markdown(view="original"))        # as if every tracked change were rejected ("final": accepted)
print(doc.to_markdown("p:3B212964", ids=False))   # one block, or a span: "p:16425160..p:3B212964"
print(doc.to_markdown(view="markup", stories="all"))   # headers, footers and notes too, their comments

for paragraph in doc.paragraphs():             # the body's paragraphs (tables' too), in order
    print(paragraph.id, paragraph.style_name, paragraph.text)

scope = doc.paragraph("p:3B212964")
scope.set_text("The supplier will host the customer portal for 36 months.")   # formatting kept
added = scope.insert_after("A new paragraph after it.", style="Normal").object   # result.object: the new one
added.format(italic=True)                      # or doc.get(result.id)
doc.undo()                                     # every edit is one undo step
doc.save("agreement-edited.docx")
```

The body alone is read unless `stories=` says otherwise (`doc.comments()` lists every
comment, a header's too). In the markup view a field's result is marked
(`<!-- field: REF RefIncident \h -->section 4<!-- /field -->`), and a heading's id comment
says whether a list numbers it (`numbered: list`) or its number is typed text
(`numbered: text`, renumbered by editing it).

`doc.state()` is the same document as JSON (styles, runs, lists, comments, revisions);
`doc.get(id)` resolves any id (`p:`, `t:`, `c:`, `rev:`, `fn:`, `d:`...).

### Find text and replace it

```python
from docx_agent import AmbiguousAnchor, Document

doc = Document.open("agreement.docx")
fee = doc.anchor("EUR 12,500")                 # the one place; AnchorNotFound or AmbiguousAnchor otherwise
print(fee.id, fee.text)                        # p:4B50F4CD@19:29 EUR 12,500
fee.replace("EUR 13,000")                      # across runs; the formatting around it kept
doc.anchor("3%").insert_after(" (CPI)")        # insert_before, delete, format(bold=True) too

for match in doc.find("months"):               # every place, as TextRanges
    print(match.id, match.context())
try:
    doc.anchor("the")
except AmbiguousAnchor as error:               # lists every candidate with its id
    print(len(error.candidates), "matches; pick one with occurrence= or within=")
first = doc.anchor("the", within="p:3B212964", occurrence=0)

result = doc.replace("customer", "client")     # every match, one undo step (regex=True for patterns)
print(result.count)
```

### Move a section

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
termination = doc.section_blocks("p:058B17C9")  # the heading and everything up to the next ## or #
print(termination)                             # ['p:058B17C9', 'p:6A70E1A8']
with doc.tracking(author="Claude"):            # tracked: Word's moves, accepted or rejected whole
    doc.move_blocks(termination, before="p:12972045")   # before "Pricing"; or after=, or "p:A..t:B"
print([change.kind for change in doc.changes()])   # ['move']
doc.move_blocks("p:73D22116..p:7C9523C2", after="p:3B212964")   # untracked: "Contacts" after "Scope"
```

### Tracked changes

`with doc.tracking(author=...)` writes every edit inside as a tracked change (a revision)
that a person can accept or reject in Word:

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
with doc.tracking(author="Claude"):
    doc.anchor("24 months").replace("36 months")
    doc.anchor("EUR 12,500").insert_after(", excluding VAT")
    service = doc.paragraph("p:69E36C9F")
    service.set_text(service.text.replace("48 hours", "five working days"))
print(doc.to_markdown("p:3B212964", view="markup"))
# ...for a period of {--24 months--}<!-- rev:5 Claude -->...{++36 months++}<!-- rev:6 Claude -->...
doc.word_tracks_changes = True                 # optional: Word's own switch, for a person's later edits
```

A replacement records the whole anchor as deleted and inserted, so anchor tightly
(`"24 months"`, not `"a period of 24 months"`); `Paragraph.set_text` records only the
words that changed. `doc.word_tracks_changes` (formerly `track_revisions`) does **not**
make your edits tracked; it is the Track Changes button a person sees in Word.

### Comments: list, add, reply, resolve

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
for comment in doc.comments(replies=False):    # each thread's first comment; replies are comments too
    print(comment.id, comment.author, repr(comment.anchor.text), "->", comment.text)
# c:76FCC91F Legal '24 months' -> The contract says 36 months, not 24. Please correct.

legal = doc.comment("c:76FCC91F")
print(legal.anchor.id, legal.anchor.paragraph_ids())   # where it is attached: p:3B212964@89:98
doc.reply_to_comment(legal.id, "Changed to 36 months.", author="Claude")
doc.resolve_comment(legal.id)                  # reopen_comment, edit_comment, delete_comment too
print(legal.done, [reply.text for reply in legal.replies])

added = doc.add_comment(doc.anchor("99.5%"), "Per month or per year?", author="Claude")
print(doc.comment(added.id).anchor.text)       # 99.5%
```

`to_markdown(view="markup")` shows each comment where it is attached, with its replies and
whether it is resolved. Comments are not tracked changes: adding, replying and resolving
work the same inside or outside `doc.tracking(...)`, and each is one undo step.

### Revisions: list, accept, reject

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
with doc.tracking(author="Claude"):
    doc.anchor("24 months").replace("36 months")
    doc.anchor("3%").replace("2.5%")
    doc.anchor("The customer may terminate immediately if availability falls below 98% for two "
               "consecutive months.").delete(collapse_space=True)   # and the space before it
for revision in doc.revisions():               # filters: author=, kind=, within=, since=, until=
    print(revision.id, revision.kind, revision.author, repr(revision.text))
for change in doc.changes():                   # the records grouped as a reviewer reads them
    print(change.kind, change.author, repr(change.old_text), "->", repr(change.text))
# replacement Claude '24 months' -> '36 months'  (a deletion and an insertion: one change)
doc.changes(kind="deletion")[0].reject()       # a change is accepted or rejected whole
doc.revisions(author="Claude")[0].accept()     # or doc.accept("rev:5"), doc.reject(["rev:6", ...])
doc.reject(within="p:4B50F4CD")                # every revision in a block or range
doc.accept_all()                               # reject_all() too; each one undo step
```

`revisions()` lists Word's records one by one, so a replacement is two (a `deletion` and an
`insertion`), a new paragraph is its text and its mark (`paragraph-mark-insertion`), and
paragraphs added at a story's end carry Word's `paragraph-properties` record on the last
one; `Revision.kind`'s docstring lists every kind. `changes()` groups them into
`replacement`, `insertion`, `deletion`, `move`, `formatting`, `table` and `section`
changes.

### Write Markdown into the document

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
result = doc.insert_markdown("""## Payment terms

Invoices are due within **30 days**.

- By bank transfer.
- In euros.
""", at="after:p:4B50F4CD")                     # "end", "before:<id>", "replace:<id>[..<id>]"
print(result.blocks)                           # the new blocks' ids, in the document's own styles
print(doc.to_markdown(f"{result.blocks[0]}..{result.blocks[-1]}", ids=False))
doc.insert_markdown("Either party may end the agreement with **three months'** notice.",
                    at="replace:p:6A70E1A8")
```

Inside `doc.tracking(...)` the whole insertion is tracked.

### Markdown in the template's own styles

`style_map` names the styles that differ from the default (`#` is Heading 1, a paragraph
Normal...) as a plain dict:

```python
from docx_agent import Document, StyleMap

doc = Document.new(template="brand.dotx", keep_content=False)
doc.insert_markdown("""# Annual report

Body text in the template's own body style.

| Region | Total |
| --- | ---: |
| North | 4.1 |
""", style_map={"h1": "Heading 1", "paragraph": "Brand Body"})
print([(p.text, p.style_name) for p in doc.paragraphs()])   # cells stay Normal: "table_cell" names theirs
print(StyleMap.DEFAULT.to_dict())              # every key and its default style
```

The keys: `paragraph` (body paragraphs only; a table cell's are `table_cell` and a
footnote's `footnote`, which keep the document's defaults unless named), `h1` to `h6`,
`quote`, `code` (a code block), `inline_code`, `bullet`, `bullet2` to `bullet5`, `number`,
`number2` to `number5`, `table`, `footnote_reference`, `emphasis`, `strong` and `link`.
A `StyleMap` (with `Rule`s) does the same: `StyleMap.from_dict({...})`,
`StyleMap.DEFAULT.with_rules(Rule("heading", "Title", 1))`.

### Tables, pictures and charts

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
table = doc.insert_table(3, 2, after="p:4B50F4CD", header_rows=1,
                         data=[["Item", "Monthly fee"], ["Hosting", "EUR 12,500"]]).id
doc.table(table).cell(2, 0).paragraphs[0].set_text("Support")
doc.format_cell(table, 0, 1, shading="D9E2F3")   # formatting only (set_cell is the same)
doc.insert_row(table)                          # below the last; delete_row, insert_column, merge_cells...
print(doc.to_markdown(table, ids=False))       # a GFM table

picture = doc.insert_picture("p:4C32E1FB@0", "logo.png", width=72, alt_text="Company logo").object
picture.alt_text, len(picture.image)           # 'Company logo'; image is a property: its bytes

charts = Document.open("charts.docx")
chart = charts.chart("d:1")                    # charts.charts() lists them
print(chart.chart_type, chart.categories, [series.name for series in chart.series])
chart.series[0].set_value(2, 4285)             # the chart's cache and its embedded workbook
chart.set_title("Sales by region")
book = chart.workbook_values()                 # what Word's Edit Data shows, read from the workbook
assert book["series"][0]["values"]["values"] == chart.series[0].values
```

### Sections, headers and footers, table of contents

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
doc.add_header("s:body", "default", "Supplier agreement summary")   # s:body: the last section
footer = doc.add_footer("s:body", "default", "Page ").id                # its story: "footer1"
doc.insert_page_number(doc.paragraphs(footer)[0].range(5))           # a PAGE field after "Page "
landscape = doc.insert_section_break(after="p:6A70E1A8", kind="nextPage").id
doc.set_section("s:body", orientation="landscape")                    # the section after the break
toc = doc.insert_toc(before="p:16425160")      # page numbers from docx2svg's layout
doc.update_fields()                            # recompute every field (TOC, PAGE, REF...)
print([section.id for section in doc.sections()], toc.id)
```

### Render and find where text landed

```python
from pathlib import Path

from docx_agent import Document

doc = Document.open("agreement.docx")
before = doc.layout()
doc.paragraph("p:3B212964").set_text("A much longer paragraph. " * 40)
after = doc.layout()
print(after.where("p:3B212964"))               # [Placement(page=1, top=..., bottom=..., lines=...)]
print(after.compare(before).changed)           # the pages the edit changed
svg = doc.render_svg(pages=[1])[0]             # each paragraph's group carries data-docx-agent-id
Path("page1.png").write_bytes(doc.render_png(pages=[1])[0])   # pip install 'docx-agent[png]'
```

Renders always show the final view (tracked changes accepted, comments hidden); review
changes with `to_markdown(view="markup")`.

### New documents and templates

```python
from docx_agent import Document

doc = Document.new(title="Meeting notes", author="Claude")   # A4, Word's own styles; page="Letter"...
doc.insert_markdown("# Meeting notes\n\nPresent: Legal, Finance, Procurement.")
doc.save("notes.docx")

branded = Document.new(template="brand.dotx", keep_content=False)   # the template's styles and headers
branded.insert_markdown("# Board summary\n\nOne page.")   # replaces the one empty paragraph
branded.copy_blocks(Document.open("agreement.docx"), "p:16425160..p:3B212964", at="end")
branded.save("summary.docx")                   # a document
branded.save_as_template("mine.dotx")          # or a template: the extension decides
```

Copying from a document in another house style, `style_map` names the destination style
for a source style, and `unmapped="body"` gives every other source style the destination
lacks the destination's body style instead of importing it:

```python
from docx_agent import Document

report = Document.new(template="brand.dotx")
source = Document.open("agreement.docx")
result = report.copy_blocks(source, source.section_blocks("p:12972045"), at="end",   # "Pricing"
                            style_map={"heading 2": "Heading 1"}, unmapped="body")
print([(p.style_name, p.text[:20]) for p in report.paragraphs()][-2:])
report.styles.purge_unused()                   # custom styles nothing uses; remove(name, replacement=) one
```

`styles="merge"` keeps a heading's level and a list item's list; `doc.styles` is a
property (`for style in doc.styles`).

**Opening a template is not making a document from it.** `Document.new(template="brand.dotx")`
is Word's File > New: a new *document* with the template's styles, headers and content.
`Document.open("brand.dotx")` opens the *template itself*, to edit it, and warns
(`TemplateOpened`). Either way `save()` writes the kind of file its extension names -- a
`.docx` or `.docm` is a document, a `.dotx` or `.dotm` a template -- since Word refuses a file
whose kind and extension disagree; `validate(target="x.docx")` reports such a mismatch in
bytes written another way.

### Save and validate

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
with doc.tracking(author="Claude"):
    doc.anchor("24 months").replace("36 months")
problems = doc.validate()                      # what Word would repair or refuse; [] is clean
assert not problems, "\n".join(map(str, problems))
doc.save("agreement-reviewed.docx", validate=True)   # refuses to write when validate() finds any
```
