# Changelog

docx-agent has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail of every phase.

## Unreleased

- Date formats (round 6, R9; ooxml-edit 0.38.0, `ooxml_edit.dates`). Excel's date tokens
  (`d dd ddd dddd m mm mmm mmmm mmmmm yy yyyy`, either case), everything else written as
  given; month and day names measured in Excel for Mac, 2026-10-16, for en-US, en-GB,
  nl-NL, nl-BE, de-DE, fr-FR, es-ES and it-IT; another locale refuses a name token and
  takes the numeric ones.
  - A reference in any text tool takes a date pattern: `{{d:events/#1/when:d mmmm yyyy}}`
    is `2 March 2026`, `2 maart 2026` in nl-BE; `:dd/mm/yyyy` is `02/03/2026` in every
    locale (fixed text). A date pattern on a non-date and a number format on a date are
    refused, naming the other. The prompt says so in one clause.
  - `word_insert_markdown` from and `word_edit_table`'s `set_cells` from take `date_format`:
    the dates of the columns placed in that pattern; refused without `from`, or when no
    column placed holds a date.
  - `edit_chart` add from takes `date_format`: the category dates are kept as dates
    (serials) in the pattern's Excel code for the session's locale (`[$-40C]d\ mmmm\ yyyy`),
    in the chart and in its workbook, on a text axis; Word shows the code's `/` as the
    viewer's date separator and the names by the prefix (measured in Excel). `render`
    draws them as Word does (docx2svg on ooxml-common 0.17.0: `5 januari 2026` under
    `[$-413]d\ mmmm\ yyyy`, measured in Word for Mac, 2026-10-16); `edit_chart read` lists
    them as ISO dates (`2026-01-05`) with `category_format`, the code (ooxml-edit 0.39.0), a
    category is named by that date, and the Markdown outline's chart line lists them so.
    ISO rather than the code's text: a model reads it whatever the locale, and it is the
    form `from` dates take. No definition changes.
  - Without a pattern nothing changes: dates are written in the locale's day order.
  - Definitions: +84 estimated tokens to every definition (12,228 -> 12,312), none to the
    core (2,570); the guard rises to 12,315. The strict set is unchanged. Needs ooxml-edit
    0.38.0.

- A deck's slide-number reference in a Word text tool's text (`{{slide:s:262}}`, round 6,
  R8) is refused: "slide references are for decks; a Word document has no slides. Write the
  text itself, or \{{ for the braces" -- `FORMAT.own_facts` names the prefix with that text
  (ooxml-edit 0.37.0), so it holds with no deck format loaded. Before, the braces were
  written as text. No change to any definition. Needs ooxml-edit 0.37.0.

- From ooxml-edit 0.36.0 (round 6, R5, R6): a text fact with a suffix that is no number
  format (`{{d:x/#1/avg:0.20}}`) says "'0.20' is not a number format" and lists the forms,
  not "no column 'avg:0.20'"; find_text's and replace_text's `slides` (decks only) take
  slide ids, which adds 12 estimated tokens to the core (2,558 -> 2,570) and to every
  definition (12,216 -> 12,228; the guard 12,220 -> 12,230).

- Forgiving the model's small mistakes (round 6, R2, R3, R7; ooxml-edit 0.35.0), each read
  with a line in the result's `warnings` and no change to any definition: a ref named
  `vatNote` or `vat-note` is kept as `vat_note` and `$vatNote` read as `$vat_note`, in the
  declared targets and through the tools' own `resolve` alike (refused when the folded name
  already names another range); `functions.word_set_text` as a batch op's `tool`; a key one
  level off says where it belongs (`word_sections` `top`: "top belongs one level down:
  margins.top"). `word_format` reads `top` as `space_before` and `bottom` as
  `space_after`, as pptx-agent's `ppt_format_text` does: it has no other top or bottom.

- Figures in text from a source (round 5, T4; ooxml-edit 0.33.0): `word_set_text`,
  `word_insert_text`, `word_insert_markdown` (`markdown`) and `replace_text` (`replace`)
  take references to one value, `{{d:by_topic/door/mean}}` or `{{t:checks/N-1/rating}}`
  (`:0.0`, `:#,##0`, `:0.0%` after one to format a number), and write the value in the
  session's locale (`2,2` in nl-BE, dates in its day order); the result's `data.facts`
  lists each value with its sources. `\{{` writes one literally and `{{name}}` stays as it
  is; a reference to no single value refuses the call, changing nothing. In Markdown the
  value is Markdown text, as the rest is. Not yet kept in the file, unlike a deck's: the
  provenance item keys a record by a table's or a drawing's durable id, and a paragraph's
  id changes when it is edited. The prompt: one sentence on references (199 characters,
  about 57 estimated tokens). The definitions are unchanged (every definition 12,216, guard
  12,220); the strict set is unchanged.

- A chart added with its look in one call (round 5, T2; ooxml-edit 0.32.0): `edit_chart`'s
  `add` also takes `data_labels: true`, `colors` (each series' colour in order, a pie's
  slices') and `style` with `part` (the text), the chart byte for byte the one `add` and
  the separate `show_data_labels`, `format_series`/`format_point` and `set_text_style`
  calls write (tested for column with a line series, line, pie and scatter); one undo
  step. A document's chart is sized by `width` (its height keeps Word's 432 x 252 ratio);
  `box` stays the deck's. The definitions: 63 estimated tokens in every definition
  (12,153 -> 12,216; the guard 12,155 -> 12,220), none in the core (2,558). The strict set
  is unchanged.

- Fewer rows of a source (round 5, T3; ooxml-edit 0.31.0), none of it in the definitions:
  `from` takes a row range -- `d:checks/#2:` (row 2 to the last), `d:checks/#1:#9`, and a
  workbook's `t:Sales/#1:#9` (xlsx-agent) -- each value with its sources. A table made from
  one keeps the header row repeating on each page (`w:tblHeader`), as every table from a
  source does: Word flows a long table over pages itself, so nothing like the deck's
  `continue` is needed. `set_cells` refused for want of rows now offers the row range that
  fits as `valid_options` ("..., or place fewer with d:checks/#1:#2 (nothing is cut)").
  The tool definitions, their estimated tokens and the strict set are unchanged.
- `describe` lists the session's datasets (round 5, T11; ooxml-edit 0.30.0), none of it in
  the definitions: a document's `describe` result gets `datasets` -- each one's id, columns,
  `types`, `rows` (a count) and `from` (`d1/t:Sales`, `d:by_rep`, or `application` for one
  the application kept), `stale`/`closed` when true -- so a figure placed from `d:<id>` is
  seen to have a source. Never a value; no key when the session has none.
- Dates from a source written one way (round 5, T18b; ooxml-edit 0.29.0), none of it in the
  definitions. `word_insert_markdown` and `word_edit_table` `set_cells` from a source write a `date`, `datetime` or `time` column's ISO
  8601 text -- a workbook's date column through `t:`, a derive's dates, a dataset labelled
  `date` -- in the session locale's day order, as a Python date already was (`1/5/2026`
  en-US, `5/1/2026` nl-BE), so one document no longer mixes `1/3/2026` and `2026-01-20`;
  unlabelled ISO text is written as it is. `edit_chart` categories from such a column are
  written the same way, and `set_values` matches its rows to categories written ISO (an
  older document) or in the locale. The kept provenance names cells, never values:
  unchanged. Needs ooxml-edit 0.29.0.

- Fewer wasted turns on names a model naturally writes (round 5, T6 and T8; ooxml-edit
  0.27.0), none of it in the definitions: the estimated tokens and the strict set are
  unchanged.
  - Run colours: the DrawingML names of the theme colours a deck or chart uses are read as
    Word's -- `tx1`, `tx2`, `bg1`, `bg2`, `dk1`...`lt2`, `hlink`, `folHlink` as `text1`,
    `text2`, `background1`, `background2`, `dark1`...`light2`, `hyperlink`,
    `followedHyperlink` -- and Word's own in any case and with a space before the number
    (`"Text 1"`, `"Accent 3"`; `formatting.theme_color_name`). The refusal still lists
    Word's names, and says the DrawingML ones are read as them. Chart colours take
    PowerPoint's and Word's names too (ooxml-edit's `charts.format`).
  - `word_insert_markdown` reads `columns` as `from_columns`, the name `edit_chart` uses
    for the same thing, and `word_edit_table` does when `columns` is a list of names (its
    own `columns` is split's count); each with a warning, `columns read as from_columns`.
  - Refusals say what would do (the dispatcher's): an enum value or unknown argument close
    to an allowed one says "did you mean"; `key` on a tool that takes none says keys are for
    the tools that create something; a ref name with capitals says the rule in words and a
    name that would do.

- `batch` applies its ops one by one by default (ooxml-edit 0.26.0, round 5 T1): a failing
  op -- an anchor that names nothing, an enum the schema refuses, a `$ref` a failed op would
  have made -- is skipped with its own error and the others still apply as one undo step;
  `data` then lists each op's outcome with `applied` and `failed`, and the summary names the
  ops to resend. `atomic: true` keeps the all-or-none batch. The format's new `restored`
  hook drops the document's indexes after a failed op is rolled back inside the batch.
  The definitions grow by 22 estimated tokens (core 2,535 -> 2,558, about 3,786 counted;
  every definition 12,131 -> 12,153, guard 12,150 -> 12,155); the strict set is unchanged.
  The golden transcripts replay unchanged: a batch whose ops all apply returns what it did.
- Source addresses in the saved file (ooxml-edit 0.25.0): where a table's or a chart's values
  came from -- the record of the result's `from` (`ooxml_edit.provenance`: the source
  without the session's handle, the source document's name, its version, per column the
  rows and their sources) -- is now kept in the document, where X1 kept it in the result
  only. One custom XML item related from the document part, the shape Word for Mac keeps
  through open, edit and save (measured 2026-10-14): `customXml/itemN.xml`
  (`<prov:sources xmlns:prov="urn:uvrt:docx-agent:provenance:v1">`, one
  `<prov:source kind="table|chart" key="...">` per object holding the compact JSON record),
  `itemPropsN.xml` with `ds:schemaRef` that namespace, a `customXml` relationship from
  `word/document.xml`. A table is keyed by its `t:` id (its first row's `w14:paraId`, kept
  by Word because every paragraph and row is stamped), a chart by its `d:` id (`docPr`).
  Other custom XML items are left as they are; the last record dropped takes the item, its
  properties and the relationship away, all in the call's undo step.
  - The docx format's `store_provenance` and `read_provenance` hooks:
    `word_insert_markdown`'s and `set_cells`' `from` call `call.keep_provenance`, and the
    shared `edit_chart` keeps a chart's. `session.provenance(doc)` and the new
    `Document.provenance()` read `{address: record}` for the tables and charts there are
    with tracked changes accepted; its `ignored` counts the entries left out (object gone,
    or not a record). `Document.store_provenance(address, record)` keeps one, `None` drops
    it.
  - Kept true, or dropped: `set_cells` without `from` over the block drops those columns'
    entries; a row inserted or deleted above the block shifts it, below it leaves it, inside
    it drops the record; a column inserted shifts the entries right of it, one deleted drops
    its entry and shifts the rest; merge and split drop it; a record with no column left is
    dropped. Rows are counted with tracked changes accepted. Accepting changes keeps the
    record; rejecting any change in a recorded table drops it (its old values may be
    back). A table renamed by a row at its top takes its record to the new id.
  - Deletes and copies: every outermost edit and batch prunes the entries whose table or
    chart is no longer in the file, so a deletion drops its record in its own undo step and
    undo brings both back; a save prunes nothing, and an entry whose object another program
    deleted is left out on reading and pruned by the next edit. A copied table (or chart)
    gets new ids, so no record -- never a shared one.
  - `word_edit_table`'s `insert_row` before row 0, and `delete_row` of row 0, failed with
    `not_found` (the table's id is its first row's, and the result looked the old one up);
    they now return the table under its new id, with `renamed` saying so.
  - No tool definition changes (about 12,131 estimated tokens, the guard 12,150); every
    `strict_problems` stays empty. ooxml-edit 0.25.0 is the floor.

- Tables from data (X1): `word_insert_markdown` (as a third choice beside `markdown` and
  `blob`) and `word_edit_table`'s `set_cells` take `from` -- a dataset `d:<id>` or another
  open document's table `t:<name>`, as ooxml-edit 0.22.0's `edit_chart` does -- with
  `from_columns` (the source's columns by name, in order; default all) and `header` (the
  header row's texts; `[]` for none; default the column names), the behaviour and result
  of pptx-agent's tables from data. Every source row is placed; numbers are written as
  chart category labels are (General, the session locale's decimal separator), dates as the
  source gives them, line breaks as line breaks in the cell's one paragraph, the text
  literally. A new table is the source's size (no `rows`/`columns`), written as a Markdown
  table is (`insert_blocks`, a no-header option in the Markdown table model: no
  `w:tblHeader`, `tblLook` `firstRow` 0); `set_cells` writes from row `index`. A source over
  63 columns, 4,096 rows or 5,000 cells, or too big for the table written into, is refused,
  never cut; a block over a merged cell is refused. `data.from` gives the source and, per
  placed column, its table column, source rows and their sources; none of it is saved.
  The result gives the table's id and size, not its Markdown (the caller has the data).
  One call is one undo step, tracked while tracking is on.
- `d:` is a drawing (`d:3`) and a dataset (`d:<id>`): the docx format's `source_hint`
  (ooxml-edit 0.23.0) says `d:3 is a drawing in this document, not a dataset` when a source
  argument (`from`, `edit_chart`'s too) is given a drawing's id. No argument that takes a
  drawing is a source.
- Every definition grows by 194 estimated tokens (about 11,933 -> 12,127; the guard 11,950 ->
  12,150), the core not at all. word_insert_markdown (seven optional parameters) stays
  strict; edit_smartart no longer fits Claude's 24 optional parameters and is sent
  non-strict. ooxml-edit 0.23.0 is the floor.
- A new table's paragraph and row ids are drawn once per seed instead of two at a time
  (the same ids): writing a 5,000-cell table took 80 s, now under one.

- The user's locale (ooxml-common 0.15.0's `Locale`, ooxml-edit 0.19.0's
  `Toolbox(locale=...)`): `Document.locale` and `Document.new(locale=...)`, set from the
  session for every document that joins it. Every layout and render passes it to docx2svg
  (`ConvertOptions(locale=...)`) -- `render_svg`, `render_png`, `layout()`, the reflow in
  each edit's checks, field and TOC page numbers, the `render` tool's worker (a subprocess
  worker too) and the layout worker; the locale keys the render and layout caches. A new
  chart (`insert_chart`) gets `add_chart(..., locale=)`, and its text takes the locale's
  language where the document states none, where it took en-US. en-US without one, never
  the server's. No tool definition changes.
- **Changed output:** a new blank document's `w:decimalSymbol` and `w:listSeparator` are
  the locale's, so `Document.new()` without one writes en-US's `.` and `,` where it wrote
  `,` and `;` (the measuring Word's, Dutch, metric) -- accepted. A locale with a decimal
  comma writes those again, part for part. A new document's language, not given, is the
  locale's (en-US without one). A Letter page wrote `.` and `,` already: unchanged.
- `blank.Locale` is split: the page's margins, distances, tab stop and hyphenation zone are
  `blank.PageDefaults` (`Document.new(page_defaults=...)`; `blank.METRIC`, `blank.US`), the
  separators come from the shared `Locale`. `blank.Locale`, and `METRIC` or `US` given as
  `locale=`, are deprecated (`DeprecationWarning`) and write what they wrote. With docx2svg
  #8 a decimal tab stop aligns on the locale's separator: `.` by default where docx2svg
  aligned on `,`.
- ooxml-edit 0.18.0 makes workbooks a shared-tool kind, and the shared definitions every
  Word request carries grow with it (`describe`, `find_text`, `new_document`'s `kind`,
  `save_document`'s `format`): the core 2,471 -> 2,535 estimated (about 3,752 counted by
  the proxy, within 6,000), every definition 11,880; the guard rises to 11,950.
  `save_document` refuses a format that is not a document's (`xlsx`, `xlsm`, `xltx`) with
  `invalid_arguments` naming docx, dotx and markdown, where it would have failed as
  `not_found`.
- Rendering and layout in a daemonic process (a Celery prefork worker's child; ooxml-edit
  0.17.0): `render_pages` and `layout.convert_bytes` are registered worker tasks, so
  ooxml-edit's subprocess workers run them there -- in parallel, with deadlines that kill --
  instead of in-process.
- Chart formatting (production feedback round 3, G1-G4; ooxml-edit 0.15.0): axis scales,
  gridlines, series and point colours and markers, combination charts, chart text and point
  labels through `document.chart(...)` and `edit_chart`'s new actions, in Word's sizes;
  `insert_chart` takes `radar_filled` and `bubble`, `category_order`, `gridlines` and
  `bubble_scale`. Word 16 opens and draws every new construct unprompted and keeps it on a
  save. The every-definition guard rises to 11,850 (11,774 estimated, +646), the core
  unchanged; w3's golden records `read`'s new keys (`scales`, `types`).
- `render` rasterises, and `check` lays out, a snapshot with the document's lock let go:
  the bytes and version are taken under the lock (with `validate` and `fields`, which a
  check reads from the same version), the worker draws or lays them out while other calls
  may edit the document, and the result -- cached under the snapshot's version -- reports
  that version. A layout finished after an edit is mapped to ids on a copy opened from the
  snapshot's bytes. Inside a batch, unchanged. Needs ooxml-edit 0.14.0 (production
  feedback round 3, A6).
- `render`'s and `check`'s `slides` (read as pages) take numbers only: ooxml-edit 0.14 lets
  a deck name slides by id there, and an id given for a document is refused, naming
  `pages`; `pages` itself stays numeric (`"1"` is refused: `pages[0] must be a number`).
- Python 3.14 and 3.15: CI runs the suite on both, on Linux, macOS and Windows, and the
  classifiers declare them. `requires-python` stays `>=3.10`. Python 3.14 made
  `forkserver` Linux's default start method (`fork` before); the worker pool names `spawn`
  itself, so nothing changes, and a test now runs the layout and render tools, with the
  application's font folders, in a worker under every start method the platform has.
- The golden transcripts hold on every CI runner, not only on a Mac with Office: each
  recorded result also carries `data_keys` (its data's schema) and `data_sha_core` (its data
  less what the layout measures), and `tests/test_tools_goldens.py` replays every
  transcript everywhere, comparing all but the layout (defined once in
  `tests/goldens_replay.py`). A result that gains a key -- as `coverage` did, unnoticed by
  CI -- now fails on Linux, macOS and Windows. `tools/refresh_goldens.py` re-records them,
  refusing unless only the named new keys changed and every saved document is unchanged
  (CONTRIBUTING.md, "Golden transcripts").
- The application's own font folders, for the tools and the library: `Toolbox(font_dirs=...)`
  (or a session's own, ooxml-edit 0.13) sets `Document.font_dirs` on every document the
  session opens or makes, and every layout and render uses them -- reflow, `coverage` in
  `render`, `check` and `save_document`, fields' and a TOC's page numbers, PNGs -- in the
  worker process too: the folders are resolved where the toolbox runs (`Document.font_dirs`,
  else `OOXML_FONT_DIRS`) and passed as docx2svg's `font_dirs`, so they also key the layout
  and render caches. No tool definition or prompt changes. Before, a face only in such a
  folder was reported missing and the layout incomplete. Needs ooxml-common 0.8.0 and
  ooxml-edit 0.13.0.
- `check` (with `reflow`), `render` and `save_document` return `coverage`, from docx2svg's
  coverage summary: whether the layout is complete, the blocks laid out, where it stopped
  (by id), header and footer stops, and the faces substituted with open ones or missing.
  "Check passed" can now be told apart from "couldn't check". In the library:
  `DocumentLayout.coverage` and `coverage_facts()`. Needs docx2svg with
  `docx2svg.coverage`.
- `coverage` gains `status`: `complete`, `approximate` or `partial`. When docx2svg laid a
  place out by a rule it has not measured instead of stopping (for example a floating
  drawing in a table cell at an unmeasured position), `approximations` lists the first
  five with page, reason, element path and message, and `approximations_total` gives the
  count when there are more. Needs docx2svg with `Coverage.status` (uvrt/docx2svg#5); an
  older docx2svg still gets `status` derived from `complete`. The golden transcripts'
  data digests are re-recorded for the new key.
- Markdown: a line break that ends its paragraph (a cover page's Shift+Enter) is written
  as `<br>`. It was written as `\`, which CommonMark reads back as a literal backslash at
  the end of a block. When reading Markdown, `<br>`, `<br/>` and `<br />` are line breaks,
  and a paragraph that is only `<br>` is a paragraph of line breaks.
- The invalid `\h` escape in `to_markdown`'s docstring is fixed. CI now compiles the
  sources with syntax warnings as errors.
- README: which fonts the layout needs, and the open substitutes.
- `undo` refuses ooxml-edit 0.12's `scope` (`invalid_arguments`): a document undoes
  document-wide, since every edit shares its body part.
- `edit_chart` `add` makes a radar chart (`chart_type: "radar"`), as Word inserts one:
  lines in the theme's accents, the legend at the top (measured on Office for Mac 16).
  Without `position`, a new chart's legend is where the application puts it. Needs
  ooxml-edit 0.11.0.
- Golden w14-supplier-radar (a supplier comparison with a caption): a radar chart from a CSV.

## 0.0.1 -- 2026-10-08 (initial public release)

Developed 2026-10-03 to 2026-10-08:

- **E0:** a lossless package on [`ooxml-edit`](https://github.com/uvrt/ooxml-edit), live
  views with stable ids, undo, validity checks, rendering and reflow feedback through
  [`docx2svg`](https://github.com/uvrt/docx2svg), and Microsoft Word as the oracle under a
  machine-wide lock.
- **E1-E3:** text ranges and anchors, find and replace across runs, styles and direct
  formatting, lists, hyperlinks, bookmarks, cross-references and pictures; Markdown with
  ids out and in (in the document's own styles); tracked changes and comments.
- **E4-E6:** sections, headers and footers, notes, fields and tables of contents; tables,
  floating drawings, text boxes, shapes and content controls; new documents and templates,
  blocks copied between documents, compatibility-mode upgrade, properties.
- **Charts and SmartArt** wherever Word holds them, with their data and embedded
  workbooks kept in step.
- **Agent tools** (`docx_agent.tools`) on `ooxml_edit.tools`, documented in
  `SUPPORTED.md` and `GUIDANCE.md`, with golden transcripts of the trial's Word tasks.
