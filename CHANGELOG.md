# Changelog

docx-agent has not been released to PyPI. Each version that reaches `main` is tagged
`v<version>`, from 0.1.0, the first tagged. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail of every phase.

## 0.1.17 -- unreleased

- **A table of contents lists what Word lists for `\o`, `\t` and `\u`** (production round
  21, T1): `update_fields` (and `word_fields`' update) rebuilt every table of contents
  from every outline-level paragraph, `\t` ignored -- a `TOC \h \z \t "Section title,1"`
  listed the body's Heading 2s as well. The entries are now chosen by Word's rules,
  measured on 51 made-up documents updated by Word (Dutch and English;
  `tools/toc_probe.py`, `tests/observations/toc-word.json`, the `fields` module
  docstring): `\o "a-b"` the paragraphs whose *style's* outline level is in the range
  (built-in `heading N` styles are level N whatever their definition says; a paragraph's
  own `w:outlineLvl` is not read); `\t` the paragraphs in each named style at its level,
  before `\o`'s; `\u` the paragraph's own level before its style's, and with `\u` a
  paragraph with none is never listed; neither `\o` nor `\t` is `\o "1-9"`; empty
  paragraphs never. `\t`'s list is split on the document's locale's list separator
  (`Document.locale`, en-US's `,` by default): Dutch Word takes `"Section title,1"` as one
  style name. `docx_agent.edit.fields.toc_options` reads an instruction as Word does.
  Not modelled: Word compares `\t`'s names with its own UI language's display names (Dutch
  Word: "Kop 2"); the styles part's names are compared, which are English Word's.
- **`Document.blocks()`** (T6): the body's top-level blocks in order, one dict each --
  `id`, `kind` (`paragraph`, `table`, `control`, `custom_xml`), `style` (display name),
  `level` (the outline level 1-9 a table of contents reads: `\u`'s, which for a paragraph
  stating none is `\o`'s), `section_break`, `movable` (`move_blocks` takes it). And
  `Document.outline_level(id)`. `word_describe`'s heading tree and `word_design_facts`'
  patterns read it, so their headings are now Word's outline levels: a built-in heading
  style is its level even where a paragraph states another; a body-text paragraph
  (`w:outlineLvl` 9) in a style with a level is not a heading.
- **`word_design_facts` reports every declared style** (T7): `include: ["declared_styles"]`
  (not in the default) lists every paragraph, character and table style, used or not
  (`used`), with its `basedOn` chain and its look after inheritance (what differs from the
  default paragraph style's, that style given whole), a table style its conditional
  formats (`parts`) and each place's fill, borders and text -- the resolution the styles in
  use go through. Paged by size (`cursor`, `next_cursor`): a template of 164 styles (100 of
  them table styles) is 152,000 characters of facts, about 7 pages at the default limit.
  **Colours are what Word draws**: a theme colour with its tint or shade
  (`w:themeColor`/`w:themeTint`/`w:themeShade`, `w:themeFill`...) is resolved to RGB --
  measured: Word draws the theme colour, not the `w:val` beside it, and the tint or shade
  is the HLS luminance's (ooxml-common's `lumMod`/`lumOff`, channels rounded down; 24 of
  27 swatches exact, 3 one level off; `tools/theme_tint_probe.py`,
  `tests/observations/theme-tints-word.json`) -- with `theme_color`, `theme_fill` (and in
  a border's text) naming it. This applies to the styles in use and table styles too, and
  paragraph shading and borders now appear in a style's look. word_design_facts' definition
  grows 42 offline tokens (every definition 13,110 -> 13,152, the guard 13,155; the core
  unchanged; a reviewer's largest 7,519 -> 7,560); it is not strict.
## 0.1.15 -- unreleased

- **Why a refresh was slow** (production round 20, S2): where the render worker lays a
  version out whole although it was handed the layout of the version before, `render`'s
  result says why in `data.relaid` -- `reason` (docx2svg's `relayout.FALLBACKS`: `styles`,
  `sections`, `fields-open`...), `at` (the block it names, by the id the Word tools take: the
  block in the document now, or `p:<paraId>` of one the edit deleted), `field` (the
  instruction of a field an edited block opens and another closes: the contents' `TOC`) and
  `part` (the part that changed: `word/styles.xml`). Nothing is added where only what the
  edit reaches was laid out. The worker's facts carry docx2svg's cause as it is
  (`relaid_cause`); `docx_agent.layout.relaid_facts` maps it. With docx2svg 0.8.0 a document
  with a table of contents, `PAGEREF`, `PAGE`, `SEQ` or `REF` in its body is laid out from
  the edit too (S1), where 0.7.0 laid it out whole (`relaid: "page-fields"`). Tool
  definitions unchanged. docx2svg's floor is raised to 0.8.0.

## 0.1.14 -- unreleased

- **The render worker lays out again only what an edit reaches** (production round 19, R1):
  `render` keeps the latest layout on the document (as before) and, when the document has
  changed since, hands it to its worker as the layout of an earlier version
  (`("previous", laid_out)` as the task's `state`); the worker lays the new bytes out from it
  (docx2svg 0.7.0's `lay_out_docx(previous=...)`: from the page the first changed block is
  on until the pages line up again), the pages byte for byte those of a layout of the bytes
  alone. The pages an edit left as they were then key to the images the store holds and are
  not drawn again, as before; the worker's facts say which pages changed (`changed`) and why
  the whole document was laid out again where it was (`relaid`). The first edit of a
  document stamps an id on every paragraph, which docx2svg measures nothing again for. A
  different font folder or locale lays the document out whole, as before. Tool definitions
  unchanged. docx2svg's floor is raised to 0.7.0.

## 0.1.13 -- unreleased

- **Pages and blocks both ways, in this API's ids** (production round 19, R3):
  `DocumentLayout.blocks_on(page)` is the ids of the body's top-level blocks (paragraphs and
  tables, the ids the Word tools take) that a 1-based page draws something of, in document
  order; `DocumentLayout.pages_of(id)` every page a paragraph, table, table row or cell (or
  an alias of its id) is drawn on -- each page of one split across pages, each page a header
  row is drawn again on; `[]` for a block nothing of which is drawn. Both read docx2svg's own
  page <-> block map (`Page.blocks`, docx2svg 0.6.0: element paths derived from the XML) and
  map its keys onto ids through the live tree, once per layout. `page_of` and
  `ids_on_page` are unchanged. No tool needs it: `check`'s facts already say each entry's
  pages and `render` returns images, so the definitions are unchanged; an application
  showing which pages an agent's section covers, or jumping from a reviewer's entry to its
  page, calls `doc.layout()`. docx2svg's floor is raised to 0.6.0.

## 0.1.12 -- unreleased

- **A render worker does docx2svg's one-time work at its start** (production round 19, R2):
  `PageRenderer.warm_up` -- a new worker's first task -- calls `docx2svg.warm()`, which
  imports docx2svg's modules and scans the font folders (the system's, Word's and
  `OOXML_FONT_DIRS`), so the worker's first layout does not. docx2svg 0.5.0 also reads only
  each font file's header tables for that scan, not the whole file: a worker's first
  16-page layout is about a second faster here, and its memory 1.4 GB smaller. The pages
  are the same bytes (`tests/test_parallel_render.py`). docx2svg >= 0.5.0.

## 0.1.11 -- unreleased

- **Empty paragraphs between text** (production round 18, Q4): `check`'s layout facts
  (`include: ["layout"]`, the family the spacers are in) list `empty_paragraphs`: per table
  cell of the main story and for the body, the empty paragraphs with a paragraph of text on
  both sides -- `{"kind": "empty_paragraphs", "at": "t:1A/c1,0" | "body", "paragraphs":
  [positions among the container's paragraphs, 0-based], "ids", "fix"}`, every run of a
  container in its one entry. Empty: no text but whitespace (tracked deleted text counts as
  text) and no tab, break, drawing, object, field, note or comment reference, content
  control or equation. A table, a block-level content control, and an empty paragraph that
  carries a section break, a frame or a page break before end a run and are not
  neighbours, so the empties beside them are not listed; leading and trailing empties
  (the paragraph a cell or the body ends on among them) never are; a paragraph whose mark
  is a tracked deletion is passed over. `fix`: `word_delete` calls, one span per run whose
  paragraphs are siblings, else one per paragraph, the last run first; under tracking they
  delete as tracked changes, as every `word_delete`. Read from the document, so listed when
  the layout failed too; `pages` keeps those on the pages asked for. Headers, footers,
  notes and text boxes are not looked at (the layout facts' spacers are the body's too). A
  run in the body is also among the `spacers` (with its height and a `space_before`
  alternative); either fix deletes it. Kept in docx-agent (`layout_facts.empty_paragraphs`):
  pptx-agent's rule is over DrawingML, and the two share only the grouping of a run.
  Definitions unchanged; strict sets unchanged.

## 0.1.10 -- unreleased

- **A reviewer's `word_changes` without `ids` and `all`** (round 17; needs ooxml-edit
  0.69.0, `Tool.action_fields`): a role set offered `list` alone (`review`,
  `review`+`resolve`, `review`+`edit`) is no longer sent the two parameters only `accept`
  and `reject` take. The full definition keeps them. Offline estimates: review core 2,595 ->
  2,544, all 7,506 -> 7,455; review+resolve core 2,603 -> 2,552, all 7,514 -> 7,463; the
  default and accept sets unchanged. Strict sets: unchanged except `review`+`edit`, which
  now sends `edit_smartart` strict too (the optional parameters freed make room).
- Online strict test (`tests/test_tools_online.py`): several calls of the same tool are
  taken together, and a definition with `items` is asked for its first action other than
  `list` (which takes no items).

## 0.1.9 -- unreleased

Needs ooxml-edit 0.68.0 (`ooxml_edit.pictures`, the SVG recolouring shared with pptx-agent).

- **An icon from the application's asset catalogue** (round 17, P6a). `docx_format(assets=...)`
  (new, `deck_format`'s counterpart; the catalogue protocol is ooxml-edit's
  `tools.assets`, the one pptx-agent takes) offers `word_find_asset` (deferred, in
  `word_objects`) and `word_drawings` `insert_picture` `asset` in place of `image`; without
  a catalogue neither is sent. Inline, or floating with `wrap`, as any picture; `at` may
  be a cell (`t:1A/c0,0`: the start of its first paragraph); `size_pt` is read as
  `width`; an asset's `alt_text` defaults to its id's last part in words.
- **SVG pictures and their tint**: `Document.insert_picture` takes an SVG and stores it as
  Word stores one it inserts (measured in Word for Mac, 2026-10-26: the SVG part by
  extension behind an image relationship and the blip's SVG extension, a PNG fallback at
  twice the SVG's own size, the picture at its own size at 96 dpi, named "Graphic N").
  `color` (on `word_drawings` too) tints it as PowerPoint's Graphics Fill does -- a theme
  colour with its `MsftOfcThm_..._Fill_v2` class and the theme's value, or `RRGGBB` --
  and draws the fallback from the tinted SVG. Measured: Word draws the SVG, not the
  fallback, and resolves the class against the document's theme. A bitmap cannot be
  tinted (`invalid_arguments`, field `color`).
- **Minimum text sizes by kind** (P6b): `docx_format(min_text_size=...)` -- one number, or
  `{"body", "table", "footnote", "endnote", "header", "footer", "caption"}` points
  (`docx_agent.edit.minimums`; `caption` is the Caption style's own) -- sets
  `Document.min_text_size`, which `check` reads when it is given no `min_size` (its layout
  facts' `small_text`, each entry with its kind's `min` for a mapping) and the builders
  read too: `word_insert_markdown` and `word_edit_table` `set_cells` report `small_text`
  (and a warning) for the text they wrote that its styles make smaller. Unset (the
  default), nothing changes.
- Definitions: word_drawings' `color` +27 to every definition (13,083 -> 13,110; guard
  13,115), none to the core (2,816); with a catalogue configured, +151 more (13,261).
  Reviewers: 7,492 -> 7,519 (guard 7,520). The strict set is unchanged.

## 0.1.8 -- unreleased

- **`word_compare_blocks`** (round 17, P3; deferred, in `word_style`), the Word counterpart
  of `ppt_compare_slides` with a reference: `{doc, blocks: [...], reference: {doc?, block}}`
  compares each table or paragraph with the reference block -- a sample, in this document or
  another open one -- and lists every difference as `{block, property, cells | rows |
  paragraphs | columns, value, reference, fix}`, `fix` `{tool, calls}` (each call's
  arguments but `doc`) where a Word tool writes the reference's value, else `null`; paged
  by difference (`cursor`). Compared (`docx_agent.edit.compare`, as the tables state
  themselves): `header_rows`, `columns`, `rows` (body rows, the header rows apart; no fix:
  a filled copy has more by design), `merges` per row (split and merge calls), `widths` (the
  grid, within 0.5 pt: lengths are whole twips and every route to a width rounds to them;
  `WIDTH_TOLERANCE` says why), `width`, `table_style`, `table_alignment`, `table_shading`,
  `table_borders`; per cell `shading`, `borders`, `vertical_alignment`; per paragraph
  `paragraph_style` (by name, so another document's compare) and its direct formatting in
  `word_format`'s names (`clear` where the reference states none). Rows correspond by the
  reference's structure: its header rows one to one, then body row *i* to its body row *i*
  or its last. The tests run every fix call through the tools and compare again: nothing
  with a fix is left. New: `Document.compare_blocks`.
- **Fix: `word_format_table` `scope: "table"` with `borders`** set the table's style to the
  border's name ("the document has no table style 'single'"): the border style shadowed
  the `style` argument. Borders alone now leave the style as it was.
- Definitions: word_compare_blocks adds 258 to every definition (12,825 -> 13,083 by the
  offline estimate; the guard rises to 13,085), none to the core (2,816). A reviewer is
  sent it too (a reading tool): every definition a reviewer is sent 7,234 -> 7,492 (guard
  7,495). The strict set is unchanged.

## 0.1.7 -- unreleased

- **`word_copy_from` within one document** (round 17, P4): `source_doc` may be `doc` (it was
  refused, "copy from another open document"). `Document.copy_blocks(self, ...)` copies from
  a snapshot of the document as it is before the call and inserts as from another
  document: bookmarks renamed (or dropped), colliding paragraph, drawing, note, comment and
  control ids re-issued, lists, comments, notes and tracked changes as a copy from another
  document takes them; one undo step. A table can be copied after itself
  (`at: "after:t:1A"`), or over itself (`replace:`). The result is the same, part for part,
  as copying from a second open copy of the document (`tests/test_e6_copy.py`).
- `copy_blocks` with `replace:` (untracked) now takes the replaced blocks out before it puts
  the copy in, as the tracked replace did, unless nothing follows them in their container
  (a cell's last blocks): a copy that kept the ids of the blocks it replaced had them
  re-issued after the result's `copied` and `blocks` named them, so those named ids that
  were no longer there.
- `word_copy_from`'s description says "an open document, this one too" (every definition
  12,824 -> 12,825 by the offline estimate, the guard; the core unchanged).

## 0.1.6 -- unreleased

- **What names a block, in `word_read` and `describe`** (round 17, P5), so a sample table or
  control can be found without reading the XML. A table's id comment adds its alt text when
  it has some: `alt: ...` (its title, `w:tblCaption`) and `description: ...`
  (`w:tblDescription`), after `table style: ...`, on one line and cut at 120 characters;
  `<!-- t:3C table style: Report table alt: sample:record-table -->`. An inline content
  control's comment adds `tag: ...` and `title: ...` (`w:alias`), as a block control's
  already did. `describe` adds `alt` and `description` to each table that has them, and
  `tagged` (`[{id, kind, tag?, title?}]`: each content control with a tag or a title, at
  most 30; left out when there is none). New: `Table.alt_text`
  (`{"alt"?, "description"?}`). Golden transcripts: w11-contract-form's `describe` result
  is re-recorded with `--new-keys tagged` (its data less that key as recorded).
  Tool definitions are unchanged.

## 0.1.5 -- unreleased

- **`edit_chart` and `edit_smartart` as in a deck** (review, S10; needs ooxml-edit 0.67.0).
  The handlers (`tools/shared.py`) are now ooxml-edit's `charts.actions` -- the data
  actions, the missing-argument messages (`_chart_need` is gone), `add`'s checks and look,
  the SmartArt actions and the results -- with only finding the drawing and its `id` here.
  What a model sees change, as the user chose for both formats:
  - `add_category` takes the new label from `text`, as the shared definition says (it read
    `category`; a call giving only `category` is now "add_category needs text"), and its
    `values` may be left out (blanks), as `add_series`' may;
  - `set_legend` without `position` puts the legend on the right (it was refused), and a
    label position (`center`...) is refused (it was taken as the right);
  - `edit_smartart` `set_text` without `text` sets empty text (it was refused), and every
    result's `nodes` is `[{n, text, level}]` (it was each node's text, cut at 80
    characters);
  - every `edit_chart` result has `number_formats` (`values`, `data_labels`, `value_axis`)
    and, for a bar or column chart, `gap_width`; numeric categories are listed as numbers
    (they were text; dates are still ISO dates); `read` adds `axes` (each axis' title) and
    drops `number_format`, which is `number_formats.values`.
  Golden transcripts: w3-figures-update's `add_category` call now gives the label as
  `text` (it gave `category`), and its saved document is byte for byte the one recorded;
  the `data_sha`/`data_keys` of w3's, w9's and w14's chart results are re-recorded, each
  checked to differ only by the result keys above. `tests/test_chart_cases.py` runs
  ooxml-edit's table (`charts.cases`) through these tools, as pptx-agent's does through its
  own. Tool definitions are unchanged.

## 0.1.4 -- unreleased

- **Error codes with ooxml-edit's typed refusals** (review, S9; with ooxml-edit 0.66.0;
  the tests skip on an older one, and the requirement stays `>=0.65.0`).
  ooxml-edit now raises an `ArgumentError` for a value it refuses and a `UnitError` for one
  out of its range (`ooxml_edit.errors`), and its dispatcher says `invalid_arguments` and
  `unit` for them. What a model sees change: `edit_smartart` removing a diagram's last node
  is `invalid_arguments`, not `unit`, as in a deck. The library's own errors keep their
  codes (`ERRORS` is unchanged: `FormattingError` `unit`, `SpanError`, `PictureError` and
  `EditError` `refused`, `AnchorNotFound` `not_found`, `AmbiguousAnchor` `ambiguous`), and
  nothing here mapped an ooxml-edit `ValueError` that the typing makes redundant
  (`edit_chart`'s catch of `ChartDataError`, not one of the typed classes, stays). New
  `tests/test_tool_errors.py`. Tool definitions are unchanged.

## 0.1.3 -- unreleased

- **A linter, and `py.typed`** (review, S8): ruff checks pyflakes and a few bug-finding rules
  (`[tool.ruff]` in pyproject.toml; no formatter, no import sorting), as a step of CI's build
  job, pinned in the dev extra. What it found is fixed: unused imports and locals (the render
  tool's unused `entry`, `revisions/track.py`'s unused `rest`, the tests and the golden
  graders) and an f-string without a placeholder. `edit/charts.py`'s `_no_edit` keeps the
  `yield` after its `raise`: never reached, it is what makes the function a generator, which
  `@contextmanager` needs; a comment now says so. The wheel and sdist now carry `py.typed`,
  so a caller's type checker reads the annotations. mypy is not run here yet. Tools behave
  as before.

## 0.1.2 -- unreleased

- **The render plumbing is ooxml-edit's** (review, S5): a new `tools/render.py` is a thin
  adapter, `PageRenderer` (`RENDERER`), over docx2svg for ooxml-edit 0.65.0's shared
  pipeline (`ooxml_edit.tools.render`), which pptx-agent's slides use too. The adapter says
  what is a document's: laid out whole, once, in one worker, and its pages rasterised over
  the workers at once; the raster fingerprint (the rasteriser's font files, embedded faces,
  the sans-serif face); the `data-docx-*` attributes left out of the key; the layout kept
  on the document and handed back to the next call's worker. The rest -- `draw`'s cache
  lookups and keeping, the chunks, the worker tasks, the keys, the warm-up, the contact
  image's task, `png_size` -- is the shared copy; `tools/shared.py`'s `lay_out_pages`,
  `rasterise_pages`, `rasterise_tiles`, `warm_up`, `_code_stamp` and `png_size` are gone
  (`render_pages`, the serial renderer, stays). The same images byte for byte, the same
  keys, the same tool definitions and token counts; what changes, as pptx-agent did
  already: `png_size` and every PNG the renderer returns are checked to be a PNG
  (signature and header), and the warm-up's and the contact image's warnings are left out
  in the worker. Render timings unchanged within noise (4 pages at 1,000 px: median
  2,713 ms before, 2,709 ms after; a 3-page contact image 2,406 ms before, 2,354 ms
  after). `tests/test_parallel_render.py` calls the shared `lay_out_views` with `RENDERER`
  (the same assertions, the shared shapes). ooxml-edit floor 0.65.0.

## 0.1.1 -- unreleased

- **The Office theme read from ooxml-common, not copied** (review, S3): `edit/blank.py`
  writes a new document's theme -- its colours, Aptos Display / Aptos, the per-script
  faces, the format scheme's gradients and line widths -- from
  `ooxml_common.drawingml.theme`, shared with pptx-agent's new deck, instead of its own
  copy (`THEME_COLORS`, `MAJOR_FONT`, `MINOR_FONT`, `SCRIPT_FONTS` and `LINE_WIDTHS` stay,
  as the shared values). A test compared the copy with the shared data, once, on the pull
  request, before it went. The same theme as before, byte for byte. ooxml-common floor
  0.31.0.

- **One lookup for `edit_chart`'s series and category, name first** (review, N4): the
  handler finds them, and `edit_smartart` its node, with ooxml-edit 0.63.0's
  `charts.actions.resolve_series`, `resolve_category` and `resolve_node`, the lookup
  pptx-agent's handler uses too; this repository's own `_series` and `_category` are gone.
  What changes: a series NAME now wins over a position -- a number from 0 is a position
  only when no series has that name, so a chart whose series are called "1" and "2"
  means the same in a document and a deck (`series: "1"` took series 1 here before). A
  name is matched with whitespace collapsed, then full-width forms and case ignored; a
  name two series or categories share is refused (`label_not_found`, give the position)
  instead of taking the first; an unknown name is `label_not_found` (was `not_found`),
  with the names in `valid_options`; a missing series or category lists them too; a
  position past the end stays `not_found`. A formatting action's `category` is read only
  with a `series`, as in a deck. `edit_smartart`'s `node` past the end is `not_found`
  with the node numbers. The tool definitions are unchanged. ooxml-edit floor 0.62.0 ->
  0.63.0.

## 0.1.0 -- 2026-10-25

- **Versions that mean something** (review, N3): this is 0.1.0, the first version tagged
  (`v0.1.0`). It asks for `docx2svg>=0.2.0` (the `png` extra too), docx2svg's first tagged
  version, which always gives a layout its coverage, so the paths for a docx2svg without
  one are gone: `DocumentLayout.coverage` is docx2svg's, never `None`; `coverage_facts` no
  longer takes `stopped` (it answered from the stop alone when there was no coverage); a
  render's `coverage` is always given. A coverage given as a dictionary still may leave
  out `status`, `approximations` and `placeholders`. The same facts for every supported
  install. Each version that reaches `main` is tagged by the new tag workflow, and CI's
  `build` job checks that `pyproject.toml`'s version is the one this file's top version
  section names (`tools/check_version.py`). The dependency notes moved from
  `pyproject.toml` to CONTRIBUTING.md, "Dependencies".

- No private name of a sibling is imported any more (review, N2): docx2svg's
  `parse.document.cell_conditions` (was `_conditions`), ooxml-common 0.28.0's
  `imagemeta.pixels_and_density`, ooxml-edit 0.62.0's `charts.create.check_options` and,
  in `tools/charts_probe.py`, `charts.diagram.DiagramModel`; the tests patch docx2svg's
  `lay_out_docx` only. Each owner's `tests/test_public_api.py` now holds every name
  docx-agent imports from it. No behaviour change.

- CI runs nightly (03:17 UTC) and on demand, so a sibling's change that breaks this
  suite shows the next morning; ooxml-edit's own CI now runs this suite on its pull
  requests too. No code change.

- `word_design_facts` (round 16, F9), the counterpart of pptx-agent's `ppt_design_facts` for
  analysing a template: measured facts with ids, no verdicts (`Document.design_facts`,
  `docx_agent.edit.design`). `styles`: every paragraph style in use per role (`title`,
  `heading 1`-`9`, `body`, `list`, `caption`, `table`, `toc`, `header`, `footer`,
  `footnote`, `endnote`), its paragraphs, empty ones and an example id, and its look after
  inheritance from docx2svg's cascade (font, size, colour, bold and italic, alignment,
  spacing before, after and between lines, indents, keep options, outline level, numbering
  format and label; in a cell with the table style). `table_styles`: each table style in
  use, its tables, their style options, and its places -- header row, row bands, first
  column, last row and column when on, the corner -- each with the conditional formats
  applied, fill, the four borders after the conflicts between neighbours and the text, all
  from docx2svg's table resolver (`resolve_table`, `edge_winner`, its conditions) over a
  five-by-four sample. `sections`: page size, orientation, margins, columns, start and the
  header and footer each variant shows. `theme`: name, the twelve colour slots, the major
  and minor faces. `patterns`: one block shape under two or more headings of a level (a
  paragraph by its style, a run of one style once, a table by style and columns), with
  counts and example spans. Lists are capped (`omitted` counts the rest) and a result over
  the call's limit keeps fewer entries. `effective.style_look` is the cascade's look of a
  style with no direct formatting. Deferred, in `word_style`; a reading tool, so a
  reviewer's too. Definitions +169 estimated tokens over F8's: every definition 12,655 ->
  12,824 (guard 12,656 -> 12,825), the core unchanged (2,816), every reviewer definition
  7,064 -> 7,234 (guard 7,065 -> 7,235), the reviewer's core unchanged (2,608); not strict,
  the strict set unchanged.

- `check` takes `layout` in `include`, and `min_size` and `styles` (round 16, F8; ooxml-edit
  0.61.0, now the floor): layout facts, measured from the one layout reflow uses and render
  draws, and from docx2svg's style cascade (`docx_agent.layout_facts`) -- the smallest text
  drawn per kind (body, table, footnote, endnote, header, footer) and, with `min_size`,
  every paragraph's text below it (`fix`: clear a direct size where the style's is at
  least the minimum); direct formatting that makes a paragraph or all its text differ from
  its styles (paragraph, table, character), in word_format's names with the styles' values
  and the `word_format` call that clears it (bold on part of the text named in `in_part`,
  not cleared); with `styles`, the paragraph and table styles in use that the list does not
  name, and the listed names the document lacks; runs of empty top-level paragraphs used
  as space, with how far the next block would move without them, the `word_delete` call and
  the `space_before` that keeps the next paragraph in place (between paragraphs the larger
  of the two spaces counts, docx2svg's measured rule); keep-with-next paragraphs drawn last
  on a page, their next block on a later one; rows split across pages (no tool sets a row's
  "allow to break" yet: no fix) and tables on several pages whose first row is not drawn
  again (`word_format_table` `header_rows: 1` where none is marked); pictures, drawings and
  tables (as drawn) past the margins (`word_drawings` `resize` for an inline picture past
  the right one; a table only beyond Word's own outdent -- its first cell's left margin and
  half its outer border, the last cell's on the right, from the resolved table -- measured
  in Word 16.106: below mode 15 a default table's text is on the margin, in mode 15 its
  border, and a tblInd of -720 puts it 36 pt out in both); and `uncertain`, the pages with placeholders, approximations or stories
  not drawn, with `unknown_from` where a layout made to stop stopped. At most 20 of each,
  most telling first, `<fact>_left_out` the rest; `pages` keeps the facts on those pages.
  Giving `min_size` or `styles` asks for the facts; `layout` alone sets no baseline for a
  later reflow's `since_last_check`. On the 40-page made-up document of round 16, check with
  reflow and layout took 7.20 s against 6.89 s for reflow alone (best of 3; 0.38 s for the
  facts on a layout already made). Measured in Word 16.106 for Mac on made-up probes: a
  keep-with-next paragraph before a page-break-before paragraph stays at its page's foot, a
  90-line chain of them breaks where docx2svg breaks it, one that fits moves; a row taller
  than a page runs over four pages as docx2svg draws it, and a marked header row repeats.
- `word_format` takes `clear` (round 16, F8): the direct properties named (word_format's own:
  `size`, `font`, `bold`, `alignment`, `space_before`, `indent_first`...) removed, back to the
  style's; giving a property a value and clearing it is refused. What check's
  `direct_formatting` fix calls.
- Definitions: core 2,749 -> 2,816 estimated (+67, check; about 4,168 counted), every
  definition 12,493 -> 12,655 (+95 for word_format's `clear`; guard 12,494 -> 12,656), a
  reviewer's core 2,541 -> 2,608 (guard 2,542 -> 2,609), every reviewer definition 6,997 ->
  7,064 (guard 6,998 -> 7,065). The strict set's members are unchanged.

- **docx2svg carries on past what it cannot lay out** (round 16, docx2svg's `carry-on-r16`):
  the coverage of `check` (`reflow`), `render` and `save_document` lists `placeholders` --
  page, reason, block id -- of what it drew as a marked box instead of laying it out (any
  makes it `partial`), and a frame or a nested floating table that stopped the layout is now
  an approximation; pages past them are known. `Unknown` for a block past a stop remains for
  a docx2svg made to stop (`docx2svg.paginate.CARRY_ON = False`).

- `save_document` takes `clean` (round 16, F11; ooxml-edit 0.60.0, now the floor): the file
  saved without what Word keeps but never draws -- orphaned parts (a zip entry no content
  type is declared for, `[trash]/0000.dat`; a part nothing reaches), every `w:rsid*`
  attribute and settings' `w:rsids` list, webSettings' pasted `w:divs` with every
  `w:divId` (kept whole when a division sets a margin or a border: Word draws those,
  measured), an external `attachedTemplate` and its relationship
  (`docx_agent.edit.cleaning.clean`). The result's `cleaned` gives each kind's count and
  bytes and the file's sizes before and after; the summary the sizes. The session's
  document is unchanged. Measured in Word 16.106 on made-up documents: rsids, an external
  template link, 400 pasted divisions and an unreferenced part taken out (64,769 -> 13,066
  bytes), the cleaned file opens without a prompt and its PDF is the original's pixel for
  pixel on all 9 pages; with a `[trash]/0000.dat` entry too Word does not open the original
  without a dialog and opens the cleaned copy (130,433 -> 13,066 bytes); a division with a
  margin and a border is drawn, so it is kept. A Markdown save refuses `clean`. Definitions
  +42 estimated tokens: core 2,707 -> 2,749, every definition 12,451 -> 12,493 (guard
  12,452 -> 12,494), reviewer core 2,499 -> 2,541 (guard 2,500 -> 2,542), every reviewer
  definition 6,955 -> 6,997 (guard 6,956 -> 6,998); the strict set's members unchanged
  (save_document is strict on OpenAI only, `clean` a nullable boolean there).

- `Result.touched` names a body change's heading sections (round 16, F12):
  `{"doc", "part": "word/document.xml", "heading": "p:301", "level": 2, "blocks":
  ["p:301", "t:7"]}` -- the innermost heading before each changed block (the block itself
  when it is one), its level and the first and last block of the section it owns, as
  `section_blocks` and `describe` read them; `heading: null, level: 0` before the first
  heading; a change in no block (the last section's properties) and every other part as
  `{"doc", "part"}`, as before (`docx_agent.tools.touched`). Blocks are compared by their
  XML less the ids stamping writes, so the first edit's stamping reports no section; the
  blocks are found from the part's bytes, only the stretch a call changed read again: on a
  made-up 41-page document (444 blocks, 272 KB) 4.5 ms a call (median; 10.7 ms the first),
  and an edit's median 105 ms with it, 100-132 ms without. Pages are not named: which
  pages a change moved is known only from a new layout (render after the change and
  compare the pages' images by content). Application-only; no token change.

- `new_document`'s result has the new document as its `doc` (round 16, F10; ooxml-edit
  0.60.0's dispatcher), `open_document`'s too; it was `None`.

- `SEQ` with its switches before the identifier (round 16, F4; docx2svg's
  `fields.sequence_switches` and `sequence_step`, shared): `SEQ \# "00" Item` is 01, 02 as
  Word draws it, where it counted a sequence named nothing and drew 1. Measured in Word
  16.106 (45 fields, docx2svg ROADMAP H.8): a switch's argument is the next token whatever
  it is, a `\#` after a `\*` is Word's error text (left as it was and listed as unknown),
  a numeric picture's `#` with no digit a space.

- `check` takes `contrast` (ooxml-edit 0.59.0, round 15, T2): a deck's text below that
  contrast ratio with what is behind it; a document refuses it (`invalid_arguments`,
  `field: "contrast"`), as it does `spacing`. ooxml-edit 0.59.0 is the floor. Definitions
  +37 estimated tokens: core 2,670 -> 2,707, every definition 12,414 -> 12,451 (guard
  12,415 -> 12,452), reviewer core 2,462 -> 2,499 (guard 2,465 -> 2,500), every reviewer
  definition 6,919 -> 6,955 (guard 6,920 -> 6,956).

- Tools docs: layout runs in worker processes, so a script keeps its work under
  `if __name__ == "__main__":`; without the guard the toolbox no longer waits two minutes a
  call but lays out in subprocess workers, saying why (round 13, R6; ooxml-edit 0.58.0, no
  change here).

- What a call changed (round 13, R2; ooxml-edit 0.57.0): every tool result's `touched`, for
  the application and never in what the model reads, lists `{"doc", "part"}` for each part
  whose bytes changed (`word/document.xml`, a header, the comments, the numbering), undo
  and redo included -- ooxml-edit's default, no docx-agent code. Finer entries (the
  paragraphs and tables by id) would need a format mapping that compares a changed part's
  blocks by their `w14:paraId` and table ids; not built. No token change.

- `check` takes `spacing` (ooxml-edit 0.56.0, round 12, Q1): a deck's tightest gaps,
  padding and frame clearances in points; a document refuses it (`invalid_arguments`,
  field `spacing`), as it refuses `boxes`. The shared definition costs 35 estimated tokens:
  the core 2,635 -> 2,670, every definition 12,379 -> 12,414 (guard 12,415); a reviewer's
  core 2,427 -> 2,462 (guard 2,465), all 6,884 -> 6,919 (guard 6,920); the strict set
  unchanged.

- A contact sheet whose tiles do not fit the 2,576 px edge is drawn with narrower tiles,
  not refused (round 10, O4; ooxml-edit 0.55.0): too wide, the widest that fit the columns;
  too tall, the widest for the pages -- before drawing at the shortest page the sections
  declare (for the pages named, or the page count a render of this version found), and
  once drawn at the pages' own heights, drawn again narrower if need be (the layout kept:
  rasterising alone). `sheet.shrunk` (`{"from", "why"}`) and the summary say so; only tiles
  that would not fit even at 160 px are refused. Pages drawn with two faces the rasteriser
  files under one name (Aptos and Aptos Display) are drawn apart by docx2svg now
  (`face-not-addressable` no longer raised for them).

- `render` lays a version out once, not once a call (round 10, item 3): the layout comes
  back from the worker with the pages (pickled: the layout and the faces it was measured
  with) and is kept on the document by the bytes' digest, the latest only; the next call
  drawing other pages of the same bytes hands it to its worker, which writes the pages from
  it (`lay_out_pages(..., laid_out=...)`). 36-page sample-long, nine calls of four pages
  after a one-page edit: 5.0 s (17.4 s before; each call but the first 0.1 s); the first
  render of all 36 in nine calls 41.8 -> 29.5 s. Same PNGs. docx-agent now calls docx2svg's public
  `lay_out_docx`, `page_svgs` and `user_font_dirs` (docx2svg's round-10 API) instead of its
  private names.

- `render` lays a document out once and rasterises its pages over every worker at once,
  and keeps each image by its content (round 10, O1 and O3; ooxml-edit 0.52.0,
  ooxml-common 0.24.0): one worker lays out and writes the pages' SVGs (`lay_out_pages`),
  each keyed by its SVG -- less the `data-docx-*` attributes the rasteriser never reads --
  and the rasteriser's arguments (font files with their times and sizes, embedded faces by
  digest, options, the renderers' code); the pages whose image the toolbox's memory store
  does not hold are rasterised in chunks over the workers (`rasterise_pages`). A repeat of
  an unchanged document asks no worker; after an edit only the pages whose key changed are
  rasterised. The page count comes from the same layout (no second conversion). The result
  now has `cached` (the pages not drawn by the call); a contact sheet reuses the tiles and
  says which (`cached_tiles`). PNGs are `render_pages`' byte for byte.
- Refusals before drawing: the image budget (`call.check_images`), and a sheet of named
  pages too tall even at the shortest page the sections declare (`w:pgSz`); one still too
  tall once laid out is refused as before. `FORMAT.warm_up` warms each new worker outside
  any call's deadline. The main process makes a version's bytes once
  (`DocumentEntry.snapshot`). Definitions are unchanged; three golden transcripts
  re-recorded for the new `cached` key (`tools/refresh_goldens.py --new-keys cached`).

- `render` with `sheet` (round 9, N1; ooxml-edit 0.49.0), as for decks: up to 30 pages as
  one image, tiled in order, each labelled `page 3`; `pages` left out is every page up to
  30. `sheet.width` is each tile's (160-1280, default 360; the call's `width` stands in),
  `columns` 1-8 (default 4). Refused, naming the bound and what to do: more than 30, a page
  that is not there (as now), a sheet over 2,576 px wide or -- once the pages are drawn and
  their heights known -- tall, naming how many pages fit. `coverage` as for a render; the
  sheet is cached whole. Definitions: +64 to the core (2,570 -> 2,634) and every definition
  (12,314 -> 12,378; guard 12,380); a reviewer's core guard 2,370 -> 2,430, all 6,825 ->
  6,885; strict set unchanged.
- `edit_chart` (round 9, N2; ooxml-edit 0.48.0): a `scale` without `axis` -- in
  `set_axis_scale` or `add` -- is the value axis where it is the chart's one axis that takes
  bounds and units (column, bar, line, radar, dot); a scatter's x and y, a date axis or a
  secondary value axis must still be named, and the refusal names them, as it does for a
  scale of only `reverse` or `number_format`. Every "needs" refusal names what to give, where
  it said "series is required": `set_value needs category: 'A', 'B' or 'C' (or a number from
  0 to 2)` (the first 12 and how many more), the series, the chart's axes, the values' count,
  the legend positions, what `add` needs. Definitions unchanged.
- What a name that is no Word tool means (round 9, N5; ooxml-edit 0.47.0): "no tool
  'word_add_comment': use word_comments with action add", in a batch's op too, with the
  best few tools in `valid_options` instead of every one; an editor without the accept role
  is told `word_changes accept` needs it. ooxml-edit reads the tools' names and actions;
  `TOOL_HINTS` (`FORMAT.tool_hints`) adds the names it misses or settles wrongly: a link is
  `word_links add_hyperlink` (not a header's `link`), an image `word_drawings
  insert_picture`, a table, heading or paragraph `word_insert_markdown`, a page break
  `word_format` (`page_break_before`), margins, orientation and columns `word_sections
  set`, a checkbox or drop-down `word_controls insert`, a header, footer or page number
  `word_headers_footers set`. A batch op may carry `comment`/`note` (N4), kept in the call
  log. Data and tests only: definitions and the strict set unchanged.
- Several agents on one document (round 8, M1 and M2; ooxml-edit 0.46.0): a builder's, a
  reviewer's and an editor-in-chief's views of one shared session at once
  (`session.view(roles=..., context=...)`), each held to its own roles, in a batch too; the
  role line's accept sentence decided over the tools sent (`system_prompt(tools=...)`).
  Tests and the tools README only: no Word tool or definition changed.
- Tool roles (round 7, item 9; ooxml-edit 0.45.0): the Word tools are tagged for the
  additive roles `review`, `resolve`, `accept` and `edit` (`Toolbox(roles=...)`,
  `toolbox.session(roles=...)`). `word_comments` takes its actions' roles from
  `shared.COMMENT_ACTION_ROLES` -- list, add, reply `review`; resolve, reopen, delete (and a
  reply with `resolve: true`) `resolve`; editing a comment's text every role's (the user's
  decision: its author is not known) -- and joins the core of a role set without `edit`.
  `word_changes`' accept and reject are the `accept` role's, its list `review`'s, and it
  joins that core too (an editor-in-chief, `["review", "accept"]`, settles the tracked
  changes and changes nothing else). The `list` actions of `word_sections`, `word_fields`,
  `word_notes`, `word_links`, `word_drawings` and `word_controls`, and `word_styles`' list
  and describe, are `review`'s. `word_tool` takes `action_roles`, `call_role` and
  `review_core`. The default, every role, sends the definitions byte for byte as before
  (core 2,570, every definition 12,314 estimated). Estimated core / every definition:
  `["review"]` 2,350 (about 3,478 counted by the proxy) / 6,807; `["review", "resolve"]`
  2,358 / 6,815; `["review", "accept"]` 2,356 / 6,812; all three 2,364 / 6,820 -- guards
  `REVIEW_CORE_GUARD` 2,370 and `REVIEW_ALL_GUARD` 6,825 (tests/test_tools_roles.py);
  `["accept"]` 814 / 1,418; `["edit"]` 1,418 / 10,653. Claude's strict set: a reviewer's
  `word_comments`, `word_changes`, `word_notes`; `["edit"]` alone loses `word_changes` and
  gains `edit_smartart` and `set_properties`; the default's and `["review", "edit"]`'s are
  unchanged. The online check counts the review, review+resolve and review+accept cores
  and sends their strict tools.
- The loaded core's budget is now 6,500 tokens as count-tokens counts them (raised from 6,000
  by the user on 2026-10-17); the offline proxy and the online check use it.

- An empty `items` list is a no-op with a warning, not a silent or refused call (round 7,
  L7; ooxml-edit 0.43.0's `Result.nothing_given`): `word_set_text`, `word_insert_text`
  and `word_comments` (other than `list`) given `items: []` change nothing, record no
  undo step and warn `items is empty: nothing was written` (inserted, changed) -- what a
  model sends after trimming a batch. `word_comments` without `items` still needs them;
  `word_format`'s and `word_lists`' `targets` keep `minItems: 1`; a drop-down's choices
  are not a list of work and are unchanged. Tool definitions unchanged.
- `batch` runs a tool the application registered in the toolbox with these (round 7, L6;
  ooxml-edit 0.43.0), tested here with a made-up one.

- A dot chart (round 6, R10; ooxml-edit 0.40.0): `edit_chart` `add` with `chart_type:
  "dot"` (and `Document.insert_chart(..., "dot", ...)`) draws one marker per category on
  the value axis, no line -- a fixed 1-5 scale with the values as labels is one call
  (`axis: "value"`, `scale: {"min": 1, "max": 5}`, `data_labels: true`). A native line
  chart, so Word edits it as one: each series' line off and a 5 pt circle marker in the
  series colour (the marker Office gives a new scatter series, measured); `colors` colour
  the markers. `edit_chart read` reports any line chart whose series all have no line and
  a marker as `dot` (`Chart.is_dot`); `set_values` and `add_series` keep it one. The
  Markdown outline's chart line still says `line` (it reads the chart model's plot types).
  The `chart_type` enum gains `dot`: +2 estimated tokens to every definition (12,312 ->
  12,314), the guard stays 12,315; edit_chart is sent non-strict.
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
