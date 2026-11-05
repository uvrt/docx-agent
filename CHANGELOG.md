# Changelog

docx-agent has not been released to PyPI. Each version that reaches `main` is tagged
`v<version>`, from 0.1.0, the first tagged. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail of every phase.

## 0.1.42 -- unreleased

- **Spacing around punctuation by the text's language** (production round 35, R21;
  `docx_agent.edit.typography`). Text a call writes -- every tool that writes text:
  `word_insert_markdown`, `word_set_text`, `word_insert_text`, `replace_text`, table cells,
  comments, notes, headers and footers, text boxes, content control values -- in a run whose language is
  French gets a no-break space (U+00A0) before `:` `;` `!` `?` and inside `« »`: an ordinary
  space it wrote there becomes U+00A0, a missing one is put in next to what it wrote. Text
  that was there before is not changed; marks between words (`12:30`, `x?y`, URLs), code
  (Markdown code, the code styles), text Word does not proof (`w:noProof`) and field
  results are left alone. The result says how many spaces it wrote. Canadian French gets
  the guillemets only; other French variants take French's rule. `docx_format(typography=False)`
  (`Document.typography`) writes text as given. Measured in Word (`tools/typography_probe.py`,
  `tests/observations/typography-word.json`): French proofing flags a missing space and
  accepts an ordinary one, U+00A0 and U+202F; Word breaks a line before a mark after an
  ordinary space and never after U+00A0 or U+202F; AutoFormat writes U+00A0 inside
  guillemets; a run's language is its own `w:lang`, its styles', `w:docDefaults'` -- not
  `w:themeFontLang`.
- **`check`'s `punctuation_spacing`** (in `layout`; also where the layout fails): per
  paragraph of any story, the places still without the spacing -- `breakable` (an ordinary
  space), `missing` -- and, for English, German, Dutch, Spanish, Italian and Portuguese,
  the spaces their proofing flags (`space_before`, and `space_inside` guillemets for the
  last three; reported, never removed on writing), with `replace_text` fix calls. Scoped
  by `pages` and `range` as the other facts.

## 0.1.41 -- unreleased

- **`check` and `render` on some blocks** (production round 35, R18). `check(range=[...])`
  -- each range any address `word_read` takes: a block, a span `p:A..t:B` (`p:A..` to the
  end), a cell, a section `s:2`, a bookmark; several in one call -- reports the facts about
  those blocks: the layout facts of their paragraphs, tables and drawings, measured on the
  pages they fall on (`smallest_text` of their text only), `validate`'s problems found in
  them (`problems`, with `elsewhere` counting the rest; `docx_agent.validate.problems_in`:
  what one of their elements carries, or what their XML alone shows again), and each
  range's pages (`range: [{"range", "blocks", "pages"}]`). A range adds `layout` to the
  default facts. `render(range=[...])` draws the pages the blocks fall on, as `pages`
  does; on more pages than a call's images, one sheet (said in the warnings). With
  `pages` too, those of them. The layout is the document's shared one, laid out with the
  lock let go where the call may and on its own snapshot where it holds the lock (round
  34's rule; a test runs scoped checks lone and in a `read_batch` at once). A range that
  names nothing is refused with the body's sections as spans (`docx_agent.scope`).
- **Copies against their reference** (R19). `check(copies=[{"blocks", "reference": {"doc",
  "block"}}])`: each copy -- a block or span made from a sample, the reference in this or
  another open document -- compared with its reference block by block
  (`docx_agent.edit.copies`): whether it `follows` (every difference its own), the
  differences with their fix calls and `kind` (`own`, `structure`, `changed_by_copy`,
  `format`), and `example_text`: the copy's paragraphs still holding the reference's own
  sentences (of three words or twenty characters; not its header rows' labels). One
  `word_follow_sample` call per reference block makes every copy that drifted follow
  (`fix`). Each comparison is kept by the digests of both sides' blocks and of each
  document's styles and theme, not by the document's version, so an edit elsewhere finds
  it kept (`kept`, per copy `cached`); 128 per document, least recently used dropped.
  The reply keeps to a budget (20 differences a copy, 100 a reply; the copy's own
  counted, not listed) and says how to read the rest.
- **A comparison reads each style sheet once.** `word_compare_blocks`, `word_follow_sample`
  and the copies above asked the style sheet for a style's name for every paragraph
  compared, each time listing every style: 1.6 s for a copy of 10 blocks, now 0.13 s
  (`compare.styles_read`).
- The tool definitions are ooxml-edit's: `check`'s `range` and `copies` and `render`'s
  `range` reach the model once ooxml-edit's shared definitions carry them; the handlers
  take them now (tests add them to the definitions as ooxml-edit is to).

## 0.1.40 -- unreleased

- **No wait for a layout in progress while holding the document** (production round 34,
  R17). 0.1.38's shared layouts and 0.1.39's unlocked layouts could deadlock: a lone
  `check` laid the document out with its lock let go, then waited to take the lock back
  *before* ending its layout's flight, while a call holding the lock -- a `read_batch`'s
  `word_inspect` or `check`, a batch's `check`, any layout under the lock -- waited for
  that flight. Every other call on the document then queued for the lock. Now:
  - a flight's owner ends it (result or error) before it takes the document's lock back;
  - only a caller that let the document's lock go waits for a layout in progress; a
    caller holding it (and `Document.layout()`, which cannot tell) lays out on its own,
    from the document's last layout.
  The order of the locks and waits is written down in `docx_agent.layout`'s docstring.
  Tests run lone checks beside a `read_batch` of `check` and `word_inspect`, a batch with
  a check between its changes, `word_fields` update and nine callers at once, each joined
  with a deadline (a deadlock fails the test, never hangs the run).

## 0.1.39 -- unreleased

- **No layout under the lock** (production round 33, R16c; ooxml-edit 0.79.0). A `check`
  inside a batch and `word_fields` update laid the document out holding its lock, so
  every other call on it waited.
  - A `check` after a batch's last change now runs once the batch is done, as a check
    alone does: it lets the document go while it lays out (ooxml-edit runs a batch's
    reads after its last change after its undo step). A check *between* a batch's changes
    still holds the document, since it reads the state between them, and says so with
    the way out (`check laid the document out between the batch's changes, holding it:
    ... Put check after the batch's last change (or call it alone)`).
  - `word_fields` update does its work *ahead* of its change (ooxml-edit's `Tool.aheads`):
    every pass and its layout runs on a copy opened from the document's bytes, with the
    lock let go, and the copy's layouts become the document's
    (`docx_agent.layout.twin`/`adopt`, each layout kept under the bytes it is of). The
    update under the lock then writes what the copy wrote and finds each pass's layout
    cached. If another call changed the document meanwhile, the work ahead is done again
    from there (at most twice); after that the update lays out what changed, under the
    lock.
  - A page cross-reference (`insert_cross_reference` with `kind: "page"`) and
    `word_template` `upgrade_to_modern` lay the document out ahead the same way.
  - The work ahead never changes what a call does: anything that fails there is left to
    the call, which meets it again and reports it.
  - The layout engine is taken with the bytes, before the lock is let go: another call
    ending meanwhile pointed the document's converter elsewhere.
  - On the made-up 79-page report (tool layer, worker pool), with a one-paragraph read
    started 0.3 s into each call:
    - batch `[edit, check]`: the document held 3.8-5.9 s -> 0.7-0.8 s, the read waited
      3.5-5.6 s -> 0.3-0.5 s;
    - `word_fields` update after one edit: held 8.9-11.2 s -> 1.9-2.3 s, the read waited
      8.6-10.9 s -> 0 s. The update itself takes about as long as before (9-12 s), since
      it now works on the copy as well.
  - Tool definitions unchanged.

## 0.1.38 -- unreleased

- **Incremental, shared layouts; page-scoped facts** (production round 33, R16). A document
  is laid out from its last layout (docx2svg's `lay_out_docx(previous=...)`: from the page
  an edit reaches until the pages line up again, page for page a layout of the new bytes
  alone), for `check`, `layout_facts`, `update_fields`' passes and the Python API alike
  (`docx_agent.layout.lay_out_bytes`, the tool layer's worker task); never a stale one --
  where docx2svg cannot lay out from it (a style, a section, a header changed...) it lays
  the document out whole. The last layout is kept per document (one, as a pickle: about
  6 MB for a 79-page report) and is the same one the `render` tool keeps, so a check after
  a render of the same state lays nothing out, and the other way round. Callers asking for
  one state while it is laid out wait for that one layout (the lock let go where the call
  may let it go) instead of laying it out again. A layout is now made with docx2svg's
  `keep=True`, which keeps its measurements across the passes a body with page-number
  fields needs (9.1 s against `convert_docx`'s 13.0 s on the report). Page SVGs are drawn from
  the layout when asked for (`render_svg`), not all with it. `check(pages=...)`'s
  `direct_formatting` and `column_fit` work only on the blocks and tables drawn on those
  pages (the same entries as before for them). Less held under the document lock: a
  check's `validate` ran twice (once now), its field facts read every part again per field
  (once now); a long table's paths are listed once, not once per paragraph in it
  (`word_edit_table`'s Markdown, `keep_together`'s rows, the layout's ids, the layout
  facts: they grew with the square of the rows). On a made-up 79-page report (25 tables,
  one of 432 rows; tool layer, worker pool): a check after one edit 18-20 s -> 3-4 s; a
  check after a render of the same state 18 s -> 1.3 s; nine concurrent checks of one
  state 52 s -> 3.9 s (each); a batch with a check 19 s -> 3.6 s; `word_fields` update
  past its 30 s deadline -> 4-8 s; the layout facts, whole and for three pages,
  4.7 s / 0.9 s -> 1.6 s / 0.3 s; `word_edit_table` on the long table 3.6 s -> 1.5-2.2 s.
  Every page's SVG, signature, warning, coverage and layout fact the same as before on
  the 151-document corpus, and an incremental layout the same as a whole one there (278
  layouts from an earlier version). Tool definitions unchanged.

## 0.1.37 -- unreleased

- **`read_batch` in the core: many reads in one call, a reader's too** (production round 32,
  R15; ooxml-edit 0.78.0). `describe`, `word_read` (ranges, views), `word_inspect`,
  `check`, `find_text`, the listing tools... in one call, each answered as
  the call alone, within its own `max_result_chars`; the reply within 96,000 characters,
  what does not fit listed with the call that asks for it. It changes nothing and is the
  `review` role's, so a reader (`roles=["review"]`) has it in its core; a change among its
  ops is refused with the way out (`batch`). The prompt's line on `word_read` says several
  reads go in one `read_batch`; ooxml-edit's shared fragment names it too. Offline tokens:
  the core 2,458 -> 2,642 (about 3,910 counted by the proxy), every definition 13,756 ->
  13,939 (guard 13,940), a reviewer's core (review, resolve, accept) 2,333 -> 2,516 (guard
  2,520) and every definition 7,625 -> 7,808 (guard 7,810). Sent non-strict: the strict set
  is unchanged (6 tools, 23 optional parameters). `ooxml-edit>=0.78.0` (was 0.73.0).

## 0.1.36 -- unreleased

- **Numbered headings in a table of contents** (production round 32, R12): `update_fields`
  writes a heading's list number in its entry, as Word's own update does (measured,
  `tools/toc_numbers_probe.py`, `tests/observations/toc-numbers-word.json`, 27 made-up
  documents): the number (`3`, `3.1`, as the list draws it) in its own Hyperlink run, a tab
  run in the paragraph mark's formatting, then the heading's text; a number followed by a
  space -- or by nothing -- is the number, a space and the text in one run. A tab gets a
  left stop in the entry: the number's end (the line's start plus its width in the entry's
  face) plus 240 twips, moved to the left indent of the first `toc 1` ... `toc 9` style, in
  level order, at or past it (240 twips per level for a style the document lacks), else
  that point itself. Every one of the 27 documents' entries as Word wrote them.
- **The right tab stop next to the style's own** (production round 32, R14): an entry gets
  no right dot stop of its own when its `toc N` style has a tab stop (its own or inherited,
  any alignment or leader), as Word writes it -- no second stop 0.5 pt from the style's.
  A style with none (or one that clears its inherited stop) still gets the column's end less
  10 twips.
- **`not_followed` kinds** (production round 32, R13, the names pptx-agent's
  `ppt_follow_sample` uses): every entry `word_follow_sample` leaves has a `kind`: `own`
  (the copy's own text, lines and list membership), `changed_by_copy` (a style this
  document lacks that Word has no built-in for), `structure` (columns, widths of another
  column count, merges, another kind of block). `report: "actionable"` leaves the `own` ones
  out, counted per copy in `own`; over the reply's budget the `own` ones become counts
  before anything else is cut.

## 0.1.35 -- unreleased

- **Copies follow a revised sample** (production round 31, R8): `Document.follow_sample(sample,
  copies, sample_document=None, preview=False)` and `word_follow_sample` (deferred, in
  `word_style`) make each named copy of a sample table -- or paragraph -- take the sample's
  formatting in one undo step, keeping its own text and values
  (`docx_agent.edit.follow`): the table's properties (style, width, alignment, borders,
  shading, layout, margins, style options; not its alt text), the grid's column widths
  where the column counts agree, each row's and cell's properties from the corresponding
  sample row and cell (as `compare_blocks` pairs them: header rows one to one, the last body
  row the pattern for more), each paired paragraph's properties (not its list membership)
  and its runs' formatting property by property -- the copy's runs with its paragraph's
  usual value take the sample's, a word in its own emphasis keeps it; a paragraph with no
  text gets it on its mark. Never changed: text, rows, columns, merges, lists. The report
  says per copy what changed and `not_followed`: each difference `compare_blocks` still
  finds (less the copy's own row count), with `why` and `fix` calls where a tool makes it;
  20 per copy, 100 per reply, the rest counted (`more`, `left_out`). `preview` reports the
  same and changes nothing. A sample in another open document: styles by name, a built-in
  one added, one Word lacks reported. Tracked: `w:tblPrChange`, `w:tblGridChange`,
  `w:tcPrChange`, `w:pPrChange`, `w:rPrChange` as each edit records them; rows untracked,
  as Word (a warning). Addresses as the other tools take them: a cell or row address or a
  paragraph in a table names its table, an empty copy address and a repeat are ignored with
  a note, the sample itself is a no-op; an unknown one is refused naming the tables there
  with their rows and columns. +260 to every definition (13,459 -> 13,719), none to the core;
  not strict.

## 0.1.34 -- unreleased

- **The table style new tables get** (production round 31, R3), as Word's Table Design >
  Set as Default ("this document only") writes it (measured in Word:
  `feedback-round31/measurements-C.md`): `Styles.set_default(name)` and `word_styles`
  `set_default` write the settings' `w:defaultTableStyle` naming the style -- nothing else:
  not the styles' `w:default`, not the tables there are (Word draws a table with no style of
  its own as before, and reads its default table style from that setting, not from a table
  style's `w:default`). A built-in table style the document lacks is written as Word writes
  it; the same style again is no change, said. `insert_table(style="default")` and
  `insert_markdown`'s tables take it first, then the most used one, then Table Grid.
  `Styles.table_default()` reads it; `word_styles` `list` marks it `"default": true` (and
  its summary names it), `describe` too.
- **A style's tab stops** (R4): `Styles.modify(name, tabs=[{"position", "alignment",
  "leader"}])` (`add` too) and `word_styles` `add`/`modify` `tabs` replace a paragraph style's
  own stops (`[]` removes them), written as Word writes a style's tab stop (`w:val`,
  `w:leader`, `w:pos`, in order of position); `{"position": pt, "alignment": "clear"}`
  clears a stop the style inherits through `based_on` (Word's own clear writes `w:tab
  w:val="clear"`, and Word honours one written so: measured), with a warning when it
  inherits none there. A stop with no `position` is at the text column's end, worked out when
  written: the widest section's width between its margins (a warning names the widths where
  sections differ). Measured: Word draws a right tab past the right margin at the
  paragraph's right edge (the margin less a right indent) and one within the margin where it
  is, so a stop far past any margin would follow the margins from the file alone -- but
  docx2svg draws a stop where it is, past the page, so that form waits for docx2svg.
  `word_styles` `describe` lists a style's own stops. With `set_default`, +168 to every
  definition (word_styles is not strict); the caption element +8 (word_insert_markdown is
  strict: an enum value, so the strict set stays 6 tools, 23 optional parameters).
- **Frequent wrong calls** (R5a-d): `word_delete` with `find` and `target` finds within the
  target (a note says so; with a different `within` too, refused). A Markdown paragraph
  starting `: ` -- or, as Pandoc marks a table's caption, `Table: ` before or after a table
  -- is a caption paragraph in the Caption style (written as Word writes it where missing),
  its marker taken off; `style_map`'s new `caption` element names another style (in a quote,
  a list item or a note it is text; reading shows the paragraph's style on its block as
  before). A cell a table lacks (`word_inspect` `t:X/c1,2`, the cell edits) says the table's
  rows and columns, counted from 0, its merged cells and the positions no cell covers, with
  its cells' ids as the options. `keep_together` on a one-row table is no change with a note
  (`cant_split` keeps a row on one page) instead of a refusal that failed a whole batch.
- **`check`'s `column_fit`** (R6, in `layout`): a body table's column whose widest line uses
  at most 60% of its text width (every paragraph on one line), beside a column with a
  paragraph drawn on 3 or more lines, freeing at least 18 pt: both columns' width, widest
  line and longest word (pt), the wrapping one's lines, every column's suggested width (the
  roomy one its widest line or word, its cell margins and 3 pt; the rest to the wrapping
  one) and the `word_format_table` calls that set them. Thresholds decided, not measured.
- **`numbering_skips_level`'s `first` is gone** (its name in 0.1.32): `para` names the
  paragraph.

## 0.1.33 -- unreleased

- **"Numbering: None" stops a list level a style inherits** (production round 30, W13), as
  Word writes it (measured in Word: `feedback-round30/measurements-X.md`): a style based on
  a style linked to a level numbers at that level too (Word numbers it, 1. a. 2. a.), and
  `Styles.modify(name, numbering=None)` (`{"list": "none"}` in `word_styles`) now stops that
  as well: where the style still numbers through `based_on` after its own link goes, it gets
  `w:numPr` with `w:numId` 0 and no `w:ilvl` (what Word writes for a derived style's
  "Numbering: None"; Word keeps it on saving and draws no number, the paragraphs after it
  counting on). A style that numbers nothing, its own or inherited, is left as it was:
  `changed` false and a warning saying so (the tool's summary too). Linking a style to a
  level warns, like the previous holder's warning, naming each style based on it that numbers
  at the level too, with its paragraph count. `numId` 0 was already no numbering in
  docx2svg and in docx-agent's own numbering (`list_of`, `check`, `numbering_skips_level`,
  a style's list); tests hold it for a style's `w:numId` 0.
- **A list level's positions and what follows its number** (W14), as Word's "Define new
  multilevel list" writes them (measured): `set_list_format(..., number_at=, text_at=,
  follow=, tab_at=)` and `word_lists` `format` with `number_at`, `text_at` (points: the
  number's "Aligned at" and "Text indent at", the level's `w:ind` -- `left` the text,
  `hanging` the difference, `firstLine` where the number is right of the text, `firstLine` 0
  where they are one; the one not given keeps its place), `follow` (`tab`, `space`,
  `nothing`: `w:suff`, none for a tab) and `tab_at` (the level's `num` tab stop, "Add tab
  stop at"). A style linked to the level takes this indent. docx2svg draws every label and
  text where Word's PDF has them (ten cases). +90 to every definition (13,192 -> 13,282),
  none to the core; word_lists is not strict (the strict set: 6 tools, 23 optional
  parameters).
- **`word_styles` `describe` says a paragraph style's list level** (W15): `numbering`
  `{"list", "level", "own"}` and, for a level inherited through `based_on`, `from` (the
  style that holds it); `{"list": "none"}` for a style whose own `w:numId` 0 stops
  numbering; absent for a style that numbers nothing.
- **`numbering_skips_level` names its paragraph `para`** (W15), as the other layout facts
  do; `first`, its name in 0.1.32, stays with the same id for readers of that.

## 0.1.32 -- unreleased

- **A paragraph style linked to a list level** (production round 29, W12), as Word's "Link
  level to style" does it (measured in Word: `feedback-round29/measurements-W.md`).
  `Styles.modify(name, numbering={"list": <numId, or a style in that list>, "level": n})`,
  and `word_styles` `modify` with `numbering` (`{"list": "3" | "Chapter title", "level":
  1}`): the level names the style (`w:pStyle`, the definition's and an instance's override
  of it), the style numbers its paragraphs at that level (its `w:numPr`) with the level's
  indent (its own `w:ind` goes), and the style that held the level before is unlinked -- its
  `w:numPr` and `w:ind` go, so its paragraphs are no longer numbered, which a warning says
  with their count; a style linked to another level of the list moves. `numbering=None`
  (`{"list": "none"}` in the tool) unlinks the style: no level names it, and its
  `w:numPr` (and, where a level was linked, its `w:ind`) goes. A list or level the
  document lacks, a style in no list and a style not of paragraphs are refused before
  anything changes. Formatting and numbering in one call are one undo step. docx2svg draws
  the numbers as Word does (1, 1.1, 1.2, 2, 2.1, 2.1.1). +84 to every definition, none to
  the core; word_styles is not strict (the strict set: 6 tools, 23 optional parameters).
- **`check`'s `numbering_skips_level`** (W12, in `layout`): paragraphs numbered at a list
  level whose label shows a level above that no paragraph is numbered at -- a section
  style's `1.1` with no chapter before it (Word counts the unused level from its start:
  1.1, 1.2) -- with the style, the level and the unused levels (from 0), the label, the
  number the first one shows (where every level above is unused), how many and the first
  one's id. Read from the document, so listed where the layout failed too.
- **A row-limit refusal with no smaller range to offer** (W11): `word_edit_table`'s and
  `word_insert_markdown`'s refusal of a source with more rows than a table takes (or has
  room for) passed `valid_options=None` when the source could not take a row range (one
  row by key, a selection) and failed as `internal`; it is the refusal (`limit`, or
  `refused` for an existing table's room) with its advice and no options.
- Tests only (production round 29, C1; ooxml-edit 0.75.0): the shared edit_chart's `blanks`
  and `null` for a blank point in its values add 49 offline tokens to every definition, and
  nothing to the core, which does not load edit_chart. The token guards in `tests/` rise
  with it. edit_chart is never sent strict, so the strict set is unchanged (23 optional
  parameters). `read` now always says `blanks` (how the chart draws blank points), so
  w3-figures-update's two chart reads were re-recorded with `tools/refresh_goldens.py w3
  --new-keys blanks`: the data less `blanks` hashes as before.

## 0.1.31 -- unreleased

- **A render worker's kept layouts are safe from threads** (free threading, step 1). The
  layouts a worker keeps for the next call (`tools/render.py`, `_RECENT`) were an unlocked
  dict, and two tasks handed the same pickle got the same layout and the same
  `InstalledFonts`, which laying out and drawing write to: in-process threads with the GIL
  off (Python 3.15t) could raise from its eviction or write to one set of fonts at once.
  Every use now holds a lock, and a task takes its layout out until it is done, so a second
  task at the same time reads its own from the pickle; a layout handed as `previous` is not
  put back, since the new layout goes on with its fonts and measurements. A warm render (the
  layout held) costs what it did; copying the fonts per task instead would have added 10 to
  90 ms a call. New `tests/test_threads.py`; a thread note in
  [docs/api-tour.md](docs/api-tour.md). New `.github/workflows/freethreading.yml`: the suite
  on Python 3.15t, Linux and macOS, with lxml 7 (a pre-release, in that workflow only), the
  GIL off and `RuntimeWarning` an error, after checking that importing the package and its
  dependencies leaves the GIL off; not a required check yet.

## 0.1.30 -- unreleased

- **Every place-taking parameter reads the same ids** (production round 28, W9). One
  resolver (`docx_agent/edit/addresses.py`) behind every `target`, `range`, `within`,
  `targets`, `blocks`, `at`, `after`/`before` and `section`: a cell id (`t:X/c1,1`), a row
  (`tr:`, `t:X/r1`), a section (`s:2`, `s:body`), a bookmark (`bm:Name` or its plain name), a
  story, a text range and a span are read the same everywhere they make sense
  (ROADMAP.md, "One resolver for every place", has the table). `word_inspect` on a cell
  inspects the cell: its paragraphs, and `cell` with its id, table, row, column and the
  properties it states. `word_read` reads a cell's blocks, a row as its table with that row,
  a bookmark's or range's blocks. Scopes (`find_text`/`replace_text` `range`, every `within`,
  `word_format` and `word_lists` `targets`, `word_changes` `within`) take cells, rows,
  sections and bookmarks; text targets and the `at` of `word_fields`, `word_notes`,
  `word_drawings` and the content-control tool take a paragraph (its text), a cell (its paragraphs'
  text) and a bookmark as well as a range; a `section` takes any id in the body (its
  section). **Open spans:** `p:A..` runs from `p:A` to the end of its container (the body,
  or the cell it is in), `..t:B` from the start; refs at either end (`$a..$b`, `$a..`). A
  malformed span (`p:A...p:B`, `..`, three ends) is refused with the syntax, never as an empty
  id; a cell, row, span, section, bookmark, story or text range where one block is wanted, and a
  paragraph in a table given as a table, is `invalid_arguments` with the ids that would do in
  `valid_options`. Library: `AddressError` (a `KeyError`), `Document.range` takes a
  paragraph, cell, bookmark, hyperlink or run id, `section_of` any id in the body, `get` a
  bookmark's plain name. Calls that worked do what they did (but a reversed span, below); refusals that change: a cell,
  row, span, section, bookmark or text range where a paragraph or table is wanted was
  `not_found` ("no paragraph or table ..."), now `invalid_arguments` naming what it is.
- **`word_set_text` writes each `\n` as a paragraph** (W10). A `\n` in the text, refused
  before, now starts a new paragraph after the target (`Document.set_text_paragraphs`);
  `\v` stays a line break. No parameter: what was refused now works. The first line
  rewrites the target as before (its formatting kept); each new paragraph takes the target's
  paragraph properties -- style, alignment, spacing, indents, keeps, list membership --
  without a section break or a revision record, and its last text run's formatting, as
  Enter at the end of the paragraph and `insert_paragraph` after it do. A cell target whose
  text has a `\n` gets the lines as its paragraphs (its others go, as `word_edit_table`
  `set_cells` writes them); without one, its first paragraph is rewritten as before. With
  `items`, each target is split on its own. A repeat changes nothing; a paragraph that ends
  a section is refused (the new paragraphs would start the next one). Tracked like any
  insertion; one undo step.
- **A reversed span is refused everywhere alike.** `word_read`'s `range` and
  `word_insert_markdown`'s `replace:` read `p:B..p:A` (its last end before its first) the
  other way round; now they refuse it as `word_delete` and `word_move` do (`'p:B..p:A' ends
  before it starts`).
- **Tool definitions.** `word_set_text`'s `text` says a `\n` starts a paragraph (+14) and
  `word_read`'s `range` names a cell and an open span (+4): the core 2,440 -> 2,458
  estimated (about 3,638 counted by the proxy); every definition, with `word_inspect`'s
  `target` (+12), 13,028 -> 13,059. The strict set is unchanged: 6 tools, 23 optional
  parameters.

## 0.1.29 -- unreleased

- **A batch of Word ops costs about the same per op at any document size** (production
  round 28, W8: 31 tables, 124 `word_format_table` / `word_edit_table` ops took 122 s, three
  ops in four not run). Measured on made-up 23- and 24-page reports with 12 and 31 tables
  (`feedback-round28/timings-R.md`): what each op repeated over the whole document is done
  once per batch, or only for what it reads:
  - `word_format_table` `keep_together: true` lays the document out for its warning (a table
    taller than a page) once the document is as the call leaves it -- in a batch after the
    last op, one layout for the batch, where each op laid it out (ooxml-edit 0.73.0's
    `Call.defer`). The warning keeps its place among the batch's warnings. **Behaviour
    change, by decision:** in a batch it describes the document after the batch, not after
    that op -- the two differ only where a later op of the batch changes that table's height
    or place, and the state after the batch is the one the caller goes on with. A call alone
    is as before.
  - `word_edit_table` returns its table as Markdown from the table read alone
    (`Reader.block`): the blocks before it only have their list items counted (a list item
    in a cell numbers on from the body's list as before), the blocks after it are not read.
    `Document.to_markdown(<table or cell id>)` the same.
  - The effective formatting (`paragraph.effective`, the Markdown reader's emphasis, a
    table's row keeps) finds docx2svg's paragraph of the one asked by its path, checked by
    docx2svg's own `resolve_path`, instead of resolving every paragraph's path after every
    edit (work that grew with the square of the document); a paragraph not found so maps
    them all, as before.
  - The id index, the drawings without ids, the ids in use and the section breaks are found
    by lxml's walk or XPath, in C, not by a Python walk over every element.
  Every result, warning and byte is the one before: a golden-style test runs a mixed batch
  and each of its ops alone with all of this off and on, and the walks are compared with
  the old ones on every fixture.
- **A batch past its deadline says what was not run** (ooxml-edit 0.73.0): how many ops,
  the first and last, and a batch size that would have fitted at the rate the ops ran
  (`data.not_run`), in the summary too.
- Requires ooxml-edit 0.73.0.

## 0.1.28 -- unreleased

- **Asking for what is so already succeeds** (production round 27, W7). `word_format_table`
  `keep_together: false` on a table not kept together (one row, too) and `keep_with_next:
  false` where the last row does not keep were refused; now they change nothing, the summary
  is `No change to <table>`, `changed` is empty and a warning says why (`t:X was not kept
  together`, `t:X's last row did not keep with the next`); with another property in the call,
  that is set and the warning stays. A table its rows' style keeps together keeps its round-25
  warning; `keep_with_next: true` without `keep_together` stays refused. The library's
  `set_keep_together(..., False)` and `set_table_keep_with_next(..., False)` return the same
  unchanged `EditResult` with the warning instead of raising `EditError`. On a table not
  kept together whose last row keeps with the next on its own, `keep_with_next: false` (with
  or without `keep_together: false`) releases that row, the warning saying so (`t:X was not
  kept together: its last row's keep with the next is released`).
- **The same for the like refusals elsewhere.** `remove_from_list` on a paragraph in no list
  (`word_lists` `remove`): no change and a warning, where `EditError` stopped the whole call
  at the first such paragraph; the tool now takes the items in a span out of their lists and
  names the others in one warning. `inline_drawing` on a drawing in the line and
  `float_drawing` on a floating one (`word_drawings` `inline`, `float`): no change and a
  warning; floating one with settings sets those (as `set_drawing`), where it stands.
  `word_headers_footers` `remove` or `link` where the section has no such header or footer
  of its own, and `unlink` where it has one, changed nothing already but said `Removed...`;
  now the summary and a warning say no change. `add_bookmark` (`word_links`) under a name
  already marking exactly that text: no change and a warning (`bm:X marks this text
  already`); the name elsewhere stays refused. Refusals that guard against a mistake stay
  (a wrong target kind, a missing name, a one-row table kept together, a level or restart for
  a paragraph in no list, an inline shape). Tool definitions unchanged.

## 0.1.27 -- unreleased

- **`word_compare_blocks` pairs a cell's paragraphs by an alignment** (production round 26,
  W6). A cell's paragraph *k* was compared with the sample cell's paragraph *k* (or its
  last), so a copy whose cell lacked the sample's label line had its body line restyled as
  the label. Now the cell's paragraphs and the sample cell's are aligned in order by style
  and text (case and white space ignored): pairs whose texts are the same or 0.8 alike
  (`TEXT_SIMILARITY`, difflib's ratio) first, then pairs of the same style; two paragraphs
  matching neither way never pair. A sample paragraph left unpaired is a
  `missing_paragraph` (the cell; `reference` its `paragraph_style`, the start of its `text`,
  its `index`), a copy paragraph left unpaired an `extra_paragraph` (the paragraph; `value`
  its style and text), and where the counts differ the cell's `paragraphs` differ (`value`
  the copy's count, `reference` the sample's); none has a fix. A copy paragraph after a
  paired one, in that pair's sample style, repeats it (a sample line is the pattern for like
  lines after it, as the last body row is for more rows). Equal counts pair by position as
  before unless the alignment crosses a pair matching by text (reordered lines: missing and
  extra, no style fix). So a style fix rests on matching text, or on position where the
  counts are equal. Tool definitions unchanged.

## 0.1.26 -- unreleased

- **A table kept together and with the next block** (production round 25, W5).
  `word_format_table` `keep_with_next: true` with `keep_together: true` (scope `table`) keeps
  the last row with the next as well -- `w:keepNext` on its deciding paragraph, directly only
  where its styles do not keep it -- so a table and the box after it go to the next page
  together. On a table already kept together, `keep_with_next` alone keeps or releases the
  last row; on a table not kept together it is refused (a last row that keeps alone takes
  only itself to the next page with the block after it, the rest left behind: measured),
  as is `keep_together: false` with `keep_with_next: true`. `keep_together: false` removes
  the last row's keep with the others. A paragraph between that does not keep with the
  next lets the second table start a page alone (measured): the call warns and names it.
  `table_properties` gains `keep_with_next`; `describe` says `keep_with_next: true` beside
  `keep_together: true`. A table kept together is now one whose every row but the last
  keeps, the last too or not -- so a label style that keeps every row is reported as kept
  together `with_next` (it is, in Word), and `keep_together: false` on it warns that the
  style still keeps the rows.
- **check's `layout.settings`** (W5): the tables kept together moved from
  `layout.keep_together` to `layout.settings.keep_together` -- what the document is set to
  do, not facts to act on; a summary counts none of it. Entries gain `with_next: true`. A
  table kept together taller than a page keeps `taller_than_page` there and is also listed
  once in `keep_with_next_at_foot` (its first split, `taller_than_page: true`), so it stays
  among the facts to look at. `keep_with_next_at_foot` also reports a table whose last row
  keeps with the next, drawn last on its page with the next block on a later page (a chain
  of keeps longer than a page, broken there: `{"table", "row", "page", "next", "next_page"}`).
  Measured in Word for Mac 16.106 (`feedback-round25/probe-M/w5.py`, 37 cases): docx2svg
  places every one as Word does.

## 0.1.25 -- unreleased

- **Keep lines together** (production round 25, W4). `word_format` `keep_lines` (Word's
  "Keep lines together", `w:keepLines`; the Python API's `keep_together`, which the tool
  also reads as `keep_lines`, saying so) sets it on or off, and `clear: ["keep_lines"]`
  takes it back to the style's. check's `direct_formatting` reports a direct `w:keepLines`
  that differs from the styles' as `keep_lines`, as it does `keep_with_next`, and its fix
  clears it; `word_styles` describe and `word_compare_blocks` name it too. The tools say
  `keep_lines` for a paragraph everywhere, since a table's `keep_together` (`word_format_table`,
  check's fact) is another thing: `word_inspect`'s `effective` and `word_design_facts`' looks
  said `keep_together` and now say `keep_lines` (mapped at the tool layer; the Python API,
  `Document.state()`, `design_facts()` and `Paragraph.keep_together`, keeps its names). docx2svg already paginates by it (a paragraph so
  kept that does not fit at a page's foot moves whole); a made-up probe in Word for Mac
  16.106 drew all 7 cases as docx2svg does (`feedback-round25/probe-L/keep.py`): at a
  page's foot a paragraph splits, kept it moves whole, a style's keep released directly
  splits, a keep-with-next paragraph before it moves with it, and one taller than a page
  starts the next page and breaks where that is full.

## 0.1.24 -- unreleased

- **check's kept rows by Word's rule** (production round 25, W1). `keep_with_next_at_foot`'s
  table entries asked every paragraph of a row to keep with the next; Word keeps a row by
  the first paragraph of its first cell (a nested table first in that cell: its first
  paragraph), measured in round 24. The fact now asks docx2svg's rule
  (`docx2svg.table.row_keeps_with_next`), the one its layout paginates by, so a table whose
  first column alone keeps (a label style) and runs longer than a page is reported.
- **A table kept together** (W2). `word_format_table` `keep_together: true` (scope `table`)
  keeps every row but the last with the next -- `w:keepNext` on each row's deciding
  paragraph, only where its styles do not keep it (a direct release of a style's keep is
  removed) -- and releases the last row (`w:keepNext w:val="0"` only where its style keeps
  it), so the table goes to the next page whole and does not take the block after it along.
  `false` removes what it writes and only that, recognised by the rows' keeps -- on a table so
  kept, each deciding paragraph's direct keep that differs from its styles'. Measured in
  Word (11 cases, all as docx2svg draws them): a table that does not fit moves whole; a
  table whose label rows chained it to a heading and the next table now stays and the next
  one moves; a table longer than a page starts the next page and breaks where it is full
  (with or without a header row). check reports such a table as `keep_together`
  (`{"table", "pages", "direct"}`), the keeps and the release are not `direct_formatting`
  and no fix clears them; `table_properties()["keep_together"]`, and `describe`'s tables
  say `keep_together: true`. `Document.set_keep_together(table, on)`. One row: refused.
  A table kept together that is taller than a page is still kept so (Word starts it on a
  new page and breaks it where that is full): the result warns "taller than a page", decided
  from the layout after the edit (a kept-together table still drawn on more than one page;
  where the document cannot be laid out to it, the warning says that is not known), and
  check marks its `keep_together` entry `taller_than_page: true` instead of listing a
  `keep_with_next_at_foot` split after each kept row of it.
- **`past_margins` to the last character** (W3). A table's lines are measured from their
  first to their last character that is not a space: a table as wide as the column with no
  cell margins reported its lines' last spaces (0.8-1 pt past the edge). Measured in Word
  (`feedback-round25/measurements-J.md`, 20 characters at 9 widths each): only spaces hang
  past a cell's edge; punctuation, quotes, dashes and brackets need their whole advance
  inside, as docx2svg breaks them (180 of 180 as Word).
- **The render workers share one font scan** (R1). The first worker to start scans the font
  folders (`docx2svg.scan_fonts`, each file opened once) and saves the scan in a folder of
  the user's own in the temporary folder; the workers starting with it wait for it (a file
  lock) and every later worker takes it to `docx2svg.warm(scan=...)`, which reads only the
  files changed or added since (and saves the scan again), so a changed font folder is
  still seen. Renders are the same bytes with and without it. Needs docx2svg 0.11.1.

## 0.1.23 -- unreleased

- **A table without a header row** (production round 24, W2). Measured in Word: a table
  style's first-row formatting follows `w:tblLook`'s `w:firstRow` (the attribute where it is
  written, else its bit in `w:val`, and on for a table with no `w:tblLook`), never a row's
  `w:tblHeader`, which only repeats the row; Word's "Header Row" off writes `w:val` 0480 with
  `w:firstRow="0"`. Markdown whose header row is empty (`| | |`) is a table with no header
  row: no `w:tblHeader`, the first-row look off, every row a body row; the reader writes such
  a table back with an empty header row, so it round-trips. `word_format_table` takes `look`,
  the style options on by Table Design's names (`header_row`, `first_column`,
  `banded_rows`, `total_row`, `last_column`, `banded_columns`; the rest off), and
  `cant_split` on a row (`scope: "row"`) or every row (`scope: "table"`).
  `table_properties()["look"]` of a table with no `w:tblLook` is the default look (header
  row, first column, banded rows), as Word and docx2svg draw it, no longer all off.
- **Tables back to back stay two tables.** Markdown with two tables in a row wrote them
  adjacent, and Word joins adjacent tables into one (measured: the second's rows drew as rows
  of the first, its own look gone). An empty paragraph in the body's paragraph style (the
  mapped `paragraph` style, else the default) now goes between them, as `insert_table` keeps
  one -- also where an inserted table would touch a table already there. Measured: Word
  keeps the tables apart, each with its own look; docx2svg draws them alike. The separator
  reads back as nothing (an `empty` id comment with ids), so the Markdown round-trips to one
  separator, not two.
- **check's `keep_with_next_at_foot`** (W3): a keep-with-next paragraph whose next block
  starts a new page anyway -- a page break before it (its style's or its own), a page break
  at its start or at the end of the paragraph, a section break that starts a new page -- is
  no longer reported; a table split after a row whose paragraphs all keep with the next is
  (`{"table", "row", "page", "next_row", "next_page"}`): Word keeps such rows with the next.
  Read from docx2svg's layout, as the rest.
- **check's `header_not_repeated`** (W4): `fix` (`header_rows: 1`) only where the first row
  looks like a header -- the table style's first-row formatting applies to it, or all its
  text is bold, or each of its cells has a fill, unlike the second row -- named in
  `header_look`. A two-column table whose first row looks like the rest (labels and values)
  is not reported; another such table gets the call as `fix_if_header`.
- **Shading on styles** (W6): `styles.add(..., shading=...)` and `styles.modify`, and
  `word_styles` `add`/`modify` `shading`: `RRGGBB`, a theme colour, `"accent1 lighter
  80%"` / `"accent2 darker 25%"` (or `{"theme_color", "tint" | "shade"}`), `none` to remove;
  a paragraph style's behind its paragraphs (`w:pPr`), a character style's behind its text
  (`w:rPr`). Written as Word writes it (measured: Word's own theme shading on a style, and
  what Word saves): `w:shd w:val="clear" w:color="auto"`, `w:fill` the colour drawn and
  `w:themeFill` with `w:themeFillTint` or `w:themeFillShade`. `describe` reports it.
- **A restarted style-numbered item keeps its style's indent** (W7). Direct numbering takes
  the list level's indent, so where the style's differs, the restart's instance carries the
  whole level with the style's indent and the restart number as its start beside the
  `w:startOverride` (measured: Word draws the item at the style's indent, keeps the form on
  saving, and the items after it count on from it; with a level in the override Word counts
  from the level's start, and a level without the `w:startOverride` does not restart).
  `word_lists` `restart` and Markdown lists that restart (a list starting at 4, a later list
  of one insertion) alike; the warning about the level's indent is gone with the cause.
  docx2svg draws every form as Word does.
- Tool definitions, offline: every definition 12,824 -> 12,944 (guard 12,945), the core
  unchanged (2,440); a reviewer's all 7,281 -> 7,310 (guard 7,315). The strict set is
  unchanged: 6 tools, 23 optional parameters.

## 0.1.22 -- unreleased

- **A Word toolbox is sent the shared tools with a document's parameters only**
  (production round 23, V3; ooxml-edit 0.70.0, `Tool.kind_fields` and `kind_values`):
  `check` without a deck's `slides`, `boxes`, `spacing` and `contrast`, its `include`
  offering `validate`, `reflow`, `fields`, `layout` and `app` (not a deck's `fit`,
  `collisions`, `facts`, `design`, which this handler refuses), its default described for a
  document; `find_text`, `replace_text` and `render` without `slides`; `stories` offering
  `body` and `all` (a deck's `layouts` and `masters` were read as the body); `undo` without
  a deck's or workbook's `scope` (a document undoes document-wide, and still refuses it).
  `contrast`
  could apply to Word text (its colour against shading or a page colour), but nothing here
  measures it and the handler refuses it, so it is left out until it is measured. A call
  that still passes one is answered as before: `boxes`, `spacing` and `contrast` refused
  with `<name> is for decks; for a document use include`, `slides` refused by find and
  replace and read as `pages` (with a warning) by render and check. A toolbox with
  pptx-agent's tools too still sends all of them. Offline, from 0.1.21: the core 2,820 ->
  2,440, every definition 13,204 -> 12,824 (guard 12,825); a reviewer's core 2,612 -> 2,328
  (guard 2,330), all 7,564 -> 7,281 (guard 7,285). `replace_text` is strict and loses `slides`:
  6 tools, 24 -> 23 optional parameters.
- **`block_of(session, doc, id)` and `blocks_between(session, doc, a, b)`** (production
  round 23, V5): `docx_agent.tools.block_of` and `blocks_between` take the application's
  session and a document id outside a call -- each `$name` resolved by ooxml-edit 0.70.0's
  `Session.resolve_ref`, the document read under its lock -- beside a handler's
  `(call, ...)` forms, one implementation for both. An unknown ref or document is
  `not_found`; a document of another kind `invalid_arguments`. `session.refs(doc)` lists a
  document's refs. Requires ooxml-edit 0.70.0.

## 0.1.21 -- unreleased

- **Markdown lists in a style that numbers itself take its numbering** (production round 23,
  V1). Where the style a list maps to carries `w:numPr` -- its own or through `basedOn` -- at
  the list's level, `insert_markdown` writes no `w:numPr` and no indent: the style's
  numbering and indent apply, and the items count on from the style's earlier items, across
  calls and the paragraphs between, as Word numbers them (measured: 1, 2, a paragraph in
  another style, 3). A template's List Number 2 tied to level 1 of List Number's definition
  nests so; Word's own List Number 2, a list of its own at level 0, keeps direct numbering
  at level 1, as does a style with `numId` 0 or one whose numbering level is tied
  (`w:pStyle`) to another style (Word draws no number for it: measured). A list starting at
  another number, or a later list of the same insertion, restarts on its first item alone,
  Word's form. `word_lists` `restart` on an item its style numbers gives that item alone an
  instance over the same definition with a `w:startOverride` (what Word writes; the items
  after it count on from it), with a warning when the level's indent differs from the
  style's; `continue` on items with direct numbering that their style numbers too drops
  their `w:numPr`, or says exactly that when there is nothing earlier to continue. The
  Markdown reader reads Word's restart form as one list. `numbering.style_list(document,
  style_id)` is the rule.
- **A block-level content control is a block** (V2): `delete_block`, `move_block`,
  `move_blocks`, `copy_blocks` (and `at` positions, `word_delete`, `word_move`,
  `word_copy_from`) take its `cc:` id and act on the control with its content, as on a
  table. A control locked against deletion or editing (`w:lock`) is refused with the lock
  named, as is one holding a section break; tracked, a move is untracked and a deletion
  deletes the content as tracked text (Word has no revision for a control), with a warning.
  `blocks()` marks such a control `movable`; a copy's `blocks` and `copied` name the new
  control. An inline control is refused, with a pointer to `remove_control`.
- **Sections by number** (V4): `s:<n>` (from 1) names a section wherever a section id is
  taken -- `word_read` and `word_inspect` (the same resolver: the section's blocks),
  `word_sections`, `word_headers_footers`, `block_of`; `s:body` reads the body of a document
  with no body-level `w:sectPr` too. A number past the last says how many there are.
  `word_read`'s `range` says so: core +4 tokens (2,816 -> 2,820), every definition +10 with
  `word_delete`'s `cc:` (13,194 -> 13,204, guard 13,205); the strict set is unchanged.
- **`Style.outline_level`** (V6): a paragraph style's outline level, 1-9, by the rule
  `blocks()` uses for a paragraph that states none (built-in heading N is N; else
  `w:outlineLvl` through `basedOn`), `None` for body text and other kinds; one function,
  `styles.style_outline_level`, now serves both.

## 0.1.20 -- unreleased

- **Markdown column alignment through styles** (production round 22, Q3): `insert_markdown`'s
  `style_map` (and `word_insert_markdown`'s, and a table from `from`) takes
  `table_cell_center` and `table_cell_right`. A cell of a column aligned `:-:` or `-:`
  takes the style named for it and no direct `w:jc`, so `check` reports no direct
  formatting there; with no style named, the direct `w:jc` stays as before and the result's
  warning names the role that would avoid it. Left and unaligned columns are unchanged; no
  default names either role. Every tool definition +11 tokens (13,183 -> 13,194, guard
  13,195), the core unchanged.
- **`Document.block_of(id)` and `Document.blocks_between(a, b)`** (production round 22, U1):
  the top-level body block holding any id the library takes -- a nested paragraph (cell,
  nested table, text box, content control), range, run, table, row, cell, drawing (a
  group's member too), hyperlink, field, revision, content control, a note (the block of
  its reference), a comment (the block its range starts in), a bookmark (`bm:<name>` or the
  name: its start), a section (its break's paragraph) -- or `None` outside the body (a
  header's paragraph, a note's own text, `s:body`); an unknown id is a `KeyError`. And the
  inclusive span of top-level blocks between two such ids, through `_blocks_of`, so a
  reversed span is the block tools' `EditError`. Refs (`$name`) live in the tools' session:
  `docx_agent.tools.block_of(call, id)` and `blocks_between(call, a, b)` resolve them first.

## 0.1.19 -- unreleased

- **A table of contents' `\t` names built-in styles in the document's language too**
  (production round 21, D1). Word compares `\t`'s names with its own display names --
  Dutch Word lists a `\t "Kop 2;1"` heading and not a `\t "Heading 2;1"` one -- while the
  styles part stores built-in styles under their English names. A name is now taken as a
  stored name or alias, then as the built-in style the document's language shows under it,
  then as English's (`ooxml_common.stylenames.find_style`, Word's display names measured
  in English, Dutch, French, German, Spanish and Italian). The document's language is
  `Document.locale`'s, else the settings' `w:themeFontLang`, else `w:docDefaults`' `w:lang`.
  Without a locale the list is split on the separator Word uses in that language (`;` for
  the five, `,` in English and in a language not measured). Lane B's three cases that were
  expected failures pass: "Kop 2" as Dutch Word, and "Heading 2" in a Dutch document as
  English Word lists it (Dutch Word lists nothing there: the Word that wrote a field cannot
  be known, so both are taken).
- **`STYLEREF` names read as `\t`'s** (with docx2svg 0.10): `update_fields` looks a
  `STYLEREF`'s style up in the same language -- the document's locale, else its
  `w:themeFontLang` / default `w:lang` -- and in English, untrimmed as Word reads it, so
  `STYLEREF "Kop 1"` in a Dutch document stores the heading's text where the English Word
  measured stored its error (`tests/test_styleref_update.py`, `DOCUMENT_LANGUAGE`).
- **Theme colours with a tint or shade are Word's to the level** (production round 21, D2):
  `word_design_facts`' theme colours come from `ooxml_common.themetint.word_theme_tint`,
  Windows' integer HLS on a scale of 1530, which draws all 27 swatches of T7 (and 559 more)
  as Word does; three were a level off. Needs ooxml-common 0.34.0.

## 0.1.18 -- unreleased

- **`update_fields` writes `STYLEREF` results** (production round 21, T2): the text docx2svg
  draws for the field (its 0.9.0) -- in the body, the last text of the style before the
  field, else the first after it; in a header or footer, the text on **the first page that
  shows it**. Measured (`tools/styleref_probe.py`, `tests/observations/styleref-word.json`,
  nine made-up documents updated and saved by Word): Word stores the body's values so, and
  a header's from the pages its window drew (page 2's for a header shown on pages 1 and 2;
  for a header not drawn yet, neither the first page's nor page 2's) -- not knowable from the
  file, so the first page that shows it is written. The paragraph number switches (`\n`,
  `\r`, `\w`) and what docx2svg does not compute are left as they were and listed in
  `unknown`. The live view (`render`, `layout`) shows each page's text through docx2svg.

## 0.1.17

- **`replace_text` and `find_text` take a span** (production round 21, T3): `range`
  `p:A..t:B` is the blocks from A to B inclusive, their tables' cells too, resolved by the
  block tools' own resolver, with their errors for a reversed (`refused`) or unknown
  (`not_found`) span; `expect` counts over the whole span. `word_insert_text`'s and
  `word_delete`'s `within` take one the same way. Definitions unchanged.
- **Cells take their style as it is, and a style that does not fit a cell is named**
  (round 21, T4): `word_insert_markdown` writes no direct indent on a cell paragraph (it
  wrote none on main either: the report did not reproduce; the only direct paragraph
  property it writes there is a Markdown column alignment's `w:jc`). Where the `table_cell`
  style's effective indents run the text past the cell's edge (an indent more negative
  than the cell margin: Word draws the text past it, into the next cell, measured; docx2svg
  the same) or leave it no width, the result warns once per table and style, naming the
  style, the indent and the style it comes from. `check`'s `direct_formatting` leaves out
  the footnote and endnote separators' paragraphs (Word's own spacing; not content), so
  Markdown with every style role mapped checks clean.
- **A table as wide as the body's lines** (round 21, T5): `word_format_table` `width_of`
  `"body"` takes `width_percent` (default 100) of the lines of the paragraph around the
  table (the text width less its effective left and right indents,
  `Document.body_lines(table)`), stored as a width in points with the table's indent on the
  lines' left edge (centre and right placed by the indent), the columns' and cells' widths
  shared to it. Word draws its borders on the lines' edges, as it draws a text-width table's
  on the margins, and keeps it when it saves (measured; docx2svg draws it identically).
  `check`'s `past_margins` leaves out a table's side that stays within the body's lines
  (a negative indent's reach), and reports one reaching further as before. Tool
  definitions: +31 estimated tokens over every definition (13,152 -> 13,183, after T7's), none in the
  core; word_format_table is not strict.
## 0.1.16 -- unreleased

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
