# Changelog

docx-agent has not been released to PyPI. Each version that reaches `main` is tagged
`v<version>`, from 0.1.0, the first tagged. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail of every phase.

## 0.1.81 -- unreleased

- **A text field with Markdown is written as its Markdown** (production round 48, R86: a
  workbook's notes, written from Markdown with nested bullets, came in as Normal paragraphs).
  A dataset's text value whose source holds more than its characters carries its Markdown
  (ooxml-edit 0.106.0's `Dataset.markdown`; xlsx-agent 0.1.34 fills it from a cell's bold
  and italic runs and its list lines). `word_insert_markdown` writes that Markdown for the
  one field `from` names, and for a reference standing alone between blank lines in
  `markdown`: list items in the template's list styles at their levels, emphasis in its
  character styles, as the same Markdown typed in the call (`field_text.with_markdown`). A
  reference inside a sentence writes the value's text on one line, every character as text
  (a list's lines joined by spaces). A value without Markdown is written as before, and a
  table's cells from rows take each value's text. The facts name the blocks each made and
  count the Markdown's characters. With an ooxml-edit before 0.106.0 nothing changes.
  tests/test_dataset_markdown.py; no tool definition changed (0 tokens).

## 0.1.79 -- unreleased

- **The application's own edits under `session.lock` are committed** (round 48, R84;
  ooxml-edit 0.104.0). An application that edits a document through `Document` while holding
  `session.lock` no longer sends the next edit on a draft through three conflicts ("the
  document changed outside the tools") to the whole-document lock: letting the lock go
  commits the change (by the parts it changed), and `session.commit(doc)` does it without
  letting go. The serial-replay stress gains such edits (`tests/host_stress.py`: a paragraph
  added, a block's text set, both in one `batch()`, or the lock held with nothing changed);
  `python tests/host_stress.py --seeds 0-19 --threads 8 --calls 16` replays every run the same.

## 0.1.78 -- unreleased

- Tests: the check-timing test that tells the wait for the layout's worker from the layout
  stretches both by sleeps far apart (0.3 s before the task, 2 s in the layout), so a
  loaded runner's slack no longer fails it; it proves the same thing (round 47, R82). Its
  lower bounds allow 10 ms for a clock that reads a sleep short (Windows: 0.2 s as 0.199).

## 0.1.77 -- unreleased

- **An edit conflicts only with changes to what it depends on** (round 47, R78:
  callers writing into neighbouring, all but empty parts of one document ended on the lock).
  A draft's windows are what its change depends on, no longer two blocks of context either
  side: blocks changed under their own ids are their own window; blocks put between two are
  placed by the block the call named (`before:X` depends on X, `after:P` on P) -- unless they
  hold a table, which reads both neighbours -- and blocks removed or replaced keep their
  neighbours; a block named, alone. The provenance item (values' sources) is merged entry by
  entry with the document's, where two calls' entries differ, instead of a conflict; an id
  the call made itself is not looked for in the document -- which, alone, made every insert
  or fill writing text by reference (`from: d:notes/#1/body`) conflict three times with
  itself ("p:... is gone") and run under the lock, one caller or many. And a call that adds
  the provenance item while its ids are drawn again names the item again from its renamed
  first key (its properties' `itemID` is a hash of it): 0.1.74 kept the hash of the draft's
  key, so the bytes differed from the calls run one at a time (found by the stress). A call
  that drew ids on its draft and then failed (an atomic batch whose last op fails) runs
  under the lock, where its ids are spent as before (`Blocks.drew`; ooxml-edit 0.103.0);
  the stress plans such late failures. On the round's stress shape (ten
  callers, one all but empty part each, tables from a dataset, checks between), 0.1.74
  ended 24-28 of about 90 drafts on the lock after 100+ conflicts; now none, 0-3 conflicts.
  `tests/block_stress.py --shape short`; the serial replay holds for both shapes.
- **"changed while this call read it" only when the change touched what was read** (round
  47): a ranged `check` or `render` says the blocks it read (`docx_agent.scope.ReadScope`,
  ooxml-edit's `Call.read_scope`), and `docx_agent.blocks.touches` says whether a change
  made meanwhile touched them -- changed or removed one, put blocks between two, or changed
  a part a read of blocks shows (not the registrations, the document's properties, the
  provenance item, or lists only added to the numbering); and since both report pages, any
  change at or before the range's last block, or to the last section, too -- a change after
  the range never. A read of the whole document still
  hears of any change.

## 0.1.76 -- unreleased

- **Refs at the committed version, and refs by caller** (round 47, R79 R80; ooxml-edit
  0.102.0). `Session.read_refs` and `Session.read_resolve_ref` read a document's refs at its
  committed version, never waiting for an edit in flight; a call that only names something is
  committed at once. A ref belongs to the agent (call context) that defined it: another agent
  giving the same `ref` to what it makes is refused before anything is made, with a free name
  to use instead (`table_2`); using another agent's `$name` works, and a ref defined with no
  context is shared, as before. The serial-replay stress gains refs by several callers
  (`tests/ref_stress.py`: own names, names every thread gives, names nobody gave yet, as
  `ref`, `at` and `target`); `python tests/ref_stress.py --seeds 0-19 --threads 8 --calls 16`
  replays every run the same (seeds 0-19, Python 3.12 and 3.15t).

## 0.1.75 -- unreleased

- **`check` says where its time went** (production round 47, R82), as an edit does:
  `data.timing`, seconds per part that ran -- `resolve` (the ranges), `validate`, `copies`,
  `layout` (laying the version out, in the worker holding its layout), `facts` (the layout's
  facts read there and here, the comparison with the last check), `fields` -- and the waits
  apart, named: `worker_wait` (the worker's queue, a layout of the document in progress
  there, the bytes there and the facts back) and `document_wait` (taking the document back
  after laying out with it let go). The wait before the call ran is the result's duration
  less these. A check in a batch, which holds the document, times its `layout` where it
  is made. No new parameters; everything else in the result is as before (the same calls'
  results byte for byte, the timing taken out), and the timing costs no measurable time.
- **`markdown_from` is no longer read as `from`** (R81): `word_insert_markdown`'s `from` is
  the one name, as its definition says. A call sending `markdown_from` is refused as an
  unknown argument, with the tool's arguments, naming `from` (the nearest spelling,
  `markdown`, would write the reference as text).

## 0.1.74 -- unreleased

- **Edits of different blocks of one document at the same time** (production round 45, R64
  stage 3; ooxml-edit 0.100.0, `DocumentFormat.blocks`; `docx_agent.blocks`). With another
  edit of the document in flight (and a fill from a dataset always: it may run long),
  `word_set_text`, `word_insert_markdown` (`end`, `before:`,
  `after:`; no pictures or footnotes), `word_edit_table` `set_cells` and batches of them --
  text by reference only from datasets -- run on a *draft*: an editable copy of the last
  committed version (`layout.read_copy`'s `make.draft`), no lock of the document held. The
  draft records what the call does: the blocks its ids named, every id it drew, the facts of
  the whole document it used (the table style the tables use most, the list definition the
  lists use most, a body of one empty paragraph), and anything a patch cannot carry (ids
  stamped or renamed, tracked changes, annotation ids, revision ids repeated, ids kept from a
  copy, drawing numbers, media, a layout), which runs the call under the lock instead. Its
  patch: the body's blocks before and after, by their bytes; each run that differs, and each
  block it named, with two blocks either side, is a window (more after a paragraph written
  in several lines). Under the lock, briefly: every other part unchanged since its version;
  the body's namespaces and last section as they were; each window found by its first id and
  byte for byte as it was; the facts the same; then the windows replaced, other parts
  replaced, ids drawn again where another call drew them first (the result renamed with
  them, never text), the ledger read at the new version, the reflow measured from the
  document as it was just before. Another call's change first: prepared again, at most
  three times (once only after a preparation of more than 2 s of CPU), then under the lock;
  a call that takes the lock (a fields update, a style change) waits for a draft in flight
  that long to apply first, so the team's slow fill runs once. Ids another call drew first
  are drawn again and renamed in every XML part the call changed or added -- the provenance
  item a document's first fill adds too -- so a first fill beside other edits runs once. Serializable: a randomized run of eight callers (their
  chapters and each other's, batches with reads after their change, fills, values by
  reference, formatting and undo under the lock) gives the same results, bytes and id state
  as its commit order run alone (tests/block_stress.py; tests/test_block_edits_serial.py
  small, `--run-slow` longer). Any other call with no other edit in flight runs as before
  (ooxml-edit's `block_edits="auto"`). tests/test_block_edits.py: what a draft takes, a draft's result and
  bytes those of the document's, edits beside a held call, ids drawn again, the same and
  neighbouring blocks prepared again, past the tries, and each rule of the shared-state
  inventory in `docx_agent/blocks.py`. A body is read for its blocks once per state, and
  from the state read last only between the first and last byte that differ (as `touched`
  reads a change).
- A table cell's `from: ""` (a strict decoder's placeholder beside its text) and
  `word_edit_table`'s own `from: ""` are no source (round 45, lane D's find; ooxml-edit
  0.100.0 no longer refuses it either).

## 0.1.73 -- unreleased

- **Checks, renders and reads never wait for an edit in flight; an edit's checks run after
  it lets the document go** (production round 45, R64 stages 1 and 2; ooxml-edit 0.97.0).
  The Word format copies documents for ooxml-edit's reads of the committed version
  (`DocumentFormat.read_copy`: `docx_agent.layout.read_copy`): under the lock, when a
  changing call ends, it copies the id ledger, the renames, the version numbers and the
  settings (font folders, locale, stamping, how a gone id is described); the copy is opened
  later from the version's bytes, with no lock held. It answers every reading call as the
  document did at that version -- the same ids, refusals, facts and pages -- and its
  layouts are the document's own (the layouts, their summaries, the facts of checks, the
  layout the last check saw, all keyed by the bytes laid out), so a check of the copy lays
  out where the document's layouts are held, and the next version is laid out from it. Its
  bytes are those it was opened from (no second serialization). A changing call's checks --
  the validation delta (seconds on a long document) and the reflow -- run on the copy of
  the version it made, with the call's own layout engine, after the lock is let go; a
  change's reflow is measured from the bytes its version kept rather than serialized
  again. On the made-up 60-page report with twelve builders and two layout workers (round
  45 harness), the checks' wait for the lock fell from 5,138 s in all to 156 s. Results and
  bytes are 0.1.70's for the same calls one after another (tests/test_committed_reads.py
  compares every read and edit with the copies on and off; the harness's ordered run gives
  the same results and final bytes as 0.1.70).

## 0.1.72 -- unreleased

- **A table filled from a dataset in one pass** (production round 45, R67). `word_edit_table`
  `set_cells` `from` wrote its block as one edit per row added and one or more per cell, each
  throwing every index away, so the next read the whole document again: 431 rows in a long
  report took minutes (821 CPU-s on a made-up 212-page report; now 3.4). The rows added and
  the cells written are the same edits in the same order, made in one pass
  (`Document._rows_after_last`, `Document._fill_cells`): the ids in use read once, the
  table's grid once, the index rebuilt only where an edit changes the blocks. The bytes,
  ids, results and undo are 0.1.71's; where something still needs stamping the edits are
  made one by one as before. `set_cells` with `cells` that all name their cell by position
  takes the same pass (a thousand cells: about 80 s before); a label is still read cell by
  cell. `Cell.paragraphs` looks its own paragraphs up rather than testing every paragraph of
  the story, and a result's ids are gathered in linear time.
- **`word_fields` update reads less** (R67): a field's result is read from its start onward,
  not from the part's root; where a block is drawn comes from an index of the layout's lines
  (made again when the aliases change), not a search of every line per heading; the
  numbering part is listed once for a table of contents' labels; the toc styles' looks are
  kept by the content of the styles, numbering, settings and theme parts they are resolved
  from, not parsed from the whole package each pass; a layout's paths resolve each parent's
  children once. The same bytes and results (6.95 -> 3.2 CPU-s for an update with nothing
  to lay out on a 62-page report). An update after a large change is still mostly docx2svg's
  layout.
- **Where a long call's time went**: `set_cells` `from` and `word_fields` update say it in
  `data.timing`, seconds per part (`resolve`, `check`, `rows`, `cells`, `record`, `result`,
  `spacing`, `validate`, `reflow`; `ahead`, `update`), the wait for the document's lock in
  none. Tool definitions unchanged.
- The spacing of written text is not looked for a second time at a batch's end when nothing
  changed since the tool layer applied it.

## 0.1.71 -- unreleased

- **A dataset's text field written as Markdown, by reference** (production round 45, R65).
  `word_insert_markdown`'s `from` names one field of one row too -- `d:notes/#12/body`, a
  row by key `d:notes/Doors/body`, another document's `t:checks/Alpha/note` -- and writes
  its text as `markdown` would, verbatim: its paragraphs, its lists as lists in the
  template's styles with nested items at their level, its emphasis (the same document as
  typing it). `markdown_from` is read as `from`. Several fields go in one call as
  references each standing alone between blank lines in `markdown`, with headings or text
  around them. The result's `data.facts` names the blocks each field made (`blocks`) and
  gives a long or several-line value's length (`chars`) instead of echoing it; a figure
  inside a sentence is listed as before. Each paragraph a field made keeps it as its
  source in the file (`Document.text_facts`), so a check of numbers reads its numbers as
  the data's. A value that is not text is refused with the row's text fields and the
  `{{...}}` form for a figure; an unknown dataset, row or field names what there is, as
  whole addresses. `word_set_text` writes a field the same way (`{{d:notes/#2/summary}}`
  as an item's text), its facts naming the paragraphs it went into.
- **`set_cells` takes one value per cell** (R66): a cell's `from` names one value of a
  dataset or another document's table (`d:scores/#12/rating:0.0`, with a number format or
  a date pattern as a text reference takes), and a cell's `text` takes references. The
  result's `facts` name each value's cell; the cell keeps its source. A top-level `from`
  naming one value with cells that give only places is read as each cell's `from`, with a
  note; alone it is refused with the call that works. A whole dataset or a run of its rows
  fills the table as before.
- **A comment on a whole table or row** (R68): `word_comments` add (and
  `Document.add_comment`) takes a table's id or a row's. The comment covers it as Word
  16.106 anchors one on a selected table or row (measured: the range from the start of the
  first cell's first paragraph to right after the table -- or the row, unless it is the
  last -- and the reference opening the next paragraph); Word resaves the result
  unchanged. Deleting the paragraph after such a table, or the row after such a row, moves
  the reference on rather than refusing.

## 0.1.70 -- unreleased

- **A list that counts on through the copies of a sample, as the caller states it**
  (production round 43). Word lets one plain numbered list run on through every copy of a
  sample in a document on purpose (1-2 in the first, 3-4 in the next). The sample alone
  cannot show it, and in Word's XML such a copy looks the same as one that should have
  begun again, so it is never inferred: a `check` `copies` item takes `"lists": "run_on"`
  (ooxml-edit 0.94.0; `sample`, the default, keeps 0.1.69's rule). The copies of one
  reference stated so are a run, in document order: the first numbers as its sample shows,
  each later one goes on from the copy before it, in the same list (`runs_on_from` names
  it). A copy that restarts, or starts a list of its own, is `list_numbering` `starts at
  1` against `runs on`, fixed by `word_lists` `continue` -- never a restart; no restart is
  offered inside a running list, and the level and the label compare as before. Headings
  numbered by the outline are unaffected. A run of one copy is said in the reply's `note`.
- **`word_lists` `continue` takes `join`**: a paragraph of the list to go on from, where it
  is not the nearest earlier list over the same definition (`Document.continue_numbering`'s
  `from_id`). Given to another action it is left out, with a warning.

## 0.1.69 -- unreleased

- **`check(copies=...)` no longer counts numbered headings as list differences** (production
  round 43). Since 0.1.68 a copy's title numbered by its heading style through the
  template's outline (1, 1.1, 1.1.1) was told as `list_numbering` `starts at 2` against
  the sample's `starts at 1`, with a `word_lists` `restart` that would break the outline's
  numbers: every copy after the first failed to follow. Where a numbered item stands in
  its count is now compared only where the sample's own count begins inside the sample (a
  numbered list of its own: each copy's begins and goes on alike, a copy running on from
  the one before is still told, fixed by a restart). Where the sample's item is numbered
  by the outline (an outline level with numbering from its style, or a level tied to a
  style -- read from the structure, not the style's name), or its list runs on from before
  the sample, level and label are compared and the number is not; a copy's item numbered
  by the outline is never offered a restart.

## 0.1.68 -- unreleased

- **`check(copies=...)` compares lists and page settings as Word applies them** (production
  round 42). The copies comparison read a paragraph's indents and keeps as its own XML
  states them and its list membership not at all, so copies with dashes among round bullets,
  an item a level too high or a keep lost were said to follow their sample. Now each paired
  paragraph (a block, a table's cell paragraphs) is read through the style cascade --
  docDefaults, table style, paragraph style, list level, direct formatting -- and compared
  with its reference's: `list_level`, `list_format` (the label: `bullet •`, `decimal %1.`),
  `list_numbering` (`continues` or `starts at N`, counted over the body as the layout
  counts: a copy whose numbered list goes on from the chapter before says so),
  `indent_left`, `indent_first`, `keep_with_next`, `keep_lines`, `page_break_before` and
  `widow_control`. Values, not XML: a keep stated directly and by the style, an indent
  stated as the list gives it, a restarted bullet list or another list drawing the same
  label are no difference. One difference per cause (a level differing hides its label's
  and indents', a label differing hides its indents'; a pair whose styles differ waits for
  the style fix). Each with its fix -- `word_lists` `level`, `add`, `remove`, `format` (with
  `text_at`/`number_at` where the indents differ), `restart`, `continue`, or `word_format`
  (`clear` where the style gives the sample's value) -- and the reply's `list_fix` joins every
  copy's `word_lists` calls (at most 40, `more` past them); `word_follow_sample`'s `fix` no
  longer names a copy for list differences it keeps by design. A kept comparison is also
  keyed by the list items before the copy (an XML digest, nothing parsed). The paragraphs
  are read without parsing the package: each from its own XML with docx2svg's property
  reader, its table's style and cell conditions, resolved against the styles, numbering and
  settings alone (parsed once while those parts are the same); where docx2svg's final view
  joins paragraphs (a deleted or moved mark, a deleted row) the whole parse, as before -- the
  same values either way, paragraph for paragraph, on every fixture document. A reference in the
  same document gets `word_format` `like` its paragraph as the fix (round 42's `like`), one
  call per reference paragraph, wherever that takes nothing else compared -- not where a
  numbered item's numbering starts or a numbered item would join the reference's list
  (counting on from it), nor a page setting of a paragraph in another list; there, and for
  a reference in another document, the explicit calls. Cost, 22
  copies of a chapter with lists in a 70-block report, CPU per check: after an edit 0.307 s
  on 0.1.65 and 0.312 s here (medians of 8 runs each), the same version again 0.074 s both.
  No tool definition changes.
- Tests: `tests/test_shared_layout.py` shows a check after a render of the same version, and
  a render after a check, lay nothing out (the worker holding the layout answers, `"same"`),
  in `read_batch` and after a batch too; and the one case that lays out again, the holder
  busy longer than the pool's wait.

## 0.1.67 -- unreleased

- **Paragraphs made to look like a given one** (production round 42). `word_format` takes
  `like`: a paragraph whose look each target takes -- its paragraph style; its list at its
  level, numbered on from it (its own `numId` and `ilvl` where it has them, none where its
  style numbers it, `numId` 0 where it is taken out of its style's list); its direct indents;
  its page settings (keep with next, keep lines together, page break before, widow control).
  Each is taken whole (where the sample has none, the target's own goes); run formatting,
  spacing, alignment and tabs stay the target's (the result says so). A list a target leaves
  unused is removed. The call's other settings win where both say something (an explicit
  `paragraph_style`, keep, `keep_together` or `clear`). The result lists per target what was
  set (`data.like.set`) and the targets already alike (`alike`, no change; a call where every
  target is alike changes nothing and says so). A range, a cell or a span of one paragraph
  as the sample is that paragraph (a note says so); several are refused with their ids.
  Python: `Document.format_like(identifier, sample)` (`docx_agent.edit.like`).
- **New list items like the sample's, by level.** `word_insert_markdown` takes `like`: the
  paragraphs list items look like, by level (`["p:3", "p:4"]`: the top level like p:3, the
  second like p:4; one id for the top level; `""` for a level written as before). A level
  below the given ones is the nearest given one's list a level deeper per level, as Word's
  Tab demotes an item (the same style, the level's own indent). A style map naming a
  paragraph for a list role (`{"element": "bullet2", "style": "p:4"}`) is read as `like` for
  that role, with a note. Warnings say where a sample is in no list, or a bullet item stands
  for numbered ones. Python: `insert_markdown(..., like={1: "p:3", 2: "p:4"})`.
- **`word_lists` `remove` on items numbered by a list of their own over a numbering style**
  -- a builder removing a stray list's numbering, where `force` then takes every marker --
  now also names the `word_format` `like` call that joins the nearest earlier list at their
  level instead, keeping their marker and level.
- With a replacement (`at: "replace:..."`, 0.1.66's kept settings), list items written
  `like` a paragraph take its look, not the replaced items' settings.
- Tool definitions: `like` on `word_format` and `word_insert_markdown`, +66 tokens offline
  to every definition, none to the core. `word_insert_markdown` is sent strict: its tenth
  optional parameter takes the strict set to 6 tools and 24 optional parameters, Claude's
  limit (no tool leaves it).

## 0.1.66 -- unreleased

- **`word_set_text` writes a paragraph's style and its emphasis in the same call**
  (production round 42). Each item (and the single form) takes `paragraph_style`, a
  paragraph style the template has, by name or id; one it lacks is refused with the
  paragraph styles there are, and the same style again is a note, not a change. The text
  may carry inline emphasis in the template's character styles, as `word_insert_markdown`
  writes it: `**...**` the strong role's style (Strong), `*...*` the emphasis role's
  (Emphasis), and a span `[text]{Character style}` (also `{style="..."}` and Pandoc's
  `{custom-style="..."}`) any character style the template has -- one it lacks is refused
  with the ones there are. Marks inside a span are direct bold or italic beside its style,
  as `insert_markdown` writes a run with several marks. The result's `data.styled` lists the
  styled runs (`range`, `text`, `style`), at most 40, the rest counted. The text is read
  as CommonMark reads `*` (`5 * 3` and an unclosed `**` stay text) and as nothing else:
  `_` is never emphasis, links, code, HTML, entities and line-start blocks are their
  characters; a backslash before punctuation writes it (`\*` is an asterisk, `\[` a
  bracket). **Text with markup states the paragraph's emphasis:** its marked pieces get
  their styles and every other run loses its character style and direct bold and italic
  (other run formatting -- colour, size, font, language -- is kept as the diff keeps it; a
  hyperlink's runs keep their look); a repeat changes nothing. **Text without markup is
  written exactly as before**, its formatting kept. One undo step; tracked, the style
  changes are formatting revisions (`w:rPrChange`). Python: `markdown.styled.parse`,
  `Document.style_spans`. word_set_text has five optional parameters now and is sent
  non-strict to Claude (the dispatcher still validates it); edit_smartart takes its place
  in the strict set (6 tools, 23 optional parameters). +69 estimated tokens in the core.
- **`word_insert_markdown` at `replace:` keeps the replaced paragraphs' settings**
  (production round 42). Before, the new paragraphs had the style map's defaults only:
  replacing paragraphs lost their keeps, page break, indents, paragraph style and list
  membership (a replaced nested item's list became a new list). Now each new body
  paragraph takes the paragraph properties of the replaced paragraph of its kind -- a
  paragraph, a heading of the same level, a list item of the same kind (bullet or
  numbered) and level -- at its position, the last one's past the last (without its page
  break); a list's nesting counts from the replaced items' level, so a flat list replacing
  nested items stays nested, and items join the replaced items' list and count on in it.
  Markdown's word wins: the kind, a numbered list's start, a style the style map names
  (then only the keeps, page break and widow control are kept). The result's `data.kept`
  says per new paragraph what it kept and from which (at most 30, the rest counted);
  instances the Markdown list would have used and nothing names are removed. Python:
  `Document.insert_markdown(..., keep_settings=True)`, `EditResult.kept`
  (`docx_agent.markdown.keep`). No tool definition changes.

## 0.1.65 -- unreleased

- **A check's facts read the parse its layout was made from** (production round 41). A check
  made in the worker holding the document's layout opened a copy of the document from the
  bytes laid out, and its layout facts parsed those bytes with docx2svg a second time
  (`edit.effective`'s model; about a quarter of the facts' CPU). The copy now takes docx2svg's
  parse of the very bytes the layout was made from (`docx2svg.parsed_document`, docx2svg
  0.11.10) where the copy's bytes are those; with an earlier docx2svg, or other bytes, it
  parses them as before. The facts are the same: every call's result of a made-up build --
  two 11-page documents and the 46-page report, edited and checked, through a Toolbox with 2
  workers -- is 0.1.64's byte for byte, and
  `tests/test_worker_facts.py::test_the_facts_read_the_parse_the_layout_was_made_from`
  compares the facts both ways. With docx2svg 0.11.10 the worker's layout of each new version
  also builds the model only of the blocks that changed (docx2svg's kept parse, held with the
  layout in the worker, within what a held layout costs). No tool definition changes.

## 0.1.64 -- unreleased

- **`check(range=...)` says what a range's keep does to the pages, laid out** (production
  round 40). A range of whole blocks parted by a page break, short enough for one page, has
  a `fix` (`word_format` `keep_together: true`); what that does was not said: Word moves the
  range whole to the next page, which can leave much of the page before it blank and push
  what follows down -- seen only in the next render, then released, then offered again by
  the next check. Now each such range has `if_kept`: the document with that call's change
  alone, laid out from the layout checked (docx2svg's incremental layout, in the worker
  holding it; page for page a layout of those bytes alone), and read there -- `starts`,
  `ends`, `one_page`, `blank` (the page before it: `pt` and `share` of the body's height
  left blank at its foot, `now_pt` and `now_share` as it is now), `document_pages` (now,
  then), `after` (the page what follows starts on, now and then) and `said`, e.g. "keeping
  it whole (this range alone): starts on page 14, on one page; page 13 left 20% blank
  (142.2 pt), now 2%; the document stays at 46 pages; what follows stays on page 14". The
  other side for a range kept already that starts a page only because of its keep (no
  page break or section start before it): `if_released`, the same facts with the keep
  released and that call as its `fix` (`keep_together: false`) -- so an agent sees both
  sides and decides once. Each range is tried alone. At most four trials a check, none
  started past 6 s after the first; a range past them says so (`effect_left_out`) and how
  to ask (check it alone). Where a check reads the layout in the caller's process (a
  converter that cannot make the facts where the layout is held), no trial: the range says
  so.
- What a trial costs, measured as CPU time (the machine was shared): 0.79 s a range
  (median; 0.69-1.00 s over 19 parted ranges) on a made-up 46-page report with a contents,
  page cross-references and a page count, 0.71 s (0.53-0.88 s, 4 ranges) on a made-up
  11-page document -- about what laying the document out again after the keep costs, and
  against 0.16-0.19 s and 0.08-0.11 s for the same scoped check without it. A range on one
  page, taller than a page, or kept and not starting a page costs nothing. On by default
  for the ranges a check is given (`docx_agent.tools.shared.KEEP_TRIALS`); the caps are
  `docx_agent.tools.keep_trials.TRIALS` and `TRIAL_SECONDS`. No tool definition changes.
- A trial's result is kept in the worker that made it (`keep_trials.TRIAL_CACHE`, 256, least
  recently used out; never pickled), keyed by the digest of the bytes checked, the options,
  the range, kept or released and the caps: the same range of an unchanged document
  checked again lays nothing out, and an edit is another key. On the 46-page report, a
  repeated check of a parted table (other pages asked, so not the caller's kept facts)
  took 0.074 s of CPU (median, 6 tables), against 0.079 s without trials and 0.56 s for
  the first, tried.

## 0.1.63 -- unreleased

- **`word_format` formats text it finds** (production round 40). `find` takes items of
  `text`, `within` (a block or span; the body by default; `range` read as `within`),
  `occurrence` (from 1) and `expect` (`one`, the default, or `all`), as `replace_text` finds
  text, so a text and its character style go in one call instead of a search for offsets
  first. Every item is found before anything changes: a text found nowhere is refused with
  where it occurs written otherwise (case, spaces) and what `within` holds; one matching
  several, with each match's `occurrence`, range and context. `targets` is optional now (one
  of `targets` or `find`), and either takes one value in place of a list.
- **One way to name a place** (`docx_agent.tools.positions`). `at` -- `end`, `after:<id>`,
  `before:<id>`, `replace:<id>`, a position `p:A@5`, or a bare id read as after it -- is the
  form six of the nine placing tools took; `word_insert_text` takes it too, beside `target`
  and `where`, and `word_move`, `word_sections` and `word_fields`' TOC and caption read an
  `at` (`before:p:A`, `end`) as their `after`/`before`. Each placing tool reads another's
  name for a place as its own where that says the same thing (`after`, a `target` alone,
  `position`, `location`; a warning says so), and refuses one whose side is in its name
  (`before`, `where` with a `target`) with the arguments that work in `details.call`
  (`details.item` for an item). `word_insert_markdown` and `word_copy_from` take `ref`, the
  blocks made named as one id or a span; `after:$ref` is after a span's last block,
  `before:$ref` before its first. A `ref` holding an address says the place is `at`.
- **`word_edit_table` `set_cells` `from` adds the rows the data needs** with `grow: true`:
  after the last row, each as `insert_row` makes one (the last row's properties and cells,
  never a header row), in the call's undo step, tracked while tracking is on; `index` may be
  the row after the last, to append. Without it the refusal gives that call in
  `details.call`. A source with fewer rows than the table leaves the rows below as they
  were, and the warning now names them and the `delete_row` that removes them.
- Definitions: +236 estimated tokens to every definition (15,050 -> 15,286), none to the
  core; the strict set is 6 tools and 24 optional parameters (`word_insert_markdown`'s
  `ref`): it is full -- Claude's limit -- so the next optional parameter on a strict tool
  sends a tool non-strict.

## 0.1.61 -- unreleased

- **A check's summary digests only what was laid out again** (production round 39). Where
  the worker holding a document's layout lays the next version out from it, docx2svg keeps
  the pages and lines the edit did not reach -- the very objects -- and the summary a check
  compares by (`layout.LayoutSummary`, round 38) took a digest of every page and line again
  each time: now it takes those of the kept pages and lines from the summary of the layout
  before (`LayoutSummary(layout, before, keep=True)`) and digests only the rest. The summary
  is the one made whole, field for field (`tests/test_worker_facts.py`); it is pickled
  without what it holds for the next, and the summary before is let go when the task ends.
  Measured on a made-up 46-page report, a paragraph near the middle edited each time: 190
  to 210 ms of a worker's CPU per check that asks for a summary, now 10 to 12 ms.
- With docx2svg 0.11.9 (not required): a layout takes about two thirds of the CPU and a
  page's SVG about two fifths. Every call's result of a made-up build -- two 11-page
  documents (chapters, checks, a copied section, a table, fields, renders, save) and a
  46-page report edited and checked -- is 0.1.60's on 0.11.8, byte for byte.

## 0.1.60 -- unreleased

- **`render`'s `fast`: a page for about a third of the CPU while building** (production round
  39; ooxml-edit 0.92.0). An exact page is drawn by resvg from its SVG, which shapes each
  positioned glyph of the page alone (four fifths of a text-heavy page's time), and twice
  where two faces answer to one name (Aptos and Aptos Display: a page with a heading is drawn
  in two passes, the second over the first). `fast: true` outlines the text first (docx2svg's
  `outline_svg`: the faces, fallbacks, shaping and positions resvg would use, every pass at
  once) and rasterises once: 0.83 s against 2.06 s on a report page with a heading, 0.89 s
  against 1.37 s a page on average over 28 pages. The layout, page breaks and positions are
  the same SVG's; the pixels differ along glyph edges only. The result says so:
  `"approximate": ["pixels"]`, `approximate_pages`, each such image labelled
  `page 3 (approximate)`, and the summary the way to the exact image (render without
  `fast`). A fast image is kept apart from the exact ones; a page whose exact image is held
  gets that one. Exact renders are unchanged, byte for byte.
- Rendering runs in the pool's render lane (ooxml-edit 0.92.0): it yields to layouts and
  checks waiting for a worker, its deadline counts from when a worker takes it, and a worker
  past it finishes the page in the background (kept for the next call) rather than being
  stopped and started again with its layouts lost.

## 0.1.59 -- unreleased

- **A worker keeps as many documents' layouts as its memory budget holds** (production
  round 39; docx2svg 0.11.8). A process held two layouts (`_HELD`), two pickled ones
  (`_RECENT`) and two documents' measurements (`_MEASURES`): with six documents built at once
  on two workers, nearly every check laid its document out whole. They are now bounded by
  `docx_agent.layout.LAYOUT_MEMORY` -- 256 MB a process by default,
  `DOCX_AGENT_LAYOUT_MEMORY` (MB) to set it -- at a megabyte a page held (`PAGE_COST`,
  measured: 1.05 MB a page on a made-up 11-page document, 0.95 MB on a 46-page report):
  held layouts up to the budget (about 23 ten-page documents, or five 46-page reports), the
  measurements kept by lineage up to half of it (`MEASURE_COST`, 96 kB a block measured), a
  caller's own pickled layouts up to a quarter; the least recently used dropped first, the
  newest kept whatever its size. At most 1.75 times the budget in all, plus the layouts the
  tasks running at that moment hold. The study's made-up documents, two workers, half a core
  of other load (on a machine busy with other work, so absolute times are high): six
  builders 32.7 -> 18.2 s of CPU a document (worker CPU halved; documents per CPU-minute
  1.84 -> 3.30), twelve builders 31.2 -> 17.7; one builder and the 11-writer report on one
  document unchanged. Workers' largest resident memory with twelve builders 144 -> 250 MB.
- With docx2svg 0.11.8, a document's first save (a settings part written again, revision
  ids) no longer lays it out whole, and `word_fields` writing the page count into a footer
  lays the document out whole with the body's measurements (where its headers and footers
  hold no drawing, picture, object or frame).
- Pages and facts are the same, byte for byte.

## 0.1.58 -- unreleased

- A table cell merged with another after its values were labelled (a scale's `show: label`)
  holds both labels, one per paragraph, and at most one record: `check(scale=...)` now lists
  it under `unrecorded` (its steps in `label_of`; ooxml-edit 0.90.1), no longer skips it as
  "2 paragraphs of text"; `style_by_value` still writes nothing there. A cell rewritten from a
  source keeps its label's record, as any edit of a label does: a value written there is
  listed as a value typed over its label, with the fix.

## 0.1.57 -- unreleased

- **A document's layout stays in the worker that made it; a check's facts are made there**
  (production round 38; ooxml-edit 0.90.0). Every layout task pickled the layout back to the
  calling process, which read it (about 0.4 s of its CPU on a 46-page report), mapped it to
  ids and made the facts there, in the one interpreter every caller shares; the next task
  was handed it again. Now a layout is held in the process that made it (one per document,
  two documents a worker, as before) and named by its token; the document's layout tasks go
  to that worker (ooxml-edit's pool affinity, keyed by the document's lineage); a task
  waits for that worker, while it is busy, at most as long as laying the document out whole
  took last (going elsewhere costs at most that: the other worker lays out from the older
  layout it holds, or whole), then takes another. A lone check's layout facts -- page count,
  coverage, each range's pages and extent, the layout facts of the pages asked for -- are
  made in that worker on a copy of the document opened from the very bytes laid out (ids
  are the bytes' own), and only they come back; `since_last_check` compares small summaries
  of the two layouts here, ids followed through the document's aliases (as before). A
  render's layout task is the same: nothing pickled back. What stays here: validate, the
  fields, copies, a scale's values, and the reads that need the layout itself (word_inspect,
  a reflow in a short document's edits, update_fields), which fetch it from the worker
  holding it. Saving still writes the pages and lines of the layout checked last. Facts
  are the same, byte for byte: 131 documents compared through both paths (checks whole, by
  page, by range, after an edit), and a test replays checks, edits and ranges through both.
  Scoped facts resolve a table's row and cell keys from their row's element (half a
  scoped check's time went to resolving each from the root). Made-up 46-page report, two
  workers, 11 threads each editing its chapter, checking it and rendering a page, 4 rounds,
  4 CPUs with half a core of other load: check median 10-13 s -> 4-6 s, longest 31-36 s
  -> 11-30 s, wall 115-121 s -> 67-86 s, this process's CPU 48-51 s -> 13 s, workers' CPU
  196-204 s -> 124-160 s, layouts pickled 105-129 MB -> none; one writer: check 1.6-1.7 s
  -> 1.1-1.2 s.
## 0.1.56 -- unreleased

- **Retired ids stay retired across a save and a reopen** (production round 38; ooxml-edit
  0.89.0). Opened from its file alone, a document could give again an id deleted after the
  last surviving one of its sequence (the round 37 limit). `Document.id_state()` exports what
  the document's ids have been -- every value it had, each deleted id with the change that
  deleted it and who made it, drawing numbers, renames, the version -- as a short opaque
  string to keep beside the bytes; `Document.open(data, ids=state)` takes it back: no id is
  given again, naming a deleted one says `... was deleted by word_delete (agent=writer-2),
  the change that made version 7, before the document was saved and opened again`
  (`DeletedId.earlier`), and versions go on from the saved one (the next change gets one
  more than the highest the earlier session gave). Word's bytes are not touched. The tools:
  `Session.save(doc)` gives `data` and `state` at one version, `save_document`'s output
  carries `Output.state`, `Session.open(data, name, doc_id=..., state=...)` takes it back.
  A state of other bytes (a digest of their ids), damaged or of a later form is refused
  (`IdStateError`; the tool layer's `StateRefused`) with the way out; without a state,
  nothing changes. Size: 294 bytes for a 300-paragraph document after 1,500 edits and 600
  deletions; 427 bytes for 2,000 paragraphs and 1,000 deletions by 11 callers; about 5
  bytes per retired id Word wrote. No tool definition changed: offline tokens unchanged.
## 0.1.54 -- unreleased

- **A step's label in place of its value** (production round 38; ooxml-edit 0.87.0): a scale
  with `"show": "label"` makes `style_by_value` replace each value's characters with its
  step's label (a cell reads `High`, not `4.5`): the label takes the value's first character's
  formatting, then the step's character style (and a cell its shading), and the paragraph's
  entry in the provenance item keeps it beside its text facts with the value it stands for and
  its reference (`Document.text_labels`, `Document.keep_text_label`) -- through a save, a
  reopen and Word, as text facts are. An edit of a label keeps its record, stretched over the
  new text (a paragraph's or cell's whole text stays whole), so `check(scale=...)` can list the
  `labels` that are not their value's -- the value changed at its source, the label edited by
  hand, a value typed over it -- and the `unrecorded` ones (a step's label with no value kept),
  with the one fix. A value written by reference (`{{t:...}}`, `{{d:...}}`) found as a label
  is read again from its source where the session can. Exact text steps are unchanged; a cell
  takes a label in one paragraph of text. Tracked, the label is a deletion and an insertion.
- Tokens (offline estimate): the scale's `show` +14 in the core (3,034 -> 3,048) and +60 in
  every definition (14,950 -> 15,010, guard 15,010); a reviewer's core 2,908 -> 2,923, all
  8,256 -> 8,270. The strict set unchanged (6 tools, 23 optional parameters).

## 0.1.51 -- unreleased

- **Layouts after an edit cost a third as much in a worker pool** (production round 37; docx2svg
  0.11.7). A document is laid out from its last layout in whichever worker is free, handed
  over as a pickle without what its flow measured (most of it); the worker then measured
  every block again, about a whole layout. Each worker process now keeps what it measured
  of the document (docx2svg's `MeasureStore`, by lineage: two documents' worth, about 25 MB
  for a 46-page report) and measures only the blocks that changed: one paragraph edited,
  1.9-2.5 s of CPU -> 0.6-0.7 s. A check no longer has docx2svg find the pages whose drawing
  changed (nothing read them; after an edit that adds a page under a "Page X of Y" footer it
  drew every page: 2.65 -> 0.77 s). Layouts are pickled with their fractions rebuilt from
  their terms (a third less time each way). Many versions laid out at once: the layout
  the next version starts from is the newest started, not the last finished.
  Made-up 46-page report, 11 threads each editing its chapter then checking and rendering
  its page, 4 rounds, 4 CPUs: wall 164 -> 103 s, check median 18.1 -> 11.7 s (longest 43 ->
  29), render median 13.9 -> 10.0 s, workers' CPU 466 -> 276 s. Edits never wait for a
  layout (median 0.2 s throughout): a check and a render let the document go while they
  lay out. Pages, facts and renders are unchanged.

## 0.1.50 -- unreleased

- **An id is never given to another block** (production round 37, R36). A deleted block's
  id was free again, so the next new block got it, and a caller holding the old id reached a
  block it had never read. Now every id a document's session had -- in any part, in any
  state its history went through -- is kept beside the ids in use (`docx_agent.edit.ids.Ledger`)
  and never generated again; an undone deletion brings its block back with its id, and redo
  writes the same bytes. Naming a removed id fails as deleted (`DeletedId`, a `KeyError`;
  the tools' `not_found`), with the change that removed it -- "p:16425160 was deleted by
  word_delete (agent=writer-2), the change that made version 7" -- and how to read again;
  a block replaced by one block names its replacement. Copies keep their own ids only where
  the document never had them; drawing numbers (`d:`) are never reused either. A document
  saved and opened again goes on after the last generated id its file still has, so the
  holes deletions left are not filled (ROADMAP.md, "An id is never given to another block").
- What stays the same block keeps its id: Markdown written into a new document's one empty
  paragraph continues it (its paraId and textId, as typing into it in Word), and a table of
  contents rebuilt by `update_fields` keeps each line's ids in its place (so an update that
  changes nothing writes the same bytes); lines it drops are retired, lines it adds are new.
  A copy that replaces the blocks it was copied from (`copy_blocks(self, X, at="replace:X")`)
  now gets new ids, and `X` says it was replaced with the copy.
- Generating ids is now linear in the ids drawn: the seed's sequence is cached and each
  document remembers where it got to (it walked the sequence from its start on every call);
  reading the ids in use walks the tree once (`used_long_hex`; it walked it sixteen times).
  Each edit reads the block ids of the state it made once (about 9 ms on a document of 2,100
  paragraphs and 600 cells, on a loaded machine).
- Golden transcripts w5, w6, w9, w14 re-recorded for the ids now generated (holes skipped,
  retired values not reissued); w3 and batch-30-edits for the saved bytes (textIds).
- Tokens: no definition changed (error text only).

## 0.1.49 -- unreleased

- **Values styled by a scale, applied once** (production round 36, R31; ooxml-edit 0.85.0):
  `style_by_value(doc, scale)` gives each value of the scale's `targets` its step's style --
  a paragraph, a range (`p:A@4:5`), a cell, a table's column (`t:X/col1`: its body cells, a
  repeated header row, or the first where the table's look has a header row, left out), or
  the values written by reference (`{{t:checks/*/rating}}`, `Document.text_facts()`: the
  value's characters, or its cell where it is the cell's whole text). One character style
  per step, named from the scale (`Score 1 Low`), made once and reused, as Word writes one
  made with a font colour and shading (measured in Word for Mac 16.106: a made-up document,
  the style made and applied by script and saved -- `w:customStyle`, the id without spaces,
  based on Default Paragraph Font, `w:color` and `w:shd w:val="clear" w:color="auto"` in its
  `w:rPr`, no `w:qFormat`; Word keeps ours as written on a save and draws them, as docx2svg
  does); a cell is shaded as Word shades one and its text gets the style; a run's own colour,
  shading and highlight are taken off. A style whose colours are not the step's any more is
  set to them, with a warning (its uses follow). One undo step, tracked as formatting where
  tracking is on; `preview` says which styles it would make. `check(scale=...)` lists the
  values styled with another step (a value changed after styling), with none, or with the
  step but other colours, and the fix (`docx_agent.tools.scales`).
- Tokens (offline estimate): check's `scale` +158 in the core (2,876 -> 3,034, about 4,490
  counted by the proxy); `style_by_value` (deferred, `shared_misc`) +464: every definition
  14,328 -> 14,950 (guard 14,950); a reviewer's core 2,750 -> 2,908 (guard 2,910), all 8,096
  -> 8,256 (guard 8,260). The strict set unchanged (6 tools, 23 optional parameters).

## 0.1.48 -- unreleased

- **The outline: what each heading owns** (production round 36, R32). `Document.outline()`
  gives every heading of the body in order with its `level`, `text`, `range` (`h:p:A`), its
  `last` block, how many `blocks` it owns and the `section` it is in (`s:2`; `to_section`
  where its last block is in another). A heading's range -- the heading and every
  top-level block after it up to the next heading of its level or higher, tables, content-control blocks,
  deeper headings and section breaks included -- is `h:<heading>`, taken wherever a span is
  (`p:A..t:B`): `to_markdown` and `state`, `move_blocks`, `keep_blocks_together`, and the
  tools' `range`, `target`, `targets` and `blocks` (`word_read`, `word_inspect`, `find_text`,
  `check` and `render` `range`, `check`'s `copies`, `word_format`, `word_delete`, `word_move`,
  `word_lists`); a span's end may be one (`h:p:A..h:p:B`). `Document.heading_span` gives it
  as the span. A heading is a paragraph of the body itself with an outline level as Word
  reads it for a table of contents (`Document.blocks`' `level`): a Title is one only where
  the document gives it a level; a heading in a table, text box, content control, header or
  note owns nothing (`h:` on one is refused, naming the headings). `describe`'s `headings`
  add `section` in a document of several sections; a range that names nothing is refused
  with the headings as `h:p:A  Text` (shorter than the spans before).
- **Guards for a heading's content and a style's numbers** (R33). `word_delete` of a
  heading whose content stays (a paragraph, or a span that leaves some of what it owns) is
  refused, saying how many blocks would go under which heading, with the ways out: its
  range `h:<heading>`, or `force` (the warning then says where they went). `word_format`
  changing a heading's level (`paragraph_style`, `clear_direct`, `clear`) warns -- it does
  not refuse -- which headings and blocks now go under which heading
  (`docx_agent.edit.outline.moves`). `word_lists` `remove` on a paragraph its style numbers
  is refused (taking it out writes an override the style no longer decides), naming the
  paragraph styles like it that number nothing -- one based on it, one at its outline level,
  its base, one beside it, the default -- and `force`, which writes the override as before.
- Tokens (offline estimate, on 0.1.47): core 2,874 -> 2,876, every definition 14,270 -> 14,328
  (guard 14,275 -> 14,330), a reviewer's core 2,748 -> 2,750 (its guard); the strict set
  unchanged (6 tools, 23 optional parameters).
## 0.1.47 -- unreleased

- **Numbers in text no value holds** (production round 36, R28; ooxml-edit 0.84.0): `check`
  with `include: ["numbers"]` reads each body paragraph's and table cell's text
  (`FORMAT.number_texts`, `docx_agent.tools.textfacts`) as `find_text` reads it -- deleted
  text left out, a note's or drawing's mark a space -- and lists the decimals and
  percentages no value of the session's datasets or open workbooks' tables matches within
  the rounding the text shows (`unverified`, with the nearest values), and the typed
  numbers a value matches (`typed`, with the `replace_text` call that writes each by
  reference, scoped to its paragraph or cell). `range` narrows it as `check` resolves a
  range, `pages` by the layout (laid out with the lock let go, as `check` does); each text
  says its `page`. Left out: tables placed with `from`, a table of contents, headers,
  footers and notes, and the text written by reference. Opt in, capped; whole numbers and
  the ignore list are the application's (`Toolbox(numbers=NumberRules(...))`).
- **Text written by reference is kept in the file** (R28): a value a text tool writes by
  reference (`{{d:...}}`, `{{t:...}}` in `word_set_text`, `word_insert_text`,
  `replace_text`, `word_insert_markdown`) is kept on its paragraph in the provenance item
  (`kind="text"`, keyed by the paragraph's id, its text and each value's place, source,
  format and sources; `Document.text_facts()`, `Document.keep_text_facts()`), found by the
  text around it in the argument; `replace_text` keeps every place it wrote. It follows its
  text: every edit keeps the values it did not touch, moved, and drops those it touched or
  borders; a paragraph renamed keeps them, one deleted or left without any drops its
  entry; undo brings them back. Read after an edit made elsewhere (Word keeps the
  paragraph's id through an edit and a save), the values the changed span left alone stay.
  A split keeps the part with the paragraph's id; copies and joined paragraphs' second
  halves keep none.
- Offline tokens: check's include gains `numbers`: the core 2,871 -> 2,874 (about 4,254
  counted by the proxy), every definition 14,267 -> 14,270 (guard 14,270 -> 14,275); a
  reviewer's core 2,746 -> 2,748, all 8,095 -> 8,096 (guards unchanged). No parameter added:
  the strict set and its optional count are unchanged.

## 0.1.46 -- unreleased

- **A table parted from the line that introduces it or the line after it** (production
  round 36, R29). `check` with `layout` lists `table_parted`, beside `keep_with_next_at_foot`:
  a top-level table drawn on a later page than the paragraph just before it
  (`{"table", "para", "side": "lead_in", "kind", "para_page", "table_page", "fix"}`), or the
  paragraph just after it drawn on a later page than its last row (`{"table", "para", "side":
  "line_after", "kind", "row", "row_page", "para_page", "fix"}`) -- where that paragraph
  belongs with the table: `kind` `caption` (the Caption style), `colon` (a lead-in ending in
  a colon) or `one_line` (drawn on one line: a title, a source line). A longer paragraph of
  text is left out, as are a paragraph or a last row that keeps with the next already
  (`keep_with_next_at_foot`'s) and a page break or a section start between them. The fixes,
  measured in Word for Mac 16.106 (made-up probe, 33 cases, all drawn as docx2svg draws
  them): the lead-in keeps with the next (`word_format`) and goes to the table's page with
  its first row; a table on one page is kept together and with the next
  (`word_format_table`) and goes whole to the line's page; a table over several pages keeps
  its last row's deciding paragraph (the first of its first cell) with the next
  (`word_format`), and that row alone goes along (every paragraph of the row keeping is the
  same). Scoped by `pages` and `range`, at most 20, `table_parted_left_out` the rest.
- **A span of blocks kept on one page** (R30). `word_format(targets=["p:A..t:B"],
  keep_together=true)` -- also a table, a section or a block-level control -- writes what Word
  needs on each block (`Document.keep_blocks_together`): every paragraph keeps its lines
  together and, but the span's last, keeps with the next; every row of a table keeps with
  the next (its deciding paragraph), but the last row of a table that ends the span. Directly
  only where the styles do not give it; the span's last block keeps with what follows as it
  did. Measured in Word: a heading, a paragraph of seven lines, a five-row table and a line
  after move to the next page whole at every one of seven positions across the page's foot
  (21 cases with and without the flags: keep with the next alone lets the paragraph split;
  with its lines kept, whole); the documents the tool writes, 7 of 7 as docx2svg draws them.
  A row that keeps with the next is not broken over a page (Word moves it whole: 25 cases),
  so no `cantSplit` is written. `keep_together: false` removes what it writes and only that
  (a heading style's keep stays). Already kept, or not kept: no change, said in `warnings`.
  On one paragraph (or a range, a cell) `keep_together` keeps its lines together, as
  `keep_lines` -- as the Python API's name was read before, now without the warning. Where
  the document's current layout is at hand and draws the span taller than a page, the call
  warns: it cannot be kept (Word starts it on a new page and breaks it where the page is
  full).
- **Each range in `check(range=[...])` says whether it is on one page and kept** (R30: the
  caller's spans, nothing stored in the document; no new field): `one_page`, `kept` (every
  block holds the keeps `keep_together` writes) and, where it is not on one page, `breaks` --
  `{"after", "before", "page"}` between two blocks (`page_break: true` where the later one
  starts a page anyway, which no keep undoes), `{"inside", "pages"}` within one; at most 10,
  `breaks_more` the rest -- and `fix` (the `word_format` `keep_together` call) where its
  height fits a page's body, else `taller_than_page`. A range that is not whole blocks (a
  cell, a text range) gets none of these.
- **A table's frame in `word_design_facts`** (R34): `table_frames`, each top-level body
  table's stated width against the body between its section's margins (`body`; `over_body`
  where its indent and width reach past the right margin), its indent, its grid columns'
  widths, its style, and the rows whose last cell is empty with the height each states
  (`height`, `rule` `at_least`/`exact`; at most 10 per table); `usual_table_width`, the width
  most body tables share, and `narrower_than_usual` on a narrower one. Reported by default,
  after `table_styles`; at most 30 tables, `omitted` the rest.
- Not matched in docx2svg (measured, open): a row whose deciding paragraph keeps with the
  next is not split over a page in Word; docx2svg splits it (a kept-together table of tall
  rows at a page's foot: Word moves it whole). A chain of keeps taller than a page: Word
  breaks it before a kept table, docx2svg inside it. `check`'s facts follow docx2svg there.
- Offline tokens: every definition +43 (14,224 -> 14,267: `word_format` +38,
  `word_design_facts` +5), the core unchanged (2,871); the strict set unchanged.
  Replies: `check` with one span over two pages 56 -> 106 (its `one_page`, `kept`, one
  break and the fix; about +8 for a span on one page); `layout` +5 (`table_parted`, empty);
  `word_design_facts` of a one-table document 633 -> 664.

## 0.1.45 -- unreleased

- **A range's extent on its pages** (production round 35, R26). `check(range=[...])` gives
  each range, on each page it is drawn on, `extent`: `{"page", "top", "bottom", "body"}` in
  points -- the top of its first line and the bottom of its last there, from the page's top
  edge, as `DocumentLayout.where` measures a block, and the page's height less the top and
  bottom margins its section sets -- with `area: [top, bottom]` only on a page whose header
  or footer reaches past its margin and pushes the body in (docx2svg lays the body out
  there; `area` is the room it has), `columns` only where it is drawn in more than one; and
  `height`, `bottom - top` summed over its pages. At most 12 pages listed, the rest counted
  in `extent_more`. A block's own extent: name it as its own range beside the span
  (`range=["p:A..p:C", "p:A", "t:B", "p:C"]`) -- no new field. Python API:
  `DocumentLayout.extent(ids, pages=None)` -> `Extent(page, top, bottom, lines, columns,
  body, area)`.
- **The live view from the shared layout** (R27; `docx_agent.liveview`, an application API,
  not a tool). `view(session, doc, since=token)` (or `view(document)`) -> every page in
  order with the ids of the blocks it draws and its size, the pages whose drawing changed
  since `since` (all on a first call; `stale` with a note, and all, for a token no longer
  kept), a new token. `page_svg(session, doc, page, token, text="text"|"outlines",
  images="external"|"embedded", glyphs=None)` draws one page of the token's version -- even
  after the document moved on -- as `render_svg` draws it, pictures by name
  (`externalise_images`, pptx2svg's names), outlined text in one glyph set. Both read the
  layout the tools made of that version: `view` takes the document's lock for the bytes,
  lets it go while the version is laid out (waiting for a check's layout in progress there),
  and lays out on its own instead when its thread holds the lock already; `page_svg` holds
  no document lock. Changed pages are docx2svg's `changed_pages` for the same versions
  (tested over six kinds of edit).
- `DocumentLayout.blocks_on` maps the body's keys from one listing of the body (each key
  resolved alone listed it again: 0.7 s -> 0.09 s for 79 pages). A layout's pages are drawn
  under its own lock (the fourth in `docx_agent.layout`'s order).
- Measured on a made-up 79-page report, in process: a first view after a check 0.67 s, no
  layout (a third layout of their own: 12.4 s); after one edit and a check, view plus the
  changed page's SVG 0.27 s, no layout, against 1.49 s for docx2svg's layout from the
  previous one, its SVG and the blocks per page.
- Offline tokens: definitions unchanged (0). `check` with a range: +38 for one range over
  two pages (292 in all), +143 for a span and its four blocks (450).

## 0.1.44 -- unreleased

- **`set_language` retags the template's own language** (production round 35, R25;
  `docx_agent.edit.language.own_languages`). `styles="matching"` (still the default) now
  means the styles stating the template's own language, not only the old default's: in the
  attribute the new language is written in (`w:val`, `w:eastAsia` or `w:bidi`; the other two
  are not read), (1) `w:docDefaults`' language, (2) the Normal style's own language (without
  one, the default's), and (3) only when neither states one, the language more than half the
  styles stating one carry, when at least two do (a tie or a single style decides nothing).
  A template made in an English Word with a Dutch default -- Normal, headings, their
  character styles, footer and date styles in `en-US` -- is retagged whole; a style in
  another language (a German quotation style, an Italian character style) stays and is
  named, as are styles based on it. The warnings say which language was taken as the
  template's own and why when it is more than the default. `runs=True` retags runs in any of
  those languages. Runs are still untouched by default.
- `check`'s `punctuation_spacing` `language_mismatch` on a style in the template's own
  language (`style:Normal` of an English template) proposes the document's `set_language`,
  not that style's alone; a style in another language keeps its own fix.
- `word_styles` `set_language`'s summary names the tag written, not the default's other
  attributes.
- Offline tokens: definitions unchanged (0). On a made-up English template with French
  text (characters / 3.5): `set_language`'s reply 307 -> 265 (twelve styles kept became two),
  `check` (layout) before 289 -> 283, after 289 -> 148 (no `language_mismatch` left).

## 0.1.43 -- unreleased

- **The language of an existing document** (production round 35, R24;
  `docx_agent.edit.language`). `Document.set_language("fr-BE")` writes what a new document
  in that language has -- `w:docDefaults`' `w:lang` and the settings' `w:themeFontLang`, in
  the attribute Word writes the language in (`w:val`; `w:eastAsia` for Chinese, Japanese,
  Korean, Yi; `w:bidi` for the complex scripts) -- and the styles whose own language was the
  old default (`styles="matching"`; `"all"`, `"none"` or style names); styles in another
  language and runs' own tags stay (`runs=True` also sets the runs tagged with the old
  default; `default=False` sets the styles named only). One undo step; a no-op says so. The
  tag is checked against the tags Word writes (any case, `_` for `-`, `fr-Latn-BE` read as
  `fr-BE`, with a note); a bare language (`fr`), a region Word does not know (`fr-XX`) or an
  unknown language is refused with the tags that would do. `Document.languages()` says where
  the document states languages. Text written afterwards gets the language's spacing (R21).
  For the model: `word_styles` `set_language` with `language` (and `name` for a style's
  alone). Measured in Word (`tools/language_probe.py`, `tests/observations/language-word.json`):
  the tag and attribute Word writes for each language of its dictionary; how it reads tags;
  setting the Normal style's language writes the Normal style's `w:lang`, all the text's
  writes `w:lang` on every run, the Normal template's writes Normal.dotm only -- none
  rewrites `w:docDefaults`, `w:themeFontLang` or another style's language.
- **`check`'s `punctuation_spacing` no longer offers to delete French spaces** in text tagged
  English, German, Dutch, Spanish, Italian or Portuguese: such a paragraph is an entry of its
  own, `"kind": "language_mismatch"`, with the tag's `source` (`default`, `style:<name>`,
  `run`) and a `fix` that sets the language (`word_styles` `set_language`, French: the
  document's own French tag if it has one, else `fr-FR`); a run's own tag gets a note and no
  fix. French text keeps its `replace_text` fixes.
- Offline tokens: `word_styles`' `set_language` and `language` +56 to every definition
  (14,168 -> 14,224), none to the core (2,871); a reviewer's everything 8,037 -> 8,089.

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
- `ooxml-edit>=0.81.0` (was 0.79.0): its shared `check` carries `range` and `copies` and
  its `render` `range` (production round 35, R18/R19), so 0.1.41's scoped checks, scoped
  renders and copies reach the model. The tests use the toolbox's own definitions (no
  longer adding them); offline tokens +229 to the core (2,642 -> 2,871) and to every
  definition (13,939 -> 14,168), a reviewer's core 2,746 and everything 8,037.

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
