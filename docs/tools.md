# Tools for an agent

`docx_agent.tools` puts this API behind tool definitions a model calls, for the tool layer
in [`ooxml-edit`](https://github.com/uvrt/ooxml-edit) (`ooxml_edit.tools`, 0.4 or later): the
model never runs Python and never sees a path. The application opens documents from bytes,
registers other inputs (a template, Markdown, an image) as blobs, sends the definitions to
Claude (or, through the same canonical schemas, OpenAI) and runs the calls:

```python
from ooxml_edit.tools import Toolbox
from docx_agent.tools import FORMAT, GROUPS, TOOLS

with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
    session = toolbox.session(clock=my_clock)          # dates written come from here
    d1 = session.open(docx_bytes, name="agreement.docx")
    tools = toolbox.definitions("anthropic")   # core loaded, the rest by tool search (the default)
    system = toolbox.system_prompt(extra=my_guidance)  # mechanics only; house style is yours
    result = toolbox.dispatch(session, "describe", {"doc": d1})
    for output in session.take_outputs():              # what save_document wrote, as bytes
        store(output.name, output.data)
```

- **The tools:** the Word handlers of the shared tools (`open_document`, `new_document`,
  `save_document`, `undo`, `describe`, `find_text`, `replace_text`, `render`, `check`,
  charts, SmartArt, properties, `close_document`, `read_blob`, `batch`) and 25 Word tools:
  `word_read` (Markdown with ids,
  paged), `word_set_text`, `word_insert_text`, `word_delete`, `word_insert_markdown`,
  `word_format` (every run and paragraph property, in points; or the look of a given paragraph, `like`: its style, list level, indents and keeps), tracking mode, tracked
  changes (list, accept, reject), comments, moving and copying sections, sections, headers
  and footers, fields and the TOC, notes, links, lists, styles, tables (a new one is Markdown,
  through `word_insert_markdown`, or a dataset's or another document's table's rows with its
  `from`, as `word_edit_table`'s `set_cells` takes them), pictures and other drawings, content controls and the
  compatibility mode; a template's design facts (`word_design_facts`; every declared style resolved with `include: ["declared_styles"]`, paged); blocks compared with a sample block, the fixes as calls (`word_compare_blocks`); copies made to follow a revised sample in one call, with a preview and per copy what could not follow, why and its kind (`own`, `changed_by_copy`, `structure`; `report: "actionable"` leaves the copies' own out) (`word_follow_sample`). Core: `describe`, `word_read`, `word_set_text` and the shared core;
  the rest load on demand in five groups. What is supported and what is not:
  [SUPPORTED.md](../src/docx_agent/tools/SUPPORTED.md); for the application's thinking layer:
  [GUIDANCE.md](../src/docx_agent/tools/GUIDANCE.md).
- **Every changing call** is one undo step (all of a `batch` is one: a failing op is skipped
  with its error and the rest apply; `atomic: true` for all or none), written as tracked
  changes while `word_set_tracking` is on, and returns `checks`: the `validate()` delta
  against the document as opened and, once the document was laid out, the pages the change
  reflowed (for 20 pages or fewer; beyond, `check` reports them).
- **`check` (with `reflow`), `render` and `save_document` return `coverage`**: how much
  docx2svg could lay out. `complete` is false when a block, a header or footer, or a face
  could not be laid out. Coverage also gives the stop and the faces substituted with open
  ones. A check on a partial layout is not a passed check.
- **`check` and `render` take `range`** (round 35): one or several blocks, spans
  (`p:A..t:B`), cells or sections; `check` reports the facts about those blocks on the
  pages they fall on, the problems found in them and each range's pages, `render` draws
  those pages; each range of whole blocks says whether it is on one page (`one_page`),
  set to stay so (`kept`, as `word_format` `keep_together` keeps a span), where it breaks
  and the call that keeps it -- with what that call does, laid out (round 40, `if_kept`:
  where the range then starts, how much of the page before it is left blank, the page
  count and the page what follows starts on, now and then); a kept range that starts a page
  because of its keep says what releasing it gives (`if_released`, with that call). Each
  range is tried alone, at most four a check (`effect_left_out` past them). **`check` takes `copies`**: each copy (a block or span) compared with the
  reference it was made from, here or in another open document -- whether it follows,
  what differs with its fix and kind, where the reference's example text is still in it,
  and one `word_follow_sample` call per reference to make the copies follow. Each list item
  and paragraph is also compared as Word applies it (round 42): its list level, its label
  (`list_format`, e.g. `bullet •` beside `bullet –`), where its numbering starts
  (`list_numbering`: `continues`, `starts at 3`), its indents and its keeps and widow
  control, with the `word_lists` or `word_format` call that fixes each (`list_fix`: every
  copy's list calls together); two items that look the same never differ. A comparison
  is kept while neither side's blocks, nor the list items before them, change.
- **`render` takes `fast`** (round 39): while building, a page for about a third of the CPU:
  the same layout and page breaks, the text drawn as outlines in one raster pass, the
  pixels approximate along glyph edges. The result says so (`"approximate": ["pixels"]`,
  `approximate_pages`, labels `page 3 (approximate)`); render without `fast` for the exact
  image (a review's, a final one). A page whose exact image is held is returned exact.
- **Text is spaced as its language wants** (round 35): what a call writes in French gets a
  no-break space before `: ; ! ?` and inside `« »` (measured in Word: an ordinary space lets
  the mark start the next line), and the result says how many; text already there is not
  changed. `check`'s `punctuation_spacing` lists what is left, and the spaces English,
  German, Dutch, Spanish, Italian and Portuguese proofing flags, with `replace_text` fixes.
  `docx_format(typography=False)` turns the writing off.
- **Layout runs in the toolbox's worker pool** through `Document.converter`, under a 30 s
  deadline: past it the worker is killed, the call reports `timeout` and changes nothing.
  The workers start with `spawn`, which imports the program's main module again in each,
  so a script keeps its work under `if __name__ == "__main__":`. Without the guard the
  workers stop as they start and the toolbox lays out in subprocess workers instead, with
  a warning that says why (ooxml-edit 0.58.0); `Toolbox(workers=0)` lays out in the
  script's own process.
  Each worker keeps the layouts it made, with what it measured, for the documents' next
  versions, within a memory budget: `DOCX_AGENT_LAYOUT_MEMORY` (MB, in the environment the
  workers start with; 256 by default) for the layouts held -- about a megabyte a page -- and
  up to 1.75 times it with the measurements and pickled layouts kept beside them, the least
  recently used dropped first (`docx_agent.layout.LAYOUT_MEMORY`, round 39).
- **Saving refuses new validation problems**; only the application may allow them
  (`Toolbox(allow_new_problems=True)`). `save_document` with `clean: true` takes out of
  the file saved what Word keeps but never draws -- orphaned parts (`[trash]/0000.dat`, a
  part nothing reaches), every `w:rsid*` attribute and settings' rsid list, webSettings'
  pasted web divisions (unless one sets a margin or a border), an external
  `attachedTemplate` -- and reports each kind's count and bytes and the file's sizes as
  `cleaned`; the document in the session is unchanged (`docx_agent.edit.cleaning`).
- **Every result's `touched`** (the application's, never in what the model reads) names
  what the call changed: a body change by its **heading section** --
  `{"doc", "part": "word/document.xml", "heading": "p:301", "level": 2, "blocks":
  ["p:301", "t:7"]}`, the innermost heading before each changed block and the first and
  last block of the section it owns (`heading: null, level: 0` before the first heading)
  -- every other part as `{"doc", "part"}` (`docx_agent.tools.touched`). Pages are not
  named: which pages moved is known only from a new layout; render after the change and
  compare the pages' images by content instead.

`tests/goldens/transcripts` hold the end-to-end trial's Word tasks done with the tools alone,
and w10-w13 (a board pack from five reports, a contract form, a picture with alt text, a
legacy document tidied);
`tests/test_tools_goldens.py` replays them to byte-identical outputs that pass the trial's
own checks (where Office's faces are installed; elsewhere it compares every result less what
the layout measures, CONTRIBUTING.md "Golden transcripts"), and `tests/test_oracle_tools.py` has Word open each.
