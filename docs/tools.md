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
  `word_format` (every run and paragraph property, in points), tracking mode, tracked
  changes (list, accept, reject), comments, moving and copying sections, sections, headers
  and footers, fields and the TOC, notes, links, lists, styles, tables (a new one is Markdown,
  through `word_insert_markdown`, or a dataset's or another document's table's rows with its
  `from`, as `word_edit_table`'s `set_cells` takes them), pictures and other drawings, content controls and the
  compatibility mode; a template's design facts (`word_design_facts`). Core: `describe`, `word_read`, `word_set_text` and the shared core;
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
- **Layout runs in the toolbox's worker pool** through `Document.converter`, under a 30 s
  deadline: past it the worker is killed, the call reports `timeout` and changes nothing.
  The workers start with `spawn`, which imports the program's main module again in each,
  so a script keeps its work under `if __name__ == "__main__":`. Without the guard the
  workers stop as they start and the toolbox lays out in subprocess workers instead, with
  a warning that says why (ooxml-edit 0.58.0); `Toolbox(workers=0)` lays out in the
  script's own process.
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
