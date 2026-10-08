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
  charts, SmartArt, properties, `close_document`, `read_blob`, `batch`) and 24 Word tools:
  `word_read` (Markdown with ids,
  paged), `word_set_text`, `word_insert_text`, `word_delete`, `word_insert_markdown`,
  `word_format` (every run and paragraph property, in points), tracking mode, tracked
  changes (list, accept, reject), comments, moving and copying sections, sections, headers
  and footers, fields and the TOC, notes, links, lists, styles, tables (a new one is Markdown,
  through `word_insert_markdown`), pictures and other drawings, content controls and the
  compatibility mode. Core: `describe`, `word_read`, `word_set_text` and the shared core;
  the rest load on demand in five groups. What is supported and what is not:
  [SUPPORTED.md](../src/docx_agent/tools/SUPPORTED.md); for the application's thinking layer:
  [GUIDANCE.md](../src/docx_agent/tools/GUIDANCE.md).
- **Every changing call** is one undo step (all of a `batch` is one), written as tracked
  changes while `word_set_tracking` is on, and returns `checks`: the `validate()` delta
  against the document as opened and, once the document was laid out, the pages the change
  reflowed (for 20 pages or fewer; beyond, `check` reports them).
- **Layout runs in the toolbox's worker pool** through `Document.converter`, under a 30 s
  deadline: past it the worker is killed, the call reports `timeout` and changes nothing.
- **Saving refuses new validation problems**; only the application may allow them
  (`Toolbox(allow_new_problems=True)`).

`tests/goldens/transcripts` hold the end-to-end trial's Word tasks done with the tools alone,
and w10-w13 (a board pack from five reports, a contract form, a picture with alt text, a
legacy document tidied);
`tests/test_tools_goldens.py` replays them to byte-identical outputs that pass the trial's
own checks, and `tests/test_oracle_tools.py` has Word open each.
