# Using the Word tools with a model provider

The tools in `docx_agent.tools` are defined once, in a provider-neutral canonical schema
([`ooxml_edit.tools`](https://github.com/uvrt/ooxml-edit/blob/main/docs/TOOLS.md)). A
`Toolbox` turns them into each provider's definitions, runs the model's calls with
`toolbox.dispatch(session, name, arguments)`, and turns the results back into that
provider's messages. What the tools do: [SUPPORTED.md](SUPPORTED.md); how to use them in an application:
[GUIDANCE.md](GUIDANCE.md) and [docs/tools.md](../../../docs/tools.md).

| Call | Sends |
| --- | --- |
| `definitions("anthropic")` | every tool, strict on as many as Claude's limits allow (writing tools first); the non-core groups deferred behind the BM25 tool-search tool, `cache_control` on the last core tool |
| `definitions("openai-responses")` | the Responses API (GPT-6): the core as functions, each other group a deferred `namespace`, and `tool_search` |
| `definitions("openai-chat", groups=[...])` | Chat Completions, which has no tool search: the core and the groups named; `toolbox.allowed_tools([...], provider="openai-chat")` narrows a turn's calls without changing `tools` |

**Status.** This repository's own model trials ran on Claude. A production user ran these
tools live on Azure OpenAI's Responses API (October 2026: the default
`definitions("openai-responses")`, `store=False` with
`include=["reasoning.encrypted_content"]`, images in the outputs and in a user message), and
every task completed with the loop below as written. `items += response.output` passes the
reasoning items (with their `encrypted_content`) and a hosted tool search's
`tool_search_call` / `tool_search_output` back as they came; only `function_call` items are
dispatched, by their bare `name` (a deferred tool's group arrives separately, as
`namespace`). The OpenAI definitions and results are also checked offline against OpenAI's
documented rules (`ooxml_edit.tools.adapters.openai_problems`, `openai_input_problems`, and
ooxml-edit's Responses round-trip test). Chat Completions has not been run live.

## Groups

| Group | Tools |
| --- | --- |
| `core` (always loaded) | `open_document`, `new_document`, `save_document`, `undo`, `find_text`, `replace_text`, `render`, `check`, `describe`, `word_read`, `word_set_text`, `batch` |
| `shared_misc` | `close_document`, `read_blob`, `set_properties`, `edit_chart`, `edit_smartart` |
| `word_text` | `word_inspect`, `word_insert_text`, `word_delete`, `word_insert_markdown` |
| `word_review` | `word_set_tracking`, `word_changes`, `word_comments` |
| `word_structure` | `word_move`, `word_copy_from`, `word_sections`, `word_headers_footers`, `word_fields`, `word_notes`, `word_links` |
| `word_objects` | `word_edit_table`, `word_format_table`, `word_drawings`, `word_controls` |
| `word_style` | `word_format`, `word_lists`, `word_styles`, `word_template` |

## Setup (every provider)

```python
from ooxml_edit.tools import Toolbox
from docx_agent.tools import FORMAT, GROUPS, TOOLS

toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)   # toolbox.close() when done
session = toolbox.session()                                 # clock=... sets the dates written
session.open(docx_bytes, name="agreement.docx")             # the model calls it "d1"
system = toolbox.system_prompt(extra=HOUSE_RULES)          # your guidance after the tools'
task = "In d1, change the term to 36 months, tracked, then save it as agreement.docx."
```

**Your own fonts.** Faces kept in a folder of the application's own (licensed fonts the
system does not search) are named once, as configuration -- never by the model:

```python
toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, font_dirs=["/srv/app/fonts"])
session = toolbox.session()                          # the toolbox's folders
other = toolbox.session(font_dirs=["/srv/b/fonts"])  # or a session's own
```

Every layout and render -- the reflow in each edit's checks and in `check`, the
`coverage` that `render`, `check` and `save_document` report, field and TOC page numbers,
the PNGs -- then finds those faces, in the worker process too (the folders are resolved
where the toolbox runs and handed over). They are added to the folders Word uses and
searched after them. Without `font_dirs` the environment variable `OOXML_FONT_DIRS`
(folders separated by `os.pathsep`) is read; `font_dirs=[]` reads neither. No tool
definition or prompt changes. A document opened outside the tools takes the same setting
as `Document.font_dirs`.

**The user's locale.** The regional settings a document is drawn and edited in -- an
`ooxml_common.locales.Locale` -- are configuration too, per toolbox or per session:

```python
from ooxml_common.locales import Locale

dutch = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-NL", list_separator=";")
toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, locale=dutch)
other = toolbox.session(locale=None)                 # en-US
```

Every layout and render (in the worker process too, where the locale is handed over and
keys the render cache) passes it to docx2svg -- the separator a decimal tab stop aligns
on, a chart's decimal and thousands separators -- and the document edits in it: a new
chart's language where the document states none, and `new_document`'s language and number
separators (`w:decimalSymbol`, `w:listSeparator`). Without one it is en-US, never the
server's own. No tool definition or prompt changes. A document outside the tools takes it
as `Document.locale` (or `Document.new(locale=...)`).

### Tables from data

`word_insert_markdown` (in place of `markdown` or `blob`) and `word_edit_table`'s `set_cells`
take a table's text from a source, as `edit_chart` takes a chart's (ooxml-edit 0.22.0):
`from` is a dataset (`d:<id>`) or a table of another open document (`t:<name>`),
`from_columns` picks its columns by name, in the order placed (default: all), and every row
is placed. A header row of the column names goes above them -- repeated on each page, as a
Markdown table's is -- unless `header` gives its texts, or `[]` for none. Numbers are written
as a chart's category labels write them (General, the session locale's decimal separator),
dates as the source gives them, a blank as an empty cell, a value's line breaks as line
breaks in the cell's one paragraph; the text goes in literally, never read as Markdown.
`word_insert_markdown` makes the table the source's size at `at`, in the document's table
style (`style_map`'s `table` and `table_cell` apply); `set_cells` writes from row `index`
(default 0), column 0, and refuses a source the table has no room for, or a block over a
merged cell. Nothing is cut: past 63 columns, 4,096 rows or 5,000 cells the call is refused
(`limit`). A whole call is one undo step, and tracked while tracking is on.

```json
{"doc": "d1", "at": "after:p:3B212964", "from": "d:checks",
 "from_columns": ["check", "rating", "due"]}
```

The result's `data` has the table (`id`, `rows`, `columns`; not its Markdown, which the
caller has) and `from`, where
every placed value came from -- the source, its document and version, `header_row` and
`first_row` (table rows from 0; source row `n`, from 1, is table row `first_row + n - 1`),
and per column its table column, source column, source rows as runs and their sources:

```json
{"source": "d:checks", "doc": "d2", "version": 3, "header_row": 0, "first_row": 1,
 "columns": [{"col": 1, "column": "rating", "rows": "1-12",
              "sources": ["Checks!C2:C13"]}, ...]}
```

It is not written into the document. A Word document names its drawings with the same
`d:` prefix as datasets (`d:3`; a dataset's id starts with a letter), so a `from` only
ever names a dataset, and given a drawing's id it says so: `no dataset d:3; d:3 is a
drawing in this document, not a dataset: ...`.

The provider SDKs read their keys from the environment (`ANTHROPIC_API_KEY`,
`OPENAI_API_KEY`). When the loop ends, `session.take_outputs()` holds what `save_document`
wrote, as bytes.

## Claude (Messages API)

```python
import anthropic

client = anthropic.Anthropic()
tools = toolbox.definitions("anthropic")
messages = [{"role": "user", "content": task}]
while True:
    response = client.messages.create(model="claude-sonnet-5-5", max_tokens=8000,
                                      system=system, tools=tools, messages=messages)
    messages.append({"role": "assistant", "content": response.content})
    calls = [block for block in response.content if block.type == "tool_use"]
    if not calls:
        break
    results = [(c.id, toolbox.dispatch(session, c.name, c.input)) for c in calls]
    messages.append(toolbox.render_results("anthropic", results))   # tool_result blocks
```

## OpenAI (Responses API)

```python
from openai import OpenAI

client = OpenAI()
tools = toolbox.definitions("openai-responses")
items = [{"role": "user", "content": task}]
while True:
    response = client.responses.create(model="gpt-6", instructions=system, tools=tools,
                                       input=items)
    items += response.output
    calls = [item for item in response.output if item.type == "function_call"]
    if not calls:
        break
    results = [(c.call_id, toolbox.dispatch(session, c.name, c.arguments)) for c in calls]
    items += toolbox.render_results("openai-responses", results)    # function_call_output items
```

## OpenAI (Chat Completions)

```python
from openai import OpenAI

client = OpenAI()
tools = toolbox.definitions("openai-chat", groups=["word_text", "word_review"])
messages = [{"role": "system", "content": system}, {"role": "user", "content": task}]
while True:
    reply = client.chat.completions.create(model="gpt-6", tools=tools,
                                           messages=messages).choices[0].message
    messages.append(reply)
    if not reply.tool_calls:
        break
    results = [(c.id, toolbox.dispatch(session, c.function.name, c.function.arguments))
               for c in reply.tool_calls]
    messages += toolbox.render_results("openai-chat", results)
```

A Chat Completions tool message is text only, so `render_results` follows the tool
messages with one user message carrying the images (a `render`'s pages), each labelled with
its call. `dispatch` takes the arguments as a dict or as the JSON string OpenAI sends, and
never raises for the model's mistakes: an error comes back as a result the model can read,
naming the valid options where there are some.
