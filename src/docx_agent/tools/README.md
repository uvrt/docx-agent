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
| `word_style` | `word_format`, `word_lists`, `word_styles`, `word_template`, `word_design_facts`, `word_compare_blocks` |

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

**Roles: a reviewer's tools.** A toolbox or session loads roles, additive toolsets
(ooxml-edit 0.45.0, [TOOLS.md, "Roles"](https://github.com/uvrt/ooxml-edit/blob/main/docs/TOOLS.md)):
`review` reads, checks, compares, renders and adds and replies to comments; `resolve`
resolves, reopens and deletes them; `accept` accepts and rejects tracked changes; `edit`
makes every other change. Editing a comment's text is every role's. The default is all
four, the toolbox as without roles.

```python
reviewer = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, roles=["review"])
tools = reviewer.definitions("anthropic")       # 22 tools; word_comments list/add/reply/edit
system = reviewer.system_prompt(extra=HOUSE_RULES)
session = reviewer.session()                    # word_set_text here is refused
second_pass = toolbox.session(roles=["review", "resolve"])   # per session, same toolbox
chief = toolbox.session(roles=["review", "accept"])  # settles tracked changes, nothing else
```

A reviewer's core is open, save, `describe`, `word_read`, `find_text`, `render`, `check`,
`batch`, `word_comments` and `word_changes` (2,350 estimated tokens against the full core's
2,570); the mixed tools offer only `list` (`word_styles` `describe` too), `edit_chart`
only `read`; accepting or rejecting tracked changes is `accept`'s. A call outside the session's roles --
in a `batch` too, op by op -- is refused naming the role it needs; a reviewer reports each
finding with the call that would fix it, for a builder (`roles=["review", "edit"]`, or the
default) to apply in one batch.

Several agents on one document at once share one session -- the document, its undo
history, refs and outputs -- each through a view with roles of its own (ooxml-edit
0.46.0); a view narrows the session's roles, never widens them, and its `context` reaches
the handlers of the application's own tools as `call.caller`:

```python
session = toolbox.session()                     # every role
builder = session.view(roles=["review", "edit"], context={"agent": "builder-1"})
chief = session.view(roles=["review", "accept"], context={"agent": "chief-1"})
toolbox.dispatch(chief, "word_set_text", args)  # refused: needs the edit role
sent = toolbox.definitions("anthropic", roles=["review", "accept"])
system = toolbox.system_prompt(roles=["review", "accept"], tools=sent)  # follows `sent`
```

### The application's settings: an asset catalogue, minimum text sizes

`docx_format(...)` is `FORMAT` with settings of the application's own -- never tool
arguments, so the definitions stay the same unless a setting adds a tool (pptx-agent's
`deck_format` takes the same two):

```python
from ooxml_edit.tools import FolderCatalogue
from docx_agent.tools import docx_format

icons = FolderCatalogue("brand-icons")             # brand-icons/index.json: id, file, name, keywords
fmt = docx_format(assets=icons, min_text_size={"body": 10, "table": 9, "footnote": 8, "caption": 9})
toolbox = Toolbox(TOOLS, formats=[fmt], groups=GROUPS)
```

- `assets` (ooxml-edit's catalogue protocol, `search` and `get`): `word_find_asset` finds
  ids by words, and `word_drawings` `insert_picture` takes `asset` in place of `image` --
  inline, or floating with `wrap`; `at` a position or a cell (`t:1A/c0,0`); `size_pt` or
  `width`; `color` a theme colour (`accent2`) or `RRGGBB` to tint an SVG, as Word draws a
  recoloured SVG (its theme class kept, so the icon follows the document's theme).
- `min_text_size`: one number, or points per kind (`body`, `table`, `footnote`,
  `endnote`, `header`, `footer`, `caption`; a kind left out is not checked). `check`
  reports `small_text` against it, and `word_insert_markdown` and `word_edit_table`
  `set_cells` report the text they wrote that its styles make smaller.

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
(`limit`); a refusal for want of rows offers the row range of the source that fits. To
place some of a source's rows, name them: `d:checks/#2:` (row 2 to the last),
`t:Sales/#1:#9` (a workbook's rows 1 to 9). Word flows a long table over pages itself,
the header row repeated. A whole call is one undo step, and tracked while tracking is on.

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

Its record -- `from` without the session's document handle, with the source document's
name (ooxml-edit 0.25.0's `ooxml_edit.provenance`) -- is also kept in the document: see
"Source addresses in the saved file" below. A Word document names its drawings with the same
`d:` prefix as datasets (`d:3`; a dataset's id starts with a letter), so a `from` only
ever names a dataset, and given a drawing's id it says so: `no dataset d:3; d:3 is a
drawing in this document, not a dataset: ...`.

### Source addresses in the saved file

Where a table's or a chart's values came from is kept in the document itself, so it can be
read back after a save, in another session, or after Word has opened and saved the file
(measured, Word for Mac, 2026-10-14). It is one custom XML item related from the document
part -- `customXml/itemN.xml`, its `itemPropsN.xml` naming the schema
`urn:uvrt:docx-agent:provenance:v1`, and a `customXml` relationship from
`word/document.xml` -- with one entry per table or chart, keyed by the address the tools
take: a table's `t:<paraId>` (its first row's `w14:paraId`, which Word keeps because
docx-agent stamps every paragraph and row) and a chart's `d:<docPr id>`. Each holds the
compact JSON record:

```xml
<prov:sources xmlns:prov="urn:uvrt:docx-agent:provenance:v1"><prov:source kind="table"
 key="t:60D6847F">{"v":1,"source":"d:checks","header_row":0,"first_row":1,"columns":[...]}</prov:source>
<prov:source kind="chart" key="d:2">{"v":1,"source":"d:checks","categories":{...},"series":[...]}</prov:source></prov:sources>
```

Other custom XML items (a bibliography, cover-page properties, a content control's data)
are left as they are. `word_insert_markdown`'s and `set_cells`' `from`, and `edit_chart`'s,
write the record in the call's undo step; the last one dropped takes the item, its
properties and the relationship away. `session.provenance(doc)` (or `Document.provenance()`)
reads `{address: record}`, for the tables and charts the document has with its tracked
changes accepted. No tool or parameter is added.

The record never says what is no longer true. A record's rows count the table's rows with
its tracked changes accepted; an edit without `from` keeps it true or drops it:

| Edit | The table's record |
| --- | --- |
| `set_cells` (no `from`) over the block, header row included | that column's entry dropped |
| `insert_row`/`delete_row` above the block / below it / inside it | shifted / kept / dropped |
| `insert_column`; `delete_column` | entries right of it shifted; its entry dropped, the rest shifted |
| `merge`, `split` | dropped |
| a row inserted or deleted at the top (the table's id changes) | moved to the new id |
| `word_changes` accept | kept (under the new id if accepting renamed the table) |
| `word_changes` reject of anything in the table | dropped: the old values may be back |
| the table or chart deleted (`word_delete`, a replace) | dropped in that call's undo step |
| copied (`word_copy_from`) | the copy has none: it gets new ids, never a recorded one |

A record left with no column entry is dropped. A chart's record follows `edit_chart`'s own
rules (ooxml-edit's docs/TOOLS.md). Text changed in a recorded table by the text tools
(`word_set_text`, `replace_text`...) is not followed: the record says where the values were
placed from. A table deleted as a tracked change keeps its entry, left out on reading; a
save prunes nothing; an entry whose object is gone (another program deleted it) is left
out on reading (`Document.provenance().ignored` counts them) and pruned by the next edit.

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
