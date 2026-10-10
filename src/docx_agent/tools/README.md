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
