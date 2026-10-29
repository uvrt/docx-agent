# docx-agent

[![CI](https://github.com/uvrt/docx-agent/actions/workflows/ci.yml/badge.svg)](https://github.com/uvrt/docx-agent/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

An AI-editable Word (`.docx`) layer for Python: inspect a document, change it through a
semantic API (optionally as tracked changes), read and write it through Markdown that
carries stable ids, render the pages back, repeat. The counterpart of
[`pptx-agent`](https://github.com/uvrt/pptx-agent) for Word.

It builds on [`ooxml-edit`](https://github.com/uvrt/ooxml-edit) (the lxml-based OPC
package, ordered insertion and undo, and its charts subpackage) and renders and lays out
through [`docx2svg`](https://github.com/uvrt/docx2svg).

To give a model the document as tool calls (Claude, or OpenAI's Responses and Chat
Completions APIs): [using the tools with a model provider](src/docx_agent/tools/README.md).

## Install

Not on PyPI yet. Python 3.10+. The siblings install from git:

```bash
pip install "ooxml-common @ git+https://github.com/uvrt/ooxml-common@main" \
            "ooxml-edit @ git+https://github.com/uvrt/ooxml-edit@main" \
            "docx2svg[png] @ git+https://github.com/uvrt/docx2svg@main"
pip install "docx-agent @ git+https://github.com/uvrt/docx-agent@main"
```

## Fonts

Rendering and reflow feedback lay the document out with docx2svg, which measures every
glyph with the face Word uses. Install the document's faces where the layout runs:

- **Office installed:** nothing to do.
- **No Office:** install the open metric compatible substitutes (Debian/Ubuntu: `sudo
  apt-get install fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation`;
  macOS: `brew install --cask font-carlito font-caladea font-liberation`). With Carlito in place
  of Calibri, and Liberation Sans, Serif and Mono in place of Arial, Times New Roman and
  Courier New, lines and pages break where they do in Word. Symbol and Wingdings bullets are laid out from
  metrics recorded from Word's copies.
- **Without a face and its substitute** (Aptos, Cambria, Calibri Light...), the layout
  stops at the first paragraph in that face, and everything after it is unknown, never
  guessed.

docx2svg reports every substitution, and never uses a substitute when the real face is
installed. Details and the fidelity to expect:
[docx2svg's README](https://github.com/uvrt/docx2svg#fonts).

`check` (with `reflow`), `render` and `save_document` return the layout's **coverage**,
so a check that passed is told apart from one that could not see everything. It holds:

- `complete`
- the pages laid out, with Word's saved page count when the layout stopped
- the blocks laid out, out of the total
- where the layout stopped, by block id, and header and footer stops (docx2svg no longer
  stops by default: what it cannot lay out it approximates or draws as a placeholder, and
  goes on)
- the placeholders: what was drawn as a marked box of its height instead of laid out, by
  page, reason and block id (any makes `complete` false)
- the approximations: what was laid out by a rule not measured against Word
- the faces substituted or missing

In the library, `doc.layout().coverage_facts()` returns the same.

`check` with `layout` (or `min_size`, `styles`) measures the pages from that same layout --
text below a size by kind, direct formatting over the style with the call that clears it,
styles not in an allowed list, empty paragraphs used as space, keep-with-next paragraphs at
a page's foot, rows split across pages and header rows that do not repeat, pictures and
tables past the margins, and the pages a reviewer cannot judge from the render -- as facts
with block ids and pages; `docx_agent.layout_facts.layout_facts(doc.layout())` in the
library.
A Word toolbox is sent the shared tools with a document's parameters only: `check` has
`pages`, `include` (`validate`, `reflow`, `fields`, `layout`, `app`), `min_size` and
`styles`, and `find_text`, `replace_text` and `render` have no `slides` (ooxml-edit's
`Tool.kind_fields`). A call that still passes a deck's parameter gets the answer it always
got (`render`'s and `check`'s `slides` read as `pages`, with a warning).

Refs (`$name`) live in the tools' session, not the document. Outside a call the application
reads them with `session.refs("d1")` (`{name: address}`, a copy) and finds the top-level
block holding one with `docx_agent.tools.block_of(session, "d1", "$intro")` or the blocks
between two with `blocks_between(session, "d1", "$intro", "$closing")`; a handler passes
its `call` instead (`block_of(call, "$intro")`).

The user's regional settings -- the separator a decimal tab stop aligns on, a chart's
decimal and thousands separators, the language of new chart text in a document that states
none, a new document's language and number separators -- are `Document.locale` (or
`Toolbox(locale=...)`), an `ooxml_common.locales.Locale`; en-US unless given, never the
server's own.

## Example

```python
from docx_agent import Document

doc = Document.open("agreement.docx")
print(doc.to_markdown())                       # every block, its id in a comment: <!-- p:3B212964 -->
print(doc.to_markdown(view="markup"))          # tracked changes {++ins++}{--del--} and comments {>>...<<} in place

scope = doc.paragraph("p:3B212964")
scope.set_text("The supplier will host the customer portal for 36 months.")   # formatting kept
added = scope.insert_after("A new paragraph after it.", style="Normal").object
added.format(italic=True)
doc.undo()                                     # every edit is one undo step
doc.save("agreement-edited.docx")
```

A table or chart placed from a source by the agent tools (`from`) keeps where its values
came from in the saved file, in a custom XML item Word keeps; read it back with
`Document.provenance()`, `{address: record}` (`t:<paraId>` for a table, `d:<id>` for a
chart, each an `ooxml_edit.provenance` record). `ignored` counts the entries whose table or
chart is gone; `Document.store_provenance(address, record)` keeps one (`None` drops it):

```python
doc = Document.open("report.docx")
for address, record in doc.provenance().items():
    print(address, record["source"], record.get("document"))   # t:60D6847F d:checks budget.xlsx
```

More recipes -- find and replace, tracked changes, comments, revisions, Markdown in the
template's own styles, tables, pictures, charts, sections, rendering, templates,
validation -- each run as written by the tests: [docs/common-tasks.md](docs/common-tasks.md).

## What is supported

- **Library:** phases E0 to E6 of [ROADMAP.md](ROADMAP.md), and charts and SmartArt: text
  ranges and anchors, find and replace across runs, styles and direct formatting, lists,
  hyperlinks, bookmarks, cross-references, pictures; Markdown with ids in and out, or JSON;
  tracked changes and comments; sections, headers and footers, notes, fields and tables of
  contents; tables, floating drawings, text boxes, shapes, content controls; new documents
  and templates, blocks copied between documents, compatibility-mode upgrade, properties.
  Each written as Word writes it, tracked where Word tracks. The full account:
  [docs/api-tour.md](docs/api-tour.md).
- **Agent tools** (`docx_agent.tools`, on `ooxml_edit.tools`): the document through tool
  calls for a model; it never runs Python and never sees a path
  ([docs/tools.md](docs/tools.md)). What is supported and what is not:
  [SUPPORTED.md](src/docx_agent/tools/SUPPORTED.md); guidance for the application's
  thinking layer: [GUIDANCE.md](src/docx_agent/tools/GUIDANCE.md).

## Status

Pre-release (0.1.0); the API may still change. Golden transcripts replay the end-to-end
trial's Word tasks with the tools alone to byte-identical outputs. The Microsoft Word
oracle tests are local-only (macOS with Word) and skip elsewhere, including CI.

## Documentation

- [docs/common-tasks.md](docs/common-tasks.md) -- short recipes for what an agent does most
- [docs/api-tour.md](docs/api-tour.md) -- what is covered, and a tour of the API
- [docs/tools.md](docs/tools.md) -- the tool layer for a model
- [tools/README.md](src/docx_agent/tools/README.md), [SUPPORTED.md](src/docx_agent/tools/SUPPORTED.md), [GUIDANCE.md](src/docx_agent/tools/GUIDANCE.md) -- the agent tools
- [ROADMAP.md](ROADMAP.md) -- phases, decisions and what Word was measured to write
- [CHANGELOG.md](CHANGELOG.md), [CONTRIBUTING.md](CONTRIBUTING.md) (running the tests)

## Family

- [pptx2svg](https://github.com/uvrt/pptx2svg) -- renders PowerPoint (`.pptx`) slides to SVG and PNG.
- [docx2svg](https://github.com/uvrt/docx2svg) -- renders Word (`.docx`) documents to SVG, page by page.
- [ooxml-common](https://github.com/uvrt/ooxml-common) -- the format-neutral reading, DrawingML, fonts and text metrics both renderers share.
- [ooxml-edit](https://github.com/uvrt/ooxml-edit) -- lossless, undoable editing of OOXML packages, shared by both agent layers.
- [pptx-agent](https://github.com/uvrt/pptx-agent) -- an AI-editable PowerPoint layer: inspect, edit, re-render.
- [docx-agent](https://github.com/uvrt/docx-agent) (this repo) -- an AI-editable Word layer: inspect, edit (optionally as tracked changes), re-render.

## License

MIT, see [LICENSE](LICENSE). `tests/corpus/commonmark/` (examples of the CommonMark Spec)
is under CC BY-SA 4.0, and the third-party fixtures keep their own terms; each directory's
`PROVENANCE.md` says which.
