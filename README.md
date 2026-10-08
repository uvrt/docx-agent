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

## Install

Not on PyPI yet. Python 3.10+. The siblings install from git:

```bash
pip install "ooxml-common @ git+https://github.com/uvrt/ooxml-common@main" \
            "ooxml-edit @ git+https://github.com/uvrt/ooxml-edit@main" \
            "docx2svg[png] @ git+https://github.com/uvrt/docx2svg@main"
pip install "docx-agent @ git+https://github.com/uvrt/docx-agent@main"
```

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

Pre-release (0.0.1); the API may still change. Golden transcripts replay the end-to-end
trial's Word tasks with the tools alone to byte-identical outputs. The Microsoft Word
oracle tests are local-only (macOS with Word) and skip elsewhere, including CI.

## Documentation

- [docs/common-tasks.md](docs/common-tasks.md) -- short recipes for what an agent does most
- [docs/api-tour.md](docs/api-tour.md) -- what is covered, and a tour of the API
- [docs/tools.md](docs/tools.md) -- the tool layer for a model
- [SUPPORTED.md](src/docx_agent/tools/SUPPORTED.md), [GUIDANCE.md](src/docx_agent/tools/GUIDANCE.md) -- the agent tools
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
