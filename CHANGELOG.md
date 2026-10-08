# Changelog

docx-agent has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail of every phase.

## Unreleased

- `check` (with `reflow`), `render` and `save_document` return `coverage`, from docx2svg's
  coverage summary: whether the layout is complete, the blocks laid out, where it stopped
  (by id), header and footer stops, and the faces substituted with open ones or missing.
  "Check passed" can now be told apart from "couldn't check". In the library:
  `DocumentLayout.coverage` and `coverage_facts()`. Needs docx2svg with
  `docx2svg.coverage`.
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
