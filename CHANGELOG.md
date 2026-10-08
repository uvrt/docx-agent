# Changelog

docx-agent has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail of every phase.

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
