# Roadmap

Working plan for `docx-agent`: an AI-editable Word layer, the counterpart of `pptx-agent`.
Written to be picked up cold: every phase says what is missing, where it goes, how to
approach it, what must be measured before it, and how to know it is done.

**Effort key:** S ≈ half a day · M ≈ 1–3 days · L ≈ 1–2 weeks · XL ≈ 3+ weeks.

The loop is the sibling's. The agent never writes XML. It reads the document (as Markdown
with ids, or as structured JSON), changes it through a semantic API, looks at the rendered
pages, and repeats. Two things make Word harder than PowerPoint, and most of this file is
about them:

1. **A document is a flow, not a set of canvases.** An edit on page 2 can move every line
   after it. The agent needs to be told where things landed, not just that the edit
   succeeded.
2. **Word documents are negotiated.** People exchange them with tracked changes and
   comments, so an agent's edit often has to be a *proposal* (a revision with an author and
   a date), not a silent rewrite.

---

## Decisions already made

These are settled and are not reopened anywhere below.

- **Modern files only.** Every document docx-agent *creates* is a modern Word file in
  compatibility mode 15 (`w:compatSetting w:name="compatibilityMode" w:val="15"`). It never
  writes the binary `.doc` format. An existing document keeps its own compatibility mode,
  because upgrading reflows the layout; `Document.upgrade_to_modern()` is an explicit,
  opt-in operation (E6), and its reflow is reported like any other edit's.
- **Tracked changes are in scope.** Any edit can be written as a revision (`w:ins`,
  `w:del`, `w:rPrChange`, `w:pPrChange`, moves...) with an author and a date, and there are
  APIs to list, accept and reject revisions. Comments belong with them (E3).
- **Markdown is an authoring and reading layer, never the source of truth.** The `.docx`
  is the truth. `insert_markdown(at, md)` maps Markdown onto the document's *own* styles;
  `to_markdown(range)` is a readable projection in which every block carries its stable id;
  refinement goes through the semantic API (E2).
- **Rendering goes through docx2svg**, and the agent gets reflow feedback from its layout:
  where a paragraph landed, before/after renders, which pages changed (E0, E1).
- **The oracle is Word itself.** An edited document must open in Word and export to PDF
  without a repair prompt. Word is one instance per machine, which constrains how the
  oracle tests run (see *Testing strategy*).

---

## Prior art, and what to take from it

Surveyed October 2026. The question for each: how does it represent the document, how does
it address things, what does it do with tracked changes, and what does it lose?

### Libraries that edit `.docx`

**python-docx** (Python, MIT) — <https://python-docx.readthedocs.io/>

The closest relative in design: its `Paragraph`, `Run` and `Table` objects are proxies over
lxml elements (`CT_P`, `CT_R`), not an owned model, so unknown markup survives a save.
Version 1.2.0 (June 2025) added comments (`Document.add_comment`,
`Run.mark_comment_range`), without replies or resolution, and refusing comments in headers
and footers ([docs](https://python-docx.readthedocs.io/en/latest/user/comments.html)). It
has no tracked-change support, and no notion of a stable address: a paragraph is "the
element you happen to hold".

- **Take:** live proxies over lxml (the same conclusion pptx-agent reached
  independently); its pragmatic reading of `w:t`, `w:tab`, `w:br` into Python text.
- **Avoid:** text setters that replace a paragraph's runs wholesale (`paragraph.text = ...`
  drops mixed formatting); no addressing that survives an edit; editing that ignores
  `w:ins`/`w:del`, which several MCP servers below cite as the reason they did not build on
  it ("silently drops tracked changes").

**Open XML SDK** (C#, MIT) and **Open-XML-PowerTools** (C#, MIT) —
<https://github.com/dotnet/Open-XML-SDK>, <https://github.com/OpenXmlDev/Open-Xml-PowerTools>

The SDK is a strongly typed DOM generated from the schemas, which keeps unknown elements as
`OpenXmlUnknownElement` and processes Markup Compatibility. Its issue tracker is the best
public record of what Word does with `w14:paraId`: there is no generator for valid ids
([#962](https://github.com/dotnet/Open-XML-SDK/issues/962)), and a reporter saw Word
**replace SDK-generated paraIds on the first save** after adding a comment, after which they
stayed stable ([#925](https://github.com/dotnet/Open-XML-SDK/issues/925), unresolved).

PowerTools is where the hard algorithms already exist, in MIT code: `RevisionAccepter`
(accept all), `RevisionProcessor` (accept *and* reject), `OpenXmlRegex` (search and replace
across run boundaries), `DocumentBuilder` (assemble documents from pieces of others,
importing styles, numbering, comments, footnotes and images), and `WmlComparer` (diff two
documents into revision markup).

- **Take:** PowerTools' algorithms as the *reference* for E1's find/replace, E3's
  accept/reject and E6's cross-document copy — read for their edge cases (a deleted
  paragraph mark, moves with range markers, fields split across revisions), then written
  here against our own tests. Issue #925 as the reason paraId durability must be
  *measured* in E0 before it is relied on.
- **Avoid:** a generated typed DOM; PowerTools' habit of normalising a document (merging
  runs, stripping rsids) before processing it, which makes every untouched part change.

**docx4j** (Java, Apache-2.0) — <https://github.com/plutext/docx4j>

JAXB-based: every part is unmarshalled into a typed object tree and marshalled back on
save. Strong at content-control data binding (OpenDoPE conventions: conditionals, repeats,
XHTML import) and document merging (MergeDocx, commercial). Its forum records it generating
invalid paraIds/textIds.

- **Take:** the data-binding model — a content control bound with `w:dataBinding` to a
  custom XML part is filled by writing the *XML part*, because Word re-populates the
  control from it on open. E5 must respect that even though binding is not authored here.
- **Avoid:** marshal/unmarshal round trips (the owned-model weakness pptx-agent's roadmap
  describes for pptx-svg).

**Aspose.Words** (commercial, closed) — <https://reference.aspose.com/words/net/aspose.words.layout/>

The most complete API, and the one that already answers this project's reflow question:
`LayoutCollector.GetStartPageIndex(node)` gives the 1-based page a node starts on, and
`LayoutEnumerator` walks the page layout (page, rectangle, text of each entity). It also has
`Range.Replace` across runs, `StartTrackRevisions(author, date)` so ordinary edits become
revisions, `Revisions.AcceptAll/RejectAll`, `Document.Compare`, and Markdown load/save.

- **Take:** the API *shape* for E1's reflow feedback (`page_of(id)`, an enumerator of
  layout entities keyed to document nodes) and E3's tracking mode (a mode the ordinary
  edits consult, not a parallel set of "tracked" methods).
- **Avoid:** its fidelity claims are its own; here, layout comes from docx2svg, which is
  measured against Word glyph by glyph and *stops* rather than guess.

**LibreOffice UNO** (MPL) — <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Text/Redline>

A full word processor driven over a bridge: `XTextCursor` editing, the `RecordChanges`
property for tracking, `XRedlinesSupplier` to enumerate redlines. It is an independent
layout engine, so it can say what page a paragraph is on.

- **Take:** nothing structural; it is a useful *secondary* reader in tests (does a
  non-Microsoft consumer still open the file?).
- **Avoid:** as an editing path. Importing into Writer and exporting `.docx` again rewrites
  the whole package in LibreOffice's dialect; a round trip through it is not lossless.

### Converters

**Pandoc** (Haskell, GPL-2.0+) — <https://pandoc.org/MANUAL.html>

The reference Markdown↔docx mapping. Its writer styles output from a `--reference-doc`,
using named paragraph styles (Normal, Body Text, First Paragraph, Compact, Title, Subtitle,
Heading 1–9, Block Text, Source Code, Footnote Text, Definition Term, Definition, Caption,
Table Caption, Image Caption, TOC Heading...), character styles (Verbatim Char, Hyperlink,
Footnote Reference...) and one table style, "Table". Its reader's `--track-changes` is
`accept` (default), `reject` or `all` (insertions, deletions and comments wrapped in spans
with classes), and `docx+styles` keeps each paragraph's and run's style as a
`custom-style` attribute.

- **Take:** the idea that Markdown elements map to *named styles* the document defines
  (E2's style map), and its names as the fallback vocabulary for things Word has no
  built-in style for (code); the three revision views (accept, reject, all) for E2's
  projection.
- **Avoid:** its writer produces a *new* document; it cannot insert into an existing one,
  its tight lists use a "Compact" style that is not Word's, and it discards page geometry
  by design. Not a dependency (and its GPL data files, the default reference.docx among
  them, are not copied here).

**mammoth** (BSD-2-Clause) — <https://github.com/mwilliamson/python-mammoth>

Semantic docx→HTML with an explicit style map (`p[style-name='Heading 1'] => h1`). Its own
Markdown output is deprecated in favour of HTML then a separate converter.

- **Take:** the style map as a user-overridable, declarative table — E2's `StyleMap` reads
  in both directions with the same table.
- **Avoid:** projection without ids; mammoth's output cannot be mapped back to the document.

### AI-editing tools

**Microsoft Copilot in Word.** Agent Mode reached general availability on 22 April 2026
and edits documents directly, multi-step; Copilot can record its edits as tracked changes
at word level, and the Legal Agent works inside documents that already have revisions,
"separating prior revisions from its own new proposals"
([MC1384426](https://mc.merill.net/message/MC1384426),
[windowsforum](https://windowsforum.com/threads/copilot-in-word-gets-governed-track-changes-for-trustworthy-contract-edits.411092/)).
([GA date](https://pasqualepillitteri.it/en/news/1401/microsoft-copilot-agent-mode-word-excel-powerpoint-april-2026)). There is no public API to it; what it confirms is the product shape: **edits as reviewable
revisions, attributed, on top of other people's revisions.**

**Office.js, the Word JavaScript API** — <https://learn.microsoft.com/en-us/javascript/api/word/word.trackedchange>

The public API for code running inside Word. WordApi 1.4 added bookmarks, change tracking,
comments and fields, and `getReviewedText("Original" | "Current")`; WordApi 1.6 added
`TrackedChange` (`author`, `date`, `type`, `accept()`, `reject()`, `getRange()`),
`acceptAll`/`rejectAll` on collections, `Document.changeTrackingMode` (`Off`, `TrackAll`,
`TrackMineOnly`) and `Paragraph.uniqueLocalId` with
`Document.getParagraphByUniqueLocalId` — an id that is explicitly *per session* and differs
between coauthors. **There is no `insertMarkdown`**: neither the released sets nor the
preview list has one (checked October 2026); content goes in as text, HTML or OOXML, and an
archived sample (`Word-Add-in-MarkdownConversion`) converts Markdown by hand with the
paragraph, list and table APIs. An open issue asks for a range that ignores deleted text
([office-js#5874](https://github.com/officedev/office-js/issues/5874)).

- **Take:** the revision vocabulary (author, date, type; accept/reject one, all, or a
  collection) and the "Original"/"Current" text views; `TrackMineOnly` as a filter.
- **Note:** Microsoft's own paragraph id is session-scoped. That is honest about the problem
  E0 has to solve, and a hint that `w14:paraId` alone may not be durable (see *Addressing*).

**docx MCP servers and agent skills.** A crowded field, all of the same family — unzip,
patch XML directly (not python-docx), re-zip:

| Project | Addressing | Revisions | What to note |
| --- | --- | --- | --- |
| [DocxEngine](https://github.com/ruwadgroup/docxengine) (Apache-2.0) | stable paragraph ids (`P4#d4e5`) in Markdown-like text | `w:ins`/`w:del`, accept/reject by author or date | coalesces runs; the same "Markdown with ids" reading layer as E2 |
| [Adeu](https://github.com/dealfluence/adeu) (MIT) | text matching, ids for changes and comments | edits projected back as tracked changes | docx→Markdown for reading, edits as redlines; blocks ambiguous matches |
| [docx-mcp](https://github.com/SecurityRonin/docx-mcp) (MIT) | `w14:paraId` | insert/delete/replace tracked, accept/reject by author | `validate_paraids`, footnote and bookmark audits |
| [docx-mcp-server](https://github.com/knorq-ai/docx-mcp-server) | paragraph index | `w:ins`/`w:del` with author and timestamp | 40 tools |
| [word-mcp-semantic](https://github.com/LePhilippeDucTai/word-mcp-semantic) | — | tracked, per-action undo | edits a document open in Word, or the package at rest |
| Anthropic's [docx skill](https://github.com/anthropics/skills/blob/main/skills/docx/SKILL.md) | raw XML | hand-written `w:ins`/`w:del` "redlining" | reads through pandoc; validates every edit is tracked under the given author |

- **Take:** ids in the reading projection (DocxEngine, docx-mcp); refusing an ambiguous text
  anchor rather than picking one (Adeu); validation as a tool, not a hope (docx-mcp's
  paraId and footnote checks); per-action undo.
- **Improve on:** none of them renders, so none can say where an edit landed or that it
  pushed a table onto the next page; none states a byte-level round-trip guarantee; and
  most address paragraphs by position or by paraId without dealing with paraIds that are
  missing, duplicated or rewritten by Word.

### What follows from the survey

The field has every *piece*: lxml proxies (python-docx), revision algorithms (PowerTools),
layout feedback (Aspose), Markdown style maps (Pandoc, mammoth), ids in the projection
(DocxEngine). What no single open tool has is the combination pptx-agent already has for
slides — **lossless by construction, addressable across edits, undoable byte for byte,
and rendered against the real application** — extended with revisions and reflow feedback.

---

## Where we are

**Phases E0 to E6 are done** (2026-10-03 and -04; see *Phase E0* to *Phase E6* for what
landed and what was measured). E0: the package and its lossless round trip, live views (stories,
paragraphs, runs, tables as grids, sections), ids with lazy stamping and aliases, undo,
`set_text`/`insert_paragraph`/`delete_block`, the validity checks, rendering and reflow
feedback through docx2svg (one conversion per state, since docx2svg's `convert_docx`), the
Word oracle under a machine-wide lock, and the paraId measurement. E1: the semantic API --
text ranges and anchors, find and replace across runs, styles by name (built-ins written as
Word writes them, measured) and direct formatting, effective values through docx2svg's
resolver, lists and numbering, hyperlinks, bookmarks, cross-references, inline pictures,
`coalesce_runs`, `move_block` -- held to every fixture, a randomised model and Word. E2's
read half: `to_markdown` with ids in the final, original and markup views, read back by
markdown-it-py against the model it was written from; the read-only JSON state (schema
`docx-agent/state`, version 1); read-only views of notes, comments, revisions, content
controls and drawings, so every id either writes resolves through `Document.get`. E3:
tracking mode over every primitive (`doc.tracking(...)`, `track=`), written as Word 16.106
writes each edit (measured by `tools/e3_probe.py`); listing, accepting and rejecting
revisions, held to accept-all = the untracked edit and reject-all = the original on every
fixture and to Word's own Accept All and Reject All; comments with replies and resolution
in the modern parts; table rows, columns and merges, tracked in the forms Word accepts;
Word's `w:trackRevisions`. List numbers are now docx2svg's measured counter. E2's write
half: `insert_markdown` writes CommonMark and GFM as the document's own styles, resolved by
name through the same `StyleMap` the reader reads (a missing one added as Word writes it,
Table Grid newly measured), lists as instances of their own, tables as Word makes them,
real footnotes, pictures from what the caller gives; one undo step, one revision group when
tracked; held to the CommonMark AST round trip over the spec's examples and hand-written
documents in a blank document and every fixture, and to Word. E4: sections (breaks
inserted and removed, page setup, columns, line numbering, vertical alignment, title page,
page numbering), headers and footers (default, first, even; linked, unlinked), footnotes
and endnotes (inserted, edited, moved, deleted, numbered per section), fields (`PAGE`,
`NUMPAGES`, `SECTIONPAGES`, `PAGEREF`, `REF`, `SEQ` and captions, `DATE`, hyperlinks) and
tables of contents, every result computed from docx2svg's layout and never left for Word to
update -- written as Word 16.106 writes each (`tools/e4_probe.py`), tracked where Word
tracks, and held to Word: a TOC's and page references' numbers agree with Word's own update,
143 of 143. E5: tables (made as Word makes them, every table, row, column and cell
property, merges grown to whole cells, splits, rows and columns across spans and merges,
nested and floating tables, the grid held consistent by a randomised sweep), drawings (a
picture floated and put inline again, every wrap type, position, order, lock and overlap,
text boxes whose paragraphs the paragraph API edits with their VML fallback kept in step,
shapes of every preset, groups moved and resized) and content controls (each kind read,
filled, inserted and removed; a bound one's custom XML node written with its fill) --
written as Word 16.106 writes each (`tools/e5_probe.py`), tracked where Word tracks, and
decision 7's structural edits in forms **Word's own Accept All and Reject All give back
exactly**, the merge across E3 recorded as a difference among them. E6: authoring --
`Document.new()` writes every part of Word's own new blank document from measured values
(part for part the same, its locale's page and measurements as parameters, A4/metric by
default), or makes one from a template as File > New does (the body kept unless asked, a
`.dotm`'s macros dropped with a warning, no template attached unless asked: Word for Mac
asks for access to an attached one outside its sandbox); `save_as_template`;
`copy_blocks` copies blocks from another document as Word pastes them (styles by name with
Use Destination Styles, Keep Source Formatting and Merge Formatting, lists apart, media
shared, notes and comments renumbered, bookmarks renamed, colliding ids re-issued,
tracked as Word tracks a paste); `upgrade_to_modern()` changes `w:compat` as Convert does
(measured over all 65 legacy options) and reports the reflow; core properties and a kept-up
`app.xml` -- measured by `tools/e6_probe.py`, held to Word, and the new document added to
the corpus every phase's gates run on. Charts and SmartArt (decision 11, unblocked by
ooxml-edit 0.2's `ooxml_edit.charts`): every chart's series, categories, values, titles,
axis titles and legend and every diagram's nodes, edited through the shared subpackage
wherever Word holds them -- the body, headers, footers, notes, text boxes, groups -- each
edit one undo step, the caches and the embedded workbook in step, read in `to_markdown`
and `state()`, measured on Word first (`tools/charts_probe.py`) and held to it. Every phase
of this roadmap is done; what is left is in each phase's proposals.

What it builds on:

| Piece | State | Used for |
| --- | --- | --- |
| `ooxml-edit` (uvrt/ooxml-edit, public, lxml) | **Extracted** from pptx-agent's `core/` with its history: `ooxml_edit.opc`, `.xml`, `.history`, `.stamp`; free of every format's vocabulary, guarded by its `tests/test_neutrality.py`; pptx-agent depends on it (`pptx_agent.core` is a shim). 0.2 adds the optional `ooxml_edit.charts` (charts, their workbooks, SmartArt), extracted from pptx-agent's E4 | the OPC package, ordered insertion, undo/redo/batches, part reaping, id stamping; charts and SmartArt |
| `ooxml-common` (uvrt/ooxml-common, standard library only) | Public, used by the renderers | not a dependency of the editor (it is the renderers'); docx2svg brings it |
| `docx2svg` | Draws documents page by page, measured against Word; a hard dependency | rendering, layout positions, the effective-property resolver, the Word oracle tools |

**What docx2svg gives the agent today** (from its README and ROADMAP):

- `convert_docx_to_svg(source, options) -> list[str]`, one SVG per page;
  `convert_docx_to_png`; `convert_docx_to_layout(source, options) -> Layout`, the laid-out
  pages without SVG. A `Layout` has `pages`, each with `lines` (`Line.path`,
  `Line.paragraph_id`, `baseline`, `top`, `pitch`, `story`, `part`), `floats`, `pictures`,
  `rules`, and `stop` when the layout could not go on; `convert_docx(source, options) ->
  Conversion(layout, svgs, page_numbers)` lays out once and gives both; `Line.column` is the
  text column. `ConvertOptions.warnings` carries stable codes (`layout-stopped:table`...).
- **Identity:** every drawn element carries `data-docx-path`, a structural path counted
  XPath-style among same-named siblings by *local* name, always written with a `w:` prefix
  (`w:body/w:p[3]/w:hyperlink[1]/w:r[2]`, `w:body/w:tbl[1]/w:tr[2]/w:tc[1]/w:p[1]`); a
  paragraph's group also carries `data-docx-id` (its `w14:paraId`, "which Word copies with
  the paragraph, so it is not unique; hence both"), and a header's or footer's
  `data-docx-story` and `data-docx-part`.
- **What it draws:** text in Word's face, size and position (glyph-exact on every committed
  document), decorations, list labels, paragraph borders and shading, tab leaders, inline
  pictures, tables (including nested and autofit where measured), text columns, footnotes
  and endnotes, headers and footers with `PAGE`/`NUMPAGES`/`SECTIONPAGES` computed (other
  fields as cached), floating drawings, shapes, text boxes, charts and SmartArt (from the
  cached drawing). It hyphenates as Word does for English, Dutch, German and French.
- **What still stops it** — a chart's 3-D scene (drawn flat), SmartArt without a cached
  drawing, picture fills, a header's drawing the body wraps around, a floating table or
  drawing not positioned against its column in a multi-column section, an endnote in one
  in mode 15, a footnote in one a continuous break joins to another section, a cell
  spanning columns wider than they are, frames other than drop caps, and any paragraph it
  cannot measure (a face this machine lacks). Where it stops it draws a marked band,
  invents no page after it, and warns. **Reflow feedback has to carry that through:** a
  page past a stop is *unknown*, never guessed.
- **Revisions** — since docx2svg's final-view work (October 2026: moved text drawn at its
  destination, paragraphs joined across deleted marks, every content control's content;
  docx-agent's own markup fixture renders that way) the gaps below are closed. As first
  surveyed, read from `parse/document.py` and checked with a five-line probe document (not
  committed), it drew the **final** view, partly. Text in `w:ins` is drawn
  (unmarked) and `w:del` is skipped, which is right; but text in `w:moveTo` is **dropped**
  (the final view should show it), a paragraph whose mark is deleted is **not joined** to
  the next, and formatting changes simply show the new formatting. Separately, the text of
  an inline content control (`w:sdt` inside a paragraph) is **not drawn**: the parser
  descends into the `w:sdt` but not its `w:sdtContent`. Each is a docx2svg change for its
  own roadmap; this file proposes them (E3), it does not make them.

---

## The document model: live views over lxml

As in pptx-agent, **a typed node is a view over an lxml element**, never a copy. Reading
`paragraph.style` reads `w:pPr/w:pStyle/@w:val` on demand; nothing is lifted out of the
tree, so everything the model does not know — `mc:AlternateContent`, `w14`/`w15`/`w16*`
extensions, custom XML parts, `w:customXml`, permission ranges, proofing marks, math —
survives because nothing removed it.

| View | Over | Notes |
| --- | --- | --- |
| `Document` | the package | stories, styles, numbering, settings, undo, render, layout |
| `Story` | `w:body`, `w:hdr`, `w:ftr`, `w:footnote`, `w:endnote`, `w:comment`, `w:txbxContent` | anything that holds blocks; each has its part |
| `Block` (`Paragraph`, `Table`, `ContentControl`) | `w:p`, `w:tbl`, block-level `w:sdt` | blocks in order, descending `w:sdtContent` and `w:customXml` transparently |
| `Section` | `w:sectPr` (in the last paragraph's `w:pPr`, or the body's last child) | page size and orientation, margins, columns, header/footer references, `titlePg`, numbering |
| `Paragraph` | `w:p` | style, direct `w:pPr`, text (see *Text ranges*), runs, list membership, section break |
| `Run` and inline items | `w:r` and its `w:t`, `w:tab`, `w:br`, `w:sym`, `w:noBreakHyphen`, `w:softHyphen`, `w:drawing`, note and comment references | positional; containers (`w:hyperlink`, `w:ins`, `w:del`, `w:moveTo`, `w:moveFrom`, `w:smartTag`, inline `w:sdt`, `w:customXml`, `w:fldSimple`) are walked, not flattened |
| `Table`, `Row`, `Cell` | `w:tbl`, `w:tr`, `w:tc` | a *grid* view resolving `w:gridSpan` and `w:vMerge`; a cell holds blocks, so tables nest |
| `ContentControl` | `w:sdt` at block, run, row or cell level | kind (`w:text`, `w:richText`, checkbox, drop-down, date, picture, repeating section), tag, alias, lock, data binding |
| `Field` | `w:fldChar` begin/separate/end with `w:instrText`, or `w:fldSimple` | instruction, cached result, dirty flag; may span runs and paragraphs |
| `Hyperlink` | `w:hyperlink` | `r:id` (external) or `w:anchor` (bookmark) |
| `Bookmark` | `w:bookmarkStart`/`w:bookmarkEnd` | name, id; may span paragraphs and table cells |
| `Note` | `w:footnote`/`w:endnote` and its reference run | separators (`w:type`) are not notes |
| `Comment` | `w:comment` plus `commentsExtended`/`commentsIds`/`commentsExtensible`/`people` | range markers, reference run, replies, resolved state, durable id |
| `Drawing` | `w:drawing` → `wp:inline` or `wp:anchor` | `wp:docPr` (id, name, alt text), extent, wrap, position; picture, shape, group, text box, chart, SmartArt |
| `Revision` | `w:ins`, `w:del`, `w:moveFrom`/`w:moveTo`, `w:rPrChange`, `w:pPrChange`, `w:sectPrChange`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, `w:tblGridChange`, `w:numberingChange`, paragraph-mark and row `w:ins`/`w:del`, `w:cellIns`/`w:cellDel`/`w:cellMerge` | E3 |
| `Numbering` | `numbering.xml`: `w:abstractNum`, `w:num`, `w:lvlOverride` | list definitions and instances |
| `Styles` | `styles.xml`: `w:docDefaults`, `w:latentStyles`, `w:style` | lookup by name *and* id (ids are localised: a Dutch template's "heading 1" has id `Kop1`) |
| `Settings` | `settings.xml` | compatibility mode, `w:trackRevisions`, `w:evenAndOddHeaders`, `w:updateFields`, `w:rsids`, note properties; a strictly ordered sequence |

**Effective values come from docx2svg, not a second cascade.** docx2svg's `resolve` has
the style cascade measured against Word — including where Word departs from ECMA-376
(toggle properties, numbering precedence, complex-script classification, theme fonts) — and
says which level every value came from. `run.effective.bold` asks it, over the current
bytes of the part (cached per undo state). Writing is always to a *declared* level: a style,
or direct formatting.

**Compatibility mode** is read from `w:compat/w:compatSetting[@w:name="compatibilityMode"]`
(absent means Word 2007's mode 12). docx2svg's layout already differs by mode (justification,
autospacing, footnotes in columns), so the mode is part of every layout request.

### Code layout (proposed)

Mirroring pptx-agent (E0 built `oxml`, `edit`, `layout` and `validate`): `src/docx_agent/oxml`
(WordprocessingML namespaces, child-order tables, `WordPackage` on ooxml-edit's package),
`edit` (the views and operations),
`revisions` (E3), `markdown` (E2: `stylemap.py`, the one table; `read.py`, the document into
records and the neutral `model.py`; `render.py`, the model as text; `parse.py`, Markdown
back into the model's normalised AST through markdown-it-py; `html.py`, the table fallback;
`write.py` with the write half), `layout` (the docx2svg bridge), `state` (the JSON
projection). Tests as in pptx-agent: `test_roundtrip`,
`test_validity`, `test_ids`, `tests/oracle.py`, and one file per phase.

---

## Addressing and stable ids

An agent holds a reference across edits: "make the paragraph after the table bold, then
comment on it". pptx-agent's rule carries over: **ids name things, positions name places,
and the API says which it is giving you.**

### Paragraphs: `w14:paraId`

[MS-DOCX](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/a0e7d2e2-2246-44c6-96e8-1cf009823615)
defines `w14:paraId` on `w:p` (and `w:tr`) as an `ST_LongHexNumber` "unique within the
document part", except across the choices of an `mc:AlternateContent`, with values greater
than 0 and less than `0x80000000`. `w14:textId` has the same range, requires a `paraId`
beside it, and versions the paragraph's text. Word generates both randomly.

What is known, and what is not:

| Question | Evidence | Status |
| --- | --- | --- |
| Does Word write them? | Reportedly Word 2010 and later write them on every paragraph on save; **Word 16.106 for Mac does not**: a document it makes has none, and it writes them only when every paragraph has one or comments need them | **measured** (E0) |
| Are they unique in real files? | Not reliably: content copied by tools (and, per docx2svg, by Word copying a paragraph) repeats them; files from Word 2007, LibreOffice, Google Docs and generators often have none | known; corpus to confirm |
| Unique across parts? | No: only within a part, so a header and the body may share one | by specification |
| Does Word keep ids it did not write? | Open XML SDK #925: Word **replaced** SDK-generated paraIds on the first save after a comment was added, then kept them. Measured: kept, *only if* every paragraph, row and drawing in every story has its ids (below) | **measured** (E0) |
| Does Word keep them through an edit of the paragraph in Word? | `paraId` kept, `textId` renewed | **measured** (E0) |
| What does Word do with duplicates and out-of-range values? | A repeat: replaced, the first kept. One out of range: every paraId in the document dropped | **measured** (E0) |

**The id.** A paragraph's id is `p:<paraId>` in the body and `<story>/p:<paraId>`
elsewhere, where the story is the part's stem (`header2/p:1A2B3C4D`, `footnotes/p:...`,
`comments/p:...`). The paraId is used verbatim, so the id an agent sees is the same one
docx2svg puts in `data-docx-id`.

**Missing.** A paragraph without a paraId has a *positional* id, `p@<story>/<n>` — honest
about being volatile — until it is first edited. Then it is **stamped**: given a paraId and a
textId, deterministically generated (a counter-seeded sequence, checked for uniqueness in
the part and below `0x80000000`), with `w14` declared on the part's root and listed in
`mc:Ignorable` if it is not already. Reading never stamps, or the byte-identity guarantee
would fail. This is pptx-agent's lazy stamping, with the stamp in Word's own attribute
rather than an `extLst` (a `w:p` has none).

**Duplicated.** The first occurrence in document order keeps `p:<paraId>`; later ones are
`p:<paraId>#<n>` (occurrence index, volatile). The first edit of a duplicate gives it a
fresh paraId — a duplicate is invalid by the specification — and the session keeps an
alias from the old id, so a reference the agent holds still resolves; every edit result
reports renames.

**Durable across a Word save — measured** (E0, Word 16.106 for Mac, 2026-10-03;
`tools/paraid_probe.py`, observations in `tests/observations/paraid-durability.json`).
The three ways on this section first listed -- re-anchoring by `textId` and text, hidden
`_dxa_` bookmarks, or both -- were not needed. What decides it is not *who* wrote a paraId
but whether the document is **complete**:

| Probe (saved by Word) | What Word wrote |
| --- | --- |
| Every paragraph with a valid, unique paraId (another tool's, and one docx-agent stamped) | every paraId and textId kept |
| ... and one paragraph without a paraId | **no paraId at all**, anywhere |
| ... and one paraId out of range (`80000005`) | no paraId at all |
| ... and one table row (`w:tr`) without a paraId | no paraId at all |
| ... and one *header* paragraph without a paraId (the body complete) | no paraId at all, the body's included |
| ... and one picture (`wp:inline`) without `wp14:anchorId` and `wp14:editId` | no paraId at all (with both: kept) |
| ... and one repeated paraId | the repeat replaced, the first and all others kept |
| ... and a comment, an insertion, a bookmark, a footnote | every paraId kept |
| Mixed (missing, out of range, repeated, another tool's, ours) with a comment part | every paraId **replaced** -- not dropped, presumably because the comments need them -- textIds Word had no reason to change kept |
| A new document Word made | no paraIds (Word writes none); one paragraph stamped by docx-agent: dropped on save |

Typing into a paragraph in Word keeps its paraId and renews its textId; adding a comment,
typing elsewhere and saving twice change nothing else. Modes 15, 14 and none (12) behave
alike. Note, comment, revision and bookmark ids Word **renumbers** on every save (footnote
5 → 1, comment 7 → 0, insertion 42 → 0, bookmark 9 → 0).

**The way on, chosen on that evidence:** the first edit stamps the **whole document**
(`Document(stamping="document")`, the default): every paragraph and table row, in every
story, without a valid and unique paraId gets one (and a textId), out-of-range values and
repeats are re-issued, and every drawing without `wp14:anchorId`/`wp14:editId` gets them,
`w14`/`wp14` declared and ignorable, each rename kept as an alias. Measured the same way,
those ids then survive every scenario -- save, save again, typing elsewhere, typing into a
stamped paragraph, a comment -- in modes 15, 14 and none, in a generated document, one with
a picture and one Word made. Stamping is still lazy (reading never stamps) and
deterministic; its cost is that the first edit rewrites every story part that lacked ids.
`stamping="paragraph"` stamps only what an edit touches, for callers who prefer fewer
changed bytes to ids that outlive a Word save. Two residues, both Word's own: the
paragraphs of footnote and endnote *separators* get new paraIds on every save (Word writes
those notes itself), and a document another tool later breaks (a paragraph added without a
paraId) loses its paraIds on Word's next save -- re-anchoring by `textId` and text, kept as
an option, would recover from that; it is not built. `w:customXml` markup cannot be used
as a stamp either way: current Word removes it on open.

### Everything else

| Thing | Id | Durable? | Notes |
| --- | --- | --- | --- |
| Run | `p:<id>/r<n>` | positional | re-resolved per call, as in pptx-agent: runs have no identity in OOXML. Counts runs inside `w:hyperlink`, `w:ins`, inline `w:sdt` in document order; `w:del` runs count only in the `original` and `markup` views |
| Text range | `p:<id>@<start>:<end>` | relative to a durable paragraph | see below |
| Table | `t:<paraId of its first row>` | as rows | `w:tbl` has no id; `w:tr` carries `w14:paraId` |
| Row | `tr:<paraId>` | as paragraphs | positional fallback `t:<id>/r<n>` |
| Cell | `t:<id>/c<row>,<col>` | positional in the *grid* | a merged cell answers to its origin |
| Section | `s:<paraId>` of the paragraph that ends it, `s:body` for the last | as paragraphs | a section break *is* that paragraph's `w:sectPr` |
| Header/footer story | `header2`, `footer1` (part stem); also `s:<id>/header/first|default|even` | part names are stable unless parts are renamed | the second form says which section uses it, resolving inheritance |
| Footnote, endnote | `fn:<w:id>`, `en:<w:id>` | for the session: Word renumbers note ids on save (measured, 5 → 1) | E4; a tracked move is a copy, a new id |
| Field | `fld:<story>/<n>` | positional: the n-th field of its story (nested ones counted), as runs are | E4 |
| Comment | `c:<w16cid:durableId>` when `commentsIds.xml` has one, else `c#<w:id>` | durableId is Word's own durable id; `w:id` Word renumbers on save (measured, 7 → 0) | replies link through `commentsExtended` by the comment's paragraph `w15:paraId`; Word adds `commentsIds.xml` on saving |
| Drawing | `d:<wp:docPr/@id>` | `docPr@id` should be unique in the document; collisions get `#n`, as pptx-agent's `cNvPr@id` | |
| Group member | `d:<group's docPr id>/<its wpg:cNvPr id>`, a nested group's members a step further (`d:5/9/7`) | Word keeps a member's `cNvPr` id through a save (measured); a repeat within one group gets `#n` | a member has no `docPr`; a chart in a group is addressed this way (`Document.chart("d:5/7")`, `Drawing.members`) |
| Bookmark | `bm:<name>` | names are unique by rule; Word renumbers bookmark `w:id`s on save (measured, 9 → 0) | hidden bookmarks start with `_` |
| Hyperlink | `hl:<paragraph id>/<n>` | positional: the n-th `w:hyperlink` of its paragraph | E1 |
| Style | `style:<w:styleId>` (from `Styles.add`); found by `w:name` | Word renames ids into its interface language on save (measured, `Heading2` → `Kop2`): address styles by name | E1 |
| Content control | `cc:<w:sdtPr/w:id>` | Word writes random ids; tags are not unique | |
| Revision | `rev:<w:id>` | for the session: Word renumbers them on save (measured, 42 → 0) | E3 |

### Text ranges and anchors

A paragraph's **text** is defined once and used everywhere (ranges, find, Markdown):
`w:t` characters; `w:tab` → `\t`; a line break `w:br` → `\v` (pptx-agent's convention); a
page or column break → `\f`; `w:noBreakHyphen` → U+2011; `w:softHyphen` → U+00AD;
`w:sym` → its character; a field → its *result* text (instructions are never text); a
drawing, note reference or other object → U+FFFC, one character, so offsets past it do not
depend on its contents. Offsets are Python string indices into that text.

The **view** decides what counts: `current` (default: insertions in, deletions out — what
Word's "No Markup" shows), `original` (the reverse), or `markup` (both, as
`getReviewedText` does not offer). A range is valid in the view it was taken in and is
rejected, not reinterpreted, in another.

**Find-by-text anchors** for agents that read Markdown: `doc.find("net revenue", within=
"p:3A1F09C2")` returns ranges; `doc.anchor("net revenue", occurrence=None)` returns *one*
range or raises `AmbiguousAnchor` listing the candidates with their ids (Adeu's rule).
Matching normalises nothing by default; `whitespace="collapse"` and `case=False` are
explicit.

---

## Edit operations

Every operation is one undo step (ooxml-edit's batches), byte-exact on undo, validates
before it changes anything, and returns what it touched (ids created, renamed, removed).
**Each is built on a handful of primitives** — insert inline content at a position, remove
a range, set properties on a run/paragraph/table/row/cell/section, insert or remove a
block — and *tracking is a mode of those primitives* (E3), so every operation is trackable
without a second implementation.

**Text.** `paragraph.set_text(text)` keeps mixed formatting by diffing old and new text as
pptx-agent's E1 does (each new character takes the formatting of the character it replaces
or follows); `range.replace(text)`, `range.delete()`, `insert_text(at, text)`;
`insert_paragraph(after|before, text, style=)`; `delete_block(id)`; `move_block(id, to)`.

**Formatting, styles first.** `paragraph.style = "Heading 2"` resolves the *name* (then the
id, then a built-in latent style, whose definition is written as Word writes it — measured);
`run.style` for character styles; `Styles.add/modify` for new or changed styles.
**Direct formatting second:** `run.bold`, `italic`, `underline`, `strike`, `size`, `font`,
`color` (theme colours kept as `w:themeColor`), `highlight`, `paragraph.alignment`,
`indent`, `spacing`, `keep_with_next`, `page_break_before`, and `clear_direct_formatting()`,
which is how an agent turns a mess back into a styled document. Reads distinguish *declared*
(`run.bold` → `None` when inherited) from *effective* (`run.effective.bold`, from docx2svg).

**Lists and numbering.** Apply a list style ("List Bullet", "List Number" and their 2–5
levels) or a numbering instance; `indent()`/`outdent()` (`w:ilvl`); `restart(at=1)` — a new
`w:num` over the same `w:abstractNum` with `w:startOverride`, as Word does; `continue_from
(id)`; change a level's format or bullet, copy-on-write when the abstract definition is
shared by other lists. Integrity is checked after every edit: every `w:numId` resolves, every
`w:num` names an existing abstract definition, `w:numId="0"` means "no list".

**Tables.** Cell text through the text API; insert/delete rows and columns, merge/split
(`w:gridSpan`, `w:vMerge`), with pptx-agent's merge-aware rules and its randomised sweep;
`w:tblGrid` kept in step; widths, borders, shading, table style and `w:tblLook`; header
rows (`w:tblHeader`); nested tables (a cell holds blocks); a cell always ends in a `w:p`.

**Sections and page setup.** Size and orientation, margins, columns, section breaks
(`nextPage`, `continuous`, `evenPage`, `oddPage`) inserted by splitting — a copy of the
section's `w:sectPr` goes into a paragraph of its own before the break, as Word writes a
break put at a paragraph's start (measured, E4) — page numbering (`w:pgNumType` format and
restart), `titlePg`, line numbering, vertical alignment. Built in E4: `insert_section_break`,
`remove_section_break`, `set_section` (see *Phase E4*).

**Headers and footers.** Create, edit (through the same text API on their story), link to
previous (remove the reference), unlink (copy the inherited part), first-page and
even/odd (`w:evenAndOddHeaders` is document-wide: the API says so when it is set). Built in
E4: `add_header`/`add_footer`, `link_to_previous`, `unlink_from_previous`, `Section.header`.

**Fields.** Read instruction and result; insert `PAGE`, `NUMPAGES`, `SECTIONPAGES`, `DATE`,
`REF`, `PAGEREF`, `SEQ`, `HYPERLINK`, `TOC`. **Update** is the hard part: docx-agent can
compute `REF`, `SEQ` and cross-reference text itself, and `PAGE`/`PAGEREF`/`TOC` page
numbers from docx2svg's layout (a TOC changes the layout it reports, so it is laid out
until the numbers stop moving, at most a few passes); where the layout cannot say, the
result is left and the field reported unknown. ~~Marked dirty~~: measured in E4, a
`w:dirty` field makes Word ask on opening exactly as `w:updateFields` does (its export
blocks on the question), so docx-agent writes neither. Setting `w:updateFields` makes Word
ask the user a question on open; it is an option, never the default (see Decisions).

**Images and drawings.** Insert an inline picture (media de-duplicated by content, as in
pptx-agent; `wp:docPr` id unique in the document; size from pixels and density), replace an
image keeping the frame, alt text, resize; floating drawings (E5): position
(`wp:positionH/V`, `relativeFrom`), wrap type, behind/in front; delete with part reaping.

**Footnotes and endnotes.** Insert at a position (the reference run with the Footnote
Reference style, the note with Footnote Text and its `w:footnoteRef` run), edit, delete,
creating `footnotes.xml` with Word's separator notes if it is missing; ids one above the
largest in use. Built in E4, with moving a note and its numbering per section.

**Comments.** Add on a range (start/end markers, reference run, `comments.xml`; the modern
parts `commentsExtended`, `commentsIds`, `commentsExtensible` and `people.xml` as Word 365
writes them — measured), reply, resolve (`w15:done`), edit, delete with its replies. Built
in E3 (`edit/comments.py`): `add_comment(range, text, author=, initials=, date=)` returns
`c:<durableId>`; `reply_to_comment`, `resolve_comment`/`reopen_comment` (a thread's first
comment carries `w15:done`), `edit_comment` (the last paragraph keeps the paraId the modern
parts key it by), `delete_comment` (with its replies; the last one gone, the comment parts
go, as Word writes a document without comments); `Comment` views have `reply`, `resolve`,
`reopen`, `edit`, `delete`, `replies`. As Word writes them (measured): each comment
paragraph in Comment Text ("annotation text"), the first opening with `w:annotationRef` in
Comment Reference ("annotation reference") -- both added from Word's own definitions,
`word_comment_styles.py`, with English Word's ids `CommentText`/`CommentReference`; the
reference run after the range at the size of the text it follows; a reply anchored where
its parent is; `commentsExtended` (`w15:paraId` of the last paragraph, `w15:done`,
`w15:paraIdParent`), `commentsIds` (`w16cid:durableId`), `commentsExtensible`
(`w16cex:durableId`, `w16cex:dateUtc`), `people.xml`. Adding a comment to a document whose
comments lack those entries gives them theirs, as Word does on saving.

**Bookmarks and cross-references.** Add/rename/remove a bookmark on a range (names start
with a letter, at most 40 characters, unique); a cross-reference is a `REF` or `PAGEREF`
field with `\h` to it; an internal hyperlink is `w:hyperlink w:anchor`.

**Content controls.** Read and fill plain-text, rich-text, checkbox, drop-down and date
controls, keeping `w:sdtPr`; respect `w:lock`; and **if the control is data-bound**
(`w:dataBinding`), write the bound node of the custom XML part as well — otherwise Word
overwrites the control from the XML on open. Creating bindings is out of scope.

**Find and replace across run boundaries** — the classic docx problem: Word splits text into
runs at every formatting, rsid, spell-check (`w:proofErr`) and revision boundary, so "net
revenue" may be three runs. The algorithm (after PowerTools' `OpenXmlRegex`): build the
paragraph's text map (character → run and offset, in the chosen view), match a literal or
regex over the joined text, then split runs at the match's edges and replace the
characters; the replacement takes the formatting of the match's first character (Word's
rule). Matches that would cut a field in two, or run from a field's result into ordinary
text, are refused; matches across paragraphs are opt-in and done as a delete and insert.
Under tracking a replacement is a `w:del` plus a `w:ins`. **Runs are never coalesced
implicitly** (it would change untouched bytes); `paragraph.coalesce_runs()` merges adjacent
runs whose properties differ only in rsids, explicitly.

**Moving and copying blocks between documents** (E6). `target.insert_blocks(at,
source.blocks(range))` imports what the blocks depend on, after PowerTools'
`DocumentBuilder`: styles by name, with a policy — `use_destination` (Word's default when
pasting) or `keep_source` (renamed on conflict); numbering (abstract definitions and
instances cloned, `numId`s remapped, so a copied list does not continue the destination's);
images and other related parts (copied, de-duplicated, `r:id`s remapped); footnotes,
endnotes and comments the blocks reference (copied and renumbered); bookmarks (renamed on
conflict); paraIds and `docPr` ids (re-issued on conflict). Within one document, `move_block`
keeps ids; a copy gets fresh ones.

---

## Tracked changes

```python
with doc.tracking(author="Claude", date=None, initials="CL"):   # date None: now, UTC
    doc.paragraph("p:3A1F09C2").set_text("Revenue grew 14%.")
    doc.anchor("draft").replace("final")
    doc.paragraph("p:0C11D2E7").style = "Heading 2"

for rev in doc.revisions(author="Claude"):   # kind, author, date, story, range, ids
    ...
doc.revision("rev:12").accept(); doc.revisions(author="Alice").reject()
doc.settings.track_revisions = True          # w:trackRevisions: Word keeps tracking
```

`doc.tracking(...)` is a mode the primitives consult (Aspose's `StartTrackRevisions`, not a
second API). `w:trackRevisions` in settings is a different thing — it tells *Word* to
track the next person's edits — and the two are set independently. As built (E3):
`doc.tracking(author=, date=, initials=)` (date `None` is now, to the minute, UTC); every
edit also takes `track=True` (the context's author, else `docx-agent`) or `track=False`;
`doc.revisions(author=, kind=, within=, since=, until=)` lists `Revision`s (`id`, `kind`,
`author`, `date`, `text`, `range` in the markup view, `ids` it affects); `doc.accept(...)`
and `doc.reject(...)` take an id, a list, a range or block id, or nothing (all, after the
filters), `accept_all()`/`reject_all()`; `doc.word_tracks_changes` (or
`set_word_tracks_changes`; `track_revisions` and `set_track_revisions` until the pilot
renamed them, now deprecated aliases) is the setting.

### How each primitive is expressed

Measured on Word 16.106 for Mac (`tools/e3_probe.py`: Word itself, tracking, typing and
deleting by AppleScript, then saving; observations in `tests/observations/e3-word.json`),
and written so (E3):

| Edit | What Word writes, and docx-agent with it |
| --- | --- |
| Insert text | a run of its own in `w:ins` (`w:id`, `w:author`, `w:date`, `w16du:dateUtc`); text inserted into one's own insertion joins it, into another author's splits it (each part its own id) |
| Delete text | the runs in `w:del`, `w:t` → `w:delText`, a field's begin/separate/end runs with them and `w:instrText` → `w:delInstrText`; **one's own insertion deleted goes outright** (measured: typing then deleting part of it leaves only the rest, in `w:ins`); another author's insertion gets a `w:del` inside it |
| Replace text | the deletion, then the insertion after it |
| Insert a paragraph | Word, typing Enter at a paragraph's end: *that* paragraph's mark `w:ins`, the new one holding its old mark and the text in `w:ins`. docx-agent: the new paragraph's own mark inserted where a block follows it (Word's Reject All handles that cleanly: measured), Word's form where it is last in its container (where Word cannot accept or reject a last mark: measured, a cell's pulls the cell's first paragraph out of the table, the body's keeps its record) |
| Delete a paragraph | its content `w:del` and its mark `w:pPr/w:rPr/w:del`, no change to the next; last in its container, the mark before it is the deleted one and it takes that paragraph's properties as a `w:pPrChange` |
| Delete across a paragraph mark | the first paragraph's mark `w:del`, and **the next paragraph given the first one's properties, its own kept in a `w:pPrChange`** (measured), so accepting makes what an untracked join makes |
| Run formatting | `w:rPrChange` holding the complete old `w:rPr` (an empty one when there was none); formatting set back drops the record |
| Paragraph formatting, style, list membership | `w:pPrChange` with the old `w:pPr` (empty when none); a list is `w:pStyle` List Paragraph and `w:numPr` in the new properties -- Word writes no `w:numberingChange` |
| Section, table, row, cell properties | Word writes `w:sectPrChange` with only the values it changed; `w:tblPrChange` with the whole old `w:tblPr` and a `w:tblGridChange` (`w:id` alone); `w:tcPrChange` per cell -- **every cell of the table**, changed or not (measured in E5), and where the grid changed a `w:trPrChange` per row (`w:gridAfter w:val="0"`); a row's height and alignment it changes without a record. E4 writes the section's, E5 the table's and cells' in Word's whole-table form (a row's alignment also as a `w:trPrChange`, so a review gives it back) |
| Insert/delete a table row | `w:trPr/w:ins` with each new cell's mark inserted; `w:trPr/w:del` with each cell's marks and content deleted |
| Insert a column | Word: the new cells, each with its paragraph mark inserted, the grid widened -- no `w:cellIns`, no `w:tblGridChange`; its Reject All removes such cells and their grid column (measured). docx-agent writes the same |
| Delete a column, merge cells | **Word tracks none of them: it applies them untracked** (measured; a cell split too). Decided: supported, never refused. Written in the nearest forms Word accepts (measured on files written so): a deleted column as `w:cellDel` on its cells with their content deleted (Word's Accept All removes them and their grid column, its Reject All restores them); a merge down as `w:cellMerge` `rest`/`cont` (Word's Accept All merges, its Reject All restores); a merge across as the first cell's new span in a `w:tcPrChange` and `w:cellDel` on the others; merged cells' content deleted where it was and inserted in the first cell. Each with a comment saying what changed (`Tracked structural change: ...`), which accepting or rejecting it here removes. ~~Word's Accept All and Reject All of a merge across do not give the same table~~ -- E5: with the merged-away cells' paragraph marks deleted and Word's whole-table record (every cell's old properties, the old grid) they give exactly ours; splits, a column across a span and a merge's first row deleted likewise (a divided merge as `w:cellMerge` `vMerge`/`vMergeOrig` beside the old properties) |
| Move a paragraph | the paragraph stays, its mark and runs in `w:moveFrom` (text stays `w:t`), between `w:moveFromRangeStart` (after its properties; `w:id`, `w:author`, `w:date`, `w:name`) and `w:moveFromRangeEnd` (after it); a new paragraph at the destination, the same in `w:moveTo`, the ranges sharing a name. The caller's id follows to the destination (an alias). A table, a paragraph holding other revisions, one last in its container moves as a deletion and an insertion; one holding a note or comment reference moves untracked, with a warning |
| Insert a picture, a field | the run(s) in `w:ins`; a picture's resize and alt text are not tracked (Word tracks no resize: measured); a picture replaced is the old deleted and the new inserted |
| Hyperlinks | Word inserts a `HYPERLINK` field over the text, the text itself counted inserted (so its Reject All loses the text: measured). docx-agent: the text deleted and a `w:hyperlink` with its copy inserted after it -- never a hyperlink inside `w:ins` (Word drops the link then); Word re-saves a hyperlink holding a revision as a field, and the element once accepted |
| Bookmarks, style definitions, list definitions | not revisions in Word: written as untracked |
| Comments | not revisions; written the same in any mode (below) |
| People | Word lists every author of a revision or comment in `people.xml` (`w15:person`, `w15:presenceInfo` provider `None`, user id the name) and drops the part when no one is left: so does docx-agent |

**Revision ids.** `w:id` on annotations is allocated one above the largest id in use across
*all* annotation kinds in *all* stories (revisions, comments, bookmarks), which is stricter
than the specification and what Word does (measured: Word numbers revisions and comments in
one sequence). Word renumbers on save (measured in E0), so a `rev:` id is valid for the
session; the durable handle for a revision is its range's paragraph id plus author and
date. A record an edit copies (a split run's `w:rPrChange`, a split insertion) gets an id
of its own.

**Dates.** Word writes its local time in `w:date` -- with a `Z` all the same -- and the
true UTC in `w16du:dateUtc` (measured); docx-agent writes the UTC in both, so its output
does not depend on the machine's time zone.

**`w:rsid`.** Revision save ids say in which editing *session* a change was made; Word adds
them on save and uses them when combining documents. docx-agent never strips or rewrites
existing ones. By default it writes none (output stays deterministic and the session needs
no `w:rsids` bookkeeping); an option allocates one session rsid, registers it in
`settings.xml`'s `w:rsids`, and stamps `w:rsidR`/`w:rsidRPr` on what it creates. Open
question below.

**Dates** are ISO 8601 in UTC with a `Z`; tests inject a fixed date so outputs are
reproducible.

### Accept and reject

Accept or reject by id, by author, by date range, by kind, within a range, or all — the
filters Office.js and the MCP servers converged on. The algorithms follow PowerTools'
`RevisionProcessor` and are verified, not trusted:

- accept insertion → unwrap; reject → remove (and an inserted paragraph mark merges the
  paragraphs);
- accept deletion → remove; reject → unwrap, `w:delText` → `w:t`;
- **a paragraph mark that goes** (a deletion accepted, an insertion rejected) **joins its
  paragraph into the next: the next keeps its own element, paraId and properties**, the
  first paragraph's content moving to its start (measured, `JoinX`: a deleted mark with no
  `w:pPrChange` accepted keeps the second paragraph's properties; `JoinA`: tracking, Word
  gives the second the first's properties beforehand, so accepting keeps the first's). This
  is also docx2svg's final view (R.2: the next paragraph's properties): accepting and the
  final view agree here. The removed paragraph's id becomes an alias of the survivor.
  **Untracked**, Word joins the same way (measured, `Mark`): the second paragraph's
  element and id, the first one's properties -- E1's join now does that (it kept the
  first paragraph's element);
- with no paragraph after it in its container, an empty paragraph goes and one with content
  keeps its mark (Word: a cell's first paragraph leaves the table, the body's last mark
  stays; docx-agent writes forms that avoid those marks, above); a cell whose only,
  empty paragraph's mark was inserted goes with the grid column it leaves, as Word's Reject
  All of its own column insert does;
- property changes → accept drops the `*Change` element, reject restores the old properties
  (a `w:sectPrChange` holds only what changed: its values replace their own kind);
- moves → both ends together, marks and range markers included;
- rows: a deleted row accepted or an inserted row rejected goes (the table too, with its
  last row); cells: `w:cellDel` accepted or `w:cellIns` rejected removes the cell and any
  grid column no remaining cell covers; `w:cellMerge` accepted sets `w:vMerge`;
- a record goes with the others made with it (same author and date): a row's or table's
  cell records and their content, a paragraph's content with its mark;
- a comment whose reference a review removes goes too (measured: Word's Reject All drops a
  comment on rejected inserted text); pictures' and hyperlinks' relationships are released;
  numbering instances the review left unused are removed (a style an edit added stays, as
  in Word); people with nothing left are dropped.

**The invariant the tests hold:** for every operation `op` and fixture `d`,
`accept_all(tracked(op)(d))` equals `op(d)`, and `reject_all(tracked(op)(d))` equals `d`, in
canonical form — and both export from Word unprompted.

### How docx2svg should render revisions

docx2svg draws the final view today, with the gaps listed under *Where we are* (`w:moveTo`
dropped, deleted paragraph marks not joined). Proposed, for docx2svg's own roadmap:

1. **Final** — fix the gaps; this is what `No Markup` shows and the default.
Only the final view is rendered (decided by the user, 2026-10-03): tracked changes and
comments are **not visualised** -- no original view, no inline markup, no balloons, no
comment highlights. An agent reviews revisions and comments through the API and
`to_markdown`'s revision views, which are text, not through renders.

What Word's *PDF export* shows for a document with revisions depends on the markup view
saved with it: its default export prints the markup (docx2svg R.0); `tests/oracle.py`
exports a document with markup in Word's *No Markup* view, and E3's oracle holds docx2svg's
drawing to it page by page.

---

## The Markdown layer

### The mapping

`StyleMap` is one declarative table read in both directions (mammoth's idea); a document or
template can override any row. Style names are *names* (`w:name`), resolved to the
document's ids, so a localised template works.

| Markdown (CommonMark + GFM) | Word, by default | Falls back to |
| --- | --- | --- |
| paragraph | the document's default paragraph style (usually Normal; Body Text if the template's body uses it) | Normal |
| `#`–`######` | Heading 1–6 | built-in definition written as Word writes it |
| `-`/`*`/`+` list, nesting | List Bullet, List Bullet 2–5 | the document's bullet list definition, else a new one |
| `1.` list, nesting, start number | List Number, List Number 2–5; a start other than 1 → `w:startOverride` | as above |
| block quote | Quote (Word built-in; decided) | — |
| fenced/indented code | HTML Preformatted (Word built-in; decided) | — |
| `` `code` `` | HTML Code (Word built-in; decided) | — |
| `*em*`, `**strong**` | Emphasis / Strong character styles (decided; follows the document's styles) | direct `w:i`, `w:b` by option |
| `~~strike~~` (GFM) | direct `w:strike` | — |
| `[text](https://...)` | `w:hyperlink` with an external relationship, Hyperlink character style | — |
| `[text](#name)` | `w:hyperlink w:anchor="name"` to a bookmark | — |
| `![alt](file)` | inline picture, `wp:docPr/@descr` = alt | only files under a caller-given directory |
| GFM table | `w:tbl` with the document's table style (the one its tables use most, else Table Grid); header row → `w:tblHeader` + `firstRow`; column alignment → cell paragraph `w:jc` | — |
| hard line break | `w:br` | — |
| thematic break `---` | empty paragraph with a bottom border (what Word's AutoFormat makes of `---`) | — |
| footnote `[^1]` (extension) | a footnote | — |
| raw HTML | refused, or kept as literal text by option | — |
| task list `- [ ]` (GFM) | checkbox content control (`w14:checkbox`) | — |

`insert_markdown(at, md, *, track=None, style_map=None)` inserts blocks *at* a block id
(before, after, or replacing it) or inline Markdown *into* a paragraph range; with `track`
it is one tracked insertion. It is one undo step and returns the new blocks' ids, so the
agent can refine them through the semantic API.

`to_markdown(range=None, *, view="final", ids=True, headers=False, notes=True)` projects
blocks (built in E2's read half; `view="current"` is accepted for `final`):

```markdown
<!-- p:3A1F09C2 -->
## Results

<!-- p:0C11D2E7 -->
Revenue grew **14%** in the third quarter.[^fn:2]

<!-- t:5D10A7E1 -->
| Region | Q3 |
| --- | ---: |
| North | 4.1 |

<!-- d:7 picture "Figure 1" 6.5x3.2in -->
```

Every block carries its id; what Markdown cannot say becomes a comment with an id (a
drawing that is not a picture, a content control, a section break, a page break, an
unsupported block), so the agent can still address it. `view="markup"` writes revisions
in CriticMarkup (`{++inserted++}`, `{--deleted--}`, a paragraph mark `{--¶--}`,
`{>>comment<<}`), each followed by a comment with its id (`<!-- rev:12 Alice -->`,
`<!-- c#3 -->`); `final` and `original` are the two clean views. A style the map does not
know is named on its block (`<!-- p:3A1F09C2 style: Callout -->`). What was built, and how
each construct is written, is under *Phase E2*.

### Guarantees, and explicit non-guarantees

**Guaranteed** (and tested):

- `to_markdown` never changes the document, not one byte.
- For the supported subset, `to_markdown(insert_markdown(blank, md))` is equivalent to `md`
  — the same CommonMark AST after normalisation (list markers, emphasis delimiters,
  whitespace; `parse.roundtrip_ast` says the rest) — in a blank document and in every
  template fixture. The CommonMark spec examples outside it are listed, with their reasons,
  under *Phase E2* (59 of 575).
- Ids in the projection resolve through the API to the blocks they label, in the same
  session and undo state.
- `insert_markdown` writes only the document's styles (or creates the mapped ones it lacks,
  and says so), never a hidden format of its own: direct formatting only where Markdown
  has no style (strikethrough, a thematic break's border, a table cell's alignment) and a
  run's second mark beside its one character style.

**Not guaranteed:**

- docx → Markdown → docx is not a round trip. Markdown has no colours, fonts, alignment,
  spacing, sections, headers, floating objects or most field types; those are what the
  semantic API is for.
- Markdown → docx → Markdown is not byte-identical (only AST-equivalent).
- `to_markdown` is a projection, not a rendering: direct formatting Markdown has no word
  for (underline, colour, size, font, highlight, super- and subscript, alignment, spacing)
  is not in it -- `state()` has it; nor is a list's label format (a numbered item is
  written with its number, whatever `a.` or `iv.` Word draws), hidden text, a paragraph's
  leading whitespace or a table's header row when the table has none (GFM needs one: the
  first row stands in). Its `original` view undoes the *text* of revisions, not formatting
  or style changes.
- A field's cached result is the text in the `final` and `original` views, with nothing to
  tell it from typed text; only `view="markup"` marks it, inline
  (`<!-- field: REF RefIncident \h -->4 Data retention<!-- /field -->`; at the start of a
  line the mark joins the block's id comment, a field spanning paragraphs -- a table of
  contents -- is marked where it starts, and a field nested in another's instruction is
  not marked). The result is the one cached in the file: `update_fields()` recomputes it.
- A heading's number is written before its text in every view, Word's list number or typed
  text alike; only the id comment tells them apart (`numbered: list`, `numbered: text` for
  text starting with a number such as `4`, `4.1`, `IV.` or `A.`), so with `ids=False` they
  read the same.
- The body only, by default. Headers, footers and the notes nothing in the body references
  are read with `stories="all"` (the stories with content, each after a `<!-- story: ... -->`
  comment) or `stories=[...]`; their comments then show in the markup view where they are
  attached. `doc.comments()` lists every comment, in any story, whatever is read.
- A chart or a SmartArt diagram is a drawing, not a block: its content is one line in its
  drawing's comment (`<!-- d:3 chart "Chart 3" [chart: bar; title: ...; series: ...;
  categories: ...] -->`, `[smartart: Goals [Faster edits], Risks]`, pptx-agent's outline
  line), read-only -- the values are in `state()`'s `chart` JSON, which `Chart.apply`
  takes back -- and `insert_markdown` makes neither.
- There is no `apply_markdown(edited)` that diffs a whole edited projection back into the
  document. An agent edits *by id*: Markdown in for new content, the API for refinement.
  (A diff-and-patch layer was considered and rejected for now: matching blocks without ids
  is guesswork, and with ids it is just the API with extra steps.)

### Dialect and parser

**CommonMark 0.31 plus the GFM table, strikethrough and task-list extensions, plus
footnotes.** CommonMark because it is specified and has a conformance suite; GFM tables
because agents write them constantly.

| Option | For | Against |
| --- | --- | --- |
| **markdown-it-py** (MIT, pure Python) | CommonMark-compliant, the fastest compliant Python parser; GFM tables and strikethrough built in; token stream with source line maps; footnotes and task lists via `mdit-py-plugins` | two to three small runtime dependencies (`mdurl`, plugins) |
| mistletoe (MIT, pure Python) | compliant, tables and strikethrough built in, AST | footnotes not built in; smaller ecosystem |
| marko (MIT, pure Python) | compliant (0.31.2), GFM mode, extensible | slowest of the three |
| our own parser | no dependency | CommonMark's emphasis and list rules are hundreds of spec examples; a second-rate parser is a correctness bug in every insert |

**Recommendation: markdown-it-py**, behind a small internal interface (tokens → our block
model) so it can be swapped, as a runtime dependency of the `markdown` module or an extra
(decided: markdown-it-py). lxml is already a dependency through ooxml-edit; no other dependency is
proposed. The *writer* (`to_markdown`) needs no library.

---

## Reflow feedback

The layout comes from docx2svg (`convert_docx` over the saved bytes, which lays the
document out once and returns both the layout of every page and the pages' SVG); docx-agent
maps it to its own ids. Because docx-agent serialises the bytes docx2svg parses, a
`data-docx-path` / `Line.path` names exactly one element of the live tree, found with
docx2svg's exported `docx2svg.paths.resolve_path` (and `path_part` for the part), and
`Line.paragraph_id` cross-checks it. The map is built once per layout.

```python
lay = doc.layout()                     # cached per undo state; ConvertOptions passed through
lay.where("p:0C11D2E7")                # [Placement(page=3, top=412.5, bottom=468.1, column=0,
                                       #            lines=4, story="body")]  (points)
lay.page_of("t:5D10A7E1")              # 3, or a range for a block that breaks across pages
lay.pages_known                        # 12; None past a stop
lay.stopped                            # Stop(page=12, reason="table", at="t:...") or None

before = doc.layout()
doc.anchor("draft").replace("a considerably longer final version")
diff = doc.layout().compare(before)    # Reflow(first_changed_page=3, changed=[3, 4],
                                       #        page_count=(12, 13), moved={"p:...": (4, 5)})
doc.render_svg(pages=diff.changed)     # or render_png; ids on every paragraph group
```

- **Placement** in points (and device pixels on request) per block: page, top and bottom,
  column, line count, and for a block split across pages one placement per page. A run or
  text range resolves to its lines' boxes.
- **Changed pages:** each page gets a signature of what it draws (glyphs, positions, ids;
  not the SVG text), and two layouts are compared page by page; `moved` lists blocks whose
  page changed. Pages are compared as content, so an edit that only changes `NUMPAGES` in
  every footer correctly marks every page changed — and says why.
- **Renders** carry `data-docx-agent-id` on each paragraph, table and drawing group
  (added to docx2svg's SVG by docx-agent, from the same map), so the agent can point at what
  it sees, as with `data-pptx-id`.
- **Stops are carried through.** If docx2svg stops on page 12, `where()` of a block past it
  is `Unknown(after_page=12, reason=...)`, `compare` reports pages past either layout's
  stop as unknown, and nothing is extrapolated.
- **Built in E0** (`docx_agent.layout`): `where` and `page_of` for paragraphs and tables
  in every story docx2svg draws (body, headers, footers, notes; a header is on every page
  it is drawn on), one `Placement` per page and text column (`Line.column`; 0 in a section
  of one column and for a header's or footer's line); a block docx2svg draws nothing of is `Unknown("not-drawn")` (a vertically merged cell's
  continuation, a hidden paragraph); `pages_known` is `None` after a stop and `stopped`
  says where (`Stop(page, reason, at=<id>)`); `compare` signs each page by what it draws
  (lines keyed by id through the session's aliases, glyphs, positions, rules, pictures,
  floats, the stop), and says why a page changed: `content`, `flow` (other blocks on it),
  `stories` (only a header, footer or note differs) or `added`/`removed`. Held to an
  independent SVG diff: a one-word edit on page 9 of a 36-page fixture changes exactly page
  9, an insertion's reflow exactly the pages it reaches. `ids_on_page(n)` lists what is on
  a page; `unmapped` lists any line path no paragraph answers to (none on any fixture).
  Renders carry `data-docx-agent-id` on paragraph groups; tables and drawings, which
  docx2svg draws without a group of their own, wait for E5.
- **Cost.** One full layout per document state, cached by the bytes' hash: `layout()` and
  `render_svg()` of the same state read one `docx2svg.convert_docx` (every page laid out
  and written as SVG; writing the SVG costs little beside the layout), tested by counting
  the calls. Incremental layout
  (re-lay from the page before the first changed block) is a docx2svg proposal for later,
  not a requirement: page-number fields and footnotes make "earlier pages are unchanged"
  false often enough to need care.

---

## A full-state SVG? No. A structured state, yes

pptx-agent's E3 puts the whole slide state on SVG attributes and reads edits back from it.
**For Word it is the wrong vehicle**, and the argument is structural:

- A slide is a closed canvas; a shape is drawn once, in one place. A paragraph is split
  across lines, columns and pages; a table across pages with repeated header rows; a
  footnote is drawn on whatever page its reference lands. There is no one SVG element to
  hang a block's state on, and per-page SVGs would have to agree about blocks they share.
- The SVG is a *consequence* of state the agent changes, and a change to page 2 changes
  pages 3–40. Applying edits made to page 3's SVG while page 2 is being edited is a merge
  problem with no good answer.
- Agents read text. A 40-page document as SVG attributes is an order of magnitude more
  tokens than the same document as Markdown with ids.

What *is* useful, and is planned (E2, read-only first; the read-only state is built, see
*Phase E2*):

- **Markdown with ids** — the reading layer above.
- **A structured JSON state** (`doc.state(range)`): the block tree with ids, declared and
  effective styles and properties, runs with formatting, lists, tables as grids with merges,
  revisions and comments with their ranges, drawings with their placement, and the layout's
  placements beside them. It is what an agent inspects when Markdown is not enough ("why is
  this paragraph 11 pt?"), and it shares pptx-agent's JSON conventions (`"t"`, `"b"`,
  absent means inherited). Built as `doc.state(range, *, view, layout, xml)`, schema
  `docx-agent/state` version 1, documented field by field in `docx_agent/state.py`.
- **A raw-XML floor in that JSON**, per block on request (base64 of the block's XML), so
  anything the vocabulary does not know can be inspected; *writing* raw XML back
  (`apply_state`) is deferred until a real need shows up — the semantic API and Markdown
  cover authoring, and an applied raw floor is exactly the security surface pptx-agent's
  `fullstate/safe.py` had to build.
- **The SVG carries ids, not state** (`data-docx-agent-id`), so a picture is addressable.

---

## Phase E0 — foundation: package, round trip, ids, undo, render (L)

**Prerequisite, outside this repository: extract `ooxml-edit`** (M) -- **done**. pptx-agent's
`core/` moved, with its history, into uvrt/ooxml-edit as `ooxml_edit.opc`, `ooxml_edit.xml`,
`ooxml_edit.history` and `ooxml_edit.stamp`; pptx-agent depends on it and keeps
`pptx_agent.core` as a re-export shim. pptx-agent's whole suite and its PowerPoint oracle
pass on ooxml-edit, with every fixture's output byte-identical per part to before, and the
neutrality test moved with the code (`tests/test_neutrality.py`, widened from PowerPoint's
vocabulary to every format's). docx-agent builds on it and does not fork it.

**Here:**

- `oxml/`: WordprocessingML namespaces (`w`, `w14`, `w15`, `w16cid`, `w16cex`, `wp`,
  `wp14`, `a`, `pic`, `r`, `mc`...), registered with ooxml-edit; **child-order tables** for
  `w:pPr`, `w:rPr`, `w:tblPr`, `w:trPr`, `w:tcPr`, `w:sectPr`, `w:settings` (all strict
  sequences in the schema, `settings` notoriously long) and the block containers;
  `WordPackage` (main part, stories, styles, numbering, settings, notes, comments parts;
  `.docx`, `.docm`, `.dotx`, `.dotm` content types).
- `Document.open/save`: **every part byte-identical** on an unedited round trip.
- The views needed to address things: stories, blocks (through `w:sdt`/`w:customXml`),
  paragraphs, runs, tables as grids.
- Ids: paraId derivation, positional fallback, duplicate handling, lazy stamping, aliases;
  the other id forms in *Addressing*.
- Undo/redo/nested batches from ooxml-edit: undo of anything is the original bytes.
- `paragraph.set_text` (formatting-keeping) and `delete_block` — the two edits E0 needs to
  test ids and undo meaningfully.
- Render and layout: `render_svg`/`render_png` through docx2svg (a hard dependency, by
  decision 9; PNG through its `png` extra), `layout()` with `where`, `page_of`, `compare`,
  and ids on the SVG.
- **Measured before stamping is relied on** (Word probe, generated by a committed script):
  which paraIds Word keeps on save — absent, in range, stamped by us, out of range,
  duplicated, with and without `textId`, after the paragraph is edited in Word, after a
  comment is added — and whether it renumbers note, comment and revision ids. The results
  decide the way on for durable ids and are recorded here with their evidence.

**Done when:** every fixture round-trips byte for byte; stamping, then undo, gives the
original bytes; ids resolve after unrelated inserts and deletes, through undo and redo; the
validity checks pass after E0's edits on every fixture; a rendered page carries the API's
ids on every paragraph docx2svg draws; `compare` flags exactly the pages a one-word edit
changed on a multi-page fixture; Word exports every fixture with E0's edits unprompted; and
the paraId measurement is in this file.

### E0 — done (2026-10-03)

Every "done when" holds, on all thirteen fixtures (docx2svg's eleven, two of our own):

- **Package.** `docx_agent.oxml`: Word's namespaces and the child sequences of `w:body`,
  `w:p`, `w:r`, `w:pPr`, `w:rPr`, `w:sectPr`, `w:tbl`, `w:tblPr`, `w:tr`, `w:trPr`, `w:tc`,
  `w:tcPr`, the story roots, `w:settings` and `w:style`, registered with ooxml-edit;
  `WordPackage` finds the main part, the story parts (headers, footers, notes, comments),
  styles, numbering, settings and the compatibility mode. Open and save is byte-identical
  per part, in entry order; reading anything dirties nothing.
- **Views** (`docx_agent.edit`): `Story`, `Paragraph`, `Run`, `Table` (a grid resolving
  `w:gridSpan` and `w:vMerge`, rows and cells through row-level content controls), `Row`,
  `Cell`, `Section`. A view holds an id and resolves it per access. Text is read across
  runs, hyperlinks, fields (their results), simple fields, content controls and revisions,
  in the `current`, `original` and `markup` views, by the rules of *Text ranges*.
- **Ids** as *Addressing* specifies, for paragraphs, runs, tables, rows, cells and
  sections; the other kinds wait for the phases that edit them. Stamping is lazy and, after
  the measurement, whole-document by default (*Durable across a Word save — measured*).
- **Edits.** `paragraph.set_text` is an in-place diff: surviving characters stay in their
  `w:t`, a replacement takes the formatting of what it replaces and an insertion that of
  the character before it (pptx-agent's E1 rule; E1 adds ranges and find/replace on the
  same text map), tabs and breaks written as elements, a new textId as Word gives one.
  `insert_paragraph(after=|before=, style=)` continues the paragraph beside it (its
  properties without a section break or revision record, its nearest run's formatting).
  `delete_block` deletes a paragraph or table, moves a lone range marker to the next
  paragraph, and refuses -- before anything changes -- a section break, a note or comment
  reference, half a field, or leaving a cell or story without a final paragraph. An edit
  never changes the compatibility mode. Each is one undo step; undo gives the original
  bytes, redo the edited ones, a failed batch rolls back.
- **Validity** (`docx_agent.validate`): the checks of *Testing strategy* layer 2 that E0's
  edits can disturb; no edit adds a problem on any fixture.
- **Render and reflow** (`docx_agent.layout`): see *Reflow feedback* for what is built and
  where it stops.
- **The oracle** (`tests/oracle.py`, `pytest -m oracle`): every fixture with E0's edits
  exports from Word unprompted with the edited text in its PDF, and a Word re-save keeps
  every paragraph's text, the compatibility mode and every paraId docx-agent wrote.

**What Word changes on re-saving an edited fixture** (recorded by `test_oracle.py`): it
rewrites every XML part (rsids added, its own namespace set and formatting); adds
`docProps/app.xml`, `docProps/core.xml`, `fontTable.xml`, `webSettings.xml` and a theme
where a fixture had none, `commentsExtended.xml`/`commentsIds.xml` beside comments and an
`endnotes.xml` beside footnotes; drops `docProps/thumbnail.jpeg` and Word 2010's
`stylesWithEffects.xml`; renumbers note, comment, revision and bookmark ids; writes new
paraIds into footnote separators. It keeps the compatibility mode (12, 14, 15), every
paragraph's text and every other paraId.

**Measured on the way, beyond paraIds:** Word would not open a first version of our
generated fixtures (the open blocked on a dialog) -- a style's `w:next` after `w:qFormat`,
and settings naming footnote separators of a package without footnotes; `w:style`'s
sequence is now registered so the validity checks catch the first. Word drops a trailing
space of a `w:t` without `xml:space="preserve"`, as the specification allows.

**Proposals for docx2svg** -- all four **done** in docx2svg, and docx-agent moved onto them:

1. *Export the path function* -- done (docx2svg `8b224a6`: `docx2svg.paths.resolve_path`,
   `path_part`, `locate`); docx-agent's mirrored `resolve_path` is gone.
2. *One call for layout and SVG* -- done (`49974fd`: `docx2svg.convert_docx` returns a
   `Conversion(layout, svgs, page_numbers)`); `layout()` and `render_svg()` share one.
3. *A column per line* -- done (`8b66b29`: `Line.column`); `Placement` is per page and
   column.
4. *The machine-wide Word lock* -- done (`ae0eb95`: docx2svg takes `word-oracle.lock`,
   blocking for up to 600 s, around its Word work, and never quits a Word it did not
   start).

No change to ooxml-edit was needed.

## Phase E1 — the semantic API (L)

Text ranges and anchors; formatting (styles first, direct second, effective reads through
docx2svg); lists and numbering; find and replace across runs; hyperlinks; bookmarks and
cross-reference fields; inline pictures (insert, replace, alt text, delete with reaping);
`coalesce_runs`. All built on the primitives, designed so E3 can track them.

**Done when:** each operation, on every fixture, passes edit → save → reopen → read back,
the validity checks (numbering integrity included), undo to the original bytes and redo;
find/replace is held to an independent model over a randomised sweep of matches crossing
run, hyperlink, revision and field boundaries; and Word exports every fixture with all of
E1's edits unprompted, the replaced text present in the PDF's text.

### E1 — done (2026-10-03)

Every "done when" holds, on all fourteen fixtures (E0's thirteen and `lists-and-styles.docx`,
generated: a template localised to Dutch, lists sharing a definition and restarted, a list
from a style, mixed formatting split by rsids and a proofing mark, links, a bookmark, a
picture):

- **Ranges and anchors** (`edit/ranges.py`, `edit/inline.py`): `p:<id>@<start>:<end>`,
  positions `p:<id>@<n>`, ranges across paragraphs `p:<id>@<a>..p:<id>@<b>`, each in a
  view; `find` (case, whitespace collapse and regex explicit; across paragraph ends on
  request), `anchor` refusing ambiguity with every candidate's id, `replace` (a whole
  document's in one undo step, groups expanded), `insert_text`, `delete_text`. A
  replacement takes the formatting of the match's first character (`keep="first"`, Word's
  rule) or keeps each surviving character's (`keep="characters"`, `set_text`'s diff);
  text inserted beside only a field's result gets a run of its own outside the field; a
  span that cuts a field, or would delete a note reference, is refused. A range across
  paragraphs is deleted and the first paragraph takes in the rest of the last (its id and
  properties kept); across a table, a content control or a section break it is refused.
- **Formatting, styles first** (`edit/styles.py`, `edit/builtin_styles.py`,
  `edit/formatting.py`, `edit/formatops.py`): paragraph and character styles by `w:name`
  (any case), alias or id -- "Heading 1" finds a Dutch template's `Kop1`; a built-in style
  the document lacks is written from the measured table (below) with what it points at;
  `Styles.add`/`modify`. Direct run formatting (bold, italic, underline, strike, double
  strike, caps, small caps, size, font, colour -- theme colours as `w:themeColor` --
  highlight, vertical alignment) and paragraph formatting (alignment, indents, spacing,
  line spacing, keep with next, keep together, page break before, widow control), checked
  before anything changes; a range is split at its edges and what comes out alike merged
  back, so no edit leaves an empty or redundant run. `clear_direct_formatting` keeps
  styles, list membership, language and revision records. Reads are declared (`None` when
  inherited); `effective` asks docx2svg's resolver over the current bytes (one parse per
  document state) and `explain()`s each value's origin.
- **Lists** (`edit/numbering.py`): `add_to_list` (Word's own bullet or number list with a
  definition and `w:nsid` of its own, or joining the list just before, or a numId), levels,
  `restart_numbering` (a new `w:num` over the same definition with `w:startOverride`, given
  to the list's items from there on), `continue_numbering`, `remove_from_list` (`numId` 0
  where a style lists the paragraph), `set_list_format` copy-on-write when another
  instance shares the definition. An instance or definition an edit leaves unused is
  removed; ones the document had unused are left. A Normal paragraph made a list item gets
  List Paragraph, as Word gives it (measured), and loses it again on leaving the list.
- **Hyperlinks, bookmarks, cross-references** (`edit/links.py`): external links reuse a
  relationship to the same address and release it when unused; internal links need their
  bookmark; links get the Hyperlink style; change and remove. Bookmarks by name (valid
  names only, unique regardless of case, `w:id` one above every annotation id), renamed
  with the hyperlinks and `REF`/`PAGEREF`/`NOTEREF` instructions that name them, removed
  (what still points at them in `result.warnings`), resolved to a range. `REF` and
  `PAGEREF` fields inserted as Word inserts them, the result computed: the bookmark's text,
  or its page from docx2svg's layout (dirty where the layout cannot say).
- **Pictures** (`edit/pictures.py`): inline pictures from bytes or a file at their natural
  size, `docPr` ids one above the document's largest, `wp14:anchorId`/`editId` stamped;
  media shared by content, reaped through ooxml-edit's `release` when unused; replace
  (keeping the frame, its width, or natural size), resize, alt text, delete.
- `coalesce_runs` (runs differing only in `w:rsid*`, on request only), `move_block` within a
  story keeping ids, `append_paragraph` for a story with no paragraph to insert beside.
  `EditResult` gains `count` and `warnings`. Every operation is one undo step; no edit
  changes the compatibility mode.
- **Validity** gained numbering integrity (numIds resolve, instances' definitions exist,
  levels 0-8, `w:nsid` unique), relationship ids per part, style references and `basedOn`
  cycles, unique `docPr` ids and valid bookmark names.

**Tests.** 728 in the default run: each of 17 operations on each fixture through edit,
save, reopen, read back, validity and numbering integrity, undo to the original bytes and
redo; find and replace against an independent model over a randomised sweep (480 spans in
generated paragraphs of formatted runs, hyperlinks, insertions, deletions, complex and
simple fields, proofing marks and bookmarks: 433 replaced as the model says, 47 refused for
cutting a field, nothing else changed); mixed formatting kept by `set_text` and replace on
every fixture; every fixture rendered by docx2svg after an E1 edit set with the API's ids
on every paragraph and no new warning but the cached `REF` the edits add; a style change
reflows from its page, a same-length replacement changes only its page, a picture pushes
the text below it down.

**The oracle** (`pytest -m oracle`, 56 tests, all passing): every fixture with E1's edit
set (`tests/e1_edits.py`) exports from Word without a prompt; its PDF has the replaced
word, the heading, the list numbers `1.`, `2.` and `7.` (the restart), the bullet's
cross-reference; pdfium reads the heading at 13 pt over the body's 11, the direct
formatting at 20 pt in C00000, and Strong bold where the resolver says it is (Word draws
bold in a face without a bold variant, Calibri Light, by stroking the glyphs). A re-save
keeps every paraId and paragraph's text (E0's guarantee holds after E1's edits), the lists
and their restart, every hyperlink, bookmark and picture (`docPr` ids too).

**What Word changes on re-saving E1's edits** (Word 16.106, Dutch interface; printed by
`test_oracle.py`): it **renames every style id into its interface's language** --
`Heading2` → `Kop2`, `Normal` → `Standaard`, `ListParagraph` → `Lijstalinea`, `Strong` →
`Zwaar` -- with every reference, and the *names* of linked character styles ("Heading 2
Char" → "Kop 2 Char"); `w:name` of built-in styles is kept, which is why styles are found
by name. It keeps every style property but refreshes a theme colour's cached `w:val` from
the document's theme (`2F5496` → `0F4761` where the document had no theme, Hyperlink's
`0563C1` → `0000FF` under a 2010 theme). It keeps our list definitions verbatim (`nsid`,
`tmpl`, levels), adding `w15:restartNumberingAfterBreak="0"` to each and a
`w16cid:durableId` to each instance. It adds no bookmark and keeps our `docPr` ids.

**Measured for E1** (`tools/e1_probe.py`, `tests/observations/e1-word.json`): the
definitions Word writes on a built-in style's first use, for 35 paragraph and 12 character
styles, with the linked character styles and, for List Bullet/Number 1-5, the single-level
list whose level names the style (`src/docx_agent/edit/word_styles.py`, keyed by name,
references by name, rsids dropped); that Word gives a listed Normal paragraph List
Paragraph; its default lists' first level (Symbol U+F0B7 or `%1.`, 720/360 twips); a
picture stating no density placed at 96 dpi, one stating 300 dpi at 300.

**Approximations, and why:**

- *Built-in definitions are one Word's.* Word 16.106 with its 2023 Office theme; the
  header and footer styles' tab stops (4536/9072 twips, 8 and 16 cm) are this Word's
  locale's, an English one writes 4680/9360. A theme colour's cached `w:val` is the probe's;
  Word refreshes it on save. New styles get English Word's ids.
- *New lists are nine-level.* Word's AppleScript default lists are single-level; ours are
  the `hybridMultilevel` nine levels Word's Bullets and Numbering buttons write, whose first
  level matches what was measured and whose deeper levels (0.5 inch apart; `o` Courier New
  and Wingdings bullets; `a.` and `i.`) are not measured here. Word keeps them verbatim.
- ~~*A whole-paragraph join keeps the first paragraph's properties.*~~ Measured in E3: Word
  keeps the *second* paragraph's element and paraId and gives it the first one's
  properties; the join does that now (*Accept and reject*).
- *Text typed at a hyperlink's end joins the link* (it takes the character before it);
  Word's typing rule there is not measured.
- *Leaving a list returns List Paragraph to the default style*, undoing what making it a
  list item did; not measured.
- *A stated picture density is rounded to whole dots per inch*, which reproduces the
  measured 300 dpi; other densities are not measured.
- *Effective values for a paragraph docx2svg does not lay out on its own* (one joined across
  a deleted mark, one in a text box) are resolved without its table style's conditional
  formats.
- *Undo depth* is ooxml-edit's default, 50 steps.

## Phase E2 — Markdown out and in (M–L)

`to_markdown` with ids and views; the JSON state (read-only); then `StyleMap`,
`insert_markdown` (blank documents and templates, inline and block), creating mapped
styles that are missing. **Done** (2026-10-03 and -04): both halves below.

**Done when:** `to_markdown` changes no byte; the AST round trip holds for a Markdown corpus
built from the CommonMark spec's examples of the supported constructs plus hand-written
documents, in a blank document and every template fixture; inserted content uses only the
document's styles (checked by listing every `pStyle`/`rStyle` written); a Dutch-localised
template maps headings by name; Word exports every result unprompted; and with tracking on (after E3) the insertion is
one revision.

### E2, the read half — done (2026-10-03)

`to_markdown` and the JSON state; the write half (`insert_markdown`) follows E3, by the
suggested order, and reads the same `StyleMap`. What "done when" asks of the read half
holds: `to_markdown` changes no byte, on every fixture, in every view, with the state too;
the reverse-parse half of the AST round trip holds everywhere; a Dutch-localised template
maps headings by name and by outline level. The AST round trip through `insert_markdown`,
"inserted content uses only the document's styles" and Word's export of inserted content
are the write half's.

- **`to_markdown(range=None, *, view="final", ids=True, headers=False, notes=True,
  style_map=None)`** (`docx_agent.markdown`): CommonMark with GFM tables, strikethrough and
  footnotes. `range` is the body (default), a story (`header1`, `footnotes`), a section
  (`s:<id>`, `s:body`), a block id (a paragraph in a table names its table, `cc:` a
  control) or `<id>..<id>`; notes referenced in it come with it. `ids=False` writes plain
  Markdown.
- **Ids.** A block's comment sits on the line above it (`<!-- p:3A1F09C2 -->`); a list
  item's and a note paragraph's at its end (`- Item <!-- p:... -->`, where a leading
  comment would make the line an HTML block); a code block's lists one id per line. The
  comment also says `volatile` for a positional or repeated id (reading never stamps),
  `joins <ids>` for paragraphs the view joined, `style: <name>` for a style no rule names,
  `empty` for a paragraph with no text, and takes in a bookmark or page break standing at
  the start of the line. A GFM table's cells are named by the ids its own implies
  (`t:<id>/c<r>,<c>`); the HTML fallback writes `data-id` on the table and every cell and
  the paragraphs' comments inside. Every id written resolves through `Document.get`, which
  now also answers `fn:`/`en:` notes, `c:`/`c#` comments, `rev:` revisions, `cc:` content
  controls, `d:` drawings that are not pictures, `hl:`, `bm:` and story names
  (`edit/annotations.py`, read-only views for E3 and E4 to make editable).
- **The mapping** (`markdown/stylemap.py`, `StyleMap`: one declarative table of `Rule`s,
  `construct` -> Word style by `w:name`; `reads` are names read as the construct but never
  written; `with_rules` overrides rows; `style_for` is what the write half will apply):
  - *Headings* by style name (`heading 1`-`heading 9`, any case), else by the outline
    level the style declares through its `basedOn` chain -- "Kop 1" with `w:outlineLvl` 0 is
    `#`. Levels 7-9 are `######` with `level 7` and the style said. A numbered heading
    keeps its number as text.
  - *Paragraph*: Normal, and read also Body Text, List Paragraph, and the story styles
    (footnote text, header...). Anything else is a paragraph with its style named.
  - *Lists* from list membership, not style names: a level's `numFmt` `bullet` is `-`, any
    other a numbered item, `none` no list; nesting from `w:ilvl`, relative to the list's
    first level. **Numbers are Word's**, counted by docx2svg's measured counter
    (`ListCounters` over `parse_numbering`; its ROADMAP, "Numbering and lists --
    measured", N.1-N.4): definitions sharing a `w:nsid` are one list under the first's
    levels, instances over one definition count on from each other, a `w:startOverride`
    restarts at the instance's first item at its level (at an overriding `w:lvl`'s
    `w:start` when there is one) and is the level's start there, `w:lvlRestart` and
    `w:isLgl` are honoured, a shallower level with no count is set to its start and
    counted as used, and text boxes are a story of their own. `tests/test_numbering_counts.py`
    holds the state's labels and the Markdown's numbers to Word's on docx2svg's probe
    (its recording, read from its checkout), in all three compatibility settings. A
    different instance or kind at the same level starts a new Markdown list, kept apart by
    the other marker (`*`, `)`). List Bullet/List Number without numbering are still lists.
  - *Quote* (and Block Text) is `>`; *HTML Preformatted* (and Source Code) a fenced code
    block, consecutive paragraphs one block; an empty paragraph with a bottom border `---`.
  - *Emphasis and strong are read from the effective formatting* (docx2svg's resolver),
    against the paragraph's baseline -- what a run with no properties of its own gets
    there. Decided: **direct `w:b`/`w:i` read exactly as the Strong and Emphasis styles
    do**, so `**x**` means "bold beyond its paragraph" however it was made (the state says
    which: a run's `style` and `direct`). A heading's bold or a quote's italic is not
    emphasis. Strikethrough (`w:strike`, `w:dstrike`) is `~~`; HTML Code (and Verbatim Char)
    is inline code.
  - *Links*: external `[text](url)`, internal `[text](#bookmark)`; *bookmarks* as
    `<!-- bm:name -->` where they start (hidden `_` ones only when a link targets them).
  - *Notes*: `[^fn:1]`/`[^en:1]` with GFM definitions after the body, in first-reference
    order, a note's paragraphs indented.
  - *Pictures*: `![alt](d:7 "name")`, inline or floating; other drawings
    `<!-- d:5 chart "Chart 2" -->`; a text box's paragraphs follow its paragraph between
    `<!-- d:3 text-box -->` and `<!-- /d:3 -->`.
  - *Fields*: the cached result is the text. *Content controls*: their content between
    `<!-- cc:301 text -->` and `<!-- /cc:301 -->`, inline or as blocks (kind, tag, title).
    *Breaks*: a line break `\` + newline, a page or column break `<!-- page break -->`, a
    section break `<!-- s:<id> section break: continuous -->` after its paragraph.
    *Headers and footers* after the body with `headers=True`, each under
    `<!-- story: header1 (header) -->`.
  - *Tables*: GFM when every row fills the grid without merges and every cell holds at
    most one plain paragraph with no line break; the first row is the header (GFM needs
    one), a column whose cells all centre or right-align says so. Otherwise HTML: `rowspan`,
    `colspan`, nested tables, several paragraphs per cell.
  - *Escaping*: `\`, `` ` ``, `*`, `_`, `[`, `]`, `<`, `~` always; `|` in cells; `&` before an
    entity; `{` before CriticMarkup; at a line's start what would open a block (`#`, `>`, `-`,
    `+`, `=`, `1.`). Each block's inline Markdown is parsed back before it is written; if
    CommonMark's flanking rules would lose a mark next to punctuation inside a word,
    strikethrough and then emphasis are dropped from that block, never its text.
- **Views.** `final` is docx2svg's measured final view: insertions and moves' destinations
  in, deletions and sources out; a paragraph whose mark is deleted or moved away joins the
  next (the first one's id, the last one's style and list); an empty one before a table
  goes; a deleted row goes. `original` mirrors it (an inserted mark joins, an inserted row
  goes). `markup` shows both with CriticMarkup and ids, comments as
  `{>>Bob, in reply: Yes.<<}<!-- c:2B3C4D5E -->`, formatting changes as
  `<!-- rev:113 formatting Bob -->` after the run and paragraph-property changes on the
  paragraph's comment.
- **The reverse-parse check** (`markdown.check`, `parse.py`): the model, the renderer and
  markdown-it-py meet in one normalised AST (blocks; inline text with marks and link
  targets; comments; images; footnote references), forgetting only what Markdown does not
  keep (markers, delimiters, whitespace at line ends, tightness). Every fixture, every view,
  with and without ids, after E1's edit set too, reads back as the model it was written
  from. `parse.py` is also the "tokens -> our block model" interface `insert_markdown` will
  consume.
- **The JSON state** (`docx_agent.state`): see *A structured state*. Blocks (paragraphs,
  tables as grids with spans, block content controls) with ids, styles by name, the
  construct, list place and Word's number, the text, effective formatting (the paragraph's,
  and each run's difference from its baseline), links, bookmarks, notes, comments,
  revisions, pictures, drawings, inline controls, fields, text boxes, section breaks;
  notes, comments (author, date, parent, resolved) and revisions (kind, author, date,
  paragraph or table) of the range; docx2svg's placements with `layout=True`; each block's
  XML in base64 with `xml=True`.

**Tests** (232 new; 960 in the default run): reading changes no byte -- every XML part's
tree, `changed_parts()`, history and aliases checked after every view of `to_markdown` and
the state -- on all seventeen fixtures (the corpus and E2's three: `generated/markdown/`,
below the corpus so E0's and E1's suites and the oracle keep theirs; not opened in Word);
the reverse-parse check and id resolution on each in each view; **the final view's text is
docx2svg's drawn text, paragraph by paragraph**, on every fixture docx2svg lays out (body
and notes; whitespace, soft hyphens and object marks aside, note marks matched as numbers,
paragraphs with the fields docx2svg computes -- PAGE, NUMPAGES, SECTIONPAGES -- left out);
per construct on `constructs.docx`, `review.docx` (two authors' revisions of every kind,
comments with a reply and a resolution, `commentsIds` durable ids) and
`dutch-template.docx`; the state's ids, styles, lists, spans, revisions, comments,
placements and raw XML. Word was not used.

**Non-guarantees of the read half**, besides *Guarantees, and explicit non-guarantees*:
the original view keeps today's formatting, style and list membership (it undoes the text
of revisions; property changes are E3's); effective values of a paragraph docx2svg does not
lay out on its own are resolved without its table style's conditions (E1's
approximation); a joined paragraph's emphasis is measured against each part's own
paragraph; hidden text is left out.

**Proposals for docx2svg** (not made here; docx2svg's to decide):

1. *`w:startOverride` is not applied.* `parse/styles.py`'s `parse_numbering` keeps a
   `w:lvlOverride`'s `w:lvl` but not its `w:startOverride`, so docx2svg draws a restarted
   list from 1: `lists-and-styles.docx` after E1's `restart_numbering(at=7)` is drawn `1.`
   where Word's PDF shows `7.` (E1's oracle), `constructs.docx`'s restart at 5 is drawn
   `1.`. Same visible class as the final-view fixes.
2. *Counting per instance.* `ListCounters` counts per `numId`; Word counts instances over
   one abstract definition on from each other unless one restarts (E1's model; a probe
   with two instances over one definition and no override would settle it). docx-agent's
   reader counts Word's way.

Both **done** in docx2svg (`aa98caf`, measured: N.1-N.4), which also found docx-agent's
own model wrong in seven places (`w:nsid` sharing, a `w:lvl` with a `w:startOverride`,
an override's level restarted, `w:lvlRestart`, unused shallower levels, `w:isLgl`, text
boxes); the reader now counts with docx2svg's `ListCounters`, and `Numbering.level`
takes a level from the first definition with its `w:nsid`.

### E2, the write half — done (2026-10-04)

`insert_markdown`, after E3 as the suggested order has it, reading the read half's
`StyleMap`. Every "done when" of the phase holds: the AST round trip over the corpus in a
blank document and every fixture (the read half's 17 and the blank one), only the
document's styles, the Dutch template's headings by name, Word's export of every result,
and one revision group with tracking on.

- **Measured first** (`tools/e2_probe.py`, Word 16.106 for Mac, Dutch interface, under the
  machine-wide lock): Word's **Table Grid** as it writes it on first use (set by its
  localised name; Word's dictionary has no constant for it) -- now
  `edit/word_table_styles.py`, beside E1's and E3's measured styles -- and the table it
  makes (`w:tblW` auto, `tblLook` `04A0`, the text width shared among the columns: 9026
  twips as 3008/3009/3009); and a document's **first footnote**: Word writes
  `footnotes.xml` *and* `endnotes.xml`, each with its separator (`w:id` -1) and
  continuation separator (0) in paragraphs spaced 0 after, single, the settings'
  `w:footnotePr`/`w:endnotePr` naming both, and the note as a footnote-text paragraph
  starting with the footnote-reference run and a space. Not measurable this way: the
  border Word's AutoFormat makes of a `---` line (AppleScript's typing does not trigger
  AutoFormat), so the thematic break is its documented form (below).
- **`Document.insert_markdown(md, *, at="end", style_map=None, images=None, fetch=None,
  html="refuse", emphasis="styles", track=None)`** (`markdown/write.py`). `at`: `"end"` or
  `"end:<story>"` (a header, a footer), `"after:<id>"`, `"before:<id>"` (a bare id is
  after it), `"replace:<id>"` or `"replace:<id>..<id>"` (blocks of one container; a
  paragraph in a cell is replaced in its cell). The result is one undo step: `created`
  lists every new paragraph and table id (cells' paragraphs and notes' too, and `fn:<n>`
  for each note), `blocks` the top-level ones in order, `removed` what a replacement
  deleted, `warnings` each style added. Nothing to write (only link definitions, comments,
  an empty quote) changes nothing (`changed=False`).
- **The mapping, written** (the read half reads each back as the same construct):

  | Markdown | Written | Read back from |
  | --- | --- | --- |
  | paragraph | the default paragraph style, without `w:pStyle` (as Word writes Normal) | Normal, Body Text, List Paragraph, a story's own styles |
  | `#`-`######` | the style named `heading N`, else a style whose outline level is N-1 (through `basedOn`), else Word's built-in | the name, else the outline level |
  | `-` list, nested | List Bullet, List Bullet 2-5 (deeper: 5) with its own `w:numPr`: one `w:num` per Markdown list, `w:ilvl` its depth | list membership, `w:ilvl` |
  | `1.` list, start | List Number, List Number 2-5; the instance's `w:lvlOverride/w:startOverride` at its level is the start, so every list restarts | membership; the number Word counts (docx2svg's `ListCounters`) |
  | the lists' definition | per insertion and kind, a copy (new `w:nsid`) of the document's own nine-level bullet or numbered definition its paragraphs use most, else Word's own (E1's) -- never one the document's lists count through | -- |
  | `>` | Quote (each paragraph of the quote) | Quote, Block Text |
  | fenced, indented code | HTML Preformatted, a paragraph per line | consecutive HTML Preformatted (Source Code) paragraphs |
  | `---` | an empty paragraph with `w:pBdr/w:bottom` single, `sz` 6, `space` 1, `auto`: Word's AutoFormat of `---`, not measured | an empty paragraph with a bottom border |
  | GFM table | `w:tbl` in the style the document's tables use most, else Table Grid (measured); `tblW` auto, `tblLook` `04A0`, the section's text width shared; header row `w:tblHeader`; alignment the cells' `w:jc` (left written, read as none) | a grid without merges, one paragraph a cell |
  | `*em*`, `**strong**` | Emphasis, Strong | the effective italic or bold beyond the paragraph's |
  | `` `code` `` | HTML Code | HTML Code, Verbatim Char |
  | two or more marks on a run | one character style -- code, then link, then strong, then emphasis -- and the rest direct `w:b`/`w:i` (a run has one style); `emphasis="direct"` writes all direct | as above |
  | `~~strike~~` | direct `w:strike` | `w:strike`, `w:dstrike` |
  | `[t](url)`, `<url>` | `w:hyperlink` with an external relationship, Hyperlink style | the hyperlink |
  | `[t](#name)` | `w:hyperlink w:anchor="name"` | the anchor |
  | `![alt](src "title")` | an inline picture, natural size (capped to the text width), `descr` the alt text, `name` the title | `![alt](d:<id> "name")` |
  | `[^label]` | a footnote: footnote text, the reference run in footnote reference (as Word writes it), parts made as measured | `[^fn:<id>]` with its definition |
  | hard break | `w:br` | `\` at a line's end |
  | raw HTML | refused (`html="refuse"`), or its text (`"text"`); comments (ids) dropped | -- |

  Footnote paragraphs are in footnote text (a row the map now has; the reader already read
  it as a paragraph). A mapped style the document lacks is added from Word's measured definition (and what it
  names: a heading's linked character style, a list style's own list) and the result says
  so; "Normal" missing means the document's default (layout-sweep has no styles part).
- **Tracking** (`track=` or `doc.tracking(...)`): the whole insertion is one revision group
  -- every record one author and date. Paragraphs are inserted with their marks
  (`w:pPr/w:rPr/w:ins`), runs in `w:ins` (inside a hyperlink, never around it), rows with
  `w:trPr/w:ins` and their cells' paragraphs; at a container's end in Word's typing form
  (E3's: the mark before the insertion is the inserted one, the last new paragraph keeps the
  old mark with its properties recorded as a change -- from that paragraph's original ones
  when E3 recorded a change on it). A note's content is inserted with its reference, and
  rejecting the reference now removes the note (`review.py`, as a comment goes). A
  replacement deletes the replaced blocks as `delete_block` does and inserts after them;
  untracked, it inserts before and deletes. Tracked, replacing a container's last
  paragraph that follows a table is refused, as `delete_block` refuses it.
- **The round trip's normalisation** (`parse.roundtrip_ast`): the read half's
  (`parse_ast`: markers, delimiters, line-end whitespace, tightness), and HTML comments (dropped
  on writing), a soft break as a space (Word has none), footnote labels by order of first
  reference (written `fn:<id>`), a picture as its alt text and link (its address becomes the
  drawing's id, its title its name), a left-aligned column as unaligned.
- **Validity** gained one check: every prefix a part's `mc:Ignorable` lists is declared on
  its root. Word does not open a part that lists one it cannot resolve: the oracle found
  a notes part listing `wp14` (a picture beside a footnote's reference) undeclared.
- **The read half, fixed on the way:** emphasis whose run holds spaces inside a longer
  emphasis (`*foo **bar** baz*`) was dropped by the self-check (the spaces were moved out
  of every run); a mark spanning a link (`*foo [bar](/url)*`) was closed at it; a `!`
  before a link read back as an image; a no-break space at a line's ends was trimmed by the
  parser (now a character reference); an image's alt text lost its escaped characters
  (`text_special` tokens).

**Tests** (768 new in `tests/test_markdown_write.py`; 2,355 passing in the default run, 160
skipped where a fixture has no such place or table; 8 minutes under `pytest -n auto` on four
cores, about 25 minutes of CPU).
The corpus: 575 of the CommonMark Spec 0.31.2's 652 examples (`tests/corpus/commonmark`,
CC BY-SA 4.0, committed with its attribution, selected by
`tools/select_commonmark_examples.py`: the 77 holding raw HTML -- *HTML blocks*, *Raw
HTML* and 13 others -- left out, and 0 matching the trace check) and six hand-written
documents (`tests/corpus/handwritten`: a report, lists of every shape, tables, inline
syntax, notes, structure). Each example in a fresh blank document (`generated/markdown/blank.docx` until E6; since E6, `Document.new()`: Word's own); in each of the 18
fixtures, every example and document inserted one after another at the end of the body,
each read back over its own blocks -- with every style named defined and the mapped one,
no direct formatting but the decided kind, and no validity problem added. The Dutch
template's headings are Kop1, Kop2 (by outline level) and Kop3 (by name), its quote,
lists, emphasis and strong its own (Citaat, Lijstopsomteken, Lijstnummering, Nadruk,
Zwaar), no heading style added. Undo is one step to the original bytes and redo to the
edited, on every fixture; the result renders through docx2svg with every new paragraph's
id and `where()` finds each block; tracked at the end, after, before, replacing a block, a
range and a story's last paragraph, on every fixture: accept-all is the untracked insertion
and reject-all the original (forgiving, as Word keeps them, the styles added, the lists
those styles carry, and notes parts left with only their separators), every record ours at
one date, other authors' untouched. Units: lists' levels, instances and start override,
the numbers the state gives, a table as Word makes one, notes and their parts, the break's
border, links, pictures from each source and refused outside them, raw HTML, places, the
style map overridden, emphasis direct.

**Outside the guarantee** (59 of the 575 examples, each listed in the test with its reason,
and each asserted still to fail so the list stays exact):

- *A list item holds one paragraph and the lists nested in it* (31: 4, 5, 7, 61, 108, 254,
  256, 258, 262-264, 270, 271, 273, 274, 277, 278, 286-288, 290, 300, 307, 309, 316,
  318-321, 324, 325): a second paragraph, code block, quote, heading or rule in an item is
  written after the list's paragraph and reads back after the list. Word's form for it
  (List Continue, or a list paragraph without a number) would need the reader to read it
  as the item's and a measurement of List Continue; not done.
- *A block quote holds paragraphs* (16: 6, 128, 228-230, 232, 235-237, 250-252, 259, 260,
  292, 293): a heading, list, code block or nested quote in a quote is written as itself,
  outside it; nested quotes are one.
- *Adjacent block quotes, or code blocks, with nothing between* read back as one (242).
- *Empty constructs* (10: 126, 130, 144, 200, 218, 239, 240, 485, 486, 567): an empty code
  block is one empty paragraph, read back as one blank line; an empty quote writes nothing;
  a link with an empty destination is its text.
- *A code span of only whitespace* (334) is whitespace at the paragraph's end.

Raw HTML is refused by design; a footnote referenced twice is two notes (Word's notes have
one reference each); a reference inside a note stays text (Word has no note in a note);
pictures are inline only.

**The oracle** (`pytest -m oracle`, `tests/test_oracle_e2.py`, 73 tests, all passing; the
whole oracle, 205, passes): with a Markdown document of every construct inserted at the
end of each of the 18 documents, Word opens and exports every result unprompted, and its
PDF shows the headings, the list items with the numbers our state gives (`3.`, `4.`, `5.`:
the start override drawn as Word draws it), the table's cells, the code, the quote, the
link text and both footnotes' text; docx2svg's pages have Word's characters, with no
difference the unedited document lacks; a Word re-save keeps the final view's Markdown,
the notes, the pictures, every new paragraph's paraId, text and style; tracked, Word's own
Accept All and Reject All read as ours, but for what is recorded below.

**What Word does otherwise, recorded:**

- *Reject All of paragraphs inserted at a body's end* (two or more): Word leaves the last
  paragraph's property change unrejected -- that paragraph keeps the inserted one's style,
  with a pending `w:pPrChange` -- and the same when *Word* typed the paragraphs
  (`tools/e2_probe.py`, `typing_at_the_end`: Word's form is ours, record for record). Our
  `reject` gives the original. The text is the same.
- *A body with no paragraph at all* (samplelib's blank): the last new paragraph's inserted
  mark is the container's last, which Word neither accepts nor rejects (E3's measurement);
  it stays a revision in Word, text right.
- *A re-save* renames the Dutch template's "Kop 1"/"Kop 2" (ids `Kop1`, `Kop2`: Dutch Word's
  ids for its built-in headings) to "heading 1"/"heading 2"; where the insertion wrote the
  styles part (layout-sweep has none), Word gives the styles its own localised ids and names
  the default paragraph style Normal.
- *A hyperlink inside an insertion* is re-saved as a `HYPERLINK` field (E3's record); its
  Reject All leaves the field's empty skeleton (or an empty hyperlink), invisible.

**Approximations, and why:** a list item's or quote's blocks other than paragraphs are
written beside them (above); a picture is inline at its natural size, capped to the text
width; Markdown heading anchors are not bookmarks (`[x](#name)` links to a bookmark of that
name, which may not exist: Word ignores such a link); the paragraph style for plain text is
the document's default, not Body Text where a template's body uses it (the table's
fallback: not done); a footnote referenced twice is two notes.

**Proposals** (not made here): for docx2svg, a footnote referenced in a table cell is not
drawn (its known gap, which it warns of: `notes.md`'s cell note shows it), and in
layout-sweep (mode 12) the footnote separator is drawn as a line group with
`data-docx-path="w:footnotes/separator"` and no paragraph id -- a group the identity
contract does not expect (docx-agent's render test now skips separator groups); for
ooxml-edit, `release` leaves a relationships part empty when its last relationship goes (a
rejected note's link), where removing the part would be what Word writes.

## Phase E3 — tracked changes and comments (L–XL)

Tracking mode over every primitive; revision listing, accept and reject; comments with
replies and resolution in the modern parts; the measurements (deleted-mark merges, deleting
one's own insertion, cell revisions, what Word's PDF shows); and the docx2svg proposals
(the final view fixed; nothing visualised) handed to that project.

**Done when:** the accept/reject invariant holds for every E1/E2 operation on every fixture;
revisions by other authors in a fixture survive our edits untouched; Word exports every
tracked fixture unprompted, shows our author and date (read from a Word re-save of the
file, see *Testing strategy*), and accepting all in Word gives the same text as our
`accept_all`; comments and replies appear in Word with their threads.

### E3 — done (2026-10-03)

Every "done when" holds, on all fourteen fixtures (E1's), for every E1 operation and E3's
own (E2's write half, `insert_markdown`, is not built yet: it will track through the same
primitives).

- **Measured first** (`tools/e3_probe.py`, Word 16.106 for Mac, Dutch interface, under the
  machine-wide lock): Word itself, tracking, made 21 edits in a text probe and 8 in a table
  probe by AppleScript -- typing over a selection, Backspace, Enter, bold, a style, an
  alignment, `apply number default`, a relocated paragraph, a deleted field, an inserted and
  a resized picture, a hyperlink, a comment, a section margin; rows and columns inserted
  and deleted (a column by cutting it: Backspace opens a dialog), cells merged across and
  down and split, a table's alignment, a cell's shading, a row's height -- and saved; then
  its own Accept All and Reject All of each, saved; the same paragraph-mark deletions
  untracked; and a document written here with the cell forms Word does not write
  (`w:cellIns`, `w:cellDel`, `w:cellMerge`, a span change, `w:hMerge`) and the paragraph
  marks at a container's end, exported, accepted and rejected. What it wrote is the table
  under *How each primitive is expressed*; how its Accept All and Reject All join is under
  *Accept and reject*. Comment styles as Word adds them with a first comment are
  `src/docx_agent/edit/word_comment_styles.py`, generated from that probe.
- **Tracking mode** (`docx_agent.revisions`: `stamp.py` the records' attributes and ids,
  `track.py` the tracked primitives, `mode.py` the mode and the setting, `review.py`
  listing, accepting and rejecting): text insert, delete and replace (and `set_text`'s
  diff), paragraphs inserted and deleted, joins across marks, formatting by snapshot (every
  paragraph and run of the part tagged, the edit run untracked, each change recorded with
  the complete old properties), styles and lists, moves, pictures, hyperlinks,
  cross-references; rows, columns and merges (`edit/tables.py`: also new untracked
  operations, `insert_row`, `delete_row`, `insert_column`, `delete_column`, `merge_cells`,
  and `Table.insert_row`...). Each is one undo step, undone to the original bytes.
- **Comments** (`edit/comments.py`) and **people** as *Edit operations* says; **Word's
  setting** `w:trackRevisions` (`doc.word_tracks_changes`, then `doc.track_revisions`), an empty element in its place in
  the settings' sequence, as Word writes it.
- **Reading.** `to_markdown`'s markup view shows E3's revisions and comments with their
  ids; the original view now shows a paragraph's recorded old properties (its style and
  list from a `w:pPrChange`) and drops an inserted paragraph with nothing of the original
  in it at a container's end; the state lists them all. **Validity** gained revision and
  comment integrity: every record with an author and an id no other has, move ranges paired
  and named on both sides, every comment referenced, the modern comment parts agreeing with
  `comments.xml`, no hyperlink inside a revision container.

**Tests** (633 new; 1,593 passing in the default run, 144 skipped where a fixture has no plain table). The invariant: every E1 operation and
E3's, on every fixture, untracked; the same tracked, saved and reopened; `accept(author=)`
equal to the untracked edit and `reject(author=)` to the original in canonical form
(`tests/canonical.py`: run boundaries, rsids, relationship ids and annotation numbers
aside; paraIds for the operations whose surviving element is the measured one, not the
original); and the whole edit set composed (`tests/e3_edits.py`: every kind of edit, each
on what the ones before made). Every tracked operation on every fixture: saved, reopened,
no validity problem added, every new revision ours at our date, other authors' untouched,
the final view the untracked edit's (but for a deleted column and merged cells: Word's
final view shows those as they were, R.6), the original view the original's text, undo
and redo exact, and accepting it reads back as the untracked edit does. The edit set on
every fixture renders through docx2svg with the API's ids and no new warning, reads back
in every view, its final view what docx2svg draws. And unit tests of what each edit writes,
the filters, partial reviews, moves accepted whole, the join, comments' parts, threads,
resolution, editing and deletion, the setting, and the table operations.

**The oracle** (`pytest -m oracle`, `tests/test_oracle_e3.py`, 76 tests, all passing):
every fixture with the edit set exports from Word unprompted; Word's own Accept All and
Reject All of it read as ours (the final view's Markdown) on every fixture; a Word re-save
keeps our author and date on every revision, every comment, reply, resolution and durable
id, and the final view; docx2svg's drawing of the tracked document has, page by page, the
characters of Word's final-view export, with no difference the unedited fixture lacks.
The structural forms on a table: a deleted column and a merge down accepted and rejected
by Word as by us; a merge across differs both ways (recorded below).

**What Word does otherwise, recorded:**

- ~~*A merge across*: Word's Accept All rebuilds the grid from the cells' widths and leaves
  the merged cell spanning a column more beside a cell it should have replaced; its Reject
  All mangles the row (a `w:gridAfter`, spans of three). Word has no form for it~~ -- found
  in E5: Word widens the cell before a deleted cell whose paragraph marks are not deleted,
  and rebuilds the grid from the widths unless every cell records its old properties;
  with both, its Accept All and Reject All are ours (*Phase E5*).
- *A formatting change that changes nothing* (bold on a heading already bold): Word's
  re-save drops it, record and all.
- *A hyperlink holding a revision*: Word re-saves it as a `HYPERLINK` field (text and
  address kept), and as a `w:hyperlink` again once accepted.
- *The describing comments* stay after Word's own Accept All or Reject All; ours go with
  the changes they describe.

**Approximations, and why:**

- *Dates* are UTC in `w:date` too (Word: local time there): output independent of the
  machine.
- *Where Word's own form cannot be accepted or rejected* (a mark at a container's end),
  the form chosen is one Word handles (above); Word's typing form (the previous mark
  inserted) is used only at a container's end.
- *Which element survives* a tracked move (the destination: a new paragraph, the caller's
  id an alias of it), a join (the last) or an end-of-container insertion is Word's; a
  reject therefore keeps the content and properties, not always the paraId, of what was
  there.
- *A tracked move* of a table, of a paragraph holding revisions or last in its container is
  a deletion and an insertion; of a paragraph holding a note or comment reference,
  untracked (a warning says so). The copy carries no bookmark or comment range.
- *Not revisions*, as in Word: bookmarks, style and list definitions, a picture's size and
  alt text. Rejecting keeps a style an edit added (as Word does); numbering instances it
  left unused go.
- *Records grouped* for `accept(id)`: a move's both ends; a row's, a table's cell records
  and a paragraph's mark with what the same author made at the same date inside them.
- *The original view* undoes paragraph-property changes now, not run formatting (a bold
  change still shows bold).
- ~~*Columns* are inserted and deleted where no cell spans the boundary; merges refuse a
  rectangle a cell lies partly in or one holding a vertical merge~~ -- E5 widened both:
  columns go in and out across spans and skips, a merge grows to whole cells.

**Proposals** (not made here): for docx2svg, ~~`ListCounters.label` asks for docx2svg's own
document and paragraph objects; docx-agent passes a stand-in with the numbering and a
paragraph-properties mapping -- a function taking ``(numbering, numId, ilvl)`` would make
the counter a public API~~ -- done: docx2svg's `ListCounters.item(numbering, numId, ilvl)`,
which the reader now calls with `parse_numbering`'s result, no stand-in left; the counts
against docx2svg's recording of Word still hold. For ooxml-edit, `reap` leaves a `Default` content type for an
extension no part has any more (a last picture removed); harmless, but a package Word
never writes.

## Phase E4 — structure: sections, headers and footers, notes, fields (L)

Sections and page setup, section breaks; headers and footers (create, link, unlink, first
page, even/odd); footnotes and endnotes; fields, including TOC generation from headings with
page numbers from docx2svg's layout, iterated to a fixed point.

**Done when:** each operation passes the standard gates; a section break inserted
mid-document gives docx2svg and Word the same page count and the same section on each page;
headers show on the right pages in Word's PDF; a generated TOC's page numbers equal Word's
own after Word updates the field on a re-save.

### E4 — done (2026-10-04)

Every "done when" holds, on all fourteen fixtures (E1's), and on the probe documents.

- **Measured first** (`tools/e4_probe.py`, Word 16.106 for Mac, Dutch interface, under the
  machine-wide lock; `tests/observations/e4-word.json`): Word itself, by AppleScript, made
  twelve probes' edits and saved them -- section breaks of each kind, a section's page
  setup, headers and footers of each kind, page-number fields in footers, a footer
  unlinked and a header linked again, footnotes and endnotes with section numbering and a
  note deleted, a caption, a bookmark with `REF` and `PAGEREF`, a `DATE`, a `HYPERLINK`
  field, a table of contents over three and over nine levels on two text widths, a break
  deleted, and the same kinds of edit tracked. What it wrote, and docx-agent with it:

  | Edit | What Word writes (and docx-agent) |
  | --- | --- |
  | Section break at a paragraph's start | an empty paragraph of its own holding the `w:sectPr`, with the next paragraph's properties (before a heading: a heading); docx-agent writes every break so (list membership left out: not measured) |
  | Section break at a paragraph's text's end | Word moves the text into a new paragraph that ends the section and leaves its old mark, empty, after it; docx-agent uses the form above instead (no empty line in the next section; the break paragraph takes no room: docx2svg's 4.11) |
  | The split section | the new `w:sectPr` a copy of the split one; the split one gets the break's kind as `w:type` (`nextPage` written as none) and **gives its header and footer references to the new section**, inheriting them from there |
  | Removing a break | its paragraph mark deleted: the paragraph joins the next, which keeps its element and paraId; the text takes the next section's properties and headers; the removed section's own header part goes (`removal`). Tracked: the mark `w:del`, the `w:sectPr` staying on it (`tracked-removal`) |
  | Page setup | `w:pgSz` (`w:orient="landscape"`, sides swapped), `w:pgMar` (all seven), `w:lnNumType w:countBy`, `w:pgNumType w:fmt w:start`, `w:cols w:num w:sep w:space`, `w:vAlign`, `w:titlePg`, in the schema's order |
  | Header or footer | a part of its own (`word/header<n>.xml`), one paragraph in the Header (Footer) style; references ordered even, default header; even, default footer; first header, first footer; `w:evenAndOddHeaders` in the settings, document-wide |
  | Unlink / link | unlinking copies the story before it into a new part (an unlinked footer starts as a copy); linking again removes the reference and the part |
  | Page numbers | `PAGE  \* MERGEFORMAT` complex, its result `w:noProof`; `NUMPAGES` and `SECTIONPAGES` the same as `w:fldSimple` |
  | Notes | a document's first note brings both notes parts with separators (E2); the reference run in Footnote Reference; the note's paragraph in Footnote Text opening with `w:footnoteRef` and a space. Numbering, restart and position are the **section's** `w:footnotePr` (`pos`, `numFmt`, `numStart`, `numRestart`), never the settings'; a deleted note goes with its reference |
  | Caption | a Caption paragraph: `Figure `, `SEQ Figure \* ARABIC` as `w:fldSimple` (result `w:noProof`), the title as a plain run |
  | `DATE`, `REF`, `PAGEREF` | `DATE \@ "d MMMM yyyy"` with the date cached (English month names in the en-GB probe); `REF Name \h` (no result when the bookmark is empty); `PAGEREF Name \h` result `w:noProof` |
  | `HYPERLINK` field | Word saves it as a `w:hyperlink` in the Hyperlink style: docx-agent writes that |
  | Table of contents | `TOC \o "1-3" \h \z \u` (the probe's Insert Field added `\* MERGEFORMAT`); begin, instruction and separate open the first entry's paragraph, the end opens the paragraph it was put before (or an empty one of its own); entries in TOC *n* with a right dot-leader tab at **the text width less 10 twips** (9016 of 9026, 9376 of 9386) and a `w:noProof` mark; with `\h` a hyperlink to `_Toc<9 digits>` holding the text (Hyperlink, `w:noProof`), a tab and a `PAGEREF _Toc… \h` (`w:noProof`, `w:webHidden`, and an empty run between instruction and separate); a `_Toc` bookmark around each heading's text; the TOC 1-9 styles (`word_toc_styles.py`: spacing after 100, indent 220 a level) |
  | Tracked | a section change is a `w:sectPrChange` holding the old values of what changed (margins whole; of the columns only `w:num`); a break's paragraph mark inserted; a header's text `w:ins` (creating, linking, unlinking are not tracked); a note's reference and content `w:ins`; a field's runs `w:ins` (Word leaves the separate and result outside); a table of contents' entries inserted, its links as `HYPERLINK` fields, the end's run inserted |

- **Sections** (`edit/sections.py`): `Section` reads `start`, `page_width`/`page_height`,
  `orientation`, `margins`, `text_width`, `columns`, `line_numbering`,
  `vertical_alignment`, `title_page`, `page_numbering`, `notes(kind)`, `paragraph_ids`,
  `header(kind)`/`footer(kind)` (inheritance resolved) and `stories()`;
  `doc.section_of(id)`; `insert_section_break(after=|before=, kind=)` (the result's `id`
  the new section's); `remove_section_break(id, join=True)` (`join=False` keeps the
  paragraph); `set_section(id, **values)` (points: `start`, `page_width`, `page_height`,
  `orientation`, `margin_*`, `header_distance`, `footer_distance`, `gutter`, `columns`,
  `column_space`, `column_separator`, `column_widths`, `line_numbering`,
  `vertical_alignment`, `title_page`, `page_number_format`, `page_number_start`,
  `footnote_*`, `endnote_*`), checked before anything changes;
  `even_and_odd_headers` (document-wide, said in `warnings`).
- **Headers and footers**: `add_header`/`add_footer(section, kind, text)` (a first one turns
  on `w:titlePg`, an even one `w:evenAndOddHeaders`, each said), `link_to_previous` (and
  `remove_header`/`remove_footer`), `unlink_from_previous`; a story is edited through the
  paragraph API by its ids (`header3/p:…`), page numbers by `insert_page_number`.
- **Notes** (`edit/notes.py`): `insert_footnote`/`insert_endnote(at, text)` (a line per
  paragraph; `fn:<id>`), `edit_note`, `move_note`, `delete_note`, `Note.edit/move/delete`;
  numbering per section through `set_section`; the endnotes' position by
  `set_note_settings("endnote", position="sectEnd")` (the settings', where docx2svg
  measured Word reading it); a settings part is made where a document has none.
- **Fields** (`edit/fields.py`): `fields(story)`, `field(id)` (`Field.instruction`,
  `keyword`, `result`, `dirty`, `paragraph_id`, `update()`); `insert_field(at,
  instruction)`, `insert_page_number`, `insert_date`, `insert_sequence`,
  `insert_caption(after=|before=, label=, text=)`, `insert_hyperlink`,
  `insert_cross_reference` (E1's, its page now the number Word shows: the page's number in
  its section's format), `insert_toc(before=|after=, levels=, hyperlinks=, hide_in_web=,
  outline_levels=)`, `update_fields(ids, date=, passes=)` and `page_label(id)`. Results:
  `PAGE`, `NUMPAGES`, `SECTIONPAGES` and `PAGEREF` from docx2svg's pages (`PageInfo`'s
  number and section; switches through docx2svg's `fields.field_text`), `REF` from the
  bookmark, `SEQ` counted (`\r`, `\c`, `\h`, `\*`), `DATE` from the date given, a TOC
  rebuilt from the headings (outline levels from styles with `\o`, the paragraph's own
  with `\u`; empty headings left out; Word's `_Toc` bookmarks reused), the document laid
  out again until the results hold (at most `passes`; `insert_toc` runs it).
  `EditResult.unknown` lists what the layout cannot say (a target past a stop, a header
  no page shows): left as it was, **never marked dirty** (measured: a `w:dirty` field makes
  Word ask on opening; its export blocks).
- **Tracking** (`track=True` on every E4 edit, or `doc.tracking`): what Word tracks is
  written as a revision -- section properties (`w:sectPrChange`: the old value of each kind
  that changed, a kind that was absent in its "off" form, which rejecting removes again --
  review reads that form), a break (its mark inserted, the split section's kind recorded),
  a break removed (its mark deleted), notes (inserted with their reference; deleted as the
  reference's deletion, which accepting turns into the note's; moved as a deleted reference
  and an inserted copy), fields and captions (`w:ins`; a `w:fldSimple` kept out of revision
  containers, its result run inserted), a table of contents (entries and the end inserted,
  hyperlinks around their insertions), text in headers and footers. Not revisions, as in
  Word: creating, linking and unlinking stories, `w:evenAndOddHeaders`, bookmarks, field
  results (`update_fields` writes caches). Review gained: a section change's "off" form
  rejected away, an empty break paragraph's mark accepted or rejected with its section,
  a simple field a review empties pruned, a note's paragraph gone with its note.
- **Undo**: every operation is one undo step, undone to the original bytes, redone to the
  edited ones (tested on every fixture).

**Tests** (333 new; 2,675 passing in the default run, 173 skipped where a fixture has no
such place; 9.5 minutes under `pytest -n auto` on four cores). Every E4 operation
(`tests/e4_edits.py`: a break, page setup, a break removed, headers and footers with page
numbers, linking and unlinking, notes inserted, edited, moved and deleted, fields with
`update_fields`, a table of contents) on every fixture: edit, save, reopen, read back; no
validity problem added; the compatibility mode kept; undo to the original bytes and redo to
the edited ones; laid out by docx2svg with every paragraph it made placed by `where()` (or
said unknown for a reason). Tracked (`tests/test_e4_tracked.py`): a break, page setup, a
break removed, notes, fields and a table of contents on every fixture -- accept-all is the
untracked edit and reject-all the original in canonical form (forgiving bookmarks, styles,
emptied notes parts, notes' numbering, and a tracked break's shared references, whose
stories are checked instead). And `tests/test_e4_word_forms.py` rebuilds the probes and
holds what each edit writes to what Word wrote: page setup, breaks of each kind, a break
before a heading, removing a break, headers' parts and references, unlinking and linking,
page-number footers, a note, section note numbering, a caption, `DATE`, a hyperlink field,
`SEQ` counting and restart, a `PAGEREF` in a lower-Roman section, a field past a stop, the
table of contents (every entry's XML equal to Word's, and every page number), TOC 1-9, a
rebuilt TOC, and the tracked forms.

**TOC and page references against Word.** Offline, on the probe Word itself updated
(`tocfield`): every entry and every page number of our table of contents equals Word's
(15 / 15). In the oracle, Word opened each document, updated every field itself
(AppleScript `update field` on each) and saved it, and the cached results were compared
with ours: on all fourteen fixtures with the edit set and a `PAGEREF` to the last heading,
on sample-long's 18 chapters over 36 pages, and on the probe with a section restarting its
numbers in lower Roman -- **143 of 143 entries and page references agree, 0 differ**; none
was past a stop (docx2svg laid every one of these documents out to its end), so none had to
be left out.

**The oracle** (`pytest -m oracle`, `tests/test_oracle_e4.py`, 88 tests, all passing; under
the shared lock, serially). Every fixture with the edit set exports from Word unprompted,
its PDF showing the footnote, the endnote, the first-page header and the headings; a break
inserted mid-document gives Word and docx2svg the same page count and the same section on
every page, on all fourteen fixtures (each section's footer names it); headers and footers
show on the right pages -- the first-page header on page 1 only, the even footer on even
pages, the default footer with its computed page number on the other odd pages -- in Word's
PDF and in docx2svg's pages alike; the TOC and page references above; a Word re-save keeps
every section's start, orientation, columns and title page, the stories each section shows
(their text), the notes, the fields and the table of contents with its numbers; tracked,
Word's own Accept All and Reject All read as ours (final view's Markdown, notes by order,
and the sections' start, columns and top margin), but for the blank body, recorded below.
A dirty field, last: Word's export blocks on its question.

**What Word does otherwise, recorded:**

- *A break at a paragraph's text's end* (Word's form, above) is not docx-agent's: the
  break paragraph is always one of its own.
- *Tracked section changes*: Word records of the columns only `w:num` (its old
  `<w:cols w:space="708"/>` as `<w:cols w:num="1"/>`), so its own rejection would lose the
  spacing; docx-agent records the old element whole.
- *A tracked break* leaves the split section's references where they are (references are
  not revisions): the two sections share the parts until it is accepted, where the
  untracked break moves them (the stories each section shows are the same).
- *AppleScript's `get header`/`get footer`* makes empty headers and footers of every kind
  for the section asked about (seen in every probe); not something Word's interface does,
  and not mirrored.
- *Footnote options set on a section's range* applied to that section and the next in one
  probe (`notes`), to the last section alone in the other (`notes2`); docx-agent writes the
  section asked for. Word's AppleScript refuses to set the endnotes' location or format
  (`notes2`'s failed steps, recorded).
- *A dirty field* (`w:dirty="true"`) makes Word ask whether to update the document's
  fields on opening: the export blocks (`test_oracle_e4`, twice), and the Word the recovery
  quits can hold the question into its next launches in the same run (every save after it
  blocked; once a Word left behind finished the export minutes later). The test now makes
  that Word go before it ends.
- *Reject All with a blank body* (samplelib's, no paragraph): Word leaves the break joined
  into the body's section with the break's kind and the tracked change kept (continuous,
  two columns, 54 pt) -- the last mark it cannot review (E2, E3); ours gives the original.
- *A tracked section change* Word rejects by setting what it recorded over what is there
  (measured: a record of `<w:cols w:space=…/>` left two columns), so a change that adds an
  attribute records its default value (`w:num="1"`), which rejecting here drops again.

**Approximations, and why:**

- *A break's paragraph* takes the next paragraph's properties without its list membership
  (a numbered heading would number an empty paragraph; Word's choice not measured).
- *A new section copies everything* of the one it splits but its references (moved, as
  measured), its page-number restart included (not measured).
- *Removing a break joins* its paragraph into the next keeping the next's properties
  (Word's tracked form, measured; the untracked join with two styled paragraphs not
  measured); `join=False` keeps the paragraph.
- *A header added to a section* is one paragraph in the Header style; a first-page one
  sets `w:titlePg`, an even one the document-wide setting (Word's interface does the same
  through its checkboxes).
- *Line numbering's start* is written 0-based (`w:start` = start - 1), *distance* and
  *restart* as given: only `w:countBy` measured. docx2svg does not draw line numbers.
- *TOC entries* hold the heading's text (field results in, object marks out); a numbered
  heading's list label is not in it (Word writes it, with a tab: not measured here), nor
  `\t`, `\b`, `\f`, `\l`, `\p`, `\s`, `\d` switches; with no headings, Word's English
  "No table of contents entries found."; in a section of several columns the tab is at the
  page's text width (not measured).
- *A `PAGEREF`* names the page the bookmark's paragraph starts on, not the line its start
  is on (the same for a heading).
- *`DATE` pictures* are written with English names; Word writes the run's language's.
- *A note's numbering in the settings* (`w:footnotePr`'s format, start, restart) is not
  written: Word ignores it there (docx2svg's 4.12, 4.13).

**Proposals** (not made here): for docx2svg -- `PageInfo` (`page.info`: the page's number,
section, story kind, blank) is what PAGE-like results need; it would help to document it
as public, and `fields.field_text` to take `PAGEREF` (docx-agent passes the instruction
with `PAGEREF` read as `PAGE`); line numbering (`w:lnNumType`) is neither drawn nor
warned, where `w:vAlign` is warned; `PAGEREF`, `REF` and `SEQ` could be computed as
`PAGE` is (its H.7), now that docx-agent writes their results. For ooxml-edit --
`insert_in_order` leaves a detached child unattached when the parent has no registered
sequence (`append_in_order` attaches it): raising, or appending, would be safer.

## Phase E5 — tables and drawings (L–XL)

Tables: rows, columns, merges and splits (merge-aware, randomised sweep as in pptx-agent),
widths, borders and shading, styles and `tblLook`, header rows, nested tables. Drawings:
floating positioning and wrap, z-order (`relativeHeight`), text boxes' text through the
story API, charts' data through the shared chart code if pptx-agent's `edit/chart.py` has
moved to a shared package by then (DrawingML charts are the same part in Word). Content
controls: fill, including data-bound ones.

**Done when:** standard gates; the table sweep keeps the grid consistent after each step;
docx2svg and Word agree on the pages of edited tables; every content-control fill shows in
Word's PDF, bound ones too.

### E5 — done (2026-10-04)

Every "done when" holds, on all fourteen fixtures (E1's) and on the probe documents; chart
editing waited for the shared package (decision 11) and is in *Charts and SmartArt*.

- **Measured first** (`tools/e5_probe.py`, Word 16.106 for Mac, Dutch interface, under the
  machine-wide lock, no clipboard; `tests/observations/e5-word.json`): Word itself, by
  AppleScript, set every table property on tables made as Word makes them (E2's), merged,
  split and inserted and deleted rows and columns across spans and merges, floated pictures
  and set every wrap type, position, order and lock, made a text box and a shape, and typed
  into content controls of each kind (written here: Word's dictionary has no
  content-control object) and a data-bound one -- each untracked and tracked; then reviewed
  five tracked forms of a merge across with its own Accept All and Reject All. What it
  wrote, and docx-agent with it:

  | Edit | What Word writes (and docx-agent) |
  | --- | --- |
  | A new table | E2's (Word's Insert Table): the document's table style, `w:tblW` auto, `w:tblLook` 04A0, the room shared among the columns (a nested table: its cell's width less 216 twips of margins, as Word's 2142 + 2142 of a 4500-twip cell), `w:tcW` in each cell; a paragraph between it and a table beside it, and after it where it would end a cell |
  | Header rows, layout, width, alignment, indent, look | `w:trPr/w:tblHeader`; `w:tblLayout w:type="fixed"`; `w:tblW w:type="pct"` in fiftieths with the grid shared again (80% of a 9026-twip column, less its outer borders' halves: 2405/2404/2404, each line at the ceiling of its cumulative share) and the cells' `w:tcW` left; `w:jc` in the table's properties **and every row's**; `w:tblInd`; `w:tblLook` with `w:val` the flags' bits |
  | A row | `w:trHeight` (at least: no `w:hRule`; exactly: `w:hRule="exact"`); `w:cantSplit` |
  | A cell | `w:shd w:val="clear" w:color="auto" w:fill`; `w:tcBorders`, **an edge written on both cells** (the neighbour's opposite side); `w:tcMar` (only the sides set); `w:vAlign`; `w:textDirection` (`btLr` upward, `tbRl` downward) -- and with it the row's `w:cantSplit`, an at-least height of 1134 where it has none, and the cell's paragraphs indented 113 each side |
  | A column's width | its `w:gridCol` and every cell's `w:tcW` over it |
  | Floating | `w:tblpPr` (`leftFromText`... `vertAnchor`, `horzAnchor`, `tblpX`/`tblpY` -- **a positive offset a twip more than given**: 100 pt is 2001, docx2svg's F.18 places it a twip less -- or `tblpXSpec`/`tblpYSpec`) and `w:tblOverlap w:val="never"` |
  | Merge across, down, a rectangle | the first cell's `w:gridSpan` (`w:tcW` summed) and `w:vMerge` restart/continue; every other cell's paragraphs appended to the first **in reading order** (row by row, left to right: E3's order was column-first, now Word's); a continuation keeps one empty paragraph |
  | Split | across: a grid column divided (1000/1000 of 2000) and every other row's cell over it spanning both; a cell spanning enough columns shares them instead; down: a merge divided, or new rows below with the row's other cells merged down into them (Word also gives the rows an at-least height from its layout -- not written here); the content stays in the first cell |
  | A column beside a cell spanning the boundary | **the new cell beside the spanning cell, which keeps its span** -- Word's Insert Columns works on cells, not the grid; a row skipping the grid there skips a column more |
  | A row inside a merge / its first row deleted | the merge carried on; deleting its first row makes the next one its start (the first's text goes with it) |
  | Floating a picture | `wp:anchor`: `distL`/`distR` 114300, `relativeHeight` 1024 above the highest (Word's first 251658240), `layoutInCell="1"`, `allowOverlap="1"`, against the column and the paragraph, **`wp:wrapNone`** (in front of the text), the ids kept |
  | Wrapping, distances | `wp:wrapSquare wrapText="bothSides"` (`left`, `right`, `largest`); `wp:wrapTight`/`wp:wrapThrough` with Word's polygon (0,0 0,21000 21300,21000 21300,0); `wp:wrapTopAndBottom`; front `wp:wrapNone behindDoc="0"`, behind `behindDoc="1"`; `distT`... in EMU |
  | Position | `wp:posOffset` in EMU as given, or `wp:align` (`center`, `inside`...) against `relativeFrom` |
  | Order, lock, overlap, size | to the front: 1024 above the highest; to the back: **1025 below the lowest**; `locked="1"`; `allowOverlap="0"`; a resize adds `wp14:sizeRelH`/`sizeRelV` of 0 |
  | Inline again | `wp:inline`, every distance 0, the ids kept |
  | Text box | Insert Text Box's: `mc:AlternateContent`, a `wps:wsp` with `txBox="1"`, no fill, a 0.5 pt black line, `w:txbxContent`, `wps:bodyPr` (insets 91440/45720, anchored top), wrapped square; the fallback a VML `v:shape` of `#_x0000_t202` holding **the same `w:txbxContent`, paraIds and all**, lengths in inches where whole (`2in`) |
  | Shape | `make new shape`'s: the preset, `wps:style` (line `accent1` shaded 15000, fill `accent1`, font `lt1`), an effect extent of 12700 for its line, centred text, the fallback `v:rect`/`v:oval`/`v:roundrect` |
  | Content controls | re-saving: `w:tag` before `w:id`, an empty `w:sdtEndPr`; typing: the content's text replaced in its first run; a date control's `w:fullDate` follows a date typed; a drop-down refuses typing; **a bound control shows its node** (a stale cache replaced on opening) and Word's own typing left the node as it was |
  | Tracked | a table's or a cell's properties: `w:tblPrChange` (the whole old `w:tblPr`), `w:tblGridChange` (the old grid) and **`w:tcPrChange` on every cell of the table, changed or not**; where the grid changed, `w:trPrChange` on every row (`w:gridAfter w:val="0"`); a row's height, breaking, header and alignment, a table's floating, every drawing setting: **no record** (applied); a new text box or shape: its run `w:ins`, its text `w:ins` (the fallback's copy with ids of its own); typing into a control: `w:del` and `w:ins` inside `w:sdtContent`; merges, splits and a column's deletion: untracked (E3) |

- **Tables** (`edit/table_format.py`, `edit/tables.py`): `insert_table(rows, columns, after=|before=, data=, style=, widths=, width=, layout=, header_rows=, alignment=, indent=, look=)`; `set_table(id, style=, width=, layout=, alignment=, indent=, look=, borders=, shading=, cell_margins=, header_rows=, floating=)` (`None` removes; floating `{x, y, x_align, y_align, horizontal_anchor, vertical_anchor, left, right, top, bottom, overlap}`); `set_column_width`; `set_row(id, row, height=, height_rule=, cant_split=, header=)`; `set_cell(id, row, column, shading=, borders=, margins=, vertical_alignment=, text_direction=, width=)`; `table_properties`, `row_properties`, `cell_properties` (declared values, points); `merge_cells` (the rectangle grows to whole cells and merges, and says so), `split_cell(id, row, column, rows=, columns=)`, `insert_column`/`delete_column` across spans and skips, `insert_row`/`delete_row` in merges; nested tables by inserting after a cell's paragraph. `Table.set`, `.split`, `.set_row`, `.set_cell`, `.properties`. Points throughout.
- **Drawings** (`edit/drawings.py`): `float_drawing(id, **settings)`, `inline_drawing(id)`, `set_drawing(id, wrap=, side=, distances=, x=, y=, x_align=, y_align=, horizontal_from=, vertical_from=, z_order=, lock_anchor=, allow_overlap=, layout_in_cell=, width=, height=)`, `move_drawing`, `resize_drawing` (a group's members scale with it: its child extent kept), `insert_text_box(at, text, width=, height=, x=, y=, wrap=)`, `insert_shape(at, preset, ..., text=)` (every `ST_ShapeType` name, ooxml-common's table); a text box's paragraphs are a story (`Drawing.paragraphs`, `append_paragraph(story="d:<id>")`) edited through the paragraph API, its VML fallback brought in step after every edit (`Document._edit`) and its style after every move; `Drawing.position`, `.wrapping`, `.size`, `.text`, `Picture.float`/`.set`.
- **Content controls** (`edit/controls.py`): `fill_control(id, value)` for plain and rich text (lines as paragraphs in a block one), drop-down (an item's text or value), combo box, date (the control's picture, `w:fullDate`), check box (`w14:checked` and its state's glyph); a bound control's node written with it (`w:dataBinding`'s XPath into the custom XML part named by `w:storeItemID`), and after a review the node set from what the control shows; `w:lock` respected; a placeholder shown goes. `insert_control(at, kind, text=, tag=, alias=, items=, date=, checked=)` (in the line, or a block one `after=`/`before=`; never inside another control's content), `remove_control(id, keep_content=)`; `ContentControl.value`, `.items`, `.binding`, `.lock`, `.fill`, `.remove`.
- **Validity** gained E5's checks: every row covers the grid exactly (`w:gridBefore`, spans, `w:gridAfter`) and a vertical merge goes on only under a cell of its columns (`validate.grid_problems`; a table holding cell or grid revisions once reviewed), and the child orders of borders, margins, `w:sdtPr`, `w:sdt`, `wp:anchor` and `wp:inline`.
- **Review** gained: a table's old grid reviewed with any record of its table; a rejected span's grid rebuilt from the cells' widths where no old grid was recorded (as Word); a row's recorded `w:gridAfter` of 0 dropped; a text box's fallback re-copied; bound controls' nodes set.

**Tracking** (`track=True` or `doc.tracking`). What Word tracks is a revision; what it
does not is applied untracked and said in `warnings`. Decision 7's structural edits are
written in forms **Word's own Accept All and Reject All give back exactly** (measured on
each, below): a merge across as E3's (the first cell's span in a `w:tcPrChange`, the others
`w:cellDel` with their content and **their paragraph marks** deleted -- Word accepts a
deleted cell whose marks stay by widening the cell before it) with **every cell's old
properties and the old grid recorded**, Word's own form for a table's properties -- without
them Word's Reject All rebuilds the grid from the widths; a merge down as E3's `w:cellMerge`;
a split's new cells' and rows' marks inserted, every cell's old properties recorded, **no
old grid** (with one Word's Reject All keeps the split column), and a divided merge as
**`w:cellMerge` (`vMerge` the new, `vMergeOrig` the old) beside the old properties** -- a
`w:tcPrChange` alone Word accepts as the old merge; a merge's first row deleted likewise;
a column across a span: Word's own column form, with the old grid and rows recorded.
**The merge across E3 recorded as a difference is now Word's both ways.**

**Tests** (623 new; 3,298 passing in the default run, 173 skipped; 13.4 minutes under `pytest -n auto` on four cores). Every E5 operation (`tests/e5_edits.py`: a table with a header row; its properties, rows and cells; merges, splits and columns across them; a nested table; a floating table; a picture floated, wrapped, placed, ordered, locked; one floated and put inline again; a text box edited through the paragraph API; shapes; content controls of each kind inserted, filled and removed) on every fixture: edit, save, reopen, read back; no validity problem added (grid consistency, unique drawing ids, controls and custom XML among them); the compatibility mode kept; undo to the original bytes and redo; laid out by docx2svg with every paragraph it made placed by `where()` or said unknown, page 1 rendered. Tracked (`tests/test_e5_tracked.py`), on every fixture, eighteen cases -- a table inserted; a table's and a cell's properties; a column's width; merges across, of a rectangle and grown to whole cells; splits across, down, of a merge and both; a column inserted and deleted across spans; a row inserted into and deleted from a merge; a text box; a shape; a bound control's fill -- accept-all is the untracked edit and reject-all the original in canonical form, the grid consistent; what Word does not track applied with a warning; a bound control's node following the review. The randomised table sweep (`tests/test_e5_table_sweep.py`: 25 random steps on 12 seeds, rows and columns inserted and deleted, rectangles merged, cells split, widths, borders and shading, on tables with spans and merges) keeps the grid consistent after every step, adds no validity problem and undoes exactly; tracked, each step on 8 seeds accepts to the step and rejects to the table before it. `tests/test_e5_word_forms.py` rebuilds the probes and holds what each edit writes to what Word wrote, block for block: every table property, merges and splits, columns and rows across spans, every float, wrap, position, order, lock, overlap and size, putting a picture inline again, the text box's whole form. And `tests/test_e5_units.py`: groups moved and resized, a text box's fallback through edits and tracking, the VML style after a move, refusals.

**The oracle** (`pytest -m oracle`, `tests/test_oracle_e5.py`, 84 tests, all passing --
377 with E0-E4's, all passing; under the shared lock, serially). Every fixture with E5's
edit set (`e5_edits.e5_edit_set`) exports from Word unprompted, its PDF showing the tables'
cells (nested ones too), the text box, the shape and every control's value. **Word's pages
are docx2svg's**: the same page count and, page by page, the same characters, for the
table probe with every property edit, the merge probe with every merge, split and column
and row edit, and the edit set on 11 of the 14 fixtures; on the other three
(`ids-and-markup`, `sample-long`, `sample-simple`) a line at a page's foot changes pages
(each edit group exported alone agrees there: small differences adding up), recorded and
asserted as at most a line. Every content-control fill shows in Word's PDF -- plain, rich,
drop-down, combo box, date (in its picture), check box (the glyph), a block one, **two
bound ones** -- and Word shows a bound control's node, not a stale cache. A Word re-save
keeps every table's cells, spans, merges and floating, every drawing's kind, place,
wrapping, size and text, every control's kind and text, and every edit reads back from
Word's file as from ours; a bound fill survives a re-save in the control and the node.
**Word's own Accept All and Reject All read as ours** for all eighteen tracked cases (36
of 36): a table inserted, its and a cell's properties, a column's width, merges across, of
a rectangle and grown, splits across, down, of a merge and both, a column inserted and
deleted across spans, a row inserted into and deleted from a merge, a text box, a shape,
a bound fill. E3's oracle now holds its own merge across to "same" both ways.

**What Word changes on re-saving E5's edits**: it keeps every `docPr` id,
`wp14:anchorId`, `relativeHeight` and content-control id; adds `w:shapeDefaults` to the
settings with a text box or shape; writes no `o:gfxdata` it was not given; renumbers
nothing else of E5's; narrows a nested table's grid by its outer borders' halves (2052 to
2042 twips a column in a 4320-twip cell).

**Found on the way, and what the oracle's edit set keeps out of** (docx2svg does not model
them; proposed below): a vertical merge whose row Word splits at a page's foot (docx2svg
moves the row whole: its tables' stage 4, "not measured"); a picture that may not overlap
(Word moves it clear; docx2svg warns); a text box wrapped square beside tables (Word
fitted three lines more on the page); a heading's style carried into new cells (it kept
the table with what followed in Word; new cells are now in the default style, as E2's).

**What Word does otherwise, recorded:**

- *Splitting a cell down*, Word gives the rows an at-least height from its own layout (130 twips here); *splitting into rows and columns at once* it widens the grid by 33 twips. docx-agent writes neither.
- *A column left of a cell another row spans*, Word's AppleScript refused to insert; *deleting a column a span covers* and *deleting a row of a merge* it refused by `delete` (row deletion went by selection and Backspace). docx-agent's forms there are its own: the span narrowed by the column, the grid column gone.
- *Shading a cell solid*, Word also gave the cell above it an explicit bottom border of the table's own line; not mirrored.
- *A vertical text direction*: Word changes the row (`w:cantSplit`, a height) without a record, so its Reject All leaves them; so does docx-agent's.
- *Tight wrapping* came back with `behindDoc="1"` once, through the same steps as through wrapping (0): not mirrored.
- *Word's own tracked table records* are its whole-table form; docx-agent writes the same now (a review needs it), and a row's alignment change as a `w:trPrChange` Word does not write (so a review gives it back).
- *Typing into a bound control* in Word left the custom XML node unchanged in the saved file; opening it again shows the node's value. docx-agent writes the node with every fill.

**Approximations, and why:**

- *A new table's style* is the one the document's tables use most, else Table Grid (E2); only styles the document defines or Word's measured Table Grid are written.
- *A floating table's default distances* (9 pt left and right, Table Properties' "Around") and *its default anchors* (the margin across, the text down) are not measured; the probe set its own.
- *Cell spacing* (`w:tblCellSpacing`) is read and kept, not set (Word writes it on every row too; measured, not built).
- *A split across* divides the cell's width evenly at whole twips; *a split's grid* is rebuilt from the cells' widths when a tracked one is rejected, so grid lines no cell's edge used come back merged (as Word does).
- *A text box or shape is placed* at the offset given from the column and the paragraph (Word's convert puts a floated picture where its layout had it: the file's 0,0 here); its fallback has no `o:gfxdata` (Word's binary cache for old readers), and presets VML cannot draw (all but the rectangle, ellipse and rounded rectangle) have no fallback.
- *Forward and backward* in the order are a step past the neighbour's height (not measured; front and back are).
- *A content control's id* is a 31-bit number of docx-agent's sequence; *a date control* writes `en-GB` and its picture's English names; *a check box* Word's MS Gothic glyphs; *inserting and removing a control* is no revision (Word has none for the wrapper): tracked, its content is.
- *Groups* are read, moved and resized; their members are not edited one by one.

**Proposals** (not made here): for docx2svg -- (1) **a row of a vertical merge split at a
page's foot**: Word splits it, the merge's content going on at the next page's top
(`sample-simple` with E5's merge and split: `m11` on the next page); docx2svg moves the
row whole (its tables' stage 4 leaves it unmeasured) -- a probe of merges at a page's foot
would settle it; (2) **a text box wrapped square beside tables** (`layout-sweep` with the
box anchored in the paragraph before three tables): Word fits three lines more on the
page than docx2svg, each alone agreeing -- the tables beside a drawing (F.15) with a nested
table and spans among them; (3) a header row (`w:tblHeader`) of two-line cells at a page's
foot: Word kept it whole where docx2svg split it (seen once, with a heading's style in the
cells; not isolated further). For ooxml-edit -- `OpcPackage.add_part` with a content type
its extension's `Default` already has registers an `Override` too (harmless; Word writes
none). None for ooxml-common: its preset table names every `ST_ShapeType`, which
`insert_shape` checks against. The oracle helper here (`tests/oracle.py`) now kills a Word
that ignores the polite signal (seen once: a Word hung on a dialog after a tracked row
insertion beside a span).

## Phase E6 — authoring (L)

- **From scratch, mode 15:** a minimal package as Word 365 writes it — measured from a new
  blank document Word saves: `styles.xml` (doc defaults, latent styles, Normal and the
  defaults Word writes), `settings.xml` (`compatibilityMode` 15 and the `w:compatSetting`s
  Word writes with it), `fontTable.xml`, `webSettings.xml`, a theme, `docProps`. The theme's
  and default faces' *names* are facts; the theme XML is written here from measured facts (decided).
- **From a template (`.dotx`, `.dotm`):** copy, switch the main part's content type to the
  document type, keep or drop `w:attachedTemplate`; from `.dotm`, drop the VBA project
  (a `.docx` with macros does not open).
- **Cross-document copy** with style, numbering, media, notes and comments import.
- **`upgrade_to_modern()`:** set mode 15 and drop the compatibility options mode 15
  ignores, as Word's Convert does (measured), reporting the reflow through `compare`.

**Done when:** a document built from scratch with E1–E5's operations, and one built from
each template fixture, open and export in Word unprompted, and Word's re-save changes no
style definition we wrote; copied blocks keep their look in the destination (docx2svg and
Word); an upgraded fixture reports its reflow and exports unprompted.

### E6 — done (2026-10-04)

Every "done when" holds: a document built from scratch -- now in the corpus every phase's
suite runs on (`new/document`), so E1-E5's edit sets and `insert_markdown` are held to it
like any fixture -- and one built from each template open and export in Word unprompted,
Word's re-save changes no style definition docx-agent wrote, copied blocks keep their look
(docx2svg and Word), and every upgraded fixture reports its reflow and exports unprompted.

- **Measured first** (`tools/e6_probe.py`, Word 16.106 for Mac, Dutch interface, metric,
  under the machine-wide lock; `tests/observations/e6-word.json`), by AppleScript: a new
  blank document (`make new document`), saved as it is and with text typed; a document
  made from a template written here, as a `.dotx` and as a `.docx` (`create new document
  attached template`: File > New from a template); a document saved as a template; a
  source copied into a destination whose styles, list, footnote, comment, bookmark,
  picture, content control and paraIds all collide -- by `formatted text` (Word's own copy
  between ranges, no clipboard), by `insert file`, tracked, and through the clipboard with
  each paste option (default, Keep Source Formatting, Use Destination Styles, Merge
  Formatting), **only because the clipboard was empty**, and emptied again after each
  (the probe refuses otherwise: the clipboard's contents could not be put back as they
  were); Convert (`upgrade`) on documents in modes 11, 12 and 14 -- one with every one of
  the 65 legacy compatibility options -- and on three corpus fixtures, beside a plain
  re-save of the same documents. Nothing Word wrote is committed: the new document's
  style set, theme and font table are Word's Normal template, so the observations hold only
  their comparison with `Document.new()`; `src/docx_agent/edit/word_new.py` holds the
  values (document defaults, the 376 latent styles, each style's properties, the font table,
  the compatibility and math settings) and `edit/blank.py` writes every part from them.

  | What | What Word does (and docx-agent) |
  | --- | --- |
  | A new blank document | `document.xml` (one empty paragraph, the section), `styles.xml`, `settings.xml`, `webSettings.xml`, `fontTable.xml`, the Office theme, `docProps`: **every part the same** as `Document.new(language="nl-BE")`'s, once rsids, the document id, the window's view, dates and authors are set aside and Word's localised names (the theme's "Kantoorthema", the linked styles' "Kop 1 Char", the ids `Standaard`, `Kop1`) read as English Word's. Its styles: Normal, heading 1-9 and their linked character styles, Default Paragraph Font, Normal Table, No List, Title, Subtitle, Quote, List Paragraph, Intense Emphasis, Intense Quote, Intense Reference -- the Aptos generation's (12 pt, 8 pt after, line 278 auto, headings in Aptos Display and accent 1 shaded) |
  | ... the locale | this Word: A4, 2.5 cm margins (1417), header and footer 708, tab stop 708, hyphenation zone 425, `,` and `;`, the language nl-BE. `blank.Locale` holds these: `METRIC` (measured, the default) and `US` (Letter, 1440, 720, no zone, `.` and `,`: English Word's known defaults, not measured here); the language is its own parameter, `en-US` by default |
  | ... what Word adds on saving | rsids, `w15:docId`, `w:view w:val="normal"` (its window's Draft view), its own `app.xml` counts and `Application`: docx-agent writes none of them (a file without a view opens in Print Layout) |
  | From a template | the whole package kept -- styles, numbering, theme, settings, sections, headers and footers **and the body**; the main part a document's; `w:attachedTemplate` to the `.dotx` (`file:///` and its path) and `app.xml`'s `Template` its name (from a `.docx`: no attachment, `Normal.dotm`); core properties: title, subject, keywords kept, author and last modifier Word's user, revision 1, created and modified now; `app.xml`'s company kept; Word's own settings added on saving (the compatibility settings, `themeFontLang`...) and the notes parts |
  | Save as template | the main part's content type, and nothing else of the document's own |
  | Copy, use destination (Word's default paste, Use Destination Styles, `insert file`, `formatted text`: all alike) | a style the destination has by name is the destination's (the copy changes look); one it lacks is imported, a paragraph style with the source's document defaults where they differ (the face and size), a character style without; the list a new instance over a copy of its definition (its `nsid` kept); footnote and comment renumbered; both pictures one media part; the copy's `docPr` id and content-control id re-issued; the copy's colliding **bookmark dropped**; **every paraId of the document rewritten** |
  | Copy, Keep Source Formatting | no style imported: each paragraph Normal, the source's effective paragraph, mark and run formatting that differs from the destination's default written direct (a character style too: the footnote reference superscript) |
  | Copy, Merge Formatting | Normal, and only bold, italic and underline, direct |
  | Copy, tracked | every paragraph's mark and runs `w:ins`, the row `w:trPr/w:ins` |
  | Convert | `w:compat` only: `compatibilityMode` 15 and `overrideTableStyleFontSizeAndJustification`, `enableOpenTypeFeatures`, `doNotFlipMirrorIndents`, `differentiateMultirowTableHeaders` (1) and `useWord2013TrackBottomHyphenation` (0); of the 65 legacy options it keeps ten -- `spaceForUL`, `balanceSingleByteDoubleByteWidth`, `noExtraLineSpacing`, `doNotLeaveBackslashAlone`, `ulTrailSpace`, `doNotExpandShiftReturn`, `suppressBottomSpacing`, `adjustLineHeightInTable`, `doNotUseHTMLParagraphAutoSpacing`, `applyBreakingRules` -- in every mode, and drops the rest. `useFELayout` goes on *any* save of a document without East Asian languages (a plain re-save drops it too): a save's doing, kept by docx-agent's Convert. A plain re-save adds `useWord2013TrackBottomHyphenation` 1 |

- **The API.**
  - `Document.new(*, page=None, orientation=None, language=None, locale=None, title=None,
    author=None, created=None, template=None, keep_content=True, attach=None)`: blank --
    `page` "A4" (default), "Letter", "Legal", "A5", "A3" or `(width, height)` in points,
    `orientation` "portrait"/"landscape" (the longer side across, `w:orient`), `language`
    (`w:lang`, `w:themeFontLang`; default en-US), `locale` (`blank.Locale`; `METRIC` for
    every page but Letter, which takes `US`), title and author in `core.xml`, `created`
    the creation and modification date; deterministic given `created`; the empty paragraph
    has its paraId from the start. From a template (`.dotx`, `.dotm`, `.docx`, `.docm`; a
    path or bytes): as Word's File > New; `keep_content=False` drops the body (one empty
    paragraph in the default style and the last section stay; notes and comments it
    referenced go, and headers only dropped sections used); `page`, `orientation` and
    `language` replace the template's when given. A `.dotm`'s VBA project, its
    `vbaData.xml` and their content type are **dropped with a `MacrosDropped` warning**
    (refusing would make every macro-enabled house template unusable; a `.docx` carrying
    them does not open). A template below mode 15 gives a document converted as Convert
    does (`ModeUpgraded`): every document docx-agent creates is mode 15. The new document
    is the base state undo returns to.
  - `doc.save_as_template(path)`: `.dotx` (`.dotm` from a macro-enabled document), the
    document unchanged.
  - `doc.copy_blocks(source, ids, *, at="end", styles="use_destination", lists="separate",
    bookmarks="rename")` and `docx_agent.copy_blocks(source, ids, to=doc, at=...)`: `ids`
    a block id, `first..last` or a list; `at` as `insert_markdown` takes it; `styles`
    `use_destination` / `keep_source` / `merge`; `lists` `separate` / `continue` (joins the
    destination list before the insertion when its first level counts alike); `bookmarks`
    `rename` / `drop` (Word's). One undo step, tracked where the document tracks. The
    result's `copied` maps each source id to its copy's (`p:`, `t:`, `fn:`, `comment#`,
    `sdt#`); `warnings` name the styles imported and bookmarks renamed.
  - `doc.upgrade_to_modern()`: opt-in only (no other edit changes the mode); one undo step;
    `result.reflow` is docx2svg's `compare` of the pages before and after, `warnings` the
    options removed.
  - `doc.properties` (title, subject, author, keywords, description, last modified by,
    revision, created, modified, category, status, language -- `dc:language`, the
    metadata's -- and `app.xml`'s statistics) and `doc.set_properties(**values)` (one undo
    step; `None` removes). `save()` brings `app.xml`'s pages, words, characters (with and
    without spaces), lines and paragraphs up to date when the body changed: the text counts
    always (the body's current view), pages and lines from docx2svg when the saved state
    was laid out (a save never lays out on its own; otherwise they stay). The edits never
    touch `app.xml`, so undo stays exact. Word's own counts are a background estimate
    (it saved a four-paragraph document with `Paragraphs` 1 and `Lines` 1).
- **Copying, as built** (`edit/importing.py`): styles by name, an imported style's
  `basedOn`, `next` and `link` imported with it; `keep_source` resolves each paragraph's and
  run's formatting at the XML level -- document defaults, the paragraph style's chain, the
  character style's chain (toggles flipping along both), direct -- and writes what differs
  from the destination's default, theme fonts and colours as the source's faces and values
  where the two themes differ; lists' abstract definitions copied (`nsid` kept unless
  taken), instances made with their overrides; every relationship made again (media by
  content; charts and other parts copied with their own relationships, renumbered where
  they must be; hyperlinks external); notes copied into the destination's notes part
  (made as Word makes it) and renumbered; comments with their `commentsExtended` state and
  reply links, new durable ids, their authors in `people.xml`; bookmarks renamed
  `Name_1`... with the copies' internal links and `REF`/`PAGEREF`/`HYPERLINK \l` fields
  following; `docPr`, `wp14` ids, paraIds and textIds kept unless the destination (or an
  earlier copy) uses them; content-control ids re-issued on collision.
- **docx2svg and Aptos Display.** A new document's headings are in Aptos Display, one of
  Office's *cloud* fonts, which Office for Mac keeps in its group container
  (`FontCache/4/CloudFonts/<family>`), not where docx2svg looks: its layout stopped at the
  first heading. docx-agent gave docx2svg those folders after its own
  (`layout.cloud_font_dirs`, unless `font_dirs` is given) until docx2svg searched them
  itself (its `73b758d`, proposal 1 below); now it passes nothing, and every page of the
  corpus, the chart fixtures and a new document's Aptos Display heading renders the same.
- **Found on the way, in E3** (the new document is the first fixture whose only paragraph
  is empty): a paragraph typed tracked at a container's end after one whose properties
  already carry a recorded change took the changed properties as the old ones (Reject All
  left a heading style); deleting, tracked, a paragraph whose mark was the author's own
  insertion joined the paragraph before into it and lost that one's id. Both fixed. One
  case is recorded, not fixed: tracked text written into the empty paragraph and two
  paragraphs typed after it -- Reject All gives the original text and properties, but the
  surviving paragraph is the last one (which carried the old mark), with its paraId
  (`test_e3_invariant.ID_ONLY`, held as a strict expected failure).

**Tests** (`tests/test_e6_new.py`, `test_e6_template.py`, `test_e6_copy.py`,
`test_e6_upgrade.py`: 120; and the new document in every per-fixture suite). A new document
part by part against the probe's comparison, its package, settings and section against
Word's, its styles, theme, locale, pages, orientation, language and properties,
deterministic given its date; from every template (the three of
`tests/fixtures/generated/templates`, written by `tools/make_template_fixtures.py`, and the
corpus's two Dutch templates): kept and dropped bodies, attachment, properties, macros,
mode, overrides, edits saved, read back, undone and redone, laid out; save as template;
copying against each of Word's observations (styles per policy, the direct formatting Keep
Source and Merge write -- property for property against Word's paste --, lists, media,
notes, comments, bookmarks, ids, tracked: accept-all the untracked copy and reject-all the
destination), docx2svg drawing each Keep Source copy in the source's faces, sizes,
weights and colours, and every fixture copied into a new document and into a template;
Convert's `w:compat` against Word's on the eight converted documents, only the settings
changed, every fixture upgraded with its reflow, valid, undone and redone; properties.
**The suite**: 2,123 passed, 1 expected failure, in the default run (3 min 15 s with
`pytest -n auto` on four cores); 3,562 with `--run-slow` (14 min 8 s). The heaviest
per-fixture suites -- E1-E5's operations, renders and tracked invariants, the Markdown
corpus and tracked insertions in every fixture, the copy and upgrade sweeps -- hold four
representative fixtures by default (`sample-with-table`, `ids-and-markup`,
`lists-and-styles`, the new document; `conftest.REPRESENTATIVE`) and every fixture with
`--run-slow`, the randomised table sweep its first seeds; no gate is dropped from the slow
run or the oracle. Everything: `pytest -q -n auto --run-slow`, then `pytest -m oracle -q`.

**The oracle** (`pytest -m oracle`, `tests/test_oracle_e6.py`, 54 tests: 50 passing, 4
skipped -- samplelib's blank has no paragraph to copy, three fixtures are mode 15 already;
the whole oracle, 448 passing, 4 skipped, the new document in every phase's per-fixture
oracle among them). **Word's own new document with three paragraphs typed and
`Document.new()` with the same text are drawn alike by docx2svg** -- line for line, glyph
for glyph, every position -- **and Word's PDFs of the two have the same characters in the
same faces, sizes and colours.** Word exports unprompted: new documents (A4, Letter,
landscape in Dutch, one with Markdown), a document from each of the five templates with its
body and without, the copied content with each style policy, every fixture copied into a
new document, and every fixture below mode 15 upgraded (its re-save keeping mode 15), and a
document attached to a template inside its sandbox. **Keep Source Formatting keeps the
look in Word**: each copied paragraph's characters in Word's PDF of the destination have
the faces, sizes, colours and weights they have in Word's PDF of the source. A re-save keeps
every style definition a new document has (by name; ids and linked names localised, below),
its text and paraIds; a document from a template its styles, body, header and footer; the
copied content its text, paraIds, both comments, both footnotes, the renamed bookmark and
the two lists counting apart (1, 2 and 1, 2). E4's oracle records one more case: Word's
Reject All of the tracked structure in the new document -- whose only paragraph is empty,
so everything in its body is the insertion -- leaves the break's section properties, as
it does in samplelib's blank (E4).

**What Word changes on re-saving, recorded**: a new document's built-in styles get the
interface's ids (`Standaard`, `Kop1`, `Tabelraster`...) and its linked styles the
interface's names (`Kop 1 Char`), every definition otherwise as written; it adds rsids, a
document id and `w:view w:val="normal"`, and writes its own `app.xml` (`Application`
"Microsoft Office Word", its counts); paraIds are kept.

**What Word does otherwise, recorded:**

- *On a paste with colliding paraIds*, Word rewrites every paraId in the document (E0's
  measurement: a document with a repeated paraId gets fresh ones at its next save);
  docx-agent re-issues only the colliding copies'.
- *A colliding bookmark*: Word drops the copy's; docx-agent renames it by default (the
  copy's links and fields kept working), `bookmarks="drop"` as Word.
- *An attached template*: Word attaches a `.dotx`; **attached, Word for Mac asks for
  access to it on opening when it is outside its sandbox** (the oracle's export blocked on
  the question; one in the Office group container opens unprompted), so docx-agent
  attaches none unless `attach=True` -- and never a macro-enabled one by default, whose
  macros would come back from it.
- *From a template, Word drops the template's paraIds* (it writes none in a new document);
  docx-agent keeps them: complete and unique, Word keeps them on later saves (E0).

**Approximations, and why:** the US locale's values are English Word's known defaults,
not measured on this machine; style ids and linked styles' names are English Word's (Dutch
Word renames them on saving; nothing else changes); copied content takes the source's
current view (insertions kept, deletions dropped, change records dropped); a copied
paragraph's section break is not copied (the copy takes the destination's sections); a
copied data-bound control loses its binding (its text stays); `keep_source` does not flatten
a table style's conditional formatting into its cells' paragraphs; a picture bullet's
`lvlPicBulletId` is dropped from a copied list definition; `app.xml`'s pages and lines are
docx2svg's, and only when the saved state was laid out.

**Proposals** (not made here): for docx2svg -- (1) ~~look for faces in Office for Mac's cloud
font cache (`~/Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts/`), where
Word keeps Aptos Display, the heading face of every new document Word 365 makes; docx-agent
passes the folders itself meanwhile~~ -- done in docx2svg; docx-agent no longer passes them. For ooxml-edit -- (2) `declare_content_type` appends a
second `Override` when the part already has one with another type (a template's main part
becoming a document's), and `content_type` then reads the first: it should change the
existing one (docx-agent does so itself). None for ooxml-common.

---

## Charts and SmartArt (after E6)

Decision 11 held chart editing until pptx-agent's chart and workbook code moved into a
shared package. It did: ooxml-edit 0.2's optional `ooxml_edit.charts` (`Chart`,
`Diagram`, `model`, `GraphicHost`), which pptx-agent hosts in a slide's graphic frame and
docx-agent now hosts in Word's drawings.

### Charts and SmartArt — done (2026-10-04)

```python
chart = doc.chart("d:3")                    # or doc.drawing("d:3").chart; doc.charts()
chart.categories, [s.name for s in chart.series], chart.series[0].values, chart.title
chart.series[0].set_value(2, 4285)          # the cache and the workbook cell, one undo step
chart.series[1].name = "Gross margin"       # cache, header cell, table column name
chart.add_category("Q4", [4400, 530])       # cells move, formulas and the table grow
chart.add_series("Net", [7.4, 7.5, 7.7, 7.9]); chart.remove_series("Cost")
chart.set_title("Margins"); chart.set_axis_title("value", "%"); chart.set_legend(False)
doc.chart("d:5/7")                          # a chart in a group: the group's id / its own
model = chart.model; model["series"][0]["values"][0] = 12; chart.apply(model)

diagram = doc.diagram("d:4")                # or doc.drawing("d:4").diagram; doc.diagrams()
diagram.set_text(1, "Much faster edits")    # data model, and the cached drawing where exact
diagram.node(0).add_child("Smaller files")  # the drawing dropped: Word lays it out again
diagram.remove_node(diagram.nodes[-1])
```

- **The adapter** (`edit/charts.py`): `graphic_host` builds ooxml-edit's `GraphicHost`
  from a Word drawing -- the frame is the `wp:inline`/`wp:anchor` (or a group member's
  `wpg:graphicFrame`), the part whose relationships name the chart is the one holding the
  drawing (`document.xml` for the body and its text boxes, a header, a footer, the notes),
  the edit is docx-agent's one undo step (`Document._edit`), `application="Word"`,
  `document="document"`, `lang` the document's default language (`w:docDefaults`'
  `w:lang`, else `en-US`), and a title template of Word's. `Document.chart(id)`,
  `.diagram(id, on_inexact_drawing=)`, `.charts()`, `.diagrams()`; `Drawing.chart`,
  `.diagram`, `.members`; `Member` for a group's members; `Chart.apply(model)` and
  `Diagram.apply(model)` take the JSON back, ooxml-edit's `ChartModelError` becoming
  `ChartEditError` (an `EditError`), nothing changed. Every call of the subpackage's API is
  one undo step; a chart edit changes only its chart part and its workbook, and stamps no
  paraId (the document part's bytes stay).
- **Reading.** `state()`'s drawings carry `chart` (ooxml-edit's chart JSON: `types`,
  `title`, `axis_titles`, `legend`, `format`, `categories`, `series`), `diagram` (`layout`,
  `nodes` with `id`, `lvl`, `t`) and, for a group, `members`. `to_markdown` writes a chart
  as pptx-agent's outline line and a diagram as its node text, both **in the drawing's
  comment** -- the drawing stands in a paragraph's line, not as a block of its own, so
  Markdown's non-guarantees list it (*Guarantees, and explicit non-guarantees*) and the
  E2 reverse-parse check holds (a title full of `-->` too).
- **The registry.** docx-agent registers none of DrawingML's, the charts' or the diagrams'
  child orders: `tests/test_child_order.py` snapshots `CHILD_ORDER` in fresh interpreters --
  docx-agent's vocabulary alone, the subpackage alone, both in either order -- and holds
  every one of the subpackage's 26 sequences as the subpackage wrote it, the union nothing
  more (docx-agent's own 39 unchanged from before). The validity checks now hold chart and
  diagram parts to those orders too.

**Measured first** (`tools/charts_probe.py`, Word 16.106 for Mac, Dutch interface, under the
machine-wide lock, no clipboard; `tests/observations/charts-word.json`), on documents
built by `tools/chart_parts.py` and on the fixtures edited through the API above:

| Question | What Word does |
| --- | --- |
| Caches against the workbook (the cache `Q1`... `Cached`, 10-40; the workbook `W1`... `Booked`, 900s) | **draws the cache**, never the workbook, and **a save keeps the cache** as it was: Word does not refresh a cache from an embedded workbook on opening or saving. Its workbooks come back byte for byte |
| `c:autoUpdate` absent, `0`, `1` (embedded workbook) | the same drawing each way; **written back as `0`** every time |
| Extensions after docx-agent's edits | `c14:style`, the `c16r2` declaration and the `c16r3` display options kept; every `c16:uniqueId` kept; **a series docx-agent added (it writes none) gets a fresh `c16:uniqueId`** (`{00000000-...}`) on Word's save; caches, formulas and workbooks byte for byte |
| A title's form | no `a:defRPr`: **Arial 18 pt, not bold** (docx2svg agrees); an empty `a:defRPr` (ooxml-edit's default): **Aptos bold 18 pt** (docx2svg: 10 pt bold -- proposed); Office's chart-style form (docx-agent's template): **Aptos 14 pt, not bold, `595959`** (docx2svg agrees); a `c:title` with no text: the single series' name, Aptos bold 18 pt (docx2svg draws none -- proposed). Axis titles: the empty form Aptos bold 10 pt, docx-agent's Aptos 10 pt `595959` (docx2svg draws no axis title -- proposed). Word saves each form back as given, but drops a title paragraph's `a:endParaRPr` after its run |
| Tracking | with `w:trackRevisions` on and the edits made in docx-agent's tracking mode, **Word's review counts no revision** (`count of revisions` 0) for chart data or a diagram's text, and its PDF shows both: there is no revision form to write |
| Charts in a header, a footnote, a text box, a group, floating | every edit shows in Word's PDF; Word keeps every `docPr` id and **the group member's `wpg:cNvPr` id**; a text box's VML fallback repeats its chart's drawing (same id, same relationship); a group's fallback is a PNG of the chart, **not redrawn** when the chart changes (only readers of VML see it) |
| SmartArt: the data model's text changed, the drawing stale | the PDF shows the **data model**; the save rewrites the drawing from it |
| ... the drawing's text alone changed | the PDF shows the data model; the save rewrites the drawing |
| ... the drawing removed (relationship and `dataModelExt` too) | laid out from the data model; a new drawing written |
| ... a node added (no presentation points), the drawing dropped or kept stale | the node shown, laid out; its presentation points and the drawing written |
| ... a node removed with its presentation points, the drawing dropped or kept stale | the diagram without it; the drawing rewritten |
| ... a text edit the drawing follows exactly | shown; the drawing kept as written |
| Word's AppleScript dictionary | **no chart, chart-data or SmartArt object**: no class, no `chart data`/`activate chart data`; only `has chart`, the inline-shape and shape type enumerators, SmartArt node enumerations no class uses, and `run VB macro` (a macro the document would have to carry). So Word's own Edit Data view cannot be asked for -- the question pptx-agent left open stays open, and the evidence is the files, as there |

**The SmartArt policy, chosen on that evidence: drop.** Word, like PowerPoint, lays every
diagram out again from its data model on opening and rewrites the drawing on saving, so
for Word dropping and keeping a drawing the edit cannot keep exactly in step come to the
same. docx2svg draws only the cache: kept, it would draw the old diagram (old text, a
removed node, no new one) as if current; dropped, it draws a placeholder and warns
(`diagram-no-cached-drawing`) until Word saves the document again. So `DIAGRAM_POLICY`
is `drop`, as pptx-agent's, and `DiagramDrawingDropped` says so; `keep` and `refuse` are
there per call. A text edit the drawing can follow exactly (the node's `presOf` shape
holding exactly the texts it presents) keeps the drawing, patched, and docx2svg draws it.

**Tracking**: as measured, applied untracked, `UntrackedChartEdit` warning (Python's
`warnings`, as the subpackage's own `ChartDataWarning`: the chart API returns the chart,
not an `EditResult`).

**Edit Data, argued from the files** (pptx-agent's E4 reasoning, which this repeats with
Word): every edit is checked against an independent reading of the embedded workbook
(`tests/xlsx.py`, ooxml-edit's checker ported): every formula's cells hold what its cache
holds, tables cover the data and are named after their headers, shared-string counts are
right; Word keeps caches, formulas and workbooks byte for byte, and the check holds on what
Word saved. Edit Data opens that workbook.

**Fixtures** (`tests/fixtures/generated/charts`, `tools/make_chart_fixtures.py`, provenance
in its `PROVENANCE.md`): written here, then saved by Word -- `charts.docx` (clustered
columns, a line, a pie, a doughnut, a scatter, columns with a line), `chart-places.docx`
(a chart in the header, a footnote, a text box, a group, and floating), `smartart.docx`
(Basic Block List, Vertical Bullet List, laid out by Word from data models alone). One
directory below the corpus, read by every reader suite (`conftest.reading_paths`).

**Tests** (239 new, and the reading suites on the three fixtures: 2,398 passing in the
default run, 1,662 skipped, 3 min 33 s under `pytest -n auto` on four cores, against 3 min
27 s before; nothing new under `--run-slow`). `tests/test_charts.py`: fifteen edits (a value, a blank, a whole series, a
category label, a series name, categories added last and first and removed, series added
last and first and removed, a title, an axis title, the legend, the JSON applied) on all
eleven charts of the fixtures -- every kind and every place -- each one undo step, edited,
saved, reopened and read back, the workbook checked independently, no validity problem added, only chart
and workbook parts changed, undone to the original bytes and redone to the edited ones;
docx2svg's render of four of them on every chart (the new label drawn, the value axis
reaching the new value, no new warning); refusals change nothing; a linked workbook is
edited in its cache with ooxml-edit's warning; tracking warns and writes no revision; the
title's form and language; `to_markdown`'s summaries, read back; `state()`'s JSON and its
application; a chart in a nested group (`d:5/9/7`). `tests/test_diagrams.py`: text edits the drawing follows, nodes added at the
top and as children, removed, the JSON applied -- the same gates, the drop's warning, the
part reaped, docx2svg's placeholder -- and `keep`, `refuse` and the last node.
`tests/test_child_order.py`: the registry.

**The oracle** (`pytest -m oracle`, `tests/test_oracle_charts.py`, 9 tests, all passing --
467 with E0-E6's, all passing, 4 skipped as before; E2's runs on the chart fixtures too, as
reading fixtures: Markdown inserted into each exports, survives a re-save and reviews as
ours -- only its page-by-page character comparison leaves them out, since docx2svg's lines
hold no chart text and `chart-places`' floating drawings are paged apart, proposal 7).
Every chart fixture with an edit set on every chart (its first value ten times above its
largest, a category relabelled and one added, the first series renamed and one added, a
title) and the SmartArt fixture with its own (a text set, a node added at the top and one
as a child, a node removed) **export from Word unprompted**; the PDF has every new label,
series name and title, **every value axis reaches the new value**, the removed node is
gone -- each read inside its own chart, wherever Word put it. **Word's chart text is docx2svg's** by docx2svg's F.20 method -- each span Word draws
in a chart (string, face, weight, device size, the anchor of its advance box and its
baseline, within 0.5 pt, each chart's text taken against its title so that where the page
puts the chart does not count), and whatever docx2svg draws that Word does not: **85 of 85
in `charts.docx`'s six charts**, every text of the group's and the floating chart, and the
header's but for five (below); Word's fractions are written with its host's decimal comma,
which the comparison reads as a point. **A Word re-save keeps** every chart's data, formulas
and workbook (byte for byte), and the workbook check holds on Word's file; what Word
changes is `c:lang` (its interface's) and a `c16:uniqueId` for each added series. Word
**rewrites each edited diagram's drawing** with the data model's texts. Tracked, Word counts
no revision.

**Approximations, and why:**

- *A new title's form* is Office's default chart style's (14 pt, not bold, `595959`, the
  minor face); Word's AppleScript cannot add a title, so what Word writes for one it adds
  was not observed -- the form was measured to draw as Word's own and to survive a save.
- *A new series* has no `c16:uniqueId` until Word saves (Word then gives it one); no
  reader is known to need it before.
- *A group's VML fallback picture* is not redrawn after a chart edit (neither does Word).
- *Word's `c:lang`* is its interface's; docx-agent leaves the chart's own, and writes the
  document's default language into new title text.
- *A cache-only chart* (linked, OLE, missing workbook) is edited in its cache, with
  ooxml-edit's warning: Edit Data then shows the old values (by construction, not measured
  in Word).

**Found on the way:** the Markdown reader wrote a note's paragraph that holds only a drawing
as a bare comment line, which markdown-it reads as an HTML block (the reverse-parse check
failed on a chart in a footnote); fixed. **Clean-ups:** docx-agent no longer gives docx2svg
Office's cloud-font folders (docx2svg searches them itself; every page of the corpus, the
chart fixtures and a new document's Aptos Display heading renders the same), and the
Markdown reader counts list items through docx2svg's public `ListCounters.item(numbering,
numId, ilvl)` over `parse_numbering`, no stand-in document or paragraph (the counts against
docx2svg's recording of Word hold).

**Proposals** (not made here). For docx2svg -- (1) **a chart in a footnote** is related
from the main part (`chart-unreadable:no c:chart`): a note's drawing should resolve against
the notes part; (2) **a chart inline in a text box** is neither drawn nor warned; (3)
**axis titles** are not drawn; (4) **a title with no text** (`c:title` alone) Word draws as
its single series' name, bold 18 pt, docx2svg not at all; (5) **an empty `a:defRPr`** in a
title, in a chart whose `c:txPr` states no size, Word draws at 18 pt bold, docx2svg at
10 pt (F.20's "c:txPr's size or 10 pt" holds where `c:txPr` states one); (6) in **a short
chart** (126 x 63 pt, the header's) docx2svg sets the category labels 1 pt lower against
the title than Word and draws a `0` on the value axis Word leaves out; (7) **floating
drawings wrapped top and bottom**, anchored at offset 0 in consecutive paragraphs
(`chart-places.docx`), Word stacks below their anchors' text onto two pages, docx2svg onto
one. For ooxml-common -- (8) **a dropped diagram drawing**: `diagram_drawing_part`'s last
fallback hands a diagram without `dataModelExt` the part's only other diagram drawing,
which belongs to another diagram (docx2svg then draws that diagram twice and warns of
nothing; `tests/test_diagrams.py` pins it): it should skip a drawing another data model
names. For ooxml-edit -- (9) `default_title_text` puts `lang` on an empty run, which
`replace_body_text` drops when it writes the title's text, so a new title's run has no
language (docx-agent's template carries it on `a:endParaRPr`, where the text takes it
from); (10) a new series could get a `c16:uniqueId` as Word gives one, so a document
docx-agent saved has the ids Word would write.

---

## Usability (end-to-end pilot)

### The pilot (2026-10-04)

A Sonnet agent did a real task with docx-agent as a black box, from the README, `help()`
and `dir()` only: apply five reviewers' comments to a supplier agreement summary
(`tests/fixtures/generated/pilot`, written by docx-agent) as tracked changes by "Claude",
reply to each and resolve it. It succeeded, first run, but found the API slowly. Its
friction log, in order:

1. **No README example for comments or tracked changes**: the whole surface found by
   guessing names in `dir(Document)`.
2. **`track_revisions` / `set_track_revisions` read as "record my edits as tracked
   changes"**; they only set Word's `w:trackRevisions`, for a person's later edits in Word.
   The mechanism is `doc.tracking()`. One failed attempt.
3. **`Comment.paragraph_ids` is the comment's own paragraph** (`comments/p:...`), not the
   text it is attached to; nothing said where a comment is anchored. Found through
   `to_markdown(view="markup")` after a dead end.
4. **`render_png(view="markup")`** raised docx2svg's bare `ConvertOptions` `TypeError`,
   naming no accepted option. Renders show the final view only, by decision ("How docx2svg
   should render revisions").
5. **`docx_agent.validate.check` takes a package** the `Document` API does not offer; the
   README's validity guarantee could not be checked from the public surface.
6. **A tracked replacement records the whole anchor** (`"a period of 24 months"` deleted
   and inserted): a tight anchor is the caller's job.

### What changed

- **README, "Common tasks"**: a recipe for each thing an agent does most -- reading in
  every view and editing by id, anchors and replacing, tracked changes, comments,
  revisions, `insert_markdown`, tables, pictures, charts, sections, headers and footers,
  the table of contents, renders and `layout().where`, `Document.new` and templates, saving
  and validating. `tests/test_readme.py` runs every block as written, on the pilot's
  document and three other fixtures, so the recipes cannot rot. The phases' examples
  follow as "A tour of the API", illustrative.
- **`doc.word_tracks_changes`** (and `set_word_tracks_changes`) is Word's switch; the old
  names are deprecated aliases raising a `DeprecationWarning` that points to
  `doc.tracking()`, and the two docstrings refer to each other.
- **`Comment.anchor`**: the `TextRange` the comment's markers hold, in whatever story it
  covers (its text, id and paragraphs); a reply with no range of its own is attached where
  its thread is; a comment on a place is an empty range. `state()` gives each comment an
  `anchor` (`{"id", "t"}`, in the state's view). The comment's own paragraphs are
  `content_paragraph_ids`; `paragraph_ids` is its deprecated alias.
- **`doc.validate()`** returns the problem list (`docx_agent.validate.check` on the
  package); `save(..., validate=True)` refuses to write a document it finds problems in.
- **Render and layout options are checked** against docx2svg's `ConvertOptions` before
  docx2svg sees them; an unknown one raises a `TypeError` that lists the accepted ones, and
  for `view=`-like names says renders show the final view and points to
  `to_markdown(view="markup")`. The render docstrings say so too.
- **Tight changes, documented**: `TextRange.replace` says a tracked replacement records the
  whole range and points to `Paragraph.set_text`, which records only the words that
  changed; the README's tracked-changes recipe says so.
- **Docstrings** on every public class, method and property an agent reaches, with a
  one-line example where it helps.

### Run again, from the README alone

An independent agent, given only the new README and `help()`/`dir()` (no source), wrote the
task again: 106 lines, every comment located through `comment.anchor`, the indexation
change narrowed to `"3%"` because `replace`'s docstring said to, `validate()` clean. It
reported three more frictions, all fixed: `doc.comments()` lists replies alongside
threads (now `comments(replies=False)` gives each thread once); `Comment.reply` took
`**who` (now `author=`, `initials=`, `date=`); nothing said whether comments are tracked
(the README now says they are not, inside `tracking()` or out). Tightened, it is
`examples/review_comments.py` (48 lines, against the pilot's 69, and it finds each
comment's place through the API where the pilot hard-coded ids and anchor strings);
`tests/test_examples.py` runs it on the pilot's input.

### Left, with reasons

- **A markup render** (insertions underlined, deletions struck, comment balloons): out by
  decision; `to_markdown(view="markup")` is the review surface.
- **`replace` diffing old against new itself**: it would change the measured edit model
  (the replacement takes the formatting of the match's first character; the tracked forms
  are held to Word's), so it stays a caller's choice between a tight anchor and
  `Paragraph.set_text`.
- **Where a tracked replacement of exactly a comment's text puts the insertion**: after
  the comment's reference, so the comment's anchor then holds only the deleted text (its
  current-view text is empty). Where Word puts it is not measured yet (no Word in this
  pass); measure it before changing it.
- **Whether to turn on Word's own Track Changes switch** for the person who opens the
  result: a judgement for the task, not the library; the README says what the switch does.

---

## Trial findings

### The trial (2026-10-05)

Twelve tasks (six for Word), each run twice by an independent Sonnet agent with docx-agent as
a black box (README, `help()` and `dir()` only), graded by checks through the public API,
by Word itself and by visual graders. All twelve Word runs succeeded and Word opened every
output unprompted, but the docx-agent findings below cost the agents time, and one defect
was latent everywhere: a template opened and saved as a `.docx` kept the template's content
type, which Word refuses while `validate()` said nothing (its pptx-agent twin failed a
PowerPoint run). The merge-a-section task's inputs and the reproduction are in
`tests/fixtures/generated/trial` (written by docx-agent; `PROVENANCE.md`).

### What changed

1. **Templates (S1).** `save()` sets the main part's content type from the extension:
   `.docx` and `.docm` are documents, `.dotx` and `.dotm` templates; nothing else changes,
   but a VBA project written to a `.docx` or `.dotx` is dropped (`MacrosDropped`).
   `save_as_template` takes the template extension and refuses a document one.
   `validate(target=)` reports `content-type-extension` (and `macros-in-macro-free-file`).
   `Document.open` on a template warns `TemplateOpened`, pointing to
   `Document.new(template=)`; the README says opening a template edits it and `new` makes a
   document. Measured (`tools/trial_probe.py`, `tests/observations/trial-word.json`): Word
   refuses a template's type in a `.docx`, a document's type in a `.docm` and the
   macro-enabled type in a `.docx`; its own Save As `.docm` writes the macro-enabled type
   with no VBA project, which opens. The oracle (`tests/test_oracle_trial.py`) opens the
   trial's file saved again as a `.docx` and as a `.dotx`, both unprompted.
2. **`insert_markdown(style_map=)`** takes a plain dict (`{"h1": "Report Title",
   "paragraph": "Report Body"}`; keys `paragraph`, `h1`-`h6`, `quote`, `code`,
   `inline_code`, `bullet`..., `number`..., `table`, `table_cell`, `footnote`,
   `footnote_reference`, `emphasis`, `strong`, `link`). `StyleMap`, `Rule` and
   `DEFAULT_STYLE_MAP` (also `StyleMap.DEFAULT`) are exported. `paragraph` is the body's
   paragraphs only: table cells (`table_cell`) and a footnote's text (`footnote`) keep the
   defaults unless named. A mapped built-in body style the document lacks is added, where
   it was silently Normal. A README recipe.
3. **Copying in the destination's styles.** `copy_blocks(style_map={source name: destination
   name}, unmapped="body")`: a source style neither mapped nor in the destination by name
   becomes the destination's body style (the style most of its plain body paragraphs use) --
   a built-in one is added as Word writes it -- instead of being imported. `styles="merge"`
   keeps a heading's level and a list item's list (a deliberate departure from Word's Merge
   Formatting, which flattens them; `test_e6_copy` says so). `Styles.remove(name,
   replacement=)` refuses a style in use without a replacement; `Styles.usage(name)`;
   `Styles.purge_unused()` removes the custom styles nothing uses (a linked pair together).
   The trial's W6 inputs are the test (`tests/test_trial_copy.py`), and Word exports the
   merged handbook.
4. **`to_markdown`.** `stories="all"` (or a list of names) adds every header and footer
   with content and every note, with their comments in the markup view; the markup view
   marks a field's result (`<!-- field: REF ... -->...<!-- /field -->`); a heading's id
   comment says `numbered: list` or `numbered: text`. "Guarantees, and explicit
   non-guarantees" says what is still not shown.
5. **Moving a section.** `section_blocks(heading)`: the heading and everything to the next
   heading of its level or higher. `move_blocks("p:A..t:B" | [ids], after=|before=)`: one
   undo step, every block keeping its id; tracked, Word's moves (a table as a deletion and
   an insertion), and Word exports a tracked section move unprompted.
6. **Revisions.** `doc.changes(author=, kind=)` groups the records into `replacement`,
   `insertion`, `deletion`, `move`, `formatting`, `table` and `section` changes, each
   accepted or rejected whole; `Revision.kind` documents every kind (`REVISION_KINDS` in
   `docx_agent.revisions.changes` says what each is, `paragraph-mark-insertion` and Word's
   `paragraph-properties` record at a story's end among them);
   `TextRange.delete(collapse_space=True)` deletes the doubled or stranded space too.
7. **Charts.** `chart.workbook_values()`: the embedded workbook's cells for the chart's
   ranges (series names, values, categories, with their formulas), read-only.
8. **API shape.** `EditResult.object` (`doc.get(result.id)`) for every edit; `doc.stories`,
   `doc.styles` and `Note.paragraph_ids` answer as properties and as calls, and
   `doc.styles["Heading 1"]` finds one; `format_cell` is `set_cell` by the name that says
   it formats (both kept); `Picture.image` is a property, the old call deprecated.
9. **Small items.** Markdown at the end of a body that is one empty paragraph (a new
   document's) replaces it, as typing does (ending in a table, it stays after the table;
   tracked, rejecting gives it back); `_Ref` bookmark names are taken; `bookmarks()` says
   it leaves out the hidden `_` ones unless `hidden=True`; `Document.new(created=)` takes an
   ISO 8601 string as `tracking(date=)` does.

### Proposals for the other repositories

- **ooxml-edit**: `Chart.workbook_values()` belongs in `ooxml_edit.charts.chart.Chart`
  (pptx-agent's charts want it too); docx-agent's is written on ooxml-edit's reading
  helpers (`workbook_part`, `Workbook`, the series' formulas) and moves there unchanged.
  Done in ooxml-edit 0.2.2, merged with pptx-agent's (its `sheet` added; a rectangle reads
  as rows; a formula that is not one range or names a missing sheet reads `None`);
  docx-agent's `Chart` inherits it.
- **ooxml-common / ooxml-edit**: the kind-from-extension rule (`kind_for`, `write_as`,
  the content-type check) is the same for every OPC format; pptx-agent's `.potx` twin of
  finding 1 wants it, so a shared home in ooxml-edit's package layer would serve both.
  Done, split by what each part is: the facts (extension, kind, template and macro flags,
  main content type, `kind_for`, `kind_mismatch`) are standard-library data and went to
  ooxml-common 0.4.6 (`ooxml_common.kinds`), where ooxml-edit's format-neutral core could
  not name them; the retyped copy of `[Content_Types].xml` went to ooxml-edit 0.2.2
  (`OpcPackage.content_types_with`).  `TemplateOpened`, `MacrosDropped` and the validate
  codes stay here, as does what makes a VBA project Word's (its key maps and `wordVbaData`).

### Left, with reasons

- **A `.docx` written with a macro-enabled document's macros dropped** follows Word's Save
  As (which asks first); refusing instead is a choice docx-agent did not take, since the
  extension says what the caller wants.
- **Field marks in the `final` view**: left out by decision; the final view is plain text,
  and the markup view is the review surface.

---

## Agent tools (the tool layer's T2)

### Word core — done (2026-10-07)

`docx_agent.tools` is the Word half of the agent tool layer (the tool-layer roadmap's T2), on
`ooxml_edit.tools` 0.4 (the shared tool definitions, refs and `batch`): the docx
`DocumentFormat` (open, detect and validate from bytes, the error map with valid options,
the warnings collected, a prompt fragment of mechanics only, `summary` and the per-edit
`checks`), the Word handlers of the shared tools S1-S13 and the Word tools W1-W27 and W29.
W28 (`word_insert_chart`) waits for a new chart from data (LE3, LW3); `word_template`'s
`apply_styles_from` for LW5.

- **Lengths are points, everything is in memory:** no tool takes a path; inputs are blob
  handles, and saved files go to the application.
- **Tracking mode:** `word_set_tracking` keeps the author on the session's entry and every
  later edit runs in `doc.tracking(author=..., date=<session clock>)`; comments are never
  tracked.
- **Layout in the worker pool:** every layout goes through `Document.converter` (new: a
  callable laying bytes out elsewhere, `docx_agent.layout.convert_bytes`), which the tools
  point at the toolbox's process pool under `Limits.layout_timeout` (30 s). On the 36-page
  `sample-long.docx` a reflow check takes 3.7 s (6.9 s with the worker's start), a new TOC
  6.0 s and updating the fields 3.5 s; a timeout rolls the call back.
- **Checks in every changing result:** the `validate()` delta and, when the state before
  was laid out, the pages reflowed (20 pages or fewer; `check` for longer documents).
- **Plural inputs and refs:** `word_set_text`, `word_insert_text` and `word_comments` take
  `items`, `word_format` and `word_lists` `targets`, `word_edit_table` many `cells`; created
  objects take a `ref`, and targets accept `$name`, inside a `batch` too.
- **Library changes for it:** the empty-TOC warning (trial 2's N8: `insert_toc` and
  `update_fields` warn while a TOC has no entries) and `Document.converter`. In ooxml-edit,
  a batch that changes nothing records no undo step (so a list action records none).
- **Golden transcripts** (`tests/goldens`): the trial's Word tasks w1-w6 as tool calls only,
  replayed to byte-identical outputs that pass the trial's checks (25/25, 40/40, 14/14,
  17/17, 20/20, 24/24), and one `batch` of 30 tracked insertions and 10 comments; Word opens
  every output (`tests/test_oracle_tools.py`).
- **Budgets:** core definitions about 2,900 tokens and all 43 about 11,500, by the offline
  estimate in `tests/test_tools_schema.py` (the online count waits for an API key).

Left: no tool reads a text blob (a CSV the user supplied reaches the model only through the
application's message); `LW6` (`doc.outline()`) stays composed in the tools; field staleness
(LW2) and layout facts (LW1) are T4.

---

## Suggested order

```
ooxml-edit ─▶ E0 ─▶ E1 ─▶ E2 (read) ─▶ E3 ─▶ E2 (write) ─▶ E4 ─▶ E5 ─▶ E6
                      └─▶ cross-document copy (E6's import machinery), whenever needed
```

1. **ooxml-edit, then E0.** Nothing else is testable without a lossless package, ids and
   the paraId measurement.
2. **E1**, with the primitives designed for tracking from the start.
3. **E2's read half** (`to_markdown`, JSON) early: it is the agent's eyes, it is cheap, and
   it changes nothing.
4. **E3 before E2's write half**, so `insert_markdown` is trackable the day it lands, and
   because legal and business documents — the ones agents are asked to edit — arrive with
   revisions already in them.
5. **E4, E5, E6** in that order: structure is what most edits disturb; tables and drawings
   are the largest surface; authoring needs all of them, and template work needs E6's
   import machinery anyway.

Cross-document copy could move earlier (after E1) if agents need to assemble documents from
pieces sooner; it depends only on E1 and ooxml-edit's part copying.

---

## Testing strategy

Five layers, in decreasing order of authority and increasing order of availability:

1. **The Word oracle** (`pytest -m oracle`, opt-in, as in pptx-agent). Word opens the edited document and
   exports it to PDF; no repair prompt; the PDF's text and page count are checked, and
   rasterised pages where a visible marker is the evidence (pptx-agent's method). And
   **re-save**: Word saves the file again as `.docx` (docx2svg's
   `tools/word_save_docx.applescript`), and the test diffs what Word changed — ids kept or
   rewritten, revisions kept with our author and date, styles untouched, numbering intact.
   The re-save diff is how "Word accepted it as it is" is told apart from "Word silently
   repaired it". The scripts are docx2svg's, found through a sibling checkout or
   `DOCX2SVG_ORACLE_SCRIPT`, as pptx-agent finds pptx2svg's.
2. **Structural validity** (`tests/test_validity.py`), everywhere, after every edit set on
   every fixture: well-formed parts; children in schema order; relationships resolve and
   every part has a content type, no `Override` without a part, no orphaned part; paraIds
   unique per part and in range, each `textId` with a paraId; `docPr` ids unique; bookmarks
   paired, names valid and unique; annotation ids unique; comment markers and references
   paired; every note reference has its note; fields balanced (begin, separate, end); no
   `w:t` inside `w:del` and no `w:delText` outside; move ranges paired; **numbering
   integrity** (every `numId` resolves, every abstract definition exists, levels in range);
   style references resolve and `basedOn` chains are acyclic; every `w:tc` ends in a `w:p`;
   the body's `w:sectPr` is last; `settings.xml` in order. Optionally, schema validation
   against the ECMA-376 transitional XSDs with the MS-DOCX extensions (a development tool;
   fetched for a development-time check only (decided)).
3. **Round-trip fidelity** (`tests/test_roundtrip.py`), the gate everything rests on:
   byte-identity for untouched parts, canonical-form equality for edited ones, undo to the
   original bytes after every edit set (one batch and step by step), redo to the edited
   bytes. Plus E3's accept/reject invariant and E2's Markdown AST round trip.
4. **docx2svg render checks**: every edited fixture renders; the SVG carries the API's ids
   on every drawn paragraph; `layout().where` agrees with the SVG's lines; the warnings are
   no worse than the unedited fixture's (an edit must not make docx2svg stop earlier unless
   it added the obstacle). Pixel comparison against Word is docx2svg's job, not this one's.
5. **Unit and property tests** per module, including randomised sweeps (find/replace across
   boundaries, table merges, revision accept/reject).

### Word is one instance per machine

docx2svg's measuring work drives the same Word, and its recovery path `pkill`s Word and
deletes `~$` lock files — which would destroy a docx-agent export in flight, and vice versa.
So:

- **The default run and the slow run.** `pytest -n auto` holds every gate on every
  fixture, but the heaviest per-fixture suites (E1-E5's operations, renders and tracked
  invariants, the Markdown corpus and tracked insertions, E6's copy and upgrade sweeps)
  on four representative fixtures (`conftest.REPRESENTATIVE`) and the randomised table
  sweep on its first seeds: about three and a half minutes on four cores. `--run-slow`
  runs those on every fixture and seed (about fourteen minutes); nothing else differs (E6).
- **Never in the default run.** `pytest` runs layers 2–5; `pytest -m oracle` adds the oracle,
  and never under xdist's parallel workers: the default run is xdist-safe (`pytest -n auto`,
  `pytest-xdist` a dev dependency since E2's write half), and on a worker the oracle's tests
  skip (`tests/conftest.py`); run `pytest -m oracle` without `-n`.
- **One machine-wide lock.** docx-agent's oracle helper takes an advisory `flock` on one
  well-known file in the Office group container (Word's sandbox already reads and writes
  there) before it launches, exports, saves or recovers Word, and releases it after.
  docx2svg takes the same lock around its exports and recovery (since its commit
  `ae0eb95`: blocking, for up to 600 s) and never quits a Word it did not start.
  docx-agent's helper still **refuses to start if Word is already running** — the sign a
  person, or a tool that does not take the lock, is using it — and skips with "Word is in
  use" rather than wait, kill or share.
- **Recovery only under the lock.** Dismissing a repair dialog, quitting Word and deleting
  `~$` files happen only while holding it.
- **Built in E0** (`tests/oracle.py`): the lock is `word-oracle.lock` in the Office group
  container, taken non-blocking; a held lock or a running Word skips the oracle ("Word is
  in use"); recovery quits only the Word the session started (asking first, killing only if
  it will not quit) and deletes `~$` files only in docx-agent's own staging directory,
  `docx-agent-oracle/` in the group container, where answers are cached by content.
- **Cached by content.** Exports and re-saves are cached by the SHA-256 of the input
  `.docx`, outside the repository (a Word PDF carries font subsets and is never committed),
  so a re-run that changed nothing needs no Word, and the oracle can be run in one batch at
  a quiet moment.
- **Paths under the sandbox.** Inputs are staged in the Office group container, as
  docx2svg's are.

### Fixtures

- **Reused from docx2svg's committed fixtures**, with their provenance notes copied
  alongside: `layout-sweep.docx` and `style-document.docx` (generated by docx2svg's own
  scripts), the four `samplelib/` documents (licence: "do whatever you want with the
  files"), and the five `wordto/` documents ("free to download and use for any purpose").
  Between them: compatibility modes 14 and 15, Word 14 and 15 authorship, tables, an
  image, multi-page flow.
- **Generated here**, by committed scripts and then opened and saved by Word so they are
  Word-authored (pptx-agent's SmartArt-deck method): tracked changes of every kind by two
  authors, comments with replies and resolution, footnotes and endnotes, several sections
  with distinct headers and footers, nested and merged tables, content controls (one
  data-bound), fields and a TOC, floating drawings, a localised (Dutch) template with
  translated style ids, a `.dotx`, and a document with no paraIds and one with duplicated
  ones.
- **Mixed provenance**, authored here so the licence is ours: the same content exported by
  LibreOffice and Google Docs, because the id edge cases (no paraIds, duplicates, odd
  numbering) come from other producers, as pptx-agent's duplicate `cNvPr@id` came from a
  Google Slides deck.
- **Only committed, licensed documents are used.** Material docx2svg keeps outside git is
  out of bounds for this project's tests, its documentation and its commits.

---

## Non-goals

- **Rendering.** docx2svg's job; docx-agent depends on it and maps ids onto it.
- **The binary `.doc` format**, and writing any compatibility mode below 15 for a new
  document.
- **Editing a document that is open in Word** (co-authoring, AutoSave). docx-agent edits
  files at rest and warns when Word's `~$` owner file sits beside the input.
- **Macros and VBA.** A `.docm`'s project is kept untouched; nothing reads or writes it.
- **Authoring data bindings, XML mapping, mail merge, forms protection, digital
  signatures** (an edit invalidates a signature; the API says so instead of pretending).
- **A full-state SVG**, and importing documents from SVG, HTML or PDF.
- **Markdown as storage**, and diffing an edited projection back into the document.
- **Document comparison** (PowerTools' `WmlComparer`); a later candidate, not a goal.
- **Pixel-exact layout claims.** The fidelity target and its evidence live in docx2svg.

---

## Risks and decisions

**Risks**

1. **paraId durability.** Measured in E0: Word keeps ids it did not write when the document
   is complete (every paragraph, row and drawing with its ids) and writes none at all when
   it is not, so the first edit stamps the whole document. Remaining risk: a document
   another tool breaks after us loses its paraIds at Word's next save (re-anchoring by
   `textId` and text would recover; not built), and a Word version that behaves otherwise
   (measured on 16.106 for Mac only; re-run `tools/paraid_probe.py`).
2. **Layout coverage.** Reflow feedback is only as far as docx2svg's layout reaches; real
   documents with features it stops at give partial answers. That is honest, but agents will
   meet it, and docx2svg's priorities may need to follow docx-agent's corpus.
3. **Revision edge cases.** Paragraph-mark deletions, moves, revisions inside tables and
   fields are where every implementation surveyed has had bugs. The invariant tests and the
   Word re-save diff are the defence; the measurements come first.
4. **Word-oracle contention.** Shared with docx2svg's measuring; the lock needs docx2svg's
   cooperation to be complete.
5. **Template variety.** Localised style ids, missing built-in styles, list definitions
   shared by unrelated lists: the style map must fail loudly rather than invent formatting.

**Decisions** (the user, 2026-10-03, on the questions this draft raised)

1. **Durable ids after a Word save:** E0 measures whether Word rewrites ids it did not
   write; the way on (re-anchor by `textId` and text, hidden `_dxa_` bookmarks as a second
   stamp, or both) is chosen on that evidence. *Outcome (E0):* none of the three -- Word
   keeps every paraId of a complete document, so the first edit stamps the whole document
   (*Durable across a Word save — measured*).
2. **Markdown parser:** markdown-it-py (with `mdit-py-plugins` for footnotes and task lists)
   as a runtime dependency. pptx-agent's planned E6 outline shares the choice.
3. **Default style mapping:** Word's own styles — block quotes to "Quote", code to "HTML
   Preformatted"/"HTML Code", emphasis through the Emphasis/Strong character styles (not
   direct `w:i`/`w:b`), so the formatting follows the document's styles.
4. **A from-scratch document's theme and styles:** written here from measured facts, as
   pptx-agent's `Document.new()` does; never copied from the user's `Normal.dotm`.
5. **Field updates:** TOC and page references are computed here from docx2svg's layout;
   `w:updateFields` is not set (Word would prompt on every open).
6. **rsids:** none written.
7. **Tracked column inserts and cell merges:** supported, **never refused** — refusing would
   leave the agent unable to restructure a table in a tracked document. Measure what Word
   writes; where it has no revision form, use the nearest accepted form plus a describing
   comment (see the tracked-change table above).
8. **docx2svg changes:** onto docx2svg's roadmap; the tracked-change final-view fixes
   (`w:moveTo`, deleted paragraph marks) and inline content-control text at high priority,
   as visible rendering bugs. Tracked changes and comments are not visualised: no
   original or markup view. The shared Word lock and the exported path function follow
   when docx-agent needs them.
9. **docx2svg is a hard dependency** (standard-library only, so the cost is small).
10. **Schema validation:** the ECMA-376 and MS-DOCX schemas are fetched for a
    development-time check only, never shipped.
11. **Charts in Word:** wait until pptx-agent's chart and workbook code moves into a shared
    package; no separate chart editing here before that. *Outcome:* it moved into
    ooxml-edit 0.2's optional `ooxml_edit.charts`, which docx-agent now hosts in Word's
    drawings (*Charts and SmartArt*); no chart code of its own.
