# Guidance for the application's thinking layer (Word documents)

`docx_agent.tools` gives a model the means to edit a Word document and the facts about the
result. It ships no judgement: no house style, no tone of voice, no rules about headings or
length. What a good document is for *your* users is the application's to say, in its own
prompt and its own review pass. This page is for the developer of that layer. The numbers
come from trial 3 (Sonnet 5.5), its fix round and the Haiku 5.5 trial; the tool roadmap in
`ooxml-edit` (`docs/TOOLS-ROADMAP.md`) has the detail.

## The facts the tools return

| Fact | Where | What it says |
|---|---|---|
| `validate` | every changing call, `check` | validation problems the call added (`new`) or removed (`fixed`) against the document as opened |
| reflow | changing calls (documents of 20 pages or fewer, once laid out), `check` with `reflow` | the pages the change moved; whether layout completed (`complete`, `stopped`) |
| `coverage` | `check` with `reflow`, `render`, `save_document` | how much of the document the layout covers. `complete` is false when a block, a header or footer, or a face could not be laid out. `status` is `complete`, `approximate` (everything laid out, but `approximations` lists places laid out by an unmeasured rule, with their page and reason) or `partial`. It also gives the blocks laid out, `placeholders` (what docx2svg drew as a marked box of its height instead of laying it out, with page, reason and block id: the layout goes on past them), where the layout stopped if it did, and the faces substituted or missing. |
| layout | `check` with `layout` (or `min_size`, `styles`) | what the pages show, measured from the layout reflow uses: the smallest text per kind and, with `min_size` (or the application's `min_text_size`, per kind), text below it; direct formatting over the style, with the `word_format` `clear` call that takes it off; with `styles`, paragraph and table styles not in the list; empty paragraphs used as space (with the `word_delete` and `space_before` calls that replace them); `empty_paragraphs`, per cell and for the body (`at`), the empty ones between two paragraphs of text (`paragraphs` their positions, `ids`, `fix` the `word_delete` calls; a body run is also among the spacers); keep-with-next paragraphs left at a page's foot; `table_parted`, a table on another page than its caption, lead-in or source line, with the keep that joins them; `numbering_skips_level`, a style numbered at a list level whose label shows a level above that nothing uses (`1.1` with no chapter); rows split across pages and header rows that do not repeat; `column_fit`, a table column whose text uses well under its width beside one that wraps heavily, with widths that fit both and the `word_format_table` calls; pictures and tables past the margins; and `uncertain`, the pages with placeholders or approximations. Each list at most 20, most telling first, `<fact>_left_out` the rest; `pages` keeps the facts on those pages. Facts, not verdicts: which of them matter is the application's rule. |
| `numbers` | `check` with `include: ["numbers"]` (opt in; `range` or `pages` narrow it) | per body paragraph or table cell (`at`, its `page`): `unverified`, decimals and percentages that no value of the session's datasets or open workbooks' tables matches within the rounding the text shows (`2,7` matches 2.66), each with the two nearest values (`value`, `from`, `sources`); `typed`, numbers a value matches that were typed, with `from` and, when it writes the same number, the `fix` -- a `replace_text` call writing it by reference (`fix_hint`: several in one `batch`). Unverified is not wrong: say what the data hold. Left aside and counted (`left_aside`): dates, times, years, list numbers, ids, whole numbers (the application's setting), the application's ignore list. Not read: text written by reference (kept on its paragraph, `Document.text_facts`), tables placed with `from`, a table of contents, deleted text, headers, footers and notes. At most 20 texts, 5 numbers a list (`texts_left_out`, `more`) |
| fields | `check` with `fields` | the fields (TOC, page numbers, cross-references, captions) and their cached results: a cache is only as fresh as the last `word_fields update` |
| `warnings` | every call | what the library noticed and did (an unknown construct kept, a style imported, a legacy option removed) |
| `renamed` | changing calls | old paragraph id -> new one; the old id keeps working |
| markup | `word_read` with `view: "markup"` | the tracked changes and comments as Word shows them in review |
| tracked changes | `word_changes` `list` | each revision: author, date, kind, old and new text |
| comments | `word_comments` `list` | each thread: anchor text, author, replies, resolved or open |
| formatting | `word_inspect` | a block's style, its runs' direct formatting, the effective font, size, spacing and indents |
| design facts | `word_design_facts` | for analysing a template: each style in use per role (body, headings by level, captions, lists, table text, headers and footers) with its usage and effective look; each table style's header row, bands, first column, borders and text as drawn; each body table's frame -- its width against the body between the margins and the width most tables share, its columns, the rows whose last cell is empty with their stated height; the page grid and header and footer variants per section; the theme's palette and fonts; block shapes repeated under headings, with example spans to copy |
| copies against their sample | `word_compare_blocks` | tables or paragraphs copied from a sample, compared with it (`reference`, here or in another open document): header rows, body rows, columns, merges, column widths (within 0.5 pt), table style, alignment, shading and borders, each cell's shading, borders and vertical alignment, each paragraph's style and direct formatting; every difference with the reference's value and, where a tool writes it, the calls that give it. Rows by the reference's structure: its header rows one to one, its last body row the pattern for any more. A cell's paragraphs pair by style and text, not position: a sample line the cell lacks is `missing_paragraph`, a line the sample lacks `extra_paragraph` (no fix: add or remove it yourself); a style fix only where the texts match, or the counts are equal |
| copies following a revised sample | `word_follow_sample` | the copies you name (tables or paragraphs made from the sample; a cell or a paragraph in a table names its table) take the sample's table, row, cell and paragraph properties, column widths, styles (by name from another document) and run formatting, rows and cells corresponding as `word_compare_blocks` pairs them; each keeps its text, rows, columns, merges, lists and its own emphasis (a word in italics). One undo step; `preview: true` changes nothing. Per copy: what changed, and `not_followed` -- each difference left with `why`, its `kind` (`own`: the copy's own text, lines, list; `changed_by_copy`; `structure`: columns, merges) and, where a tool makes it, `fix` calls; `report: "actionable"` leaves the `own` ones out, counted. Tracked: as Word records each, rows untracked (a warning) |
| `validate` on save | `save_document` | the save refuses new validation problems; only the application can allow them |

How to use them:

- **A check is only as good as its coverage.** A reflow or a render of a partial layout
  says nothing about the pages past the stop. When `coverage.complete` is false, report
  what could not be checked; don't present it as checked. When `coverage.status` is
  `approximate`, the pages are laid out but the places in `approximations` (and what
  follows them on their page) may be off: say so when a check depends on them.
- **Validation is the gate.** Every Word output in the trials opened in Word without a
  prompt, because the save refuses what Word would repair. Keep it on.
- **Fields are caches.** After moving headings or adding a caption or cross-reference, a
  model should run `word_fields update`; the shipped prompt says so, and `check` with
  `fields` shows what each field holds now. In the trials the models updated fields
  themselves.
- **Renders show the final view** (changes accepted, comments hidden). To see what a reviewer
  will see, read `word_read` with `view: "markup"`.
- **Formatting differences are facts too:** `word_inspect` shows direct formatting a pasted
  paragraph carries (font, size, colour on its runs) against its style; `word_format` with
  `clear_direct` takes it off, with `clear` only the properties named. `check` with
  `layout` lists every paragraph that differs from its style, with that call.
- **Templates whose body lines pass a margin** (a negative indent on Normal): tables as wide
  as those lines are `word_format_table` `width_of: "body"`, and `past_margins` does not
  count them; a cell style inheriting such an indent is named in `word_insert_markdown`'s
  warnings (its text runs past the cell, in Word too): map `table_cell` to a style that fits.
- **A table that should not split or pull the next block along:** `word_format_table`
  `keep_together: true` (scope table) keeps every row but the last with the next and
  releases the last, as Word reads row keeps; `check` reports it as the table's
  `keep_together`, not as direct formatting. Longer than a page, it starts a page and breaks.
  With `keep_with_next: true` as well, the last row keeps too: the table stays with the
  block after it (a table and the box after it move to the next page together). The
  paragraph between two tables must keep with the next as well (`word_format`), or the
  second table can still start a page alone: the call's warning names it. A label style
  that keeps every row's first paragraph is the same thing: `check` reports such a table
  with `with_next`. Chains of these longer than a page break where Word breaks them, and
  `keep_with_next_at_foot` says where.
- **Numbered headings from a template's list:** `word_styles` `modify` with `numbering`
  (`{"list": <its number or a style in it>, "level": n}`) links a style to a level, as Word's
  "Link level to style": the style that held the level is unlinked and its paragraphs lose
  their numbers (the warning counts them); `{"list": "none"}` unlinks, and stops a level
  the style inherits from the style it is based on (a warning names those when you link).
  `describe` shows a style's level, own or inherited. A level's indent is the list's:
  `word_lists` `format` sets `number_at`, `text_at`, `follow` and `tab_at` (a label wider
  than its hanging indent runs into the text: widen it). `check`'s `numbering_skips_level`
  names a level used while the level above never is.
- **Blocks that belong on one page** (a heading, its text, a table and its source line):
  `word_format` `keep_together: true` on the span (`p:A..t:B`) keeps them together as Word
  does -- lines kept, each block with the next, every row with the next -- and moves them
  to the next page whole where they do not fit; `false` releases what it wrote. On one
  paragraph it keeps that paragraph's lines. `check(range=[span])` says per span
  `one_page`, `kept`, the `breaks` in it and the `fix`; `taller_than_page` where no keep
  can hold it. Nothing is stored: name the spans each check. With the `fix` comes
  `if_kept` (round 40): the keep laid out -- the page the span then starts on, the share of
  the page before it left blank, whether the document grows -- and on a kept span that
  starts a page, `if_released`: what releasing it gives, with that call. Read both and
  decide once; setting and releasing the same keep in turn costs a layout and a check each.
- **A heading's range** (`h:p:A`): the heading and every block under it, to the next
  heading of its level or higher -- what `describe`'s `headings` count in `blocks` (with the
  `section` each is in, when there are several). Give it wherever a span is taken:
  `word_read`, `check` and `render` `range`, `word_format` `targets`, `word_delete`,
  `word_move` `blocks`, `copies`. A heading is a paragraph of the body itself with an outline
  level (a heading style's, or one it or its style states); a Title has none unless the
  template gives it one, and a heading inside a table or text box owns nothing.
- **Headings keep their content:** `word_delete` of a heading whose content would stay is
  refused (it would go under the heading before): delete `h:<heading>`, or `force` to keep
  the content. Changing a heading's level (`word_format` `paragraph_style`) warns which
  headings and blocks now go under which. `word_lists` `remove` on a paragraph its style
  numbers is refused, naming styles like it that number nothing: give one of those
  (`word_format` `paragraph_style`), or `force` to override the style.
- **`check`'s `layout.settings`** holds what the document is set to do (the tables kept
  together), not facts to act on: a summary counts none of it. A table kept together that
  is taller than a page is said there and, once, in `keep_with_next_at_foot`
  (`taller_than_page`), which is worth a look.
- **The template's own table style for new tables:** `word_styles` `set_default` (a table
  style) is Word's Table Design > Set as Default; `word_insert_markdown` and new tables use it
  before the most used one, and `list` marks it `default`.
- **The document's language** (what its text is proofed and spaced in): `word_styles`
  `set_language` with `language` (`fr-BE`, `nl-BE`...) sets it where the template states
  it -- its default and the styles in its own language (Normal's, the default's); styles in
  another language stay and are named -- with `name`, that style's alone. `check`'s `punctuation_spacing` entries of kind
  `language_mismatch` (French spacing in text tagged another language) carry that call as
  their fix: the spaces are the text's, not to be deleted.
- **Tab stops of a style** (a contents style's right-aligned page numbers): `word_styles`
  `modify` with `tabs` replaces its own stops; `{"position": pt, "alignment": "clear"}` clears
  one it inherits (without that, the inherited stop is the first one the tab reaches). A
  stop with no `position` is at the text column's end as the sections are now (worked out
  when written: widen the margins later and set it again).
- **Captions in Markdown:** a paragraph starting `: ` (or `Table: ` next to a table) is a
  caption paragraph in the Caption style (`style_map` `caption` names another); for a
  numbered caption, `word_fields` `insert_caption`.
- **Markdown columns aligned `:-:` or `-:`** are direct alignment that `check` reports unless
  `style_map` names `table_cell_center` / `table_cell_right` styles; the warning names the role.
- **A layout fact is only as good as the page it is on.** `layout.uncertain` names the
  pages the layout drew a placeholder on or laid out by an unmeasured rule: judge those
  from the facts, not the render, or say they could not be judged.

## A review pass

The application calls the tools itself, through `toolbox.dispatch`, after the model says it is
done:

```python
ALLOWED_STYLES = ["Normal", "Heading 1", "Heading 2", "Heading 3", "Caption", "Table Grid"]


def review(toolbox, session, doc, pages, house_rules, model_call):
    """Facts, the reviewer's view and renders of the changed pages, judged by the application's
    own rules in its own model call."""
    facts = toolbox.dispatch(session, "check", {"doc": doc, "include": ["validate", "reflow",
                                                                         "fields", "layout"],
                                                 "min_size": 9, "styles": ALLOWED_STYLES})
    markup = toolbox.dispatch(session, "word_read", {"doc": doc, "view": "markup"})
    images = toolbox.dispatch(session, "render", {"doc": doc, "pages": pages[:4]}).images
    return model_call(images=images, text=markup.data, facts=facts.data, rules=house_rules)
```

- **Run it as a step of the application, not a choice of the model.** In a production
  user's live runs on Azure OpenAI (Oct 2026, both libraries, 8 runs), gpt6-luna looked at
  the render in every run; gpt6.1-sol skipped "look at the render" in all four of its runs.
  A prompt that asks for a review is not a review pass: the application calls `render` and
  `check` itself after the edit loop (or requires a review turn before `save_document`), and
  treats what is still unresolved as blocking or reports it.
- **Coverage is one of those facts.** `save_document`, `render` and `check` with `reflow`
  return `coverage.status`: `complete`; `approximate` (laid out, with the approximations
  reported in `approximations`); or `partial` (part of the document could not be laid out,
  so could not be checked). Anything other than `complete` blocks the output or goes into
  the report: say what was not checked.
- **What to send:** the brief, the markup view (pages through `next_cursor` for long
  documents), the facts, renders of the pages the edit touched (1000 px wide: about 1,700
  tokens a page on Claude) and your house rules. Ask for findings with paragraph ids.
- **What to do with findings:** send them back as one user turn and let the model fix them
  through the tools; one more pass at most.
- **An agent with tools reads in one call.** `read_batch` (ooxml-edit 0.78.0) changes
  nothing (`mutates` is false) and is the `review` role's: give it to a read-only agent with
  the other reading tools, and to builders. Many reads -- `describe`, ranges of
  `word_read`, `word_inspect`, a `check` -- go in one call, each answered as alone within its
  own `max_result_chars`, the reply within 96,000 characters; what did not fit is listed
  with the call that asks for it. Production builders made 40-90 reads a run, a model turn
  each, before it.
- In trial 3 every Word run scored 10 without a review pass; the pass earns its cost on
  documents that go outside the company, or where your house rules are detailed.
- `check`'s `include: ["app"]` is reserved for a critique hook the application registers; it
  returns `app_findings: null` today. Run the pass yourself as above.

## Model routing

| Work | Model | Evidence (main arm, per run) |
|---|---|---|
| Word editing: a report from a template and Markdown, review comments, figures from a CSV, restructuring, selective review of tracked changes, merging a section from another document, a landscape page, a footer with page numbers, a footnote and cross-reference (w1-w9) | **Haiku 5.5** | 18/18 succeed, every one graded 10; USD 0.006 against Sonnet 5.5's 0.113 (about 5%) |
| Decks with a designed graphic | Sonnet 5.5 | see `pptx_agent/tools/GUIDANCE.md` |

- Haiku's slips on Word were argument shapes (both alternatives of an either/or field, `doc`
  at the top level of `batch`), each recovered within one turn from the error's
  `valid_options`.
- Both models found deferred tools through tool search; use the default
  `toolbox.definitions("anthropic")`.

## House rules: in the application's prompt

House rules -- which styles to use for what, how captions read, what a footer says, whether
edits are tracked and under which author -- go in the application's own guidance, after the
shipped fragments, written against what the tools can see and do:

```python
HOUSE_RULES = '''House rules (Acme):
- Edit with tracking on, as "Acme Legal" (word_set_tracking), unless asked otherwise.
- Headings use the Heading 1-3 styles; body text uses Normal; no direct formatting
  (check with layout and styles lists what differs; word_format clear removes it).
- Every figure and table has a caption (word_fields insert_caption) above tables, below figures.
- Before saving, update fields and check with validate, reflow and fields.'''
system = toolbox.system_prompt(extra=HOUSE_RULES)
```

- Keep rules checkable against a tool's facts, short, and the same in the reviewer's brief.
- Do not put them in tool descriptions or patch them into the library: the tools stay the
  same for every application, and a rule belongs to one.
