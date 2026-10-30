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
| layout | `check` with `layout` (or `min_size`, `styles`) | what the pages show, measured from the layout reflow uses: the smallest text per kind and, with `min_size` (or the application's `min_text_size`, per kind), text below it; direct formatting over the style, with the `word_format` `clear` call that takes it off; with `styles`, paragraph and table styles not in the list; empty paragraphs used as space (with the `word_delete` and `space_before` calls that replace them); `empty_paragraphs`, per cell and for the body (`at`), the empty ones between two paragraphs of text (`paragraphs` their positions, `ids`, `fix` the `word_delete` calls; a body run is also among the spacers); keep-with-next paragraphs left at a page's foot; rows split across pages and header rows that do not repeat; pictures and tables past the margins; and `uncertain`, the pages with placeholders or approximations. Each list at most 20, most telling first, `<fact>_left_out` the rest; `pages` keeps the facts on those pages. Facts, not verdicts: which of them matter is the application's rule. |
| fields | `check` with `fields` | the fields (TOC, page numbers, cross-references, captions) and their cached results: a cache is only as fresh as the last `word_fields update` |
| `warnings` | every call | what the library noticed and did (an unknown construct kept, a style imported, a legacy option removed) |
| `renamed` | changing calls | old paragraph id -> new one; the old id keeps working |
| markup | `word_read` with `view: "markup"` | the tracked changes and comments as Word shows them in review |
| tracked changes | `word_changes` `list` | each revision: author, date, kind, old and new text |
| comments | `word_comments` `list` | each thread: anchor text, author, replies, resolved or open |
| formatting | `word_inspect` | a block's style, its runs' direct formatting, the effective font, size, spacing and indents |
| design facts | `word_design_facts` | for analysing a template: each style in use per role (body, headings by level, captions, lists, table text, headers and footers) with its usage and effective look; each table style's header row, bands, first column, borders and text as drawn; the page grid and header and footer variants per section; the theme's palette and fonts; block shapes repeated under headings, with example spans to copy |
| copies against their sample | `word_compare_blocks` | tables or paragraphs copied from a sample, compared with it (`reference`, here or in another open document): header rows, body rows, columns, merges, column widths (within 0.5 pt), table style, alignment, shading and borders, each cell's shading, borders and vertical alignment, each paragraph's style and direct formatting; every difference with the reference's value and, where a tool writes it, the calls that give it. Rows by the reference's structure: its header rows one to one, its last body row the pattern for any more |
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
- **`check`'s `layout.settings`** holds what the document is set to do (the tables kept
  together), not facts to act on: a summary counts none of it. A table kept together that
  is taller than a page is said there and, once, in `keep_with_next_at_foot`
  (`taller_than_page`), which is worth a look.
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
