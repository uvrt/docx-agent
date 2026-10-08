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
| `coverage` | `check` with `reflow`, `render`, `save_document` | how much of the document the layout covers. `complete` is false when a block, a header or footer, or a face could not be laid out. It also gives the blocks laid out, where the layout stopped, and the faces substituted or missing. |
| fields | `check` with `fields` | the fields (TOC, page numbers, cross-references, captions) and their cached results: a cache is only as fresh as the last `word_fields update` |
| `warnings` | every call | what the library noticed and did (an unknown construct kept, a style imported, a legacy option removed) |
| `renamed` | changing calls | old paragraph id -> new one; the old id keeps working |
| markup | `word_read` with `view: "markup"` | the tracked changes and comments as Word shows them in review |
| tracked changes | `word_changes` `list` | each revision: author, date, kind, old and new text |
| comments | `word_comments` `list` | each thread: anchor text, author, replies, resolved or open |
| formatting | `word_inspect` | a block's style, its runs' direct formatting, the effective font, size, spacing and indents |
| `validate` on save | `save_document` | the save refuses new validation problems; only the application can allow them |

How to use them:

- **A check is only as good as its coverage.** A reflow or a render of a partial layout
  says nothing about the pages past the stop. When `coverage.complete` is false, report
  what could not be checked; don't present it as checked.
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
  `clear_direct` takes it off.

## A review pass

The application calls the tools itself, through `toolbox.dispatch`, after the model says it is
done:

```python
def review(toolbox, session, doc, pages, house_rules, model_call):
    """Facts, the reviewer's view and renders of the changed pages, judged by the application's
    own rules in its own model call."""
    facts = toolbox.dispatch(session, "check", {"doc": doc, "include": ["validate", "reflow",
                                                                         "fields"]})
    markup = toolbox.dispatch(session, "word_read", {"doc": doc, "view": "markup"})
    images = toolbox.dispatch(session, "render", {"doc": doc, "pages": pages[:4]}).images
    return model_call(images=images, text=markup.data, facts=facts.data, rules=house_rules)
```

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
  (check with word_inspect; remove with word_format clear_direct).
- Every figure and table has a caption (word_fields insert_caption) above tables, below figures.
- Before saving, update fields and check with validate, reflow and fields.'''
system = toolbox.system_prompt(extra=HOUSE_RULES)
```

- Keep rules checkable against a tool's facts, short, and the same in the reviewer's brief.
- Do not put them in tool descriptions or patch them into the library: the tools stay the
  same for every application, and a rule belongs to one.
