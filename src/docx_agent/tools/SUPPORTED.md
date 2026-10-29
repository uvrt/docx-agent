# What the Word tools support (overview)

The tools in `docx_agent.tools`, by area. Lengths are points. Every changing call is one undo
step, written as tracked changes while tracking is on, and returns the validation delta and,
for laid-out documents, the pages it reflowed (see [GUIDANCE.md](GUIDANCE.md)).

| Area | Tools | Notes |
|---|---|---|
| Documents | `open_document`, `new_document`, `save_document`, `close_document`, `read_blob` | new from a template; save as docx, dotx or Markdown; refuses new validation problems |
| Reading | `describe`, `word_read`, `word_inspect`, `find_text`, `render` | Markdown with ids (final, markup or original view, every story; a block, a span or a section, `s:2` or `s:body`); a block's exact formatting; page renders (with `sheet`, up to 30 pages tiled into one labelled image) |
| Text | `word_set_text`, `word_insert_text`, `word_delete`, `replace_text`, `word_insert_markdown` | within a block or a span `p:A..t:B` (`range`); a figure from a dataset or another document's table by reference, `{{d:<id>/<key>/<column>}}` (`:0.0` to format, `:d mmmm yyyy` for a date), written in the session's locale and listed in the result (`data.facts`), not kept in the file; tracked edits change only the words that changed; Markdown in the document's styles, tables included (a cell style whose indents run past the cell is named in a warning; a centred or right column takes the `table_cell_center` / `table_cell_right` style, else direct alignment and a warning naming the role; a list in a style that numbers itself takes the style's numbering and indent and counts on as Word counts it); a table from a dataset or another document's table (`from`; its dates in a pattern with `date_format`) |
| Review | `word_set_tracking`, `word_changes`, `word_comments` | tracked changes listed, accepted, rejected by author, kind or range; comment threads |
| Structure | `word_move`, `word_copy_from`, `word_sections`, `word_headers_footers` | sections and blocks (a block-level content control with its content too, as `word_delete` deletes one) moved or copied from another document with their styles; page setup, breaks, columns; headers, footers, page X of Y |
| Fields and references | `word_fields`, `word_notes`, `word_links` | TOC, captions, cross-references, dates, update; footnotes and endnotes; hyperlinks and bookmarks |
| Formatting and styles | `word_format`, `word_lists`, `word_styles`, `word_template` | run and paragraph formatting (`top`/`bottom` read as the paragraph's spacing), clearing direct formatting, all of it or the properties named; lists (restart and continue as Word writes them, for items their style numbers too); the style sheet; out of compatibility mode |
| Objects | `word_edit_table`, `word_format_table`, `word_drawings`, `word_controls`, `edit_chart`, `edit_smartart` | tables edited and formatted (new ones are Markdown), cells set from a dataset or another document's table (`set_cells` `from`); pictures, text boxes, shapes, inline or floating; content controls; charts from data (column, stacked column, bar, stacked bar, line, pie, scatter, radar), with their scale, gridlines, legend, data labels, series colours and text style in one call; SmartArt text |
| Properties | `set_properties` | title, author, language, subject |
| Session | `undo`, `check`, `batch` | `check`: validation, reflow, fields and, with `layout`, the pages' layout facts (small text, direct formatting with the call that clears it, styles outside a list, empty paragraphs as space, empty paragraphs between text in the body and in cells, keep-with-next at a page's foot, split rows, header rows not repeated, past the margins, uncertain pages); `batch`: a failing op is skipped with its error, the rest apply as one step; `atomic: true` for all or none |

A table's width is a share of the text width or, with `word_format_table` `width_of: "body"`, of
the body's lines beside it (a template whose Normal runs past a margin); `check` does not count
such a table as past the margins.

Not supported, in outline: macros (dropped on save), equations and ink beyond keeping them,
building SmartArt, editing a template's styles from another file (`word_template` only
upgrades the compatibility mode), and anything Word does on its own when it opens a file
(updating fields, repaginating) beyond what `word_fields update` and the layout facts give.
