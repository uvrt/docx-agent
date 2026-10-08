# E2's fixtures

| File | SHA-256 |
| --- | --- |
| `constructs.docx` | `970a8dfc6bd1c682b174c1ac8440c462ba137756769d93f6c408c25aedd962a1` |
| `review.docx` | `cfc65d864152dd3e319b5b311a9e7f19a33a318c05e3158b29ed58fb61d38943` |
| `dutch-template.docx` | `aea529307424031e2769497811952565dc96fecd2eccbd9221813f92ab0fafeb` |
| `blank.docx` | `bc77a32eed81f2a555d312fe4ed677892f75453e2b1a1bb9becdb20139a7114f` |

Written by `tools/make_markdown_fixtures.py` in this repository, deterministically (running
it again writes the same bytes), every part by hand with the helpers of
`tools/make_fixtures.py`: the styles, numbering, notes and comments are written there, not
taken from Word, a template or any other application; the two pictures are one-colour
4x4 PNGs written by the same script. Licence: this repository's, MIT.

They are what `to_markdown` and the JSON state are tested on (`tests/test_markdown*.py`,
`tests/test_state.py`), beside the corpus. They sit one directory below it, so E0's and
E1's suites and the Word oracle keep the corpus they were measured on; none of the three
has been opened in Word.

`constructs.docx` (mode 15) holds every construct the Markdown layer maps: headings 1-3
and 7, a title, direct and styled emphasis, strong, strikethrough and inline code, text
full of Markdown's special characters, a quote, a code block, a thematic break, nested,
adjacent and restarted lists, links, a bookmark, footnotes and an endnote, inline and
floating pictures, a text box, fields, a GFM-shaped table and one GFM cannot hold, content
controls, a section break, a page break, a header and a footer. `review.docx` (mode 15)
holds revisions by two authors -- insertions, deletions, a move, deleted and inserted
paragraph marks, table rows, a formatting and a style change -- and comments, one a reply
and one resolved. `dutch-template.docx` (mode 15) is a style set localised to Dutch, with
headings named "Kop 1" and "Kop 2" that only their outline level makes headings.

`blank.docx` (mode 15) was the blank document `insert_markdown` was held to until E6
measured the one Word makes (`Document.new()`, which `tests/test_markdown_write.py` now
uses): the document defaults, the four styles every document has (Normal, Default
Paragraph Font, Normal Table, No List), one empty paragraph and an A4 section; no
numbering, notes or theme. It stays with the reading fixtures -- a minimal document every
reader is held to.
