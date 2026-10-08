# CommonMark spec examples

`examples.json` is a selection of the examples of the **CommonMark Spec, version 0.31.2**,
by John MacFarlane (Copyright (C) 2014-24 John MacFarlane), fetched from
<https://spec.commonmark.org/0.31.2/spec.json>.

**Licence:** the spec, and so this selection, is licensed under the Creative Commons
Attribution-ShareAlike 4.0 International licence (CC BY-SA 4.0,
<https://creativecommons.org/licenses/by-sa/4.0/>). This directory is under that licence,
not under the repository's MIT licence.

**Changes:** only each example's Markdown, number and section are kept (not its HTML); the
examples of the raw-HTML constructs (the *HTML blocks* and *Raw HTML* sections, and any
other example holding raw HTML that is not a comment) are left out, because
`insert_markdown` refuses raw HTML by design, and so would any example matching the
repository's trace check (none does). `examples.json` records the counts. Written by
`tools/select_commonmark_examples.py`, which fetches the spec and makes the selection again.

The examples are the corpus of `tests/test_markdown_write.py`'s round trip: `insert_markdown`
of each, then `to_markdown`, must give the same CommonMark AST after the round trip's
normalisation, in a blank document and in every template fixture -- except the examples that
test lists as outside the guarantee, each with its reason.
