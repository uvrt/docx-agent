# The end-to-end pilot's input

| File | SHA-256 |
| --- | --- |
| `agreement-summary.docx` | `0eaf103091e0dfcc2a4ee953a9de351d8c1e6d98d76b2fdceb09a149ef78c676` |

**Written by docx-agent** through its own API (a new document from `Document.new`, its
headings and paragraphs, and five comments by four authors), for the first end-to-end usability pilot
(ROADMAP.md, "Usability (end-to-end pilot)"): a short, invented supplier agreement summary
with five reviewers' comments, each asking for one change. Nothing in it comes from Word or
any other application; the names, amounts and the address `procurement@example.com` are
made up. It is committed as the pilot used it, so `examples/review_comments.py` runs on the
same bytes. Licence: this repository's, MIT.

It sits one level below `generated/`, so the corpus the other suites hold to every fixture
(`tests/conftest.py`, `fixture_paths`) does not take it in.
