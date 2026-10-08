# The end-to-end trial's inputs

| File | SHA-256 |
| --- | --- |
| `field-handbook.docx` | `9f646be95854bc8c27082aad88c08a327985168d1b3046b9fcf9c11828aaec5d` |
| `safety-bulletin.docx` | `1b5f67aa4967aa311c268f15140ec8cc8c9dd3b28ab1ac8d1498a4076ae423ba` |
| `template-saved-as-docx.docx` | `cbd62499c472168ffe9b8ebe840beee80f8bf4e2be00fca704206e1af25d5282` |

**Written by docx-agent** through its own API, for the full end-to-end trial (ROADMAP.md,
"Trial findings"), and committed as the trial used them:

- `field-handbook.docx` and `safety-bulletin.docx` are the merge-a-section task's two inputs:
  each a new document from `Document.new`, its styles changed and added with `styles.modify`
  and `styles.add` ("Handbook Body"; "Bulletin Body", "Bulletin Note"), its text written with
  `insert_markdown` (headings, paragraphs, lists, footnotes) and, in the bulletin, a picture
  inserted with `insert_picture`: a ladder drawn with Pillow by the trial's input builder.
  The handbook and the bulletin, their organisations and their text are invented.
- `template-saved-as-docx.docx` is the trial's reproduction of its template finding:
  a `.dotx` docx-agent wrote (`save_as_template`) opened with `Document.open` and saved
  as a `.docx` before `save` set the kind from the extension -- its main part still a
  template's, which Word refuses. It is the regression test's input.

Nothing in them comes from Word or any other application. Licence: this repository's, MIT.

They sit one level below `generated/`, so the corpus the other suites hold to every fixture
(`tests/conftest.py`, `fixture_paths`) does not take them in.
