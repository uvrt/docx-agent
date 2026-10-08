# E6's templates

| File | SHA-256 |
| --- | --- |
| `brand.dotx` | `8687bec707ad0da4c083c5c8a46920a9a7224dc9de0df8a651ec26a52c516123` |
| `brand-macros.dotm` | `305007f658887b54f8447ee8cbae3d479bb55018c75602087d89938981d59696` |
| `brand-mode14.dotx` | `4c22c2081ba576c9486baeff014ccbde6cee005c2373ab426a3458ba78a5848a` |

Written by `tools/make_template_fixtures.py` in this repository, deterministically, every
part by hand with the helpers of `tools/e6_probe.py` (whose `template` Word made a document
from in the probe): nothing in them comes from Word, Word's templates or any other
application. The VBA project in `brand-macros.dotm` is a stand-in of a few bytes, not a
VBA project. Licence: this repository's, MIT.

`brand.dotx` (mode 15): heading 1 redefined, a paragraph and a character style of its own,
header and footer styles, a bullet list, a header and a footer, Letter paper with a
1000-twip top margin, even and odd headers, four body paragraphs, and properties (title,
subject, author, keywords, company). `brand-macros.dotm` is it as a macro-enabled template;
`brand-mode14.dotx` is it in compatibility mode 14. They sit two levels below the corpus, so
no per-fixture suite opens them as documents.
