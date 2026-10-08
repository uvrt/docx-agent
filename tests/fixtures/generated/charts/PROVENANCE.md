# The chart and SmartArt fixtures

| File | SHA-256 |
| --- | --- |
| `charts.docx` | `c56ad5c42a29a51c0c5e1224fa3ea6ef3400d03a4d9f6563c37cfe9fa292e027` |
| `chart-places.docx` | `b668fb1bf0580eee70876c70b9af15a7314f3e6e021dd73a433cec299b6c45a7` |
| `smartart.docx` | `070846f4ed4d43e80dad34c661b0046d60c718866e7a1337e7c64a6cac73f093` |

**Written here, then saved by Word** -- pptx-agent's SmartArt-fixture method.
`tools/make_chart_fixtures.py` builds each input from `Document.new()`'s parts, with the
charts, their embedded workbooks and the SmartArt data models written by
`tools/chart_parts.py`, every byte of them by that script; Microsoft Word 16.106 for Mac
(Dutch interface) then opened each and saved it again, through `tests/oracle.py`'s
machine-wide lock (no clipboard, no other document open). What is committed is what Word
wrote, with one change: `docProps/core.xml`'s author and dates replaced by `docx-agent` and
`2026-10-04T12:00:00Z` (`normalise_properties`, which writes every other entry's bytes as
Word wrote them). Running the script again gives the same inputs; Word's output differs in
its random parts (rsids, the VML fallback's `o:gfxdata`), so the files are committed rather
than regenerated. Licence: this repository's, MIT.

The SmartArt layout, quick-style and colour definitions in the *input* are those
pptx-agent's committed `tests/fixtures/powerpoint-smartart.pptx` carries (PowerPoint wrote
them there): a definition must be whole for Word to lay out anything but Basic Block List
(docx2svg's ROADMAP, F.21). The data models had no presentation points and no drawing;
Word laid both diagrams out and wrote the presentation points, the cached drawings
(`word/diagrams/drawing*.xml`) and the definitions back.

What Word wrote, beside the parts given: `c:lang` as its interface's `nl-NL`; numbers in
the caches as doubles (`4.4` became `4.4000000000000004`); an empty `c:dLbls` in each plot
and `c:showDLblsOverMax`; the embedded workbooks renamed `Microsoft_Excel-werkblad*.xlsx`
(its interface's word) and **kept byte for byte**; `w:noProof` on each drawing's run and
rsids; a text box's and a group's VML fallback -- the group's with a PNG of the chart
(`word/media/image1.png`), which Word does not redraw when the chart changes (measured,
`tools/charts_probe.py`, `places`); and no paraIds (Word 16.106 writes none for a document
that has none).

`charts.docx` (mode 15) holds six inline charts, each with its workbook (categories down
column A, a series to a column, a table over them): clustered columns with a title, a line,
a pie with a title, a doughnut of two rings, a scatter (numeric x values), and columns with
a line over the same categories. `chart-places.docx` holds a chart in each place a Word
chart stands outside the body's line: the default header (`d:1`), a footnote (`d:2`), a
floating text box (`d:4`, in `d:3`), a group beside a rectangle (`d:5/7`, a
`wpg:graphicFrame`) and a floating chart (`d:8`). `smartart.docx` holds Basic Block List
(`d:1`, five nodes) and Vertical Bullet List (`d:2`, two items with two bulleted items
each).

They sit one directory below the corpus (`tests/conftest.py` globs `*/*.docx`), so the
other phases' suites and their Word oracle keep the corpus they were measured on; the
readers (`conftest.reading_paths`), the chart and SmartArt suites and the chart oracle hold
them to every gate.
