# Contributing

## Running the tests

```bash
pip install "ooxml-common @ git+https://github.com/uvrt/ooxml-common@main" \
            "ooxml-edit @ git+https://github.com/uvrt/ooxml-edit@main" \
            "docx2svg[png] @ git+https://github.com/uvrt/docx2svg@main"
pip install -e .[dev,png]
```

or, with sibling checkouts next to this one (which some local-only tests need):

```bash
python3 -m venv --system-site-packages .venv
.venv/bin/pip install -e ../ooxml-edit -e ../ooxml-common -e ../docx2svg -e '.[dev]'
```

```bash
python -m pytest -q -n auto               # the gates: round trip, ids, undo, validity, render
python -m pytest -q -n auto --run-slow    # and the exhaustive sweeps marked slow
python -m pytest -m oracle -q             # Microsoft Word itself (macOS), skipped if Word is in use
```

The default run holds every gate on every fixture, the heaviest per-fixture suites (each
phase's operations, renders and tracked invariants, the Markdown corpus, the copy and
upgrade sweeps) on four representative fixtures and the randomised table sweep on its
first seeds -- about three and a half minutes on four cores; `--run-slow` holds those on every
fixture and seed (about fourteen minutes). Everything is `pytest -q -n auto --run-slow`
and then `pytest -m oracle -q` (the oracle never runs under `-n`).

**Local-only tests skip cleanly elsewhere, including on CI:**

- The **Word oracle** (`-m oracle`, opt-in) runs docx2svg's AppleScripts (found in the
  sibling checkout, or through `DOCX2SVG_ORACLE_SCRIPT`) under one machine-wide lock, and
  never shares or kills a Word someone else is using.
- Tests that read docx2svg's probe observations need its checkout next to this one (or
  `DOCX2SVG_REPO`).
- **Provider tests** (`-m provider`) call a model provider's API; skipped without
  `ANTHROPIC_API_KEY`.

`tools/paraid_probe.py` repeats the paraId measurement, and `tools/e1_probe.py` to
`tools/e6_probe.py` the measurements of what Word writes for each phase's edits,
`tools/charts_probe.py` those of what Word does with charts and SmartArt.

The recipes in [docs/common-tasks.md](docs/common-tasks.md) are run as written by
`tests/test_readme.py`; keep them working.

## What never goes into the repository

- **No Microsoft font file**, in any form, and no font data in any fixture.
- **No Office output**: no PDF or raster Word exported.
- **No third-party document** without a licence that allows redistribution, kept in its
  own directory with a `PROVENANCE.md` (source, date, licence, hashes).
- No API keys or other secrets.

## Pull requests

Open pull requests against `main`. CI runs the default suite on Linux, macOS and Windows,
Python 3.10 to 3.13, with the siblings installed from their `main` branches.
