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

## Golden transcripts

`tests/goldens/transcripts` are tasks done with the tools alone, each call recorded with its
normalised result (`tests/goldens_replay.py`). Every call's `expect` holds `ok`, `summary`,
the ids and warnings, and three things about its data: `data_sha` (all of it), `data_keys`
(its schema: its keys and, for a dict-valued one such as `coverage` or `validate`, that
dict's keys) and `data_sha_core` (all of it less what the layout measures).

`tests/test_tools_goldens.py` replays them two ways:

- **Where Office's faces are installed** (Calibri, Cambria, Georgia, Aptos: the Mac they were
  recorded on), every result must be as recorded and every saved document byte for byte.
- **Everywhere, CI's runners included**, the *core* replay: the same calls with whatever
  faces are there (the open substitutes, or none), comparing everything except what the
  layout measures. A result that gains, loses or renames a key, or whose other content
  changes, fails on every runner.

What the layout measures is defined once, at the top of the split in
`tests/goldens_replay.py`: the data fields `reflow`, `coverage`, `pages`, `pages_estimated`
and a saved document's `size` (its `app.xml` carries the page count); the keys those fields
have only when the layout stopped or substituted a face (`coverage.stop`,
`coverage.missing_fonts`, `reflow.stopped`, ...), left out of `data_keys`; and the calls that
write measured page numbers into the document (`word_fields` `update`, `insert_toc`, ...),
after which a session's results are compared by `ok` and `data_keys` only, and the trial's
check (which reads those page numbers) runs only where the faces are. If a new transcript
fails the core replay on CI but passes on the Mac, a result depends on the layout in a way
the split does not name: add it there, with why.

**Refreshing them.** When a change adds a key to results (as `coverage` and `coverage.status`
did), re-record on a Mac with Office:

```bash
python tools/refresh_goldens.py --dry-run --new-keys coverage.status   # what would change
python tools/refresh_goldens.py --new-keys coverage.status             # write it
```

It writes nothing unless, for every call, `ok`, `summary` and the other recorded fields are
unchanged, the data less the keys named by `--new-keys` hashes to the recorded `data_sha`,
and every saved document is byte for byte the one recorded. A change that alters a result
or a saved document is not a refresh: record that transcript again
(`goldens_replay.run(..., record=True)`) and review the diff as part of the change.

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
Python 3.10 to 3.15, with the siblings installed from their `main` branches.
