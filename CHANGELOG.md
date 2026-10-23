# Changelog

## 0.1.2 — 2026-10-23

- `excel_round` returns the value unchanged when it has more digits than can be rounded (`excel_round(1e20, 10)`
  raised before).
- Packaging: setuptools 77 or later (the licence is an SPDX string, PEP 639); classifiers for Python 3.12-3.14,
  status and audience.
- The release workflow runs the tests, checks that a tag matches the version and runs `twine check` before
  publishing.

## 0.1.1 — 2026-10-09

- `check` without `n_checks` now checks exactly the labelled rows of the Reconciliation sheet, so a blank
  row inside the sheet is skipped instead of being reported as a mismatch (and the last row checked). With
  `n_checks`, rows are still consecutive from `first_row`, so a missing row fails. New: `labelled_rows`.
- Typed: strict mypy in CI, a `py.typed` marker for users' type checkers, and a `Values` type alias.
- CI also runs ruff and measures coverage.

## 0.1.0 — 2026-10-08

First release, extracted from five analysis repositories that each carried their own copy:
reconciliation check with a configurable layout, pure-Python evaluation with Excel's ROUND,
LibreOffice recalculation, byte-stable Office files, and a command line.
