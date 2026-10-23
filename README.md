<picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/D0M3N1C0X/excel-twin/main/site-assets/cover-dark.png">
  <img alt="excel-twin: prove the spreadsheet matches the model. Five analysis repositories reconcile through it, 3,134 figures checked between Python and Excel." src="https://raw.githubusercontent.com/D0M3N1C0X/excel-twin/main/site-assets/cover-light.png">
</picture>

<p align="center">
  <a href="#use"><b>Use it</b></a> ·
  <a href="#the-idea">How it works</a> ·
  <a href="#used-in">Where it runs</a> ·
  <a href="CHANGELOG.md">Changelog</a>
</p>

<p align="center">
  <a href="https://github.com/D0M3N1C0X/excel-twin/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/D0M3N1C0X/excel-twin/actions/workflows/ci.yml/badge.svg"></a>
  <img alt="Python 3.12+" src="https://img.shields.io/badge/python-3.12%2B-0071e3">
  <img alt="Typed" src="https://img.shields.io/badge/typing-strict%20mypy-5e5ce6">
  <img alt="MIT licence" src="https://img.shields.io/badge/licence-MIT-6e6e73">
</p>

## Why this matters

Finance and HR teams receive models as Excel workbooks, because that is the file they can open, audit and change.
When the analysis behind it was built in pandas, the two engines can drift apart: a formula typed once and never
checked, a rounding rule that differs, a range that stops one row short. Nobody notices until a number in a board
pack is wrong. **excel-twin makes the workbook prove it agrees with the model**, figure by figure, on every change.
It is the small toolkit behind the reconciliation in five of my analysis repositories (3,134 figures checked in all),
extracted so each one tests the same way.

## The idea

The workbook carries a **Reconciliation** sheet: one row per figure, with the value Python computed next to a
live formula that points at the cell which should equal it, and a match column.

| Area | Item | | Python value | Workbook value | Difference | Match |
|---|---|---|---:|---:|---:|---|
| Model | Revenue | | 34.25 | `=Model!B3` | 0 | Yes |

A spreadsheet engine calculates the file, and excel-twin fails if any match cell is not *Yes* or any cell
anywhere holds a formula error (`#REF!`, `#DIV/0!`, …).

## What's in it

| Module | Does |
|---|---|
| `reconcile` | The sheet's layout (`Layout`, configurable: sheet name, columns, the word for a match) and `check`, which returns a `Report` |
| `evaluate` | Reads values a spreadsheet saved, or calculates every formula in pure Python with [`formulas`](https://pypi.org/project/formulas/), with Excel's `ROUND` |
| `recalc` | Recalculates a copy with LibreOffice through a one-line Basic macro in a throwaway profile, for CI |
| `determinism` | Byte-stable `.xlsx`: fixed zip timestamps and document dates, and a content comparison that ignores compression |

## One bug it corrects

`formulas` rounds half-cent ties down: `ROUND(34.245, 2)` gives 34.24, because 34.245 × 100 is
3424.4999… in binary. Excel and LibreOffice give 34.25: they cut to 15 significant digits first.
excel-twin replaces `ROUND` with the Excel behaviour, so a Python check agrees with the spreadsheet a client
will open. Found in a payroll model where it moved three totals by one cent.

## Use

```bash
pip install "excel-twin[formulas] @ git+https://github.com/D0M3N1C0X/excel-twin@v0.1.2"
```

```python
import excel_twin as et

report = et.check(et.values_from_formulas("model.xlsx"))          # pure Python
report = et.check(et.values_from_saved("recalculated.xlsx"),      # after LibreOffice or Excel
                  layout=et.Layout(sheet="Riconciliazione", yes="Sì"))
report.print()
assert report.ok
```

```bash
excel-twin recalc deliverables/model.xlsx build/        # LibreOffice
excel-twin check build/model.xlsx                        # saved values
excel-twin check deliverables/model.xlsx --formulas      # pure Python
excel-twin same-as-git deliverables/model.xlsx           # rebuilt file has the committed content
```

## Used in

[pay-transparency-readiness-kit](https://github.com/D0M3N1C0X/pay-transparency-readiness-kit) ·
[workforce-cost-model](https://github.com/D0M3N1C0X/workforce-cost-model) ·
[equal-value-pay-ranges](https://github.com/D0M3N1C0X/equal-value-pay-ranges) ·
[controlli-cedolini](https://github.com/D0M3N1C0X/controlli-cedolini) ·
[vimeo-growth-case](https://github.com/D0M3N1C0X/vimeo-growth-case)

## Development

```bash
pip install -e ".[test,lint]"
pytest --cov       # coverage is also reported in each CI run
ruff check .       # style and likely bugs
mypy               # strict types; the package ships a py.typed marker
```

## Limits

- `formulas` is slow on large workbooks: use it on samples locally and LibreOffice on the full file in CI.
- Dynamic-array functions are not supported by `formulas`; the repositories above avoid them on purpose.

## About

Built by **Domenico Perroni**: HR advisory for the Italian market, people analytics and reward, based in Kraków,
with an AI coding assistant; the design, the tests and the bug hunt are mine.
[GitHub](https://github.com/D0M3N1C0X) · [LinkedIn](https://www.linkedin.com/in/domenico-perroni) · [Portfolio](https://d0m3n1c0x.github.io/)

MIT licensed.
