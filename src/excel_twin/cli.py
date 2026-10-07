"""
    excel-twin check BOOK.xlsx [--formulas] [--sheet NAME] [--item-col B] [--summary-cell B3] [--yes Yes]
    excel-twin recalc BOOK.xlsx OUT_DIR
    excel-twin same-as-git FILE [FILE ...] [--ref HEAD]
"""
import argparse
import sys
from pathlib import Path

from . import determinism, evaluate, reconcile


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="excel-twin", description="Prove a workbook's formulas reproduce a Python model.")
    sub = ap.add_subparsers(dest="command", required=True)
    c = sub.add_parser("check", help="check the Reconciliation sheet and look for formula errors")
    c.add_argument("workbook", type=Path)
    c.add_argument("--formulas", action="store_true", help="calculate in pure Python instead of reading saved values")
    for name, default in (("sheet", "Reconciliation"), ("item-col", "B"), ("summary-cell", "B3"), ("yes", "Yes")):
        c.add_argument(f"--{name}", default=default)
    c.add_argument("--first-row", type=int, default=7)
    r = sub.add_parser("recalc", help="recalculate a copy with LibreOffice")
    r.add_argument("workbook", type=Path)
    r.add_argument("out_dir", type=Path)
    g = sub.add_parser("same-as-git", help="compare files with their committed version, ignoring timestamps")
    g.add_argument("files", nargs="+")
    g.add_argument("--ref", default="HEAD")
    a = ap.parse_args(argv)

    if a.command == "check":
        layout = reconcile.Layout(sheet=a.sheet, item_col=a.item_col, summary_cell=a.summary_cell, yes=a.yes,
                                  first_row=a.first_row)
        values = evaluate.values_from_formulas(a.workbook) if a.formulas else evaluate.values_from_saved(a.workbook)
        report = reconcile.check(values, layout=layout)
        report.print()
        return 0 if report.ok else 1
    if a.command == "recalc":
        from . import recalc
        print(f"recalculated -> {recalc.recalc(a.workbook, a.out_dir)}")
        return 0
    return 1 if determinism.check_against_git(a.files, a.ref) else 0


if __name__ == "__main__":
    sys.exit(main())
