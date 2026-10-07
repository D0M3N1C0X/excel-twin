"""
The Reconciliation sheet and the check that reads it.

The sheet has one row per figure from ``first_row``: an item label, the value Python computed, a live
formula pointing at the cell that should equal it, and a match column that says whether they agree.
``check`` fails when any match cell is not the agreed word, or when any cell anywhere in the workbook
holds a formula error.
"""
from dataclasses import dataclass, field

ERRORS = ("#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "Err:")


@dataclass(frozen=True)
class Layout:
    sheet: str = "Reconciliation"
    first_row: int = 7
    area_col: str = "A"
    item_col: str = "B"
    python_col: str = "D"
    workbook_col: str = "E"
    match_col: str = "G"
    summary_cell: str = "B3"
    yes: str = "Yes"


@dataclass
class Report:
    checks: int
    formula_errors: list = field(default_factory=list)
    mismatches: list = field(default_factory=list)
    summary: object = None

    @property
    def ok(self) -> bool:
        return self.checks > 0 and not self.formula_errors and not self.mismatches

    def print(self) -> None:
        for s, r, v in self.formula_errors[:20]:
            print(f"::error::formula error {v} at {s}!{r}")
        for m in self.mismatches[:30]:
            print("::error::mismatch", m)
        print(f"formula errors: {len(self.formula_errors)}")
        print(f"reconciliation: {self.checks - len(self.mismatches)} of {self.checks} match")
        print(f"summary cell: {self.summary}")


def count_checks(values: dict, layout: Layout = Layout()) -> int:
    """Rows of the Reconciliation sheet that carry an item label."""
    s = layout.sheet.upper()
    col = layout.item_col
    return sum(1 for (sheet, ref) in values if sheet == s and ref.startswith(col) and ref[len(col):].isdigit()
               and int(ref[len(col):]) >= layout.first_row)


def check(values: dict, n_checks: int | None = None, layout: Layout = Layout()) -> Report:
    """Every formula error in the workbook, and every reconciliation row that does not match."""
    s = layout.sheet.upper()
    get = lambda ref: values.get((s, ref))
    n = count_checks(values, layout) if n_checks is None else n_checks
    errors = [(sh, r, v) for (sh, r), v in values.items() if not isinstance(v, (int, float)) and str(v).startswith(ERRORS)]
    mismatches = []
    for row in range(layout.first_row, layout.first_row + n):
        if get(f"{layout.match_col}{row}") != layout.yes:
            mismatches.append(tuple(get(f"{c}{row}") for c in (layout.area_col, layout.item_col, layout.python_col,
                                                                  layout.workbook_col)))
    return Report(checks=n, formula_errors=errors, mismatches=mismatches, summary=get(layout.summary_cell))
