"""
Values of a workbook's cells, keyed by (SHEET NAME IN UPPER CASE, "A1").

``values_from_saved`` reads what a spreadsheet engine stored when it last calculated the file.
``values_from_formulas`` calculates the formulas in pure Python with the ``formulas`` package, with one
correction: Excel's ROUND.
"""
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path


def excel_round(x, d):
    """Excel's ROUND: the value is first cut to 15 significant digits, then rounded half away from zero,
    so ROUND(34.245, 2) is 34.25. `formulas` multiplies the float by 100 and gets 3424.4999..., which
    gives 34.24 on every half-cent tie; Excel and LibreOffice agree with this version."""
    v = Decimal(f"{float(x):.15g}").quantize(Decimal(1).scaleb(-int(d)), rounding=ROUND_HALF_UP)
    return float(v)




def values_from_saved(path: Path) -> dict[tuple[str, str], object]:
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    out[(ws.title.upper(), cell.coordinate)] = cell.value
    return out


def values_from_formulas(path: Path) -> dict[tuple[str, str], object]:
    """Calculate every formula in pure Python. ROUND is replaced by ``excel_round`` for the whole process:
    `formulas` caches functions after their first use, so the replacement is made once, before any model."""
    try:
        import formulas
        from formulas.functions import FUNCTIONS, wrap_ufunc
    except ImportError as e:  # pragma: no cover
        raise RuntimeError("install the optional dependency: pip install 'excel-twin[formulas]'") from e
    FUNCTIONS["ROUND"] = wrap_ufunc(excel_round)
    solution = formulas.ExcelModel().loads(str(path)).finish().calculate()
    out = {}
    for key, ranges in solution.items():
        if "!" not in key:
            continue
        sheet, ref = key.rsplit("!", 1)
        sheet = sheet.strip("'").split("]", 1)[-1]
        if ":" in ref:
            continue
        value = ranges.value[0][0] if hasattr(ranges, "value") else ranges
        if hasattr(value, "item"):
            value = value.item()
        out[(sheet.upper(), ref.replace("$", ""))] = value
    return out
