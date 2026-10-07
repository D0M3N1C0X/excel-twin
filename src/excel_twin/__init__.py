"""
excel-twin: prove that a spreadsheet's live formulas reproduce a Python model.

A model built in Python and delivered as an Excel workbook has two engines that can drift apart.
excel-twin checks that they agree: the workbook carries a Reconciliation sheet (one row per figure,
the Python value next to a live formula), a spreadsheet engine calculates it, and every row must
match while no cell holds a formula error.

    determinism   byte-stable .xlsx files: fixed timestamps, content comparison
    recalc        recalculate a copy with LibreOffice, so formulas carry values
    evaluate      read calculated values, or calculate in pure Python with Excel's rounding
    reconcile     the layout of the Reconciliation sheet and the check itself
"""
from .determinism import differences, normalise, same_content
from .evaluate import excel_round, values_from_formulas, values_from_saved
from .reconcile import ERRORS, Layout, Report, check, count_checks

__all__ = ["ERRORS", "Layout", "Report", "check", "count_checks", "differences", "excel_round", "normalise",
           "same_content", "values_from_formulas", "values_from_saved"]
__version__ = "0.1.0"
