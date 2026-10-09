import shutil
import time

import pytest

import excel_twin as et
from excel_twin import cli


def test_a_model_that_agrees_passes(make):
    path, n = make()
    report = et.check(et.values_from_formulas(path))
    assert report.ok and report.checks == n
    assert report.summary == "All 3 checks match"


def test_excel_rounding_of_half_cent_ties():
    assert et.excel_round(34.245, 2) == 34.25
    assert et.excel_round(22.83 * 1.5, 2) == 34.25
    assert et.excel_round(2.675, 2) == 2.68
    assert et.excel_round(-1.005, 2) == -1.01


def test_formulas_own_round_breaks_the_tie_and_ours_does_not(make):
    from formulas.functions import math as formulas_math
    assert formulas_math.xround(34.245, 2) == 34.24          # the bug excel-twin corrects
    path, _ = make()
    assert et.values_from_formulas(path)[("MODEL", "B3")] == 34.25


def test_a_wrong_python_value_is_a_mismatch(make):
    path, _ = make(python_values={"Revenue": 34.26})
    report = et.check(et.values_from_formulas(path))
    assert not report.ok and len(report.mismatches) == 1 and report.mismatches[0][1] == "Revenue"


def test_a_formula_error_anywhere_fails(make):
    path, _ = make(break_formula=True, python_values={"Ratio": 0.0})
    report = et.check(et.values_from_formulas(path))
    assert report.formula_errors and not report.ok


def test_an_empty_reconciliation_is_not_a_pass():
    assert not et.check({}, 0).ok


def test_a_blank_row_inside_the_sheet_is_skipped_unless_rows_are_counted():
    s = "RECONCILIATION"
    values = {(s, "B7"): "first", (s, "G7"): "Yes", (s, "B9"): "second", (s, "G9"): "Yes"}
    report = et.check(values)
    assert report.ok and report.checks == 2
    # told to expect three consecutive rows, the build is missing one: that must fail
    assert not et.check(values, 3).ok


def test_layout_is_configurable(make):
    path, _ = make()
    values = et.values_from_formulas(path)
    renamed = {("RICONCILIAZIONE" if s == "RECONCILIATION" else s, r): v for (s, r), v in values.items()}
    assert et.check(renamed, layout=et.Layout(sheet="Riconciliazione")).ok


def test_builds_are_byte_stable_in_content(make):
    a, _ = make("a.xlsx")
    time.sleep(1.1)
    b, _ = make("b.xlsx")
    assert et.same_content(a.read_bytes(), b.read_bytes())
    assert et.differences(a.read_bytes(), b.read_bytes()) == []


def test_differences_name_the_entry(make):
    a, _ = make("a.xlsx")
    b, _ = make("b.xlsx", python_values={"Revenue": 1.0})
    assert any(line.startswith("xl/worksheets/") for line in et.differences(a.read_bytes(), b.read_bytes()))


def test_cli_check(make, capsys):
    path, _ = make()
    assert cli.main(["check", str(path), "--formulas"]) == 0
    bad, _ = make("bad.xlsx", python_values={"Label": "low"})
    assert cli.main(["check", str(bad), "--formulas"]) == 1
    assert "reconciliation: 2 of 3 match" in capsys.readouterr().out


@pytest.mark.skipif(not (shutil.which("soffice") or shutil.which("libreoffice")), reason="LibreOffice not installed")
def test_libreoffice_recalculates_and_agrees(make, tmp_path):
    from excel_twin.recalc import recalc
    path, n = make()
    out = recalc(path, tmp_path / "out")
    report = et.check(et.values_from_saved(out))
    assert report.ok and report.checks == n


def test_values_from_saved_reads_stored_cells_only(make):
    path, _ = make()
    values = et.values_from_saved(path)
    assert values[("MODEL", "B1")] == 22.83 and values[("RECONCILIATION", "B7")] == "Revenue"
    assert ("MODEL", "B3") not in values          # never calculated by a spreadsheet engine: no stored value


def test_check_against_git_flags_a_changed_rebuild(make, tmp_path, monkeypatch):
    import subprocess

    from conftest import build
    if shutil.which("git") is None:
        pytest.skip("git not installed")
    path, _ = make()
    monkeypatch.chdir(tmp_path)
    git = ["git", "-c", "user.name=t", "-c", "user.email=t@example.com"]
    subprocess.run([*git, "init", "-q"], check=True)
    subprocess.run([*git, "add", path.name], check=True)
    subprocess.run([*git, "commit", "-q", "-m", "workbook"], check=True)
    build(path)                                   # the same build again: same content
    assert et.determinism.check_against_git([path.name]) == 0
    build(path, python_values={"Revenue": 34.26})  # a different build
    assert et.determinism.check_against_git([path.name]) == 1

