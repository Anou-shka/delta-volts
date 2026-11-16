"""The Excel model builds, recalculates with zero errors, and every check is TRUE."""
import re
import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "model" / "delta_model.xlsx"

pytestmark = pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")


@pytest.fixture(scope="module")
def recalculated():
    from src.delta import build_excel, recalc
    assert build_excel.main() == 0
    result = recalc.recalc(MODEL)
    assert result.get("status") == "success", result
    return load_workbook(MODEL, data_only=True), load_workbook(MODEL)


def test_all_checks_true(recalculated):
    values, _ = recalculated
    cover = values["Cover"]
    flag = next(r[2].value for r in cover.iter_rows() if r[1].value == "ALL CHECKS PASS")
    assert flag is True
    checks = values["Checks"]
    for row in checks.iter_rows(min_row=8):
        for c in row[2:]:
            assert c.value is not False, f"Checks!{c.coordinate} is FALSE"


@pytest.mark.parametrize("sheet,label,col,value", [
    ("IS", "Revenue", "G", 554_885.168),
    ("IS", "Operating profit (EBIT)", "G", 83_932.074),
    ("IS", "EPS, basic (NT$)", "O", 17.59),
    ("BS", "Total assets", "G", 639_618.600),
    ("CF", "Cash at end of period", "O", 182_350.204),
    ("IS", "Revenue", "H", 539_668 + 218_000),      # FY2026E Base = 9M26 actual + 4Q26E
])
def test_key_values(recalculated, sheet, label, col, value):
    values, _ = recalculated
    ws = values[sheet]
    row = next(r[0].row for r in ws.iter_rows() if r[0].value == label)
    assert ws[f"{col}{row}"].value == pytest.approx(value, abs=0.01)


def test_statements_have_no_hardcoded_numbers(recalculated):
    """Inputs live on Historicals/Segments/Assumptions; statements and schedules are formulas only."""
    _, formulas = recalculated
    for sheet in ("IS", "BS", "CF", "Schedules"):
        ws = formulas[sheet]
        for row in ws.iter_rows(min_row=7, min_col=3, max_col=17):
            for c in row:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    pytest.fail(f"hard-coded number at {sheet}!{c.coordinate}: {c.value}")
                if isinstance(c.value, str) and c.value.startswith("="):
                    assert not re.search(r"(?<![A-Z$\d.])\d{3,}", c.value), f"constant in {sheet}!{c.coordinate}"


@pytest.mark.parametrize("scenario", ["Bear", "Bull"])
def test_every_scenario_balances(recalculated, scenario, tmp_path):
    """Switching the scenario selector keeps every check TRUE."""
    from src.delta import recalc
    path = tmp_path / f"model_{scenario}.xlsx"
    shutil.copy(MODEL, path)
    wb = load_workbook(path)
    wb["Assumptions"]["B3"] = scenario
    wb.save(path)
    assert recalc.recalc(path).get("status") == "success"
    checks = load_workbook(path, data_only=True)["Checks"]
    flag = next(r[18].value for r in checks.iter_rows() if r[0].value == "ALL CHECKS")
    assert flag is True


def test_assumptions_file_is_complete():
    import pandas as pd
    a = pd.read_csv(ROOT / "model" / "assumptions.csv")
    years = ["FY2026", "FY2027", "FY2028", "FY2029", "FY2030"]
    for key, g in a.groupby("key"):
        assert sorted(g.scenario) == ["Base", "Bear", "Bull"], key
        filled = g[years].notna()
        assert (filled.nunique() == 1).all(), f"{key}: scenarios fill different years"
        assert g.rationale.notna().all(), f"{key}: rationale missing"
