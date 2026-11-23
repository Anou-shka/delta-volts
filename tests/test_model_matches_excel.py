"""The Excel workbook (recalculated by LibreOffice) and the Python mirror agree within 0.1%.

Covers every forecast year in all three scenarios, the DCF and price target, the reverse DCF (entering the solved
shift in the workbook makes the DCF value equal the share price) and a corner of the AI x margin sensitivity grid.
"""
import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.delta import model as mirror

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "model" / "delta_model.xlsx"
FC = mirror.FC
COLS = dict(zip(FC, "HIJKL"))
REL = 0.001

pytestmark = pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")


def _row(ws, label):
    return next(r[0].row for r in ws.iter_rows() if r[0].value == label)


def _workbook(tmp_path, name, scenario="Base", ai_shift=0.0, gm_shift=0.0):
    """Copy the built model, set scenario and shifts, recalculate, return values."""
    from src.delta import recalc
    path = tmp_path / f"{name}.xlsx"
    shutil.copy(MODEL, path)
    wb = load_workbook(path)
    a = wb["Assumptions"]
    a["B3"] = scenario
    a[f"C{_row(a, 'AI data-centre capex growth shift, FY2027E-FY2030E (pp, every year)')}"] = ai_shift
    a[f"C{_row(a, 'Gross margin shift, flowing through to EBIT (pp)')}"] = gm_shift
    wb.save(path)
    assert recalc.recalc(path).get("status") == "success"
    return load_workbook(path, data_only=True)


@pytest.fixture(scope="module", autouse=True)
def built():
    from src.delta import build_excel, recalc
    assert mirror.main() == 0
    assert build_excel.main() == 0
    assert recalc.recalc(MODEL).get("status") == "success"  # leave the committed model with cached values


def _dcf(wb, label):
    ws = wb["DCF"]
    return ws[f"C{_row(ws, label)}"].value


@pytest.mark.parametrize("scenario", ["Bear", "Base", "Bull"])
def test_forecast_and_dcf_match(tmp_path, scenario):
    wb = _workbook(tmp_path, scenario, scenario)
    py = mirror.valuation(scenario)
    f = py["forecast"]
    for p in FC:
        c = COLS[p]
        assert wb["IS"][f"{c}{_row(wb['IS'], 'Revenue')}"].value == pytest.approx(f["revenue"][p], rel=REL)
        assert wb["IS"][f"{c}{_row(wb['IS'], 'Operating profit (EBIT)')}"].value == pytest.approx(f["ebit"][p], rel=REL)
        assert wb["IS"][f"{c}{_row(wb['IS'], 'EPS, basic (NT$)')}"].value == pytest.approx(f["eps"][p], rel=REL)
        assert wb["DCF"][f"{c}{_row(wb['DCF'], 'Free cash flow to the firm')}"].value == pytest.approx(
            py["dcf"]["fcff"][p], rel=REL)
        assert abs(wb["BS"][f"{c}{_row(wb['BS'], 'Balance check: assets - (liabilities + equity)')}"].value) < 0.01
        assert abs(f["bs_diff"][p]) < 0.01
    assert _dcf(wb, "WACC") == pytest.approx(py["wacc"]["wacc"], rel=REL)
    assert _dcf(wb, "DCF value per share today (NT$)") == pytest.approx(py["dcf"]["value_today"], rel=REL)
    assert _dcf(wb, "Comps value per share today (NT$)") == pytest.approx(py["comps"]["value_today"], rel=REL)
    assert _dcf(wb, "12-month price target (NT$)") == pytest.approx(py["target"], rel=REL)
    assert _dcf(wb, "Rating (Buy > +15%, Sell < -10%, else Hold)") == py["rating"]


def test_reverse_dcf_reproduces_share_price(tmp_path):
    rd = mirror.reverse_dcf("Base")
    wb = _workbook(tmp_path, "reverse", ai_shift=rd["shift"])
    price = mirror.load_market()["delta_close"]
    assert _dcf(wb, "DCF value per share today (NT$)") == pytest.approx(price, rel=0.005)


def test_sensitivity_corner_reproduces(tmp_path):
    grid = mirror.sensitivity_ai_gm("Base")
    ai, gm = mirror.AI_SHIFTS[-1], mirror.GM_SHIFTS[0]
    wb = _workbook(tmp_path, "corner", ai_shift=ai, gm_shift=gm)
    assert _dcf(wb, "DCF value per share today (NT$)") == pytest.approx(grid.loc[gm, ai], rel=REL)
