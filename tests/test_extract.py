"""Historicals extracted from the filings tie out and match known reported totals."""
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
PROCESSED = ROOT / "data" / "processed"


@pytest.fixture(scope="module", autouse=True)
def build():
    from src.delta import extract
    assert extract.main() == 0


def test_all_checks_pass():
    checks = pd.read_csv(PROCESSED / "checks.csv")
    failing = checks[~checks["pass"]]
    assert failing.empty, failing.to_string()


@pytest.mark.parametrize("item,period,value", [
    ("revenue", "FY2025", 554_885.168),      # FY2025 report p.10
    ("revenue", "FY2021", 314_670.796),      # FY2021 report p.10
    ("ni_parent", "FY2025", 60_108.399),     # FY2025 report p.11
    ("cash", "FY2025", 151_172.175),         # FY2025 report p.8
    ("revenue", "1H26", 342_608.732),        # Q2 2026 report p.7
    ("eps_basic", "FY2025", 23.14),
])
def test_known_values(item, period, value):
    h = pd.read_csv(PROCESSED / "historicals.csv", index_col=0)
    assert h.at[item, period] == pytest.approx(value, abs=0.001)


def test_segments_cover_every_period():
    s = pd.read_csv(PROCESSED / "segments.csv")
    latest = s[s.latest & (s["item"] == "revenue")]
    assert {"FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "1H26", "2Q26"} <= set(latest.period)
