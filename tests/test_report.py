"""The report is generated from the model: no missing values, every cited footnote defined, numbers in sync."""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_report_generates_cleanly():
    from src.delta import report
    md = report.write()
    assert not re.search(r"\bnan\b|\bNone\b|\binf\b", md), "missing value rendered in report"
    cited = set(re.findall(r"\[\^([a-z0-9]+)\](?!:)", md))
    defined = set(re.findall(r"^\[\^([a-z0-9]+)\]:", md, flags=re.M))
    assert cited <= defined, f"undefined footnotes: {cited - defined}"
    s = json.loads((ROOT / "model" / "valuation_summary.json").read_text())
    assert f"NT${s['target']:,.0f}" in md and f"NT${s['dcf']['value_today']:,.0f}" in md
    assert md.count("[ANALYST TO REVIEW]") >= 5
