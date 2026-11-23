"""Snapshot market data for the valuation into data/raw/market/ (run deliberately: `make market`).

Writes:
  prices_daily.csv      daily closes, 2308.TW and TAIEX (^TWII), 5 years (for the beta regression, 52-week range)
  peers_snapshot.csv    Delta and peers: price, shares, market cap, debt, cash, TTM revenue/EBITDA, consensus
                        EPS and revenue (current and next fiscal year), FX to NT$
  snapshot.json         the date and time of the snapshot

Source: Yahoo Finance via yfinance (consensus figures are Yahoo's analyst averages). The valuation reads only
these files, so it reproduces exactly until the snapshot is refreshed.
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

import pandas as pd
import yfinance as yf

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "raw" / "market"

DELTA = "2308.TW"
INDEX = "^TWII"
PEERS = {
    "2308.TW": ("Delta Electronics", "Subject company"),
    "VRT": ("Vertiv", "Data-centre power and thermal pure play (US)"),
    "SU.PA": ("Schneider Electric", "Electrical equipment incl. data-centre power and cooling (EU)"),
    "ETN": ("Eaton", "Electrical equipment incl. data-centre power (US)"),
    "2301.TW": ("Lite-On Technology", "Taiwan power-supply peer"),
    "6282.TW": ("AcBel Polytech", "Taiwan power-supply peer"),
}
FX = {"TWD": None, "USD": "TWD=X", "EUR": "EURTWD=X"}


def _est(df, col):
    try:
        return float(df.loc[col, "avg"])
    except Exception:
        return float("nan")


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    now = dt.datetime.now(dt.timezone.utc)
    px = {}
    for t in (DELTA, INDEX):
        h = yf.Ticker(t).history(period="5y", auto_adjust=False)
        px[t] = h["Close"].tz_localize(None)
    prices = pd.DataFrame(px).dropna(how="all")
    prices.index.name = "date"
    prices.round(4).to_csv(OUT / "prices_daily.csv")

    fx = {}
    for ccy, t in FX.items():
        fx[ccy] = 1.0 if t is None else float(yf.Ticker(t).history(period="5d")["Close"].iloc[-1])

    rows = []
    for t, (name, role) in PEERS.items():
        tk = yf.Ticker(t)
        i = tk.info
        rev, eps = tk.revenue_estimate, tk.earnings_estimate
        rows.append({
            "ticker": t, "name": name, "role": role, "currency": i.get("currency"),
            "price": i.get("currentPrice") or i.get("regularMarketPrice"),
            "shares_m": (i.get("sharesOutstanding") or float("nan")) / 1e6,
            "market_cap_m": (i.get("marketCap") or float("nan")) / 1e6,
            "total_debt_m": (i.get("totalDebt") or 0) / 1e6,
            "total_cash_m": (i.get("totalCash") or 0) / 1e6,
            "revenue_ttm_m": (i.get("totalRevenue") or float("nan")) / 1e6,
            "ebitda_ttm_m": (i.get("ebitda") or float("nan")) / 1e6,
            "eps_ttm": i.get("trailingEps"),
            "eps_fy0": _est(eps, "0y"), "eps_fy1": _est(eps, "+1y"),
            "revenue_fy0_m": _est(rev, "0y") / 1e6, "revenue_fy1_m": _est(rev, "+1y") / 1e6,
            "low_52w": i.get("fiftyTwoWeekLow"), "high_52w": i.get("fiftyTwoWeekHigh"),
            "last_fiscal_year_end": dt.datetime.fromtimestamp(i.get("lastFiscalYearEnd", 0), dt.timezone.utc).date(),
            "most_recent_quarter": dt.datetime.fromtimestamp(i.get("mostRecentQuarter", 0), dt.timezone.utc).date(),
            "fx_to_twd": fx.get(i.get("currency"), float("nan")),
        })
    pd.DataFrame(rows).to_csv(OUT / "peers_snapshot.csv", index=False)
    last = prices[DELTA].dropna()
    meta = {"retrieved_utc": now.isoformat(timespec="seconds"), "source": "Yahoo Finance via yfinance",
            "price_date": str(last.index[-1].date()), "delta_close": float(last.iloc[-1]), "fx": fx}
    (OUT / "snapshot.json").write_text(json.dumps(meta, indent=2))
    print(json.dumps(meta, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
