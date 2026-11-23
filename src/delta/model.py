"""Python mirror of the Excel model: forecast, DCF, comps, sensitivities and reverse DCF.

Reads the same inputs as build_excel.py (data/processed/, model/assumptions.csv, model/valuation_inputs.csv,
data/raw/market/) and reproduces the workbook's formulas line by line. tests/test_model_matches_excel.py
recalculates the workbook and asserts both agree within 0.1%.

The mirror also produces what a spreadsheet cannot do live: the reverse DCF (solving for the AI capex growth the
share price implies), the AI-growth x margin sensitivity grid and the scenario values for the football field.

Run: python -m src.delta.model   -> prints the valuation summary and writes model/valuation_summary.json
"""
from __future__ import annotations

import datetime as dt
import json
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
PROCESSED = ROOT / "data" / "processed"
MARKET = ROOT / "data" / "raw" / "market"
FC = ["FY2026", "FY2027", "FY2028", "FY2029", "FY2030"]
SEGS = ["pe", "infra", "auto", "mob"]
SEG_NAME = {"pe": "Power Electronics", "infra": "Infrastructure", "auto": "Automation", "mob": "Mobility"}
BS_DATE = dt.date(2025, 12, 31)
DAYS = 365.0
DPS_2026_PAID = "dps_2026"  # assumption key


# ------------------------------------------------------------------ inputs
def load_assumptions(scenario: str = "Base") -> dict[str, dict[str, float]]:
    a = pd.read_csv(ROOT / "model" / "assumptions.csv")
    out: dict[str, dict[str, float]] = {}
    for _, r in a[a.scenario == scenario].iterrows():
        scale = 0.01 if r.unit == "pct" else 1.0
        out[r.key] = {p: float(r[p]) * scale for p in FC if not pd.isna(r[p])}
    return out


def load_valuation_inputs() -> dict[str, float]:
    v = pd.read_csv(ROOT / "model" / "valuation_inputs.csv")
    return {r.key: float(r.value) * (0.01 if r.unit == "pct" else 1.0) for _, r in v.iterrows()}


def load_market() -> dict:
    meta = json.loads((MARKET / "snapshot.json").read_text())
    peers = pd.read_csv(MARKET / "peers_snapshot.csv")
    px = pd.read_csv(MARKET / "prices_daily.csv", parse_dates=["date"], index_col="date")
    last_year = px["2308.TW"].dropna()
    last_year = last_year[last_year.index > last_year.index[-1] - pd.DateOffset(years=1)]
    meta.update({"peers": peers, "low_52w": float(last_year.min()), "high_52w": float(last_year.max())})
    return meta


@dataclass
class Base:
    """FY2025A balance sheet and P&L items, and 3Q26/9M26 segment figures (NT$ m)."""
    bs: dict[str, float]
    ni_parent: float
    ni_nci: float  # positive
    shares: float
    seg_9m: dict[str, float]
    seg_3q: dict[str, float]


def load_base() -> Base:
    h = pd.read_csv(PROCESSED / "historicals.csv", index_col=0)["FY2025"]
    s = lambda *k: float(sum(h[x] for x in k))  # noqa: E731
    bs = {
        "cash": float(h["cash"]),
        "st_investments": s("fvtpl_current", "fvoci_current", "amortised_cost_current"),
        "receivables": s("notes_receivable", "accounts_receivable", "ar_related"),
        "inventories": float(h["inventories"]),
        "lt_investments": s("fvtpl_noncurrent", "fvoci_noncurrent", "amortised_cost_noncurrent",
                            "equity_method_investments"),
        "ppe": float(h["ppe"]), "rou": float(h["right_of_use"]), "intangibles": float(h["intangibles"]),
        "st_debt": s("st_borrowings", "lt_debt_current"),
        "payables": s("notes_payable", "accounts_payable", "ap_related"),
        "lt_debt": s("bonds_payable", "lt_borrowings"), "leases": float(h["lease_liab_noncurrent"]),
        "equity_parent": float(h["equity_parent"]), "nci": float(h["nci"]),
    }
    bs["other_ca"] = float(h["total_current_assets"]) - (bs["cash"] + bs["st_investments"] + bs["receivables"]
                                                         + bs["inventories"])
    bs["other_nca"] = float(h["total_noncurrent_assets"]) - (bs["lt_investments"] + bs["ppe"] + bs["rou"]
                                                             + bs["intangibles"])
    bs["other_cl"] = float(h["total_current_liabilities"]) - bs["st_debt"] - bs["payables"]
    bs["other_ncl"] = float(h["total_noncurrent_liabilities"]) - bs["lt_debt"] - bs["leases"]
    seg = pd.read_csv(PROCESSED / "segments.csv")
    seg = seg[seg.latest & (seg["item"] == "revenue")]

    def seg_rev(period):
        g = seg[seg.period == period].set_index("segment").value / 1000
        out = {k: float(g[SEG_NAME[k]]) for k in SEGS}
        out["others"] = float(g["Others (derived: group revenue - segment total)"])
        return out
    return Base(bs=bs, ni_parent=float(h["ni_parent"]), ni_nci=float(h["ni_nci"]),
                shares=float(h["share_capital"]) / 10.0, seg_9m=seg_rev("9M26"), seg_3q=seg_rev("3Q26"))


# ------------------------------------------------------------------ forecast
@dataclass
class Forecast:
    rows: dict[str, dict[str, float]] = field(default_factory=dict)

    def set(self, key, p, v):
        self.rows.setdefault(key, {})[p] = v

    def __getitem__(self, key):
        return self.rows[key]


def forecast(a: dict, base: Base, ai_shift: float = 0.0, gm_shift: float = 0.0) -> Forecast:
    """Replicates the Segments, IS, Schedules, BS and CF forecast formulas of the workbook."""
    f = Forecast()
    prev_bs = dict(base.bs)
    prev = {"ni_parent": base.ni_parent, "ni_nci": -base.ni_nci, "shares": base.shares}
    seg_prev: dict[str, float] = {}
    q3_total = sum(base.seg_3q.values())
    for i, p in enumerate(FC):
        A = lambda k: a[k][p]  # noqa: E731
        A26 = lambda k: a[k]["FY2026"]  # noqa: E731
        s: dict[str, float] = {}
        if p == "FY2026":
            for k in SEGS + ["others"]:
                s[k] = base.seg_9m[k] + A26("q4_revenue") * base.seg_3q[k] / q3_total
            s["total"] = sum(s[k] for k in SEGS + ["others"])
            for line, share in (("ai_power", "ai_power_share"), ("ai_cooling", "ai_cooling_share"),
                                ("ai_other", "ai_other_share")):
                s[line] = A26(share) * s["total"]
        else:
            g = A("dc_capex_growth") + ai_shift
            for line, content in (("ai_power", "content_power"), ("ai_cooling", "content_cooling"),
                                  ("ai_other", "content_other")):
                s[line] = seg_prev[line] * (1 + g) * (1 + A(content))
        s["ai_total"] = s["ai_power"] + s["ai_cooling"] + s["ai_other"]
        s["ai_pe"] = (s["ai_power"] * A26("alloc_power_pe") + s["ai_cooling"] * A26("alloc_cooling_pe")
                      + s["ai_other"] * A26("alloc_other_pe"))
        s["ai_infra"] = s["ai_total"] - s["ai_pe"]
        if p == "FY2026":
            s["nonai_pe"] = s["pe"] - s["ai_pe"]
            s["nonai_infra"] = s["infra"] - s["ai_infra"]
        else:
            s["nonai_pe"] = seg_prev["nonai_pe"] * (1 + A("g_pe_nonai"))
            s["nonai_infra"] = seg_prev["nonai_infra"] * (1 + A("g_infra_nonai"))
            s["pe"] = s["nonai_pe"] + s["ai_pe"]
            s["infra"] = s["nonai_infra"] + s["ai_infra"]
            s["auto"] = seg_prev["auto"] * (1 + A("g_auto"))
            s["mob"] = seg_prev["mob"] * (1 + A("g_mob"))
            s["others"] = A("others_pct") * sum(s[k] for k in SEGS)
            s["total"] = sum(s[k] for k in SEGS + ["others"])
        for k in SEGS:
            s[f"op_{k}"] = s[k] * (A(f"m_{k}") + gm_shift)
        s["op_total"] = sum(s[f"op_{k}"] for k in SEGS)
        seg_prev = s

        rev = s["total"]
        cogs = -rev * (1 - (A("gross_margin") + gm_shift))
        gp = rev + cogs
        rnd = -A("rnd_pct") * rev
        sga = s["op_total"] - gp - rnd
        ebit = gp + sga + rnd
        total_debt_prev = prev_bs["st_debt"] + prev_bs["lt_debt"]
        int_inc = A("cash_yield") * prev_bs["cash"]
        fin = -A("cost_debt") * total_debt_prev
        other_nonop = A("other_nonop_pct") * rev
        pbt = ebit + int_inc + fin + other_nonop
        tax = -A("etr") * pbt
        ni = pbt + tax
        ni_nci = -A("nci_pct") * ni
        ni_parent = ni + ni_nci
        shares = prev["shares"]

        capex_total = A26("capex_2026") if p == "FY2026" else A("capex_pct") * rev
        capex_int = A("capex_int_pct") * rev
        capex_ppe = capex_total - capex_int
        dep = -A("dep_rate") * prev_bs["ppe"]
        ppe = prev_bs["ppe"] + capex_ppe + dep
        amort = A("amort_rate") * prev_bs["intangibles"]
        intangibles = prev_bs["intangibles"] + capex_int - amort
        da = -dep + amort

        bs = dict(prev_bs)
        bs["receivables"] = A("dso") / DAYS * rev
        bs["inventories"] = A("dio") / DAYS * -cogs
        bs["payables"] = A("dpo") / DAYS * -cogs
        bs["other_ca"] = A("other_ca_pct") * rev
        bs["other_cl"] = A("other_cl_pct") * rev
        bs["ppe"], bs["intangibles"] = ppe, intangibles
        dividends = (-A26(DPS_2026_PAID) * base.shares if p == "FY2026" else -A("payout") * prev["ni_parent"])
        div_nci = A("nci_div_pct") * prev["ni_nci"]
        bs["equity_parent"] = prev_bs["equity_parent"] + ni_parent + dividends
        bs["nci"] = prev_bs["nci"] - ni_nci + div_nci
        cfo = (pbt + da - (bs["receivables"] - prev_bs["receivables"]) - (bs["inventories"] - prev_bs["inventories"])
               + (bs["payables"] - prev_bs["payables"]) + tax - (bs["other_ca"] - prev_bs["other_ca"])
               + (bs["other_cl"] - prev_bs["other_cl"]))
        cfi = -capex_ppe - capex_int
        cff = 0.0 + dividends + div_nci
        bs["cash"] = prev_bs["cash"] + cfo + cfi + cff
        ta = sum(bs[k] for k in ("cash", "st_investments", "receivables", "inventories", "other_ca", "lt_investments",
                                 "ppe", "rou", "intangibles", "other_nca"))
        tle = sum(bs[k] for k in ("st_debt", "payables", "other_cl", "lt_debt", "leases", "other_ncl",
                                  "equity_parent", "nci"))
        nwc = bs["receivables"] + bs["inventories"] - bs["payables"] + bs["other_ca"] - bs["other_cl"]
        nwc_prev = (prev_bs["receivables"] + prev_bs["inventories"] - prev_bs["payables"] + prev_bs["other_ca"]
                    - prev_bs["other_cl"])
        vals = {**{f"seg_{k}": v for k, v in s.items()}, "revenue": rev, "cogs": cogs, "gross_profit": gp,
                "sga": sga, "r_and_d": rnd, "ebit": ebit, "interest_income": int_inc, "finance_costs": fin,
                "other_nonop": other_nonop, "pbt": pbt, "tax": tax, "net_income": ni, "ni_nci": ni_nci,
                "ni_parent": ni_parent, "shares": shares, "eps": ni_parent / shares, "d_and_a": da,
                "ebitda": ebit + da, "capex": capex_total, "capex_ppe": capex_ppe, "capex_int": capex_int,
                "dep": dep, "amort": amort, "cfo": cfo, "cfi": cfi, "cff": cff, "dividends": dividends,
                "div_nci": div_nci, "fcf": cfo - capex_total, "total_assets": ta, "bs_diff": ta - tle,
                "total_debt": bs["st_debt"] + bs["lt_debt"], "nwc": nwc, "delta_nwc": nwc - nwc_prev,
                "etr": A("etr"), "nci_pct": A("nci_pct"), "dps_paid": -dividends / shares,
                **{f"bs_{k}": v for k, v in bs.items()}}
        for k, v in vals.items():
            f.set(k, p, v)
        prev_bs = bs
        prev = {"ni_parent": ni_parent, "ni_nci": ni_nci, "shares": shares}
    return f


# ------------------------------------------------------------------ valuation
def wacc_build(v: dict, base: Base, market: dict) -> dict:
    debt = base.bs["st_debt"] + base.bs["lt_debt"] + base.bs["leases"]
    equity = market["delta_close"] * base.shares
    de = debt / equity
    beta_l = v["beta_u"] * (1 + (1 - v["tax_marginal"]) * de)
    ke = v["rf"] + beta_l * (v["mature_erp"] + v["crp"])
    kd_at = v["kd"] * (1 - v["tax_marginal"])
    wd = debt / (debt + equity)
    wacc = (1 - wd) * ke + wd * kd_at
    return {"debt": debt, "equity_mkt": equity, "de": de, "beta_l": beta_l, "ke": ke, "kd_at": kd_at, "wd": wd,
            "wacc": wacc}


def dcf(f: Forecast, base: Base, v: dict, w: dict, market: dict, wacc: float | None = None,
        g: float | None = None) -> dict:
    wacc = w["wacc"] if wacc is None else wacc
    g = v["g"] if g is None else g
    fcff, pv = {}, {}
    for i, p in enumerate(FC):
        nopat = f["ebit"][p] * (1 - f["etr"][p])
        fcff[p] = nopat + f["d_and_a"][p] - f["capex"][p] - f["delta_nwc"][p]
        pv[p] = fcff[p] / (1 + wacc) ** (i + 0.5)
    tv = fcff["FY2030"] * (1 + g) / (wacc - g)
    pv_tv = tv / (1 + wacc) ** 5
    ev = sum(pv.values()) + pv_tv
    b = base.bs
    net_cash = b["cash"] + b["st_investments"] + b["lt_investments"] - b["st_debt"] - b["lt_debt"] - b["leases"]
    equity = ev + net_cash
    parent = equity * (1 - f["nci_pct"]["FY2026"])
    per_share_bs = parent / base.shares
    price_date = dt.date.fromisoformat(market["price_date"])
    t = (price_date - BS_DATE).days / DAYS
    dps26 = -f["dividends"]["FY2026"] / base.shares
    today = per_share_bs * (1 + w["ke"]) ** t - dps26
    ebitda30 = f["ebitda"]["FY2030"]
    tv_exit = ebitda30 * v["exit_multiple"]
    return {"fcff": fcff, "pv": pv, "tv": tv, "pv_tv": pv_tv, "ev": ev, "tv_share": pv_tv / ev,
            "net_cash": net_cash, "equity": equity, "equity_parent": parent, "per_share_bs": per_share_bs,
            "years_to_price_date": t, "value_today": today, "implied_exit_multiple": tv / ebitda30,
            "tv_exit": tv_exit, "implied_g_exit": (tv_exit * wacc - fcff["FY2030"]) / (tv_exit + fcff["FY2030"]),
            "ev_exit": sum(pv.values()) + tv_exit / (1 + wacc) ** 5}


def comps(f: Forecast, base: Base, market: dict) -> dict:
    p = market["peers"].copy()
    p["ev_m"] = p.market_cap_m + p.total_debt_m - p.total_cash_m
    p["ev_ebitda_ttm"] = p.ev_m / p.ebitda_ttm_m
    p["ebitda_fy1_m"] = p.ebitda_ttm_m * p.revenue_fy1_m / p.revenue_ttm_m
    p["ev_ebitda_fy1"] = p.ev_m / p.ebitda_fy1_m
    p["ev_sales_ttm"] = p.ev_m / p.revenue_ttm_m
    p["ev_sales_fy1"] = p.ev_m / p.revenue_fy1_m
    p["pe_ttm"] = p.price / p.eps_ttm
    p["pe_fy1"] = p.price / p.eps_fy1
    peers = p[p.ticker != "2308.TW"]
    stats = {m: {"q1": float(peers[m].quantile(0.25)), "median": float(peers[m].median()),
                 "q3": float(peers[m].quantile(0.75))} for m in ("ev_ebitda_fy1", "pe_fy1", "ev_ebitda_ttm", "pe_ttm")}
    # Delta on FY2027E (the peers' next fiscal year); net cash and NCI share at FY2026E year end
    ebitda27, eps27 = f["ebitda"]["FY2027"], f["eps"]["FY2027"]
    nc26 = (f["bs_cash"]["FY2026"] + f["bs_st_investments"]["FY2026"] + f["bs_lt_investments"]["FY2026"]
            - f["bs_st_debt"]["FY2026"] - f["bs_lt_debt"]["FY2026"] - f["bs_leases"]["FY2026"])
    nci = f["nci_pct"]["FY2027"]

    def from_ev_ebitda(mult):
        return (mult * ebitda27 + nc26) * (1 - nci) / base.shares
    out = {"table": p, "stats": stats,
           "ev_ebitda": {k: from_ev_ebitda(stats["ev_ebitda_fy1"][k]) for k in ("q1", "median", "q3")},
           "pe": {k: stats["pe_fy1"][k] * eps27 for k in ("q1", "median", "q3")}}
    out["value_today"] = (out["ev_ebitda"]["median"] + out["pe"]["median"]) / 2
    return out


def valuation(scenario: str = "Base", ai_shift: float = 0.0, gm_shift: float = 0.0, wacc: float | None = None,
              g: float | None = None) -> dict:
    a, v, base, market = load_assumptions(scenario), load_valuation_inputs(), load_base(), load_market()
    f = forecast(a, base, ai_shift, gm_shift)
    w = wacc_build(v, base, market)
    d = dcf(f, base, v, w, market, wacc, g)
    c = comps(f, base, market)
    price = market["delta_close"]
    blend_today = v["weight_dcf"] * d["value_today"] + (1 - v["weight_dcf"]) * c["value_today"]
    dps_next = f["dps_paid"]["FY2027"]
    target = blend_today * (1 + w["ke"]) - dps_next
    upside = target / price - 1
    rating = "Buy" if upside > v["upside_buy"] else "Sell" if upside < -v["downside_sell"] else "Hold"
    return {"forecast": f, "wacc": w, "dcf": d, "comps": c, "price": price, "price_date": market["price_date"],
            "blend_today": blend_today, "target": target, "upside": upside, "rating": rating, "market": market,
            "inputs": v}


def reverse_dcf(scenario: str = "Base", lo: float = -0.6, hi: float = 0.6) -> dict:
    """Uniform shift to AI data-centre capex growth (FY2027E-FY2030E) at which DCF value today = share price."""
    a, v, base, market = load_assumptions(scenario), load_valuation_inputs(), load_base(), load_market()
    w = wacc_build(v, base, market)
    price = market["delta_close"]

    def gap(shift):
        return dcf(forecast(a, base, ai_shift=shift), base, v, w, market)["value_today"] - price
    if gap(lo) > 0 or gap(hi) < 0:
        return {"shift": float("nan")}
    for _ in range(80):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if gap(mid) < 0 else (lo, mid)
    shift = (lo + hi) / 2
    f = forecast(a, base, ai_shift=shift)
    base_g = [a["dc_capex_growth"][p] for p in FC[1:]]
    implied = [x + shift for x in base_g]
    cagr = lambda xs: float(np.prod([1 + x for x in xs]) ** (1 / len(xs)) - 1)  # noqa: E731
    return {"shift": shift, "base_capex_growth": base_g, "implied_capex_growth": implied,
            "base_cagr": cagr(base_g), "implied_cagr": cagr(implied),
            "ai_revenue_2030": f["seg_ai_total"]["FY2030"], "ai_share_2030": f["seg_ai_total"]["FY2030"]
            / f["revenue"]["FY2030"], "revenue_2030": f["revenue"]["FY2030"], "eps_2030": f["eps"]["FY2030"]}


AI_SHIFTS = [-0.20, -0.10, 0.0, 0.10, 0.20]
GM_SHIFTS = [-0.02, -0.01, 0.0, 0.01, 0.02]


def sensitivity_ai_gm(scenario: str = "Base") -> pd.DataFrame:
    a, v, base, market = load_assumptions(scenario), load_valuation_inputs(), load_base(), load_market()
    w = wacc_build(v, base, market)
    grid = {gm: [dcf(forecast(a, base, ai, gm), base, v, w, market)["value_today"] for ai in AI_SHIFTS]
            for gm in GM_SHIFTS}
    return pd.DataFrame(grid, index=AI_SHIFTS).T  # rows: gm shift, cols: ai shift


def summary() -> dict:
    out = {}
    for scen in ("Bear", "Base", "Bull"):
        r = valuation(scen)
        f = r["forecast"]
        out[scen] = {"dcf_value_today": r["dcf"]["value_today"], "comps_value_today": r["comps"]["value_today"],
                     "target": r["target"], "eps": {p: f["eps"][p] for p in FC},
                     "revenue": {p: f["revenue"][p] for p in FC}}
    base = valuation("Base")
    rd = reverse_dcf("Base")
    sens = sensitivity_ai_gm("Base")
    wacc_steps = [-0.01, -0.005, 0.0, 0.005, 0.01]
    g_steps = [-0.01, -0.005, 0.0, 0.005, 0.01]
    a, v, b, m = load_assumptions("Base"), load_valuation_inputs(), load_base(), load_market()
    f0 = base["forecast"]
    w0 = base["wacc"]["wacc"]
    wg = [[dcf(f0, b, v, base["wacc"], m, w0 + dw, v["g"] + dg)["value_today"] for dg in g_steps] for dw in wacc_steps]
    return {"price_date": base["price_date"], "price": base["price"], "wacc": base["wacc"],
            "sensitivity_wacc_g": {"wacc": [w0 + d for d in wacc_steps], "g": [v["g"] + d for d in g_steps],
                                   "values": wg},
            "dcf": {k: v for k, v in base["dcf"].items() if not isinstance(v, dict)},
            "comps": {"ev_ebitda": base["comps"]["ev_ebitda"], "pe": base["comps"]["pe"],
                      "stats": base["comps"]["stats"], "value_today": base["comps"]["value_today"]},
            "blend_today": base["blend_today"], "target": base["target"], "upside": base["upside"],
            "rating": base["rating"], "scenarios": out, "reverse_dcf": rd,
            "sensitivity_ai_gm": {"ai_shifts": AI_SHIFTS, "gm_shifts": GM_SHIFTS, "values": sens.values.tolist()},
            "low_52w": base["market"]["low_52w"], "high_52w": base["market"]["high_52w"]}


def main() -> int:
    s = summary()
    path = ROOT / "model" / "valuation_summary.json"
    path.write_text(json.dumps(s, indent=2, default=float))
    w, d = s["wacc"], s["dcf"]
    print(f"Price {s['price']:.0f} ({s['price_date']}) | WACC {w['wacc']:.2%} (Ke {w['ke']:.2%}, beta {w['beta_l']:.2f})")
    print(f"DCF: EV {d['ev']:,.0f}m, TV share {d['tv_share']:.0%}, value today NT${d['value_today']:.0f} "
          f"(implied exit {d['implied_exit_multiple']:.1f}x)")
    print(f"Comps value today NT${s['comps']['value_today']:.0f} | blend {s['blend_today']:.0f} | "
          f"12m target NT${s['target']:.0f} ({s['upside']:+.0%}) -> {s['rating']}")
    for scen, x in s["scenarios"].items():
        print(f"  {scen}: DCF {x['dcf_value_today']:.0f}, comps {x['comps_value_today']:.0f}, target {x['target']:.0f}")
    rd = s["reverse_dcf"]
    print(f"Reverse DCF: AI capex growth shift {rd['shift']:+.1%} p.a. -> implied 2027-30 CAGR {rd['implied_cagr']:.1%} "
          f"(Base {rd['base_cagr']:.1%}); FY2030E AI share {rd['ai_share_2030']:.0%}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
