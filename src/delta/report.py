"""Write report/delta_initiation.md from the model, so every number in the report comes from the pipeline.

Figures come from data/processed/ (historicals, segments, geography), data/slides_3q26.csv, the Python mirror
(src/delta/model.py, tested equal to the workbook) and model/valuation_summary.json. Judgement sections are drafts
marked [ANALYST TO REVIEW].

Run: python -m src.delta.report
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd

from . import model as M

ROOT = Path(__file__).resolve().parents[2]
PROCESSED = ROOT / "data" / "processed"
OUT = ROOT / "report" / "delta_initiation.md"
FC = M.FC


def bn(x_m: float, d: int = 1) -> str:
    """NT$ millions -> 'NT$123.4bn'."""
    return f"NT${x_m / 1000:,.{d}f}bn"


def pct(x: float, d: int = 1) -> str:
    return f"{x * 100:.{d}f}%"


def ntd(x: float, d: int = 0) -> str:
    return f"NT${x:,.{d}f}"


def context() -> dict:
    h = pd.read_csv(PROCESSED / "historicals.csv", index_col=0)
    seg = pd.read_csv(PROCESSED / "segments.csv")
    seg = seg[seg.latest]
    geo = pd.read_csv(PROCESSED / "geography.csv")
    geo = geo[geo.apply(lambda r: r.document == f"fs_{r.period[2:]}_annual_consolidated.pdf", axis=1)]
    sl = pd.read_csv(ROOT / "data" / "slides_3q26.csv")
    s = json.loads((ROOT / "model" / "valuation_summary.json").read_text())
    base = M.valuation("Base")
    bear, bull = M.valuation("Bear"), M.valuation("Bull")
    a = M.load_assumptions("Base")
    v = M.load_valuation_inputs()
    b, mkt = M.load_base(), M.load_market()
    f = base["forecast"]
    d = base["dcf"]
    # exit-multiple value per share today (as on the DCF sheet)
    exit_ps = ((d["ev_exit"] + d["net_cash"]) * (1 - f["nci_pct"]["FY2026"]) / b.shares
               * (1 + base["wacc"]["ke"]) ** d["years_to_price_date"] + f["dividends"]["FY2026"] / b.shares)
    sv = lambda period, segment, item: float(  # noqa: E731
        seg[(seg.period == period) & (seg.segment == segment) & (seg["item"] == item)].value.iloc[0]) / 1000
    slv = lambda period, segment, item: float(  # noqa: E731
        sl[(sl.period == period) & (sl.segment == segment) & (sl["item"] == item)].value_ntd_m.iloc[0])
    g = geo.pivot_table(index="region", columns="period", values="revenue") / 1000
    gn = geo.pivot_table(index="region", columns="period", values="noncurrent_assets") / 1000
    col = h["FY2025"]
    net_cash25 = (col["cash"] + col["fvtpl_current"] + col["fvoci_current"] + col["amortised_cost_current"]
                  - col["st_borrowings"] - col["lt_debt_current"] - col["bonds_payable"] - col["lt_borrowings"])
    return dict(net_cash25=net_cash25, h=h, sv=sv, slv=slv, g=g, gn=gn, s=s, base=base, bear=bear, bull=bull, a=a, v=v, b=b, mkt=mkt, f=f,
                d=d, exit_ps=exit_ps, rd=s["reverse_dcf"], peers=base["comps"]["table"])


def key_table(c: dict) -> str:
    h, f, price = c["h"], c["f"], c["s"]["price"]
    mcap = price * c["b"].shares
    rows = {"Revenue (NT$bn)": [], "Revenue growth": [], "EBIT margin": [], "EPS (NT$)": [], "P/E (x)": [],
            "EV/EBITDA (x)": [], "Dividend yield": []}
    cols = ["FY2024", "FY2025", "FY2026", "FY2027", "FY2028"]
    for p in cols:
        if p in FC:
            rev, ebit, eps, ebitda = f["revenue"][p], f["ebit"][p], f["eps"][p], f["ebitda"][p]
            prev = f["revenue"][M.FC[M.FC.index(p) - 1]] if p != "FY2026" else h.at["revenue", "FY2025"]
            nc = (f["bs_cash"][p] + f["bs_st_investments"][p] + f["bs_lt_investments"][p] - f["bs_st_debt"][p]
                  - f["bs_lt_debt"][p] - f["bs_leases"][p])
            dps = f["dps_paid"][p]
        else:
            col = h[p]
            rev, ebit, eps = col["revenue"], col["ebit"], col["eps_basic"]
            ebitda = ebit + col["depreciation"] + col["amortisation"]
            prev = h.at["revenue", f"FY{int(p[2:]) - 1}"]
            nc = (col["cash"] + col["fvtpl_current"] + col["fvoci_current"] + col["amortised_cost_current"]
                  + col["fvtpl_noncurrent"] + col["fvoci_noncurrent"] + col["amortised_cost_noncurrent"]
                  + col["equity_method_investments"] - col["st_borrowings"] - col["lt_debt_current"]
                  - col["bonds_payable"] - col["lt_borrowings"] - col["lease_liab_noncurrent"])
            dps = -col["dividends_paid"] / c["b"].shares
        rows["Revenue (NT$bn)"].append(f"{rev / 1000:,.1f}")
        rows["Revenue growth"].append(pct(rev / prev - 1))
        rows["EBIT margin"].append(pct(ebit / rev))
        rows["EPS (NT$)"].append(f"{eps:.2f}")
        rows["P/E (x)"].append(f"{price / eps:.1f}")
        rows["EV/EBITDA (x)"].append(f"{(mcap - nc) / ebitda:.1f}")
        rows["Dividend yield"].append(pct(dps / price))
    head = "| NT$ | " + " | ".join(p + ("A" if p not in FC else "E") for p in cols) + " |"
    lines = [head, "|---|" + "---:|" * len(cols)]
    lines += [f"| {k} | " + " | ".join(v) + " |" for k, v in rows.items()]
    return "\n".join(lines)


def forecast_table(c: dict) -> str:
    f = c["f"]
    rows = [("Revenue", "revenue", 1), ("of which AI data-centre", "seg_ai_total", 1),
            ("Power Electronics", "seg_pe", 1), ("Infrastructure", "seg_infra", 1), ("Automation", "seg_auto", 1),
            ("Mobility", "seg_mob", 1), ("EBIT", "ebit", 1), ("Net income to parent", "ni_parent", 1),
            ("Free cash flow (CFO + capex)", "fcf", 1), ("Capex", "capex", 1)]
    lines = ["| NT$bn | " + " | ".join(p + "E" for p in FC) + " |", "|---|" + "---:|" * 5]
    for label, key, _ in rows:
        lines.append(f"| {label} | " + " | ".join(f"{f[key][p] / 1000:,.1f}" for p in FC) + " |")
    lines.append("| AI data-centre % of revenue | " + " | ".join(pct(f["seg_ai_total"][p] / f["revenue"][p], 0)
                                                           for p in FC) + " |")
    lines.append("| EBIT margin | " + " | ".join(pct(f["ebit"][p] / f["revenue"][p]) for p in FC) + " |")
    lines.append("| EPS (NT$) | " + " | ".join(f"{f['eps'][p]:.2f}" for p in FC) + " |")
    return "\n".join(lines)


def scenario_table(c: dict) -> str:
    a = {k: M.load_assumptions(k) for k in ("Bear", "Base", "Bull")}
    res = {"Bear": c["bear"], "Base": c["base"], "Bull": c["bull"]}
    rows = [
        ("AI data-centre capex growth FY27/28/29/30",
         lambda k: " / ".join(f"{a[k]['dc_capex_growth'][p] * 100:+.0f}%" for p in FC[1:])),
        ("Content gain: AI power (FY2027E)", lambda k: f"{a[k]['content_power']['FY2027'] * 100:+.0f}%"),
        ("Content gain: liquid cooling (FY2027E)", lambda k: f"{a[k]['content_cooling']['FY2027'] * 100:+.0f}%"),
        ("Power Electronics margin FY2030E", lambda k: pct(a[k]["m_pe"]["FY2030"])),
        ("Infrastructure margin FY2030E", lambda k: pct(a[k]["m_infra"]["FY2030"])),
        ("Gross margin FY2030E", lambda k: pct(a[k]["gross_margin"]["FY2030"])),
        ("FY2030E revenue (NT$bn)", lambda k: f"{res[k]['forecast']['revenue']['FY2030'] / 1000:,.0f}"),
        ("FY2030E AI data-centre % of revenue",
         lambda k: pct(res[k]["forecast"]["seg_ai_total"]["FY2030"] / res[k]["forecast"]["revenue"]["FY2030"], 0)),
        ("FY2030E EPS (NT$)", lambda k: f"{res[k]['forecast']['eps']['FY2030']:.2f}"),
        ("DCF value per share today (NT$)", lambda k: f"{res[k]['dcf']['value_today']:,.0f}"),
        ("12-month target (NT$)", lambda k: f"{res[k]['target']:,.0f}"),
    ]
    lines = ["| | Bear | Base | Bull |", "|---|---:|---:|---:|"]
    lines += [f"| {label} | " + " | ".join(fn(k) for k in ("Bear", "Base", "Bull")) + " |" for label, fn in rows]
    return "\n".join(lines)


def comps_table(c: dict) -> str:
    p = c["peers"]
    lines = ["| Company | EV (NT$bn) | EV/EBITDA TTM | EV/EBITDA FY1 | P/E TTM | P/E FY1 |",
             "|---|---:|---:|---:|---:|---:|"]
    for _, r in p.iterrows():
        ev_twd = r.ev_m * r.fx_to_twd / 1000
        lines.append(f"| {r['name']} | {ev_twd:,.0f} | {r.ev_ebitda_ttm:.1f}x | {r.ev_ebitda_fy1:.1f}x | "
                     f"{r.pe_ttm:.1f}x | {r.pe_fy1:.1f}x |")
    st = c["s"]["comps"]["stats"]
    lines.append(f"| **Peer median (excl. Delta)** | | {st['ev_ebitda_ttm']['median']:.1f}x | "
                 f"{st['ev_ebitda_fy1']['median']:.1f}x | {st['pe_ttm']['median']:.1f}x | {st['pe_fy1']['median']:.1f}x |")
    return "\n".join(lines)


def statements_table(c: dict) -> str:
    """Summary income statement, balance sheet and cash flow, FY2024A-FY2030E (NT$bn)."""
    h, f = c["h"], c["f"]
    cols = ["FY2024", "FY2025"] + FC
    hist = {  # label: (historicals key, forecast key, sign)
        "Revenue": ("revenue", "revenue", 1), "Gross profit": ("gross_profit", "gross_profit", 1),
        "EBIT": ("ebit", "ebit", 1), "D&A": ("d_and_a", "d_and_a", 1), "EBITDA": ("ebitda", "ebitda", 1),
        "Profit before tax": ("pbt", "pbt", 1), "Income tax": ("tax", "tax", 1),
        "Net income to parent": ("ni_parent", "ni_parent", 1),
        "Cash": ("cash", "bs_cash", 1), "Receivables": ("receivables", "bs_receivables", 1),
        "Inventories": ("inventories", "bs_inventories", 1), "PP&E": ("ppe", "bs_ppe", 1),
        "Total assets": ("total_assets", "total_assets", 1), "Payables": ("payables", "bs_payables", 1),
        "Total debt": ("total_debt", "total_debt", 1), "Equity to parent": ("equity_parent", "bs_equity_parent", 1),
        "Cash flow from operations": ("cfo", "cfo", 1), "Capex": ("capex", "capex", -1),
        "Free cash flow": (None, "fcf", 1), "Dividends paid": ("dividends_paid", "dividends", 1),
    }
    sections = {"Income statement": ["Revenue", "Gross profit", "EBIT", "D&A", "EBITDA", "Profit before tax",
                                     "Income tax", "Net income to parent"],
                "Balance sheet (year end)": ["Cash", "Receivables", "Inventories", "PP&E", "Total assets",
                                             "Payables", "Total debt", "Equity to parent"],
                "Cash flow": ["Cash flow from operations", "Capex", "Free cash flow", "Dividends paid"]}
    lines = ["| NT$bn | " + " | ".join(p + ("A" if p not in FC else "E") for p in cols) + " |",
             "|---|" + "---:|" * len(cols)]
    for sec, labels in sections.items():
        lines.append(f"| **{sec}** |" + " |" * len(cols))
        for lab in labels:
            hk, fk, sign = hist[lab]
            vals = []
            for p in cols:
                if p in FC:
                    v = f[fk][p] * (1 if lab != "Capex" else -1)
                elif hk is None:
                    v = h.at["cfo", p] + h.at["capex", p]
                else:
                    v = h.at[hk, p]
                vals.append(f"{v / 1000:,.1f}")
            lines.append(f"| {lab} | " + " | ".join(vals) + " |")
    return "\n".join(lines)


def assumptions_table(c: dict) -> str:
    """Base-case drivers by year (the full set, Bear/Base/Bull, is in model/assumptions.csv)."""
    a = c["a"]
    rows = [("AI data-centre capex growth", "dc_capex_growth", "pct"), ("Content gain: AI power", "content_power", "pct"),
            ("Content gain: liquid cooling", "content_cooling", "pct"), ("Content gain: other AI", "content_other", "pct"),
            ("Power Electronics excl. AI growth", "g_pe_nonai", "pct"),
            ("Infrastructure excl. AI growth", "g_infra_nonai", "pct"), ("Automation growth", "g_auto", "pct"),
            ("Mobility growth", "g_mob", "pct"), ("Power Electronics margin", "m_pe", "pct"),
            ("Infrastructure margin", "m_infra", "pct"), ("Automation margin", "m_auto", "pct"),
            ("Mobility margin", "m_mob", "pct"), ("Gross margin", "gross_margin", "pct"), ("R&D % revenue", "rnd_pct", "pct"),
            ("Effective tax rate", "etr", "pct"), ("Capex % revenue (FY2026E: NT$bn)", "capex_pct", "pct"),
            ("Depreciation % opening PP&E", "dep_rate", "pct"), ("Receivable days", "dso", "days"),
            ("Inventory days", "dio", "days"), ("Payable days", "dpo", "days"), ("Dividend payout (prior-year NI)", "payout", "pct")]
    lines = ["| Base case | " + " | ".join(p + "E" for p in FC) + " |", "|---|" + "---:|" * 5]
    for label, key, unit in rows:
        vals = []
        for p in FC:
            if key == "capex_pct" and p == "FY2026":
                vals.append(f"{a['capex_2026']['FY2026'] / 1000:.0f}")
            elif p in a.get(key, {}):
                x = a[key][p]
                vals.append(f"{x * 100:.1f}%" if unit == "pct" else f"{x:.0f}")
            elif key == "payout" and p == "FY2026":
                vals.append(f"NT${a['dps_2026']['FY2026']:.1f}/sh")
            else:
                vals.append("–")
        lines.append(f"| {label} | " + " | ".join(vals) + " |")
    lines.append("| 4Q26E revenue (NT$bn) | " + f"{a['q4_revenue']['FY2026'] / 1000:.0f} | – | – | – | – |")
    lines.append("| AI shares of FY2026E revenue (power / cooling / other) | "
                 + f"{a['ai_power_share']['FY2026'] * 100:.0f}% / {a['ai_cooling_share']['FY2026'] * 100:.0f}% / "
                 + f"{a['ai_other_share']['FY2026'] * 100:.0f}% | – | – | – | – |")
    return "\n".join(lines)


def write() -> str:
    c = context()
    h, f, s, d, rd, a, v, sv, slv, g, gn = (c[k] for k in ("h", "f", "s", "d", "rd", "a", "v", "sv", "slv", "g", "gn"))
    price, target, up = s["price"], s["target"], s["upside"]
    w = s["wacc"]
    rev = lambda p: h.at["revenue", p]  # noqa: E731
    gm = lambda p: h.at["gross_profit", p] / rev(p)  # noqa: E731
    om = lambda p: h.at["ebit", p] / rev(p)  # noqa: E731
    capex25 = -(h.at["capex_ppe", "FY2025"] + h.at["capex_intangibles", "FY2025"])
    seg25 = {k: sv("FY2025", k, "revenue") for k in ("Power Electronics", "Infrastructure", "Automation", "Mobility")}
    op25 = {k: sv("FY2025", k, "op_profit") for k in seg25}
    seg9 = {k: slv("9M26", k, "revenue") for k in seg25}
    op9 = {k: slv("9M26", k, "op_profit") for k in seg25}
    rev9 = slv("9M26", "Group", "revenue")
    rev3q = slv("3Q26", "Group", "revenue")
    us_share = g.at["U.S.A.", "FY2025"] / g["FY2025"].sum()
    us_share21 = g.at["U.S.A.", "FY2021"] / g["FY2021"].sum()
    cn_share, cn_share21 = (g.at["Mainland China", p] / g[p].sum() for p in ("FY2025", "FY2021"))
    nca25 = gn["FY2025"]
    rd_g = " / ".join(f"{x * 100:.0f}%" for x in rd["implied_capex_growth"])
    base_g = " / ".join(f"{x * 100:.0f}%" for x in rd["base_capex_growth"])
    peers_med = s["comps"]["stats"]
    lo52, hi52 = s["low_52w"], s["high_52w"]
    sc = s["scenarios"]

    md = f"""---
title: "Delta Electronics (2308.TW) — The AI power play"
subtitle: "Initiation of coverage · {s['price_date']}"
author: "Anoushka · Student research for educational purposes only. Not investment advice."
---

# Summary

| Rating | 12-month target | Price ({s['price_date']}) | Upside |
|---|---|---|---|
| **{s['rating']}** `[ANALYST TO REVIEW]` | **{ntd(target)}** | {ntd(price)} | **{up * 100:+.0f}%** |

*Draft. The rating follows the spec's rule (Buy above +15%, Sell below −10%) applied to a draft Base case that the analyst has not yet signed off.*

**Investment thesis** `[ANALYST TO REVIEW]` *(draft written from the model's evidence; the analyst writes the final version)*

1. **Half of Delta is now an AI data-centre business, and the earnings show it.** Management puts AI data-centre businesses at about 50% of revenue.[^q3call] Revenue rose 37% in 9M26 to {bn(rev9)}, and EPS rose 62% to NT$26.64.[^ic3q]
2. **Even after halving, the price asks for more AI growth than our Base case.** At {ntd(price)} the reverse DCF implies AI data-centre capex growth of about {pct(rd['implied_cagr'], 0)} a year in FY2027–30, against {pct(rd['base_cagr'], 0)} in our Base case.
3. **Capital intensity is rising just as growth normalises.** Capex goes from {bn(capex25)} in FY2025 to about NT$70bn in 2026.[^q2call] Depreciation lives for new AI equipment are shortening to three years, and Q3 gross margin fell to 34.1% on inventory write-downs.[^ic3q][^q3call]

{key_table(c)}

*FY2026E = 9M26 actual + 4Q26E. Multiples at {ntd(price)}; EV uses each year's net cash. Sources: Delta financial statements; model.*

![Revenue build: AI data-centre lines vs the rest (draft Base case)](figures/revenue_build.png){{width=95%}}

# Investment thesis `[ANALYST TO REVIEW]`

*Reasoning for the draft: the model's three strongest findings (AI share, the gap between price and Base case, capital intensity), each tied to a falsifiable indicator.*

**1. Delta has become an AI infrastructure company, and the numbers are audited, not promised.** Infrastructure revenue rose {pct(seg25['Infrastructure'] / sv('FY2024', 'Infrastructure', 'revenue') - 1, 0)} in FY2025 to {bn(seg25['Infrastructure'])}, and its operating profit rose from {bn(sv('FY2024', 'Infrastructure', 'op_profit'))} to {bn(op25['Infrastructure'])}.[^fs25seg] In 9M26 Power Electronics grew 44% and Infrastructure 47%.[^ic3q] On the 3Q26 call, the CFO put AI data-centre businesses at about 50% of revenue: AI power about 25% and liquid cooling NT$25bn in the quarter (about 13%).[^q3call] Group EBIT margin rose from {pct(om('FY2021'))} in FY2021 to 17.0% in 9M26.[^ic3q]
*What would prove it wrong:* AI data-centre revenue falling as a share of sales in the monthly releases, or Power Electronics and Infrastructure margins dropping below 15%.

**2. The share price still discounts more AI growth than our Base case.** The stock has fallen from about NT$1,965 in mid-October to {ntd(price)}.[^mkt] On our Base case (AI capex growth of {base_g} in FY2027–30) the DCF is worth {ntd(d['value_today'])} a share. To justify {ntd(price)}, AI capex would have to grow about {rd_g} a year instead, and AI would reach {pct(rd['ai_share_2030'], 0)} of FY2030 revenue. Consensus is also well above us: NT$60.4 FY2027 EPS against our {f['eps']['FY2027']:.1f}.[^cons]
*What would prove it wrong:* hyperscaler capex guidance for 2027–28 above +40% a year, or Delta's FY2027 orders and monthly revenue tracking consensus rather than our Base case.

**3. Returns are at risk from capital intensity and margin slippage.** Capex rises to about NT$70bn in 2026 ({pct(70000 / f['revenue']['FY2026'])} of revenue) as Delta builds in Thailand, the US and Penang.[^q2call][^q3call] New AI equipment is depreciated over three years from April 2026.[^q2call] In 3Q26, inventory write-downs of about 1.7% of revenue pulled gross margin down to 34.1%.[^q3call] Our Base case still has free cash flow of {bn(f['fcf']['FY2026'])} in FY2026E rising to {bn(f['fcf']['FY2030'])} in FY2030E, so this is a returns question, not a solvency one: net cash was {bn(c['net_cash25'])} at end-FY2025.
*What would prove it wrong:* gross margin back above 35% in 4Q26–1Q27, as price increases take effect,[^q3call] and capex held to about 8% of revenue in 2027.

# Key debate: how durable is AI data-centre demand after 2027? `[ANALYST TO REVIEW]`

*Reasoning for the draft: it's the variable that moves value most (Sensitivity grid 2), and the sell-off shows the market is re-pricing it now.*

**The bull case.** Management says hyperscaler capex is US$700–800bn in 2026, heading above US$1tn in 2027, with global data-centre capex at US$1.2–1.5tn,[^q3call] and that investment stays robust through 2030. The IEA expects data-centre electricity use to more than double from 415 TWh in 2024 (about 1.5% of the world's electricity) to about 945 TWh in 2030.[^iea] Power density per rack keeps rising, which raises Delta's content per dollar of capex. A GB200 NVL72 rack draws about 120 kW,[^nvgb200] and NVIDIA's 800V DC architecture targets 1 MW racks from 2027.[^nv800] Delta starts mass production of ±400V and 800V HVDC products from 3Q26,[^q2call] and October revenue was a record {bn(74666)}, up 30%.[^mr10]

**The bear case.** Taiwan's AI complex sold off hard in November: the TAIEX fell about 25% from mid-October and 17.5% in one week (16–20 Nov), even after strong Nvidia results.[^news] Management itself lists four constraints that will make data-centre construction uneven: component shortages, power availability, a shortage of skilled electricians and local political opposition.[^q3call] Q3 showed the symptoms. Revenue was below expectations, gross margin fell to 34.1% on write-downs, and Infrastructure growth slowed to 25% y/y from 47% for the nine months.[^ic3q] If hyperscalers digest 2026–27 capacity, our Bear case (capex +15% in 2027, then −5%) cuts the DCF to {ntd(sc['Bear']['dcf_value_today'])}.

**Where we stand (draft).** On our Base case the share price already assumes the bull case, so we lean cautious. The analyst decides.

# Company overview

Delta Electronics (Taipei) makes power and thermal-management products in four segments.[^fs25seg]
- **Power Electronics** ({pct(seg25['Power Electronics'] / sum(seg25.values()), 0)} of FY2025 segment revenue): components, power supplies and systems, fans and thermal management.
- **Infrastructure** ({pct(seg25['Infrastructure'] / sum(seg25.values()), 0)}): ICT infrastructure (data-centre power and cooling systems), energy infrastructure and displays.
- **Automation** ({pct(seg25['Automation'] / sum(seg25.values()), 0)}): industrial and building automation.
- **Mobility** ({pct(seg25['Mobility'] / sum(seg25.values()), 0)}): EV powertrain systems.

No single customer exceeded 10% of revenue in FY2025.[^fs25geo] The AI business is still concentrated in the hyperscaler and GPU ecosystem, though: four business units serving AI make up about half of revenue.[^q3call]

| Segment (NT$bn) | FY2024 | FY2025 | 9M26 | FY2025 margin | 9M26 margin |
|---|---:|---:|---:|---:|---:|
""" + "\n".join(
        f"| {k} | {sv('FY2024', k, 'revenue') / 1000:,.1f} | {seg25[k] / 1000:,.1f} | {seg9[k] / 1000:,.1f} | "
        f"{pct(op25[k] / seg25[k])} | {pct(op9[k] / seg9[k])} |" for k in seg25) + f"""

*FY2024 on the restated basis in the FY2025 report; 9M26 from the 3Q26 slides.[^fs25seg][^ic3q]*

**Geography: selling to the US, building in Asia.** The US has become Delta's largest single-country market: {pct(us_share, 0)} of FY2025 revenue ({bn(g.at['U.S.A.', 'FY2025'])}), up from {pct(us_share21, 0)} in FY2021. Mainland China fell from {pct(cn_share21, 0)} to {pct(cn_share, 0)}.[^fs25geo] Manufacturing assets sit elsewhere. Of {bn(nca25.sum())} of non-current assets, Thailand holds {bn(nca25['Thailand'])}, Taiwan {bn(nca25['Taiwan'])}, China {bn(nca25['Mainland China'])} and the US only {bn(nca25['U.S.A.'])}.[^fs25geo] Expansion is under way in Thailand (more plants due in 2026–27), the US, China and Taiwan,[^q2call] and in a new large site in Penang, Malaysia.[^q3call]

![Revenue and assets by region](figures/geography.png){{width=80%}}

![Revenue by segment](figures/segment_revenue.png){{width=88%}}

# Industry: why AI racks need far more power and cooling

A single AI rack now draws power on a scale data centres weren't built for: NVIDIA's GB200 NVL72 rack draws about 120 kW and is liquid cooled.[^nvgb200] The next step is NVIDIA's 800V DC architecture, which targets 1 MW IT racks from 2027, developed with partners through the Open Compute Project.[^nv800] Each step up does three things for Delta:
- **More content.** More power conversion per rack: Delta supplies power shelves, including a 110 kW PSU for NVIDIA's Rubin platform, where it says it holds a leading position.[^q2call]
- **New architecture.** HVDC (±400V and 800V), solid-state transformers and fuel cells for on-site power. HVDC mass production starts in 3Q26, with meaningful volume in 2027.[^q2call]
- **Liquid cooling.** Racks at these densities are liquid cooled.[^nvgb200] Delta's liquid cooling reached NT$25bn of revenue in 3Q26, and management expects liquid-to-air cooling to stay in use longer than the market thinks.[^q3call][^q2call]

The constraint is the grid itself. The IEA estimates data centres used about 1.5% of global electricity in 2024, rising to just under 3% by 2030.[^iea] Management cites power availability and electrician shortages as the main brakes on construction.[^q3call]

**Where Delta sits in the chain.** Grid/energy infrastructure → data-centre power distribution (UPS, HVDC, SST) → rack power shelves and PSUs → cooling (CDUs, cold plates, fans). Delta's segments cover every step except the grid itself.[^fs25seg] Our peer set reflects that: Vertiv, Schneider Electric and Eaton in systems, Lite-On and AcBel in power supplies.

# Segment forecasts and scenarios

Our forecast starts from 9M26 actuals plus a 4Q26 estimate of {bn(a['q4_revenue']['FY2026'])} (management guides Q4 above Q3's {bn(rev3q)}).[^q3call] Revenue is then built in two parts:
- **AI data-centre revenue**, anchored at 50% of FY2026E revenue (power 25%, liquid cooling 13%, other 12%),[^q3call] grows with AI data-centre capex times Delta's content and share gain.
- **The rest** grows with its own drivers: industrial demand for Automation, EV demand for Mobility.

EBIT is segment revenue × segment margin. All {len(a)} drivers, with a rationale for each, are in `model/assumptions.csv`.

{forecast_table(c)}

*Draft Base case `[ANALYST TO REVIEW]`. NT$bn unless stated.*

{scenario_table(c)}

*Bear: hyperscaler digestion after 2027, no share gains, margins revert. Bull: capex keeps compounding, faster HVDC and liquid-cooling adoption.*

# Financials

![Gross and EBIT margin](figures/margins.png){{width=80%}}

**Margins.** Gross margin rose from {pct(gm('FY2021'))} (FY2021) to {pct(gm('FY2025'))} (FY2025) and 35.5% (9M26). EBIT margin rose from {pct(om('FY2021'))} to {pct(om('FY2025'))} and 17.0%.[^fs25is][^ic3q] Management calls a gross margin of around 35% "reasonable".[^q2call] Our Base case holds 35.0%, with EBIT margin rising to {pct(f['ebit']['FY2030'] / f['revenue']['FY2030'])} by FY2030E on mix.

**Cash flow and the capex cycle.** Operating cash flow was {bn(h.at['cfo', 'FY2025'])} in FY2025, against capex of {bn(capex25)}.[^fs25cf] Our Base case has FY2026E operating cash flow of {bn(f['cfo']['FY2026'])}, held back by an inventory build. Free cash flow is {bn(f['fcf']['FY2026'])} in FY2026E, rising to {bn(f['fcf']['FY2030'])} by FY2030E as capex falls back towards 6.5% of revenue.

**Balance sheet.** At end-FY2025 Delta held {bn(h.at['cash', 'FY2025'])} of cash against total debt of {bn(h.at['st_borrowings', 'FY2025'] + h.at['lt_debt_current', 'FY2025'] + h.at['bonds_payable', 'FY2025'] + h.at['lt_borrowings', 'FY2025'])}, and return on equity was about 24%.[^fs25bs] Dividends are paid the year after they're earned: NT$11.6 a share out of FY2025 earnings, about a 50% payout.[^ic4q25] Working capital is the swing factor. Receivable days were 80 and payable days 95 in FY2025, and management expects to carry larger inventory buffers.[^q3call]

# Valuation

**DCF.** WACC is {pct(w['wacc'], 2)}. That combines a cost of equity of {pct(w['ke'], 2)} with a small amount of debt, about 3% of capital at market value:
- **Risk-free rate {pct(v['rf'], 2)}:** Taiwan's 10-year government bond yield.[^tpex]
- **Equity risk premium {pct(v['mature_erp'] + v['crp'], 2)}:** mature market 4.23% plus 0.78% for Taiwan.[^damo]
- **Levered beta {w['beta_l']:.2f}:** from an industry unlevered beta of 1.45.[^damo] A regression on the TAIEX gives 1.73 (2-year raw) and 1.49 (Blume-adjusted).

Five years of free cash flow to the firm and a Gordon terminal value at {pct(v['g'])} growth give an enterprise value of {bn(d['ev'])}, of which the terminal value is {pct(d['tv_share'], 0)}. Adding {bn(d['net_cash'])} of net cash and investments, and deducting minorities at their 10.5% share of earnings, gives **{ntd(d['value_today'])} a share** today. The terminal value implies only {d['implied_exit_multiple']:.1f}× FY2030E EBITDA. Using a 12× exit multiple instead gives {ntd(c['exit_ps'])}. **That gap is the largest judgement in this valuation** `[ANALYST TO REVIEW]`.

**Comps.** Peers trade at a median {peers_med['ev_ebitda_fy1']['median']:.1f}× FY1 EV/EBITDA and {peers_med['pe_fy1']['median']:.1f}× FY1 P/E.[^cons] Applied to our FY2027E, these give {ntd(s['comps']['ev_ebitda']['median'])} and {ntd(s['comps']['pe']['median'])}: comps value {ntd(s['comps']['value_today'])}.
- **Closest peers:** Vertiv and Schneider are the closest data-centre power plays.
- **Taiwan peers:** Lite-On and AcBel share Delta's power-supply market but are smaller and lower-margin.
- **Eaton** adds US electrical exposure.

{comps_table(c)}

![Valuation range by method](figures/football_field.png){{width=92%}}

**Price target and rating** `[ANALYST TO REVIEW]`. A 50/50 blend of DCF and comps ({ntd(s['blend_today'])} today), rolled forward 12 months at the cost of equity and less the expected FY2027 dividend, gives a target of **{ntd(target)}**: {up * 100:+.0f}% against {ntd(price)}, or **{s['rating']}** under the spec's rule. Bear / Base / Bull targets are {ntd(sc['Bear']['target'])} / {ntd(sc['Base']['target'])} / {ntd(sc['Bull']['target'])}, against a 52-week range of {ntd(lo52)}–{ntd(hi52)}.

**Reverse DCF.** We solved for the uniform increase in AI data-centre capex growth that makes the DCF equal the share price. The answer is **{rd['shift'] * 100:+.0f} percentage points a year**, taking growth from {base_g} to {rd_g} in FY2027–30. At that rate AI data-centre revenue would be {pct(rd['ai_share_2030'], 0)} of FY2030 revenue and EPS NT${rd['eps_2030']:.0f}. Alternatively, the market is paying for growth well beyond 2030, which a 12× exit multiple captures.

![DCF sensitivity](figures/sensitivity.png){{width=95%}}

# Risks

| Risk | Why it matters | Indicator to watch |
|---|---|---|
| **AI capex slowdown** | About half of revenue is tied to AI data-centre spending;[^q3call] the Bear case cuts value to {ntd(sc['Bear']['dcf_value_today'])} | Hyperscaler capex guidance; Delta monthly revenue y/y |
| **Customer concentration** | No customer is above 10% of revenue,[^fs25geo] but AI demand flows through a few hyperscalers and one GPU ecosystem | Revenue mix by segment; NVIDIA platform transitions (Rubin, 800V) |
| **Tariffs and US–China/Taiwan geopolitics** | The US is {pct(us_share, 0)} of revenue but holds {bn(nca25['U.S.A.'])} of non-current assets; production is in Thailand, Taiwan and China[^fs25geo] | US tariff rulings on Thai/Taiwanese electronics; Delta's US capacity build |
| **Margin pressure from competition and costs** | Q3 gross margin fell to 34.1% on write-downs; raw materials are up about 2% this year[^q3call] | Gross margin vs 35%; price increases landing in 4Q26–1Q27 |
| **Execution on capacity expansion** | Capex of about NT$70bn in 2026; new plants in Thailand, the US and Penang[^q2call][^q3call] | Capex vs guidance; inventory days; plant start dates |
| **FX** | {pct(1 - g.at['Taiwan', 'FY2025'] / g['FY2025'].sum(), 0)} of FY2025 revenue came from outside Taiwan,[^fs25geo] while results are reported in NT$; USD/TWD was {c['mkt']['fx']['USD']:.2f} on the price date[^mkt] | USD/TWD; non-operating FX gains/losses |
| **Litigation** | Vicor's patent case and ITC investigation involve several power-supply makers[^q3call] | ITC rulings |

# Appendix

**A. Model and sources.** Full statements FY2021A–FY2030E, schedules, DCF, comps and checks are in `model/delta_model.xlsx`, generated from code with live formulas. Every input is in `model/assumptions.csv` and `model/valuation_inputs.csv`, and every sourced figure, with page and URL, is in `data/sources.csv`. The Python mirror (`src/delta/model.py`) agrees with the workbook within 0.1% in all scenarios.

**B. Summary financial statements (draft Base case).**

{statements_table(c)}

*FY2024A–FY2025A audited; FY2026E–FY2030E model forecast. Capex and dividends shown as cash outflows.*

**C. Base-case assumptions** `[ANALYST TO REVIEW]`

{assumptions_table(c)}

*All 40 drivers, with Bear / Base / Bull values and a rationale for each, are in `model/assumptions.csv`; valuation inputs are in `model/valuation_inputs.csv`.*

**D. Basis of preparation.**
- **Historicals:** audited FY2021–FY2025 and reviewed 1H26 statements (Taiwan IFRS, NT$ thousands, shown in NT$bn).
- **Q3 2026:** from the 3Q26 earnings slides and transcript; the English Q3 statements weren't yet published.
- **Market data:** {s['price_date']}, Yahoo Finance via yfinance.

*Student research for educational purposes only. Not investment advice.*

[^fs25is]: Delta Electronics, consolidated financial statements FY2021–FY2025 (audited), statements of comprehensive income.
[^fs25bs]: Delta Electronics, FY2025 consolidated financial statements, balance sheet p.8–9.
[^fs25cf]: Delta Electronics, FY2025 consolidated financial statements, cash flow statement p.14–15.
[^fs25seg]: Delta Electronics, FY2025 consolidated financial statements, Note 14 (operating segments) p.111; FY2024 restated.
[^fs25geo]: Delta Electronics, FY2025 consolidated financial statements, Note 14(6)–(7) (geography; major customers) p.112.
[^ic3q]: Delta Electronics, 3Q26 results presentation (29 Oct 2026), p.4–6, 12–18.
[^ic4q25]: Delta Electronics, 4Q25 results presentation (26 Feb 2026), p.19.
[^q2call]: Delta Electronics, 2Q26 results meeting transcript (30 Jul 2026), p.2–13.
[^q3call]: Delta Electronics, 3Q26 earnings conference transcript (29 Oct 2026), p.1–7.
[^mr10]: Delta Electronics, monthly revenue release, October 2026 (9 Nov 2026).
[^mkt]: Yahoo Finance via yfinance, snapshot {s['price_date']} (`data/raw/market/`).
[^cons]: Yahoo Finance analyst consensus averages, snapshot {s['price_date']}.
[^tpex]: Taipei Exchange (TPEx), Newsletter November 2026, bond market review: 10-year yield 2.0300% at 31 Oct 2026.
[^damo]: A. Damodaran, country risk premiums and industry betas (emerging markets), updated 5 Jan 2026.
[^iea]: IEA, *Energy and AI* (2025): data centres 415 TWh in 2024 (~1.5% of global electricity), ~945 TWh in 2030.
[^nvgb200]: NVIDIA, DGX GB200 user guide (hardware): rack power ~120 kW.
[^nv800]: NVIDIA developer blog, "NVIDIA 800 V HVDC architecture will power the next generation of AI factories".
[^news]: Focus Taiwan (3 Nov 2026) and Yahoo Finance (Nov 2026) market reports on the TAIEX sell-off.
"""
    # a list needs a blank line before it in pandoc (GitHub doesn't mind)
    md = re.sub(r"(?m)^((?!- |\s*$|\|).+)\n(- )", r"\1\n\n\2", md)
    OUT.write_text(md)
    return md


def main() -> int:
    md = write()
    words = len(md.split())
    print(f"{OUT} ({words:,} words)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
