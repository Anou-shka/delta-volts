"""Charts for the README and report, rendered in light and dark variants.

Run: python -m src.delta.charts   ->  report/figures/*.png
Palette: validated categorical slots (CVD-safe adjacent order), light and dark steps.
"""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib import font_manager  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
PROCESSED = ROOT / "data" / "processed"
FIG = ROOT / "report" / "figures"

THEMES = {
    "light": {"surface": "#fcfcfb", "text": "#0b0b0b", "text2": "#52514e", "grid": "#e4e3df",
              "series": ["#2a78d6", "#eb6834", "#1baf7a", "#eda100"]},
    "dark": {"surface": "#1a1a19", "text": "#f0efec", "text2": "#c3c2b7", "grid": "#33332f",
             "series": ["#3987e5", "#d95926", "#199e70", "#c98500"]},
}
SEGMENTS = ["Power Electronics", "Infrastructure", "Automation", "Mobility"]  # stack order = slot order
FY = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]
LTM = "LTM Jun-26"


def _font() -> None:
    names = {f.name for f in font_manager.fontManager.ttflist}
    plt.rcParams["font.family"] = "Inter" if "Inter" in names else "DejaVu Sans"


def _style(ax, t) -> None:
    ax.set_facecolor(t["surface"])
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(t["grid"])
    ax.tick_params(colors=t["text2"], length=0, labelsize=10)
    ax.yaxis.grid(True, color=t["grid"], linewidth=0.8)
    ax.set_axisbelow(True)


def segment_table() -> pd.DataFrame:
    """Segment revenue (NT$ bn) on the latest basis, FY2021-FY2025 plus LTM to June 2026."""
    s = pd.read_csv(PROCESSED / "segments.csv")
    s = s[s.latest & (s["item"] == "revenue") & s.segment.isin(SEGMENTS)]
    p = s.pivot_table(index="period", columns="segment", values="value") / 1e6
    p = p.reindex(columns=SEGMENTS).fillna(0.0)
    p.loc[LTM] = p.loc["FY2025"] - p.loc["1H25"] + p.loc["1H26"]
    return p.loc[FY + [LTM]]


def margin_table() -> pd.DataFrame:
    h = pd.read_csv(PROCESSED / "historicals.csv", index_col=0)
    rows = {}
    for p in FY:
        rows[p] = h[p]
    rows[LTM] = h["FY2025"] - h["1H25"] + h["1H26"]
    d = pd.DataFrame(rows)
    return pd.DataFrame({"Gross margin": d.loc["gross_profit"] / d.loc["revenue"] * 100,
                         "EBIT margin": d.loc["ebit"] / d.loc["revenue"] * 100})


def chart_segments(mode: str) -> Path:
    t = THEMES[mode]
    data = segment_table()
    fig, ax = plt.subplots(figsize=(8, 4.6), dpi=200)
    fig.patch.set_facecolor(t["surface"])
    _style(ax, t)
    x = range(len(data))
    bottom = [0.0] * len(data)
    for seg, color in zip(SEGMENTS, t["series"]):
        vals = data[seg].tolist()
        ax.bar(x, vals, bottom=bottom, width=0.62, color=color, edgecolor=t["surface"], linewidth=1.5,
               label=seg)
        bottom = [b + v for b, v in zip(bottom, vals)]
    for i, total in enumerate(bottom):
        ax.text(i, total + 8, f"{total:,.0f}", ha="center", va="bottom", fontsize=10, color=t["text"],
                fontweight="semibold")
    # direct labels on the right-most bar (relief for low-contrast fills)
    last = data.iloc[-1]
    y = 0.0
    for seg in SEGMENTS:
        v = last[seg]
        if v > 25:
            ax.text(len(data) - 1 + 0.36, y + v / 2, f"{seg}  {v:,.0f}", va="center", ha="left",
                    fontsize=9, color=t["text2"])
        y += v
    ax.set_xticks(list(x), [c.replace("FY", "FY ") for c in data.index])
    ax.set_xlim(-0.6, len(data) - 1 + 1.9)
    ax.set_ylim(0, max(bottom) * 1.16)
    ax.set_ylabel("NT$ billion", color=t["text2"], fontsize=10)
    ax.legend(loc="upper left", frameon=False, ncol=4, fontsize=9, labelcolor=t["text2"],
              bbox_to_anchor=(0, 1.02), handlelength=1, handleheight=1)
    fig.suptitle("Revenue by segment", x=0.02, ha="left", fontsize=13, color=t["text"], fontweight="semibold")
    fig.text(0.02, 0.005, "Source: Delta consolidated financial statements, segment note. Excludes unallocated 'Others' "
             "(<0.2% of revenue). FY2021-22: Mobility inside Power Electronics.", fontsize=7.5, color=t["text2"])
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    out = FIG / f"segment_revenue{'' if mode == 'light' else '_dark'}.png"
    fig.savefig(out, facecolor=t["surface"])
    plt.close(fig)
    return out


def chart_margins(mode: str) -> Path:
    t = THEMES[mode]
    m = margin_table()
    fig, ax = plt.subplots(figsize=(8, 4.6), dpi=200)
    fig.patch.set_facecolor(t["surface"])
    _style(ax, t)
    x = list(range(len(m)))
    for col, color in zip(m.columns, t["series"]):
        ax.plot(x, m[col], color=color, linewidth=2, marker="o", markersize=6,
                markeredgecolor=t["surface"], markeredgewidth=1.5, label=col)
        ax.text(x[-1] + 0.15, m[col].iloc[-1], f"{col}  {m[col].iloc[-1]:.1f}%", va="center",
                fontsize=9, color=t["text2"])
        ax.text(x[0] - 0.15, m[col].iloc[0], f"{m[col].iloc[0]:.1f}%", va="center", ha="right",
                fontsize=9, color=t["text2"])
    ax.set_xticks(x, [c.replace("FY", "FY ") for c in m.index])
    ax.set_xlim(-0.8, len(m) - 1 + 1.9)
    ax.set_ylim(0, 40)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax.legend(loc="upper left", frameon=False, ncol=2, fontsize=9, labelcolor=t["text2"],
              bbox_to_anchor=(0, 1.02))
    fig.suptitle("Gross and EBIT margin stepped up from FY2024", x=0.02, ha="left", fontsize=13,
                 color=t["text"], fontweight="semibold")
    fig.text(0.02, 0.005, "Source: Delta consolidated financial statements. EBIT = operating profit. "
             "LTM = FY2025 - 1H25 + 1H26.", fontsize=7.5, color=t["text2"])
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    out = FIG / f"margins{'' if mode == 'light' else '_dark'}.png"
    fig.savefig(out, facecolor=t["surface"])
    plt.close(fig)
    return out


MODEL = ROOT / "model" / "delta_model.xlsx"
HISTORY_GREY = {"light": "#a8a69f", "dark": "#77756e"}


def forecast_table() -> pd.DataFrame:
    """Base-case revenue build (NT$ bn) from the recalculated model: history total, forecast split AI / non-AI."""
    from openpyxl import load_workbook
    ws = load_workbook(MODEL, data_only=True)["Segments"]
    rows: dict[str, int] = {}
    for r in ws.iter_rows():
        if isinstance(r[0].value, str):
            rows.setdefault(r[0].value, r[0].row)  # first occurrence: the build, not the disclosures table
    cols = {"FY2024": "F", "FY2025": "G", "FY2026": "H", "FY2027": "I", "FY2028": "J", "FY2029": "K", "FY2030": "L"}
    lines = {"Non-AI revenue": "Group revenue excl. AI data-centre", "Other AI data-centre": "Other AI data-centre",
             "Liquid cooling": "Liquid cooling", "AI power solutions": "AI power solutions",
             "Group revenue": "Group revenue"}
    out = {k: [ws[f"{c}{rows[v]}"].value for c in cols.values()] for k, v in lines.items()}
    df = pd.DataFrame(out, index=list(cols)) / 1000
    return df


def chart_forecast(mode: str) -> Path:
    t = THEMES[mode]
    df = forecast_table()
    fig, ax = plt.subplots(figsize=(8, 4.6), dpi=200)
    fig.patch.set_facecolor(t["surface"])
    _style(ax, t)
    stack = ["Non-AI revenue", "Other AI data-centre", "Liquid cooling", "AI power solutions"]
    colors = dict(zip(stack, [t["series"][3], t["series"][2], t["series"][1], t["series"][0]]))
    x = list(range(len(df)))
    for i, p in enumerate(df.index):
        if pd.isna(df.at[p, "AI power solutions"]):
            ax.bar(i, df.at[p, "Group revenue"], width=0.62, color=HISTORY_GREY[mode], edgecolor=t["surface"],
                   linewidth=1.5, label="Reported (AI not disclosed)" if i == 0 else None)
            continue
        bottom = 0.0
        for k in stack:
            ax.bar(i, df.at[p, k], bottom=bottom, width=0.62, color=colors[k], edgecolor=t["surface"],
                   linewidth=1.5, label=k if p == "FY2026" else None)
            bottom += df.at[p, k]
    for i, p in enumerate(df.index):
        ax.text(i, df.at[p, "Group revenue"] + 15, f"{df.at[p, 'Group revenue']:,.0f}", ha="center", va="bottom",
                fontsize=10, color=t["text"], fontweight="semibold")
    last, y = df.iloc[-1], 0.0
    for k in stack:
        ax.text(len(df) - 1 + 0.36, y + last[k] / 2, f"{k}  {last[k]:,.0f}", va="center", ha="left", fontsize=9,
                color=t["text2"])
        y += last[k]
    ax.axvline(1.5, color=t["grid"], linewidth=1, linestyle=(0, (3, 3)))
    ax.text(1.62, df["Group revenue"].max() * 0.98, "Forecast  ->", fontsize=9, color=t["text2"])
    ax.set_xticks(x, [p.replace("FY", "FY ") + ("A" if i < 2 else "E") for i, p in enumerate(df.index)])
    ax.set_xlim(-0.6, len(df) - 1 + 2.3)
    ax.set_ylim(0, df["Group revenue"].max() * 1.28)
    ax.set_ylabel("NT$ billion", color=t["text2"], fontsize=10)
    ax.legend(loc="upper left", frameon=False, ncol=3, fontsize=8.5, labelcolor=t["text2"], bbox_to_anchor=(0, 1.02),
              handlelength=1, handleheight=1)
    fig.suptitle("Revenue build: AI data-centre lines vs the rest", x=0.02, ha="left", fontsize=13, color=t["text"],
                 fontweight="semibold")
    fig.text(0.02, 0.005, "Draft Base case [ANALYST TO REVIEW]. FY2026E = 9M26 actual + 4Q26E; AI split anchored on "
             "the 3Q26 call. AI revenue not reported before 2026.", fontsize=7.5, color=t["text2"])
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    out = FIG / f"revenue_build{'' if mode == 'light' else '_dark'}.png"
    fig.savefig(out, facecolor=t["surface"])
    plt.close(fig)
    return out


SUMMARY = ROOT / "model" / "valuation_summary.json"
DIVERGING = {"light": ("#e34948", "#f0efec", "#2a78d6"), "dark": ("#e66767", "#383835", "#3987e5")}


def _summary() -> dict:
    import json
    return json.loads(SUMMARY.read_text())


def chart_football(mode: str) -> Path:
    """Horizontal range bars: value per share by method, against the share price and the target."""
    t = THEMES[mode]
    s = _summary()
    c = s["comps"]
    wg = [v for row in s["sensitivity_wacc_g"]["values"] for v in row]
    ag = [v for row in s["sensitivity_ai_gm"]["values"] for v in row]
    rows = [("52-week trading range", s["low_52w"], s["high_52w"]),
            ("Comps: EV/EBITDA FY1 (peer IQR)", c["ev_ebitda"]["q1"], c["ev_ebitda"]["q3"]),
            ("Comps: P/E FY1 (peer IQR)", c["pe"]["q1"], c["pe"]["q3"]),
            ("DCF: WACC +/-1pp, growth +/-1pp", min(wg), max(wg)),
            ("DCF: AI growth +/-20pp, margin +/-2pp", min(ag), max(ag)),
            ("DCF: Bear to Bull scenario", s["scenarios"]["Bear"]["dcf_value_today"],
             s["scenarios"]["Bull"]["dcf_value_today"])]
    fig, ax = plt.subplots(figsize=(8, 4.4), dpi=200)
    fig.patch.set_facecolor(t["surface"])
    _style(ax, t)
    ax.yaxis.grid(False)
    ax.xaxis.grid(True, color=t["grid"], linewidth=0.8)
    for i, (label, lo, hi) in enumerate(rows[::-1]):
        ax.barh(i, hi - lo, left=lo, height=0.5, color=t["series"][0], edgecolor=t["surface"], linewidth=1.5)
        ax.text(lo - 12, i, f"{lo:,.0f}", va="center", ha="right", fontsize=8.5, color=t["text2"])
        ax.text(hi + 12, i, f"{hi:,.0f}", va="center", ha="left", fontsize=8.5, color=t["text2"])
    ax.set_yticks(range(len(rows)), [r[0] for r in rows[::-1]])
    above_target = s["price"] >= s["target"]
    for x, label, style, ha in ((s["price"], f"Share price {s['price']:,.0f}", "-", "left" if above_target else "right"),
                                (s["target"], f"Target {s['target']:,.0f} (draft)", (0, (3, 3)),
                                 "right" if above_target else "left")):
        ax.axvline(x, color=t["text"], linewidth=1.2, linestyle=style)
        ax.text(x + (12 if ha == "left" else -12), len(rows) - 0.35, label, ha=ha, va="bottom", fontsize=9,
                color=t["text"])
    hi_all = max(r[2] for r in rows)
    ax.set_xlim(0, hi_all * 1.1)
    ax.set_ylim(-0.6, len(rows) - 0.1)
    ax.set_xlabel("Value per share, NT$", color=t["text2"], fontsize=10)
    fig.suptitle("Valuation range by method", x=0.02, ha="left", fontsize=13, color=t["text"], fontweight="semibold")
    fig.text(0.02, 0.005, f"Price {s['price_date']}. Draft Base case [ANALYST TO REVIEW]. DCF values rolled to the "
             "price date; comps on FY2027E.", fontsize=7.5, color=t["text2"])
    fig.tight_layout(rect=(0, 0.03, 1, 0.97))
    out = FIG / f"football_field{'' if mode == 'light' else '_dark'}.png"
    fig.savefig(out, facecolor=t["surface"])
    plt.close(fig)
    return out


def chart_sensitivity(mode: str) -> Path:
    """Two 5x5 value grids, coloured by distance from the share price (diverging around it)."""
    import numpy as np
    from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
    t = THEMES[mode]
    s = _summary()
    lo_c, mid_c, hi_c = DIVERGING[mode]
    cmap = LinearSegmentedColormap.from_list("div", [lo_c, mid_c, hi_c])
    grids = [(np.array(s["sensitivity_wacc_g"]["values"]), "WACC x terminal growth",
              [f"{x:.1%}" for x in s["sensitivity_wacc_g"]["g"]], [f"{x:.2%}" for x in s["sensitivity_wacc_g"]["wacc"]],
              "Terminal growth", "WACC"),
             (np.array(s["sensitivity_ai_gm"]["values"]), "AI capex growth x gross margin",
              [f"{x:+.0%}" for x in s["sensitivity_ai_gm"]["ai_shifts"]],
              [f"{x:+.0%}" for x in s["sensitivity_ai_gm"]["gm_shifts"]], "AI capex growth shift (pp p.a.)",
              "Margin shift (pp)")]
    allv = np.concatenate([g[0].ravel() for g in grids])
    span = max(abs(allv.max() - s["price"]), abs(allv.min() - s["price"]))
    norm = TwoSlopeNorm(vmin=s["price"] - span, vcenter=s["price"], vmax=s["price"] + span)
    fig, axes = plt.subplots(1, 2, figsize=(9, 4.3), dpi=200)
    fig.patch.set_facecolor(t["surface"])
    for ax, (v, title, xt, yt, xl, yl) in zip(axes, grids):
        ax.set_facecolor(t["surface"])
        ax.imshow(v, cmap=cmap, norm=norm, aspect="auto")
        for (i, j), val in np.ndenumerate(v):
            ax.text(j, i, f"{val:,.0f}", ha="center", va="center", fontsize=9,
                    color=t["text"], fontweight="semibold" if (i, j) == (2, 2) else "normal")
        ax.set_xticks(range(5), xt, fontsize=8.5, color=t["text2"])
        ax.set_yticks(range(5), yt, fontsize=8.5, color=t["text2"])
        ax.set_xlabel(xl, color=t["text2"], fontsize=9)
        ax.set_ylabel(yl, color=t["text2"], fontsize=9)
        ax.set_title(title, fontsize=10.5, color=t["text"], loc="left")
        for sp in ax.spines.values():
            sp.set_visible(False)
        ax.tick_params(length=0)
        ax.set_xticks(np.arange(-0.5, 5), minor=True)
        ax.set_yticks(np.arange(-0.5, 5), minor=True)
        ax.grid(which="minor", color=t["surface"], linewidth=2)
    fig.suptitle(f"DCF value per share (NT$) vs share price {s['price']:,.0f}: red below, blue above", x=0.02,
                 ha="left", fontsize=12.5, color=t["text"], fontweight="semibold")
    fig.text(0.02, 0.01, "Draft Base case [ANALYST TO REVIEW]; centre cell = Base. Right grid from the Python mirror "
             "(full model re-run per cell).", fontsize=7.5, color=t["text2"])
    fig.tight_layout(rect=(0, 0.04, 1, 0.94))
    out = FIG / f"sensitivity{'' if mode == 'light' else '_dark'}.png"
    fig.savefig(out, facecolor=t["surface"])
    plt.close(fig)
    return out


def main() -> int:
    FIG.mkdir(parents=True, exist_ok=True)
    _font()
    for mode in THEMES:
        print(chart_segments(mode))
        print(chart_margins(mode))
        if MODEL.exists():
            print(chart_forecast(mode))
        if SUMMARY.exists():
            print(chart_football(mode))
            print(chart_sensitivity(mode))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
