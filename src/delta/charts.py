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


def main() -> int:
    FIG.mkdir(parents=True, exist_ok=True)
    _font()
    for mode in THEMES:
        print(chart_segments(mode))
        print(chart_margins(mode))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
