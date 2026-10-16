"""Extract historical financials from Delta's consolidated financial reports.

Outputs (data/processed/):
  fs_lines.csv     every as-reported statement line, every period column (NT$ thousands)
  segments.csv     segment revenue and operating profit from the segment note (NT$ thousands)
  historicals.csv  standardised model lines by period (NT$ millions)
  checks.csv       tie-out checks (subtotals, statement links, restatements)
and appends one row per standardised figure to data/sources.csv.

Run: python -m src.delta.extract
"""
from __future__ import annotations

import csv
import re
import warnings
from collections import defaultdict
from pathlib import Path

import pandas as pd
import pymupdf

from .pdftable import _lines, parse_page

warnings.filterwarnings("ignore", message="This pattern is interpreted as a regular expression")

ROOT = Path(__file__).resolve().parents[2]
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed"

TITLES = {
    "BS": "CONSOLIDATED BALANCE SHEETS",
    "IS": "CONSOLIDATED STATEMENTS OF COMPREHENSIVE INCOME",
    "CF": "CONSOLIDATED STATEMENTS OF CASH FLOWS",
}

# Column index (0 = US$ convenience translation, ignored) -> period key, per document and statement.
# Balance-sheet period keys are the period whose closing date the column shows.
DOCS: dict[str, dict[str, list[str | None]]] = {}
for y in range(2021, 2026):
    DOCS[f"fs_{y}_annual_consolidated.pdf"] = {s: [None, f"FY{y}", f"FY{y - 1}"] for s in TITLES}
DOCS["fs_2026q1_consolidated.pdf"] = {"BS": [None, "1Q26", "FY2025", "1Q25"],
                                      "IS": [None, "1Q26", "1Q25"], "CF": [None, "1Q26", "1Q25"]}
DOCS["fs_2026q2_consolidated.pdf"] = {"BS": [None, "1H26", "FY2025", "1H25"],
                                      "IS": [None, "2Q26", "2Q25", "1H26", "1H25"],
                                      "CF": [None, "1H26", "1H25"]}

# Which document is the primary source for each period (its own report, or the
# latest one that presents it). Comparatives elsewhere are used for restatement checks.
PRIMARY = {f"FY{y}": f"fs_{y}_annual_consolidated.pdf" for y in range(2021, 2026)}
PRIMARY.update({"1Q26": "fs_2026q1_consolidated.pdf", "1Q25": "fs_2026q1_consolidated.pdf",
                "2Q26": "fs_2026q2_consolidated.pdf", "2Q25": "fs_2026q2_consolidated.pdf",
                "1H26": "fs_2026q2_consolidated.pdf", "1H25": "fs_2026q2_consolidated.pdf"})
PERIOD_ORDER = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "1Q25", "2Q25", "1H25", "1Q26", "2Q26", "1H26"]
BS_PERIODS = {"FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "1Q25", "1H25", "1Q26", "1H26"}


def statement_pages(doc: pymupdf.Document) -> dict[str, list[int]]:
    pages: dict[str, list[int]] = defaultdict(list)
    for i in range(min(25, doc.page_count)):
        head = doc[i].get_text()[:400]
        for k, title in TITLES.items():
            if title in head:
                pages[k].append(i)
    return pages


def extract_lines() -> pd.DataFrame:
    recs = []
    for fname, cols in DOCS.items():
        doc = pymupdf.open(RAW / fname)
        for stmt, pages in statement_pages(doc).items():
            order = 0
            for pno in pages:
                rows, ncol = parse_page(doc[pno])
                if ncol != len(cols[stmt]):
                    raise ValueError(f"{fname} p{pno + 1} {stmt}: {ncol} columns, expected {len(cols[stmt])}")
                for r in rows:
                    order += 1
                    for ci, period in enumerate(cols[stmt]):
                        if period is None or ci not in r.values:
                            continue
                        recs.append({"document": fname, "statement": stmt, "page": pno + 1, "order": order,
                                     "label": r.label, "period": period, "value": r.values[ci]})
    return pd.DataFrame(recs)


# --------------------------------------------------------------------------- segments
SEG_NAMES = {5: ["Power Electronics", "Mobility", "Automation", "Infrastructure", "Total"],
             4: ["Power Electronics", "Automation", "Infrastructure", "Total"]}
SEG_BASIS = {5: "4 segments (Mobility separate; FY2024 report onward)",
             4: "3 segments (Mobility inside Power Electronics; FY2023 report and earlier)"}
SEG_PAGES_HINT = "OPERATING SEGMENT INFORMATION"


def _num(tok: str) -> float | None:
    t = tok.replace("$", "").replace(",", "").strip()
    neg = t.endswith(")") or t.startswith("(")
    t = t.strip("()")
    if not re.fullmatch(r"\d+(\.\d+)?", t):
        return None
    return -float(t) if neg else float(t)


def extract_segments(fname: str, ebit: dict[str, float]) -> list[dict]:
    """Parse 'Revenue from external customers' and 'Segment income' rows of the segment note.

    Each block of the table holds five figures (four segments + total) for one
    period. A block's period is identified by matching its total segment profit
    to the operating profit reported on the income statement (ebit, NT$ thousands).
    """
    doc = pymupdf.open(RAW / fname)
    start = next(i for i in range(doc.page_count) if SEG_PAGES_HINT in doc[i].get_text())
    blocks: list[dict[str, list[float]]] = []
    for pno in range(start, min(start + 3, doc.page_count)):
        words = doc[pno].get_text("words")
        for line in _lines(words, tol=3):
            text = " ".join(w[4] for w in line)
            kind = ("revenue" if text.startswith("Revenue from external")
                    else "op_profit" if re.match(r"Segment (income|profit)", text) else None)
            if not kind:
                continue
            ys = line[0][1]
            cand = [w for w in words if ys - 3 <= w[1] <= ys + 24 and w[0] > 180 and _num(w[4]) is not None]
            rows: dict[int, list] = defaultdict(list)
            for w in cand:
                rows[round(w[1])].append(w)
            for y in sorted(rows):
                if len(rows[y]) in (4, 5):
                    vals = [_num(w[4]) for w in sorted(rows[y], key=lambda w: w[0])]
                    if kind == "revenue":
                        blocks.append({"revenue": vals})
                    elif blocks and "op_profit" not in blocks[-1]:
                        blocks[-1]["op_profit"] = vals
                    break
    out = []
    for blk in blocks:
        if "op_profit" not in blk:
            continue
        period = next((p for p, v in ebit.items() if abs(v - blk["op_profit"][-1]) <= 2), None)
        if period is None:
            raise ValueError(f"{fname}: segment block with total profit {blk['op_profit'][-1]} matches no period")
        for kind, vals in blk.items():
            n = len(vals)
            for name, v in zip(SEG_NAMES[n], vals):
                out.append({"document": fname, "period": period, "basis": SEG_BASIS[n], "segment": name,
                            "item": kind, "value": v})
    return out


# --------------------------------------------------------------------------- monthly revenue
def extract_monthly() -> pd.DataFrame:
    """Monthly revenue releases (unaudited): revenue, year-to-date revenue and segment mix (%)."""
    recs = []
    for f in sorted(RAW.glob("mr_*_EN.pdf")):
        text = " ".join(pymupdf.open(f)[0].get_text().split())
        ym = f.stem.split("_")[1]
        m = re.search(r"revenues for (\w+ \d{4}) totaled NT\$ ?([\d,]+) million", text)
        ytd = re.search(r"from January to \w+ in \d{4} were NT\$ ?([\d,]+) million", text)
        rec = {"month": f"{ym[:4]}-{ym[4:]}", "document": f.name,
               "revenue_ntd_m": float(m.group(2).replace(",", "")),
               "ytd_revenue_ntd_m": float(ytd.group(1).replace(",", "")) if ytd else float("nan")}
        for seg in ("Power Electronics", "Mobility", "Automation", "Infrastructure", "Others"):
            sm = re.search(seg + r"\s+(\d+)%", text)
            rec[f"mix_{seg.lower().replace(' ', '_')}_pct"] = float(sm.group(1)) if sm else float("nan")
        recs.append(rec)
    return pd.DataFrame(recs)


# --------------------------------------------------------------------------- standardisation
def _find(df: pd.DataFrame, stmt: str, pattern: str, nth: int = 0) -> pd.Series:
    """Values by period for the nth line (in statement order) whose label matches pattern."""
    sub = df[(df.statement == stmt) & df.label.str.contains(pattern, case=False, regex=True)]
    out = {}
    for period, g in sub.groupby("period"):
        g = g.sort_values("order")
        if len(g) > nth:
            out[period] = g.iloc[nth]["value"]
    return pd.Series(out, dtype=float)


IS_MAP = {
    "revenue": r"^operating revenue$|(^| )operating revenue$",
    "cogs": r"operating costs$",
    "gross_profit": r"^gross profit",
    "selling": r"selling expenses$",
    "g_and_a": r"general and administrative expenses$",
    "r_and_d": r"research and development expenses$",
    "credit_loss": r"expected credit impairment (loss|gain|\(loss\) gain|gain \(loss\))",
    "opex_total": r"^total operating expenses$",
    "ebit": r"^operating profit$|^operating income$",
    "interest_income": r"interest income$",
    "other_income": r"other income$",
    "other_gains_losses": r"other gains and losses$",
    "finance_costs": r"finance costs$",
    "equity_method_income": r"^share of (profit|loss|\(loss\)|profit \(loss\)|\(loss\) profit).*equity method$",
    "non_operating_total": r"^total non-operating income and expenses$",
    "pbt": r"^profit before income tax$",
    "tax": r"income tax expense$",
    "net_income": r"^profit for the (year|period)$",
    "ni_parent": r"owners of the parent$",
    "ni_nci": r"non-controlling interest$",
    "eps_basic": r"basic earnings per share$",
    "eps_diluted": r"diluted earnings per share$",
}
BS_MAP = {
    "cash": r"cash and cash equivalents$",
    "fvtpl_current": r"profit or loss - current$",
    "fvoci_current": r"other comprehensive income - current$",
    "amortised_cost_current": r"amortised cost - current$",
    "contract_assets_current": r"^contract assets - current$",
    "notes_receivable": r"^notes receivable, net$",
    "accounts_receivable": r"^accounts receivable, net$",
    "ar_related": r"^accounts receivable - related parties, net$",
    "inventories": r"inventories$",
    "total_current_assets": r"^total current assets$",
    "fvtpl_noncurrent": r"profit or loss - non-current$",
    "fvoci_noncurrent": r"other comprehensive income - non-current$",
    "amortised_cost_noncurrent": r"amortised cost - non-current$",
    "equity_method_investments": r"investments accounted for (under|using) (the )?equity method$",
    "ppe": r"property, plant and equipment$",
    "right_of_use": r"right-of-use assets$",
    "investment_property": r"investment property, net$",
    "intangibles": r"intangible assets$",
    "total_noncurrent_assets": r"^total non-current assets$",
    "total_assets": r"^total assets$",
    "st_borrowings": r"short-term borrowings$",
    "notes_payable": r"^notes payable$",
    "accounts_payable": r"^accounts payable$",
    "ap_related": r"^accounts payable - related parties$",
    "lt_debt_current": r"long-term liabilities, current portion$",
    "lease_liab_current": r"lease liabilities - current$",
    "total_current_liabilities": r"^total current liabilities$",
    "bonds_payable": r"bonds payable$",
    "lt_borrowings": r"long-term borrowings$",
    "lease_liab_noncurrent": r"lease liabilities - non-current$",
    "total_noncurrent_liabilities": r"^total non-current liabilities$",
    "total_liabilities": r"^total liabilities$",
    "share_capital": r"common stock$",
    "equity_parent": r"^equity attributable to owners of the parent$",
    "nci": r"non-controlling interest$",
    "total_equity": r"^total equity$",
    "total_liabilities_equity": r"^total liabilities and equity$",
}
CF_MAP = {
    "cf_pbt": r"profit before tax",
    "depreciation": r"depreciation( expense)?$",
    "amortisation": r"amortization( expense)?$|amortisation( expense)?$",
    "cf_change_ar": r"^accounts receivable$",
    "cf_change_inventories": r"^inventories$",
    "cf_change_ap": r"^accounts payable$",
    "interest_paid": r"^interest paid$",
    "taxes_paid": r"^income taxes? paid$",
    "cfo": r"^net cash flows .*operating activities$",
    "capex_ppe": r"^acquisition of property, plant and equipment$",
    "proceeds_ppe": r"^proceeds from disposal of property, plant and equipment$",
    "capex_intangibles": r"^acquisition of intangible assets$",
    "acquisitions": r"acquisition of subsidiaries",
    "cfi": r"^net cash flows .*investing activities$",
    "lease_principal": r"lease principal repayment$|repayment of lease",
    "dividends_paid": r"^cash dividends paid$",
    "dividends_paid_nci": r"cash dividends paid to (minority|non-controlling)",
    "cff": r"^net cash flows .*financing activities$",
    "fx_effect": r"^effects? (due to|of) (changes in )?exchange rate",
    "net_change_cash": r"^(net )?(increase|decrease|increase \(decrease\)|\(decrease\) increase) in cash and cash equivalents$",
    "cash_begin": r"cash and cash equivalents at beginning of (year|period)$",
    "cash_end": r"cash and cash equivalents at end of (year|period)$",
}


def standardise(lines: pd.DataFrame) -> pd.DataFrame:
    prim = lines[lines.apply(lambda r: PRIMARY.get(r.period) == r.document, axis=1)]
    # FY2025 balance sheet: its own annual report is primary (the interim comparatives must agree).
    data: dict[str, pd.Series] = {}
    for stmt, mp in (("IS", IS_MAP), ("BS", BS_MAP), ("CF", CF_MAP)):
        for key, pat in mp.items():
            data[key] = _find(prim, stmt, pat)
    df = pd.DataFrame(data).T
    df = df.reindex(columns=[p for p in PERIOD_ORDER if p in df.columns]).fillna(0.0)
    # cash flow statements are year-to-date: derive the second quarter as 1H - 1Q
    for yy in ("25", "26"):
        for k in CF_MAP:
            df.loc[k, f"2Q{yy}"] = df.loc[k, f"1H{yy}"] - df.loc[k, f"1Q{yy}"]
        df.loc["cash_begin", f"2Q{yy}"] = df.loc["cash_end", f"1Q{yy}"]
        df.loc["cash_end", f"2Q{yy}"] = df.loc["cash_end", f"1H{yy}"]
        df.loc["net_change_cash", f"2Q{yy}"] = df.loc["cash_end", f"2Q{yy}"] - df.loc["cash_begin", f"2Q{yy}"]
    bs_cols = [c for c in df.columns if c not in BS_PERIODS]
    df.loc[list(BS_MAP), bs_cols] = float("nan")  # no balance sheet for quarter-only periods

    m = df.copy()
    m.loc["st_investments"] = m.loc[["fvtpl_current", "fvoci_current", "amortised_cost_current"]].sum()
    m.loc["receivables"] = m.loc[["notes_receivable", "accounts_receivable", "ar_related"]].sum()
    m.loc["other_current_assets"] = (m.loc["total_current_assets"] - m.loc["cash"] - m.loc["st_investments"]
                                     - m.loc["receivables"] - m.loc["inventories"])
    m.loc["lt_investments"] = m.loc[["fvtpl_noncurrent", "fvoci_noncurrent", "amortised_cost_noncurrent",
                                     "equity_method_investments"]].sum()
    m.loc["other_noncurrent_assets"] = (m.loc["total_noncurrent_assets"] - m.loc["lt_investments"] - m.loc["ppe"]
                                        - m.loc["right_of_use"] - m.loc["intangibles"])
    m.loc["payables"] = m.loc[["notes_payable", "accounts_payable", "ap_related"]].sum()
    m.loc["st_debt"] = m.loc[["st_borrowings", "lt_debt_current"]].sum()
    m.loc["other_current_liabilities"] = (m.loc["total_current_liabilities"] - m.loc["payables"]
                                          - m.loc["st_debt"])
    m.loc["lt_debt"] = m.loc[["bonds_payable", "lt_borrowings"]].sum()
    m.loc["total_debt"] = m.loc["st_debt"] + m.loc["lt_debt"]
    m.loc["other_noncurrent_liabilities"] = (m.loc["total_noncurrent_liabilities"] - m.loc["lt_debt"]
                                             - m.loc["lease_liab_noncurrent"])
    m.loc["d_and_a"] = m.loc["depreciation"] + m.loc["amortisation"]
    m.loc["ebitda"] = m.loc["ebit"] + m.loc["d_and_a"]
    m.loc["capex"] = m.loc["capex_ppe"] + m.loc["capex_intangibles"]  # negative = outflow
    for p in df.columns:
        if p not in BS_PERIODS:
            for k in ("st_investments", "receivables", "other_current_assets", "lt_investments",
                      "other_noncurrent_assets", "payables", "st_debt", "other_current_liabilities", "lt_debt",
                      "total_debt", "other_noncurrent_liabilities"):
                m.loc[k, p] = float("nan")
    eps = m.loc[["eps_basic", "eps_diluted"]]
    m = m / 1000.0  # NT$ thousands -> NT$ millions
    m.loc[["eps_basic", "eps_diluted"]] = eps  # EPS stays in NT$ per share
    return m


# --------------------------------------------------------------------------- checks
def _block_sum_checks(lines: pd.DataFrame) -> list[dict]:
    """Sum of detail lines equals each subtotal, per document/statement/period."""
    rules = {
        "BS": [("^total current assets$", None), ("^total non-current assets$", "^total current assets$"),
               ("^total current liabilities$", "^total assets$"),
               ("^total non-current liabilities$", "^total current liabilities$"),
               ("^equity attributable to owners of the parent$", "^total liabilities$")],
        "IS": [("^total operating expenses$", "^gross profit"),
               ("^total non-operating income and expenses$", "^operating profit$"),
               ("^profit for the (year|period)$", "^total operating expenses$")],
        "CF": [("^net cash flows .*investing activities$", "operating activities$"),
               ("^net cash flows .*financing activities$", "investing activities$")],
    }
    sub_skip = re.compile(r"^(total |net cash flows|cash inflow generated|cash (inflow|outflow) (generated|used)|"
                          r"gross profit|operating profit|profit before income tax)", re.I)
    out = []
    for (doc, stmt, period), g in lines.groupby(["document", "statement", "period"]):
        g = g.sort_values("order")
        labels = list(g.label)
        vals = list(g.value)
        for total_pat, after_pat in rules.get(stmt, []):
            ti = next((i for i, l in enumerate(labels) if re.search(total_pat, l, re.I)), None)
            if ti is None:
                continue
            si = 0
            if after_pat:
                si = next((i for i, l in enumerate(labels[:ti]) if re.search(after_pat, l, re.I)), -1) + 1
            parts = [v for l, v in zip(labels[si:ti], vals[si:ti]) if not sub_skip.search(l)]
            if stmt == "IS" and total_pat.startswith("^profit for"):
                # operating profit + non-operating total + tax
                pick = lambda p: next(v for l, v in zip(labels, vals) if re.search(p, l, re.I))
                parts = [pick("^operating profit$"), pick("^total non-operating"), pick("income tax expense$")]
            diff = sum(parts) - vals[ti]
            out.append({"check": f"{stmt}: lines sum to '{labels[ti]}'", "document": doc, "period": period,
                        "expected": vals[ti], "actual": sum(parts), "diff": diff, "pass": abs(diff) <= 2})
    return out


def run_checks(lines: pd.DataFrame, std: pd.DataFrame, segs: pd.DataFrame) -> pd.DataFrame:
    out = _block_sum_checks(lines)
    tol = 0.002  # NT$ millions (= NT$2k rounding)

    def add(name, period, a, b, t=tol):
        out.append({"check": name, "document": PRIMARY.get(period, ""), "period": period, "expected": a,
                    "actual": b, "diff": b - a, "pass": abs(b - a) <= t})

    s = std
    for p in s.columns:
        add("IS: revenue + COGS = gross profit", p, s.at["gross_profit", p], s.at["revenue", p] + s.at["cogs", p])
        add("IS: gross profit + opex = EBIT", p, s.at["ebit", p], s.at["gross_profit", p] + s.at["opex_total", p])
        add("IS: EBIT + non-operating = PBT", p, s.at["pbt", p], s.at["ebit", p] + s.at["non_operating_total", p])
        add("IS: PBT + tax = net income", p, s.at["net_income", p], s.at["pbt", p] + s.at["tax", p])
        add("IS: parent + NCI = net income", p, s.at["net_income", p], s.at["ni_parent", p] + s.at["ni_nci", p])
        if p[:2] in ("FY", "1H", "1Q", "2Q"):
            add("CF: PBT in CF = PBT in IS", p, s.at["pbt", p], s.at["cf_pbt", p])
            add("CF: CFO + CFI + CFF + FX = change in cash", p, s.at["net_change_cash", p],
                s.at["cfo", p] + s.at["cfi", p] + s.at["cff", p] + s.at["fx_effect", p])
            add("CF: opening cash + change = closing cash", p, s.at["cash_end", p],
                s.at["cash_begin", p] + s.at["net_change_cash", p])
        if p in BS_PERIODS:
            add("BS: total assets = total liabilities + equity", p, s.at["total_assets", p],
                s.at["total_liabilities", p] + s.at["total_equity", p])
            add("BS: total assets = current + non-current", p, s.at["total_assets", p],
                s.at["total_current_assets", p] + s.at["total_noncurrent_assets", p])
            add("BS: equity = parent + NCI", p, s.at["total_equity", p], s.at["equity_parent", p] + s.at["nci", p])
            add("BS/CF: closing cash on CF = cash on BS", p, s.at["cash", p], s.at["cash_end", p])
    # quarters add up: 1Q + 2Q = 1H
    for yy in ("25", "26"):
        for k in ("revenue", "gross_profit", "ebit", "net_income", "ni_parent"):
            add(f"IS: 1Q{yy} + 2Q{yy} = 1H{yy} ({k})", f"1H{yy}", s.at[k, f"1H{yy}"],
                s.at[k, f"1Q{yy}"] + s.at[k, f"2Q{yy}"], t=0.002)
    # opening cash of each year = closing cash of the prior year
    for y in range(2022, 2026):
        add("CF: opening cash = prior year closing cash", f"FY{y}", s.at["cash_end", f"FY{y - 1}"],
            s.at["cash_begin", f"FY{y}"])
    # segments: four segments sum to total; segment operating profit total = group EBIT
    for (doc, period, item), g in segs.groupby(["document", "period", "item"]):
        tot = g[g.segment == "Total"].value.sum()
        parts = g[~g.segment.isin(["Total"]) & ~g.segment.str.startswith("Others")].value.sum()
        out.append({"check": f"Segments: four segments sum to total ({item})", "document": doc, "period": period,
                    "expected": tot, "actual": parts, "diff": parts - tot, "pass": abs(parts - tot) <= 2})
        if item == "op_profit" and period in s.columns:
            add(f"Segments: segment profit total = EBIT ({doc})", period, s.at["ebit", period], tot / 1000)
    # segment basis change: old Power Electronics = new Power Electronics + Mobility (FY2023)
    seg = segs.set_index(["document", "period", "segment", "item"]).value
    for item in ("revenue",):  # profit was also reclassified between segments, so only revenue must tie
        old = seg[("fs_2023_annual_consolidated.pdf", "FY2023", "Power Electronics", item)]
        new = (seg[("fs_2024_annual_consolidated.pdf", "FY2023", "Power Electronics", item)]
               + seg[("fs_2024_annual_consolidated.pdf", "FY2023", "Mobility", item)])
        out.append({"check": f"Segments: FY2023 old-basis PE = new-basis PE + Mobility ({item})",
                    "document": "fs_2023 vs fs_2024", "period": "FY2023", "expected": old, "actual": new,
                    "diff": new - old, "pass": abs(new - old) <= 2})
    # restatements: a period reported in two documents (prior-year comparatives)
    key = ["statement", "label", "period"]
    dup = lines.groupby(key).filter(lambda g: g.document.nunique() > 1)
    for (stmt, label, period), g in dup.groupby(key):
        if not re.search(r"^total|profit for the|^operating revenue|net cash flows|cash and cash equivalents$",
                         label, re.I):
            continue
        v = g.drop_duplicates("document").set_index("document").value
        out.append({"check": f"Restatement: {stmt} '{label}' same in all reports", "document": " vs ".join(v.index),
                    "period": period, "expected": v.iloc[0], "actual": v.iloc[-1],
                    "diff": v.iloc[-1] - v.iloc[0], "pass": abs(v.iloc[-1] - v.iloc[0]) <= 2,
                    })
    return pd.DataFrame(out)


# --------------------------------------------------------------------------- sources
def write_sources(lines: pd.DataFrame, std_keys: dict[str, tuple[str, str]], std: pd.DataFrame,
                  segs: pd.DataFrame) -> None:
    manifest = {r["file"]: r for r in csv.DictReader(open(RAW / "manifest.csv"))}
    path = ROOT / "data" / "sources.csv"
    cols = ["figure", "value", "unit", "period", "document", "page_or_url", "date_accessed", "notes"]
    # hand-entered figures (transcripts, slides) live in sources_manual.csv and are merged in
    manual = ROOT / "data" / "sources_manual.csv"
    existing = list(csv.DictReader(open(manual))) if manual.exists() else []
    rows = []
    prim = lines[lines.apply(lambda r: PRIMARY.get(r.period) == r.document, axis=1)]
    for key, (stmt, pat) in std_keys.items():
        sub = prim[(prim.statement == stmt) & prim.label.str.contains(pat, case=False, regex=True)]
        for period, g in sub.groupby("period"):
            r = g.sort_values("order").iloc[0]
            mf = manifest[r.document]
            rows.append({"figure": f"{key} ({r.label})", "value": r.value,
                         "unit": "NT$ per share" if key.startswith("eps") else "NT$ thousands",
                         "period": period, "document": r.document,
                         "page_or_url": f"p.{r.page} | {mf['url']}", "date_accessed": mf["date_accessed"],
                         "notes": "extract.py: as reported"})
    for _, r in segs.iterrows():
        mf = manifest[r.document]
        rows.append({"figure": f"segment {r['item']} - {r.segment}", "value": r.value, "unit": "NT$ thousands",
                     "period": r.period, "document": r.document,
                     "page_or_url": f"segment note | {mf['url']}", "date_accessed": mf["date_accessed"],
                     "notes": "extract.py: operating segment note"})
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in existing + rows:
            w.writerow({c: r.get(c, "") for c in cols})


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    lines = extract_lines()
    lines.to_csv(OUT / "fs_lines.csv", index=False)

    seg_rows = []
    for fname in DOCS:
        sub = lines[(lines.document == fname) & (lines.statement == "IS")
                    & lines.label.str.match(r"^operating (profit|income)$", case=False)]
        seg_rows += extract_segments(fname, dict(zip(sub.period, sub.value)))
    segs = pd.DataFrame(seg_rows)
    # latest = the most recent report presenting the period (restated basis)
    doc_rank = {d: i for i, d in enumerate(DOCS)}
    segs["rank"] = segs.document.map(doc_rank)
    latest = segs.groupby("period")["rank"].transform("max")
    segs["latest"] = segs["rank"] == latest
    segs = segs.drop(columns="rank")
    # revenue not allocated to the four segments ("Others" on Delta's slides) = group revenue - segment total
    rev = (lines[(lines.statement == "IS") & lines.label.str.match(r"^operating revenue$", case=False)]
           .drop_duplicates(["document", "period"]).set_index(["document", "period"]).value)
    extra = []
    for (doc, period), g in segs[segs["item"] == "revenue"].groupby(["document", "period"]):
        tot = g[g.segment == "Total"].value.iloc[0]
        extra.append({**g.iloc[0].to_dict(), "segment": "Others (derived: group revenue - segment total)",
                      "value": rev[(doc, period)] - tot})
    segs = pd.concat([segs, pd.DataFrame(extra)], ignore_index=True)
    segs.to_csv(OUT / "segments.csv", index=False)

    std = standardise(lines)
    std.index.name = "item"
    std.round(3).to_csv(OUT / "historicals.csv")

    monthly = extract_monthly()
    monthly.to_csv(OUT / "monthly_revenue.csv", index=False)

    checks = run_checks(lines, std, segs)
    # unaudited monthly revenue should sum to the reported quarters (within NT$3m of rounding)
    mm = monthly.set_index("month").revenue_ntd_m
    extra = []
    for q, months, yy in (("1Q", ("01", "02", "03"), "25"), ("2Q", ("04", "05", "06"), "25"),
                          ("1Q", ("01", "02", "03"), "26"), ("2Q", ("04", "05", "06"), "26")):
        tot = sum(mm[f"20{yy}-{m}"] for m in months)
        rep = std.at["revenue", f"{q}{yy}"]
        extra.append({"check": "Monthly: three monthly releases sum to quarterly revenue", "document":
                      "mr_*.pdf", "period": f"{q}{yy}", "expected": rep, "actual": tot, "diff": tot - rep,
                      "pass": abs(tot - rep) <= 3})
    checks = pd.concat([checks, pd.DataFrame(extra)], ignore_index=True)
    checks.to_csv(OUT / "checks.csv", index=False)

    keys = {k: ("IS", v) for k, v in IS_MAP.items()}
    keys.update({k: ("BS", v) for k, v in BS_MAP.items()})
    keys.update({k: ("CF", v) for k, v in CF_MAP.items()})
    write_sources(lines, keys, std, segs)

    bad = checks[~checks["pass"]]
    print(f"{len(lines)} statement values, {len(segs)} segment values, "
          f"{len(checks)} checks, {len(bad)} failing")
    if len(bad):
        print(bad.to_string(max_colwidth=70))
    return 0 if bad.empty else 1


if __name__ == "__main__":
    raise SystemExit(main())
