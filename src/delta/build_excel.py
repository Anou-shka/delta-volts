"""Build model/delta_model.xlsx with live formulas from data/processed/ and model/assumptions.csv.

Sheets: Cover, Assumptions, Historicals, Segments, IS, BS, CF, Schedules, Checks.
Columns: FY2021A-FY2025A (C-G), FY2026E-FY2030E forecast (H-L), interim 1H25A/1H26A
(N-O) and, on Segments, 3Q26A/9M26A from the 3Q26 earnings slides (P-Q).

Forecast logic (M3):
  revenue  = segment build: FY2026E = 9M26 actual + 4Q26E (split by 3Q26 mix);
             FY2027E+ = AI data-centre lines (capex growth x Delta content gain) + non-AI drivers
  EBIT     = sum of segment revenue x segment margin; gross margin and R&D are inputs, SG&A balances
  BS / CF  = working-capital days, capex and depreciation schedules; cash is the balancing item
             from the cash flow statement, so "assets = liabilities + equity" tests the model.

Conventions (banking): blue = hard-coded input, black = formula, green = link to another sheet,
yellow fill = Base-case input set by the analyst. NT$ million unless stated.

Run: python -m src.delta.build_excel   (then python -m src.delta.recalc model/delta_model.xlsx)
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .extract import BS_MAP, CF_MAP, IS_MAP

ROOT = Path(__file__).resolve().parents[2]
PROCESSED = ROOT / "data" / "processed"
ASSUMPTIONS = ROOT / "model" / "assumptions.csv"
SLIDES = ROOT / "data" / "slides_3q26.csv"
OUT = ROOT / "model" / "delta_model.xlsx"

FY = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]
FC = ["FY2026", "FY2027", "FY2028", "FY2029", "FY2030"]
INTERIM = ["1H25", "1H26"]
SLIDE = ["3Q26", "9M26"]
PERIODS = FY + INTERIM
COL = {"FY2021": "C", "FY2022": "D", "FY2023": "E", "FY2024": "F", "FY2025": "G",
       "FY2026": "H", "FY2027": "I", "FY2028": "J", "FY2029": "K", "FY2030": "L",
       "1H25": "N", "1H26": "O", "3Q26": "P", "9M26": "Q"}
HEAD = {**{p: f"{p}A" for p in FY + INTERIM + SLIDE}, **{p: f"{p}E" for p in FC}}
PERIOD_END = {"FY2021": "31-Dec-21", "FY2022": "31-Dec-22", "FY2023": "31-Dec-23", "FY2024": "31-Dec-24",
              "FY2025": "31-Dec-25", "FY2026": "31-Dec-26", "FY2027": "31-Dec-27", "FY2028": "31-Dec-28",
              "FY2029": "31-Dec-29", "FY2030": "31-Dec-30", "1H25": "30-Jun-25", "1H26": "30-Jun-26",
              "3Q26": "Jul-Sep 26", "9M26": "30-Sep-26"}
ALL_FY = FY + FC
PREV = {**{b: a for a, b in zip(ALL_FY, ALL_FY[1:])}, "1H26": "1H25"}
PREV_BS = {**{b: a for a, b in zip(ALL_FY, ALL_FY[1:])}, "1H25": "FY2024", "1H26": "FY2025"}
CHECK_ALL = "S"  # "All periods" column on the Checks sheet

# ---------------------------------------------------------------- styles
ARIAL = "Arial"
BLUE, BLACK, GREEN, GREY, WHITE = "0000FF", "000000", "008000", "595959", "FFFFFF"
NAVY, TEAL = "1F3864", "2F5597"
F_NUM = '#,##0.0;(#,##0.0);"-"'
F_PCT = '0.0%;(0.0%);"-"'
F_EPS = '0.00;(0.00);"-"'
F_DAYS = '0.0;(0.0);"-"'
F_X = '0.00"x";(0.00"x");"-"'
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_FC_HEAD = PatternFill("solid", fgColor=TEAL)
FILL_FC = PatternFill("solid", fgColor="F2F6FC")
FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
FILL_TRUE = PatternFill("solid", fgColor="C6EFCE")
FILL_FALSE = PatternFill("solid", fgColor="FFC7CE")
THIN_TOP = Border(top=Side(style="thin", color="808080"))
STYLE_COLOR = {"input": BLUE, "formula": BLACK, "link": GREEN}

LINK_RE = re.compile(r"^=-?'[^']+'!\$?[A-Z]+\$?\d+$")
REG: dict[tuple[str, str], int] = {}  # (sheet, key) -> row
PENDING: list = []  # forecast cells, written after every row is registered (allows forward references)


def font(color=BLACK, bold=False, italic=False, size=10):
    return Font(name=ARIAL, color=color, bold=bold, italic=italic, size=size)


def tf_format(ws, rng: str, top_left: str) -> None:
    """Green fill for TRUE, red for FALSE; blank cells stay unfilled."""
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISLOGICAL({top_left}),{top_left})"], fill=FILL_TRUE))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISLOGICAL({top_left}),NOT({top_left}))"],
                                                   fill=FILL_FALSE))


def ref(sheet: str, key: str, p: str, absolute: bool = False) -> str:
    c, r = COL[p], REG[(sheet, key)]
    return f"'{sheet}'!${c}${r}" if absolute else f"'{sheet}'!{c}{r}"


def A(key: str, p: str) -> str:
    """Live (scenario-selected) assumption for period p."""
    return ref("Assumptions", key, p)


def A26(key: str) -> str:
    """A one-off FY2026E assumption, anchored absolutely."""
    return ref("Assumptions", key, "FY2026", absolute=True)


def color_for(style: str, v) -> str:
    color = STYLE_COLOR[style]
    if style == "link" and isinstance(v, str) and not LINK_RE.match(v):
        color = BLACK  # a formula combining links is a calculation, not a link
    return color


class Sheet:
    """Period-grid sheet: labels in A, notes in B, historicals C-G, forecast H-L, interim N-Q."""

    def __init__(self, wb: Workbook, name: str, title: str, note: str, periods=PERIODS, forecast=True):
        self.ws = wb.create_sheet(name)
        self.name, self.periods, self.forecast, self.r = name, periods, forecast, 7
        ws = self.ws
        ws["A1"] = title
        ws["A1"].font = font(NAVY, bold=True, size=14)
        ws["A2"] = note
        ws["A2"].font = font(GREY, italic=True)
        ws.column_dimensions["A"].width = 46
        ws.column_dimensions["B"].width = 30
        for p in COL:
            ws.column_dimensions[COL[p]].width = 12
        ws.column_dimensions["M"].width = 2
        ws["A5"], ws["B5"] = "NT$ million", "Note / source"
        for c in ("A5", "B5"):
            ws[c].font = font(WHITE, bold=True)
            ws[c].fill = FILL_HEAD
        cols = list(periods) + (FC if forecast else [])
        for p in cols:
            h = ws[f"{COL[p]}5"]
            h.value = HEAD[p]
            h.font, h.alignment = font(WHITE, bold=True), Alignment(horizontal="right")
            h.fill = FILL_FC_HEAD if p in FC else FILL_HEAD
            d = ws[f"{COL[p]}6"]
            d.value = PERIOD_END[p]
            d.font, d.alignment = font(GREY, italic=True), Alignment(horizontal="right")
        if forecast:
            ws["H4"] = "Forecast (scenario: Assumptions!B3)"
            ws["H4"].font = font(TEAL, bold=True, size=9)
        if any(p in periods for p in INTERIM):
            ws["N4"] = "Interim"
            ws["N4"].font = font(GREY, italic=True)
        ws.freeze_panes = "C7"
        self.cols = cols

    def section(self, text: str):
        self.r += 1
        for c in ["A", "B"] + [COL[p] for p in self.cols]:
            self.ws[f"{c}{self.r}"].fill = FILL_SECTION
        self.ws[f"A{self.r}"] = text
        self.ws[f"A{self.r}"].font = font(NAVY, bold=True)
        self.r += 1

    def line(self, key, label, fn=None, fc=None, fmt=F_NUM, style="formula", fc_style="formula", bold=False,
             note="", periods=None, total=False):
        """fn(p) -> value for historical/interim periods (written now); fc(p) -> forecast formula (deferred)."""
        row = self.r
        REG[(self.name, key)] = row
        ws = self.ws
        ws[f"A{row}"] = label
        ws[f"A{row}"].font = font(bold=bold or total)
        if note:
            ws[f"B{row}"] = note
            ws[f"B{row}"].font = font(GREY, italic=True, size=9)
        if fn is not None:
            for p in periods or self.periods:
                self._write(row, p, fn(p), fmt, style, bold or total, total)
        if fc is not None and self.forecast:
            PENDING.append((self, row, fc, fmt, fc_style, bold or total, total))
        self.r += 1
        return row

    def _write(self, row, p, v, fmt, style, bold, total):
        if v is None:
            return
        c = self.ws[f"{COL[p]}{row}"]
        c.value = v
        c.font = font(color_for(style, v), bold=bold)
        c.number_format = fmt
        if p in FC:
            c.fill = FILL_FC
        if total:
            c.border = THIN_TOP

    def ref(self, key, p):
        return f"{COL[p]}{REG[(self.name, key)]}"

    def blank(self, n=1):
        self.r += n


def finalize() -> None:
    for sh, row, fc, fmt, style, bold, total in PENDING:
        for p in FC:
            sh._write(row, p, fc(p), fmt, style, bold, total)
    PENDING.clear()


# ---------------------------------------------------------------- data
def load():
    h = pd.read_csv(PROCESSED / "historicals.csv", index_col=0)
    s = pd.read_csv(PROCESSED / "segments.csv")
    lines = pd.read_csv(PROCESSED / "fs_lines.csv")
    checks = pd.read_csv(PROCESSED / "checks.csv")
    a = pd.read_csv(ASSUMPTIONS)
    return h, s, lines, checks, a


HIST_LABELS = {
    # IS
    "revenue": "Operating revenue", "cogs": "Operating costs", "gross_profit": "Gross profit (reported)",
    "selling": "Selling expenses", "g_and_a": "General and administrative expenses",
    "r_and_d": "Research and development expenses", "credit_loss": "Expected credit impairment (loss) gain",
    "opex_total": "Total operating expenses (reported)", "ebit": "Operating profit (reported)",
    "interest_income": "Interest income", "other_income": "Other income",
    "other_gains_losses": "Other gains and losses", "finance_costs": "Finance costs",
    "equity_method_income": "Share of profit (loss) of associates and JVs",
    "non_operating_total": "Total non-operating income and expenses (reported)",
    "pbt": "Profit before income tax (reported)", "tax": "Income tax expense",
    "net_income": "Profit for the period (reported)", "ni_parent": "Profit attributable to owners of the parent",
    "ni_nci": "Profit attributable to non-controlling interests", "eps_basic": "Basic EPS (NT$)",
    "eps_diluted": "Diluted EPS (NT$)",
    # BS
    "cash": "Cash and cash equivalents", "fvtpl_current": "Financial assets at FVTPL - current",
    "fvoci_current": "Financial assets at FVOCI - current",
    "amortised_cost_current": "Financial assets at amortised cost - current",
    "contract_assets_current": "Contract assets - current", "notes_receivable": "Notes receivable, net",
    "accounts_receivable": "Accounts receivable, net", "ar_related": "Accounts receivable - related parties",
    "inventories": "Inventories", "total_current_assets": "Total current assets (reported)",
    "fvtpl_noncurrent": "Financial assets at FVTPL - non-current",
    "fvoci_noncurrent": "Financial assets at FVOCI - non-current",
    "amortised_cost_noncurrent": "Financial assets at amortised cost - non-current",
    "equity_method_investments": "Investments accounted for under the equity method",
    "ppe": "Property, plant and equipment", "right_of_use": "Right-of-use assets",
    "investment_property": "Investment property, net", "intangibles": "Intangible assets (incl. goodwill)",
    "total_noncurrent_assets": "Total non-current assets (reported)", "total_assets": "Total assets (reported)",
    "st_borrowings": "Short-term borrowings", "notes_payable": "Notes payable", "accounts_payable": "Accounts payable",
    "ap_related": "Accounts payable - related parties", "lt_debt_current": "Long-term liabilities, current portion",
    "lease_liab_current": "Lease liabilities - current (if shown separately)",
    "total_current_liabilities": "Total current liabilities (reported)", "bonds_payable": "Bonds payable",
    "lt_borrowings": "Long-term borrowings", "lease_liab_noncurrent": "Lease liabilities - non-current",
    "total_noncurrent_liabilities": "Total non-current liabilities (reported)",
    "total_liabilities": "Total liabilities (reported)", "share_capital": "Share capital - common stock",
    "equity_parent": "Equity attributable to owners of the parent", "nci": "Non-controlling interest",
    "total_equity": "Total equity (reported)", "total_liabilities_equity": "Total liabilities and equity (reported)",
    # CF
    "cf_pbt": "Profit before tax (cash flow statement)", "depreciation": "Depreciation",
    "amortisation": "Amortisation", "cf_change_ar": "Change in accounts receivable",
    "cf_change_inventories": "Change in inventories", "cf_change_ap": "Change in accounts payable",
    "interest_paid": "Interest paid", "taxes_paid": "Income taxes paid",
    "cfo": "Net cash flows from operating activities (reported)",
    "capex_ppe": "Acquisition of property, plant and equipment",
    "proceeds_ppe": "Proceeds from disposal of PP&E", "capex_intangibles": "Acquisition of intangible assets",
    "acquisitions": "Acquisition of subsidiaries, net of cash acquired",
    "cfi": "Net cash flows from investing activities (reported)",
    "st_borrowings_change": "Increase (decrease) in short-term borrowings",
    "bonds_issued": "Issuance of bonds payable", "bonds_repaid": "Repayment of bonds payable",
    "lt_debt_proceeds": "Proceeds from long-term borrowings", "lt_debt_repaid": "Repayment of long-term borrowings",
    "lease_principal": "Lease principal repayment", "dividends_paid": "Cash dividends paid (owners of the parent)",
    "dividends_paid_nci": "Cash dividends paid to non-controlling interests",
    "cff": "Net cash flows from financing activities (reported)",
    "fx_effect": "Effect of exchange rate changes on cash",
    "net_change_cash": "Net increase (decrease) in cash (reported)",
    "cash_begin": "Cash at beginning of period (reported)", "cash_end": "Cash at end of period (reported)",
}
SKIP_BS_PERIODS = set()  # all historical columns carry a balance sheet (1H = 30 June)


# ---------------------------------------------------------------- sheets
def build_historicals(wb, h, checks):
    sh = Sheet(wb, "Historicals", "Historical financial statements - as reported",
               "NT$ million (reported in NT$ thousands, shown /1,000). Taiwan IFRS, consolidated. Blue = hard-coded "
               "from the filings; page reference for every value in data/sources.csv.", forecast=False)
    ws = sh.ws
    docs = {**{p: f"fs_{p[2:]}_annual_consolidated.pdf" for p in FY},
            "1H25": "fs_2026q2_consolidated.pdf (comparative)", "1H26": "fs_2026q2_consolidated.pdf"}
    ws["A3"] = "Source document"
    ws["A3"].font = font(GREY, italic=True)
    for p in PERIODS:
        ws[f"{COL[p]}3"] = docs[p].replace("_consolidated.pdf", "").replace("fs_", "")
        ws[f"{COL[p]}3"].font = font(GREY, italic=True, size=8)
    for title, mp in (("Income statement", IS_MAP), ("Balance sheet (period end)", BS_MAP),
                      ("Cash flow statement", CF_MAP)):
        sh.section(title)
        for key in mp:
            if key not in h.index:
                continue
            vals = h.loc[key]
            fmt = F_EPS if key.startswith("eps") else F_NUM
            sh.line(key, HIST_LABELS.get(key, key), lambda p, v=vals: (None if pd.isna(v[p]) else float(v[p])),
                    fmt=fmt, style="input", note="extracted; see sources.csv")
    sh.section("Other inputs")
    sh.line("par_value", "Par value per share (NT$)", lambda p: 10.0, fmt=F_EPS, style="input",
            note="Taiwan standard par; 2,598m shares per ic_4Q25 p.10")
    sh.line("checks_total", "Python extraction tie-outs: total", lambda p: float(len(checks)), fmt="0",
            style="input", note="data/processed/checks.csv", periods=["FY2021"])
    sh.line("checks_passed", "Python extraction tie-outs: passed", lambda p: float(checks["pass"].sum()), fmt="0",
            style="input", note="data/processed/checks.csv", periods=["FY2021"])
    return sh


def H(key, p):
    return ref("Historicals", key, p)


def X(sheet, key, p):
    return ref(sheet, key, p)


def build_is(wb):
    sh = Sheet(wb, "IS", "Income statement",
               "NT$ million unless stated. Historicals link to Historicals and recompute every subtotal; forecast "
               "EBIT is built from segments (revenue x margin), with gross margin and R&D as inputs and SG&A the balance.")
    lk = lambda k: (lambda p: f"={H(k, p)}")  # noqa: E731
    S = lambda k, p: X("Segments", k, p)  # noqa: E731
    sh.section("Income statement")
    sh.line("revenue", "Revenue", lk("revenue"), fc=lambda p: f"={S('rev_total', p)}", style="link",
            fc_style="link", bold=True)
    sh.line("cogs", "Cost of revenue", lk("cogs"),
            fc=lambda p: f"=-{sh.ref('revenue', p)}*(1-{A('gross_margin', p)})", style="link")
    sh.line("gross_profit", "Gross profit", lambda p: f"={sh.ref('revenue', p)}+{sh.ref('cogs', p)}",
            fc=lambda p: f"={sh.ref('revenue', p)}+{sh.ref('cogs', p)}", total=True)
    sh.line("sga", "Selling, general and administrative",
            lambda p: f"={H('selling', p)}+{H('g_and_a', p)}",
            fc=lambda p: f"={S('op_total', p)}-{sh.ref('gross_profit', p)}-{sh.ref('r_and_d', p)}",
            note="forecast: balancing item to segment EBIT")
    sh.line("r_and_d", "Research and development", lk("r_and_d"),
            fc=lambda p: f"=-{A('rnd_pct', p)}*{sh.ref('revenue', p)}", style="link")
    sh.line("credit_loss", "Expected credit loss", lk("credit_loss"), style="link")
    sum_opex = lambda p: f"=SUM({sh.ref('sga', p)}:{sh.ref('credit_loss', p)})"  # noqa: E731
    sh.line("opex", "Total operating expenses", sum_opex, fc=sum_opex, total=True)
    ebit = lambda p: f"={sh.ref('gross_profit', p)}+{sh.ref('opex', p)}"  # noqa: E731
    sh.line("ebit", "Operating profit (EBIT)", ebit, fc=ebit, total=True)
    sh.line("interest_income", "Interest income", lk("interest_income"),
            fc=lambda p: f"={A('cash_yield', p)}*{X('BS', 'cash', PREV_BS[p])}", style="link",
            note="forecast: yield x opening cash")
    sh.line("finance_costs", "Finance costs", lk("finance_costs"),
            fc=lambda p: f"=-{A('cost_debt', p)}*{X('BS', 'total_debt', PREV_BS[p])}", style="link",
            note="forecast: cost of debt x opening debt")
    sh.line("other_nonop", "Other non-operating items",
            lambda p: f"={H('other_income', p)}+{H('other_gains_losses', p)}+{H('equity_method_income', p)}",
            fc=lambda p: f"={A('other_nonop_pct', p)}*{sh.ref('revenue', p)}",
            note="other income, gains/losses, associates")
    sum_nonop = lambda p: f"=SUM({sh.ref('interest_income', p)}:{sh.ref('other_nonop', p)})"  # noqa: E731
    sh.line("nonop", "Total non-operating income", sum_nonop, fc=sum_nonop, total=True)
    pbt = lambda p: f"={sh.ref('ebit', p)}+{sh.ref('nonop', p)}"  # noqa: E731
    sh.line("pbt", "Profit before tax", pbt, fc=pbt, total=True)
    sh.line("tax", "Income tax", lk("tax"), fc=lambda p: f"=-{A('etr', p)}*{sh.ref('pbt', p)}", style="link")
    ni = lambda p: f"={sh.ref('pbt', p)}+{sh.ref('tax', p)}"  # noqa: E731
    sh.line("net_income", "Net income", ni, fc=ni, total=True)
    sh.line("ni_nci", "Less: non-controlling interests", lambda p: f"=-{H('ni_nci', p)}",
            fc=lambda p: f"=-{A('nci_pct', p)}*{sh.ref('net_income', p)}", style="link")
    nip = lambda p: f"={sh.ref('net_income', p)}+{sh.ref('ni_nci', p)}"  # noqa: E731
    sh.line("ni_parent", "Net income to owners of the parent", nip, fc=nip, total=True)
    sh.blank()
    sh.line("shares", "Shares outstanding, period end (m)",
            lambda p: f"={H('share_capital', p)}/{H('par_value', p)}",
            fc=lambda p: f"={sh.ref('shares', PREV[p])}", note="share capital / NT$10 par; held flat")
    eps = lambda p: f"={sh.ref('ni_parent', p)}/{sh.ref('shares', p)}"  # noqa: E731
    sh.line("eps", "EPS, basic (NT$)", eps, fc=eps, fmt=F_EPS, bold=True, note="reported EPS on Historicals")
    sh.section("Memo")
    sh.line("d_and_a", "Depreciation and amortisation",
            lambda p: f"={H('depreciation', p)}+{H('amortisation', p)}",
            fc=lambda p: f"=-{X('Schedules', 'dep', p)}+{X('Schedules', 'amort', p)}",
            note="historical: cash flow statement")
    ebitda = lambda p: f"={sh.ref('ebit', p)}+{sh.ref('d_and_a', p)}"  # noqa: E731
    sh.line("ebitda", "EBITDA (EBIT + D&A)", ebitda, fc=ebitda, bold=True, note="operating definition (decision D3)")
    sh.section("Ratios")
    growth = lambda k: (lambda p: f"={sh.ref(k, p)}/{sh.ref(k, PREV[p])}-1" if p in PREV else None)  # noqa: E731
    pct = lambda k, sign="": (lambda p: f"={sign}{sh.ref(k, p)}/{sh.ref('revenue', p)}")  # noqa: E731
    for key, label, f, kw in (
            ("rev_growth", "Revenue growth (y/y)", growth("revenue"), {"note": "1H26 vs 1H25"}),
            ("gm", "Gross margin", pct("gross_profit"), {}),
            ("rnd_pct", "R&D % of revenue", pct("r_and_d", "-"), {}),
            ("sga_pct", "SG&A % of revenue", pct("sga", "-"), {}),
            ("ebit_margin", "EBIT margin", pct("ebit"), {"bold": True}),
            ("ebitda_margin", "EBITDA margin", pct("ebitda"), {}),
            ("etr", "Effective tax rate", lambda p: f"=-{sh.ref('tax', p)}/{sh.ref('pbt', p)}", {}),
            ("ni_margin", "Net margin (to parent)", pct("ni_parent"), {}),
            ("eps_growth", "EPS growth (y/y)", growth("eps"), {})):
        sh.line(key, label, f, fc=f, fmt=F_PCT, **kw)
    return sh


def build_bs(wb):
    sh = Sheet(wb, "BS", "Balance sheet",
               "NT$ million, period end. Historical 'other' lines tie each section to a separately reported subtotal. "
               "Forecast: working capital from days, PP&E and intangibles from Schedules, cash from the cash flow "
               "statement, other items held flat. Assets = liabilities + equity is tested on Checks.")
    lk = lambda k: (lambda p: f"={H(k, p)}")  # noqa: E731
    s = lambda *ks: (lambda p: "=" + "+".join(H(k, p) for k in ks))  # noqa: E731
    flat = lambda k: (lambda p: f"={sh.ref(k, PREV_BS[p])}")  # noqa: E731
    IS = lambda k, p: X("IS", k, p)  # noqa: E731
    days = "'Schedules'!$B$4"
    sh.section("Assets")
    sh.line("cash", "Cash and cash equivalents", lk("cash"), fc=lambda p: f"={X('CF', 'cash_close', p)}",
            style="link", fc_style="link", note="forecast: from cash flow statement")
    sh.line("st_investments", "Short-term investments", s("fvtpl_current", "fvoci_current", "amortised_cost_current"),
            fc=flat("st_investments"), note="FVTPL + FVOCI + amortised cost")
    sh.line("receivables", "Receivables", s("notes_receivable", "accounts_receivable", "ar_related"),
            fc=lambda p: f"={A('dso', p)}/{days}*{IS('revenue', p)}", note="notes + trade + related; DSO")
    sh.line("inventories", "Inventories", lk("inventories"), fc=lambda p: f"={A('dio', p)}/{days}*-{IS('cogs', p)}",
            style="link", note="forecast: DIO")
    sh.line("other_ca", "Other current assets",
            lambda p: f"={H('total_current_assets', p)}-SUM({sh.ref('cash', p)}:{sh.ref('inventories', p)})",
            fc=lambda p: f"={A('other_ca_pct', p)}*{IS('revenue', p)}", note="contract assets, prepayments, tax")
    tca = lambda p: f"=SUM({sh.ref('cash', p)}:{sh.ref('other_ca', p)})"  # noqa: E731
    sh.line("tca", "Total current assets", tca, fc=tca, total=True)
    sh.line("lt_investments", "Long-term investments",
            s("fvtpl_noncurrent", "fvoci_noncurrent", "amortised_cost_noncurrent", "equity_method_investments"),
            fc=flat("lt_investments"))
    sh.line("ppe", "Property, plant and equipment", lk("ppe"), fc=lambda p: f"={X('Schedules', 'ppe_close', p)}",
            style="link", fc_style="link")
    sh.line("rou", "Right-of-use assets", lk("right_of_use"), fc=flat("rou"), style="link")
    sh.line("intangibles", "Intangible assets (incl. goodwill)", lk("intangibles"),
            fc=lambda p: f"={X('Schedules', 'int_close', p)}", style="link", fc_style="link")
    sh.line("other_nca", "Other non-current assets",
            lambda p: f"={H('total_noncurrent_assets', p)}-SUM({sh.ref('lt_investments', p)}:{sh.ref('intangibles', p)})",
            fc=flat("other_nca"), note="deferred tax, investment property")
    tnca = lambda p: f"=SUM({sh.ref('lt_investments', p)}:{sh.ref('other_nca', p)})"  # noqa: E731
    sh.line("tnca", "Total non-current assets", tnca, fc=tnca, total=True)
    ta = lambda p: f"={sh.ref('tca', p)}+{sh.ref('tnca', p)}"  # noqa: E731
    sh.line("ta", "Total assets", ta, fc=ta, bold=True, total=True)
    sh.section("Liabilities and equity")
    sh.line("st_debt", "Short-term debt", s("st_borrowings", "lt_debt_current"), fc=flat("st_debt"),
            note="ST borrowings + current LT debt; held flat")
    sh.line("payables", "Payables", s("notes_payable", "accounts_payable", "ap_related"),
            fc=lambda p: f"={A('dpo', p)}/{days}*-{IS('cogs', p)}", note="notes + trade + related; DPO")
    sh.line("other_cl", "Other current liabilities",
            lambda p: f"={H('total_current_liabilities', p)}-{sh.ref('st_debt', p)}-{sh.ref('payables', p)}",
            fc=lambda p: f"={A('other_cl_pct', p)}*{IS('revenue', p)}", note="other payables, tax, contract liab.")
    tcl = lambda p: f"=SUM({sh.ref('st_debt', p)}:{sh.ref('other_cl', p)})"  # noqa: E731
    sh.line("tcl", "Total current liabilities", tcl, fc=tcl, total=True)
    sh.line("lt_debt", "Long-term debt", s("bonds_payable", "lt_borrowings"), fc=flat("lt_debt"),
            note="bonds + LT borrowings; held flat")
    sh.line("leases", "Lease liabilities (non-current)", lk("lease_liab_noncurrent"), fc=flat("leases"), style="link")
    sh.line("other_ncl", "Other non-current liabilities",
            lambda p: f"={H('total_noncurrent_liabilities', p)}-{sh.ref('lt_debt', p)}-{sh.ref('leases', p)}",
            fc=flat("other_ncl"), note="deferred tax liabilities")
    tncl = lambda p: f"=SUM({sh.ref('lt_debt', p)}:{sh.ref('other_ncl', p)})"  # noqa: E731
    sh.line("tncl", "Total non-current liabilities", tncl, fc=tncl, total=True)
    tl = lambda p: f"={sh.ref('tcl', p)}+{sh.ref('tncl', p)}"  # noqa: E731
    sh.line("tl", "Total liabilities", tl, fc=tl, total=True)
    sh.line("equity_parent", "Equity attributable to owners of the parent", lk("equity_parent"),
            fc=lambda p: f"={sh.ref('equity_parent', PREV_BS[p])}+{IS('ni_parent', p)}+{X('CF', 'dividends', p)}",
            style="link", note="forecast: + net income - dividends")
    sh.line("nci", "Non-controlling interest", lk("nci"),
            fc=lambda p: f"={sh.ref('nci', PREV_BS[p])}-{IS('ni_nci', p)}+{X('CF', 'dividends_nci', p)}", style="link")
    te = lambda p: f"={sh.ref('equity_parent', p)}+{sh.ref('nci', p)}"  # noqa: E731
    sh.line("te", "Total equity", te, fc=te, total=True)
    tle = lambda p: f"={sh.ref('tl', p)}+{sh.ref('te', p)}"  # noqa: E731
    sh.line("tle", "Total liabilities and equity", tle, fc=tle, bold=True, total=True)
    diff = lambda p: f"={sh.ref('ta', p)}-{sh.ref('tle', p)}"  # noqa: E731
    sh.line("bs_diff", "Balance check: assets - (liabilities + equity)", diff, fc=diff, note="must be zero")
    sh.section("Memo")
    td = lambda p: f"={sh.ref('st_debt', p)}+{sh.ref('lt_debt', p)}"  # noqa: E731
    sh.line("total_debt", "Total debt", td, fc=td, bold=True)
    nd = lambda p: f"={sh.ref('total_debt', p)}-{sh.ref('cash', p)}-{sh.ref('st_investments', p)}"  # noqa: E731
    sh.line("net_debt", "Net debt / (net cash)", nd, fc=nd, note="debt - cash - ST investments; excl. leases")
    nwc = lambda p: f"={sh.ref('receivables', p)}+{sh.ref('inventories', p)}-{sh.ref('payables', p)}"  # noqa: E731
    sh.line("nwc", "Trade working capital", nwc, fc=nwc, note="receivables + inventories - payables")
    return sh


def build_cf(wb):
    sh = Sheet(wb, "CF", "Cash flow statement",
               "NT$ million. Historicals as reported (Taiwan IFRS: interest and tax paid sit in operating cash flow); "
               "'other' lines tie each section to the reported total. Forecast: indirect method from the IS and BS; "
               "closing cash feeds the balance sheet.")
    lk = lambda k: (lambda p: f"={H(k, p)}")  # noqa: E731
    BS = lambda k, p: X("BS", k, p)  # noqa: E731
    IS = lambda k, p: X("IS", k, p)  # noqa: E731
    SC = lambda k, p: X("Schedules", k, p)  # noqa: E731

    def chg(k, sign=-1):
        return lambda p: f"={'-' if sign < 0 else ''}({BS(k, p)}-{BS(k, PREV_BS[p])})"
    sh.section("Operating activities")
    sh.line("pbt", "Profit before tax", lk("cf_pbt"), fc=lambda p: f"={IS('pbt', p)}", style="link", fc_style="link")
    sh.line("d_and_a", "Depreciation and amortisation", lambda p: f"={H('depreciation', p)}+{H('amortisation', p)}",
            fc=lambda p: f"={IS('d_and_a', p)}", fc_style="link")
    sh.line("chg_ar", "(Increase) decrease in receivables", lk("cf_change_ar"), fc=chg("receivables"), style="link")
    sh.line("chg_inv", "(Increase) decrease in inventories", lk("cf_change_inventories"), fc=chg("inventories"),
            style="link")
    sh.line("chg_ap", "Increase (decrease) in payables", lk("cf_change_ap"), fc=chg("payables", +1), style="link")
    sh.line("interest_paid", "Interest paid", lk("interest_paid"), style="link", note="forecast: within PBT")
    sh.line("taxes_paid", "Income taxes paid", lk("taxes_paid"), fc=lambda p: f"={IS('tax', p)}", style="link",
            fc_style="link", note="forecast: cash tax = book tax")
    sh.line("other_op", "Other operating items",
            lambda p: f"={H('cfo', p)}-SUM({sh.ref('pbt', p)}:{sh.ref('taxes_paid', p)})",
            fc=lambda p: f"=-({BS('other_ca', p)}-{BS('other_ca', PREV_BS[p])})+({BS('other_cl', p)}-{BS('other_cl', PREV_BS[p])})",
            note="hist: non-cash items, other WC; fcst: other WC")
    cfo = lambda p: f"=SUM({sh.ref('pbt', p)}:{sh.ref('other_op', p)})"  # noqa: E731
    sh.line("cfo", "Cash flow from operations", cfo, fc=cfo, total=True)
    sh.section("Investing activities")
    sh.line("capex_ppe", "Purchase of PP&E", lk("capex_ppe"), fc=lambda p: f"=-{SC('capex_ppe', p)}", style="link",
            fc_style="link")
    sh.line("proceeds_ppe", "Disposal of PP&E", lk("proceeds_ppe"), style="link")
    sh.line("capex_int", "Purchase of intangible assets", lk("capex_intangibles"),
            fc=lambda p: f"=-{SC('capex_int', p)}", style="link", fc_style="link")
    sh.line("acquisitions", "Acquisitions, net of cash acquired", lk("acquisitions"), style="link")
    sh.line("other_inv", "Other investing items",
            lambda p: f"={H('cfi', p)}-SUM({sh.ref('capex_ppe', p)}:{sh.ref('acquisitions', p)})",
            note="financial assets, prepayments, other")
    cfi = lambda p: f"=SUM({sh.ref('capex_ppe', p)}:{sh.ref('other_inv', p)})"  # noqa: E731
    sh.line("cfi", "Cash flow from investing", cfi, fc=cfi, total=True)
    sh.section("Financing activities")
    sh.line("net_borrowing", "Net borrowing / (repayment)",
            lambda p: "=" + "+".join(H(k, p) for k in ("st_borrowings_change", "bonds_issued", "bonds_repaid",
                                                       "lt_debt_proceeds", "lt_debt_repaid")),
            fc=lambda p: f"={BS('total_debt', p)}-{BS('total_debt', PREV_BS[p])}", note="ST + bonds + LT, net")
    sh.line("lease_principal", "Lease principal repaid", lk("lease_principal"), style="link")
    sh.line("dividends", "Dividends paid to shareholders", lk("dividends_paid"),
            fc=lambda p: (f"=-{A26('dps_2026')}*{IS('shares', 'FY2025')}" if p == "FY2026"
                          else f"=-{A('payout', p)}*{IS('ni_parent', PREV[p])}"),
            style="link", note="paid in Q3; FY2026 = NT$11.6 x shares")
    sh.line("dividends_nci", "Dividends paid to non-controlling interests", lk("dividends_paid_nci"),
            fc=lambda p: f"={A('nci_div_pct', p)}*{IS('ni_nci', PREV[p])}", style="link")
    sh.line("other_fin", "Other financing items",
            lambda p: f"={H('cff', p)}-SUM({sh.ref('net_borrowing', p)}:{sh.ref('dividends_nci', p)})",
            note="deposits, NCI transactions")
    cff = lambda p: f"=SUM({sh.ref('net_borrowing', p)}:{sh.ref('other_fin', p)})"  # noqa: E731
    sh.line("cff", "Cash flow from financing", cff, fc=cff, total=True)
    sh.blank()
    sh.line("fx", "Effect of exchange rates", lk("fx_effect"), style="link", note="forecast: nil")
    net = lambda p: f"={sh.ref('cfo', p)}+{sh.ref('cfi', p)}+{sh.ref('cff', p)}+{sh.ref('fx', p)}"  # noqa: E731
    sh.line("net_change", "Net change in cash", net, fc=net, total=True)
    sh.line("cash_open", "Cash at beginning of period",
            lambda p: f"={BS('cash', PREV_BS[p])}" if p in PREV_BS else f"={H('cash_begin', p)}",
            fc=lambda p: f"={BS('cash', PREV_BS[p])}", style="link", fc_style="link",
            note="prior balance sheet; FY2021 from CF")
    close = lambda p: f"={sh.ref('cash_open', p)}+{sh.ref('net_change', p)}"  # noqa: E731
    sh.line("cash_close", "Cash at end of period", close, fc=close, bold=True, total=True)
    sh.section("Memo")
    capex = lambda p: f"={sh.ref('capex_ppe', p)}+{sh.ref('capex_int', p)}"  # noqa: E731
    sh.line("capex", "Capex (PP&E + intangibles)", capex, fc=capex, note="decision D4")
    fcf = lambda p: f"={sh.ref('cfo', p)}+{sh.ref('capex', p)}"  # noqa: E731
    sh.line("fcf", "Free cash flow (CFO + capex)", fcf, fc=fcf, bold=True)
    fm = lambda p: f"={sh.ref('fcf', p)}/{IS('revenue', p)}"  # noqa: E731
    sh.line("fcf_margin", "FCF margin", fm, fc=fm, fmt=F_PCT)
    cc = lambda p: f"={sh.ref('cfo', p)}/{IS('net_income', p)}"  # noqa: E731
    sh.line("cash_conversion", "Cash conversion (CFO / net income)", cc, fc=cc, fmt=F_X)
    return sh


SEGS = ["Power Electronics", "Infrastructure", "Automation", "Mobility"]
SEG_KEY = {"Power Electronics": "pe", "Infrastructure": "infra", "Automation": "auto", "Mobility": "mob"}


def build_segments(wb, segs):
    periods = PERIODS + SLIDE
    sh = Sheet(wb, "Segments", "Segment build: revenue, AI data-centre line and profit",
               "NT$ million. Historicals: segment note (latest restated basis; FY2021-22 on the old 3-segment basis). "
               "3Q26A/9M26A: 3Q26 earnings slides. FY2026E = 9M26 actual + 4Q26E split by 3Q26 mix; FY2027E+ = "
               "AI data-centre lines + non-AI drivers.", periods=periods)
    ws = sh.ws
    ws["P4"] = "3Q26 earnings slides"
    ws["P4"].font = font(GREY, italic=True)
    latest = segs[segs.latest]
    IS = lambda k, p: X("IS", k, p)  # noqa: E731

    def val(seg, item):
        def f(p):
            r = latest[(latest.segment == seg) & (latest["item"] == item) & (latest.period == p)]
            return None if r.empty else float(r.value.iloc[0]) / 1000
        return f

    def fy26(k):
        return f"={COL['9M26']}{REG[('Segments', k)]}+{A26('q4_revenue')}*{COL['3Q26']}{REG[('Segments', k)]}" \
               f"/{COL['3Q26']}{REG[('Segments', 'rev_total')]}"

    sh.section("Revenue from external customers")
    seg_fc = {
        "pe": lambda p: fy26("rev_pe") if p == "FY2026" else f"={sh.ref('nonai_pe', p)}+{sh.ref('ai_pe', p)}",
        "infra": lambda p: fy26("rev_infra") if p == "FY2026" else f"={sh.ref('nonai_infra', p)}+{sh.ref('ai_infra', p)}",
        "auto": lambda p: fy26("rev_auto") if p == "FY2026" else f"={sh.ref('rev_auto', PREV[p])}*(1+{A('g_auto', p)})",
        "mob": lambda p: fy26("rev_mob") if p == "FY2026" else f"={sh.ref('rev_mob', PREV[p])}*(1+{A('g_mob', p)})",
    }
    for seg in SEGS:
        k = SEG_KEY[seg]
        sh.line(f"rev_{k}", seg, val(seg, "revenue"), fc=seg_fc[k], style="input",
                note="inside Power Electronics before FY2023" if seg == "Mobility" else "segment note / slides")
    others_sl = val("Others (derived: group revenue - segment total)", "revenue")
    sh.line("rev_others", "Others / unallocated",
            lambda p: (others_sl(p) if p in SLIDE else
                       f"={IS('revenue', p)}-SUM({sh.ref('rev_pe', p)}:{sh.ref('rev_mob', p)})"),
            fc=lambda p: fy26("rev_others") if p == "FY2026" else
            f"={A('others_pct', p)}*SUM({sh.ref('rev_pe', p)}:{sh.ref('rev_mob', p)})",
            note="group revenue less segments")
    tot = lambda p: f"=SUM({sh.ref('rev_pe', p)}:{sh.ref('rev_others', p)})"  # noqa: E731
    sh.line("rev_total", "Group revenue", tot, fc=tot, total=True)

    sh.section("AI data-centre revenue (forecast; decision D1 anchors on Assumptions)")
    sh.line("ai_capex_g", "Hyperscaler / AI data-centre capex growth",
            fc=lambda p: None if p == "FY2026" else f"={A('dc_capex_growth', p)}", fmt=F_PCT, fc_style="link")

    def ai_line(share_key, content_key):
        return lambda p: (f"={A26(share_key)}*{sh.ref('rev_total', p)}" if p == "FY2026" else
                          f"={sh.ref(content_key.replace('content_', 'ai_'), PREV[p])}*(1+{A('dc_capex_growth', p)})"
                          f"*(1+{A(content_key, p)})")
    sh.line("ai_power", "AI power solutions", fc=ai_line("ai_power_share", "content_power"),
            note="FY2026E: ~25% of revenue (mgmt)")
    sh.line("ai_cooling", "Liquid cooling", fc=ai_line("ai_cooling_share", "content_cooling"),
            note="FY2026E: ~13% of revenue (mgmt)")
    sh.line("ai_other", "Other AI data-centre", fc=ai_line("ai_other_share", "content_other"),
            note="FY2026E: ~50% total less power, cooling")
    ait = lambda p: f"=SUM({sh.ref('ai_power', p)}:{sh.ref('ai_other', p)})"  # noqa: E731
    sh.line("ai_total", "AI data-centre revenue", fc=ait, total=True)
    sh.line("ai_share", "AI data-centre % of group revenue",
            fc=lambda p: f"={sh.ref('ai_total', p)}/{sh.ref('rev_total', p)}", fmt=F_PCT, bold=True)
    sh.line("ai_pe", "of which booked in Power Electronics",
            fc=lambda p: f"={sh.ref('ai_power', p)}*{A26('alloc_power_pe')}+{sh.ref('ai_cooling', p)}*"
                         f"{A26('alloc_cooling_pe')}+{sh.ref('ai_other', p)}*{A26('alloc_other_pe')}",
            note="allocation not disclosed (analyst input)")
    sh.line("ai_infra", "of which booked in Infrastructure",
            fc=lambda p: f"={sh.ref('ai_total', p)}-{sh.ref('ai_pe', p)}")

    sh.section("Revenue excluding AI data-centre (forecast)")
    for k, g in (("pe", "g_pe_nonai"), ("infra", "g_infra_nonai")):
        sh.line(f"nonai_{k}", f"{'Power Electronics' if k == 'pe' else 'Infrastructure'} excl. AI",
                fc=lambda p, k=k, g=g: (f"={sh.ref('rev_' + k, p)}-{sh.ref('ai_' + k, p)}" if p == "FY2026" else
                                        f"={sh.ref('nonai_' + k, PREV[p])}*(1+{A(g, p)})"),
                note="FY2026E: segment less allocated AI")
    sh.line("nonai_total", "Group revenue excl. AI data-centre",
            fc=lambda p: f"={sh.ref('rev_total', p)}-{sh.ref('ai_total', p)}", total=True)

    sh.section("Segment operating profit")
    for seg in SEGS:
        k = SEG_KEY[seg]
        sh.line(f"op_{k}", seg, val(seg, "op_profit"),
                fc=lambda p, k=k: f"={sh.ref('rev_' + k, p)}*{A('m_' + k, p)}", style="input",
                note="segment note / slides; fcst rev x margin")
    opt = lambda p: f"=SUM({sh.ref('op_pe', p)}:{sh.ref('op_mob', p)})"  # noqa: E731
    sh.line("op_total", "Total segment profit (= EBIT)", opt, fc=opt, total=True,
            note="Others assumed nil profit")

    sh.section("Revenue growth (y/y)")
    comparable = {"FY2022": SEGS[:3], "FY2023": ["Infrastructure", "Automation"], "FY2024": SEGS,
                  "FY2025": SEGS, "1H26": SEGS}
    for seg in SEGS:
        k = SEG_KEY[seg]
        g = lambda p, k=k: f"={sh.ref('rev_' + k, p)}/{sh.ref('rev_' + k, PREV[p])}-1"  # noqa: E731
        sh.line(f"g_{k}", seg, lambda p, s=seg, g=g: g(p) if s in comparable.get(p, []) else None, fc=g, fmt=F_PCT)
    gt = lambda p: f"={sh.ref('rev_total', p)}/{sh.ref('rev_total', PREV[p])}-1"  # noqa: E731
    sh.line("g_total", "Group", lambda p: gt(p) if p in PREV else None, fc=gt, fmt=F_PCT, bold=True)
    sh.line("g_ai", "AI data-centre revenue",
            fc=lambda p: None if p == "FY2026" else f"={sh.ref('ai_total', p)}/{sh.ref('ai_total', PREV[p])}-1",
            fmt=F_PCT)
    sh.line("g_nonai", "Revenue excl. AI data-centre",
            fc=lambda p: None if p == "FY2026" else f"={sh.ref('nonai_total', p)}/{sh.ref('nonai_total', PREV[p])}-1",
            fmt=F_PCT)

    sh.section("Revenue mix (% of group)")
    for seg in SEGS:
        k = SEG_KEY[seg]
        mx = lambda p, k=k: f"={sh.ref('rev_' + k, p)}/{sh.ref('rev_total', p)}"  # noqa: E731
        sh.line(f"mix_{k}", seg, lambda p, s=seg, mx=mx: None if s == "Mobility" and p in ("FY2021", "FY2022")
                else mx(p), fc=mx, fmt=F_PCT)

    sh.section("Segment margin (profit / revenue)")
    for seg in SEGS:
        k = SEG_KEY[seg]
        m = lambda p, k=k: f"={sh.ref('op_' + k, p)}/{sh.ref('rev_' + k, p)}"  # noqa: E731
        sh.line(f"m_{k}", seg, lambda p, s=seg, m=m: None if s == "Mobility" and p in ("FY2021", "FY2022") else m(p),
                fc=m, fmt=F_PCT)
    mt = lambda p: f"={sh.ref('op_total', p)}/{sh.ref('rev_total', p)}"  # noqa: E731
    sh.line("m_total", "Group EBIT margin", mt, fc=mt, fmt=F_PCT, bold=True)

    # disclosed AI anchors (documentation; the inputs themselves live on Assumptions)
    sh.blank()
    sh.section("AI data-centre disclosures by management (the evidence behind decision D1)")
    hdr = ["Disclosure", "Source", "Period", "Share of revenue", "Qualifier", "Implied NT$ m"]
    for c, t in zip("ABCDEF", hdr):
        ws[f"{c}{sh.r}"] = t
        ws[f"{c}{sh.r}"].font, ws[f"{c}{sh.r}"].fill = font(WHITE, bold=True), FILL_HEAD
    sh.r += 1
    anchors = [
        ("AI data-centre-related businesses (4 BUs)", "3Q26 transcript p.7", "2026", 0.50, "approx.", None),
        ("AI power-related solutions", "3Q26 transcript p.7", "2026", 0.25, "approx.", None),
        ("Liquid cooling (NT$25bn in 3Q26)", "3Q26 transcript p.7", "3Q26", 0.13, "approx.", None),
        ("Liquid cooling", "2Q26 transcript p.12", "FY2025", 0.10, "approx.", "FY2025"),
        ("Data-centre-related business", "2Q26 transcript p.2", "1H26", 0.50, "more than", "1H26"),
    ]
    for name, src, per, share, qual, base in anchors:
        r = sh.r
        ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"], ws[f"E{r}"] = name, src, per, share, qual
        ws[f"D{r}"].font, ws[f"D{r}"].number_format = font(BLUE), F_PCT
        ws[f"D{r}"].comment = Comment(f"Management statement, {src}. Rounded; translated transcript "
                                      "(Mandarin prevails).", "model")
        for c in "ABCE":
            ws[f"{c}{r}"].font = font(size=9)
        if base:
            ws[f"F{r}"] = f"=D{r}*{IS('revenue', base)}"
            ws[f"F{r}"].number_format, ws[f"F{r}"].font = F_NUM, font()
        sh.r += 1
    ws[f"A{sh.r}"] = ("Clarified in 3Q26: the earlier '>25% AI-related' refers to power solutions only. Not used: the "
                      "press-reported FY2025 split (AI power >20%, AI cooling ~9%) - not in Delta's own documents.")
    ws[f"A{sh.r}"].font = font(GREY, italic=True, size=9)
    return sh


def build_schedules(wb):
    sh = Sheet(wb, "Schedules", "Supporting schedules",
               "NT$ million unless stated. Ratios use year-end balances. Historical ratios are the reference for the "
               "forecast drivers on Assumptions; forecast columns apply those drivers.", periods=FY)
    ws = sh.ws
    IS = lambda k, p: X("IS", k, p)  # noqa: E731
    BS = lambda k, p: X("BS", k, p)  # noqa: E731
    CF = lambda k, p: X("CF", k, p)  # noqa: E731
    ws["A4"], ws["B4"] = "Days in year", 365
    ws["B4"].font, ws["B4"].fill = font(BLUE), FILL_KEY
    days = "$B$4"
    hp = lambda f: (lambda p: f(p) if p in PREV else None)  # noqa: E731
    both = lambda f, **kw: dict(fn=f, fc=f, **kw)  # noqa: E731

    sh.section("Working capital")
    sh.line("revenue", "Revenue", **both(lambda p: f"={IS('revenue', p)}"), style="link", fc_style="link")
    sh.line("cogs", "Cost of revenue (positive)", **both(lambda p: f"=-{IS('cogs', p)}"), style="link", fc_style="link")
    for k in ("receivables", "inventories", "payables"):
        sh.line(k, k.capitalize(), **both(lambda p, k=k: f"={BS(k, p)}"), style="link", fc_style="link")
    sh.line("dso", "Receivable days (DSO)", **both(lambda p: f"={sh.ref('receivables', p)}/{sh.ref('revenue', p)}*{days}"),
            fmt=F_DAYS)
    sh.line("dio", "Inventory days (DIO)", **both(lambda p: f"={sh.ref('inventories', p)}/{sh.ref('cogs', p)}*{days}"),
            fmt=F_DAYS)
    sh.line("dpo", "Payable days (DPO)", **both(lambda p: f"={sh.ref('payables', p)}/{sh.ref('cogs', p)}*{days}"),
            fmt=F_DAYS)
    sh.line("ccc", "Cash conversion cycle (days)",
            **both(lambda p: f"={sh.ref('dso', p)}+{sh.ref('dio', p)}-{sh.ref('dpo', p)}"), fmt=F_DAYS, bold=True)
    sh.line("other_ca_pct", "Other current assets % revenue",
            **both(lambda p: f"={BS('other_ca', p)}/{sh.ref('revenue', p)}"), fmt=F_PCT)
    sh.line("other_cl_pct", "Other current liabilities % revenue",
            **both(lambda p: f"={BS('other_cl', p)}/{sh.ref('revenue', p)}"), fmt=F_PCT)

    sh.section("Capex")
    sh.line("capex_total", "Total capex (PP&E + intangibles)", lambda p: f"=-{CF('capex', p)}",
            fc=lambda p: f"={A26('capex_2026')}" if p == "FY2026" else f"={A('capex_pct', p)}*{sh.ref('revenue', p)}",
            style="link", fc_style="link", note="FY2026E: mgmt ~NT$70bn", bold=True)
    sh.line("capex_int", "of which intangibles", lambda p: f"=-{CF('capex_int', p)}",
            fc=lambda p: f"={A('capex_int_pct', p)}*{sh.ref('revenue', p)}", style="link")
    sh.line("capex_ppe", "of which PP&E", lambda p: f"=-{CF('capex_ppe', p)}",
            fc=lambda p: f"={sh.ref('capex_total', p)}-{sh.ref('capex_int', p)}", style="link")
    sh.line("capex_total_pct", "Total capex % revenue",
            **both(lambda p: f"={sh.ref('capex_total', p)}/{sh.ref('revenue', p)}"), fmt=F_PCT, bold=True)

    sh.section("PP&E roll-forward")
    sh.line("ppe_open", "Opening PP&E", hp(lambda p: f"={BS('ppe', PREV[p])}"), fc=lambda p: f"={BS('ppe', PREV[p])}",
            style="link", fc_style="link")
    sh.line("ppe_add", "Add: purchases of PP&E", **both(lambda p: f"={sh.ref('capex_ppe', p)}"))
    sh.line("dep", "Less: depreciation", lambda p: f"=-{H('depreciation', p)}",
            fc=lambda p: f"=-{A('dep_rate', p)}*{sh.ref('ppe_open', p)}", style="link",
            note="hist: all depreciation incl. right-of-use")
    sh.line("ppe_other", "Other movements (plug)",
            hp(lambda p: f"={BS('ppe', p)}-{sh.ref('ppe_open', p)}-{sh.ref('ppe_add', p)}-{sh.ref('dep', p)}"),
            note="disposals, FX, M&A, ROU add-back; fcst nil")
    sh.line("ppe_close", "Closing PP&E", lambda p: f"={BS('ppe', p)}",
            fc=lambda p: f"={sh.ref('ppe_open', p)}+{sh.ref('ppe_add', p)}+{sh.ref('dep', p)}", style="link",
            total=True)
    sh.line("dep_rate", "Depreciation % opening PP&E",
            hp(lambda p: f"=-{sh.ref('dep', p)}/{sh.ref('ppe_open', p)}"),
            fc=lambda p: f"=-{sh.ref('dep', p)}/{sh.ref('ppe_open', p)}", fmt=F_PCT)
    sh.line("da_pct", "D&A % revenue", **both(lambda p: f"={IS('d_and_a', p)}/{sh.ref('revenue', p)}"), fmt=F_PCT)
    sh.line("capex_da", "Capex / D&A", **both(lambda p: f"={sh.ref('capex_total', p)}/{IS('d_and_a', p)}"), fmt=F_X)

    sh.section("Intangibles")
    sh.line("int_open", "Opening intangibles", hp(lambda p: f"={BS('intangibles', PREV[p])}"),
            fc=lambda p: f"={BS('intangibles', PREV[p])}", style="link", fc_style="link")
    sh.line("amort", "Amortisation", lambda p: f"={H('amortisation', p)}",
            fc=lambda p: f"={A('amort_rate', p)}*{sh.ref('int_open', p)}", style="link")
    sh.line("int_close", "Closing intangibles (incl. goodwill)", lambda p: f"={BS('intangibles', p)}",
            fc=lambda p: f"={sh.ref('int_open', p)}+{sh.ref('capex_int', p)}-{sh.ref('amort', p)}", style="link",
            total=True)
    sh.line("amort_rate", "Amortisation % opening intangibles",
            hp(lambda p: f"={sh.ref('amort', p)}/{sh.ref('int_open', p)}"),
            fc=lambda p: f"={sh.ref('amort', p)}/{sh.ref('int_open', p)}", fmt=F_PCT)
    sh.line("capex_int_pct", "Intangibles capex % revenue",
            **both(lambda p: f"={sh.ref('capex_int', p)}/{sh.ref('revenue', p)}"), fmt=F_PCT)

    sh.section("Debt and interest")
    sh.line("debt", "Total debt", **both(lambda p: f"={BS('total_debt', p)}"), style="link", fc_style="link")
    avg = lambda k: (lambda p: f"=AVERAGE({sh.ref(k, PREV[p])},{sh.ref(k, p)})")  # noqa: E731
    sh.line("avg_debt", "Average debt", hp(avg("debt")), fc=avg("debt"))
    sh.line("fin_cost", "Finance costs", **both(lambda p: f"=-{IS('finance_costs', p)}"), style="link",
            fc_style="link", note="incl. lease interest")
    sh.line("cost_debt", "Implied cost of debt (on average debt)",
            hp(lambda p: f"={sh.ref('fin_cost', p)}/{sh.ref('avg_debt', p)}"),
            fc=lambda p: f"={sh.ref('fin_cost', p)}/{sh.ref('avg_debt', p)}", fmt=F_PCT, bold=True)
    sh.line("cash", "Cash and cash equivalents", **both(lambda p: f"={BS('cash', p)}"), style="link", fc_style="link")
    sh.line("avg_cash", "Average cash", hp(avg("cash")), fc=avg("cash"))
    sh.line("int_inc", "Interest income", **both(lambda p: f"={IS('interest_income', p)}"), style="link",
            fc_style="link")
    sh.line("cash_yield", "Yield on average cash", hp(lambda p: f"={sh.ref('int_inc', p)}/{sh.ref('avg_cash', p)}"),
            fc=lambda p: f"={sh.ref('int_inc', p)}/{sh.ref('avg_cash', p)}", fmt=F_PCT)

    sh.section("Tax")
    sh.line("etr", "Effective tax rate", **both(lambda p: f"={IS('etr', p)}"), style="link", fc_style="link",
            fmt=F_PCT)
    sh.line("cash_tax", "Cash taxes paid % PBT", lambda p: f"=-{CF('taxes_paid', p)}/{IS('pbt', p)}",
            fc=lambda p: f"=-{CF('taxes_paid', p)}/{IS('pbt', p)}", fmt=F_PCT)

    sh.section("Shares, dividends and minorities")
    sh.line("shares", "Shares outstanding (m)", **both(lambda p: f"={IS('shares', p)}"), style="link",
            fc_style="link")
    sh.line("div_paid", "Dividends paid in year", **both(lambda p: f"=-{CF('dividends', p)}"), style="link",
            fc_style="link")
    sh.line("dps_paid", "Dividend per share paid in year (NT$)",
            **both(lambda p: f"={sh.ref('div_paid', p)}/{sh.ref('shares', p)}"), fmt=F_EPS)
    sh.line("payout", "Payout of prior-year net income",
            hp(lambda p: f"={sh.ref('div_paid', p)}/{IS('ni_parent', PREV[p])}"),
            fc=lambda p: f"={sh.ref('div_paid', p)}/{IS('ni_parent', PREV[p])}", fmt=F_PCT, bold=True,
            note="Taiwan: dividend paid in year after earned")
    sh.line("nci_div_pct", "Dividends to NCI % prior-year NCI profit",
            hp(lambda p: f"={CF('dividends_nci', p)}/{IS('ni_nci', PREV[p])}"),
            fc=lambda p: f"={CF('dividends_nci', p)}/{IS('ni_nci', PREV[p])}", fmt=F_PCT)
    sh.line("nci_share", "Non-controlling interests % net income",
            **both(lambda p: f"=-{IS('ni_nci', p)}/{IS('net_income', p)}"), fmt=F_PCT)

    sh.section("Returns")
    roe = lambda p: f"={IS('ni_parent', p)}/AVERAGE({BS('equity_parent', PREV[p])},{BS('equity_parent', p)})"  # noqa
    sh.line("roe", "Return on average equity (parent)", hp(roe), fc=roe, fmt=F_PCT, bold=True)
    at = lambda p: f"={sh.ref('revenue', p)}/AVERAGE({BS('ta', PREV[p])},{BS('ta', p)})"  # noqa: E731
    sh.line("asset_turn", "Revenue / average total assets", hp(at), fc=at, fmt=F_X)
    return sh


UNIT_FMT = {"pct": F_PCT, "days": F_DAYS, "ntd_m": F_NUM, "ntd": F_EPS}
UNIT_LABEL = {"pct": "%", "days": "days", "ntd_m": "NT$ m", "ntd": "NT$"}
HIST_REF = {
    "gross_margin": ("IS", "gm"), "rnd_pct": ("IS", "rnd_pct"), "etr": ("IS", "etr"),
    "nci_pct": ("Schedules", "nci_share"), "dso": ("Schedules", "dso"), "dio": ("Schedules", "dio"),
    "dpo": ("Schedules", "dpo"), "capex_pct": ("Schedules", "capex_total_pct"), "dep_rate": ("Schedules", "dep_rate"),
    "amort_rate": ("Schedules", "amort_rate"), "cash_yield": ("Schedules", "cash_yield"),
    "cost_debt": ("Schedules", "cost_debt"), "payout": ("Schedules", "payout"),
    "nci_div_pct": ("Schedules", "nci_div_pct"), "other_ca_pct": ("Schedules", "other_ca_pct"),
    "other_cl_pct": ("Schedules", "other_cl_pct"), "capex_int_pct": ("Schedules", "capex_int_pct"),
    "dps_2026": ("Schedules", "dps_paid"), "capex_2026": ("Schedules", "capex_total"),
    "m_pe": ("Segments", "m_pe"), "m_infra": ("Segments", "m_infra"), "m_auto": ("Segments", "m_auto"),
    "m_mob": ("Segments", "m_mob"), "g_auto": ("Segments", "g_auto"), "g_mob": ("Segments", "g_mob"),
    "g_pe_nonai": ("Segments", "g_pe"), "g_infra_nonai": ("Segments", "g_infra"),
}


def build_assumptions(wb, a: pd.DataFrame):
    sh = Sheet(wb, "Assumptions", "Assumptions",
               "Every forecast driver, by year, with Bear / Base / Bull inputs. The 'Live' row (bold) follows the "
               "scenario in B3 and feeds the model. Source of truth: model/assumptions.csv. Historical columns show "
               "the same ratio as reported, for reference.", periods=FY)
    ws = sh.ws
    ws["A3"], ws["B3"] = "Scenario (select)", "Base"
    ws["B3"].font, ws["B3"].fill = font(BLUE, bold=True), FILL_KEY
    dv = DataValidation(type="list", formula1='"Bear,Base,Bull"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B3")
    ws["A4"], ws["B4"] = "Scenario index (1 Bear, 2 Base, 3 Bull)", '=MATCH(B3,{"Bear","Base","Bull"},0)'
    ws["B4"].font = font()
    ws.column_dimensions["N"].width = 110
    ws["N5"] = "Rationale ([ANALYST TO REVIEW] = draft for the analyst to accept, change or reject)"
    ws["N5"].font, ws["N5"].fill = font(WHITE, bold=True), FILL_HEAD
    ws["N4"] = None
    order = list(dict.fromkeys(a.key))
    sections = list(dict.fromkeys(a.section))
    for section in sections:
        sh.section(section)
        for key in [k for k in order if a[a.key == k].section.iloc[0] == section]:
            rows = a[a.key == key].set_index("scenario")
            unit = rows.unit.iloc[0]
            fmt = UNIT_FMT[unit]
            scale = 0.01 if unit == "pct" else 1.0
            live = sh.r
            first = live + 1
            years = [p for p in FC if not rows[p].isna().all()]
            src = HIST_REF.get(key)

            def hist(p, src=src):
                if src is None:
                    return None
                cell = wb[src[0]][f"{COL[p]}{REG[src]}"].value
                return None if cell is None else f"={ref(src[0], src[1], p)}"
            sh.line(key, rows.label.iloc[0], hist, fc=lambda p, first=first, years=years: (
                f"=CHOOSE($B$4,{COL[p]}{first},{COL[p]}{first + 1},{COL[p]}{first + 2})" if p in years else None),
                fmt=fmt, style="link", bold=True, note=f"live ({UNIT_LABEL[unit]})")
            for i, scen in enumerate(("Bear", "Base", "Bull")):
                r = sh.r
                ws[f"A{r}"] = f"    {scen}"
                ws[f"A{r}"].font = font(GREY if scen != "Base" else BLACK, italic=scen != "Base")
                for p in years:
                    v = rows.at[scen, p]
                    if pd.isna(v):
                        continue
                    c = ws[f"{COL[p]}{r}"]
                    c.value = float(v) * scale
                    c.font, c.number_format = font(BLUE), fmt
                    if scen == "Base":
                        c.fill = FILL_KEY
                ws[f"N{r}"] = rows.at[scen, "rationale"]
                ws[f"N{r}"].font = font(GREY, size=9)
                sh.r += 1
    return sh


def build_checks(wb):
    sh = Sheet(wb, "Checks", "Model integrity checks",
               "Every check must be TRUE. Tolerance in NT$ million (B3) and NT$ per share (B4). The overall result "
               "rolls up to the Cover sheet.")
    ws = sh.ws
    ws["A3"], ws["B3"] = "Tolerance (NT$ m)", 0.01
    ws["A4"], ws["B4"] = "EPS tolerance (NT$)", 0.02
    for c in ("B3", "B4"):
        ws[c].font = font(BLUE)
    ws.column_dimensions[CHECK_ALL].width = 12
    ws[f"{CHECK_ALL}5"] = "All periods"
    ws[f"{CHECK_ALL}5"].font, ws[f"{CHECK_ALL}5"].fill = font(WHITE, bold=True), FILL_HEAD
    tol, etol = "$B$3", "$B$4"
    R = lambda sheet, k, p: X(sheet, k, p)  # noqa: E731
    eq = lambda a, b, t=tol: (lambda p: f"=ABS({a(p)}-{b(p)})<={t}")  # noqa: E731
    rows = []

    def chk(key, label, fn=None, fc=None, periods=None, note="", fc_periods=FC):
        row = sh.line(key, label, fn, fc=fc, fmt="General", note=note, periods=periods)
        cells = [f"{COL[p]}{row}" for p in (periods or PERIODS)] if fn else []
        cells += [f"{COL[p]}{row}" for p in fc_periods] if fc else []
        ws[f"{CHECK_ALL}{row}"] = f"=AND({','.join(cells)})"
        ws[f"{CHECK_ALL}{row}"].font = font(bold=True)
        rows.append(row)

    sh.section("Balance sheet")
    bal = lambda p: f"=ABS({R('BS', 'bs_diff', p)})<={tol}"  # noqa: E731
    chk("bs_balance", "Assets = liabilities + equity", bal, fc=bal)
    for k, h, lab in (("ta", "total_assets", "Total assets"), ("tl", "total_liabilities", "Total liabilities"),
                      ("te", "total_equity", "Total equity")):
        chk(f"bs_{k}", f"{lab} = reported", eq(lambda p, k=k: R("BS", k, p), lambda p, h=h: R("Historicals", h, p)))
    sh.section("Income statement")
    for k, h, lab in (("gross_profit", "gross_profit", "Gross profit"), ("opex", "opex_total", "Operating expenses"),
                      ("ebit", "ebit", "EBIT"), ("nonop", "non_operating_total", "Non-operating total"),
                      ("pbt", "pbt", "Profit before tax"), ("net_income", "net_income", "Net income"),
                      ("ni_parent", "ni_parent", "Net income to parent")):
        chk(f"is_{k}", f"{lab} = reported", eq(lambda p, k=k: R("IS", k, p), lambda p, h=h: R("Historicals", h, p)))
    chk("is_eps", "EPS (computed) = reported basic EPS",
        eq(lambda p: R("IS", "eps", p), lambda p: R("Historicals", "eps_basic", p), etol),
        note="period-end shares vs weighted average")
    sh.section("Cash flow")
    chk("cf_pbt", "PBT in cash flow = PBT in income statement",
        eq(lambda p: R("CF", "pbt", p), lambda p: R("IS", "pbt", p)))
    chk("cf_change", "Net change in cash = reported",
        eq(lambda p: R("CF", "net_change", p), lambda p: R("Historicals", "net_change_cash", p)))
    chk("cf_cash", "Closing cash (CF) = cash on balance sheet",
        eq(lambda p: R("CF", "cash_close", p), lambda p: R("BS", "cash", p)),
        fc=eq(lambda p: R("CF", "cash_close", p), lambda p: R("BS", "cash", p)))
    sh.section("Segments")
    seg_ebit = eq(lambda p: R("Segments", "op_total", p), lambda p: R("IS", "ebit", p))
    chk("seg_ebit", "Segment profit total = EBIT", seg_ebit, fc=seg_ebit)
    rev_eq = eq(lambda p: R("Segments", "rev_total", p), lambda p: R("IS", "revenue", p))
    chk("seg_rev", "Segment revenue total = IS revenue", rev_eq, fc=rev_eq)
    chk("seg_others", "Unallocated revenue < 0.5% of group revenue",
        lambda p: f"=ABS({R('Segments', 'rev_others', p)})<0.005*{R('IS', 'revenue', p)}",
        fc=lambda p: f"=ABS({R('Segments', 'rev_others', p)})<0.005*{R('IS', 'revenue', p)}")
    sl_rows = [REG[("Segments", f"{x}_{k}")] for x in ("rev", "op") for k in ("pe", "infra", "auto", "mob")]
    chk("seg_slides", "3Q26 slides: 1H26 (audited) + 3Q26 = 9M26, every segment",
        lambda p: "=AND(" + ",".join(f"ABS('Segments'!O{r}+'Segments'!P{r}-'Segments'!Q{r})<=2" for r in sl_rows) + ")",
        periods=["1H26"], note="slides in NT$ m: tolerance NT$2m")
    sh.section("Forecast sanity")
    chk("fc_nonai", "Non-AI revenue stays positive (PE and Infrastructure)",
        fc=lambda p: f"=AND({R('Segments', 'nonai_pe', p)}>0,{R('Segments', 'nonai_infra', p)}>0)")
    chk("fc_cash", "Cash stays positive", fc=lambda p: f"={R('BS', 'cash', p)}>0")
    chk("fc_sga", "Implied SG&A between 5% and 15% of revenue",
        fc=lambda p: f"=AND({R('IS', 'sga_pct', p)}>=0.05,{R('IS', 'sga_pct', p)}<=0.15)",
        note="SG&A balances GP and segment EBIT")
    chk("fc_q4", "4Q26E revenue above 3Q26 (management guidance)",
        fc=lambda p: (f"={A26('q4_revenue')}>'Segments'!P{REG[('Segments', 'rev_total')]}" if p == "FY2026"
                      else None), note="3Q26 transcript p.6", fc_periods=["FY2026"])
    sh.section("Source extraction (Python)")
    chk("py_checks", "All tie-outs in data/processed/checks.csv pass",
        lambda p: f"={R('Historicals', 'checks_passed', 'FY2021')}={R('Historicals', 'checks_total', 'FY2021')}",
        periods=["FY2021"], note="subtotals, restatements, slides, monthly")
    sh.blank()
    r = sh.r
    ws[f"A{r}"] = "ALL CHECKS"
    ws[f"A{r}"].font = font(NAVY, bold=True, size=12)
    ws[f"{CHECK_ALL}{r}"] = f"=AND({','.join(f'{CHECK_ALL}{x}' for x in rows)})"
    ws[f"{CHECK_ALL}{r}"].font = font(bold=True, size=12)
    REG[("Checks", "all")] = r
    tf_format(ws, f"C8:{CHECK_ALL}{r}", "C8")
    return sh


def build_cover(wb):
    ws = wb.create_sheet("Cover", 0)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 30
    ws["B2"] = "Delta Electronics, Inc. (2308.TW)"
    ws["B2"].font = font(NAVY, bold=True, size=18)
    ws["B3"] = "Initiation of coverage - financial model"
    ws["B3"].font = font(GREY, size=12)
    ws["B4"] = "Student research for educational purposes only. Not investment advice."
    ws["B4"].font = font(GREY, italic=True, size=9)
    items = [
        ("Company", "Delta Electronics, Inc."),
        ("Ticker / exchange", "2308 / Taiwan Stock Exchange"),
        ("Units", "NT$ million unless stated (reported in NT$ thousands, Taiwan IFRS)"),
        ("Latest reported period", "9M26 (3Q26 earnings slides, 29 Oct 2026); latest full statements 1H26"),
        ("Forecast basis", "FY2026E = 9M26 actual + 4Q26E; balance sheet rolls forward from FY2025A"),
        ("Share price and date", "Set in M4 (yfinance, 2308.TW)"),
        ("Rating", "[ANALYST TO REVIEW] - set in M5"),
        ("12-month price target", "[ANALYST TO REVIEW] - set in M4/M5"),
        ("Scenario", "='Assumptions'!B3"),
        ("Model status", "M3 - forecasts FY2026E-FY2030E; draft assumptions [ANALYST TO REVIEW]"),
        ("Last updated", dt.date.today().isoformat()),
    ]
    r = 6
    for k, v in items:
        ws[f"B{r}"], ws[f"C{r}"] = k, v
        ws[f"B{r}"].font = font(bold=True)
        ws[f"C{r}"].font = font(GREEN if str(v).startswith("=") else BLACK)
        r += 1
    r += 1
    ws[f"B{r}"] = "ALL CHECKS PASS"
    ws[f"B{r}"].font = font(NAVY, bold=True, size=12)
    ws[f"C{r}"] = f"='Checks'!{CHECK_ALL}{REG[('Checks', 'all')]}"
    ws[f"C{r}"].font = font(bold=True, size=12)
    ws[f"C{r}"].alignment = Alignment(horizontal="left")
    tf_format(ws, f"C{r}", f"C{r}")
    REG[("Cover", "checks")] = r
    r += 2
    ws[f"B{r}"] = "Key financials (NT$ m; live scenario)"
    ws[f"B{r}"].font = font(NAVY, bold=True)
    r += 1
    heads = ["FY2024", "FY2025", "FY2026", "FY2027", "FY2028"]
    key_rows = [("Revenue", "IS", "revenue", F_NUM), ("Revenue growth", "IS", "rev_growth", F_PCT),
                ("AI data-centre % of revenue", "Segments", "ai_share", F_PCT),
                ("EBIT margin", "IS", "ebit_margin", F_PCT), ("EPS, basic (NT$)", "IS", "eps", F_EPS),
                ("EPS growth", "IS", "eps_growth", F_PCT), ("Dividend per share paid (NT$)", "Schedules", "dps_paid", F_EPS),
                ("Free cash flow", "CF", "fcf", F_NUM), ("Net debt / (net cash)", "BS", "net_debt", F_NUM)]
    for i, p in enumerate(heads):
        c = ws.cell(row=r, column=4 + i, value=HEAD[p])
        c.font, c.alignment = font(WHITE, bold=True), Alignment(horizontal="right")
        c.fill = FILL_FC_HEAD if p in FC else FILL_HEAD
        ws.column_dimensions[c.column_letter].width = 12
    ws[f"B{r}"].fill = ws[f"C{r}"].fill = FILL_HEAD
    r += 1
    for label, sheet, key, fmt in key_rows:
        ws[f"B{r}"] = label
        ws[f"B{r}"].font = font()
        for i, p in enumerate(heads):
            cell_ref = f"{COL[p]}{REG[(sheet, key)]}"
            if wb[sheet][cell_ref].value is None and p not in FC:
                continue
            c = ws.cell(row=r, column=4 + i, value=f"='{sheet}'!{cell_ref}")
            c.font, c.number_format = font(GREEN), fmt
        r += 1
    r += 1
    ws[f"B{r}"] = "Sheet guide"
    ws[f"B{r}"].font = font(NAVY, bold=True)
    r += 1
    guide = [("Assumptions", "Scenario selector; every forecast driver by year, Bear / Base / Bull"),
             ("Historicals", "Extracted as-reported figures (inputs), FY2021A-FY2025A, 1H25A, 1H26A"),
             ("Segments", "Segment revenue build incl. AI data-centre line; segment profit"),
             ("IS / BS / CF", "Three statements: historicals by formula, FY2026E-FY2030E forecast"),
             ("Schedules", "Working capital, capex, PP&E, intangibles, debt, tax, dividends, returns"),
             ("Checks", "Integrity tests; all must be TRUE")]
    for k, v in guide:
        ws[f"B{r}"], ws[f"C{r}"] = k, v
        ws[f"B{r}"].font, ws[f"C{r}"].font = font(bold=True), font()
        r += 1
    r += 1
    ws[f"B{r}"] = "Colour key"
    ws[f"B{r}"].font = font(NAVY, bold=True)
    r += 1
    for txt, col, fill in (("Blue: hard-coded input (sourced in data/sources.csv or model/assumptions.csv)", BLUE, None),
                           ("Black: formula", BLACK, None), ("Green: link to another sheet", GREEN, None),
                           ("Yellow fill: Base-case input set by the analyst", BLACK, FILL_KEY),
                           ("Shaded columns: forecast", BLACK, FILL_FC)):
        ws[f"B{r}"] = txt
        ws[f"B{r}"].font = font(col)
        if fill:
            ws[f"B{r}"].fill = fill
        r += 1
    ws.sheet_view.showGridLines = False


def main() -> int:
    REG.clear()
    PENDING.clear()
    h, segs, lines, checks, a = load()
    wb = Workbook()
    wb.remove(wb.active)
    build_historicals(wb, h, checks)
    build_is(wb)
    build_bs(wb)
    build_cf(wb)
    build_segments(wb, segs)
    build_schedules(wb)
    build_assumptions(wb, a)
    build_checks(wb)
    finalize()
    build_cover(wb)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_options.gridLines = False
        ws.oddFooter.center.text = "Delta Electronics (2308.TW) model - &A - page &P of &N"
    order = ["Cover", "Assumptions", "Historicals", "Segments", "IS", "BS", "CF", "Schedules", "Checks"]
    wb._sheets = [wb[n] for n in order]
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = {"Cover": NAVY, "Assumptions": "FFC000", "Checks": "70AD47"}.get(ws.title)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(OUT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
