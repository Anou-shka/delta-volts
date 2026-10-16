# Project spec: Delta Electronics (2308.TW) initiation of coverage — "The AI power play"

Paste this whole file into Claude Code as the first message of the project, then save it in the repo as `SPEC.md`.

---

## 0. Who I am and how we work

I'm Anoushka, a CS + Business graduate (NTU) and CFA Level II candidate, building this for my portfolio and for interviews in equity research and AI investment research. I interned at Delta Electronics Int'l (Singapore) in 2022 as a process analyst.

**Working rules (follow these throughout):**

1. **You build the machinery; I make the investment calls.** Every forecast assumption lives in one assumptions sheet that I set. Where the report needs a judgment (rating, thesis, key debate, risks), write a clearly marked draft labelled `[ANALYST TO REVIEW]` and explain your reasoning in one line, so I can accept, change or reject it.
2. **Every number has a source.** Any figure taken from a filing, a website or a dataset gets an entry in `data/sources.csv` (figure, value, unit, document, page or URL, date accessed). If you can't source a number, don't invent it: mark it `TODO-SOURCE` and tell me.
3. **Public information only.** Use only published filings, investor presentations, earnings call materials and public data. Never use or ask about anything from my internship.
4. **Teach as you go.** At the end of each milestone, give me a short summary: what you built, the 3 most important modelling choices, and 3 questions an interviewer could ask me about them. Put these in `docs/milestone-notes.md`.
5. **Ask before guessing** on anything that changes the valuation by more than ~5%.
6. Stop at the end of each milestone and wait for me to say "continue".

---

## 1. The deliverables

| # | Deliverable | Format |
|---|---|---|
| 1 | Linked 3-statement model with segment-driven revenue, DCF, comps and scenarios | Excel (`model/delta_model.xlsx`) with **live formulas**, built with openpyxl |
| 2 | Python mirror of the model that recomputes key outputs and checks the Excel | `src/` package + `tests/` |
| 3 | Initiation report, 10–12 pages, with rating and 12-month price target | PDF (`report/delta_initiation.pdf`) built from Markdown + charts |
| 4 | Charts: revenue mix, margins, AI exposure, valuation football field, sensitivity heatmap | PNG in `report/figures/` |
| 5 | README for GitHub + a one-page case study for my website | `README.md`, `docs/case-study.md` |
| 6 | Interview defence sheet: likely questions, with space for my answers | `docs/defence.md` |

---

## 2. Starting context (verify everything against the filings)

From a Taipei Times report on the Q4 2025 results (27 Feb 2026), to be verified against Delta's own filings before use:

- FY2025 revenue about NT$554.9bn, up ~32% y/y. Q4 2025 revenue NT$161.6bn, net profit NT$17.3bn, EPS NT$6.67, gross margin 34.6%, operating margin 16.35%.
- Q4 2025 revenue mix: Power Electronics ~50%, Infrastructure ~37%, Automation ~9%, Mobility ~4%.
- AI-related power supplies were more than 20% of 2025 revenue; AI-related cooling about 9%.
- 2025 capex about NT$46.1bn, planned to rise in 2026. New Thailand plants started in Q4 2025; expansion planned in Thailand, India and the US. EV demand expected to stay weak for 1–2 years.

Use the **latest reported period available today** (at least FY2025 and 1H/Q2 2026; check whether Q3 2026 results are out and include them if so). Note in the report which period the model is based on.

Reported figures are in **NT$ thousands** under Taiwan IFRS. Present the model in **NT$ millions**, and keep a unit note on every sheet.

---

## 3. Data sources

| Data | Where | Notes |
|---|---|---|
| Annual reports, quarterly consolidated financial statements, investor presentations, earnings call slides | Delta investor relations site (deltaww.com → Investors) and TWSE / MOPS | Primary source for all historicals |
| Share price, shares outstanding, market cap, price history | yfinance (`2308.TW`) | Record the price date used for the target |
| Peer financials and multiples | yfinance for Vertiv (VRT), Schneider Electric (SU.PA), Eaton (ETN), Lite-On Technology (2301.TW), AcBel Polytech (6282.TW); filings where needed | Convert to a common currency; state FX rates and dates |
| Taiwan 10-year government bond yield | Central Bank of Taiwan or a market data source | Risk-free rate input |
| Equity risk premium, country risk premium, beta references | Damodaran Online datasets | Cite the dataset and date |
| Industry context (AI server rack power, data-centre capex) | Hyperscaler earnings releases, Delta and peer investor materials | Cite every claim |

**Network check first:** before anything else, test whether you can reach Delta's IR site, MOPS and Yahoo Finance from this environment. If any download fails, **stop and give me a list of exact files to download manually**. I'll put them in `data/raw/`. Don't use mirrors or scraped copies.

---

## 4. Repo structure

```
delta-initiation/
├── SPEC.md
├── README.md
├── data/
│   ├── raw/            # original PDFs and downloads, untouched
│   ├── processed/      # cleaned CSVs of historicals
│   └── sources.csv     # every sourced figure
├── src/delta/
│   ├── extract.py      # pull historicals from filings into CSVs
│   ├── model.py        # forecast engine (Python mirror)
│   ├── valuation.py    # DCF, comps, scenarios
│   ├── build_excel.py  # writes model/delta_model.xlsx with formulas
│   └── charts.py
├── model/delta_model.xlsx
├── report/
│   ├── delta_initiation.md
│   ├── figures/
│   └── delta_initiation.pdf
├── docs/
│   ├── milestone-notes.md
│   ├── defence.md
│   └── case-study.md
└── tests/
```

---

## 5. The model (Excel, live formulas)

### Sheets

1. **Cover**: company, ticker, price and date, rating, target, scenario selector, units, last updated.
2. **Assumptions**: every forecast driver, by year, with **Bear / Base / Bull** columns and a selector cell. Blue font for inputs, black for formulas (banking convention). One-line rationale column beside each driver.
3. **Historicals**: 5 years (FY2021–FY2025) plus the latest interim period, as reported.
4. **Segments**: revenue build by segment (below).
5. **IS / BS / CF**: income statement, balance sheet, cash flow statement; historical + 5-year forecast (FY2026E–FY2030E).
6. **Schedules**: working capital, PP&E and capex, D&A, debt and interest, tax, dividends, share count.
7. **DCF**: FCFF, WACC build, terminal value (Gordon growth + exit multiple cross-check), equity bridge.
8. **Comps**: peer table with EV/EBITDA, EV/Sales and P/E on trailing and forward (NTM) basis; implied Delta value.
9. **Sensitivity**: WACC × terminal growth grid; AI revenue growth × gross margin grid.
10. **Football field**: valuation ranges from DCF, comps and 52-week range.
11. **Checks**: balance sheet balances every year, cash on BS = cash in CF, historicals match reported to within rounding. All checks show TRUE/FALSE and roll up to one cell on Cover.

### Revenue build (the heart of the thesis)

- Model the four reported segments: **Power Electronics, Infrastructure, Automation, Mobility**.
- Within them, split out an **AI data-centre line** (AI power supplies + AI cooling, using the disclosed shares of revenue) from the rest, so the model can answer "how much of Delta is an AI bet?"
- Drive the AI line top-down from a short chain I can change: hyperscaler data-centre capex growth → AI server/rack shipments → Delta's content per rack (power shelf, power distribution, liquid cooling) → Delta's share. Where public data can't support a link, simplify to growth rates and say so.
- Non-AI lines: growth rates linked to simple, stated drivers (e.g. industrial cycle for Automation, EV demand for Mobility).

### Rest of the model

- Margins by segment where disclosed, otherwise at group level, with a mix effect from AI growth.
- Working capital via days (receivables, inventory, payables) from historical averages.
- Capex as % of revenue with an explicit 2026 step-up (management guided higher capex). D&A from the PP&E schedule.
- Dividends from payout ratio history. Share count flat unless filings show changes.
- No circular references. If interest on cash/debt would create one, use the prior-year balance.

### Valuation

- **WACC:** risk-free (Taiwan 10-year), equity risk premium and beta (cite Damodaran or a regression on 2308.TW vs TAIEX, show both), cost of debt from interest expense / average debt, target capital structure.
- **DCF:** 5-year FCFF + terminal value. Show the share of value coming from the terminal value.
- **Comps:** forward EV/EBITDA and P/E. Explain why each peer is or isn't a fair comparison (Vertiv and Schneider = data-centre power pure plays; Lite-On and AcBel = Taiwan power-supply peers).
- **Price target:** a stated blend of DCF and comps. Rating rule: Buy if upside > 15%, Sell if downside > 10%, otherwise Hold. The rating and target are `[ANALYST TO REVIEW]`.
- **Reverse DCF:** what AI revenue growth does the current share price imply? This is a key interview answer.

### Python mirror

`src/delta/model.py` recomputes revenue, EBIT, FCFF, DCF value per share and the balance check from the same assumptions. A test (`tests/test_model_matches_excel.py`) recalculates the Excel (use LibreOffice headless to recalculate) and asserts the Excel and Python outputs agree within 0.1%.

---

## 6. The report (10–12 pages)

1. **Page 1 — summary:** rating, target, upside, price and date; three thesis bullets; key financials table (FY2024A–FY2028E: revenue, EBIT margin, EPS, P/E, EV/EBITDA, dividend yield); one chart.
2. **Investment thesis:** 3 points, each with evidence and what would prove it wrong.
3. **Key debate:** the single question the market is arguing about (e.g. how durable AI power demand is after the current hyperscaler capex cycle), with both sides.
4. **Company overview:** segments, customers, geography, manufacturing footprint.
5. **Industry:** why AI racks need far more power and cooling per rack, and where Delta sits in that chain. Every claim cited.
6. **Segment forecasts:** drivers and how the Bear/Base/Bull cases differ.
7. **Financials:** margins, cash flow, capex cycle, balance sheet strength.
8. **Valuation:** DCF, comps, football field, sensitivity, reverse DCF.
9. **Risks:** at least 5 (customer concentration, AI capex slowdown, tariffs and US–China/Taiwan geopolitics, margin pressure from competition, execution on capacity expansion, FX), each with an indicator to watch.
10. **Appendix:** full financial statements, assumptions table, sources.

Writing style: plain, specific, numbers in every paragraph, no hype. Footnote every sourced figure. Add the disclaimer: *"Student research for educational purposes only. Not investment advice."*

Build the PDF from Markdown with a clean template (WeasyPrint or Typst, whichever installs cleanly), A4, page numbers, a header with ticker and date.

---

## 7. Milestones

| Day | Milestone | Done when |
|---|---|---|
| Oct 17 | **M1. Data:** network check, collect filings, extract 5 years of historicals + latest interim into CSVs, build `sources.csv` | Historicals tie to the filings' totals; I've reviewed a sample of 10 figures |
| Oct 18 | **M2. Historical model:** Excel sheets 1–6 with historicals, schedules, checks passing | All checks TRUE for historical years |
| Oct 19 | **M3. Forecasts:** segment and AI revenue build, Bear/Base/Bull assumptions (drafted for my review), full 3-statement forecast | Checks TRUE for all forecast years; I've set the Base assumptions |
| Oct 20 | **M4. Valuation:** WACC, DCF, comps, sensitivity, football field, reverse DCF, Python mirror + test | Excel and Python agree; test passes |
| Oct 21 | **M5. Report:** charts and full draft with `[ANALYST TO REVIEW]` sections | I've written the thesis, debate and rating sections myself |
| Oct 22 | **M6. Polish:** PDF, README, case study, defence sheet, final checks | PDF renders cleanly; repo runs from a fresh clone with one command |

---

## 8. Quality bar

- `make all` (or one script) rebuilds everything from `data/raw/`: CSVs → Excel → charts → PDF.
- Excel opens with no errors and no hard-coded numbers in forecast cells (inputs only on Assumptions).
- Every check on the Checks sheet is TRUE.
- Every sourced number in the report appears in `sources.csv`.
- README covers: what it is, key findings (rating, target, reverse-DCF result), method, how to run it, limitations.
- `docs/defence.md` lists at least 25 likely interview questions grouped by topic (business, forecasts, valuation, risks, AI thesis, model mechanics), with an empty answer space under each for me. Don't write the answers.

---

## 9. Start now

Begin with **M1**: run the network check, tell me what you can and can't reach, then propose the exact list of filings you'll use before downloading anything.
