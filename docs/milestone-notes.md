# Milestone notes

## M1 — Data (16 Oct 2026)

### What was built

- **`src/delta/fetch.py`** downloads 39 primary documents from Delta's IR file centre into `data/raw/`. Each download is logged in `data/raw/manifest.csv` with its URL, access date and SHA-256 hash.
- **`src/delta/pdftable.py`** rebuilds statement rows from word coordinates on each PDF page. It doesn't trust the PDF text order, because a value can print one line below its label and brackets for negatives are separate glyphs.
- **`src/delta/extract.py`** produces the processed data:
  - every as-reported line (`fs_lines.csv`, 2,360 values);
  - standardised model lines in NT$ millions (`historicals.csv`, FY2021–FY2025 plus 1Q/2Q/1H for 2025 and 2026);
  - the segment note (`segments.csv`);
  - monthly revenue to Sep 2026 (`monthly_revenue.csv`);
  - 422 tie-out checks (`checks.csv`), **all passing**;
  - `sources.csv`, with 920 rows: one per figure, with page and URL.
- **`tests/test_extract.py`** reruns the extraction and asserts that every check passes and that six known values match.
- `make data` rebuilds everything from scratch.

**Model base period:** Q2 2026 (six months to 30 Jun 2026), the latest audited/reviewed report. Q3 2026 results are not out; Jul–Sep monthly revenue gives unaudited Q3 2026 revenue of NT$197,059m (+31% y/y).

### What the data says (headline numbers, NT$)

| | FY2021 | FY2022 | FY2023 | FY2024 | FY2025 | 1H26 |
|---|---:|---:|---:|---:|---:|---:|
| Revenue (bn) | 314.7 | 384.4 | 401.2 | 421.1 | 554.9 | 342.6 |
| Gross margin | 28.7% | 28.8% | 29.2% | 32.4% | 34.3% | 36.3% |
| EBIT margin | 10.0% | 10.8% | 10.2% | 11.3% | 15.1% | 17.2% |
| EPS (basic) | 10.32 | 12.58 | 12.86 | 13.56 | 23.14 | 17.59 |

FY2025 Infrastructure revenue grew 82% to NT$182.0bn, and its segment profit rose from NT$5.5bn to NT$28.4bn. That segment is where most of the AI data-centre story shows up.

### The 3 most important modelling choices

1. **Extract by machine and prove it with tie-outs, instead of typing numbers in.** Every subtotal is re-added from its detail lines. The three statements are linked (PBT, closing cash). Quarters sum to half-years, segments sum to EBIT, monthly releases sum to quarters, and each comparative matches the prior year's report. Standardised lines such as "other current assets" are computed as residuals against reported totals, so the standardised balance sheet ties by construction.
2. **Restatement and segment policy.**
   - Each fiscal year comes from its own audited report. No statement restatements were found.
   - Segments are different. Delta reported three segments until the FY2023 report, then carved Mobility out of Power Electronics, restating FY2023, and restated FY2024 again in the FY2025 report.
   - I use the latest restated basis for FY2023 onward. FY2021–FY2022 stay on the 3-segment basis, so the model can't show a clean Mobility history before 2023.
   - Segment revenue is 0.1–0.2% below group revenue; that gap is Delta's "Others" bucket and is kept as an explicit line.
3. **Definitions that differ from Delta's own presentation.**
   - **EBITDA = EBIT + D&A** (NT$111.8bn FY2025), not Delta's slide figure of PBT + finance costs + D&A (NT$117.9bn), which includes non-operating income. The operating definition is the right one for EV/EBITDA comps.
   - **Capex = cash paid for PPE + intangibles** (NT$47.2bn FY2025). Management quoted NT$46.6bn and the Taipei Times NT$46.1bn (PPE + land prepayments). I'll use the cash-flow definition and footnote the difference.
   - **2Q cash flows = 1H − 1Q**, because Delta only reports year-to-date cash flows.

### 3 questions an interviewer could ask

1. "Your FY2023 segment numbers differ between the 2023 and 2024 annual reports. Which did you use and why?"
2. "Delta says its FY2025 EBITDA was NT$117.9bn, but your model says NT$111.8bn. Which is right for valuation?"
3. "How do you know your extracted historicals are correct?"

### Decisions taken after M1 (see `docs/decisions.md`)

- **D1:** the AI line is anchored only on Delta's own disclosures (2Q26 transcript):
  - AI-related products >25% of FY2026 revenue;
  - liquid cooling ~10% of FY2025 revenue and >12% of FY2026;
  - data-centre business >50% of 1H26 revenue.

  The press split (AI power >20%, AI cooling ~9%) stays `TODO-SOURCE` and isn't used.
- **D2:** wait for Q3 2026 results (expected ~29–30 Oct 2026) before locking the base period. It will then move from 1H26 to 9M26.

### Still open

- Review the 10-figure sample in `docs/m1-review-sample.md`.
- **Management statements worth knowing for M3** (2Q26 transcript):
  - FY2026 capex about NT$70bn, with NT$27.5bn already spent in 1H.
  - New AI equipment has been depreciated over 3 years since about April 2026 (it was 3–5 years). This raises D&A and lowers reported margins.
  - Gross margin "around 35% is reasonable".
  - Tax rate about 22%.
  - HVDC mass production starts in 3Q26, with meaningful volume in 2027.

---

## M2 — Historical model (16 Oct 2026)

### What was built

- **`src/delta/build_excel.py`** writes `model/delta_model.xlsx` with openpyxl. It has nine sheets: Cover, Assumptions, Historicals, Segments, IS, BS, CF, Schedules and Checks.
- **Size and checks:** 1,195 live formulas and zero formula errors. All 18 checks are TRUE for FY2021A–FY2025A, 1H25A and 1H26A, and they roll up to one cell on the Cover.
- **`src/delta/recalc.py`** recalculates the workbook in headless LibreOffice and scans every formula for errors. **`src/delta/preview.py`** renders sheets to PNG for the README.
- **`tests/test_excel.py`** (runs in CI) checks four things:
  - the workbook recalculates cleanly;
  - every check is TRUE;
  - five key values match the filings;
  - **no hard-coded number appears on the IS, BS, CF or Schedules sheets**.
- `make model` rebuilds it; `make all` runs the whole pipeline.

### What the schedules show (FY2021 → FY2025)

| | FY2021 | FY2022 | FY2023 | FY2024 | FY2025 |
|---|---:|---:|---:|---:|---:|
| Receivable days | 82 | 83 | 72 | 78 | 80 |
| Inventory days | 107 | 106 | 98 | 108 | 102 |
| Payable days | 89 | 84 | 69 | 89 | 95 |
| Total capex % revenue | 7.7% | 5.8% | 7.1% | 8.3% | 8.5% |
| Depreciation % opening PP&E | – | 19.6% | 20.4% | 21.5% | 20.8% |
| Effective tax rate | 20.0% | 19.7% | 20.1% | 21.3% | 22.7% |
| Payout of prior-year earnings | – | 53% | 78% | 50% | 52% |
| Return on average equity | – | 19% | 17% | 16% | 24% |

### The 3 most important modelling choices

1. **One input layer; everything else is a formula.**
   - The only hard-coded financials are on Historicals (blue), and each one is traceable to a page through `sources.csv`.
   - The statements link to those inputs and **recompute** every subtotal rather than copying the reported one. That lets the Checks sheet compare each recomputed subtotal with the reported figure.
   - Each "other" line is anchored to a *different* reported subtotal (current assets, non-current assets, current liabilities and so on). So *assets = liabilities + equity* is a genuine test, not true by construction.
   - A test fails the build if a number is ever typed into a statement sheet.
2. **Ratio definitions that follow how Delta actually operates.**
   - Working-capital days use year-end balances and a 365-day year, kept as an input.
   - **Payout is measured against the prior year's earnings**, because Taiwanese companies pay dividends in the year after they're earned (Delta pays in Q3). That's why 1H cash flow shows no dividend to shareholders.
   - It's also why June balance sheets are distorted. Other current liabilities jump from NT$97.9bn (Dec-25) to NT$165.2bn (Jun-26), most likely the declared-but-unpaid NT$11.6/share dividend (about NT$30bn). I'll confirm against the Q2 2026 notes in M3. **So interim balance sheets aren't used to set working-capital assumptions.**
3. **Interim periods sit beside the annual series, not inside it.** 1H25A and 1H26A get their own columns to the right, so forecasts can run FY2025A → FY2026E with consistent formulas. Share count = share capital / NT$10 par: 2,597.5m, unchanged since FY2021. Computed EPS matches reported basic EPS within NT$0.02 in every period, which shows there are no material buybacks or issuance to model.

### 3 questions an interviewer could ask

1. "Why does your balance-sheet check prove anything if you built the balance sheet from the reported numbers?"
2. "Delta's payable days rose from 69 to 95 between FY2023 and FY2025. Is that sustainable, and how will you forecast it?"
3. "Why measure payout against last year's earnings rather than this year's?"

### Still open

- Q3 2026 results (decision D2): the forecast base period moves to 9M26 when they're published (~29–30 Oct 2026).
- The 10-figure review in `docs/m1-review-sample.md`.
- Confirm the June 2026 dividend payable in the Q2 2026 notes (M3).
