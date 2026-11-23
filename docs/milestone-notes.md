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

---

## M3 — Forecasts (16 Nov 2026) — *draft, waiting for the analyst to set the Base case*

### What was built

- **Q3 2026 data.** The 3Q26 slides, the 3Q26 transcript and October 2026 revenue were added to the pipeline. 3Q26/9M26 segment and group figures are in `data/slides_3q26.csv`. They tie to the audited 1H26 segment note to within NT$1m for every segment, and the monthly releases tie to Q3 revenue. That makes 435 extraction checks, all passing.
- **`model/assumptions.csv`:** 40 drivers × Bear / Base / Bull × FY2026E–FY2030E, each with a rationale. Drafts are tagged `[ANALYST TO REVIEW]`. This one file feeds the Assumptions sheet, and the scenario selector (Assumptions!B3) switches the whole model.
- **Workbook:** forecast columns FY2026E–FY2030E on Segments, IS, BS, CF and Schedules, built from formulas only. Size: 2,292 formulas and zero errors. **Every check is TRUE in all three scenarios**, including *assets = liabilities + equity* in each forecast year, with cash as the balancing item from the cash flow statement.
- **Tests (19, all passing in CI):**
  - Bear and Bull both balance;
  - the assumptions file is complete;
  - there are no typed-in numbers in forecast cells.
- **Chart:** `report/figures/revenue_build.png` shows the Base-case revenue build.

### How revenue is built

1. **FY2026E = 9M26 actual + 4Q26E.** The Q4 estimate is split by the 3Q26 segment mix. Management guided Q4 revenue above Q3's NT$197.1bn, and October alone was NT$74.7bn.
2. **AI data-centre line**, anchored at FY2026E on management's 3Q26 statements: about 50% of revenue, made up of AI power about 25%, liquid cooling about 13% and other AI data-centre about 12%. From FY2027E each line grows at (1 + AI data-centre capex growth) × (1 + Delta's content/share gain) − 1.
3. **AI allocation:** each AI line is split between Power Electronics and Infrastructure. Delta doesn't disclose this, so the split is an analyst input.
4. **Non-AI revenue:** the remainder of Power Electronics and Infrastructure grows with its own driver. Automation and Mobility grow on industrial-cycle and EV assumptions.
5. **EBIT** = Σ segment revenue × segment margin. Gross margin and R&D are inputs; SG&A balances.

### Draft outputs (before the analyst's review)

| | Bear | Base | Bull |
|---|---:|---:|---:|
| FY2026E revenue (NT$bn) | 739.7 | 757.7 | 767.7 |
| FY2026E EPS (NT$) | 32.24 | 36.72 | 39.33 |
| FY2028E revenue (NT$bn) | 767.7 | 1,051.9 | 1,280.9 |
| FY2030E revenue (NT$bn) | 778.8 | 1,257.6 | 1,758.0 |
| FY2030E AI data-centre % of revenue | 52% | 63% | 70% |
| FY2030E EBIT margin | 14.2% | 17.9% | 19.7% |
| FY2030E EPS (NT$) | 29.68 | 62.63 | 98.10 |

### The 3 most important modelling choices

1. **Calibrate the current year to actuals, not to a growth rate.** Three quarters of FY2026 are already reported, so FY2026E is 9M26 actual plus one Q4 assumption. That leaves one judgement for the current year, and management has bounded it (Q4 above Q3).
2. **A two-link AI chain instead of a false-precision unit model.** Delta doesn't disclose rack units or content per rack. A chain of shipments × content × share would have been mostly invented numbers. Capex growth (which management quotes) times a content/share gain keeps the economic logic (Delta grows with hyperscaler spend, plus or minus share and content) while keeping each assumption visible and debatable.
3. **Segment margins drive EBIT; cash balances the balance sheet.** Margins are set where Delta discloses profit (by segment), so AI mix shows up in group margin automatically. Debt and other long-term items are held flat, working capital follows days and PP&E follows capex and depreciation. Cash is whatever the cash flow statement leaves, which makes the balance-sheet check a real test of the model's plumbing.

### 3 questions an interviewer could ask

1. "Your Base case has AI data-centre revenue at 63% of sales by 2030. What has to be true for that, and what breaks it?"
2. "Why is SG&A a plug? Doesn't that hide cost assumptions?"
3. "Delta changed AI equipment depreciation to 3 years. Where does that show up in your model?" (Answer: in D&A, PP&E and EBITDA. EBIT is margin-driven, so the margin inputs must allow for it.)

### Decisions needed from the analyst (largest valuation impact first)

1. **AI data-centre capex growth, FY2027E–FY2030E.** Base is +30% / +15% / +8% / +5%. Nothing moves the valuation more.
2. **Segment margins:** Power Electronics 20%, Infrastructure 18.5%, Automation recovering to 6%, Mobility to +3%.
3. **Content/share gains:** power +5% a year, cooling +10% to +6%, other +3%.
4. **4Q26E revenue:** Base NT$218bn.
5. **AI allocation to Power Electronics** (power 85%, cooling 40%, other 30%), which affects segment mix and so margins.
6. **Payout:** at 50%, net cash rises to about NT$309bn by FY2030E. Is that credible, or should the payout rise?

---

## M4 — Valuation (23 Nov 2026) — *on the draft Base case*

### What was built

- **`src/delta/market.py`:** a dated market snapshot in `data/raw/market/`, taken 23 Nov 2026. It covers Delta's price of NT$990, the five peers with consensus estimates, FX, and five years of prices for Delta and the TAIEX. The risk-free rate comes from the TPEx bond-market review; the ERP, country risk premium and industry betas come from Damodaran (January 2026).
- **`model/valuation_inputs.csv`:** single-value valuation inputs, each with source and rationale.
- **Four new workbook sheets:**
  - **DCF:** WACC build, FCFF, Gordon and exit-multiple terminal values, equity bridge, value per share, target and rating;
  - **Comps:** peer multiples and the implied value for Delta;
  - **Sensitivity:** a live WACC × growth grid, an AI-growth × margin grid and the reverse DCF;
  - **Football field:** valuation ranges by method.

  The workbook now has 2,562 formulas, zero errors and all checks TRUE.
- **`src/delta/model.py`:** the Python mirror. It reproduces the full forecast, DCF, comps and target, and adds the reverse DCF, the AI × margin grid and the scenario values. `tests/test_model_matches_excel.py` proves it agrees with the recalculated workbook within 0.1% for every forecast year in Bear, Base and Bull. Entering the reverse-DCF answer in the workbook reproduces the share price, and a corner of the sensitivity grid reproduces too. That makes 24 tests, all passing.
- **Charts:** `report/figures/football_field.png` and `sensitivity.png`.

### Results (draft Base case; price NT$990 on 23 Nov 2026)

| | |
|---|---|
| WACC | 9.26% (cost of equity 9.46%; levered beta 1.48) |
| DCF value per share today | NT$675 (Gordon, g 2.0%; terminal value 78% of EV) |
| DCF with 12× exit EBITDA | NT$1,010 |
| Comps value per share | NT$859 (EV/EBITDA NT$982; P/E NT$737) |
| 12-month target (50/50, rolled) | NT$821, −17% → **Sell** under the spec rule |
| Bear / Base / Bull target | NT$562 / 821 / 1,033 |
| Reverse DCF | the price implies AI data-centre capex growth of **48% a year in FY2027–30**, against 14% in the draft Base, with AI at 83% of FY2030 revenue |

### The 3 most important modelling choices

1. **Terminal value: Gordon at a growth rate capped by the NT$ risk-free rate, with an exit multiple beside it.** In a NT$ model, perpetual growth above the 2.03% risk-free rate is hard to defend. But it implies 7.3× FY2030E EBITDA, well below where peers trade. Showing both numbers (NT$675 vs NT$1,010) puts the question where it belongs: how much growth is left after 2030?
2. **Minority interests valued at their share of earnings, not book.** Delta Thailand is listed and worth far more than its book minority interest. Deducting book NCI (NT$57bn) would overstate Delta's equity, so the bridge removes 10.5% of equity value, the minorities' share of net income.
3. **A mirror model that proves the spreadsheet.** Every number the report quotes is computed twice, in the workbook's formulas and in Python, and a test fails if they disagree by more than 0.1%. The mirror also does the reverse DCF and the full-model sensitivity grid, and the workbook can reproduce any of them through two shift inputs.

### 3 questions an interviewer could ask

1. "Your DCF says NT$675 and the stock is NT$990 after halving. Why isn't that a screaming Buy for you?" (Answer: on the draft Base case even the stressed grids top out below NT$990. The market is pricing AI capex growth of about 48% a year, or a much higher terminal value. Then say which of those you believe.)
2. "Why cap terminal growth at the risk-free rate?"
3. "How do you value Delta Thailand's minority stake?"

### Decisions needed from the analyst

1. **Terminal method and growth (D14):** Gordon at 2%, a higher g, or the exit multiple. The value range is NT$675–1,010.
2. **Base-case AI path (M3 drafts):** the draft FY2027E EPS of NT$43.9 sits well below consensus (NT$60.4). Either defend the gap or move towards consensus.
3. **Beta (D12):** bottom-up 1.48 vs regression 1.73 (2y raw).
4. **Target blend and rating (D16).** The rating is yours.

---

## M5 — Report draft (23 Nov 2026) — *thesis, debate and rating are for the analyst to write*

### What was built

- **`report/delta_initiation.md`:** the initiation report, about 3,200 words, in the spec's structure: summary page, thesis, key debate, company overview, industry, segment forecasts and scenarios, financials, valuation, risks and appendix. It has six figures, eight tables and 18 footnoted sources.
- **`src/delta/report.py` generates it.** Every number comes from the processed data, the Python mirror (tested equal to the workbook) or the valuation summary, so changing an assumption and running `make model report` updates the whole report. `tests/test_report.py` fails if a value is missing (`nan`/`None`), a cited footnote is undefined, or the target in the text differs from the model.
- **New data for the report:**
  - revenue and non-current assets by geography from each annual report's segment note, with a check that regions sum to group revenue (440 extraction checks);
  - industry facts with primary citations: NVIDIA (GB200 NVL72 ~120 kW per rack; 800V DC for 1 MW racks from 2027) and the IEA (data centres ~1.5% of global electricity in 2024, ~945 TWh by 2030).
- **Charts used:** revenue build (page 1), revenue by segment, margins, football field, sensitivity heatmaps.
- **Length:** a test render with a plain stylesheet comes to 14 A4 pages. The M6 template (two-column summary page, smaller figures, compact notes) is meant to bring it to 10–12.

### Findings the report adds

- **The US is now Delta's largest single-country market:** 36% of FY2025 revenue (NT$197bn), up from 23% in FY2021. Mainland China fell from 33% to 15%.
- **The assets are elsewhere.** The US holds just NT$7.3bn of NT$228bn non-current assets; Thailand holds NT$71bn and Taiwan NT$62bn. That's the core of the tariff risk.

### The 3 most important choices

1. **The report is code, not a document.** Generating it from the model means the numbers in the text can't disagree with the spreadsheet, and the test proves it.
2. **Only sourced claims in the industry section.** Rack power, the 800V architecture and grid demand come from NVIDIA and the IEA directly. Comparative claims with no source were rewritten as numbers that have one.
3. **Judgement sections are drafts with stated reasoning.** Each `[ANALYST TO REVIEW]` section says why it was drafted that way, so the analyst can accept, change or reject the logic, not just the words.

### 3 questions an interviewer could ask

1. "What's the single question the market is arguing about on Delta, and which side are you on?"
2. "Delta sells 36% of its revenue to the US but has almost no assets there. How do you think about tariff risk?"
3. "Your page-1 table shows 27× FY2026E P/E even after the stock halved. Why isn't it cheap?"

### For the analyst

- Rewrite **Investment thesis**, **Key debate** and **Price target and rating** in `src/delta/report.py` (the drafts are there with their reasoning), then run `make report`.
- The rating follows from the draft Base case and the terminal-value choice (D14). Set those first.

---

## M6 — Polish (23 Nov 2026)

### What was built

- **`report/delta_initiation.pdf`:** A4, 10 pages. It has a running header (ticker, initiation, date), "Page n of N", the disclaimer on every page, and a styled title banner. Valuation and the appendix start on new pages, and sources are numbered once each. `src/delta/pdf.py` renders it from the generated Markdown (pandoc, then WeasyPrint). `tests/test_pdf.py` checks it renders at 10–12 pages.
- **Report appendix completed to spec:** summary income statement, balance sheet and cash flow FY2024A–FY2030E, plus the Base-case assumptions table. There's also a new chart of revenue vs non-current assets by region (the tariff-exposure picture).
- **`docs/defence.md`:** 33 interview questions in six groups (business, AI thesis, forecasts, valuation, risks, model mechanics). Each has a pointer to the evidence and space for my answer; no answers are written.
- **`docs/case-study.md`:** a one-page website draft. The disclosure line about AI tools is tagged `[ANALYST TO REVIEW]`.
- **README:** rewritten for a reader arriving cold. It covers the answer first, key findings, deliverables, method, how to run, what's left and limitations.
- **One command:** `make all` runs filings → extract → mirror + Excel → charts → report → PDF; `make test` runs 26 tests. CI installs LibreOffice and pandoc and runs everything on every push.

### The 3 most important choices

1. **The PDF is built from the same generated Markdown,** so the GitHub page, the PDF and the model can't diverge. Fixing one bug (pandoc reading "NT$" as maths) fixed it everywhere.
2. **Length was reached with content, not padding.** The spec asks for 10–12 pages. The gap was closed by the appendix the spec requires (full statements, assumptions) and one more chart that carries a finding.
3. **The defence sheet points to evidence, not answers.** Each question names the decision, sheet or source that holds the answer, so preparation means re-reading my own work rather than memorising lines.

### 3 questions an interviewer could ask

1. "Talk me through your report in two minutes."
2. "If I change your AI capex assumption, how long until your report reflects it?" (One command; everything regenerates.)
3. "What would you do differently with another month?"
