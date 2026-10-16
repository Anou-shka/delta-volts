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
