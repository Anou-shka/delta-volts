# Data

| Folder / file | Contents |
|---|---|
| `raw/` | Original PDFs from Delta's IR file centre, untouched. `raw/manifest.csv` records each file's URL, access date, size and SHA-256. Re-download with `make fetch`. |
| `processed/fs_lines.csv` | Every line of the consolidated balance sheet, income statement and cash flow statement, every period column, **as reported** (NT$ thousands). |
| `processed/historicals.csv` | Standardised model lines by period (**NT$ millions**; EPS in NT$). Periods: FY2021–FY2025, 1Q25, 2Q25, 1H25, 1Q26, 2Q26, 1H26. |
| `processed/segments.csv` | Segment revenue and operating profit from the segment note (NT$ thousands). `latest = True` marks the most recent (restated) presentation of each period. |
| `processed/monthly_revenue.csv` | Unaudited monthly revenue releases, Jan 2025 – Sep 2026 (NT$ millions), with segment mix %. |
| `processed/checks.csv` | Every tie-out check (subtotals, statement links, cash, restatements, segments, monthly vs quarterly). |
| `sources.csv` | Every sourced figure: value, unit, period, document, page and URL, date accessed. Built by `extract.py` from the filings plus `sources_manual.csv` (figures from transcripts and slides). |

## Source documents (all from https://www.deltaww.com/en-US/investors)

- Consolidated financial statements: FY2021–FY2025 (audited), Q1 2026 and Q2 2026 (reviewed).
- Annual reports 2024 and 2025.
- Analyst meeting slides 4Q21–4Q25, 1Q26, 2Q26; 2Q26 meeting transcript (the only transcript published).
- Monthly revenue releases Jan 2025 – Sep 2026.
- Minutes of the 2026 AGM.

## Conventions

- Signs follow the statements: costs, expenses, tax, capex and dividends paid are negative.
- Balance-sheet values sit under the period whose closing date they show (FY2025 = 31 Dec 2025, 1H26 = 30 Jun 2026, 1Q26 = 31 Mar 2026).
- Cash flow statements are year-to-date; 2Q cash flows are derived as 1H − 1Q.
- Each fiscal year is taken from its own audited report; prior-year comparatives in the next report are used only as a restatement check (none found in the statements).
- Segments: FY2021–FY2022 are on Delta's old 3-segment basis (Mobility inside Power Electronics). FY2023 onward use the 4-segment basis from the latest report.
