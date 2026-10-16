# M1 data: network check and download list

## Network check (16 Oct 2026, from the cloud build environment)

Every data host returned **HTTP 403 on CONNECT** from the environment's egress proxy (a policy denial, not a site outage). Package installs from PyPI work.

| Source | Host(s) tested | Result |
|---|---|---|
| Delta IR site | `www.deltaww.com` | Blocked (403) |
| Delta IR file centre (where the PDFs live) | `filecenter.deltaww.com` | Blocked (403) |
| MOPS / TWSE | `mops.twse.com.tw`, `mopsov.twse.com.tw`, `emops.twse.com.tw`, `doc.twse.com.tw`, `www.twse.com.tw` | Blocked (403) |
| Yahoo Finance / yfinance | `finance.yahoo.com`, `query1.finance.yahoo.com`, `query2.finance.yahoo.com`, `fc.yahoo.com` | Blocked (403) |
| Central Bank of Taiwan | `www.cbc.gov.tw` | Blocked (403) |
| Damodaran Online | `pages.stern.nyu.edu` | Blocked (403) |
| PyPI | `pypi.org` | OK (200) |

**Two ways forward (pick one):**

- **A. Allowlist the hosts (recommended).** In the cloud environment settings → Network access, add the hosts above under Allowed domains. Then the pipeline can fetch everything itself, record exact URLs and access dates in `sources.csv`, and `make all` stays reproducible.
- **B. Manual download.** Download the files below into `data/raw/` with the file names shown. Don't rename the contents or edit them.

## Files to download into `data/raw/` (option B)

Status of Q3 2026: **not out yet.** A third-party calendar lists Delta's Q3 2026 results for about 29 Oct 2026 (not confirmed by Delta). The latest reported period is therefore **Q2 2026 (1H 2026)**.

### 1. Audited / reviewed consolidated financial statements (English, Taiwan IFRS, NT$ thousands)

Source: Delta IR → Financial Information → Financial Reports (or MOPS → Financial statements). Each annual report carries the prior year as a comparative, so these cover FY2021–FY2025.

| Save as | Document | Why |
|---|---|---|
| `fs_2025_annual_consolidated.pdf` | FY2025 consolidated financial statements with independent auditors' report | FY2025 + FY2024 IS/BS/CF, segment note, capex, debt |
| `fs_2024_annual_consolidated.pdf` | FY2024 consolidated financial statements | FY2024 + FY2023 |
| `fs_2023_annual_consolidated.pdf` | FY2023 consolidated financial statements | FY2023 + FY2022 (restatement check) |
| `fs_2022_annual_consolidated.pdf` | FY2022 consolidated financial statements | FY2022 + FY2021 |
| `fs_2026q2_consolidated.pdf` | Q2 2026 reviewed consolidated financial statements (six months to 30 Jun 2026) | Latest interim period |
| `fs_2026q1_consolidated.pdf` | Q1 2026 reviewed consolidated financial statements | Quarterly split of 1H 2026 |

### 2. Annual reports to shareholders (English)

The FY2024 report is at `https://filecenter.deltaww.com/IR/download/annual_report/2024annual.pdf`; FY2025 likely follows the same pattern (`.../2025annual.pdf`, unverified).

| Save as | Document | Why |
|---|---|---|
| `ar_2025.pdf` | 2025 Annual Report | Segments, plants and capacity, customer concentration, dividends, share count, capex plans |
| `ar_2024.pdf` | 2024 Annual Report | Dividend and payout history, prior-year segment detail |

### 3. Investor conference presentations (earnings slides)

Source: Delta IR → Investor Conference.

| Save as | Document | Why |
|---|---|---|
| `ic_2026q2.pdf` | Q2 2026 investor conference slides | Latest segment mix, margins, capex guidance |
| `ic_2026q1.pdf` | Q1 2026 investor conference slides | |
| `ic_2025q4.pdf` | Q4 2025 investor conference slides (Feb 2026) | Verifies the Taipei Times figures: AI power >20%, AI cooling ~9%, capex NT$46.1bn, Q4 mix |
| `ic_2024q4.pdf`, `ic_2023q4.pdf`, `ic_2022q4.pdf`, `ic_2021q4.pdf` | Q4 slides for earlier years | Segment revenue mix and margin history if the FS segment note is less detailed |

If an earnings-call transcript or recording summary is published on the IR site for Q4 2025, Q1 2026 or Q2 2026, include it as `call_<period>.pdf`.

### 4. Monthly revenue reports

URL pattern confirmed: `https://filecenter.deltaww.com/ir/download/monthly_revenues/YYYYMM_EN.pdf`.

| Save as | Files | Why |
|---|---|---|
| `mr_202601_EN.pdf` … `mr_202609_EN.pdf` | Jan–Sep 2026 (9 files) | Gives Q3 2026 revenue ahead of the full results; cross-check for the FY2026E forecast |

### 5. Market data (only needed under option B; yfinance does this under A)

| Save as | Data | Where |
|---|---|---|
| `px_2308TW.csv` | Daily prices for 2308.TW, last 3 years (Date, Close, Adj Close, Volume) | Yahoo Finance → Historical Data → Download |
| `px_TWII.csv` | TAIEX (^TWII) daily, same dates | Yahoo Finance |
| `px_peers.csv` or one CSV per ticker | VRT, SU.PA, ETN, 2301.TW, 6282.TW: price, shares, net debt, trailing and forward EBITDA/EPS | Yahoo Finance quote + Statistics/Analysis pages (screenshot or CSV) |
| `fx_rates.csv` | USD/TWD and EUR/TWD on the price date | Yahoo Finance (TWD=X, EURTWD=X) or CBC |
| `tw10y_yield.csv` | Taiwan 10-year government bond yield, latest | CBC statistics or TPEx |

### 6. Damodaran Online datasets (latest update, likely Jan 2026)

| Save as | Dataset |
|---|---|
| `damodaran_ctryprem.xlsx` | Country default spreads and risk premiums (`ctryprem.xlsx`) |
| `damodaran_histimpl.xlsx` | Implied ERP history for the US (`histimpl.xls`) |
| `damodaran_betaGlobal.xlsx` | Betas by sector, Global (`betaGlobal.xls`) |
| `damodaran_betaemerg.xlsx` | Betas by sector, Emerging markets (`betaemerg.xls`) |

### 7. Industry context (needed for M5, not M1)

Hyperscaler earnings releases with capex figures (Microsoft, Alphabet, Amazon, Meta; latest quarter) and NVIDIA product material on rack power for GB200/GB300 NVL72. These can wait until the report milestone.
