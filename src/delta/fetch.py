"""Download the primary-source files listed in FILES into data/raw/.

Each download is logged in data/raw/manifest.csv with its URL, access date,
size and SHA-256, so every figure extracted later can be traced to a file.
Files already present are skipped (re-run with --force to refresh).
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
RAW = ROOT / "data" / "raw"
MANIFEST = RAW / "manifest.csv"

FC = "https://filecenter.deltaww.com/ir/download"

# (saved file name, URL, description)
FILES: list[tuple[str, str, str]] = []

# Annual consolidated financial statements (audited). Each carries the prior year as comparative.
for y in (2021, 2022, 2023, 2024):
    FILES.append((f"fs_{y}_annual_consolidated.pdf",
                  f"{FC}/financial_report/Q4{y}_eng_Consolidated Financial Report.pdf",
                  f"FY{y} consolidated financial statements (audited)"))
FILES.append(("fs_2025_annual_consolidated.pdf",
              f"{FC}/financial_report/Q4 2025_eng_ Consolidated Financial Report.pdf",
              "FY2025 consolidated financial statements (audited)"))

# Interim consolidated financial statements (reviewed).
FILES += [
    ("fs_2026q1_consolidated.pdf", f"{FC}/financial_report/Q1 2026_eng_ Consolidated Financial Report.pdf",
     "Q1 2026 consolidated financial statements (reviewed)"),
    ("fs_2026q2_consolidated.pdf", f"{FC}/financial_report/Q22026_eng_Consolidated Financial Report.pdf",
     "Q2 2026 consolidated financial statements (reviewed)"),
]

# Annual reports to shareholders.
for y in (2024, 2025):
    FILES.append((f"ar_{y}.pdf", f"{FC}/annual_report/{y}annual.pdf", f"{y} Annual Report"))

# Analyst meeting (earnings) slides: Q4 of each year plus 2026 quarters, and the latest transcript.
for q in ("4Q21", "4Q22", "4Q23", "4Q24", "4Q25", "1Q26", "2Q26"):
    FILES.append((f"ic_{q}.pdf", f"{FC}/calendar/{q}_Analyst Meeting.pdf", f"{q} analyst meeting slides"))
FILES.append(("call_2Q26_transcript.pdf", f"{FC}/calendar/2Q26_Transcript.pdf", "2Q26 analyst meeting transcript"))

# Monthly revenue reports, 2025-2026 (Q3 2026 revenue ahead of results; quarterly cross-checks).
for y, months in ((2025, range(1, 13)), (2026, range(1, 10))):
    for m in months:
        FILES.append((f"mr_{y}{m:02d}_EN.pdf", f"{FC}/monthly_revenues/{y}{m:02d}_EN.pdf",
                      f"Monthly revenue report {y}-{m:02d}"))

# AGM minutes 2026 (dividend approved out of FY2025 earnings).
FILES.append(("agm_2026_minutes.pdf", f"{FC}/calendar/20260528_Minutes of 2026AGM_Eng.pdf",
              "Minutes of 2026 AGM"))


def _quote(url: str) -> str:
    parts = urllib.parse.urlsplit(url)
    return urllib.parse.urlunsplit(parts._replace(path=urllib.parse.quote(parts.path)))


def download(url: str, dest: Path, tries: int = 4) -> bytes:
    req = urllib.request.Request(_quote(url), headers={"User-Agent": "Mozilla/5.0 (student research)"})
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                data = r.read()
            dest.write_bytes(data)
            return data
        except Exception as e:  # network hiccup: back off and retry
            if i == tries - 1:
                raise
            print(f"  retry {i + 1} after error: {e}", file=sys.stderr)
            time.sleep(2 ** (i + 1))
    raise RuntimeError("unreachable")


def load_manifest() -> dict[str, dict]:
    if not MANIFEST.exists():
        return {}
    with MANIFEST.open() as f:
        return {row["file"]: row for row in csv.DictReader(f)}


def save_manifest(rows: dict[str, dict]) -> None:
    cols = ["file", "description", "url", "date_accessed", "bytes", "sha256"]
    with MANIFEST.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for k in sorted(rows):
            w.writerow({c: rows[k].get(c, "") for c in cols})


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--force", action="store_true", help="re-download files already present")
    args = ap.parse_args(argv)

    RAW.mkdir(parents=True, exist_ok=True)
    rows = load_manifest()
    failed = []
    for name, url, desc in FILES:
        dest = RAW / name
        if dest.exists() and name in rows and not args.force:
            continue
        print(f"get {name}")
        try:
            data = download(url, dest)
        except Exception as e:
            print(f"  FAILED: {e}", file=sys.stderr)
            failed.append((name, url))
            continue
        if not data.startswith(b"%PDF"):
            print(f"  WARNING: {name} is not a PDF", file=sys.stderr)
        rows[name] = {"file": name, "description": desc, "url": url,
                      "date_accessed": dt.date.today().isoformat(), "bytes": len(data),
                      "sha256": hashlib.sha256(data).hexdigest()}
    save_manifest(rows)
    for name, url in failed:
        print(f"MISSING {name}: download manually from {url}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
