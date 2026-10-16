"""Render selected sheets of the model to PNG for the README (needs LibreOffice).

Run: python -m src.delta.preview   ->  report/figures/model_<sheet>.png
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parents[2]
MODEL = ROOT / "model" / "delta_model.xlsx"
FIG = ROOT / "report" / "figures"
SHEETS = {"IS": "model_income_statement.png", "Checks": "model_checks.png"}


def main() -> int:
    with tempfile.TemporaryDirectory() as tmp:
        shutil.copy(MODEL, tmp)
        env = {**os.environ, "SAL_USE_VCLPLUGIN": "svp"}
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(Path(tmp) / MODEL.name)],
                       check=True, capture_output=True, env=env, timeout=180)
        doc = pymupdf.open(Path(tmp) / "delta_model.pdf")
        for sheet, out in SHEETS.items():
            page = next(p for p in doc if f"model - {sheet} - page" in p.get_text())
            text_rects = [b[:4] for b in page.get_text("blocks") if "model - " not in b[4]]
            x0, y0 = min(r[0] for r in text_rects) - 12, min(r[1] for r in text_rects) - 12
            x1, y1 = max(r[2] for r in text_rects) + 12, max(r[3] for r in text_rects) + 12
            page.get_pixmap(dpi=150, clip=pymupdf.Rect(x0, y0, x1, y1)).save(FIG / out)
            print(FIG / out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
