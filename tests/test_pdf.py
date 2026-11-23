"""The report renders to an A4 PDF of 10-12 pages (page count checked when the report font, Inter, is installed)."""
import shutil
import subprocess

import pytest

pytestmark = pytest.mark.skipif(shutil.which("pandoc") is None, reason="pandoc not installed")


def test_pdf_renders(tmp_path):
    pytest.importorskip("weasyprint")
    from src.delta import pdf, report
    report.write()
    pages = pdf.render(tmp_path / "report.pdf")
    fonts = subprocess.run(["fc-list"], capture_output=True, text=True).stdout
    if "Inter" in fonts:
        assert 10 <= pages <= 12, f"{pages} pages"
    else:
        assert pages >= 8
