"""Render report/delta_initiation.md to report/delta_initiation.pdf (A4, header with ticker and date, page numbers).

Markdown -> HTML with pandoc, then HTML + print CSS -> PDF with WeasyPrint.
Run: python -m src.delta.pdf
"""
from __future__ import annotations

import html
import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REPORT = ROOT / "report"
SRC = REPORT / "delta_initiation.md"
OUT = REPORT / "delta_initiation.pdf"

CSS = """
@page {
  size: A4; margin: 19mm 15mm 16mm 15mm;
  @top-left { content: "Delta Electronics (2308.TW) · Initiation of coverage"; font: 600 7.5pt Inter, sans-serif;
              color: #1F3864; }
  @top-right { content: "%(date)s"; font: 7.5pt Inter, sans-serif; color: #595959; }
  @bottom-left { content: "Student research for educational purposes only. Not investment advice.";
                 font: italic 7pt Inter, sans-serif; color: #595959; }
  @bottom-right { content: "Page " counter(page) " of " counter(pages); font: 7.5pt Inter, sans-serif; color: #595959; }
}
@page :first { @top-left { content: none; } @top-right { content: none; } }
html { font-family: Inter, "DejaVu Sans", sans-serif; font-size: 9.4pt; line-height: 1.42; color: #1a1a1a; }
body { margin: 0; }
.banner { background: #1F3864; color: white; padding: 10px 14px 9px; margin: 0 0 8px; }
.banner .t { font-size: 17pt; font-weight: 700; letter-spacing: -0.2px; }
.banner .s { font-size: 9pt; opacity: 0.9; margin-top: 2px; }
.banner .a { font-size: 7.5pt; opacity: 0.8; margin-top: 4px; }
h1 { font-size: 12pt; color: #1F3864; border-bottom: 1.5px solid #1F3864; padding-bottom: 2px;
     margin: 13px 0 6px; break-after: avoid; }
h1:first-of-type { margin-top: 4px; }
p { margin: 0 0 5px; text-align: left; orphans: 3; widows: 3; }
ul { margin: 0 0 5px 0; padding-left: 15px; } li { margin: 0 0 2px; }
strong { color: #111; }
code { font-family: Inter, sans-serif; font-size: 7.2pt; font-weight: 600; background: #FFF3BF; color: #7a5a00;
       padding: 0 3px; border-radius: 2px; }
table { border-collapse: collapse; width: 100%%; margin: 4px 0 6px; font-size: 8.2pt; break-inside: avoid; }
th { background: #1F3864; color: white; font-weight: 600; padding: 2.5px 5px; text-align: left; }
td { padding: 2px 5px; border-bottom: 0.5px solid #d9d9d9; vertical-align: top; }
tr:nth-child(even) td { background: #F5F7FB; }
td:first-child, th:first-child { width: 30%%; }
td[style*="right"], th[style*="right"] { text-align: right; }
figure { margin: 6px 0 8px; text-align: center; break-inside: avoid; }
figure img { width: 82%% !important; }
figure.wide img { width: 94%% !important; }
figcaption { display: none; }
em { color: #444; }
section.footnotes { font-size: 7.4pt; line-height: 1.3; color: #444; columns: 2; column-gap: 14px; margin-top: 6px; }
section.footnotes hr { display: none; }
section.footnotes ol { padding-left: 14px; margin: 0; } section.footnotes li { margin: 0 0 1px; }
section.footnotes a.footnote-back { display: none; }
a { color: #1F3864; text-decoration: none; }
sup a { font-size: 6.5pt; }
h1#sources + ol { font-size: 7.4pt; line-height: 1.3; color: #444; columns: 2; column-gap: 14px; }
figure.p1 img { width: 68%% !important; }
h1#valuation, h1#appendix { break-before: page; }
"""


def front_matter(md: str) -> tuple[dict, str]:
    m = re.match(r"^---\n(.*?)\n---\n", md, flags=re.S)
    meta = {}
    if m:
        for line in m.group(1).splitlines():
            k, _, v = line.partition(":")
            meta[k.strip()] = v.strip().strip('"')
        md = md[m.end():]
    return meta, md


def number_sources(md: str) -> str:
    """Replace named footnotes with numbered references to a de-duplicated source list (one entry per source)."""
    defs = dict(re.findall(r"^\[\^([a-z0-9]+)\]:\s*(.+)$", md, flags=re.M))
    md = re.sub(r"^\[\^[a-z0-9]+\]:.*$\n?", "", md, flags=re.M)
    order: list[str] = []
    for key in re.findall(r"\[\^([a-z0-9]+)\]", md):
        if key not in order:
            order.append(key)
    md = re.sub(r"\[\^([a-z0-9]+)\]", lambda m: f"<sup>[{order.index(m.group(1)) + 1}]</sup>", md)
    md = re.sub(r"</sup><sup>", "", md.replace("]</sup><sup>[", ","))
    notes = "\n".join(f"{i + 1}. {defs[k]}" for i, k in enumerate(order))
    return md.rstrip() + "\n\n# Sources {#sources}\n\n" + notes + "\n"


def render(out: Path = OUT) -> int:
    from weasyprint import CSS as WCSS
    from weasyprint import HTML
    meta, body_md = front_matter(SRC.read_text())
    body_md = number_sources(body_md)
    body = subprocess.run(["pandoc", "--from", "markdown-tex_math_dollars-tex_math_single_backslash", "--to", "html5",
                           "--wrap=none"], input=body_md,
                          capture_output=True, text=True, check=True).stdout
    # wide charts (multi-panel) get a little more width
    body = re.sub(r'<figure>(\s*<img [^>]*revenue_build[^>]*>)', r'<figure class="p1">\1', body, count=1)
    body = re.sub(r'<figure>(\s*<img [^>]*sensitivity[^>]*>)', r'<figure class="wide">\1', body)
    date = meta.get("subtitle", "").split("·")[-1].strip()
    banner = (f'<div class="banner"><div class="t">{html.escape(meta.get("title", ""))}</div>'
              f'<div class="s">{html.escape(meta.get("subtitle", ""))}</div>'
              f'<div class="a">{html.escape(meta.get("author", ""))}</div></div>')
    doc = f'<!doctype html><html lang="en"><head><meta charset="utf-8"><title>{html.escape(meta.get("title", ""))}' \
          f'</title></head><body>{banner}{body}</body></html>'
    pdf = HTML(string=doc, base_url=str(REPORT)).render(stylesheets=[WCSS(string=CSS % {"date": date})])
    pdf.write_pdf(out)
    return len(pdf.pages)


def main() -> int:
    n = render()
    print(f"{OUT} ({n} pages)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
