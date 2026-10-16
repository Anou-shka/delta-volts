"""Position-based parser for the statement pages of Delta's financial reports.

The PDFs' text order is unreliable (a value can sit one line below its label,
labels wrap, the closing bracket of a negative number is a separate glyph), so
rows are rebuilt from word coordinates:

* words left of the notes column are label text;
* numeric tokens right of the "US Dollars" header are values, assigned to a
  column by their right edge (figures are right-aligned);
* an item stays open until a new label line starts after at least one value
  has been seen, so wrapped labels and values printed a line low both work.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import pymupdf

NUM = re.compile(r"^\(?\$?\(?(-|[\d,]+(\.\d+)?)\)?$")


@dataclass
class Row:
    label: str
    values: dict[int, float] = field(default_factory=dict)  # column index -> value
    y: float = 0.0


def _is_num(tok: str) -> bool:
    t = tok.strip()
    return bool(NUM.match(t)) and any(c.isdigit() for c in t) or t == "-"


def _value(tok: str) -> float:
    t = tok.replace("$", "").replace(",", "").strip("()")
    return 0.0 if t in ("-", "") else float(t)


def _lines(words, tol=2.5):
    """Group words into visual lines by their top coordinate."""
    lines: list[list] = []
    for w in sorted(words, key=lambda w: (round(w[1], 1), w[0])):
        if lines and abs(lines[-1][0][1] - w[1]) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(l, key=lambda w: w[0]) for l in lines]


def column_edges(page: pymupdf.Page, value_left: float) -> list[float]:
    """Right edges of the value columns, found by clustering numeric tokens."""
    rights = sorted(w[2] for w in page.get_text("words")
                    if w[0] >= value_left and _is_num(w[4]) and w[4] != "-")
    clusters: list[list[float]] = []
    for r in rights:
        if clusters and r - clusters[-1][-1] <= 12:
            clusters[-1].append(r)
        else:
            clusters.append([r])
    # ignore stray clusters (page numbers, footnote markers)
    return [max(c) for c in clusters if len(c) >= 3]


HEADER_WORDS = re.compile(r"^(US|Dollars|New|Taiwan|Notes|Items|Assets|Liabilities|and|Equity|"
                          r"December|June|March|September|\d{1,2},?|(19|20)\d\d,?|Six|Three|Nine|"
                          r"months|ended|Year|years|ended,|June|30,|31,)$", re.I)
NOTE_TOKEN = re.compile(r"^(\d+(\(\d+\))*|and|(\(\d+\))+|\d+\(\d+\)(\(\d+\))*)$")


def _is_open(tok: str) -> bool:
    t = tok.replace("$", "").replace("\u00a0", "").strip()
    return t != "" and set(t) <= {"("}


def parse_page(page: pymupdf.Page) -> tuple[list[Row], int]:
    words = page.get_text("words")
    us = [w for w in words if w[4] == "US" and w[1] < 200]
    notes = [w for w in words if w[4] == "Notes" and w[1] < 200]
    if not us:
        return [], 0
    value_left = us[0][0] - 30
    notes_x = notes[0][0] if notes else value_left - 60
    edges = column_edges(page, value_left)
    header_top = max(w[3] for w in us + notes)

    rows: list[Row] = []
    cur: Row | None = None
    cur_x = 0.0
    for line in _lines(words):
        y = line[0][1]
        if y <= header_top + 1 or line[0][4].startswith("~"):
            continue
        text = " ".join(w[4] for w in line)
        if "accompanying notes" in text or text.strip() == "(Continued)":
            continue
        if y < header_top + 40 and all(HEADER_WORDS.match(w[4]) for w in line):
            continue
        lab_words = [w for w in line if w[2] < value_left and not (
            w[0] >= notes_x - 30 and NOTE_TOKEN.match(w[4].strip("()").replace(")(", "") or "x"))]
        lab_words = [w for w in lab_words if not (w[0] >= notes_x - 30 and NOTE_TOKEN.match(w[4]))]
        label = " ".join(w[4] for w in lab_words).strip()
        vals: dict[int, float] = {}
        pending_open = False
        for i, w in enumerate(line):
            if w[0] < value_left:
                continue
            if _is_open(w[4]):
                pending_open = True
                continue
            if not _is_num(w[4]):
                continue
            col = min(range(len(edges)), key=lambda k: abs(edges[k] - w[2]))
            if abs(edges[col] - w[2]) > 14:
                continue
            v = _value(w[4])
            nxt = line[i + 1][4] if i + 1 < len(line) else ""
            neg = pending_open or "(" in w[4] or w[4].endswith(")") or (
                nxt.startswith(")") and line[i + 1][0] - w[2] <= 8)
            pending_open = False
            vals[col] = -v if neg and v else v
        if label:
            x0 = lab_words[0][0]
            if cur is not None and not cur.values and (x0 > cur_x + 3 or label[0].islower()):
                cur.label = f"{cur.label} {label}"      # wrapped label
            else:
                cur = Row(label=label, y=y)
                cur_x = x0
                rows.append(cur)
        if vals and cur is not None:
            for k, v in vals.items():
                cur.values.setdefault(k, v)
    out = []
    for r in rows:
        r.label = re.sub(r"\s+", " ", r.label).strip()
        if r.values:
            out.append(r)
    return out, len(edges)
