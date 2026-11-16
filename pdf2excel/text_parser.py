"""Turn pasted text into spreadsheet rows.

Handles text copied from Excel or web pages (tabs), CSV, Markdown or ASCII tables
(pipes), semicolon lists, columns lined up with spaces (common when copying from
PDFs), "Label: value" lists, and lines that end in numbers such as statements.
"""

from __future__ import annotations

import csv
import io
import re
from collections import Counter

from .converter import to_number

DELIMITERS = {
    "auto": None,
    "tab": "\t",
    "comma": ",",
    "semicolon": ";",
    "pipe": "|",
    "spaces": "spaces",
    "label": "label",
    "numbers": "numbers",
}

_MULTI_SPACE = re.compile(r"\s{2,}")
_PIPE_RULE = re.compile(r"^\s*\|?[\s:+\-|=]+\|?\s*$")  # |---|:---:| or +----+----+
_LABEL = re.compile(r"^\s*([^:\t]{1,60}?)\s*:\s+(.+?)\s*$")
_VALUE_TOKEN = re.compile(
    r"^(?:"
    r"\(?-?[$€£]?[\d,]*\.?\d+\)?%?-?"  # 1,200.00  (350)  $9  15%  12.50-
    r"|\d{1,4}[/.-]\d{1,2}(?:[/.-]\d{2,4})?"  # 01/02/2024  2024-01-02  3/15
    r")$"
)


def _split(line: str, mode: str) -> list[str]:
    if mode == "\t":
        return [c.strip() for c in line.split("\t")]
    if mode in (",", ";"):
        return [c.strip() for c in next(csv.reader([line], delimiter=mode, skipinitialspace=True), [""])]
    if mode == "|":
        cells = [c.strip() for c in line.strip().split("|")]
        if line.strip().startswith("|"):
            cells = cells[1:]
        if line.strip().endswith("|") and cells:
            cells = cells[:-1]
        return cells
    if mode == "spaces":
        return _MULTI_SPACE.split(line.strip())
    if mode == "label":
        m = _LABEL.match(line)
        return [m.group(1), m.group(2)] if m else [line.strip()]
    if mode == "numbers":
        return _split_trailing_values(line)
    return [line.strip()]


def _split_trailing_values(line: str) -> list[str]:
    """'Rent payment 01/02 1,200.00 3,400.00' -> ['Rent payment', '01/02', '1,200.00', '3,400.00']."""
    tokens = line.split()
    values: list[str] = []
    while len(tokens) > 1 and _VALUE_TOKEN.match(tokens[-1]):
        values.insert(0, tokens.pop())
    # a leading date or number (e.g. a transaction date) becomes its own column too
    lead: list[str] = []
    while len(tokens) > 1 and _VALUE_TOKEN.match(tokens[0]):
        lead.append(tokens.pop(0))
    return lead + [" ".join(tokens)] + values


def _consistency(lines: list[str], mode: str) -> float:
    """Share of lines that split into the most common column count (only if > 1 column)."""
    counts = Counter(len(_split(line, mode)) for line in lines)
    width, hits = counts.most_common(1)[0]
    if width < 2:
        multi = sum(n for w, n in counts.items() if w > 1)
        return 0.0 if multi == 0 else multi / len(lines) * 0.5
    return hits / len(lines)


def detect_mode(lines: list[str]) -> str:
    """Pick the splitting rule that gives the most consistent columns."""
    if not lines:
        return "spaces"
    sample = lines[:200]
    candidates: list[tuple[float, int, str]] = []
    for priority, mode in enumerate(("\t", "|", ",", ";")):
        if sum(mode in line for line in sample) >= max(1, len(sample) * 0.6):
            candidates.append((_consistency(sample, mode), -priority, mode))
    if sum(bool(_LABEL.match(line)) for line in sample) >= len(sample) * 0.8:
        candidates.append((0.95, -10, "label"))
    candidates.append((_consistency(sample, "spaces") * 0.95, -20, "spaces"))
    candidates.append((_consistency(sample, "numbers") * 0.85, -30, "numbers"))
    best = max(candidates)
    return best[2] if best[0] > 0 else "spaces"


def parse_text(text: str, delimiter: str = "auto", convert_numbers: bool = True) -> dict:
    """Parse pasted text. Returns {'rows': [[...]], 'mode': str, 'header': bool}."""
    lines = [line.rstrip("\r") for line in text.split("\n")]
    lines = [line for line in lines if line.strip()]
    mode = DELIMITERS.get(delimiter, None) if delimiter != "auto" else None
    if mode is None:
        mode = detect_mode(lines)
    if mode == "|":
        lines = [line for line in lines if not _PIPE_RULE.match(line)]
    if mode == "," and lines:
        rows = [[c.strip() for c in r] for r in csv.reader(io.StringIO("\n".join(lines)), skipinitialspace=True)]
    else:
        rows = [_split(line, mode) for line in lines]

    width = max((len(r) for r in rows), default=0)
    rows = [r + [""] * (width - len(r)) for r in rows]
    if convert_numbers:
        rows = [[to_number(c) for c in r] for r in rows]

    header = mode != "label" and _looks_like_header(rows)
    name = {v: k for k, v in DELIMITERS.items() if v}.get(mode, "spaces")
    return {"rows": rows, "mode": name, "header": header}


def _looks_like_header(rows: list[list[object]]) -> bool:
    if len(rows) < 2:
        return False
    first = rows[0]
    if any(isinstance(c, (int, float)) for c in first):
        return False
    if sum(1 for c in first if c not in ("", None)) < max(1, len(first) // 2):
        return False
    return any(isinstance(c, (int, float)) for r in rows[1:] for c in r) or len(rows) > 2
