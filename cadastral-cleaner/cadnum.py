"""Clean a list of Ukrainian cadastral numbers.

Reads numbers from CSV or Excel, normalises them to XXXXXXXXXX:XX:XXX:XXXX,
removes duplicates and writes an Excel + JSON report in which every input row
is accounted for: accepted, duplicate, or rejected with the reason.

Usage:
    python cadnum.py numbers.xlsx --out out
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path

log = logging.getLogger("cadnum")

GROUPS = (10, 2, 3, 4)  # KOATUU code : zone : quarter : parcel
TOTAL_DIGITS = sum(GROUPS)

OK = "ok"
DUPLICATE = "duplicate"
INVALID = "invalid"
EMPTY = "empty"

# Every kind of whitespace, including non-breaking and zero-width characters
# that arrive with text copied from web pages and PDFs.
_WHITESPACE_RE = re.compile(r"[\s ​‌‍﻿]+")
_SCIENTIFIC_RE = re.compile(r"^\d+([.,]\d+)?[eE]\+?\d+$")


@dataclass
class Row:
    row: int  # row number in the source file, 1-based
    raw: str  # value exactly as read
    number: str | None  # normalised number, None when rejected
    status: str
    reason: str = ""


def normalize(value: object) -> tuple[str | None, str]:
    """Return (normalised number, "") or (None, reason it was rejected)."""
    # Excel stores long digit strings as floats and silently drops digits
    # beyond the 15th, so a numeric cell can never be trusted.
    if isinstance(value, float) or (isinstance(value, int) and not isinstance(value, bool)):
        return None, "numeric cell: Excel keeps only 15 digits, the number is damaged; store it as text"
    if value is None:
        return None, "empty value"

    text = _WHITESPACE_RE.sub("", str(value))
    if not text:
        return None, "empty value"
    if _SCIENTIFIC_RE.match(text):
        return None, "scientific notation: the number was damaged by a spreadsheet"

    if ":" in text:
        parts = text.split(":")
        if len(parts) != len(GROUPS):
            return None, f"expected {len(GROUPS)} groups separated by ':', got {len(parts)}"
        for part, size in zip(parts, GROUPS):
            if not part.isascii() or not part.isdigit():
                return None, f"non-digit characters in group '{part}'"
            if len(part) != size:
                return None, f"group '{part}' has {len(part)} digits, expected {size}"
        return ":".join(parts), ""

    if not text.isascii() or not text.isdigit():
        return None, "contains characters other than digits and ':'"
    if len(text) != TOTAL_DIGITS:
        return None, f"{len(text)} digits, expected {TOTAL_DIGITS}"
    parts, start = [], 0
    for size in GROUPS:
        parts.append(text[start : start + size])
        start += size
    return ":".join(parts), ""


def process(values: list[tuple[int, object]]) -> list[Row]:
    """Classify every input value; the first occurrence of a number wins."""
    seen: dict[str, int] = {}
    rows: list[Row] = []
    for row_number, value in values:
        raw = "" if value is None else str(value)
        number, reason = normalize(value)
        if number is None:
            status = EMPTY if reason == "empty value" else INVALID
            rows.append(Row(row_number, raw, None, status, reason))
        elif number in seen:
            rows.append(Row(row_number, raw, number, DUPLICATE, f"same as row {seen[number]}"))
        else:
            seen[number] = row_number
            rows.append(Row(row_number, raw, number, OK))
    return rows


def _looks_like_header(value: object) -> bool:
    """A first cell with no digits at all is a column title, not data."""
    return isinstance(value, str) and bool(value.strip()) and not any(ch.isdigit() for ch in value)


def _pick_column(header: list[object], column: str | None) -> int:
    if column is None:
        return 0
    if column.isdigit():
        return int(column) - 1
    names = [str(h).strip().casefold() if h is not None else "" for h in header]
    if column.strip().casefold() not in names:
        raise ValueError(f"column '{column}' not found; available: {[n for n in names if n]}")
    return names.index(column.strip().casefold())


def _select(table: list[list[object]], column: str | None) -> list[tuple[int, object]]:
    if not table:
        return []
    has_header = _looks_like_header(table[0][0] if table[0] else None) or (
        column is not None and not column.isdigit()
    )
    index = _pick_column(table[0], column)
    if index < 0:
        raise ValueError("column number must be 1 or greater")
    start = 1 if has_header else 0
    return [
        (number, line[index] if index < len(line) else None)
        for number, line in enumerate(table[start:], start=start + 1)
    ]


def read_csv(path: Path, column: str | None = None) -> list[tuple[int, object]]:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("could not decode the file as UTF-8 or Windows-1251")
    first_line = text.split("\n", 1)[0]
    # ':' is part of the number itself, so only these can be delimiters.
    delimiter = max(";,\t", key=first_line.count)
    table = [list(line) for line in csv.reader(text.splitlines(), delimiter=delimiter)]
    return _select(table, column)


def read_excel(path: Path, column: str | None = None) -> list[tuple[int, object]]:
    from openpyxl import load_workbook

    sheet = load_workbook(path, read_only=True, data_only=True).active
    table = [list(line) for line in sheet.iter_rows(values_only=True)]
    return _select(table, column)


def read_input(path: Path, column: str | None = None) -> list[tuple[int, object]]:
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return read_excel(path, column)
    if suffix in (".csv", ".txt"):
        return read_csv(path, column)
    raise ValueError(f"unsupported file type '{suffix}', expected .csv, .txt or .xlsx")


def write_json(rows: list[Row], path: Path) -> None:
    payload = {
        "summary": dict(Counter(r.status for r in rows)),
        "numbers": [r.number for r in rows if r.status == OK],
        "rows": [asdict(r) for r in rows],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_excel(rows: list[Row], path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    workbook = Workbook()
    numbers = workbook.active
    numbers.title = "Numbers"
    numbers.append(["Cadastral number"])
    for r in rows:
        if r.status == OK:
            numbers.append([r.number])
    numbers.column_dimensions["A"].width = 28

    report = workbook.create_sheet("Report")
    report.append(["Source row", "Raw value", "Normalised", "Status", "Reason"])
    for r in rows:
        report.append([r.row, r.raw, r.number, r.status, r.reason])
    for column, width in zip("ABCDE", (11, 30, 28, 11, 70)):
        report.column_dimensions[column].width = width

    for sheet in (numbers, report):
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        sheet.freeze_panes = "A2"
        # Keep every value as text so Excel never turns a number into 5.62E+18.
        for line in sheet.iter_rows(min_row=2):
            for cell in line:
                if isinstance(cell.value, str):
                    cell.number_format = "@"
    workbook.save(path)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("input", help="CSV or Excel file with cadastral numbers")
    parser.add_argument("--column", help="column name or 1-based number (default: first column)")
    parser.add_argument("--out", default="out", help="output folder")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    source = Path(args.input)
    try:
        rows = process(read_input(source, args.column))
    except (OSError, ValueError) as exc:
        log.error("cannot read %s: %s", source, exc)
        return 2

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    write_json(rows, out_dir / f"{source.stem}-cleaned.json")
    write_excel(rows, out_dir / f"{source.stem}-cleaned.xlsx")

    counts = Counter(r.status for r in rows)
    log.info(
        "rows: %d | unique valid: %d | duplicates: %d | invalid: %d | empty: %d",
        len(rows), counts[OK], counts[DUPLICATE], counts[INVALID], counts[EMPTY],
    )
    for r in rows:
        if r.status == INVALID:
            log.warning("row %d rejected: %r - %s", r.row, r.raw, r.reason)
    return 0


if __name__ == "__main__":
    sys.exit(main())
