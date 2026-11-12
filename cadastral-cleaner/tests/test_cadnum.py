import json
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import cadnum  # noqa: E402

GOOD = "5622681900:01:001:0087"


@pytest.mark.parametrize(
    "raw",
    [
        "5622681900:01:001:0087",
        "5622681900010010087",
        "  5622681900:01:001:0087  ",
        "5622681900 : 01 : 001 : 0087",
        "5622681900 01 001 0087",
        "﻿5622681900010010087",
        "5622681900:01:001:0087\n",
    ],
)
def test_normalize_accepts_both_formats_and_stray_whitespace(raw):
    assert cadnum.normalize(raw) == (GOOD, "")


def test_leading_zeros_are_kept():
    assert cadnum.normalize("0520880300020010015") == ("0520880300:02:001:0015", "")


@pytest.mark.parametrize(
    "raw, fragment",
    [
        ("", "empty"),
        ("   ", "empty"),
        (None, "empty"),
        ("562268190001001008", "18 digits"),
        ("56226819000100100871", "20 digits"),
        ("5622681900:01:001", "got 3"),
        ("5622681900:1:001:0087", "has 1 digits, expected 2"),
        ("5622681900:01:001:008A", "non-digit"),
        ("5622681900-01-001-0087", "other than digits"),
        ("5.6226819E+18", "scientific"),
        ("5,62E+18", "scientific"),
        ("５６２２６８１９０００１００１００８７", "other than digits"),
    ],
)
def test_normalize_rejects_with_reason(raw, fragment):
    number, reason = cadnum.normalize(raw)
    assert number is None
    assert fragment in reason


@pytest.mark.parametrize("value", [5622681900010010087, 5.6226819e18])
def test_numeric_cells_are_rejected_not_guessed(value):
    number, reason = cadnum.normalize(value)
    assert number is None
    assert "15 digits" in reason


def test_process_marks_duplicates_across_formats():
    rows = cadnum.process(
        [(1, GOOD), (2, "5622681900010010087"), (3, "bad"), (4, ""), (5, "0520880300020010015")]
    )
    assert [r.status for r in rows] == ["ok", "duplicate", "invalid", "empty", "ok"]
    assert rows[1].number == GOOD
    assert rows[1].reason == "same as row 1"
    assert rows[2].number is None and rows[2].raw == "bad"


def test_read_csv_with_header_semicolons_and_cp1251(tmp_path):
    path = tmp_path / "in.csv"
    path.write_bytes("Назва;Кадастровий номер\nДілянка 1;5622681900:01:001:0087\n".encode("cp1251"))
    values = cadnum.read_csv(path, column="кадастровий номер")
    assert values == [(2, GOOD)]


def test_read_csv_without_header_uses_first_column(tmp_path):
    path = tmp_path / "in.csv"
    path.write_text("5622681900010010087\n5622681900:01:001:0088\n", encoding="utf-8")
    assert [v for _, v in cadnum.read_csv(path)] == ["5622681900010010087", "5622681900:01:001:0088"]
    assert [n for n, _ in cadnum.read_csv(path)] == [1, 2]


def test_unknown_column_is_an_error(tmp_path):
    path = tmp_path / "in.csv"
    path.write_text("number\n5622681900010010087\n", encoding="utf-8")
    with pytest.raises(ValueError, match="not found"):
        cadnum.read_csv(path, column="missing")


def test_cli_excel_round_trip(tmp_path):
    source = tmp_path / "list.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Cadastral number"])
    sheet.append([GOOD])
    sheet.append(["5622681900010010087"])  # duplicate in the other format
    sheet.append([5622681900010010088])  # numeric cell: damaged by Excel
    sheet.append(["0520880300020010015"])
    workbook.save(source)

    out = tmp_path / "out"
    assert cadnum.main([str(source), "--out", str(out)]) == 0

    data = json.loads((out / "list-cleaned.json").read_text(encoding="utf-8"))
    assert data["summary"] == {"ok": 2, "duplicate": 1, "invalid": 1}
    assert data["numbers"] == [GOOD, "0520880300:02:001:0015"]
    assert [r["row"] for r in data["rows"]] == [2, 3, 4, 5]

    result = load_workbook(out / "list-cleaned.xlsx")
    numbers = [row[0] for row in result["Numbers"].iter_rows(min_row=2, values_only=True)]
    assert numbers == [GOOD, "0520880300:02:001:0015"]
    assert result["Numbers"]["A3"].number_format == "@"
    report = list(result["Report"].iter_rows(min_row=2, values_only=True))
    assert [r[3] for r in report] == ["ok", "duplicate", "invalid", "ok"]


def test_cli_returns_2_for_missing_or_unsupported_file(tmp_path):
    assert cadnum.main([str(tmp_path / "none.csv"), "--out", str(tmp_path)]) == 2
    other = tmp_path / "data.pdf"
    other.write_text("x")
    assert cadnum.main([str(other), "--out", str(tmp_path)]) == 2
