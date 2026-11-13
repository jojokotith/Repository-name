"""Open the workbook in LibreOffice, let it recalculate, and compare every
summary number with the same figures computed independently in Python."""

import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import build as report  # noqa: E402

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
pytestmark = pytest.mark.skipif(SOFFICE is None, reason="LibreOffice is needed to recalculate formulas")


def recalculate(source: Path, out_dir: Path) -> Path:
    subprocess.run(
        [SOFFICE, "--headless", f"-env:UserInstallation=file://{out_dir}/lo-profile",
         "--convert-to", "xlsx", "--outdir", str(out_dir), str(source)],
        check=True, capture_output=True, timeout=180,
    )
    return out_dir / source.name


def read_summary(path: Path) -> tuple[dict, list]:
    sheet = load_workbook(path, data_only=True)[report.SUMMARY]
    rows = {}
    for line in sheet.iter_rows(min_row=5, max_row=4 + len(report.MANAGERS), values_only=True):
        rows[line[0]] = dict(zip(("orders", "contracts", "type2", "net", "open", "amount"), line[1:7]))
    total = [c.value for c in sheet[5 + len(report.MANAGERS)]][:7]
    return rows, total


def build_and_read(tmp_path, orders, start=None, end=None):
    source = tmp_path / "src" / "report.xlsx"
    source.parent.mkdir()
    report.build(source, orders)
    if start:
        workbook = load_workbook(source)
        workbook[report.SUMMARY]["B1"], workbook[report.SUMMARY]["B2"] = start, end
        workbook.save(source)
    out = tmp_path / "out"
    out.mkdir()
    return read_summary(recalculate(source, out))


def test_summary_matches_independent_calculation(tmp_path):
    orders = report.sample_orders()
    rows, total = build_and_read(tmp_path, orders)
    expected = report.expected_summary(orders, report.PERIOD_START, report.PERIOD_END)
    assert rows == expected
    assert total[0] == "Разом"
    for index, key in enumerate(("orders", "contracts", "type2", "net", "open", "amount"), start=1):
        assert total[index] == sum(r[key] for r in expected.values())
    # The demo data must actually exercise the logic.
    assert sum(r["type2"] for r in expected.values()) > 0
    assert sum(r["open"] for r in expected.values()) > 0


def test_changing_the_period_changes_the_numbers(tmp_path):
    orders = report.sample_orders()
    start, end = date(2026, 11, 10), date(2026, 11, 16)
    rows, _ = build_and_read(tmp_path, orders, start, end)
    assert rows == report.expected_summary(orders, start, end)
    full = report.expected_summary(orders, report.PERIOD_START, report.PERIOD_END)
    assert rows != full


def test_client_example_manager_509(tmp_path):
    """Two orders, one contract, one not calculated - the case from the brief."""
    day = date(2026, 11, 5)
    orders = [
        {"date": day, "manager": 509, "order": "З-0001", "amount": 12000, "contract": report.TYPE_1, "calculated": report.YES},
        {"date": day, "manager": 509, "order": "З-0002", "amount": None, "contract": "", "calculated": report.NO},
        {"date": day, "manager": 510, "order": "З-0003", "amount": 5000, "contract": report.TYPE_2, "calculated": report.YES},
    ]
    rows, _ = build_and_read(tmp_path, orders)
    assert rows[509] == {"orders": 2, "contracts": 1, "type2": 0, "net": 1, "open": 1, "amount": 12000}
    assert rows[510] == {"orders": 1, "contracts": 1, "type2": 1, "net": 0, "open": 0, "amount": 5000}
    assert rows[501]["orders"] == 0
