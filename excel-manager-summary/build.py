"""Build an Excel workbook with a self-updating summary by manager.

The "Дані" sheet holds one order per row. The "Зведення" sheet counts orders,
contracts and not-yet-calculated orders per manager for a chosen period using
plain formulas (no macros), so it also works in Excel for the web and in
Google Sheets. New rows are picked up automatically because every formula
covers the whole column.

Usage:
    python build.py report.xlsx
"""

from __future__ import annotations

import random
import sys
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

DATA = "Дані"
SUMMARY = "Зведення"
TYPE_1 = "Тип 1"
TYPE_2 = "Тип 2"
YES = "Так"
NO = "Ні"

MANAGERS = [501, 502, 503, 504, 505, 506, 507, 508, 509, 510, 511, 512]
PERIOD_START = date(2026, 11, 1)
PERIOD_END = date(2026, 11, 30)

HEADER_FILL = PatternFill("solid", fgColor="1F7A4D")
HEADER_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF4CC")
TOTAL_FILL = PatternFill("solid", fgColor="E8F3EC")
THIN = Side(style="thin", color="C9D3CD")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def sample_orders(seed: int = 7, count: int = 90) -> list[dict]:
    """Deterministic demo data: some orders fall outside the default period."""
    rng = random.Random(seed)
    orders = []
    for number in range(1, count + 1):
        day = date(2026, 10, 20) + timedelta(days=rng.randrange(0, 52))
        calculated = rng.random() < 0.8
        contract = ""
        if calculated and rng.random() < 0.6:
            contract = TYPE_2 if rng.random() < 0.25 else TYPE_1
        orders.append(
            {
                "date": day,
                "manager": rng.choice(MANAGERS),
                "order": f"З-{number:04d}",
                "amount": rng.randrange(8, 120) * 500 if calculated else None,
                "contract": contract,
                "calculated": YES if calculated else NO,
            }
        )
    orders.sort(key=lambda o: (o["date"], o["order"]))
    return orders


def expected_summary(orders: list[dict], start: date, end: date) -> dict[int, dict]:
    """The numbers the summary sheet must show, computed independently in Python."""
    result = {m: {"orders": 0, "contracts": 0, "type2": 0, "net": 0, "open": 0, "amount": 0} for m in MANAGERS}
    for o in orders:
        if not (start <= o["date"] <= end):
            continue
        row = result[o["manager"]]
        row["orders"] += 1
        row["contracts"] += o["contract"] in (TYPE_1, TYPE_2)
        row["type2"] += o["contract"] == TYPE_2
        row["open"] += o["calculated"] == NO
        row["amount"] += o["amount"] or 0
    for row in result.values():
        row["net"] = row["contracts"] - row["type2"]
    return result


def _style_header(sheet, row: int, columns: int) -> None:
    for column in range(1, columns + 1):
        cell = sheet.cell(row=row, column=column)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def build(path: Path, orders: list[dict] | None = None) -> None:
    orders = sample_orders() if orders is None else orders
    workbook = Workbook()

    data = workbook.active
    data.title = DATA
    data.append(["Дата", "Менеджер", "Замовлення", "Вартість, грн", "Договір", "Прораховано"])
    for o in orders:
        data.append([o["date"], o["manager"], o["order"], o["amount"], o["contract"] or None, o["calculated"]])
    _style_header(data, 1, 6)
    for column, width in zip("ABCDEF", (13, 12, 14, 16, 12, 14)):
        data.column_dimensions[column].width = width
    for row in data.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "DD.MM.YYYY"
    data.freeze_panes = "A2"
    data.auto_filter.ref = f"A1:F{len(orders) + 1}"

    # Dropdowns keep the values the formulas rely on free of typos.
    contract_list = DataValidation(type="list", formula1=f'"{TYPE_1},{TYPE_2}"', allow_blank=True)
    done_list = DataValidation(type="list", formula1=f'"{YES},{NO}"', allow_blank=True)
    contract_list.add("E2:E5000")
    done_list.add("F2:F5000")
    data.add_data_validation(contract_list)
    data.add_data_validation(done_list)

    summary = workbook.create_sheet(SUMMARY)
    summary["A1"], summary["B1"] = "Період з", PERIOD_START
    summary["A2"], summary["B2"] = "Період по", PERIOD_END
    for cell in (summary["B1"], summary["B2"]):
        cell.number_format = "DD.MM.YYYY"
        cell.fill = INPUT_FILL
        cell.border = BORDER
    summary["A1"].font = summary["A2"].font = Font(bold=True)
    summary["D1"] = "Змініть дати в жовтих клітинках — таблиця перерахується сама."

    header_row = 4
    headers = ["Менеджер", "Замовлень", "Договорів усього", "з них Тип 2",
               "Договорів без Тип 2", "Не прораховано", "Вартість, грн"]
    for column, title in enumerate(headers, start=1):
        summary.cell(row=header_row, column=column, value=title)
    _style_header(summary, header_row, len(headers))
    summary.row_dimensions[header_row].height = 32

    d = f"'{DATA}'"
    period = f'{d}!$A:$A,">="&$B$1,{d}!$A:$A,"<="&$B$2'
    first = header_row + 1
    for offset, manager in enumerate(MANAGERS):
        r = first + offset
        who = f"{d}!$B:$B,$A{r}"
        summary.cell(row=r, column=1, value=manager)
        summary.cell(row=r, column=2, value=f"=COUNTIFS({who},{period})")
        summary.cell(row=r, column=3, value=f'=COUNTIFS({who},{d}!$E:$E,"<>",{period})')
        summary.cell(row=r, column=4, value=f'=COUNTIFS({who},{d}!$E:$E,"{TYPE_2}",{period})')
        summary.cell(row=r, column=5, value=f"=C{r}-D{r}")
        summary.cell(row=r, column=6, value=f'=COUNTIFS({who},{d}!$F:$F,"{NO}",{period})')
        summary.cell(row=r, column=7, value=f"=SUMIFS({d}!$D:$D,{who},{period})")
        for column in range(1, len(headers) + 1):
            summary.cell(row=r, column=column).border = BORDER
    last = first + len(MANAGERS) - 1

    total = last + 1
    summary.cell(row=total, column=1, value="Разом")
    for column in range(2, len(headers) + 1):
        letter = summary.cell(row=total, column=column).column_letter
        summary.cell(row=total, column=column, value=f"=SUM({letter}{first}:{letter}{last})")
    for column in range(1, len(headers) + 1):
        cell = summary.cell(row=total, column=column)
        cell.font = Font(bold=True)
        cell.fill = TOTAL_FILL
        cell.border = BORDER
    for r in range(first, total + 1):
        summary.cell(row=r, column=7).number_format = "#,##0"

    for column, width in zip("ABCDEFG", (14, 13, 13, 13, 15, 15, 16)):
        summary.column_dimensions[column].width = width
    summary.freeze_panes = summary.cell(row=first, column=2)

    workbook.active = workbook.index(summary)
    workbook.save(path)


if __name__ == "__main__":
    build(Path(sys.argv[1] if len(sys.argv) > 1 else "report.xlsx"))
