"""Append parsed ODS rows to the Pending sheet of Order.xlsx."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from ods_parser import OdsRow

PENDING_SHEET = "Pending"
FIRST_DATA_ROW = 4  # rows 1-3 are header/title/sum
EXCEL_EPOCH = dt.date(1899, 12, 30)  # accounts for the 1900 leap year bug


@dataclass
class OrderBatch:
    order_number: str
    date: dt.date
    rows: list[OdsRow]


def _excel_serial(d: dt.date) -> int:
    return (d - EXCEL_EPOCH).days


def _existing_order_numbers(ws) -> set[str]:
    seen: set[str] = set()
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW, min_col=5, max_col=5, values_only=True):
        val = row[0]
        if val:
            seen.add(str(val).strip())
    return seen


def _next_empty_row(ws) -> int:
    n = FIRST_DATA_ROW
    while ws.cell(row=n, column=1).value not in (None, ""):
        n += 1
    return n


def write_orders(
    source_xlsx: Path,
    batches: list[OrderBatch],
    backups_dir: Path,
) -> tuple[Path, int, int]:
    """Append batches to a copy of source_xlsx. Returns (output_path, rows_added, orders_added).

    Skips any batch whose order_number is already present in column E.
    """
    wb = load_workbook(source_xlsx)
    if PENDING_SHEET not in wb.sheetnames:
        raise ValueError(f"{source_xlsx} has no '{PENDING_SHEET}' sheet")
    ws = wb[PENDING_SHEET]

    seen = _existing_order_numbers(ws)
    next_row = _next_empty_row(ws)

    rows_added = 0
    orders_added = 0
    for batch in batches:
        if batch.order_number in seen:
            continue
        if not batch.rows:
            continue
        serial = _excel_serial(batch.date)
        for r in batch.rows:
            ws.cell(row=next_row, column=1, value=r.item_no)
            ws.cell(row=next_row, column=2, value=r.name)
            ws.cell(
                row=next_row,
                column=3,
                value=f"=VLOOKUP(A{next_row},SizeData!A$1:B$3974,2,FALSE)",
            )
            ws.cell(row=next_row, column=4, value=r.reserved_qty)
            ws.cell(row=next_row, column=5, value=batch.order_number)
            ws.cell(row=next_row, column=6, value=serial)
            next_row += 1
            rows_added += 1
        seen.add(batch.order_number)
        orders_added += 1

    backups_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    out_path = backups_dir / f"Order_{stamp}.xlsx"
    wb.save(out_path)
    return out_path, rows_added, orders_added
