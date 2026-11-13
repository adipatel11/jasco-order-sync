"""Append parsed ODS rows to the Pending sheet of Order.xlsx.

Optionally replaces whole dates first: every data row whose column-F date is one of
`replace_dates` is removed (the rows below close up the gap) before the fresh batches
are appended, so re-copying a day yields exactly what TAP holds now — no stale rows.
"""

from __future__ import annotations

import datetime as dt
import shutil
from copy import copy
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from ods_parser import OdsRow

PENDING_SHEET = "Pending"
# The sheet's real layout: row 1 is a stray leftover data row, row 2 the "Pending"
# title, row 3 the =TODAY()/=SUM() line, and row 4 the column headers ("Item code",
# "Name", …). So appended data starts at row 5. This was 4 — which happened to work,
# because _next_empty_row only scans for the first blank in column A and row 4's
# header is not blank, but it made the dedupe read the header cell as an order.
FIRST_DATA_ROW = 5
ORDER_HEADER_LABEL = "Order"  # the column-E header, never a real order number
LAST_FORMAT_COL = 7  # copy formatting for columns A-G (G stays empty but needs borders)
DATE_FORMAT = "m/d/yyyy"

# The size lookup table the generated column-C formula points at.
SIZE_DATA_SHEET = "SizeData"
# Used only if that sheet is missing; matches the range the formula used to hardcode.
SIZE_DATA_FALLBACK_ROWS = 3974


def _size_data_rows(wb) -> int:
    """Last populated row of SizeData, for the VLOOKUP range.

    This used to be hardcoded to 3974, which silently stopped working the moment a
    3975th item was added: the lookup would miss and every new row's size would come
    back #N/A with nothing raising an error. Measuring the sheet keeps the generated
    formula correct as the catalogue grows.
    """
    if SIZE_DATA_SHEET not in wb.sheetnames:
        return SIZE_DATA_FALLBACK_ROWS
    ws = wb[SIZE_DATA_SHEET]
    # Walk up from max_row: it can be inflated by stray formatting on empty cells.
    for row in range(ws.max_row, 0, -1):
        if ws.cell(row=row, column=1).value not in (None, ""):
            return row
    return SIZE_DATA_FALLBACK_ROWS


def _row_style(ws, row: int) -> list[tuple]:
    """Snapshot the formatting of cols A-G of `row`, to paint onto other rows later.

    A snapshot rather than a row reference because replacing a date can delete the
    very row that would otherwise have served as the formatting template.
    """
    style = []
    for col in range(1, LAST_FORMAT_COL + 1):
        src = ws.cell(row=row, column=col)
        style.append((copy(src.font), copy(src.fill), copy(src.border),
                      copy(src.alignment), copy(src.protection), src.number_format))
    return style


def _apply_row_style(ws, row: int, style: list[tuple]) -> None:
    """Mirror of the owner's manual "format painter" step: give `row` the snapshot's
    look (font, fill, border, alignment, number format)."""
    for col, (font, fill, border, alignment, protection, number_format) in enumerate(
        style, start=1
    ):
        dst = ws.cell(row=row, column=col)
        dst.font = copy(font)
        dst.fill = copy(fill)
        dst.border = copy(border)
        dst.alignment = copy(alignment)
        dst.protection = copy(protection)
        dst.number_format = number_format


@dataclass
class OrderBatch:
    order_number: str
    date: dt.date
    rows: list[OdsRow]


def _existing_order_numbers(ws) -> set[str]:
    """Every order number already in column E, wherever in the sheet it sits.

    Deliberately scans the whole column instead of starting at FIRST_DATA_ROW. The
    sheet carries a stray data row at row 1 whose order would otherwise be invisible
    here, and the cost of the two mistakes is wildly asymmetric: an over-inclusive
    set can only cause an order to be skipped as already-present, while a missing one
    double-enters it into the owner's book. The header cell is the one value in the
    column that is never an order number, so it is the only thing filtered out.
    """
    seen: set[str] = set()
    for row in ws.iter_rows(min_col=5, max_col=5, values_only=True):
        val = row[0]
        if not val:
            continue
        text = str(val).strip()
        if text and text != ORDER_HEADER_LABEL:
            seen.add(text)
    return seen


def _next_empty_row(ws) -> int:
    n = FIRST_DATA_ROW
    while ws.cell(row=n, column=1).value not in (None, ""):
        n += 1
    return n


def _cell_date(value) -> dt.date | None:
    """Column F as a date. Rows we wrote hold real dates; a hand-typed one may be text."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _remove_rows_for_dates(ws, dates: set[dt.date]) -> int:
    """Delete every data row whose column-F date is in `dates`; return how many.

    Not ws.delete_rows: openpyxl doesn't rewrite formulas when it shifts rows, so
    every column-C VLOOKUP below a deleted row would keep pointing at its OLD row
    number — i.e. at a different item's code. Instead the kept rows are slid up with
    move_range(translate=True), which re-targets each formula to its new row exactly
    as Excel's own "Delete row" would.

    Only the data block (FIRST_DATA_ROW down to the first blank in column A — the
    same block appends extend) is touched; the title/header rows above never move.
    """
    last = _next_empty_row(ws) - 1
    if last < FIRST_DATA_ROW or not dates:
        return 0
    keep = [r for r in range(FIRST_DATA_ROW, last + 1)
            if _cell_date(ws.cell(row=r, column=6).value) not in dates]
    removed = (last - FIRST_DATA_ROW + 1) - len(keep)
    if removed == 0:
        return 0

    # Move the whole used width, not just A-G, so anything the owner keeps in a
    # column further right stays on the same line as its order.
    width = LAST_FORMAT_COL
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW, max_row=last):
        for cell in row:
            if cell.value not in (None, "") and cell.column > width:
                width = cell.column
    last_col = get_column_letter(width)
    # What an empty row below the data looks like, for the rows the block gives up.
    blank_style = _row_style(ws, last + 1)

    dest = FIRST_DATA_ROW
    for r in keep:
        if r != dest:
            # move_range moves every cell in the range (creating empty ones as
            # needed), so the destination row is fully overwritten — nothing from
            # the deleted row it lands on survives.
            ws.move_range(f"A{r}:{last_col}{r}", rows=dest - r, translate=True)
        dest += 1
    for r in range(dest, last + 1):
        for c in range(1, width + 1):
            ws.cell(row=r, column=c).value = None
        _apply_row_style(ws, r, blank_style)
    return removed


def write_orders(
    master_xlsx: Path,
    batches: list[OrderBatch],
    backups_dir: Path,
    replace_dates: set[dt.date] | None = None,
) -> tuple[Path | None, int, int, int]:
    """Write batches to the master workbook IN PLACE.

    If replace_dates is given, every existing data row dated one of those days is
    deleted first, so the batches for those days replace them rather than being
    skipped as already present.

    Before overwriting, a pristine copy of the master is saved to backups_dir as a
    restore point. Returns (backup_path, rows_added, orders_added, rows_removed);
    backup_path is None when nothing changed (master left untouched).

    Skips any batch whose order_number is (still) present in column E.
    """
    wb = load_workbook(master_xlsx)
    if PENDING_SHEET not in wb.sheetnames:
        raise ValueError(f"{master_xlsx} has no '{PENDING_SHEET}' sheet")
    ws = wb[PENDING_SHEET]

    # Template for formatting: the last existing data row, snapshotted BEFORE any
    # removal (which may delete that row). None if the sheet has no data rows.
    last_data_row = _next_empty_row(ws) - 1
    template = _row_style(ws, last_data_row) if last_data_row >= FIRST_DATA_ROW else None

    rows_removed = _remove_rows_for_dates(ws, replace_dates or set())

    seen = _existing_order_numbers(ws)
    next_row = _next_empty_row(ws)
    size_rows = _size_data_rows(wb)

    rows_added = 0
    orders_added = 0
    for batch in batches:
        if batch.order_number in seen:
            continue
        if not batch.rows:
            continue
        for r in batch.rows:
            ws.cell(row=next_row, column=1, value=r.item_no)
            ws.cell(row=next_row, column=2, value=r.name)
            ws.cell(
                row=next_row,
                column=3,
                value=(
                    f"=VLOOKUP(A{next_row},{SIZE_DATA_SHEET}!A$1:B${size_rows},2,FALSE)"
                ),
            )
            ws.cell(row=next_row, column=4, value=r.reserved_qty)
            ws.cell(row=next_row, column=5, value=batch.order_number)
            ws.cell(row=next_row, column=6, value=batch.date)
            if template is not None:
                _apply_row_style(ws, next_row, template)
            else:
                ws.cell(row=next_row, column=6).number_format = DATE_FORMAT
            next_row += 1
            rows_added += 1
        seen.add(batch.order_number)
        orders_added += 1

    if rows_added == 0 and rows_removed == 0:
        # Nothing changed — leave the master completely untouched (no backup, no save).
        return None, 0, 0, 0

    # Save a pristine copy of the master BEFORE overwriting it, as a restore point.
    backups_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_path = backups_dir / f"{master_xlsx.stem}_backup_{stamp}{master_xlsx.suffix}"
    shutil.copy2(master_xlsx, backup_path)

    wb.save(master_xlsx)
    return backup_path, rows_added, orders_added, rows_removed
