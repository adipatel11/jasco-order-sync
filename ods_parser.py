"""Parse a TAP Export.ods and return the rows we care about."""

from __future__ import annotations

import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path


@dataclass
class OdsRow:
    item_no: int
    name: str
    reserved_qty: float


_ROW_RE = re.compile(r"<table:table-row[^>]*>(.*?)</table:table-row>", re.DOTALL)
_CELL_RE = re.compile(r"<table:table-cell[^>]*?(?:\s/>|>(.*?)</table:table-cell>)", re.DOTALL)
_TEXT_RE = re.compile(r"<text:p[^>]*>(.*?)</text:p>", re.DOTALL)
_TAG_RE = re.compile(r"<[^>]+>")
_REPEAT_RE = re.compile(r'table:number-columns-repeated="(\d+)"')


def _cell_text(cell_xml: str) -> str:
    parts = _TEXT_RE.findall(cell_xml or "")
    return " ".join(_TAG_RE.sub("", p).replace("&amp;", "&").strip() for p in parts).strip()


def _row_cells(row_xml: str) -> list[str]:
    """Return cell text values for a row, expanding repeated empty cells."""
    cells: list[str] = []
    for match in re.finditer(
        r"<table:table-cell([^>]*?)(?:\s/>|>(.*?)</table:table-cell>)",
        row_xml,
        re.DOTALL,
    ):
        attrs, body = match.group(1), match.group(2)
        text = _cell_text(body or "")
        repeat_match = _REPEAT_RE.search(attrs)
        repeat = int(repeat_match.group(1)) if repeat_match else 1
        # Cap absurd repeats (some ODS files pad rows with thousands of empties).
        repeat = min(repeat, 50)
        cells.extend([text] * repeat)
    return cells


def parse_ods(path: str | Path) -> list[OdsRow]:
    """Return data rows from the ODS, skipping header (row 0) and totals row (last)."""
    with zipfile.ZipFile(path) as z:
        with z.open("content.xml") as f:
            content = f.read().decode("utf-8")

    rows_xml = _ROW_RE.findall(content)
    parsed: list[OdsRow] = []

    for i, row_xml in enumerate(rows_xml):
        if i == 0:
            continue  # header
        cells = _row_cells(row_xml)
        if len(cells) < 8:
            continue
        item_no_raw = cells[0].strip()
        name = cells[1].strip()
        reserved_raw = cells[7].strip()  # column H

        if not item_no_raw or not name:
            # totals row at the end has blank item# and name
            continue

        try:
            item_no = int(item_no_raw)
        except ValueError:
            continue
        try:
            reserved_qty = float(reserved_raw)
        except ValueError:
            continue

        parsed.append(OdsRow(item_no=item_no, name=name, reserved_qty=reserved_qty))

    return parsed


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: python ods_parser.py <path-to.ods>", file=sys.stderr)
        return 1
    rows = parse_ods(sys.argv[1])
    print(f"Parsed {len(rows)} rows:")
    for r in rows:
        print(f"  {r.item_no:>7}  {r.reserved_qty:>5}  {r.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
