"""Entrypoint: pull yesterday's TAP orders and append to Order.xlsx."""

from __future__ import annotations

import datetime as dt
import logging
import os
import shutil
import sys
from pathlib import Path

from dotenv import load_dotenv

from ods_parser import parse_ods
from tap_scraper import apply_filter, chrome_session, iter_orders, load_or_login
from xlsx_writer import OrderBatch, write_orders

ROOT = Path(__file__).parent
DOWNLOADS = ROOT / "downloads"
BACKUPS = ROOT / "backups"
LOGS = ROOT / "logs"


def setup_logging() -> None:
    LOGS.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d")
    handler_file = logging.FileHandler(LOGS / f"run_{stamp}.log")
    handler_stream = logging.StreamHandler(sys.stdout)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[handler_file, handler_stream],
    )


def clear_downloads() -> None:
    if DOWNLOADS.exists():
        shutil.rmtree(DOWNLOADS)
    DOWNLOADS.mkdir(parents=True, exist_ok=True)


def main() -> int:
    load_dotenv(ROOT / ".env")
    setup_logging()
    log = logging.getLogger("run")

    username = os.environ.get("TAP_USERNAME")
    password = os.environ.get("TAP_PASSWORD")
    xlsx_path_raw = os.environ.get("ORDER_XLSX_PATH")
    headless = os.environ.get("HEADLESS", "false").lower() == "true"

    if not (username and password and xlsx_path_raw):
        log.error("Missing TAP_USERNAME, TAP_PASSWORD, or ORDER_XLSX_PATH in .env")
        return 2
    xlsx_path = Path(xlsx_path_raw).expanduser()
    if not xlsx_path.exists():
        log.error("ORDER_XLSX_PATH does not exist: %s", xlsx_path)
        return 2

    target_date = dt.date.today() - dt.timedelta(days=1)
    log.info("Target submitted date: %s", target_date.strftime("%m-%d-%Y"))

    clear_downloads()

    batches: list[OrderBatch] = []
    with chrome_session(DOWNLOADS, headless=headless) as page:
        load_or_login(page, username, password)
        apply_filter(page, target_date)

        for handle in iter_orders(page, DOWNLOADS):
            try:
                rows = parse_ods(handle.ods_path)
            except Exception as e:
                log.exception("Failed to parse %s: %s", handle.ods_path, e)
                continue
            log.info(
                "Order %s: %d line items from %s",
                handle.order_number,
                len(rows),
                handle.ods_path.name,
            )
            batches.append(
                OrderBatch(order_number=handle.order_number, date=target_date, rows=rows)
            )

    if not batches:
        log.info("No orders found for %s — nothing to do", target_date)
        return 0

    backup_path, rows_added, orders_added = write_orders(xlsx_path, batches, BACKUPS)
    skipped = len(batches) - orders_added
    if rows_added == 0:
        log.info(
            "Nothing new for %s — master left untouched (%d duplicate orders skipped)",
            target_date,
            skipped,
        )
        return 0
    log.info(
        "Backed up master to %s; appended %d rows from %d orders into %s "
        "(%d duplicate orders skipped)",
        backup_path,
        rows_added,
        orders_added,
        xlsx_path,
        skipped,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
