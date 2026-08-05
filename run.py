"""Entrypoint: pull yesterday's TAP orders and append to Order.xlsx."""

from __future__ import annotations

import datetime as dt
import logging
import os
import shutil
import sys
from pathlib import Path

try:
    import fcntl  # Unix only; the Windows picker machine has no fcntl (see _lock)
except ImportError:
    fcntl = None

from dotenv import load_dotenv

import master_workbook
from ods_parser import parse_ods
from tap_scraper import apply_filter, chrome_session, iter_orders, load_or_login
from xlsx_writer import OrderBatch

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


def notify(text: str) -> None:
    """Best-effort Telegram ping. A silent no-op unless TG_* is configured.

    An unattended VPS run has nobody watching the terminal, so a failure that only
    lands in a log file is a failure nobody hears about. Imported lazily so the Mac
    and the Windows picker, which don't configure Telegram, never need `requests`.
    """
    log = logging.getLogger("run")
    if not (os.environ.get("TG_BOT_TOKEN") and os.environ.get("TG_CHAT_ID")):
        return
    try:
        import telegram

        telegram.send(text)
    except Exception:
        log.exception("Could not send the Telegram status message")


def _lock():
    """Take the single-instance lock, or return None if someone else holds it.

    Two processes driving one Chromium profile corrupt it. On the VPS TAP_LOCK
    points at the stock radar's own lock file, so the two projects share a single
    mutex over the shared browser profile and can never run at the same time.
    Holding the open fd is what holds the flock; it's released when we exit.
    """
    if fcntl is None:  # Windows: only the picker runs there, never concurrently
        return open(ROOT / ".lock", "w")
    lock_fd = open(Path(os.environ.get("TAP_LOCK", ROOT / ".lock")), "w")
    try:
        fcntl.flock(lock_fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        return None
    return lock_fd


def main() -> int:
    load_dotenv(ROOT / ".env")
    setup_logging()
    log = logging.getLogger("run")

    username = os.environ.get("TAP_USERNAME")
    password = os.environ.get("TAP_PASSWORD")
    headless = os.environ.get("HEADLESS", "false").lower() == "true"

    if not (username and password):
        log.error("Missing TAP_USERNAME or TAP_PASSWORD in .env")
        return 2
    # Check the workbook is reachable BEFORE scraping — a bad OneDrive config or a
    # lapsed sign-in should cost a second, not a full run against TAP.
    problem = master_workbook.check_config()
    if problem:
        log.error("%s", problem)
        return 2

    lock_fd = _lock()
    if lock_fd is None:
        log.error("Another TAP automation is already running (lock: %s) — two would "
                  "fight over the browser profile. On the VPS the stock radar holds "
                  "this lock; stop it first (systemctl stop stock-radar).",
                  os.environ.get("TAP_LOCK", ROOT / ".lock"))
        return 1

    target_date = dt.date.today() - dt.timedelta(days=1)
    log.info("Target submitted date: %s", target_date.strftime("%m-%d-%Y"))

    clear_downloads()

    # No terminal (e.g. a scheduled launchd run) → don't block on the MFA prompt.
    interactive = sys.stdin.isatty()

    batches: list[OrderBatch] = []
    with chrome_session(DOWNLOADS, headless=headless) as page:
        load_or_login(page, username, password, interactive=interactive)
        expected = apply_filter(page, target_date)

        for handle in iter_orders(page, DOWNLOADS, expected_total=expected):
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

    result = master_workbook.append(batches, BACKUPS)
    skipped = len(batches) - result.orders_added
    if result.rows_added == 0:
        log.info(
            "Nothing new for %s — workbook left untouched (%d duplicate orders skipped)",
            target_date,
            skipped,
        )
        return 0
    log.info(
        "Backed up to %s; appended %d rows from %d orders (%d duplicate orders skipped)",
        result.backup_path,
        result.rows_added,
        result.orders_added,
        skipped,
    )
    # The owner reads this on their phone, so include the link — it's the whole point
    # of the workbook living in OneDrive rather than on a server they can't reach.
    message = (
        f"Jasco order sync: appended {result.rows_added} rows from "
        f"{result.orders_added} order(s) for {target_date} "
        f"({skipped} duplicate order(s) skipped)."
    )
    if result.web_url:
        message += f"\n\nOpen the workbook: {result.web_url}"
    notify(message)
    return 0


if __name__ == "__main__":
    try:
        code = main()
    except Exception as e:
        # An unattended VPS run has nobody watching stdout; make sure a crash is heard
        # about rather than dying quietly in a log file.
        logging.getLogger("run").exception("Run failed")
        notify(f"Jasco order sync FAILED: {type(e).__name__}: {e}")
        raise
    raise SystemExit(code)
