"""The staging workbook, wherever it lives — one call site for run.py and pick.py.

Both entry points used to take a local Path and hand it straight to write_orders.
Now the workbook normally lives in the owner's OneDrive, so this module owns the
"where is it, fetch it, put it back" part and leaves xlsx_writer to do exactly what
it always did to a local .xlsx file.

Local mode is still supported (ORDER_XLSX_PATH with no OneDrive settings) because it
is how you test a change without touching the real workbook.
"""

from __future__ import annotations

import logging
import os
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path

import onedrive
from onedrive import GraphAuthRequiredError, WorkbookChangedError  # re-exported for callers
from xlsx_writer import OrderBatch, write_orders

# run.py on the VPS and pick.py on a desktop are different machines, so the repo's
# flock cannot keep them apart. If they collide, the loser's upload is refused on its
# eTag and it starts over against the fresh copy. write_orders is idempotent and
# append-only, so replaying it is safe — the retry simply re-skips whatever the
# winner already wrote.
MAX_UPLOAD_ATTEMPTS = 3

log = logging.getLogger("workbook")


@dataclass
class WriteResult:
    backup_path: Path | None
    rows_added: int
    orders_added: int
    web_url: str | None  # the owner's bookmark, when the workbook is in OneDrive


def in_cloud() -> bool:
    return bool(os.environ.get("ONEDRIVE_SHARE_LINK") or os.environ.get("ONEDRIVE_FILE_PATH"))


def check_config() -> str | None:
    """Validate that the workbook is reachable. Returns an error message, or None.

    Called at startup by both entry points so a misconfiguration fails immediately,
    rather than after a full scrape has already been done.
    """
    if in_cloud():
        if not os.environ.get("GRAPH_CLIENT_ID"):
            return "ONEDRIVE_* is set but GRAPH_CLIENT_ID is missing from .env"
        try:
            ref = onedrive.resolve()
        except GraphAuthRequiredError:
            return onedrive.REAUTH_MESSAGE
        except Exception as e:  # noqa: BLE001 — surfaced verbatim to the operator
            return f"Could not find the workbook in OneDrive: {e}"
        log.info("Workbook: %s in OneDrive (%s)", ref.name, ref.web_url)
        return None

    raw = os.environ.get("ORDER_XLSX_PATH")
    if not raw:
        return "Set ONEDRIVE_SHARE_LINK (or ORDER_XLSX_PATH for a local copy) in .env"
    if not Path(raw).expanduser().exists():
        return f"ORDER_XLSX_PATH does not exist: {raw}"
    return None


def web_url() -> str | None:
    """The owner's bookmark link, or None in local mode / if it can't be resolved.

    Best-effort by design: this only ever decorates the UI and the Telegram ping, so
    a lookup failure must never be what stops a sync from running.
    """
    if not in_cloud():
        return None
    try:
        return onedrive.resolve().web_url
    except Exception:  # noqa: BLE001 — cosmetic; real failures surface in check_config
        log.debug("Could not resolve the workbook link", exc_info=True)
        return None


def append(batches: list[OrderBatch], backups_dir: Path) -> WriteResult:
    """Append batches to the staging workbook and return what changed."""
    if in_cloud():
        return _append_cloud(batches, backups_dir)

    path = Path(os.environ["ORDER_XLSX_PATH"]).expanduser()
    backup, rows, orders = write_orders(path, batches, backups_dir)
    return WriteResult(backup, rows, orders, None)


def _append_cloud(batches: list[OrderBatch], backups_dir: Path) -> WriteResult:
    ref = onedrive.resolve()

    for attempt in range(1, MAX_UPLOAD_ATTEMPTS + 1):
        with tempfile.TemporaryDirectory(prefix="jasco-wb-") as tmp:
            local = Path(tmp) / ref.name
            etag = onedrive.download(ref, local)

            # write_orders snapshots the pristine download into backups_dir before
            # saving, so the restore point captures the exact cloud state we edited.
            backup, rows, orders = write_orders(local, batches, backups_dir)
            if rows == 0:
                # Nothing new — leave the cloud copy completely untouched.
                return WriteResult(None, 0, 0, ref.web_url)

            try:
                onedrive.upload(ref, local, if_match=etag)
            except WorkbookChangedError:
                # Someone else wrote between our download and our upload. Drop the
                # backup for this attempt — it documents a state we never modified —
                # and redo the append against the copy that actually won.
                if backup and backup.exists():
                    backup.unlink()
                if attempt == MAX_UPLOAD_ATTEMPTS:
                    raise
                log.warning(
                    "%s changed in OneDrive mid-edit; retrying (attempt %d/%d)",
                    ref.name, attempt, MAX_UPLOAD_ATTEMPTS,
                )
                continue

            return WriteResult(backup, rows, orders, ref.web_url)

    raise WorkbookChangedError(f"Gave up after {MAX_UPLOAD_ATTEMPTS} attempts on {ref.name}")


def fetch_copy(dest: Path) -> Path:
    """Download the current cloud workbook to dest, for inspection. Local mode copies.

    Not used by the sync itself — it exists so `python master_workbook.py` can pull
    the workbook down to look at without going through OneDrive in a browser.
    """
    if in_cloud():
        ref = onedrive.resolve()
        onedrive.download(ref, dest)
    else:
        shutil.copy2(Path(os.environ["ORDER_XLSX_PATH"]).expanduser(), dest)
    return dest


def main() -> int:
    import sys

    from dotenv import load_dotenv

    load_dotenv(Path(__file__).parent / ".env")
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    problem = check_config()
    if problem:
        print(problem, file=sys.stderr)
        return 2
    print("Workbook configuration OK —", "OneDrive" if in_cloud() else "local file")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
