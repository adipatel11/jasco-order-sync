"""Interactive picker: choose a date, see that day's TAP orders, tick which to copy.

Companion to run.py (the unattended midnight job). It shares the SAME scraper, ODS
parser, and xlsx writer — the only difference is that the owner chooses the date and
chooses *which* orders get appended to Order.xlsx, through a small Tkinter window.

Two phases, each a short self-contained browser session that reuses the persistent
`.browser_profile` (so in normal use there is no re-login or MFA):
  1. Fetch — filter the chosen date and list its order numbers (no downloads yet).
  2. Copy  — export only the ticked orders, parse them, append to Order.xlsx.

write_orders is idempotent (it skips order numbers already present in column E), so
picking an order the midnight job already captured is harmless — it just gets skipped.

The browser work runs on background threads; results come back to the Tk main loop
through a queue, so the window never freezes while TAP is being driven.
"""

from __future__ import annotations

import calendar
import datetime as dt
import logging
import os
import queue
import shutil
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from dotenv import load_dotenv

from ods_parser import parse_ods
from tap_scraper import (
    MFARequiredError,
    apply_filter,
    chrome_session,
    iter_orders,
    list_order_numbers,
    load_or_login,
)
from xlsx_writer import OrderBatch, write_orders

ROOT = Path(__file__).parent
DOWNLOADS = ROOT / "downloads"
BACKUPS = ROOT / "backups"
LOGS = ROOT / "logs"

# The picker drives the browser headless: in normal use the persistent profile already
# holds the trusted-device cookie (established by the daily run.py), so there is no MFA
# to interact with. If the session has expired we surface a clear message instead of
# popping a browser the owner would have to wrangle. Flip to False to watch it run.
HEADLESS = True

# Shown to the owner whenever TAP needs a fresh login the picker can't perform headless.
REAUTH_MESSAGE = (
    "TAP needs a fresh login.\n\n"
    "Run the daily sync once in a terminal (python run.py) and complete the text-message "
    "code, ticking “Trust this device”. Then try the picker again."
)

log = logging.getLogger("pick")


def setup_logging() -> None:
    LOGS.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[
            logging.FileHandler(LOGS / f"pick_{stamp}.log"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def clear_downloads() -> None:
    if DOWNLOADS.exists():
        shutil.rmtree(DOWNLOADS)
    DOWNLOADS.mkdir(parents=True, exist_ok=True)


class PickerApp:
    """Tkinter front-end. All TAP/browser work happens on worker threads; those push
    tuples onto self.q and the Tk loop drains them in _drain (the only safe way to
    touch widgets from outside the main thread)."""

    def __init__(self, root: tk.Tk, username: str, password: str, xlsx_path: Path):
        self.root = root
        self.username = username
        self.password = password
        self.xlsx_path = xlsx_path
        self.q: queue.Queue = queue.Queue()
        self.busy = False
        self.fetched_date: dt.date | None = None
        self.check_vars: dict[str, tk.BooleanVar] = {}
        self._cal_win: tk.Toplevel | None = None  # the open calendar popup, if any
        self._cal_view = (dt.date.today().year, dt.date.today().month)

        root.title("Jasco Order Picker")
        root.geometry("440x560")
        root.minsize(380, 420)

        self._build_widgets()
        self.root.after(100, self._drain)

    # --- widget construction ----------------------------------------------
    def _build_widgets(self) -> None:
        yesterday = dt.date.today() - dt.timedelta(days=1)

        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(top, text="Date (MM-DD-YYYY):").pack(side="left")
        self.date_var = tk.StringVar(value=yesterday.strftime("%m-%d-%Y"))
        self.date_entry = ttk.Entry(top, textvariable=self.date_var, width=12)
        self.date_entry.pack(side="left", padx=6)
        # Our own calendar popup (a plain Toplevel of buttons) — tkcalendar's
        # drop-down is unusable on macOS, so we roll a small reliable one.
        self.cal_btn = ttk.Button(top, text="Pick…", width=6, command=self._open_calendar)
        self.cal_btn.pack(side="left")
        ttk.Button(top, text="Yesterday",
                   command=lambda: self._set_date(yesterday)).pack(side="left", padx=(6, 0))
        ttk.Button(top, text="Today",
                   command=lambda: self._set_date(dt.date.today())).pack(side="left", padx=4)

        self.fetch_btn = ttk.Button(self.root, text="Fetch orders", command=self.on_fetch)
        self.fetch_btn.pack(fill="x", padx=10)

        mid = ttk.LabelFrame(self.root, text="Orders")
        mid.pack(fill="both", expand=True, padx=10, pady=6)
        self.canvas = tk.Canvas(mid, highlightthickness=0)
        scroll = ttk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        self.list_frame = ttk.Frame(self.canvas)
        self.list_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        sel = ttk.Frame(self.root)
        sel.pack(fill="x", padx=10)
        self.all_btn = ttk.Button(sel, text="Select all",
                                  command=lambda: self._set_all(True), state="disabled")
        self.all_btn.pack(side="left")
        self.none_btn = ttk.Button(sel, text="Clear",
                                   command=lambda: self._set_all(False), state="disabled")
        self.none_btn.pack(side="left", padx=4)

        self.copy_btn = ttk.Button(self.root, text="Copy selected → Excel",
                                   command=self.on_copy, state="disabled")
        self.copy_btn.pack(fill="x", padx=10, pady=(6, 0))

        self.status = tk.StringVar(value="Pick a date and click Fetch.")
        ttk.Label(self.root, textvariable=self.status, relief="sunken",
                  anchor="w").pack(fill="x", side="bottom")

    # --- small helpers ----------------------------------------------------
    def _on_mousewheel(self, event) -> None:
        """Scroll the order list one notch per wheel step on both OSes.

        Tk reports wheel deltas differently per platform: Windows sends multiples of
        ±120, while macOS sends small raw values. The darwin branch keeps the Mac's
        original behavior unchanged; Windows divides by 120 so one notch ≈ one step
        instead of a 120-line jump.
        """
        if sys.platform == "darwin":
            self.canvas.yview_scroll(-event.delta, "units")
        else:
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _set_date(self, d: dt.date) -> None:
        self.date_var.set(d.strftime("%m-%d-%Y"))

    def _set_all(self, value: bool) -> None:
        for var in self.check_vars.values():
            var.set(value)

    def _parse_date(self) -> dt.date:
        return dt.datetime.strptime(self.date_var.get().strip(), "%m-%d-%Y").date()

    # --- calendar popup ---------------------------------------------------
    # A self-contained month grid in a plain Toplevel. We deliberately avoid any
    # focus grab (tkcalendar's grab is what wedges on macOS): the popup is just a
    # window of buttons that closes when a day is clicked or the window is closed.
    def _open_calendar(self) -> None:
        if self.busy:
            return
        if self._cal_win is not None and self._cal_win.winfo_exists():
            self._cal_win.lift()
            return
        try:
            base = self._parse_date()
        except ValueError:
            base = dt.date.today()
        self._cal_view = (base.year, base.month)

        win = tk.Toplevel(self.root)
        win.title("Pick a date")
        win.transient(self.root)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", self._close_calendar)
        self._cal_win = win
        self._render_calendar()

        win.update_idletasks()
        x = self.cal_btn.winfo_rootx()
        y = self.cal_btn.winfo_rooty() + self.cal_btn.winfo_height() + 2
        win.geometry(f"+{x}+{y}")

    def _render_calendar(self) -> None:
        win = self._cal_win
        if win is None:
            return
        for child in win.winfo_children():
            child.destroy()
        year, month = self._cal_view

        try:
            selected = self._parse_date()
        except ValueError:
            selected = None

        hdr = ttk.Frame(win)
        hdr.pack(fill="x", padx=6, pady=6)
        ttk.Button(hdr, text="◀", width=3,
                   command=lambda: self._shift_month(-1)).pack(side="left")
        ttk.Label(hdr, text=f"{calendar.month_name[month]} {year}",
                  anchor="center").pack(side="left", expand=True, fill="x")
        ttk.Button(hdr, text="▶", width=3,
                   command=lambda: self._shift_month(1)).pack(side="left")

        grid = ttk.Frame(win)
        grid.pack(padx=6, pady=(0, 8))
        for col, name in enumerate(("Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")):
            ttk.Label(grid, text=name, width=3, anchor="center").grid(
                row=0, column=col, padx=1, pady=1)

        # firstweekday=6 → weeks start on Sunday (US convention).
        weeks = calendar.Calendar(firstweekday=6).monthdayscalendar(year, month)
        for r, week in enumerate(weeks, start=1):
            for c, day in enumerate(week):
                if day == 0:
                    continue
                d = dt.date(year, month, day)
                btn = tk.Button(grid, text=str(day), width=2,
                                relief=("sunken" if d == selected else "raised"),
                                command=lambda dd=d: self._pick_day(dd))
                if d == selected:
                    btn.config(font=("TkDefaultFont", 0, "bold"))
                btn.grid(row=r, column=c, padx=1, pady=1)

    def _shift_month(self, delta: int) -> None:
        year, month = self._cal_view
        month += delta
        if month < 1:
            month, year = 12, year - 1
        elif month > 12:
            month, year = 1, year + 1
        self._cal_view = (year, month)
        self._render_calendar()

    def _pick_day(self, d: dt.date) -> None:
        self._set_date(d)
        self._close_calendar()

    def _close_calendar(self) -> None:
        if self._cal_win is not None:
            try:
                self._cal_win.destroy()
            except tk.TclError:
                pass
            self._cal_win = None

    def _clear_orders(self) -> None:
        for child in self.list_frame.winfo_children():
            child.destroy()
        self.check_vars.clear()

    def _populate_orders(self, numbers: list[str]) -> None:
        self._clear_orders()
        for n in numbers:
            var = tk.BooleanVar(value=False)
            self.check_vars[n] = var
            ttk.Checkbutton(self.list_frame, text=n, variable=var).pack(
                anchor="w", padx=6, pady=1
            )

    def _set_busy(self, busy: bool) -> None:
        """Disable inputs while a worker runs; re-enable order-dependent buttons after."""
        self.busy = busy
        ready = (not busy) and bool(self.check_vars)
        self.fetch_btn.config(state="disabled" if busy else "normal")
        self.date_entry.config(state="disabled" if busy else "normal")
        self.cal_btn.config(state="disabled" if busy else "normal")
        if busy:
            self._close_calendar()
        for btn in (self.copy_btn, self.all_btn, self.none_btn):
            btn.config(state="normal" if ready else "disabled")

    # --- actions ----------------------------------------------------------
    def on_fetch(self) -> None:
        try:
            target = self._parse_date()
        except ValueError:
            messagebox.showerror("Bad date", "Please enter the date as MM-DD-YYYY.")
            return
        self._clear_orders()
        self.fetched_date = None
        self._set_busy(True)
        self.status.set(f"Opening TAP and filtering {target:%m-%d-%Y}…")
        threading.Thread(target=self._fetch_worker, args=(target,), daemon=True).start()

    def _fetch_worker(self, target: dt.date) -> None:
        try:
            with chrome_session(DOWNLOADS, headless=HEADLESS) as page:
                load_or_login(page, self.username, self.password, interactive=False)
                expected = apply_filter(page, target)
                numbers = list_order_numbers(page, expected_total=expected)
            self.q.put(("orders", target, numbers))
        except MFARequiredError:
            self.q.put(("error", REAUTH_MESSAGE))
        except Exception as e:  # noqa: BLE001 — surfaced to the owner, logged in full
            log.exception("Fetch failed")
            self.q.put(("error", f"Couldn't fetch orders:\n{e}"))
        finally:
            self.q.put(("done",))

    def on_copy(self) -> None:
        selected = [n for n, v in self.check_vars.items() if v.get()]
        if not selected:
            messagebox.showinfo("Nothing selected", "Tick at least one order to copy.")
            return
        if self.fetched_date is None:
            return
        self._set_busy(True)
        self.status.set(f"Exporting {len(selected)} order(s)…")
        threading.Thread(
            target=self._copy_worker, args=(self.fetched_date, selected), daemon=True
        ).start()

    def _copy_worker(self, target: dt.date, selected: list[str]) -> None:
        try:
            clear_downloads()
            only = set(selected)
            batches: list[OrderBatch] = []
            with chrome_session(DOWNLOADS, headless=HEADLESS) as page:
                load_or_login(page, self.username, self.password, interactive=False)
                expected = apply_filter(page, target)
                for handle in iter_orders(page, DOWNLOADS, only=only, expected_total=expected):
                    try:
                        rows = parse_ods(handle.ods_path)
                    except Exception:  # noqa: BLE001 — one bad ODS shouldn't sink the rest
                        log.exception("Failed to parse %s", handle.ods_path)
                        self.q.put(("status", f"Skipped {handle.order_number}: parse error"))
                        continue
                    batches.append(OrderBatch(handle.order_number, target, rows))
                    self.q.put(("status", f"Exported {handle.order_number} ({len(rows)} items)…"))

            if not batches:
                self.q.put(("error", "Nothing was exported — the selected orders had no items."))
                return
            _backup, rows_added, orders_added = write_orders(self.xlsx_path, batches, BACKUPS)
            self.q.put(("copied", rows_added, orders_added, len(selected)))
        except MFARequiredError:
            self.q.put(("error", REAUTH_MESSAGE))
        except Exception as e:  # noqa: BLE001 — surfaced to the owner, logged in full
            log.exception("Copy failed")
            self.q.put(("error", f"Couldn't copy orders:\n{e}"))
        finally:
            self.q.put(("done",))

    # --- queue pump (runs on the Tk main thread) --------------------------
    def _drain(self) -> None:
        try:
            while True:
                self._handle(self.q.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._drain)

    def _handle(self, msg: tuple) -> None:
        kind = msg[0]
        if kind == "status":
            self.status.set(msg[1])
        elif kind == "orders":
            _, target, numbers = msg
            self.fetched_date = target
            self._populate_orders(numbers)
            self.status.set(
                f"Found {len(numbers)} order(s) for {target:%m-%d-%Y} — tick the ones to copy."
                if numbers else f"No orders found for {target:%m-%d-%Y}."
            )
        elif kind == "copied":
            _, rows_added, orders_added, requested = msg
            skipped = requested - orders_added
            if rows_added == 0:
                self.status.set(f"Nothing new — all {requested} were already in the sheet.")
                messagebox.showinfo(
                    "Done", "Those orders were already in the sheet — nothing was added."
                )
            else:
                text = f"Added {rows_added} row(s) from {orders_added} order(s)."
                if skipped:
                    text += f"\n{skipped} were already in the sheet and skipped."
                self.status.set(text.replace("\n", "  "))
                messagebox.showinfo("Done", text)
        elif kind == "error":
            self.status.set("Error — see the dialog.")
            messagebox.showerror("Error", msg[1])
        elif kind == "done":
            self._set_busy(False)


def _fatal_dialog(title: str, message: str) -> None:
    """Show an error before the main window exists (config problems)."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(title, message)
    root.destroy()


def main() -> int:
    load_dotenv(ROOT / ".env")
    setup_logging()

    username = os.environ.get("TAP_USERNAME")
    password = os.environ.get("TAP_PASSWORD")
    xlsx_raw = os.environ.get("ORDER_XLSX_PATH")
    if not (username and password and xlsx_raw):
        _fatal_dialog(
            "Missing configuration",
            "Set TAP_USERNAME, TAP_PASSWORD and ORDER_XLSX_PATH in .env first.",
        )
        return 2
    xlsx_path = Path(xlsx_raw).expanduser()
    if not xlsx_path.exists():
        _fatal_dialog("Missing workbook", f"ORDER_XLSX_PATH does not exist:\n{xlsx_path}")
        return 2

    root = tk.Tk()
    PickerApp(root, username, password, xlsx_path)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
