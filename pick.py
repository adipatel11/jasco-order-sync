"""Interactive picker: choose up to 10 dates, see their TAP orders, tick which to copy.

Companion to run.py (the unattended midnight job). It shares the SAME scraper, ODS
parser, and xlsx writer — the only difference is that the owner chooses the dates and
chooses *which* orders get written to Order.xlsx, through a small Tkinter window.

Two phases, each a short self-contained browser session that reuses the persistent
`.browser_profile` (so in normal use there is no re-login or MFA):
  1. Fetch — filter each chosen date in turn to its orders that aren't canceled, and
     list each order's number and status (no downloads yet). One filter per date so
     every order's date is known for certain.
  2. Copy  — filter all the ticked orders' dates at once
     ((d1 OR d2 …) AND status<>'Canceled'),
     export only the ticked orders, parse them, and write them to Order.xlsx.

Copying REPLACES a date: every row already in the sheet dated one of the ticked
orders' days is deleted first, then the fresh export is appended. So re-copying a day
always leaves the sheet matching TAP, rather than skipping orders already present.

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
import webbrowser
from pathlib import Path
from tkinter import messagebox, ttk

from dotenv import load_dotenv

import master_workbook
from master_workbook import GraphAuthRequiredError, WorkbookLockedError
from ods_parser import parse_ods
from tap_scraper import (
    MFARequiredError,
    apply_filter,
    chrome_session,
    iter_orders,
    list_orders,
    load_or_login,
)
from xlsx_writer import OrderBatch

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

# Each date is its own filter pass on TAP during Fetch, so keep the list short.
MAX_DATES = 10
DATE_FMT = "%m-%d-%Y"

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

    def __init__(self, root: tk.Tk, username: str, password: str):
        self.root = root
        self.username = username
        self.password = password
        self.workbook_url: str | None = None  # set after a successful copy
        self.q: queue.Queue = queue.Queue()
        self.busy = False
        self.dates: list[dt.date] = []  # the chosen dates, kept sorted
        # What the last Fetch returned: which date each listed order belongs to, and
        # every order per date (to spot unticked orders on a date being replaced).
        self.order_date: dict[str, dt.date] = {}
        self.fetched: dict[dt.date, list[str]] = {}
        self.check_vars: dict[str, tk.BooleanVar] = {}
        self._cal_win: tk.Toplevel | None = None  # the open calendar popup, if any
        self._cal_view = (dt.date.today().year, dt.date.today().month)

        root.title("Jasco Order Picker")
        root.geometry("460x700")
        root.minsize(400, 560)

        self._build_widgets()
        self.root.after(100, self._drain)
        # Resolve the workbook link off the main thread: it's a network round-trip
        # and the window should appear instantly whether or not it succeeds.
        threading.Thread(target=self._resolve_url_worker, daemon=True).start()

    def _resolve_url_worker(self) -> None:
        url = master_workbook.web_url()
        if url:
            self.q.put(("workbook_url", url))

    # --- widget construction ----------------------------------------------
    def _build_widgets(self) -> None:
        yesterday = dt.date.today() - dt.timedelta(days=1)

        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(top, text="Date (MM-DD-YYYY):").pack(side="left")
        self.date_var = tk.StringVar(value=yesterday.strftime(DATE_FMT))
        self.date_entry = ttk.Entry(top, textvariable=self.date_var, width=12)
        self.date_entry.pack(side="left", padx=6)
        self.date_entry.bind("<Return>", lambda e: self._add_typed_date())
        self.add_btn = ttk.Button(top, text="Add", width=5, command=self._add_typed_date)
        self.add_btn.pack(side="left")

        # Its own row: on the top row next to the entry it got clipped to "Calenc".
        quick = ttk.Frame(self.root)
        quick.pack(fill="x", padx=10, pady=(0, 4))
        # Our own calendar popup (a plain Toplevel of buttons) — tkcalendar's
        # drop-down is unusable on macOS, so we roll a small reliable one.
        self.cal_btn = ttk.Button(quick, text="Calendar…", command=self._open_calendar)
        self.cal_btn.pack(side="left")
        self.yesterday_btn = ttk.Button(quick, text="+ Yesterday",
                                        command=lambda: self._add_date(yesterday))
        self.yesterday_btn.pack(side="left", padx=(4, 0))
        self.today_btn = ttk.Button(quick, text="+ Today",
                                    command=lambda: self._add_date(dt.date.today()))
        self.today_btn.pack(side="left", padx=4)

        self.dates_frame = ttk.LabelFrame(self.root)
        self.dates_frame.pack(fill="x", padx=10, pady=(0, 6))
        self.dates_list = tk.Listbox(self.dates_frame, height=5, selectmode="extended",
                                     exportselection=False, activestyle="none")
        dates_scroll = ttk.Scrollbar(self.dates_frame, orient="vertical",
                                     command=self.dates_list.yview)
        self.dates_list.configure(yscrollcommand=dates_scroll.set)
        self.dates_list.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        dates_scroll.pack(side="left", fill="y", pady=6)
        for key in ("<Delete>", "<BackSpace>"):
            self.dates_list.bind(key, lambda e: self._remove_selected_dates())
        date_btns = ttk.Frame(self.dates_frame)
        date_btns.pack(side="left", fill="y", padx=6, pady=6)
        self.remove_btn = ttk.Button(date_btns, text="Remove",
                                     command=self._remove_selected_dates)
        self.remove_btn.pack(fill="x")
        self.clear_dates_btn = ttk.Button(date_btns, text="Clear all",
                                          command=self._clear_dates)
        self.clear_dates_btn.pack(fill="x", pady=(4, 0))

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

        # The workbook now lives in OneDrive, so the owner can go straight from
        # copying to looking at it without hunting for the link.
        self.open_btn = ttk.Button(self.root, text="Open workbook in browser",
                                   command=self._open_workbook, state="disabled")
        self.open_btn.pack(fill="x", padx=10, pady=(4, 0))

        self.status = tk.StringVar(value="Add up to 10 dates and click Fetch.")
        ttk.Label(self.root, textvariable=self.status, relief="sunken",
                  anchor="w").pack(fill="x", side="bottom")

        # Start with yesterday listed so the everyday case stays a single click.
        self._add_date(yesterday)

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

    def _open_workbook(self) -> None:
        if self.workbook_url:
            webbrowser.open(self.workbook_url)

    def _set_all(self, value: bool) -> None:
        for var in self.check_vars.values():
            var.set(value)

    def _parse_date(self) -> dt.date:
        return dt.datetime.strptime(self.date_var.get().strip(), DATE_FMT).date()

    # --- the chosen-dates list --------------------------------------------
    def _add_date(self, d: dt.date) -> bool:
        """Add d to the list (no-op if already there). False if the list is full."""
        if self.busy:
            return False
        if d in self.dates:
            return True
        if len(self.dates) >= MAX_DATES:
            messagebox.showwarning(
                "Too many dates",
                f"You can pick up to {MAX_DATES} dates at a time. Remove one first.",
                parent=self._cal_win if self._cal_win is not None else self.root,
            )
            return False
        self.dates.append(d)
        self.dates.sort()
        self._refresh_dates()
        return True

    def _remove_date(self, d: dt.date) -> None:
        if self.busy or d not in self.dates:
            return
        self.dates.remove(d)
        self._refresh_dates()

    def _add_typed_date(self) -> None:
        try:
            d = self._parse_date()
        except ValueError:
            messagebox.showerror("Bad date", "Please enter the date as MM-DD-YYYY.")
            return
        self._add_date(d)

    def _remove_selected_dates(self) -> None:
        if self.busy:
            return
        picked = {self.dates[i] for i in self.dates_list.curselection()}
        if not picked:
            return
        self.dates = [d for d in self.dates if d not in picked]
        self._refresh_dates()

    def _clear_dates(self) -> None:
        if self.busy:
            return
        self.dates.clear()
        self._refresh_dates()

    def _refresh_dates(self) -> None:
        """Redraw the list box (and the open calendar's highlights) from self.dates."""
        self.dates_list.delete(0, "end")
        for d in self.dates:
            self.dates_list.insert("end", f"{d:%a}  {d:%m-%d-%Y}")
        self.dates_frame.config(text=f"Dates to fetch ({len(self.dates)}/{MAX_DATES})")
        if self._cal_win is not None:
            self._render_calendar()

    # --- calendar popup ---------------------------------------------------
    # A self-contained month grid in a plain Toplevel. We deliberately avoid any
    # focus grab (tkcalendar's grab is what wedges on macOS): the popup is just a
    # window of buttons. Clicking a day adds it to the list (or removes it if it is
    # already there); the popup stays open so several days can be picked in a row.
    def _open_calendar(self) -> None:
        if self.busy:
            return
        if self._cal_win is not None and self._cal_win.winfo_exists():
            self._cal_win.lift()
            return
        base = self.dates[-1] if self.dates else dt.date.today()
        self._cal_view = (base.year, base.month)

        win = tk.Toplevel(self.root)
        win.title("Pick dates")
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
        chosen = set(self.dates)

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
                picked = d in chosen
                btn = tk.Button(grid, text=str(day), width=2,
                                relief=("sunken" if picked else "raised"),
                                command=lambda dd=d: self._toggle_day(dd))
                if picked:
                    btn.config(font=("TkDefaultFont", 0, "bold"))
                btn.grid(row=r, column=c, padx=1, pady=1)

        foot = ttk.Frame(win)
        foot.pack(fill="x", padx=6, pady=(0, 8))
        ttk.Label(foot, text=f"{len(self.dates)}/{MAX_DATES} picked").pack(side="left")
        ttk.Button(foot, text="Done", command=self._close_calendar).pack(side="right")

    def _shift_month(self, delta: int) -> None:
        year, month = self._cal_view
        month += delta
        if month < 1:
            month, year = 12, year - 1
        elif month > 12:
            month, year = 1, year + 1
        self._cal_view = (year, month)
        self._render_calendar()

    def _toggle_day(self, d: dt.date) -> None:
        if d in self.dates:
            self._remove_date(d)
        else:
            self._add_date(d)

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
        self.order_date.clear()
        self.fetched.clear()

    def _populate_orders(self, results: list[tuple[dt.date, dict[str, str]]]) -> None:
        """One bold heading per date, then a checkbox per order (number — status)."""
        self._clear_orders()
        for d, orders in results:
            numbers = list(orders)
            self.fetched[d] = numbers
            count = f"{len(numbers)} order(s)" if numbers else "no orders"
            ttk.Label(self.list_frame, text=f"{d:%a %m-%d-%Y} — {count}",
                      font=("TkDefaultFont", 0, "bold")).pack(
                anchor="w", padx=6, pady=(6, 1))
            for n in numbers:
                var = tk.BooleanVar(value=False)
                self.check_vars[n] = var
                self.order_date[n] = d
                label = f"{n}  —  {orders[n]}" if orders[n] else n
                ttk.Checkbutton(self.list_frame, text=label, variable=var).pack(
                    anchor="w", padx=18, pady=1
                )

    def _set_busy(self, busy: bool) -> None:
        """Disable inputs while a worker runs; re-enable order-dependent buttons after."""
        self.busy = busy
        ready = (not busy) and bool(self.check_vars)
        state = "disabled" if busy else "normal"
        for widget in (self.fetch_btn, self.date_entry, self.add_btn, self.cal_btn,
                       self.yesterday_btn, self.today_btn, self.remove_btn,
                       self.clear_dates_btn, self.dates_list):
            widget.config(state=state)
        if busy:
            self._close_calendar()
        for btn in (self.copy_btn, self.all_btn, self.none_btn):
            btn.config(state="normal" if ready else "disabled")

    # --- actions ----------------------------------------------------------
    def on_fetch(self) -> None:
        if not self.dates:
            messagebox.showinfo("No dates", "Add at least one date to fetch.")
            return
        dates = list(self.dates)
        self._clear_orders()
        self._set_busy(True)
        self.status.set(f"Opening TAP to fetch {len(dates)} date(s)…")
        threading.Thread(target=self._fetch_worker, args=(dates,), daemon=True).start()

    def _fetch_worker(self, dates: list[dt.date]) -> None:
        try:
            results: list[tuple[dt.date, dict[str, str]]] = []
            with chrome_session(DOWNLOADS, headless=HEADLESS) as page:
                load_or_login(page, self.username, self.password, interactive=False)
                # One filter per date rather than a single OR-ed query: the grid rows
                # don't reliably say which day an order is from, and the copy step has
                # to know that to write column F and to replace the right days.
                for i, d in enumerate(dates, start=1):
                    self.q.put(("status", f"Filtering {d:%m-%d-%Y} ({i}/{len(dates)})…"))
                    expected = apply_filter(page, d)
                    results.append((d, list_orders(page, expected_total=expected)))
            self.q.put(("orders", results))
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
        order_date = {n: self.order_date[n] for n in selected}
        replace = sorted(set(order_date.values()))

        # Replacing a date wipes ALL of its rows, so an order left unticked on a date
        # that is being copied will not be in the sheet afterwards. That's the one
        # surprising outcome, so it's the one we stop and ask about.
        partial = [d for d in replace
                   if any(not self.check_vars[n].get() for n in self.fetched.get(d, []))]
        if partial:
            days = "\n".join(f"   • {d:%a %m-%d-%Y}" for d in partial)
            if not messagebox.askokcancel(
                "Some orders are unticked",
                "Copying replaces every row already in the sheet for these dates:\n\n"
                f"{days}\n\n"
                "Those dates have orders you didn't tick, and they won't be in the "
                "sheet afterwards. Continue?",
                icon="warning",
            ):
                return

        self._set_busy(True)
        self.status.set(f"Exporting {len(selected)} order(s)…")
        threading.Thread(
            target=self._copy_worker, args=(order_date, replace), daemon=True
        ).start()

    def _copy_worker(self, order_date: dict[str, dt.date], replace: list[dt.date]) -> None:
        try:
            clear_downloads()
            only = set(order_date)
            batches: list[OrderBatch] = []
            failed: list[str] = []
            with chrome_session(DOWNLOADS, headless=HEADLESS) as page:
                load_or_login(page, self.username, self.password, interactive=False)
                # All the dates in one query: (d1 OR d2 …) AND status<>'Canceled'.
                expected = apply_filter(page, replace)
                for handle in iter_orders(page, DOWNLOADS, only=only, expected_total=expected):
                    day = order_date.get(handle.order_number)
                    if day is None:
                        # The detail page labelled it differently from the list; with
                        # one date there is no doubt which day it is, otherwise there is.
                        if len(replace) != 1:
                            log.error("Exported %s, which wasn't in the list", handle.order_number)
                            failed.append(handle.order_number)
                            continue
                        day = replace[0]
                    try:
                        rows = parse_ods(handle.ods_path)
                    except Exception:  # noqa: BLE001 — reported below, logged in full
                        log.exception("Failed to parse %s", handle.ods_path)
                        failed.append(handle.order_number)
                        continue
                    batches.append(OrderBatch(handle.order_number, day, rows))
                    self.q.put(("status", f"Exported {handle.order_number} ({len(rows)} items)…"))

            # Replacing a date deletes its existing rows, so writing an incomplete set
            # would quietly lose the orders that didn't come through. All or nothing.
            missing = sorted(only - {b.order_number for b in batches})
            if missing or failed:
                bad = sorted(set(missing) | set(failed))
                self.q.put(("error",
                            "Couldn't export every ticked order, so the workbook was left "
                            "unchanged:\n\n" + "\n".join(bad) + "\n\nTry Copy again."))
                return
            self.q.put(("status", "Updating the workbook in OneDrive…"))
            result = master_workbook.append(batches, BACKUPS, replace_dates=set(replace))
            self.q.put(("copied", result, len(only), replace))
        except MFARequiredError:
            self.q.put(("error", REAUTH_MESSAGE))
        except GraphAuthRequiredError:
            self.q.put(("error", master_workbook.onedrive.REAUTH_MESSAGE))
        except WorkbookLockedError as e:
            # The likeliest failure by far in day-to-day use: the owner is looking at
            # the workbook while copying into it. Say exactly what to do about it.
            self.q.put(("error", str(e)))
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
            results = msg[1]
            self._populate_orders(results)
            total = sum(len(numbers) for _, numbers in results)
            self.status.set(
                f"Found {total} order(s) across {len(results)} date(s) — tick the ones to copy."
                if total else f"No orders found for those {len(results)} date(s)."
            )
        elif kind == "workbook_url":
            self.workbook_url = msg[1]
            self.open_btn.config(state="normal")
        elif kind == "copied":
            _, result, requested, replaced = msg
            if result.web_url:
                self.workbook_url = result.web_url
                self.open_btn.config(state="normal")
            skipped = requested - result.orders_added
            if result.rows_added == 0 and result.rows_removed == 0:
                self.status.set("Nothing changed in the sheet.")
                messagebox.showinfo(
                    "Done", "Those orders were already in the sheet — nothing was changed."
                )
            else:
                days = ", ".join(f"{d:%m-%d}" for d in replaced)
                text = f"Added {result.rows_added} row(s) from {result.orders_added} order(s)"
                if result.rows_removed:
                    text += (f", replacing {result.rows_removed} old row(s) dated "
                             f"{days}.")
                else:
                    text += "."
                if skipped:
                    text += (f"\n{skipped} order(s) were already in the sheet under "
                             "another date, or had no items, and were skipped.")
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
    if not (username and password):
        _fatal_dialog(
            "Missing configuration", "Set TAP_USERNAME and TAP_PASSWORD in .env first."
        )
        return 2
    # Fail before the window opens if the workbook is unreachable — better than
    # letting the owner pick orders and only then discovering nothing can be saved.
    problem = master_workbook.check_config()
    if problem:
        _fatal_dialog("Can't reach the workbook", problem)
        return 2

    root = tk.Tk()
    PickerApp(root, username, password)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
