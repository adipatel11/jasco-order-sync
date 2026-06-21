# jasco-order-sync

Automates the daily routine of pulling submitted orders from the Mississippi DOR
Taxpayer Access Point (https://tap.dor.ms.gov/_/) and appending them to the **Pending**
sheet of `Order.xlsx`.

> Built for a small wholesale business to replace a ~1-hour daily manual data-entry
> task with a one-command (or one-click) sync. Stack: Python, Playwright (headless
> Chromium), openpyxl, Tkinter, and a launchd schedule. Credentials and business data
> live only in a git-ignored `.env` / local files — nothing sensitive is committed.

## What it does

1. Logs into TAP (a persistent browser profile keeps the "Trust this device" cookie, so the MFA text-message step only happens on the first run).
2. Opens the account's **Add/View Retail Orders** list and filters to `submitted = yesterday`.
3. For each order: clicks View, clicks Export (a popup triggers the ODS download), and recovers the order number from the file/page.
4. Appends rows to the `Pending` sheet of `Order.xlsx`:
   - `A` = Item # (from ODS col A)
   - `B` = Name (from ODS col B)
   - `C` = `=VLOOKUP(A{row},SizeData!A$1:B$3974,2,FALSE)`
   - `D` = Reserved Quantity (from ODS col H)
   - `E` = Order # (scraped from order page)
   - `F` = Date (yesterday, as Excel serial)
5. Writes a timestamped copy to `backups/Order_YYYY-MM-DD_HHMMSS.xlsx`. The source
   xlsx is never overwritten.
6. Skips orders whose order # is already present in column E (idempotent).

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium

cp .env.example .env
# edit .env: TAP_USERNAME, TAP_PASSWORD, ORDER_XLSX_PATH
```

For dev, point `ORDER_XLSX_PATH` at `Practice Files/Order.xlsx`. For prod, point at
the owner's OneDrive-synced copy.

## Run

```bash
python run.py
```

First run will open Chromium, log in, pause for MFA — click **Trust this device**
in the browser, then press Enter in the terminal. The full browser profile is
stored under `.browser_profile/`, so subsequent runs reuse the trust-device cookie
and skip MFA entirely.

## Picking specific orders on demand (`pick.py`)

The midnight job grabs *every* order from *yesterday*. When the owner instead wants to
choose a **specific date** and hand-pick **which** orders to copy, run the picker:

```bash
python pick.py
```

Or, so the owner never has to touch a terminal, **double-click the launcher** — it
opens the same window:
- **Mac:** `Order Picker.command` in Finder. (First time only, macOS may ask to
  confirm opening it: right-click → **Open** → **Open**.)
- **Windows:** `Order Picker.bat` in File Explorer. (First time only, SmartScreen
  may warn — click **More info → Run anyway**.)

A window opens. Enter a date (or click **Yesterday** / **Today**), click **Fetch
orders** — it lists that day's order numbers without downloading anything — tick the
ones you want, then click **Copy selected → Excel**. Only the ticked orders are
exported and appended to the same `Pending` sheet of `Order.xlsx`.

It reuses the exact same scraper, parser, and writer as `run.py`, so the output format,
backups, and `=VLOOKUP(...)` column are identical. Because the writer skips order
numbers already in column E, picking an order the midnight job already captured is
safe — it's simply skipped. The two tools never conflict.

**Notes**
- The picker shares the daily job's `.browser_profile`, so it needs no separate login.
  If TAP asks for a fresh login (cookie expired), the window says so — do one
  `python run.py` in a terminal to re-establish trust, then reopen the picker.
- It drives the browser **headless** (no window pops up). Set `HEADLESS = False` at the
  top of `pick.py` to watch it work while debugging.
- Tkinter ships with the python.org installer. If you used Homebrew Python and get
  `ModuleNotFoundError: No module named 'tkinter'`, install it with `brew install python-tk`.

## Layout

| File              | Role                                                              |
| ----------------- | ----------------------------------------------------------------- |
| `run.py`          | Daily entrypoint: yesterday → every order → append (unattended)   |
| `pick.py`         | Interactive Tkinter picker: choose a date, pick which orders       |
| `tap_scraper.py`  | Playwright driver: login, filter, list/iterate orders, export     |
| `ods_parser.py`   | Reads rows from an Export ODS, drops totals row                   |
| `xlsx_writer.py`  | Loads `Order.xlsx`, appends to `Pending`, saves timestamped copy  |
| `backups/`        | Output xlsx files                                                 |
| `downloads/`      | Temp ODS landing zone (cleared each run)                          |
| `logs/`           | Run logs                                                          |

## Deploying on the owner's Mac

1. **Install Python** (3.11+): from [python.org](https://www.python.org/downloads/) or `brew install python`.
2. **Clone and set up:**
   ```bash
   git clone https://github.com/adipatel11/jasco-order-sync.git
   cd jasco-order-sync
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   playwright install chromium
   ```
3. **Configure `.env`:**
   ```bash
   cp .env.example .env
   ```
   Edit it with the user's TAP credentials, `TAP_ACCOUNT_NAME` (the business
   account name shown after login), and `ORDER_XLSX_PATH` pointing at the OneDrive-synced
   `Order.xlsx`. Leave `HEADLESS=false`.
4. **First supervised run** (clears MFA once, establishes the trusted device):
   ```bash
   python run.py
   ```
   Complete MFA in the browser window, ticking **Trust this device**. Confirm it
   appends to his file and writes a backup to `backups/`.

## Daily scheduling (launchd, after the supervised run is clean)

A LaunchAgent template lives at `launchd/com.jasco.order-sync.plist`. It runs the
script headless at **midnight**; if the Mac is asleep, launchd runs the missed job
on the next wake (fine, since it pulls *yesterday's* orders).

```bash
# from inside the repo:
pwd     # copy this absolute path
# edit launchd/com.jasco.order-sync.plist: replace every
#   /Users/OWNER/PATH/TO/jasco-order-sync  with that path
cp launchd/com.jasco.order-sync.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.jasco.order-sync.plist

# test it immediately without waiting for midnight:
launchctl start com.jasco.order-sync
cat logs/launchd.err.log     # check for errors / MFARequiredError
```

To change later: `launchctl unload` the agent, edit, `launchctl load` again.

**Notes**
- The owner must be logged in for the agent to run (it drives a browser).
- If a scheduled run ever logs `MFARequiredError`, the trusted-device cookie
  expired — do one manual `python run.py` to re-establish it.
- If the Mac is off/closed for a *full* calendar day, that day's run is skipped
  (the script only ever fetches the single prior day).

## Running the picker on the owner's Windows PC

The daily unattended job (`run.py` + launchd) stays on the **Mac**. The Windows PC
on-site only runs the on-demand **picker** (`pick.py`), launched by double-clicking
`Order Picker.bat`. Set it up once:

1. **Install Python** (3.11+) from [python.org](https://www.python.org/downloads/).
   On the installer's first screen tick **"Add python.exe to PATH"**, and keep the
   default **"tcl/tk and IDLE"** component checked (that's Tkinter — the picker needs it).
2. **Clone and set up** (in PowerShell or Command Prompt, from where you want the repo):
   ```bat
   git clone https://github.com/adipatel11/jasco-order-sync.git
   cd jasco-order-sync
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   playwright install chromium
   ```
3. **Configure `.env`:** copy `.env.example` to `.env` and edit it with the owner's
   TAP credentials, `TAP_ACCOUNT_NAME` (the business account name shown after login), and
   `ORDER_XLSX_PATH` pointing at his OneDrive-synced `Order.xlsx` — use the real
   Windows path, e.g. `C:\Users\Owner\OneDrive\...\Order.xlsx`.
   ```bat
   copy .env.example .env
   ```
4. **Establish the trusted device once** (so the picker, which runs headless, never
   hits MFA). The picker can't do MFA itself, so run the daily sync once from a
   terminal to clear it:
   ```bat
   python run.py
   ```
   Complete the text-message code in the browser window, ticking **Trust this device**.
   The cookie is saved under `.browser_profile\` and reused from then on.
5. **Everyday use:** double-click **`Order Picker.bat`** — no terminal needed. Pick a
   date, **Fetch orders**, tick the ones to copy, **Copy selected → Excel**.

   To put it on the Desktop, double-click **`Create Desktop Shortcut.bat`** once — it
   drops a **Jasco Order Picker** icon on the Desktop that launches the picker. (Or by
   hand: right-click `Order Picker.bat` → **Send to → Desktop (create shortcut)**.)

**Notes**
- `Order Picker.bat` uses the repo's `.venv` automatically, so always keep the `.bat`
  inside the repo folder (the Desktop *shortcut* is fine — it still points back here).
- If the picker ever says TAP needs a fresh login, the trusted-device cookie expired —
  redo step 4 (`python run.py`) once, then reopen the picker.
- The Mac and Windows machines each keep their **own** `.browser_profile\` and `.env`
  (both are git-ignored), so trusting one device never affects the other.
