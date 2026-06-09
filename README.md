# jasco-order-sync

Automates the daily routine of pulling submitted orders from the Mississippi DOR
Taxpayer Access Point (https://tap.dor.ms.gov/_/) and appending them to the **Pending**
sheet of `Order.xlsx`.

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

## Layout

| File              | Role                                                              |
| ----------------- | ----------------------------------------------------------------- |
| `run.py`          | Entrypoint, orchestrates the full flow                            |
| `tap_scraper.py`  | Playwright driver: login, filter, iterate orders, export          |
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
   Edit it with the owner's TAP credentials, `TAP_ACCOUNT_NAME` (if not
   `ACME RETAIL LLC`), and `ORDER_XLSX_PATH` pointing at his OneDrive-synced
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
