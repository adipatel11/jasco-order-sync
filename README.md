# jasco-order-sync

Automates the daily routine of pulling submitted orders from the Mississippi DOR
Taxpayer Access Point (https://tap.dor.ms.gov/_/) and appending them to the **Pending**
sheet of `Order.xlsx`.

## What it does

1. Logs into TAP (cookies persisted so MFA only triggers on first run after "trust device").
2. Filters orders to `submitted = yesterday`.
3. For each order: scrapes the order number, downloads the Export ODS.
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

cp .env.example .env
# edit .env: TAP_USERNAME, TAP_PASSWORD, ORDER_XLSX_PATH
```

For dev, point `ORDER_XLSX_PATH` at `Practice Files/Order.xlsx`. For prod, point at
the owner's OneDrive-synced copy.

## Run

```bash
python run.py
```

First run will open Chrome, log in, prompt for MFA — click **Trust this device**.
The script saves `cookies.pkl`; subsequent runs reuse the session.

## Layout

| File              | Role                                                              |
| ----------------- | ----------------------------------------------------------------- |
| `run.py`          | Entrypoint, orchestrates the full flow                            |
| `tap_scraper.py`  | Selenium driver: login, filter, iterate orders, export            |
| `ods_parser.py`   | Reads rows from an Export ODS, drops totals row                   |
| `xlsx_writer.py`  | Loads `Order.xlsx`, appends to `Pending`, saves timestamped copy  |
| `backups/`        | Output xlsx files                                                 |
| `downloads/`      | Temp ODS landing zone (cleared each run)                          |
| `logs/`           | Run logs                                                          |

## Daily scheduling (optional, set up after manual run is clean)

Use macOS `launchd`. Create `~/Library/LaunchAgents/com.jasco.order-sync.plist`
that calls `python /path/to/run.py` once daily, then `launchctl load` it.
