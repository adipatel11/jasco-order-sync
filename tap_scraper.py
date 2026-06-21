"""Drive the TAP website with Playwright: log in, filter, iterate orders, export ODS.

Selectors here come from a real `playwright codegen` recording of the live site, so
they use role/name locators (resilient to markup churn) rather than guessed CSS.

A *persistent* browser profile is used (BROWSER_PROFILE dir). When you tick
"Trust this device" during the first MFA, Chromium stores that cookie in the profile
and reuses it on every later run — so MFA's text-message step only happens once.
"""

from __future__ import annotations

import logging
import os
import re
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

from playwright.sync_api import Download, Page, TimeoutError as PWTimeout, sync_playwright

log = logging.getLogger(__name__)

TAP_URL = "https://tap.dor.ms.gov/_/"
BROWSER_PROFILE = Path(".browser_profile")

# The business/account link shown after login (e.g. "ACME RETAIL LLC").
# Set this per deployment via the TAP_ACCOUNT_NAME env var (see .env.example).
ACCOUNT_NAME = os.environ.get("TAP_ACCOUNT_NAME", "ACME RETAIL LLC")

# The detail page shows "<store> Retail Order - Order ID <value>". We read the value
# that follows this label; the regexes below are loose fallbacks if that text moves.
ORDER_ID_LABEL = "Order ID"
ORDER_ID_AFTER_LABEL_RE = re.compile(r"Order ID\s*#?[:\s\-]*([A-Za-z0-9\-]+)")
ORDER_NUMBER_RE = re.compile(r"[A-Z]\d{5,}")


@dataclass
class OrderHandle:
    order_number: str
    ods_path: Path


@contextmanager
def chrome_session(downloads_dir: Path, headless: bool = False):
    """Open a persistent-profile Chromium context. Yields a Page.

    The persistent profile is what makes "Trust this device" stick across runs.
    """
    downloads_dir.mkdir(parents=True, exist_ok=True)
    BROWSER_PROFILE.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_PROFILE.resolve()),
            headless=headless,
            accept_downloads=True,
            viewport={"width": 1400, "height": 1000},
        )
        page = context.pages[0] if context.pages else context.new_page()
        try:
            yield page
        finally:
            context.close()


def _on_login_page(page: Page) -> bool:
    """True if the login form renders (i.e. we are not authenticated).

    TAP is a JS app, so the username field appears a beat after navigation. We poll
    for it; if it never shows within the timeout, we're already logged in.
    """
    try:
        page.get_by_role("textbox", name="Username").wait_for(
            state="visible", timeout=15000
        )
        return True
    except PWTimeout:
        return False


class MFARequiredError(RuntimeError):
    """Raised when MFA is needed but no terminal is attached to complete it."""


def login(page: Page, username: str, password: str, interactive: bool = True) -> None:
    """Username/password + first-time MFA. Pauses for the SMS code if prompted.

    When `interactive` is False (e.g. a scheduled launchd run with no terminal), an
    MFA prompt raises MFARequiredError instead of hanging on input(), so the run
    fails fast and the log tells you to re-authenticate manually.
    """
    page.get_by_role("textbox", name="Username").fill(username)
    pw = page.get_by_role("textbox", name="Password")
    pw.fill(password)
    pw.press("Enter")

    # If the device is already trusted, TAP skips straight past MFA.
    try:
        page.get_by_role("textbox", name="Security Code").wait_for(timeout=6000)
    except PWTimeout:
        log.info("No MFA prompt — device already trusted")
        return

    if not interactive:
        raise MFARequiredError(
            "MFA is required but this run is unattended (no terminal). The trusted-"
            "device cookie has likely expired. Run `python run.py` manually once to "
            "re-establish trust, then scheduled runs will work again."
        )

    print("\n=== MFA required ===")
    print("In the browser window:")
    print("  1. Click 'Get a text message…' if needed and enter the Security Code.")
    print("  2. Tick 'Trust this device'.")
    print("  3. Click 'Log In'.")
    input("Then press Enter here once you're logged in... ")


def _filter_box(page: Page):
    return page.get_by_role("textbox", name="Filter Retail Orders")


def _orders_list_ready(page: Page, timeout: int = 3000) -> bool:
    """True if the Retail Orders filter box is visible (we're on the list page)."""
    try:
        _filter_box(page).wait_for(state="visible", timeout=timeout)
        return True
    except PWTimeout:
        return False


def _go_to_orders(page: Page) -> None:
    """Ensure we land on the Retail Orders list.

    The post-login landing varies: sometimes the session restores us directly onto
    the orders list, sometimes onto a dashboard where we must click the account then
    "Add/View Retail Orders". So we check for the filter box at each step instead of
    blindly clicking (which could navigate *away* from an already-loaded list).
    """
    for attempt in (1, 2, 3):
        if _orders_list_ready(page):
            return

        acct = page.get_by_role("link", name=ACCOUNT_NAME)
        if acct.count() > 0:
            acct.first.click()
            page.wait_for_load_state("networkidle")
            log.info("Clicked account; now at %s", page.url)

        orders = page.get_by_role("link", name="Add/View Retail Orders")
        try:
            orders.first.wait_for(state="visible", timeout=10000)
            orders.first.click()
            page.wait_for_load_state("networkidle")
            log.info("Clicked Add/View Retail Orders; now at %s", page.url)
        except PWTimeout:
            log.warning("Add/View Retail Orders link not found (attempt %d, url=%s)",
                        attempt, page.url)

        if _orders_list_ready(page, timeout=15000):
            return

        log.warning("Orders list not ready (attempt %d, url=%s); resetting", attempt, page.url)
        page.goto(TAP_URL, wait_until="domcontentloaded")

    raise RuntimeError(f"Could not reach the Retail Orders list (url={page.url})")


def load_or_login(page: Page, username: str, password: str, interactive: bool = True) -> None:
    """Ensure we end up on the Retail Orders list, logging in only if needed."""
    page.goto(TAP_URL, wait_until="domcontentloaded")
    if _on_login_page(page):
        log.info("No active session — performing login")
        login(page, username, password, interactive=interactive)
    else:
        log.info("Reusing saved browser profile session")
    _go_to_orders(page)


# Order View links live inside the grid cells (ids like "Dc-u-7"). Other "View" links
# on the page (toolbar/sidebar) are NOT inside those cells, so scoping here excludes
# them and leaves exactly one View link per order row.
def _order_views(page: Page):
    return page.locator("[id^='Dc-u-']").get_by_role("link", name="View")


def _order_link_count(page: Page) -> int:
    return _order_views(page).count()


def _wait_orders_stable(page: Page, *, differ_from: int | None = None,
                        timeout_ms: int = 20000) -> int:
    """Poll the 'View'-link count until it stops changing.

    The grid re-renders asynchronously after a filter/back-nav, so a single read can
    catch the stale list. If `differ_from` is given, first wait for the count to move
    off that value (the old list lingers for a moment after a filter is applied).
    """
    deadline = time.time() + timeout_ms / 1000
    last: int | None = None
    moved = differ_from is None
    while time.time() < deadline:
        n = _order_link_count(page)
        if differ_from is not None and n != differ_from:
            moved = True
        if moved and n == last:
            return n
        last = n
        page.wait_for_timeout(500)
    return last if last is not None else 0


def _wait_for_orders(page: Page, minimum: int, timeout_ms: int = 15000) -> bool:
    """Wait until at least `minimum` order rows are rendered on the current page.

    The grid re-renders in stages after a filter or a return-from-order, so a read taken
    too early can see only some of the rows. Waiting for the known count keeps us from
    mistaking a half-rendered list for a short one (which would silently drop orders).
    """
    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        if len(_current_page_numbers(page)) >= minimum:
            return True
        page.wait_for_timeout(300)
    log.warning("Only %d/%d order rows rendered after %dms",
                len(_current_page_numbers(page)), minimum, timeout_ms)
    return False


def apply_filter(page: Page, target_date) -> int:
    """Filter Retail Orders to a day's submitted reserve-inventory orders.

    Returns the stabilised order count on page 1 — the authoritative "how many orders
    are there" number that callers pass back in so iteration/listing can wait for the
    grid to fully render before trusting it.
    """
    date_str = target_date.strftime("%m-%d-%Y")
    filter_text = f'submitted="{date_str}" AND status="reserve inventory"'
    before = _order_link_count(page)
    box = _filter_box(page)
    box.click()
    box.fill(filter_text)
    box.press("Enter")
    page.wait_for_load_state("networkidle")
    n = _wait_orders_stable(page, differ_from=before)
    log.info("Filter narrowed list from %d to %d orders (page 1)", before, n)
    return n


# --- order list rows & pagination --------------------------------------------
# Each order is a <tr class="TDR ...">. The order number sits in one cell and the
# "View" link in another cell of the SAME row. The list reshuffles its row order
# every time we enter and leave an order, so we always select by order number,
# never by position.

def _current_page_numbers(page: Page) -> list[str]:
    """Order numbers (e.g. R3144144) visible on the current list page, in DOM order."""
    rows = page.locator("tr.TDR")
    numbers: list[str] = []
    for i in range(rows.count()):
        m = ORDER_NUMBER_RE.search(rows.nth(i).inner_text())
        if m:
            numbers.append(m.group(0))
    return numbers


def _click_order_view(page: Page, order_number: str) -> bool:
    """Click the View link in the row whose order-number cell matches exactly."""
    row = page.locator("tr.TDR").filter(
        has=page.get_by_role("cell", name=order_number, exact=True)
    )
    link = row.get_by_role("link", name="View").first
    if link.count() == 0:
        return False
    link.click()
    return True


def _goto_next_page(page: Page) -> bool:
    """Advance to the next list page. Returns False if there is no next page."""
    nxt = page.get_by_role("link", name="Next")
    if nxt.count() == 0:
        return False
    before = _current_page_numbers(page)
    try:
        nxt.first.click(timeout=5000)
    except PWTimeout:
        return False
    page.wait_for_load_state("networkidle")
    _wait_orders_stable(page)
    return _current_page_numbers(page) != before


def _goto_first_page(page: Page) -> None:
    """Walk back to the first list page via the Prev link (no-op if single page)."""
    for _ in range(100):  # safety bound
        prev = page.get_by_role("link", name="Prev")
        if prev.count() == 0:
            return
        before = _current_page_numbers(page)
        try:
            prev.first.click(timeout=5000)
        except PWTimeout:
            return
        page.wait_for_load_state("networkidle")
        _wait_orders_stable(page)
        if _current_page_numbers(page) == before:
            return


def list_order_numbers(page: Page, expected_total: int | None = None) -> list[str]:
    """Enumerate order numbers across all list pages WITHOUT entering/exporting them.

    The interactive picker uses this to show the day's orders cheaply so the owner can
    choose which to export — exporting every order just to pick a couple would be slow.
    Reuses the same row scraping and pagination as iter_orders; order is preserved and
    duplicates (a number that re-appears after the list reshuffles) are collapsed.

    `expected_total` (the post-filter count from apply_filter) makes us wait for the grid
    to finish rendering before scanning, so we never show a short list by mistake.
    """
    _goto_first_page(page)
    if expected_total:
        _wait_for_orders(page, expected_total)
    seen: set[str] = set()
    ordered: list[str] = []
    while True:
        for n in _current_page_numbers(page):
            if n not in seen:
                seen.add(n)
                ordered.append(n)
        if not _goto_next_page(page):
            break
    return ordered


def _extract_order_number(download: Download, page: Page) -> str:
    """Order # from the 'Order ID' label on the detail page, with loose fallbacks."""
    # 1. Preferred: the text that follows the "Order ID" label on the detail page.
    try:
        label = page.get_by_text(ORDER_ID_LABEL).first
        container_text = label.locator("xpath=..").inner_text(timeout=2000)
        log.debug("Order ID container text: %r", container_text)
        m = ORDER_ID_AFTER_LABEL_RE.search(container_text)
        if m:
            return m.group(1).strip()
    except PWTimeout:
        pass
    # 2. Fallbacks: the export filename, then the URL.
    for source in (download.suggested_filename, page.url):
        m = ORDER_NUMBER_RE.search(source or "")
        if m:
            return m.group(0)
    # Last resort: the raw filename stem so nothing is silently lost.
    return Path(download.suggested_filename).stem


def _export_current_order(page: Page, downloads_dir: Path) -> OrderHandle:
    """On an order detail page: click Export, save the ODS, read the order number.

    The export is finicky: it may open a transient popup or download directly, the
    download may attach to the popup rather than the main page, and the first click
    can be a no-op if the detail page isn't fully interactive yet. So we capture
    downloads from ANY page (main or popup) and retry the Export click.
    """
    context = page.context
    captured: list[Download] = []
    handlers: list = []

    def _attach(p: Page) -> None:
        handler = lambda d: captured.append(d)  # noqa: E731
        p.on("download", handler)
        handlers.append((p, handler))

    _attach(page)
    context.on("page", _attach)  # catch downloads that fire on a popup window
    try:
        export = page.get_by_role("link", name="Export")
        export.first.wait_for(state="visible", timeout=15000)
        for attempt in range(1, 4):
            export.first.click()
            deadline = time.time() + 12
            while not captured and time.time() < deadline:
                page.wait_for_timeout(200)
            if captured:
                break
            log.warning("Export click %d produced no download; retrying", attempt)
            page.wait_for_timeout(1000)
        if not captured:
            raise RuntimeError(f"Export produced no download (url={page.url})")
        download = captured[0]
    finally:
        context.remove_listener("page", _attach)
        for p, handler in handlers:
            try:
                p.remove_listener("download", handler)
            except Exception:
                pass

    for other in list(context.pages):
        if other is not page:
            try:
                other.close()
            except Exception:
                pass

    order_number = _extract_order_number(download, page)
    out_path = downloads_dir / f"{order_number}.ods"
    download.save_as(out_path)
    log.info(
        "Order %s exported (file=%r url=%s)",
        order_number, download.suggested_filename, page.url,
    )
    return OrderHandle(order_number=order_number, ods_path=out_path)


def _on_order_detail(page: Page, timeout: int = 10000) -> bool:
    """True once an order's detail page is up (its 'Go back to Request' link shows)."""
    try:
        page.get_by_role("link", name="Go back to Request").wait_for(
            state="visible", timeout=timeout
        )
        return True
    except PWTimeout:
        return False


def _return_to_list(page: Page) -> None:
    """Best-effort: get back to the Retail Orders list from wherever we are.

    Fast no-op when the list is already showing. If we're on an order detail page,
    click its 'Go back to Request' link. Always settles the grid before returning, so
    the next row read/click sees a stable list.
    """
    if _orders_list_ready(page, timeout=2000):
        return
    back = page.get_by_role("link", name="Go back to Request")
    if back.count() > 0:
        try:
            back.first.click(timeout=5000)
            page.wait_for_load_state("networkidle")
        except PWTimeout:
            pass
    _wait_orders_stable(page)


def _open_order_detail(page: Page, order_number: str, attempts: int = 3) -> bool:
    """Click an order's View link and confirm its detail page loaded, retrying on miss.

    The grid reshuffles every time we leave an order, and the live site occasionally
    swallows a View click while the grid is still re-rendering — leaving us on the list
    with no detail page. Rather than let a single timeout kill the whole run, we
    re-settle the list and retry; the caller skips the order if we truly can't get in.
    """
    for attempt in range(1, attempts + 1):
        _return_to_list(page)  # start each try from a settled list
        if not _click_order_view(page, order_number):
            log.warning("Order %s: no View link on the list (attempt %d, url=%s)",
                        order_number, attempt, page.url)
            continue
        page.wait_for_load_state("networkidle")
        if _on_order_detail(page):
            return True
        log.warning("Order %s: detail page didn't load (attempt %d, url=%s); retrying",
                    order_number, attempt, page.url)
    return False


def iter_orders(page: Page, downloads_dir: Path, only: set[str] | None = None,
                expected_total: int | None = None):
    """Yield an OrderHandle per order on the filtered list.

    Two quirks of the live site shape this:
      * The grid reshuffles whenever we enter/leave an order, so we never select by
        position — only by order number — re-locating each order from a fresh scan.
      * The grid sometimes re-renders only *some* rows after we return from an order. A
        naive "no rows left to do → finished" check then stops early and silently drops
        orders. So we first snapshot the full set of order numbers to process (waiting
        for the grid to reach `expected_total` rows), then loop until every one has been
        handled — re-settling the list before each scan and giving up only after several
        fruitless sweeps.

    `only` restricts processing to those order numbers (the picker passes the owner's
    selection). `expected_total` is the filtered count from apply_filter; it lets us wait
    for a fully-rendered grid before trusting what's on it.
    """
    # Snapshot the authoritative set of orders to process.
    if only is not None:
        targets = set(only)
        _goto_first_page(page)
        if expected_total:
            _wait_for_orders(page, expected_total)
    else:
        targets = set(list_order_numbers(page, expected_total=expected_total))
    log.info("iter_orders: %d order(s) to process", len(targets))

    processed: set[str] = set()
    stalls = 0
    while targets - processed:
        _goto_first_page(page)
        if expected_total:
            _wait_for_orders(page, expected_total)

        # Find an unprocessed target anywhere across the list pages.
        number: str | None = None
        while True:
            here = [n for n in _current_page_numbers(page)
                    if n in targets and n not in processed]
            if here:
                number = here[0]
                break
            if not _goto_next_page(page):
                break

        if number is None:
            # A full sweep found none of the remaining targets — likely a transient
            # under-render. Retry a few times before accepting they're unreachable.
            stalls += 1
            if stalls >= 3:
                log.error("Gave up locating orders on the list: %s",
                          sorted(targets - processed))
                break
            log.warning("Sweep found no remaining target (stall %d/3); retrying", stalls)
            continue
        stalls = 0

        # Enter the order's detail page, retrying the click if the flaky grid swallows
        # it. If we still can't get in, skip this order rather than killing the run.
        if not _open_order_detail(page, number):
            log.error("Could not open order %s after retries; skipping it", number)
            processed.add(number)
            _return_to_list(page)
            continue

        handle = _export_current_order(page, downloads_dir)
        processed.add(number)
        processed.add(handle.order_number)  # in case list/detail labels differ
        yield handle
        _return_to_list(page)

    log.info("Processed %d of %d targeted order(s)", len(processed & targets), len(targets))
