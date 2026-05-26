"""Drive the TAP website: log in, filter, iterate orders, export ODS.

Site-specific selectors live in the SELECTORS dict near the top of this file.
They are best-guess placeholders — confirm and tweak them on the first live run
by opening the site in Chrome devtools and matching against the actual DOM.
"""

from __future__ import annotations

import datetime as dt
import logging
import pickle
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

log = logging.getLogger(__name__)

TAP_URL = "https://tap.dor.ms.gov/_/"
COOKIES_FILE = Path("cookies.pkl")

# --- SITE-SPECIFIC SELECTORS --------------------------------------------------
# These are placeholders. Update on first live run.
SELECTORS = {
    "login_username": (By.ID, "Username"),
    "login_password": (By.ID, "Password"),
    "login_submit": (By.CSS_SELECTOR, "button[type='submit']"),
    "mfa_trust_device": (By.XPATH, "//input[@type='checkbox' and contains(@id, 'trust')]"),
    "mfa_submit": (By.CSS_SELECTOR, "button[type='submit']"),
    # The "Submitted" date filter input on the orders list:
    "submitted_filter": (By.CSS_SELECTOR, "input[name='submitted']"),
    "filter_apply": (By.XPATH, "//button[contains(., 'Search') or contains(., 'Apply')]"),
    # Each order row in the filtered list — link/anchor to detail page:
    "order_row_links": (By.CSS_SELECTOR, "a.order-link, table tbody tr a"),
    # On the order detail page:
    "detail_order_number": (By.CSS_SELECTOR, "[data-field='orderNumber'], .order-number"),
    "detail_export_button": (
        By.XPATH,
        "//a[contains(., 'Export')] | //button[contains(., 'Export')]",
    ),
}
# -----------------------------------------------------------------------------


@dataclass
class OrderHandle:
    order_number: str
    ods_path: Path


def build_driver(downloads_dir: Path, headless: bool = False) -> webdriver.Chrome:
    """Chrome with automatic ODS downloads into downloads_dir."""
    downloads_dir.mkdir(parents=True, exist_ok=True)
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1400,1000")
    prefs = {
        "download.default_directory": str(downloads_dir.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    opts.add_experimental_option("prefs", prefs)
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=opts)


def _save_cookies(driver: webdriver.Chrome) -> None:
    with COOKIES_FILE.open("wb") as f:
        pickle.dump(driver.get_cookies(), f)
    log.info("Saved %d cookies to %s", len(driver.get_cookies()), COOKIES_FILE)


def _load_cookies(driver: webdriver.Chrome) -> bool:
    if not COOKIES_FILE.exists():
        return False
    driver.get(TAP_URL)
    with COOKIES_FILE.open("rb") as f:
        cookies = pickle.load(f)
    for c in cookies:
        # selenium rejects cookies with sameSite values it doesn't expect
        c.pop("sameSite", None)
        try:
            driver.add_cookie(c)
        except Exception as e:
            log.debug("skip cookie %s: %s", c.get("name"), e)
    driver.get(TAP_URL)
    return True


def _is_logged_in(driver: webdriver.Chrome) -> bool:
    """Heuristic: if the login form isn't on the page, assume we're past it."""
    try:
        driver.find_element(*SELECTORS["login_username"])
        return False
    except NoSuchElementException:
        return True


def login(driver: webdriver.Chrome, username: str, password: str) -> None:
    """Fresh login flow. Prompts at the terminal for MFA confirmation."""
    driver.get(TAP_URL)
    wait = WebDriverWait(driver, 20)
    wait.until(EC.visibility_of_element_located(SELECTORS["login_username"]))
    driver.find_element(*SELECTORS["login_username"]).send_keys(username)
    driver.find_element(*SELECTORS["login_password"]).send_keys(password)
    driver.find_element(*SELECTORS["login_submit"]).click()

    print("\n=== If MFA is shown in the browser, complete it now. ===")
    print("=== Be sure to tick 'Trust this device' before submitting. ===")
    input("Press Enter once you're past MFA and on the post-login page... ")
    _save_cookies(driver)


def load_or_login(driver: webdriver.Chrome, username: str, password: str) -> None:
    if _load_cookies(driver) and _is_logged_in(driver):
        log.info("Reusing saved session cookies")
        return
    log.info("No usable session, performing full login")
    login(driver, username, password)


def apply_filter(driver: webdriver.Chrome, target_date: dt.date) -> None:
    """Set submitted=MM-DD-YYYY and apply."""
    wait = WebDriverWait(driver, 20)
    field = wait.until(EC.visibility_of_element_located(SELECTORS["submitted_filter"]))
    field.clear()
    field.send_keys(target_date.strftime("%m-%d-%Y"))
    driver.find_element(*SELECTORS["filter_apply"]).click()
    # Give the table a moment to refresh.
    time.sleep(1.5)


def _wait_for_ods(downloads_dir: Path, before: set[Path], timeout: float = 30.0) -> Path:
    """Block until a new .ods file lands in downloads_dir, return its path."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        current = {p for p in downloads_dir.iterdir() if p.suffix.lower() == ".ods"}
        # Ignore Chrome's partial-download placeholder .crdownload files via suffix check.
        new = current - before
        if new:
            path = next(iter(new))
            # wait for the size to stabilize (download fully flushed)
            last_size = -1
            for _ in range(20):
                size = path.stat().st_size
                if size > 0 and size == last_size:
                    return path
                last_size = size
                time.sleep(0.25)
            return path
        time.sleep(0.5)
    raise TimeoutException(f"No new ODS appeared in {downloads_dir} within {timeout}s")


def iter_orders(driver: webdriver.Chrome, downloads_dir: Path):
    """Yield OrderHandle for each order in the filtered list.

    Each iteration:
      1. clicks into the order detail page,
      2. scrapes the order number,
      3. triggers Export, waits for the ODS download,
      4. navigates back to the list.
    """
    wait = WebDriverWait(driver, 20)
    list_url = driver.current_url
    # snapshot link count up-front; re-fetch by index since DOM changes after nav
    links = driver.find_elements(*SELECTORS["order_row_links"])
    n = len(links)
    log.info("Filter matched %d orders", n)

    for i in range(n):
        driver.get(list_url)
        links = driver.find_elements(*SELECTORS["order_row_links"])
        if i >= len(links):
            log.warning("Order list shrank between iterations at index %d", i)
            break
        links[i].click()

        # scrape order number
        order_el = wait.until(EC.visibility_of_element_located(SELECTORS["detail_order_number"]))
        order_number = order_el.text.strip()

        # trigger export
        before = {p for p in downloads_dir.iterdir() if p.suffix.lower() == ".ods"}
        export_btn = wait.until(EC.element_to_be_clickable(SELECTORS["detail_export_button"]))
        export_btn.click()
        ods_path = _wait_for_ods(downloads_dir, before)

        yield OrderHandle(order_number=order_number, ods_path=ods_path)


@contextmanager
def chrome_session(downloads_dir: Path, headless: bool = False):
    driver = build_driver(downloads_dir, headless=headless)
    try:
        yield driver
    finally:
        driver.quit()
