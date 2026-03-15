"""
Instrumentl → Google Sheets Automation
---------------------------------------
For each grant on Instrumentl:
  1. Click the grant → Funding Opportunity tab → extract website URL
  2. Switch to Google Sheets tab → type URL into the next row
  3. Switch back to Instrumentl → close modal → next grant

Sheets writing strategy
-----------------------
We navigate to the starting cell ONCE at startup using the Name Box.
Each sheets_type_url() call types the URL then presses Enter, which
advances the cursor down one row automatically.  This avoids the bug
where the cell address (e.g. "B46") was being typed as literal text
instead of used for navigation.
"""

import csv
import io
import os
import pathlib
import re
import time
import requests
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementClickInterceptedException
)

# ── Configuration ────────────────────────────────────────────────────────────
SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1hc_Ehb2evMR5h5kbQKfrWmiz53C-DO4Aa410u2yxa0E/edit?gid=729130956#gid=729130956"
)
INSTRUMENTL_URL = "https://www.instrumentl.com/projects#/all-projects"

SHEET_START_ROW = 2    # first row to write into (1-indexed)
SHEET_COLUMN    = "A"  # starting column (name goes here, URL goes in the next column)

# ── Resume support ─────────────────────────────────────────────────────────
# Set SKIP_FIRST_N > 0 to fast-scroll past grants already in the sheet.
# The script marks the first N grants as "already processed" without
# opening them, then starts writing at SHEET_START_ROW + SKIP_FIRST_N.
SKIP_FIRST_N = 0   # starting fresh — process every grant from the top

# Local file that persists processed grant names across runs.
# Delete this file to start completely fresh.
PROGRESS_FILE = pathlib.Path(__file__).parent / "processed_grants.txt"
# ──────────────────────────────────────────────────────────────────────────

SHORT_WAIT  = 3   # seconds for short pauses
LONG_WAIT   = 10  # seconds for WebDriverWait timeout
# ─────────────────────────────────────────────────────────────────────────────


def load_processed_names() -> set:
    """Load the set of already-processed grant names from disk."""
    if not PROGRESS_FILE.exists():
        return set()
    names = {line.strip() for line in PROGRESS_FILE.read_text(encoding="utf-8").splitlines() if line.strip()}
    print(f"  Loaded {len(names)} previously processed grants from {PROGRESS_FILE.name}")
    return names


def save_processed_name(name: str):
    """Append a single grant name to the progress file."""
    with PROGRESS_FILE.open("a", encoding="utf-8") as f:
        f.write(name + "\n")


def read_existing_sheet_names(driver, sheets_handle) -> tuple[set, int]:
    """
    Download the Google Sheet as CSV using the active browser session and
    return (set_of_existing_names_in_col_A, number_of_filled_data_rows).

    This lets the script skip grants that are already in the sheet even when
    the local processed_grants.txt cache is missing or out of date.
    Falls back to (empty set, 0) on any error so the script still runs.
    """
    driver.switch_to.window(sheets_handle)

    match = re.search(r'/spreadsheets/d/([^/]+)', SPREADSHEET_URL)
    if not match:
        print("  Warning: could not parse spreadsheet ID — skipping sheet read.")
        return set(), 0
    sheet_id = match.group(1)

    gid_match = re.search(r'gid=(\d+)', SPREADSHEET_URL)
    gid = gid_match.group(1) if gid_match else '0'

    export_url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}"
        f"/export?format=csv&gid={gid}"
    )

    # Copy the browser's current cookies into a requests session so the
    # export request is authenticated with the same Google account.
    session = requests.Session()
    for c in driver.get_cookies():
        session.cookies.set(c['name'], c['value'])
    session.headers['User-Agent'] = driver.execute_script('return navigator.userAgent')

    try:
        resp = session.get(export_url, timeout=20, allow_redirects=True)
        resp.raise_for_status()
    except Exception as exc:
        print(f"  Warning: could not read sheet via export ({exc}) — relying on local cache only.")
        return set(), 0

    names: set = set()
    last_filled_row = 0
    for row_idx, row in enumerate(csv.reader(io.StringIO(resp.text)), start=1):
        if row_idx < SHEET_START_ROW:
            continue  # skip header rows
        if row and row[0].strip():
            names.add(row[0].strip())
            last_filled_row = row_idx

    filled_count = (
        last_filled_row - SHEET_START_ROW + 1
        if last_filled_row >= SHEET_START_ROW
        else 0
    )
    print(
        f"  Google Sheet: {len(names)} existing entries "
        f"(last data row: {last_filled_row or 'none'})"
    )
    return names, filled_count


def make_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(1200, 900)
    return driver


def wait_for(driver, timeout=LONG_WAIT) -> WebDriverWait:
    return WebDriverWait(driver, timeout)


def sheets_go_to_start(driver, cell_address: str):
    """
    Navigate to the starting cell exactly once using the Name Box.
    After this, every sheets_type_url() call advances one row via Enter.

    The Name Box in Google Sheets is a <div> (not <input>), so we search
    by class/aria-label without restricting to input elements.
    """
    # Wait until the spreadsheet grid is actually rendered before proceeding
    grid_selectors = [
        ".waffle-name-box", ".cell-input",
        "canvas#waffle-grid-container", ".grid-container",
    ]
    for sel in grid_selectors:
        try:
            wait_for(driver, timeout=30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, sel))
            )
            print(f"  Sheet ready (found '{sel}')")
            break
        except TimeoutException:
            continue

    # Escape out of any cell-edit mode
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(0.5)

    # Find the Name Box — it is a <div>, NOT an <input>
    name_box = None
    for by, selector in [
        (By.CSS_SELECTOR, ".waffle-name-box"),
        (By.CSS_SELECTOR, ".cell-input"),
        (By.CSS_SELECTOR, "[aria-label='Name Box']"),
        (By.XPATH,        "//*[@aria-label='Name Box']"),
        (By.XPATH,        "//*[contains(@class,'waffle-name-box')]"),
        (By.XPATH,        "//*[contains(@class,'cell-input')]"),
    ]:
        try:
            name_box = wait_for(driver, timeout=8).until(
                EC.element_to_be_clickable((by, selector))
            )
            print(f"  Name Box found via: {selector}")
            break
        except TimeoutException:
            continue

    if name_box is None:
        raise RuntimeError("Could not find Google Sheets Name Box — is the sheet loaded?")

    name_box.click()
    time.sleep(0.4)
    name_box.send_keys(Keys.CONTROL + "a")
    name_box.send_keys(cell_address)
    name_box.send_keys(Keys.RETURN)
    time.sleep(0.8)


def sheets_write_row(driver, name: str, url: str):
    """
    Write grant name to the current cell (col A) and URL to the next cell (col B),
    then press Enter to move down to the next row.
    Assumes the cursor is already on column A of the target row.
    """
    active = driver.switch_to.active_element
    active.send_keys(name)
    active.send_keys(Keys.TAB)      # move right to col B
    active = driver.switch_to.active_element
    active.send_keys(url)
    active.send_keys(Keys.RETURN)   # confirm + move down (returns to col A)
    time.sleep(0.5)


def instrumentl_sort_by_grant_name(driver):
    """
    Click the sort dropdown and choose 'Grant Name'.
    The hashed Ember class (._trigger_XXXXX) changes between deployments,
    so we try multiple stable selectors and skip gracefully if none work.
    """
    wait = wait_for(driver)

    # Try stable selectors for the sort trigger (most to least specific)
    trigger = None
    trigger_selectors = [
        # The label span inside any trigger-like element
        (By.CSS_SELECTOR, ".table-sort-item-label"),
        # Buttons/divs that contain the sort label text
        (By.XPATH, "//*[contains(@class,'table-sort-item-label')]"),
        # Any element whose text looks like a sort control
        (By.XPATH, "//span[normalize-space()='Grant Name' or normalize-space()='Date Added' or normalize-space()='Sort']"),
        # Fallback: the old hashed class (works if Ember version hasn't changed)
        (By.CSS_SELECTOR, "._trigger_163dkf > .table-sort-item-label"),
    ]
    for by, sel in trigger_selectors:
        try:
            trigger = wait_for(driver, timeout=5).until(
                EC.element_to_be_clickable((by, sel))
            )
            break
        except TimeoutException:
            continue

    if trigger is None:
        print("  ⚠ Could not find sort trigger — skipping sort step.")
        print("    Grants will be processed in default order.")
        return

    trigger.click()
    time.sleep(0.5)

    try:
        grant_name_opt = wait_for(driver, timeout=5).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(.,'Grant Name') and not(self::html) and not(self::body)]")
            )
        )
        grant_name_opt.click()
        time.sleep(SHORT_WAIT)
        print("  Sorted by Grant Name.")
    except TimeoutException:
        print("  ⚠ 'Grant Name' option not found after clicking trigger — skipping sort.")


def get_grant_rows(driver):
    """Return all visible grant row elements."""
    return driver.find_elements(By.CSS_SELECTOR, ".name-and-owner-column")


def scroll_element_into_view(driver, element, block: str = "center"):
    """Scroll an element into the viewport."""
    driver.execute_script(
        f"arguments[0].scrollIntoView({{block:'{block}', inline:'nearest'}});", element
    )
    time.sleep(0.5)


def _find_scroll_container(driver):
    """Return JS expression that resolves to the scrollable list container."""
    return """(function() {
        var el = document.getElementById('0-scrollable');
        if (!el) {
            var divs = Array.from(document.querySelectorAll('div'))
                            .filter(function(d) {
                                return d.scrollHeight > d.clientHeight + 100;
                            })
                            .sort(function(a, b) {
                                return b.scrollHeight - a.scrollHeight;
                            });
            el = divs[0] || null;
        }
        return el;
    })()"""


def scroll_to_bottom(driver):
    """
    Scroll the grants list to its absolute bottom so Instrumentl's
    infinite-scroll intersection observer fires and loads the next batch.
    Pumps scroll three times so the observer reliably fires even after
    a modal close that resets the scroll position.
    """
    script = f"""
        var el = {_find_scroll_container(driver)};
        if (el) {{
            el.scrollTop = el.scrollHeight;
        }} else {{
            window.scrollTo(0, document.body.scrollHeight);
        }}
    """
    for _ in range(3):
        driver.execute_script(script)
        time.sleep(1.5)
    time.sleep(2)   # give Ember time to render the next batch


def get_scroll_top(driver) -> int:
    """Return the current scrollTop of the grants container (0 if not found)."""
    return driver.execute_script(f"""
        var el = {_find_scroll_container(driver)};
        return el ? el.scrollTop : 0;
    """) or 0


def open_grant_and_get_name_and_url(driver, grant_row) -> tuple[str | None, str | None]:
    """
    Click a grant row, go to the Funding Opportunity tab, and return (name, url).
    Name is read from the row element before clicking (most reliable source).
    """
    wait = wait_for(driver)

    # Get name from the row element text before clicking
    name = grant_row.text.strip().splitlines()[0] if grant_row.text.strip() else None

    grant_row.click()
    time.sleep(SHORT_WAIT)

    # Click the Funding Opportunity tab
    try:
        funding_tab = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(.,'Funding Opportunity')]")
            )
        )
        funding_tab.click()
        time.sleep(1.5)
    except TimeoutException:
        print("  ✗ 'Funding Opportunity' tab not found.")
        return name, None

    # Get URL from the View website link href (no need to open the tab)
    try:
        view_website = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".grant-website-url"))
        )
        url = view_website.get_attribute("href")
    except TimeoutException:
        print("  ✗ 'View website' link not found.")
        url = None

    return name, url


def save_grant(driver) -> bool:
    """
    Click the Save button on the open grant modal.
    Uses the same selector priority as the original bookmarklet:
      1. .save-button-container > .btn  (most specific)
      2. Any visible button whose text is exactly "save"
      3. [aria-label="Save"]
    Returns True if the save button was found and clicked, False otherwise.
    """
    # Try the specific container selector first
    for by, sel in [
        (By.CSS_SELECTOR, ".save-button-container > .btn"),
        (By.CSS_SELECTOR, "[aria-label='Save'], [aria-label='save']"),
        (By.XPATH, "//button[normalize-space(translate(text(),'SAVE','save'))='save']"),
    ]:
        try:
            btn = wait_for(driver, timeout=5).until(
                EC.element_to_be_clickable((by, sel))
            )
            driver.execute_script(
                "arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", btn
            )
            time.sleep(0.4)
            btn.click()
            time.sleep(1.0)   # brief pause after saving
            print("  ✓ Grant saved.")
            return True
        except TimeoutException:
            continue

    print("  ⚠ Save button not found — grant not saved.")
    return False


def close_grant_modal(driver):
    """Close the grant detail modal."""
    for sel in [
        ".in > .modal-dialog > .modal-content > .modal-header > .close > span",
        ".modal-lg .modal-header span",
    ]:
        try:
            close_btn = wait_for(driver, timeout=5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
            )
            close_btn.click()
            time.sleep(1)
            return
        except TimeoutException:
            continue
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(1)


def instrumentl_login(driver):
    """
    Log in to Instrumentl. Reads from .env if present, otherwise uses
    the hardcoded fallback credentials.
    """
    load_dotenv()
    email    = os.environ.get("INSTRUMENTL_EMAIL",    "darian@aracities.org")
    password = os.environ.get("INSTRUMENTL_PASSWORD", "ZE,EP2MLv3r=kh]")

    driver.get("https://www.instrumentl.com/login")
    wait = WebDriverWait(driver, 15)

    email_field = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "input[type='email'], input[name='email'], input[id*='email']")
    ))
    email_field.clear()
    email_field.send_keys(email)

    password_field = driver.find_element(
        By.CSS_SELECTOR, "input[type='password']"
    )
    password_field.clear()
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    # Wait until we're redirected away from the login page
    wait.until(EC.url_changes("https://www.instrumentl.com/login"))
    print("  Logged in successfully.")
    time.sleep(2)


def main():
    driver = make_driver()

    # ── 1. Open Google Sheets ────────────────────────────────────────────────
    print("Opening Google Sheets …")
    driver.get(SPREADSHEET_URL)
    sheets_handle = driver.current_window_handle
    time.sleep(SHORT_WAIT + 3)   # give Sheets extra time to paint

    # Check the sheet for entries already written so we can skip them
    print("Checking Google Sheet for existing entries …")
    sheet_names, sheet_count = read_existing_sheet_names(driver, sheets_handle)

    # Navigate to the starting cell once — cursor will advance on its own
    start_cell = f"{SHEET_COLUMN}{SHEET_START_ROW + SKIP_FIRST_N}"
    print(f"Navigating to starting cell {start_cell} …")
    sheets_go_to_start(driver, start_cell)

    # ── 2. Open Instrumentl in a new tab and log in ──────────────────────────
    print("Opening Instrumentl …")
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[-1])
    print("  Logging in …")
    instrumentl_login(driver)
    driver.get(INSTRUMENTL_URL)
    instrumentl_handle = driver.current_window_handle
    time.sleep(SHORT_WAIT + 1)

    # ── 3. Sort grants by Grant Name ─────────────────────────────────────────
    print("Sorting by Grant Name …")
    instrumentl_sort_by_grant_name(driver)

    # ── 4. Load previously processed grants & merge with sheet contents ──────
    processed_names = load_processed_names()

    # Add any names found directly in the sheet that aren't in the local cache.
    # This ensures grants already written to the sheet are never re-processed,
    # even when processed_grants.txt is missing or was created on another machine.
    new_from_sheet = sheet_names - processed_names
    if new_from_sheet:
        print(f"  {len(new_from_sheet)} sheet entries not in local cache — adding to skip list …")
        processed_names |= new_from_sheet
        for name in new_from_sheet:
            save_processed_name(name)

    # Use the sheet's actual row count as the authoritative resume position.
    already_done = sheet_count if sheet_count > 0 else len(processed_names)

    # Jump the sheet cursor past rows already written in previous runs
    if already_done > 0:
        resume_cell = f"{SHEET_COLUMN}{SHEET_START_ROW + already_done}"
        print(f"  Resuming — navigating sheet to {resume_cell} …")
        driver.switch_to.window(sheets_handle)
        sheets_go_to_start(driver, resume_cell)
        driver.switch_to.window(instrumentl_handle)

    if SKIP_FIRST_N > 0:
        print(f"Skipping first {SKIP_FIRST_N} grants (already in sheet) …")
        skipped    = 0
        skip_empty = 0
        while skipped < SKIP_FIRST_N:
            rows = get_grant_rows(driver)
            new_in_batch = 0
            for row in rows:
                text = row.text.strip().splitlines()[0] if row.text.strip() else ""
                if text and text not in processed_names:
                    processed_names.add(text)
                    skipped      += 1
                    new_in_batch += 1
                    if skipped >= SKIP_FIRST_N:
                        break
            if new_in_batch == 0:
                skip_empty += 1
                if skip_empty > 5:
                    print(f"  Warning: only found {skipped} to skip (expected {SKIP_FIRST_N}).")
                    break
                if rows:
                    scroll_element_into_view(driver, rows[-1], block="end")
                scroll_to_bottom(driver)
            else:
                skip_empty = 0
                if skipped < SKIP_FIRST_N and rows:
                    scroll_element_into_view(driver, rows[-1], block="end")
                    scroll_to_bottom(driver)
        print(f"  Skipped {skipped} grants. Resuming at {start_cell}.")

    # ── 5. Iterate over all remaining grants ─────────────────────────────────
    total_processed   = 0
    no_new_rows_count = 0
    MAX_EMPTY_SCROLLS = 40

    while True:
        grant_rows = get_grant_rows(driver)

        # Find the first unprocessed row in the current DOM snapshot
        next_row      = None
        next_row_text = None
        for row in grant_rows:
            text = row.text.strip().splitlines()[0] if row.text.strip() else ""
            if text and text not in processed_names:
                next_row      = row
                next_row_text = text
                break

        # Nothing new — scroll to trigger the next batch
        if next_row is None:
            no_new_rows_count += 1
            if no_new_rows_count > MAX_EMPTY_SCROLLS:
                print(f"\nNo new grants after {MAX_EMPTY_SCROLLS} scroll attempts. All done.")
                break
            print(
                f"\n  ↓ Scroll attempt {no_new_rows_count}/{MAX_EMPTY_SCROLLS} "
                f"(processed {total_processed} so far) …"
            )
            # Every 3rd attempt: scroll UP first to reset the observer, then back down
            if no_new_rows_count % 3 == 0:
                driver.execute_script(f"""
                    var el = {_find_scroll_container(driver)};
                    if (el) el.scrollTop = 0;
                """)
                time.sleep(1)
            if grant_rows:
                scroll_element_into_view(driver, grant_rows[-1], block="end")
            scroll_to_bottom(driver)
            time.sleep(3)   # extra wait for Ember to render
            continue
        processed_names.add(next_row_text)
        save_processed_name(next_row_text)
        sheet_row = SHEET_START_ROW + already_done + total_processed - 1
        print(f"\n[{already_done + total_processed}] {next_row_text}")

        scroll_element_into_view(driver, next_row)
        modal_name, website_url = open_grant_and_get_name_and_url(driver, next_row)
        grant_name = modal_name or next_row_text  # fall back to row text if modal name missing

        if website_url:
            print(f"  NAME: {grant_name}")
            print(f"  URL:  {website_url}  → row {sheet_row}")
            driver.switch_to.window(sheets_handle)
            time.sleep(1)
            sheets_write_row(driver, grant_name, website_url)
            driver.switch_to.window(instrumentl_handle)
            time.sleep(1)

        # Close modal, then nudge scroll so the list stays positioned
        close_grant_modal(driver)
        grant_rows = get_grant_rows(driver)
        if grant_rows:
            scroll_element_into_view(driver, grant_rows[-1], block="end")

    print(f"\nAll done! {total_processed} grants processed. Check your Google Sheet.")
    input("Press Enter to close the browser …")
    driver.quit()


if __name__ == "__main__":
    main()
