"""
Instrumentl → Google Sheets Automation
---------------------------------------
For each grant on Instrumentl:
  1. Click the grant → Funding Opportunity tab → extract website URL
  2. Switch to Google Sheets tab → paste URL into next row
  3. Switch back to Instrumentl → close modal → next grant
"""

import time
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
    "1hc_Ehb2evMR5h5kbQKfrWmiz53C-DO4Aa410u2yxa0E/edit?gid=0"
)
INSTRUMENTL_URL = "https://www.instrumentl.com/projects#/all-projects"

SHEET_START_ROW = 2   # first row to write URLs into (1-indexed)
SHEET_COLUMN    = "B" # column to paste URLs

# ── Resume support ────────────────────────────────────────────────────────────
# Set SKIP_FIRST_N > 0 to fast-scroll past grants already in the sheet.
# The script will mark the first N grants as "already processed" without
# opening them, then start writing at SHEET_START_ROW + SKIP_FIRST_N.
SKIP_FIRST_N = 50   # 50 grants already collected; resume from grant 51
# ─────────────────────────────────────────────────────────────────────────────

SHORT_WAIT  = 3   # seconds for short pauses
LONG_WAIT   = 10  # seconds for WebDriverWait timeout
# ─────────────────────────────────────────────────────────────────────────────


def make_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    # Remove 'enable-automation' banner so Google Sheets behaves normally
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(1200, 900)
    return driver


def wait_for(driver, timeout=LONG_WAIT) -> WebDriverWait:
    return WebDriverWait(driver, timeout)


def sheets_navigate_to_cell(driver, cell_address: str):
    """Click the Google Sheets Name Box and jump to a cell (e.g. 'B2')."""
    # The Name Box is an input inside .cell-input or near the top toolbar.
    # Using keyboard shortcut: Escape first, then click the Name Box.
    try:
        name_box = wait_for(driver).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".cell-input"))
        )
        name_box.click()
        time.sleep(0.3)
        name_box.send_keys(Keys.CONTROL + "a")
        name_box.send_keys(cell_address)
        name_box.send_keys(Keys.RETURN)
        time.sleep(0.5)
    except TimeoutException:
        # Fallback: use Ctrl+Home then keyboard navigation isn't practical,
        # so try an alternate Name Box selector.
        name_box = driver.find_element(
            By.XPATH, "//input[contains(@class,'cell-input') or @aria-label='Name Box']"
        )
        name_box.click()
        name_box.send_keys(Keys.CONTROL + "a")
        name_box.send_keys(cell_address)
        name_box.send_keys(Keys.RETURN)
        time.sleep(0.5)


def sheets_type_url(driver, url: str):
    """Type (not paste) a URL into the currently selected Sheets cell."""
    active = driver.switch_to.active_element
    active.send_keys(url)
    active.send_keys(Keys.RETURN)   # confirm and move down one row
    time.sleep(0.5)


def instrumentl_sort_by_grant_name(driver):
    """Click the NAME sort trigger → select 'Grant Name'."""
    wait = wait_for(driver)
    trigger = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "._trigger_163dkf > .table-sort-item-label")
        )
    )
    trigger.click()
    time.sleep(0.5)

    grant_name_opt = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//p[contains(.,'Grant Name')]")
        )
    )
    grant_name_opt.click()
    time.sleep(SHORT_WAIT)


def get_grant_rows(driver):
    """Return all visible grant row elements."""
    return driver.find_elements(By.CSS_SELECTOR, ".name-and-owner-column")


def scroll_element_into_view(driver, element, block: str = "center"):
    """Scroll an element into the viewport."""
    driver.execute_script(
        f"arguments[0].scrollIntoView({{block:'{block}', inline:'nearest'}});", element
    )
    time.sleep(0.5)


def scroll_to_bottom(driver):
    """
    Scroll the grants list container to its absolute bottom.
    Instrumentl's infinite-scroll observer fires when the last row
    enters the viewport — scrolling to scrollHeight guarantees that.
    Tries #0-scrollable first, then the tallest scrollable div, then window.
    """
    driver.execute_script("""
        var el = document.getElementById('0-scrollable');
        if (!el) {
            // pick the tallest scrollable div (most likely the list container)
            var divs = Array.from(document.querySelectorAll('div'))
                            .filter(function(d) {
                                return d.scrollHeight > d.clientHeight + 100;
                            })
                            .sort(function(a, b) {
                                return b.scrollHeight - a.scrollHeight;
                            });
            el = divs[0] || null;
        }
        if (el) {
            el.scrollTop = el.scrollHeight;
        } else {
            window.scrollTo(0, document.body.scrollHeight);
        }
    """)
    time.sleep(3)   # give Ember time to render the next batch


def open_grant_and_get_url(driver, grant_row) -> str | None:
    """
    Click a grant row, navigate to Funding Opportunity tab,
    and return the website URL (or None if not found).
    """
    wait = wait_for(driver)
    grant_row.click()
    time.sleep(SHORT_WAIT)

    # Click "Funding Opportunity" tab
    try:
        funding_tab = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(.,'Funding Opportunity')]")
            )
        )
        funding_tab.click()
        time.sleep(1.5)
    except TimeoutException:
        print("  ✗ 'Funding Opportunity' tab not found, skipping.")
        return None

    # Get href from "View website" link without opening it
    try:
        view_website = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".grant-website-url"))
        )
        url = view_website.get_attribute("href")
        return url
    except TimeoutException:
        print("  ✗ 'View website' link not found.")
        return None


def close_grant_modal(driver):
    """Close the grant detail modal."""
    try:
        close_btn = wait_for(driver, timeout=5).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".modal-lg .modal-header span")
            )
        )
        close_btn.click()
        time.sleep(1)
    except TimeoutException:
        # Try pressing Escape as a fallback
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)


def main():
    driver = make_driver()
    wait = wait_for(driver)

    # ── 1. Open Google Sheets ────────────────────────────────────────────────
    print("Opening Google Sheets …")
    driver.get(SPREADSHEET_URL)
    sheets_handle = driver.current_window_handle
    time.sleep(SHORT_WAIT + 1)   # let Sheets fully load

    # ── 2. Open Instrumentl in a new tab ────────────────────────────────────
    print("Opening Instrumentl …")
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[-1])
    driver.get(INSTRUMENTL_URL)
    instrumentl_handle = driver.current_window_handle
    time.sleep(SHORT_WAIT + 1)

    # ── 3. Sort grants by Grant Name ─────────────────────────────────────────
    print("Sorting by Grant Name …")
    instrumentl_sort_by_grant_name(driver)

    # ── 4. Fast-skip already-processed grants ────────────────────────────────
    processed_names   = set()
    current_sheet_row = SHEET_START_ROW + SKIP_FIRST_N

    if SKIP_FIRST_N > 0:
        print(f"Skipping first {SKIP_FIRST_N} grants (already in sheet) …")
        skipped = 0
        skip_empty = 0
        while skipped < SKIP_FIRST_N:
            rows = get_grant_rows(driver)
            new_in_batch = 0
            for row in rows:
                text = row.text.strip().splitlines()[0] if row.text.strip() else ""
                if text and text not in processed_names:
                    processed_names.add(text)
                    skipped += 1
                    new_in_batch += 1
                    if skipped >= SKIP_FIRST_N:
                        break
            if new_in_batch == 0:
                skip_empty += 1
                if skip_empty > 5:
                    print(f"  Warning: only found {skipped} grants to skip (expected {SKIP_FIRST_N}).")
                    break
                if rows:
                    scroll_element_into_view(driver, rows[-1], block="end")
                scroll_to_bottom(driver)
            else:
                skip_empty = 0
                if skipped < SKIP_FIRST_N and rows:
                    scroll_element_into_view(driver, rows[-1], block="end")
                    scroll_to_bottom(driver)
        print(f"  Skipped {skipped} grants. Resuming at sheet row {current_sheet_row}.")

    # ── 5. Iterate over remaining grants ─────────────────────────────────────
    total_processed   = 0
    no_new_rows_count = 0       # consecutive scrolls that produced nothing new
    MAX_EMPTY_SCROLLS = 10      # give up only after 10 fruitless scroll attempts

    while True:
        grant_rows = get_grant_rows(driver)

        # ── Find the first unprocessed row currently in the DOM ──────────────
        next_row      = None
        next_row_text = None
        for row in grant_rows:
            text = row.text.strip().splitlines()[0] if row.text.strip() else ""
            if text and text not in processed_names:
                next_row      = row
                next_row_text = text
                break

        # ── Nothing new visible yet → scroll to bottom to trigger next batch ─
        if next_row is None:
            no_new_rows_count += 1
            if no_new_rows_count > MAX_EMPTY_SCROLLS:
                print(f"\nNo new grants after {MAX_EMPTY_SCROLLS} scroll attempts. All done.")
                break

            print(
                f"\n  ↓ Scroll attempt {no_new_rows_count}/{MAX_EMPTY_SCROLLS} "
                f"(processed {total_processed} so far) …"
            )

            # Scroll the last visible row into view first (fires the intersection
            # observer that Instrumentl uses), then go to absolute bottom.
            if grant_rows:
                scroll_element_into_view(driver, grant_rows[-1], block="end")
            scroll_to_bottom(driver)   # wait=3 s baked in
            continue

        # ── We have a new row — reset the empty-scroll counter ───────────────
        no_new_rows_count = 0
        total_processed  += 1
        processed_names.add(next_row_text)
        print(f"\n[{total_processed}] {next_row_text}")

        # Bring the row to the centre of the viewport before clicking
        scroll_element_into_view(driver, next_row)
        website_url = open_grant_and_get_url(driver, next_row)

        if website_url:
            print(f"  URL: {website_url}")

            # ── Switch to Google Sheets ──────────────────────────────────────
            driver.switch_to.window(sheets_handle)
            time.sleep(1)

            cell_address = f"{SHEET_COLUMN}{current_sheet_row}"
            print(f"  → Writing to cell {cell_address}")
            sheets_navigate_to_cell(driver, cell_address)
            sheets_type_url(driver, website_url)

            current_sheet_row += 1

            # ── Back to Instrumentl ──────────────────────────────────────────
            driver.switch_to.window(instrumentl_handle)
            time.sleep(1)

        # Close the modal, then scroll the last visible row into view so the
        # intersection observer fires and loads the next batch when needed.
        close_grant_modal(driver)
        grant_rows = get_grant_rows(driver)
        if grant_rows:
            scroll_element_into_view(driver, grant_rows[-1], block="end")

    print(f"\nAll done! {total_processed} grants processed. Check your Google Sheet.")
    input("Press Enter to close the browser …")
    driver.quit()


if __name__ == "__main__":
    main()
