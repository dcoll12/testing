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

SHEET_START_ROW = 2    # first row to write URLs into (1-indexed)
SHEET_COLUMN    = "B"  # column to paste URLs

# ── Resume support ─────────────────────────────────────────────────────────
# Set SKIP_FIRST_N > 0 to fast-scroll past grants already in the sheet.
# The script marks the first N grants as "already processed" without
# opening them, then starts writing at SHEET_START_ROW + SKIP_FIRST_N.
SKIP_FIRST_N = 0   # starting fresh — process every grant from the top
# ──────────────────────────────────────────────────────────────────────────

SHORT_WAIT  = 3   # seconds for short pauses
LONG_WAIT   = 10  # seconds for WebDriverWait timeout
# ─────────────────────────────────────────────────────────────────────────────


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


def sheets_type_url(driver, url: str):
    """Type a URL into the active Sheets cell, then press Enter to move down."""
    active = driver.switch_to.active_element
    active.send_keys(url)
    active.send_keys(Keys.RETURN)   # confirm + advance to next row
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


def scroll_to_bottom(driver):
    """
    Scroll the grants list to its absolute bottom so Instrumentl's
    infinite-scroll intersection observer fires and loads the next batch.
    """
    driver.execute_script("""
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
        if (el) {
            el.scrollTop = el.scrollHeight;
        } else {
            window.scrollTo(0, document.body.scrollHeight);
        }
    """)
    time.sleep(3)   # give Ember time to render the next batch


def open_grant_and_get_url(driver, grant_row) -> str | None:
    """
    Click a grant row, go to the Funding Opportunity tab,
    and return the website URL (or None if not found).
    """
    wait = wait_for(driver)
    grant_row.click()
    time.sleep(SHORT_WAIT)

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

    try:
        view_website = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".grant-website-url"))
        )
        return view_website.get_attribute("href")
    except TimeoutException:
        print("  ✗ 'View website' link not found.")
        return None


def save_grant(driver):
    """
    Click the Save button on the open grant modal.
    Uses the same selector priority as the original bookmarklet:
      1. .save-button-container > .btn  (most specific)
      2. Any visible button whose text is exactly "save"
      3. [aria-label="Save"]
    Logs a warning but does NOT raise if the button isn't found.
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
            return
        except TimeoutException:
            continue

    print("  ⚠ Save button not found — grant not saved.")


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
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)


def main():
    driver = make_driver()

    # ── 1. Open Google Sheets ────────────────────────────────────────────────
    print("Opening Google Sheets …")
    driver.get(SPREADSHEET_URL)
    sheets_handle = driver.current_window_handle
    time.sleep(SHORT_WAIT + 3)   # give Sheets extra time to paint

    # Navigate to the starting cell once — cursor will advance on its own
    start_cell = f"{SHEET_COLUMN}{SHEET_START_ROW + SKIP_FIRST_N}"
    print(f"Navigating to starting cell {start_cell} …")
    sheets_go_to_start(driver, start_cell)

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
    processed_names = set()

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
    MAX_EMPTY_SCROLLS = 10

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
            if grant_rows:
                scroll_element_into_view(driver, grant_rows[-1], block="end")
            scroll_to_bottom(driver)
            continue

        # Process the grant
        no_new_rows_count = 0
        total_processed  += 1
        processed_names.add(next_row_text)
        sheet_row = SHEET_START_ROW + SKIP_FIRST_N + total_processed - 1
        print(f"\n[{total_processed}] {next_row_text}")

        scroll_element_into_view(driver, next_row)
        website_url = open_grant_and_get_url(driver, next_row)

        # Save the grant on Instrumentl (mirrors the bookmarklet behaviour)
        save_grant(driver)

        if website_url:
            print(f"  URL: {website_url}  → row {sheet_row}")
            driver.switch_to.window(sheets_handle)
            time.sleep(1)
            sheets_type_url(driver, website_url)   # Enter advances the row
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
