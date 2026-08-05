"""
Cuckoo+ Service Specialist — monthly list scraper (Appium / UiAutomator2)

HOW THIS APP'S SCREENS WORK (based on your Appium Inspector capture)
----------------------------------------------------------------------
- The CUSTOMER LIST screen shows, per row: an order number, a product
  code, a customer name, and a "View Order" button.
- Tapping "View Order" opens ONE scrollable detail screen (not tabs)
  with labeled sections in this order:
      Billing Address & Contact       -> name, phone
      Installation/Service Address & Contact -> address, name, phone
      Sales Info                      -> dealer/agent code
- The section headers (e.g. "Billing Address & Contact") are the same
  wording on every record — only the values change. So instead of
  matching on a specific customer's data (which only works for the one
  record you inspected), this script finds each header and reads
  whatever TextView comes right after it. That's what makes it reusable
  across all your customers.

IMPORTANT — VALIDATE BEFORE TRUSTING THIS
-------------------------------------------
The app has no resource-ids, so there's no rock-solid way to identify
elements — everything here is a best-effort structural guess based on
ONE record you captured. Two things to check on your first test run:

1. LIST ROW GROUPING: the script assumes each row contributes exactly
   3 TextViews (order number, product, name) in that order, matching
   the order "View Order" buttons appear on screen. If your list also
   shows other TextViews (e.g. a status label, a date), this grouping
   will be wrong and rows will look shuffled. The script prints a
   warning if the counts don't divide evenly — read it if it appears.

2. DETAIL FIELD OFFSETS: the "following TextView number N after this
   header" offsets below match the ONE record you captured. If a
   different customer's record has an extra line (e.g. a second phone
   number, or a blank/missing field), offsets could shift. Test on a
   few different customers, not just one, before trusting a full run.

Run with LIST_LIMIT = 3 first. Check the CSV. Only then raise the limit.
"""

import csv
import time
from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.common.appiumby import AppiumBy
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ============================================================
# CONFIG
# ============================================================

APP_PACKAGE = "cuckoo.doctress"          # confirm with: adb shell pm list packages | findstr cuckoo
APP_ACTIVITY = ".MainActivity"           # confirm with: adb shell dumpsys window | findstr mCurrentFocus

# Start small. Raise only after checking the output CSV looks correct.
LIST_LIMIT = 3

# Button that opens a record's detail screen from the list.
# NOTE the leading space before "View Order" — that's exactly what
# Inspector captured, and text matches must be exact.
VIEW_ORDER_BUTTON = (AppiumBy.XPATH, '//android.widget.Button[@text=" View Order"]')

# How many plain TextViews make up one row on the LIST screen, and in
# what order. Adjust if your list shows more/fewer fields per row.
ROW_FIELDS_PER_CUSTOMER = ["order_number", "product_code", "name"]

# Section headers on the DETAIL screen, and how many TextViews after
# each header to read, in order. These offsets came directly from your
# capture — verify against a second customer record before trusting them.
DETAIL_SECTIONS = {
    "Billing Address & Contact": {
        "billing_name": 1,
        "billing_phone": 2,
    },
    "Installation/Service Address & Contact": {
        "install_address": 1,
        "install_name": 2,
        "install_phone": 3,
    },
    "Sales Info": {
        "sales_agent": 1,
    },
}

OUTPUT_CSV = "cuckoo_export.csv"
WAIT_SECONDS = 10


# ============================================================
# Helpers
# ============================================================

def build_driver():
    options = UiAutomator2Options()
    options.platform_name = "Android"
    options.app_package = APP_PACKAGE
    options.app_activity = APP_ACTIVITY
    options.no_reset = True  # keep you logged in between runs
    return webdriver.Remote("http://127.0.0.1:4723", options=options)


def wait_for(driver, selector, timeout=WAIT_SECONDS):
    by, value = selector
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


def read_text_safe(el):
    try:
        return el.text.strip()
    except Exception:
        return ""


# ============================================================
# List screen: read all rows without opening any of them
# ============================================================

def read_list_rows(driver):
    """
    Returns a list of dicts, one per customer, built from the flat set of
    TextViews on the list screen, grouped in chunks of len(ROW_FIELDS_PER_CUSTOMER).
    """
    wait_for(driver, VIEW_ORDER_BUTTON)
    all_texts = driver.find_elements(AppiumBy.CLASS_NAME, "android.widget.TextView")
    buttons = driver.find_elements(*VIEW_ORDER_BUTTON)

    per_row = len(ROW_FIELDS_PER_CUSTOMER)
    expected = per_row * len(buttons)

    if len(all_texts) != expected:
        print(
            f"  !! WARNING: found {len(all_texts)} TextViews but expected "
            f"{expected} ({per_row} x {len(buttons)} buttons). "
            f"Row grouping is probably WRONG — inspect the list screen again "
            f"before trusting this output."
        )

    rows = []
    for i in range(len(buttons)):
        chunk = all_texts[i * per_row:(i + 1) * per_row]
        row = {ROW_FIELDS_PER_CUSTOMER[j]: read_text_safe(chunk[j]) for j in range(min(per_row, len(chunk)))}
        rows.append(row)
    return rows


# ============================================================
# Detail screen: read sections by anchoring on header text
# ============================================================

def read_detail_screen(driver):
    record = {}
    for header_text, fields in DETAIL_SECTIONS.items():
        try:
            wait_for(driver, (AppiumBy.XPATH, f'//android.widget.TextView[@text="{header_text}"]'), timeout=5)
        except TimeoutException:
            print(f"  !! Header not found on this record: '{header_text}' — skipping its fields.")
            for field_name in fields:
                record[field_name] = ""
            continue

        for field_name, offset in fields.items():
            xpath = (
                f'//android.widget.TextView[@text="{header_text}"]'
                f'/following::android.widget.TextView[{offset}]'
            )
            try:
                el = driver.find_element(AppiumBy.XPATH, xpath)
                record[field_name] = read_text_safe(el)
            except NoSuchElementException:
                record[field_name] = ""
    return record


# ============================================================
# Main loop
# ============================================================

def run():
    driver = build_driver()
    time.sleep(3)  # let the app finish loading

    all_records = []

    try:
        list_rows = read_list_rows(driver)
        total = min(len(list_rows), LIST_LIMIT)
        print(f"Found {len(list_rows)} rows in the list — processing {total}.")

        for i in range(total):
            # Re-fetch buttons each loop: navigating away invalidates old
            # element references (a common Appium gotcha).
            buttons = driver.find_elements(*VIEW_ORDER_BUTTON)
            label = list_rows[i].get("name") or list_rows[i].get("order_number") or f"row {i+1}"
            print(f"[{i+1}/{total}] Opening: {label}")

            buttons[i].click()
            time.sleep(1.5)

            try:
                detail = read_detail_screen(driver)
                merged = {**list_rows[i], **detail}
                all_records.append(merged)
            except (NoSuchElementException, TimeoutException) as e:
                print(f"  !! Skipped row {i+1} due to error: {e}")

            driver.back()   # no in-app back button was found in your capture;
            time.sleep(1)   # using the phone's native back instead.

    finally:
        driver.quit()

    write_csv(all_records)
    print(f"Done. Wrote {len(all_records)} records to {OUTPUT_CSV}")


def write_csv(records):
    if not records:
        print("No records captured — nothing written.")
        return
    fieldnames = sorted({key for r in records for key in r.keys()})
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)


if __name__ == "__main__":
    run()
