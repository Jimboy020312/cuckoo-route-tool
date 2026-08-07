"""
Cuckoo+ Service Specialist — monthly list scraper (Appium / UiAutomator2)

CONFIRMED SCREEN FLOW (from real XML captures)
--------------------------------------------------
1. LIST screen: each customer is a card with label/value pairs (NS No,
   Sales No, NS Date, Status, Appt Date, Product Name, Cust Name) plus
   one button.
2. Tapping that button does NOT open the detail screen directly — it
   opens a "Choose Option" POPUP with several choices (View Order, CCS
   Note, Appointment, Contact List, Cancel Appointment, Cancel). The
   script taps "View Order" from that popup.
3. That opens a "Customer Information" screen with TWO TABS:
     - "Address/Contact Info" (shown by default) — has THREE sections:
       Billing Address & Contact, Installation/Service Address &
       Contact, and Emergency Contact. The first two have normal
       label/value pairs PLUS a few fields with no visible label at
       all (customer name, a masked ID number, and the address) —
       these are identified by position instead of a label. Emergency
       Contact is fully labeled, no unlabeled fields.
     - "Sales Info" — reached by tapping its tab header. Normal
       label/value pairs, no unlabeled fields.
4. One driver.back() from the detail screen returns to the list.

HOW LABEL/VALUE PAIRING WORKS
---------------------------------
On every screen here, a label sits to the LEFT of its value, and both
share the same vertical position (y-coordinate) — but the raw reading
order in the XML doesn't alternate label-value cleanly, and the exact
pixel positions differ between devices/screen resolutions. So instead
of matching on fixed x-coordinates (which broke the first time this
was tested on a different device — a 720x1520 screen instead of the
1600x2560 one the fields were originally mapped from), the script
groups elements into rows by shared y-coordinate, then within each row
takes the leftmost text as the label and the rightmost as its value.
This holds regardless of screen size. Unlabeled fields (no matching
text in the known label list) are collected separately, in top-to-
bottom order, and assigned fixed names based on where they fall in
that section.

ONE THING WORTH NOTING
--------------------------
The list screen's "Product Name" field (e.g. "CP-XN501HW") is the
product model/SKU. The Sales Info tab's "Product" field (e.g. "XCEL")
is a different, separate product-name value from that section of the
app — both are genuine product names from different parts of the
record, not a mislabeled dealer code.

Start with LIST_LIMIT = 3. Once the CSV looks right, set it to None to
process the entire list.
"""

import csv
import re
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

APP_PACKAGE = "cuckoo.doctress"
# confirmed: the list screen
APP_ACTIVITY = "cuckoo.doctress.naturalcareservicelist"

LIST_LIMIT = None   # was 3 for testing — now processes the whole list
# set True again only if something breaks and you need to see raw element data
DEBUG = False

# --- List screen (confirmed from XML) ---
LABEL_FIELD_MAP = {
    "NS No": "ns_no",
    "Sales No": "sales_no",
    "NS Date": "ns_date",
    "Status": "status",
    "Appt Date": "appt_date",
    "Product Name": "product_name",
    "Cust Name": "cust_name",
}
KEY_FIELD = "ns_no"   # unique per card — avoids re-scraping someone after scrolling

# --- Detail screen: Address/Contact Info tab (confirmed from XML) ---
BILLING_LABEL_MAP = {
    "Doc No.": "billing_doc_no",
    "Sales No.": "billing_sales_no",
    "Contact Person": "billing_contact_person",
    "Tel No (Mobile 1)": "billing_mobile1",
    "Tel No (Mobile 2)": "billing_mobile2",
    "Tel No (Office)": "billing_office_phone",
    "Email": "billing_email",
}
BILLING_ORPHAN_FIELDS = ["billing_customer_name",
                         "billing_ic_number", "billing_address"]

INSTALL_LABEL_MAP = {
    "Contact Person": "install_contact_person",
    "Tel No (Mobile 1)": "install_mobile1",
    "Tel No (Mobile 2)": "install_mobile2",
    "Tel No (House)": "install_house_phone",
    "Tel No (Office)": "install_office_phone",
    "Email": "install_email",
}
INSTALL_ORPHAN_FIELDS = ["install_address"]

EMERGENCY_LABEL_MAP = {
    "Contact Name": "emergency_contact_name",
    "Contact Number": "emergency_contact_number",
    "Relation": "emergency_relation",
}

HEADER_TEXTS = {
    "Billing Address & Contact",
    "Installation/Service Address & Contact",
    "Emergency Contact",
}

# --- Detail screen: Sales Info tab (confirmed from XML) ---
SALES_INFO_LABEL_MAP = {
    "Sales No.": "sales_info_sales_no",
    "Current Stage": "current_stage",
    "Product": "sales_info_product",
    "OutRight Price": "outright_price",
    "Sales Date": "sales_date",
    "Rental Fees (Monthly)": "rental_fee_monthly",
    "Rental Processing Fees": "rental_processing_fee",
    "Application Type": "application_type",
    "Sales Status": "sales_status",
    "PO Number": "po_number",
    "Promo Code": "promo_code",
    "Rental Scheme": "rental_scheme",
}

OUTPUT_CSV = "cuckoo_export.csv"
WAIT_SECONDS = 10
# raised since smaller, controlled scroll steps need more of them to reach the bottom
MAX_SCROLLS = 300
MAX_STAGNANT_ROUNDS = 2
# fallback step, only used when there's no measured position to target yet (e.g. the very first scroll)
SCROLL_STEP_PERCENT = 0.32
SCROLL_REGION_TOP_FRACTION = 0.40     # stays below the pinned filter header
SCROLL_REGION_HEIGHT_FRACTION = 0.48


# ============================================================
# Low-level helpers
# ============================================================

def build_driver():
    options = UiAutomator2Options()
    options.platform_name = "Android"
    options.app_package = APP_PACKAGE
    options.app_activity = APP_ACTIVITY
    options.no_reset = True
    return webdriver.Remote("http://127.0.0.1:4723", options=options)


def wait_for(driver, selector, timeout=WAIT_SECONDS):
    by, value = selector
    return xWebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


def read_text_safe(el):
    try:
        return el.text.strip()
    except Exception:
        return ""


def parse_bounds(bounds_str):
    """'[x1,y1][x2,y2]' -> (x1, y1, x2, y2), or None if unparseable."""
    if not bounds_str:
        return None
    m = re.match(r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", bounds_str)
    if not m:
        return None
    return tuple(int(g) for g in m.groups())


def tap_element(driver, element):
    """
    Taps by screen coordinates instead of calling .click() directly —
    some Appium-Python-Client / Selenium version combinations throw
    "Wrong parameters applied for elementClick" on UiAutomator2. This
    sidesteps that bug entirely.
    """
    parsed = parse_bounds(element.get_attribute("bounds"))
    if not parsed:
        raise RuntimeError(
            "Could not read bounds — can't compute a tap point for this element.")
    x1, y1, x2, y2 = parsed
    cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
    driver.execute_script("mobile: clickGesture", {"x": cx, "y": cy})


def scroll_down(driver, percent=None):
    """
    Uses "mobile: scrollGesture" rather than "mobile: swipeGesture" —
    swipeGesture performs a FLING, which has momentum: Android keeps
    scrolling after the simulated finger lifts, and how far it travels
    depends on gesture velocity in a way that's hard to predict. That
    inconsistency was very likely why scrolling sometimes jumped past
    several customers at once. scrollGesture instead moves a fixed,
    controlled fraction of the given area with no fling.

    `percent`, when given, comes from compute_scroll_percent() — a
    measured amount based on exactly where the last already-processed
    customer's card sits on screen right now, rather than a guessed
    fixed distance. Falls back to SCROLL_STEP_PERCENT only when there's
    no measurement to base it on yet (the very first scroll).

    Note: for scrollGesture, "direction" describes which way the
    CONTENT moves (not the simulated finger) — "down" reveals further/
    later items in the list, which is what swipeGesture called "up".
    """
    if percent is None:
        percent = SCROLL_STEP_PERCENT
    size = driver.get_window_size()
    width, height = size["width"], size["height"]
    driver.execute_script("mobile: scrollGesture", {
        "left": int(width * 0.1),
        "top": int(height * SCROLL_REGION_TOP_FRACTION),
        "width": int(width * 0.8),
        "height": int(height * SCROLL_REGION_HEIGHT_FRACTION),
        "direction": "down",
        "percent": percent,
    })
    time.sleep(0.6)
    dismiss_keyboard_if_present(driver)


def compute_scroll_percent(driver, customers, margin_fraction=0.04):
    """
    Measures exactly where the LAST customer currently on screen sits
    (customers are already top-to-bottom order), and computes the exact
    scroll fraction needed to bring that card's top boundary up near
    the top of the scroll region — i.e. "scroll to the line separating
    this card from the next," rather than a blind fixed distance.

    A small margin is subtracted so the boundary lands just inside the
    visible region instead of exactly on the edge, avoiding the
    partially-rendered-row problem from earlier.
    """
    if not customers:
        return SCROLL_STEP_PERCENT

    size = driver.get_window_size()
    height = size["height"]
    region_top = height * SCROLL_REGION_TOP_FRACTION
    region_height = height * SCROLL_REGION_HEIGHT_FRACTION

    last_card_top_y = customers[-1]["card_top_y"]
    desired_shift = (last_card_top_y - region_top) - \
        (margin_fraction * region_height)
    percent = desired_shift / region_height

    # Guardrails: never scroll by ~nothing (no progress) or overshoot
    # past the measured area (which would defeat the point of measuring).
    return max(0.08, min(0.95, percent))


def dismiss_keyboard_if_present(driver):
    """Defensive safety net — closes the keyboard if anything ever
    accidentally focuses a text field again, so it doesn't silently
    corrupt the next read."""
    try:
        driver.execute_script("mobile: hideKeyboard")
    except Exception:
        pass  # no keyboard was showing, or the command isn't supported — fine either way


def pair_fields(elements, known_labels=None, skip_texts=None, y_tol=8):
    """
    Generic label/value pairing based on RELATIVE position, not fixed
    pixel coordinates — this makes it work across different screen
    resolutions/devices, since it never assumes a label sits at a
    specific x value. Elements are grouped into rows by shared
    y-coordinate; within each row, the first element (left to right)
    whose text matches a KNOWN label is treated as the label, and the
    element immediately after it becomes its value. Anything sitting
    further left than the label (e.g. a small numbered badge next to
    each card) is simply ignored rather than mistaken for the label.

    Returns (row, orphans):
      row     -> {label_text: value_text} for every recognized label
      orphans -> list of value texts (top-to-bottom order) for rows
                 that don't match a known label — i.e. unlabeled
                 fields like a bare address or name with no caption.
    """
    skip_texts = skip_texts or set()
    entries = []
    for el in elements:
        parsed = parse_bounds(el.get_attribute("bounds"))
        if not parsed:
            continue
        x1, y1, _, _ = parsed
        text = read_text_safe(el)
        if text in skip_texts:
            continue
        entries.append((y1, x1, text))

    entries.sort(key=lambda e: (e[0], e[1]))

    rows = []
    for entry in entries:
        placed = False
        for row in rows:
            if abs(row[0][0] - entry[0]) <= y_tol:
                row.append(entry)
                placed = True
                break
        if not placed:
            rows.append([entry])

    row_dict = {}
    orphans = []
    for row in rows:
        row_sorted = sorted(row, key=lambda e: e[1])  # left to right
        row_y = row_sorted[0][0]

        if known_labels is not None:
            # Find the first element in this row whose text is a genuine
            # known label — this way something sitting further left (like
            # a small numbered badge next to each card) gets ignored
            # instead of being mistaken for the label itself.
            label_idx = next((i for i, e in enumerate(
                row_sorted) if e[2] in known_labels), None)
            if label_idx is not None:
                label_text = row_sorted[label_idx][2]
                value_text = row_sorted[label_idx +
                                        1][2] if label_idx + 1 < len(row_sorted) else ""
                row_dict[label_text] = value_text
                # anything before label_idx (e.g. a badge number) is just ignored
            else:
                for _, _, text in row_sorted:
                    orphans.append((row_y, text))
        else:
            if len(row_sorted) >= 2:
                row_dict[row_sorted[0][2]] = row_sorted[-1][2]
            else:
                orphans.append((row_y, row_sorted[0][2]))

    orphans.sort(key=lambda t: t[0])
    return row_dict, [t for _, t in orphans]


# ============================================================
# List screen
# ============================================================

def group_into_rows(entries, y_tol=8):
    """entries: list of (y1, x1, text). Returns rows: list of rows,
    each row a list of (y1, x1, text) sorted left-to-right, rows
    themselves sorted top-to-bottom."""
    entries = sorted(entries, key=lambda e: (e[0], e[1]))
    rows = []
    for e in entries:
        placed = False
        for row in rows:
            if abs(row[0][0] - e[0]) <= y_tol:
                row.append(e)
                placed = True
                break
        if not placed:
            rows.append([e])
    rows = [sorted(r, key=lambda e: e[1]) for r in rows]
    rows.sort(key=lambda r: r[0][0])
    return rows


def rows_to_fields(rows_chunk, known_labels):
    """Same label-search-per-row logic as pair_fields, applied to an
    already-sliced chunk of rows belonging to one customer."""
    result = {}
    for row in rows_chunk:
        label_idx = next((i for i, e in enumerate(
            row) if e[2] in known_labels), None)
        if label_idx is not None:
            label_text = row[label_idx][2]
            value_text = row[label_idx + 1][2] if label_idx + \
                1 < len(row) else ""
            result[label_text] = value_text
    return result


def get_visible_customers(driver):
    """
    Reads EVERY TextView and Button on screen in one global query (no
    per-card scoped searches — those don't reliably scope on this
    Appium/UiAutomator2 setup, which is what broke the earlier version),
    then figures out which elements belong to which customer purely by
    on-screen position: each "NS No" row marks where a new card starts.

    Returns a list of {"row": {...parsed fields...}, "button": element_or_None}
    for every customer currently visible.
    """
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="NS No"]'))

    text_elements = driver.find_elements(
        AppiumBy.CLASS_NAME, "android.widget.TextView")
    entries = []
    for el in text_elements:
        parsed = parse_bounds(el.get_attribute("bounds"))
        if not parsed:
            continue
        x1, y1, _, _ = parsed
        entries.append((y1, x1, read_text_safe(el)))
    rows = group_into_rows(entries)

    marker_indices = [i for i, row in enumerate(
        rows) if any(t == "NS No" for _, _, t in row)]
    if DEBUG:
        print(
            f"  [debug] found {len(rows)} row(s) total, {len(marker_indices)} 'NS No' marker(s)")

    buttons = driver.find_elements(AppiumBy.XPATH, '//android.widget.Button')
    button_positions = []
    for b in buttons:
        parsed = parse_bounds(b.get_attribute("bounds"))
        if parsed:
            button_positions.append((parsed[1], b))
    button_positions.sort(key=lambda t: t[0])

    known_labels = set(LABEL_FIELD_MAP.keys())
    customers = []
    for idx, start in enumerate(marker_indices):
        end = marker_indices[idx + 1] if idx + \
            1 < len(marker_indices) else len(rows)
        chunk = rows[start:end]
        card_top_y = chunk[0][0][0]
        next_top_y = rows[marker_indices[idx + 1]][0][0] if idx + \
            1 < len(marker_indices) else float("inf")

        field_dict = rows_to_fields(chunk, known_labels)
        parsed_row = {LABEL_FIELD_MAP[k]: v for k,
                      v in field_dict.items() if k in LABEL_FIELD_MAP}

        matching_button = next(
            (b_el for b_y, b_el in button_positions if card_top_y <= b_y < next_top_y), None)
        customers.append(
            {"row": parsed_row, "button": matching_button, "card_top_y": card_top_y})

    if DEBUG:
        for c in customers:
            print(
                f"  [debug] parsed customer: {c['row']}  (button found: {c['button'] is not None})")

    return customers


def get_visible_customers_stable(driver, max_attempts=4, settle_delay=0.4):
    """
    Reads the screen repeatedly until two consecutive reads agree on
    which NS numbers are visible. A single read can catch the view
    mid-render — Android list rows get REUSED as you scroll, so reading
    too early can show a row still holding the PREVIOUS customer's name
    while its NS No has already updated to the new one. That's what was
    causing blank names and, worse, a name attached to the wrong NS
    number. Waiting for two matching reads in a row avoids trusting a
    transitional, half-updated state.
    """
    prev_signature = None
    customers = []
    for _ in range(max_attempts):
        customers = get_visible_customers(driver)
        signature = tuple(c["row"].get(KEY_FIELD, "") for c in customers)
        if signature == prev_signature and signature:
            return customers
        prev_signature = signature
        time.sleep(settle_delay)
    return customers


# ============================================================
# Popup menu ("Choose Option") -> detail screen
# ============================================================

def select_popup_option(driver, option_text, timeout=WAIT_SECONDS):
    # normalize-space() handles the extra leading spaces the app puts in these labels
    xpath = f'//android.widget.TextView[normalize-space(@text)="{option_text}"]'
    el = wait_for(driver, (AppiumBy.XPATH, xpath), timeout=timeout)
    tap_element(driver, el)


# ============================================================
# Detail screen: Address/Contact Info tab
# ============================================================

def get_billing_elements(driver):
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="Billing Address & Contact"]'))
    return driver.find_elements(
        AppiumBy.XPATH,
        '//android.widget.TextView[@text="Billing Address & Contact"]'
        '/following-sibling::android.view.ViewGroup[1]//android.widget.TextView'
    )


def get_installation_elements(driver):
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="Installation/Service Address & Contact"]'))
    return driver.find_elements(
        AppiumBy.XPATH,
        '//android.widget.TextView[@text="Installation/Service Address & Contact"]'
        '/parent::android.view.ViewGroup//android.widget.TextView'
    )


def get_emergency_elements(driver):
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="Emergency Contact"]'))
    return driver.find_elements(
        AppiumBy.XPATH,
        '//android.widget.TextView[@text="Emergency Contact"]'
        '/parent::android.view.ViewGroup//android.widget.TextView'
    )


def get_sales_info_elements(driver):
    wait_for(driver, (AppiumBy.XPATH,
             '//android.widget.TextView[@text="Current Stage"]'), timeout=8)
    return driver.find_elements(AppiumBy.XPATH, '//android.widget.TextView')


def read_full_detail(driver):
    record = {}

    # --- Address/Contact Info tab (shown by default) ---
    billing_row, billing_orphans = pair_fields(
        get_billing_elements(driver),
        known_labels=set(BILLING_LABEL_MAP.keys()), skip_texts=HEADER_TEXTS
    )
    for label_text, field_name in BILLING_LABEL_MAP.items():
        record[field_name] = billing_row.get(label_text, "")
    for i, field_name in enumerate(BILLING_ORPHAN_FIELDS):
        record[field_name] = billing_orphans[i] if i < len(
            billing_orphans) else ""

    install_row, install_orphans = pair_fields(
        get_installation_elements(driver),
        known_labels=set(INSTALL_LABEL_MAP.keys()), skip_texts=HEADER_TEXTS
    )
    for label_text, field_name in INSTALL_LABEL_MAP.items():
        record[field_name] = install_row.get(label_text, "")
    for i, field_name in enumerate(INSTALL_ORPHAN_FIELDS):
        record[field_name] = install_orphans[i] if i < len(
            install_orphans) else ""

    emergency_row, _ = pair_fields(
        get_emergency_elements(driver),
        known_labels=set(EMERGENCY_LABEL_MAP.keys()), skip_texts=HEADER_TEXTS
    )
    for label_text, field_name in EMERGENCY_LABEL_MAP.items():
        record[field_name] = emergency_row.get(label_text, "")

    # --- Switch to Sales Info tab ---
    sales_tab = wait_for(
        driver, (AppiumBy.XPATH, '//android.widget.TextView[@text="Sales Info"]'))
    tap_element(driver, sales_tab)
    # get_sales_info_elements() below already waits/polls for "Current Stage"
    # to appear, so no fixed sleep needed here.

    sales_row, _ = pair_fields(
        get_sales_info_elements(driver),
        known_labels=set(SALES_INFO_LABEL_MAP.keys())
    )
    for label_text, field_name in SALES_INFO_LABEL_MAP.items():
        record[field_name] = sales_row.get(label_text, "")

    return record


# ============================================================
# Main loop: scroll + scrape until nothing new appears
# ============================================================

def run():
    driver = build_driver()
    time.sleep(3)

    all_records = []
    seen_keys = set()
    stagnant_rounds = 0
    scroll_count = 0

    try:
        while True:
            if LIST_LIMIT and len(all_records) >= LIST_LIMIT:
                print(f"Reached LIST_LIMIT of {LIST_LIMIT} — stopping.")
                break

            customers = get_visible_customers_stable(driver)

            next_customer = None
            for c in customers:
                key = c["row"].get(KEY_FIELD)
                if not key or key in seen_keys:
                    continue
                if c["button"] is None:
                    # Card is likely only partially scrolled into view (its
                    # NS No label rendered, but the button below it hasn't
                    # yet). Don't mark it seen — leave it to be retried once
                    # it's fully visible, instead of losing it permanently.
                    continue
                next_customer = c
                break

            if next_customer is None:
                stagnant_rounds += 1
                if stagnant_rounds >= MAX_STAGNANT_ROUNDS:
                    print(
                        "No new customers found after scrolling — reached the end of the list.")
                    break
                scroll_count += 1
                if scroll_count >= MAX_SCROLLS:
                    print(
                        "Hit the safety scroll limit — stopping to avoid an infinite loop.")
                    break
                target_percent = compute_scroll_percent(driver, customers)
                if DEBUG:
                    print(
                        f"  [debug] scrolling by measured percent={target_percent:.3f}")
                scroll_down(driver, percent=target_percent)
                continue

            stagnant_rounds = 0
            next_row = next_customer["row"]
            key = next_row[KEY_FIELD]
            seen_keys.add(key)
            print(
                f"[{len(all_records) + 1}] Opening: {key} ({next_row.get('cust_name', '')})")

            try:
                button = next_customer["button"]
                if button is None:
                    raise RuntimeError(
                        "No 'View Order' button found near this customer's row")

                try:
                    tap_element(driver, button)
                except Exception as e:
                    raise RuntimeError(
                        f"[stage: tapping row button] {type(e).__name__}: {e}")

                try:
                    # select_popup_option() already waits/polls for the popup
                    # to appear — no fixed sleep needed before it.
                    select_popup_option(driver, "View Order")
                except Exception as e:
                    raise RuntimeError(
                        f"[stage: selecting 'View Order' from popup] {type(e).__name__}: {e}")

                try:
                    wait_for(
                        driver, (AppiumBy.XPATH, '//android.widget.Button[@text="Customer Information"]'))
                except Exception as e:
                    raise RuntimeError(
                        f"[stage: waiting for Customer Information screen] {type(e).__name__}: {e}")

                try:
                    detail = read_full_detail(driver)
                except Exception as e:
                    raise RuntimeError(
                        f"[stage: reading detail screen fields] {type(e).__name__}: {e}")

                all_records.append({**next_row, **detail})
            except Exception as e:
                print(f"  !! Skipped {key} due to error: {e}")
            finally:
                driver.back()
                # get_visible_customers() at the top of the next loop already
                # waits/polls for "NS No" to reappear, so no sleep needed here.

    finally:
        driver.quit()

    write_csv(all_records)
    print(f"Done. Wrote {len(all_records)} records to {OUTPUT_CSV}")


def write_csv(records):
    if not records:
        print("No records captured — nothing written.")
        return
    fieldnames = sorted({key for r in records for key in r.keys()})

    # Excel treats any CSV cell starting with +, -, or = as the start of a
    # formula (e.g. "+60-127156860" gets evaluated as 60 minus 127156860).
    # A leading apostrophe tells Excel "this is literal text" and is not
    # shown in the cell — this keeps phone numbers displaying correctly.
    def excel_safe(value):
        if isinstance(value, str) and value[:1] in ("+", "-", "="):
            return "'" + value
        return value

    safe_records = [{k: excel_safe(v) for k, v in r.items()} for r in records]

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(safe_records)


if __name__ == "__main__":
    run()
