import re
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto("https://services.mpigr.gov.in/svamitva_dashboard/", wait_until="networkidle")

    refreshed_locator = page.locator("span.navbar-refreshed-value")
    refreshed_locator.wait_for(state="visible")
    
    full_text = refreshed_locator.inner_text().strip()
    
    # Replace minutes with :00
    full_text_hour = re.sub(r":\d{2}", ":00", full_text)
    print("Full text:", full_text_hour)  # Output: 23 Nov 2026, 05:00 PM IST

    # Time only
    time_only = full_text_hour.split(", ")[-1]
    print("Time only:", time_only)       # Output: 05:00 PM IST

    browser.close()
