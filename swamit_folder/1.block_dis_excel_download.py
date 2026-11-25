import re
import os
import sys
import json
from playwright.sync_api import sync_playwright

# Ensure UTF-8 output encoding for Windows terminal
sys.stdout.reconfigure(encoding='utf-8')


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_DIR = os.path.join(BASE_DIR, 'file')
os.makedirs(FILE_DIR, exist_ok=True)


def download_handle_district(download):
    print('download started')
    district_file = os.path.join(FILE_DIR, 'district.xls')
    district_old_file = os.path.join(FILE_DIR, 'district_old.xls')
    if os.path.exists(district_file):
        if os.path.exists(district_old_file):
            os.remove(district_old_file)
        os.rename(district_file, district_old_file)
    download.save_as(district_file)
    print('download completed District')


def download_handle_block(download):
    print('download started')
    block_file = os.path.join(FILE_DIR, 'block.xls')
    block_old_file = os.path.join(FILE_DIR, 'block_old.xls')
    if os.path.exists(block_file):
        if os.path.exists(block_old_file):
            os.remove(block_old_file)
        os.rename(block_file, block_old_file)
    download.save_as(block_file)
    print('download completed Block')


def download_svamitva_excel(headless=False):
    """Downloads district and block-level Svamitva Excel reports and fetches the last refreshed time.

    Launches a Chromium browser via Playwright to navigate the MPIGR Svamitva dashboard,
    downloads the district-level report ('./file/district.xls'), opens the Jabalpur
    block-level view to download the block report ('./file/block.xls'), and extracts
    the dashboard's last refreshed timestamp with minutes rounded to ':00'.
    Existing report files are rotated to '_old.xls' before saving new ones.

    Args:
        headless (bool, optional): If True, runs the browser in headless mode without
            a visible UI. Defaults to False.

    Returns:
        str | None: The formatted last refreshed timestamp (e.g., '23 Nov 2026, 05:00 PM IST'),
            or None if an error occurs during execution.
    """

    
    os.makedirs('file', exist_ok=True)
    full_text_hour = None

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()

        try:
            page.goto('https://services.mpigr.gov.in/svamitva_dashboard/district-registration', timeout=30000)
            page.wait_for_load_state('networkidle', timeout=15000)

            # 1. Download District level report
            excel_download = page.wait_for_selector("button.btn-download-excel", timeout=10000)
            with page.expect_download() as download_info:
                excel_download.click(force=True)
            download_handle_district(download_info.value)
            page.wait_for_load_state('networkidle', timeout=15000)

            # 2. Click district row to open block-level view
            page.wait_for_selector("//a[contains(text(),'जबलपुर')]").click()
            page.wait_for_load_state('networkidle', timeout=15000)

            # 3. Download Block level report
            block_download = page.wait_for_selector("button.btn-download-excel", timeout=30000)
            with page.expect_download() as download_info:
                block_download.click(force=True)
            download_handle_block(download_info.value)

            # 4. Get last refreshed timestamp
            page.goto("https://services.mpigr.gov.in/svamitva_dashboard/", wait_until="networkidle")
            refreshed_locator = page.locator("span.navbar-refreshed-value")
            refreshed_locator.wait_for(state="visible")

            full_text = refreshed_locator.inner_text().strip()
            full_text_hour = re.sub(r":\d{2}", ":00", full_text)
            print("Full text:", full_text_hour)

            # Persist and rotate refreshed timestamp in refresh_time.json
            time_file = os.path.join(FILE_DIR, 'refresh_time.json')
            old_time = None
            if os.path.exists(time_file):
                try:
                    with open(time_file, 'r', encoding='utf-8') as f:
                        prev_data = json.load(f)
                        old_time = prev_data.get('current_time')
                except Exception:
                    pass

            # Fallback to the modification time of the previous downloaded file
            if not old_time:
                district_old = os.path.join(FILE_DIR, 'district_old.xls')
                if os.path.exists(district_old):
                    from datetime import datetime
                    dt = datetime.fromtimestamp(os.path.getmtime(district_old))
                    old_time = dt.strftime("%d %b %Y, %I:00 %p IST")

            time_data = {
                'current_time': full_text_hour,
                'old_time': old_time
            }
            try:
                with open(time_file, 'w', encoding='utf-8') as f:
                    json.dump(time_data, f, indent=2)
                print(f"Saved refreshed time to {time_file}")
            except Exception as e:
                print(f"Warning: could not save {time_file}: {e}")

        except Exception as e:
            print(f"\n[X] Error during execution: {e}")
            print("\nPossible reasons: website server error, network connectivity issue, or element not found.")

        finally:
            browser.close()

    return full_text_hour


if __name__ == "__main__":
    refreshed_time = download_svamitva_excel()
    print("Returned value:", refreshed_time)