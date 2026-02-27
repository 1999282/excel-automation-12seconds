from playwright.sync_api import sync_playwright
import time
import sys

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    page.on("console", lambda msg: print(f"Browser console: {msg.text}"))
    page.on("pageerror", lambda err: print(f"Browser error: {err}"))

    print("Navigating to production server...")
    try:
        page.goto("https://1999282.github.io/excel-automation-12seconds/web/", timeout=20000)
    except Exception as e:
        print(f"Failed to connect to production server: {e}")
        sys.exit(1)

    print("Uploading 50,000 row enterprise dataset...")
    # Target the file input directly
    page.set_input_files("#file-input", "c:/Users/deepa/OneDrive - Hult Students/Desktop/Agent Google Antigravity/excel-automation-12seconds/complex_enterprise_data.csv")

    print("Waiting for Web Worker to process the data. This may take up to 2 minutes due to the 50,000 rows...")
    
    # Wait for the results section to become visible
    try:
        page.wait_for_selector("#results-section", state="visible", timeout=120000) # Give it 120 seconds
        print("Processing complete! Application did not freeze.")
    except Exception as e:
        print("Timeout or error waiting for results:", e)
        page.screenshot(path="stress_test_error.png", full_page=True)
        sys.exit(1)

    print("Taking a screenshot of the processed 50k row dashboard...")
    # Let charts render
    time.sleep(2)
    
    # Take full page screenshot
    page.screenshot(path="stress_test_50k_success.png", full_page=True)
    print("Test passed! Screenshot saved as stress_test_50k_success.png")

    # Get some stats to verify
    try:
        rows_processed = page.locator("#totalRows").inner_text()
        print(f"Verified UI shows: {rows_processed}")
    except:
        pass

    browser.close()

with sync_playwright() as playwright:
    run(playwright)
