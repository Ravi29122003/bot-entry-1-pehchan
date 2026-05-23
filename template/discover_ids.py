"""
Discover all form element IDs on the Pehchan portal.
Run this once to get the exact element IDs we need.
"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import time
import getpass
from playwright.sync_api import sync_playwright

BASE_URL = "https://pehchan.rajasthan.gov.in"
USERNAME = "JPRU00001_OPR"

def main():
    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=False, slow_mo=100)
    page = browser.new_page()
    page.set_default_timeout(15000)
    page.set_default_navigation_timeout(60000)

    # Login
    for ver in ["pehchan7", "pehchan3", "pehchan2"]:
        try:
            print(f"Trying {ver}...")
            page.goto(f"{BASE_URL}/{ver}/Mainpage.aspx", timeout=15000)
            page.wait_for_load_state("networkidle", timeout=15000)
            if "Server Error" not in page.content():
                break
        except Exception:
            continue

    page.locator("text=लॉगिन").or_(page.locator("text=Login")).first.click()
    page.wait_for_load_state("networkidle")
    portal_base = page.url.split("/Admin")[0]

    page.locator("input[type='text']").first.fill(USERNAME)
    password = getpass.getpass("Password: ")
    page.locator("input[type='password']").first.fill(password)
    captcha = input("CAPTCHA: ").strip()
    page.locator("input[type='text']").nth(1).fill(captcha)
    page.locator("text=लॉगिन करे").first.click()

    for i in range(10):
        time.sleep(2)
        if "Home" in page.url:
            break

    if "Home" not in page.url:
        input("Press Enter after login completes...")

    portal_base = page.url.split("/Admin")[0]
    print(f"Portal: {portal_base}")

    # Dismiss popup
    time.sleep(2)
    for sel in ["button.close", ".close", "text=×"]:
        try:
            el = page.locator(sel).first
            if el.is_visible(timeout=3000):
                el.click()
                break
        except Exception:
            continue

    # Navigate to form
    page.evaluate(f"window.location.href = '{portal_base}/Admin/frmBirthEntryOld.aspx'")
    time.sleep(5)
    page.wait_for_load_state("networkidle", timeout=30000)

    # Fill pre-form to load main form
    page.locator("input[type='text']").first.fill("9999")
    page.locator("input[type='text']").nth(1).fill("1972")
    page.locator("select").first.select_option(index=1)
    time.sleep(1)
    page.locator("text=प्रवेश करे").first.click()
    page.wait_for_load_state("networkidle", timeout=30000)
    time.sleep(3)

    # NOW DUMP ALL FORM ELEMENTS
    print("\n" + "=" * 80)
    print("ALL INPUT ELEMENTS ON THE FORM")
    print("=" * 80)

    elements = page.evaluate("""() => {
        const results = [];
        // Text inputs
        document.querySelectorAll('input[type="text"]').forEach(el => {
            const label = el.closest('tr') ? el.closest('tr').innerText.substring(0, 60) : '';
            results.push({type: 'text', id: el.id, name: el.name, value: el.value, label: label.replace(/\\n/g, ' | ')});
        });
        // Radio buttons
        document.querySelectorAll('input[type="radio"]').forEach(el => {
            const label = el.closest('tr') ? el.closest('tr').innerText.substring(0, 60) : '';
            results.push({type: 'radio', id: el.id, name: el.name, value: el.value, checked: el.checked, label: label.replace(/\\n/g, ' | ')});
        });
        // Selects
        document.querySelectorAll('select').forEach(el => {
            const label = el.closest('tr') ? el.closest('tr').innerText.substring(0, 60) : '';
            const options = Array.from(el.options).map(o => o.text).slice(0, 5);
            results.push({type: 'select', id: el.id, name: el.name, options: options, label: label.replace(/\\n/g, ' | ')});
        });
        // Textareas
        document.querySelectorAll('textarea').forEach(el => {
            const label = el.closest('tr') ? el.closest('tr').innerText.substring(0, 60) : '';
            results.push({type: 'textarea', id: el.id, name: el.name, label: label.replace(/\\n/g, ' | ')});
        });
        return results;
    }""")

    for el in elements:
        print(f"\n  [{el['type'].upper()}] id={el['id']}")
        print(f"    name={el.get('name', '')}")
        if el['type'] == 'select':
            print(f"    options={el.get('options', [])}")
        if el['type'] == 'radio':
            print(f"    value={el.get('value', '')} checked={el.get('checked', False)}")
        if el['type'] == 'text':
            print(f"    value={el.get('value', '')}")
        print(f"    label={el.get('label', '')[:80]}")

    print("\n\nDone. Close browser manually.")
    input("Press Enter to exit...")
    browser.close()
    pw.stop()

if __name__ == "__main__":
    main()
