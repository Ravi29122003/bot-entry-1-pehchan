"""Discover element IDs for late registration fields (फीस, radio buttons)"""
import sys, io, time, getpass
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from playwright.sync_api import sync_playwright

pw = sync_playwright().start()
br = pw.chromium.launch(headless=False, slow_mo=100)
pg = br.new_page()
pg.set_default_timeout(15000)
pg.set_default_navigation_timeout(60000)

for v in ["pehchan7", "pehchan3", "pehchan2"]:
    try:
        pg.goto(f"https://pehchan.rajasthan.gov.in/{v}/Mainpage.aspx", timeout=15000)
        pg.wait_for_load_state("networkidle", timeout=15000)
        if "Server Error" not in pg.content():
            break
    except:
        continue

pg.locator("text=Login").or_(pg.locator("text=लॉगिन")).first.click()
pg.wait_for_load_state("networkidle")
pg.locator("input[type=text]").first.fill("JPRU00001_OPR")
pwd = getpass.getpass("Password: ")
pg.locator("input[type=password]").first.fill(pwd)
cap = input("CAPTCHA: ")
pg.locator("input[type=text]").nth(1).fill(cap)
pg.locator("text=लॉगिन करे").first.click()

for i in range(10):
    time.sleep(2)
    if "Home" in pg.url:
        break
if "Home" not in pg.url:
    input("Press Enter after login...")

base = pg.url.split("/Admin")[0]
time.sleep(2)
for s in ["button.close", ".close", "text=×"]:
    try:
        e = pg.locator(s).first
        if e.is_visible(timeout=3000):
            e.click()
            break
    except:
        continue

# Navigate to form
pg.evaluate(f"window.location.href='{base}/Admin/frmBirthEntryOld.aspx'")
time.sleep(5)
pg.wait_for_load_state("networkidle", timeout=30000)

# Use reg 400/1988 - this is Row 5 which has late registration fields
pg.locator("input[type=text]").first.fill("400")
pg.locator("input[type=text]").nth(1).fill("1988")
pg.locator("select").first.select_option(index=1)
time.sleep(1)
pg.locator("text=प्रवेश करे").first.click()
pg.wait_for_load_state("networkidle", timeout=30000)
time.sleep(3)

# Now fill birth date (04/12/1987) and reg date (06/01/1988) to trigger late reg fields
pg.evaluate("""() => {
    const birth = document.getElementById('ContentPlaceHolder1_txtbirth');
    birth.value = '04/12/1987';
    birth.dispatchEvent(new Event('change', {bubbles: true}));
}""")
time.sleep(5)

pg.evaluate("""() => {
    const reg = document.getElementById('ContentPlaceHolder1_txtregisdt');
    reg.value = '06/01/1988';
    reg.dispatchEvent(new Event('change', {bubbles: true}));
}""")
time.sleep(5)

# Now dump ALL elements - especially looking for fee field and new radio buttons
print("\n" + "=" * 80)
print("ALL VISIBLE ELEMENTS AFTER LATE REGISTRATION DATES")
print("=" * 80)

elements = pg.evaluate("""() => {
    const results = [];
    document.querySelectorAll('input[type="text"], textarea').forEach(el => {
        if (el.offsetParent !== null) {
            const tr = el.closest('tr');
            const label = tr ? tr.innerText.substring(0, 80).replace(/\\n/g, ' | ') : '';
            results.push({type: 'TEXT', id: el.id, value: el.value, label: label});
        }
    });
    document.querySelectorAll('input[type="radio"]').forEach(el => {
        if (el.offsetParent !== null) {
            const tr = el.closest('tr');
            const label = tr ? tr.innerText.substring(0, 80).replace(/\\n/g, ' | ') : '';
            results.push({type: 'RADIO', id: el.id, value: el.value, checked: el.checked, label: label});
        }
    });
    document.querySelectorAll('select').forEach(el => {
        if (el.offsetParent !== null) {
            const tr = el.closest('tr');
            const label = tr ? tr.innerText.substring(0, 80).replace(/\\n/g, ' | ') : '';
            const opts = Array.from(el.options).slice(0, 5).map(o => o.text);
            results.push({type: 'SELECT', id: el.id, options: opts, label: label});
        }
    });
    return results;
}""")

for el in elements:
    if el['type'] == 'RADIO':
        print(f"\n  [{el['type']}] id={el['id']}  value={el['value']}  checked={el.get('checked', False)}")
    elif el['type'] == 'SELECT':
        print(f"\n  [{el['type']}] id={el['id']}  options={el.get('options', [])}")
    else:
        print(f"\n  [{el['type']}] id={el['id']}  value=[{el['value']}]")
    print(f"    label: {el.get('label', '')[:100]}")

input("\nPress Enter to close...")
br.close()
pw.stop()
