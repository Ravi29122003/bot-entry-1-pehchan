import sys, io, time, getpass
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from playwright.sync_api import sync_playwright
pw = sync_playwright().start()
br = pw.chromium.launch(headless=False, slow_mo=100)
pg = br.new_page()
pg.set_default_timeout(15000)
pg.set_default_navigation_timeout(60000)
for v in ["pehchan7","pehchan3","pehchan2"]:
    try:
        pg.goto(f"https://pehchan.rajasthan.gov.in/{v}/Mainpage.aspx", timeout=15000)
        pg.wait_for_load_state("networkidle", timeout=15000)
        if "Server Error" not in pg.content(): break
    except: continue
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
    if "Home" in pg.url: break
if "Home" not in pg.url: input("Press Enter...")
base = pg.url.split("/Admin")[0]
time.sleep(2)
for s in ["button.close", ".close", "text=×"]:
    try:
        e = pg.locator(s).first
        if e.is_visible(timeout=3000): e.click(); break
    except: continue
pg.evaluate(f"window.location.href='{base}/Admin/frmBirthEntryOld.aspx'")
time.sleep(5)
pg.wait_for_load_state("networkidle", timeout=30000)
pg.locator("input[type=text]").first.fill("9998")
pg.locator("input[type=text]").nth(1).fill("1972")
pg.locator("select").first.select_option(index=1)
time.sleep(1)
pg.locator("text=प्रवेश करे").first.click()
pg.wait_for_load_state("networkidle", timeout=30000)
time.sleep(3)
pg.locator("#ContentPlaceHolder1_rdinstitanal_1").click()
time.sleep(5)
ids = pg.evaluate("""() => {
    const r = [];
    document.querySelectorAll('textarea, input[type=text]').forEach(el => {
        if (el.offsetParent !== null) r.push(el.id + ' = [' + el.value + ']');
    });
    return r;
}""")
print("\n=== ALL VISIBLE FIELDS AFTER GAIR SANSTHAGAT ===")
for i in ids: print(i)
input("Press Enter to close...")
br.close()
pw.stop()
