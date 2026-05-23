"""
Pehchan Portal Bot — Browser automation using Playwright.

Flow per record:
  1. Pre-form: reg number + year + registrar dropdown -> pravesh kare
  2. Main form: fill all 23 sections
  3. Submit: indraaj kare -> handle JS alert -> verify success -> punah shuru kare
"""

import time
import os
from datetime import datetime
from playwright.sync_api import sync_playwright, Page, TimeoutError as PwTimeout
from config.settings import (
    BASE_URL, USERNAME,
    PRE_FORM_REGISTRAR_INDEX, REGISTRATION_YEAR,
    GENDER_MAP, FIXED,
    PAGE_LOAD_TIMEOUT, ELEMENT_TIMEOUT, POST_CLICK_DELAY,
    POST_DATE_DELAY, TYPING_DELAY, BETWEEN_RECORDS_DELAY,
    POST_DROPDOWN_DELAY, SCREENSHOT_DIR,
)


class PehchanBot:

    def __init__(self, headless: bool = False):
        self.headless = headless
        self.playwright = None
        self.browser = None
        self.page = None
        self.is_logged_in = False
        self.records_this_session = 0
        self.portal_base = None
        os.makedirs(SCREENSHOT_DIR, exist_ok=True)

    def start(self):
        print("\nStarting browser...")
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=self.headless,
            slow_mo=100,
            args=["--disable-blink-features=AutomationControlled"],
        )
        self.page = self.browser.new_page()
        self.page.set_default_timeout(ELEMENT_TIMEOUT)
        self.page.set_default_navigation_timeout(PAGE_LOAD_TIMEOUT)
        print("  Browser ready")

    def stop(self):
        print("\nShutting down...")
        try:
            if self.browser:
                self.browser.close()
        except Exception:
            pass
        try:
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass
        self.browser = None
        self.playwright = None
        print("  Done")

    def screenshot(self, name: str) -> str:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(SCREENSHOT_DIR, f"{ts}_{name}.png")
        try:
            self.page.screenshot(path=path, full_page=True, timeout=60000)
        except Exception:
            try:
                self.page.screenshot(path=path, full_page=False, timeout=30000)
            except Exception as e2:
                print(f"  Screenshot failed ({name}): {e2}")
                return ""
        return path

    # ---- LOGIN ----

    def login(self, password: str) -> bool:
        print("\nLogging in...")
        # Try multiple portal versions - NIC rotates between them
        for version in ["pehchan7", "pehchan3", "pehchan2"]:
            try:
                print(f"  Trying {version}...")
                self.page.goto(f"{BASE_URL}/{version}/Mainpage.aspx", timeout=15000)
                self._wait()
                if "Server Error" not in self.page.content():
                    break
            except Exception:
                continue
        self.page.locator("text=लॉगिन").or_(self.page.locator("text=Login")).first.click()
        self._wait()
        self.portal_base = self.page.url.split("/Admin")[0]
        print(f"  Portal: {self.portal_base}")

        self.page.locator("input[type='text']").first.fill(USERNAME)
        print(f"  Username: {USERNAME}")

        self.page.locator("input[type='password']").first.fill(password)
        print(f"  Password: ********")

        print("\n  === LOOK AT BROWSER - READ CAPTCHA ===")
        captcha = input("  Type CAPTCHA here -> ").strip()
        self.page.locator("input[type='text']").nth(1).fill(captcha)
        print(f"  CAPTCHA: {captcha}")

        self.page.locator("text=लॉगिन करे").first.click()
        self._wait()
        # Wait and keep checking for redirect to Home page
        for i in range(10):
            time.sleep(2)
            if "Home" in self.page.url:
                break

        self.portal_base = self.page.url.split("/Admin")[0]
        print(f"  Portal: {self.portal_base}")

        if "Home" in self.page.url:
            self.is_logged_in = True
            self.records_this_session = 0
            print("  Login successful!")
            return True

        print(f"  Login may have failed. URL: {self.page.url}")
        self.screenshot("login_check")
        input("  Press Enter after verifying in browser... ")
        self.portal_base = self.page.url.split("/Admin")[0]
        if "Home" in self.page.url:
            self.is_logged_in = True
            print("  Login confirmed after manual check")
            return True
        return False

    # ---- POST-LOGIN ----

    def post_login_setup(self):
        print("\nPost-login setup...")
        time.sleep(2)
        for selector in ["button.close", ".close", "text=×", "[aria-label='Close']"]:
            try:
                el = self.page.locator(selector).first
                if el.is_visible(timeout=3000):
                    el.click()
                    print("  Popup dismissed")
                    time.sleep(1)
                    break
            except Exception:
                continue
        if "/Admin/" in self.page.url:
            self.portal_base = self.page.url.split("/Admin")[0]
        print("  Registrar: using default (Greater Jaipur)")

    # ---- NAVIGATE ----

    def navigate_to_entry_form(self) -> bool:
        print("\nNavigating to Legacy Birth Entry...")
        if "/Admin/" in self.page.url:
            self.portal_base = self.page.url.split("/Admin")[0]

        # Method 1: Click through menu using JavaScript to force visibility
        try:
            self.page.evaluate("""() => {
                const links = document.querySelectorAll('a');
                for (const a of links) {
                    if (a.textContent.includes('लिगेसी जन्म') || a.textContent.includes('Legacy Birth')) {
                        a.click();
                        return true;
                    }
                }
                return false;
            }""")
            time.sleep(3)
            self._wait()
            if "frmBirthEntryOld" in self.page.url:
                print("  On entry form (JS click)")
                return True
        except Exception:
            pass

        # Method 2: Hover menu then click
        try:
            entry_menu = self.page.locator("a").filter(has_text="इन्द्राज").or_(self.page.locator("a").filter(has_text="Entry"))
            entry_menu.first.hover()
            time.sleep(2)
            reg_menu = self.page.locator("a").filter(has_text="पंजीकरण").or_(self.page.locator("a").filter(has_text="Registration"))
            reg_menu.first.hover()
            time.sleep(2)
            legacy = self.page.locator("a").filter(has_text="लिगेसी जन्म").or_(self.page.locator("a").filter(has_text="Legacy Birth"))
            legacy.first.click()
            self._wait()
            if "frmBirthEntryOld" in self.page.url:
                print("  On entry form (menu hover)")
                return True
        except Exception:
            pass

        # Method 3: Navigate using window.location instead of page.goto
        print("  Trying JS navigation...")
        try:
            self.page.evaluate(f"window.location.href = '{self.portal_base}/Admin/frmBirthEntryOld.aspx'")
            time.sleep(5)
            self._wait()
            if "frmBirthEntryOld" in self.page.url:
                print("  On entry form (JS nav)")
                return True
        except Exception as e:
            print(f"  JS nav failed: {e}")

        print(f"  Could not reach entry form. Current: {self.page.url}")
        return False

    # ---- PRE-FORM ----

    def fill_pre_form(self, record: dict) -> bool:
        print(f"\nPre-form: Reg#{record['reg_number']}/{record['year']}")
        try:
            self.page.locator("input[type='text']").first.fill(record["reg_number"])
            self.page.locator("input[type='text']").nth(1).fill(record["year"])
            self.page.locator("select").first.select_option(index=PRE_FORM_REGISTRAR_INDEX)
            time.sleep(1)
            self.page.locator("text=प्रवेश करे").first.click()
            self._wait()
            time.sleep(POST_CLICK_DELAY / 1000)

            if self.page.locator("text=जन्म दिनांक").count() > 0:
                print("  Main form loaded")
                return True

            print("  Form may not have loaded")
            self.screenshot(f"preform_{record['reg_number']}")
            return False
        except Exception as e:
            print(f"  Pre-form error: {e}")
            self.screenshot(f"preform_err_{record['reg_number']}")
            return False

    # ---- MAIN FORM ----

    def fill_main_form(self, record: dict) -> bool:
        print(f"\nFilling form for Row {record['excel_row']}...")
        try:
            p = self.page
            CP = "ContentPlaceHolder1_"

            # ============================================
            # PHASE 1: DATES FIRST (they trigger postbacks for late registration)
            # Using JS with __doPostBack to properly trigger server-side validation
            # ============================================

            p.evaluate(f"""() => {{
                const el = document.getElementById('{CP}txtbirth');
                el.value = '{record["birth_date"]}';
                el.dispatchEvent(new Event('change', {{bubbles: true}}));
            }}""")
            time.sleep(3)
            print(f"  Birth date: {record['birth_date']}")

            p.evaluate(f"""() => {{
                const el = document.getElementById('{CP}txtregisdt');
                el.value = '{record["reg_date"]}';
                el.dispatchEvent(new Event('change', {{bubbles: true}}));
            }}""")
            time.sleep(3)
            print(f"  Reg date: {record['reg_date']}")

            # Verify dates actually stuck
            actual_birth = p.evaluate(f"document.getElementById('{CP}txtbirth').value")
            actual_reg = p.evaluate(f"document.getElementById('{CP}txtregisdt').value")
            if not actual_reg:
                # Re-enter reg date
                p.evaluate(f"""() => {{
                    const el = document.getElementById('{CP}txtregisdt');
                    el.value = '{record["reg_date"]}';
                    el.dispatchEvent(new Event('change', {{bubbles: true}}));
                }}""")
                time.sleep(3)
                print(f"  Reg date RE-ENTERED: {record['reg_date']}")

            # Handle late registration fee field if it appeared
            try:
                fee_visible = p.evaluate(f"""() => {{
                    const el = document.getElementById('{CP}txtfees');
                    return el && el.offsetParent !== null;
                }}""")
                if fee_visible:
                    print("  Late registration detected - fees=50, clicking affidavit radio")
                    # Keep default fees (50), click 1st affidavit radio option
                    p.evaluate(f"""() => {{
                        const el = document.getElementById('{CP}rdaffdvityesno_0');
                        if (el) el.click();
                    }}""")
                    time.sleep(3)  # Wait for any postback from affidavit radio
            except Exception:
                pass

            # ============================================
            # PHASE 2: ALL RADIO BUTTONS AND DROPDOWNS
            # Using JavaScript for reliability
            # ============================================

            # Gender
            gender_id = f"{CP}rdsex_1" if record["gender"] == "महिला" else f"{CP}rdsex_0"
            p.evaluate(f"document.getElementById('{gender_id}').click()")
            time.sleep(0.5)
            print(f"  Gender: {record['gender']}")

            # Residency - अन्य स्थान
            p.evaluate(f"document.getElementById('{CP}rdlocalother_1').click()")
            time.sleep(1)
            print("  Residency: अन्य स्थान")

            # Same address at birth - हाँ
            p.evaluate(f"document.getElementById('{CP}rdaddress_0').click()")
            time.sleep(1)
            print("  Same address at birth: हाँ")

            # Birth place - गैर संस्थागत
            p.evaluate(f"document.getElementById('{CP}rdinstitanal_1').click()")
            time.sleep(2)
            print("  Birth place: गैर संस्थागत")

            # District dropdown
            try:
                p.locator(f"#{CP}ddldistrict").select_option(label="जयपुर")
                time.sleep(1)
                print("  District: जयपुर")
            except Exception as e:
                print(f"  District issue: {e}")

            # Father religion
            dharm_id = f"{CP}RadDharm_1" if record["religion"] == "मुस्लिम" else f"{CP}RadDharm_0"
            p.evaluate(f"document.getElementById('{dharm_id}').click()")
            time.sleep(0.5)
            print(f"  Father religion: {record['religion']}")

            # Mother religion
            mdharm_id = f"{CP}RadDharmMother_1" if record["religion"] == "मुस्लिम" else f"{CP}RadDharmMother_0"
            p.evaluate(f"document.getElementById('{mdharm_id}').click()")
            time.sleep(0.5)
            print(f"  Mother religion: {record['religion']}")

            # Education dropdowns
            try:
                p.locator(f"#{CP}cmbFatherEdu").select_option(label="अवर्णित")
                p.locator(f"#{CP}cmbMotherEdu").select_option(label="अवर्णित")
            except Exception:
                pass

            # Occupation dropdowns
            try:
                p.locator(f"#{CP}cmbFatherBusiness").select_option(label="अकर्मकार")
                p.locator(f"#{CP}cmbMotherBusiness").select_option(label="अकर्मकार")
            except Exception:
                pass
            print("  Education/Occupation: set")

            # Delivery care - अवर्णित
            p.evaluate(f"document.getElementById('{CP}radHelpIfAny_2').click()")
            time.sleep(0.5)
            print("  Delivery care: अवर्णित")

            # Pregnancy duration dropdown
            p.locator(f"#{CP}cmbPregTime").select_option(label=FIXED["pregnancy_weeks"])
            time.sleep(0.5)
            print(f"  Pregnancy: {FIXED['pregnancy_weeks']}")

            # ============================================
            # PHASE 3: WAIT FOR ALL POSTBACKS TO SETTLE
            # ============================================
            print("  Waiting for postbacks to settle...")
            time.sleep(3)

            # ============================================
            # PHASE 4: RE-VERIFY AND FIX RADIO STATES
            # (postbacks can reset these)
            # ============================================

            # Re-check residency
            res_checked = p.evaluate(f"document.getElementById('{CP}rdlocalother_1').checked")
            if not res_checked:
                p.evaluate(f"document.getElementById('{CP}rdlocalother_1').click()")
                time.sleep(1)
                print("  Residency RE-CLICKED: अन्य स्थान")

            # Re-check birth place
            bp_checked = p.evaluate(f"document.getElementById('{CP}rdinstitanal_1').checked")
            if not bp_checked:
                p.evaluate(f"document.getElementById('{CP}rdinstitanal_1').click()")
                time.sleep(1)
                print("  Birth place RE-CLICKED: गैर संस्थागत")

            # Re-check same address
            sa_checked = p.evaluate(f"document.getElementById('{CP}rdaddress_0').checked")
            if not sa_checked:
                p.evaluate(f"document.getElementById('{CP}rdaddress_0').click()")
                time.sleep(1)
                print("  Same address RE-CLICKED: हाँ")

            # Re-verify gender radio
            gender_recheck_id = f"{CP}rdsex_1" if record["gender"] == "महिला" else f"{CP}rdsex_0"
            gender_checked = p.evaluate(f"document.getElementById('{gender_recheck_id}').checked")
            if not gender_checked:
                print(f"  Re-clicking gender: {gender_recheck_id}")
                p.evaluate(f"document.getElementById('{gender_recheck_id}').click()")
                time.sleep(3)

            # Wait for Phase 4 re-click postbacks to settle
            time.sleep(3)

            # ============================================
            # PHASE 5: ALL TEXT FIELDS (after everything has settled)
            # ============================================

            # Re-verify dates are still there
            actual_birth = p.evaluate(f"document.getElementById('{CP}txtbirth').value")
            actual_reg = p.evaluate(f"document.getElementById('{CP}txtregisdt').value")
            if not actual_birth:
                p.evaluate(f"""() => {{
                    document.getElementById('{CP}txtbirth').value = '{record["birth_date"]}';
                }}""")
                print(f"  Birth date RE-SET: {record['birth_date']}")
            if not actual_reg:
                p.evaluate(f"""() => {{
                    document.getElementById('{CP}txtregisdt').value = '{record["reg_date"]}';
                }}""")
                print(f"  Reg date RE-SET: {record['reg_date']}")

            # Child name
            p.locator(f"#{CP}txtbabyHindi").fill(record["child_name"])
            print(f"  Child: {record['child_name']}")

            # Father name
            p.locator(f"#{CP}txtfatherHindi").fill(record["father_name"])
            print(f"  Father: {record['father_name']}")

            # Mother name
            p.locator(f"#{CP}txtmotherHindi").fill(record["mother_name"])
            print(f"  Mother: {record['mother_name']}")

            # Permanent address (textarea)
            p.locator(f"#{CP}txtaddressHindi").fill(record["address"])
            print(f"  Address: {record['address']}")

            # Birth-time address (textarea)
            p.locator(f"#{CP}txtbirthaddressHindi").fill(record["address"])
            print(f"  Birth-time address: {record['address']}")

            # Informant name
            # Informant: use mother if father is अप्राप्त and mother is not
            if record["father_name"] == "अप्राप्त" and record["mother_name"] != "अप्राप्त":
                informant_name = record["mother_name"]
            else:
                informant_name = record["father_name"]
            p.locator(f"#{CP}txtinformer").fill(informant_name)
            print(f"  Informant: {informant_name}")

            # Information date
            p.evaluate(f"""() => {{
                document.getElementById('{CP}txtinfodt').value = '{record["reg_date"]}';
            }}""")
            print(f"  Info date: {record['reg_date']}")

            # Baby weight
            p.locator(f"#{CP}txtbabyweight").click(click_count=3)
            p.locator(f"#{CP}txtbabyweight").fill(FIXED["baby_weight"])
            print(f"  Weight: {FIXED['baby_weight']}")

            # Birth place address
            p.locator(f"#{CP}txtbpalaceHindi").fill(record["address"])
            print(f"  Birth place addr: {record['address']}")

            # Informant mobile
            p.locator(f"#{CP}txtmobile").fill(FIXED["informant_mobile"])
            print("  Informant mobile: 0")

            # Informant address
            p.locator(f"#{CP}txtinformeradd").fill(record["address"])
            print(f"  Informant address: {record['address']}")

            # Marriage age, mother age, live births - ABSOLUTE LAST
            p.locator(f"#{CP}txtMotherAgeOnMar").fill(FIXED["marriage_age"])
            p.locator(f"#{CP}txtMotherAgeOnBirth").fill(FIXED["mother_birth_age"])
            p.locator(f"#{CP}txtTotalChild").fill(FIXED["live_births"])
            print(f"  Marriage/Age/Births: {FIXED['marriage_age']}/{FIXED['mother_birth_age']}/{FIXED['live_births']}")

            # === ABSOLUTE LAST: Use JavaScript to set all fragile fields ===
            time.sleep(2)
            p.evaluate(f"""() => {{
                const set = (id, val) => {{
                    const el = document.getElementById(id);
                    if (el) {{ el.value = val; el.dispatchEvent(new Event('change')); }}
                }};
                set('{CP}txtMotherVillageName', '{record["address"]}');
                set('{CP}txtbpalaceHindi', '{record["address"]}');
                set('{CP}txtmobile', '{FIXED["informant_mobile"]}');
                set('{CP}txtinformeradd', '{record["address"]}');
                set('{CP}txtMotherAgeOnMar', '{FIXED["marriage_age"]}');
                set('{CP}txtMotherAgeOnBirth', '{FIXED["mother_birth_age"]}');
                set('{CP}txtTotalChild', '{FIXED["live_births"]}');
                set('{CP}txtbabyweight', '{FIXED["baby_weight"]}');
            }}""")
            print(f"  JS batch fill: village, birth place, informant, age, weight")

            print(f"  Form filled for Row {record['excel_row']}")
            return True

        except Exception as e:
            print(f"  Form error: {e}")
            self.screenshot(f"form_err_{record['excel_row']}")
            return False

    # ---- SUBMIT ----

    def submit_and_verify(self, record: dict) -> bool:
        print(f"\nSubmitting Row {record['excel_row']}...")
        try:
            dialog_messages = []

            def handle_dialog(dialog):
                try:
                    dialog_messages.append(dialog.message)
                    print(f"  ALERT: {dialog.message[:100]}")
                    dialog.accept()
                except Exception:
                    pass

            self.page.on("dialog", handle_dialog)
            self.screenshot(f"before_submit_row{record['excel_row']}")

            # Click submit using JavaScript to avoid any locator issues
            self.page.evaluate("""() => {
                const btns = document.querySelectorAll('input[type="submit"], input[type="button"], button');
                for (const b of btns) {
                    if (b.value && (b.value.includes('इंद्राज') || b.value.includes('इन्द्राज'))) {
                        b.click();
                        return 'clicked: ' + b.value;
                    }
                }
                // Fallback: find by ID pattern
                const sub = document.querySelector('[id*="btnSubmit"], [id*="btnsave"], [id*="Button"]');
                if (sub) { sub.click(); return 'clicked fallback: ' + sub.id; }
                return 'not found';
            }""")
            print("  Submit clicked, waiting...")

            # Wait for response - alerts may take a moment
            time.sleep(5)

            # Check dialog messages
            if dialog_messages:
                for msg in dialog_messages:
                    print(f"  Dialog received: {msg[:80]}")
                    if "पहले से उपलब्ध" in msg:
                        print("  POST-SUBMIT DUPLICATE - record already existed")
                        return "already_exists"
                    if "पंजीकृत" in msg:
                        print("  DUPLICATE detected")
                        return False

            # Check page content for success
            self.screenshot(f"after_submit_row{record['excel_row']}")

            # Check page content for success - retry up to 3 times
            success = False
            already_exists = False
            page_text = ""
            for attempt in range(3):
                page_text = self.page.content()
                if "इस विवरण का रिकॉर्ड पहले से उपलब्ध है" in page_text or "पहले से उपलब्ध" in page_text or "सेव नहीं किया जा सकता" in page_text:
                    already_exists = True
                    break
                if "संग्रहित" in page_text or "सूचना संग्रहित कर ली गयी है" in page_text:
                    success = True
                    break
                # Also check if पुनः शुरू करे button appeared (only shows after success)
                if "पुनः शुरू करे" in page_text:
                    success = True
                    break
                if attempt < 2:
                    time.sleep(3)  # Wait 3 more seconds before retrying

            # Re-check dialogs (alert may have fired during page content checks)
            for msg in dialog_messages:
                if "पहले से उपलब्ध" in msg or "सेव नहीं किया जा सकता" in msg:
                    print(f"  Dialog (late): {msg[:80]}")
                    print("  POST-SUBMIT DUPLICATE - record already existed")
                    try:
                        restart = self.page.locator("text=पुनः शुरू करे")
                        if restart.count() > 0:
                            restart.first.click()
                            self._wait()
                    except Exception:
                        pass
                    return "already_exists"

            if already_exists:
                print("  POST-SUBMIT DUPLICATE - record already existed")
                try:
                    restart = self.page.locator("text=पुनः शुरू करे")
                    if restart.count() > 0:
                        restart.first.click()
                        self._wait()
                        time.sleep(1)
                except Exception:
                    pass
                return "already_exists"

            if success:
                print("  SUCCESS - record saved!")
                self.records_this_session += 1

                # Click पुनः शुरू करे
                try:
                    restart = self.page.locator("text=पुनः शुरू करे")
                    if restart.count() > 0:
                        restart.first.click()
                        self._wait()
                        time.sleep(1)
                        print("  Clicked punah shuru kare")
                except Exception:
                    pass
                return True

            if "पंजीकृत" in page_text:
                print("  DUPLICATE on page")
                return False

            # If we got dialog alerts with संग्रहित, count as success even without page text
            for msg in dialog_messages:
                if "संग्रहित" in msg:
                    print("  SUCCESS (from alert)")
                    self.records_this_session += 1
                    try:
                        restart = self.page.locator("text=पुनः शुरू करे")
                        if restart.count() > 0:
                            restart.first.click()
                            self._wait()
                            time.sleep(1)
                    except Exception:
                        pass
                    return True

            print("  No success indicator found")
            print(f"  Dialogs received: {len(dialog_messages)}")
            self.page.remove_listener("dialog", handle_dialog)
            return False

        except Exception as e:
            print(f"  Submit error: {e}")
            self.screenshot(f"submit_err_{record['excel_row']}")
            return False

    # ---- NAVIGATION BACK ----

    def ensure_on_pre_form(self) -> bool:
        if "frmBirthEntryOld" in self.page.url:
            self.page.evaluate("window.scrollTo(0, 0)")
            time.sleep(1)
            if self.page.locator("text=रजिस्ट्रेशन संख्या डाले").count() > 0:
                return True
        # Use JS navigation instead of page.goto to avoid redirect loops
        try:
            self.page.evaluate(f"window.location.href = '{self.portal_base}/Admin/frmBirthEntryOld.aspx'")
            time.sleep(5)
            self._wait()
            self.page.evaluate("window.scrollTo(0, 0)")
            return "frmBirthEntryOld" in self.page.url
        except Exception:
            return False

    # ---- SESSION CHECK ----

    def is_session_alive(self) -> bool:
        url = self.page.url
        if "Default.aspx" in url or "login" in url.lower():
            self.is_logged_in = False
            return False
        return True

    # ---- HELPERS ----

    def _wait(self):
        try:
            self.page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT)
        except PwTimeout:
            self.page.wait_for_load_state("domcontentloaded", timeout=PAGE_LOAD_TIMEOUT)

    def _type_date(self, label: str, date_str: str):
        try:
            row = self.page.locator(f"text={label}").first.locator("xpath=ancestor::tr")
            date_input = row.locator("input[type='text']").first
            date_input.click()
            date_input.fill(date_str)
            self.page.locator("text=जन्म रिपोर्ट").first.click()
            time.sleep(5)
            print(f"  {label}: {date_str}")
        except Exception as e:
            print(f"  Date field '{label}' issue: {e}")

    def _fill_field_in_row(self, label: str, value: str):
        if not value or value.lower() == "nan":
            return
        try:
            row = self.page.locator(f"text={label}").first.locator("xpath=ancestor::tr")
            field = row.locator("input[type='text'], textarea").first
            field.fill(value)
            print(f"  {label}: {value[:30]}{'...' if len(value) > 30 else ''}")
        except Exception as e:
            print(f"  Field '{label}' issue: {e}")

    def _click_radio_in_row(self, label: str, option: str):
        try:
            row = self.page.locator(f"text={label}").first.locator("xpath=ancestor::tr")
            opt = row.locator(f"text={option}")
            if opt.count() > 0:
                opt.first.click()
                print(f"  {label}: {option}")
                return
            lbl = row.locator(f"label:has-text('{option}')")
            if lbl.count() > 0:
                lbl.first.click()
                print(f"  {label}: {option}")
        except Exception as e:
            print(f"  Radio '{label}' -> '{option}' issue: {e}")
