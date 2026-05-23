"""
Pehchan Portal Automation — Configuration
"""

BASE_URL = "https://pehchan.rajasthan.gov.in"
USERNAME = "YOUR_OPERATOR_ID"  # set to your pehchan operator account ID before running

REGISTRATION_YEAR = "1972"
PRE_FORM_REGISTRAR_INDEX = 2  # 0="--विकल्प चुने--", 1="नगर निगम ग्रेटर जयपुर (पुराना)", 2="नगर निगम जयपुर"

GENDER_MAP = {"पुरूष": "पुरूष", "स्त्री": "महिला"}

FIXED = {
    "residency": "अन्य स्थान",
    "birth_address_same": "हाँ",
    "birth_place_type": "गैर संस्थागत",
    "mother_mobile": "",
    "father_mobile": "",
    "informant_mobile": "0",
    "district": "जयपुर",
    "marriage_age": "18",
    "mother_birth_age": "21",
    "live_births": "0",
    "baby_weight": "2.5",
    "pregnancy_weeks": "36 सप्ताह",
}

PAGE_LOAD_TIMEOUT = 60000
ELEMENT_TIMEOUT = 15000
POST_CLICK_DELAY = 2000
POST_DATE_DELAY = 3000
TYPING_DELAY = 50
BETWEEN_RECORDS_DELAY = 500
POST_DROPDOWN_DELAY = 2000

MAX_RECORDS_PER_SESSION = 50
SESSION_CHECK_INTERVAL = 5

PROGRESS_FILE = "logs/progress.json"
SCREENSHOT_DIR = "logs/screenshots"
