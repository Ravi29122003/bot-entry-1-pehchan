# Pehchan Portal Automation — Birth Entry Bot

Automates legacy birth record entry into the Rajasthan Civil Registration System (Pehchan) portal.

## Setup (Windows)

### Step 1: Install Python
Download Python 3.10+ from https://python.org/downloads
During install, CHECK "Add Python to PATH".

### Step 2: Clone / Copy this project
Copy the entire `pehchan-automation` folder to your Windows machine.

### Step 3: Install dependencies
Open Command Prompt (cmd) in the project folder:

```
cd pehchan-automation
pip install -r requirements.txt
playwright install chromium
```

### Step 4: Place your data file
Copy `1972_baki_pdf.xlsx` into the `data/` folder.

### Step 5: Set your password
Edit `config/settings.py` — the password is asked at runtime (never stored in code).

## Usage

### Dry Run (validate data, no browser)
```
python run.py --dry-run
```

### POC Test (first 10 records)
```
python run.py --start 0 --end 10
```

### Full Run (all records)
```
python run.py
```

### Resume after crash
```
python run.py
```
(Automatically resumes from last successful row)

### Reset progress
```
python run.py --reset
```

## How It Works

1. Opens a Chrome browser (you can see it)
2. Goes to Pehchan portal login page
3. Fills username — PAUSES for you to type CAPTCHA
4. After login, dismisses popup, selects Heritage registrar
5. Navigates to Legacy Birth Entry form
6. For each Excel row:
   - Fills pre-form (reg number, year, registrar)
   - Fills main form (all fields)
   - Clicks submit
   - Logs success/failure
   - Moves to next row
7. If session expires, pauses for re-login

## File Structure

```
pehchan-automation/
├── config/
│   ├── settings.py          # All configurations, URLs, mappings
│   └── __init__.py
├── data/
│   └── 1972_baki_pdf.xlsx   # Your Excel data file
├── logs/
│   ├── progress.json        # Tracks completed/failed rows
│   └── screenshots/         # Auto-captured screenshots
├── excel_reader.py          # Reads and cleans Excel data
├── portal_bot.py            # Browser automation logic
├── progress_tracker.py      # Resume/progress management
├── run.py                   # Main entry point
├── requirements.txt         # Python dependencies
└── README.md                # This file
```

## What's TODO (needs portal screenshots)

- Post-submission handling (what happens after "इंद्राज करे")
- Hospital search dropdown exact behavior
- Exact HTML element IDs (will be refined on first run)
- Form number field clarification
