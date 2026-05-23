"""
Pehchan Portal Automation - Main Runner

Usage:
  python run.py --dry-run                # Validate data, no browser
  python run.py --start 0 --end 5        # POC: first 5 records
  python run.py                          # Run all, resume from last
  python run.py --reset                  # Clear progress, start fresh
"""

import sys
import os
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# When spawned via subprocess.Popen with CREATE_NEW_CONSOLE, Python's sys.stdin
# is wired to the parent's stdin handle, not the new console window's input.
# Reopen it from CONIN$ so input() reads from THIS console window.
if sys.platform == "win32":
    try:
        sys.stdin = open("CONIN$", "r", encoding="utf-8", errors="replace")
    except Exception:
        pass

import argparse
import getpass
import time
import time as _time

from excel_reader import load_records, get_record_summary
from portal_bot import PehchanBot
from progress_tracker import ProgressTracker
from config.settings import (
    BETWEEN_RECORDS_DELAY, MAX_RECORDS_PER_SESSION,
    SESSION_CHECK_INTERVAL, PROGRESS_FILE,
)


def main():
    args = parse_args()

    print("\n" + "=" * 60)
    print("  PEHCHAN PORTAL AUTOMATION - Birth Entry Bot")
    print("=" * 60)

    if not os.path.exists(args.file):
        print(f"\nFile not found: {args.file}")
        sys.exit(1)

    records = load_records(args.file)
    if not records:
        print("No records. Exiting.")
        sys.exit(1)

    start = args.start or 0
    end = args.end or len(records)
    target = records[start:end]
    print(f"\n  Target: rows {start} to {end - 1} ({len(target)} records)")
    overall_start = time.time()

    if args.dry_run:
        print("\nDRY RUN - data validation only\n")
        for i, r in enumerate(target):
            addr = r['address'][:25] + '...' if len(r['address']) > 25 else r['address']
            print(f"  [{r['excel_row']:>4}] {get_record_summary(r)}")
            print(f"         religion={r['religion']} | address={addr}")
        print(f"\n  Total: {len(target)} records")
        return

    tracker = ProgressTracker(PROGRESS_FILE)
    if args.reset:
        confirm = input("\n  Type 'RESET' to confirm: ")
        if confirm == "RESET":
            if os.path.exists(PROGRESS_FILE):
                os.remove(PROGRESS_FILE)
            tracker = ProgressTracker(PROGRESS_FILE)
            print("  Progress reset")
        else:
            print("  Cancelled.")
            return

    print("\nCredentials")
    password = getpass.getpass(f"  Password for operator account: ")
    if not password:
        print("  Empty password")
        sys.exit(1)

    bot = PehchanBot(headless=False)
    tracker.mark_started()

    try:
        bot.start()

        if not bot.login(password):
            print("\nLogin failed. Wrong CAPTCHA? Try again.")
            bot.stop()
            return

        bot.post_login_setup()

        if not bot.navigate_to_entry_form():
            print("\nCould not reach entry form.")
            bot.stop()
            return

        print("\n" + "=" * 60)
        print("  STARTING DATA ENTRY")
        print("=" * 60)

        for i, record in enumerate(target):
            row_idx = record["excel_row"]

            if tracker.is_completed(row_idx):
                print(f"\n  [{row_idx:>4}] Already done - skipping")
                continue

            if i > 0 and i % SESSION_CHECK_INTERVAL == 0:
                if not bot.is_session_alive():
                    print("\nSession expired. Re-login needed...")
                    if not bot.login(password):
                        print("Re-login failed. Saving progress.")
                        break
                    bot.post_login_setup()
                    bot.navigate_to_entry_form()

            print(f"\n{'=' * 60}")
            print(f"  [{i + 1}/{len(target)}] {get_record_summary(record)}")
            print(f"{'=' * 60}")
            entry_start = time.time()

            try:
                if not bot.ensure_on_pre_form():
                    bot.navigate_to_entry_form()

                duplicate_detected = []

                def on_duplicate(dialog):
                    try:
                        if "पंजीकृत" in dialog.message or "पहले से संग्रहित" in dialog.message:
                            duplicate_detected.append(dialog.message)
                        dialog.accept()
                    except Exception:
                        pass

                bot.page.on("dialog", on_duplicate)
                success = bot.fill_pre_form(record)
                bot.page.remove_listener("dialog", on_duplicate)

                if duplicate_detected:
                    print(f"  DUPLICATE - already registered")
                    tracker.mark_skipped(row_idx, "Duplicate")
                    bot.ensure_on_pre_form()
                    continue

                if not success:
                    tracker.mark_failed(row_idx, "Pre-form failed", detail="Portal didn't load main form")
                    bot.ensure_on_pre_form()
                    continue

                if not bot.fill_main_form(record):
                    tracker.mark_failed(row_idx, "Form fill failed", detail="Form fill returned False")
                    bot.ensure_on_pre_form()
                    continue

                submit_result = bot.submit_and_verify(record)
                entry_duration = round(time.time() - entry_start, 1)

                if submit_result == "already_exists":
                    print(f"  Row {row_idx} POST-SUBMIT DUPLICATE in {entry_duration}s")
                    tracker.mark_post_submit_duplicate(row_idx)
                    bot.ensure_on_pre_form()
                    time.sleep(BETWEEN_RECORDS_DELAY / 1000)
                    continue

                if submit_result is False:
                    tracker.mark_failed(row_idx, "Submit failed", detail="submit_and_verify returned False")
                    bot.ensure_on_pre_form()
                    continue

                tracker.mark_success(row_idx, duration=entry_duration)
                print(f"  Row {row_idx} DONE in {entry_duration}s")
                time.sleep(BETWEEN_RECORDS_DELAY / 1000)

            except KeyboardInterrupt:
                print("\n\nStopped by user")
                print(tracker.get_summary())
                break
            except Exception as e:
                print(f"\n  Error on row {row_idx}: {e}")
                tracker.mark_failed(row_idx, "Exception", detail=str(e))
                bot.screenshot(f"error_row{row_idx}")
                try:
                    bot.ensure_on_pre_form()
                except Exception:
                    print("  Recovery failed.")
                    break

    except KeyboardInterrupt:
        print("\n\nStopped by user")
    except Exception as e:
        print(f"\nFatal: {e}")
        bot.screenshot("fatal")
    finally:
        total_time = round(time.time() - overall_start, 1) if 'overall_start' in dir() else 0
        print(tracker.get_summary())
        if total_time > 0:
            attempted = len(target)
            print(f"  Total time: {total_time}s ({round(total_time/60, 1)} min)")
            if attempted > 0:
                avg = round(total_time / attempted, 1)
                remaining = len(records) - len(tracker.data.get('completed_rows', []))
                est_hours = round((avg * remaining) / 3600, 1)
                print(f"  Avg per record: {avg}s (across {attempted} attempted)")
                print(f"  Remaining records: {remaining}")
                print(f"  Estimated time remaining: {est_hours} hours")
        bot.stop()


def parse_args():
    p = argparse.ArgumentParser(description="Pehchan Birth Entry Bot")
    p.add_argument("--file", default="data/1972_baki_pdf.xlsx")
    p.add_argument("--start", type=int, default=None)
    p.add_argument("--end", type=int, default=None)
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--reset", action="store_true")
    return p.parse_args()


if __name__ == "__main__":
    main()
