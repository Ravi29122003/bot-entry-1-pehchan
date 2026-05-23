"""
Convert pre-form duplicates to success entries in progress.json.

Run this AFTER the rerun of rows 0-761 completes.
Pre-form duplicates ("पंजीकृत") during rerun are entries WE made earlier.
This script moves them from skipped_rows to completed_rows so billing is correct.

Usage: python convert_dups_to_success.py
"""

import json
import os
from datetime import datetime

PROGRESS_FILE = "logs/progress.json"

def main():
    if not os.path.exists(PROGRESS_FILE):
        print("ERROR: progress.json not found")
        return

    with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    skipped = data.get("skipped_rows", [])
    completed_set = set()
    for entry in data.get("completed_rows", []):
        if isinstance(entry, dict):
            completed_set.add(entry["row"])
        elif isinstance(entry, int):
            completed_set.add(entry)

    # Find skipped rows that are NOT already in completed
    to_convert = [s for s in skipped if s["row"] not in completed_set and s.get("reason") == "Duplicate"]
    
    if not to_convert:
        print("No pre-form duplicates to convert.")
        return

    print(f"Found {len(to_convert)} pre-form duplicates to convert to success.")
    print(f"Sample rows: {[s['row'] for s in to_convert[:10]]}")
    
    confirm = input("Convert these to success entries? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("Cancelled.")
        return

    # Move from skipped to completed
    now = datetime.now().isoformat()
    for s in to_convert:
        data["completed_rows"].append({
            "row": s["row"],
            "duration": 0,  # No timing data for these
            "timestamp": now,
        })
        data["total_success"] += 1

    # Remove converted entries from skipped
    converted_rows = {s["row"] for s in to_convert}
    data["skipped_rows"] = [s for s in skipped if s["row"] not in converted_rows]
    data["total_skipped"] = len(data["skipped_rows"])

    # Save
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Done! Converted {len(to_convert)} duplicates to success.")
    print(f"New totals: success={data['total_success']}, skipped={data['total_skipped']}")


if __name__ == "__main__":
    main()