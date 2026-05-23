"""
Progress Tracker — Saves and loads automation progress so we can
resume from where we left off after crashes, session timeouts, etc.
"""

import json
import os
from datetime import datetime


class ProgressTracker:
    def __init__(self, progress_file: str = "logs/progress.json"):
        self.progress_file = progress_file
        self.data = self._load()
        self._completed_set = self._build_completed_set()

    def _build_completed_set(self) -> set:
        rows = set()
        for entry in self.data.get("completed_rows", []):
            if isinstance(entry, dict) and isinstance(entry.get("row"), int):
                rows.add(entry["row"])
            elif isinstance(entry, int):
                rows.add(entry)
        return rows

    def _load(self) -> dict:
        """Load progress from disk, or create fresh."""
        if os.path.exists(self.progress_file):
            with open(self.progress_file, "r", encoding="utf-8") as f:
                data = json.load(f)
        else:
            data = {}

        data.setdefault("completed_rows", [])           # List of {row, duration, timestamp}
        data.setdefault("failed_rows", [])              # List of {row, error, detail, timestamp}
        data.setdefault("skipped_rows", [])             # List of {row, reason}
        data.setdefault("post_submit_duplicates", [])   # List of {row, timestamp}
        data.setdefault("last_completed_row", -1)
        data.setdefault("total_success", 0)
        data.setdefault("total_failed", 0)
        data.setdefault("total_skipped", 0)
        data.setdefault("total_post_submit_dup", 0)
        data.setdefault("started_at", None)
        data.setdefault("last_updated", None)
        return data

    def _save(self):
        """Write progress to disk."""
        os.makedirs(os.path.dirname(self.progress_file), exist_ok=True)
        self.data["last_updated"] = datetime.now().isoformat()
        with open(self.progress_file, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def mark_started(self):
        """Mark the start of a new run."""
        if not self.data["started_at"]:
            self.data["started_at"] = datetime.now().isoformat()
        self._save()

    def is_completed(self, excel_row: int) -> bool:
        """Check if a row has already been successfully entered."""
        return (excel_row + 2) in self._completed_set

    def mark_success(self, excel_row: int, duration: float = 0):
        """Mark a row as successfully entered."""
        row = excel_row + 2
        if row not in self._completed_set:
            self.data["completed_rows"].append({
                "row": row,
                "duration": duration,
                "timestamp": datetime.now().isoformat(),
            })
            self._completed_set.add(row)
            self.data["total_success"] += 1
            self.data["last_completed_row"] = row
        self._save()

    def mark_post_submit_duplicate(self, excel_row: int):
        """Mark a row that was detected as duplicate AFTER submit went through."""
        self.data["post_submit_duplicates"].append({
            "row": excel_row + 2,
            "timestamp": datetime.now().isoformat(),
        })
        self.data["total_post_submit_dup"] += 1
        self._save()

    def mark_failed(self, excel_row: int, error: str, detail: str = ""):
        """Mark a row as failed."""
        self.data["failed_rows"].append({
            "row": excel_row + 2,
            "error": error,
            "detail": detail,
            "timestamp": datetime.now().isoformat(),
        })
        self.data["total_failed"] += 1
        self._save()

    def mark_skipped(self, excel_row: int, reason: str):
        """Mark a row as intentionally skipped."""
        self.data["skipped_rows"].append({
            "row": excel_row + 2,
            "reason": reason,
        })
        self.data["total_skipped"] += 1
        self._save()

    def get_summary(self) -> str:
        """Print a human-readable summary."""
        return (
            f"\n{'='*50}\n"
            f"  PROGRESS SUMMARY\n"
            f"{'='*50}\n"
            f"  ✓ Success           : {self.data['total_success']}\n"
            f"  ✗ Failed            : {self.data['total_failed']}\n"
            f"  ⊘ Skipped           : {self.data['total_skipped']}\n"
            f"  ⟳ Post-submit dups  : {self.data['total_post_submit_dup']}\n"
            f"  Last row            : {self.data['last_completed_row']}\n"
            f"  Started             : {self.data['started_at']}\n"
            f"  Updated             : {self.data['last_updated']}\n"
            f"{'='*50}\n"
        )

    def get_next_row_index(self) -> int:
        """Get the next row to process (for resuming)."""
        if not self._completed_set:
            return 0
        return max(self._completed_set) + 1
