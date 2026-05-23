"""
Pehchan Portal Automation — UNIFIED Live Tracking Dashboard
Auto-discovers every subfolder with logs/progress.json under the parent dir,
merges progress with priority (completed > post_dup > skipped > failed),
and serves a real-time dashboard at http://localhost:8060.
"""

import argparse
import json
import os
import re
import signal
import subprocess
import sys
import threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
from datetime import datetime, timedelta
from io import BytesIO
from urllib.parse import urlparse, parse_qs
import pandas as pd
from openpyxl import Workbook

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

# ─── CONFIGURATION ───
PROGRESS_REL = os.path.join("logs", "progress.json")
EXCEL_REL_CANDIDATES = [
    os.path.join("data", "1999-entries-new-1.xlsx"),
]
DEFAULT_DASHBOARD_PORT = 8060
AUTO_REFRESH_SECONDS = 10
ACTIVE_WINDOW_SECONDS = 120  # bot is "active" if last_updated within this window
RECENT_START_GRACE_SECONDS = 90  # treat a just-spawned bot as active for this long

BOTS_CONFIG_FILE = "bots-config.json"
TEMPLATE_DIR_NAME = "template"
# Excel row range covered by the dataset (header is row 1, data rows 2..3640)
EXCEL_FIRST_DATA_ROW = 2
EXCEL_LAST_DATA_ROW = 3640
# Bot --start arg corresponds to excel row (--start + 2); --end arg is exclusive
# of the last row, so a bot processes excel rows [start+2, end-1] inclusive.
ADD_BOT_CHUNK_SIZE = 200

PARENT_DIR = os.path.dirname(os.path.abspath(__file__))

# folder_name -> datetime of last successful spawn
RECENTLY_STARTED = {}
# folder_name -> pid of the bot we launched from the dashboard
SPAWNED_PIDS = {}
_SPAWN_LOCK = threading.Lock()


_BOT_NAME_RE = re.compile(r"^Bot-(\d+)$")


def _bot_sort_key(name):
    """Sort Bot-1 < Bot-2 < ... < Bot-10; non-matching names go to the end."""
    m = _BOT_NAME_RE.match(name)
    if m:
        return (0, int(m.group(1)))
    return (1, name)


def load_bots_config():
    """Load bots-config.json; returns empty dict if missing or invalid."""
    path = os.path.join(PARENT_DIR, BOTS_CONFIG_FILE)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_bots_config(config):
    path = os.path.join(PARENT_DIR, BOTS_CONFIG_FILE)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def discover_bot_folders():
    """Return list of absolute paths for bots listed in bots-config.json.
    Folders that don't exist on disk are skipped silently."""
    config = load_bots_config()
    folders = []
    for name in sorted(config.keys(), key=_bot_sort_key):
        full = os.path.join(PARENT_DIR, name)
        if os.path.isdir(full):
            folders.append(full)
    return folders


def find_excel_in_folder(folder):
    for rel in EXCEL_REL_CANDIDATES:
        p = os.path.join(folder, rel)
        if os.path.exists(p):
            return p
    data_dir = os.path.join(folder, "data")
    if os.path.isdir(data_dir):
        for f in os.listdir(data_dir):
            if f.endswith(".xlsx"):
                return os.path.join(data_dir, f)
    return None


def find_excel_anywhere(folders):
    for f in folders:
        p = find_excel_in_folder(f)
        if p:
            return p
    return None


def load_excel_data(excel_path):
    df = pd.read_excel(excel_path)
    records = []
    for i, row in df.iterrows():
        excel_row = i + 2

        birth_date = None
        reg_date = None
        gap_days = None
        is_late_reg = False

        try:
            bd = row.get('जन्म दिनांक')
            rd = row.get('रजि दिनांक')
            if isinstance(bd, datetime):
                birth_date = bd
            else:
                birth_date = pd.to_datetime(bd, dayfirst=True)
            if isinstance(rd, datetime):
                reg_date = rd
            else:
                reg_date = pd.to_datetime(rd, dayfirst=True)
            if birth_date and reg_date:
                gap_days = (reg_date - birth_date).days
                is_late_reg = gap_days > 20
        except Exception:
            pass

        reg_num = str(row.get('रजि क्रमांक', '')).strip()
        year = str(row.get('Unnamed: 1', '')).strip()
        child = str(row.get('बालक का नाम', '')).strip()
        father = str(row.get('पिता का नाम', '')).strip()
        mother = str(row.get('माता का नाम', '')).strip()
        gender = str(row.get('लिंग', '')).strip()
        religion = str(row.get('हिन्दू मुस्लिम', '')).strip()
        address = str(row.get('पता', row.get('पता जन्म स्थान', row.get('पता/जन्म स्थान', '')))).strip()

        records.append({
            "excel_row": excel_row,
            "bot_row": i,
            "reg_num": reg_num,
            "year": year,
            "child": child,
            "father": father,
            "mother": mother,
            "gender": gender,
            "religion": religion,
            "address": address[:60],
            "birth_date": birth_date.strftime("%d/%m/%Y") if birth_date else "",
            "reg_date": reg_date.strftime("%d/%m/%Y") if reg_date else "",
            "gap_days": gap_days,
            "is_late_reg": is_late_reg,
        })
    return records


def _load_progress_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    return data


def load_progress_for(folder):
    """Load logs/progress.json plus any logs/progress-*.json archives, merged.
    Dedup by row: latest timestamp wins for completed/post_dup/failed;
    first-seen wins for skipped (no timestamp). Uses earliest started_at and
    latest last_updated across all files. Returns None if no files found."""
    logs_dir = os.path.join(folder, "logs")
    if not os.path.isdir(logs_dir):
        return None

    paths = []
    main_path = os.path.join(logs_dir, "progress.json")
    if os.path.exists(main_path):
        paths.append(main_path)
    try:
        for entry in sorted(os.listdir(logs_dir)):
            if entry.startswith("progress-") and entry.endswith(".json"):
                paths.append(os.path.join(logs_dir, entry))
    except OSError:
        pass

    if not paths:
        return None

    merged_completed = {}   # row -> entry with latest timestamp
    merged_post_dup = {}    # row -> entry with latest timestamp
    merged_skipped = {}     # row -> first occurrence (no timestamps)
    merged_failed = {}      # row -> entry with latest timestamp
    started_at = None
    last_updated = None

    def _newer(cur, ts):
        return cur is None or (ts and ts > cur.get("timestamp", ""))

    for path in paths:
        data = _load_progress_json(path)
        if data is None:
            continue

        for entry in data.get("completed_rows", []):
            if isinstance(entry, dict):
                r = entry.get("row")
                if r is None:
                    continue
                if _newer(merged_completed.get(r), entry.get("timestamp", "")):
                    merged_completed[r] = entry
            elif isinstance(entry, int):
                merged_completed.setdefault(entry, {"row": entry, "duration": 0, "timestamp": ""})

        for entry in data.get("post_submit_duplicates", []):
            r = entry.get("row")
            if r is None:
                continue
            if _newer(merged_post_dup.get(r), entry.get("timestamp", "")):
                merged_post_dup[r] = entry

        for entry in data.get("skipped_rows", []):
            r = entry.get("row")
            if r is None:
                continue
            merged_skipped.setdefault(r, entry)

        for entry in data.get("failed_rows", []):
            r = entry.get("row")
            if r is None:
                continue
            if _newer(merged_failed.get(r), entry.get("timestamp", "")):
                merged_failed[r] = entry

        s = data.get("started_at")
        if s and (started_at is None or s < started_at):
            started_at = s
        l = data.get("last_updated")
        if l and (last_updated is None or l > last_updated):
            last_updated = l

    return {
        "completed_rows": list(merged_completed.values()),
        "post_submit_duplicates": list(merged_post_dup.values()),
        "skipped_rows": list(merged_skipped.values()),
        "failed_rows": list(merged_failed.values()),
        "started_at": started_at,
        "last_updated": last_updated,
    }


def _is_recent(ts_str, window_seconds):
    if not ts_str:
        return False
    try:
        ts = datetime.fromisoformat(ts_str)
    except Exception:
        return False
    return (datetime.now() - ts) <= timedelta(seconds=window_seconds)


def _was_recently_started(folder_name):
    ts = RECENTLY_STARTED.get(folder_name)
    if ts is None:
        return False
    return (datetime.now() - ts) <= timedelta(seconds=RECENT_START_GRACE_SECONDS)


def spawn_bot(folder_path, start_row, end_row):
    """Spawn the bot as a detached subprocess. Returns its PID.
    stdout/stderr are inherited so bot output is visible live."""
    cmd = [
        "python", "run.py",
        "--file", "data/1999-entries-new-1.xlsx",
        "--start", str(start_row),
        "--end", str(end_row),
    ]

    kwargs = {
        "cwd": folder_path,
    }

    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.CREATE_NEW_CONSOLE
    else:
        kwargs["preexec_fn"] = os.setpgrp
        kwargs["start_new_session"] = True

    proc = subprocess.Popen(cmd, **kwargs)
    folder_name = os.path.basename(folder_path)
    SPAWNED_PIDS[folder_name] = proc.pid
    return proc.pid


_PY_NAMES = {"python.exe", "python", "python3", "python3.exe", "pythonw.exe"}


def find_bot_pids(folder_path):
    """Find PIDs of python processes whose cwd is folder_path.
    Returns a list of PIDs (empty if none found or psutil unavailable)."""
    if not HAS_PSUTIL:
        return []
    target = os.path.normcase(os.path.normpath(folder_path))
    pids = []
    for p in psutil.process_iter(["name", "cwd"]):
        try:
            name = (p.info.get("name") or "").lower()
            if name not in _PY_NAMES:
                continue
            cwd = p.info.get("cwd")
            if cwd and os.path.normcase(os.path.normpath(cwd)) == target:
                pids.append(p.pid)
        except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
            continue
    return pids


def _kill_pid(pid):
    """Kill a single PID and its descendants. Returns (ok, error_message)."""
    if sys.platform == "win32":
        rc = subprocess.call(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        # 0 = success, 128 = process not found (already dead — treat as ok)
        if rc in (0, 128):
            return True, ""
        return False, f"taskkill rc={rc}"
    else:
        try:
            if HAS_PSUTIL:
                parent = psutil.Process(pid)
                for child in parent.children(recursive=True):
                    try:
                        child.terminate()
                    except psutil.NoSuchProcess:
                        pass
                parent.terminate()
            else:
                try:
                    pgid = os.getpgid(pid)
                    os.killpg(pgid, signal.SIGTERM)
                except ProcessLookupError:
                    pass
            return True, ""
        except psutil.NoSuchProcess:
            return True, ""  # already dead
        except Exception as e:
            return False, str(e)


def stop_bot(folder_name, folder_path):
    """Kill any python bot process running in folder_path. Returns (ok, message)."""
    pids = find_bot_pids(folder_path)

    # Fall back to tracked PID for dashboard-spawned bots (e.g. when psutil missing)
    tracked = SPAWNED_PIDS.get(folder_name)
    if tracked and tracked not in pids:
        pids.append(tracked)

    if not pids:
        if not HAS_PSUTIL:
            return False, "psutil not installed and no tracked PID — install psutil to stop externally-started bots"
        return False, "No running python process found in this folder"

    errors = []
    for pid in pids:
        ok, err = _kill_pid(pid)
        if not ok:
            errors.append(f"pid {pid}: {err}")

    SPAWNED_PIDS.pop(folder_name, None)
    RECENTLY_STARTED.pop(folder_name, None)

    if errors:
        return False, "; ".join(errors)
    return True, f"stopped {len(pids)} process(es): {pids}"


def _iter_progress_rows(prog):
    """Yield (bucket, row) for each completed/post_dup/skipped entry in prog.
    Bucket is one of 'completed', 'post_dup', 'skipped'."""
    for entry in prog.get("completed_rows", []):
        r = entry.get("row") if isinstance(entry, dict) else entry
        if r is not None:
            yield "completed", r
    for entry in prog.get("post_submit_duplicates", []):
        r = entry.get("row")
        if r is not None:
            yield "post_dup", r
    for entry in prog.get("skipped_rows", []):
        r = entry.get("row")
        if r is not None:
            yield "skipped", r


def discover_all_progress_folders():
    """Scan PARENT_DIR for every subfolder whose logs/ directory has at least
    one progress.json or progress-*.json file. Unlike discover_bot_folders()
    (which is limited to entries in bots-config.json), this also includes
    folders for bots that have been removed from config — so their archived
    work still counts toward done_rows / row→bot mapping."""
    result = []
    try:
        entries = os.listdir(PARENT_DIR)
    except OSError:
        return result
    for name in sorted(entries):
        full = os.path.join(PARENT_DIR, name)
        if not os.path.isdir(full):
            continue
        logs_dir = os.path.join(full, "logs")
        if not os.path.isdir(logs_dir):
            continue
        try:
            for entry in os.listdir(logs_dir):
                if entry == "progress.json" or (entry.startswith("progress-") and entry.endswith(".json")):
                    result.append(full)
                    break
        except OSError:
            continue
    return result


def compute_done_rows():
    """Set of excel rows completed/post_dup/skipped by any bot — INCLUDING
    bots that have been removed from bots-config.json (so their work still
    counts). Failed rows are NOT included; they remain retry candidates."""
    done = set()
    for folder in discover_all_progress_folders():
        prog = load_progress_for(folder)
        if not prog:
            continue
        for _, r in _iter_progress_rows(prog):
            done.add(r)
    return done


def compute_row_to_bot_map():
    """Build excel_row -> bot_name for every row processed (completed/post_dup/
    skipped) by any bot, including bots no longer in bots-config.json. When
    multiple bots touched the same row, the bot with the lower _bot_sort_key
    wins (first-seen in sorted order)."""
    row_to_bot = {}
    sorted_folders = sorted(
        discover_all_progress_folders(),
        key=lambda p: _bot_sort_key(os.path.basename(p)),
    )
    for folder in sorted_folders:
        bot_name = os.path.basename(folder)
        prog = load_progress_for(folder)
        if not prog:
            continue
        for _, r in _iter_progress_rows(prog):
            row_to_bot.setdefault(r, bot_name)
    return row_to_bot


def pick_next_chunk(done_rows):
    """Collect the next ADD_BOT_CHUNK_SIZE pending excel rows starting from
    EXCEL_FIRST_DATA_ROW. A row is pending iff it is NOT in done_rows. Returns
    (cmd_start, cmd_end, excel_start, excel_end, pending_count) or None.
    The returned --start/--end span first and last pending rows; the bot skips
    rows it has already processed when iterating."""
    pending = []
    for r in range(EXCEL_FIRST_DATA_ROW, EXCEL_LAST_DATA_ROW + 1):
        if r in done_rows:
            continue
        pending.append(r)
        if len(pending) >= ADD_BOT_CHUNK_SIZE:
            break

    if not pending:
        return None
    excel_start = pending[0]
    excel_end = pending[-1]
    return excel_start - 2, excel_end + 1, excel_start, excel_end, len(pending)


def next_bot_name(config):
    n = 1
    while f"Bot-{n}" in config:
        n += 1
    return f"Bot-{n}"


def create_bot_from_template(bot_name):
    """Copy template/ into <PARENT>/<bot_name>/. Creates an empty logs/ dir."""
    import shutil
    src = os.path.join(PARENT_DIR, TEMPLATE_DIR_NAME)
    dst = os.path.join(PARENT_DIR, bot_name)
    if not os.path.isdir(src):
        raise RuntimeError(f"Template not found at {src}")
    if os.path.exists(dst):
        raise RuntimeError(f"Target folder already exists: {dst}")
    shutil.copytree(src, dst)
    os.makedirs(os.path.join(dst, "logs"), exist_ok=True)
    return dst


def merge_progress(folders):
    """
    Merge progress across all bot folders with priority:
      completed > post_dup > skipped > failed
    Returns merged dict shaped like a single progress.json (for re-use of
    build_dashboard_data), plus per-bot summaries.
    """
    completed_lookup = {}   # row -> {duration, timestamp, bot}
    post_dup_lookup = {}    # row -> {timestamp, bot}
    skipped_set_global = set()
    failed_lookup = {}      # row -> latest {error, detail, timestamp, bot}

    per_bot = []
    earliest_start = None
    latest_update = None

    config = load_bots_config()

    for folder in folders:
        folder_name = os.path.basename(folder)
        prog = load_progress_for(folder)
        if prog is None:
            # Newly-added bot that hasn't started yet — emit a per-bot entry
            # with zero counts so it still shows up in the UI.
            prog = {
                "completed_rows": [],
                "failed_rows": [],
                "skipped_rows": [],
                "post_submit_duplicates": [],
                "started_at": None,
                "last_updated": None,
            }

        # Merge completed (latest timestamp wins for duration/timestamp)
        bot_completed_rows = set()
        for entry in prog.get("completed_rows", []):
            if isinstance(entry, dict):
                r = entry.get("row")
                if r is None:
                    continue
                ts = entry.get("timestamp", "")
                dur = entry.get("duration", 0)
                bot_completed_rows.add(r)
                cur = completed_lookup.get(r)
                if cur is None or (ts and ts > cur["timestamp"]):
                    completed_lookup[r] = {"duration": dur, "timestamp": ts, "bot": folder_name}
            elif isinstance(entry, int):
                bot_completed_rows.add(entry)
                completed_lookup.setdefault(entry, {"duration": 0, "timestamp": "", "bot": folder_name})

        # Merge post-submit duplicates
        bot_post_dup_rows = set()
        for entry in prog.get("post_submit_duplicates", []):
            r = entry.get("row")
            if r is None:
                continue
            bot_post_dup_rows.add(r)
            ts = entry.get("timestamp", "")
            cur = post_dup_lookup.get(r)
            if cur is None or (ts and ts > cur["timestamp"]):
                post_dup_lookup[r] = {"timestamp": ts, "bot": folder_name}

        # Merge skipped
        bot_skipped_rows = set()
        for s in prog.get("skipped_rows", []):
            r = s.get("row")
            if r is None:
                continue
            bot_skipped_rows.add(r)
            skipped_set_global.add(r)

        # Merge failed (latest timestamp wins)
        bot_failed_rows = set()
        for f in prog.get("failed_rows", []):
            r = f.get("row")
            if r is None:
                continue
            bot_failed_rows.add(r)
            ts = f.get("timestamp", "")
            cur = failed_lookup.get(r)
            if cur is None or ts > cur["timestamp"]:
                failed_lookup[r] = {
                    "error": f.get("error", "Unknown"),
                    "detail": f.get("detail", ""),
                    "timestamp": ts,
                    "bot": folder_name,
                }

        # Track session-level timestamps
        s_at = prog.get("started_at")
        l_at = prog.get("last_updated")
        if s_at:
            if earliest_start is None or s_at < earliest_start:
                earliest_start = s_at
        if l_at:
            if latest_update is None or l_at > latest_update:
                latest_update = l_at

        cfg = config.get(folder_name, {})
        rng_start = cfg.get("start")
        rng_end = cfg.get("end")
        if rng_start is not None and rng_end is not None:
            # Bot processes excel rows [start+2, end-1] inclusive
            assigned_total = rng_end - rng_start - 2
            in_range = set(range(rng_start + 2, rng_end))
            bot_completed_cur = bot_completed_rows & in_range
            bot_post_dup_cur = bot_post_dup_rows & in_range
            bot_skipped_cur = bot_skipped_rows & in_range
            bot_failed_cur = bot_failed_rows & in_range
        else:
            # No current assignment (bot was reset). Card shows zeros; archive
            # data still feeds the global merge above.
            assigned_total = 0
            bot_completed_cur = set()
            bot_post_dup_cur = set()
            bot_skipped_cur = set()
            bot_failed_cur = set()

        # Per-bot effective status counts (priority within the current assignment only)
        bot_failed_effective = bot_failed_cur - bot_completed_cur - bot_post_dup_cur - bot_skipped_cur
        bot_skipped_effective = bot_skipped_cur - bot_completed_cur - bot_post_dup_cur

        processed = (
            len(bot_completed_cur)
            + len(bot_post_dup_cur)
            + len(bot_skipped_effective)
            + len(bot_failed_effective)
        )
        pending = max(0, assigned_total - processed) if assigned_total is not None else None

        last_processed_row = max(
            bot_completed_cur | bot_post_dup_cur | bot_failed_cur,
            default=None,
        )

        recently_started = _was_recently_started(folder_name)
        is_active = _is_recent(l_at, ACTIVE_WINDOW_SECONDS) or recently_started

        per_bot.append({
            "folder": folder_name,
            "folder_path": folder,
            "range_start": rng_start,
            "range_end": rng_end,
            "assigned_total": assigned_total,
            "success": len(bot_completed_cur),
            "post_dup": len(bot_post_dup_cur),
            "skipped": len(bot_skipped_effective),
            "failed": len(bot_failed_effective),
            "pending": pending,
            "last_processed_row": last_processed_row,
            "started_at": s_at,
            "last_updated": l_at,
            "active": is_active,
            "recently_started": recently_started,
        })

    # Apply global priority: remove from lower buckets anything in higher buckets
    completed_set = set(completed_lookup.keys())
    post_dup_set = set(post_dup_lookup.keys()) - completed_set
    skipped_set = skipped_set_global - completed_set - post_dup_set
    for r in completed_set | post_dup_set | skipped_set:
        failed_lookup.pop(r, None)

    # Rebuild merged progress.json-shaped dict for build_dashboard_data
    merged_completed = [
        {"row": r, "duration": info["duration"], "timestamp": info["timestamp"]}
        for r, info in completed_lookup.items()
    ]
    # Sort by timestamp to keep first-entry-skip logic meaningful
    merged_completed.sort(key=lambda x: x.get("timestamp", ""))

    merged_post_dup = [
        {"row": r, "timestamp": info["timestamp"]}
        for r, info in post_dup_lookup.items()
        if r in post_dup_set
    ]
    merged_skipped = [{"row": r} for r in skipped_set]
    merged_failed = [
        {
            "row": r,
            "error": info["error"],
            "detail": info["detail"],
            "timestamp": info["timestamp"],
        }
        for r, info in failed_lookup.items()
    ]

    merged = {
        "completed_rows": merged_completed,
        "failed_rows": merged_failed,
        "skipped_rows": merged_skipped,
        "post_submit_duplicates": merged_post_dup,
        "total_success": len(merged_completed),
        "total_failed": len(merged_failed),
        "total_skipped": len(merged_skipped),
        "total_post_submit_dup": len(merged_post_dup),
        "started_at": earliest_start,
        "last_updated": latest_update,
    }

    resume_commands = build_resume_commands(per_bot)
    return merged, per_bot, resume_commands


def build_resume_commands(per_bot):
    """For each bot, compute its resume command or running/completed status.
    Skips bots without a current range (they've been reset and have nothing
    to resume — they need to be assigned work via the Work Dispatcher)."""
    items = []
    for b in per_bot:
        rng_start = b["range_start"]
        rng_end = b["range_end"]
        if rng_start is None or rng_end is None:
            continue
        last = b["last_processed_row"]

        if last is not None:
            resume_start = last - 2
        elif rng_start is not None:
            resume_start = rng_start
        else:
            resume_start = None

        if b["active"]:
            status = "running"
        elif b["pending"] is not None and b["pending"] == 0:
            status = "completed"
        else:
            status = "resume"

        if status == "resume" and rng_end is not None and resume_start is not None:
            command = (
                f'cd "{b["folder_path"]}"; python run.py '
                f'--file data/1999-entries-new-1.xlsx '
                f'--start {resume_start} --end {rng_end}'
            )
        else:
            command = None

        items.append({
            "folder": b["folder"],
            "status": status,
            "last_processed_row": last,
            "resume_start": resume_start,
            "range_start": rng_start,
            "range_end": rng_end,
            "command": command,
            "can_stop": status == "running",
        })
    return items


def build_dashboard_data(excel_records, progress):
    # Build completed lookup
    completed_lookup = {}
    for entry in progress.get("completed_rows", []):
        if isinstance(entry, dict):
            completed_lookup[entry["row"]] = {
                "duration": entry.get("duration", 0),
                "timestamp": entry.get("timestamp", ""),
            }
        elif isinstance(entry, int):
            completed_lookup[entry] = {"duration": 0, "timestamp": ""}
    completed_set = set(completed_lookup.keys())

    post_dup_lookup = {}
    for entry in progress.get("post_submit_duplicates", []):
        post_dup_lookup[entry["row"]] = {"timestamp": entry.get("timestamp", "")}
    post_dup_set = set(post_dup_lookup.keys())

    skipped_set = set()
    for s in progress.get("skipped_rows", []):
        skipped_set.add(s["row"])
    skipped_set -= completed_set
    skipped_set -= post_dup_set

    failed_lookup = {}
    for f in progress.get("failed_rows", []):
        r = f["row"]
        if r not in failed_lookup or f["timestamp"] > failed_lookup[r]["timestamp"]:
            failed_lookup[r] = {
                "error": f.get("error", "Unknown"),
                "detail": f.get("detail", ""),
                "timestamp": f["timestamp"],
            }
    for r in completed_set | post_dup_set | skipped_set:
        failed_lookup.pop(r, None)

    # Per-bot one entry has login+captcha cost; with N bots, the first N entries
    # carry that cost. Skip the first entry of each bot from the avg calculation.
    durations = []
    for entry in progress.get("completed_rows", []):
        if isinstance(entry, dict) and entry.get("duration", 0) > 0:
            durations.append(entry["duration"])

    durations_for_avg = durations[1:] if len(durations) > 1 else durations
    avg_duration = round(sum(durations_for_avg) / len(durations_for_avg), 1) if durations_for_avg else 0
    min_duration = round(min(durations_for_avg), 1) if durations_for_avg else 0
    max_duration = round(max(durations_for_avg), 1) if durations_for_avg else 0
    total_bot_time = round(sum(durations), 1)

    rows_data = []
    for rec in excel_records:
        er = rec["excel_row"]
        if er in completed_set:
            status = "success"
            error = None
            detail = None
            timestamp = completed_lookup[er]["timestamp"]
            duration = completed_lookup[er]["duration"]
        elif er in post_dup_set:
            status = "post_dup"
            error = None
            detail = "Record already existed on portal"
            timestamp = post_dup_lookup[er]["timestamp"]
            duration = 0
        elif er in failed_lookup:
            status = "failed"
            error = failed_lookup[er]["error"]
            detail = failed_lookup[er]["detail"]
            timestamp = failed_lookup[er]["timestamp"]
            duration = 0
        elif er in skipped_set:
            status = "skipped"
            error = None
            detail = None
            timestamp = None
            duration = 0
        else:
            status = "pending"
            error = None
            detail = None
            timestamp = None
            duration = 0
        rows_data.append({
            **rec,
            "status": status,
            "error": error,
            "detail": detail,
            "timestamp": timestamp,
            "duration": duration,
        })

    total = len(excel_records)
    excel_row_set = {r["excel_row"] for r in excel_records}
    success_count = len(completed_set & excel_row_set)
    post_dup_count = len(post_dup_set & excel_row_set)
    failed_count = len([r for r in rows_data if r["status"] == "failed"])
    skipped_count = len([r for r in rows_data if r["status"] == "skipped"])
    pending_count = len([r for r in rows_data if r["status"] == "pending"])

    form_fill_fails = len([r for r in rows_data if r["status"] == "failed" and r.get("error") == "Form fill failed"])
    submit_fails = len([r for r in rows_data if r["status"] == "failed" and r.get("error") == "Submit failed"])
    preform_fails = len([r for r in rows_data if r["status"] == "failed" and r.get("error") == "Pre-form failed"])
    exception_fails = len([r for r in rows_data if r["status"] == "failed" and r.get("error") == "Exception"])

    late_reg_total = len([r for r in excel_records if r["is_late_reg"]])
    late_reg_done = len([r for r in rows_data if r["is_late_reg"] and r["status"] == "success"])
    late_reg_failed = len([r for r in rows_data if r["is_late_reg"] and r["status"] == "failed"])
    late_reg_pending = len([r for r in rows_data if r["is_late_reg"] and r["status"] == "pending"])
    late_reg_post_dup = len([r for r in rows_data if r["is_late_reg"] and r["status"] == "post_dup"])

    billable = success_count + post_dup_count

    est_remaining_seconds = avg_duration * pending_count if avg_duration > 0 else 0
    est_remaining_hours = round(est_remaining_seconds / 3600, 1)

    return {
        "stats": {
            "total": total,
            "success": success_count,
            "post_dup": post_dup_count,
            "failed": failed_count,
            "skipped": skipped_count,
            "pending": pending_count,
            "billable": billable,
            "form_fill_fails": form_fill_fails,
            "submit_fails": submit_fails,
            "preform_fails": preform_fails,
            "exception_fails": exception_fails,
            "late_reg_total": late_reg_total,
            "late_reg_done": late_reg_done,
            "late_reg_failed": late_reg_failed,
            "late_reg_pending": late_reg_pending,
            "late_reg_post_dup": late_reg_post_dup,
            "started_at": progress.get("started_at"),
            "last_updated": progress.get("last_updated"),
            "avg_duration": avg_duration,
            "min_duration": min_duration,
            "max_duration": max_duration,
            "total_bot_time_min": round(total_bot_time / 60, 1),
            "est_remaining_hours": est_remaining_hours,
            "success_rate": round(success_count * 100 / max(1, success_count + failed_count + post_dup_count), 1),
        },
        "rows": rows_data,
    }


EXCEL_RECORDS = None
EXCEL_RAW_DF = None


class DashboardHandler(SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/api/data":
            self.send_json_response()
        elif path == "/api/next-chunk":
            self.handle_next_chunk()
        elif path == "/api/download-excel":
            self.handle_download_excel(parsed.query)
        elif path == "/" or path == "/index.html":
            self.send_dashboard_html()
        else:
            self.send_error(404)

    def do_POST(self):
        if self.path == "/api/start-bot":
            self.handle_start_bot()
        elif self.path == "/api/stop-bot":
            self.handle_stop_bot()
        elif self.path == "/api/add-bot":
            self.handle_add_bot()
        elif self.path == "/api/delete-bot":
            self.handle_delete_bot()
        elif self.path == "/api/reassign-bot":
            self.handle_reassign_bot()
        elif self.path == "/api/reset-bot":
            self.handle_reset_bot()
        else:
            self.send_error(404)

    def _send_json(self, code, obj):
        payload = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def handle_start_bot(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            req = json.loads(body)
            folder = str(req.get("folder", "")).strip()
            start_row = int(req.get("start"))
            end_row = int(req.get("end"))
        except Exception as e:
            self._send_json(400, {"ok": False, "error": f"Bad request: {e}"})
            return

        if not folder or start_row < 0 or end_row < start_row:
            self._send_json(400, {"ok": False, "error": "Invalid folder or range"})
            return

        folders = discover_bot_folders()
        folder_map = {os.path.basename(f): f for f in folders}
        if folder not in folder_map:
            self._send_json(404, {"ok": False, "error": f"Unknown folder: {folder}"})
            return
        folder_path = folder_map[folder]

        with _SPAWN_LOCK:
            prog = load_progress_for(folder_path)
            l_at = prog.get("last_updated") if prog else None
            if _is_recent(l_at, ACTIVE_WINDOW_SECONDS):
                self._send_json(409, {"ok": False, "error": "Bot appears to be active already"})
                return
            if _was_recently_started(folder):
                self._send_json(409, {"ok": False, "error": "Bot was just started — give it a moment"})
                return

            try:
                pid = spawn_bot(folder_path, start_row, end_row)
                RECENTLY_STARTED[folder] = datetime.now()
            except Exception as e:
                self._send_json(500, {"ok": False, "error": f"Spawn failed: {e}"})
                return

        self._send_json(200, {"ok": True, "pid": pid, "folder": folder, "start": start_row, "end": end_row})

    def handle_stop_bot(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            req = json.loads(body)
            folder = str(req.get("folder", "")).strip()
        except Exception as e:
            self._send_json(400, {"ok": False, "error": f"Bad request: {e}"})
            return

        if not folder:
            self._send_json(400, {"ok": False, "error": "Missing folder"})
            return

        folders = discover_bot_folders()
        folder_map = {os.path.basename(f): f for f in folders}
        if folder not in folder_map:
            self._send_json(404, {"ok": False, "error": f"Unknown folder: {folder}"})
            return

        ok, msg = stop_bot(folder, folder_map[folder])
        if not ok:
            self._send_json(500, {"ok": False, "error": msg})
            return
        self._send_json(200, {"ok": True, "folder": folder, "message": msg})

    def handle_add_bot(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            req = json.loads(body) if body.strip() else {}
        except Exception as e:
            self._send_json(400, {"ok": False, "error": f"Bad request: {e}"})
            return

        manual_start = req.get("start")
        manual_end = req.get("end")

        pending_count = None
        with _SPAWN_LOCK:
            config = load_bots_config()

            if manual_start is not None and manual_end is not None:
                try:
                    cmd_start = int(manual_start)
                    cmd_end = int(manual_end)
                except (TypeError, ValueError):
                    self._send_json(400, {"ok": False, "error": "start/end must be integers"})
                    return
                if cmd_start < 0 or cmd_end > (EXCEL_LAST_DATA_ROW + 1) or cmd_start >= cmd_end:
                    self._send_json(400, {
                        "ok": False,
                        "error": f"Invalid range — need 0 <= start < end <= {EXCEL_LAST_DATA_ROW + 1}",
                    })
                    return
                excel_start = cmd_start + 2
                excel_end = cmd_end - 1
            else:
                done_rows = compute_done_rows()
                chunk = pick_next_chunk(done_rows)
                if chunk is None:
                    self._send_json(400, {"ok": False, "error": "No pending rows left — every excel row has already been processed"})
                    return
                cmd_start, cmd_end, excel_start, excel_end, pending_count = chunk

            bot_name = next_bot_name(config)
            try:
                folder_path = create_bot_from_template(bot_name)
            except Exception as e:
                self._send_json(500, {"ok": False, "error": f"Failed to copy template: {e}"})
                return

            config[bot_name] = {"start": cmd_start, "end": cmd_end}
            try:
                save_bots_config(config)
            except Exception as e:
                self._send_json(500, {"ok": False, "error": f"Folder created but config save failed: {e}"})
                return

        span = excel_end - excel_start + 1
        self._send_json(200, {
            "ok": True,
            "bot": bot_name,
            "folder_path": folder_path,
            "start": cmd_start,
            "end": cmd_end,
            "excel_rows": [excel_start, excel_end],
            "row_count": pending_count if pending_count is not None else span,
            "span": span,
            "pending_count": pending_count,
        })

    def handle_delete_bot(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            req = json.loads(body)
            folder = str(req.get("folder", "")).strip()
        except Exception as e:
            self._send_json(400, {"ok": False, "error": f"Bad request: {e}"})
            return

        if not folder:
            self._send_json(400, {"ok": False, "error": "Missing folder"})
            return

        config = load_bots_config()
        if folder not in config:
            self._send_json(404, {"ok": False, "error": f"Unknown bot: {folder}"})
            return

        # Safety: refuse if the bot looks running
        folder_path = os.path.join(PARENT_DIR, folder)

        prog = load_progress_for(folder_path) if os.path.isdir(folder_path) else None
        l_at = prog.get("last_updated") if prog else None
        if _is_recent(l_at, ACTIVE_WINDOW_SECONDS):
            self._send_json(409, {"ok": False, "error": "Bot has recent activity — stop it before deleting"})
            return
        if _was_recently_started(folder):
            self._send_json(409, {"ok": False, "error": "Bot was just started — wait or stop it before deleting"})
            return
        if folder in SPAWNED_PIDS:
            self._send_json(409, {"ok": False, "error": "Dashboard spawned this bot — stop it before deleting"})
            return
        if HAS_PSUTIL and os.path.isdir(folder_path):
            pids = find_bot_pids(folder_path)
            if pids:
                self._send_json(409, {
                    "ok": False,
                    "error": f"Python process(es) {pids} running in this folder — stop them first",
                })
                return

        del config[folder]
        try:
            save_bots_config(config)
        except Exception as e:
            self._send_json(500, {"ok": False, "error": f"Failed to save config: {e}"})
            return

        self._send_json(200, {
            "ok": True,
            "folder": folder,
            "message": "Removed from config; folder and progress data preserved on disk",
        })

    def handle_reset_bot(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            req = json.loads(body)
            folder = str(req.get("folder", "")).strip()
        except Exception as e:
            self._send_json(400, {"ok": False, "error": f"Bad request: {e}"})
            return

        if not folder:
            self._send_json(400, {"ok": False, "error": "Missing folder"})
            return

        with _SPAWN_LOCK:
            config = load_bots_config()
            if folder not in config:
                self._send_json(404, {"ok": False, "error": f"Unknown bot: {folder}"})
                return

            folder_path = os.path.join(PARENT_DIR, folder)
            if not os.path.isdir(folder_path):
                self._send_json(404, {"ok": False, "error": f"Folder missing on disk: {folder}"})
                return

            # Same safety gates as delete/reassign
            prog = load_progress_for(folder_path)
            l_at = prog.get("last_updated") if prog else None
            if _is_recent(l_at, ACTIVE_WINDOW_SECONDS):
                self._send_json(409, {"ok": False, "error": "Bot has recent activity — stop it before resetting"})
                return
            if _was_recently_started(folder):
                self._send_json(409, {"ok": False, "error": "Bot was just started — stop it before resetting"})
                return
            if folder in SPAWNED_PIDS:
                self._send_json(409, {"ok": False, "error": "Dashboard spawned this bot — stop it before resetting"})
                return
            if HAS_PSUTIL:
                pids = find_bot_pids(folder_path)
                if pids:
                    self._send_json(409, {
                        "ok": False,
                        "error": f"Python process(es) {pids} running in this folder — stop them first",
                    })
                    return

            # Archive current progress.json (if any) — preserves work history
            progress_path = os.path.join(folder_path, PROGRESS_REL)
            archive_name = None
            if os.path.exists(progress_path):
                ts = datetime.now().strftime("%Y%m%d-%H%M%S")
                archive_name = f"progress-{ts}.json"
                archive_path = os.path.join(folder_path, "logs", archive_name)
                try:
                    os.rename(progress_path, archive_path)
                except Exception as e:
                    self._send_json(500, {"ok": False, "error": f"Failed to archive progress.json: {e}"})
                    return

            # Clear the range — bot stays in config but has no assignment
            config[folder] = {"start": None, "end": None}
            try:
                save_bots_config(config)
            except Exception as e:
                self._send_json(500, {"ok": False, "error": f"Config save failed: {e}"})
                return

            SPAWNED_PIDS.pop(folder, None)
            RECENTLY_STARTED.pop(folder, None)

        self._send_json(200, {
            "ok": True,
            "folder": folder,
            "archived": archive_name,
            "message": "Range cleared; archived progress preserved on disk",
        })

    def handle_reassign_bot(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
            req = json.loads(body) if body.strip() else {}
            folder = str(req.get("folder", "")).strip()
        except Exception as e:
            self._send_json(400, {"ok": False, "error": f"Bad request: {e}"})
            return

        if not folder:
            self._send_json(400, {"ok": False, "error": "Missing folder"})
            return

        with _SPAWN_LOCK:
            config = load_bots_config()
            if folder not in config:
                self._send_json(404, {"ok": False, "error": f"Unknown bot: {folder}"})
                return

            folder_path = os.path.join(PARENT_DIR, folder)
            if not os.path.isdir(folder_path):
                self._send_json(404, {"ok": False, "error": f"Folder missing on disk: {folder}"})
                return

            # Safety: refuse if bot is running
            prog = load_progress_for(folder_path)
            l_at = prog.get("last_updated") if prog else None
            if _is_recent(l_at, ACTIVE_WINDOW_SECONDS):
                self._send_json(409, {"ok": False, "error": "Bot has recent activity — stop it before reassigning"})
                return
            if _was_recently_started(folder):
                self._send_json(409, {"ok": False, "error": "Bot was just started — stop it before reassigning"})
                return
            if folder in SPAWNED_PIDS:
                self._send_json(409, {"ok": False, "error": "Dashboard spawned this bot — stop it before reassigning"})
                return
            if HAS_PSUTIL:
                pids = find_bot_pids(folder_path)
                if pids:
                    self._send_json(409, {
                        "ok": False,
                        "error": f"Python process(es) {pids} running in this folder — stop them first",
                    })
                    return

            # Determine new range — manual override or auto-pick (excluding self)
            manual_start = req.get("start")
            manual_end = req.get("end")
            pending_count = None
            if manual_start is not None and manual_end is not None:
                try:
                    cmd_start = int(manual_start)
                    cmd_end = int(manual_end)
                except (TypeError, ValueError):
                    self._send_json(400, {"ok": False, "error": "start/end must be integers"})
                    return
                if cmd_start < 0 or cmd_end > (EXCEL_LAST_DATA_ROW + 1) or cmd_start >= cmd_end:
                    self._send_json(400, {
                        "ok": False,
                        "error": f"Invalid range — need 0 <= start < end <= {EXCEL_LAST_DATA_ROW + 1}",
                    })
                    return
                excel_start = cmd_start + 2
                excel_end = cmd_end - 1
            else:
                done_rows = compute_done_rows()
                chunk = pick_next_chunk(done_rows)
                if chunk is None:
                    self._send_json(400, {"ok": False, "error": "No pending rows left to reassign"})
                    return
                cmd_start, cmd_end, excel_start, excel_end, pending_count = chunk

            # Archive old progress.json (if any), then remove so bot starts fresh
            progress_path = os.path.join(folder_path, PROGRESS_REL)
            archive_name = None
            if os.path.exists(progress_path):
                ts = datetime.now().strftime("%Y%m%d-%H%M%S")
                archive_name = f"progress-{ts}.json"
                archive_path = os.path.join(folder_path, "logs", archive_name)
                try:
                    os.rename(progress_path, archive_path)
                except Exception as e:
                    self._send_json(500, {"ok": False, "error": f"Failed to archive progress.json: {e}"})
                    return

            # Update config
            config[folder] = {"start": cmd_start, "end": cmd_end}
            try:
                save_bots_config(config)
            except Exception as e:
                self._send_json(500, {"ok": False, "error": f"Config save failed: {e}"})
                return

            # Clear in-memory traces of the previous run
            SPAWNED_PIDS.pop(folder, None)
            RECENTLY_STARTED.pop(folder, None)

        span = excel_end - excel_start + 1
        self._send_json(200, {
            "ok": True,
            "folder": folder,
            "start": cmd_start,
            "end": cmd_end,
            "excel_rows": [excel_start, excel_end],
            "row_count": pending_count if pending_count is not None else span,
            "span": span,
            "pending_count": pending_count,
            "archived": archive_name,
        })

    def handle_next_chunk(self):
        config = load_bots_config()
        done_rows = compute_done_rows()
        chunk = pick_next_chunk(done_rows)
        bot_name = next_bot_name(config)
        if chunk is None:
            self._send_json(200, {
                "ok": True,
                "free": False,
                "bot": bot_name,
                "start": None,
                "end": None,
            })
            return
        cs, ce, es, ee, pending_count = chunk
        span = ee - es + 1
        self._send_json(200, {
            "ok": True,
            "free": True,
            "bot": bot_name,
            "start": cs,
            "end": ce,
            "excel_start": es,
            "excel_end": ee,
            "row_count": pending_count,
            "span": span,
            "pending_count": pending_count,
        })

    def handle_download_excel(self, query):
        params = parse_qs(query or "")
        status_filter = (params.get("status", ["all"])[0] or "all").lower()
        bot_filter = (params.get("bot", [""])[0] or "").strip()

        config = load_bots_config()

        # Build per-row data the same way /api/data does
        folders = discover_bot_folders()
        merged, _per_bot, _resume = merge_progress(folders)
        data = build_dashboard_data(EXCEL_RECORDS, merged)
        rows = data["rows"]

        # Apply status filter
        if status_filter and status_filter != "all":
            rows = [r for r in rows if r.get("status") == status_filter]

        # Apply bot filter (uses range from config, +2/-1 offsets per spec)
        if bot_filter:
            cfg = config.get(bot_filter)
            if cfg is None:
                self._send_json(404, {"ok": False, "error": f"Unknown bot: {bot_filter}"})
                return
            s, e = cfg["start"], cfg["end"]
            lo, hi = s + 2, e - 1
            rows = [r for r in rows if lo <= r["excel_row"] <= hi]

        # Pre-compute which bot actually processed each row (from progress data)
        row_to_bot = compute_row_to_bot_map()

        # Build the workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Pehchan Export"
        headers = [
            "रजि क्रमांक", "Year", "जन्म दिनांक", "रजि दिनांक", "लिंग",
            "बालक का नाम", "पिता का नाम", "माता का नाम", "हिन्दू मुस्लिम",
            "पता/जन्म स्थान", "पता",
            "STATUS", "ROW", "ERROR", "BOT",
        ]
        ws.append(headers)

        status_label_map = {
            "success": "ENTERED",
            "post_dup": "POST-DUP",
            "skipped": "PRE-DUP",
            "failed": "FAILED",
            "pending": "PENDING",
        }

        # Resolve which raw Excel columns are present for the address-like fields
        raw_df = EXCEL_RAW_DF
        addr_birthplace_col = None
        addr_full_col = None
        if raw_df is not None:
            for c in ("पता/जन्म स्थान", "पता जन्म स्थान"):
                if c in raw_df.columns:
                    addr_birthplace_col = c
                    break
            if "पता" in raw_df.columns:
                addr_full_col = "पता"

        for r in rows:
            er = r["excel_row"]
            addr_birthplace = ""
            addr_full = ""
            if raw_df is not None:
                df_idx = er - 2  # pandas is 0-indexed; excel row 2 → df row 0
                if 0 <= df_idx < len(raw_df):
                    row_obj = raw_df.iloc[df_idx]
                    if addr_birthplace_col:
                        v = row_obj.get(addr_birthplace_col, "")
                        addr_birthplace = "" if pd.isna(v) else str(v).strip()
                    if addr_full_col:
                        v = row_obj.get(addr_full_col, "")
                        addr_full = "" if pd.isna(v) else str(v).strip()
            # Fall back to truncated address from EXCEL_RECORDS if neither column was found
            if not addr_birthplace and not addr_full:
                addr_full = r.get("address", "")

            ws.append([
                r["reg_num"], r["year"], r["birth_date"], r["reg_date"], r["gender"],
                r["child"], r["father"], r["mother"], r["religion"],
                addr_birthplace, addr_full,
                status_label_map.get(r["status"], r["status"].upper()),
                er,
                r.get("error") or "",
                row_to_bot.get(er, ""),
            ])

        # Light column widths so the file is readable on open
        widths = [12, 8, 12, 12, 8, 22, 22, 22, 12, 40, 40, 10, 8, 30, 12]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(ord("A") + i - 1)].width = w

        buf = BytesIO()
        wb.save(buf)
        payload = buf.getvalue()

        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        suffix = status_filter
        if bot_filter:
            suffix = f"{bot_filter}-{status_filter}"
        fname = f"pehchan-entries-{suffix}-{ts}.xlsx"

        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(payload)

    def send_json_response(self):
        folders = discover_bot_folders()
        merged, per_bot, resume_commands = merge_progress(folders)
        data = build_dashboard_data(EXCEL_RECORDS, merged)
        data["bots"] = per_bot
        data["resume_commands"] = resume_commands

        # Sheet-wide pending count for the Work Dispatcher: total data rows
        # minus anything any bot (including orphans) has finished.
        done = compute_done_rows()
        total_data_rows = EXCEL_LAST_DATA_ROW - EXCEL_FIRST_DATA_ROW + 1
        in_range_done = sum(1 for r in done if EXCEL_FIRST_DATA_ROW <= r <= EXCEL_LAST_DATA_ROW)
        data["stats"]["global_pending"] = total_data_rows - in_range_done

        payload = json.dumps(data, ensure_ascii=False)
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(payload.encode("utf-8"))

    def send_dashboard_html(self):
        html = get_dashboard_html()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html.encode("utf-8"))


def get_dashboard_html():
    return '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pehchan Unified Tracker</title>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&family=DM+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{
  --bg:#F8F7F4;--card:#FFF;--hover:#F1F0ED;--border:#E5E3DE;--blt:#EDEBE7;
  --txt:#1A1A18;--t2:#6B6960;--t3:#9C9889;
  --ok:#2D7A3A;--ok-bg:#E8F5EA;--ok-bd:#A8D5AE;
  --err:#C53030;--err-bg:#FEE8E8;--err-bd:#F5A8A8;
  --wrn:#B7791F;--wrn-bg:#FEFCE8;--wrn-bd:#F0D58C;
  --pnd:#6B6960;--pnd-bg:#F1F0ED;--pnd-bd:#E5E3DE;
  --blu:#2563EB;--blu-bg:#EFF6FF;
  --prp:#9333EA;--prp-bg:#F3E8FF;
  --idle:#9C9889;
  --mn:'JetBrains Mono',monospace;--sn:'DM Sans',sans-serif
}
body{font-family:var(--sn);background:var(--bg);color:var(--txt);line-height:1.5}

.hdr{background:var(--card);border-bottom:1px solid var(--border);padding:18px 32px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:100}
.hdr h1{font-family:var(--mn);font-size:17px;font-weight:700;letter-spacing:-.5px}
.hdr .sub{font-size:12px;color:var(--t2);margin-top:1px}
.hr{display:flex;align-items:center;gap:14px;font-size:12px;color:var(--t2)}
.dot{width:8px;height:8px;background:var(--ok);border-radius:50%;animation:p 2s infinite}
@keyframes p{0%,100%{opacity:1}50%{opacity:.4}}

.w{max-width:1440px;margin:0 auto;padding:20px 32px}

.bill{background:var(--ok-bg);border:1px solid var(--ok-bd);border-radius:8px;padding:16px 24px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center}
.bill .bl{font-size:13px;font-weight:600;color:var(--ok)}
.bill .bv{font-family:var(--mn);font-size:32px;font-weight:700;color:var(--ok)}
.bill .bb{font-size:12px;color:var(--t2)}

.sg{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:20px}
.sc{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:14px 18px}
.sc .lb{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--t3);margin-bottom:4px}
.sc .vl{font-family:var(--mn);font-size:26px;font-weight:700;line-height:1}
.sc .dt{font-size:11px;color:var(--t2);margin-top:5px}
.sc.co .vl{color:var(--ok)}.sc.ce .vl{color:var(--err)}.sc.cw .vl{color:var(--wrn)}.sc.cp .vl{color:var(--t2)}.sc.cv .vl{color:var(--prp)}

.pg{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:18px 24px;margin-bottom:20px}
.pb{height:28px;background:var(--pnd-bg);border-radius:5px;overflow:hidden;display:flex;margin-top:10px}
.pb>div{height:100%;transition:width .6s}
.pb .so{background:var(--ok)}.pb .sd{background:var(--prp)}.pb .ss{background:var(--wrn)}.pb .se{background:var(--err)}
.lg{display:flex;gap:20px;margin-top:10px;font-size:12px}
.li{display:flex;align-items:center;gap:5px}
.ld{width:10px;height:10px;border-radius:3px}

.ps{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:20px}
.pn{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:18px 22px}
.pn h3{font-size:13px;font-weight:600;margin-bottom:14px}
.br{display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid var(--blt);font-size:12px}
.br:last-child{border-bottom:none}
.br .ct{font-family:var(--mn);font-weight:600;font-size:13px}

.bots-wrap{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:18px 22px;margin-bottom:20px}
.bots-wrap h3{font-size:13px;font-weight:600;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center;gap:10px}
.bots-wrap h3 .meta{font-weight:400;color:var(--t3);font-size:11px}
.bots-wrap h3 .rt{margin-left:auto;display:flex;align-items:center;gap:10px}

.disp-summary{display:flex;gap:12px;margin-bottom:14px}
.disp-metric{flex:1;background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:14px 18px;text-align:center}
.disp-metric .v{font-family:var(--mn);font-size:26px;font-weight:700;line-height:1;color:var(--txt)}
.disp-metric .l{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--t3);margin-top:6px}
.disp-metric.pending .v{color:var(--wrn)}
.disp-metric.failed .v{color:var(--err)}
.disp-metric.total .v{color:var(--blu)}
.disp-empty{font-size:12px;color:var(--t2);padding:14px 12px;text-align:center;font-style:italic}
.disp-done{font-size:13px;color:var(--ok);padding:14px 12px;text-align:center;font-weight:600;font-family:var(--mn)}
.disp-bot-row{display:flex;align-items:center;justify-content:space-between;padding:10px 14px;border:1px solid var(--border);border-radius:6px;background:var(--bg);margin-bottom:6px}
.disp-bot-row:last-child{margin-bottom:0}
.disp-bot-row .info{display:flex;flex-direction:column;gap:2px}
.disp-bot-row .name{font-family:var(--mn);font-size:12px;font-weight:700;color:var(--txt)}
.disp-bot-row .sub{font-size:10px;color:var(--t3);font-family:var(--mn)}
.assignbtn{flex-shrink:0;padding:7px 16px;border:1px solid var(--ok);border-radius:4px;background:var(--ok);color:#fff;font-size:11px;font-family:var(--sn);font-weight:700;letter-spacing:.4px;cursor:pointer;transition:all .15s;text-transform:uppercase}
.assignbtn:hover:not(:disabled){background:#1f5c2a;border-color:#1f5c2a}
.assignbtn:disabled{opacity:.55;cursor:wait}
.assignbtn.err{background:var(--err);border-color:var(--err)}
.addbot{padding:6px 14px;border:1px solid var(--blu);border-radius:4px;background:var(--blu-bg);color:var(--blu);font-size:11px;font-family:var(--sn);font-weight:700;letter-spacing:.3px;cursor:pointer;transition:all .15s;text-transform:uppercase}
.addbot:hover:not(:disabled){background:var(--blu);color:#fff}
.addbot:disabled{opacity:.6;cursor:wait}
.addbot.err{background:var(--err);border-color:var(--err);color:#fff}

.modal{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.45);z-index:1000;display:none;align-items:center;justify-content:center}
.modal.open{display:flex}
.modal-content{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:22px 26px;min-width:380px;max-width:90vw;box-shadow:0 8px 32px rgba(0,0,0,.25)}
.modal-content h3{font-size:14px;font-weight:700;margin-bottom:6px}
.modal-content p{font-size:12px;color:var(--t2);margin-bottom:14px}
.modal-content p code{font-family:var(--mn);background:var(--bg);padding:1px 6px;border-radius:3px;color:var(--txt);font-weight:600}
.form-row{display:flex;gap:12px;margin-bottom:10px}
.form-row label{flex:1;display:flex;flex-direction:column;gap:5px;font-size:10px;color:var(--t3);font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.form-row input{padding:7px 10px;border:1px solid var(--border);border-radius:4px;font-family:var(--mn);font-size:13px;outline:none;color:var(--txt);background:var(--card)}
.form-row input:focus{border-color:var(--blu)}
.form-hint{font-size:11px;color:var(--t3);margin-bottom:16px;font-family:var(--mn);min-height:14px}
.form-hint.err{color:var(--err)}
.modal-buttons{display:flex;gap:8px;justify-content:flex-end}
.modal-buttons .cancel{padding:7px 16px;border:1px solid var(--border);border-radius:4px;background:var(--card);color:var(--t2);font-size:11px;font-family:var(--sn);font-weight:600;cursor:pointer;text-transform:uppercase;letter-spacing:.3px}
.modal-buttons .cancel:hover{background:var(--hover)}
.bg-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px}
.bcard{position:relative;border:1px solid var(--border);border-radius:6px;padding:12px 14px;background:var(--bg);cursor:pointer;transition:all .15s;user-select:none}
.delbot{position:absolute;top:5px;right:5px;width:20px;height:20px;border:1px solid transparent;border-radius:50%;background:transparent;color:var(--t3);font-size:14px;line-height:1;cursor:pointer;font-weight:700;display:flex;align-items:center;justify-content:center;padding:0;transition:all .15s}
.delbot:hover{background:var(--err);color:#fff;border-color:var(--err)}
.resetbtn{display:inline-block;margin-top:8px;padding:4px 12px;border:1px solid var(--border);border-radius:4px;background:transparent;color:var(--t2);font-size:10px;font-family:var(--sn);font-weight:600;letter-spacing:.4px;cursor:pointer;transition:all .15s;text-transform:uppercase}
.resetbtn:hover{background:var(--hover);color:var(--txt);border-color:var(--t3)}
.bcard .stats + .resetbtn{margin-top:8px}
.bcard:hover{background:var(--hover);border-color:var(--blu)}
.bcard.selected{border-color:var(--blu);background:var(--blu-bg);box-shadow:0 0 0 2px rgba(37,99,235,.15)}
.bcard.locked{cursor:default}
.bcard.locked:hover{background:var(--bg);border-color:var(--border)}
.bcard .top{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.bcard .sdot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.bcard .sdot.on{background:var(--ok);box-shadow:0 0 0 3px rgba(45,122,58,.18);animation:p 2s infinite}
.bcard .sdot.off{background:var(--idle)}
.bcard .fn{font-family:var(--mn);font-size:11px;font-weight:600;color:var(--txt);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bcard .rng{font-family:var(--mn);font-size:10px;color:var(--t3);margin-bottom:8px}
.bcard .stats{display:grid;grid-template-columns:repeat(5,1fr);gap:6px;margin-bottom:8px}
.bcard .st{text-align:center;padding:5px 2px;background:var(--card);border:1px solid var(--border);border-radius:4px}
.bcard .st .v{font-family:var(--mn);font-size:14px;font-weight:700;line-height:1}
.bcard .st .l{font-size:9px;color:var(--t3);text-transform:uppercase;letter-spacing:.4px;margin-top:2px}
.bcard .st.s .v{color:var(--ok)}.bcard .st.f .v{color:var(--err)}.bcard .st.d .v{color:var(--prp)}.bcard .st.k .v{color:var(--wrn)}.bcard .st.p .v{color:var(--t2)}
.bcard .la{font-size:10px;color:var(--t2);font-family:var(--mn)}
.bcard .la b{color:var(--txt);font-weight:600}

.rlist{display:flex;flex-direction:column;gap:8px}
.rrow{display:flex;align-items:center;gap:14px;padding:10px 12px;border:1px solid var(--border);border-radius:6px;background:var(--bg)}
.rleft{flex-shrink:0;min-width:260px;max-width:320px}
.rfn{font-family:var(--mn);font-size:11px;font-weight:600;color:var(--txt);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.rlr{font-family:var(--mn);font-size:10px;color:var(--t3);margin-top:2px}
.rright{flex:1;display:flex;align-items:center;gap:8px;min-width:0}
.rbadge{display:inline-flex;align-items:center;padding:5px 12px;border-radius:4px;font-size:11px;font-weight:700;font-family:var(--mn);letter-spacing:.5px}
.rbadge.running{background:var(--ok-bg);color:var(--ok);border:1px solid var(--ok-bd)}
.rbadge.running::before{content:"";display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--ok);margin-right:7px;animation:p 2s infinite}
.rbadge.completed{background:var(--prp-bg);color:var(--prp);border:1px solid #D8B4FE}
.rbadge.unknown{background:var(--pnd-bg);color:var(--pnd);border:1px solid var(--pnd-bd)}
.cmdbox{flex:1;font-family:var(--mn);font-size:11px;background:var(--card);border:1px solid var(--border);border-radius:4px;padding:7px 10px;color:var(--txt);overflow-x:auto;white-space:nowrap;min-width:0}
.cpy{flex-shrink:0;padding:7px 14px;border:1px solid var(--border);border-radius:4px;background:var(--card);font-size:11px;font-family:var(--sn);font-weight:600;cursor:pointer;color:var(--t2);transition:all .15s}
.cpy:hover{background:var(--txt);color:#fff;border-color:var(--txt)}
.cpy.done{background:var(--ok);color:#fff;border-color:var(--ok)}
.startbtn{flex-shrink:0;padding:7px 16px;border:1px solid var(--ok);border-radius:4px;background:var(--ok);color:#fff;font-size:11px;font-family:var(--sn);font-weight:700;letter-spacing:.4px;cursor:pointer;transition:all .15s;text-transform:uppercase}
.startbtn:hover:not(:disabled){background:#1f5c2a;border-color:#1f5c2a}
.startbtn:disabled{opacity:.55;cursor:wait}
.startbtn.err{background:var(--err);border-color:var(--err)}
.startbtn.err:disabled{opacity:1}
.stopbtn{flex-shrink:0;padding:7px 16px;border:1px solid var(--err);border-radius:4px;background:var(--err);color:#fff;font-size:11px;font-family:var(--sn);font-weight:700;letter-spacing:.4px;cursor:pointer;transition:all .15s;text-transform:uppercase}
.stopbtn:hover:not(:disabled){background:#8b2020;border-color:#8b2020}
.stopbtn:disabled{opacity:.55;cursor:wait}

.clrbtn{display:none;margin-left:10px;padding:3px 9px;border:1px solid var(--blu);border-radius:4px;background:var(--blu-bg);color:var(--blu);font-size:11px;font-family:var(--sn);font-weight:600;cursor:pointer}
.clrbtn:hover{background:var(--blu);color:#fff}
.clrbtn.on{display:inline-flex;align-items:center;gap:5px}

.ts{background:var(--card);border:1px solid var(--border);border-radius:8px;overflow:hidden}
.th{padding:14px 22px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center}
.th h3{font-size:13px;font-weight:600}
.tc{display:flex;gap:6px;align-items:center}
.fb{padding:4px 10px;border:1px solid var(--border);border-radius:4px;background:var(--bg);font-size:11px;font-family:var(--sn);cursor:pointer;color:var(--t2);transition:all .15s}
.fb:hover{background:var(--hover)}.fb.on{background:var(--txt);color:#fff;border-color:var(--txt)}
.fb.dl{background:var(--blu-bg);color:var(--blu);border-color:var(--blu);font-weight:700;margin-left:6px}
.fb.dl:hover{background:var(--blu);color:#fff;border-color:var(--blu)}
.si{padding:5px 10px;border:1px solid var(--border);border-radius:4px;font-size:12px;font-family:var(--sn);width:180px;outline:none}
.si:focus{border-color:var(--blu)}
.tw{max-height:620px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
thead{position:sticky;top:0;z-index:10}
th{background:var(--bg);padding:8px 10px;text-align:left;font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--t3);border-bottom:1px solid var(--border);white-space:nowrap}
td{padding:7px 10px;border-bottom:1px solid var(--blt);white-space:nowrap}
tr:hover td{background:var(--hover)}
.bd{display:inline-flex;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:600;font-family:var(--mn)}
.bd.success{background:var(--ok-bg);color:var(--ok);border:1px solid var(--ok-bd)}
.bd.failed{background:var(--err-bg);color:var(--err);border:1px solid var(--err-bd)}
.bd.skipped{background:var(--wrn-bg);color:var(--wrn);border:1px solid var(--wrn-bd)}
.bd.pending{background:var(--pnd-bg);color:var(--pnd);border:1px solid var(--pnd-bd)}
.bd.post_dup{background:var(--prp-bg);color:var(--prp);border:1px solid #D8B4FE}
.lt{display:inline-flex;padding:1px 5px;border-radius:3px;font-size:9px;font-weight:600;background:var(--prp-bg);color:var(--prp);margin-left:3px}
.m{font-family:var(--mn);font-size:11px}
.et{color:var(--err);font-size:10px}
.dtt{color:var(--t3);font-size:10px;max-width:200px;overflow:hidden;text-overflow:ellipsis}
</style>
</head>
<body>

<div class="hdr">
  <div><h1>PEHCHAN UNIFIED TRACKER</h1><div class="sub">Merged from all bot instances</div></div>
  <div class="hr"><div class="dot"></div><span id="lu">Loading...</span><span>|</span><span id="tm">—</span></div>
</div>

<div class="w">

<div class="bill">
  <div><div class="bl">TOTAL BILLABLE ENTRIES (CA Invoice)</div><div class="bb" id="bk">—</div></div>
  <div class="bv" id="bv">—</div>
</div>

<div class="sg">
  <div class="sc"><div class="lb">Total Records</div><div class="vl" id="sT">—</div><div class="dt" id="sF">—</div></div>
  <div class="sc co"><div class="lb">Entered Successfully</div><div class="vl" id="sO">—</div><div class="dt" id="sR">—</div></div>
  <div class="sc cv"><div class="lb">Post-Submit Duplicate</div><div class="vl" id="sD">—</div><div class="dt">Already existed on portal</div></div>
  <div class="sc ce"><div class="lb">Failed</div><div class="vl" id="sE">—</div><div class="dt" id="sED">—</div></div>
  <div class="sc cw"><div class="lb">Pre-form Duplicate</div><div class="vl" id="sS">—</div><div class="dt">Detected before form load</div></div>
  <div class="sc cp"><div class="lb">Remaining</div><div class="vl" id="sP">—</div><div class="dt" id="sET">—</div></div>
</div>

<div class="pg">
  <div style="display:flex;justify-content:space-between;align-items:center">
    <span style="font-weight:600;font-size:13px">Overall Progress</span>
    <span id="pp" class="m" style="font-size:13px;font-weight:600">0%</span>
  </div>
  <div class="pb">
    <div class="so" id="bo" style="width:0%"></div>
    <div class="sd" id="bd" style="width:0%"></div>
    <div class="ss" id="bs" style="width:0%"></div>
    <div class="se" id="be" style="width:0%"></div>
  </div>
  <div class="lg">
    <div class="li"><div class="ld" style="background:var(--ok)"></div>Entered</div>
    <div class="li"><div class="ld" style="background:var(--prp)"></div>Post-Submit Dup</div>
    <div class="li"><div class="ld" style="background:var(--wrn)"></div>Pre-form Dup</div>
    <div class="li"><div class="ld" style="background:var(--err)"></div>Failed</div>
    <div class="li"><div class="ld" style="background:var(--pnd-bg);border:1px solid var(--border)"></div>Pending</div>
  </div>
</div>

<div class="ps">
  <div class="pn">
    <h3>Failure Breakdown</h3>
    <div class="br"><span>Form Fill Failed</span><span class="ct" id="fF" style="color:var(--err)">—</span></div>
    <div class="br"><span>Submit Failed</span><span class="ct" id="fS" style="color:var(--err)">—</span></div>
    <div class="br"><span>Pre-form Failed</span><span class="ct" id="fP" style="color:var(--err)">—</span></div>
    <div class="br"><span>Exception</span><span class="ct" id="fX" style="color:var(--err)">—</span></div>
  </div>
  <div class="pn">
    <h3>Late Registration Records (gap &gt;20 days)</h3>
    <div class="br"><span>Total in Excel</span><span class="ct" id="lT" style="color:var(--prp)">—</span></div>
    <div class="br"><span>Entered</span><span class="ct" id="lD" style="color:var(--ok)">—</span></div>
    <div class="br"><span>Post-Submit Dup</span><span class="ct" id="lP" style="color:var(--prp)">—</span></div>
    <div class="br"><span>Failed</span><span class="ct" id="lF" style="color:var(--err)">—</span></div>
    <div class="br"><span>Pending</span><span class="ct" id="lR" style="color:var(--t2)">—</span></div>
  </div>
  <div class="pn">
    <h3>Bot Performance</h3>
    <div class="br"><span>Avg time/entry</span><span class="ct m" id="tA">—</span></div>
    <div class="br"><span>Fastest</span><span class="ct m" id="tN">—</span></div>
    <div class="br"><span>Slowest</span><span class="ct m" id="tX">—</span></div>
    <div class="br"><span>Total bot runtime</span><span class="ct m" id="tT">—</span></div>
    <div class="br"><span>Est. remaining</span><span class="ct m" id="tE">—</span></div>
  </div>
</div>

<div class="ps" style="grid-template-columns:1fr 1fr">
  <div class="pn">
    <h3>Session</h3>
    <div class="br"><span>Started at</span><span class="m" id="xS">—</span></div>
    <div class="br"><span>Last activity</span><span class="m" id="xL">—</span></div>
  </div>
  <div class="pn">
    <h3>Processed vs Pending</h3>
    <div class="br"><span>Processed (all categories)</span><span class="ct" id="xT" style="color:var(--blu)">—</span></div>
    <div class="br"><span>Yet to process</span><span class="ct" id="xP" style="color:var(--t2)">—</span></div>
  </div>
</div>

<div class="bots-wrap">
  <h3>Work Dispatcher <span class="meta" id="dispMeta">—</span></h3>
  <div class="disp-summary">
    <div class="disp-metric pending"><div class="v" id="dispPending">—</div><div class="l">Entries Pending</div></div>
    <div class="disp-metric failed"><div class="v" id="dispFailed">—</div><div class="l">Entries Failed</div></div>
    <div class="disp-metric total"><div class="v" id="dispTotal">—</div><div class="l">Total in Sheet</div></div>
  </div>
  <div id="dispList"></div>
</div>

<div class="bots-wrap">
  <h3>Per-Bot Status <span class="rt"><span class="meta" id="botMeta">—</span><button class="addbot" onclick="openAddBotModal()">+ Add Bot</button></span></h3>
  <div class="bg-grid" id="botGrid"></div>
</div>

<div class="bots-wrap">
  <h3>Resume Commands <span class="meta" id="resumeMeta">—</span></h3>
  <div class="rlist" id="resumeList"></div>
</div>

<div class="ts">
  <div class="th">
    <h3>All Records <span id="tC" style="color:var(--t3);font-weight:400"></span><button id="clrBot" class="clrbtn" onclick="clearBotFilter()">× <span id="clrBotLbl">Show All</span></button></h3>
    <div class="tc">
      <input type="text" class="si" placeholder="Search reg#, name, row..." id="sb" oninput="fl()">
      <button class="fb on" onclick="sf('all',this)">All</button>
      <button class="fb" onclick="sf('success',this)">Entered</button>
      <button class="fb" onclick="sf('post_dup',this)">Post-Dup</button>
      <button class="fb" onclick="sf('failed',this)">Failed</button>
      <button class="fb" onclick="sf('skipped',this)">Pre-Dup</button>
      <button class="fb" onclick="sf('pending',this)">Pending</button>
      <button class="fb dl" onclick="downloadExcel()" title="Download the currently filtered rows as .xlsx">⤓ Download</button>
    </div>
  </div>
  <div class="tw">
    <table>
      <thead><tr>
        <th>Row</th><th>Reg #</th><th>Year</th><th>Status</th><th>Duration</th>
        <th>Child</th><th>Father</th><th>Mother</th><th>Gender</th>
        <th>Birth Date</th><th>Reg Date</th><th>Gap</th><th>Religion</th>
        <th>Error</th><th>Detail</th><th>Timestamp</th>
      </tr></thead>
      <tbody id="tb"></tbody>
    </table>
  </div>
</div>

</div>

<script>
let R=[],cf='all',BOTS=[],BOT_FILTER=null;
function fm(i){if(!i)return'—';try{return new Date(i).toLocaleString('en-IN',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:true})}catch(e){return i}}
function fmShort(i){if(!i)return'—';try{return new Date(i).toLocaleString('en-IN',{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit',hour12:true})}catch(e){return i}}
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function sf(f,b){cf=f;document.querySelectorAll('.fb:not(.dl)').forEach(x=>x.classList.remove('on'));b.classList.add('on');fl()}
function downloadExcel(){
  const params=new URLSearchParams();
  params.set('status',cf||'all');
  if(BOT_FILTER&&BOT_FILTER.folder)params.set('bot',BOT_FILTER.folder);
  const a=document.createElement('a');
  a.href='/api/download-excel?'+params.toString();
  a.style.display='none';
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>a.remove(),1500);
}
function fl(){
  const q=document.getElementById('sb').value.toLowerCase(),tb=document.getElementById('tb');
  let h='',c=0;
  for(const r of R){
    if(BOT_FILTER&&(r.excel_row<BOT_FILTER.start||r.excel_row>BOT_FILTER.end))continue;
    if(cf!=='all'&&r.status!==cf)continue;
    if(q){const s=(r.reg_num+' '+r.child+' '+r.father+' '+r.mother+' '+r.excel_row+' '+r.address).toLowerCase();if(!s.includes(q))continue}
    c++;
    const sl={success:'ENTERED',failed:'FAILED',skipped:'PRE-DUP',pending:'PENDING',post_dup:'POST-DUP'}[r.status]||r.status;
    h+='<tr><td class="m">'+r.excel_row+'</td><td class="m">'+r.reg_num+'</td><td class="m">'+r.year+'</td>'
      +'<td><span class="bd '+r.status+'">'+sl+'</span>'+(r.is_late_reg?'<span class="lt">LATE</span>':'')+'</td>'
      +'<td class="m">'+(r.duration>0?r.duration+'s':'')+'</td>'
      +'<td>'+r.child+'</td><td>'+r.father+'</td><td>'+r.mother+'</td><td>'+r.gender+'</td>'
      +'<td class="m">'+r.birth_date+'</td><td class="m">'+r.reg_date+'</td>'
      +'<td class="m">'+(r.gap_days!==null?r.gap_days+'d':'')+'</td>'
      +'<td>'+r.religion+'</td>'
      +'<td>'+(r.error?'<span class="et">'+r.error+'</span>':'')+'</td>'
      +'<td>'+(r.detail?'<span class="dtt" title="'+(r.detail||'').replace(/"/g,'&quot;')+'">'+r.detail+'</span>':'')+'</td>'
      +'<td class="m" style="font-size:10px">'+(r.timestamp?fm(r.timestamp):'')+'</td></tr>';
  }
  tb.innerHTML=h;
  document.getElementById('tC').textContent='('+c+' shown)';
}
function renderDispatcher(stats, bots){
  const pending=stats.pending||0;
  const failed=stats.failed||0;
  const total=stats.total||0;
  const gp=stats.global_pending!=null?stats.global_pending:(pending+failed);
  document.getElementById('dispPending').textContent=pending.toLocaleString();
  document.getElementById('dispFailed').textContent=failed.toLocaleString();
  document.getElementById('dispTotal').textContent=total.toLocaleString();
  const meta=document.getElementById('dispMeta');
  const list=document.getElementById('dispList');
  if(gp<=0){
    meta.textContent='nothing left to do';
    list.innerHTML='<div class="disp-done">✓ All entries processed!</div>';
    return;
  }
  const available=(bots||[]).filter(b=>b.pending===0).sort((a,b)=>{
    const ka=parseInt((a.folder.match(/Bot-(\\d+)/)||[])[1]||0,10);
    const kb=parseInt((b.folder.match(/Bot-(\\d+)/)||[])[1]||0,10);
    return ka-kb;
  });
  if(!available.length){
    meta.textContent=gp.toLocaleString()+' to dispatch';
    list.innerHTML='<div class="disp-empty">All bots are busy — wait for one to finish or add a new bot.</div>';
    return;
  }
  meta.textContent=available.length+' bot(s) idle • '+gp.toLocaleString()+' to dispatch';
  const chunk=Math.min(200,gp);
  const label=(gp<200)?('Assign all ('+gp+' remaining)'):'Assign next 200';
  let h='';
  for(const b of available){
    const lastDone=b.last_processed_row?(' • last processed row '+b.last_processed_row):'';
    h+='<div class="disp-bot-row">'
      +'<div class="info"><div class="name">'+esc(b.folder)+'</div>'
      +'<div class="sub">finished current assignment'+lastDone+'</div></div>'
      +'<button class="assignbtn" data-folder="'+esc(b.folder)+'" onclick="assignWork(this)">'+esc(label)+'</button>'
      +'</div>';
  }
  list.innerHTML=h;
}
async function assignWork(btn){
  const folder=btn.dataset.folder;
  const orig=btn.textContent;
  btn.disabled=true;btn.textContent='Assigning...';
  try{
    const r=await fetch('/api/reassign-bot',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({folder:folder})});
    const d=await r.json();
    if(r.ok&&d.ok){
      btn.textContent='Assigned '+d.row_count+' rows';
      if(BOT_FILTER&&BOT_FILTER.folder===folder)clearBotFilter();
      setTimeout(ld,800);
    }else{
      btn.classList.add('err');btn.textContent='Failed';btn.title=d.error||('HTTP '+r.status);
      setTimeout(()=>{btn.classList.remove('err');btn.disabled=false;btn.textContent=orig;btn.removeAttribute('title')},4000);
    }
  }catch(e){
    btn.classList.add('err');btn.textContent='Error';btn.title=String(e);
    setTimeout(()=>{btn.classList.remove('err');btn.disabled=false;btn.textContent=orig;btn.removeAttribute('title')},4000);
  }
}
function renderResume(items){
  const list=document.getElementById('resumeList');
  const meta=document.getElementById('resumeMeta');
  if(!items||!items.length){list.innerHTML='<div style="color:var(--t2);font-size:12px">No bots detected.</div>';meta.textContent='0 bots';return}
  const running=items.filter(i=>i.status==='running').length;
  const completed=items.filter(i=>i.status==='completed').length;
  const torun=items.length-running-completed;
  meta.textContent=items.length+' bots • '+running+' running • '+completed+' completed • '+torun+' to resume';
  let h='';
  for(const r of items){
    let right='';
    if(r.status==='running'){
      right='<span class="rbadge running">RUNNING</span>';
      if(r.can_stop){
        right+='<button class="stopbtn" data-folder="'+esc(r.folder)+'" onclick="stopBot(this)">Stop</button>';
      }
    }else if(r.status==='completed'){
      right='<span class="rbadge completed">COMPLETED</span>';
    }else if(r.command){
      right='<code class="cmdbox" title="'+esc(r.command)+'">'+esc(r.command)+'</code>'
           +'<button class="cpy" data-cmd="'+esc(r.command)+'" onclick="cpy(this)">Copy</button>'
           +'<button class="startbtn" data-folder="'+esc(r.folder)+'" data-start="'+r.resume_start+'" data-end="'+r.range_end+'" onclick="startBot(this)">Start</button>';
    }else{
      right='<span class="rbadge unknown">RANGE UNKNOWN</span>';
    }
    const lr=(r.last_processed_row!=null)?('last processed row: '+r.last_processed_row):'no rows processed yet';
    h+='<div class="rrow"><div class="rleft"><div class="rfn" title="'+esc(r.folder)+'">'+esc(r.folder)+'</div><div class="rlr">'+esc(lr)+'</div></div><div class="rright">'+right+'</div></div>';
  }
  list.innerHTML=h;
}
async function openAddBotModal(){
  document.getElementById('newStartInput').value='';
  document.getElementById('newEndInput').value='';
  document.getElementById('newBotName').textContent='Bot-?';
  document.getElementById('newRangeHint').textContent='Loading next free chunk...';
  document.getElementById('newRangeHint').classList.remove('err');
  document.getElementById('addBotModal').classList.add('open');
  try{
    const r=await fetch('/api/next-chunk');
    const d=await r.json();
    document.getElementById('newBotName').textContent=d.bot||'Bot-?';
    if(d.free){
      document.getElementById('newStartInput').value=d.start;
      document.getElementById('newEndInput').value=d.end;
    }
    updateRangeHint();
  }catch(e){
    document.getElementById('newRangeHint').textContent='Could not load next-chunk suggestion';
  }
  document.getElementById('newStartInput').focus();
}
function closeAddBotModal(){document.getElementById('addBotModal').classList.remove('open')}
function updateRangeHint(){
  const sV=document.getElementById('newStartInput').value;
  const eV=document.getElementById('newEndInput').value;
  const hint=document.getElementById('newRangeHint');
  if(sV===''||eV===''){hint.textContent='Enter start and end';hint.classList.remove('err');return}
  const s=parseInt(sV,10),e=parseInt(eV,10);
  if(isNaN(s)||isNaN(e)||s<0||e>3641||s>=e){
    hint.textContent='Invalid: need 0 ≤ start < end ≤ 3641';
    hint.classList.add('err');return;
  }
  hint.classList.remove('err');
  hint.textContent='Will process excel rows '+(s+2)+'–'+(e-1)+' ('+(e-s-2)+' rows)';
}
async function confirmAddBot(){
  const s=parseInt(document.getElementById('newStartInput').value,10);
  const e=parseInt(document.getElementById('newEndInput').value,10);
  if(isNaN(s)||isNaN(e)||s<0||e>3641||s>=e){updateRangeHint();return}
  const btn=document.getElementById('addBotConfirm');
  btn.disabled=true;btn.textContent='Creating...';
  try{
    const r=await fetch('/api/add-bot',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({start:s,end:e})});
    const d=await r.json();
    if(r.ok&&d.ok){
      btn.textContent='Created '+d.bot;
      setTimeout(()=>{closeAddBotModal();btn.disabled=false;btn.textContent='Create';ld()},900);
    }else{
      document.getElementById('newRangeHint').textContent=d.error||('HTTP '+r.status);
      document.getElementById('newRangeHint').classList.add('err');
      btn.disabled=false;btn.textContent='Create';
    }
  }catch(err){
    document.getElementById('newRangeHint').textContent=String(err);
    document.getElementById('newRangeHint').classList.add('err');
    btn.disabled=false;btn.textContent='Create';
  }
}
async function resetBot(folder,ev){
  if(ev){ev.stopPropagation()}
  if(!confirm('Reset '+folder+'?\\n\\nCurrent progress will be archived. Bot will be ready for new work.'))return;
  try{
    const r=await fetch('/api/reset-bot',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({folder:folder})});
    const d=await r.json();
    if(r.ok&&d.ok){
      if(BOT_FILTER&&BOT_FILTER.folder===folder)clearBotFilter();
      ld();
    }else{
      alert('Reset failed: '+(d.error||'HTTP '+r.status));
    }
  }catch(e){alert('Reset failed: '+e)}
}
async function deleteBot(folder,ev){
  if(ev){ev.stopPropagation()}
  if(!confirm('Delete '+folder+'?\\n\\nProgress data will be preserved on disk but the bot will be removed from the dashboard.'))return;
  try{
    const r=await fetch('/api/delete-bot',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({folder:folder})});
    const d=await r.json();
    if(r.ok&&d.ok){
      if(BOT_FILTER&&BOT_FILTER.folder===folder)clearBotFilter();
      ld();
    }else{
      alert('Delete failed: '+(d.error||'HTTP '+r.status));
    }
  }catch(e){alert('Delete failed: '+e)}
}
async function stopBot(btn){
  const folder=btn.dataset.folder;
  btn.disabled=true;btn.textContent='Stopping...';
  try{
    const r=await fetch('/api/stop-bot',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({folder:folder})});
    const d=await r.json();
    if(r.ok&&d.ok){
      btn.textContent='Stopped';
      setTimeout(ld,800);
    }else{
      btn.textContent='Failed';btn.title=d.error||('HTTP '+r.status);
      setTimeout(()=>{btn.disabled=false;btn.textContent='Stop';btn.removeAttribute('title')},4000);
    }
  }catch(e){
    btn.textContent='Error';btn.title=String(e);
    setTimeout(()=>{btn.disabled=false;btn.textContent='Stop';btn.removeAttribute('title')},4000);
  }
}
async function startBot(btn){
  const folder=btn.dataset.folder;
  const start=parseInt(btn.dataset.start,10);
  const end=parseInt(btn.dataset.end,10);
  btn.disabled=true;btn.textContent='Starting...';
  try{
    const r=await fetch('/api/start-bot',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({folder:folder,start:start,end:end})});
    const d=await r.json();
    if(r.ok&&d.ok){
      btn.textContent='Started • pid '+d.pid;
      setTimeout(ld,1500);
    }else{
      btn.classList.add('err');btn.textContent='Failed';btn.title=d.error||('HTTP '+r.status);
      setTimeout(()=>{btn.classList.remove('err');btn.disabled=false;btn.textContent='Start';btn.removeAttribute('title')},4000);
    }
  }catch(e){
    btn.classList.add('err');btn.textContent='Error';btn.title=String(e);
    setTimeout(()=>{btn.classList.remove('err');btn.disabled=false;btn.textContent='Start';btn.removeAttribute('title')},4000);
  }
}
function cpy(btn){
  const cmd=btn.getAttribute('data-cmd');
  const done=()=>{const orig=btn.textContent;btn.textContent='Copied!';btn.classList.add('done');setTimeout(()=>{btn.textContent='Copy';btn.classList.remove('done');},1500);};
  if(navigator.clipboard&&navigator.clipboard.writeText){
    navigator.clipboard.writeText(cmd).then(done).catch(()=>{fallback(cmd,done)});
  }else{fallback(cmd,done)}
}
function fallback(text,done){
  const ta=document.createElement('textarea');ta.value=text;ta.style.position='fixed';ta.style.left='-9999px';
  document.body.appendChild(ta);ta.select();try{document.execCommand('copy');done()}catch(e){}finally{document.body.removeChild(ta)}
}
function renderBots(bots){
  BOTS=bots||[];
  const grid=document.getElementById('botGrid');
  const meta=document.getElementById('botMeta');
  if(!bots||!bots.length){grid.innerHTML='<div style="color:var(--t2);font-size:12px">No bot folders detected.</div>';meta.textContent='0 bots';return}
  const active=bots.filter(b=>b.active).length;
  meta.textContent=bots.length+' bots • '+active+' active • click a card to filter table';
  let h='';
  for(let i=0;i<bots.length;i++){
    const b=bots[i];
    const dotCls=b.active?'sdot on':'sdot off';
    const hasRange=(b.range_start!=null&&b.range_end!=null);
    const rng=hasRange
      ?('rows '+(b.range_start+2)+'–'+(b.range_end-1)+(b.assigned_total?' • '+b.assigned_total+' assigned':''))
      :'no range assigned — ready for work';
    const pendingDone=(hasRange && b.pending===0);
    const pendCell=hasRange
      ?(pendingDone
          ?'<div class="st s"><div class="v">✓</div><div class="l">Done</div></div>'
          :'<div class="st p"><div class="v">'+(b.pending==null?'—':b.pending)+'</div><div class="l">Pending</div></div>')
      :'<div class="st p"><div class="v">—</div><div class="l">Ready</div></div>';
    const isSel=(BOT_FILTER&&BOT_FILTER.folder===b.folder);
    const cls='bcard'+(isSel?' selected':'')+(hasRange?'':' locked');
    const title=hasRange?'Click to filter table to this bot\\'s rows':'Bot has no assignment — use the Work Dispatcher';
    h+='<div class="'+cls+'" data-i="'+i+'" title="'+title+'" onclick="selectBot('+i+')">'
      +'<button class="delbot" title="Remove '+esc(b.folder)+' from dashboard" data-folder="'+esc(b.folder)+'" onclick="deleteBot(this.dataset.folder,event)">×</button>'
      +'<div class="top"><div class="'+dotCls+'" title="'+(b.active?'Active':'Idle')+'"></div>'
      +'<div class="fn" title="'+esc(b.folder)+'">'+esc(b.folder)+'</div></div>'
      +'<div class="rng">'+esc(rng)+'</div>'
      +'<div class="stats">'
      +'<div class="st s"><div class="v">'+b.success+'</div><div class="l">Entered</div></div>'
      +'<div class="st d"><div class="v">'+b.post_dup+'</div><div class="l">Post-Dup</div></div>'
      +'<div class="st k"><div class="v">'+b.skipped+'</div><div class="l">Pre-Dup</div></div>'
      +'<div class="st f"><div class="v">'+b.failed+'</div><div class="l">Failed</div></div>'
      +pendCell
      +'</div>'
      +(!b.active?'<button class="resetbtn" data-folder="'+esc(b.folder)+'" onclick="resetBot(this.dataset.folder,event)" title="Archive current progress and clear the bot\\'s assignment">Reset</button>':'')
      +'<div class="la">Last activity: <b>'+fmShort(b.last_updated)+'</b></div>'
      +'</div>';
  }
  grid.innerHTML=h;
}
function selectBot(i){
  const b=BOTS[i];if(!b)return;
  if(b.range_start==null||b.range_end==null)return;
  if(BOT_FILTER&&BOT_FILTER.folder===b.folder){clearBotFilter();return}
  BOT_FILTER={folder:b.folder,start:b.range_start+2,end:b.range_end-1};
  document.querySelectorAll('.bcard').forEach(c=>c.classList.toggle('selected',parseInt(c.dataset.i,10)===i));
  const btn=document.getElementById('clrBot');btn.classList.add('on');
  document.getElementById('clrBotLbl').textContent='Showing rows '+BOT_FILTER.start+'–'+BOT_FILTER.end;
  fl();
  document.querySelector('.ts').scrollIntoView({behavior:'smooth',block:'start'});
}
function clearBotFilter(){
  BOT_FILTER=null;
  document.querySelectorAll('.bcard').forEach(c=>c.classList.remove('selected'));
  document.getElementById('clrBot').classList.remove('on');
  fl();
}
function up(d){
  const s=d.stats;
  document.getElementById('sT').textContent=s.total.toLocaleString();
  document.getElementById('sO').textContent=s.success.toLocaleString();
  document.getElementById('sD').textContent=s.post_dup.toLocaleString();
  document.getElementById('sE').textContent=s.failed.toLocaleString();
  document.getElementById('sS').textContent=s.skipped.toLocaleString();
  document.getElementById('sP').textContent=s.pending.toLocaleString();
  document.getElementById('sR').textContent='Success rate: '+s.success_rate+'%';
  document.getElementById('sED').textContent=s.form_fill_fails+' form / '+s.submit_fails+' submit / '+s.preform_fails+' pre-form';
  document.getElementById('sET').textContent=s.est_remaining_hours>0?'~'+s.est_remaining_hours+'h remaining':'—';
  document.getElementById('sF').textContent=s.total.toLocaleString()+' rows in Excel';
  document.getElementById('bv').textContent=s.billable.toLocaleString();
  document.getElementById('bk').textContent=s.success+' entered + '+s.post_dup+' post-submit duplicates';
  const pc=n=>(n/Math.max(1,s.total)*100).toFixed(2)+'%';
  document.getElementById('bo').style.width=pc(s.success);
  document.getElementById('bd').style.width=pc(s.post_dup);
  document.getElementById('bs').style.width=pc(s.skipped);
  document.getElementById('be').style.width=pc(s.failed);
  const pr=s.success+s.post_dup+s.failed+s.skipped;
  document.getElementById('pp').textContent=((pr/s.total)*100).toFixed(1)+'% processed';
  document.getElementById('fF').textContent=s.form_fill_fails;
  document.getElementById('fS').textContent=s.submit_fails;
  document.getElementById('fP').textContent=s.preform_fails;
  document.getElementById('fX').textContent=s.exception_fails;
  document.getElementById('lT').textContent=s.late_reg_total;
  document.getElementById('lD').textContent=s.late_reg_done;
  document.getElementById('lP').textContent=s.late_reg_post_dup;
  document.getElementById('lF').textContent=s.late_reg_failed;
  document.getElementById('lR').textContent=s.late_reg_pending;
  document.getElementById('tA').textContent=s.avg_duration>0?s.avg_duration+'s':'—';
  document.getElementById('tN').textContent=s.min_duration>0?s.min_duration+'s':'—';
  document.getElementById('tX').textContent=s.max_duration>0?s.max_duration+'s':'—';
  document.getElementById('tT').textContent=s.total_bot_time_min>0?s.total_bot_time_min+' min':'—';
  document.getElementById('tE').textContent=s.est_remaining_hours>0?s.est_remaining_hours+'h':'—';
  document.getElementById('xS').textContent=fm(s.started_at);
  document.getElementById('xL').textContent=fm(s.last_updated);
  document.getElementById('lu').textContent='Updated '+new Date().toLocaleTimeString();
  document.getElementById('xT').textContent=pr.toLocaleString();
  document.getElementById('xP').textContent=s.pending.toLocaleString();
  renderBots(d.bots||[]);
  renderResume(d.resume_commands||[]);
  renderDispatcher(s, d.bots||[]);
  R=d.rows;fl();
}
async function ld(){try{const r=await fetch('/api/data');up(await r.json())}catch(e){console.error(e)}}
let cd=''' + str(AUTO_REFRESH_SECONDS) + ''';
setInterval(()=>{cd--;if(cd<=0){ld();cd=''' + str(AUTO_REFRESH_SECONDS) + ''';}document.getElementById('tm').textContent='Refresh in '+cd+'s'},1000);
ld();
</script>

<div class="modal" id="addBotModal" onclick="if(event.target===this)closeAddBotModal()">
  <div class="modal-content">
    <h3>Add new bot</h3>
    <p>Will be created as <code id="newBotName">Bot-?</code></p>
    <div class="form-row">
      <label>Start (--start)<input type="number" id="newStartInput" min="0" max="3641" oninput="updateRangeHint()"></label>
      <label>End (--end)<input type="number" id="newEndInput" min="0" max="3641" oninput="updateRangeHint()"></label>
    </div>
    <div class="form-hint" id="newRangeHint">—</div>
    <div class="modal-buttons">
      <button class="cancel" onclick="closeAddBotModal()">Cancel</button>
      <button class="addbot" id="addBotConfirm" onclick="confirmAddBot()">Create</button>
    </div>
  </div>
</div>

</body>
</html>'''


def main():
    global EXCEL_RECORDS, EXCEL_RAW_DF

    try:
        sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="Pehchan Unified Live Tracking Dashboard")
    parser.add_argument("--port", type=int, default=DEFAULT_DASHBOARD_PORT,
                        help=f"Port to serve dashboard on (default: {DEFAULT_DASHBOARD_PORT})")
    args = parser.parse_args()
    port = args.port

    folders = discover_bot_folders()
    if not folders:
        print(f"ERROR: No subfolders with {PROGRESS_REL} found under {PARENT_DIR}")
        sys.exit(1)

    print(f"Parent dir: {PARENT_DIR}")
    print(f"Discovered {len(folders)} bot folders:")
    for f in folders:
        print(f"  - {os.path.basename(f)}")

    excel_path = find_excel_anywhere(folders)
    if not excel_path:
        print("ERROR: No Excel file found in any subfolder's data/ directory")
        sys.exit(1)

    print(f"\nLoading Excel: {excel_path}")
    EXCEL_RECORDS = load_excel_data(excel_path)
    EXCEL_RAW_DF = pd.read_excel(excel_path)
    print(f"Loaded {len(EXCEL_RECORDS)} records")

    server = HTTPServer(("0.0.0.0", port), DashboardHandler)
    print(f"\n{'='*50}")
    print(f"  UNIFIED DASHBOARD: http://localhost:{port}")
    print(f"  Auto-refreshes every {AUTO_REFRESH_SECONDS}s")
    print(f"  Ctrl+C to stop")
    print(f"{'='*50}\n")

    try:
        import webbrowser
        webbrowser.open(f"http://localhost:{port}")
    except Exception:
        pass

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDashboard stopped.")
        server.server_close()


if __name__ == "__main__":
    main()
