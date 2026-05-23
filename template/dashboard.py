"""
Pehchan Portal Automation — Live Tracking Dashboard
Serves a real-time dashboard at http://localhost:8050
Reads progress.json + Excel file for 100% accurate data.
"""

import argparse
import json
import os
import sys
import time
from http.server import HTTPServer, SimpleHTTPRequestHandler
from datetime import datetime
import pandas as pd

# ─── CONFIGURATION ───
PROGRESS_FILE = "logs/progress.json"
DEFAULT_DASHBOARD_PORT = 8050
AUTO_REFRESH_SECONDS = 10


def find_excel_file():
    candidates = [
        "1999-entries-new-1.xlsx"
    ,
    ]
    for f in candidates:
        if os.path.exists(f):
            return f
    if os.path.exists("data"):
        for f in os.listdir("data"):
            if f.endswith(".xlsx"):
                return os.path.join("data", f)
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


def load_progress():
    if not os.path.exists(PROGRESS_FILE):
        return {"completed_rows": [], "failed_rows": [], "skipped_rows": [],
                "post_submit_duplicates": [],
                "total_success": 0, "total_failed": 0, "total_skipped": 0,
                "total_post_submit_dup": 0,
                "started_at": None, "last_updated": None}
    with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    data.setdefault("post_submit_duplicates", [])
    data.setdefault("total_post_submit_dup", 0)
    return data


def build_dashboard_data(excel_records, progress):
    # Build completed lookup: row -> {duration, timestamp}
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

    # Build post-submit duplicate lookup
    post_dup_lookup = {}
    for entry in progress.get("post_submit_duplicates", []):
        post_dup_lookup[entry["row"]] = {"timestamp": entry.get("timestamp", "")}
    post_dup_set = set(post_dup_lookup.keys())

    # Build skipped lookup
    skipped_set = set()
    for s in progress.get("skipped_rows", []):
        skipped_set.add(s["row"])
    skipped_set -= completed_set
    skipped_set -= post_dup_set

    # Build failed lookup: row -> latest {error, detail, timestamp}
    failed_lookup = {}
    for f in progress.get("failed_rows", []):
        r = f["row"]
        if r not in failed_lookup or f["timestamp"] > failed_lookup[r]["timestamp"]:
            failed_lookup[r] = {
                "error": f.get("error", "Unknown"),
                "detail": f.get("detail", ""),
                "timestamp": f["timestamp"],
            }
    # Remove from failed if also in completed, post-dup, or skipped
    for r in completed_set | post_dup_set | skipped_set:
        failed_lookup.pop(r, None)

    # Compute per-entry durations (skip first entry — has login+captcha)
    durations = []
    for entry in progress.get("completed_rows", []):
        if isinstance(entry, dict) and entry.get("duration", 0) > 0:
            durations.append(entry["duration"])

    durations_for_avg = durations[1:] if len(durations) > 1 else durations
    avg_duration = round(sum(durations_for_avg) / len(durations_for_avg), 1) if durations_for_avg else 0
    min_duration = round(min(durations_for_avg), 1) if durations_for_avg else 0
    max_duration = round(max(durations_for_avg), 1) if durations_for_avg else 0
    total_bot_time = round(sum(durations), 1)

    # Classify each row
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

    # Stats
    total = len(excel_records)
    success_count = len(completed_set & {r["excel_row"] for r in excel_records})
    post_dup_count = len(post_dup_set & {r["excel_row"] for r in excel_records})
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


class DashboardHandler(SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        if self.path == "/api/data":
            self.send_json_response()
        elif self.path == "/" or self.path == "/index.html":
            self.send_dashboard_html()
        else:
            self.send_error(404)

    def send_json_response(self):
        progress = load_progress()
        data = build_dashboard_data(EXCEL_RECORDS, progress)
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
<title>Pehchan Tracker</title>
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

.ts{background:var(--card);border:1px solid var(--border);border-radius:8px;overflow:hidden}
.th{padding:14px 22px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center}
.th h3{font-size:13px;font-weight:600}
.tc{display:flex;gap:6px;align-items:center}
.fb{padding:4px 10px;border:1px solid var(--border);border-radius:4px;background:var(--bg);font-size:11px;font-family:var(--sn);cursor:pointer;color:var(--t2);transition:all .15s}
.fb:hover{background:var(--hover)}.fb.on{background:var(--txt);color:#fff;border-color:var(--txt)}
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
  <div><h1>PEHCHAN ENTRY TRACKER</h1><div class="sub">Rajasthan Civil Registration — Birth Record Digitization</div></div>
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

<div class="ts">
  <div class="th">
    <h3>All Records <span id="tC" style="color:var(--t3);font-weight:400"></span></h3>
    <div class="tc">
      <input type="text" class="si" placeholder="Search reg#, name, row..." id="sb" oninput="fl()">
      <button class="fb on" onclick="sf('all',this)">All</button>
      <button class="fb" onclick="sf('success',this)">Entered</button>
      <button class="fb" onclick="sf('post_dup',this)">Post-Dup</button>
      <button class="fb" onclick="sf('failed',this)">Failed</button>
      <button class="fb" onclick="sf('skipped',this)">Pre-Dup</button>
      <button class="fb" onclick="sf('pending',this)">Pending</button>
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
let R=[],cf='all';
function fm(i){if(!i)return'—';try{return new Date(i).toLocaleString('en-IN',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:true})}catch(e){return i}}
function sf(f,b){cf=f;document.querySelectorAll('.fb').forEach(x=>x.classList.remove('on'));b.classList.add('on');fl()}
function fl(){
  const q=document.getElementById('sb').value.toLowerCase(),tb=document.getElementById('tb');
  let h='',c=0;
  for(const r of R){
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
  R=d.rows;fl();
}
async function ld(){try{const r=await fetch('/api/data');up(await r.json())}catch(e){console.error(e)}}
let cd=''' + str(AUTO_REFRESH_SECONDS) + ''';
setInterval(()=>{cd--;if(cd<=0){ld();cd=''' + str(AUTO_REFRESH_SECONDS) + ''';}document.getElementById('tm').textContent='Refresh in '+cd+'s'},1000);
ld();
</script>
</body>
</html>'''


def main():
    global EXCEL_RECORDS

    parser = argparse.ArgumentParser(description="Pehchan Live Tracking Dashboard")
    parser.add_argument("--port", type=int, default=DEFAULT_DASHBOARD_PORT,
                        help=f"Port to serve dashboard on (default: {DEFAULT_DASHBOARD_PORT})")
    args = parser.parse_args()
    port = args.port

    excel_path = find_excel_file()
    if not excel_path:
        print("ERROR: No Excel file found in data/ folder")
        sys.exit(1)

    print(f"Loading Excel: {excel_path}")
    EXCEL_RECORDS = load_excel_data(excel_path)
    print(f"Loaded {len(EXCEL_RECORDS)} records")

    if not os.path.exists(PROGRESS_FILE):
        print(f"WARNING: {PROGRESS_FILE} not found — all rows show as pending.")

    server = HTTPServer(("0.0.0.0", port), DashboardHandler)
    print(f"\n{'='*50}")
    print(f"  DASHBOARD: http://localhost:{port}")
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