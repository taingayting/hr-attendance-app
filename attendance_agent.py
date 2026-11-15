"""
Attendance Agent  v1.6  (day AND night shifts; Sunday and holiday columns; stops only when a day is BOTH Sunday and holiday)  -  built for payroll safety

  python attendance_agent.py check   2026-10-10     -> shows the report, changes NOTHING
  python attendance_agent.py fill    2026-10-10     -> shows every cell it will write, asks YES, backs up, saves, double-checks
  python attendance_agent.py answers 2026-10-10 answers.csv   -> same, for answers you got from workers/supervisors

Files:    put your files in the folder "inbox" (or paper / fingerprint / attendance - the folder does not matter).
          The agent opens every file and recognises it by what is INSIDE. If it is not 100% sure, it STOPS and asks.
          Paper list = columns ID,HOURS,OT and the date in its name, e.g. paper_2026-10-10.csv
          HOURS = number in the 8H column (blank = 8; 4 = half day). OT = tick(s), e.g. 2 or 2/4 (two ticks: biggest counts, then compared with the fingerprint).
          Optional: --fp NAME  --att NAME  --paper NAME   to tell it exactly which file to use.
Safety:   nothing is saved without you typing YES | backup before every save | never overwrites a different value
          every written cell is read back and checked | every change is saved in memory/audit_log.csv
"""
import sys, os, glob, math, shutil, csv
from datetime import datetime, time, timedelta
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string as C

# ================= RULES  (change numbers here) =================
SHIFT_START      = time(7, 0)    # day shift start
SHIFT_END        = time(16, 0)   # OT starts after this
LATE_GRACE_MIN   = 3             # minutes allowed before "late"
OT_TOLERANCE_MIN = 10            # 17:56 counts as 18:00
EARLY_GRACE_MIN  = 3             # leaving up to this many minutes before 16:00 is still OK
MIN_PUNCHES      = 4             # a normal day has 4 or more punches: IN, lunch out, lunch in, OUT (my rule)
# (no 60-minute limit: confirmed by HR - late/early always gets the hours filled AND a remark; HR checks later)
NIGHT_START      = time(18, 30)  # night shift start
NIGHT_END        = time(3, 30)   # night shift end, next day (real data: many leave at 06:00 = 2.5H OT)
NIGHT_WINDOW_END = time(6, 45)   # punches until this time next morning still belong to last night's shift
DUP_SECONDS      = 90            # two punches closer than this = the same scan
SHEET            = "人事考勤"      # sheet name in your attendance Excel
CODES = ["A","HP","HH","PC","M","L1","L2","AL1","AL2","MC","HW","SW","NP","Holiday"]
# ================================================================
CHANGES = []                     # every cell we plan to write: (worker, cell, old, new)
SHIFT_PREF = {}                  # {worker ID: "Day" / "Night" / "Evening"}: the shift HR chose before (the website fills it from its memory)
TIMES_PREF = {}                  # {worker ID: ("04:00", "13:00")}: work times remembered for a worker (the website fills it from its memory)
NO_PAPER_IDS = {}                # {worker ID: department}: departments with NO paper today. The website fills it; they are filled from the fingerprint
RESIGNED_PREF = {}               # {worker ID: "2026-10-19"}: remembered as resigned since that day (the website fills it from its memory)
CV_PATH = None                   # the attendance Excel: its sheet CV has the resign dates (the website fills it)
NEW_WORKERS = set()              # new workers whose fingerprint is NOT registered yet (the website fills this from the worker book): paper is trusted for them

def norm(s):
    return "".join(str(s).split()).lower()

def fix_id(x):                   # "DMO10003" -> "DMO-10003"
    x = str(x).strip().upper()
    return x.replace("DMO", "DMO-") if x.startswith("DMO") and "-" not in x else x

SCAN_FOLDERS = ["inbox", "paper", "fingerprint", "attendance"]     # "backup" is never scanned
FILES = {}                                                          # filled in by identify()

def classify(path):
    """Open a file and say what it is, by its CONTENT. Returns (kind, reason)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".jpg", ".jpeg", ".png", ".heic", ".pdf"):
        return "photo", "picture/PDF of a paper sheet - I cannot read it yet"
    try:
        if ext == ".csv":
            cols = {str(c).strip() for c in pd.read_csv(path, nrows=0, encoding="utf-8-sig").columns}
        elif ext in (".xlsx", ".xlsm"):
            wb = openpyxl.load_workbook(path, read_only=True)
            if SHEET in wb.sheetnames:
                wb.close(); return "attendance", f"has the sheet '{SHEET}'"
            first = next(wb[wb.sheetnames[0]].iter_rows(max_row=1, values_only=True), ())
            cols = {str(c).strip() for c in first if c is not None}
            wb.close()
        else:
            return "unknown", "file type not recognised"
    except Exception as e:
        return "unknown", f"cannot open it ({type(e).__name__})"
    if {"Name", "No.", "Date/Time"} <= cols: return "fingerprint", "has columns Name, No., Date/Time"
    if "ID" in cols and cols & {"LATE_MIN", "CODE", "REASON"}: return "answers", "has columns ID + LATE_MIN/CODE/REASON"
    if {"ID", "OT"} <= cols: return "paper", "has columns ID, OT" + (", HOURS" if "HOURS" in cols else "")
    return "unknown", "columns not recognised"

def identify(day, need, opts):
    kinds = ("fingerprint", "attendance", "paper", "answers", "photo", "unknown")
    found = {k: [] for k in kinds}
    for folder in SCAN_FOLDERS:
        for f in sorted(glob.glob(os.path.join(folder, "*"))):
            if os.path.isdir(f) or os.path.basename(f).startswith("~$"): continue
            kind, why = classify(f)
            found[kind].append((f, why))
    print("\n===== FILE RECOGNITION (by what is inside each file) =====")
    for k in kinds:
        for f, why in found[k]:
            print(f"  {f:<48} -> {k.upper():<12} ({why})" + ("  [ignored]" if k in ("unknown", "photo", "answers") else ""))
    label = {"fingerprint": "fingerprint export", "attendance": "monthly attendance Excel", "paper": "paper list (columns ID,OT)"}
    flag = {"fingerprint": "fp", "attendance": "att", "paper": "paper"}
    chosen = {}
    for kind in need:
        if kind in flag and flag[kind] in opts:
            p = opts[flag[kind]]
            if not os.path.exists(p): sys.exit(f"STOP: you named '{p}' but I cannot find that file.")
            k, why = classify(p)
            if k != kind: sys.exit(f"STOP: you said '{p}' is the {label[kind]}, but inside it looks like: {k.upper()} ({why}).")
            chosen[kind] = p; continue
        cands = found[kind]
        if kind == "paper" and cands:
            dated = [c for c in cands if str(day) in os.path.basename(c[0])]
            if not dated:
                sys.exit(f"STOP: I found paper list(s) {[c[0] for c in cands]} but none has the date {day} in its file name. "
                         f"Rename it like paper_{day}.csv so I do not guess the date.")
            cands = dated
        if not cands:
            hint = ""
            if kind == "paper" and found["photo"]:
                hint = " I see a photo of the paper, but I cannot read photos yet - please type ID,OT into paper_" + str(day) + ".csv."
            sys.exit(f"STOP: no {label[kind]} found.{hint} Put it in the 'inbox' folder.")
        if len(cands) > 1:
            sys.exit(f"STOP: I found {len(cands)} files that look like the {label[kind]}: {[c[0] for c in cands]}. "
                     f"I will not guess. Keep only one, or tell me which: --{flag[kind]} FILENAME")
        chosen[kind] = cands[0][0]
    print("\nUSING:")
    for k, p in chosen.items(): print(f"  {label[k]:<28} = {p}")
    return chosen

# ---------------- read the inputs ----------------
_FP_CACHE = {}


def read_fingerprint(path):
    """Reading the Excel takes seconds, and the page asks for it again at every click: remember it until the file changes."""
    try:
        sig = (os.path.abspath(path), os.path.getmtime(path), os.path.getsize(path))
    except OSError:
        return _read_fingerprint(path)
    hit = _FP_CACHE.get(sig)
    if hit is None:
        _FP_CACHE.clear()                       # only the newest file is kept
        hit = _FP_CACHE[sig] = _read_fingerprint(path)
    return hit.copy()


def _read_fingerprint(path):
    df = pd.read_excel(path)
    df["dt"] = pd.to_datetime(df["Date/Time"])
    df["ID Number"] = df["No."].map(fix_id)      # "No." is more reliable than "ID Number"
    df = df.sort_values(["ID Number", "dt"])
    gap = df.groupby("ID Number")["dt"].diff().dt.total_seconds()
    return df[~(gap <= DUP_SECONDS)]               # drop double scans of the same finger

def load_history():
    p = "memory/worker_reasons.csv"; h = {}
    if os.path.exists(p):
        for _, a in pd.read_csv(p, dtype=str, encoding="utf-8-sig").fillna("").iterrows():
            h.setdefault(a["worker_id"], []).append((a["work_date"], a["reason"]))
    return h

OT_TICKS = {0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0}   # the OT columns on the paper
HALF_MORNING_OUT  = (time(10, 45), time(12, 20))       # half day, morning only: leaves around lunch
HALF_AFTERNOON_IN = (time(11, 45), time(13, 45))       # half day, afternoon only: arrives after lunch

def parse_num_list(v):
    # '' -> []   '2' -> [2.0]   '2/4' -> [2.0, 4.0]   '?' or other text -> None (unclear)
    if v is None or pd.isna(v) or str(v).strip() == "": return []
    out = []
    for p in str(v).replace(";", "/").replace("|", "/").split("/"):
        p = p.strip()
        if p == "": continue
        try: out.append(float(p))
        except ValueError: return None
    return out

def g(x):                                  # 2.0 -> "2"
    return f"{x:g}" if isinstance(x, (int, float)) else str(x)

def d(x):                                  # None -> "?"
    return "?" if x is None else g(x)

def _shift_text(v):
    """'Night', 'night - today only' -> 'Night'. Anything else -> ''."""
    v = "" if v is None or (isinstance(v, float) and v != v) else str(v).strip().lower()
    for k in ("Day", "Night", "Evening"):
        if v.startswith(k.lower()): return k
    return ""

def _span_ok(a, b):
    """Start a, End b ('HH:MM'). Same day (End later) or ending after midnight (End earlier), at most 14 hours long."""
    if a == b: return False
    m = lambda x: int(x[:2]) * 60 + int(x[3:])
    span = m(b) - m(a) if m(b) > m(a) else m(b) + 1440 - m(a)
    return 0 < span <= 14 * 60

def _hhmm_text(v):
    """'4:00' -> '04:00'. Empty -> ''. Something wrong -> '?'."""
    v = "" if v is None or (isinstance(v, float) and v != v) else str(v).strip()
    if not v: return ""
    import re as _re
    m = _re.fullmatch(r"(\d{1,2})[:.](\d{2})", v) or _re.fullmatch(r"(\d{2})(\d{2})", v)
    if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59: return "?"
    return f"{int(m.group(1)):02d}:{m.group(2)}"

def read_paper(path):
    if not os.path.exists(path): sys.exit(f"STOP: paper list not found: {path}")
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig") if path.endswith(".csv") else pd.read_excel(path, dtype=str)
    df = df[df["ID"].notna()]
    ids = [fix_id(i) for i in df["ID"]]
    dups = sorted({i for i in ids if ids.count(i) > 1})
    if dups: sys.exit("STOP: the same worker is on the paper list twice: " + ", ".join(dups))
    out = {}
    for pid, (_, r) in zip(ids, df.iterrows()):
        hv = parse_num_list(r["HOURS"]) if "HOURS" in df.columns else []
        if hv is None or len(hv) > 1 or (hv and not (0 < hv[0] <= 8)): hours = None      # unclear
        else: hours = hv[0] if hv else 8.0
        ticks = parse_num_list(r["OT"])
        if ticks is not None and any(t not in OT_TICKS for t in ticks): ticks = None     # not a real OT column
        code = str(r["CODE"]).strip().upper() if "CODE" in df.columns and pd.notna(r["CODE"]) else ""
        if code and code not in [c.upper() for c in CODES]: code = ""
        code = next((c for c in CODES if c.upper() == code), "")
        remark = str(r["REMARK"]).strip() if "REMARK" in df.columns and pd.notna(r["REMARK"]) else ""
        t_s, t_e = (_hhmm_text(r["START"]) if "START" in df.columns else ""), (_hhmm_text(r["END"]) if "END" in df.columns else "")
        times = (t_s, t_e) if (t_s and t_e and t_s != "?" and t_e != "?" and _span_ok(t_s, t_e)) else None      # both must be good times, 14 hours at most
        out[pid] = {"hours": hours, "ticks": ticks, "code": code, "remark": remark, "blank_hours": not hv, "times": times,
                    "shift": _shift_text(r["SHIFT"]) if "SHIFT" in df.columns else "",
                    "times_bad": (f"{t_s or '-'} to {t_e or '-'}" if (t_s or t_e) and not times else "")}
    return out

# ---------------- the thinking part ----------------
def hm(p, day):                                    # "06:52"  or "03:40+1" when it is the next calendar day
    return f"{p:%H:%M}" + ("+1" if p.date() > day else "")

def night_window(pts, d, force=False, strict=False):
    # A night shift that STARTS on date d = an IN punch 16:00-19:59 on d (workers may come early) AND a break punch 23:00-23:59 on d.
    # All its punches (also after midnight, until 06:45 next day) belong to work date d.
    # If the worker forgot the IN scan (no punch 16:00-19:59) it is still a night shift when: there is a break punch 23:00-23:59 AND
    # no punch at all in the daytime of d (a day worker always has morning punches). The window then starts at the first evening punch.
    on_d = [p for p in pts if p.date() == d]
    end = datetime.combine(d + timedelta(days=1), NIGHT_WINDOW_END)
    ins = [p for p in on_d if time(16, 0) <= p.time() <= time(19, 59)]
    if not any(p.hour == 23 for p in on_d):
        # He forgot the 23:xx break scan. It is still his night shift when: an IN scan 16:00-19:59, NO scan in the daytime, and
        #  - HR chose Night for him, or - he has a scan after midnight (a day worker never scans at 00:15).
        if not ins or any(NIGHT_WINDOW_END < p.time() < time(16, 0) for p in on_d): return None
        after_midnight = any(p.date() == d + timedelta(days=1) and p.hour < 4 for p in pts)
        if force or after_midnight: return [p for p in pts if ins[0] <= p <= end]
        if (not strict and any(p.date() == d + timedelta(days=1) and p.time() <= NIGHT_WINDOW_END for p in pts)
                and not any(p.date() == d + timedelta(days=1) and time(8, 0) <= p.time() <= time(14, 0) for p in pts)):     # next day has lunch scans: it is a DAY, leave it
            # no break scan and he left at 04:00-06:45 next morning (like 18:20 -> 05:30). It is his night shift only when he was a night worker the day before
            # or two days before (a day worker who stayed late for OT, then comes at 06:30 next morning, is NOT caught here).
            if any(night_window(pts, d - timedelta(days=k), strict=True) for k in (1, 2)): return [p for p in pts if ins[0] <= p <= end]
        return None
    if ins: return [p for p in pts if ins[0] <= p <= end]
    if any(NIGHT_WINDOW_END < p.time() < time(16, 0) for p in on_d): return None
    evening = [p for p in on_d if p.time() >= time(20, 0)]
    return [p for p in pts if evening[0] <= p <= end]

def analyse_night(nw, day):
    first, last = nw[0], nw[-1]
    start = datetime.combine(day, NIGHT_START)
    end = datetime.combine(day + timedelta(days=1), NIGHT_END)
    late = int((first - start).total_seconds() // 60)
    done = last.date() > day and last.time() >= time(3, 0)          # there is an OUT punch after midnight
    ot = 0.0
    if done and last > end:
        ot = max(0.0, math.floor(((last - end).total_seconds() / 60 + OT_TOLERANCE_MIN) / 30) * 0.5)
    early = int((end - last).total_seconds() // 60) if done and last < end else 0
    in_missing = first.time() >= time(22, 30) or first.date() > day          # the first scan is already the break scan: nobody scanned IN
    return {"kind": "NIGHT", "first": first, "last": last, "n": len(nw), "pts": nw, "done": done, "ot": ot, "in_missing": in_missing,
            "late": 0 if in_missing else (late if late > LATE_GRACE_MIN else 0), "early": early if early > EARLY_GRACE_MIN else 0}

VERSION = "2026-11-07 night without break scans (history), questions list"            # the app shows this, so you can see which file is really running
SWAP_DAY = 15                                   # night group -> day group change: on this date the night group works 16:00-24:00

def analyse_short_night(pts, d):
    """The 15th: workers leaving the night group work 16:00-24:00 (8H, OT up to about 02:30). Returns the same kind of dict as a night shift."""
    on_d = sorted(p for p in pts if p.date() == d and p.time() >= time(15, 0))
    if not on_d or not (time(15, 0) <= on_d[0].time() <= time(17, 30)): return None
    nxt = sorted(p for p in pts if d < p.date() <= d + timedelta(days=1) and p <= datetime.combine(d + timedelta(days=1), NIGHT_WINDOW_END))
    after = [p for p in nxt if p.time() > time(2, 30)]
    if after: return None                                           # scans late at night / early morning = a normal long night
    allp = on_d + nxt
    last = allp[-1]
    if not (datetime.combine(d, time(23, 30)) <= last <= datetime.combine(d + timedelta(days=1), time(2, 30))): return None
    start, end = datetime.combine(d, time(16, 0)), datetime.combine(d + timedelta(days=1), time(0, 0))
    first = allp[0]
    late = int((first - start).total_seconds() // 60)
    early = int((end - last).total_seconds() // 60) if last < end else 0
    ot = max(0.0, math.floor(((last - end).total_seconds() / 60 + OT_TOLERANCE_MIN) / 30) * 0.5) if last > end else 0.0
    return {"kind": "NIGHT", "short": True, "first": first, "last": last, "n": len(allp), "pts": allp, "done": True, "ot": ot,
            "late": late if late > LATE_GRACE_MIN else 0, "early": early if early > EARLY_GRACE_MIN else 0}

def analyse_custom(pts, day, hours, times):
    """A worker with his OWN work times (told to HR), for example 04:00-13:00 on the same day. Late counts after the Start, OT after the End.
    Hours stay as on the paper. Returns the same kind of dict as a normal day, or None when the times are not usable."""
    try:
        s_, e_ = [time(int(x.split(":")[0]), int(x.split(":")[1])) for x in times]
    except Exception:
        return None
    if e_ == s_: return None
    if e_ < s_:                                                             # ends after midnight: the same rules as the night shift, with HIS hours
        start, end = datetime.combine(day, s_), datetime.combine(day + timedelta(days=1), e_)
        if end - start > timedelta(hours=14): return None
        t = sorted(p for p in pts if start - timedelta(hours=2) <= p <= end + timedelta(hours=4))
        if not t:
            return {"kind": "NO_FP"}
        first, last = t[0], t[-1]
        late = int((first - start).total_seconds() // 60)
        done = last >= end - timedelta(minutes=30)                          # there is an OUT punch near his End (the normal night: 03:00 for 03:30)
        ot = max(0.0, math.floor(((last - end).total_seconds() / 60 + OT_TOLERANCE_MIN) / 30) * 0.5) if (done and last > end) else 0.0
        early = int((end - last).total_seconds() // 60) if done and last < end else 0
        in_missing = (first - start) >= timedelta(hours=4)                  # the first scan is already the break scan: nobody scanned IN
        return {"kind": "NIGHT", "first": first, "last": last, "n": len(t), "pts": t, "done": done, "ot": ot, "in_missing": in_missing,
                "late": 0 if in_missing else (late if late > LATE_GRACE_MIN else 0), "early": early if early > EARLY_GRACE_MIN else 0,
                "custom": f"{times[0]}-{times[1]}", "end_dt": end}
    start, end = datetime.combine(day, s_), datetime.combine(day, e_)
    used = set(night_window(pts, day - timedelta(days=1)) or [])           # punches of last night's shift are not his
    hi = min(end + timedelta(hours=8), datetime.combine(day, time(23, 59, 59)))           # OT after his End is counted until the end of that day
    t = sorted(p for p in pts if start - timedelta(hours=2) <= p <= hi and p not in used)
    if not t:
        return {"kind": "NO_FP"}
    first, last, n = t[0], t[-1], len(t)
    late = int((first - start).total_seconds() // 60)
    early = int((end - last).total_seconds() // 60)
    ot = 0.0
    if last > end:
        ot = max(0.0, math.floor(((last - end).total_seconds() / 60 + OT_TOLERANCE_MIN) / 30) * 0.5)
    base = {"first": first, "last": last, "n": n, "pts": t, "custom": f"{times[0]}-{times[1]}", "end_dt": end,
            "late": late if late > LATE_GRACE_MIN else 0, "early": early if early > EARLY_GRACE_MIN else 0}
    if hours is not None and hours < 8:
        return {**base, "kind": "HALF_UNSURE"}                              # half day on special times: you decide
    if n == 1:
        return {**base, "kind": "MISSING"}
    return {**base, "kind": "DAY", "ot": ot}

def analyse(pts, day, hours=8.0, pref=None, times=None):
    """pref = the shift HR chose for this worker before ('Day', 'Night', 'Evening') or None = work it out from the punches.
    times = ("04:00", "13:00") when the worker has his own start and end time (then the normal 07:00-16:00 / night times are not used)."""
    if times:
        c = analyse_custom(pts, day, hours, times)
        if c: return c
    if pref != "Day" and pref != "Night":                 # evening shift 16:00-00:00 (usually only 2 scans); on the 15th the night group works it too
        sn = analyse_short_night(pts, day)
        if sn: return sn
    nw = None if pref in ("Day", "Evening") else night_window(pts, day, force=(pref == "Night"))
    if nw: return analyse_night(nw, day)
    used = set([] if pref == "Day" else (night_window(pts, day - timedelta(days=1), force=(pref == "Night")) or []))      # punches of last night's shift
    t = sorted(p for p in pts if p.date() == day and p.hour >= 4 and p not in used)
    if not t:
        return {"kind": "NO_FP"}
    first, last, n = t[0], t[-1], len(t)
    end = datetime.combine(day, SHIFT_END)
    start = datetime.combine(day, SHIFT_START)
    def ot_of():
        mins = (last - end).total_seconds() / 60 + OT_TOLERANCE_MIN
        return max(0.0, math.floor(mins / 30) * 0.5) if last > end else 0.0
    late = int((first - start).total_seconds() // 60)
    early = int((end - last).total_seconds() // 60)
    base = {"first": first, "last": last, "n": n, "pts": t,
            "late": late if late > LATE_GRACE_MIN else 0, "early": early if early > EARLY_GRACE_MIN else 0}
    morning = n >= 2 and first.time() < time(11, 0) and HALF_MORNING_OUT[0] <= last.time() <= HALF_MORNING_OUT[1]
    afternoon = n >= 2 and HALF_AFTERNOON_IN[0] <= first.time() <= HALF_AFTERNOON_IN[1] and last.time() >= time(15, 45)
    if hours is not None and hours < 8:               # paper says half day: fingerprint must LOOK like a half day
        if morning:   return {**base, "kind": "HALF_OK", "ot": 0.0}
        if afternoon: return {**base, "kind": "HALF_OK", "ot": ot_of(), "late": 0}
        return {**base, "kind": "HALF_UNSURE"}
    if morning or afternoon:                           # paper says full day but fingerprint looks like a half day
        return {**base, "kind": "FP_HALF", "which": "morning only" if morning else "afternoon only"}
    if n == 1:
        return {**base, "kind": "MISSING"}
    return {**base, "kind": "DAY", "ot": ot_of()}

def calc_hours(first, last):
    """Only 2 or 3 punches: hours worked = first to last punch, minus 1 hour lunch if the worker was there from before 11:30 to after 12:30.
    Rounded DOWN to half hours, at most 8. It is only a SUGGESTION: you confirm it in Step 3."""
    h = (last - first).total_seconds() / 3600
    lunch = datetime.combine(first.date(), time(11, 30)); lunch_end = datetime.combine(first.date(), time(12, 30))
    if first <= lunch and last >= lunch_end: h -= 1
    return max(0.0, min(8.0, math.floor(h * 2) / 2))

LAST_SHIFT_MAX = {"Day": 0.0, "Night": 0.0}   # longest OT of each shift for the last day checked (shown in Step 3)

def build_report(day, fp_path, paper_path):
    fp, paper = read_fingerprint(fp_path), read_paper(paper_path)
    today = fp[fp["dt"].dt.date == day]
    if today.empty:
        sys.exit(f"STOP: the fingerprint file has NO records for {day}. It covers {fp['dt'].min():%Y-%m-%d %H:%M} to "
                 f"{fp['dt'].max():%Y-%m-%d %H:%M}. Download a newer export.")
    lastrec = fp["dt"].max()
    if lastrec.date() == day and lastrec.time() < time(21, 0):
        print(f"WARNING: the last record in the fingerprint file is {lastrec:%H:%M} on {day}. "
              f"If the file was downloaded before the day ended, OT may be missing.")
    names = dict(zip(fp["ID Number"], fp["Name"]))
    pts_of = {pid: sorted(grp["dt"].tolist()) for pid, grp in fp.groupby("ID Number")}
    hist = load_history()
    # longest OT of each shift on this date = the biggest fingerprint OT among workers whose punches are complete
    shift_max = {"Day": 0.0, "Night": 0.0}
    for _pid, _info in paper.items():
        if _info.get("times") or TIMES_PREF.get(_pid): continue          # a worker with his own work times does not count for "the longest OT today"
        _f = analyse(pts_of.get(_pid, []), day, _info["hours"], _info.get("shift") or SHIFT_PREF.get(_pid))
        if _f["kind"] == "NIGHT" and _f["done"] and _f["n"] >= MIN_PUNCHES and not _f.get("short"):
            shift_max["Night"] = max(shift_max["Night"], _f["ot"])
        elif _f["kind"] == "DAY" and _f["n"] >= MIN_PUNCHES:
            shift_max["Day"] = max(shift_max["Day"], _f["ot"])
    global LAST_SHIFT_MAX
    LAST_SHIFT_MAX = dict(shift_max)
    for _pid, _dept in NO_PAPER_IDS.items():                  # no paper for his department: the fingerprint IS the paper (8 hours + the fingerprint OT)
        if _pid in paper: continue
        _f = analyse(pts_of.get(_pid, []), day, 8.0, SHIFT_PREF.get(_pid), TIMES_PREF.get(_pid))
        if _f["kind"] == "NO_FP": continue
        _ot = float(_f.get("ot") or 0.0) if _f["kind"] in ("DAY", "NIGHT") else 0.0
        paper[_pid] = {"hours": 8.0, "ticks": ([_ot] if _ot else []), "code": "", "remark": "", "blank_hours": False, "times": None, "times_bad": "",
                       "shift": "", "no_paper": _dept}
    rows = []
    cpts = {}                                          # punches of workers with their own times (for the suggested hours)
    for pid, info in paper.items():
        hrs, ticks = info["hours"], info["ticks"]
        pot = None if ticks is None else (max(ticks) if ticks else 0.0)
        r = {"ID": pid, "Name": names.get(pid, ""), "Paper_Hours": hrs, "Paper_OT": pot, "In": "", "Out": "",
             "Finger_OT": "", "Late_min": 0, "Punches": "", "History": "", "hours": None, "ot_ok": False,
             "code": (info.get("code") or None), "ask": True, "night": False, "ot_fill": None, "Action": "", "Shift": "",
             "calc_hours": None, "Remark": info.get("remark", "")}
        rows.append(r)
        old = [h for h in hist.get(pid, []) if h[0] != str(day)]
        if old: r["History"] = f"{len(old)} earlier remark(s); last {old[-1][0]}: {old[-1][1]}"
        _by_paper = str(info.get("remark", "")).strip().lower().startswith("resign")
        _since = RESIGNED_PREF.get(pid)
        if _by_paper or (_since and _since <= str(day)):
            # RESIGNED: nothing is filled for him. The sheet CV of the Excel must show his resign date.
            txt, need = resigned_check(pid, day)
            scans = [p for p in pts_of.get(pid, []) if p.date() == day]
            why = "paper remark" if _by_paper else f"remembered since {_since}"
            if not _by_paper and scans:
                txt += f" | BUT he has {len(scans)} fingerprint scan(s) today: check. If he works again, press Forget in Step 2."; need = True
            r.update(resigned=True, ask=need, hours=None, code=None, ot_fill=None, ot_ok=False)
            if scans: r.update(In=hm(min(scans), day), Out=hm(max(scans), day), Punches=" ".join(hm(p, day) for p in sorted(scans)))
            r["Action"] = f"RESIGNED ({why}): nothing is filled. {txt}"
            continue
        if info.get("code") and (hrs is None or info.get("blank_hours")):
            # a CODE is written on the paper (A, HP, P, MC ...) and no hours: just fill the code. Only HP asks, and only when the worker HAS fingerprint scans.
            code = info["code"]; fc = analyse(pts_of.get(pid, []), day, 8.0, info.get("shift") or SHIFT_PREF.get(pid), info.get("times") or TIMES_PREF.get(pid))
            has_fp = fc["kind"] != "NO_FP"
            r.update(code=code, code_only=True, Shift=("" if not has_fp else ("Night" if fc["kind"] == "NIGHT" else "Day")))
            if has_fp: r.update(In=hm(fc["first"], day), Out=hm(fc["last"], day), Punches=" ".join(hm(p, day) for p in fc["pts"]))
            if code == "HP" and has_fp:
                r["Action"] = f"PAPER CODE HP but the worker HAS {fc['n']} fingerprint scan(s) -> ask the worker (HP is filled only after you decide)"; r["code"] = None
            else:
                r["ask"] = False
                r["Action"] = f"CODE {code} from the paper: filled" + (" (the worker has fingerprint scans, not asked)" if has_fp else " (no fingerprint)")
            continue
        if hrs is None:
            r["Action"] = "PAPER HOURS UNCLEAR (8H column) - check the paper"; continue
        _tm = info.get("times") or TIMES_PREF.get(pid)
        f = analyse(pts_of.get(pid, []), day, hrs, info.get("shift") or SHIFT_PREF.get(pid), _tm)
        k = f["kind"]
        if f.get("custom"):
            r["Times"] = f["custom"]
            r["Times_from"] = "from the paper table" if info.get("times") else "remembered for this worker"
            cpts[pid] = f["pts"]
        if k == "NO_FP":
            if pid in NEW_WORKERS and hrs is not None and pot is not None:
                r.update(hours=(int(hrs) if hrs == int(hrs) else hrs), ot_fill=pot, ot_ok=False, ask=False, Shift="New", new_nofp=True,
                         code=("HH" if hrs < 8 else (info.get("code") or None)))
                r["Action"] = f"NEW WORKER, NO FINGERPRINT YET: hours {g(hrs)}H and OT {g(pot)}H are filled from the PAPER only (nothing to compare)"
            elif info.get("code"):
                r["ask"] = False; r["code_only"] = True
                r["Action"] = f"CODE {info['code']} from the paper, no fingerprint: filled"
            else:
                r["Action"] = "NO FINGERPRINT -> ask the supervisor" + (" (new worker, but the paper hours or OT are unclear)" if pid in NEW_WORKERS else "")
            continue
        r.update(In=hm(f["first"], day), Out=hm(f["last"], day), Punches=" ".join(hm(p, day) for p in f["pts"]),
                 Shift=(("Evening" if f.get("short") else "Night") if k == "NIGHT" else "Day"))
        if k == "NIGHT":
            notes = []
            if hrs < 8: notes.append(f"PAPER says {g(hrs)}H (half day) on a NIGHT shift - not built, ask the worker")
            if not f["done"]:
                if f["last"].date() > day:
                    _cut = (f["end_dt"] - timedelta(minutes=30)).strftime("%H:%M") if f.get("end_dt") else "03:00"
                    notes.append(f"LAST SCAN {hm(f['last'], day)} (before {_cut}) - left early, or forgot to scan - ask the worker")
                else:
                    notes.append("NO SCAN after midnight - forgot to scan, or the file ends too early - ask the worker")
            if f.get("in_missing"): notes.append(f"IN SCAN MISSING (the first scan is {hm(f['first'], day)}: nobody scanned at the start of the night shift) - ask the worker")
            need_n = 2 if f.get("short") else MIN_PUNCHES          # the evening shift 16:00-00:00 has no break scan: 2 scans are normal
            if f["n"] < need_n: notes.append(f"ONLY {f['n']} PUNCHES (need {need_n}+)")
            if f["late"]: notes.append(f"LATE {f['late']} min")
            if f["early"]: notes.append(f"EARLY LEAVE {f['early']} min")
            multi = ticks is not None and len(ticks) > 1
            tick_txt = f" (ticks {', '.join(g(t) for t in ticks)} - biggest used)" if multi else ""
            if pot is None: notes.append("PAPER OT UNCLEAR - check the paper")
            blocked = hrs < 8 or not f["done"] or f["n"] < need_n or bool(f.get("in_missing"))
            smax = shift_max["Night"]
            if pot is not None and hrs >= 8 and not f["done"] and pot > 0:
                # OUT missing or early: nothing is filled, you decide. Only a hint about the longest night OT.
                if pot > smax: notes.append(f"PAPER OT {g(pot)}H is MORE than the longest night OT today ({g(smax)}H) - check the paper")
                else: notes.append(f"(paper OT {g(pot)}H is within the longest night OT today {g(smax)}H)")
            if pot is not None and pot != f["ot"]:
                if not blocked:
                    r["ot_fill"] = f["ot"]
                    notes.append(f"OT DIFFERENT (paper {g(pot)}H{tick_txt}, finger {g(f['ot'])}H) -> FILLED fingerprint OT {g(f['ot'])}H, CHECK LATER")
                else: notes.append(f"OT DIFFERENT (paper {g(pot)}H{tick_txt}, finger {g(f['ot'])}H)")
            r.update(Finger_OT=f["ot"], Late_min=f["late"], night=True,
                     hours=(None if blocked else (int(hrs) if hrs == int(hrs) else hrs)),
                     ot_ok=(pot is not None and pot == f["ot"] and not blocked))
            r["ask"] = bool(notes)
            tag = (f"EVENING shift 16:00-00:00{'' if (info.get("shift") or SHIFT_PREF.get(pid)) == 'Evening' else ' (worked out from the punches; if wrong, change Shift)'}" if f.get("short") else "NIGHT shift") + f" (OT {g(f['ot'])}H, {f['n']} punches)"
            r["Action"] = tag + (": " + "; ".join(notes) + (" -> ask the worker" if r["ot_fill"] is None else " (remark for you)") if notes else ": MATCH" + (f"{tick_txt}, same as fingerprint" if multi else ""))
            continue
        if k == "MISSING":
            r["Action"] = "ONLY 1 PUNCH -> ask the worker"; continue
        if k == "HALF_UNSURE":
            r["Action"] = f"PAPER says {g(hrs)}H (half day) but the fingerprint does not look like a half day -> ask the worker"; continue
        if k == "FP_HALF":
            r["Action"] = f"FINGERPRINT looks like a HALF DAY ({f['which']}) but the paper says {g(hrs)}H -> ask the worker"; continue
        half = hrs < 8
        notes, fill_hours = [], True
        if k == "DAY" and f["n"] < MIN_PUNCHES:
            notes.append(f"ONLY {f['n']} PUNCHES (need {MIN_PUNCHES}+)"); fill_hours = False
        if f["late"]: notes.append(f"LATE {f['late']} min")
        if k == "DAY" and f["early"] and f["last"] < f.get("end_dt", datetime.combine(day, SHIFT_END)):
            notes.append(f"EARLY LEAVE {f['early']} min")
        multi = ticks is not None and len(ticks) > 1
        tick_txt = f" (ticks {', '.join(g(t) for t in ticks)} - biggest used)" if multi else ""
        if pot is None: notes.append("PAPER OT UNCLEAR - check the paper")
        elif half and pot > 0: notes.append(f"HALF DAY with OT {g(pot)}H")
        elif pot != f["ot"]:
            if fill_hours and not half and k == "DAY":      # HR rule: fingerprint OT is filled, difference is remarked
                r["ot_fill"] = f["ot"]
                notes.append(f"OT DIFFERENT (paper {g(pot)}H{tick_txt}, finger {g(f['ot'])}H) -> FILLED fingerprint OT {g(f['ot'])}H, CHECK LATER")
            else:
                notes.append(f"OT DIFFERENT (paper {g(pot)}H{tick_txt}, finger {g(f['ot'])}H)")
        else: r["ot_ok"] = True
        r.update(Finger_OT=f["ot"], Late_min=f["late"], code=("HH" if half else (info.get("code") or None)),
                 hours=((int(hrs) if hrs == int(hrs) else hrs) if fill_hours else None))
        r["ask"] = bool(notes)
        if notes: r["Action"] = "; ".join(notes) + (" -> ask the worker" if r["ot_fill"] is None else " (remark for you)")
        else: r["Action"] = "MATCH" + (f" (half day {g(hrs)}H, code HH)" if half else "") + (f"{tick_txt}, same as fingerprint" if multi else "")
    for r, (pid, info) in zip(rows, paper.items()):
        if info.get("no_paper"):
            r["no_paper"] = info["no_paper"]
            r["Action"] = f"NO PAPER for {info['no_paper']} (filled from the fingerprint): " + ("MATCH: 8H" + (f" + OT {g(r['Finger_OT'])}H" if r.get("Finger_OT") else "") if r["Action"].strip() == "MATCH" else r["Action"])
        if r.get("resigned"): continue
        if info.get("times_bad"):
            r["Action"] = f"START/END NOT USED ({info['times_bad']}): they must be real times and End later than Start, so I used the normal shift - " + r["Action"]; r["ask"] = True
        if r.get("Times"):
            r["Action"] = f"OWN WORK TIMES {r['Times'].replace('-', ' to ')} ({r['Times_from']}) - " + r["Action"]
        if info.get("code") and "PAPER CODE" not in r["Action"] and not r.get("code_only"):
            r["Action"] = f"PAPER CODE {info['code']} - " + r["Action"]
        if info.get("remark"):
            r["Action"] = f"PAPER REMARK: {info['remark']} - " + r["Action"]; r["ask"] = True
        # only 2 or 3 punches: suggest the hours from the punches (you confirm in Step 3)
        f = None
        pts = cpts.get(pid) or [p for p in pts_of.get(pid, []) if p.date() == day and p.hour >= 4]
        if 2 <= len(pts) <= 3 and not r.get("code_only") and r["Shift"] not in ("Night", "Evening") and r["calc_hours"] is None:
            ch = calc_hours(pts[0], pts[-1])
            r["calc_hours"] = ch
            r["ask"] = True
            r["Action"] += f" | SUGGESTED HOURS {g(ch)}H from {pts[0]:%H:%M} to {pts[-1]:%H:%M} (1h lunch removed when covered) - you confirm"
    return rows

def show(rows, day):
    print(f"\n===== Attendance check for {day} =====")
    for r in rows:
        print(f"{r['ID']:<10} {str(r['Name'])[:18]:<18} paper {d(r['Paper_Hours'])+'H':<4} OT {d(r['Paper_OT']):<4} in {r['In']:<6} out {r['Out']:<6} -> {r['Action']}")
        if r["ask"] and r["Punches"]: print(f"{'':<11}fingerprint punches: {r['Punches']}")
        if r["History"]: print(f"{'':<11}history: {r['History']}")
    ask = [r for r in rows if r["ask"]]
    nights = [r for r in ask if r["night"]]
    print(f"\nMatched: {len(rows)-len(ask)}   Need to ask: {len(ask)}   Total on paper: {len(rows)}")
    if nights: print(f"   (of the {len(ask)}: {len(nights)} are night-shift workers)")
    os.makedirs("reports", exist_ok=True)
    old = f"reports/ask_list_{day}.csv"
    if os.path.exists(old):        # never lose the reasons you typed: keep the old list under another name
        os.replace(old, f"reports/ask_list_{day}_old_{datetime.now():%H%M%S}.csv")
    df = pd.DataFrame(ask)
    df["OT_filled_from_fingerprint"] = df["ot_fill"].apply(lambda v: "" if v is None else (f"{v:g}H" if v else "0H"))
    df = df.rename(columns={"Action": "REMARK (what is different)"})
    df.drop(columns=["hours", "ot_ok", "ask", "code", "night", "ot_fill", "calc_hours", "Remark", "Times_from"], errors="ignore").assign(Reason="", Final_OT_if_changed="").to_csv(
        f"reports/ask_list_{day}.csv", index=False, encoding="utf-8-sig")
    print(f"Ask list saved: reports/ask_list_{day}.csv  (type the worker's reason in the 'Reason' column)")

# ---------------- the Excel part ----------------
REUSE = None             # a dict while one person is planned for many days: the Excel is opened ONCE and all the days are written into the same open copy
ONLY_ID = None           # the ID while one person is planned: nothing is written for the other workers (no Holiday / A codes for everybody)

def open_attendance(day):
    path = FILES["attendance"]
    if REUSE is not None and path in REUSE:
        wb, cws = REUSE[path]
    else:
        wb = openpyxl.load_workbook(path)
        cws = openpyxl.load_workbook(path, data_only=True)[SHEET] if SHEET in wb.sheetnames else None
        if REUSE is not None: REUSE[path] = (wb, cws)
    if SHEET not in wb.sheetnames: sys.exit(f"STOP: sheet '{SHEET}' not found in {os.path.basename(path)}. Did the sheet name change?")
    ws = wb[SHEET]
    a2 = ws["A2"].value
    if isinstance(a2, datetime) and (a2.year, a2.month) != (day.year, day.month):
        sys.exit(f"STOP: this Excel is for {a2:%Y-%m}, but you asked for {day}. Wrong month file?")
    blocks = sorted([m for m in ws.merged_cells.ranges if m.min_row == 5 and m.min_col >= C("N") and m.max_col <= C("OY")],
                    key=lambda m: m.min_col)
    if day.day > len(blocks): sys.exit("STOP: cannot find this day's columns in the Excel.")
    b = blocks[day.day - 1]
    cached = cws.cell(5, b.min_col).value
    if isinstance(cached, datetime) and cached.date() != day:
        sys.exit(f"STOP: the columns found are for {cached.date()}, not {day}.")
    cols = {norm(ws.cell(7, c).value): c for c in range(b.min_col, b.max_col + 1)}
    for need in ("dayhours", "nighthours", "(a)dayotx1.5", "(a)nightotx2", "l/e", "code"):
        if need not in cols: sys.exit(f"STOP: column '{need}' not found in this day's block. Did the headers change?")
    rows = {ws.cell(r, 1).value: r for r in range(8, ws.max_row + 1) if ws.cell(r, 1).value}
    return path, wb, ws, cols, rows

OVERWRITE = False        # True only while History redoes an old day: then a cell that holds another value is replaced (and shown as Was -> New)
TARGETS = {}             # every cell the plan wants (even when it already has that value)

def put(ws, row, col, value, log):
    if col is None or value in (None, "", 0): return
    cell = ws.cell(row, col)
    wid = ws.cell(row, 1).value
    TARGETS[cell.coordinate] = value
    if cell.value == value: return
    if cell.value not in (None, "") and not OVERWRITE:
        log.append(f"  SKIPPED {cell.coordinate} ({wid}): already has {cell.value}, wanted {value}"); return
    CHANGES.append((wid, cell.coordinate, cell.value, value))
    cell.value = value

def write_day(ws, cols, row, hours=None, ot=None, late=None, code=None, log=None, night=False, daytype=None):
    if daytype in ("sunday", "holiday"):    # Sunday / holiday: hours + OT are ONE number in ONE column (day: (d) or (b), night: (N) or (n))
        if hours is not None:
            total = (hours or 0) + (ot or 0)
            nm = lambda v: int(v) if v == int(v) else v
            if daytype == "holiday":
                # HR rule: the first 8 hours go in (b) day / (n) night holiday OT; everything over 8 hours goes in 超时加班时数 (c) OT 2
                put(ws, row, cols.get(OFF_COLS[(daytype, bool(night))]), nm(min(total, 8)), log)
                if total > 8: put(ws, row, cols.get("(c)ot2"), nm(total - 8), log)
                put(ws, row, cols.get("l/e"), late, log)      # your Excel keeps the late/early minutes on a holiday too
            else:
                put(ws, row, cols.get(OFF_COLS[(daytype, bool(night))]), nm(total), log)
        if code:
            if code not in CODES: log.append(f"  code '{code}' is not in the dropdown list, skipped"); return
            put(ws, row, cols.get("code"), code, log)
        return
    if night:       # night shift: night hours + (a) night OT x2 (first 2h) + (c) OT 2 (after 2h)
        put(ws, row, cols.get("nighthours"), hours, log)
        if ot:      # same idea as day: first 2h in (a) night OTx2, everything after 2h in (c) OT 2
            put(ws, row, cols.get("(a)nightotx2"), min(ot, 2), log)
            put(ws, row, cols.get("(c)ot2"),       max(ot - 2, 0), log)
        put(ws, row, cols.get("l/e"), late, log)
        if code:
            if code not in CODES: log.append(f"  code '{code}' is not in the dropdown list, skipped"); return
            put(ws, row, cols.get("code"), code, log)
        return
    put(ws, row, cols.get("dayhours"), hours, log)
    if ot:
        put(ws, row, cols.get("(a)dayotx1.5"), min(ot, 2), log)
        put(ws, row, cols.get("(c)ot1.5"),     min(max(ot - 2, 0), 4), log)
        put(ws, row, cols.get("(c)ot2"),       max(ot - 6, 0), log)
    put(ws, row, cols.get("l/e"), late, log)
    if code:
        if code not in CODES: log.append(f"  code '{code}' is not in the dropdown list, skipped"); return
        put(ws, row, cols.get("code"), code, log)

def backup(path):
    os.makedirs("backup", exist_ok=True)
    name = os.path.basename(path).replace(".xlsx", f"_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    shutil.copy(path, os.path.join("backup", name)); print("Backup saved:", name)

def save_changes(path, wb, day, cmd):
    """Backup, save, read every cell back, write the audit log. Returns (ok, list_of_bad_cells)."""
    backup(path); wb.save(path)
    check = openpyxl.load_workbook(path)[SHEET]
    bad = [(c, n, check[c].value) for _, c, _, n in CHANGES if check[c].value != n]
    print(f"VERIFIED: all {len(CHANGES)} cells were read back from the saved file and are correct." if not bad
          else f"PROBLEM - these cells are not as expected: {bad}   Restore the backup!")
    os.makedirs("memory", exist_ok=True)
    new_file = not os.path.exists("memory/audit_log.csv")
    with open("memory/audit_log.csv", "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new_file: w.writerow(["saved_at", "work_date", "command", "worker_id", "cell", "old", "new"])
        for wid, c, o, n in CHANGES: w.writerow([datetime.now().isoformat(timespec="seconds"), day, cmd, wid, c, o, n])
    print("Audit log updated: memory/audit_log.csv")
    return (not bad), bad

def review_and_save(path, wb, day, cmd, log):
    print(f"\n===== Review before saving ({cmd} {day}) =====")
    for l in log: print(l)
    if not CHANGES: print("Nothing new to write."); return False
    for wid, coord, old, new in CHANGES:
        print(f"  {wid:<10} {coord:<7} {'(empty)' if old in (None, '') else old} -> {new}")
    print(f"\nThis will write {len(CHANGES)} cells for {len({c[0] for c in CHANGES})} workers.")
    if input("Type YES to save into your Excel (anything else cancels): ").strip() != "YES":
        print("Cancelled. Nothing was changed."); return False
    save_changes(path, wb, day, cmd)
    return True

# ---------------- commands ----------------
HOLIDAYS_2026 = [   # Cambodia paid public holidays 2026 (Prakas No. 167 of the Ministry of Labour, as listed by PPCBank; same dates on Emerhub and Siemreap.net). YOU confirm them in the website before saving.
    ("2026-01-01", "New Year's Day"), ("2026-01-07", "Victory over Genocide Day"), ("2026-03-08", "International Women's Day"),
    ("2026-04-14", "Khmer New Year day 1"), ("2026-04-15", "Khmer New Year day 2"), ("2026-04-16", "Khmer New Year day 3"),
    ("2026-05-01", "Labour Day and Visak Bochea"), ("2026-05-05", "Royal Ploughing Ceremony"), ("2026-05-14", "King's Birthday"),
    ("2026-06-18", "Queen Mother's Birthday"), ("2026-09-24", "Constitution Day"),
    ("2026-10-10", "Pchum Ben day 1"), ("2026-10-11", "Pchum Ben day 2"), ("2026-10-12", "Pchum Ben day 3"),
    ("2026-10-15", "Commemoration of the late King Father"), ("2026-10-29", "King's Coronation Day"), ("2026-11-09", "Independence Day"),
    ("2026-11-23", "Water Festival day 1"), ("2026-11-24", "Water Festival day 2"), ("2026-11-25", "Water Festival day 3"),
    ("2026-12-29", "Peace Day")]
HOLIDAY_FILE = "memory/holidays.csv"

def load_holidays():
    """{ 'YYYY-MM-DD': name } from memory/holidays.csv (made empty if missing)."""
    hol = {}
    os.makedirs("memory", exist_ok=True)
    if not os.path.exists(HOLIDAY_FILE):
        with open(HOLIDAY_FILE, "w", encoding="utf-8-sig") as f:
            f.write("DATE,NAME\n")
    try:
        for _, h in pd.read_csv(HOLIDAY_FILE, dtype=str, encoding="utf-8-sig").fillna("").iterrows():
            hol[str(h["DATE"]).strip()] = h.get("NAME", "")
    except Exception as e:
        sys.exit(f"STOP: cannot read {HOLIDAY_FILE} ({e}).")
    return hol

def set_holiday(day, on, name="", pay="paid"):
    """Make a day a holiday (on=True) or a normal day (on=False) in memory/holidays.csv. The other columns (PAY) of the other days are kept."""
    load_holidays()
    df = pd.read_csv(HOLIDAY_FILE, dtype=str, encoding="utf-8-sig").fillna("")
    if "PAY" not in df.columns: df["PAY"] = ""
    df = df[df["DATE"].str.strip() != str(day)]
    if on:
        df = pd.concat([df, pd.DataFrame([{"DATE": str(day), "NAME": name or "Holiday", "PAY": "payslip" if pay == "payslip" else ""}])], ignore_index=True)
    df = df.sort_values("DATE")
    if (df["PAY"] == "").all(): df = df.drop(columns=["PAY"])
    df.to_csv(HOLIDAY_FILE, index=False, encoding="utf-8-sig")

def holiday_pay(day):
    """'paid' (everyone who stays home gets the code Holiday) or 'payslip' (not paid ahead: the holiday pay is added later on the payslip, so no code)."""
    try:
        for _, h in pd.read_csv(HOLIDAY_FILE, dtype=str, encoding="utf-8-sig").fillna("").iterrows():
            if str(h["DATE"]).strip() == str(day):
                return "payslip" if str(h.get("PAY", "")).strip().lower().startswith("pay") else "paid"
    except Exception:
        pass
    return "paid"

def day_type(day):
    """'normal', 'sunday', 'holiday' or 'both' (a holiday that falls on a Sunday)."""
    hol = load_holidays()
    sun, h = day.weekday() == 6, str(day) in hol
    return "both" if (sun and h) else ("holiday" if h else ("sunday" if sun else "normal"))

def stop_if_off_day(day):
    """Sunday and holiday are filled now (all hours in ONE number in the Sunday/holiday column). Only a day that is BOTH a Sunday and a holiday stops: ask HR."""
    t = day_type(day)
    if t == "both":
        name = load_holidays().get(str(day), "")
        sys.exit(f"STOP: {day} is a Sunday AND a holiday ({name}). I do not know which column and pay to use. Tell me how your factory pays this day.")
    return t

# the ONE column each kind of day uses (all hours worked in one number, no split): Sunday day (d) / night (N), holiday day (b) / night (n)
OFF_COLS = {("sunday", False): "(d)ot2", ("sunday", True): "(n)ot2", ("holiday", False): "(b)holidayot", ("holiday", True): "(n)holidayot"}

def punched_ids(day):
    """IDs of workers who have punches for this work day (day or night shift). None if the fingerprint file is not known."""
    fpp = FILES.get("fingerprint")
    if not fpp or not os.path.exists(fpp): return None
    fp = read_fingerprint(fpp)
    came = set()
    for pid, grp in fp.groupby("ID Number"):
        if analyse(sorted(grp["dt"].tolist()), day, 8.0)["kind"] != "NO_FP": came.add(pid)
    return came

EXTRA_SKIP = set()     # worker IDs that must NOT get the code Holiday (the website fills it from the worker book: resigned / not started)

def _cv_dates(path):
    """{worker ID: (entry date, stop date as written)} read from the sheet CV. STOP if the sheet is missing."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "CV" not in wb.sheetnames:
        wb.close(); sys.exit("STOP: the Excel has no sheet CV, so I cannot tell who already resigned. I will not write Holiday codes.")
    out = {}
    for r in wb["CV"].iter_rows(values_only=True):
        if r and len(r) > 18 and r[1] and str(r[1]).strip().upper().startswith("DMO-"):
            out[str(r[1]).strip().upper()] = (_as_date(r[6]), r[18])
    wb.close()
    return out

def _as_date(v):
    from datetime import date as _date
    if isinstance(v, datetime): return v.date()
    if isinstance(v, _date): return v
    for f in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try: return datetime.strptime(str(v).strip(), f).date()
        except ValueError: pass
    return None

_CV_CACHE = {}

def _cv_stop_info():
    """({ID: (entry date, stop date as written)}, '') from the sheet CV of CV_PATH, or (None, why not). Read again only when the file changes."""
    p = CV_PATH
    if not p or not os.path.exists(p): return None, "no attendance Excel to look in"
    key = (os.path.abspath(p), os.path.getmtime(p), os.path.getsize(p))
    if key in _CV_CACHE: return _CV_CACHE[key]
    try: res = (_cv_dates(p), "")
    except SystemExit: res = (None, "the Excel has no sheet CV")
    except Exception as e: res = (None, f"could not read the sheet CV: {type(e).__name__}")
    _CV_CACHE.clear(); _CV_CACHE[key] = res
    return res

def resigned_check(pid, day):
    """What the sheet CV of the Excel says about a resigned worker. Returns (text, needs_you)."""
    cv, err = _cv_stop_info()
    if cv is None: return f"the CV sheet was not checked ({err})", True
    if pid not in cv: return "he is NOT in the CV sheet: check his resign date there", True
    raw = cv[pid][1]
    if raw in (None, ""): return "the CV sheet has NO resign date for him: fill it in the CV sheet", True
    sd = _as_date(raw)
    if sd is None: return f"the CV resign date '{raw}' cannot be read as a date: fix it in the CV sheet", True
    if sd <= day: return f"CV resign date {sd:%d-%m-%Y}: OK", False
    return f"the CV resign date is {sd:%d-%m-%Y}, LATER than this day: fix the CV sheet, or he has not resigned yet", True

def plan_fill(day, rows):
    """Work out every cell to write (nothing is saved). Returns (path, wb, log, n_filled). CHANGES holds the cells."""
    CHANGES.clear()
    dtype = stop_if_off_day(day)
    path, wb, ws, cols, xl = open_attendance(day)
    if dtype in ("sunday", "holiday"):
        miss = [OFF_COLS[(dtype, n)] for n in (False, True) if OFF_COLS[(dtype, n)] not in cols]
        if miss: sys.exit(f"STOP: this day's block has no column {', '.join(miss)} (the {dtype} column). Did the headers change?")
    log = []; n = 0
    for r in rows:
        if r["ID"] not in xl: log.append(f"  NOT IN EXCEL: {r['ID']} (new worker?) - not filled"); continue
        if dtype == "holiday" and (r["hours"] or r["code"]):
            # HR rule: on a holiday EVERY worker has the code Holiday and no other code (no HH, no leave code), whether he came or not.
            # Paid holiday: code Holiday. "Pay later" holiday: no code (the pay is added later on the payslip).
            r = dict(r, code=("Holiday" if holiday_pay(day) != "payslip" else None))
        if r["hours"] or r["code"]:          # a code alone (HP, A, MC ...) is also written, even when there are no hours
            if r["code"] == "MC" and dtype == "normal" and not r["hours"]:      # HR rule: a sick-leave (MC) working day = 8 hours in the day-hours column
                r = dict(r, hours=8, night=False, ot_fill=None, ot_ok=False)
                log.append(f"  INFO: {r['ID']} has code MC: 8 hours are written in the day hours (HR rule).")
            write_day(ws, cols, xl[r["ID"]], hours=r["hours"], ot=(r["ot_fill"] if r["ot_fill"] is not None else (r["Paper_OT"] if r["ot_ok"] else None)), code=r["code"], late=(r.get("le_fill") or None), log=log, night=r["night"], daytype=dtype); n += 1
    if ONLY_ID:
        return path, wb, log, n
    if dtype in ("sunday", "holiday"):
        listed = {r["ID"] for r in rows}
        came = punched_ids(day)
        if came is None and dtype == "holiday":
            sys.exit("STOP: I need the fingerprint file to know who worked on the holiday (everybody else gets the code Holiday).")
        if came is not None:
            for pid in sorted(came - listed):
                if pid in xl: log.append(f"  {pid} has fingerprint punches on this {dtype} but is NOT on the paper list: not filled. If the worker worked, add them to the paper list.")
        if dtype == "holiday" and holiday_pay(day) == "payslip":
            log.append("  INFO: this holiday is set to 'pay later on the payslip': workers who stayed home get NO code (left empty). Workers who came are filled as usual.")
        elif dtype == "holiday":
            cv = _cv_dates(path)                                          # {ID: (entry, stop_raw)} from the sheet CV (typed values, not formulas)
            k = 0
            for wid, row in xl.items():
                if wid in listed or wid in came or wid in EXTRA_SKIP: continue
                ent, stp = cv.get(wid, (None, None))
                if ent and ent > day: continue                            # not hired yet
                if stp not in (None, ""):
                    sd = _as_date(stp)
                    if sd is None or sd <= day: continue                  # already resigned (or a date I cannot read: not touched)
                put(ws, row, cols.get("code"), "Holiday", log); k += 1
            log.append(f"  INFO: {k} workers did not work on this holiday: they get the code Holiday (paid holiday). Workers who resigned, or have not started yet, are skipped.")
    if dtype == "normal":
        _absent_codes(day, rows, path, ws, cols, xl, log)
    return path, wb, log, n

def _absent_codes(day, rows, path, ws, cols, xl, log):
    """HR rule: a normal working day, no fingerprint punch, not on the paper, no leave paper / letter, still working there -> code A."""
    listed = {r["ID"] for r in rows}
    came = punched_ids(day)
    if came is None:
        log.append("  INFO: no fingerprint file, so I do not write the code A for workers without a punch."); return
    if len(came) < max(5, len(xl) // 10):
        log.append(f"  WARNING: the fingerprint has punches for only {len(came)} workers on {day}. I do NOT write the code A (the file may not cover this day)."); return
    cv = _cv_dates(path); k = 0
    for wid, row in xl.items():
        if wid in listed or wid in came or wid in EXTRA_SKIP: continue
        ent, stp = cv.get(wid, (None, None))
        if ent and ent > day: continue
        if stp not in (None, ""):
            sd = _as_date(stp)
            if sd is None or sd <= day: continue
        if any(ws.cell(row, cols[c]).value not in (None, "") for c in ("dayhours", "nighthours") if c in cols): continue
        put(ws, row, cols.get("code"), "A", log); k += 1
    log.append(f"  INFO: {k} workers have no punch, are not on the paper and have no leave paper or letter: they get the code A (absent).")

MANAGED = ("dayhours", "nighthours", "(a)dayotx1.5", "(a)nightotx2", "(c)ot1.5", "(c)ot2", "l/e", "code", "(b)holidayot", "(n)holidayot", "(d)ot2", "(n)ot2")

def plan_fill_redo(day, rows, clear_ids=(), clean_all=False):
    """Redo an OLD day: like plan_fill, but a cell that holds another value is replaced (CHANGES = worker, cell, old, new).
    clear_ids = workers you decided in Step 3: a cell the program writes for the day and that the new numbers leave empty is cleared (new = None).
    Nothing is saved here."""
    global OVERWRITE
    TARGETS.clear()
    OVERWRITE = True
    try:
        path, wb, log, n = plan_fill(day, rows)
        ws = wb[SHEET]
        if clear_ids or clean_all:
            _p, _wb2, _ws2, cols, xl = open_attendance(day)
            if clean_all: clear_ids = set(clear_ids) | set(xl)       # the day type changed (holiday on/off): every worker's cells of this day are made again
            for wid in clear_ids:
                row = xl.get(wid)
                if not row: continue
                for name in MANAGED:
                    c = cols.get(name)
                    if not c: continue
                    cell = ws.cell(row, c)
                    if cell.value not in (None, "") and cell.coordinate not in TARGETS:
                        CHANGES.append((wid, cell.coordinate, cell.value, None))
                        cell.value = None
        return path, wb, log, n
    finally:
        OVERWRITE = False

def plan_all_holidays(skip_ids=()):
    """Every paid public holiday inside the month of the attendance Excel: workers who did not work get the code Holiday (nothing is saved here).
    Only holidays that the fingerprint file covers are filled (the fingerprint shows who worked). A cell that has something is never touched.
    Returns (path, wb, log, {day: workers filled})."""
    CHANGES.clear()
    path = FILES["attendance"]
    wb = openpyxl.load_workbook(path)
    if SHEET not in wb.sheetnames: sys.exit(f"STOP: sheet '{SHEET}' not found in {os.path.basename(path)}.")
    ws = wb[SHEET]
    cached = openpyxl.load_workbook(path, data_only=True)[SHEET]
    blocks = sorted([m for m in ws.merged_cells.ranges if m.min_row == 5 and m.min_col >= C("N") and m.max_col <= C("OY")], key=lambda m: m.min_col)
    rows = {ws.cell(r, 1).value: r for r in range(8, ws.max_row + 1) if ws.cell(r, 1).value}
    cv = _cv_dates(path)
    fpp = FILES.get("fingerprint"); span = None
    if fpp and os.path.exists(fpp):
        fp = read_fingerprint(fpp); span = (fp["dt"].min().date(), fp["dt"].max().date())
    log, per_day = [], {}
    for b in blocks:
        d = cached.cell(5, b.min_col).value
        if not isinstance(d, datetime): continue
        d = d.date(); t = day_type(d)
        if t not in ("holiday", "both"): continue
        if t == "both": log.append(f"  INFO {d:%d %b}: Sunday AND public holiday. Only the CODE Holiday is written for workers who did not work (like your Excel does). Hours of workers who worked are not touched here.")
        if holiday_pay(d) == "payslip":
            log.append(f"  {d:%d %b}: this holiday is set to 'pay later on the payslip': no code written."); continue
        cols = {norm(ws.cell(7, c).value): c for c in range(b.min_col, b.max_col + 1)}
        if "code" not in cols: log.append(f"  {d:%d %b}: the code column was not found."); continue
        if not span or not (span[0] <= d <= span[1]):
            log.append(f"  {d:%d %b} ({load_holidays().get(str(d), 'holiday')}): the fingerprint file in Step 1 does not cover this day, so I cannot tell who worked: NOT filled."); continue
        hour_cols = [cols[n_] for n_ in ("dayhours", "nighthours", "(b)holidayot", "(n)holidayot", "(d)ot2", "(n)ot2") if n_ in cols]
        n_worked = sum(1 for row in rows.values() if any(ws.cell(row, c).value not in (None, "") for c in hour_cols))
        if rows and n_worked / len(rows) >= 0.3 and t == "holiday":
            log.append(f"  {d:%d %b} ({load_holidays().get(str(d), 'holiday')}): {n_worked} of {len(rows)} workers already have hours on this day, so it looks like a normal working day: NOT filled. "
                       "If your factory gave this day off, tell me; if it is not a holiday for you, remove it in the holiday list.")
            continue
        came = punched_ids(d) or set(); k = 0
        for wid, row in rows.items():
            if wid in came or wid in EXTRA_SKIP or wid in skip_ids: continue
            ent, stp = cv.get(wid, (None, None))
            if ent and ent > d: continue
            if stp not in (None, ""):
                sd = _as_date(stp)
                if sd is None or sd <= d: continue
            if any(ws.cell(row, c).value not in (None, "") for c in cols.values()): continue
            put(ws, row, cols["code"], "Holiday", log); k += 1
        per_day[str(d)] = k
    return path, wb, log, per_day

def cmd_fill(day, rows):
    path, wb, log, n = plan_fill(day, rows)
    print(f"\nWorkers on paper: {len(rows)}   clear enough to fill: {n}   left for you: {len(rows)-n}")
    review_and_save(path, wb, day, "fill", log)

def cmd_answers(day, csv_path):
    path, wb, ws, cols, xl = open_attendance(day)
    dtype = stop_if_off_day(day)
    log = []; notes = []
    df = pd.read_csv(csv_path, dtype=str, encoding="utf-8-sig").fillna("")
    for _, a in df.iterrows():
        pid = fix_id(a["ID"])
        if pid not in xl: log.append(f"  NOT IN EXCEL: {pid} - not filled"); continue
        num = lambda v: float(v) if v != "" else None
        write_day(ws, cols, xl[pid], hours=num(a.get("HOURS", "")), ot=num(a.get("OT", "")),
                  late=num(a.get("LATE_MIN", "")), code=a.get("CODE", "").strip() or None, log=log, daytype=dtype)
        notes.append(f"- {day} {pid}: " + ", ".join(f"{k}={v}" for k, v in a.items() if k != "ID" and v != ""))
    if review_and_save(path, wb, day, "answers", log):
        with open("memory/daily_log.md", "a", encoding="utf-8") as f: f.write("\n".join(notes) + "\n")
        reasons = [(fix_id(a["ID"]), a.get("REASON", "").strip(), a.get("CODE", "").strip()) for _, a in df.iterrows() if a.get("REASON", "").strip()]
        if reasons:
            new_file = not os.path.exists("memory/worker_reasons.csv")
            with open("memory/worker_reasons.csv", "a", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                if new_file: w.writerow(["saved_at", "work_date", "worker_id", "reason", "code"])
                for pid, reason, code in reasons: w.writerow([datetime.now().isoformat(timespec="seconds"), day, pid, reason, code])
            print("Reasons remembered: memory/worker_reasons.csv (shown next time the same worker is on your list)")

if __name__ == "__main__":
    pos, opts, i, args = [], {}, 0, sys.argv[1:]
    while i < len(args):
        if args[i] in ("--fp", "--att", "--paper") and i + 1 < len(args): opts[args[i][2:]] = args[i + 1]; i += 2
        else: pos.append(args[i]); i += 1
    if len(pos) < 2 or pos[0] not in ("check", "fill", "answers"): sys.exit(__doc__)
    cmd = pos[0]
    try: day = datetime.strptime(pos[1], "%Y-%m-%d").date()
    except ValueError: sys.exit("STOP: the date must look like 2026-10-10")
    stop_if_off_day(day)
    if cmd == "answers":
        if len(pos) < 3: sys.exit("STOP: tell me the answers file, e.g.  answers 2026-10-10 answers.csv")
        FILES = identify(day, ("attendance",), opts)
        cmd_answers(day, pos[2])
    else:
        FILES = identify(day, ("fingerprint", "attendance", "paper"), opts)
        rows = build_report(day, FILES["fingerprint"], FILES["paper"])
        show(rows, day)
        if cmd == "fill": cmd_fill(day, rows)


# ===================================================================== SOME WORKERS, MANY DAYS, FROM THE FINGERPRINT ONLY
def people_days(people, days, fp_path, att_path, empty_paper):
    """people = {worker ID: department}. For every worker and day: what the fingerprint says, with the same logic as the daily work (no paper: 8 hours + the fingerprint OT).
    Returns a list of dicts: wid, day, type, found, punches, shift, hours, ot, code, le, ask (needs your answer), note (why nothing is filled), night."""
    global NO_PAPER_IDS
    fp = read_fingerprint(fp_path)
    cv = _cv_dates(att_path) if att_path else {}
    out = []
    keep, NO_PAPER_IDS = NO_PAPER_IDS, dict(people)
    try:
        for day in days:
            t = day_type(day)
            _today = fp[fp["dt"].dt.date == day]
            no_fp_day = _today.empty
            few = (not no_fp_day) and _today["ID Number"].nunique() < 5          # the file has only a few people this day: "no punch" may just mean "not in the file"
            rows_by = None
            for wid in people:
                d = {"wid": wid, "day": day, "type": t, "found": "", "punches": "", "shift": "", "hours": None, "ot": None, "code": "", "le": None, "ask": False, "note": "", "night": False}
                out.append(d)
                ent, stp = cv.get(wid, (None, None))
                stp_d = _as_date(stp) if stp not in (None, "") else None
                if t == "both":
                    d["note"] = "Sunday AND holiday: I do not know the pay rule. Nothing is filled."; continue
                if ent and ent > day:
                    d["note"] = "Not hired yet (entry date in sheet CV). Nothing is filled."; continue
                if RESIGNED_PREF.get(wid) and RESIGNED_PREF[wid] <= str(day):
                    d["note"] = f"Remembered as resigned since {RESIGNED_PREF[wid]}. Nothing is filled."; continue
                if stp_d is not None and stp_d <= day:
                    d["note"] = f"Resign date {stp_d} in sheet CV. Nothing is filled."; continue
                if no_fp_day:
                    d["note"] = "The fingerprint file has no records for this day."; continue
                if rows_by is None:
                    try:
                        rows_by = {x["ID"]: x for x in build_report(day, fp_path, empty_paper)}
                    except SystemExit as e:
                        rows_by = {}; d["note"] = str(e).replace("STOP: ", "")[:160]; continue
                r = rows_by.get(wid)
                if r is None:                                                    # no punches at all this day
                    if few: d["found"] = f"No punch, but the file has punches for only {_today['ID Number'].nunique()} worker(s) this day: I do not write A / Holiday. Tick and type the code yourself if you are sure."; continue
                    if t == "normal": d["code"] = "A"; d["found"] = "No punch on a working day -> A (absent)"
                    elif t == "holiday" and holiday_pay(day) != "payslip": d["code"] = "Holiday"; d["found"] = "Holiday, no punch -> code Holiday"
                    else: d["found"] = "No punch (" + t + ")"
                    continue
                d.update(found=r.get("Action", ""), punches=r.get("Punches", ""), shift=r.get("Shift", ""), night=bool(r.get("night")), ask=bool(r.get("ask")), code=r.get("code") or "")
                if not r.get("ask"):
                    d["hours"] = r.get("hours"); d["le"] = r.get("le_fill") or None
                    d["ot"] = r["ot_fill"] if r.get("ot_fill") is not None else (r.get("Paper_OT") if r.get("ot_ok") else None)
                elif r.get("calc_hours"):
                    d["note"] = f"My suggestion: {r['calc_hours']:g} hours (you decide)"
    finally:
        NO_PAPER_IDS = keep
    return out


def plan_people(items, fp_path, att_path):
    """items = [{wid, day, hours, ot, code, night, le}] (what you confirmed). The Excel is opened once. Returns (path, wb, log, changes).
    Cells with another value are replaced (shown Was -> New). Nothing is saved."""
    global REUSE, ONLY_ID
    ids = {it["wid"] for it in items}
    REUSE, ONLY_ID = {}, ids
    FILES["attendance"] = att_path; FILES["fingerprint"] = fp_path
    allch, log, wb, path = [], [], None, att_path
    try:
        for day in sorted({it["day"] for it in items}):
            todo = [it for it in items if it["day"] == day]
            rows = [{"ID": it["wid"], "Name": "", "hours": it.get("hours"), "ot_fill": it.get("ot"), "Paper_OT": None, "ot_ok": False, "code": it.get("code") or None,
                     "night": bool(it.get("night")), "le_fill": it.get("le")} for it in todo]
            path, wb, lg, n = plan_fill_redo(day, rows, clear_ids=tuple(it["wid"] for it in todo))
            log += [f"{day}: {l.strip()}" for l in lg]
            allch += CHANGES
    finally:
        REUSE, ONLY_ID = None, None
    return path, wb, log, allch
